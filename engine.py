"""Veloce Commerce Operations - ledger calculation engine.

Rules that keep closing figures from drifting:
  * No floats. Money = integer paise, FX rates = integer micro-units (1e-6).
  * Gross, net, FX variance and running balance are DERIVED here, never stored.
  * Every derived figure comes from one function, used by every endpoint.
"""
from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import date
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from typing import Iterable, Optional

RATE_SCALE = 10**6
SUPPORTED_FOREIGN = {"USD", "EUR", "GBP", "AED", "SGD", "AUD", "CAD"}  # all 2-decimal currencies
# Extend only after checking the current RBI purpose-code list / with your bank (AD category-I).
PURPOSE_CODES = {"P0802", "P0807"}
REVENUE_CATEGORIES = {"Export Revenue"}

TXN_RE = re.compile(r"^[A-Za-z0-9_-]{3,40}$")
FIRA_RE = re.compile(r"^[A-Za-z0-9/_-]{6,40}$")
IFSC_RE = re.compile(r"^[A-Z]{4}0[A-Z0-9]{6}$")
SWIFT_RE = re.compile(r"^[A-Z]{4}[A-Z]{2}[A-Z0-9]{2}([A-Z0-9]{3})?$")


class ValidationError(ValueError):
    def __init__(self, errors):
        self.errors = [errors] if isinstance(errors, str) else list(errors)
        super().__init__("; ".join(self.errors))


# ---------- exact-number helpers -------------------------------------------------
def _dec(v) -> Decimal:
    if isinstance(v, bool):
        raise ValueError("not a number")
    try:
        return v if isinstance(v, Decimal) else Decimal(str(v).strip().replace(",", ""))
    except InvalidOperation:
        raise ValueError(f"not a number: {v!r}") from None


def to_minor(v) -> int:
    """'1500.25' -> 150025 (cents or paise). Rejects more than 2 decimals."""
    d = _dec(v) * 100
    if d != d.to_integral_value():
        raise ValueError(f"more than 2 decimal places: {v!r}")
    return int(d)


def to_micro(v) -> int:
    """FX rate -> micro-units. 84.15000000000001 (Excel float noise) -> 84150000."""
    return int((_dec(v) * RATE_SCALE).quantize(Decimal(1), ROUND_HALF_UP))


def div_round(n: int, d: int) -> int:
    """Integer division, half away from zero (symmetric, so a reversal negates exactly)."""
    q, r = divmod(abs(n), d)
    if 2 * r >= d:
        q += 1
    return q if n >= 0 else -q


def fmt_inr(p: int) -> str:
    sign, p = ("-" if p < 0 else ""), abs(p)
    return f"{sign}{p // 100}.{p % 100:02d}"


def fmt_rate(m: int) -> str:
    s = f"{m // RATE_SCALE}.{m % RATE_SCALE:06d}".rstrip("0")
    return s + "0" * (2 - len(s.split(".")[1])) if len(s.split(".")[1]) < 2 else s


# ---------- model ----------------------------------------------------------------
@dataclass(frozen=True)
class Entry:
    txn_id: str
    entry_date: date
    description: str
    category: str
    currency: str                      # 'INR' or an ISO code
    foreign_minor: int                 # cents; 0 for INR rows
    applied_rate_micro: int
    expected_rate_micro: int
    manual_gross_paise: Optional[int]  # INR rows only (signed); foreign rows derive it
    fees_paise: int
    purpose_code: str
    fira_ref: str
    reverses_txn_id: Optional[str] = None


def gross_paise(e: Entry) -> int:          # workbook col I
    if e.currency == "INR":
        return e.manual_gross_paise or 0
    return div_round(e.foreign_minor * e.applied_rate_micro, RATE_SCALE)


def net_paise(e: Entry) -> int:            # workbook col K
    return gross_paise(e) - e.fees_paise


def fx_variance_paise(e: Entry) -> int:    # workbook col L: (applied - expected) * foreign
    if e.currency == "INR":
        return 0
    return div_round(e.foreign_minor * (e.applied_rate_micro - e.expected_rate_micro), RATE_SCALE)


def validate(e: Entry) -> list[str]:
    errs: list[str] = []
    if not TXN_RE.match(e.txn_id):
        errs.append("txn_id must be 3-40 chars of A-Z a-z 0-9 _ -")
    if not e.description.strip():
        errs.append("description is required")
    if not e.category.strip():
        errs.append("category is required")
    if e.applied_rate_micro <= 0 or e.expected_rate_micro <= 0:
        errs.append("FX rates must be > 0")
    if e.fees_paise < 0:
        errs.append("fees must be >= 0")
    if e.currency == "INR":
        if e.foreign_minor != 0:
            errs.append("INR rows must have foreign amount 0")
        if e.applied_rate_micro != RATE_SCALE or e.expected_rate_micro != RATE_SCALE:
            errs.append("INR rows must use rate 1")
        if not e.manual_gross_paise:
            errs.append("INR rows need a non-zero gross INR amount")
        elif e.manual_gross_paise < 0 and e.fees_paise:
            errs.append("fee on a debit row: the workbook ignores it; post the fee as its own row")
        if (e.purpose_code, e.fira_ref) != ("N/A", "N/A"):
            errs.append("INR rows must use 'N/A' for purpose code and FIRA ref")
    else:
        if e.currency not in SUPPORTED_FOREIGN:
            errs.append(f"unsupported currency {e.currency}")
        if e.foreign_minor <= 0:
            errs.append("inward remittance foreign amount must be > 0")
        if e.manual_gross_paise is not None:
            errs.append("gross INR is derived for foreign rows (foreign amount x applied rate)")
        if e.purpose_code not in PURPOSE_CODES:
            errs.append(f"RBI purpose code must be one of {sorted(PURPOSE_CODES)}")
        if not FIRA_RE.match(e.fira_ref) or e.fira_ref == "N/A":
            errs.append("FIRA/FIRC reference is required for foreign remittances")
        if e.foreign_minor > 0 and e.fees_paise >= gross_paise(e):
            errs.append("fees exceed gross INR received")
    return errs


# ---------- statement + budget ---------------------------------------------------
@dataclass(frozen=True)
class Line:
    seq: int
    entry: Entry
    gross: int
    fees: int
    net: int
    fx_variance: int
    running_balance: int


def build_statement(entries: Iterable[tuple[int, Entry]]) -> list[Line]:
    """Single source of truth for gross/net/FX/running balance. Order = (date, seq)."""
    lines, balance = [], 0
    for seq, e in sorted(entries, key=lambda t: (t[1].entry_date, t[0])):
        n = net_paise(e)
        balance += n
        lines.append(Line(seq, e, gross_paise(e), e.fees_paise, n, fx_variance_paise(e), balance))
    return lines


def budget_vs_actual(lines: list[Line], budgets: dict[str, tuple[str, int]]) -> dict:
    """budgets: {category: (kind 'revenue'|'expense', budget_paise)}.

    actual_settled  = SUMIF of Net Settled (signed, same as the workbook)
    utilised        = revenue earned, or expense spent (positive number)
    remaining       = budget - utilised   (the workbook's  Budget - Actual  is wrong for expenses,
                      because expense nets are negative and the variance gets inflated)
    """
    actual, fx = {}, {}
    for ln in lines:
        actual[ln.entry.category] = actual.get(ln.entry.category, 0) + ln.net
        fx[ln.entry.category] = fx.get(ln.entry.category, 0) + ln.fx_variance
    rows = []
    for cat, (kind, budget) in budgets.items():
        a = actual.pop(cat, 0)
        used = a if kind == "revenue" else -a
        rows.append({"category": cat, "kind": kind, "budget": budget, "actual_settled": a,
                     "utilised": used, "remaining": budget - used, "fx_gain": fx.get(cat, 0)})
    for cat, a in actual.items():  # categories missing from the budget table (never silently dropped)
        rows.append({"category": cat, "kind": "unbudgeted", "budget": 0, "actual_settled": a,
                     "utilised": abs(a), "remaining": -abs(a), "fx_gain": fx.get(cat, 0)})
    closing = lines[-1].running_balance if lines else 0
    total_actual = sum(r["actual_settled"] for r in rows)
    return {"rows": rows, "total_actual_settled": total_actual,
            "total_fx_gain": sum(ln.fx_variance for ln in lines),
            "closing_balance": closing, "ties_to_ledger": total_actual == closing}
