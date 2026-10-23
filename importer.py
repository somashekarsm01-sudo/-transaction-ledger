"""Parse Veloce_Commerce_Operations_Ledger.xlsx (exact tab/column schema) into the database.

Only INPUT columns are trusted. Derived columns (I for foreign rows, K, L, O) are recomputed by the
engine; a hard-coded number in a derived column is reported as a warning, not imported.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date, datetime

from openpyxl import load_workbook

import db
from engine import (Entry, IFSC_RE, SWIFT_RE, ValidationError, fx_variance_paise, gross_paise,
                    net_paise, to_micro, to_minor)

LEDGER_COLS = ["Date", "Transaction ID", "Description", "Category", "Invoiced Currency", "Foreign Amount",
               "Applied FX Rate", "Expected Rate (Budget)", "Gross INR Received",
               "Bank Fees / Deductions (INR)", "Net Settled INR", "FX Variance (INR Gain/Loss)",
               "RBI Purpose Code", "FIRA / FIRC Ref", "Running Balance (INR)"]


@dataclass
class Parsed:
    profile: dict = field(default_factory=dict)
    budgets: dict = field(default_factory=dict)
    entries: list = field(default_factory=list)
    errors: list = field(default_factory=list)
    warnings: list = field(default_factory=list)


def _formula(v) -> bool:
    return isinstance(v, str) and v.startswith("=")


def _date(v) -> date:
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    return date.fromisoformat(str(v).strip())


def _table(ws, required):
    rows = ws.iter_rows(values_only=True)
    header = [str(h).strip() if h is not None else "" for h in next(rows, [])]
    missing = [c for c in required if c not in header]
    return header, missing, rows


def parse_workbook(src) -> Parsed:
    p = Parsed()
    wb = load_workbook(src, data_only=False)  # formulas visible, so hard-codes can be told apart
    for name in ("Banking Profile", "Remittance Ledger", "Budget vs Actual"):
        if name not in wb.sheetnames:
            p.errors.append(f"missing sheet: {name}")
    if p.errors:
        return p

    for k, v in list(wb["Banking Profile"].iter_rows(min_row=2, values_only=True)):
        if k and v is not None:
            p.profile[str(k).strip()] = str(v).strip()
    if "IFSC Central Identifier" in p.profile and not IFSC_RE.match(p.profile["IFSC Central Identifier"]):
        p.errors.append("Banking Profile: IFSC format invalid")
    if "SWIFT Network Address" in p.profile and not SWIFT_RE.match(p.profile["SWIFT Network Address"]):
        p.errors.append("Banking Profile: SWIFT/BIC format invalid")

    header, missing, rows = _table(wb["Budget vs Actual"], ["Category", "Budgeted INR"])
    if missing:
        p.errors.append(f"Budget vs Actual: missing columns {missing}")
    else:
        ci, bi = header.index("Category"), header.index("Budgeted INR")
        for n, r in enumerate(rows, start=2):
            if r[ci] and isinstance(r[bi], (int, float)):
                p.budgets[str(r[ci]).strip()] = to_minor(r[bi])

    header, missing, rows = _table(wb["Remittance Ledger"], LEDGER_COLS)
    if missing:
        p.errors.append(f"Remittance Ledger: missing columns {missing}")
        return p
    ix = {c: header.index(c) for c in LEDGER_COLS}
    for n, r in enumerate(rows, start=2):
        g = lambda c: r[ix[c]]
        if not g("Transaction ID"):
            continue  # pre-filled formula rows with no transaction
        try:
            cur = str(g("Invoiced Currency")).strip().upper()
            gross_cell = g("Gross INR Received")
            manual = to_minor(gross_cell) if cur == "INR" and gross_cell is not None and not _formula(gross_cell) else None
            if cur == "INR" and manual is None:
                raise ValueError("INR row needs a hard-coded Gross INR amount")
            e = Entry(str(g("Transaction ID")).strip(), _date(g("Date")), str(g("Description") or "").strip(),
                      str(g("Category") or "").strip(), cur, to_minor(g("Foreign Amount") or 0),
                      to_micro(g("Applied FX Rate")), to_micro(g("Expected Rate (Budget)")), manual,
                      to_minor(g("Bank Fees / Deductions (INR)") or 0),
                      str(g("RBI Purpose Code") or "").strip(), str(g("FIRA / FIRC Ref") or "").strip())
        except (ValueError, TypeError) as ex:
            p.errors.append(f"Remittance Ledger row {n}: {ex}")
            continue
        derived = {"Net Settled INR": net_paise(e), "FX Variance (INR Gain/Loss)": fx_variance_paise(e)}
        if cur != "INR":
            derived["Gross INR Received"] = gross_paise(e)
        for col, want in derived.items():
            v = g(col)
            if v is not None and not _formula(v):
                try:
                    if to_minor(v) != want:
                        p.warnings.append(f"row {n} {e.txn_id}: hard-coded '{col}' {v} != computed "
                                          f"{want/100:.2f}; engine value used")
                except ValueError:
                    pass
        p.entries.append(e)
    return p


def import_workbook(conn, src, dry_run: bool = False) -> dict:
    parsed = parse_workbook(src)
    report = {"dry_run": dry_run, "posted": [], "skipped": [], "errors": list(parsed.errors),
              "warnings": parsed.warnings}
    if parsed.errors:
        return report
    try:
        with db.tx(conn, commit=not dry_run):
            db.set_budgets(conn, parsed.budgets)
            for k, v in parsed.profile.items():
                conn.execute("INSERT INTO entity_profile(key,value) VALUES(?,?) "
                             "ON CONFLICT(key) DO UPDATE SET value=excluded.value", (k, v))
            for e in parsed.entries:
                row = conn.execute("SELECT * FROM ledger_entries WHERE txn_id=?", (e.txn_id,)).fetchone()
                if row is None:
                    db.post(conn, e)
                    report["posted"].append(e.txn_id)
                    continue
                existing = db.row_to_entry(row)
                if existing == Entry(**{**e.__dict__, "category": existing.category}):
                    report["skipped"].append(e.txn_id)          # idempotent re-import
                else:
                    raise ValidationError(f"{e.txn_id}: already posted with different values; "
                                          "ledger is append-only - post a reversal and a new entry")
    except ValidationError as ex:
        report["errors"] += ex.errors
        report["posted"] = []
    report["closing_paise"] = (db.statement(conn)[-1].running_balance if db.statement(conn) else 0)
    return report
