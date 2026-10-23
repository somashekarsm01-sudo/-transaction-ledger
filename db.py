"""SQLite persistence: append-only ledger, hash chain, period snapshots, reconciliation."""
from __future__ import annotations

import hashlib
import sqlite3
from contextlib import contextmanager
from datetime import date

from engine import (Entry, Line, ValidationError, build_statement, budget_vs_actual,
                    net_paise, validate, REVENUE_CATEGORIES)

GENESIS = "GENESIS"

SCHEMA = """
CREATE TABLE IF NOT EXISTS ledger_entries(
  seq INTEGER PRIMARY KEY AUTOINCREMENT,
  txn_id TEXT NOT NULL UNIQUE,
  entry_date TEXT NOT NULL CHECK(entry_date GLOB '[0-9][0-9][0-9][0-9]-[0-9][0-9]-[0-9][0-9]'),
  description TEXT NOT NULL,
  category TEXT NOT NULL,
  currency TEXT NOT NULL CHECK(length(currency)=3),
  foreign_minor INTEGER NOT NULL,
  applied_rate_micro INTEGER NOT NULL CHECK(applied_rate_micro>0),
  expected_rate_micro INTEGER NOT NULL CHECK(expected_rate_micro>0),
  manual_gross_paise INTEGER,
  fees_paise INTEGER NOT NULL DEFAULT 0,
  purpose_code TEXT NOT NULL,
  fira_ref TEXT NOT NULL,
  reverses_txn_id TEXT REFERENCES ledger_entries(txn_id),
  prev_hash TEXT NOT NULL,
  entry_hash TEXT NOT NULL,
  created_at TEXT NOT NULL DEFAULT (strftime('%Y-%m-%dT%H:%M:%SZ','now'))
);
CREATE UNIQUE INDEX IF NOT EXISTS ux_fira ON ledger_entries(fira_ref)
  WHERE reverses_txn_id IS NULL AND fira_ref <> 'N/A';
CREATE UNIQUE INDEX IF NOT EXISTS ux_one_reversal ON ledger_entries(reverses_txn_id)
  WHERE reverses_txn_id IS NOT NULL;
CREATE TRIGGER IF NOT EXISTS ledger_no_update BEFORE UPDATE ON ledger_entries
  BEGIN SELECT RAISE(ABORT, 'ledger is append-only: post a reversal instead'); END;
CREATE TRIGGER IF NOT EXISTS ledger_no_delete BEFORE DELETE ON ledger_entries
  BEGIN SELECT RAISE(ABORT, 'ledger is append-only: post a reversal instead'); END;

CREATE TABLE IF NOT EXISTS budgets(
  category TEXT PRIMARY KEY, kind TEXT NOT NULL CHECK(kind IN ('revenue','expense')),
  budget_paise INTEGER NOT NULL);
CREATE TABLE IF NOT EXISTS entity_profile(key TEXT PRIMARY KEY, value TEXT NOT NULL);
CREATE TABLE IF NOT EXISTS snapshots(
  id INTEGER PRIMARY KEY AUTOINCREMENT, period_end TEXT NOT NULL UNIQUE,
  entry_count INTEGER NOT NULL, closing_paise INTEGER NOT NULL, digest TEXT NOT NULL,
  taken_at TEXT NOT NULL DEFAULT (strftime('%Y-%m-%dT%H:%M:%SZ','now')));
CREATE TRIGGER IF NOT EXISTS snap_no_update BEFORE UPDATE ON snapshots
  BEGIN SELECT RAISE(ABORT, 'snapshots are immutable'); END;
CREATE TRIGGER IF NOT EXISTS snap_no_delete BEFORE DELETE ON snapshots
  BEGIN SELECT RAISE(ABORT, 'snapshots are immutable'); END;
"""


def connect(path: str) -> sqlite3.Connection:
    conn = sqlite3.connect(path, isolation_level=None, timeout=30)  # explicit transactions
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA foreign_keys=ON")
    conn.executescript(SCHEMA)
    return conn


@contextmanager
def tx(conn, commit: bool = True):
    conn.execute("BEGIN IMMEDIATE")  # one writer at a time: head hash cannot fork
    try:
        yield
    except BaseException:
        conn.execute("ROLLBACK")
        raise
    conn.execute("COMMIT" if commit else "ROLLBACK")


# ---------- row <-> entry ----------------------------------------------------------
def row_to_entry(r) -> Entry:
    return Entry(r["txn_id"], date.fromisoformat(r["entry_date"]), r["description"], r["category"],
                 r["currency"], r["foreign_minor"], r["applied_rate_micro"], r["expected_rate_micro"],
                 r["manual_gross_paise"], r["fees_paise"], r["purpose_code"], r["fira_ref"],
                 r["reverses_txn_id"])


def _digest(prev: str, e: Entry) -> str:
    parts = [e.txn_id, e.entry_date.isoformat(), e.description, e.category, e.currency,
             e.foreign_minor, e.applied_rate_micro, e.expected_rate_micro,
             "" if e.manual_gross_paise is None else e.manual_gross_paise, e.fees_paise,
             e.purpose_code, e.fira_ref, e.reverses_txn_id or ""]
    return hashlib.sha256((prev + "|" + "|".join(map(str, parts))).encode()).hexdigest()


# ---------- reads ------------------------------------------------------------------
def budgets(conn) -> dict[str, tuple[str, int]]:
    return {r["category"]: (r["kind"], r["budget_paise"])
            for r in conn.execute("SELECT * FROM budgets ORDER BY rowid")}


def statement(conn, as_of: date | None = None) -> list[Line]:
    rows = conn.execute("SELECT * FROM ledger_entries").fetchall()
    lines = build_statement((r["seq"], row_to_entry(r)) for r in rows)
    return [ln for ln in lines if as_of is None or ln.entry.entry_date <= as_of]


def last_period_end(conn) -> date | None:
    r = conn.execute("SELECT MAX(period_end) m FROM snapshots").fetchone()
    return date.fromisoformat(r["m"]) if r["m"] else None


# ---------- writes (call inside tx) ----------------------------------------------
def set_budgets(conn, items: dict[str, int]) -> None:
    for cat, paise in items.items():
        kind = "revenue" if cat in REVENUE_CATEGORIES else "expense"
        conn.execute("INSERT INTO budgets(category,kind,budget_paise) VALUES(?,?,?) "
                     "ON CONFLICT(category) DO UPDATE SET budget_paise=excluded.budget_paise", (cat, kind, paise))


def post(conn, e: Entry, *, is_reversal: bool = False) -> int:
    known = {c.lower(): c for c in budgets(conn)}
    canon = known.get(e.category.strip().lower())
    if canon is None:
        raise ValidationError(f"{e.txn_id}: category {e.category!r} is not in the budget table "
                              f"{sorted(known.values())}")
    e = Entry(**{**e.__dict__, "category": canon})      # exact-match category => SUMIF can't miss it
    errs = [] if is_reversal else validate(e)
    closed = last_period_end(conn)
    if closed and e.entry_date <= closed:
        errs.append(f"period closed through {closed}; post a dated-after-close adjustment")
    if errs:
        raise ValidationError([f"{e.txn_id}: {m}" for m in errs])
    head = conn.execute("SELECT entry_hash FROM ledger_entries ORDER BY seq DESC LIMIT 1").fetchone()
    prev = head["entry_hash"] if head else GENESIS
    try:
        cur = conn.execute(
            "INSERT INTO ledger_entries(txn_id,entry_date,description,category,currency,foreign_minor,"
            "applied_rate_micro,expected_rate_micro,manual_gross_paise,fees_paise,purpose_code,fira_ref,"
            "reverses_txn_id,prev_hash,entry_hash) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
            (e.txn_id, e.entry_date.isoformat(), e.description, e.category, e.currency, e.foreign_minor,
             e.applied_rate_micro, e.expected_rate_micro, e.manual_gross_paise, e.fees_paise,
             e.purpose_code, e.fira_ref, e.reverses_txn_id, prev, _digest(prev, e)))
    except sqlite3.IntegrityError as ex:
        raise ValidationError(f"{e.txn_id}: rejected by database constraint ({ex})") from None
    return cur.lastrowid


def reverse(conn, txn_id: str, on: date, reason: str) -> int:
    r = conn.execute("SELECT * FROM ledger_entries WHERE txn_id=?", (txn_id,)).fetchone()
    if r is None:
        raise ValidationError(f"{txn_id}: not found")
    o = row_to_entry(r)
    if o.reverses_txn_id:
        raise ValidationError(f"{txn_id}: is itself a reversal")
    rev = Entry(f"{txn_id}-REV", on, f"REVERSAL of {txn_id}: {reason}", o.category, o.currency,
                -o.foreign_minor, o.applied_rate_micro, o.expected_rate_micro,
                None if o.manual_gross_paise is None else -o.manual_gross_paise,
                -o.fees_paise, o.purpose_code, o.fira_ref, txn_id)
    return post(conn, rev, is_reversal=True)


def take_snapshot(conn, period_end: date) -> dict:
    prev = last_period_end(conn)
    if prev and period_end <= prev:
        raise ValidationError(f"period_end must be after {prev}")
    n, closing, dig = _period_fingerprint(conn, period_end)
    conn.execute("INSERT INTO snapshots(period_end,entry_count,closing_paise,digest) VALUES(?,?,?,?)",
                 (period_end.isoformat(), n, closing, dig))
    return {"period_end": period_end.isoformat(), "entry_count": n, "closing_paise": closing}


def _period_fingerprint(conn, period_end: date):
    lines = statement(conn, period_end)
    hashes = {r["txn_id"]: r["entry_hash"] for r in conn.execute("SELECT txn_id, entry_hash FROM ledger_entries")}
    dig = hashlib.sha256("\n".join(hashes[ln.entry.txn_id] for ln in lines).encode()).hexdigest()
    return len(lines), (lines[-1].running_balance if lines else 0), dig


# ---------- drift detection --------------------------------------------------------
def reconcile(conn) -> dict:
    checks, prev = {}, GENESIS
    chain_ok = True
    for r in conn.execute("SELECT * FROM ledger_entries ORDER BY seq"):
        if r["prev_hash"] != prev or r["entry_hash"] != _digest(prev, row_to_entry(r)):
            chain_ok = False
            checks["first_bad_txn"] = r["txn_id"]
            break
        prev = r["entry_hash"]
    checks["hash_chain_intact"] = chain_ok

    lines = statement(conn)
    closing = lines[-1].running_balance if lines else 0
    checks["running_balance_equals_sum_of_nets"] = closing == sum(net_paise(ln.entry) for ln in lines)
    bva = budget_vs_actual(lines, budgets(conn))
    checks["budget_actuals_tie_to_ledger"] = bva["ties_to_ledger"]
    checks["no_unbudgeted_categories"] = all(r["kind"] != "unbudgeted" for r in bva["rows"])

    snap_ok = True
    for s in conn.execute("SELECT * FROM snapshots ORDER BY period_end"):
        n, c, d = _period_fingerprint(conn, date.fromisoformat(s["period_end"]))
        if (n, c, d) != (s["entry_count"], s["closing_paise"], s["digest"]):
            snap_ok = False
    checks["snapshots_match"] = snap_ok
    return {"status": "OK" if all(v for k, v in checks.items() if isinstance(v, bool)) else "DRIFT",
            "closing_paise": closing, "checks": checks}
