"""Flask API. Run:  VELOCE_API_TOKEN=<long-random> gunicorn -w 1 'app:create_app()'  (behind HTTPS)."""
from __future__ import annotations

import hmac
import os
from datetime import date

from flask import Flask, g, jsonify, request

import db
from engine import (Entry, ValidationError, budget_vs_actual, fmt_inr, fmt_rate, to_micro, to_minor)
from importer import import_workbook


def create_app(db_path: str | None = None) -> Flask:
    app = Flask(__name__)
    app.config["MAX_CONTENT_LENGTH"] = 5 * 1024 * 1024
    path = db_path or os.environ.get("VELOCE_DB", "veloce.db")
    db.connect(path).close()  # create schema once

    @app.before_request
    def _auth():
        token = os.environ.get("VELOCE_API_TOKEN")
        if not token:
            return jsonify(error="server has no VELOCE_API_TOKEN configured"), 503  # fail closed
        given = request.headers.get("Authorization", "").removeprefix("Bearer ").strip()
        if not hmac.compare_digest(given, token):
            return jsonify(error="unauthorised"), 401
        g.conn = db.connect(path)

    @app.teardown_request
    def _close(_):
        c = g.pop("conn", None)
        if c:
            c.close()

    @app.errorhandler(ValidationError)
    def _bad(ex):
        return jsonify(errors=ex.errors), 422

    def line_json(ln):
        e = ln.entry
        return {"seq": ln.seq, "date": e.entry_date.isoformat(), "txn_id": e.txn_id, "description": e.description,
                "category": e.category, "currency": e.currency, "foreign_amount": fmt_inr(e.foreign_minor),
                "applied_rate": fmt_rate(e.applied_rate_micro), "expected_rate": fmt_rate(e.expected_rate_micro),
                "gross_inr": fmt_inr(ln.gross), "fees_inr": fmt_inr(ln.fees), "net_settled_inr": fmt_inr(ln.net),
                "fx_variance_inr": fmt_inr(ln.fx_variance), "purpose_code": e.purpose_code,
                "fira_ref": e.fira_ref, "running_balance_inr": fmt_inr(ln.running_balance),
                "reverses": e.reverses_txn_id}

    @app.get("/api/profile")
    def profile():
        p = {r["key"]: r["value"] for r in g.conn.execute("SELECT * FROM entity_profile")}
        for k, v in p.items():
            if "Account" in k and v.isdigit():
                p[k] = "*" * (len(v) - 4) + v[-4:]      # never return the full account number
        return jsonify(p)

    @app.post("/api/import")
    def import_xlsx():
        f = request.files.get("file")
        if not f or not f.filename.lower().endswith(".xlsx"):
            return jsonify(error="upload an .xlsx file as form field 'file'"), 400
        rep = import_workbook(g.conn, f.stream, dry_run=request.args.get("dry_run") == "1")
        rep["closing_balance_inr"] = fmt_inr(rep.pop("closing_paise"))
        return jsonify(rep), (422 if rep["errors"] else 200)

    @app.get("/api/ledger")
    def ledger():
        return jsonify([line_json(ln) for ln in db.statement(g.conn)])

    @app.post("/api/entries")
    def add_entry():
        b = request.get_json(force=True)
        try:
            cur = str(b["currency"]).upper()
            e = Entry(b["txn_id"], date.fromisoformat(b["date"]), b["description"], b["category"], cur,
                      to_minor(b.get("foreign_amount", 0)), to_micro(b.get("applied_rate", 1)),
                      to_micro(b.get("expected_rate", 1)),
                      to_minor(b["gross_inr"]) if cur == "INR" else None, to_minor(b.get("fees_inr", 0)),
                      b.get("purpose_code", "N/A"), b.get("fira_ref", "N/A"))
        except (KeyError, ValueError, TypeError) as ex:
            raise ValidationError(f"bad payload: {ex}") from None
        with db.tx(g.conn):
            seq = db.post(g.conn, e)
        ln = next(x for x in db.statement(g.conn) if x.seq == seq)
        return jsonify(line_json(ln)), 201

    @app.post("/api/entries/<txn_id>/reverse")
    def reverse_entry(txn_id):
        b = request.get_json(force=True)
        with db.tx(g.conn):
            seq = db.reverse(g.conn, txn_id, date.fromisoformat(b["date"]), b.get("reason", "correction"))
        return jsonify(line_json(next(x for x in db.statement(g.conn) if x.seq == seq))), 201

    @app.get("/api/budget-vs-actual")
    def bva():
        r = budget_vs_actual(db.statement(g.conn), db.budgets(g.conn))
        out = {k: fmt_inr(v) if isinstance(v, int) and not isinstance(v, bool) else v
               for k, v in r.items() if k != "rows"}
        out["rows"] = [{k: (fmt_inr(v) if isinstance(v, int) and k not in ("category", "kind") else v)
                        for k, v in row.items()} for row in r["rows"]]
        return jsonify(out)

    @app.post("/api/snapshots")
    def snapshot():
        with db.tx(g.conn):
            s = db.take_snapshot(g.conn, date.fromisoformat(request.get_json(force=True)["period_end"]))
        s["closing_inr"] = fmt_inr(s.pop("closing_paise"))
        return jsonify(s), 201

    @app.get("/api/reconcile")
    def reconcile():
        r = db.reconcile(g.conn)
        r["closing_inr"] = fmt_inr(r.pop("closing_paise"))
        return jsonify(r), (200 if r["status"] == "OK" else 409)

    return app
