"""Flask API. Run:  VELOCE_API_TOKEN=<long-random> gunicorn -w 1 'app:create_app()'  (behind HTTPS)."""
from __future__ import annotations

import hmac
import os
from datetime import date

from flask import Flask, Response, g, jsonify, request

import db
from engine import (Entry, ValidationError, budget_vs_actual, fmt_inr, fmt_rate, to_micro, to_minor)
from importer import import_workbook


PAGE = r"""<!DOCTYPE html>
<html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Veloce Ledger</title>
<style>
body{font-family:system-ui,Arial,sans-serif;margin:0;background:#f4f6f8;color:#0f172a}
header{background:#0f172a;color:#fff;padding:14px 20px;display:flex;justify-content:space-between;align-items:center}
main{max-width:1100px;margin:0 auto;padding:16px}
nav button{margin-right:6px;padding:8px 14px;border:0;border-radius:8px;background:#e2e8f0;cursor:pointer}
nav button.on{background:#10b981;color:#fff}
.card{background:#fff;border-radius:12px;padding:16px;margin-top:12px;box-shadow:0 1px 3px #0002}
table{border-collapse:collapse;width:100%;font-size:13px}th,td{padding:6px 8px;border-bottom:1px solid #e2e8f0;text-align:left;white-space:nowrap}
.sc{overflow-x:auto}input,select{padding:8px;margin:4px 0;width:100%;box-sizing:border-box;border:1px solid #cbd5e1;border-radius:6px}
.g{display:grid;grid-template-columns:repeat(auto-fit,minmax(180px,1fr));gap:10px}
.go{padding:10px 18px;border:0;border-radius:8px;background:#0f172a;color:#fff;cursor:pointer;margin-top:8px}
.msg{margin-top:10px;padding:10px;border-radius:8px;background:#fef3c7;white-space:pre-wrap}.ok{background:#d1fae5}
.sm{padding:3px 8px;font-size:12px;border:1px solid #cbd5e1;background:#fff;border-radius:6px;cursor:pointer}
</style></head><body>
<header><b>Veloce Ledger</b><span id="who"></span></header>
<main>
<div class="card" id="login"><b>Enter your password (VELOCE_API_TOKEN)</b>
<input id="tok" type="password" placeholder="password"><button class="go" onclick="login()">Open</button>
<div id="lmsg"></div></div>
<div id="app" style="display:none">
<nav><button data-t="ledger" class="on">Ledger</button><button data-t="add">Add entry</button><button data-t="imp">Import Excel</button><button data-t="bva">Budget vs actual</button></nav>
<div class="card" id="ledger"><button class="go" onclick="recon()" style="margin:0 0 10px">Check balance</button><div id="rmsg"></div><div class="sc" id="ltable">Loading…</div></div>
<div class="card" id="add" style="display:none"><div class="g">
<label>Transaction ID<input id="txn_id"></label><label>Date<input id="date" type="date"></label>
<label>Description<input id="description"></label><label>Category<input id="category"></label>
<label>Currency<input id="currency" value="INR"></label><label>Foreign amount<input id="foreign_amount" value="0"></label>
<label>Applied rate<input id="applied_rate" value="1"></label><label>Expected rate<input id="expected_rate" value="1"></label>
<label>Gross INR (INR entries)<input id="gross_inr"></label><label>Fees INR<input id="fees_inr" value="0"></label>
<label>Purpose code<input id="purpose_code" value="N/A"></label><label>FIRA ref<input id="fira_ref" value="N/A"></label></div>
<button class="go" onclick="addEntry()">Save entry</button><div id="amsg"></div></div>
<div class="card" id="imp" style="display:none"><input type="file" id="file" accept=".xlsx">
<button class="go" onclick="imp(true)">Test only (no saving)</button> <button class="go" onclick="imp(false)">Import for real</button><div id="imsg"></div></div>
<div class="card" id="bva" style="display:none"><div class="sc" id="btable">Loading…</div></div>
</div></main>
<script>
let T=sessionStorage.getItem('tok')||'';
const $=id=>document.getElementById(id);
async function api(path,opt={}){opt.headers=Object.assign({Authorization:'Bearer '+T},opt.headers||{});
 const r=await fetch(path,opt);let d;try{d=await r.json()}catch(e){d={}}
 if(r.status==401){sessionStorage.removeItem('tok');T='';show();throw new Error('Wrong password')}
 return {s:r.status,d}}
function esc(x){return String(x??'').replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]))}
function tbl(rows,extra){if(!rows.length)return'No data yet.';const k=Object.keys(rows[0]);
 return '<table><tr>'+k.map(x=>'<th>'+esc(x)+'</th>').join('')+(extra?'<th></th>':'')+'</tr>'+
 rows.map(r=>'<tr>'+k.map(x=>'<td>'+esc(r[x])+'</td>').join('')+(extra?'<td>'+extra(r)+'</td>':'')+'</tr>').join('')+'</table>'}
function show(){const ok=!!T;$('login').style.display=ok?'none':'';$('app').style.display=ok?'':'none';$('who').innerHTML=ok?'<a href="#" style="color:#fff" onclick="sessionStorage.clear();location.reload()">Log out</a>':'';if(ok)loadLedger()}
async function login(){T=$('tok').value.trim();const r=await api('/api/profile').catch(e=>null);
 if(r&&r.s==200){sessionStorage.setItem('tok',T);show()}else{T='';$('lmsg').textContent='Wrong password, or server not ready.'}}
async function loadLedger(){try{const r=await api('/api/ledger');$('ltable').innerHTML=tbl(r.d,x=>x.reverses?'':'<button class="sm" onclick="rev(\''+esc(x.txn_id)+'\')">Reverse</button>')}catch(e){$('ltable').textContent=e.message}}
async function recon(){const r=await api('/api/reconcile');$('rmsg').innerHTML='<div class="msg '+(r.s==200?'ok':'')+'">'+esc(JSON.stringify(r.d))+'</div>'}
async function rev(id){const reason=prompt('Reason for reversing '+id+'?','correction');if(reason===null)return;
 const r=await api('/api/entries/'+encodeURIComponent(id)+'/reverse',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({date:new Date().toISOString().slice(0,10),reason})});
 if(r.s>=400)alert(JSON.stringify(r.d));loadLedger()}
async function addEntry(){const f=['txn_id','date','description','category','currency','foreign_amount','applied_rate','expected_rate','gross_inr','fees_inr','purpose_code','fira_ref'];
 const b={};f.forEach(k=>{const v=$(k).value.trim();if(v!=='')b[k]=v});
 const r=await api('/api/entries',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(b)});
 $('amsg').innerHTML='<div class="msg '+(r.s==201?'ok':'')+'">'+(r.s==201?'Saved.':esc(JSON.stringify(r.d)))+'</div>';if(r.s==201)loadLedger()}
async function imp(dry){const f=$('file').files[0];if(!f){$('imsg').textContent='Choose an .xlsx file first.';return}
 const fd=new FormData();fd.append('file',f);const r=await api('/api/import'+(dry?'?dry_run=1':''),{method:'POST',body:fd});
 $('imsg').innerHTML='<div class="msg '+(r.s==200?'ok':'')+'">'+esc(JSON.stringify(r.d,null,2))+'</div>';if(r.s==200&&!dry)loadLedger()}
async function loadB(){const r=await api('/api/budget-vs-actual');$('btable').innerHTML=tbl(r.d.rows||[])}
document.querySelectorAll('nav button').forEach(b=>b.onclick=()=>{document.querySelectorAll('nav button').forEach(x=>x.classList.remove('on'));b.classList.add('on');
 ['ledger','add','imp','bva'].forEach(t=>$(t).style.display=t==b.dataset.t?'':'none');if(b.dataset.t=='bva')loadB();if(b.dataset.t=='ledger')loadLedger()});
$('tok').addEventListener('keydown',e=>{if(e.key=='Enter')login()});show();
</script></body></html>
"""


def create_app(db_path: str | None = None) -> Flask:
    app = Flask(__name__)
    app.config["MAX_CONTENT_LENGTH"] = 5 * 1024 * 1024
    path = db_path or os.environ.get("VELOCE_DB", "veloce.db")
    db.connect(path).close()  # create schema once

    @app.get("/")
    def index():
        return Response(PAGE, mimetype="text/html")

    @app.before_request
    def _auth():
        if request.path == "/":  # the screen is public; all data still needs the password
            return None
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
