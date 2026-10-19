"""
Daily Sales vs Payments Reconciliation - Flask web app (runs on Vercel or locally).
Local:  python app.py     (opens on http://localhost:5000)
Vercel: exposes `app`; set APP_PASSWORD (and optionally SECRET_KEY) in project env vars.
"""
import base64
import csv
import re
import hmac
import hashlib
import io
import os
import tempfile
from datetime import timedelta

from flask import Flask, jsonify, request, send_from_directory, session

from pipeline import run_reconciliation

BASE = os.path.dirname(os.path.abspath(__file__))
app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 4 * 1024 * 1024  # Vercel functions reject bodies > ~4.5 MB anyway
app.config["SESSION_COOKIE_HTTPONLY"] = True
app.config["SESSION_COOKIE_SAMESITE"] = "Lax"
app.config["SESSION_COOKIE_SECURE"] = bool(os.environ.get("VERCEL"))
app.permanent_session_lifetime = timedelta(hours=12)

PASSWORD = os.environ.get("APP_PASSWORD", "")
# Stable across serverless instances: explicit SECRET_KEY, else derived from the password.
app.secret_key = os.environ.get("SECRET_KEY") or hashlib.sha256(("recon-v1|" + PASSWORD).encode()).hexdigest()
LOCKED = bool(os.environ.get("VERCEL")) and not PASSWORD  # hosted without a password -> refuse


def _authed():
    return (not PASSWORD and not LOCKED) or session.get("ok") is True


@app.get("/")
def index():
    return send_from_directory(os.path.join(BASE, "web"), "index.html")


@app.get("/api/status")
def status():
    return jsonify(locked=LOCKED, needs_login=bool(PASSWORD), authed=_authed())


@app.post("/api/login")
def login():
    if LOCKED:
        return jsonify(error="Server has no APP_PASSWORD configured."), 503
    entered = (request.get_json(silent=True) or {}).get("password", "")
    if PASSWORD and hmac.compare_digest(str(entered).encode(), PASSWORD.encode()):
        session.permanent = True
        session["ok"] = True
        return jsonify(ok=True)
    return jsonify(error="Wrong password."), 401


@app.post("/api/logout")
def logout():
    session.clear()
    return jsonify(ok=True)


def _save(storage, folder):
    # never trust client file names
    name = os.path.basename(storage.filename or "upload").replace("..", "_") or "upload"
    path = os.path.join(folder, f"{len(os.listdir(folder))}_{name}")
    storage.save(path)
    return path


def _text_csv(text, folder, name):
    """Typed rows 'time,amount,ref' (header optional) -> csv file, or None if empty."""
    lines = [l.strip() for l in (text or "").splitlines() if l.strip()]
    if not lines:
        return None
    if not lines[0].lower().startswith("time"):
        lines.insert(0, "time,amount,ref")
    path = os.path.join(folder, name)
    with open(path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines) + "\n")
    return path


_AMOUNT_ONLY = re.compile(r"^[\d.,\s₹]*(?:rs\.?)?[\d.,\s]*$", re.I)


def _machine_csv(text, folder, name):
    """Card-slip entry. Easiest form: just amounts (one per line or space-separated, '1,200' ok).
    Optional richer lines 'time,amount,ref' (e.g. 17:03,620,097776) still work; time/ref may be blank."""
    rows = []
    for line in (text or "").splitlines():
        line = line.strip()
        if not line or line.lower().startswith("time"):
            continue
        if _AMOUNT_ONLY.match(line):
            for tok in line.replace("₹", " ").replace("Rs.", " ").replace("rs.", " ").split():
                tok = tok.replace(",", "")
                if tok:
                    rows.append(("", tok, ""))
        else:
            parts = [p.strip() for p in line.split(",")]
            if ":" in parts[0] and len(parts) >= 2:
                rows.append((parts[0], parts[1], ",".join(parts[2:])))
            elif len(parts) >= 1:
                rows.append(("", parts[0], ",".join(parts[1:])))
    if not rows:
        return None
    path = os.path.join(folder, name)
    with open(path, "w", encoding="utf-8", newline="") as f:
        w = csv.writer(f)
        w.writerow(["time", "amount", "ref"])
        w.writerows(rows)
    return path


@app.post("/api/run")
def run():
    if not _authed():
        return jsonify(error="Please log in."), 401
    form, files = request.form, request.files
    shops = [s.strip() for s in form.get("shops", "").split(",") if s.strip()]
    if not shops:
        return jsonify(error="Enter at least one shop name."), 400
    date_label = form.get("date", "")
    try:
        with tempfile.TemporaryDirectory() as tmp:
            pos = {}
            for i, shop in enumerate(shops):
                f = files.get(f"pos_{i}")
                if not f or not f.filename:
                    return jsonify(error=f"Upload the billing export for {shop}."), 400
                pos[shop] = _save(f, tmp)
            banks = []
            for j in range(min(int(form.get("n_banks", 0) or 0), 6)):
                label = (form.get(f"blabel_{j}") or f"Federal Bank - Account {j + 1}").strip()
                hint = (form.get(f"bhint_{j}") or "").strip()
                for k, f in enumerate(f for f in files.getlist(f"bank_{j}") if f.filename):
                    banks.append((label if k == 0 else f"{label} ({k + 1})", _save(f, tmp), hint if hint in shops else None))
            pp_up = [f for f in files.getlist("phonepe") if f.filename]
            pps = [("PhonePe" if len(pp_up) == 1 else f"PhonePe ({j + 1})", _save(f, tmp)) for j, f in enumerate(pp_up)]
            cms = []
            for j in range(min(int(form.get("n_machines", 0) or 0), 8)):
                label = (form.get(f"mlabel_{j}") or f"Card Machine {j + 1}").strip()
                hint = (form.get(f"mhint_{j}") or "").strip()
                up = files.get(f"mcsv_{j}")
                path = _save(up, tmp) if up and up.filename else _machine_csv(form.get(f"mtext_{j}"), tmp, f"m{j}.csv")
                if path:
                    cms.append((label, path, hint if hint in shops else None))
            trs = []
            tp = _machine_csv(form.get("transfers"), tmp, "transfers.csv")
            if tp:
                trs.append(("Direct Transfer", tp))
            if not (banks or pps or cms or trs):
                return jsonify(error="Add at least one payment source."), 400
            out = os.path.join(tmp, "report.xlsx")
            s = run_reconciliation(date_label, pos, banks, pps, cms, trs, out, log=lambda x: print('[reconcile]', x, flush=True))
            with open(out, "rb") as fh:
                xlsx = base64.b64encode(fh.read()).decode()
    except Exception as e:
        return jsonify(error=f"Could not complete the reconciliation: {e}"), 400
    s["unassigned"] = [dict(amount=p.amount, source=p.source, ref=p.ref, time=p.time_str) for p in s["unassigned"]]
    s["xlsx"] = xlsx
    return jsonify(s)


@app.errorhandler(413)
def too_big(_):
    return jsonify(error="Upload too large (limit about 4 MB in total)."), 413


if __name__ == "__main__":
    # 127.0.0.1 (not "localhost"): on Windows, localhost tries IPv6 first and each page load can stall.
    import threading, webbrowser
    threading.Timer(1.0, lambda: webbrowser.open("http://127.0.0.1:5000")).start()
    print("Open http://127.0.0.1:5000  (Ctrl+C to stop)")
    app.run(host="127.0.0.1", port=5000, debug=False, threaded=True)
