"""
===============================================================
  MIS Excel Automation  —  Web App (Flask)
  Cloud-deployable | Login-protected | Mobile-friendly
  v2 — fully in-memory (works on Render free tier)
===============================================================
"""

import os, math, json, io
from datetime import datetime
from flask import (Flask, render_template, request, redirect,
                   url_for, session, send_file, flash, jsonify)
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "mis-secret-key-change-in-prod-2024")

# In-memory stores (no disk needed)
FILE_STORE   = {}   # username -> {"data": bytes, "name": str}
OUTPUT_STORE = {}   # username -> {"data": bytes, "name": str}
ACTIVITY_LOG = []   # list of log dicts (in memory)

# ── Credentials ───────────────────────────────────────────────
USERS = {
    "achal13": {"password": "achal1331", "role": "admin"},
}

# ── Processing helpers ────────────────────────────────────────
IFMT = '[>=10000000]##\\,##\\,##\\,##0;[>=100000]##\\,##\\,##0;##,##0'

def num_to_words_indian(n):
    if n is None or (isinstance(n, float) and math.isnan(n)): return ""
    n = int(round(float(n)))
    if n == 0: return "Zero Rupees Only"
    ones = ['','One','Two','Three','Four','Five','Six','Seven','Eight','Nine',
            'Ten','Eleven','Twelve','Thirteen','Fourteen','Fifteen','Sixteen',
            'Seventeen','Eighteen','Nineteen']
    tens = ['','','Twenty','Thirty','Forty','Fifty','Sixty','Seventy','Eighty','Ninety']
    def td(n): return ones[n] if n < 20 else tens[n//10]+(' '+ones[n%10] if n%10 else '')
    def thd(n): return (ones[n//100]+' Hundred'+(' '+td(n%100) if n%100 else '')) if n>=100 else td(n)
    parts = []
    c = n//10000000; n %= 10000000
    l = n//100000;   n %= 100000
    t = n//1000;     n %= 1000
    if c: parts.append(thd(c)+' Crore')
    if l: parts.append(td(l)+' Lakh')
    if t: parts.append(td(t)+' Thousand')
    if n: parts.append(thd(n))
    return ' '.join(parts)+' Rupees Only'

def format_indian_number(n):
    """Return Indian comma-formatted string e.g. 15,00,000 — used for mail-merge columns."""
    if n is None or (isinstance(n, float) and math.isnan(n)):
        return ""
    n = int(round(float(n)))
    if n == 0:
        return "0"
    is_neg = n < 0
    n = abs(n)
    s = str(n)
    if len(s) <= 3:
        result = s
    else:
        result = s[-3:]
        s = s[:-3]
        while s:
            result = s[-2:] + "," + result
            s = s[:-2]
    return ("-" if is_neg else "") + result

def find_amount_columns(ws):
    AMT  = ['amount','arrears','dpi','principal','overdue','disbursement',
            'outstanding','bounce','penalty','received','interest','award','claim','sold']
    SKIP = ['in word','in words','date','name','no ','number','address',
            'status','remark','reason','url','email']
    found = []
    for cell in ws[1]:
        if not cell.value: continue
        h = str(cell.value).strip(); hl = h.lower()
        if any(s in hl for s in SKIP): continue
        if any(k in hl for k in AMT):
            for r in range(2, min(ws.max_row+1, 6)):
                v = ws.cell(row=r, column=cell.column).value
                if v is not None:
                    try: float(v); found.append((cell.column, h)); break
                    except: pass
    return found

def process_excel_bytes(input_bytes, rules):
    """Process Excel from bytes, return output as bytes. No disk I/O."""
    af  = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    afo = Font(bold=True, color="276221", name="Calibri", size=10)
    wf  = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    wfo = Font(bold=True, color="7F6000", name="Calibri", size=10)
    hf  = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    hfo = Font(bold=True, color="FFFFFF", name="Calibri", size=10)
    ctr = Alignment(horizontal='center', vertical='center', wrap_text=True)

    wb = load_workbook(io.BytesIO(input_bytes))
    for sn in wb.sheetnames:
        ws = wb[sn]
        matched = []
        for cell in ws[1]:
            if not cell.value: continue
            h = str(cell.value).strip()
            if h.lower() in rules:
                matched.append((cell.column, h, rules[h.lower()][0]))
        if not matched: continue
        for ci, cn, mode in sorted(matched, key=lambda x: x[0], reverse=True):
            hc = ws.cell(row=1, column=ci)
            hc.fill = af; hc.font = afo; hc.alignment = ctr
            ws.column_dimensions[hc.column_letter].width = 18
            for r in range(2, ws.max_row+1):
                c = ws.cell(row=r, column=ci)
                if c.value is not None:
                    try:
                        v = float(c.value)
                        if not math.isnan(v): c.number_format = IFMT
                    except: pass
            # ── Mail-merge friendly column (plain text, Indian commas) ──
            ws.insert_cols(ci+1)
            mc = ci+1
            mh = ws.cell(row=1, column=mc)
            mh.value = cn.strip() + " (Formatted)"
            mf_fill = PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid")
            mh.fill = mf_fill
            mh.font = Font(bold=True, color="1F4E79", name="Calibri", size=10)
            mh.alignment = ctr
            ws.column_dimensions[mh.column_letter].width = 20
            for r in range(2, ws.max_row+1):
                ac  = ws.cell(row=r, column=ci)
                mcc = ws.cell(row=r, column=mc)
                if ac.value is not None:
                    try:
                        v = float(ac.value)
                        if not math.isnan(v):
                            mcc.value = format_indian_number(v)
                            mcc.alignment = Alignment(horizontal='right', vertical='center')
                    except:
                        pass

            if mode == "words":
                ws.insert_cols(ci+1); wc = ci+1
                wh = ws.cell(row=1, column=wc)
                wh.value = cn.strip()+" in Word"
                wh.fill = wf; wh.font = wfo; wh.alignment = ctr
                ws.column_dimensions[wh.column_letter].width = 50
                for r in range(2, ws.max_row+1):
                    ac = ws.cell(row=r, column=ci); wcc = ws.cell(row=r, column=wc)
                    if ac.value is not None:
                        try:
                            v = float(ac.value)
                            if not math.isnan(v):
                                wcc.value = num_to_words_indian(v)
                                wcc.alignment = Alignment(wrap_text=True, vertical='center')
                        except: pass
            elif mode == "fill":
                wc = None
                for off in range(1, 4):
                    nc = ws.cell(row=1, column=ci+off)
                    if nc.value and any(k in str(nc.value).lower() for k in ['in word','in words']):
                        wc = ci+off; break
                if wc:
                    wh2 = ws.cell(row=1, column=wc)
                    wh2.fill = wf; wh2.font = wfo; wh2.alignment = ctr
                    ws.column_dimensions[wh2.column_letter].width = 50
                    for r in range(2, ws.max_row+1):
                        ac = ws.cell(row=r, column=ci); wcc = ws.cell(row=r, column=wc)
                        if ac.value is not None:
                            try:
                                v = float(ac.value)
                                if not math.isnan(v):
                                    wcc.value = num_to_words_indian(v)
                                    wcc.alignment = Alignment(wrap_text=True, vertical='center')
                            except: pass
        for cell in ws[1]:
            if cell.value:
                rgb = cell.fill.fgColor.rgb if cell.fill.fgColor.type=='rgb' else '00000000'
                if rgb in ('00000000','FFFFFFFF',''):
                    cell.fill = hf; cell.font = hfo; cell.alignment = ctr
        ws.row_dimensions[1].height = 40

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out.read()

# ── Activity log (in-memory) ──────────────────────────────────
def log_activity(user, action, detail=""):
    ACTIVITY_LOG.insert(0, {
        "time":   datetime.now().strftime("%d %b %Y  %H:%M:%S"),
        "user":   user,
        "action": action,
        "detail": detail
    })
    if len(ACTIVITY_LOG) > 200:
        ACTIVITY_LOG.pop()

# ── Auth helpers ──────────────────────────────────────────────
def current_user():
    return session.get("username")

def require_login(f):
    from functools import wraps
    @wraps(f)
    def decorated(*args, **kwargs):
        if not current_user():
            return redirect(url_for("login"))
        return f(*args, **kwargs)
    return decorated

def require_admin(f):
    from functools import wraps
    @wraps(f)
    def decorated(*args, **kwargs):
        u = current_user()
        if not u or USERS.get(u, {}).get("role") != "admin":
            return redirect(url_for("dashboard"))
        return f(*args, **kwargs)
    return decorated

# ══════════════════════════════════════════════════════════════
# ROUTES
# ══════════════════════════════════════════════════════════════

@app.route("/", methods=["GET","POST"])
def login():
    if current_user():
        return redirect(url_for("dashboard"))
    error = None
    if request.method == "POST":
        u = request.form.get("username","").strip()
        p = request.form.get("password","")
        user = USERS.get(u)
        if user and user["password"] == p:
            session["username"] = u
            session["role"]     = user["role"]
            log_activity(u, "Login", "Successful login")
            return redirect(url_for("dashboard"))
        else:
            error = "Invalid username or password."
            log_activity(u or "unknown", "Login Failed", "Bad credentials")
    return render_template("login.html", error=error)

@app.route("/logout")
def logout():
    u = current_user()
    if u:
        log_activity(u, "Logout", "")
        FILE_STORE.pop(u, None)
        OUTPUT_STORE.pop(u, None)
    session.clear()
    return redirect(url_for("login"))

@app.route("/dashboard")
@require_login
def dashboard():
    return render_template("dashboard.html",
                           username=current_user(),
                           role=session.get("role","user"))

@app.route("/detect-columns", methods=["POST"])
@require_login
def detect_columns():
    f = request.files.get("file")
    if not f or not f.filename.lower().endswith((".xlsx",".xls")):
        return jsonify({"error": "Please upload a valid .xlsx or .xls file."}), 400

    file_bytes = f.read()
    fname      = f.filename
    FILE_STORE[current_user()] = {"data": file_bytes, "name": fname}

    try:
        wb   = load_workbook(io.BytesIO(file_bytes), read_only=True, data_only=True)
        ws   = wb.active
        cols = find_amount_columns(ws)
        wb.close()
        return jsonify({"columns": [{"col": c, "name": n} for c, n in cols]})
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route("/process", methods=["POST"])
@require_login
def process():
    data  = request.json
    rules = data.get("rules", {})
    u     = current_user()
    store = FILE_STORE.get(u)

    if not store:
        return jsonify({"error": "No file found. Please upload your file again."}), 400
    if not rules:
        return jsonify({"error": "No columns selected."}), 400

    fname     = store["name"]
    name, ext = os.path.splitext(fname)
    out_name  = name + "_UPDATED" + ext

    try:
        out_bytes = process_excel_bytes(store["data"], rules)
        OUTPUT_STORE[u] = {"data": out_bytes, "name": out_name}
        log_activity(u, "Processed File", fname)
        return jsonify({"success": True, "filename": out_name})
    except Exception as e:
        log_activity(u, "Process Error", str(e))
        return jsonify({"error": str(e)}), 500

@app.route("/download")
@require_login
def download():
    u     = current_user()
    store = OUTPUT_STORE.get(u)
    if not store:
        flash("No processed file found. Please process a file first.")
        return redirect(url_for("dashboard"))
    log_activity(u, "Downloaded File", store["name"])
    return send_file(
        io.BytesIO(store["data"]),
        as_attachment=True,
        download_name=store["name"],
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

@app.route("/admin")
@require_login
@require_admin
def admin():
    users_info = [{"username": u, "role": d["role"]} for u, d in USERS.items()]
    return render_template("admin.html",
                           username=current_user(),
                           logs=ACTIVITY_LOG,
                           users=users_info)

@app.route("/admin/add-user", methods=["POST"])
@require_login
@require_admin
def add_user():
    u = request.form.get("username","").strip()
    p = request.form.get("password","").strip()
    r = request.form.get("role","user")
    if u and p and u not in USERS:
        USERS[u] = {"password": p, "role": r}
        log_activity(current_user(), "Added User", u)
        flash(f"User '{u}' added successfully.")
    else:
        flash("Username already exists or invalid input.")
    return redirect(url_for("admin"))

@app.route("/admin/remove-user/<username>")
@require_login
@require_admin
def remove_user(username):
    if username == current_user():
        flash("You cannot remove yourself.")
    elif username in USERS:
        del USERS[username]
        log_activity(current_user(), "Removed User", username)
        flash(f"User '{username}' removed.")
    return redirect(url_for("admin"))

@app.route("/admin/clear-logs")
@require_login
@require_admin
def clear_logs():
    ACTIVITY_LOG.clear()
    log_activity(current_user(), "Cleared Logs", "")
    return redirect(url_for("admin"))

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
