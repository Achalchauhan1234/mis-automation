"""
===============================================================
  MIS Excel Automation  —  Web App (Flask)
  Cloud-deployable | Login-protected | Mobile-friendly
===============================================================
  INSTALL:
      pip install flask openpyxl
  RUN LOCALLY:
      python app.py
  DEPLOY:
      Push to GitHub → connect to Render.com (free tier)
===============================================================
"""

import os, math, shutil, json
from datetime import datetime
from flask import (Flask, render_template, request, redirect,
                   url_for, session, send_file, flash, jsonify)
from werkzeug.utils import secure_filename
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "mis-secret-key-change-in-prod-2024")

UPLOAD_FOLDER = os.path.join(os.path.dirname(__file__), "uploads")
OUTPUT_FOLDER = os.path.join(os.path.dirname(__file__), "outputs")
LOG_FILE      = os.path.join(os.path.dirname(__file__), "activity.json")
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

# ── Credentials ───────────────────────────────────────────────
# Add more users here: "username": {"password": "xxx", "role": "user"/"admin"}
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
    parts=[]
    c=n//10000000; n%=10000000
    l=n//100000;   n%=100000
    t=n//1000;     n%=1000
    if c: parts.append(thd(c)+' Crore')
    if l: parts.append(td(l)+' Lakh')
    if t: parts.append(td(t)+' Thousand')
    if n: parts.append(thd(n))
    return ' '.join(parts)+' Rupees Only'

def find_amount_columns(ws):
    AMT  = ['amount','arrears','dpi','principal','overdue','disbursement',
            'outstanding','bounce','penalty','received','interest','award','claim','sold']
    SKIP = ['in word','in words','date','name','no ','number','address',
            'status','remark','reason','url','email']
    found=[]
    for cell in ws[1]:
        if not cell.value: continue
        h=str(cell.value).strip(); hl=h.lower()
        if any(s in hl for s in SKIP): continue
        if any(k in hl for k in AMT):
            for r in range(2, min(ws.max_row+1, 6)):
                v=ws.cell(row=r, column=cell.column).value
                if v is not None:
                    try: float(v); found.append((cell.column, h)); break
                    except: pass
    return found

def process_excel(input_path, output_path, rules):
    af  = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    afo = Font(bold=True, color="276221", name="Calibri", size=10)
    wf  = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    wfo = Font(bold=True, color="7F6000", name="Calibri", size=10)
    hf  = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    hfo = Font(bold=True, color="FFFFFF", name="Calibri", size=10)
    ctr = Alignment(horizontal='center', vertical='center', wrap_text=True)

    shutil.copy(input_path, output_path)
    wb = load_workbook(output_path)
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
            if mode == "words":
                ws.insert_cols(ci+1); wc = ci+1
                wh = ws.cell(row=1, column=wc)
                wh.value = cn.strip()+" in Word"; wh.fill = wf; wh.font = wfo; wh.alignment = ctr
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
    wb.save(output_path)

# ── Activity log ──────────────────────────────────────────────
def log_activity(user, action, detail=""):
    logs = []
    if os.path.exists(LOG_FILE):
        try:
            with open(LOG_FILE) as f: logs = json.load(f)
        except: pass
    logs.insert(0, {
        "time": datetime.now().strftime("%d %b %Y  %H:%M:%S"),
        "user": user,
        "action": action,
        "detail": detail
    })
    logs = logs[:200]  # keep last 200
    with open(LOG_FILE, "w") as f: json.dump(logs, f)

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
    if u: log_activity(u, "Logout", "")
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
    """Upload file, detect amount columns, return JSON list."""
    f = request.files.get("file")
    if not f or not f.filename.lower().endswith((".xlsx",".xls")):
        return jsonify({"error": "Please upload a valid .xlsx or .xls file."}), 400
    fname   = secure_filename(f.filename)
    fpath   = os.path.join(UPLOAD_FOLDER, f"{current_user()}_{fname}")
    f.save(fpath)
    session["upload_path"] = fpath
    session["upload_name"] = fname
    try:
        wb = load_workbook(fpath, read_only=True, data_only=True)
        ws = wb.active
        cols = find_amount_columns(ws)
        wb.close()
        return jsonify({"columns": [{"col": c, "name": n} for c,n in cols]})
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route("/process", methods=["POST"])
@require_login
def process():
    """Process uploaded file with selected column rules."""
    data   = request.json
    rules  = data.get("rules", {})
    fpath  = session.get("upload_path")
    fname  = session.get("upload_name","output.xlsx")
    if not fpath or not os.path.exists(fpath):
        return jsonify({"error": "No file uploaded. Please upload first."}), 400
    if not rules:
        return jsonify({"error": "No columns selected."}), 400
    name, ext = os.path.splitext(fname)
    out_name  = name + "_UPDATED" + ext
    out_path  = os.path.join(OUTPUT_FOLDER, f"{current_user()}_{out_name}")
    try:
        process_excel(fpath, out_path, rules)
        session["output_path"] = out_path
        session["output_name"] = out_name
        log_activity(current_user(), "Processed File", fname)
        return jsonify({"success": True, "filename": out_name})
    except Exception as e:
        log_activity(current_user(), "Process Error", str(e))
        return jsonify({"error": str(e)}), 500

@app.route("/download")
@require_login
def download():
    out_path = session.get("output_path")
    out_name = session.get("output_name","output_UPDATED.xlsx")
    if not out_path or not os.path.exists(out_path):
        flash("No processed file found. Please process a file first.")
        return redirect(url_for("dashboard"))
    log_activity(current_user(), "Downloaded File", out_name)
    return send_file(out_path, as_attachment=True, download_name=out_name)

@app.route("/admin")
@require_login
@require_admin
def admin():
    logs = []
    if os.path.exists(LOG_FILE):
        try:
            with open(LOG_FILE) as f: logs = json.load(f)
        except: pass
    users_info = [{"username": u, "role": d["role"]} for u, d in USERS.items()]
    return render_template("admin.html",
                           username=current_user(),
                           logs=logs,
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
    if os.path.exists(LOG_FILE):
        os.remove(LOG_FILE)
    log_activity(current_user(), "Cleared Logs", "")
    return redirect(url_for("admin"))

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
