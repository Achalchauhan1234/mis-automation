# 🚀 MIS Automation — Deployment Guide
## How to get your live link in 10 minutes (Free)

---

## What you'll get
A live URL like: **https://mis-automation.onrender.com**
- Employees open this in any browser (PC, Mac, Mobile)
- They see a login page — no code visible
- You (achal13) get an Admin panel at /admin

---

## Step 1 — Create a GitHub account (if you don't have one)
Go to https://github.com and sign up (free).

---

## Step 2 — Upload these files to GitHub

1. Click the **+** button → **New repository**
2. Name it: `mis-automation`
3. Set it to **Private** (important — keeps your code hidden)
4. Click **Create repository**
5. Click **uploading an existing file**
6. Upload ALL files from this folder:
   - `app.py`
   - `requirements.txt`
   - `render.yaml`
   - folder `templates/` (with login.html, dashboard.html, admin.html)
7. Click **Commit changes**

---

## Step 3 — Deploy on Render (free hosting)

1. Go to https://render.com and sign up (free)
2. Click **New +** → **Web Service**
3. Connect your GitHub account
4. Select your `mis-automation` repository
5. Render auto-detects settings from render.yaml
6. Click **Create Web Service**
7. Wait ~2 minutes for deployment

✅ You'll get a URL like: `https://mis-automation.onrender.com`

---

## Step 4 — Share the link

Send this link to all employees:
```
https://mis-automation.onrender.com
```

They will see the login page. Only people with credentials can enter.

---

## Your Login (Admin)
- **Username:** achal13
- **Password:** achal1331
- **Admin panel:** https://your-app.onrender.com/admin

---

## Adding new employees (giving them access)

1. Log in with your admin account
2. Click **🛡 Admin** in the top navigation
3. Scroll to the **Users** section
4. Enter new username, password, role = User
5. Click **+ Add User**

They can now log in with those credentials.

---

## Important Notes

⚠️  **Free tier on Render sleeps after 15 min of inactivity.**
    First load may take 30–60 seconds to wake up.
    Upgrade to paid ($7/month) to keep it always-on.

🔒  **Employees only see the dashboard** — they cannot:
    - View or edit any code
    - Access the admin panel
    - See other users' files

🛡  **You (admin) can:**
    - Add/remove users
    - See who logged in and when
    - See which files were processed
    - Clear activity logs

---

## Running locally (on your own PC for testing)

```bash
pip install flask openpyxl
python app.py
```
Then open: http://localhost:5000
