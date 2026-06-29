# Deploying ShipScan to Render

This is the **cloud build**. Differences from the local version:

- It is **password-protected** (HTTP Basic auth) so the public URL isn't open to anyone.
- Results are delivered as a **ZIP download** in your browser (a hosted server has no
  `C:\` folder, and its disk is wiped on every restart/redeploy).
- It binds to `0.0.0.0` and the `PORT` Render provides, and runs under **gunicorn**.

---

## Before you start — read this

You are about to put a form that accepts your **Gmail App Password** on a public URL.
Do this safely:

1. **Use a Gmail App Password, never your real password.** App Passwords are revocable
   any time at https://myaccount.google.com/apppasswords — if anything looks off, revoke it.
2. **Always set `SHIPSCAN_PASSWORD`** (step 4 below). Without it the site is wide open.
3. The free instance **sleeps after ~15 minutes** of no traffic; the next visit takes
   ~30–60 seconds to wake (a "cold start"). That's normal on the free plan.

---

## Steps

### 1. Put the code on GitHub
Create a new GitHub repo and push these files (keep the folder structure):

```
app.py
requirements.txt
runtime.txt
render.yaml
templates/index.html
```

From this folder:
```bash
git init
git add .
git commit -m "ShipScan cloud build"
git branch -M main
git remote add origin https://github.com/<you>/<repo>.git
git push -u origin main
```

### 2. Create the service on Render
- Go to https://render.com and sign in with GitHub.
- **New → Web Service**, connect the repo.
- Render reads `render.yaml` automatically. If it asks, confirm:
  - **Build command:** `pip install -r requirements.txt`
  - **Start command:** `gunicorn app:app --workers 1 --threads 4 --timeout 600 --bind 0.0.0.0:$PORT`
  - **Plan:** Free
  - **Language:** Python 3

### 3. (If you used the Blueprint path)
Pick **New → Blueprint** instead, select the repo, and Render provisions everything
from `render.yaml`.

### 4. Set your password (required)
In the service's **Environment** tab, add:
- **Key:** `SHIPSCAN_PASSWORD`
- **Value:** a strong password of your choosing

Save. Render redeploys. When you open the site, the browser will prompt for a login —
leave the username blank (or type anything) and enter this password.

### 5. Use it
- Open `https://<your-app>.onrender.com`, log in with your password.
- Fill in Gmail address, App Password, sender, subject filter, filters.
- Click **Fetch PDFs**.
- When it finishes, click **⬇ All PDFs (ZIP)** to download everything, or
  **⬇ Excel report** for just the spreadsheet.

---

## Troubleshooting

- **502 / "Application failed to respond" right after deploy:** the start command is
  wrong or a package is missing. Check the **Logs** tab. The start command must be
  exactly `gunicorn app:app ...` because the Flask object in `app.py` is named `app`.
- **Build fails on the Python version:** delete `runtime.txt` and redeploy to use
  Render's default Python.
- **Gmail "login failed":** make sure IMAP is enabled in Gmail and you're using an
  **App Password**, not your account password. A login from a new datacenter IP can
  occasionally trigger a Google security prompt — approve it from your Google account.
- **It's slow to first load:** that's the free-tier cold start after the instance slept.
  Upgrading to the Starter plan keeps it always-on.

---

## Important limits of the free plan for this app

- **One user at a time.** The app runs a single worker and serializes jobs by design.
- **Files are temporary.** Download your ZIP when the run finishes; the server copy is
  wiped on the next restart/redeploy.
- **Not built for many concurrent users.** This is a personal tool, not a multi-tenant
  product. If you need that, it needs real per-user accounts and storage first.
