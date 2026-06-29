
# ShipScan — Shipping Bill Extractor

A small web app that pulls **PDF attachments out of a Gmail inbox**, filters them by
filename / pattern / subject, downloads the matches, and logs everything to an Excel
report. It was built to extract Indian customs **shipping bills** (e.g. emails from
`noreply@icegate.gov.in`), but it works for any sender that emails you PDFs.

It connects to Gmail over IMAP, so there is no Gmail API setup — you just use a Gmail
**App Password**.

---

## Why it exists

Manually searching an inbox for hundreds of shipping-bill PDFs and saving them one by
one is slow and error-prone. ShipScan lets you paste a list of bill numbers (or just
fetch everything from a sender), and it downloads only the PDFs you asked for and
records the results in a spreadsheet.

---

## Features

- **Targeted Extract** — paste filename patterns (exact names or wildcards like
  `*3951543*`, `SB_*`, `*2024*`), optionally narrow by **email subject**, and the app
  downloads only the PDFs that match.
- **Browse & Pick** — list every PDF a sender has ever sent you, newest first, then
  tick the ones you want and save just those.
- **Excel report** — every run produces an `.xlsx` log (filename, matched filter, email
  date, subject, save path, status: Saved / Already Exists / Failed).
- **Subject filter** — an optional, server-side narrowing step that makes fetching
  faster when your emails share a common subject phrase.
- **Skips duplicates** — files already saved are detected and not re-downloaded.

---

## How it stays fast

The naive approach is to download every email in full just to read its attachment
filenames — for a mailbox with thousands of messages that means pulling hundreds of MB
to find a handful of files.

ShipScan avoids that:

1. It asks Gmail for each message's **`BODYSTRUCTURE` + `ENVELOPE` only** — tiny
   metadata that includes attachment filenames but **none of the attachment bytes** —
   and fetches it in **batches** of a few hundred messages per round-trip.
2. It matches those filenames against your filters using the metadata.
3. It downloads the actual PDF bytes **only** for the parts that matched and aren't
   already on disk, fetching just that one MIME part rather than the whole email.

The optional **subject filter** narrows the candidate set *server-side* (Gmail returns
fewer message IDs in the first place), so every later step has less to do.

---

## Tech stack

- **Backend:** Python, [Flask](https://flask.palletsprojects.com/),
  [IMAPClient](https://imapclient.readthedocs.io/) (IMAP + reliable `BODYSTRUCTURE`
  parsing), [openpyxl](https://openpyxl.readthedocs.io/) (Excel)
- **Frontend:** a single self-contained `templates/index.html` (vanilla HTML/CSS/JS,
  no build step)
- **Server (production):** [Gunicorn](https://gunicorn.org/)
- **Hosting:** [Render](https://render.com/) (free tier)

---

## Project structure

```
.
├── app.py               # Flask app: IMAP fetching, filtering, Excel + ZIP output
├── templates/
│   └── index.html       # the entire UI
├── requirements.txt     # Python dependencies (incl. gunicorn)
├── runtime.txt          # pins the Python version for Render
├── render.yaml          # Render blueprint (build/start commands, env vars)
├── .gitignore
└── README.md
```

---

## Getting a Gmail App Password (required)

ShipScan never uses your normal Gmail password. You need an **App Password**:

1. Enable 2-Step Verification on your Google account.
2. Go to <https://myaccount.google.com/apppasswords> and create an App Password.
3. Use that 16-character value in the app's **App password** field.

App Passwords are revocable at any time from the same page — if anything looks off,
revoke it and the access is gone.

---

## Run it locally

```bash
git clone https://github.com/<you>/SHIPS_SCAN.git
cd SHIPS_SCAN
pip install -r requirements.txt
python app.py
```

Open <http://localhost:5000>. With no `SHIPSCAN_PASSWORD` set, the app runs without a
login (fine for your own machine).

> Note: the hosted build delivers results as a **ZIP download** in the browser. The
> earlier local-first version of this tool saved PDFs straight to a folder on your PC;
> if you want that behaviour for purely local use, point the output at a real folder
> rather than the cloud ZIP path.

---

## Deploy to Render

This repo is set up for one-click deployment.

1. **Push to GitHub** (already done if you're reading this in the repo).
2. On Render: **New → Blueprint**, select this repo. Render reads `render.yaml`
   automatically. (Or **New → Web Service** and set the commands manually — see below.)
3. After the service is created, open its **Environment** tab and add:
   - **Key:** `SHIPSCAN_PASSWORD`
   - **Value:** a strong password of your choice
   Save — Render redeploys.
4. When the build finishes (~2–3 min), open `https://<your-app>.onrender.com`. Your
   browser shows a login box: **username can be anything**, password is the
   `SHIPSCAN_PASSWORD` you set.

If configuring manually instead of using the blueprint:

- **Build command:** `pip install -r requirements.txt`
- **Start command:** `gunicorn app:app --workers 1 --threads 4 --timeout 600 --bind 0.0.0.0:$PORT`
- **Plan:** Free

The single-worker start command is intentional — the app keeps job state in memory, so
it must not be split across workers.

---

## Environment variables

| Variable             | Required        | Purpose                                                                 |
|----------------------|-----------------|-------------------------------------------------------------------------|
| `SHIPSCAN_PASSWORD`  | Yes (on Render) | Locks the public site behind HTTP Basic auth. Without it the URL is open.|
| `OUTPUT_DIR`         | No              | Server-side folder for generated files (default `/tmp/shipscan_out`).   |
| `PORT`               | Auto (Render)   | Port to bind to; Render sets this for you.                              |

---

## Usage

1. Sign in (cloud build) and enter your **Gmail address** + **App Password**.
2. Set the **sender email** (e.g. `noreply@icegate.gov.in`).
3. *(Optional)* Add a **subject filter** — paste the constant part of the subject to
   speed things up.
4. **Targeted Extract:** paste filename patterns, one per line, then **Fetch PDFs**.
   Leave the patterns blank to fetch every PDF from the sender.
5. Download the **Excel report** and/or **All PDFs (ZIP)** when the run finishes.
6. **Browse & Pick:** load all PDFs from the sender, tick the ones you want, and save.

Tip: on a large mailbox, run with a small **Days back** value first (e.g. 7) to confirm
everything works before fetching the full history.

---

## Security notes

- Your **Gmail App Password** is entered into the form and sent to the server only to
  perform the IMAP fetch. It is **not** stored in the repo or in environment variables.
- Always set `SHIPSCAN_PASSWORD` when hosting — it's the only thing stopping a stranger
  who finds the URL from using the form.
- Keep the App Password revocable: rotate or delete it from your Google account if you
  ever suspect exposure.
- Don't commit any password to the repository.

---

## Limitations

- **Single user at a time.** Runs as one worker and processes one job at a time by
  design.
- **Free-tier cold starts.** A free Render instance sleeps after ~15 minutes of
  inactivity; the next visit takes ~30–60 seconds to wake.
- **Ephemeral storage in the cloud.** Files written on the server are wiped on restart
  or redeploy — download your ZIP when a run finishes; don't rely on server copies.
- **Personal tool, not multi-tenant.** Serving many users would require real per-user
  accounts and persistent storage.

---

## License

Add a license of your choice (e.g. MIT) if you intend others to use or contribute.
