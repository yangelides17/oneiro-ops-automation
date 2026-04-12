# Oneiro PDF Filler — Setup Guide

## What this does
`watch_and_fill.py` polls your Google Drive for JSON files that the Apps Script
exports whenever you generate daily documents or certified payroll. It downloads
each JSON, fills the matching PDF template, and uploads the filled PDF back to
the same Drive folder — no Google Drive for Desktop required.

---

## Step 1 — Install Python dependencies

```bash
pip3 install google-api-python-client google-auth-oauthlib pypdf --break-system-packages
```

---

## Step 2 — Create Google Cloud credentials (one-time, ~5 min)

1. Go to [console.cloud.google.com](https://console.cloud.google.com)
2. Click the project dropdown (top left) → **New Project** → name it `Oneiro Filler` → Create
3. In the left sidebar: **APIs & Services → Library**
4. Search **Google Drive API** → click it → **Enable**
5. In the left sidebar: **APIs & Services → OAuth consent screen**
   - User Type: **External** → Create
   - App name: `Oneiro Filler` | Support email: your Gmail
   - Click through to **Scopes** → Add scope → search `drive` → check
     `.../auth/drive` → Update → Save and Continue
   - **Test users** → Add your Gmail address → Save and Continue
6. In the left sidebar: **APIs & Services → Credentials**
   - **+ Create Credentials → OAuth client ID**
   - Application type: **Desktop app** → Name: `Oneiro Filler` → Create
   - Click **Download JSON** on the newly created credential
   - Rename the downloaded file to `credentials.json`
   - Move it into this folder (`oneiro-ops-script/`)

---

## Step 3 — Authenticate and discover folders

Run this once from the `oneiro-ops-script/` directory:

```bash
python3 watch_and_fill.py --setup
```

- A browser window will open → sign in with the same Google account that owns the Drive
- Click **Continue** through the "unverified app" warning (this is your own app)
- After login, the script saves `.token.json` and `.drive_config.json` automatically

> **Note:** If the `Production Logs` or `Certified Payroll` subfolders aren't found
> yet, that's fine — they'll be created the first time you trigger document generation
> in Apps Script. Re-run `--setup` afterward to pick them up.

---

## Step 4 — Push the updated Apps Script

```bash
# From the oneiro-ops-script/ folder:
clasp logout
clasp login
clasp push
```

---

## Step 5 — Start the watcher

```bash
python3 watch_and_fill.py
```

Leave it running in a Terminal tab. When you trigger document generation from the
Google Sheet, the filled PDFs will appear in the same Drive folder within ~20 seconds.

### One-shot mode (process now and exit)
Useful if you don't want to leave the watcher running:

```bash
python3 watch_and_fill.py --once
```

---

## Full pipeline test

1. Open the Google Sheet
2. Custom menu → **Generate Daily Documents (Custom Date)** → type `04/10/2026`
3. Watch the Terminal — you should see the JSON detected and the filled PDF uploaded
4. Check Drive → Needs Review → Production Logs for the filled PDF

---

## Files created by this script

| File | Purpose |
|------|---------|
| `credentials.json` | OAuth client secret (keep private, don't commit) |
| `.token.json` | Saved auth token (auto-refreshed) |
| `.drive_config.json` | Cached folder IDs |
| `.processed_files.json` | Tracks which Drive file IDs have been filled |
| `_tmp_fills/` | Temp folder for PDFs during fill (auto-cleaned) |
