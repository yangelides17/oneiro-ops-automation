# Oneiro Operations Automation

Automated document generation and filing system for Oneiro Collection LLC, a NYC DOT pavement marking subcontractor.

## What this does

- **Watches Google Drive** for JSON payloads exported by the Apps Script data hub
- **Fills PDF forms** automatically — Production Logs, Certified Payroll Reports, Employee Sign-In Sheets, and Monthly Workforce Utilization Tables
- **Parses scanned Work Orders** dropped into the Scan Inbox folder using Claude's Vision API
- **Uploads filled PDFs** back to the correct Drive folders for review

## Repo structure

```
workers/
  watch_and_fill.py        # Main Drive watcher and job dispatcher
  fill_production_log.py   # Metro Thermoplastic Daily Production Log filler
  fill_certified_payroll.py# NYC DOT Certified Payroll Report filler
  fill_signin.py           # Employee Daily Sign-In Log filler
  fill_utilization.py      # Monthly Workforce Utilization Form filler
  parse_work_order.py      # Claude Vision API parser for scanned Work Orders

apps-script/
  Code.js                  # Google Apps Script source (deployed separately via clasp)

templates/                 # PDF form templates (also stored in Google Drive)
Procfile                   # Railway worker process definition
requirements.txt           # Python dependencies
```

## Tech stack

- **Google Apps Script** — Sheets UI, triggers, JSON export
- **Google Drive API** — file polling, download, upload
- **pypdf** — AcroForm PDF filling
- **Claude API (Anthropic)** — Vision API for Work Order parsing
- **Railway** — cloud hosting for the Python worker process
- **QuickBooks API** — invoice integration (planned)

## Deployment

The worker runs as a persistent process on Railway, connected to this GitHub repo. Pushes to `main` trigger an automatic redeploy.

Environment variables required in Railway dashboard:
- `ANTHROPIC_API_KEY` — Claude API key
- `GOOGLE_SERVICE_ACCOUNT_JSON` — Google service account credentials (JSON, base64-encoded)

## Local development

```bash
pip install -r requirements.txt
python workers/watch_and_fill.py --setup   # first-time auth + folder discovery
python workers/watch_and_fill.py           # run watcher
```

Apps Script changes are deployed separately:
```bash
clasp push
```
