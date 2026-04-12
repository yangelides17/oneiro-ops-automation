#!/usr/bin/env python3
"""
watch_and_fill.py  —  Oneiro Operations Drive watcher + PDF filler
===================================================================
Polls Google Drive for JSON files exported by the Apps Script,
fills the matching PDF template, and uploads the filled PDF back
to the same Drive folder.

On Railway: runs as a persistent worker process (see Procfile).
Locally:    python3 workers/watch_and_fill.py [--setup] [--once]

USAGE
  python3 workers/watch_and_fill.py             # run watcher
  python3 workers/watch_and_fill.py --setup     # authenticate + discover folder IDs
  python3 workers/watch_and_fill.py --once      # process pending files then exit
"""

import argparse
import json
import os
import sys
import time
import io
from datetime import datetime
from pathlib import Path

# ── CONFIGURATION ────────────────────────────────────────────────────────────
# WORKERS_DIR  = this file's directory  (repo/workers/)
# ROOT_DIR     = repo root              (repo/)
# Config/state files live at repo root so they're easy to find and gitignore.
WORKERS_DIR     = Path(__file__).parent
ROOT_DIR        = WORKERS_DIR.parent

CREDS_FILE      = ROOT_DIR / 'credentials.json'        # OAuth client secret (local) or env var (Railway)
TOKEN_FILE      = ROOT_DIR / '.token.json'              # auto-saved OAuth token
STATE_FILE      = ROOT_DIR / '.processed_files.json'    # tracks already-filled file IDs
TEMPLATE_CACHE  = ROOT_DIR / '_template_cache'          # local copies of Drive templates
POLL_SECONDS    = int(os.environ.get('POLL_SECONDS', 20))

# Google Drive API scopes needed
SCOPES = ['https://www.googleapis.com/auth/drive']

# Subfolder names inside Needs Review that we watch for JSON payloads
WATCH_SUBFOLDERS = ['Production Logs', 'Certified Payroll']

# Maps _type → template filename on Drive (inside the Templates folder)
TEMPLATE_FILES = {
    'production_log':    'Metro Thermoplastic Daily Production Log Template_FORM.pdf',
    'certified_payroll': 'Certified Payroll Report Template_FORM.pdf',
}

# MIME types we treat as scanned Work Order documents in the Scan Inbox
SCAN_INBOX_MIMETYPES = {
    'application/pdf':  'application/pdf',
    'image/jpeg':       'image/jpeg',
    'image/jpg':        'image/jpeg',
    'image/png':        'image/png',
    'image/tiff':       'image/tiff',
}

# ─────────────────────────────────────────────────────────────────────────────

sys.path.insert(0, str(WORKERS_DIR))  # filler modules live in workers/


# ── Logging ──────────────────────────────────────────────────────────────────
def log(msg: str):
    ts = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    print(f"[{ts}]  {msg}", flush=True)


# ── Google Auth ───────────────────────────────────────────────────────────────
def get_drive_service():
    """
    Build an authenticated Google Drive service (for reading only on Railway).

    Uploads are handled by the Apps Script proxy (doPost in Code.js), so the
    Drive service only needs read access — which service accounts handle fine.

    Priority:
      1. GOOGLE_SERVICE_ACCOUNT_JSON env var (Railway)
         Set this in Railway to the full contents of your service_account.json.
         Share the Oneiro Ops Drive folder with the service account email.

      2. Local OAuth flow (development machine)
         Requires credentials.json at repo root; opens browser on first run.
         Token is cached in .token.json for subsequent runs.
    """
    try:
        from googleapiclient.discovery import build
    except ImportError:
        print("\n❌  Required packages not installed. Run:\n")
        print("    pip install -r requirements.txt\n")
        sys.exit(1)

    import json as _json

    # ── Option 1: Service account (Railway) ───────────────────────────────────
    sa_json = os.environ.get('GOOGLE_SERVICE_ACCOUNT_JSON')
    if sa_json:
        try:
            from google.oauth2 import service_account
            sa_info = _json.loads(sa_json)
            creds = service_account.Credentials.from_service_account_info(
                sa_info, scopes=SCOPES
            )
            log("🔑  Authenticated via service account.")
            return build('drive', 'v3', credentials=creds)
        except Exception as e:
            print(f"\n❌  Service account auth failed: {e}\n")
            sys.exit(1)

    # ── Option 2: OAuth (local development) ──────────────────────────────────
    from google.oauth2.credentials import Credentials
    from google_auth_oauthlib.flow import InstalledAppFlow
    from google.auth.transport.requests import Request

    creds = None
    if TOKEN_FILE.exists():
        creds = Credentials.from_authorized_user_file(str(TOKEN_FILE), SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            if not CREDS_FILE.exists():
                print(f"\n❌  No auth method available.")
                print(f"    On Railway: set GOOGLE_SERVICE_ACCOUNT_JSON environment variable.")
                print(f"    Locally:    place credentials.json at {CREDS_FILE}\n")
                sys.exit(1)
            flow = InstalledAppFlow.from_client_secrets_file(str(CREDS_FILE), SCOPES)
            creds = flow.run_local_server(port=0)
        TOKEN_FILE.write_text(creds.to_json())
        log("✅  OAuth token saved.")

    return build('drive', 'v3', credentials=creds)


# ── Drive helpers ─────────────────────────────────────────────────────────────
def find_folder(service, name: str, parent_id: str = None) -> str | None:
    """Return the Drive folder ID matching name (optionally under parent_id)."""
    q = f"mimeType='application/vnd.google-apps.folder' and name='{name}' and trashed=false"
    if parent_id:
        q += f" and '{parent_id}' in parents"
    result = service.files().list(q=q, fields='files(id,name)').execute()
    files = result.get('files', [])
    return files[0]['id'] if files else None


def list_json_files(service, folder_id: str):
    """Return list of {id, name, createdTime} for .json files in folder_id."""
    q = f"'{folder_id}' in parents and mimeType='text/plain' and name contains '.json' and trashed=false"
    result = service.files().list(
        q=q,
        fields='files(id,name,createdTime)',
        orderBy='createdTime'
    ).execute()
    return result.get('files', [])


def download_file(service, file_id: str) -> dict:
    """Download and parse a JSON file from Drive."""
    from googleapiclient.http import MediaIoBaseDownload
    request = service.files().get_media(fileId=file_id)
    buf = io.BytesIO()
    downloader = MediaIoBaseDownload(buf, request)
    done = False
    while not done:
        _, done = downloader.next_chunk()
    buf.seek(0)
    return json.loads(buf.read().decode('utf-8'))


def download_binary(service, file_id: str) -> bytes:
    """Download any Drive file as raw bytes (PDF, image, etc.)."""
    from googleapiclient.http import MediaIoBaseDownload
    request = service.files().get_media(fileId=file_id)
    buf = io.BytesIO()
    downloader = MediaIoBaseDownload(buf, request)
    done = False
    while not done:
        _, done = downloader.next_chunk()
    return buf.getvalue()


def list_scan_inbox_files(service, folder_id: str) -> list:
    """Return unprocessed WO scan files (PDF/image) in the Scan Inbox folder."""
    mime_conditions = ' or '.join(
        f"mimeType='{m}'" for m in SCAN_INBOX_MIMETYPES
    )
    q = (f"'{folder_id}' in parents and ({mime_conditions}) "
         f"and not name contains '✅' and trashed=false")
    result = service.files().list(
        q=q,
        fields='files(id,name,mimeType,createdTime)',
        orderBy='createdTime'
    ).execute()
    return result.get('files', [])


def process_wo_scan(service, file_meta: dict, tmp_dir: Path) -> bool:
    """
    Download a WO scan from the Scan Inbox, parse it with Claude Vision,
    and post the structured data to the Apps Script doPost handler which
    writes a new row to the Work Order Tracker and archives the PDF.
    """
    import requests as _requests
    from parse_work_order import parse as parse_wo

    file_id   = file_meta['id']
    file_name = file_meta['name']
    mime_type = SCAN_INBOX_MIMETYPES.get(file_meta.get('mimeType', ''), 'application/pdf')

    log(f"📋  WO Scan: {file_name}")

    # ── Download ──────────────────────────────────────────────────
    try:
        file_bytes = download_binary(service, file_id)
        log(f"  📥  Downloaded {len(file_bytes) / 1024:.0f} KB")
    except Exception as e:
        log(f"  ❌  Download failed: {e}")
        return False

    # ── Parse with Claude Vision ──────────────────────────────────
    try:
        log("  🔍  Parsing with Claude Vision...")
        wo_data = parse_wo(file_bytes, mime_type)
    except Exception as e:
        log(f"  ❌  Parse failed: {e}")
        return False

    if '_parse_error' in wo_data:
        log(f"  ❌  Claude could not parse WO: {wo_data['_parse_error']}")
        return False

    wo_id = wo_data.get('work_order_id', 'UNKNOWN')
    log(f"  ✅  Parsed: {wo_id} — {wo_data.get('prime_contractor')} / "
        f"{wo_data.get('contract_number')} / {wo_data.get('location')}")

    # ── Post to Apps Script ───────────────────────────────────────
    upload_url = os.environ.get('APPS_SCRIPT_UPLOAD_URL')
    upload_key = os.environ.get('APPS_SCRIPT_UPLOAD_KEY', '')

    if not upload_url:
        log("  ⚠️   APPS_SCRIPT_UPLOAD_URL not set — cannot write to WO Tracker.")
        log("      Set this in Railway env vars.")
        return False

    import json as _json
    payload = _json.dumps({
        'action':  'write_wo',
        'key':     upload_key,
        'file_id': file_id,
        'data':    wo_data,
    })

    try:
        resp = _requests.post(
            upload_url,
            data=payload,
            headers={'Content-Type': 'application/json'},
            timeout=60,
            allow_redirects=True,
        )
        log(f"  📡  Apps Script response: HTTP {resp.status_code}, {len(resp.content)} bytes")

        if not resp.content.strip():
            log("  ❌  Empty response from Apps Script — check doPost deployment.")
            return False

        result = resp.json()
        if 'error' in result:
            log(f"  ❌  Apps Script error: {result['error']}")
            return False

        if result.get('duplicate'):
            log(f"  ℹ️   WO {wo_id} already in tracker — archived file only.")
        else:
            log(f"  ✅  WO {wo_id} added to tracker and archived.")

        return True

    except Exception as e:
        log(f"  ❌  Apps Script POST failed: {e}")
        return False


def upload_pdf(service, local_path: Path, parent_folder_id: str) -> str:
    """Upload a filled PDF to Google Drive.

    On Railway: routes through the Apps Script Web App proxy (doPost in Code.js),
    which saves the file as the real Drive-owning Google account.
    Set APPS_SCRIPT_UPLOAD_URL and APPS_SCRIPT_UPLOAD_KEY in Railway env vars.

    Locally: uploads directly via the Drive API using the OAuth credentials
    obtained during --setup (no proxy needed when running as yourself).
    """
    upload_url = os.environ.get('APPS_SCRIPT_UPLOAD_URL')
    if upload_url:
        upload_key = os.environ.get('APPS_SCRIPT_UPLOAD_KEY', '')
        return _upload_via_proxy(local_path, parent_folder_id, upload_url, upload_key)
    else:
        return _upload_direct(service, local_path, parent_folder_id)


def _upload_via_proxy(local_path: Path, folder_id: str, url: str, key: str) -> str:
    """POST base64-encoded PDF bytes to the Apps Script upload proxy.

    The proxy (doPost in Code.js) runs as your Google account and saves the
    file to Drive — bypassing the service-account storage-quota limitation.
    Returns the created Drive file ID.
    """
    import requests, base64

    pdf_bytes = local_path.read_bytes()
    encoded   = base64.b64encode(pdf_bytes).decode('utf-8')

    # Send everything as a JSON body — Apps Script handles this natively
    # via e.postData.contents. Mixing URL params + octet-stream body is
    # unreliable across Google's infrastructure and causes HTML error pages.
    import json as _json
    payload = _json.dumps({
        'key':       key,
        'filename':  local_path.name,
        'folder_id': folder_id,
        'data':      encoded,
    })
    resp = requests.post(
        url,
        data=payload,
        headers={'Content-Type': 'application/json'},
        timeout=60,
        allow_redirects=True,
    )

    # Log the raw response so we can diagnose proxy issues clearly
    log(f"  📡  Proxy response: HTTP {resp.status_code}, {len(resp.content)} bytes")

    # Empty body = proxy not reachable or Web App not deployed with doPost
    if not resp.content.strip():
        raise RuntimeError(
            f"Apps Script proxy returned an empty response (HTTP {resp.status_code}). "
            f"Common causes:\n"
            f"    • doPost not yet deployed — go to Deploy → New deployment in Apps Script\n"
            f"    • 'Who has access' is not set to 'Anyone' on the Web App deployment\n"
            f"    • APPS_SCRIPT_UPLOAD_URL is a /dev URL (requires auth) — use the /exec URL\n"
            f"    • APPS_SCRIPT_UPLOAD_KEY doesn't match UPLOAD_SECRET in Script Properties"
        )

    resp.raise_for_status()

    # Non-JSON response = Google returned an HTML error / login page
    try:
        result = resp.json()
    except Exception:
        snippet = resp.text[:300].replace('\n', ' ')
        raise RuntimeError(
            f"Apps Script proxy returned non-JSON (HTTP {resp.status_code}): {snippet!r}\n"
            f"    This usually means the Web App redirected to a login page.\n"
            f"    Ensure 'Who has access' is set to 'Anyone' (not 'Anyone with Google account')."
        )

    if 'error' in result:
        raise RuntimeError(f"Apps Script proxy error: {result['error']}")

    log(f"  📤  Proxy upload complete → {result.get('filename')}  (Drive ID: {result.get('file_id')})")
    return result.get('file_id', '')


def _upload_direct(service, local_path: Path, parent_folder_id: str) -> str:
    """Upload directly via Drive API (local dev with OAuth credentials)."""
    from googleapiclient.http import MediaFileUpload
    meta  = {'name': local_path.name, 'parents': [parent_folder_id]}
    media = MediaFileUpload(str(local_path), mimetype='application/pdf')
    created = service.files().create(body=meta, media_body=media, fields='id').execute()
    return created.get('id')


def get_template(service, doc_type: str) -> Path:
    """Return a local path to the template PDF, downloading from Drive if needed.

    Templates are cached in _template_cache/ and re-downloaded only when
    the Drive copy is newer than the local copy.
    """
    from googleapiclient.http import MediaIoBaseDownload

    filename = TEMPLATE_FILES.get(doc_type)
    if not filename:
        raise ValueError(f"No template configured for doc_type={doc_type!r}")

    # Load templates folder ID — env var (Railway) takes priority over local config
    templates_folder_id = os.environ.get('DRIVE_TEMPLATES_ID')
    if not templates_folder_id:
        config_path = ROOT_DIR / '.drive_config.json'
        if config_path.exists():
            templates_folder_id = json.loads(config_path.read_text()).get('templates_folder_id')

    TEMPLATE_CACHE.mkdir(exist_ok=True)
    local_path = TEMPLATE_CACHE / filename

    if templates_folder_id:
        # Find the file in Drive
        q = (f"'{templates_folder_id}' in parents and name='{filename}' "
             f"and trashed=false")
        result = service.files().list(
            q=q, fields='files(id,name,modifiedTime)'
        ).execute()
        files = result.get('files', [])

        if files:
            drive_file = files[0]
            drive_modified = drive_file['modifiedTime']

            # Download if local copy is missing or stale
            needs_download = not local_path.exists()
            if not needs_download and local_path.exists():
                import email.utils, calendar
                local_mtime = local_path.stat().st_mtime
                drive_ts = calendar.timegm(
                    time.strptime(drive_modified, '%Y-%m-%dT%H:%M:%S.%fZ')
                    if '.' in drive_modified else
                    time.strptime(drive_modified, '%Y-%m-%dT%H:%M:%SZ')
                )
                needs_download = drive_ts > local_mtime

            if needs_download:
                log(f"  📥  Downloading template from Drive: {filename}")
                request = service.files().get_media(fileId=drive_file['id'])
                buf = io.BytesIO()
                dl = MediaIoBaseDownload(buf, request)
                done = False
                while not done:
                    _, done = dl.next_chunk()
                local_path.write_bytes(buf.getvalue())
                log(f"  ✅  Template cached: {local_path.name}")
        else:
            log(f"  ⚠️   Template '{filename}' not found in Drive Templates folder.")
            if not local_path.exists():
                raise FileNotFoundError(
                    f"Template '{filename}' not in Drive and no local cache. "
                    f"Upload it to the Templates folder in Drive."
                )
            log(f"      Using cached copy.")
    else:
        log(f"  ⚠️   Templates folder not configured. Using local cache if available.")
        if not local_path.exists():
            raise FileNotFoundError(
                f"Template '{filename}' not found. Re-run --setup so the "
                f"Templates folder can be discovered, or upload manually."
            )

    return local_path


# ── State tracking ────────────────────────────────────────────────────────────
def load_state() -> set:
    if STATE_FILE.exists():
        return set(json.loads(STATE_FILE.read_text()))
    return set()


def save_state(processed: set):
    STATE_FILE.write_text(json.dumps(list(processed)))


# ── PDF fillers ───────────────────────────────────────────────────────────────
def fill_production_log(service, data: dict, tmp_dir: Path) -> Path:
    from fill_production_log import fill
    template = get_template(service, 'production_log')
    date_str = data.get('date', 'unknown').replace('/', '-')
    out = tmp_dir / f"Production_Log_{date_str}_FILLED.pdf"
    fill(data, template_path=str(template), output_path=str(out))
    return out


def fill_certified_payroll(service, data: dict, tmp_dir: Path) -> Path:
    from fill_certified_payroll import fill
    template = get_template(service, 'certified_payroll')
    week = data.get('week_ending', 'unknown').replace('/', '-')
    out = tmp_dir / f"Certified_Payroll_{week}_FILLED.pdf"
    fill(data, template_path=str(template), output_path=str(out))
    return out


FILLERS = {
    'production_log':    fill_production_log,
    'certified_payroll': fill_certified_payroll,
}


# ── Core processing ───────────────────────────────────────────────────────────
def process_file(service, file_meta: dict, folder_id: str, tmp_dir: Path) -> bool:
    file_id = file_meta['id']
    name    = file_meta['name']
    log(f"📄  {name}")

    try:
        data = download_file(service, file_id)
    except Exception as e:
        log(f"  ❌  Download failed: {e}")
        return False

    doc_type = data.get('_type', '')
    filler   = FILLERS.get(doc_type)
    if not filler:
        log(f"  ⚠️   Unknown _type={doc_type!r} — skipping")
        return True  # mark as seen so we don't retry forever

    try:
        pdf_path = filler(service, data, tmp_dir)
    except Exception as e:
        log(f"  ❌  Fill failed: {e}")
        import traceback; traceback.print_exc()
        return False

    try:
        new_id = upload_pdf(service, pdf_path, folder_id)
        log(f"  ✅  Uploaded → {pdf_path.name}  (Drive ID: {new_id})")
        pdf_path.unlink()   # clean up local temp file
    except Exception as e:
        log(f"  ❌  Upload failed: {e}")
        return False

    return True


def scan_once(service, folder_map: dict, scan_inbox_id: str | None,
             processed: set, tmp_dir: Path):
    # ── JSON payloads → PDF fillers ───────────────────────────────
    for subfolder_name, folder_id in folder_map.items():
        files = list_json_files(service, folder_id)
        for f in files:
            if f['id'] in processed:
                continue
            success = process_file(service, f, folder_id, tmp_dir)
            if success:
                processed.add(f['id'])
                save_state(processed)

    # ── Scan Inbox → Claude Vision WO parser ─────────────────────
    if scan_inbox_id:
        wo_files = list_scan_inbox_files(service, scan_inbox_id)
        for f in wo_files:
            if f['id'] in processed:
                continue
            success = process_wo_scan(service, f, tmp_dir)
            # Mark as processed regardless — on failure we don't want infinite retries.
            # The file remains in Scan Inbox with its original name so the admin can
            # manually review. Apps Script will rename it with ✅ on success.
            processed.add(f['id'])
            save_state(processed)


# ── Setup mode ────────────────────────────────────────────────────────────────
def run_setup(service, folder_id: str = None, templates_id: str = None):
    log("🔍  Searching Drive for Oneiro folder structure...")

    if folder_id:
        needs_review_id = folder_id
        log(f"  ✅  Using provided folder ID: {needs_review_id}")
    else:
        needs_review_id = find_folder(service, 'Needs Review') or find_folder(service, 'Docs Needing Review')
        if not needs_review_id:
            log("❌  Could not find the review folder in your Drive.")
            log("   Run with --folder-id to specify it directly:")
            log("   python3 watch_and_fill.py --setup --folder-id <YOUR_FOLDER_ID>")
            return

    log(f"  ✅  Needs Review folder: {needs_review_id}")

    folder_map = {}
    for sub in WATCH_SUBFOLDERS:
        fid = find_folder(service, sub, parent_id=needs_review_id)
        if fid:
            folder_map[sub] = fid
            log(f"  ✅  {sub}: {fid}")
        else:
            log(f"  ⚠️   '{sub}' subfolder not found — will be created by Apps Script on first run")

    # Find Templates folder
    templates_folder_id = templates_id  # may be None if not provided
    if templates_folder_id:
        log(f"  ✅  Templates folder (provided): {templates_folder_id}")
    else:
        for name in ['⚙️ Templates', 'Templates', 'templates']:
            templates_folder_id = find_folder(service, name)
            if templates_folder_id:
                log(f"  ✅  Templates folder ({name!r}): {templates_folder_id}")
                break
    if not templates_folder_id:
        log("  ⚠️   Templates folder not found in Drive.")
        log("      Re-run with --templates-id <FOLDER_ID> to register it.")

    config = {
        'needs_review_id':    needs_review_id,
        'folders':            folder_map,
        'templates_folder_id': templates_folder_id,
    }
    config_path = ROOT_DIR / '.drive_config.json'
    config_path.write_text(json.dumps(config, indent=2))
    log(f"\n✅  Config saved to {config_path.name}")
    log("   Run  python3 watch_and_fill.py  to start watching.\n")


def load_folder_map(service) -> tuple:
    """Load folder IDs from environment variables (Railway) or .drive_config.json (local).

    Environment variables take priority — set these in the Railway dashboard:
      DRIVE_NEEDS_REVIEW_ID   — ID of the "Docs Needing Review" folder
      DRIVE_TEMPLATES_ID      — ID of the Templates folder
      DRIVE_SCAN_INBOX_ID     — ID of the "Scan Inbox" folder (WO scans)
      DRIVE_PROD_LOGS_ID      — (optional) ID of Production Logs subfolder
      DRIVE_CERT_PAYROLL_ID   — (optional) ID of Certified Payroll subfolder

    Returns:
      (folder_map, scan_inbox_id) where folder_map is {subfolder_name: id}
      and scan_inbox_id is the Drive ID of the Scan Inbox (or None).
    """
    needs_review_id     = os.environ.get('DRIVE_NEEDS_REVIEW_ID')
    templates_folder_id = os.environ.get('DRIVE_TEMPLATES_ID')
    scan_inbox_id       = os.environ.get('DRIVE_SCAN_INBOX_ID')
    prod_logs_id        = os.environ.get('DRIVE_PROD_LOGS_ID')
    cert_payroll_id     = os.environ.get('DRIVE_CERT_PAYROLL_ID')

    # Fall back to local .drive_config.json if env vars not set
    config_path = ROOT_DIR / '.drive_config.json'
    if not needs_review_id:
        if not config_path.exists():
            log("❌  No Drive config found.")
            log("    On Railway: set DRIVE_NEEDS_REVIEW_ID environment variable.")
            log("    Locally:    run  python3 workers/watch_and_fill.py --setup")
            sys.exit(1)
        config = json.loads(config_path.read_text())
        needs_review_id     = config.get('needs_review_id')
        templates_folder_id = templates_folder_id or config.get('templates_folder_id')
        saved_folders       = config.get('folders', {})
        prod_logs_id        = prod_logs_id    or saved_folders.get('Production Logs')
        cert_payroll_id     = cert_payroll_id or saved_folders.get('Certified Payroll')
    else:
        log("📋  Loaded folder IDs from environment variables.")

    # Build folder map — discover any subfolders not yet known
    folder_map = {}
    for sub, known_id in [('Production Logs', prod_logs_id), ('Certified Payroll', cert_payroll_id)]:
        if known_id:
            folder_map[sub] = known_id
        elif needs_review_id:
            fid = find_folder(service, sub, parent_id=needs_review_id)
            if fid:
                folder_map[sub] = fid
                log(f"  🔎  Discovered subfolder '{sub}': {fid}")

    # Discover Scan Inbox if not provided
    if not scan_inbox_id and needs_review_id:
        # Scan Inbox is a top-level folder, not under Needs Review — search broadly
        for name in ['📥 Scan Inbox', 'Scan Inbox', 'scan inbox']:
            scan_inbox_id = find_folder(service, name)
            if scan_inbox_id:
                log(f"  🔎  Discovered Scan Inbox: {scan_inbox_id}")
                break
    elif scan_inbox_id:
        log(f"  📥  Scan Inbox: {scan_inbox_id}")

    if not scan_inbox_id:
        log("  ⚠️   Scan Inbox not found. Set DRIVE_SCAN_INBOX_ID to enable WO parsing.")

    return ({k: v for k, v in folder_map.items() if k in WATCH_SUBFOLDERS},
            scan_inbox_id)


# ── Entry point ───────────────────────────────────────────────────────────────
def main():
    parser = argparse.ArgumentParser(description='Oneiro PDF filler — Drive watcher')
    parser.add_argument('--setup',        action='store_true', help='Authenticate + discover folder IDs')
    parser.add_argument('--once',         action='store_true', help='Process pending files then exit')
    parser.add_argument('--folder-id',    default=None,        help='Manually specify the Docs Needing Review folder ID')
    parser.add_argument('--templates-id', default=None,        help='Manually specify the Templates folder ID')
    args = parser.parse_args()

    service = get_drive_service()

    if args.setup:
        run_setup(service, folder_id=args.folder_id, templates_id=args.templates_id)
        return

    folder_map, scan_inbox_id = load_folder_map(service)
    if not folder_map and not scan_inbox_id:
        log("⚠️   No watched folders found.")
        log("    Trigger generateDailyDocuments once in Apps Script,")
        log("    then re-run  python3 watch_and_fill.py --setup  to pick up subfolders.")
        return

    tmp_dir = ROOT_DIR / '_tmp_fills'
    tmp_dir.mkdir(exist_ok=True)

    processed = load_state()

    watching = list(folder_map.keys())
    if scan_inbox_id:
        watching.append('Scan Inbox')

    if args.once:
        log("⚡  Running one-shot scan...")
        scan_once(service, folder_map, scan_inbox_id, processed, tmp_dir)
        log("Done.")
        return

    log(f"👀  Watching: {', '.join(watching)}")
    log(f"   Polling every {POLL_SECONDS}s  (Ctrl+C to stop)\n")
    try:
        while True:
            scan_once(service, folder_map, scan_inbox_id, processed, tmp_dir)
            time.sleep(POLL_SECONDS)
    except KeyboardInterrupt:
        log("\n👋  Watcher stopped.")


if __name__ == '__main__':
    main()
