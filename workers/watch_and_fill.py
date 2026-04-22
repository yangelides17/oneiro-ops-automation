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
WATCH_SUBFOLDERS = ['Production Logs', 'Certified Payroll', 'Sign-In Logs', 'Field Reports']

# Maps _type → template filename on Drive (inside the Templates folder)
TEMPLATE_FILES = {
    'production_log':           'Metro Thermoplastic Daily Production Log Template_FORM.pdf',
    'certified_payroll':        'Certified Payroll Report Template_FORM.pdf',
    'signin':                   'Employee Daily Sign In Log Template_FORM.pdf',
    'contractor_field_report':  'Thermo Contractor Field Report Template_FORM.pdf',
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
    """Return list of {id, name, createdTime} for .json files in folder_id.

    Accepts either mimeType — Apps Script writes JSON as text/plain via
    MimeType.PLAIN_TEXT, but manual uploads (or Drive API uploads) often
    use application/json. Match both."""
    q = (f"'{folder_id}' in parents and name contains '.json' and trashed=false "
         f"and (mimeType='text/plain' or mimeType='application/json')")
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


def _apps_script_post(action: str, body: dict) -> dict:
    """
    POST a JSON payload to the Apps Script Web App. Returns the parsed
    JSON response. Raises on network / non-JSON / error responses.
    """
    import requests as _requests
    import json as _json

    upload_url = os.environ.get('APPS_SCRIPT_UPLOAD_URL')
    upload_key = os.environ.get('APPS_SCRIPT_UPLOAD_KEY', '')
    if not upload_url:
        raise RuntimeError('APPS_SCRIPT_UPLOAD_URL not set')

    payload = _json.dumps({'action': action, 'key': upload_key, **body})
    resp = _requests.post(
        upload_url, data=payload,
        headers={'Content-Type': 'application/json'},
        timeout=120, allow_redirects=True,
    )
    if not resp.content.strip():
        raise RuntimeError(f'Empty Apps Script response (HTTP {resp.status_code})')
    try:
        return resp.json()
    except Exception as e:
        snippet = resp.text[:300].replace('\n', ' ')
        raise RuntimeError(f'Non-JSON Apps Script response: {snippet!r}') from e


def _log_scan_failure(file_id: str, filename: str, error: str):
    """Best-effort write of a scan-failure row to Automation Log. The
       webapp's scan-status poll keys on this."""
    try:
        _apps_script_post('log_wo_scan_failure', {'data': {
            'file_id': file_id, 'filename': filename, 'error': error,
        }})
    except Exception as e:
        log(f"  ⚠️   Could not log scan failure for {file_id}: {e}")


def _trash_via_apps_script(file_id: str, label: str = ''):
    """Trash a Drive file via the Apps Script proxy so it runs as the
       spreadsheet owner (not the Railway OAuth user) — guarantees
       permission regardless of who owns the file in Scan Inbox."""
    if not file_id:
        return
    try:
        _apps_script_post('trash_file', {'file_id': file_id})
        log(f"  🗑️   Trashed{(' ' + label) if label else ''}: {file_id}")
    except Exception as e:
        log(f"  ⚠️   Could not trash {label or file_id}: {e}")


def _pdf_page_count(file_bytes: bytes) -> int:
    """Returns PDF page count, or 0 if not a PDF / unreadable."""
    try:
        from pypdf import PdfReader
        import io
        return len(PdfReader(io.BytesIO(file_bytes)).pages)
    except Exception:
        return 0


def split_pdf_by_pages(file_bytes: bytes, page_lists):
    """
    Split a PDF into N new PDFs. `page_lists` is a list of lists of
    1-indexed page numbers. Returns a list of bytes (one per input
    page list). Pages referenced out of range are silently skipped.
    """
    from pypdf import PdfReader, PdfWriter
    import io

    reader = PdfReader(io.BytesIO(file_bytes))
    total  = len(reader.pages)
    outputs = []
    for pages in page_lists:
        writer = PdfWriter()
        for p in pages:
            # page_lists use 1-indexed page numbers; pypdf is 0-indexed
            idx = int(p) - 1
            if 0 <= idx < total:
                writer.add_page(reader.pages[idx])
        buf = io.BytesIO()
        writer.write(buf)
        outputs.append(buf.getvalue())
    return outputs


def _write_wo_from_parsed(wo_data: dict, file_id: str, combined_file_id: str = '',
                          original_filename: str = '') -> bool:
    """
    POST a parsed WO to Apps Script `write_wo`. Returns True on success
    or duplicate, False on any error.

    combined_file_id   populates tracker col 39 — only set for splits
                       from a multi-WO combined PDF.
    original_filename  populates tracker col 41 — the filename the user
                       picked in the webapp (for multi-WO splits this
                       is the combined PDF's name, so all splits from
                       one upload share it for queue grouping).
    """
    if '_parse_error' in wo_data:
        log(f"  ❌  Parse error: {wo_data['_parse_error']}")
        return False

    wo_id  = wo_data.get('work_order_id', 'UNKNOWN')
    top_ct = len(wo_data.get('top_markings') or [])
    grid_ct = len(wo_data.get('intersection_grid') or [])
    wtype  = wo_data.get('work_type') or '?'
    log(f"  ✅  Parsed: {wo_id} — {wo_data.get('prime_contractor')} / "
        f"{wo_data.get('contract_number')} / {wo_data.get('location')}")
    log(f"      work_type={wtype}, top_markings={top_ct}, intersection_grid rows={grid_ct}")

    try:
        result = _apps_script_post('write_wo', {
            'file_id':           file_id,
            'combined_file_id':  combined_file_id,
            'original_filename': original_filename,
            'data':              wo_data,
        })
    except Exception as e:
        log(f"  ❌  Apps Script POST failed: {e}")
        return False

    if 'error' in result:
        log(f"  ❌  Apps Script error: {result['error']}")
        return False
    if result.get('duplicate'):
        log(f"  ℹ️   WO {wo_id} already in tracker — archived file only.")
    else:
        log(f"  ✅  WO {wo_id} added to tracker and archived.")
    return True


def _upload_split_to_scan_inbox(service, file_bytes: bytes, filename: str) -> str:
    """
    Upload a split WO PDF to Scan Inbox via the Apps Script proxy and
    return its new file_id. We skip the Drive API direct path because
    the service account doesn't own Scan Inbox — the proxy runs as the
    spreadsheet's owner so writes always land cleanly.
    """
    import base64
    encoded = base64.b64encode(file_bytes).decode('utf-8')
    result  = _apps_script_post('upload_wo_scan', {'data': {
        'filename':  filename,
        'mime_type': 'application/pdf',
        'data':      encoded,
    }})
    if 'error' in result:
        raise RuntimeError(f'upload_wo_scan error: {result["error"]}')
    fid = result.get('file_id')
    if not fid:
        raise RuntimeError('upload_wo_scan returned no file_id')
    return fid


def process_wo_scan(service, file_meta: dict, tmp_dir: Path) -> bool:
    """
    Download a WO scan from the Scan Inbox, parse it with Claude Vision,
    and post the structured data to the Apps Script doPost handler which
    writes a new row to the Work Order Tracker and archives the PDF.

    Page-count branching:
      ≤ 2 pages: single-WO flow (original behavior — one parse, one
                 write_wo, done). Covers single-page WOs and WO+CFR.
      > 2 pages: two-pass flow — Pass 1 detects WO page groups in the
                 stack, split the PDF accordingly, then Pass 2 parses
                 each split individually. Each split is uploaded to
                 Scan Inbox (fresh file_id) so the existing archive
                 path (Apps Script archiveWOFile_) works unchanged.
    """
    from parse_work_order import parse as parse_wo, detect_wo_documents

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
        _log_scan_failure(file_id, file_name, f'download failed: {e}')
        # File may or may not exist at this point; trash defensively
        _trash_via_apps_script(file_id, label=f'failed upload {file_name}')
        return False

    # ── Page-count branch ─────────────────────────────────────────
    page_count = _pdf_page_count(file_bytes) if mime_type == 'application/pdf' else 0

    if page_count <= 2:
        # Single-WO (or single-WO + CFR) — existing flow
        log(f"  🔍  Single-WO path ({page_count or '?'} page{'s' if page_count != 1 else ''})")
        try:
            wo_data = parse_wo(file_bytes, mime_type)
        except Exception as e:
            log(f"  ❌  Parse failed: {e}")
            _log_scan_failure(file_id, file_name, f'parse failed: {e}')
            _trash_via_apps_script(file_id, label=f'failed scan {file_name}')
            return False
        if '_parse_error' in wo_data:
            _log_scan_failure(file_id, file_name, wo_data['_parse_error'])
            _trash_via_apps_script(file_id, label=f'failed scan {file_name}')
            return False
        return _write_wo_from_parsed(wo_data, file_id, original_filename=file_name)

    # ── Multi-WO: Pass 1 (detect) + split + Pass 2 (parse each) ──
    log(f"  🔍  Multi-WO path — Pass 1 detecting WO boundaries in {page_count} pages…")
    detect_result = detect_wo_documents(file_bytes)
    if '_parse_error' in detect_result:
        log(f"  ⚠️   Pass 1 detection failed: {detect_result['_parse_error']}")
        log(f"      Falling back to single-WO parse on whole PDF.")
        # Fallback: try treating the whole file as a single WO
        try:
            wo_data = parse_wo(file_bytes, mime_type)
        except Exception as e:
            _log_scan_failure(file_id, file_name, f'pass 1 + fallback both failed: {e}')
            _trash_via_apps_script(file_id, label=f'failed scan {file_name}')
            return False
        if '_parse_error' in wo_data:
            _log_scan_failure(file_id, file_name,
                              f'pass 1 failed ({detect_result["_parse_error"]}); fallback also failed ({wo_data["_parse_error"]})')
            _trash_via_apps_script(file_id, label=f'failed scan {file_name}')
            return False
        return _write_wo_from_parsed(wo_data, file_id, original_filename=file_name)

    wo_docs = detect_result.get('wo_documents') or []
    if not wo_docs:
        log(f"  ⚠️   Pass 1 returned zero WO documents — trying single-WO fallback.")
        try:
            wo_data = parse_wo(file_bytes, mime_type)
        except Exception as e:
            _log_scan_failure(file_id, file_name, f'no WOs detected; fallback failed: {e}')
            _trash_via_apps_script(file_id, label=f'failed scan {file_name}')
            return False
        if '_parse_error' in wo_data:
            _log_scan_failure(file_id, file_name,
                              f'no WOs detected; fallback parse error: {wo_data["_parse_error"]}')
            _trash_via_apps_script(file_id, label=f'failed scan {file_name}')
            return False
        return _write_wo_from_parsed(wo_data, file_id, original_filename=file_name)

    log(f"  ✅  Pass 1: {len(wo_docs)} WO document(s) detected")
    for doc in wo_docs:
        log(f"       • {doc.get('wo_id') or '?'} on pages {doc.get('pages')}")

    # Split once — returns bytes per document in the same order as wo_docs.
    try:
        split_bytes_list = split_pdf_by_pages(file_bytes, [d['pages'] for d in wo_docs])
    except Exception as e:
        log(f"  ❌  Split failed: {e}")
        _log_scan_failure(file_id, file_name, f'pdf split failed: {e}')
        _trash_via_apps_script(file_id, label=f'failed scan {file_name}')
        return False

    # Sequential Pass 2 — upload each split, parse it, write_wo.
    any_success = False
    base_name = file_name.rsplit('.', 1)[0]
    for idx, (doc, split_bytes) in enumerate(zip(wo_docs, split_bytes_list), start=1):
        hint = doc.get('wo_id') or f'part{idx}'
        split_filename = f'{base_name} — {hint}.pdf'
        log(f"  📄  Split {idx}/{len(wo_docs)}: {split_filename}")

        # Upload split → fresh file_id
        try:
            split_file_id = _upload_split_to_scan_inbox(service, split_bytes, split_filename)
        except Exception as e:
            log(f"  ❌  Split upload failed for {hint}: {e}")
            # Log against the ORIGINAL file_id so the webapp surfaces this
            # on the uploader's history item — the split didn't exist yet.
            _log_scan_failure(file_id, file_name,
                              f'split upload failed for {hint}: {e}')
            continue

        # Parse split (Pass 2, existing prompt)
        try:
            wo_data = parse_wo(split_bytes, 'application/pdf')
        except Exception as e:
            log(f"  ❌  Pass 2 parse failed for {hint}: {e}")
            # Trash the split and log against BOTH file_ids so either one
            # looked up in the webapp surfaces the error.
            _log_scan_failure(split_file_id, split_filename, f'parse failed: {e}')
            _log_scan_failure(file_id, file_name,
                              f'{hint}: parse failed: {e}')
            _trash_via_apps_script(split_file_id, label=f'failed split {hint}')
            continue
        if '_parse_error' in wo_data:
            log(f"  ❌  Pass 2 returned _parse_error for {hint}: {wo_data['_parse_error']}")
            _log_scan_failure(split_file_id, split_filename, wo_data['_parse_error'])
            _log_scan_failure(file_id, file_name,
                              f'{hint}: {wo_data["_parse_error"]}')
            _trash_via_apps_script(split_file_id, label=f'failed split {hint}')
            continue

        # Write — archive for split; combined_file_id + original_filename
        # so the webapp queue can group all splits under one upload row.
        if _write_wo_from_parsed(wo_data, split_file_id,
                                 combined_file_id=file_id,
                                 original_filename=file_name):
            any_success = True

    # Original is redundant regardless — trash it. On total failure the
    # webapp surfaces a per-file error + tells the user to re-upload.
    _trash_via_apps_script(file_id, label=f'combined PDF {file_name}')

    return any_success


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


def trash_drive_file(service, file_id: str) -> bool:
    """Trash a Drive file after we're done with it.

    On Railway the worker's `.processed_files.json` state is ephemeral —
    if the container restarts, every JSON in the watch folders gets
    processed again, producing duplicate filled PDFs. Trashing the source
    JSON right after a successful upload prevents that entirely: the next
    scan just doesn't see the file.

    Uses the Apps Script proxy on Railway (runs as the real file owner)
    and the direct Drive API locally (where OAuth creds have full access).
    """
    upload_url = os.environ.get('APPS_SCRIPT_UPLOAD_URL')
    if upload_url:
        return _trash_via_proxy(file_id, upload_url,
                                os.environ.get('APPS_SCRIPT_UPLOAD_KEY', ''))
    try:
        service.files().update(fileId=file_id, body={'trashed': True}).execute()
        return True
    except Exception as e:
        log(f"  ⚠️   Direct Drive trash failed for {file_id}: {e}")
        return False


def _trash_via_proxy(file_id: str, url: str, key: str) -> bool:
    """Ask the Apps Script proxy to trash a Drive file."""
    import requests, json as _json
    try:
        resp = requests.post(
            url,
            data=_json.dumps({'action': 'trash_file', 'key': key, 'file_id': file_id}),
            headers={'Content-Type': 'application/json'},
            timeout=30,
            allow_redirects=True,
        )
        if not resp.content.strip():
            log(f"  ⚠️   Trash proxy returned empty body (HTTP {resp.status_code})")
            return False
        result = resp.json()
        if 'error' in result:
            log(f"  ⚠️   Trash proxy error: {result['error']}")
            return False
        return True
    except Exception as e:
        log(f"  ⚠️   Trash proxy request failed: {e}")
        return False


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
        # Find the file in Drive. Ordered by modifiedTime desc so if
        # duplicates somehow exist (accidental upload without "replace"),
        # we always pick the freshest one.
        q = (f"'{templates_folder_id}' in parents and name='{filename}' "
             f"and trashed=false")
        result = service.files().list(
            q=q,
            fields='files(id,name,modifiedTime,size)',
            orderBy='modifiedTime desc'
        ).execute()
        files = result.get('files', [])
        if len(files) > 1:
            log(f"  ⚠️   Found {len(files)} copies of {filename} in Templates — using newest")

        if files:
            drive_file = files[0]
            drive_modified = drive_file['modifiedTime']
            drive_size     = int(drive_file.get('size') or 0)
            drive_id       = drive_file['id']

            # Persist the Drive file-id alongside the cached PDF so we
            # can detect a file swap (user deleted the old template and
            # uploaded a replacement → new file-id, possibly-older
            # modifiedTime if Drive preserves creation time). Without
            # this, swapping the template back to an older version
            # silently keeps using the stale local cache forever.
            marker_path = local_path.with_suffix(local_path.suffix + '.drive_id')
            cached_id   = marker_path.read_text().strip() if marker_path.exists() else None

            # Size check catches any content change — most reliable
            # busting signal when mtime is the same but bytes differ.
            local_size = local_path.stat().st_size if local_path.exists() else -1

            import email.utils, calendar
            drive_ts = 0
            if local_path.exists():
                drive_ts = calendar.timegm(
                    time.strptime(drive_modified, '%Y-%m-%dT%H:%M:%S.%fZ')
                    if '.' in drive_modified else
                    time.strptime(drive_modified, '%Y-%m-%dT%H:%M:%SZ')
                )
                local_mtime = local_path.stat().st_mtime
            else:
                local_mtime = 0

            reasons = []
            if not local_path.exists():                reasons.append('no local cache')
            if cached_id and cached_id != drive_id:    reasons.append(f'drive file-id changed ({cached_id[:8]}→{drive_id[:8]})')
            if drive_size and drive_size != local_size: reasons.append(f'size changed ({local_size}→{drive_size})')
            if drive_ts > local_mtime:                 reasons.append('drive newer')
            needs_download = bool(reasons)

            if needs_download:
                log(f"  📥  Downloading template from Drive: {filename} ({', '.join(reasons)})")
                request = service.files().get_media(fileId=drive_id)
                buf = io.BytesIO()
                dl = MediaIoBaseDownload(buf, request)
                done = False
                while not done:
                    _, done = dl.next_chunk()
                local_path.write_bytes(buf.getvalue())
                marker_path.write_text(drive_id)
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
# Fillers accept an optional `source_name` — the source JSON's filename —
# so we can build PDF filenames like `<JSON stem>_FILLED.pdf`.  The JSON
# stems are authored by Apps Script with every archive-relevant field
# already baked in (`Production_Log_YYYY-MM-DD.json`,
# `Certified_Payroll_{contractNum}_{borough}_YYYY-MM-DD.json`, etc).
# Reusing them keeps the PDF filenames parseable by archiveDocument_ in
# Apps Script without duplicating the filename-generation logic here.
def _filename_from_source(source_name: str | None, fallback: str) -> str:
    if source_name:
        stem = source_name
        if stem.lower().endswith('.json'):
            stem = stem[:-5]
        return f"{stem}_FILLED.pdf"
    return fallback


def fill_production_log(service, data: dict, tmp_dir: Path, source_name: str | None = None) -> Path:
    from fill_production_log import fill
    template = get_template(service, 'production_log')
    date_str = data.get('date', 'unknown').replace('/', '-')
    fallback = f"Production_Log_{date_str}_FILLED.pdf"
    out = tmp_dir / _filename_from_source(source_name, fallback)
    fill(data, template_path=str(template), output_path=str(out))
    return out


def fill_certified_payroll(service, data: dict, tmp_dir: Path, source_name: str | None = None) -> Path:
    from fill_certified_payroll import fill
    template = get_template(service, 'certified_payroll')
    week = data.get('week_ending', 'unknown').replace('/', '-')
    fallback = f"Certified_Payroll_{week}_FILLED.pdf"
    out = tmp_dir / _filename_from_source(source_name, fallback)
    fill(data, template_path=str(template), output_path=str(out))
    return out


def fill_signin(service, data: dict, tmp_dir: Path, source_name: str | None = None) -> Path:
    from fill_signin import fill
    template = get_template(service, 'signin')
    wo_id    = data.get('wo_id', 'unknown')
    date_str = (data.get('date', 'unknown')
                .replace('/', '-').replace(' ', '_'))
    fallback = f"SignIn_{wo_id}_{date_str}_FILLED.pdf"
    out = tmp_dir / _filename_from_source(source_name, fallback)
    fill(data, template_path=str(template), output_path=str(out))
    return out


def fill_contractor_field_report(service, data: dict, tmp_dir: Path, source_name: str | None = None) -> Path:
    from fill_contractor_field_report import fill
    template = get_template(service, 'contractor_field_report')
    wo_id    = data.get('wo_id') or data.get('work_order', 'unknown')
    date_str = str(data.get('install_to') or data.get('date_entered') or 'unknown') \
                  .replace('/', '-').replace(' ', '_')
    fallback = f"Contractor_Field_Report_{wo_id}_{date_str}_FILLED.pdf"
    out = tmp_dir / _filename_from_source(source_name, fallback)
    fill(data, template_path=str(template), output_path=str(out))
    return out


FILLERS = {
    'production_log':           fill_production_log,
    'certified_payroll':        fill_certified_payroll,
    'signin':                   fill_signin,
    'contractor_field_report':  fill_contractor_field_report,
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
        # Pass the source JSON's filename so the filler can name the
        # output PDF `<JSON stem>_FILLED.pdf` — the Apps Script archive
        # step regex-parses those stems for contract/borough/date.
        pdf_path = filler(service, data, tmp_dir, source_name=name)
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

    # Trash the source JSON so we never re-process it on the next poll,
    # even if the processed-files state gets lost (Railway restart).
    if trash_drive_file(service, file_id):
        log(f"  🗑️   Trashed source JSON: {name}")
    else:
        log(f"  ⚠️   Could not trash source JSON {name} — may re-process on restart")

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

    Per-subfolder env overrides (optional — otherwise discovered from Drive):
      DRIVE_PROD_LOGS_ID        — Production Logs
      DRIVE_CERT_PAYROLL_ID     — Certified Payroll
      DRIVE_SIGN_IN_LOGS_ID     — Sign-In Logs
      DRIVE_FIELD_REPORTS_ID    — Field Reports

    Returns:
      (folder_map, scan_inbox_id) where folder_map is {subfolder_name: id}
      and scan_inbox_id is the Drive ID of the Scan Inbox (or None).
    """
    # Maps WATCH_SUBFOLDER name → corresponding env var override
    SUBFOLDER_ENV = {
        'Production Logs':   'DRIVE_PROD_LOGS_ID',
        'Certified Payroll': 'DRIVE_CERT_PAYROLL_ID',
        'Sign-In Logs':      'DRIVE_SIGN_IN_LOGS_ID',
        'Field Reports':     'DRIVE_FIELD_REPORTS_ID',
    }

    needs_review_id     = os.environ.get('DRIVE_NEEDS_REVIEW_ID')
    templates_folder_id = os.environ.get('DRIVE_TEMPLATES_ID')
    scan_inbox_id       = os.environ.get('DRIVE_SCAN_INBOX_ID')

    # Start with env overrides for each watched subfolder
    subfolder_ids: dict[str, str | None] = {
        sub: os.environ.get(env_var)
        for sub, env_var in SUBFOLDER_ENV.items()
    }

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
        for sub in SUBFOLDER_ENV:
            if not subfolder_ids[sub]:
                subfolder_ids[sub] = saved_folders.get(sub)
    else:
        log("📋  Loaded folder IDs from environment variables.")

    # Build folder map — discover any subfolders not yet known
    folder_map: dict[str, str] = {}
    for sub in WATCH_SUBFOLDERS:
        known = subfolder_ids.get(sub)
        if known:
            folder_map[sub] = known
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

    return (folder_map, scan_inbox_id)


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
