#!/usr/bin/env python3
"""
main.py  —  Oneiro Platform PDF Worker
=======================================
Polls the PostgreSQL `jobs` table for pending fill jobs and processes them
using the same fill modules from the original workers/ directory.

Architecture: database-backed job queue (no Redis dependency).
The Node API inserts rows into `jobs` with status='pending'.
This worker picks them up, fills the PDF, uploads to R2, and
updates the document status.

Usage:
  python3 main.py                    # Run the poll loop
  python3 main.py --once             # Process one batch and exit (for testing)
"""

import json
import os
import sys
import time
import tempfile
import traceback
from pathlib import Path
from datetime import datetime

# ── CONFIGURATION ────────────────────────────────────────────────────
WORKER_DIR = Path(__file__).parent
POLL_INTERVAL = int(os.environ.get('WORKER_POLL_INTERVAL', '5'))  # seconds
BATCH_SIZE = int(os.environ.get('WORKER_BATCH_SIZE', '10'))
MAX_ATTEMPTS = int(os.environ.get('WORKER_MAX_ATTEMPTS', '3'))

# Templates: check local templates/ directory at repo root
# Worker is at platform/worker/, templates are at <repo>/templates/
TEMPLATES_DIR = WORKER_DIR.parent.parent / 'templates'

# ── TEMPLATE MAPPING ────────────────────────────────────────────────
# Maps fill _type → actual template filename on disk
TEMPLATE_FILES = {
    'production_log':           'Metro_Production_Log_Fillable.pdf',
    'certified_payroll':        'Certified_Payroll_Fillable.pdf',
    'signin':                   'Sign_In_Log_Fillable.pdf',
    'contractor_field_report':  'Thermo_Contractor_Field_Report_Fillable.pdf',
    'employee_utilization':     'Monthly_Workforce_Utilization_Fillable.pdf',
}

# ── FILL MODULE IMPORTS ──────────────────────────────────────────────
from fill_production_log import fill as fill_production_log
from fill_certified_payroll import fill as fill_certified_payroll
from fill_signin import fill as fill_signin
from fill_contractor_field_report import fill as fill_contractor_field_report
from fill_utilization import fill as fill_utilization

import storage
import db


def log(msg):
    print(f"[{datetime.now().strftime('%H:%M:%S')}] {msg}", flush=True)


# ── TEMPLATE RESOLUTION ─────────────────────────────────────────────

def get_template_path(doc_type: str) -> str:
    """Resolve the template PDF path for a given document type."""
    filename = TEMPLATE_FILES.get(doc_type)
    if not filename:
        raise ValueError(f"Unknown document type: {doc_type}")

    path = TEMPLATES_DIR / filename
    if not path.exists():
        raise FileNotFoundError(
            f"Template not found: {path}\n"
            f"Expected templates in: {TEMPLATES_DIR}"
        )
    return str(path)


# ── JOB PROCESSOR ───────────────────────────────────────────────────

def process_fill_job(job_id: str, job_data: dict) -> dict:
    """
    Process a single PDF fill job.

    1. Resolve template
    2. Call the appropriate fill module
    3. Upload filled PDF to R2
    4. Update document status in DB
    """
    org_id = job_data['orgId']
    doc_id = job_data.get('documentId', '')
    fill_data = job_data.get('fillData', {})
    doc_type = fill_data.get('_type', '')
    overwrite_key = job_data.get('overwriteStorageKey')

    log(f"  Processing: type={doc_type} doc={doc_id}")

    # Resolve template
    template_path = get_template_path(doc_type)

    # Create temp output file
    with tempfile.NamedTemporaryFile(suffix='.pdf', delete=False) as tmp:
        output_path = tmp.name

    try:
        # Call the appropriate filler
        if doc_type == 'production_log':
            fill_production_log(fill_data, template_path, output_path)
        elif doc_type == 'certified_payroll':
            fill_certified_payroll(fill_data, template_path, output_path)
        elif doc_type == 'signin':
            fill_signin(fill_data, template_path, output_path)
        elif doc_type == 'contractor_field_report':
            fill_contractor_field_report(fill_data, template_path, output_path)
        elif doc_type == 'employee_utilization':
            fill_utilization(fill_data, template_path, output_path)
        else:
            raise ValueError(f"Unknown fill type: {doc_type}")

        # Upload filled PDF to R2
        r2_path = overwrite_key or f"documents/{doc_type}/{doc_id}_{int(time.time())}.pdf"
        if overwrite_key and overwrite_key.startswith(f"{org_id}/"):
            r2_path = overwrite_key[len(f"{org_id}/"):]

        storage_key = storage.upload_from_file(org_id, r2_path, output_path)
        log(f"  Uploaded: {storage_key}")

        # Build filename for the document record
        date_part = fill_data.get('date', fill_data.get('week_ending', ''))
        filename = f"{doc_type}_{date_part}_FILLED.pdf"

        # Update document status in DB
        if doc_id:
            db.update_document_status(doc_id, 'needs_review', storage_key, filename)

        return {'storage_key': storage_key, 'filename': filename}

    finally:
        try:
            os.unlink(output_path)
        except OSError:
            pass


# ── DATABASE JOB CONSUMER ──────────────────────────────────────────

FILL_JOB_TYPES = {
    'fill_production_log',
    'fill_certified_payroll',
    'fill_signin',
    'fill_field_report',
    'fill_month_end',
}


def poll_and_process():
    """
    Poll the jobs table for pending fill jobs, claim them, and process.

    Uses SELECT ... FOR UPDATE SKIP LOCKED for safe concurrent polling
    (multiple workers can run without stepping on each other).
    """
    from db import get_conn

    with get_conn() as conn:
        cur = conn.cursor()

        # Claim a batch of pending jobs atomically
        cur.execute("""
            UPDATE jobs
            SET status = 'processing', attempts = attempts + 1
            WHERE id IN (
                SELECT id FROM jobs
                WHERE status = 'pending'
                  AND type LIKE 'fill_%%'
                  AND attempts < %s
                ORDER BY created_at ASC
                LIMIT %s
                FOR UPDATE SKIP LOCKED
            )
            RETURNING id, type, payload
        """, (MAX_ATTEMPTS, BATCH_SIZE))

        claimed = cur.fetchall()

    if not claimed:
        return 0

    log(f"Claimed {len(claimed)} job(s)")

    processed = 0
    for job_id, job_type, payload_raw in claimed:
        payload = payload_raw if isinstance(payload_raw, dict) else json.loads(payload_raw)

        try:
            result = process_fill_job(job_id, payload)

            # Mark completed
            with get_conn() as conn:
                cur = conn.cursor()
                cur.execute("""
                    UPDATE jobs
                    SET status = 'completed',
                        result = %s,
                        completed_at = NOW()
                    WHERE id = %s
                """, (json.dumps(result), job_id))

            log(f"  ✓ Job {job_id} completed")
            processed += 1

        except Exception as e:
            error_msg = f"{type(e).__name__}: {e}"
            log(f"  ✗ Job {job_id} failed: {error_msg}")

            # Check if max attempts reached
            with get_conn() as conn:
                cur = conn.cursor()
                cur.execute("SELECT attempts FROM jobs WHERE id = %s", (job_id,))
                row = cur.fetchone()
                attempts = row[0] if row else MAX_ATTEMPTS

                if attempts >= MAX_ATTEMPTS:
                    cur.execute("""
                        UPDATE jobs
                        SET status = 'failed',
                            error = %s,
                            completed_at = NOW()
                        WHERE id = %s
                    """, (error_msg, job_id))
                    log(f"  → Max attempts reached, marked as failed")
                else:
                    # Return to pending for retry
                    cur.execute("""
                        UPDATE jobs
                        SET status = 'pending',
                            error = %s
                        WHERE id = %s
                    """, (error_msg, job_id))
                    log(f"  → Returned to pending (attempt {attempts}/{MAX_ATTEMPTS})")

    return processed


def main():
    """Main entry point — poll loop or single batch."""
    log("Oneiro PDF Worker starting")
    log(f"  Templates: {TEMPLATES_DIR}")
    log(f"  Poll interval: {POLL_INTERVAL}s")
    log(f"  Batch size: {BATCH_SIZE}")
    log(f"  Max attempts: {MAX_ATTEMPTS}")

    # Verify templates exist
    for doc_type, filename in TEMPLATE_FILES.items():
        path = TEMPLATES_DIR / filename
        status = "✓" if path.exists() else "✗ MISSING"
        log(f"  Template {doc_type}: {status}")

    # Verify DB connection
    try:
        from db import get_conn
        with get_conn() as conn:
            cur = conn.cursor()
            cur.execute("SELECT 1")
        log("  Database: ✓ connected")
    except Exception as e:
        log(f"  Database: ✗ {e}")
        sys.exit(1)

    # Verify R2 connection
    try:
        storage._get_client().head_bucket(Bucket=storage.BUCKET)
        log("  R2 Storage: ✓ connected")
    except Exception as e:
        log(f"  R2 Storage: ✗ {e}")
        sys.exit(1)

    single_run = '--once' in sys.argv

    if single_run:
        count = poll_and_process()
        log(f"Processed {count} job(s). Exiting.")
        return

    log("Entering poll loop...")
    while True:
        try:
            poll_and_process()
        except Exception as e:
            log(f"Poll error: {e}")
            traceback.print_exc()
        time.sleep(POLL_INTERVAL)


if __name__ == '__main__':
    main()
