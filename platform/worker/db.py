"""
Database client for the PDF worker.

Used to update document status and job results after fill operations.
Connects directly to PostgreSQL (same DB as the API server).
"""

import os
import psycopg2
from contextlib import contextmanager
from datetime import datetime

DATABASE_URL = os.environ.get('DATABASE_URL', '')


@contextmanager
def get_conn():
    """Context manager for database connections."""
    conn = psycopg2.connect(DATABASE_URL)
    try:
        yield conn
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def update_document_status(doc_id: str, status: str, storage_key: str | None = None, filename: str | None = None):
    """Update a document's status and optional storage key after fill."""
    with get_conn() as conn:
        cur = conn.cursor()
        updates = ["status = %s", "updated_at = %s"]
        params = [status, datetime.utcnow()]

        if storage_key:
            updates.append("storage_key = %s")
            params.append(storage_key)
        if filename:
            updates.append("filename = %s")
            params.append(filename)

        params.append(doc_id)
        cur.execute(
            f"UPDATE documents SET {', '.join(updates)} WHERE id = %s",
            params,
        )


def update_job_status(job_id: str, status: str, result: dict | None = None, error: str | None = None):
    """Update a job's status after processing."""
    import json
    with get_conn() as conn:
        cur = conn.cursor()
        now = datetime.utcnow()
        if status in ('completed', 'failed'):
            cur.execute(
                "UPDATE jobs SET status = %s, result = %s, error = %s, completed_at = %s, attempts = attempts + 1 WHERE id = %s",
                (status, json.dumps(result) if result else None, error, now, job_id),
            )
        else:
            cur.execute(
                "UPDATE jobs SET status = %s, attempts = attempts + 1 WHERE id = %s",
                (status, job_id),
            )


def create_audit_entry(org_id: str, source: str, action: str, subject: str | None = None, status: str | None = None):
    """Log an audit entry from the worker."""
    with get_conn() as conn:
        cur = conn.cursor()
        cur.execute(
            "INSERT INTO audit_log (org_id, source, action, subject, status) VALUES (%s, %s, %s, %s, %s)",
            (org_id, source, action, subject, status),
        )
