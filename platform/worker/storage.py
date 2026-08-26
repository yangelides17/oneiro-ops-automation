"""
R2/S3 storage client for the PDF worker.

Replaces Google Drive API calls in the old workers/ code.
All files are stored under {org_id}/{path} for tenant isolation.
"""

import os
import boto3
from pathlib import Path

_client = None
BUCKET = os.environ.get('R2_BUCKET_NAME', 'oneiro-platform')


def _get_client():
    global _client
    if _client is None:
        _client = boto3.client(
            's3',
            endpoint_url=f"https://{os.environ['R2_ACCOUNT_ID']}.r2.cloudflarestorage.com",
            aws_access_key_id=os.environ['R2_ACCESS_KEY_ID'],
            aws_secret_access_key=os.environ['R2_SECRET_ACCESS_KEY'],
            region_name='auto',
        )
    return _client


def download(storage_key: str) -> bytes:
    """Download a file from R2 by its full storage key."""
    resp = _get_client().get_object(Bucket=BUCKET, Key=storage_key)
    return resp['Body'].read()


def upload(org_id: str, path: str, data: bytes, content_type: str = 'application/pdf') -> str:
    """Upload a file to R2. Returns the full storage key."""
    key = f"{org_id}/{path}"
    _get_client().put_object(
        Bucket=BUCKET,
        Key=key,
        Body=data,
        ContentType=content_type,
    )
    return key


def download_to_file(storage_key: str, local_path: str):
    """Download a file from R2 to a local path."""
    data = download(storage_key)
    Path(local_path).parent.mkdir(parents=True, exist_ok=True)
    with open(local_path, 'wb') as f:
        f.write(data)


def upload_from_file(org_id: str, path: str, local_path: str, content_type: str = 'application/pdf') -> str:
    """Upload a local file to R2. Returns the full storage key."""
    with open(local_path, 'rb') as f:
        return upload(org_id, path, f.read(), content_type)
