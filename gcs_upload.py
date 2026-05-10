"""
gcs_upload.py — Google Cloud Storage helper for large file uploads.

Bypasses Cloud Run's 32 MB HTTP request limit by uploading files directly
from the browser to GCS via signed URLs.

Flow:
  1. Server generates a signed PUT URL + blob name
  2. Browser uploads directly to GCS (no Cloud Run limit!)
  3. Server downloads from GCS blob → BytesIO for processing
  4. Server deletes the blob after processing

Environment variables:
  GCS_UPLOAD_BUCKET  – Required. Name of the GCS bucket.
  GOOGLE_CLOUD_PROJECT – Usually auto-detected on Cloud Run.
"""

import os
import io
import uuid
import logging
from datetime import timedelta

logger = logging.getLogger(__name__)

# ── Configuration ──────────────────────────────────────────────────────────
GCS_BUCKET = os.getenv("GCS_UPLOAD_BUCKET", "")
GCS_ENABLED = bool(GCS_BUCKET)

_client = None  # Lazy singleton


def _get_client():
    """Lazy-init GCS client (reuses across calls)."""
    global _client
    if _client is None:
        from google.cloud import storage
        _client = storage.Client()
    return _client


# ── Core functions ─────────────────────────────────────────────────────────

def generate_upload_url(filename: str,
                        content_type: str = "application/pdf",
                        expiration_min: int = 60) -> tuple[str, str]:
    """
    Generate a V4 signed URL for direct browser → GCS upload.

    Returns:
        (signed_url, blob_name)
    """
    client = _get_client()
    bucket = client.bucket(GCS_BUCKET)
    blob_name = f"uploads/{uuid.uuid4().hex[:12]}/{filename}"
    blob = bucket.blob(blob_name)

    url = blob.generate_signed_url(
        version="v4",
        expiration=timedelta(minutes=expiration_min),
        method="PUT",
        content_type=content_type,
    )
    logger.info(f"Generated signed upload URL for {blob_name}")
    return url, blob_name


def download_blob(blob_name: str) -> io.BytesIO:
    """Download a blob and return its content as a seekable BytesIO."""
    client = _get_client()
    bucket = client.bucket(GCS_BUCKET)
    blob = bucket.blob(blob_name)
    data = blob.download_as_bytes()
    buf = io.BytesIO(data)
    buf.name = blob_name.rsplit("/", 1)[-1]  # keep original filename
    logger.info(f"Downloaded {blob_name} ({len(data):,} bytes)")
    return buf


def blob_exists(blob_name: str) -> bool:
    """Check whether a blob exists in the bucket."""
    client = _get_client()
    bucket = client.bucket(GCS_BUCKET)
    return bucket.blob(blob_name).exists()


def delete_blob(blob_name: str):
    """Delete a single blob (silent on errors)."""
    try:
        client = _get_client()
        bucket = client.bucket(GCS_BUCKET)
        bucket.blob(blob_name).delete()
        logger.info(f"Deleted {blob_name}")
    except Exception as e:
        logger.warning(f"Could not delete {blob_name}: {e}")


def cleanup_blobs(blob_names: list[str]):
    """Delete a list of blobs."""
    for name in (blob_names or []):
        if name:
            delete_blob(name)


def get_blob_size(blob_name: str) -> int:
    """Return the size in bytes of a blob, or 0 if it doesn't exist."""
    try:
        client = _get_client()
        bucket = client.bucket(GCS_BUCKET)
        blob = bucket.blob(blob_name)
        blob.reload()
        return blob.size or 0
    except Exception:
        return 0
