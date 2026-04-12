"""
parse_work_order.py — Claude Vision API parser for scanned Work Order PDFs
===========================================================================
Called by watch_and_fill.py when a new file lands in the Drive Scan Inbox folder.

Flow:
  1. watch_and_fill.py downloads the scanned PDF/image from Drive
  2. Calls parse(file_bytes, mime_type) here
  3. Returns a structured dict ready to write to the Work Order Tracker sheet

Environment variables required:
  ANTHROPIC_API_KEY  — Claude API key (set in Railway dashboard)
"""

import base64
import json
import os
import logging

log = logging.getLogger(__name__)

# Work Order fields we extract from the scanned form
WO_FIELDS = [
    'work_order_id',
    'contractor',
    'contract_number',
    'borough',
    'location',
    'from_street',
    'to_street',
    'due_date',
    'pavement_work_type',
    'marking_types',   # list of {type, sqft}
]

EXTRACTION_PROMPT = """You are parsing a scanned NYC DOT Work Order Report form for a pavement marking subcontractor.

Extract the following fields from the image and return ONLY valid JSON — no explanation, no markdown, just the JSON object:

{
  "work_order_id":       "string — the Work Order ID / PT number (e.g. PT-11930)",
  "contractor":          "string — the prime contractor name",
  "contract_number":     "string — contract number (e.g. 84122MBTP496)",
  "borough":             "string — borough name or abbreviation",
  "location":            "string — street or location name",
  "from_street":         "string — from street",
  "to_street":           "string — to street",
  "due_date":            "string — due date as written on the form",
  "pavement_work_type":  "string — type of pavement work",
  "marking_types":       [{"type": "string", "sqft": "string or number"}],
  "notes":               "string — any other relevant notes or remarks"
}

If a field is not visible or not applicable, use null. Do not guess.
"""


def parse(file_bytes: bytes, mime_type: str = 'application/pdf') -> dict:
    """
    Send a scanned WO file to Claude Vision API and return structured data.

    Args:
        file_bytes: raw bytes of the PDF or image file
        mime_type:  MIME type — 'application/pdf', 'image/jpeg', 'image/png', etc.

    Returns:
        dict with extracted work order fields, or {'_parse_error': reason} on failure
    """
    try:
        import anthropic
    except ImportError:
        raise RuntimeError("anthropic package not installed — run: pip install anthropic")

    api_key = os.environ.get('ANTHROPIC_API_KEY')
    if not api_key:
        raise RuntimeError("ANTHROPIC_API_KEY environment variable not set")

    client = anthropic.Anthropic(api_key=api_key)

    # Claude accepts PDFs and images as base64-encoded content blocks
    encoded = base64.standard_b64encode(file_bytes).decode('utf-8')

    # Use 'document' type for PDFs, 'image' type for image files
    if mime_type == 'application/pdf':
        content_block = {
            "type": "document",
            "source": {
                "type": "base64",
                "media_type": mime_type,
                "data": encoded,
            }
        }
    else:
        content_block = {
            "type": "image",
            "source": {
                "type": "base64",
                "media_type": mime_type,
                "data": encoded,
            }
        }

    try:
        response = client.messages.create(
            model="claude-opus-4-5",
            max_tokens=1024,
            messages=[
                {
                    "role": "user",
                    "content": [
                        content_block,
                        {"type": "text", "text": EXTRACTION_PROMPT}
                    ]
                }
            ]
        )

        raw = response.content[0].text.strip()

        # Strip markdown code fences if present
        if raw.startswith('```'):
            raw = raw.split('```')[1]
            if raw.startswith('json'):
                raw = raw[4:]

        return json.loads(raw)

    except json.JSONDecodeError as e:
        log.error(f"Claude returned non-JSON response: {e}")
        return {'_parse_error': f'JSON decode failed: {e}', '_raw': raw}
    except Exception as e:
        log.error(f"Claude API call failed: {e}")
        return {'_parse_error': str(e)}
