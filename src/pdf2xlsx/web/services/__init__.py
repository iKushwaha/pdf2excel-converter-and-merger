"""Service layer for the web app — orchestrates the core engine.

Each public function operates on a ``SessionManager`` plus a session id and
returns plain data (JSON-serialisable dicts) so the route layer stays thin.
"""

import os
import time

from ...core import converter, extractor, merger
from .session import SessionManager


def _unique_output(out_dir, base_stem):
    """Return a non-colliding output path with a timestamp suffix."""
    stamp = time.strftime("%Y%m%d-%H%M%S")
    return os.path.join(out_dir, f"{base_stem}_{stamp}.xlsx")


def convert_upload(session: SessionManager, sid: str, upload, password=None):
    """Convert a single uploaded PDF. Returns a JSON-safe status dict."""
    import logging
    _log = logging.getLogger(__name__)

    try:
        result = converter.convert_pdf_to_excel(
            upload["path"], session.output_dir(sid),
            password=password or None)
        return {
            "ok": True,
            "name": os.path.basename(result["path"]),
            "sheets": result["sheets"],
            "warnings": result["warnings"],
        }
    except converter.EncryptedPDF:
        return {"ok": False, "error": "encrypted", "message": "Password required."}
    except converter.ConversionCancelled:
        return {"ok": False, "error": "cancelled", "message": "Cancelled."}
    except converter.ConversionError:
        return {"ok": False, "error": "conversion",
                "message": "Conversion failed. The file may be corrupted or unsupported."}
    except Exception:
        _log.exception("Unexpected error converting %s", upload.get("name", "?"))
        return {"ok": False, "error": "unexpected",
                "message": "An unexpected error occurred during conversion."}


def merge_outputs(session: SessionManager, sid: str, filenames=None, mode="sheets"):
    """Merge the given output workbooks in the session output dir.

    ``filenames`` are workbook names the user wants merged; empty means
    "merge everything generated so far".
    """
    out_dir = session.output_dir(sid)
    candidates = sorted(n for n in os.listdir(out_dir)
                        if n.endswith(".xlsx") and not n.startswith("._"))
    if filenames:
        selected = [n for n in candidates if n in filenames]
    else:
        selected = candidates
    if not selected:
        raise ValueError("No converted files to merge.")

    paths = [os.path.join(out_dir, n) for n in selected]
    out_path = _unique_output(out_dir, "Merged")
    result = merger.merge_excels(paths, out_path, mode=mode)
    return {
        "name": os.path.basename(result["path"]),
        "sheets": result["sheets"],
        "notes": result["notes"],
    }


def extract_fields(session: SessionManager, sid: str, fields, filenames=None):
    """Extract the requested field keys from session output workbooks."""
    out_dir = session.output_dir(sid)
    candidates = sorted(n for n in os.listdir(out_dir)
                        if n.endswith(".xlsx") and not n.startswith("._"))
    if not candidates:
        raise ValueError("No converted files to extract from.")
    if filenames:
        candidates = [n for n in candidates if n in filenames]

    paths = [os.path.join(out_dir, n) for n in candidates]
    out_path = _unique_output(out_dir, "Extracted_Fields")
    result = extractor.extract_fields(paths, out_path, fields)
    return {
        "name": os.path.basename(result["path"]),
        "rows": result["rows"],
        "columns": result["columns"],
        "notes": result["notes"],
    }


def list_outputs(session: SessionManager, sid: str):
    """Return generated workbook info for the session output dir."""
    out_dir = session.output_dir(sid)
    items = []
    for name in sorted(os.listdir(out_dir)):
        if name.startswith("._") or not name.endswith(".xlsx"):
            continue
        path = os.path.join(out_dir, name)
        items.append({
            "name": name,
            "size": os.path.getsize(path),
        })
    return items


def ALLOWED_FIELDS():
    return extractor.FIELD_LABELS


def validate_fields(fields):
    keys = {k for k, _ in extractor.FIELD_LABELS}
    valid = [f for f in fields if f in keys]
    if not valid:
        raise ValueError(
            "Select at least one field (material_code, item_description, "
            "ean, quantity, unit_cost).")
    return valid
