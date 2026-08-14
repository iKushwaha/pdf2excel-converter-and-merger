"""Web app configuration (read from environment with sane defaults)."""

import os
import tempfile


def _as_int(name, default):
    try:
        return int(os.environ.get(name, default))
    except (TypeError, ValueError):
        return default


class Config:
    """Flask configuration values, overridable via environment variables."""

    # Where uploaded PDFs and generated workbooks are stored per session.
    # Shared across processes; sessions are cleaned by age.
    DATA_DIR = os.environ.get("PDF2EXCEL_DATA_DIR",
                              os.path.join(tempfile.gettempdir(), "pdf2excel-web"))

    # Hard cap on a single upload (bytes). Default 100 MB.
    MAX_CONTENT_LENGTH = _as_int("PDF2EXCEL_MAX_UPLOAD", 100 * 1024 * 1024)

    # Max number of files in one upload request.
    MAX_FILES_PER_UPLOAD = _as_int("PDF2EXCEL_MAX_FILES", 50)

    # Session lifetime (seconds) before the cleanup sweep removes it.
    SESSION_TTL_SECONDS = _as_int("PDF2EXCEL_SESSION_TTL", 60 * 60 * 24)

    # How often the background cleanup sweep runs (seconds).
    CLEANUP_INTERVAL_SECONDS = _as_int("PDF2EXCEL_CLEANUP_INTERVAL", 60 * 30)

    SECRET_KEY = os.environ.get("PDF2EXCEL_SECRET_KEY", "dev-only-secret-key")

    # Accepted upload extensions (case-insensitive).
    ALLOWED_EXTENSIONS = {".pdf"}
