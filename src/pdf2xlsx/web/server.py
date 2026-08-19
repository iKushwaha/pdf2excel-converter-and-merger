"""Flask application factory for the PDF → Excel web app.

The web app is a server-side alternative to the desktop app: browsers
upload PDFs, the server converts them on the host machine with the same
core engine, and the resulting Excel files are downloaded back. Files are
kept in per-session temp directories and auto-expired.
"""

import logging
import os
import time

from flask import (Flask, jsonify, render_template, request, send_file,
                   session as flask_session)
from flask_limiter import Limiter
from flask_limiter.util import get_remote_address

from . import config
from .services import (ALLOWED_FIELDS, convert_upload, extract_fields,
                       list_outputs, merge_outputs, validate_fields)
from .services.session import SessionManager

MAX_UPLOAD = config.Config.MAX_FILES_PER_UPLOAD
PDF_MAGIC = config.Config.PDF_MAGIC

# Structured security logger — separate from app logs for easy filtering.
_sec_log = logging.getLogger("pdf2excel.security")


def create_app(data_dir=None):
    """Build the Flask app; ``data_dir`` overrides Config.DATA_DIR (tests)."""
    app = Flask(__name__,
                template_folder=os.path.join(os.path.dirname(__file__),
                                             "templates"),
                static_folder=os.path.join(os.path.dirname(__file__),
                                           "static"))
    app.config.from_object(config.Config)
    if data_dir is not None:
        app.config["DATA_DIR"] = data_dir
        app.config["TESTING"] = True

    # ---- Rate limiter (per-IP, in-memory store) ----------------------------
    limiter = Limiter(
        key_func=get_remote_address,
        app=app,
        default_limits=[f"{config.Config.RATE_LIMIT_PER_MINUTE}/minute"],
        storage_uri="memory://",
    )

    sessions = SessionManager(app.config["DATA_DIR"])

    # ---- CSRF (double-submit cookie pattern) -------------------------------
    @app.before_request
    def _csrf_protect():
        """Enforce CSRF on state-changing requests (POST/PUT/DELETE).

        Uses a double-submit cookie: the client must send a ``X-CSRF-Token``
        header whose value matches the ``_csrf_token`` session cookie.  The
        token is set automatically on the first GET to any page.
        """
        if request.method not in ("POST", "PUT", "DELETE", "PATCH"):
            return None
        # Skip CSRF for the health-check endpoint (GET-only) and tests.
        if request.endpoint == "health" or app.config.get("TESTING"):
            return None
        token = request.headers.get("X-CSRF-Token", "")
        expected = flask_session.get("_csrf_token", "")
        if not expected or not token or token != expected:
            _sec_log.warning("CSRF token mismatch from %s on %s",
                             request.remote_addr, request.path)
            return jsonify({"ok": False, "error": "Invalid or missing CSRF token."}), 403
        return None

    def _ensure_csrf():
        """Set a CSRF token in the session if one doesn't exist yet."""
        if "_csrf_token" not in flask_session:
            import secrets as _secrets
            flask_session["_csrf_token"] = _secrets.token_hex(32)

    # ---- Content-Security-Policy header ------------------------------------
    @app.after_request
    def _set_security_headers(response):
        response.headers["Content-Security-Policy"] = (
            "default-src 'self'; "
            "script-src 'self'; "
            "style-src 'self' 'unsafe-inline'; "
            "img-src 'self' data:; "
            "connect-src 'self'; "
            "font-src 'self'"
        )
        response.headers["X-Content-Type-Options"] = "nosniff"
        response.headers["X-Frame-Options"] = "DENY"
        response.headers["Referrer-Policy"] = "strict-origin-when-cross-origin"
        return response

    # ---- Session helpers ---------------------------------------------------
    def _current_session():
        """Return the session id, creating one on first visit.

        The session id is rotated on first use to prevent session fixation.
        """
        _ensure_csrf()
        if "sid" not in flask_session or not sessions.is_valid(flask_session["sid"]):
            flask_session.clear()
            flask_session["sid"] = sessions.new_session()
            _ensure_csrf()
            _sec_log.info("New session created: %s from %s",
                          flask_session["sid"][:8], request.remote_addr)
        sessions.touch(flask_session["sid"])
        return flask_session["sid"]

    # -------------------------------------------------------------- pages
    @app.route("/")
    def index():
        return render_template("index.html",
                               landing_url=app.config.get("LANDING_URL", ""))

    @app.route("/health")
    @limiter.exempt
    def health():
        return jsonify({"ok": True, "name": "pdf2excel-web"})

    @app.route("/api/fields")
    def api_fields():
        return jsonify({"fields": [{"key": k, "label": v} for k, v in ALLOWED_FIELDS()]})

    # ------------------------------------------------------------ uploads
    @app.post("/api/upload")
    def api_upload():
        _log = logging.getLogger(__name__)

        sid = _current_session()
        files = request.files.getlist("pdfs")
        files = [f for f in files if f and f.filename]
        if not files:
            return jsonify({"ok": False, "error": "No files selected."}), 400
        if len(files) > MAX_UPLOAD:
            return jsonify({"ok": False, "error": f"Too many files (max {MAX_UPLOAD})."}), 400

        # ---- File type + magic-byte validation -----------------------------
        allowed = app.config.get("ALLOWED_EXTENSIONS", {".pdf"})
        max_per_file = app.config.get("MAX_PER_FILE_LENGTH",
                                       config.Config.MAX_PER_FILE_LENGTH)
        rejected = []
        valid_files = []
        for f in files:
            ext = os.path.splitext(f.filename or "")[1].lower()
            if ext not in allowed:
                rejected.append(f.filename or "unknown")
                continue
            # Magic-byte check: read first bytes and verify PDF header.
            head = f.stream.read(5)
            f.stream.seek(0)
            if not head.startswith(PDF_MAGIC):
                _sec_log.warning("Non-PDF magic bytes from %s: %s",
                                 request.remote_addr, f.filename)
                rejected.append(f.filename or "unknown")
                continue
            # Per-file size heuristic: if content-length is set and exceeds
            # the limit, reject early.  (streaming size check is unreliable
            # with multipart, so this is a best-effort filter.)
            if hasattr(f, 'content_length') and f.content_length:
                if f.content_length > max_per_file:
                    _sec_log.warning("Oversized upload from %s: %s (%d bytes)",
                                     request.remote_addr, f.filename,
                                     f.content_length)
                    rejected.append(f.filename or "unknown")
                    continue
            valid_files.append(f)
        if rejected:
            _log.warning("Rejected uploads: %s", rejected)
        if not valid_files:
            return jsonify({"ok": False,
                            "error": "No valid PDF files in the upload."}), 400

        uploaded = []
        for index, stream in enumerate(valid_files):
            try:
                path, name = sessions.save_upload(sid, index, stream)
                uploaded.append({"path": path, "name": name})
            except Exception:
                _log.exception("Upload failed for file %s",
                               getattr(stream, "filename", "?"))
                return jsonify({"ok": False,
                                "error": "Upload failed. Check file size and try again."}), 400

        results = []
        for upload in uploaded:
            results.append(convert_upload(sessions, sid, upload,
                                          password=request.form.get("password")))
        return jsonify({"ok": True, "files": results,
                        "outputs": list_outputs(sessions, sid)})

    # ----------------------------------------------------------- downloads
    @app.get("/api/outputs")
    def api_outputs():
        sid = _current_session()
        return jsonify({"ok": True, "outputs": list_outputs(sessions, sid)})

    @app.get("/api/download/<name>")
    def api_download(name):
        sid = _current_session()
        if "/" in name or "\\" in name or not name.endswith(".xlsx"):
            return jsonify({"ok": False, "error": "Invalid file name."}), 400
        # Block path traversal via ".." sequences and null bytes.
        if ".." in name or "\x00" in name:
            return jsonify({"ok": False, "error": "Invalid file name."}), 400
        out_dir = sessions.output_dir(sid)
        path = os.path.join(out_dir, name)
        # Verify the resolved path is still inside the output directory.
        real_path = os.path.realpath(path)
        real_out = os.path.realpath(out_dir)
        if not real_path.startswith(real_out + os.sep):
            _sec_log.warning("Path traversal attempt from %s: %s",
                             request.remote_addr, name)
            return jsonify({"ok": False, "error": "Invalid file name."}), 400
        if not os.path.isfile(path):
            return jsonify({"ok": False, "error": "File not found."}), 404
        return send_file(path, as_attachment=True, download_name=name)

    # -------------------------------------------------------------- merge
    @app.post("/api/merge")
    def api_merge():
        _log = logging.getLogger(__name__)

        sid = _current_session()
        body = request.get_json(silent=True) or {}
        filenames = [n for n in body.get("files", []) if isinstance(n, str)]
        mode = body.get("mode", "sheets")
        if mode not in ("sheets", "append"):
            return jsonify({"ok": False, "error": "Unknown merge mode."}), 400
        try:
            result = merge_outputs(sessions, sid, filenames, mode=mode)
        except ValueError as exc:
            return jsonify({"ok": False, "error": str(exc)}), 400
        except Exception:
            _log.exception("Merge failed for session %s", sid)
            return jsonify({"ok": False,
                            "error": "Merge failed. Please try again."}), 500
        return jsonify({"ok": True, "result": result,
                        "outputs": list_outputs(sessions, sid)})

    # ---------------------------------------------------------- extraction
    @app.post("/api/extract")
    def api_extract():
        _log = logging.getLogger(__name__)

        sid = _current_session()
        body = request.get_json(silent=True) or {}
        fields = [f for f in body.get("fields", []) if isinstance(f, str)]
        filenames = [n for n in body.get("files", []) if isinstance(n, str)]
        try:
            fields = validate_fields(fields)
            result = extract_fields(sessions, sid, fields,
                                    filenames=filenames or None)
        except ValueError as exc:
            return jsonify({"ok": False, "error": str(exc)}), 400
        except Exception:
            _log.exception("Extraction failed for session %s", sid)
            return jsonify({"ok": False,
                            "error": "Extraction failed. Please try again."}), 500
        return jsonify({"ok": True, "result": result,
                        "outputs": list_outputs(sessions, sid)})

    # ------------------------------------------------------------ cleanup
    @app.post("/api/reset")
    def api_reset():
        sid = _current_session()
        _sec_log.info("Session reset: %s from %s",
                      sid[:8], request.remote_addr)
        sessions.destroy(sid)
        flask_session.pop("sid", None)
        flask_session.pop("_csrf_token", None)
        return jsonify({"ok": True})

    # Keep the Flask dev server importable for `python -m pdf2xlsx.web.server`.
    return app


def main():
    app = create_app()
    port = int(os.environ.get("PDF2EXCEL_PORT", "8000"))
    host = os.environ.get("PDF2EXCEL_HOST", "127.0.0.1")
    print(f" * PDF → Excel web app: http://{host}:{port}")
    app.run(host=host, port=port, debug=False, threaded=True)


if __name__ == "__main__":
    main()
