"""Flask application factory for the PDF → Excel web app.

The web app is a server-side alternative to the desktop app: browsers
upload PDFs, the server converts them on the host machine with the same
core engine, and the resulting Excel files are downloaded back. Files are
kept in per-session temp directories and auto-expired.
"""

import os

from flask import (Flask, jsonify, render_template, request, send_file,
                   session as flask_session)

from . import config
from .services import (ALLOWED_FIELDS, convert_upload, extract_fields,
                       list_outputs, merge_outputs, validate_fields)
from .services.session import SessionManager

MAX_UPLOAD = config.Config.MAX_FILES_PER_UPLOAD


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

    sessions = SessionManager(app.config["DATA_DIR"])

    def _current_session():
        """Return the session id, creating one on first visit."""
        if "sid" not in flask_session or not sessions.is_valid(flask_session["sid"]):
            flask_session.clear()
            flask_session["sid"] = sessions.new_session()
        sessions.touch(flask_session["sid"])
        return flask_session["sid"]

    # -------------------------------------------------------------- pages
    @app.route("/")
    def index():
        return render_template("index.html",
                               landing_url=app.config.get("LANDING_URL", ""))

    @app.route("/health")
    def health():
        return jsonify({"ok": True, "name": "pdf2excel-web", "version": "1.0.0"})

    @app.route("/api/fields")
    def api_fields():
        return jsonify({"fields": [{"key": k, "label": v} for k, v in ALLOWED_FIELDS()]})

    # ------------------------------------------------------------ uploads
    @app.post("/api/upload")
    def api_upload():
        sid = _current_session()
        files = request.files.getlist("pdfs")
        files = [f for f in files if f and f.filename]
        if not files:
            return jsonify({"ok": False, "error": "No files selected."}), 400
        if len(files) > MAX_UPLOAD:
            return jsonify({"ok": False, "error": f"Too many files (max {MAX_UPLOAD})."}), 400

        uploaded = []
        for index, stream in enumerate(files):
            try:
                path, name = sessions.save_upload(sid, index, stream)
                uploaded.append({"path": path, "name": name})
            except Exception as exc:
                return jsonify({"ok": False, "error": f"Upload failed: {exc}"}), 400

        results = []
        for upload in uploaded:
            results.append(convert_upload(sessions, sid, upload,
                                          password=request.form.get("password")))
        return jsonify({"ok": True, "files": results, "outputs": list_outputs(sessions, sid)})

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
        path = os.path.join(sessions.output_dir(sid), name)
        if not os.path.isfile(path):
            return jsonify({"ok": False, "error": "File not found."}), 404
        return send_file(path, as_attachment=True, download_name=name)

    # -------------------------------------------------------------- merge
    @app.post("/api/merge")
    def api_merge():
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
        return jsonify({"ok": True, "result": result,
                        "outputs": list_outputs(sessions, sid)})

    # ---------------------------------------------------------- extraction
    @app.post("/api/extract")
    def api_extract():
        sid = _current_session()
        body = request.get_json(silent=True) or {}
        fields = [f for f in body.get("fields", []) if isinstance(f, str)]
        filenames = [n for n in body.get("files", []) if isinstance(n, str)]
        try:
            fields = validate_fields(fields)
            result = extract_fields(sessions, sid, fields, filenames=filenames or None)
        except ValueError as exc:
            return jsonify({"ok": False, "error": str(exc)}), 400
        return jsonify({"ok": True, "result": result,
                        "outputs": list_outputs(sessions, sid)})

    # ------------------------------------------------------------ cleanup
    @app.post("/api/reset")
    def api_reset():
        sid = _current_session()
        sessions.destroy(sid)
        flask_session.pop("sid", None)
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
