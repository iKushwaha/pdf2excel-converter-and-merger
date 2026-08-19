"""Session-scoped storage for uploaded PDFs and generated workbooks.

Each browser session gets its own directory under ``Config.DATA_DIR``:

    DATA_DIR/<session_id>/
        uploads/<n>__<sanitized-name>.pdf     original PDFs
        output/<n>__<stem>.xlsx               generated workbooks

Sessions expire after ``SESSION_TTL_SECONDS``; a background thread sweeps
stale sessions so the data dir does not grow without bound.
"""

import os
import re
import shutil
import threading
import time
import uuid

from .. import config

_UNSAFE = re.compile(r'[^A-Za-z0-9._-]+')
_ZERO_BYTE = re.compile(r'\x00')


def _sanitize(name):
    """Keep a readable filename while removing path separators, null bytes, and junk."""
    base = os.path.basename(name or "file.pdf")
    base = _ZERO_BYTE.sub("", base)  # strip null bytes first
    stem, ext = os.path.splitext(base)
    safe = _UNSAFE.sub("_", stem).strip("._") or "file"
    if len(safe) > 80:
        safe = safe[:80]
    return f"{safe}{ext.lower()}"


def _cleanup_old_sessions(data_dir, ttl, max_sessions=None):
    """Remove session dirs whose mtime is older than ``ttl`` seconds.

    When *max_sessions* is set, the oldest sessions are also pruned once the
    count exceeds the limit — even if they haven't expired yet.
    """
    try:
        dirs = []
        for name in os.listdir(data_dir):
            if name.startswith("._"):
                continue
            path = os.path.join(data_dir, name)
            if not os.path.isdir(path):
                continue
            dirs.append(path)

        removed = 0
        # First pass: remove expired sessions.
        now = time.time()
        for path in dirs[:]:
            try:
                age = now - os.path.getmtime(path)
                if age > ttl:
                    shutil.rmtree(path, ignore_errors=True)
                    dirs.remove(path)
                    removed += 1
            except OSError:
                continue

        # Second pass: if still over limit, remove oldest first.
        if max_sessions and len(dirs) > max_sessions:
            dirs.sort(key=lambda p: os.path.getmtime(p))
            for path in dirs[: len(dirs) - max_sessions]:
                try:
                    shutil.rmtree(path, ignore_errors=True)
                    removed += 1
                except OSError:
                    continue

        return removed
    except OSError:
        pass


class SessionManager:
    """Owns session directories and the background cleanup thread."""

    def __init__(self, data_dir=None, ttl=None, max_sessions=None):
        self.data_dir = data_dir or config.Config.DATA_DIR
        self.ttl = ttl or config.Config.SESSION_TTL_SECONDS
        self.max_sessions = max_sessions or config.Config.MAX_SESSIONS
        self._stop = threading.Event()
        os.makedirs(self.data_dir, exist_ok=True)
        _cleanup_old_sessions(self.data_dir, self.ttl, self.max_sessions)
        self._thread = threading.Thread(target=self._sweep_loop, daemon=True)
        self._thread.start()

    def _sweep_loop(self):
        while not self._stop.wait(config.Config.CLEANUP_INTERVAL_SECONDS):
            try:
                _cleanup_old_sessions(self.data_dir, self.ttl, self.max_sessions)
            except Exception:
                continue

    def shutdown(self):
        self._stop.set()

    # ------------------------------------------------------------------ api
    def new_session(self):
        sid = uuid.uuid4().hex
        os.makedirs(self.uploads_dir(sid), exist_ok=True)
        os.makedirs(self.output_dir(sid), exist_ok=True)
        return sid

    def is_valid(self, sid):
        return bool(sid) and _UNSAFE.sub("", sid) == sid and len(sid) <= 64

    def touch(self, sid):
        try:
            os.utime(os.path.join(self.data_dir, sid), None)
        except OSError:
            pass

    def uploads_dir(self, sid):
        return os.path.join(self.data_dir, sid, "uploads")

    def output_dir(self, sid):
        return os.path.join(self.data_dir, sid, "output")

    # -------------------------------------------------------------- uploads
    def save_upload(self, sid, index, stream):
        """Write one uploaded PDF, returning (stored_path, safe_name)."""
        safe = _sanitize(getattr(stream, "filename", f"file{index}.pdf"))
        stored = os.path.join(self.uploads_dir(sid), f"{index:03d}__{safe}")
        with open(stored, "wb") as fh:
            while True:
                chunk = stream.read(1024 * 1024)
                if not chunk:
                    break
                fh.write(chunk)
        return stored, safe

    def list_uploads(self, sid):
        """Return [{path, name}] sorted by upload order (the n__ prefix)."""
        uploads = []
        directory = self.uploads_dir(sid)
        for name in sorted(os.listdir(directory)):
            if name.startswith("._"):
                continue
            path = os.path.join(directory, name)
            if os.path.isfile(path):
                display = re.sub(r"^\d{3}__", "", name)
                uploads.append({"path": path, "name": display})
        return uploads

    def output_name(self, sid, stem, suffix):
        """Find a generated output file by stem prefix, else None."""
        directory = self.output_dir(sid)
        for name in os.listdir(directory):
            if name.startswith("._"):
                continue
            if suffix and name.endswith(suffix):
                return os.path.join(directory, name)
            if not suffix and os.path.splitext(name)[0].startswith(stem):
                return os.path.join(directory, name)
        return None

    def destroy(self, sid):
        """Delete a session's files (best effort)."""
        shutil.rmtree(os.path.join(self.data_dir, sid), ignore_errors=True)
