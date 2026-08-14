"""General purpose helpers: paths, platform commands, formatting."""

import os
import subprocess
import sys


def unique_path(path):
    """Return ``path``, or a variant with a numeric suffix so it never
    overwrites an existing file (``file.xlsx`` -> ``file (1).xlsx``)."""
    if not os.path.exists(path):
        return path
    base, ext = os.path.splitext(path)
    n = 1
    while os.path.exists(f"{base} ({n}){ext}"):
        n += 1
    return f"{base} ({n}){ext}"


def base_stem(path):
    """File name without directory or extension."""
    return os.path.splitext(os.path.basename(path))[0]


def output_xlsx_path(pdf_path, out_dir):
    """Build the xlsx path for a PDF, deduplicating collisions."""
    name = base_stem(pdf_path) + ".xlsx"
    return unique_path(os.path.join(out_dir, name))


def format_size(num_bytes):
    """Human readable file size."""
    size = float(num_bytes)
    for unit in ("B", "KB", "MB", "GB", "TB"):
        if size < 1024.0 or unit == "TB":
            return f"{size:.1f} {unit}" if unit != "B" else f"{int(size)} B"
        size /= 1024.0
    return f"{size:.1f} TB"


def open_folder(path):
    """Open a folder in the OS file manager. Returns True on success."""
    path = os.path.abspath(path)
    if not os.path.isdir(path):
        return False
    try:
        if sys.platform == "darwin":
            subprocess.Popen(["open", path])
        elif os.name == "nt":
            os.startfile(path)  # noqa: S606
        else:
            subprocess.Popen(["xdg-open", path])
        return True
    except Exception:
        return False


def reveal_file(path):
    """Reveal a single file in the OS file manager."""
    path = os.path.abspath(path)
    if not os.path.exists(path):
        return False
    try:
        if sys.platform == "darwin":
            subprocess.Popen(["open", "-R", path])
            return True
        return open_folder(os.path.dirname(path))
    except Exception:
        return False
