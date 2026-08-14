#!/usr/bin/env python3
"""PDF → Excel Converter & Merger — desktop application entry point.

Usage:
    python main.py
"""

import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "src"))

from pdf2xlsx.desktop.main import main  # noqa: E402

if __name__ == "__main__":
    sys.exit(main())
