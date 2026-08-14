#!/usr/bin/env python3
"""PDF → Excel Converter & Merger — application entry point.

Usage:
    python main.py
"""

try:
    from tkinterdnd2 import TkinterDnD
except Exception:  # drag & drop support is optional
    TkinterDnD = None

from pdf2xlsx.ui import App


def main():
    if TkinterDnD is not None:
        root = TkinterDnD.Tk()
    else:
        import tkinter as tk

        root = tk.Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
