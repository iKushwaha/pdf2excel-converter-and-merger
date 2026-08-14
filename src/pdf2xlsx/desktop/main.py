"""Desktop application entry point (console script: pdf2excel-desktop).

Also importable: ``python -m pdf2xlsx.desktop.main``
"""

import sys

try:
    from tkinterdnd2 import TkinterDnD
except Exception:  # drag & drop support is optional
    TkinterDnD = None

from .ui import App


def main():
    if TkinterDnD is not None:
        root = TkinterDnD.Tk()
    else:
        import tkinter as tk

        root = tk.Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    sys.exit(main())
