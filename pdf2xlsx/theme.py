"""Dark, modern ttk theme (no external theme libraries needed)."""

BG = "#1e1f24"
PANEL = "#26272d"
PANEL_LIGHT = "#2d2e35"
FG = "#e6e7eb"
FG_DIM = "#9aa0ab"
BORDER = "#3c3f46"
ACCENT = "#4c9aff"
ACCENT_HOVER = "#6babff"
SELECT = "#33435c"
SELECT_FG = "#ffffff"
DANGER = "#e05656"
SUCCESS = "#7bd88f"
WARNING = "#ffd166"


def apply_theme(root, style):
    """Apply the dark theme to the app's ttk widgets."""
    root.configure(bg=BG)

    style.theme_use("clam")
    style.configure(".", background=BG, foreground=FG,
                    fieldbackground=PANEL, bordercolor=BORDER,
                    lightcolor=BG, darkcolor=BG, focusthickness=0)
    style.map(".", background=[("active", PANEL_LIGHT)])

    style.configure("TFrame", background=BG)
    style.configure("Toolbar.TFrame", background=BG)
    style.configure("Panel.TFrame", background=PANEL)

    style.configure("TLabel", background=BG, foreground=FG)
    style.configure("Panel.TLabel", background=PANEL, foreground=FG)
    style.configure("Dim.TLabel", background=BG, foreground=FG_DIM)
    style.configure("Title.TLabel", background=BG, foreground=FG,
                    font=("Helvetica Neue", 15, "bold"))
    style.configure("Hint.TLabel", background=BG, foreground=FG_DIM,
                    font=("Helvetica Neue", 10))

    style.configure("TButton", background=PANEL_LIGHT, foreground=FG,
                    bordercolor=BORDER, padding=(12, 6),
                    focuscolor=PANEL_LIGHT, relief="flat")
    style.map("TButton",
              background=[("active", ACCENT_HOVER), ("disabled", PANEL)],
              foreground=[("active", "#ffffff"), ("disabled", FG_DIM)])
    style.configure("Accent.TButton", background=ACCENT, foreground="#ffffff",
                    bordercolor=ACCENT, font=("Helvetica Neue", 11, "bold"))
    style.map("Accent.TButton",
              background=[("active", ACCENT_HOVER), ("disabled", "#2c4a6b")],
              foreground=[("active", "#ffffff"), ("disabled", "#8fa5c0")])
    style.configure("Danger.TButton", background="#7a2f2f", foreground="#ffe3e3")
    style.map("Danger.TButton", background=[("active", DANGER)])

    style.configure("Treeview", background=PANEL, fieldbackground=PANEL,
                    foreground=FG, rowheight=26, bordercolor=BORDER)
    style.map("Treeview", background=[("selected", SELECT)],
              foreground=[("selected", SELECT_FG)])
    style.configure("Treeview.Heading", background=PANEL_LIGHT, foreground=FG,
                    relief="flat", padding=(6, 5), font=("Helvetica Neue", 10, "bold"))
    style.map("Treeview.Heading", background=[("active", PANEL_LIGHT)])

    style.configure("Vertical.TScrollbar", background=PANEL_LIGHT,
                    troughcolor=PANEL, bordercolor=BORDER, arrowcolor=FG)
    style.configure("Horizontal.TScrollbar", background=PANEL_LIGHT,
                    troughcolor=PANEL, bordercolor=BORDER, arrowcolor=FG)

    style.configure("TProgressbar", background=ACCENT, troughcolor=PANEL,
                    bordercolor=BORDER, lightcolor=ACCENT, darkcolor=ACCENT)

    style.configure("TRadiobutton", background=BG, foreground=FG)
    style.map("TRadiobutton", background=[("active", BG)])
    style.configure("TCheckbutton", background=BG, foreground=FG)
    style.map("TCheckbutton", background=[("active", BG)])

    style.configure("TEntry", fieldbackground=PANEL, foreground=FG,
                    insertcolor=FG, bordercolor=BORDER)
    style.configure("TCombobox", fieldbackground=PANEL, foreground=FG,
                    background=PANEL_LIGHT, arrowcolor=FG)
