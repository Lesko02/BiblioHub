import tkinter as tk
from tkinter import font as tkfont
from tkinter import ttk

PRIMARY = "#1F3B63"
PRIMARY_DARK = "#162B49"
ACCENT = "#366092"
ACCENT_HOVER = "#2B4E78"
BG = "#F3F5F9"
CARD = "#FFFFFF"
CARD_HOVER = "#EEF3FA"
BORDER = "#D5DCE6"
TEXT = "#1E2633"
MUTED = "#5B6777"
SUCCESS = "#2E7D32"
ERROR = "#C62828"
WARNING = "#B26A00"

FONT_FAMILY = "Segoe UI"
MONO_FAMILY = "Consolas"


def apply(root: tk.Tk) -> None:
    for name in ("TkDefaultFont", "TkTextFont", "TkMenuFont", "TkHeadingFont"):
        tkfont.nametofont(name).configure(family=FONT_FAMILY, size=10)
    tkfont.nametofont("TkFixedFont").configure(family=MONO_FAMILY, size=9)

    root.configure(background=BG)

    style = ttk.Style(root)
    style.theme_use("clam")

    style.configure(".", background=BG, foreground=TEXT, font=(FONT_FAMILY, 10),
                    bordercolor=BORDER, focuscolor=ACCENT)

    style.configure("TFrame", background=BG)
    style.configure("Card.TFrame", background=CARD)
    style.configure("Header.TFrame", background=PRIMARY)

    style.configure("TLabel", background=BG, foreground=TEXT)
    style.configure("Muted.TLabel", foreground=MUTED, font=(FONT_FAMILY, 9))
    style.configure("Hero.TLabel", font=(FONT_FAMILY, 26, "bold"), foreground=PRIMARY)
    style.configure("HeroSub.TLabel", font=(FONT_FAMILY, 12), foreground=MUTED)
    style.configure("PageTitle.TLabel", font=(FONT_FAMILY, 16, "bold"), foreground=PRIMARY)
    style.configure("Section.TLabel", font=(FONT_FAMILY, 11, "bold"), foreground=PRIMARY)
    style.configure("Header.TLabel", background=PRIMARY, foreground="white",
                    font=(FONT_FAMILY, 13, "bold"))
    style.configure("HeaderSub.TLabel", background=PRIMARY, foreground="#C9D6EA",
                    font=(FONT_FAMILY, 10))
    style.configure("Card.TLabel", background=CARD)
    style.configure("CardMuted.TLabel", background=CARD, foreground=MUTED, font=(FONT_FAMILY, 9))
    style.configure("CardSection.TLabel", background=CARD, foreground=PRIMARY,
                    font=(FONT_FAMILY, 11, "bold"))
    style.configure("Status.TLabel", font=(FONT_FAMILY, 9, "bold"))

    style.configure("TButton", padding=(12, 6), background=CARD, bordercolor=BORDER,
                    lightcolor=CARD, darkcolor=CARD)
    style.map("TButton",
              background=[("disabled", BG), ("active", CARD_HOVER)],
              foreground=[("disabled", MUTED)])

    style.configure("Accent.TButton", background=ACCENT, foreground="white",
                    bordercolor=ACCENT, lightcolor=ACCENT, darkcolor=ACCENT,
                    font=(FONT_FAMILY, 10, "bold"), padding=(16, 7))
    style.map("Accent.TButton",
              background=[("disabled", "#9FB2CC"), ("pressed", PRIMARY_DARK),
                          ("active", ACCENT_HOVER)],
              bordercolor=[("disabled", "#9FB2CC"), ("active", ACCENT_HOVER)],
              lightcolor=[("disabled", "#9FB2CC"), ("active", ACCENT_HOVER)],
              darkcolor=[("disabled", "#9FB2CC"), ("active", ACCENT_HOVER)],
              foreground=[("disabled", "#EEF2F7")])

    style.configure("Danger.TButton", foreground=ERROR)

    style.configure("Nav.TButton", background=PRIMARY, foreground="white",
                    bordercolor=PRIMARY, lightcolor=PRIMARY, darkcolor=PRIMARY,
                    font=(FONT_FAMILY, 10, "bold"), padding=(10, 4))
    style.map("Nav.TButton",
              background=[("active", PRIMARY_DARK)],
              bordercolor=[("active", PRIMARY_DARK)],
              lightcolor=[("active", PRIMARY_DARK)],
              darkcolor=[("active", PRIMARY_DARK)])

    style.configure("TEntry", fieldbackground="white", bordercolor=BORDER,
                    lightcolor=BORDER, padding=4)
    style.map("TEntry", bordercolor=[("focus", ACCENT)], lightcolor=[("focus", ACCENT)])

    for widget in ("TRadiobutton", "TCheckbutton"):
        style.configure(widget, background=CARD)
        style.map(widget, background=[("active", CARD)])

    style.configure("TNotebook", background=BG, borderwidth=0, tabmargins=(0, 4, 0, 0))
    style.configure("TNotebook.Tab", padding=(16, 6), background=BG, bordercolor=BORDER,
                    font=(FONT_FAMILY, 10))
    style.map("TNotebook.Tab",
              background=[("selected", CARD)],
              foreground=[("selected", PRIMARY)],
              font=[("selected", (FONT_FAMILY, 10, "bold"))])

    style.configure("Horizontal.TProgressbar", troughcolor=BORDER, background=ACCENT,
                    bordercolor=BORDER, lightcolor=ACCENT, darkcolor=ACCENT, thickness=8)
