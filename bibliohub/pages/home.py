import tkinter as tk
from tkinter import ttk

from .. import APP_NAME, APP_SUBTITLE, APP_VERSION, theme
from ..registry import TOOLS


class ToolCard(tk.Frame):
    """Riquadro cliccabile che rappresenta uno strumento nella home."""

    def __init__(self, parent, info, on_open):
        super().__init__(parent, background=theme.CARD, highlightthickness=1,
                         highlightbackground=theme.BORDER, cursor="hand2")
        self.info = info
        self._on_open = on_open

        inner = tk.Frame(self, background=theme.CARD)
        inner.pack(fill="both", expand=True, padx=22, pady=20)

        top = tk.Frame(inner, background=theme.CARD)
        top.pack(fill="x")
        tk.Label(top, text=info.icon, font=("Segoe UI Emoji", 26),
                 background=theme.CARD).pack(side="left")
        tk.Label(top, text=f"v{info.version}", font=(theme.FONT_FAMILY, 9, "bold"),
                 background="#E4ECF7", foreground=theme.ACCENT, padx=8, pady=2).pack(
            side="right", anchor="n")

        tk.Label(inner, text=info.title, font=(theme.FONT_FAMILY, 14, "bold"),
                 foreground=theme.PRIMARY, background=theme.CARD, anchor="w").pack(
            fill="x", pady=(12, 4))
        self.desc = tk.Label(inner, text=info.description, font=(theme.FONT_FAMILY, 10),
                             foreground=theme.MUTED, background=theme.CARD,
                             justify="left", anchor="nw", wraplength=260)
        self.desc.pack(fill="both", expand=True)

        bottom = tk.Frame(inner, background=theme.CARD)
        bottom.pack(fill="x", pady=(16, 0))
        self.state_label = tk.Label(bottom, text="", font=(theme.FONT_FAMILY, 9, "bold"),
                                    foreground=theme.ACCENT, background=theme.CARD)
        self.state_label.pack(side="left")
        tk.Label(bottom, text="Apri  →", font=(theme.FONT_FAMILY, 10, "bold"),
                 foreground=theme.ACCENT, background=theme.CARD).pack(side="right")

        self._all = [self, inner, top, bottom, *inner.winfo_children(),
                     *top.winfo_children(), *bottom.winfo_children()]
        self._badge = top.winfo_children()[1]
        for w in self._all:
            w.bind("<Button-1>", lambda _e: self._on_open(self.info))
            w.bind("<Enter>", lambda _e: self._hover(True))
            w.bind("<Leave>", self._leave)
        inner.bind("<Configure>", lambda e: self.desc.configure(wraplength=max(e.width - 4, 120)))

    def _leave(self, event):
        x, y = self.winfo_pointerxy()
        if self.winfo_containing(x, y) not in self._all:
            self._hover(False)

    def _hover(self, on: bool):
        bg = theme.CARD_HOVER if on else theme.CARD
        self.configure(highlightbackground=theme.ACCENT if on else theme.BORDER)
        for w in self._all:
            if w is not self._badge:
                w.configure(background=bg)

    def set_running(self, running: bool):
        self.state_label.configure(text="● In esecuzione" if running else "")


class HomePage(ttk.Frame):
    def __init__(self, parent, app):
        super().__init__(parent, padding=(40, 34, 40, 24))
        self.app = app

        ttk.Label(self, text=f"📚 {APP_NAME}", style="Hero.TLabel").pack(anchor="w")
        ttk.Label(self, text=APP_SUBTITLE, style="HeroSub.TLabel").pack(anchor="w", pady=(2, 0))
        ttk.Label(self, text="Scegli uno strumento per iniziare.",
                  style="Muted.TLabel").pack(anchor="w", pady=(14, 18))

        grid = ttk.Frame(self)
        grid.pack(fill="x")

        self.cards = {}
        for col, info in enumerate(TOOLS):
            grid.columnconfigure(col, weight=1, uniform="cards")
            card = ToolCard(grid, info, app.open_tool)
            card.grid(row=0, column=col, sticky="nsew",
                      padx=(0 if col == 0 else 10, 0 if col == len(TOOLS) - 1 else 10))
            self.cards[info.key] = card

        ttk.Label(self, text=f"{APP_NAME} v{APP_VERSION}  ·  Università degli Studi di Milano",
                  style="Muted.TLabel").pack(side="bottom", anchor="w", pady=(22, 0))

    def refresh(self):
        for key, card in self.cards.items():
            card.set_running(self.app.is_running(key))
