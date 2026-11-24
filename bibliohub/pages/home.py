import sys
import tkinter as tk
from tkinter import ttk

from .. import APP_NAME, APP_SUBTITLE, APP_VERSION, theme
from ..registry import TOOLS
from .. import win_trackpad


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
        super().__init__(parent)
        self.app = app
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)

        self.canvas = tk.Canvas(self, background=theme.BG, highlightthickness=0, borderwidth=0)
        self.scrollbar = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        self.canvas.grid(row=0, column=0, sticky="nsew")
        self.scrollbar.grid(row=0, column=1, sticky="ns")

        body = ttk.Frame(self.canvas, padding=(40, 34, 28, 24))
        self._window = self.canvas.create_window((0, 0), window=body, anchor="nw")
        body.bind("<Configure>", self._on_body_configure)
        self.canvas.bind("<Configure>", self._on_canvas_configure)
        self.bind("<Destroy>", self._on_destroy)
        self._wheel_pixels = 0.0
        self._native_delta = 0
        self._trackpad = False
        if sys.platform == "win32" and win_trackpad.install(self._on_native_wheel):
            self._trackpad = True
            self.after(8, self._pump_native_wheel)
        else:
            self.bind_all("<MouseWheel>", self._on_wheel)
            self._hook_wheel(self)
            self.app.bind("<MouseWheel>", self._on_wheel, add="+")

        ttk.Label(body, text=f"📚 {APP_NAME}", style="Hero.TLabel").pack(anchor="w")
        ttk.Label(body, text=APP_SUBTITLE, style="HeroSub.TLabel").pack(anchor="w", pady=(2, 0))
        ttk.Label(
            body,
            text="Scegli uno strumento per iniziare.  ⓘ Guida, in alto a destra, "
                 "spiega la schermata in cui ti trovi.",
            style="Muted.TLabel",
        ).pack(anchor="w", pady=(14, 18))

        grid = ttk.Frame(body)
        grid.pack(fill="x")
        columns = 2

        self.cards = {}
        for index, info in enumerate(TOOLS):
            row, col = divmod(index, columns)
            grid.columnconfigure(col, weight=1, uniform="cards")
            card = ToolCard(grid, info, app.open_tool)
            card.grid(
                row=row, column=col, sticky="nsew",
                padx=(0 if col == 0 else 10, 0 if col == columns - 1 else 10),
                pady=(0 if row == 0 else 12, 0),
            )
            self.cards[info.key] = card

        ttk.Label(body, text=f"{APP_NAME} v{APP_VERSION}  ·  Università degli Studi di Milano",
                  style="Muted.TLabel").pack(anchor="w", pady=(22, 0))

    def _on_body_configure(self, _event):
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        self._update_scrollbar()

    def _on_canvas_configure(self, event):
        self.canvas.itemconfigure(self._window, width=event.width)
        self._update_scrollbar()

    def _update_scrollbar(self):
        bbox = self.canvas.bbox("all")
        if not bbox:
            return
        needed = (bbox[3] - bbox[1]) > self.canvas.winfo_height() + 1
        mapped = bool(self.scrollbar.winfo_ismapped())
        if needed and not mapped:
            self.scrollbar.grid()
        elif not needed and mapped:
            self.scrollbar.grid_remove()
            self.canvas.yview_moveto(0)

    def _hook_wheel(self, widget):
        widget.bind("<MouseWheel>", self._on_wheel, add="+")
        for child in widget.winfo_children():
            self._hook_wheel(child)

    def _scroll_pixels(self, pixels: int) -> None:
        bbox = self.canvas.bbox("all")
        if not bbox or pixels == 0:
            return
        height = max(bbox[3] - bbox[1], 1)
        top = self.canvas.canvasy(0) + pixels
        limit = max(height - self.canvas.winfo_height(), 0)
        top = min(max(top, 0), limit)
        self.canvas.yview_moveto(top / height)

    def _apply_wheel_delta(self, delta: int) -> None:
        if not delta or getattr(self.app, "current_key", None) != "home":
            return
        # One notch is 120. A trackpad sends a stream of smaller steps.
        self._wheel_pixels += (-delta / 120.0) * 48
        pixels = int(self._wheel_pixels)
        if pixels == 0:
            return
        self._wheel_pixels -= pixels
        self._scroll_pixels(pixels)

    def _on_native_wheel(self, delta: int) -> None:
        # Runs inside the Windows message hook. Touch Tk only from the pump.
        self._native_delta += int(delta)

    def _pump_native_wheel(self) -> None:
        if not self.winfo_exists():
            return
        delta = self._native_delta
        self._native_delta = 0
        if delta:
            self._apply_wheel_delta(delta)
        self.after(8, self._pump_native_wheel)

    def _on_wheel(self, event):
        self._apply_wheel_delta(getattr(event, "delta", 0) or 0)
        return "break"

    def _on_destroy(self, event):
        if event.widget is self:
            self.unbind_all("<MouseWheel>")
            if self._trackpad:
                win_trackpad.uninstall()

    def refresh(self):
        for key, card in self.cards.items():
            card.set_running(self.app.is_running(key))
