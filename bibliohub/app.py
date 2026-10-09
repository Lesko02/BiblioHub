import importlib
import os
import sys
import threading
import tkinter as tk
from tkinter import messagebox, ttk

from . import APP_NAME, APP_VERSION, deps, theme
from .pages.home import HomePage
from .registry import TOOLS
from .settings import Settings


def enable_dpi_awareness() -> None:
    if sys.platform != "win32":
        return
    import ctypes
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(2)
    except Exception:
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass


class InstallDialog(tk.Toplevel):
    """Finestra modale che installa i pacchetti mancanti con pip in background."""

    def __init__(self, parent, packages, on_finish):
        super().__init__(parent)
        self.title("Installazione pacchetti")
        self.configure(background=theme.BG)
        self.resizable(False, False)
        self.transient(parent)
        self.protocol("WM_DELETE_WINDOW", lambda: None)

        frame = ttk.Frame(self, padding=24)
        frame.pack(fill="both", expand=True)
        ttk.Label(frame, text="Installazione dei pacchetti necessari in corso…",
                  style="Section.TLabel").pack(anchor="w")
        ttk.Label(frame, text=", ".join(packages), style="Muted.TLabel").pack(
            anchor="w", pady=(4, 12))
        bar = ttk.Progressbar(frame, mode="indeterminate", length=380)
        bar.pack(fill="x")
        bar.start(12)

        self._result = None
        self._on_finish = on_finish
        threading.Thread(target=self._work, args=(packages,), daemon=True).start()
        self.after(200, self._poll)
        self.grab_set()

    def _work(self, packages):
        try:
            self._result = deps.install(packages)
        except Exception as exc:
            self._result = exc

    def _poll(self):
        if self._result is None:
            self.after(200, self._poll)
            return
        self.grab_release()
        self.destroy()
        self._on_finish(self._result)


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.settings = Settings()
        theme.apply(self)

        self.title(APP_NAME)
        scale = self.winfo_fpixels("1i") / 96
        width, height = int(1120 * scale), int(780 * scale)
        x = (self.winfo_screenwidth() - width) // 2
        y = max((self.winfo_screenheight() - height) // 2 - int(20 * scale), 0)
        self.geometry(f"{width}x{height}+{x}+{y}")
        self.minsize(int(900 * scale), int(640 * scale))

        self._build_header()

        self.container = ttk.Frame(self)
        self.container.pack(fill="both", expand=True)
        self.container.rowconfigure(0, weight=1)
        self.container.columnconfigure(0, weight=1)

        self.pages = {}
        self.home = HomePage(self.container, self)
        self.home.grid(row=0, column=0, sticky="nsew")
        self.show_home()

        self.protocol("WM_DELETE_WINDOW", self.on_close)

    def _build_header(self):
        header = ttk.Frame(self, style="Header.TFrame", padding=(16, 10))
        header.pack(fill="x")
        self.back_button = ttk.Button(header, text="←  Home", style="Nav.TButton",
                                      command=self.show_home)
        self.header_title = ttk.Label(header, text=APP_NAME, style="Header.TLabel")
        self.header_title.pack(side="left")
        self.header_sub = ttk.Label(header, text="", style="HeaderSub.TLabel")
        self.header_sub.pack(side="left", padx=(10, 0))
        ttk.Label(header, text=f"v{APP_VERSION}", style="HeaderSub.TLabel").pack(side="right")

    def show_home(self):
        self.back_button.pack_forget()
        self.header_title.configure(text=APP_NAME)
        self.header_sub.configure(text="")
        self.home.refresh()
        self.home.tkraise()

    def _show_page(self, info, page):
        self.back_button.pack(side="left", padx=(0, 14), before=self.header_title)
        self.header_title.configure(text=info.title)
        self.header_sub.configure(text=f"v{info.version}")
        page.tkraise()

    def open_tool(self, info):
        page = self.pages.get(info.key)
        if page is not None:
            self._show_page(info, page)
            return

        missing = deps.missing_packages(info.requires)
        if missing:
            self._offer_install(info, missing)
            return

        try:
            module_name, class_name = info.page.split(":")
            page_cls = getattr(importlib.import_module(module_name), class_name)
            page = page_cls(self.container, self)
        except Exception as exc:
            messagebox.showerror(APP_NAME, f"Impossibile aprire «{info.title}»:\n\n{exc}",
                                 parent=self)
            return

        page.grid(row=0, column=0, sticky="nsew")
        self.pages[info.key] = page
        self._show_page(info, page)

    def _offer_install(self, info, missing):
        pkgs = ", ".join(missing)
        if not deps.can_install():
            messagebox.showerror(APP_NAME, f"Pacchetti mancanti per «{info.title}»: {pkgs}",
                                 parent=self)
            return
        if not messagebox.askyesno(
                APP_NAME,
                f"«{info.title}» richiede i seguenti pacchetti Python non installati:\n\n"
                f"{pkgs}\n\nVuoi installarli ora?", parent=self):
            return

        def finished(result):
            if isinstance(result, Exception) or result.returncode != 0:
                detail = result if isinstance(result, Exception) else (
                    result.stderr.strip() or result.stdout.strip())
                messagebox.showerror(APP_NAME, f"Installazione non riuscita:\n\n{detail}",
                                     parent=self)
                return
            still_missing = deps.missing_packages(info.requires)
            if still_missing:
                messagebox.showwarning(
                    APP_NAME, "Installazione completata, ma è necessario riavviare "
                              f"{APP_NAME} per usare: {', '.join(still_missing)}", parent=self)
                return
            self.open_tool(info)

        InstallDialog(self, missing, finished)

    def is_running(self, key) -> bool:
        page = self.pages.get(key)
        return bool(page and page.running)

    def refresh_home(self):
        self.home.refresh()

    def on_close(self):
        running = [i.title for i in TOOLS if self.is_running(i.key)]
        if running and not messagebox.askyesno(
                APP_NAME, "Operazioni ancora in corso:\n\n" + "\n".join(running)
                          + "\n\nChiudere comunque?", icon="warning", parent=self):
            return
        self.settings.save()
        self.destroy()


def main():
    # Con pythonw.exe stdout/stderr sono None: alcune librerie ci scrivono comunque.
    if sys.stdout is None:
        sys.stdout = open(os.devnull, "w", encoding="utf-8")
    if sys.stderr is None:
        sys.stderr = open(os.devnull, "w", encoding="utf-8")

    enable_dpi_awareness()
    App().mainloop()
