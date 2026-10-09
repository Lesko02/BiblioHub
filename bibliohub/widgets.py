import os
import queue
import threading
import traceback
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from . import theme


def documents_dir() -> Path:
    docs = Path.home() / "Documents"
    return docs if docs.is_dir() else Path.home()


def open_path(path) -> None:
    """Apre un file o una cartella con l'applicazione predefinita di sistema."""
    path = str(path)
    if hasattr(os, "startfile"):
        os.startfile(path)
    else:
        import subprocess, sys
        subprocess.Popen(["open" if sys.platform == "darwin" else "xdg-open", path])


class Card(ttk.Frame):
    """Riquadro bianco con titolo, usato per raggruppare i campi di un form."""

    def __init__(self, parent, title=None, **kwargs):
        super().__init__(parent, style="Card.TFrame", padding=(16, 12), **kwargs)
        self.columnconfigure(1, weight=1)
        if title:
            ttk.Label(self, text=title, style="CardSection.TLabel").grid(
                row=0, column=0, columnspan=3, sticky="w", pady=(0, 8))


class PathField:
    """
    Riga di form (etichetta | campo | Sfoglia) posizionata a griglia in un Card.
    mode: 'open' (file esistente), 'save' (nuovo file), 'dir' (cartella).
    """

    def __init__(self, parent, row, label, mode="open", filetypes=None,
                 default_ext=None, initial="", hint=None):
        self.mode = mode
        self.filetypes = filetypes or [("Tutti i file", "*.*")]
        self.default_ext = default_ext
        self.label = label
        self.var = tk.StringVar(value=initial or "")

        ttk.Label(parent, text=label, style="Card.TLabel").grid(
            row=row, column=0, sticky="w", padx=(0, 12), pady=4)
        self.entry = ttk.Entry(parent, textvariable=self.var)
        self.entry.grid(row=row, column=1, sticky="ew", pady=4)
        self.button = ttk.Button(parent, text="Sfoglia…", command=self.browse)
        self.button.grid(row=row, column=2, sticky="e", padx=(8, 0), pady=4)

        if hint:
            ttk.Label(parent, text=hint, style="CardMuted.TLabel").grid(
                row=row + 1, column=1, columnspan=2, sticky="w", pady=(0, 4))

    def get(self) -> str:
        return self.var.get().strip().strip('"')

    def set(self, value) -> None:
        self.var.set("" if value is None else str(value))

    def _initial_dir(self) -> str:
        current = self.get()
        if current:
            p = Path(current)
            candidate = p if self.mode == "dir" else p.parent
            if candidate.is_dir():
                return str(candidate)
        return str(documents_dir())

    def browse(self) -> None:
        initialdir = self._initial_dir()
        if self.mode == "dir":
            path = filedialog.askdirectory(title=f"Seleziona {self.label}",
                                           initialdir=initialdir, mustexist=False)
        elif self.mode == "save":
            current = Path(self.get()).name if self.get() else ""
            path = filedialog.asksaveasfilename(title=f"Salva {self.label} come…",
                                                initialdir=initialdir, initialfile=current,
                                                defaultextension=self.default_ext,
                                                filetypes=self.filetypes)
        else:
            path = filedialog.askopenfilename(title=f"Seleziona {self.label}",
                                              initialdir=initialdir, filetypes=self.filetypes)
        if path:
            self.set(str(Path(path)))

    def set_enabled(self, enabled: bool) -> None:
        state = "normal" if enabled else "disabled"
        self.entry.configure(state=state)
        self.button.configure(state=state)


class LogConsole(ttk.Frame):
    def __init__(self, parent, **kwargs):
        super().__init__(parent, style="Card.TFrame", **kwargs)
        self.text = tk.Text(self, height=10, wrap="word", state="disabled",
                            font=(theme.MONO_FAMILY, 9), background="#FBFCFE",
                            foreground=theme.TEXT, relief="flat", borderwidth=0,
                            padx=10, pady=8, highlightthickness=1,
                            highlightbackground=theme.BORDER, highlightcolor=theme.BORDER)
        scroll = ttk.Scrollbar(self, orient="vertical", command=self.text.yview)
        self.text.configure(yscrollcommand=scroll.set)
        self.text.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")
        self.text.tag_configure("error", foreground=theme.ERROR)
        self.text.tag_configure("success", foreground=theme.SUCCESS)

    def write(self, message: str, tag: str | None = None) -> None:
        self.text.configure(state="normal")
        self.text.insert("end", f"{message}\n", (tag,) if tag else ())
        self.text.see("end")
        self.text.configure(state="disabled")

    def clear(self) -> None:
        self.text.configure(state="normal")
        self.text.delete("1.0", "end")
        self.text.configure(state="disabled")


class TaskRunner:
    """
    Esegue una funzione in un thread separato senza bloccare l'interfaccia.
    La funzione riceve log(msg), progress(corrente, totale), should_stop() come keyword.
    I messaggi vengono passati al thread della GUI tramite una coda.
    """

    POLL_MS = 80

    def __init__(self, widget: tk.Misc):
        self._widget = widget
        self._queue = queue.Queue()
        self._thread = None
        self._stop = threading.Event()
        self._handlers = {}

    @property
    def running(self) -> bool:
        return self._thread is not None and self._thread.is_alive()

    def stop(self) -> None:
        self._stop.set()

    def start(self, target, on_log, on_progress, on_done, on_error) -> None:
        self._stop.clear()
        self._handlers = {"log": on_log, "progress": on_progress,
                          "done": on_done, "error": on_error}
        q = self._queue

        def worker():
            try:
                result = target(log=lambda m: q.put(("log", m)),
                                progress=lambda cur, tot: q.put(("progress", (cur, tot))),
                                should_stop=self._stop.is_set)
                q.put(("done", result))
            except Exception as exc:
                q.put(("error", (exc, traceback.format_exc())))

        self._thread = threading.Thread(target=worker, daemon=True)
        self._thread.start()
        self._widget.after(self.POLL_MS, self._poll)

    def _poll(self) -> None:
        try:
            while True:
                kind, payload = self._queue.get_nowait()
                if kind == "progress":
                    self._handlers["progress"](*payload)
                elif kind == "error":
                    self._handlers["error"](*payload)
                    return
                else:
                    self._handlers[kind](payload)
                    if kind == "done":
                        return
        except queue.Empty:
            pass
        self._widget.after(self.POLL_MS, self._poll)


def friendly_error(exc: Exception) -> str:
    if isinstance(exc, PermissionError):
        return (f"Impossibile scrivere il file:\n{exc.filename or exc}\n\n"
                "Probabilmente è aperto in un altro programma (es. Excel o Word). "
                "Chiudilo e riprova.")
    if isinstance(exc, FileNotFoundError):
        return f"File non trovato:\n{exc.filename or exc}"
    return str(exc) or exc.__class__.__name__


class ToolPage(ttk.Frame):
    """
    Struttura comune delle pagine strumento:
    intestazione, form (self.body, riempito da build()), barra azioni,
    avanzamento e log attività.
    """

    key = ""
    title = ""
    version = ""
    description = ""

    def __init__(self, parent, app):
        super().__init__(parent, padding=(28, 18, 28, 18))
        self.app = app
        self.settings = app.settings
        self.runner = TaskRunner(self)
        self._busy_widgets = []
        self._result_path = None

        head = ttk.Frame(self)
        head.pack(fill="x")
        ttk.Label(head, text=self.title, style="PageTitle.TLabel").pack(side="left")
        ttk.Label(head, text=f"v{self.version}", style="Muted.TLabel").pack(
            side="left", padx=(10, 0), pady=(6, 0))
        desc = ttk.Label(self, text=self.description, style="Muted.TLabel", justify="left")
        desc.pack(fill="x", pady=(2, 12))
        desc.bind("<Configure>", lambda e: desc.configure(wraplength=max(e.width, 200)))

        self.body = ttk.Frame(self)
        self.body.pack(fill="x")

        actions = ttk.Frame(self)
        actions.pack(fill="x", pady=(14, 6))
        self.actions = actions
        self.stop_button = ttk.Button(actions, text="⏹ Interrompi", style="Danger.TButton",
                                      command=self._request_stop)
        self.open_button = ttk.Button(actions, text="📂 Apri risultato",
                                      command=self._open_result)

        status_row = ttk.Frame(self)
        status_row.pack(fill="x", pady=(4, 4))
        ttk.Label(status_row, text="Attività", style="Section.TLabel").pack(side="left")
        self.status = ttk.Label(status_row, text="Pronto", style="Status.TLabel",
                                foreground=theme.MUTED)
        self.status.pack(side="right")

        self.progress = ttk.Progressbar(self, mode="determinate", maximum=1)
        self.progress.pack(fill="x", pady=(0, 8))

        self.log = LogConsole(self)
        self.log.pack(fill="both", expand=True)

        self.build()

    def build(self) -> None:
        raise NotImplementedError

    def add_action(self, text, command, accent=False) -> ttk.Button:
        btn = ttk.Button(self.actions, text=text, command=command,
                         style="Accent.TButton" if accent else "TButton")
        btn.pack(side="left", padx=(0, 8))
        self._busy_widgets.append(btn)
        return btn

    def register_busy(self, *widgets) -> None:
        """Widget da disabilitare mentre un'operazione è in corso."""
        self._busy_widgets.extend(widgets)

    @property
    def running(self) -> bool:
        return self.runner.running

    def set_result(self, path) -> None:
        self._result_path = path
        if path:
            self.open_button.pack(side="right")
        else:
            self.open_button.pack_forget()

    def _open_result(self) -> None:
        if self._result_path and Path(self._result_path).exists():
            open_path(self._result_path)
        else:
            messagebox.showwarning("BiblioHub", "Il risultato non è più disponibile.",
                                   parent=self)

    def _set_status(self, text, color=theme.MUTED) -> None:
        self.status.configure(text=text, foreground=color)

    def _set_busy(self, busy: bool, stoppable: bool = False) -> None:
        for w in self._busy_widgets:
            if isinstance(w, PathField):
                w.set_enabled(not busy)
            else:
                w.configure(state="disabled" if busy else "normal")
        if busy and stoppable:
            self.stop_button.configure(state="normal")
            self.stop_button.pack(side="left", padx=(0, 8))
        else:
            self.stop_button.pack_forget()

    def _request_stop(self) -> None:
        self.runner.stop()
        self.stop_button.configure(state="disabled")
        self._set_status("Interruzione in corso…", theme.WARNING)

    def _on_progress(self, current, total) -> None:
        if str(self.progress.cget("mode")) == "indeterminate":
            self.progress.stop()
            self.progress.configure(mode="determinate")
        self.progress.configure(maximum=max(total, 1), value=current)
        self._set_status(f"In esecuzione… {current}/{total}", theme.ACCENT)

    def run_task(self, target, on_done=None, stoppable=False, start_message=None) -> None:
        """target(log=, progress=, should_stop=) viene eseguito in background."""
        if self.running:
            return
        self.log.clear()
        self.set_result(None)
        if start_message:
            self.log.write(start_message)
        self._set_busy(True, stoppable)
        self._set_status("In esecuzione…", theme.ACCENT)
        self.progress.configure(mode="indeterminate", value=0)
        self.progress.start(12)

        def finish():
            self.progress.stop()
            self._set_busy(False)
            self.app.refresh_home()

        def done(result):
            finish()
            self.progress.configure(mode="determinate", maximum=1, value=1)
            self._set_status("Completato", theme.SUCCESS)
            if on_done:
                on_done(result)

        def error(exc, tb):
            finish()
            self.progress.configure(mode="determinate", maximum=1, value=0)
            self._set_status("Errore", theme.ERROR)
            message = friendly_error(exc)
            self.log.write(f"❌ {message}", "error")
            self.log.write(tb.rstrip(), "error")
            messagebox.showerror(self.title, message, parent=self)

        self.runner.start(target, on_log=self.log.write, on_progress=self._on_progress,
                          on_done=done, on_error=error)
        self.app.refresh_home()

    def validate_input_file(self, field: PathField) -> bool:
        value = field.get()
        if not value:
            messagebox.showwarning(self.title, f"Seleziona: {field.label}", parent=self)
            return False
        if not Path(value).is_file():
            messagebox.showwarning(self.title, f"File non trovato:\n{value}", parent=self)
            return False
        return True
