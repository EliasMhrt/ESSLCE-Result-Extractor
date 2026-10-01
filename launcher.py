"""ESSLCE Result Extractor - Launcher (GUI for teachers)

Double-click Start.bat (or run "python launcher.py") to open a window with
drag-and-drop zones, a start button and a live progress log.

Improvements over the original GUI:
- DPI-aware rendering (no more blurry text on high-res displays).
- Live progress bar + "student X of Y" counter + ETA estimate.
- Colored status pill (Ready / Running / Paused / Done / Error).
- File pre-validation: dropped/selected students lists are checked for the
  required columns before Start is enabled.
- Auto-loads students.xlsx / results.xlsx when they sit next to the app.
- Timestamped, colorized log with a capped scrollback.
- Pause / Resume buttons (via pause.flag) that survive internet outages, and a
  clean Stop (via stop.flag) that lets the engine flush progress before exit.
- Keyboard shortcuts (Ctrl+O / Ctrl+S / Ctrl+P / Esc) and window state memory.
"""

import ctypes
import json
import os
import queue
import re
import subprocess
import sys
import threading
import time
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from tkinterdnd2 import DND_FILES, TkinterDnD

_IS_FROZEN = getattr(sys, "frozen", False)
BASE_DIR = os.path.dirname(os.path.abspath(sys.executable if _IS_FROZEN else __file__))
PYTHON = os.path.join(BASE_DIR, ".venv", "Scripts", "python.exe")
if not os.path.exists(PYTHON):
    PYTHON = sys.executable
EXTRACTOR = os.path.join(BASE_DIR, "extractor.py")
FAST_EXTRACTOR = os.path.join(BASE_DIR, "extractor_fast.py")
CREATE_NO_WINDOW = 0x08000000

PAUSE_FILE = os.path.join(BASE_DIR, "pause.flag")
STOP_FILE = os.path.join(BASE_DIR, "stop.flag")
STATE_FILE = os.path.join(BASE_DIR, "gui_state.json")
FORCE_STOP_AFTER = 90                  # seconds after Stop before hard-killing

ICON_FILE = os.path.join(
    getattr(sys, "_MEIPASS", BASE_DIR) if _IS_FROZEN else BASE_DIR, "app.ico")

ADMISSION_KEYS = [
    "admission", "admission_no", "admission number", "candidate", "reg",
    "registration", "id number",
]
FIRSTNAME_KEYS = ["first_name", "first name", "firstname", "fname", "name"]


def _enable_dpi():
    """Make tkinter sharp on high-DPI displays (fixes the classic blur)."""
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass


def _runner_cmd():
    """Command that runs the extraction logic.

    Prefers extractor_fast.py (pause/stop/auto-pause engine) when present,
    otherwise falls back to extractor.py. In dev this is `python <engine>`; in
    a frozen build it is the exe itself, which knows to enter extraction mode
    when given -i/--input.
    """
    if _IS_FROZEN:
        return [sys.executable]
    if os.path.exists(FAST_EXTRACTOR):
        return [PYTHON, FAST_EXTRACTOR]
    return [PYTHON, EXTRACTOR]


BG = "#0f172a"
CARD = "#1e293b"
CARD_ALT = "#273449"
FG = "#e2e8f0"
MUTED = "#94a3b8"
ACCENT = "#22c55e"
ACCENT_DARK = "#16a34a"
WARN = "#f59e0b"
DANGER = "#ef4444"
BORDER = "#334155"
DROP_BG = "#0b1220"


def _strip_dnd(path):
    if path.startswith("{") and path.endswith("}"):
        path = path[1:-1]
    return path


def _apply_icon(root):
    """Set the app icon on the window AND the taskbar button.

    Tk's iconbitmap sets the title-bar icon, but Windows keeps the taskbar
    button showing whatever WM_SETICON was last sent - so we also push the icon
    through WM_SETICON on the real HWND.
    """
    if not os.path.exists(ICON_FILE):
        return
    for default in (False, True):
        try:
            if default:
                root.iconbitmap(default=ICON_FILE)
            else:
                root.iconbitmap(ICON_FILE)
        except Exception:
            pass
    try:
        user32 = ctypes.windll.user32
        user32.LoadImageW.argtypes = [
            ctypes.c_void_p, ctypes.c_wchar_p, ctypes.c_uint,
            ctypes.c_int, ctypes.c_int, ctypes.c_uint]
        user32.LoadImageW.restype = ctypes.c_void_p
        user32.SendMessageW.restype = ctypes.c_void_p
        user32.DestroyIcon.argtypes = [ctypes.c_void_p]
        hwnd = root.winfo_id()
        for wparam, size in ((0, 32), (1, 256)):
            hic = user32.LoadImageW(
                None, ICON_FILE, 1, size, size, 0x10)  # IMAGE_ICON, LR_LOADFROMFILE
            if hic:
                user32.SendMessageW(hwnd, 0x0080, wparam, hic)  # WM_SETICON
                user32.DestroyIcon(hic)
    except Exception:
        pass


class LauncherApp:
    def __init__(self, root):
        self.root = root
        self.input_path = ""
        self.output_path = os.path.join(BASE_DIR, "results.xlsx")
        self.proc = None
        self.running = False
        self.log_queue = queue.Queue()
        self.force_stop_at = None
        self._wrap_widgets = []
        self._vbar_packed = True

        self.total = 0
        self.done_base = 0
        self.completed = 0
        self.eta_stamps = []

        root.title("ESSLCE Result Extractor - Oryares")
        _apply_icon(root)
        root.geometry("960x900")
        root.minsize(720, 480)
        root.configure(bg=BG)
        try:
            scaling = ctypes.windll.user32.GetDpiForWindow(root.winfo_id()) / 96.0
            if scaling > 1.0:
                root.tk.call("tk", "scaling", scaling)
        except Exception:
            pass
        root.option_add("*Font", ("Segoe UI", 13))

        style = ttk.Style(root)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        style.configure("Card.TFrame", background=CARD)
        style.configure("CardTitle.TLabel", background=CARD, foreground=FG,
                        font=("Segoe UI", 14, "bold"))
        style.configure("Hint.TLabel", background=CARD, foreground=MUTED,
                        font=("Segoe UI", 12))
        style.configure("Path.TLabel", background=CARD, foreground=ACCENT,
                        font=("Segoe UI", 12, "bold"), wraplength=760)
        style.configure("Drop.TLabel", background=DROP_BG, foreground=MUTED,
                        font=("Segoe UI", 13))
        style.configure("Big.TButton", font=("Segoe UI", 15, "bold"),
                        padding=(18, 12))
        style.map("Big.TButton",
                  background=[("disabled", "#334155"), ("pressed", ACCENT_DARK),
                              ("active", "#4ade80")],
                  foreground=[("disabled", "#64748b")])
        style.configure("BigAccent.TButton", background=ACCENT,
                        foreground="#052e16", font=("Segoe UI", 15, "bold"),
                        padding=(18, 12), relief="flat")
        style.map("BigAccent.TButton",
                  background=[("disabled", "#334155"), ("pressed", ACCENT_DARK),
                              ("active", "#4ade80")],
                  foreground=[("disabled", "#64748b")])
        style.configure("BigWarn.TButton", background=WARN,
                        foreground="#111827", font=("Segoe UI", 15, "bold"),
                        padding=(18, 12), relief="flat")
        style.map("BigWarn.TButton",
                  background=[("disabled", "#334155"), ("pressed", "#d97706"),
                              ("active", "#fbbf24")])
        style.configure("Small.TButton", font=("Segoe UI", 12), padding=(10, 6))
        style.configure("Accent.Horizontal.TProgressbar", background=ACCENT,
                        troughcolor=DROP_BG, bordercolor=BORDER,
                        lightcolor=ACCENT, darkcolor=ACCENT)

    # ------------------------------------------------------------------
    # Layout
    # ------------------------------------------------------------------

    def build(self):
        self._restore_geometry()
        outer = ttk.Frame(self.root, style="Shell.TFrame")
        outer.pack(fill="both", expand=True)
        self.canvas = tk.Canvas(outer, bg=BG, highlightthickness=0, bd=0,
                                takefocus=True)
        self.vbar = ttk.Scrollbar(outer, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=self.vbar.set)
        self.vbar.pack(side="right", fill="y")
        self.canvas.pack(side="left", fill="both", expand=True)
        shell = ttk.Frame(self.canvas, padding=18, style="Shell.TFrame")
        self._shell_window = self.canvas.create_window((0, 0), window=shell, anchor="nw")
        shell.bind("<Configure>", self._on_shell_resize)
        self.canvas.bind("<Configure>", self._on_canvas_resize)
        for seq in ("<MouseWheel>", "<Button-4>", "<Button-5>"):
            self.root.bind_all(seq, self._on_mousewheel, add="+")
        for seq, cmd in (("<Prior>", -1), ("<Next>", 1), ("<Up>", -1), ("<Down>", 1),
                         ("<Home>", "home"), ("<End>", "end")):
            self.root.bind_all(seq, lambda e, c=cmd: self._on_keyscroll(c), add="+")
        self.root.option_add("*TCombobox*Listbox.font", ("Segoe UI", 13))
        try:
            ttk.Style(self.root).configure("Shell.TFrame", background=BG)
        except tk.TclError:
            pass

        title = tk.Label(
            shell, text="ESSLCE Result Extractor",
            bg=BG, fg=FG, font=("Segoe UI", 26, "bold"))
        subtitle = tk.Label(
            shell, text="Donated to Tana Haik Secondary School by Oryares",
            bg=BG, fg=MUTED, font=("Segoe UI", 13))
        title.pack(anchor="w")
        subtitle.pack(anchor="w", pady=(0, 14))

        self._students_card(shell)
        self._output_card(shell)
        self._progress_card(shell)
        self._actions_card(shell)
        self._log_card(shell)

        self.root.update_idletasks()
        try:
            self.root.state("zoomed")
        except tk.TclError:
            pass
        self.root.update_idletasks()
        _apply_icon(self.root)
        self._autoload_defaults()
        self._refresh_paths()
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

        self.root.bind("<Control-o>", lambda e: self._browse_students())
        self.root.bind("<Control-s>", lambda e: self._start())
        self.root.bind("<Control-p>", lambda e: self._toggle_pause())
        self.root.bind("<Escape>", lambda e: self._stop())

        self.root.update_idletasks()
        self._update_wraplength(self.canvas.winfo_width())
        self._sync_scroll()
        self.root.after(120, self._drain_log)

    def _on_shell_resize(self, _event):
        self._sync_scroll()

    def _on_canvas_resize(self, event):
        self.canvas.itemconfigure(self._shell_window, width=event.width)
        self._update_wraplength(event.width)
        self._sync_scroll()

    def _sync_scroll(self):
        """Keep the scrollregion in sync and hide the bar when everything fits."""
        bbox = self.canvas.bbox("all")
        if bbox is None:
            self._show_scrollbar(False)
            return
        self.canvas.configure(scrollregion=bbox)
        self._show_scrollbar(bbox[3] - bbox[1] > self.canvas.winfo_height() + 2)

    def _show_scrollbar(self, needed):
        # Track our own flag: winfo_ismapped() is 0 for everything before the
        # window is first shown, so it cannot be trusted to mean "is packed".
        needed = bool(needed)
        if needed == self._vbar_packed:
            return
        if needed:
            self.vbar.pack(side="right", fill="y")
        else:
            self.vbar.pack_forget()
        self._vbar_packed = needed

    def _register_wrap(self, widget, inset):
        """Make a label re-wrap to the viewport instead of a hardcoded width."""
        self._wrap_widgets.append((widget, inset))
        return widget

    def _update_wraplength(self, width):
        for widget, inset in self._wrap_widgets:
            target = max(240, width - inset)
            match = re.search(r"-?\d+", str(widget.cget("wraplength") or "0"))
            current = int(match.group()) if match else 0
            if abs(current - target) > 4:
                widget.configure(wraplength=target)

    _SCROLL_EXEMPT = ("Text", "Listbox", "Entry", "TEntry", "TCombobox", "TSpinbox", "Spinbox")

    def _under_pointer(self, event):
        widget = getattr(event, "widget", None)
        if widget is None:
            widget = self.root.winfo_containing(event.x_root, event.y_root)
        while widget is not None:
            if widget.winfo_class() in self._SCROLL_EXEMPT:
                return True
            widget = getattr(widget, "master", None)
        return False

    def _on_mousewheel(self, event):
        if self._under_pointer(event):
            return
        num = getattr(event, "num", None)
        if num == 4:
            self.canvas.yview_scroll(-3, "units")
        elif num == 5:
            self.canvas.yview_scroll(3, "units")
        elif event.delta:
            self.canvas.yview_scroll(-int(event.delta / 120) * 3, "units")

    def _on_keyscroll(self, direction):
        widget = self.root.focus_get()
        if widget is not None and widget.winfo_class() in self._SCROLL_EXEMPT:
            return
        if direction == "home":
            self.canvas.yview_moveto(0.0)
        elif direction == "end":
            self.canvas.yview_moveto(1.0)
        else:
            page = max(1, self.canvas.winfo_height() // 24)
            self.canvas.yview_scroll(direction * page, "units")

    def _students_card(self, parent):
        card = ttk.Frame(parent, style="Card.TFrame", padding=14)
        card.pack(fill="x", pady=(0, 10))

        ttk.Label(card, text="1.  Students list", style="CardTitle.TLabel").pack(anchor="w")
        ttk.Label(
            card,
            text="Drag your students Excel file here, or browse.\n"
                 "It must have an 'Admission Number' column and a 'First Name' column.",
            style="Hint.TLabel",
        ).pack(anchor="w", pady=(2, 10))

        drop = tk.Frame(card, bg=DROP_BG, highlightbackground=ACCENT,
                        highlightthickness=2, bd=0, height=150)
        drop.pack(fill="x")
        drop.pack_propagate(False)

        self.drop_icon = tk.Label(
            drop, text="\U0001F4C4  Drop students.xlsx here", bg=DROP_BG, fg=FG,
            font=("Segoe UI", 18, "bold"))
        self.drop_icon.pack(pady=(45, 0))
        self.drop_hint = tk.Label(
            drop, text="or click Browse", bg=DROP_BG, fg=MUTED, font=("Segoe UI", 13))
        self.drop_hint.pack(pady=(4, 0))

        for widget in (drop, self.drop_icon, self.drop_hint):
            widget.drop_target_register(DND_FILES)
            widget.dnd_bind("<<Drop>>", self._on_drop_students)

        row = tk.Frame(card, bg=CARD)
        row.pack(fill="x", pady=(10, 0))
        self.students_path = self._register_wrap(tk.Label(
            row, text="No file selected", bg=CARD, fg=MUTED, font=("Segoe UI", 12),
            anchor="w", wraplength=560), 174)
        self.students_path.pack(side="left", fill="x", expand=True)
        ttk.Button(row, text="Browse", style="Small.TButton",
                   command=self._browse_students).pack(side="right")

        self.file_check = self._register_wrap(tk.Label(
            card, text="", bg=CARD, fg=MUTED, font=("Segoe UI", 12),
            anchor="w", wraplength=700), 72)
        self.file_check.pack(fill="x", pady=(8, 0))

    def _output_card(self, parent):
        card = ttk.Frame(parent, style="Card.TFrame", padding=14)
        card.pack(fill="x", pady=(0, 10))

        ttk.Label(card, text="2.  Where to save the results", style="CardTitle.TLabel").pack(anchor="w")
        ttk.Label(
            card,
            text="Default is 'results.xlsx'. Drag a different file or folder to change it.",
            style="Hint.TLabel",
        ).pack(anchor="w", pady=(2, 10))

        drop = tk.Frame(card, bg=DROP_BG, highlightbackground=BORDER,
                        highlightthickness=2, bd=0, height=150)
        drop.pack(fill="x")
        drop.pack_propagate(False)
        out = tk.Label(
            drop, text="\U0001F4BE  Drop a destination file here", bg=DROP_BG,
            fg=FG, font=("Segoe UI", 18, "bold"))
        out.pack(pady=(50, 0))
        for widget in (drop, out):
            widget.drop_target_register(DND_FILES)
            widget.dnd_bind("<<Drop>>", self._on_drop_output)

        row = tk.Frame(card, bg=CARD)
        row.pack(fill="x", pady=(10, 0))
        self.output_path_lbl = self._register_wrap(tk.Label(
            row, text="", bg=CARD, fg=ACCENT, font=("Segoe UI", 12, "bold"),
            anchor="w", wraplength=560), 174)
        self.output_path_lbl.pack(side="left", fill="x", expand=True)
        ttk.Button(row, text="Browse", style="Small.TButton",
                   command=self._browse_output).pack(side="right")

    def _progress_card(self, parent):
        card = ttk.Frame(parent, style="Card.TFrame", padding=14)
        card.pack(fill="x", pady=(0, 10))

        header = tk.Frame(card, bg=CARD)
        header.pack(fill="x")
        ttk.Label(header, text="3.  Live progress", style="CardTitle.TLabel").pack(side="left")
        self.status_lbl = tk.Label(
            card, text="Ready", bg=MUTED, fg="#0f172a",
            font=("Segoe UI", 13, "bold"), padx=14, pady=3)
        self.status_lbl.place(x=0, y=0)
        self.status_lbl.pack(side="right", padx=(8, 0))

        self.bar = ttk.Progressbar(card, style="Accent.Horizontal.TProgressbar",
                                   maximum=1, value=0, length=0)
        self.bar.pack(fill="x", pady=(12, 6))

        info = tk.Frame(card, bg=CARD)
        info.pack(fill="x")
        self.progress_lbl = tk.Label(
            info, text="Student 0 / 0", bg=CARD, fg=FG,
            font=("Segoe UI", 13, "bold"), anchor="w")
        self.progress_lbl.pack(side="left")
        self.eta_lbl = tk.Label(
            info, text="", bg=CARD, fg=MUTED, font=("Segoe UI", 12), anchor="e")
        self.eta_lbl.pack(side="right")

        self.progress_hint = self._register_wrap(tk.Label(
            card, text="Start a run to see progress here.", bg=CARD, fg=MUTED,
            font=("Segoe UI", 12), anchor="w", wraplength=700), 72)
        self.progress_hint.pack(fill="x", pady=(4, 0))

    def _actions_card(self, parent):
        card = ttk.Frame(parent, style="Card.TFrame", padding=14)
        card.pack(fill="x", pady=(0, 10))

        row = tk.Frame(card, bg=CARD)
        row.pack(fill="x")

        self.start_btn = ttk.Button(row, text="\u25B6  Start Extraction",
                                    style="BigAccent.TButton", command=self._start)
        self.start_btn.pack(side="left")

        self.stop_btn = ttk.Button(row, text="\u25A0  Stop",
                                   style="Big.TButton", command=self._stop,
                                   state="disabled")
        self.stop_btn.pack(side="left", padx=(8, 0))

        self.open_btn = ttk.Button(row, text="Open Results",
                                   style="Small.TButton", command=self._open_results,
                                   state="disabled")
        self.open_btn.pack(side="right")

    def _log_card(self, parent):
        card = ttk.Frame(parent, style="Card.TFrame", padding=10)
        card.pack(fill="x")

        self.log = tk.Text(
            card, bg="#020617", fg="#cbd5e1", insertbackground=FG,
            font=("Consolas", 12), relief="flat", padx=8, pady=6, height=10)
        self.log.tag_configure("time", foreground="#475569")
        self.log.tag_configure("sys", foreground="#94a3b8")
        self.log.tag_configure("ok", foreground=ACCENT)
        self.log.tag_configure("warn", foreground=WARN)
        self.log.tag_configure("err", foreground=DANGER)
        sb = ttk.Scrollbar(card, orient="vertical", command=self.log.yview)
        self.log.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        self.log.pack(side="left", fill="both", expand=True)
        self.log.configure(state="disabled")

    # ------------------------------------------------------------------
    # File validation & helpers
    # ------------------------------------------------------------------

    def _autoload_defaults(self):
        default_in = os.path.join(BASE_DIR, "students.xlsx")
        if not self.input_path and os.path.exists(default_in):
            self.input_path = default_in
        default_out = os.path.join(BASE_DIR, "results.xlsx")
        if os.path.exists(default_out):
            self.output_path = default_out

    def _validate_students_file(self, path):
        """Return (student_count, message). Lightweight check so the GUI
        doesn't need to import the heavy extractor engine."""
        if not path or not os.path.exists(path):
            return None, ""
        try:
            import pandas as pd
            lower = str(path).lower()
            if lower.endswith((".xlsx", ".xls")):
                df = pd.read_excel(path)
            elif lower.endswith(".csv"):
                df = pd.read_csv(path)
            else:
                with open(path, "rb") as fh:
                    head = fh.read(2)
                enc = "utf-16" if head in (b"\xff\xfe", b"\xfe\xff") else "utf-8-sig"
                try:
                    df = pd.read_csv(path, encoding=enc, sep=",")
                except (UnicodeDecodeError, Exception):
                    df = pd.read_csv(path, encoding="latin-1", sep=",")
        except Exception as exc:
            return None, f"Could not read file: {exc}"

        cols = {str(c).strip().lower(): str(c) for c in df.columns}
        found_ad = found_fn = None
        for key in ADMISSION_KEYS:
            for col in cols:
                if col == key or col.startswith(key) or key in col:
                    found_ad = cols[col]
                    break
            if found_ad:
                break
        for key in FIRSTNAME_KEYS:
            for col in cols:
                if col == key or col.startswith(key) or key in col:
                    found_fn = cols[col]
                    break
            if found_fn:
                break

        count = len(df)
        if found_ad and found_fn:
            return count, (f"{count} students found - columns OK ({found_ad!r} + {found_fn!r})")
        if found_ad or found_fn:
            missing = "First Name" if found_ad else "Admission Number"
            return count, (f"Warning: missing the {missing!r} column. "
                           "Found: " + ", ".join(str(c) for c in df.columns))
        return count, ("Warning: could not find 'Admission Number' / 'First Name' "
                       "columns. Found: " + ", ".join(str(c) for c in df.columns))

    def _refresh_paths(self):
        if self.input_path:
            self.students_path.configure(text=self.input_path, fg=ACCENT)
            self.drop_icon.configure(text="\U0001F4C4  Selected")
            self.drop_hint.configure(text=os.path.basename(self.input_path))
            count, msg = self._validate_students_file(self.input_path)
            if msg:
                color = ACCENT if "OK" in msg else WARN
                self.file_check.configure(text=msg, fg=color)
            self.progress_hint.configure(
                text=f"File ready - {count or '?'} students. Click Start when ready.")
        else:
            self.students_path.configure(text="No file selected", fg=MUTED)
            self.drop_icon.configure(text="\U0001F4C4  Drop students.xlsx here")
            self.drop_hint.configure(text="or click Browse")
            self.file_check.configure(text="")
            self.progress_hint.configure(text="Start a run to see progress here.")
        self.output_path_lbl.configure(text=self.output_path)

    def _on_drop_students(self, event):
        path = _strip_dnd(event.data)
        if os.path.exists(path):
            self.input_path = path
            self._refresh_paths()
            self._log("Student list loaded: %s" % path)

    def _on_drop_output(self, event):
        path = _strip_dnd(event.data)
        if os.path.isdir(path):
            self.output_path = os.path.join(path, "results.xlsx")
        else:
            self.output_path = path
        self._refresh_paths()
        self._log("Results will be saved to: %s" % self.output_path)

    def _browse_students(self):
        path = filedialog.askopenfilename(
            title="Select the students list",
            filetypes=[("Excel / CSV", "*.xlsx *.xls *.csv"), ("All files", "*.*")])
        if path:
            self.input_path = path
            self._refresh_paths()
            self._log("Student list loaded: %s" % path)

    def _browse_output(self):
        path = filedialog.asksaveasfilename(
            title="Choose where to save results",
            defaultextension=".xlsx",
            initialfile="results.xlsx",
            filetypes=[("Excel workbook", "*.xlsx")])
        if path:
            self.output_path = path
            self._refresh_paths()

    # ------------------------------------------------------------------
    # Logging & status
    # ------------------------------------------------------------------

    def _set_status(self, text, color):
        self.status_lbl.configure(text=text, bg=color, fg="#0f172a")

    def _color_for(self, line):
        if line.startswith("> ") or line.startswith("Starting") or line.startswith("=" * 10):
            return "sys"
        if re.search(r"PAUSED|resumed|connectivity|internet|pause \d+s", line):
            return "warn"
        if "FOUND" in line:
            return "ok"
        if "NOT FOUND" in line:
            return "warn"
        if re.search(r"ERROR|error|failed|Aborting|unreachable", line):
            return "err"
        return None

    def _log(self, text, kind=None):
        try:
            self.log.configure(state="normal")
            stamp = time.strftime("%H:%M:%S")
            self.log.insert("end", "[" + stamp + "] ", ("time",))
            self.log.insert("end", text + "\n", (kind,)) if kind else self.log.insert("end", text + "\n")
            lines = int(self.log.index("end-1c").split(".")[0])
            if lines > 2000:
                self.log.delete("1.0", "1000.0")
            self.log.see("end")
            self.log.configure(state="disabled")
        except tk.TclError:
            pass

    # ------------------------------------------------------------------
    # Progress feed (parses extractor stdout)
    # ------------------------------------------------------------------

    def _feed_progress(self, line):
        m = re.search(r"Students:\s*(\d+)\s+\(collected:\s*(\d+)\)", line)
        if m:
            self.total = int(m.group(1))
            self.done_base = int(m.group(2))
            self.completed = 0
            self.bar.configure(maximum=max(self.total, 1),
                               value=min(self.done_base, self.total))
            self.progress_lbl.configure(
                text=f"Student {self.done_base} / {self.total}  ({self.done_base} already done)")
            self.progress_hint.configure(
                text="Resuming previous run - completed students will be skipped.")
            return
        if re.search(r"PAUSED", line):
            self._set_status("Paused", WARN)
            self.progress_hint.configure(text="Paused - extractor is waiting.")
            return
        if re.search(r"resumed|connectivity restored", line):
            self._set_status("Running", ACCENT)
            self.progress_hint.configure(text="Resumed - continuing extraction.")
            return
        if re.search(r"auto-pausing|internet still unreachable", line):
            self._set_status("Paused (internet?)", WARN)
            self.progress_hint.configure(text="Internet looks down - will resume automatically.")
            return
        m = re.search(r"\[(\d+)/(\d+)\]", line)
        if m:
            cur, tot = int(m.group(1)), int(m.group(2))
            self.total = tot
            self.progress_lbl.configure(text=f"Student {cur} / {tot}")
            return
        if re.match(r"\s*(FOUND|NOT FOUND|ERROR)", line):
            self.completed += 1
            done = min(self.done_base + self.completed, self.total)
            self.bar.configure(value=done)
            self.progress_lbl.configure(
                text=f"Student {min(self.done_base + self.completed, self.total)} / {self.total}  "
                     f"done this run: {self.completed}")
            self._update_eta()
            return
        m = re.search(r"Done\.\s*(\d+)\s+checked", line)
        if m:
            self.completed = int(m.group(1))
            self.bar.configure(value=min(self.done_base + self.completed, self.total))
            self.eta_lbl.configure(text="finished")

    def _update_eta(self):
        now = time.time()
        self.eta_stamps.append(now)
        self.eta_stamps = self.eta_stamps[-8:]
        if len(self.eta_stamps) < 3:
            self.eta_lbl.configure(text="ETA: --")
            return
        gaps = [b - a for a, b in zip(self.eta_stamps, self.eta_stamps[1:])]
        avg = sum(gaps) / len(gaps)
        remaining = max(self.total - (self.done_base + self.completed), 0)
        secs = int(remaining * avg)
        m, s = divmod(secs, 60)
        self.eta_lbl.configure(text=f"ETA: about {m}m {s}s")

    # ------------------------------------------------------------------
    # Run control
    # ------------------------------------------------------------------

    def _start(self):
        if self.running:
            self._toggle_pause()
            return
        if not self.input_path:
            messagebox.showwarning(
                "No student list",
                "Please select the students file first.\n\n"
                "Drag it into the top box (or click Browse).")
            return
        if not os.path.exists(self.input_path):
            messagebox.showerror("File not found", "Cannot find:\n%s" % self.input_path)
            return

        for flag in (PAUSE_FILE, STOP_FILE):
            try:
                if os.path.exists(flag):
                    os.remove(flag)
            except OSError:
                pass

        self.open_btn.configure(state="disabled")
        self.start_btn.configure(state="normal", text="\u23F8  Pause",
                                 style="BigWarn.TButton")
        self.stop_btn.configure(state="normal")
        self.force_stop_at = None
        self.total = 0
        self.done_base = 0
        self.completed = 0
        self.eta_stamps = []
        self.bar.configure(maximum=1, value=0)
        self._set_status("Running", ACCENT)
        self.progress_lbl.configure(text="Starting \u2026")
        self.eta_lbl.configure(text="")
        self.log.configure(state="normal")
        self.log.delete("1.0", "end")
        self.log.configure(state="disabled")

        cmd = _runner_cmd() + ["-i", self.input_path, "-o", self.output_path]
        self._log("> " + " ".join(cmd), "sys")
        self._log("Starting extraction \u2026", "sys")

        self.proc = subprocess.Popen(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            encoding="utf-8",
            errors="replace",
            bufsize=1,
            stdin=subprocess.DEVNULL,
            cwd=BASE_DIR,
            creationflags=CREATE_NO_WINDOW,
        )
        self.running = True
        threading.Thread(target=self._reader, daemon=True).start()

    def _toggle_pause(self):
        if not self.running or self.proc is None:
            return
        if os.path.exists(PAUSE_FILE):
            try:
                os.remove(PAUSE_FILE)
            except OSError:
                pass
            self.start_btn.configure(text="\u23F8  Pause", style="BigWarn.TButton")
            self._set_status("Running", ACCENT)
            self._log("Resume requested.", "warn")
        else:
            try:
                open(PAUSE_FILE, "w").close()
            except OSError:
                pass
            self.start_btn.configure(text="\u25B6  Resume", style="BigAccent.TButton")
            self._log("Pause requested - the engine will pause after the current student.", "warn")

    def _stop(self):
        if self.proc is None or self.proc.poll() is not None:
            return
        if not messagebox.askyesno(
                "Stop?",
                "Stop the run? Completed students are saved to progress.csv\n"
                "and skipped when you run it again later."):
            return
        try:
            open(STOP_FILE, "w").close()
        except OSError:
            pass
        self.force_stop_at = time.time() + FORCE_STOP_AFTER
        self.stop_btn.configure(state="disabled")
        self.start_btn.configure(state="disabled")
        self._set_status("Stopping", WARN)
        self._log("Stopping - saving progress \u2026", "warn")

    def _reader(self):
        for line in iter(self.proc.stdout.readline, ""):
            if line:
                self.log_queue.put(line.rstrip("\n"))
        self.proc.stdout.close()

    def _drain_log(self):
        try:
            while True:
                line = self.log_queue.get_nowait()
                self._feed_progress(line)
                self._log(line, self._color_for(line))
        except queue.Empty:
            pass

        if self.force_stop_at is not None and time.time() > self.force_stop_at:
            if self.proc is not None and self.proc.poll() is None:
                try:
                    self.proc.terminate()
                except OSError:
                    pass
                self._log("Engine did not exit in time - forced stop.", "err")
            self.force_stop_at = None

        if self.proc is not None and self.proc.poll() is not None:
            code = self.proc.poll()
            self.proc = None
            self.running = False
            for flag in (PAUSE_FILE, STOP_FILE):
                try:
                    if os.path.exists(flag):
                        os.remove(flag)
                except OSError:
                    pass
            self.force_stop_at = None
            self.start_btn.configure(state="normal", text="\u25B6  Start Extraction",
                                     style="BigAccent.TButton")
            self.stop_btn.configure(state="disabled")
            if code == 0:
                self._set_status("Finished \u2713", ACCENT)
                self.progress_hint.configure(text="All done. You can open the results file.")
                self.open_btn.configure(state="normal")
                self._log("=" * 60, "sys")
                self._log("Finished. You can open the results file now.", "ok")
            else:
                self._set_status("Stopped / error", DANGER)
                self.progress_hint.configure(text="The run stopped - see the log for details.")
        self.root.after(120, self._drain_log)

    def _open_results(self):
        if os.path.exists(self.output_path):
            os.startfile(self.output_path)
        else:
            messagebox.showinfo("Not ready",
                                "No results file yet.\nResults appear here after the run.")

    # ------------------------------------------------------------------
    # Window state
    # ------------------------------------------------------------------

    def _restore_geometry(self):
        try:
            with open(STATE_FILE, "r", encoding="utf-8") as fh:
                geom = json.load(fh).get("geometry")
            if geom:
                m = re.match(r"(\d+)x(\d+)", geom or "")
                if m:
                    w, h = int(m.group(1)), int(m.group(2))
                    if w >= self.root.minsize()[0] and h >= self.root.minsize()[1]:
                        self.root.geometry(geom)
        except Exception:
            pass

    def _on_close(self):
        try:
            with open(STATE_FILE, "w", encoding="utf-8") as fh:
                json.dump({"geometry": self.root.geometry()}, fh)
        except Exception:
            pass
        self.root.destroy()


def main():
    if "--selftest" in sys.argv:
        import extractor  # noqa: F401  (validates the full import chain)
        with open(os.path.join(BASE_DIR, "selftest_ok.txt"), "w", encoding="utf-8") as fh:
            fh.write("ok\n")
        return
    if "-i" in sys.argv or "--input" in sys.argv:
        if not _IS_FROZEN and os.path.exists(FAST_EXTRACTOR):
            from extractor_fast import main as extractor_main
        else:
            from extractor import main as extractor_main
        sys.exit(extractor_main())
    _enable_dpi()
    root = TkinterDnD.Tk()
    app = LauncherApp(root)
    app.build()
    root.mainloop()


if __name__ == "__main__":
    main()