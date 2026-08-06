"""esa-lookup GUI -- tkinter, dark theme, background worker + queue drain.

Run: `python esa_lookup.py`
"""
from __future__ import annotations

import os
import queue
import sys
import threading
import time
import tkinter as tk
from tkinter import filedialog, scrolledtext, ttk

import pipeline


# ---------------------------------------------------------------------------
# Dark palette
# ---------------------------------------------------------------------------
BG        = "#1e1e1e"   # window bg
PANEL     = "#252526"   # frames
INPUT     = "#2d2d2d"   # entries
INPUT_BRD = "#3f3f46"
FG        = "#e0e0e0"   # normal text
MUTED     = "#909090"   # secondary text
ACCENT    = "#4a9eff"   # buttons / progress
ACCENT_HI = "#65afff"
OK        = "#4ec9b0"
WARN      = "#dcdcaa"
ERROR     = "#f48771"
LOG_BG    = "#0f0f0f"
LOG_FG    = "#d4d4d4"
FONT_UI   = ("Segoe UI", 10)
FONT_UI_B = ("Segoe UI Semibold", 10)
FONT_MONO = ("Consolas", 9)


def apply_dark_theme(root: tk.Tk) -> ttk.Style:
    root.configure(bg=BG)
    style = ttk.Style()
    style.theme_use("clam")

    style.configure(".", background=BG, foreground=FG, fieldbackground=INPUT,
                    bordercolor=INPUT_BRD, lightcolor=BG, darkcolor=BG,
                    font=FONT_UI)
    style.configure("TFrame", background=BG)
    style.configure("Panel.TFrame", background=PANEL)
    style.configure("TLabel", background=BG, foreground=FG)
    style.configure("Muted.TLabel", background=BG, foreground=MUTED)
    style.configure("Status.TLabel", background=BG, foreground=ACCENT, font=FONT_UI_B)

    style.configure("TEntry", fieldbackground=INPUT, foreground=FG,
                    bordercolor=INPUT_BRD, lightcolor=INPUT_BRD, darkcolor=INPUT_BRD,
                    insertcolor=FG)

    style.configure("TButton", background=PANEL, foreground=FG, borderwidth=0,
                    padding=(12, 6), focusthickness=0)
    style.map("TButton",
              background=[("active", "#3a3a3d"), ("disabled", "#2a2a2b")],
              foreground=[("disabled", MUTED)])

    style.configure("Accent.TButton", background=ACCENT, foreground="#0b1220",
                    font=FONT_UI_B, padding=(16, 6))
    style.map("Accent.TButton",
              background=[("active", ACCENT_HI), ("disabled", "#2f4a70")],
              foreground=[("disabled", "#7a90ad")])

    style.configure("TRadiobutton", background=BG, foreground=FG,
                    focuscolor=BG, indicatorcolor=INPUT)
    style.map("TRadiobutton",
              background=[("active", BG)],
              indicatorcolor=[("selected", ACCENT), ("!selected", INPUT_BRD)])

    style.configure("TCheckbutton", background=BG, foreground=FG,
                    focuscolor=BG, indicatorcolor=INPUT)
    style.map("TCheckbutton",
              background=[("active", BG)],
              indicatorcolor=[("selected", ACCENT), ("!selected", INPUT_BRD)])

    style.configure("TLabelframe", background=BG, foreground=MUTED, bordercolor=INPUT_BRD)
    style.configure("TLabelframe.Label", background=BG, foreground=MUTED)

    # Instructions panel and per-step captions
    style.configure("Hint.TLabel", background=BG, foreground=MUTED, font=("Segoe UI", 10))
    style.configure("Step.TLabel", background=BG, foreground=ACCENT, font=FONT_UI_B)

    style.configure("Horizontal.TProgressbar", background=ACCENT, troughcolor=INPUT,
                    bordercolor=INPUT_BRD, lightcolor=ACCENT, darkcolor=ACCENT)

    return style


# ---------------------------------------------------------------------------
# App
# ---------------------------------------------------------------------------

class App:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("esa-lookup")
        self.root.geometry("820x740")
        self.root.minsize(700, 620)
        apply_dark_theme(root)

        self.event_q: queue.Queue = queue.Queue()
        self.worker: threading.Thread | None = None
        self.stop_event = threading.Event()

        self.file_var = tk.StringVar()
        self.workflow_var = tk.StringVar(value="TO")
        self.dryrun_var = tk.BooleanVar(value=False)

        self._build_ui()
        self._drain_queue()

    # -- layout ------------------------------------------------------------

    def _build_ui(self) -> None:
        pad = {"padx": 12, "pady": 6}

        # -- Instructions panel --------------------------------------------
        instr = ttk.Labelframe(self.root, text="How to use this app  -  follow steps 1-5 in order")
        instr.pack(fill="x", **pad)
        instructions_text = (
            "Do these OUTSIDE this app FIRST:\n"
            "   1.  Log into SAP GUI.  Any session with authorization for\n"
            "        transaction ZTBV, plant ESA1 is fine.  Leave it running.\n"
            "   2.  Open the target .xlsx file in Microsoft Excel.  Keep it open.\n"
            "\n"
            "In this app:\n"
            "   3.  Click 'Browse...' below and select that SAME Excel file.\n"
            "   4.  Choose which Process to run (see the two options below).\n"
            "   5.  Click 'Run'.  Do NOT close SAP GUI or Excel while it runs -\n"
            "        the app is driving both windows for you.\n"
            "\n"
            "Watch Status / Progress / Log while it works.  Click 'Stop' to cancel -\n"
            "nothing is written to the Excel file unless every SAP step succeeds.\n"
            "When done, the file is filled in via one final write pass + saved."
        )
        ttk.Label(
            instr, text=instructions_text, style="Hint.TLabel",
            justify="left", anchor="w",
        ).pack(anchor="w", fill="x", padx=12, pady=(4, 8))

        # -- Step 3: file picker --------------------------------------------
        row1 = ttk.Frame(self.root)
        row1.pack(fill="x", **pad)
        ttk.Label(row1, text="Step 3  -  Excel file:", style="Step.TLabel").pack(side="left")
        entry = ttk.Entry(row1, textvariable=self.file_var)
        entry.pack(side="left", fill="x", expand=True, padx=(8, 8))
        ttk.Button(row1, text="Browse...", command=self._on_browse).pack(side="left")

        # -- Step 4: workflow radios ----------------------------------------
        row2 = ttk.Labelframe(self.root, text="Step 4  -  Process")
        row2.pack(fill="x", **pad)
        ttk.Radiobutton(
            row2,
            text="TO Number process  (3 steps: LTAP -> Z50CFG_ENG_CRNT -> Z50CFG_ENG_VALD)",
            value="TO", variable=self.workflow_var,
        ).pack(anchor="w", padx=12, pady=(6, 0))
        ttk.Label(
            row2, style="Hint.TLabel",
            text="        Reads columns K, N, O, C.  Writes M, A, C, D, E, F, G, H, I, J.",
        ).pack(anchor="w", padx=12)
        ttk.Radiobutton(
            row2,
            text="Notification Number process  (2 steps: Z50CFG_ENG_CRNT -> Z50CFG_ENG_VALD)",
            value="NOTIF", variable=self.workflow_var,
        ).pack(anchor="w", padx=12, pady=(6, 0))
        ttk.Label(
            row2, style="Hint.TLabel",
            text="        Reads columns A, C.  Writes C, D, E, F, G, H, I, J.",
        ).pack(anchor="w", padx=12, pady=(0, 6))

        # -- Step 5: Run / Stop ---------------------------------------------
        row3 = ttk.Frame(self.root)
        row3.pack(fill="x", **pad)
        ttk.Label(row3, text="Step 5:", style="Step.TLabel").pack(side="left")
        self.run_btn = ttk.Button(row3, text="Run", style="Accent.TButton", command=self._on_run)
        self.run_btn.pack(side="left", padx=(8, 0))
        self.stop_btn = ttk.Button(row3, text="Stop", command=self._on_stop, state="disabled")
        self.stop_btn.pack(side="left", padx=(8, 0))
        self.diag_btn = ttk.Button(
            row3, text="Diagnose", command=self._on_diagnose)
        self.diag_btn.pack(side="left", padx=(8, 0))
        ttk.Checkbutton(
            row3, text="Dry run (report matches only, write nothing)",
            variable=self.dryrun_var,
        ).pack(side="left", padx=(12, 0))
        ttk.Label(
            row3, style="Hint.TLabel",
            text="   Keep this window, SAP GUI, and Excel all visible while it runs.",
        ).pack(side="left", padx=(12, 0))
        ttk.Label(
            self.root, style="Hint.TLabel", justify="left", anchor="w",
            text=("        'Diagnose' reads the ZTBV selection screens for the "
                  "chosen Process and writes what it finds to the log.\n"
                  "        Needs SAP only - no Excel file, no query is run, "
                  "nothing is written. Use it when a run stops with\n"
                  "        \"Could not resolve the ... filter\", then send the "
                  "log file."),
        ).pack(anchor="w", padx=12)

        # -- Status + progress ----------------------------------------------
        row4 = ttk.Frame(self.root)
        row4.pack(fill="x", **pad)
        self.status_var = tk.StringVar(value="Idle  -  ready when you are")
        ttk.Label(row4, textvariable=self.status_var, style="Status.TLabel").pack(anchor="w")
        self.progress = ttk.Progressbar(row4, mode="determinate", maximum=1000)
        self.progress.pack(fill="x", pady=(4, 0))

        # -- Log -------------------------------------------------------------
        row5 = ttk.Frame(self.root)
        row5.pack(fill="both", expand=True, padx=12, pady=(6, 12))
        ttk.Label(row5, text="Log", style="Muted.TLabel").pack(anchor="w")
        self.log = scrolledtext.ScrolledText(
            row5, wrap="word", bg=LOG_BG, fg=LOG_FG, insertbackground=LOG_FG,
            relief="flat", font=FONT_MONO, borderwidth=0, height=12,
        )
        self.log.pack(fill="both", expand=True, pady=(4, 0))
        self.log.configure(state="disabled")
        self.log.tag_configure("info",  foreground=LOG_FG)
        self.log.tag_configure("ok",    foreground=OK)
        self.log.tag_configure("warn",  foreground=WARN)
        self.log.tag_configure("error", foreground=ERROR)
        self.log.tag_configure("ts",    foreground=MUTED)

        # Seed the log with the same walkthrough so it is searchable / copyable
        self._append_log("esa-lookup ready.", "ok")
        self._append_log("Step 1: log into SAP GUI (transaction ZTBV, plant ESA1).", "info")
        self._append_log("Step 2: open the target .xlsx file in Microsoft Excel.", "info")
        self._append_log("Step 3: click 'Browse...' above and select that Excel file.", "info")
        self._append_log("Step 4: choose which Process to run.", "info")
        self._append_log("Step 5: click 'Run'. Keep SAP GUI + Excel open the whole time.", "info")

    # -- handlers ----------------------------------------------------------

    def _on_browse(self) -> None:
        path = filedialog.askopenfilename(
            title="Select Excel file",
            filetypes=[("Excel workbooks", "*.xlsx *.xlsm"), ("All files", "*.*")],
        )
        if path:
            self.file_var.set(path)

    def _on_run(self) -> None:
        path = self.file_var.get().strip()
        if not path:
            self._append_log("Pick an Excel file first.", "error")
            return
        if not os.path.exists(path):
            self._append_log(f"File does not exist:\n{path}", "error")
            return

        self._launch(pipeline.RunConfig(
            excel_path=path,
            workflow=self.workflow_var.get(),
            stop_event=self.stop_event,
            dry_run=self.dryrun_var.get(),
        ))

    def _on_diagnose(self) -> None:
        # Deliberately no file check: Diagnose never opens Excel, and the
        # operator running it is usually blocked before picking a file.
        self._launch(pipeline.RunConfig(
            excel_path="",
            workflow=self.workflow_var.get(),
            stop_event=self.stop_event,
            diagnose=True,
        ))

    def _launch(self, cfg) -> None:
        self._clear_log()
        self.stop_event = threading.Event()
        cfg.stop_event = self.stop_event
        self._set_running(True)
        self.status_var.set("starting...")
        self.progress["value"] = 0

        def target():
            pipeline.run(cfg, self._enqueue_event)

        self.worker = threading.Thread(target=target, daemon=True)
        self.worker.start()

    def _on_stop(self) -> None:
        if self.worker and self.worker.is_alive():
            self.stop_event.set()
            self._append_log(
                "Stop requested -- will halt after the current SAP step; "
                "the Excel file will not be modified.", "warn")

    # -- worker <-> UI plumbing --------------------------------------------

    def _enqueue_event(self, kind: str, payload) -> None:
        """Called from the worker thread; only touches the queue."""
        self.event_q.put((kind, payload))

    def _drain_queue(self) -> None:
        try:
            while True:
                kind, payload = self.event_q.get_nowait()
                if kind == "log":
                    msg, level = payload
                    self._append_log(msg, level)
                elif kind == "status":
                    self.status_var.set(payload)
                elif kind == "progress":
                    self.progress["value"] = float(payload) * 1000
                elif kind == "done":
                    self._set_running(False)
        except queue.Empty:
            pass
        self.root.after(50, self._drain_queue)

    def _append_log(self, msg: str, level: str = "info") -> None:
        ts = time.strftime("%H:%M:%S")
        self.log.configure(state="normal")
        self.log.insert("end", f"[{ts}] ", "ts")
        self.log.insert("end", msg + "\n", level)
        self.log.see("end")
        self.log.configure(state="disabled")

    def _clear_log(self) -> None:
        self.log.configure(state="normal")
        self.log.delete("1.0", "end")
        self.log.configure(state="disabled")

    def _set_running(self, running: bool) -> None:
        self.run_btn.state(["disabled"] if running else ["!disabled"])
        self.diag_btn.state(["disabled"] if running else ["!disabled"])
        self.stop_btn.state(["!disabled"] if running else ["disabled"])


def main() -> int:
    root = tk.Tk()
    App(root)
    root.mainloop()
    return 0


if __name__ == "__main__":
    sys.exit(main())
