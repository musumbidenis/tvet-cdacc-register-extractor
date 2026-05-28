#!/usr/bin/env python3
"""
gui.py — Professional desktop GUI for the TVET CDACC Register Extractor.

Features
--------
* Open any TVET CDACC assessment-register PDF (single or merged).
* Live activity log — real-time progress as pages are read and units found.
* Animated progress bar tracking extraction page by page.
* Info panel — extracted centre name, code, course, level.
* Metric cards — units, unique candidates, total / per-type registrations.
* Export:
    - Excel workbook  (multi-sheet: Summary, All Candidates, Roster)
    - Roster PDF      — landscape A4, filterable by report type AND course
    - Summary PDF     — portrait A4, centre info + statistics + unit table

Run:
    python gui.py

Requires: ttkbootstrap, pdfplumber, openpyxl, reportlab
"""

import os
import queue
import threading
import traceback
from datetime import datetime

import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from ttkbootstrap.dialogs import Messagebox
from tkinter import filedialog
import tkinter as tk

from extract_registers import extract, build_roster
from register_excel import build_workbook

# ── Version ───────────────────────────────────────────────────────────────────
_VERSION = "v2.0"


# ─────────────────────────────────────────────────────────────────────────────
# Report-type constants
# ─────────────────────────────────────────────────────────────────────────────

RT_ALL      = "All Report Types"
RT_ASSESS   = "Assessment Registrations"
RT_REASSESS = "Re-Assessment Registrations"


# ─────────────────────────────────────────────────────────────────────────────
# RosterPDFDialog  — select report type + courses to include
# ─────────────────────────────────────────────────────────────────────────────

class RosterPDFDialog(ttk.Toplevel):
    """
    Modal dialog for choosing report-type and course filters before
    exporting the roster PDF.
    """

    def __init__(self, parent, data):
        super().__init__(parent)
        self.title("Export Roster PDF")
        self.resizable(False, False)
        self.result = None

        roster = build_roster(data)
        self._courses: list[tuple] = list(dict.fromkeys(
            (info.get("course_name", ""), info.get("course_level", ""))
            for _, info in sorted(roster.items())
        ))

        types_present = set(
            (u.get("report_type") or "").strip() for u in data["units"]
        )
        self._has_assess   = any(t.lower().startswith("assessment") for t in types_present)
        self._has_reassess = any(t.lower().startswith("re")         for t in types_present)

        self._build()
        self.grab_set()
        self.focus_set()

        self.update_idletasks()
        px, py = parent.winfo_x(), parent.winfo_y()
        pw, ph = parent.winfo_width(), parent.winfo_height()
        w, h   = self.winfo_reqwidth(), self.winfo_reqheight()
        self.geometry(f"+{px + max(0, (pw - w) // 2)}+{py + max(0, (ph - h) // 2)}")

    def _build(self):
        outer = ttk.Frame(self, padding=(26, 20, 26, 16))
        outer.pack(fill=BOTH, expand=True)

        ttk.Label(outer, text="Export Roster PDF",
                  font=("Segoe UI", 13, "bold"), bootstyle=PRIMARY).pack(anchor=W)
        ttk.Label(outer,
                  text="Choose which registrations to include in the output PDF.",
                  bootstyle=SECONDARY).pack(anchor=W, pady=(2, 12))

        ttk.Separator(outer).pack(fill=X, pady=(0, 12))

        ttk.Label(outer, text="Report Type:",
                  font=("Segoe UI", 10, "bold")).pack(anchor=W)
        self._rt_var = ttk.StringVar(value=RT_ALL)
        for val, label in [
            (RT_ALL,      "All report types"),
            (RT_ASSESS,   "Assessment Registrations only"
                          + ("" if self._has_assess   else "  (none in this file)")),
            (RT_REASSESS, "Re-Assessment Registrations only"
                          + ("" if self._has_reassess else "  (none in this file)")),
        ]:
            state = NORMAL
            if val == RT_ASSESS   and not self._has_assess:   state = DISABLED
            if val == RT_REASSESS and not self._has_reassess: state = DISABLED
            ttk.Radiobutton(
                outer, text=label, variable=self._rt_var, value=val,
                bootstyle="primary", state=state,
            ).pack(anchor=W, padx=(14, 0), pady=1)

        ttk.Separator(outer).pack(fill=X, pady=12)

        ttk.Label(outer, text="Courses to Include:",
                  font=("Segoe UI", 10, "bold")).pack(anchor=W)

        self._all_var = ttk.BooleanVar(value=True)
        ttk.Checkbutton(
            outer, text="All courses",
            variable=self._all_var, bootstyle="primary-round-toggle",
            command=self._on_all_toggle,
        ).pack(anchor=W, padx=(14, 0), pady=(4, 2))

        self._course_checks: list[tuple] = []
        for cname, clevel in self._courses:
            if clevel and not clevel.strip().lower().startswith("level"):
                label = f"{cname}  Level {clevel}"
            elif clevel:
                label = f"{cname}  {clevel}"
            else:
                label = cname
            var = ttk.BooleanVar(value=True)
            cb  = ttk.Checkbutton(
                outer, text=label, variable=var,
                bootstyle="primary", state=DISABLED,
            )
            cb.pack(anchor=W, padx=(32, 0), pady=1)
            self._course_checks.append((var, cb, (cname, clevel)))

        ttk.Separator(outer).pack(fill=X, pady=12)

        btn_row = ttk.Frame(outer)
        btn_row.pack(fill=X)
        ttk.Button(btn_row, text="Cancel",
                   bootstyle=SECONDARY, command=self.destroy,
                   width=10).pack(side=RIGHT, padx=(6, 0))
        ttk.Button(btn_row, text="Export  →",
                   bootstyle=SUCCESS, command=self._confirm,
                   width=14).pack(side=RIGHT)

    def _on_all_toggle(self):
        all_on = self._all_var.get()
        for var, cb, _ in self._course_checks:
            cb.configure(state=DISABLED if all_on else NORMAL)
            if all_on:
                var.set(True)

    def _confirm(self):
        rt = self._rt_var.get()
        report_type = None if rt == RT_ALL else rt

        if self._all_var.get():
            courses = None
        else:
            selected = [key for var, cb, key in self._course_checks if var.get()]
            if not selected:
                Messagebox.show_warning(
                    "Please select at least one course.", "No selection", parent=self)
                return
            courses = selected

        self.result = {"report_type": report_type, "courses": courses}
        self.destroy()


# ─────────────────────────────────────────────────────────────────────────────
# Main application window
# ─────────────────────────────────────────────────────────────────────────────

class RegisterApp(ttk.Window):
    def __init__(self):
        super().__init__(
            themename="flatly",
            title="TVET CDACC Register Extractor",
            size=(1120, 760),
            minsize=(940, 640),
        )

        self.data     = None
        self.pdf_path = None
        self.q        = queue.Queue()

        self._build_header()
        self._build_file_bar()
        self._build_info_panel()
        self._build_metrics()
        # Export MUST be anchored to the bottom before the Activity panel is
        # packed, otherwise expand=True on Activity consumes all remaining space.
        self._build_statusbar()
        self._build_export_panel()
        self._build_activity_panel()

    # ══════════════════════════════════════════════════════════════════════════
    # UI Construction
    # ══════════════════════════════════════════════════════════════════════════

    def _build_header(self):
        head = ttk.Frame(self, bootstyle=DARK, padding=(20, 12, 20, 12))
        head.pack(fill=X)

        left = ttk.Frame(head, bootstyle=DARK)
        left.pack(side=LEFT, fill=Y)

        ttk.Label(
            left,
            text="TVET CDACC  Register Extractor",
            font=("Segoe UI", 16, "bold"),
            bootstyle="inverse-dark",
        ).pack(anchor=W)

        ttk.Label(
            head, text=_VERSION,
            font=("Segoe UI", 8),
            bootstyle="inverse-dark",
        ).pack(side=RIGHT, anchor=NE)

    def _build_file_bar(self):
        bar = ttk.Frame(self, padding=(20, 10, 20, 6))
        bar.pack(fill=X)

        ttk.Button(
            bar, text="📂  Open PDF…",
            bootstyle=(PRIMARY, OUTLINE),
            command=self.choose_pdf,
            width=16,
        ).pack(side=LEFT)

        self.path_var = ttk.StringVar(value="No file selected — click 'Open PDF…' to begin.")
        ttk.Label(
            bar, textvariable=self.path_var,
            bootstyle=SECONDARY,
            font=("Segoe UI", 9),
        ).pack(side=LEFT, padx=14)

    def _build_info_panel(self):
        outer = ttk.LabelFrame(self, text="  Extracted Information  ")
        outer.pack(fill=X, padx=20, pady=(2, 0))
        self._info_frame = ttk.Frame(outer, padding=(16, 6, 16, 8))
        self._info_frame.pack(fill=X)

        grid = self._info_frame
        ttk.Label(grid, text="Centre Name:",
                  font=("Segoe UI", 9, "bold"), bootstyle=PRIMARY).grid(
            row=0, column=0, sticky=W, pady=2)
        self._info_centre_name = ttk.Label(grid, text="—",
                                           font=("Segoe UI", 9), bootstyle=SECONDARY)
        self._info_centre_name.grid(row=0, column=1, sticky=W, padx=(6, 30), pady=2)

        ttk.Label(grid, text="Centre Code:",
                  font=("Segoe UI", 9, "bold"), bootstyle=PRIMARY).grid(
            row=0, column=2, sticky=W, pady=2)
        self._info_centre_code = ttk.Label(grid, text="—",
                                           font=("Segoe UI", 9), bootstyle=SECONDARY)
        self._info_centre_code.grid(row=0, column=3, sticky=W, padx=6, pady=2)

        ttk.Label(grid, text="Course(s):",
                  font=("Segoe UI", 9, "bold"), bootstyle=PRIMARY).grid(
            row=1, column=0, sticky=W, pady=2)
        self._info_course = ttk.Label(grid, text="—",
                                      font=("Segoe UI", 9), bootstyle=SECONDARY)
        self._info_course.grid(row=1, column=1, sticky=W, padx=(6, 30), pady=2)

        ttk.Label(grid, text="Level:",
                  font=("Segoe UI", 9, "bold"), bootstyle=PRIMARY).grid(
            row=1, column=2, sticky=W, pady=2)
        self._info_level = ttk.Label(grid, text="—",
                                     font=("Segoe UI", 9), bootstyle=SECONDARY)
        self._info_level.grid(row=1, column=3, sticky=W, padx=6, pady=2)

        grid.columnconfigure(1, weight=1)

    def _build_metrics(self):
        self._metrics_frame = ttk.Frame(self, padding=(20, 4, 20, 4))
        self._metrics_frame.pack(fill=X)

        self.metric_vars = {}
        cards = [
            ("units",      "Units",                       PRIMARY),
            ("candidates", "Unique Candidates",           INFO),
            ("rows",       "Total Registrations",         SUCCESS),
            ("assess",     "Assessment\nRegistrations",   WARNING),
            ("reassess",   "Re-Assessment\nRegistrations",DANGER),
        ]
        for key, label, style in cards:
            card = ttk.Frame(self._metrics_frame, padding=(12, 8), bootstyle=LIGHT)
            card.pack(side=LEFT, fill=X, expand=True, padx=4)
            var = ttk.StringVar(value="—")
            ttk.Label(card, textvariable=var,
                      font=("Segoe UI", 20, "bold"),
                      bootstyle=style).pack(anchor=W)
            ttk.Label(card, text=label,
                      font=("Segoe UI", 8),
                      bootstyle=SECONDARY).pack(anchor=W)
            self.metric_vars[key] = var

    def _build_activity_panel(self):
        """Live activity log with per-page progress bar."""
        outer = ttk.LabelFrame(self, text="  ⚡ Activity  ")
        outer.pack(fill=BOTH, expand=True, padx=20, pady=(6, 4))

        inner = ttk.Frame(outer, padding=(12, 8, 12, 10))
        inner.pack(fill=BOTH, expand=True)

        # ── Progress bar row ──────────────────────────────────────────────────
        prog_row = ttk.Frame(inner)
        prog_row.pack(fill=X, pady=(0, 4))

        self._prog_bar = ttk.Progressbar(
            prog_row, mode="determinate",
            bootstyle=(STRIPED, SUCCESS),
            maximum=100, value=0,
        )
        self._prog_bar.pack(side=LEFT, fill=X, expand=True)

        self._prog_label = ttk.Label(
            prog_row, text="Waiting for input…",
            font=("Segoe UI", 8), bootstyle=SECONDARY, width=26,
            anchor=W,
        )
        self._prog_label.pack(side=LEFT, padx=(10, 0))

        ttk.Separator(inner).pack(fill=X, pady=(4, 6))

        # ── Log text widget ───────────────────────────────────────────────────
        log_frame = ttk.Frame(inner)
        log_frame.pack(fill=BOTH, expand=True)

        self._log = tk.Text(
            log_frame,
            wrap="word",
            font=("Consolas", 9),
            bg="#F8F9FA",
            fg="#444444",
            relief="flat",
            bd=0,
            padx=8,
            pady=6,
            state="disabled",
            cursor="arrow",
            selectbackground="#BBDEFB",
        )
        log_vs = ttk.Scrollbar(log_frame, orient=VERTICAL,
                               command=self._log.yview, bootstyle=ROUND)
        self._log.configure(yscrollcommand=log_vs.set)

        self._log.grid(row=0, column=0, sticky=NSEW)
        log_vs.grid(row=0, column=1, sticky=NS)
        log_frame.rowconfigure(0, weight=1)
        log_frame.columnconfigure(0, weight=1)

        # ── Colour tags ───────────────────────────────────────────────────────
        self._log.tag_configure("ts",      foreground="#BBBBBB",
                                font=("Consolas", 8))
        self._log.tag_configure("info",    foreground="#666666")
        self._log.tag_configure("step",    foreground="#1565C0",
                                font=("Consolas", 9, "bold"))
        self._log.tag_configure("found",   foreground="#2E7D32")
        self._log.tag_configure("success", foreground="#1B5E20",
                                font=("Consolas", 9, "bold"))
        self._log.tag_configure("warn",    foreground="#E65100")
        self._log.tag_configure("error",   foreground="#B71C1C",
                                font=("Consolas", 9, "bold"))

        # Initial placeholder line
        self._log_write("info", "Open a PDF file to begin.")

    def _log_write(self, level: str, message: str):
        """Append one timestamped entry to the activity log."""
        _icons = {
            "info":    "    ",
            "step":    " >> ",
            "found":   " +  ",
            "success": " ✓  ",
            "warn":    " !  ",
            "error":   " ✗  ",
        }
        ts   = datetime.now().strftime("%H:%M:%S")
        icon = _icons.get(level, "    ")

        self._log.configure(state="normal")
        self._log.insert("end", f"{ts}", "ts")
        self._log.insert("end", f"{icon}{message}\n", level)
        self._log.see("end")
        self._log.configure(state="disabled")

    def _log_clear(self):
        self._log.configure(state="normal")
        self._log.delete("1.0", "end")
        self._log.configure(state="disabled")

    def _build_export_panel(self):
        outer = ttk.LabelFrame(self, text="  Export  ")
        outer.pack(fill=X, padx=20, pady=(4, 24), side=BOTTOM)
        export = ttk.Frame(outer, padding=(16, 6, 16, 10))
        export.pack(fill=X)

        btn_row = ttk.Frame(export)
        btn_row.pack(fill=X)

        self.export_btn = ttk.Button(
            btn_row, text="⬇  Excel Workbook…",
            bootstyle=SUCCESS, command=self.export_excel,
            state=DISABLED, width=20,
        )
        self.export_btn.pack(side=LEFT, padx=(0, 8))

        self.roster_btn = ttk.Button(
            btn_row, text="📄  Roster PDF…",
            bootstyle=INFO, command=self.export_roster_pdf,
            state=DISABLED, width=18,
        )
        self.roster_btn.pack(side=LEFT, padx=(0, 8))

        self.summary_btn = ttk.Button(
            btn_row, text="📊  Summary PDF…",
            bootstyle=(WARNING, OUTLINE), command=self.export_summary_pdf,
            state=DISABLED, width=18,
        )
        self.summary_btn.pack(side=LEFT, padx=(0, 8))

        self._export_hint = ttk.Label(
            btn_row, text="Open a PDF file to enable exports.",
            font=("Segoe UI", 8), bootstyle=SECONDARY,
        )
        self._export_hint.pack(side=LEFT, padx=8)

    def _build_statusbar(self):
        # Status bar removed — kept as no-op so references to self.status_var
        # elsewhere don't break.
        self.status_var = ttk.StringVar(value="")

    # ══════════════════════════════════════════════════════════════════════════
    # Core logic — load PDF
    # ══════════════════════════════════════════════════════════════════════════

    def choose_pdf(self):
        path = filedialog.askopenfilename(
            title="Select register PDF",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")],
        )
        if not path:
            return

        self.pdf_path = path
        self.path_var.set(os.path.basename(path))
        self._set_export_state(False)
        self._reset_ui()
        self.status_var.set("Reading PDF…")

        threading.Thread(target=self._parse_worker, args=(path,), daemon=True).start()
        self.after(50, self._poll_queue)

    def _parse_worker(self, path):
        def on_log(level, msg):
            self.q.put(("log", (level, msg)))

        def on_progress(cur, total):
            self.q.put(("progress", (cur, total)))

        try:
            data = extract(path, on_log=on_log, on_progress=on_progress)
            roster     = build_roster(data)
            total_regs = sum(u["candidate_count"] for u in data["units"])
            on_log("success",
                   f"Extraction complete — {data['unit_count']} unit(s)  ·  "
                   f"{len(roster)} unique candidate(s)  ·  "
                   f"{total_regs} total registration(s)")
            self.q.put(("parse_ok", data))
        except Exception:
            tb = traceback.format_exc()
            self.q.put(("log", ("error", "Extraction failed — check the error details below")))
            self.q.put(("parse_err", tb))

    # ══════════════════════════════════════════════════════════════════════════
    # Queue polling
    # ══════════════════════════════════════════════════════════════════════════

    def _poll_queue(self):
        # Drain up to 30 messages per tick for smooth animation
        for _ in range(30):
            try:
                kind, payload = self.q.get_nowait()
            except queue.Empty:
                break

            terminal = self._handle_msg(kind, payload)
            if terminal:
                return   # stop polling on final message

        self.after(50, self._poll_queue)

    def _handle_msg(self, kind: str, payload) -> bool:
        """Process one queue message. Returns True on a terminal message."""

        if kind == "log":
            level, text = payload
            self._log_write(level, text)
            return False

        if kind == "progress":
            cur, total = payload
            self._prog_bar.configure(maximum=total, value=cur)
            self._prog_label.configure(text=f"Page {cur} / {total}")
            return False

        # ── Terminal: parse done ───────────────────────────────────────────────
        if kind == "parse_err":
            self._prog_label.configure(text="Failed")
            self.status_var.set("Failed to read PDF.")
            Messagebox.show_error(
                f"Could not extract data:\n\n{payload}",
                "Extraction error", parent=self)
            return True

        if kind == "parse_ok":
            self._prog_label.configure(text="Complete ✓")
            self._populate(payload)
            return True

        # ── Terminal: export done ─────────────────────────────────────────────
        if kind == "excel_ok":
            path = payload
            self._prog_label.configure(text="Saved ✓")
            self.status_var.set(f"Saved: {os.path.basename(path)}")
            if Messagebox.show_question(
                    f"Workbook saved to:\n{path}\n\nOpen it now?",
                    "Done", buttons=["No:secondary", "Yes:success"],
                    parent=self) == "Yes":
                self._open_file(path)
            return True

        if kind == "excel_err":
            self._prog_label.configure(text="Failed")
            self.status_var.set("Excel export failed.")
            Messagebox.show_error(payload, "Export error", parent=self)
            return True

        if kind == "pdf_ok":
            path = payload
            self._prog_label.configure(text="Saved ✓")
            self.status_var.set(f"PDF saved: {os.path.basename(path)}")
            if Messagebox.show_question(
                    f"PDF saved to:\n{path}\n\nOpen it now?",
                    "Done", buttons=["No:secondary", "Yes:success"],
                    parent=self) == "Yes":
                self._open_file(path)
            return True

        if kind == "pdf_err":
            self._prog_label.configure(text="Failed")
            self.status_var.set("PDF export failed.")
            Messagebox.show_error(payload, "PDF Export error", parent=self)
            return True

        return False

    # ══════════════════════════════════════════════════════════════════════════
    # Populate info + metrics after parse
    # ══════════════════════════════════════════════════════════════════════════

    def _populate(self, data):
        self.data  = data
        roster     = build_roster(data)
        total      = sum(u["candidate_count"] for u in data["units"])
        assess_regs = sum(
            u["candidate_count"] for u in data["units"]
            if (u.get("report_type") or "").strip().lower().startswith("assessment")
        )
        reassess_regs = total - assess_regs

        # Info panel
        centre_name = data.get("centre_name") or "—"
        centre_code = data.get("centre_code") or "—"
        seen = list(dict.fromkeys(
            (u.get("course_name", ""), u.get("course_level", ""))
            for u in data["units"]
        ))
        course_str = "; ".join(cn for cn, _ in seen if cn) or (data.get("course_name") or "—")
        level_str  = "; ".join(sorted({cl for _, cl in seen if cl})) or (data.get("course_level") or "—")

        self._info_centre_name.configure(text=centre_name)
        self._info_centre_code.configure(text=centre_code)
        self._info_course.configure(text=course_str)
        self._info_level.configure(text=level_str)

        # Metric cards
        self.metric_vars["units"].set(str(data["unit_count"]))
        self.metric_vars["candidates"].set(str(len(roster)))
        self.metric_vars["rows"].set(str(total))
        self.metric_vars["assess"].set(str(assess_regs))
        self.metric_vars["reassess"].set(str(reassess_regs))

        self._set_export_state(True)
        self.status_var.set(
            f"{centre_name}  ·  {data['unit_count']} units  ·  "
            f"{len(roster)} unique candidates  ·  {total} total registrations  "
            f"({assess_regs} Assessment / {reassess_regs} Re-Assessment)"
        )

    def _reset_ui(self):
        """Clear activity log, progress bar, info labels and metrics."""
        self._log_clear()
        self._prog_bar.configure(maximum=100, value=0)
        self._prog_label.configure(text="Starting…")
        for var in self.metric_vars.values():
            var.set("—")
        self._info_centre_name.configure(text="—")
        self._info_centre_code.configure(text="—")
        self._info_course.configure(text="—")
        self._info_level.configure(text="—")

    def _set_export_state(self, enabled: bool):
        s = NORMAL if enabled else DISABLED
        self.export_btn.configure(state=s)
        self.roster_btn.configure(state=s)
        self.summary_btn.configure(state=s)
        self._export_hint.configure(
            text="" if enabled else "Open a PDF file to enable exports."
        )

    # ══════════════════════════════════════════════════════════════════════════
    # Export — Excel
    # ══════════════════════════════════════════════════════════════════════════

    def export_excel(self):
        if not self.data:
            return
        default = os.path.splitext(os.path.basename(self.pdf_path))[0] + "_extracted.xlsx"
        path = filedialog.asksaveasfilename(
            title="Save Excel Workbook", defaultextension=".xlsx",
            initialfile=default,
            filetypes=[("Excel workbook", "*.xlsx")],
        )
        if not path:
            return
        self._prep_export("Building Excel workbook…")
        threading.Thread(target=self._excel_worker, args=(path,), daemon=True).start()
        self.after(50, self._poll_queue)

    def _excel_worker(self, path):
        def _log(level, msg):
            self.q.put(("log", (level, msg)))
        try:
            _log("step",  "Building Excel workbook…")
            self.q.put(("progress", (20, 100)))
            _log("info",  "Writing Summary sheet…")
            self.q.put(("progress", (40, 100)))
            _log("info",  "Writing All Candidates sheet…")
            self.q.put(("progress", (65, 100)))
            _log("info",  "Writing Roster sheet…")
            build_workbook(self.data, path)
            self.q.put(("progress", (100, 100)))
            _log("success", f"Workbook saved → {os.path.basename(path)}")
            self.q.put(("excel_ok", path))
        except PermissionError:
            _log("error", "Permission denied — close the file in Excel and retry")
            self.q.put(("excel_err",
                        f"Cannot save — permission denied:\n{path}\n\n"
                        "The file is probably open in Excel. Close it and try again."))
        except Exception:
            tb = traceback.format_exc()
            _log("error", "Workbook export failed")
            self.q.put(("excel_err", tb))

    # ══════════════════════════════════════════════════════════════════════════
    # Export — Roster PDF
    # ══════════════════════════════════════════════════════════════════════════

    def export_roster_pdf(self):
        if not self.data:
            return

        dlg = RosterPDFDialog(self, self.data)
        self.wait_window(dlg)
        if dlg.result is None:
            return
        report_type_filter = dlg.result["report_type"]
        course_filter      = dlg.result["courses"]

        stem = os.path.splitext(os.path.basename(self.pdf_path))[0]
        if report_type_filter:
            slug    = "assessment" if report_type_filter.lower().startswith("assessment") else "reassessment"
            default = f"{stem}_{slug}_roster.pdf"
        elif course_filter and len(course_filter) == 1:
            cslug   = course_filter[0][0].replace(" ", "_")[:18]
            default = f"{stem}_{cslug}_roster.pdf"
        else:
            default = f"{stem}_roster.pdf"

        path = filedialog.asksaveasfilename(
            title="Save Roster PDF", defaultextension=".pdf",
            initialfile=default,
            filetypes=[("PDF files", "*.pdf")],
        )
        if not path:
            return

        self._prep_export("Generating Roster PDF…")
        threading.Thread(
            target=self._pdf_worker,
            args=(path, course_filter, report_type_filter),
            daemon=True,
        ).start()
        self.after(50, self._poll_queue)

    def _pdf_worker(self, path, course_filter, report_type_filter=None):
        def _log(level, msg):
            self.q.put(("log", (level, msg)))
        try:
            from roster_pdf import build_roster_pdf
            _log("step", "Building Roster PDF…")
            self.q.put(("progress", (25, 100)))
            _log("info", "Laying out candidate tables…")
            self.q.put(("progress", (60, 100)))
            _log("info", "Rendering pages…")
            build_roster_pdf(
                self.data, path,
                course_filter=course_filter,
                report_type_filter=report_type_filter,
            )
            self.q.put(("progress", (100, 100)))
            _log("success", f"Roster PDF saved → {os.path.basename(path)}")
            self.q.put(("pdf_ok", path))
        except PermissionError:
            _log("error", "Permission denied — close the PDF in your viewer and retry")
            self.q.put(("pdf_err",
                        f"Cannot save — permission denied:\n{path}\n\n"
                        "The file may be open in a PDF viewer. Close it and try again."))
        except Exception:
            tb = traceback.format_exc()
            _log("error", "Roster PDF export failed")
            self.q.put(("pdf_err", tb))

    # ══════════════════════════════════════════════════════════════════════════
    # Export — Summary PDF
    # ══════════════════════════════════════════════════════════════════════════

    def export_summary_pdf(self):
        if not self.data:
            return

        stem    = os.path.splitext(os.path.basename(self.pdf_path))[0]
        default = f"{stem}_summary.pdf"

        path = filedialog.asksaveasfilename(
            title="Save Summary PDF", defaultextension=".pdf",
            initialfile=default,
            filetypes=[("PDF files", "*.pdf")],
        )
        if not path:
            return

        self._prep_export("Generating Summary PDF…")
        threading.Thread(target=self._summary_pdf_worker,
                         args=(path,), daemon=True).start()
        self.after(50, self._poll_queue)

    def _summary_pdf_worker(self, path):
        def _log(level, msg):
            self.q.put(("log", (level, msg)))
        try:
            from roster_pdf import build_summary_pdf
            _log("step", "Building Summary PDF…")
            self.q.put(("progress", (30, 100)))
            _log("info", "Composing statistics and unit table…")
            self.q.put(("progress", (65, 100)))
            _log("info", "Rendering document…")
            build_summary_pdf(self.data, path)
            self.q.put(("progress", (100, 100)))
            _log("success", f"Summary PDF saved → {os.path.basename(path)}")
            self.q.put(("pdf_ok", path))
        except PermissionError:
            _log("error", "Permission denied — close the PDF in your viewer and retry")
            self.q.put(("pdf_err",
                        f"Cannot save — permission denied:\n{path}\n\n"
                        "The file may be open in a PDF viewer. Close it and try again."))
        except Exception:
            tb = traceback.format_exc()
            _log("error", "Summary PDF export failed")
            self.q.put(("pdf_err", tb))

    # ══════════════════════════════════════════════════════════════════════════
    # Helpers
    # ══════════════════════════════════════════════════════════════════════════

    def _prep_export(self, label: str):
        """Reset progress bar and log a header line before starting an export."""
        self._prog_bar.configure(maximum=100, value=0)
        self._prog_label.configure(text=label)
        self._log_write("step", f"── {label} ──")

    @staticmethod
    def _open_file(path):
        import subprocess
        try:
            if os.name == "nt":
                os.startfile(path)          # type: ignore[attr-defined]
            elif os.uname().sysname == "Darwin":
                subprocess.run(["open", path])
            else:
                subprocess.run(["xdg-open", path])
        except Exception:
            pass


if __name__ == "__main__":
    RegisterApp().mainloop()
