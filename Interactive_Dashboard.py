"""COPQ Interactive Dashboard - filters, checkboxes, live charts."""
import os
import sys
import datetime
import threading

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="backslashreplace")

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import pandas as pd
import numpy as np
import matplotlib
matplotlib.use("TkAgg")
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)

# Import process_ftt BEFORE combine_files: combine_files prepends "Source/"
# to sys.path (for monthly_reporting), which would otherwise shadow the root
# process_ftt.py with the old Source/process_ftt.py (no generate_ftt_report).
import process_ftt
import combine_files
from monthly_reporting import classify_copq_type
DEFAULT_INPUT = os.path.join(BASE_DIR, "Output", "COPQ_Clean.xlsx")
DEFAULT_COPQ_INPUT_DIR = os.path.join(BASE_DIR, "COPQ_Input")
DEFAULT_FTT_INPUT_DIR = os.path.join(BASE_DIR, "FTT_Input")
DEFAULT_OUTPUT_DIR = os.path.join(BASE_DIR, "Output")
DEFAULT_COLOR_MAP = os.path.join(BASE_DIR, "BC_Color", "Color_Defect.xlsx")
DEFAULT_FTT_REPORT = os.path.join(DEFAULT_OUTPUT_DIR, "FTT_BC_Report.xlsx")
DEFAULT_FTT_REPORT_ALT = os.path.join(DEFAULT_OUTPUT_DIR, "FTT_Combined_Report.xlsx")
DEFAULT_PPTX_SCRIPT = os.path.join(DEFAULT_OUTPUT_DIR, "update_copq_pptx.py")
DEFAULT_DB_SCRIPT   = os.path.join(DEFAULT_OUTPUT_DIR, "write_db_month.py")
DEFAULT_PPTX_OUTPUT = os.path.join(DEFAULT_OUTPUT_DIR, "COPQ_Report_Sep_2026.pptx")

SITES = ["VH", "VH2", "VH3", "VH4", "JV", "JV2", "JV3", "JVB"]
TYPES = ["Touch-up", "Reinspection", "B/C"]
TYPE_COLORS = {"Touch-up": "#38BDF8",
               "Reinspection": "#FBBF24", "B/C": "#FB7185"}

BG = "#0B132B"; CARD = "#1C2541"; INNER = "#273549"
TXT = "#F8FAFC"; MUT = "#94A3B8"; ACC = "#38BDF8"


def classify(row):
    """Classify a record using standard CoPQ SOP classification rules."""
    return classify_copq_type(row)


class ProgressDialog:
    """Modal progress dialog with a smooth 'creeping' progress bar.

    Owns the Toplevel window, the tick loop and the finish/fail transitions,
    so the caller only has to drive step callbacks from a worker thread.
    """

    def __init__(self, master, steps):
        self.master = master
        self.total = max(len(steps), 1)
        self.state = {"val": 0.0, "running": False}

        dlg = tk.Toplevel(master)
        dlg.title("Pipeline Progress")
        dlg.configure(bg=CARD)
        dlg.resizable(False, False)
        dlg.transient(master)
        dlg.grab_set()
        dlg.geometry("+{}+{}".format(
            master.winfo_rootx() + max((master.winfo_width() - 420) // 2, 0),
            master.winfo_rooty() + max((master.winfo_height() - 160) // 2, 0)))
        self.dlg = dlg

        tk.Label(dlg, text="RUNNING PIPELINE", font=("Segoe UI", 11, "bold"),
                 fg=ACC, bg=CARD).pack(padx=24, pady=(16, 4))
        self.step_lbl = tk.Label(dlg, text="Starting...", font=("Segoe UI", 9),
                                 fg=TXT, bg=CARD, wraplength=380)
        self.step_lbl.pack(padx=24, pady=(0, 8))
        self.pb = ttk.Progressbar(dlg, orient="horizontal", length=400,
                                  mode="determinate", maximum=self.total)
        self.pb.pack(padx=24, pady=(0, 4))
        self.pct_lbl = tk.Label(dlg, text="0%", font=("Segoe UI", 9, "bold"),
                                fg=MUT, bg=CARD)
        self.pct_lbl.pack(pady=(0, 16))

    def _render(self):
        v = min(self.state["val"], self.total)
        self.pb.configure(value=v)
        self.pct_lbl.configure(text=f"{int(v / self.total * 100)}%")

    def _tick(self):
        """Advance progress smoothly toward the next milestone."""
        if self.state["running"]:
            target = min(int(self.state["val"]) + 1, self.total - 0.05)
            remaining = target - self.state["val"]
            if remaining > 0.01:
                self.state["val"] += max(remaining * 0.08, 0.006)
                self._render()
        self.dlg.after(100, self._tick)

    def start_step(self, i, label):
        self.step_lbl.configure(text=f"[{i + 1}/{self.total}] {label} ...")
        self.state["val"] = float(i)
        self.state["running"] = True
        self._render()
        if not getattr(self.dlg, "_ticking", False):
            self.dlg._ticking = True
            self.dlg.after(100, self._tick)

    def finish_step(self, i):
        self.state["val"] = float(i + 1)
        self.state["running"] = False
        self._render()

    def finish_all(self):
        self.state["val"] = float(self.total)
        self.state["running"] = False
        self._render()

    def set_done(self, message):
        self.step_lbl.configure(text="Done.")

    def fail(self, message):
        self.step_lbl.configure(text=f"Failed: {message}", fg="#FB7185")
        self.state["running"] = False

    def close(self, delay_ms=300):
        self.dlg.after(delay_ms, self.dlg.destroy)


class Dashboard(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("COPQ Interactive Dashboard")
        self.geometry("1360x900")
        self.minsize(1150, 780)
        self.configure(bg=BG)

        self.file_path = tk.StringVar(value=DEFAULT_INPUT)
        self.ftt_files_path = tk.StringVar(value="")
        self.df = None
        self.filtered = None
        self.type_vars = {t: tk.BooleanVar(value=True) for t in TYPES}
        self.site_vars = {s: tk.BooleanVar(value=True) for s in SITES}
        self.month_var = tk.StringVar(value="All")
        self.metric_var = tk.StringVar(value="Cost ($)")
        self.METRICS = ["Cost ($)", "Qty (pairs)", "All"]
        self.status = tk.StringVar(value="Load a workbook to begin.")

        self._style()
        self._build()
        if os.path.exists(DEFAULT_INPUT):
            self.load_data()
        self._clear_ftt_report_view()

    # ---------- styling ----------
    def _style(self):
        st = ttk.Style(self)
        try:
            st.theme_use("clam")
        except Exception:
            pass
        st.configure(".", background=BG, foreground=TXT, font=("Segoe UI", 10))
        st.configure("TNotebook", background=BG, borderwidth=0)
        st.configure("TNotebook.Tab", background=CARD, foreground=MUT,
                     padding=[14, 7], font=("Segoe UI", 9, "bold"))
        st.map("TNotebook.Tab", background=[("selected", INNER)],
               foreground=[("selected", ACC)])
        st.configure("TCheckbutton", background=INNER, foreground=TXT)
        st.map("TCheckbutton", background=[("active", INNER)])
        st.configure("Treeview", background="#1E293B", foreground=TXT,
                     fieldbackground="#1E293B", rowheight=24, borderwidth=0)
        st.configure("Treeview.Heading", background="#0F172A", foreground=ACC,
                     font=("Segoe UI", 9, "bold"), relief="flat")
        st.configure("Accent.TButton", font=("Segoe UI", 10, "bold"),
                     padding=[12, 6])
        st.configure("TCombobox", fieldbackground=INNER, background=INNER,
                     foreground=TXT)

    # ---------- layout ----------
    def _build(self):
        header = tk.Frame(self, bg=CARD)
        header.pack(fill="x", padx=12, pady=(10, 4))
        tk.Label(header, text="COPQ INTERACTIVE DASHBOARD",
                 font=("Segoe UI", 15, "bold"), fg=ACC, bg=CARD).pack(side="left", padx=16, pady=10)
        tk.Label(header, textvariable=self.status, font=("Segoe UI", 9),
                 fg=MUT, bg=CARD).pack(side="right", padx=16)

        # filter bar
        bar = tk.Frame(self, bg=INNER, padx=14, pady=10)
        bar.pack(fill="x", padx=12, pady=4)

        tk.Label(bar, text="Data file:", bg=INNER, fg=TXT,
                 font=("Segoe UI", 9, "bold")).pack(side="left")
        tk.Entry(bar, textvariable=self.file_path, width=44, bg=CARD,
                 fg=TXT, relief="flat").pack(side="left", padx=6)
        ttk.Button(bar, text="Browse", command=self.browse).pack(side="left", padx=(0, 12))
        ttk.Button(bar, text="Reload Data", style="Accent.TButton",
                   command=self.load_data).pack(side="left", padx=(0, 18))

        tk.Label(bar, text="Month:", bg=INNER, fg=TXT,
                 font=("Segoe UI", 9, "bold")).pack(side="left")
        self.month_cb = ttk.Combobox(bar, textvariable=self.month_var,
                                     state="readonly", width=10)
        self.month_cb.pack(side="left", padx=6)
        self.month_cb.bind("<<ComboboxSelected>>", lambda e: self.apply_filters())

        tk.Label(bar, text="Metric:", bg=INNER, fg=TXT,
                 font=("Segoe UI", 9, "bold")).pack(side="left", padx=(14, 0))
        self.metric_cb = ttk.Combobox(bar, textvariable=self.metric_var, state="readonly",
                                      width=12, values=self.METRICS)
        self.metric_cb.pack(side="left", padx=6)
        self.metric_cb.bind("<<ComboboxSelected>>", lambda e: self.apply_filters())

        # checkbox groups
        groups = tk.Frame(self, bg=BG)
        groups.pack(fill="x", padx=12, pady=2)

        t_box = tk.LabelFrame(groups, text="Defect Type Filters", bg=INNER,
                              fg=ACC, font=("Segoe UI", 9, "bold"), padx=10, pady=4)
        t_box.pack(side="left", fill="x", expand=True, padx=(0, 6))
        for t in TYPES:
            cb = tk.Checkbutton(t_box, text=t, variable=self.type_vars[t],
                                command=self.apply_filters, bg=INNER, fg=TXT,
                                activebackground=INNER, selectcolor="#0B132B",
                                font=("Segoe UI", 9))
            cb.pack(side="left", padx=8)

        s_box = tk.LabelFrame(groups, text="Site Filters", bg=INNER, fg=ACC,
                              font=("Segoe UI", 9, "bold"), padx=10, pady=4)
        s_box.pack(side="left", fill="x", expand=True, padx=(6, 0))
        for i, s in enumerate(SITES):
            cb = tk.Checkbutton(s_box, text=s, variable=self.site_vars[s],
                                command=self.apply_filters, bg=INNER, fg=TXT,
                                activebackground=INNER, selectcolor="#0B132B",
                                font=("Segoe UI", 9))
            cb.grid(row=0, column=i, padx=5)
        ttk.Button(s_box, text="All / None", command=self.toggle_sites).grid(row=0, column=len(SITES), padx=8)

        # KPI cards
        kpis = tk.Frame(self, bg=BG)
        kpis.pack(fill="x", padx=12, pady=(8, 2))
        self.kpi_total = self._kpi(kpis, "TOTAL FILTERED", ACC)
        self.kpi_tu = self._kpi(kpis, "TOUCH-UP", "#38BDF8")
        self.kpi_re = self._kpi(kpis, "REINSPECTION", "#FBBF24")
        self.kpi_bc = self._kpi(kpis, "B/C GRADE", "#FB7185")
        self.kpi_rows = self._kpi(kpis, "RECORDS", "#34D399")

        # automation / pipeline row
        auto = tk.LabelFrame(self, text="Automation — run pipeline & export to Output",
                             bg=INNER, fg=ACC, font=("Segoe UI", 9, "bold"), padx=10, pady=6)
        auto.pack(fill="x", padx=12, pady=(8, 2))
        self.btn_combine = ttk.Button(auto, text="1. Combine COPQ (combine_files.py)",
                                      style="Accent.TButton", command=self.run_combine)
        self.btn_combine.pack(side="left", padx=(0, 8))
        self.btn_ftt = ttk.Button(auto, text="2. Process FTT (process_ftt.py)",
                                  style="Accent.TButton", command=self.run_ftt)
        self.btn_ftt.pack(side="left", padx=(0, 8))
        self.btn_all = ttk.Button(auto, text="Run Full Pipeline",
                                  style="Success.TButton", command=self.run_full_pipeline)
        self.btn_all.pack(side="left", padx=(0, 8))
        self.btn_pptx = ttk.Button(auto, text="3. Refresh PPTX Report",
                                    style="Success.TButton", command=self.run_pptx)
        self.btn_pptx.pack(side="left", padx=(0, 8))
        ttk.Button(auto, text="Open Output Folder",
                   command=lambda: os.startfile(DEFAULT_OUTPUT_DIR) if os.path.exists(DEFAULT_OUTPUT_DIR) else messagebox.showinfo("Not Found", "Output folder does not exist yet.")).pack(side="left")

        # FTT input files row (optional multi-file selection)
        ftt_row = tk.Frame(self, bg=INNER, padx=14, pady=6)
        ftt_row.pack(fill="x", padx=12, pady=(2, 2))
        tk.Label(ftt_row, text="FTT data files:", bg=INNER, fg=TXT,
                 font=("Segoe UI", 9, "bold")).pack(side="left")
        tk.Entry(ftt_row, textvariable=self.ftt_files_path, bg=CARD,
                 fg=MUT, relief="flat").pack(side="left", padx=6, fill="x", expand=True)
        ttk.Button(ftt_row, text="Browse FTT Files…",
                   command=self.browse_ftt_files).pack(side="left", padx=(0, 8))
        ttk.Button(ftt_row, text="Reload FTT Data",
               command=self.reload_ftt_data).pack(side="left", padx=(0, 8))
        ttk.Button(ftt_row, text="Clear",
                   command=self.clear_ftt_files).pack(side="left")

        # tabs: charts + data table
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill="both", expand=True, padx=12, pady=6)

        tab_charts = ttk.Frame(self.notebook)
        tab_ftt = ttk.Frame(self.notebook)
        tab_table = ttk.Frame(self.notebook)
        self.notebook.add(tab_charts, text="Charts")
        self.notebook.add(tab_ftt, text="FTT B/C Report")
        self.notebook.add(tab_table, text="Filtered Data")

        # --- FTT B/C tab: charts on top, summary table below ---
        self.ftt_fig, self.ftt_axes = plt.subplots(1, 2, figsize=(13, 3.6),
                                                   facecolor=CARD)
        self.ftt_fig.subplots_adjust(wspace=0.35, left=0.18, right=0.97,
                                     top=0.88, bottom=0.14)
        self.ftt_canvas = FigureCanvasTkAgg(self.ftt_fig, master=tab_ftt)
        self.ftt_canvas.get_tk_widget().pack(fill="both", expand=True)

        self.ftt_status = tk.Label(tab_ftt, text="No FTT report loaded.",
                                   font=("Segoe UI", 9), fg=MUT, bg=CARD)
        self.ftt_status.pack(fill="x", padx=8)

        ftt_tbl = tk.Frame(tab_ftt, bg=BG)
        ftt_tbl.pack(fill="both", expand=True)
        self.ftt_tree = ttk.Treeview(ftt_tbl, show="headings", height=8)
        fys = ttk.Scrollbar(ftt_tbl, orient="vertical", command=self.ftt_tree.yview)
        fxs = ttk.Scrollbar(ftt_tbl, orient="horizontal", command=self.ftt_tree.xview)
        self.ftt_tree.configure(yscrollcommand=fys.set, xscrollcommand=fxs.set)
        self.ftt_tree.grid(row=0, column=0, sticky="nsew")
        fys.grid(row=0, column=1, sticky="ns")
        fxs.grid(row=1, column=0, sticky="ew")
        ftt_tbl.rowconfigure(0, weight=1)
        ftt_tbl.columnconfigure(0, weight=1)

        self.fig, self.axes = plt.subplots(2, 2, figsize=(13, 6.2), facecolor=CARD)
        self.fig.subplots_adjust(hspace=0.45, wspace=0.3,
                                 left=0.06, right=0.97, top=0.92, bottom=0.1)
        self.canvas = FigureCanvasTkAgg(self.fig, master=tab_charts)
        self.canvas.get_tk_widget().pack(fill="both", expand=True)

        self.tree = ttk.Treeview(tab_table, show="headings")
        ys = ttk.Scrollbar(tab_table, orient="vertical", command=self.tree.yview)
        xs = ttk.Scrollbar(tab_table, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=ys.set, xscrollcommand=xs.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        ys.grid(row=0, column=1, sticky="ns")
        xs.grid(row=1, column=0, sticky="ew")
        tab_table.rowconfigure(0, weight=1)
        tab_table.columnconfigure(0, weight=1)

    def _kpi(self, parent, title, color):
        card = tk.Frame(parent, bg=CARD, padx=14, pady=8)
        card.pack(side="left", fill="both", expand=True, padx=3)
        tk.Label(card, text=title, font=("Segoe UI", 8, "bold"),
                 fg=MUT, bg=CARD).pack(anchor="w")
        lbl = tk.Label(card, text="$0.00", font=("Segoe UI", 15, "bold"),
                       fg=color, bg=CARD)
        lbl.pack(anchor="w")
        return lbl

    # ---------- data ----------
    def browse(self):
        paths = filedialog.askopenfilenames(
            filetypes=[("Excel", "*.xlsx *.xls"), ("All", "*.*")])
        if paths:
            self.file_path.set(" | ".join(paths))
            self.load_data()

    def browse_ftt_files(self):
        # Accept either four raw FTT files or one already-combined FTT file.
        paths = filedialog.askopenfilenames(
            title="Select FTT data files",
            filetypes=[("Excel/CSV", "*.xlsx *.xls *.csv"), ("All", "*.*")])
        if paths:
            self.ftt_files_path.set(" | ".join(paths))
            if len(paths) not in (1, 4):
                self._clear_ftt_report_view()
                self.status.set("Select 1 combined FTT file or exactly 4 FTT files.")
            else:
                self._clear_ftt_report_view()
                count_label = "1 combined FTT file" if len(paths) == 1 else "4 FTT files"
                self.status.set(f"{count_label} selected — click '2. Process FTT' to generate the report.")

    def _selected_ftt_files(self):
        return [p.strip() for p in self.ftt_files_path.get().split("|") if p.strip()]

    def clear_ftt_files(self):
        self.ftt_files_path.set("")
        self._clear_ftt_report_view()
        self.status.set("Select 1 combined FTT file or 4 FTT files to load the current report.")

    def reload_ftt_data(self):
        selected = self._selected_ftt_files()
        if len(selected) not in (1, 4) or any(not os.path.exists(p) for p in selected):
            self._clear_ftt_report_view()
            messagebox.showwarning(
            "FTT input required",
            "Select 1 combined FTT file or exactly 4 existing FTT files before reloading.")
            return
        self.status.set("Reloading FTT data...")
        self.load_ftt_report()

    def _clear_ftt_report_view(self):
        if not hasattr(self, "ftt_tree"):
            return
        self.ftt_tree.delete(*self.ftt_tree.get_children())
        self.ftt_tree.configure(columns=[])
        for ax in self.ftt_axes:
            ax.clear()
            self._dark_axis(ax)
        self.ftt_axes[0].set_title("Top 12 Defects by B/C Qty", color=TXT,
                                   fontsize=10, fontweight="bold")
        self.ftt_axes[1].set_title("B/C Qty by Site", color=TXT,
                                   fontsize=10, fontweight="bold")
        self.ftt_status.configure(text="Select 1 combined FTT file or 4 FTT files and click '2. Process FTT'.",
                                  fg=MUT)
        self.ftt_canvas.draw_idle()

    # ---------- pipeline automation ----------
    def _set_pipeline_busy(self, busy):
        state = "disabled" if busy else "normal"
        for b in (self.btn_combine, self.btn_ftt, self.btn_all, self.btn_pptx):
            b.configure(state=state)

    def _run_in_background(self, steps, done_msg):
        """Run steps in a background thread with a modal progress dialog."""
        self._set_pipeline_busy(True)
        progress = ProgressDialog(self, steps)

        def worker():
            try:
                for i, (label, func) in enumerate(steps):
                    self.after(0, lambda i=i, l=label: progress.start_step(i, l))
                    print(f"\n===== {label} =====")
                    func()
                    self.after(0, lambda i=i: progress.finish_step(i))
                self.after(0, progress.finish_all)
                self.after(0, lambda: progress.set_done(done_msg))
                self.after(0, lambda: self.status.set(done_msg))
                self.after(300, lambda: progress.close())
                self.after(310, lambda: messagebox.showinfo("Pipeline Complete", done_msg))
            except Exception as exc:
                msg = f"Pipeline failed: {exc}"
                print(f"[!] {msg}")
                self.after(0, lambda e=exc: progress.fail(e))
                self.after(600, lambda: progress.close())
                self.after(610, lambda m=msg: messagebox.showerror("Pipeline Error", m))
            finally:
                self._set_pipeline_busy(False)

        threading.Thread(target=worker, daemon=True).start()

    def run_combine(self):
        out = os.path.join(DEFAULT_OUTPUT_DIR, "COPQ_Clean.xlsx")
        os.makedirs(DEFAULT_OUTPUT_DIR, exist_ok=True)

        def step():
            combine_files.combine_copq_files(input_dir=DEFAULT_COPQ_INPUT_DIR, output_file=out)
            self.file_path.set(out)
            self.load_data()

        self._run_in_background(
            [("Combine COPQ raw files", step)],
            f"COPQ combined successfully → {out}")

    def _run_ftt_generation(self, out):
        # run process_ftt.generate_ftt_report - same entry point as process_ftt.bat
        os.makedirs(DEFAULT_OUTPUT_DIR, exist_ok=True)
        # Optional explicit multi-file selection from the dashboard
        selected = [p.strip() for p in self.ftt_files_path.get().split("|") if p.strip()]
        process_ftt.generate_ftt_report(
            input_folder=DEFAULT_FTT_INPUT_DIR,
            output_file=out,
            color_map_file=DEFAULT_COLOR_MAP,
            file_paths=selected or None)

    def run_ftt(self):
        selected = self._selected_ftt_files()
        missing = [p for p in selected if not os.path.exists(p)]
        if len(selected) not in (1, 4) or missing:
            messagebox.showwarning(
                "FTT input required",
                "Select 1 combined FTT file or exactly 4 existing FTT files before processing.")
            return
        out = os.path.join(DEFAULT_OUTPUT_DIR, "FTT_Combined_Report.xlsx")
        self._run_in_background(
            [("Process FTT report",
              lambda: self._run_ftt_generation(out)),
             ("Refresh FTT tab", self.load_ftt_report)],
            f"FTT report generated → {out}")

    def run_full_pipeline(self):
        selected = self._selected_ftt_files()
        missing = [p for p in selected if not os.path.exists(p)]
        if len(selected) not in (1, 4) or missing:
            messagebox.showwarning(
                "FTT input required",
                "Select 1 combined FTT file or exactly 4 existing FTT files before running the full pipeline.")
            return
        copq_out = os.path.join(DEFAULT_OUTPUT_DIR, "COPQ_Clean.xlsx")
        ftt_out = os.path.join(DEFAULT_OUTPUT_DIR, "FTT_Combined_Report.xlsx")
        os.makedirs(DEFAULT_OUTPUT_DIR, exist_ok=True)

        def combine_step():
            combine_files.combine_copq_files(input_dir=DEFAULT_COPQ_INPUT_DIR, output_file=copq_out)

        def ftt_step():
            self._run_ftt_generation(ftt_out)

        def reload_step():
            self.file_path.set(copq_out)
            self.load_data()

        def ftt_refresh_step():
            self.load_ftt_report()

        self._run_in_background(
            [("Step 1/4 - Combine COPQ", combine_step),
             ("Step 2/4 - Process FTT", ftt_step),
             ("Step 3/4 - Refresh dashboard data", reload_step),
             ("Step 4/4 - Refresh FTT view", ftt_refresh_step)],
            "Full pipeline completed! COPQ_Clean.xlsx & FTT_Combined_Report.xlsx are in Output/")

    def run_pptx(self):
        """Refresh the COPQ PPTX report.

        Step 1: run write_db_month.py  -> maps COPQ_Clean + FTT_Combined_Report
                into the Database workbooks (CoPQ database 25.xlsx,
                CoPQ_type_analysis.xlsx), the accumulated source of truth.
        Step 2: run update_copq_pptx.py -> clones the template and fills the
                current period from the Database. So a user only needs to click
                Refresh after the COPQ/FTT pipeline has produced new data."""
        db_script = DEFAULT_DB_SCRIPT
        pptx_script = DEFAULT_PPTX_SCRIPT
        for script in (db_script, pptx_script):
            if not os.path.exists(script):
                messagebox.showerror(
                    "Script Not Found",
                    f"Could not find the script:\n{script}\n\n"
                    "Make sure write_db_month.py and update_copq_pptx.py are "
                    "present in the Output folder.")
                return
        # python-pptx is required by the generator script.
        try:
            import pptx  # noqa: F401
        except Exception:
            if messagebox.askyesno("Install dependency",
                                   "python-pptx is required to generate the "
                                   "report. Install it now?"):
                import subprocess
                subprocess.run([sys.executable, "-m", "pip", "install", "python-pptx"],
                               check=False)
            else:
                return

        out = DEFAULT_PPTX_OUTPUT

        def db_load_step():
            import subprocess
            # Step 1: load this period's data into the Database workbooks.
            r1 = subprocess.run([sys.executable, db_script],
                                capture_output=True, text=True,
                                cwd=os.path.dirname(db_script), timeout=600)
            if r1.returncode != 0:
                raise RuntimeError(
                    "Database load failed (exit %d):\n\n%s" %
                    (r1.returncode, (r1.stdout + "\n" + r1.stderr)[-2500:]))

        def pptx_gen_step():
            import subprocess
            # Step 2: regenerate the PPTX from the Database (source of truth).
            r2 = subprocess.run([sys.executable, pptx_script],
                                capture_output=True, text=True,
                                cwd=os.path.dirname(pptx_script), timeout=600)
            if r2.returncode != 0:
                raise RuntimeError(
                    "PPTX generation failed (exit %d):\n\n%s" %
                    (r2.returncode, (r2.stdout + "\n" + r2.stderr)[-2500:]))

        self._run_in_background(
            [("Load data into Database", db_load_step),
             ("Refresh COPQ PPTX Report", pptx_gen_step)],
            f"COPQ PPTX report refreshed -> {out}")

    def load_data(self):
        # Support multiple files: paths separated by " | "
        raw = self.file_path.get().strip()
        paths = [p.strip() for p in raw.split("|") if p.strip()]
        missing = [p for p in paths if not os.path.exists(p)]
        if not paths or missing:
            messagebox.showwarning(
                "File Not Found",
                "Workbook(s) not found:\n" + "\n".join(missing or [raw]) +
                "\n\nRun the COPQ combine process first.")
            return
        def worker():
            frames = []
            used = []
            for path in paths:
                try:
                    try:
                        df = pd.read_excel(path, sheet_name="COPQ_Clean")
                    except Exception:
                        # Raw ERP export: clean it and map to standard columns,
                        # same logic as the combine pipeline.
                        site = combine_files.extract_site_name(os.path.basename(path))
                        ext = os.path.splitext(path)[1].lower()
                        if ext in (".xlsx", ".xls"):
                            raw = pd.read_excel(path)
                        elif ext == ".csv":
                            raw = pd.read_csv(path)
                        else:
                            print(f"[!] Skipping unsupported file type {os.path.basename(path)}")
                            continue
                        cleaned = combine_files.clean_copq_dataframe(raw)
                        df = combine_files.build_copq_clean_master({site: cleaned})
                        if df.empty:
                            print(f"[!] Skipping file with no mappable data: {os.path.basename(path)}")
                            continue
                except Exception as exc:
                    print(f"[!] Skipping unreadable file {os.path.basename(path)}: {exc}")
                    continue
                df["_source_file"] = os.path.basename(path)
                frames.append(df)
                used.append(os.path.basename(path))
            if not frames:
                self.after(0, lambda: messagebox.showerror("Load Error", "No readable Excel files."))
                return
            df = pd.concat(frames, ignore_index=True, sort=False)
            if "Date" in df.columns:
                df["_date"] = pd.to_datetime(df["Date"], errors="coerce")
                df["_month"] = df["_date"].dt.strftime("%Y-%m")
            else:
                df["_month"] = ""
            for col in ("Ttl cost ($)", "Defective Qty(Pair)"):
                if col in df.columns:
                    df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0.0)
                else:
                    df[col] = 0.0
            if "_type" not in df.columns:
                df["_type"] = df.apply(classify, axis=1)
            # Exclude Rework completely
            df = df[df["_type"] != "Rework"].reset_index(drop=True)
            if "Type" in df.columns:
                df = df[~df["Type"].astype(str).str.strip().str.lower().isin(["rework"])].reset_index(drop=True)
            site_col = next((c for c in df.columns if str(c).strip() == "Site"), None)
            df["_site"] = (df[site_col].astype(str).str.strip().str.upper()
                           .replace({"JV3": "JVB"}) if site_col else "")
            self.df = df
            months = ["All"] + sorted(m for m in df["_month"].dropna().unique() if m)

            def update_month_filter(months=months):
                self.month_cb.configure(values=months)
                if self.month_var.get() not in months:
                    self.month_var.set(months[-1])

            self.after(0, update_month_filter)

            n = len(df)
            label = f"{len(used)} files" if len(used) > 1 else used[0]
            self.after(0, lambda: self.status.set(f"Loaded {n:,} records from {label}"))
            self.after(0, self.apply_filters)

        threading.Thread(target=worker, daemon=True).start()

    def toggle_sites(self):
        vals = [v.get() for v in self.site_vars.values()]
        new_val = not all(vals)
        for v in self.site_vars.values():
            v.set(new_val)
        self.apply_filters()

    def apply_filters(self):
        if self.df is None:
            return
        d = self.df
        sel_types = [t for t in TYPES if self.type_vars[t].get()]
        sel_sites = [s for s in SITES if self.site_vars[s].get()]

        mask = d["_type"].isin(sel_types) & d["_site"].isin(sel_sites)
        month = self.month_var.get()
        if month and month != "All":
            mask &= d["_month"] == month
        f = d[mask]
        self.filtered = f
        self._update_kpis(f)
        self._update_table(f)
        self._update_charts(f)

    # ---------- KPIs ----------
    def _update_kpis(self, f):
        metric = self.metric_var.get()
        use_qty = metric.startswith("Qty")
        show_all = metric == "All"

        def fmt(v):
            return f"${v:,.2f} | {v:,.0f} pr" if show_all else (
                f"{v:,.0f}" if use_qty else f"${v:,.2f}")

        def kpi_val(mask):
            if show_all:
                cost = f.loc[mask, "Ttl cost ($)"].sum()
                qty = f.loc[mask, "Defective Qty(Pair)"].sum()
                return f"${cost:,.2f} | {qty:,.0f} pr"
            col = "Defective Qty(Pair)" if use_qty else "Ttl cost ($)"
            return fmt(f.loc[mask, col].sum())

        all_mask = pd.Series(True, index=f.index)
        self.kpi_total.configure(text=kpi_val(all_mask))
        self.kpi_tu.configure(text=kpi_val(f["_type"] == "Touch-up"))
        self.kpi_re.configure(text=kpi_val(f["_type"] == "Reinspection"))
        self.kpi_bc.configure(text=kpi_val(f["_type"] == "B/C"))
        self.kpi_rows.configure(text=f"{len(f):,}")

    # ---------- table ----------
    def _update_table(self, f):
        self.tree.delete(*self.tree.get_children())
        cols = [c for c in f.columns if not c.startswith("_")]
        self.tree.configure(columns=cols)
        for c in cols:
            self.tree.heading(c, text=c)
            self.tree.column(c, width=110, stretch=False, anchor="center")
        for _, row in f.head(1000).iterrows():
            self.tree.insert("", "end", values=[
                ("" if pd.isna(row[c]) else str(row[c]))[:60] for c in cols])

    # ---------- charts ----------
    def _dark_axis(self, ax):
        ax.set_facecolor(CARD)
        ax.tick_params(colors=MUT, labelsize=8)
        for s in ax.spines.values():
            s.set_color("#334155")

    def _update_charts(self, f):
        use_qty = self.metric_var.get().startswith("Qty")
        # "All" metric: charts plot Cost ($); KPIs show both $ and pairs.
        col = "Defective Qty(Pair)" if use_qty else "Ttl cost ($)"
        unit = "" if use_qty else "$"

        def compact(v):
            # compact number label, e.g. 1.08M / 421K / 850 (unit prefix added)
            a = abs(v)
            if a >= 1_000_000:
                return f"{unit}{v/1_000_000:,.2f}M"
            if a >= 1_000:
                return f"{unit}{v/1_000:,.1f}K"
            return f"{unit}{v:,.0f}"

        (ax1, ax2), (ax3, ax4) = self.axes
        # 1. Monthly trend by type
        ax1.clear(); self._dark_axis(ax1)
        if "_month" in f.columns and f["_month"].ne("").any():
            pivot = (f[f["_month"] != ""]
                     .pivot_table(index="_month", columns="_type", values=col,
                                  aggfunc="sum").fillna(0))
            bottom = np.zeros(len(pivot))
            x = range(len(pivot))
            for t in TYPES:
                if t in pivot.columns:
                    vals = pivot[t].values
                    ax1.bar(x, vals, bottom=bottom, color=TYPE_COLORS[t], label=t, width=0.6)
                    bottom += vals
            # total labels on top of each stacked bar
            for xi, total in zip(x, bottom):
                if total > 0:
                    ax1.text(xi, total, compact(total), ha="center", va="bottom",
                             color=TXT, fontsize=8, fontweight="bold")
            ax1.set_xticks(list(x))
            ax1.set_xticklabels(pivot.index, rotation=30)
            ax1.yaxis.set_major_formatter(
                matplotlib.ticker.FuncFormatter(lambda v, _: compact(v)))
            ax1.set_ylim(top=bottom.max() * 1.15 if len(bottom) else 1)
            ax1.legend(fontsize=7, facecolor=CARD, labelcolor=TXT, framealpha=0.6)
        ax1.set_title("Monthly Trend by Type", color=TXT, fontsize=10, fontweight="bold")

        # 2. Cost share donut by type
        ax2.clear(); self._dark_axis(ax2)
        sums = [f.loc[f["_type"] == t, col].sum() for t in TYPES]
        if sum(sums) > 0:
            ax2.pie(sums, labels=TYPES, colors=[TYPE_COLORS[t] for t in TYPES],
                    autopct="%1.1f%%",
                    pctdistance=0.76,
                    textprops={"color": TXT, "fontsize": 8},
                    wedgeprops={"width": 0.45})
        ax2.set_title("Type Share", color=TXT, fontsize=10, fontweight="bold")

        # 3. Site ranking
        ax3.clear(); self._dark_axis(ax3)
        site_tot = (f.groupby("_site")[col].sum().sort_values())
        if len(site_tot):
            colors = [ACC if s.startswith("VH") else "#34D399" for s in site_tot.index]
            bars = ax3.barh(site_tot.index, site_tot.values, color=colors, height=0.6)
            for b, v in zip(bars, site_tot.values):
                ax3.text(b.get_width(), b.get_y() + b.get_height()/2,
                         compact(v), va="center", ha="left",
                         color=TXT, fontsize=8, fontweight="bold")
            ax3.set_xlim(right=site_tot.values.max() * 1.22)
            ax3.xaxis.set_major_formatter(
                matplotlib.ticker.FuncFormatter(lambda v, _: compact(v)))
        ax3.set_title("Site Ranking", color=TXT, fontsize=10, fontweight="bold")

        # 4. Top models
        ax4.clear(); self._dark_axis(ax4)
        model_col = next((c for c in ("Model", "model") if c in f.columns), None)
        if model_col:
            top = (f.groupby(model_col)[col].sum()
                   .sort_values(ascending=False).head(8).iloc[::-1])
            if len(top):
                bars = ax4.barh([str(i)[:18] for i in top.index],
                                top.values, color="#C084FC", height=0.6)
                for b, v in zip(bars, top.values):
                    ax4.text(b.get_width(), b.get_y() + b.get_height()/2,
                             compact(v), va="center", ha="left",
                             color=TXT, fontsize=8)
                ax4.set_xlim(right=top.values.max() * 1.25)
                ax4.xaxis.set_major_formatter(
                    matplotlib.ticker.FuncFormatter(lambda v, _: compact(v)))
        ax4.set_title("Top Models", color=TXT, fontsize=10, fontweight="bold")

        self.fig.tight_layout()
        self.canvas.draw()

    # ---------- FTT B/C report ----------
    def load_ftt_report(self):
        selected = self._selected_ftt_files()
        if len(selected) not in (1, 4) or any(not os.path.exists(p) for p in selected):
            self._clear_ftt_report_view()
            return
        path = (DEFAULT_FTT_REPORT if os.path.exists(DEFAULT_FTT_REPORT)
                else DEFAULT_FTT_REPORT_ALT if os.path.exists(DEFAULT_FTT_REPORT_ALT)
                else None)

        def worker():
            try:
                if not path:
                    self.after(0, lambda: self.ftt_status.configure(
                        text="No FTT report found — run '2. Process FTT' first."))
                    return
                try:
                    summary = pd.read_excel(path, sheet_name="FTT_BC_Summary")
                except Exception:
                    # Output of generate_ftt_report (process_ftt.bat):
                    # sheets Combined_All / Site_* with FTT_Qty column.
                    try:
                        data = pd.read_excel(path, sheet_name="Combined_All")
                        defect_col = next(
                            (c for c in ("Defect_Issue", "Defect") if c in data.columns), None)
                        qty_col = next(
                            (c for c in ("FTT_Qty", "B/C Quantity") if c in data.columns), None)
                        if defect_col is None or qty_col is None:
                            raise ValueError("Combined_All missing expected columns")
                        summary = (
                            data.groupby(["Site", "Model", defect_col], dropna=False)[qty_col]
                            .sum().reset_index()
                            .rename(columns={defect_col: "Defect", qty_col: "B/C Quantity"})
                            .sort_values("B/C Quantity", ascending=False))
                        summary["B/C Cost"] = 0.0
                        summary.attrs["has_cost"] = False
                    except Exception:
                        data = pd.read_excel(path, sheet_name="FTT_BC_Data")
                        summary = (data.groupby(["Site", "Model", "Defect"], dropna=False)
                                   [["B/C Quantity", "B/C Cost"]].sum().reset_index()
                                   .sort_values("B/C Cost", ascending=False))
                self.after(0, lambda: self._render_ftt(summary,
                            os.path.basename(path)))
            except Exception as exc:
                print(f"[!] FTT load failed: {exc}")
                msg = f"Failed to load FTT report: {exc}"
                self.after(0, lambda m=msg: self.ftt_status.configure(text=m, fg="#FB7185"))

        threading.Thread(target=worker, daemon=True).start()

    def _render_ftt(self, summary, fname):
        for c in ("B/C Quantity", "B/C Cost"):
            if c in summary.columns:
                summary[c] = pd.to_numeric(summary[c], errors="coerce").fillna(0.0)
        total_qty = summary["B/C Quantity"].sum()
        total_cost = summary["B/C Cost"].sum()
        n_def = summary["Defect"].nunique() if "Defect" in summary.columns else 0
        # When the report has no cost data (qty-only FTT export), plot
        # quantity instead of a wall of $0 bars.
        cost_col = "B/C Cost"
        if summary.attrs.get("has_cost", True) or total_cost > 0:
            val_label, val_fmt = "B/C Cost", "${:,.0f}"
            tick_fmt = lambda v, _: f"${v/1_000:,.0f}K" if v >= 1_000 else f"${v:,.0f}"
            chart_title_cost = "B/C Cost"
        else:
            cost_col = "B/C Quantity"
            val_label, val_fmt = "B/C Qty", "{:,.0f}"
            tick_fmt = lambda v, _: f"{v/1_000:,.0f}K" if v >= 1_000 else f"{v:,.0f}"
            chart_title_cost = "B/C Qty"
        self.ftt_status.configure(
            text=f"{fname}: {len(summary):,} rows | Total B/C Qty: {total_qty:,.0f} pairs | "
                 f"Total B/C Cost: ${total_cost:,.2f} | {n_def} defect types",
            fg=ACC)

        # table
        cols = [c for c in summary.columns]
        self.ftt_tree.delete(*self.ftt_tree.get_children())
        self.ftt_tree.configure(columns=cols)
        for c in cols:
            self.ftt_tree.heading(c, text=c)
            self.ftt_tree.column(c, width=140, stretch=False, anchor="center")
        for _, row in summary.head(2000).iterrows():
            self.ftt_tree.insert("", "end", values=[
                (f"{row[c]:,.0f}" if c == "B/C Quantity" else
                 f"${row[c]:,.2f}" if c == "B/C Cost" else
                 ("" if pd.isna(row[c]) else str(row[c])))[:60] for c in cols])
        # charts
        ax_d, ax_s = self.ftt_axes
        colors = {}
        try:
            import process_ftt as _pf
            colors = _pf.load_defect_color_map(DEFAULT_COLOR_MAP)
        except Exception:
            pass
        ax_d.clear(); self._dark_axis(ax_d)
        top = (summary.groupby("Defect")[cost_col].sum()
               .sort_values(ascending=False).head(12).iloc[::-1]) if "Defect" in summary.columns else pd.Series(dtype=float)
        if len(top):
            bar_colors = ["#" + colors.get(str(d).lower(), "5B9BD5") for d in top.index]
            bars = ax_d.barh([str(i)[:26] for i in top.index], top.values,
                             color=bar_colors, height=0.6)
            mx = top.values.max()
            for b, v in zip(bars, top.values):
                ax_d.text(b.get_width(), b.get_y() + b.get_height()/2,
                          val_fmt.format(v), va="center", ha="left",
                          color=TXT, fontsize=7, fontweight="bold")
            ax_d.set_xlim(right=mx * 1.25)
            ax_d.xaxis.set_major_formatter(matplotlib.ticker.FuncFormatter(tick_fmt))
        ax_d.set_title(f"Top 12 Defects by {val_label}", color=TXT,
                       fontsize=10, fontweight="bold")

        ax_s.clear(); self._dark_axis(ax_s)
        site_tot = (summary.groupby("Site")[cost_col].sum().sort_values()) \
            if "Site" in summary.columns else pd.Series(dtype=float)
        if len(site_tot):
            scolors = [ACC if s.startswith("VH") else "#34D399" for s in site_tot.index]
            bars = ax_s.barh(site_tot.index, site_tot.values, color=scolors, height=0.55)
            mx = site_tot.values.max()
            for b, v in zip(bars, site_tot.values):
                ax_s.text(b.get_width(), b.get_y() + b.get_height()/2,
                          val_fmt.format(v), va="center", ha="left", color=TXT, fontsize=8)
            ax_s.set_xlim(right=mx * 1.25)
            ax_s.xaxis.set_major_formatter(matplotlib.ticker.FuncFormatter(tick_fmt))
        ax_s.set_title(f"{chart_title_cost} by Site", color=TXT, fontsize=10, fontweight="bold")

        self.ftt_fig.tight_layout()
        self.ftt_canvas.draw()


if __name__ == "__main__":
    Dashboard().mainloop()
