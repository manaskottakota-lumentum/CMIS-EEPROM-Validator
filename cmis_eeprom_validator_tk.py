from __future__ import annotations

import csv
import os
import sys
import tempfile
import time
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter.scrolledtext import ScrolledText
from tkinter import filedialog, messagebox, ttk

from cmis_eeprom_validator import CmisDump, load_expected_parameters, validate_parameters
from workbook_generator import generate_workbook_from_dump


def app_root() -> Path:
    return Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))


ROOT = app_root()


def result_to_dict(result) -> dict[str, object]:
    parameter = result.parameter
    return {
        "sourceSheet": parameter.source_sheet,
        "check": parameter.check,
        "parameter": parameter.parameter,
        "page": result.page,
        "address": result.address,
        "bits": parameter.bits,
        "length": parameter.length,
        "dataType": parameter.data_type,
        "cmisType": parameter.cmis_type,
        "expected": result.expected,
        "actual": result.actual,
        "status": result.status,
        "message": result.message,
        "activeCheck": parameter.active_check,
        "sourceRow": parameter.source_row,
    }


def summarize(results) -> dict[str, int]:
    return {
        "total": len(results),
        "passed": sum(1 for result in results if result.status == "PASS"),
        "failed": sum(1 for result in results if result.status == "FAIL"),
        "read": sum(1 for result in results if result.status == "READ"),
        "notChecked": sum(1 for result in results if result.status == "NOT CHECKED"),
        "errors": sum(1 for result in results if result.status == "ERROR"),
    }


def validate_files(dump_path: str, workbook_path: str) -> dict[str, object]:
    started = datetime.now()
    start_time = time.perf_counter()
    dump = CmisDump.from_file(dump_path)
    parameters = load_expected_parameters(workbook_path)
    results = validate_parameters(dump, parameters)
    return {
        "results": [result_to_dict(result) for result in results],
        "summary": summarize(results),
        "validationTime": started.strftime("%m/%d/%Y %I:%M:%S %p").replace(" 0", " "),
        "durationSeconds": round(time.perf_counter() - start_time, 3),
    }


class ToolTip:
    def __init__(self, widget: tk.Widget, text: str) -> None:
        self.widget = widget
        self.text = text
        self.tip: tk.Toplevel | None = None
        self.widget.bind("<Enter>", self.show)
        self.widget.bind("<Leave>", self.hide)

    def show(self, _event=None) -> None:
        if self.tip or not self.text:
            return
        self.tip = tk.Toplevel(self.widget)
        self.tip.wm_overrideredirect(True)
        label = ttk.Label(self.tip, text=self.text, padding=(8, 4), background="#ffffe0", relief="solid", borderwidth=1)
        label.pack()
        self.tip.update_idletasks()
        root = self.widget.winfo_toplevel()
        x = self.widget.winfo_rootx() + 16
        y = self.widget.winfo_rooty() + self.widget.winfo_height() + 8
        max_x = root.winfo_rootx() + root.winfo_width() - self.tip.winfo_reqwidth() - 6
        max_y = root.winfo_rooty() + root.winfo_height() - self.tip.winfo_reqheight() - 6
        x = max(root.winfo_rootx() + 6, min(x, max_x))
        y = max(root.winfo_rooty() + 6, min(y, max_y))
        self.tip.wm_geometry(f"+{x}+{y}")

    def hide(self, _event=None) -> None:
        if self.tip:
            self.tip.destroy()
            self.tip = None


class CmisValidatorTk(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("CVSA CMIS EEPROM Validator v1.0.0")
        self.geometry("1180x660")
        self.minsize(1050, 560)

        self.dump_path = tk.StringVar()
        self.workbook_path = tk.StringVar()
        self.search_text = tk.StringVar()
        self.search_column = tk.StringVar(value="All Columns")
        self.show_all_rows = tk.BooleanVar(value=False)
        self.results_expanded = False

        self.results: list[dict[str, object]] = []
        self.filtered_results: list[dict[str, object]] = []
        self.expand_icon = self._make_expand_icon()
        self.collapse_icon = self._make_collapse_icon()

        self._configure_styles()
        self._build_ui()
        self._wire_events()

    def _make_expand_icon(self) -> tk.PhotoImage:
        return self._make_corner_icon(expanded=True)

    def _make_collapse_icon(self) -> tk.PhotoImage:
        return self._make_corner_icon(expanded=False)

    def _make_corner_icon(self, expanded: bool) -> tk.PhotoImage:
        image = tk.PhotoImage(width=22, height=22)
        color = "#1f2937"
        segments = (
            [(3, 3, 9, 3), (3, 3, 3, 9), (13, 3, 19, 3), (19, 3, 19, 9), (3, 13, 3, 19), (3, 19, 9, 19), (19, 13, 19, 19), (13, 19, 19, 19)]
            if expanded
            else [(3, 9, 9, 9), (9, 3, 9, 9), (13, 9, 19, 9), (13, 3, 13, 9), (9, 13, 9, 19), (3, 13, 9, 13), (13, 13, 13, 19), (13, 13, 19, 13)]
        )
        for x1, y1, x2, y2 in segments:
            self._draw_line(image, x1, y1, x2, y2, color)
        return image

    def _draw_line(self, image: tk.PhotoImage, x1: int, y1: int, x2: int, y2: int, color: str) -> None:
        if x1 == x2:
            for y in range(min(y1, y2), max(y1, y2) + 1):
                image.put(color, (x1, y))
                image.put(color, (x1 + 1, y))
        else:
            for x in range(min(x1, x2), max(x1, x2) + 1):
                image.put(color, (x, y1))
                image.put(color, (x, y1 + 1))

    def _configure_styles(self) -> None:
        self.configure(bg="#f2f2f2")
        style = ttk.Style(self)
        style.theme_use("vista" if "vista" in style.theme_names() else "clam")
        style.configure(".", font=("Segoe UI", 10))
        style.configure("TFrame", background="#f2f2f2")
        style.configure("TLabelframe", background="#f2f2f2", padding=(8, 6))
        style.configure("TLabelframe.Label", background="#f2f2f2", foreground="#1f2937", font=("Segoe UI", 10, "bold"))
        style.configure("Title.TLabel", background="#f2f2f2", foreground="#0b2545", font=("Segoe UI", 18, "bold"))
        style.configure("Subtle.TLabel", background="#f2f2f2", foreground="#526176", font=("Segoe UI", 9))
        style.configure("Field.TLabel", background="#f2f2f2", foreground="#111827", font=("Segoe UI", 9, "bold"))
        style.configure("Value.TLabel", background="#f2f2f2", foreground="#111827", font=("Segoe UI", 10))
        style.configure("Muted.TLabel", background="#f2f2f2", foreground="#526176", font=("Segoe UI", 9))
        style.configure("Primary.TButton", font=("Segoe UI", 9, "bold"), padding=(10, 5))
        style.configure("TButton", font=("Segoe UI", 9), padding=(9, 5))
        style.configure("Icon.TButton", padding=(4, 4))
        style.configure("TEntry", padding=(4, 2))
        style.configure("Treeview", rowheight=26, font=("Segoe UI", 9), background="white", fieldbackground="white")
        style.configure("Treeview.Heading", font=("Segoe UI", 9, "bold"), background="#f5f5f5")
        style.map("Treeview", background=[("selected", "#dbeafe")])

    def _build_ui(self) -> None:
        self.canvas = tk.Canvas(self, bg="#f2f2f2", highlightthickness=0)
        app_scroll = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=app_scroll.set)
        app_scroll.pack(side="right", fill="y")
        self.canvas.pack(side="left", fill="both", expand=True)

        self.page = ttk.Frame(self.canvas, padding=(10, 8, 10, 10))
        self.canvas_window = self.canvas.create_window((0, 0), window=self.page, anchor="nw")

        self.file_grid = ttk.LabelFrame(self.page, text="Input files")
        self.file_grid.pack(fill="x", pady=(0, 6))
        self.file_grid.columnconfigure(0, weight=1)
        self.file_grid.columnconfigure(1, weight=1)
        self._file_card(self.file_grid, 0, "EEPROM Dump", self.dump_path, self.browse_dump)
        self._file_card(self.file_grid, 1, "Expected Values (Excel)", self.workbook_path, self.browse_workbook)

        self.action_bar = ttk.Frame(self.page)
        self.action_bar.pack(fill="x", pady=(4, 8))
        ttk.Button(self.action_bar, text="Run Validation", style="Primary.TButton", command=self.run_validation).pack(side="left", padx=(0, 6))
        ttk.Button(self.action_bar, text="Generate Spec", command=self.generate_workbook).pack(side="left", padx=(0, 6))
        ttk.Button(self.action_bar, text="Clear", command=self.clear).pack(side="left", padx=(0, 6))
        ttk.Checkbutton(self.action_bar, text="Show all rows", variable=self.show_all_rows, command=self.apply_filters).pack(side="left")
        self.msft_sample_button = ttk.Button(self.action_bar, text="Load Arista Sample", command=self.load_arista_sample)
        self.msft_sample_button.pack(side="right")
        ToolTip(self.msft_sample_button, "Load sample Arista dump and generated spec")

        self.summary_frame = ttk.LabelFrame(self.page, text="Validation summary")
        self.summary_frame.pack(fill="x", pady=(0, 8))
        self.summary_labels: dict[str, ttk.Label] = {}
        self._build_summary()

        self.body = ttk.Frame(self.page)
        self.body.pack(fill="both", expand=True)
        self.body.columnconfigure(0, weight=3)
        self.body.columnconfigure(1, weight=2)
        self.body.rowconfigure(0, weight=1)

        self.results_card = ttk.LabelFrame(self.body, text="Validation results")
        self.results_card.grid(row=0, column=0, sticky="nsew", padx=(0, 12))
        self._build_results_table()

        self.details_card = ttk.LabelFrame(self.body, text="Parameter details")
        self.details_card.grid(row=0, column=1, sticky="nsew")
        self._build_details_panel()

    def _file_card(self, parent: ttk.Frame, column: int, title: str, path_var: tk.StringVar, browse_command, paste_command=None) -> None:
        card = ttk.Frame(parent)
        card.grid(row=0, column=column, sticky="ew", padx=(0, 8) if column == 0 else (8, 0), pady=(2, 4))
        card.columnconfigure(1, weight=1)
        ttk.Label(card, text=title, style="Field.TLabel").grid(row=0, column=0, sticky="w", padx=(0, 8))
        path_entry = ttk.Entry(card, textvariable=path_var, state="readonly")
        path_entry.grid(row=0, column=1, sticky="ew", padx=(0, 6))
        ttk.Button(card, text="Browse", command=browse_command).grid(row=0, column=2, padx=(0, 4))
        next_column = 3
        if paste_command:
            ttk.Button(card, text="Paste", command=paste_command).grid(row=0, column=next_column, padx=(0, 4))
            next_column += 1
        ttk.Button(card, text="Open", command=lambda: self.open_selected_file(path_var.get())).grid(row=0, column=next_column)

    def _build_summary(self) -> None:
        fields = [("overall", "Not run"), ("total", "0"), ("passed", "0"), ("failed", "0"), ("notChecked", "0"), ("errors", "0"), ("validationTime", "-")]
        labels = ["Result", "Total", "Passed", "Failed", "Not Checked", "Errors", "Validation Time"]
        for i, ((key, value), label) in enumerate(zip(fields, labels)):
            block = ttk.Frame(self.summary_frame)
            block.pack(side="left", fill="x", expand=True, padx=(4, 16), pady=(2, 4))
            ttk.Label(block, text=label, style="Muted.TLabel").pack(anchor="w")
            font = ("Segoe UI", 14, "bold") if key == "overall" else ("Segoe UI", 12, "bold")
            color = "#071426"
            lbl = ttk.Label(block, text=value, background="#f2f2f2", foreground=color, font=font)
            lbl.pack(anchor="w", pady=(3, 0))
            self.summary_labels[key] = lbl

    def _build_results_table(self) -> None:
        toolbar = ttk.Frame(self.results_card)
        toolbar.pack(fill="x", padx=6, pady=(2, 6))

        ttk.Label(toolbar, text="Search", style="Field.TLabel").pack(side="left", padx=(0, 4))
        self.search_entry = ttk.Entry(toolbar, textvariable=self.search_text, width=38)
        self.search_entry.pack(side="left", fill="x", expand=True, padx=(0, 8))

        ttk.Label(toolbar, text="Column", style="Field.TLabel").pack(side="left", padx=(0, 4))
        self.column_combo = ttk.Combobox(toolbar, textvariable=self.search_column, values=self._column_options(), state="readonly", width=18)
        self.column_combo.pack(side="left", padx=(0, 8))
        self.expand_button = ttk.Button(toolbar, image=self.expand_icon, command=self.toggle_results_expanded, style="Icon.TButton", width=2)
        self.expand_button.pack(side="left")
        self.expand_tooltip = ToolTip(self.expand_button, "Expand results")

        table_wrap = ttk.Frame(self.results_card)
        table_wrap.pack(fill="both", expand=True, padx=6, pady=(0, 6))
        table_wrap.rowconfigure(0, weight=1)
        table_wrap.columnconfigure(0, weight=1)

        self.compact_columns = ("page", "parameter", "address", "expected", "actual", "status")
        self.expanded_columns = ("page", "check", "parameter", "address", "addressDec", "bits", "dataType", "expected", "actual", "status")
        self.columns = self.compact_columns
        self.tree = ttk.Treeview(table_wrap, columns=self.columns, show="headings", selectmode="browse")
        self.headings = {
            "page": "Page",
            "check": "Check",
            "parameter": "Parameter",
            "address": "Address",
            "addressDec": "Address (Dec)",
            "bits": "Bits",
            "dataType": "Data Type",
            "expected": "Expected",
            "actual": "Module Value",
            "status": "Status",
        }
        self.compact_widths = {"page": 80, "parameter": 300, "address": 125, "expected": 95, "actual": 110, "status": 105}
        self.expanded_widths = {
            "page": 95,
            "check": 70,
            "parameter": 360,
            "address": 100,
            "addressDec": 105,
            "bits": 80,
            "dataType": 130,
            "expected": 115,
            "actual": 125,
            "status": 120,
        }
        for column in self.columns:
            self.tree.heading(column, text=self.headings[column], command=lambda selected=column: self.select_search_column(selected))
            self.tree.column(column, width=self.compact_widths[column], minwidth=65, stretch=column == "parameter")
        y_scroll = ttk.Scrollbar(table_wrap, orient="vertical", command=self.tree.yview)
        x_scroll = ttk.Scrollbar(table_wrap, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=y_scroll.set, xscrollcommand=x_scroll.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        y_scroll.grid(row=0, column=1, sticky="ns")
        x_scroll.grid(row=1, column=0, sticky="ew")

    def _build_details_panel(self) -> None:
        self.details_card.columnconfigure(0, weight=0)
        self.details_card.columnconfigure(1, weight=1)

        self.detail_title = ttk.Label(
            self.details_card,
            text="-",
            style="Value.TLabel",
            wraplength=300,
            justify="left",
            font=("Segoe UI", 10, "bold"),
        )
        self.detail_title.grid(row=0, column=0, columnspan=2, sticky="ew", padx=6, pady=(4, 8))

        self.detail_rows: dict[str, ttk.Label] = {}
        fields = [
            ("sourceSheet", "Source Sheet"),
            ("check", "Check"),
            ("page", "Page"),
            ("address", "Start Address"),
            ("bits", "Bits"),
            ("length", "Length"),
            ("dataType", "Data Type"),
            ("cmisType", "CMIS Type"),
            ("expected", "Expected Value"),
            ("actual", "Module Value"),
        ]
        for row_index, (key, label) in enumerate(fields, start=1):
            ttk.Label(self.details_card, text=label, style="Field.TLabel", width=14).grid(
                row=row_index,
                column=0,
                sticky="nw",
                padx=(6, 10),
                pady=1,
            )
            value = ttk.Label(self.details_card, text="-", style="Value.TLabel", wraplength=220, justify="left")
            value.grid(row=row_index, column=1, sticky="ew", padx=(0, 6), pady=1)
            self.detail_rows[key] = value

        self.details_card.bind("<Configure>", self._resize_detail_wraps, add="+")

    def _resize_detail_wraps(self, _event=None) -> None:
        card_width = self.details_card.winfo_width()
        value_width = max(120, card_width - 150)
        self.detail_title.configure(wraplength=max(120, card_width - 20))
        for label in self.detail_rows.values():
            label.configure(wraplength=value_width)

    def _wire_events(self) -> None:
        self.canvas.bind("<Configure>", self._resize_canvas_window)
        self.page.bind("<Configure>", self._sync_scroll_region)
        self.tree.bind("<<TreeviewSelect>>", self.show_selected_detail)
        self.search_text.trace_add("write", lambda *_: self.apply_filters())
        self.search_column.trace_add("write", lambda *_: self.apply_filters())

    def _resize_canvas_window(self, event) -> None:
        self.canvas.itemconfigure(self.canvas_window, width=event.width)
        if self.results_expanded:
            self.canvas.itemconfigure(self.canvas_window, height=event.height)

    def _sync_scroll_region(self, _event=None) -> None:
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        if self.results_expanded:
            self.canvas.itemconfigure(self.canvas_window, height=self.canvas.winfo_height())

    def toggle_results_expanded(self) -> None:
        self.results_expanded = not self.results_expanded
        if self.results_expanded:
            self.file_grid.pack_forget()
            self.action_bar.pack_forget()
            self.summary_frame.pack_forget()
            self.details_card.grid_remove()
            self.results_card.grid_configure(column=0, columnspan=2, padx=0, sticky="nsew")
            self.body.columnconfigure(0, weight=1)
            self.body.columnconfigure(1, weight=0)
            self.expand_button.configure(image=self.collapse_icon)
            self.expand_tooltip.text = "Collapse results"
            self.configure_table_columns()
            self.populate_tree()
            self.update_idletasks()
            self.canvas.itemconfigure(self.canvas_window, height=self.canvas.winfo_height())
            return

        self.body.pack_forget()
        self.file_grid.pack_forget()
        self.action_bar.pack_forget()
        self.summary_frame.pack_forget()
        self.file_grid.pack(fill="x", pady=(0, 6))
        self.action_bar.pack(fill="x", pady=(4, 8))
        self.summary_frame.pack(fill="x", pady=(0, 8))
        self.body.pack(fill="both", expand=True)
        self.results_card.grid_configure(column=0, columnspan=1, padx=(0, 12), sticky="nsew")
        self.details_card.grid()
        self.body.columnconfigure(0, weight=3)
        self.body.columnconfigure(1, weight=2)
        self.expand_button.configure(image=self.expand_icon)
        self.expand_tooltip.text = "Expand results"
        self.configure_table_columns()
        self.populate_tree()
        self.update_idletasks()
        self.canvas.itemconfigure(self.canvas_window, height=self.page.winfo_reqheight())
        self._sync_scroll_region()

    def configure_table_columns(self) -> None:
        self.columns = self.expanded_columns if self.results_expanded else self.compact_columns
        self.tree.configure(columns=self.columns)
        widths = self.expanded_widths if self.results_expanded else self.compact_widths
        for column in self.columns:
            self.tree.heading(column, text=self.headings[column], command=lambda selected=column: self.select_search_column(selected))
            self.tree.column(column, width=widths[column], minwidth=65, stretch=column == "parameter")

    def select_search_column(self, column: str) -> None:
        label = self.headings.get(column)
        if not label:
            return
        self.search_column.set(label)
        self.search_entry.focus_set()

    def browse_dump(self) -> None:
        path = filedialog.askopenfilename(
            title="Select EEPROM / CMIS dump",
            filetypes=[("Dump files", "*.txt *.hex *.dump *.log"), ("All files", "*.*")],
        )
        if path:
            self.dump_path.set(path)

    def paste_dump(self) -> None:
        dialog = tk.Toplevel(self)
        dialog.title("Paste EEPROM Dump")
        dialog.transient(self)
        dialog.grab_set()
        dialog.geometry("760x520")
        dialog.columnconfigure(0, weight=1)
        dialog.rowconfigure(1, weight=1)

        ttk.Label(
            dialog,
            text="Paste EEPROM / CMIS dump text below.",
            style="Field.TLabel",
        ).grid(row=0, column=0, sticky="w", padx=14, pady=(12, 6))

        text_box = ScrolledText(dialog, wrap="none", font=("Consolas", 10), undo=True)
        text_box.grid(row=1, column=0, sticky="nsew", padx=14, pady=(0, 10))
        text_box.focus_set()

        buttons = ttk.Frame(dialog)
        buttons.grid(row=2, column=0, sticky="e", padx=14, pady=(0, 12))

        def use_pasted_text() -> None:
            content = text_box.get("1.0", "end").strip()
            if not content:
                messagebox.showerror("Paste EEPROM Dump", "Paste EEPROM dump text before loading.", parent=dialog)
                return
            temp_dir = Path(tempfile.gettempdir()) / "cmis_eeprom_validator"
            temp_dir.mkdir(parents=True, exist_ok=True)
            temp_path = temp_dir / "pasted_eeprom_dump.txt"
            temp_path.write_text(content + "\n", encoding="utf-8")
            self.dump_path.set(str(temp_path))
            dialog.destroy()

        ttk.Button(buttons, text="Load Pasted Dump", style="Primary.TButton", command=use_pasted_text).pack(side="left", padx=(0, 6))
        ttk.Button(buttons, text="Cancel", command=dialog.destroy).pack(side="left")

    def browse_workbook(self) -> None:
        path = filedialog.askopenfilename(
            title="Select expected-values workbook",
            filetypes=[("Spreadsheet files", "*.xlsx *.csv *.tsv"), ("All files", "*.*")],
        )
        if path:
            self.workbook_path.set(path)

    def load_arista_sample(self) -> None:
        user_profile = Path(os.environ.get("USERPROFILE", ""))
        documents_root = user_profile / "OneDrive - Lumentum Operations LLC" / "Documents"
        output_root = documents_root / "GUI" / "outputs" / "cmis_lower_memory_overview"
        dump_candidates = [
            ROOT / "samples" / "Arista IDPRM example.txt",
            output_root / "Arista IDPRM example.txt",
        ]
        workbook_candidates = [
            ROOT / "samples" / "Arista IDPRM example_specifications_spreadsheet_page03.xlsx",
            ROOT / "samples" / "Arista IDPRM example_specifications_spreadsheet.xlsx",
            output_root / "Arista IDPRM example_specifications_spreadsheet_page03.xlsx",
            output_root / "Arista IDPRM example_specifications_spreadsheet.xlsx",
            output_root / "Arista IDPRM example_specifications_excel.xlsx",
        ]
        dump = next((path for path in dump_candidates if path.exists()), None)
        workbook = next((path for path in workbook_candidates if path.exists()), None)
        if not dump or not workbook:
            messagebox.showerror("Sample Files", "Could not find the Arista sample dump and generated spec files.")
            return
        self.dump_path.set(str(dump))
        self.workbook_path.set(str(workbook))

    def generate_workbook(self) -> None:
        dump_path = filedialog.askopenfilename(
            title="Select Arista IDPROM / EEPROM dump",
            filetypes=[("Dump files", "*.txt *.hex *.dump *.log"), ("Text files", "*.txt"), ("All files", "*.*")],
        )
        if not dump_path:
            return

        default_name = f"{Path(dump_path).stem}_specifications_excel.xlsx"
        output_path = filedialog.asksaveasfilename(
            title="Save generated specifications spreadsheet",
            initialfile=default_name,
            defaultextension=".xlsx",
            filetypes=[("Specifications spreadsheet", "*.xlsx")],
        )
        if not output_path:
            return

        try:
            row_count = generate_workbook_from_dump(dump_path, output_path)
        except PermissionError:
            output_path = self.next_available_workbook_path(output_path)
            try:
                row_count = generate_workbook_from_dump(dump_path, output_path)
            except Exception as exc:
                messagebox.showerror("Generate Specifications Spreadsheet", str(exc))
                return
        except Exception as exc:
            messagebox.showerror("Generate Specifications Spreadsheet", str(exc))
            return

        self.dump_path.set(dump_path)
        self.show_generated_workbook_dialog(output_path, row_count)

    def next_available_workbook_path(self, output_path: str) -> str:
        path = Path(output_path)
        for index in range(1, 100):
            candidate = path.with_name(f"{path.stem}_{index}{path.suffix}")
            if not candidate.exists():
                return str(candidate)
        return str(path.with_name(f"{path.stem}_{os.getpid()}{path.suffix}"))

    def show_generated_workbook_dialog(self, output_path: str, row_count: int) -> None:
        dialog = tk.Toplevel(self)
        dialog.title("Generate Specifications Spreadsheet")
        dialog.transient(self)
        dialog.grab_set()
        dialog.resizable(False, False)
        dialog.columnconfigure(0, weight=1)

        message = (
            f"Created specifications spreadsheet with {row_count} byte row(s).\n\n"
            f"{output_path}\n\n"
            "Choose what to do next."
        )
        ttk.Label(dialog, text=message, wraplength=520, justify="left").grid(
            row=0, column=0, padx=18, pady=(16, 12), sticky="ew"
        )

        buttons = ttk.Frame(dialog)
        buttons.grid(row=1, column=0, padx=18, pady=(0, 16), sticky="e")

        def open_workbook() -> None:
            try:
                os.startfile(output_path)
            except Exception as exc:
                messagebox.showerror("Open Workbook", str(exc), parent=dialog)

        def load_expected_values() -> None:
            self.workbook_path.set(output_path)
            dialog.destroy()

        ttk.Button(buttons, text="Open Specifications Spreadsheet", command=open_workbook).pack(side="left", padx=(0, 6))
        ttk.Button(buttons, text="Load to Expected Values", command=load_expected_values).pack(side="left", padx=(0, 6))

        dialog.update_idletasks()
        x = self.winfo_rootx() + (self.winfo_width() - dialog.winfo_width()) // 2
        y = self.winfo_rooty() + (self.winfo_height() - dialog.winfo_height()) // 2
        dialog.geometry(f"+{max(x, 0)}+{max(y, 0)}")

    def run_validation(self) -> None:
        if not Path(self.dump_path.get()).exists():
            messagebox.showerror("Missing Dump", "Select a valid EEPROM / CMIS dump file.")
            return
        if not Path(self.workbook_path.get()).exists():
            messagebox.showerror("Missing Workbook", "Select a valid expected-values workbook.")
            return
        try:
            response = validate_files(self.dump_path.get(), self.workbook_path.get())
        except Exception as exc:
            messagebox.showerror("Validation Error", str(exc))
            return
        self.results = list(response.get("results", []))
        self.update_summary(response)
        self.apply_filters()

    def update_summary(self, response: dict[str, object]) -> None:
        summary = response.get("summary", {}) or {}
        failed = int(summary.get("failed", 0)) + int(summary.get("errors", 0))
        overall = "FAIL" if failed else "PASS"
        self.summary_labels["overall"].configure(text=overall, foreground="#dc2626" if failed else "#00875a")
        for key in ("total", "passed", "failed", "notChecked", "errors"):
            self.summary_labels[key].configure(text=str(summary.get(key, 0)))
        self.summary_labels["passed"].configure(foreground="#00875a")
        self.summary_labels["failed"].configure(foreground="#dc2626")
        self.summary_labels["validationTime"].configure(text=str(response.get("validationTime", "-")))

    def apply_filters(self) -> None:
        query = self.search_text.get().strip().lower()
        selected_column = self.search_column.get()
        filtered: list[dict[str, object]] = []
        for result in self.results:
            if not self.show_all_rows.get() and not bool(result.get("activeCheck")):
                continue
            if query and not self._row_matches(result, selected_column, query):
                continue
            filtered.append(result)
        self.filtered_results = filtered
        self.populate_tree()

    def _row_matches(self, row: dict[str, object], selected_column: str, query: str) -> bool:
        if selected_column == "All Columns":
            keys = ["sourceSheet", "check", "parameter", "page", "address", "bits", "dataType", "expected", "actual", "status"]
        else:
            keys = [self._column_key(selected_column)]
        return any(query in str(row.get(key, "")).lower() for key in keys)

    def populate_tree(self) -> None:
        selected = self.tree.selection()
        selected_index = self.tree.index(selected[0]) if selected else None
        self.tree.delete(*self.tree.get_children())
        for index, row in enumerate(self.filtered_results):
            tag = str(row.get("status", "")).lower().replace(" ", "_")
            values = (
                *self.row_values(row),
            )
            self.tree.insert("", "end", iid=str(index), values=values, tags=(tag,))
        self.tree.tag_configure("pass", foreground="#00875a")
        self.tree.tag_configure("fail", foreground="#dc2626")
        self.tree.tag_configure("error", foreground="#dc2626")
        if self.filtered_results:
            next_index = min(selected_index or 0, len(self.filtered_results) - 1)
            self.tree.selection_set(str(next_index))
            self.tree.focus(str(next_index))
            self.show_detail(self.filtered_results[next_index])
        else:
            self.clear_detail()

    def row_values(self, row: dict[str, object]) -> tuple[object, ...]:
        page = self._short_page(row.get("sourceSheet") or row.get("page"))
        address = row.get("address", "")
        expected = self._expected_display(row)
        if not self.results_expanded:
            return (
                page,
                row.get("parameter", ""),
                address,
                expected,
                row.get("actual", ""),
                row.get("status", ""),
            )

        return (
            page,
            row.get("check", ""),
            row.get("parameter", ""),
            address,
            self._address_decimal(address),
            row.get("bits", ""),
            row.get("dataType", ""),
            expected,
            row.get("actual", ""),
            row.get("status", ""),
        )

    def show_selected_detail(self, _event=None) -> None:
        selected = self.tree.selection()
        if not selected:
            return
        index = int(selected[0])
        if 0 <= index < len(self.filtered_results):
            self.show_detail(self.filtered_results[index])

    def show_detail(self, row: dict[str, object]) -> None:
        self.detail_title.configure(text=str(row.get("parameter") or "-"))
        for key, label in self.detail_rows.items():
            value = self._expected_display(row) if key == "expected" else row.get(key, "-")
            label.configure(text=str(value if value not in (None, "") else "-"))

    def clear_detail(self) -> None:
        self.detail_title.configure(text="-")
        for label in self.detail_rows.values():
            label.configure(text="-")

    def export_results(self) -> None:
        if not self.filtered_results:
            messagebox.showinfo("Export Results", "No validation results to export.")
            return
        path = filedialog.asksaveasfilename(
            title="Export Results",
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv")],
        )
        if not path:
            return
        fieldnames = [
            "sourceSheet",
            "check",
            "parameter",
            "page",
            "address",
            "bits",
            "dataType",
            "expected",
            "actual",
            "status",
            "message",
            "rawHex",
        ]
        with open(path, "w", newline="", encoding="utf-8") as output:
            writer = csv.DictWriter(output, fieldnames=fieldnames, extrasaction="ignore")
            writer.writeheader()
            writer.writerows(self.filtered_results)
        messagebox.showinfo("Export Results", f"Exported {len(self.filtered_results)} row(s).")

    def clear(self) -> None:
        self.dump_path.set("")
        self.workbook_path.set("")
        self.search_text.set("")
        self.results.clear()
        self.filtered_results.clear()
        self.populate_tree()
        for key, value in {"overall": "Not run", "total": "0", "passed": "0", "failed": "0", "notChecked": "0", "errors": "0", "validationTime": "-"}.items():
            self.summary_labels[key].configure(text=value, foreground="#071426")

    def open_selected_file(self, path: str) -> None:
        if not path:
            messagebox.showinfo("Open File", "No file selected.")
            return
        file_path = Path(path)
        if not file_path.exists():
            messagebox.showerror("Open File", f"File not found:\n{file_path}")
            return
        os.startfile(file_path)

    def _column_options(self) -> list[str]:
        return ["All Columns", "Page", "Check", "Parameter", "Address", "Bits", "Data Type", "Expected", "Module Value", "Status"]

    def _column_key(self, label: str) -> str:
        return {
            "Page": "sourceSheet",
            "Check": "check",
            "Parameter": "parameter",
            "Address": "address",
            "Bits": "bits",
            "Data Type": "dataType",
            "Expected": "expected",
            "Module Value": "actual",
            "Status": "status",
        }.get(label, "parameter")

    def _short_page(self, value: object) -> str:
        text = str(value or "")
        return text.replace(" Memory", "").replace("Page ", "")

    def _expected_display(self, row: dict[str, object]) -> str:
        if str(row.get("check", "")).strip().lower() == "no":
            return "*"
        return str(row.get("expected", ""))

    def _address_decimal(self, value: object) -> str:
        text = str(value or "").strip()
        if not text:
            return ""
        first = text.split("-", 1)[0].split("/", 1)[-1].strip()
        try:
            return str(int(first, 16)) if first.lower().startswith("0x") else str(int(first))
        except ValueError:
            return ""


if __name__ == "__main__":
    app = CmisValidatorTk()
    app.mainloop()
