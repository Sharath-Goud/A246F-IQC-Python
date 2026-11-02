import json

import openpyxl

import requests

from urllib.parse import urlparse

import os

import queue

import threading

import time

import tkinter as tk



from datetime import datetime

from pathlib import Path

from tkinter import ttk, filedialog, messagebox





APP_NAME = "A246F Excel Watcher & Sender"



APP_DIR = Path(__file__).resolve().parent

CONFIG_FILE = APP_DIR / "config.json"



DEFAULT_CONFIG = {

    "api_url": "https://localhost:7110/api/open/ipqc-template-data-a246f/process-report",

    "watch_directory": "",

    "interval": 5,

    "header_row": 7,

    "header_start_col": 2,

    "data_start_row": 8,

    "start_on_windows_login": False,

    "metadata_mappings": [

        {"field": "Report Name", "cell": "B1"},

        {"field": "Part Name", "cell": "B2"},

        {"field": "Part No.", "cell": "B3"},

        {"field": "Operation", "cell": "B4"},

        {"field": "Company", "cell": "B5"},

        {"field": "workshop", "cell": "E1"},

        {"field": "floor", "cell": "E2"},

        {"field": "class", "cell": "E3"},

        {"field": "work", "cell": "E4"},

        {"field": "Machina", "cell": "E5"},

        {"field": "Time", "cell": "G1"},

        {"field": "Cu", "cell": "G5"},

    ],

}





class ExcelWatcherApp:

    def __init__(self, root):

        self.root = root

        self.root.title(APP_NAME)

        self.root.geometry("1000x850")

        self.root.minsize(850, 700)



        self.root.configure(bg="#F3F6FA")



        self.config = self.load_config()



        self.running = False

        self.worker_thread = None

        self.stop_event = threading.Event()

        self.log_queue = queue.Queue()

        self.known_files = set()
        self.processed_signatures = set()
        self.last_attempt = {}



        self.api_url_var = tk.StringVar(

            value=self.config["api_url"]

        )

        self.watch_directory_var = tk.StringVar(

            value=self.config["watch_directory"]

        )

        self.interval_var = tk.StringVar(

            value=str(self.config["interval"])

        )

        self.header_row_var = tk.StringVar(

            value=str(self.config["header_row"])

        )

        self.header_start_col_var = tk.StringVar(

            value=str(self.config["header_start_col"])

        )

        self.data_start_row_var = tk.StringVar(

            value=str(self.config["data_start_row"])

        )

        self.start_on_login_var = tk.BooleanVar(

            value=self.config["start_on_windows_login"]

        )



        self.configure_styles()

        self.build_interface()

        self.load_mappings()



        self.root.after(150, self.process_log_queue)

        self.root.protocol("WM_DELETE_WINDOW", self.close_application)



        self.log("Application initialized.")

        self.log("Save your settings before starting monitoring.")



    def load_config(self):

        config = DEFAULT_CONFIG.copy()

        config["metadata_mappings"] = [

            item.copy()

            for item in DEFAULT_CONFIG["metadata_mappings"]

        ]



        if CONFIG_FILE.exists():

            try:

                with open(CONFIG_FILE, "r", encoding="utf-8") as file:

                    saved = json.load(file)



                config.update(saved)



            except (OSError, json.JSONDecodeError) as exc:

                messagebox.showwarning(

                    "Configuration",

                    f"Could not read config.json.\n\n{exc}",

                )



        return config



    def configure_styles(self):

        style = ttk.Style()



        try:

            style.theme_use("clam")

        except tk.TclError:

            pass



        style.configure(

            "Header.TFrame",

            background="#163A5F",

        )



        style.configure(

            "HeaderTitle.TLabel",

            background="#163A5F",

            foreground="white",

            font=("Segoe UI", 17, "bold"),

        )



        style.configure(

            "HeaderSubtitle.TLabel",

            background="#163A5F",

            foreground="#D6E5F5",

            font=("Segoe UI", 9),

        )



        style.configure(

            "Section.TLabelframe",

            background="white",

            borderwidth=1,

            relief="solid",

        )



        style.configure(

            "Section.TLabelframe.Label",

            background="white",

            foreground="#163A5F",

            font=("Segoe UI", 10, "bold"),

        )



        style.configure(

            "TLabel",

            font=("Segoe UI", 9),

        )



        style.configure(

            "TButton",

            font=("Segoe UI", 9),

            padding=(10, 6),

        )



        style.configure(

            "Primary.TButton",

            font=("Segoe UI", 9, "bold"),

            padding=(14, 7),

        )



        style.configure(

            "Treeview",

            font=("Segoe UI", 9),

            rowheight=27,

        )



        style.configure(

            "Treeview.Heading",

            font=("Segoe UI", 9, "bold"),

        )



    def build_interface(self):

        self.root.grid_columnconfigure(0, weight=1)

        self.root.grid_rowconfigure(1, weight=1)



        header = ttk.Frame(

            self.root,

            style="Header.TFrame",

            padding=(20, 14),

        )

        header.grid(

            row=0,

            column=0,

            sticky="ew",

        )



        ttk.Label(

            header,

            text="A246F Excel Watcher & Sender",

            style="HeaderTitle.TLabel",

        ).pack(anchor="w")



        ttk.Label(

            header,

            text="OMM Excel monitoring  |  Automatic report processing",

            style="HeaderSubtitle.TLabel",

        ).pack(anchor="w", pady=(4, 0))



        content = ttk.Frame(

            self.root,

            padding=(16, 12),

        )

        content.grid(

            row=1,

            column=0,

            sticky="nsew",

        )



        content.grid_columnconfigure(0, weight=1)

        content.grid_rowconfigure(2, weight=0)

        content.grid_rowconfigure(5, weight=1)



        self.build_connection_section(content)

        self.build_excel_section(content)

        self.build_mapping_section(content)

        self.build_action_section(content)

        self.build_log_section(content)



    def build_connection_section(self, parent):

        frame = ttk.LabelFrame(

            parent,

            text=" Connection Settings ",

            style="Section.TLabelframe",

            padding=(8, 4),

        )

        frame.grid(

            row=0,

            column=0,

            sticky="ew",

            pady=(0, 10),

        )



        frame.grid_columnconfigure(1, weight=1)



        ttk.Label(

            frame,

            text="API URL",

        ).grid(

            row=0,

            column=0,

            sticky="w",

            padx=(0, 12),

            pady=5,

        )



        ttk.Entry(

            frame,

            textvariable=self.api_url_var,

        ).grid(

            row=0,

            column=1,

            sticky="ew",

            pady=5,

        )



        ttk.Label(

            frame,

            text="Watch Directory",

        ).grid(

            row=1,

            column=0,

            sticky="w",

            padx=(0, 12),

            pady=5,

        )



        ttk.Entry(

            frame,

            textvariable=self.watch_directory_var,

        ).grid(

            row=1,

            column=1,

            sticky="ew",

            pady=5,

        )



        ttk.Button(

            frame,

            text="Browse...",

            command=self.browse_directory,

        ).grid(

            row=1,

            column=2,

            padx=(8, 0),

            pady=5,

        )



        ttk.Label(

            frame,

            text="Interval (seconds)",

        ).grid(

            row=2,

            column=0,

            sticky="w",

            padx=(0, 12),

            pady=5,

        )



        ttk.Entry(

            frame,

            textvariable=self.interval_var,

            width=10,

        ).grid(

            row=2,

            column=1,

            sticky="w",

            pady=5,

        )



    def build_excel_section(self, parent):

        frame = ttk.LabelFrame(

            parent,

            text=" Excel Sheet Structure ",

            style="Section.TLabelframe",

            padding=2,

        )

        frame.grid(

            row=1,

            column=0,

            sticky="ew",

            pady=(0, 3),

        )



        fields = [

            ("Header row", self.header_row_var),

            ("Header start column", self.header_start_col_var),

            ("Data start row", self.data_start_row_var),

        ]



        for column, (label, variable) in enumerate(fields):

            field_frame = ttk.Frame(frame)

            field_frame.grid(

                row=0,

                column=column,

                sticky="w",

                padx=(0, 24),

            )



            ttk.Label(

                field_frame,

                text=label,

            ).pack(anchor="w")



            ttk.Entry(

                field_frame,

                textvariable=variable,

                width=10,

            ).pack(anchor="w", pady=(2, 0))



    def build_mapping_section(self, parent):

        frame = ttk.LabelFrame(

            parent,

            text=" Metadata Cell Mapping ",

            style="Section.TLabelframe",

            padding=12,

        )

        frame.grid(

            row=2,

            column=0,

            sticky="new",

            pady=(0, 7),

        )



        frame.grid_columnconfigure(0, weight=1)

        frame.grid_rowconfigure(0, weight=1)



        table_frame = ttk.Frame(frame)

        table_frame.grid(

            row=0,

            column=0,

            sticky="nsew",

        )



        table_frame.grid_columnconfigure(0, weight=1)

        table_frame.grid_rowconfigure(0, weight=1)



        self.mapping_tree = ttk.Treeview(

            table_frame,

            columns=("field", "cell"),

            show="headings",

            selectmode="browse",

            height=7,

        )



        self.mapping_tree.heading(

            "field",

            text="Field",

        )

        self.mapping_tree.heading(

            "cell",

            text="Excel Cell",

        )



        self.mapping_tree.column(

            "field",

            width=300,

            anchor="w",

        )

        self.mapping_tree.column(

            "cell",

            width=150,

            anchor="center",

        )



        scrollbar = ttk.Scrollbar(

            table_frame,

            orient="vertical",

            command=self.mapping_tree.yview,

        )



        self.mapping_tree.configure(

            yscrollcommand=scrollbar.set,

        )



        self.mapping_tree.grid(

            row=0,

            column=0,

            sticky="nsew",

        )



        scrollbar.grid(

            row=0,

            column=1,

            sticky="ns",

        )



        buttons = ttk.Frame(frame)

        buttons.grid(

            row=0,

            column=1,

            sticky="ns",

            padx=(12, 0),

        )



        ttk.Button(

            buttons,

            text="Add Mapping",

            command=self.add_mapping,

        ).pack(fill="x", pady=(0, 8))



        ttk.Button(

            buttons,

            text="Edit Mapping",

            command=self.edit_mapping,

        ).pack(fill="x", pady=(0, 8))



        ttk.Button(

            buttons,

            text="Delete Mapping",

            command=self.delete_mapping,

        ).pack(fill="x")



    def build_action_section(self, parent):

        frame = ttk.Frame(parent)

        frame.grid(

            row=3,

            column=0,

            sticky="ew",

            pady=(0, 10),

        )



        ttk.Checkbutton(

            frame,

            text="Start on Windows login",

            variable=self.start_on_login_var,

        ).pack(

            side="left",

            padx=(0, 20),

        )



        ttk.Button(

            frame,

            text="Save",

            command=self.save_settings,

            style="Primary.TButton",

        ).pack(

            side="left",

            padx=(0, 8),

        )



        self.start_button = ttk.Button(

            frame,

            text="Start",

            command=self.start_monitoring,

            style="Primary.TButton",

        )

        self.start_button.pack(

            side="left",

            padx=(0, 8),

        )



        self.stop_button = ttk.Button(

            frame,

            text="Stop",

            command=self.stop_monitoring,

            state="disabled",

        )

        self.stop_button.pack(side="left")



        self.status_label = ttk.Label(

            frame,

            text="● Ready",

            foreground="#64748B",

            font=("Segoe UI", 10, "bold"),

        )

        self.status_label.pack(

            side="right",

            padx=8,

        )



    def build_log_section(self, parent):

        frame = ttk.LabelFrame(

            parent,

            text=" Runtime Activity Log ",

            style="Section.TLabelframe",

            padding=10,

        )

        frame.grid(

            row=5,

            column=0,

            sticky="nsew",

        )



        frame.grid_columnconfigure(0, weight=1)

        frame.grid_rowconfigure(0, weight=1)



        self.log_text = tk.Text(

            frame,

            wrap="word",

            font=("Consolas", 9),

            bg="#101827",

            fg="#D9E5F2",

            insertbackground="white",

            relief="flat",

            padx=10,

            pady=8,

            state="disabled",

        )



        scrollbar = ttk.Scrollbar(

            frame,

            orient="vertical",

            command=self.log_text.yview,

        )



        self.log_text.configure(

            yscrollcommand=scrollbar.set,

        )



        self.log_text.grid(

            row=0,

            column=0,

            sticky="nsew",

        )



        scrollbar.grid(

            row=0,

            column=1,

            sticky="ns",

        )



    def load_mappings(self):

        for item in self.mapping_tree.get_children():

            self.mapping_tree.delete(item)



        for mapping in self.config.get("metadata_mappings", []):

            self.mapping_tree.insert(

                "",

                "end",

                values=(

                    mapping.get("field", ""),

                    mapping.get("cell", ""),

                ),

            )



    def browse_directory(self):

        directory = filedialog.askdirectory(

            title="Select OMM Excel Watch Directory",

            initialdir=self.watch_directory_var.get() or str(APP_DIR),

        )



        if directory:

            self.watch_directory_var.set(directory)

            self.log(f"Watch directory selected: {directory}")



    def mapping_dialog(self, title, field="", cell=""):

        dialog = tk.Toplevel(self.root)

        dialog.title(title)

        dialog.transient(self.root)

        dialog.grab_set()

        dialog.resizable(False, False)



        dialog.geometry("380x180")



        field_var = tk.StringVar(value=field)

        cell_var = tk.StringVar(value=cell)



        container = ttk.Frame(dialog, padding=18)

        container.pack(fill="both", expand=True)



        ttk.Label(

            container,

            text="Field Name",

        ).grid(

            row=0,

            column=0,

            sticky="w",

            pady=6,

        )



        field_entry = ttk.Entry(

            container,

            textvariable=field_var,

            width=32,

        )

        field_entry.grid(

            row=0,

            column=1,

            pady=6,

        )



        ttk.Label(

            container,

            text="Excel Cell",

        ).grid(

            row=1,

            column=0,

            sticky="w",

            pady=6,

        )



        cell_entry = ttk.Entry(

            container,

            textvariable=cell_var,

            width=32,

        )

        cell_entry.grid(

            row=1,

            column=1,

            pady=6,

        )



        result = {"value": None}



        def submit():

            field_value = field_var.get().strip()

            cell_value = cell_var.get().strip().upper()



            if not field_value or not cell_value:

                messagebox.showwarning(

                    "Validation",

                    "Enter both the field name and Excel cell.",

                    parent=dialog,

                )

                return



            result["value"] = {

                "field": field_value,

                "cell": cell_value,

            }

            dialog.destroy()



        buttons = ttk.Frame(container)

        buttons.grid(

            row=2,

            column=0,

            columnspan=2,

            pady=(14, 0),

        )



        ttk.Button(

            buttons,

            text="OK",

            command=submit,

        ).pack(side="left", padx=5)



        ttk.Button(

            buttons,

            text="Cancel",

            command=dialog.destroy,

        ).pack(side="left", padx=5)



        field_entry.focus_set()

        self.root.wait_window(dialog)



        return result["value"]



    def add_mapping(self):

        result = self.mapping_dialog("Add Metadata Mapping")



        if result:

            self.mapping_tree.insert(

                "",

                "end",

                values=(result["field"], result["cell"]),

            )



            self.log(

                f"Mapping added: {result['field']} -> {result['cell']}"

            )



    def edit_mapping(self):

        selection = self.mapping_tree.selection()



        if not selection:

            messagebox.showinfo(

                "Edit Mapping",

                "Select a mapping to edit.",

            )

            return



        item_id = selection[0]

        old_field, old_cell = self.mapping_tree.item(

            item_id,

            "values",

        )



        result = self.mapping_dialog(

            "Edit Metadata Mapping",

            old_field,

            old_cell,

        )



        if result:

            self.mapping_tree.item(

                item_id,

                values=(result["field"], result["cell"]),

            )



            self.log(

                f"Mapping updated: {result['field']} -> {result['cell']}"

            )



    def delete_mapping(self):

        selection = self.mapping_tree.selection()



        if not selection:

            messagebox.showinfo(

                "Delete Mapping",

                "Select a mapping to delete.",

            )

            return



        if not messagebox.askyesno(

            "Confirm Delete",

            "Delete the selected mapping?",

        ):

            return



        for item_id in selection:

            values = self.mapping_tree.item(item_id, "values")

            self.mapping_tree.delete(item_id)



            self.log(f"Mapping deleted: {values[0]}")



    def get_current_settings(self):

        try:

            interval = int(self.interval_var.get())

            header_row = int(self.header_row_var.get())

            header_start_col = int(self.header_start_col_var.get())

            data_start_row = int(self.data_start_row_var.get())



            if interval < 1:

                raise ValueError("Interval must be at least 1 second.")



            if min(header_row, header_start_col, data_start_row) < 1:

                raise ValueError("Excel row and column values must be positive.")



        except ValueError as exc:

            raise ValueError(f"Invalid configuration: {exc}") from exc



        mappings = []



        for item_id in self.mapping_tree.get_children():

            field, cell = self.mapping_tree.item(

                item_id,

                "values",

            )



            mappings.append({

                "field": field,

                "cell": cell,

            })



        return {

            "api_url": self.api_url_var.get().strip(),

            "watch_directory": self.watch_directory_var.get().strip(),

            "interval": interval,

            "header_row": header_row,

            "header_start_col": header_start_col,

            "data_start_row": data_start_row,

            "start_on_windows_login": self.start_on_login_var.get(),

            "metadata_mappings": mappings,

        }



    def save_settings(self, show_message=True):

        try:

            settings = self.get_current_settings()



            if not settings["api_url"]:

                raise ValueError("API URL cannot be empty.")



            if not settings["watch_directory"]:

                raise ValueError("Select a watch directory first.")



            if not Path(settings["watch_directory"]).is_dir():

                raise ValueError("The selected watch directory does not exist.")



            with open(CONFIG_FILE, "w", encoding="utf-8") as file:

                json.dump(

                    settings,

                    file,

                    indent=4,

                    ensure_ascii=False,

                )



            self.config = settings



            self.log("Configuration saved successfully.")



            if show_message:

                messagebox.showinfo(

                    "Settings Saved",

                    f"Settings saved successfully.\n\n{CONFIG_FILE}",

                )



            return True



        except (ValueError, OSError) as exc:

            messagebox.showerror(

                "Settings Error",

                str(exc),

            )

            return False



    def start_monitoring(self):

        if self.running:

            return



        if not self.save_settings(show_message=False):

            return



        directory = Path(self.config["watch_directory"])



        if not directory.is_dir():

            messagebox.showerror(

                "Directory Error",

                "The selected watch directory does not exist.",

            )

            return



        existing_files = [
            path for path in directory.iterdir()
            if path.is_file()
            and path.suffix.lower() == ".xlsx"
            and not path.name.startswith("~$")
        ]



        self.stop_event.clear()

        self.running = True



        self.start_button.configure(state="disabled")

        self.stop_button.configure(state="normal")

        self.status_label.configure(

            text="● Monitoring",

            foreground="#15803D",

        )



        self.log(f"Monitoring started: {directory}")

        self.log(f"Excel files found for processing: {len(existing_files)}")
        self.log("Excel validation and API upload are enabled.")

        self.worker_thread = threading.Thread(

            target=self.monitor_folder,

            daemon=True,

        )

        self.worker_thread.start()



    def _file_signature(self, path):
        stat = path.stat()
        return (str(path.resolve()), stat.st_size, stat.st_mtime_ns)

    def _wait_until_file_ready(self, path, checks=3, delay=1):
        previous_size = -1
        stable_count = 0
        for _ in range(30):
            if self.stop_event.is_set():
                return False
            try:
                current_size = path.stat().st_size
                if current_size > 0 and current_size == previous_size:
                    stable_count += 1
                    if stable_count >= checks:
                        with open(path, "rb"):
                            return True
                else:
                    stable_count = 0
                    previous_size = current_size
            except OSError:
                stable_count = 0
            self.stop_event.wait(delay)
        return False

    def validate_excel_file(self, path):
        workbook = openpyxl.load_workbook(path, data_only=True, read_only=True)
        try:
            sheet = workbook.worksheets[0]

            def cell_text(row, col):
                value = sheet.cell(row=row, column=col).value
                return "" if value is None else str(value).strip()

            part_name = cell_text(2, 2)
            machine = cell_text(5, 5)
            report_time = cell_text(1, 7)
            if not part_name:
                raise ValueError("Part Name is missing from cell B2.")
            if not machine:
                raise ValueError("Machine is missing from cell E5.")
            if cell_text(7, 2).casefold() != "content":
                raise ValueError("Cell B7 must contain the Content header.")

            accepted_formats = (
                "%Y/%m/%d--%H:%M", "%Y/%m/%d %H:%M",
                "%Y-%m-%d %H:%M", "%Y-%m-%dT%H:%M:%S",
                "%Y/%m/%d--%H:%M:%S",
            )
            if not any(self._valid_datetime(report_time, fmt) for fmt in accepted_formats):
                raise ValueError(f"Invalid report time in G1: {report_time}")

            readings = []
            parameter = ""
            nominal = upper_tolerance = lower_tolerance = None

            for row in range(8, sheet.max_row + 1):
                content = cell_text(row, 2)
                actual = sheet.cell(row=row, column=3).value
                if not content or content.upper().startswith("SPC A #"):
                    continue
                try:
                    actual_value = float(actual)
                except (TypeError, ValueError):
                    continue

                if not parameter:
                    parameter = content
                    try:
                        nominal = float(sheet.cell(row=row, column=4).value)
                        upper_tolerance = float(sheet.cell(row=row, column=6).value)
                        lower_tolerance = float(sheet.cell(row=row, column=7).value)
                    except (TypeError, ValueError):
                        raise ValueError(f"Invalid Nominal/UpTol/LowTol at row {row}.")
                readings.append(actual_value)

            if len(readings) != 32:
                raise ValueError(f"Expected 32 numeric measurements; found {len(readings)}.")

            return {
                "part_name": part_name,
                "machine": machine,
                "parameter": parameter,
                "measurement_count": len(readings),
                "nominal": nominal,
                "usl": nominal + upper_tolerance,
                "lsl": nominal + lower_tolerance,
            }
        finally:
            workbook.close()

    @staticmethod
    def _valid_datetime(value, fmt):
        try:
            datetime.strptime(value, fmt)
            return True
        except (TypeError, ValueError):
            return False

    def upload_excel_file(self, path):
        details = self.validate_excel_file(path)
        api_url = self.config["api_url"].strip()
        if not api_url:
            raise ValueError("API URL is empty.")

        host = (urlparse(api_url).hostname or "").lower()
        # Local ASP.NET development certificates are not normally trusted by requests.
        verify_tls = host not in {"localhost", "127.0.0.1"}

        with open(path, "rb") as excel_file:
            response = requests.post(
                api_url,
                files={"file": (
                    path.name,
                    excel_file,
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )},
                headers={"Accept": "application/json"},
                timeout=(10, 90),
                verify=verify_tls,
            )

        if response.status_code == 200:
            try:
                result = response.json()
            except ValueError:
                raise RuntimeError("API returned HTTP 200 with an invalid JSON response.")
            if not result.get("success"):
                raise RuntimeError("API did not confirm successful saving: " + response.text[:1000])
            return result

        if response.status_code == 409:
            raise FileExistsError("This report was already imported: " + response.text[:1000])

        raise RuntimeError(
            f"API returned HTTP {response.status_code}: {response.text[:2000]}"
        )

    def process_one_file(self, path):
        try:
            if not self._wait_until_file_ready(path):
                self.log_queue.put(("error", f"File is not ready: {path.name}"))
                return
            signature = self._file_signature(path)
            if signature in self.processed_signatures:
                return

            self.log_queue.put(("info", f"Validating and uploading: {path.name}"))
            result = self.upload_excel_file(path)
            self.processed_signatures.add(signature)
            self.log_queue.put((
                "success",
                f"HTTP 200 — saved '{path.name}'. Report ID: {result.get('reportId')}; "
                f"Template ID: {result.get('templateId')}; "
                f"Measurements: {result.get('measurementCount')}."
            ))
        except FileExistsError as exc:
            try:
                self.processed_signatures.add(self._file_signature(path))
            except OSError:
                pass
            self.log_queue.put(("info", str(exc)))
        except Exception as exc:
            self.log_queue.put(("error", f"Upload failed for '{path.name}': {exc}"))

    def monitor_folder(self):
        directory = Path(self.config["watch_directory"])
        interval = self.config["interval"]

        while not self.stop_event.is_set():
            try:
                for path in directory.iterdir():
                    if self.stop_event.is_set():
                        break
                    if not path.is_file() or path.suffix.lower() != ".xlsx":
                        continue
                    if path.name.startswith("~$"):
                        continue

                    try:
                        signature = self._file_signature(path)
                        if signature in self.processed_signatures:
                            continue
                        last_attempt = self.last_attempt.get(signature, 0)
                        if time.time() - last_attempt < 30:
                            continue
                        self.last_attempt[signature] = time.time()
                        self.log_queue.put(("info", f"Excel file detected: {path.name}"))
                        self.process_one_file(path)
                    except OSError as exc:
                        self.log_queue.put(("error", f"Cannot access '{path.name}': {exc}"))
            except OSError as exc:
                self.log_queue.put(("error", f"Folder monitoring error: {exc}"))

            self.stop_event.wait(interval)

        self.log_queue.put(("stopped", "Folder monitoring stopped."))

    def stop_monitoring(self):

        if not self.running:

            return



        self.stop_event.set()

        self.running = False



        self.start_button.configure(state="normal")

        self.stop_button.configure(state="disabled")



        self.status_label.configure(

            text="● Stopped",

            foreground="#B45309",

        )



        self.log("Stop requested.")



    def log(self, message):

        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")



        self.log_queue.put(

            ("info", f"[{timestamp}] {message}")

        )



    def process_log_queue(self):

        try:

            while True:

                level, message = self.log_queue.get_nowait()



                self.log_text.configure(state="normal")

                self.log_text.insert("end", message + "\n")



                if level in ("error", "success"):

                    self.log_text.tag_add(

                        "error",

                        "end-2l",

                        "end-1l",

                    )

                    self.log_text.tag_config(

                        "error",

                        foreground="#FCA5A5",

                    )



                self.log_text.see("end")

                self.log_text.configure(state="disabled")



                if level == "stopped" and not self.running:

                    self.status_label.configure(

                        text="● Stopped",

                        foreground="#B45309",

                    )



        except queue.Empty:

            pass



        if self.root.winfo_exists():

            self.root.after(150, self.process_log_queue)



    def close_application(self):

        if self.running:

            if not messagebox.askyesno(

                "Exit Application",

                "Monitoring is running. Stop monitoring and exit?",

            ):

                return



            self.stop_event.set()

            self.running = False



        self.root.destroy()



def main():

    root = tk.Tk()

    ExcelWatcherApp(root)

    root.mainloop()

if __name__ == "__main__":

    main()
