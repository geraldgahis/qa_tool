# app.py

import gc
import os
import sys
import subprocess
import ttkbootstrap as ttk
import tkinter as tk
import threading
import gspread
import asyncio
import platform
import getpass

from datetime import datetime
from urllib.parse import parse_qs, urlparse
from tkinter.scrolledtext import ScrolledText
from ttkbootstrap.constants import *
from google.oauth2.service_account import Credentials

from assets_scraper import run_assets_scraper
from code_comparison import run_code_comparison
from tracking_scraper import run_tracking_scraper
from verify_cid_gam import run_verification

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

# Update this line in your code:
SERVICE_ACCOUNT_FILE = resource_path("credentials.json")

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

def ensure_playwright_browser(log_func):
    """Check if Chromium is installed, if not, download it."""
    os.environ["PLAYWRIGHT_BROWSERS_PATH"] = os.path.join(os.environ["LOCALAPPDATA"], "ms-playwright")
    
    try:
        if not os.path.exists(os.environ["PLAYWRIGHT_BROWSERS_PATH"]):
            log_func("First time setup: Installing browser components (this may take a minute)...")
        
        subprocess.run(
            [sys.executable, "-m", "playwright", "install", "chromium"], 
            check=True, 
            capture_output=True, 
            creationflags=0x08000000 
        )
    except Exception as e:
        log_func(f"Browser check finished/skipped: {e}")

def log_to_sheet(authorized_gc, action="App Launched", extra_info="N/A"): 
    PRIVATE_LOG_SHEET_ID = "17oBaoMnCJHLzEXz5tZ9U7p2MFgrySe4emUp8D48U6GI"

    try:
        user_name = getpass.getuser()
        computer_name = platform.node()
        timestamp = datetime.now().strftime("%b. %d, %Y %H:%M:%S")

        log_sh = authorized_gc.open_by_key(PRIVATE_LOG_SHEET_ID)
        try:
            log_ws = log_sh.get_worksheet(0) 
        except:
            log_ws = log_sh.sheet1
            
        log_ws.append_row(
            [timestamp, user_name, computer_name, action, extra_info], 
            value_input_option='RAW'
        )
    except Exception as e:
        print(f"Logging failed: {e}")

class QATool(ttk.Frame):
    def __init__(self, master):
        super().__init__(master, padding=15)
        self.pack(fill=BOTH, expand=YES)

        self.option_lf = ttk.Labelframe(
            self, text="Insert DTC link to start.", padding=15
        )
        self.option_lf.pack(fill=X, expand=YES, anchor=N)

        self.create_input_row()
        self.create_mode_row()
        self.create_results_view()

        self.creds = Credentials.from_service_account_file(
            SERVICE_ACCOUNT_FILE, scopes=SCOPES
        )
        self.gc = gspread.authorize(self.creds)

    # ---------------- UI ---------------- #

    def create_input_row(self):
        row = ttk.Frame(self.option_lf)
        row.pack(fill=X, expand=YES, pady=10)

        ttk.Label(row, text="DTC", width=8).pack(side=LEFT)

        self.url_entry = ttk.Entry(row, bootstyle=DANGER)
        self.url_entry.pack(side=LEFT, fill=X, expand=YES, padx=5)

        ttk.Button(
            row,
            text="Run",
            bootstyle=DANGER,
            width=12,
            command=self.on_submit
        ).pack(side=LEFT, padx=5)

    def create_mode_row(self):
        type_row = ttk.Frame(self.option_lf)
        type_row.pack(fill=X, expand=YES, pady=(5, 10))

        type_lbl = ttk.Label(type_row, text="Tasks", width=8)
        type_lbl.pack(side=LEFT)

        self.mode_mapping = {
            "capture_assets": "Capture Assets",
            "capture_tracking": "Capture Tracking",
            "verify_cid_gam": "Verify CID & GAM",
            "code_comparison": "Code Comparison",          
        }

        self.task_vars = {}
        
        for key, text in self.mode_mapping.items():
            var = tk.BooleanVar(value=True)
            self.task_vars[key] = var
            
            cb = ttk.Checkbutton(
                master=type_row,
                text=text,
                variable=var,
                bootstyle="round-toggle" 
            )
            cb.pack(side=LEFT, padx=10)

        self.btn_retry = ttk.Button(
            master=type_row,
            text="Check Placement",
            bootstyle=WARNING,
            state="normal", 
            command=self.start_check_placement
        )
        self.btn_retry.pack(side=RIGHT, padx=5)

    def create_results_view(self):
        style = ttk.Style()
        self.textbox = ScrolledText(
            master=self,
            highlightcolor=style.colors.primary,
            highlightbackground=style.colors.border,
            highlightthickness=1,
        )
        self.textbox.pack(fill=BOTH, expand=YES, pady=10)
        self.textbox.config(state=DISABLED)

    def log(self, message: str):
        self.textbox.config(state="normal")
        self.textbox.insert("end", message + "\n")
        self.textbox.see("end")
        self.textbox.config(state="disabled")

    # ---------------- Actions ---------------- #

    def on_submit(self):
        spreadsheet_url = self.url_entry.get().strip()

        if not spreadsheet_url.startswith("https://docs.google.com/"):
            self.log("Invalid Google Sheet URL.")
            return
            
        selected_tasks = [k for k, v in self.task_vars.items() if v.get()]
        if not selected_tasks:
            self.log("[WARNING] Please select at least one task to run.")
            return
        
        task_names = [self.mode_mapping[k] for k in selected_tasks]
        log_action = " | ".join(task_names)
        
        threading.Thread(
            target=log_to_sheet, 
            args=(self.gc, log_action, spreadsheet_url), 
            daemon=True
        ).start()

        threading.Thread(
            target=self.open_spreadsheet,
            args=(spreadsheet_url, selected_tasks),
            daemon=True
        ).start()

    def execute_tasks(self, selected_tasks, cos_links, tracking_rows, tracking_header, tracking_rows_full, spreadsheet):
        try:
            if "verify_cid_gam" in selected_tasks:
                self.log("Starting CID & GAM Verification...")
                asyncio.run(run_verification(cos_links, tracking_rows_full, self.log, spreadsheet))
                self.log("CID & GAM Verification completed.\n")
                
            if "code_comparison" in selected_tasks:
                self.log("Starting Code Comparison...")
                asyncio.run(run_code_comparison(cos_links, tracking_rows_full, self.log, spreadsheet))
                self.log("Code Comparison completed.\n")

            if "capture_assets" in selected_tasks:
                self.log("Starting Asset and Textmode extraction...")
                asyncio.run(run_assets_scraper(cos_links, self.log, spreadsheet))
                self.log("Asset extraction completed.\n")

            if "capture_tracking" in selected_tasks:
                self.log("Starting Tracking extraction...")
                asyncio.run(run_tracking_scraper(cos_links, tracking_rows, tracking_header, self.log, spreadsheet))
                self.log("Tracking extraction completed.\n")

            self.log("--- 🏁 All selected tasks finished successfully! ---")

        except Exception as e:
            self.log(f"Error during execution: {e}")

    def _extract_c6_rows_full(self, values, c6_items):
        if not values or not c6_items:
            return []

        c6_items_clean = [str(item).strip() for item in c6_items if item.strip()]
        match_indices = []
        names_found = set()

        for idx, row in enumerate(values):
            row_str = [str(cell).strip() for cell in row]
            for target in c6_items_clean:
                if target in row_str:
                    match_indices.append(idx)
                    names_found.add(target)

        if len(names_found) < len(c6_items_clean):
            return []

        first_match = min(match_indices)
        last_match = max(match_indices)

        all_rows = []
        last_values = [None] * len(values[0])
        
        for idx in range(last_match + 1):
            row = values[idx]
            resolved_row = []
            
            for i, cell in enumerate(row[:-1]):
                val = str(cell).strip()
                if val:
                    last_values[i] = val
                    resolved_row.append(val)
                else:
                    resolved_row.append(last_values[i] if i < len(last_values) else "")
            
            if idx >= max(0, first_match - 1):
                all_rows.append(resolved_row)

        footer_idx = last_match + 1
        if footer_idx < len(values):
            raw_footer = [str(c).strip() for c in values[footer_idx][:-1]]
            while len(raw_footer) < len(all_rows[0]):
                raw_footer.append("")
            all_rows.append(raw_footer)

        return all_rows

    def _extract_c6_rows(self, values, c6_items):
        found_rows = []
        seen_keys = set()
        tracking_header = None
        
        c6_items_clean = [str(item).strip() for item in c6_items if item.strip()]
        names_found_in_sheet = set()

        if not values:
            return None, []

        end_index = None
        for idx, row in enumerate(values):
            clean_row = [str(cell).strip() for cell in row]
            if any(item in clean_row for item in c6_items_clean):
                if idx > 0:
                    header_row = values[idx-1]
                    for i, col in enumerate(header_row):
                        if str(col).strip() in ("Small CID", "CID"):
                            end_index = i
                            break
                    tracking_header = header_row[:end_index] if end_index else header_row
                break

        last_values = [None] * len(values[0])
        for row in values:
            resolved_row = []
            for i, cell in enumerate(row):
                val = str(cell).strip()
                if val:
                    last_values[i] = val
                resolved_row.append(last_values[i] if i < len(last_values) else "")

            for item in c6_items_clean:
                if item in resolved_row:
                    names_found_in_sheet.add(item)
                    subset = resolved_row[:end_index] if end_index else resolved_row
                    row_key = tuple(subset)
                    if row_key not in seen_keys:
                        found_rows.append(subset)
                        seen_keys.add(row_key)

        missing = set(c6_items_clean) - names_found_in_sheet
        if missing:
            return tracking_header, []

        return tracking_header, found_rows

    # ---------------- Spreadsheet Logic ---------------- #

    def open_spreadsheet(self, url, selected_tasks):
        try:
            self.log("Opening spreadsheet...")
            sh = self.gc.open_by_url(url)

            worksheet = self._get_worksheet_from_url(sh, url)
            if not worksheet:
                self.log("Could not determine worksheet from URL.")
                return

            cos_value = worksheet.acell("C3").value
            c6_value = worksheet.acell("C6").value

            if not cos_value or not c6_value:
                self.log("[ERROR] COS links or Creative Names are empty.")
                return

            self.cos_links = [line.strip() for line in cos_value.splitlines() if line.strip()]
            c6_items = [line.strip() for line in c6_value.splitlines() if line.strip()]

            self.log(f"Checking {len(c6_items)} Creative Names against 'Final CIDs & Tracking'...")

            try:
                final_sheet = sh.worksheet("Final CIDs & Tracking")
                final_values = final_sheet.get_all_values()
            except Exception:
                self.log("[ERROR] Could not find sheet named 'Final CIDs & Tracking'.")
                return

            self.tracking_header, self.tracking_rows = self._extract_c6_rows(final_values, c6_items)
            
            if not self.tracking_header:
                self.log("\n[✕][STOPPED] Could not find the tracking header (Small CID/CID).")
                return
            
            if not self.tracking_rows:
                self.log("\n[✕][STOPPED] One or more Creative Names not found. Please check if Creative Names in your DTC match exactly with those in 'Final CIDs & Tracking'.")
                return
        
            self.tracking_rows_full = self._extract_c6_rows_full(final_values, c6_items)
            self.spreadsheet = sh
            
            self.log(f"[✓][Success] All items verified.")

            # ---------------- Run Core Tasks ---------------- #
            threading.Thread(
                target=self.execute_tasks,
                args=(selected_tasks, self.cos_links, self.tracking_rows, self.tracking_header, self.tracking_rows_full, self.spreadsheet),
                daemon=True
            ).start()

        except Exception as e:
            self.log(f"Error: {e}")

    def _get_worksheet_from_url(self, sh, url):
        parsed = urlparse(url)
        qs = parse_qs(parsed.fragment)
        gid_list = qs.get("gid")

        if not gid_list:
            return sh.sheet1

        gid = int(gid_list[0])
        for ws in sh.worksheets():
            if ws.id == gid:
                return ws
        return None

    # ---------------- CORE: CHECK PLACEMENTS LOGIC ---------------- #

    def start_check_placement(self):
        spreadsheet_url = self.url_entry.get().strip()

        if not spreadsheet_url.startswith("https://docs.google.com/"):
            self.log("Invalid Google Sheet URL.")
            return
            
        if hasattr(self, 'tracking_rows_full') and self.tracking_rows_full:
            self.open_selection_modal()
        else:
            self.log("Fetching placements from sheet to check...")
            self.btn_retry.config(state="disabled")
            threading.Thread(
                target=self.fetch_and_open_modal,
                args=(spreadsheet_url,),
                daemon=True
            ).start()

    def fetch_and_open_modal(self, url):
        try:
            sh = self.gc.open_by_url(url)
            worksheet = self._get_worksheet_from_url(sh, url)
            if not worksheet:
                self.log("Could not determine worksheet from URL.")
                self.after(0, lambda: self.btn_retry.config(state="normal"))
                return

            cos_value = worksheet.acell("C3").value
            c6_value = worksheet.acell("C6").value

            if not cos_value or not c6_value:
                self.log("[ERROR] COS links or Creative Names are empty.")
                self.after(0, lambda: self.btn_retry.config(state="normal"))
                return

            self.cos_links = [line.strip() for line in cos_value.splitlines() if line.strip()]
            c6_items = [line.strip() for line in c6_value.splitlines() if line.strip()]

            final_sheet = sh.worksheet("Final CIDs & Tracking")
            final_values = final_sheet.get_all_values()

            self.tracking_header, self.tracking_rows = self._extract_c6_rows(final_values, c6_items)
            
            if not self.tracking_header:
                self.log("\n[✕][STOPPED] Could not find the tracking header.")
                self.after(0, lambda: self.btn_retry.config(state="normal"))
                return
            
            if not self.tracking_rows:
                self.log("\n[✕][STOPPED] Creative Names not found.")
                self.after(0, lambda: self.btn_retry.config(state="normal"))
                return
        
            self.tracking_rows_full = self._extract_c6_rows_full(final_values, c6_items)
            self.spreadsheet = sh
            
            self.after(0, lambda: self.btn_retry.config(state="normal"))
            self.after(0, self.open_selection_modal)

        except Exception as e:
            self.log(f"Error fetching data: {e}")
            self.after(0, lambda: self.btn_retry.config(state="normal"))


    def open_selection_modal(self):
        if not hasattr(self, 'tracking_rows_full') or not self.tracking_rows_full:
            self.log("[ERROR] No data available.")
            return

        modal = tk.Toplevel(self)
        modal.title("Select Placements to Check")
        
        # Center the modal dynamically
        modal.update_idletasks()
        w, h = 600, 650
        x = (self.winfo_screenwidth() // 2) - (w // 2)
        y = (self.winfo_screenheight() // 2) - (h // 2)
        modal.geometry(f"{w}x{h}+{x}+{y}")
        modal.transient(self.winfo_toplevel())
        modal.grab_set()

        # Data Setup
        header_row_full = self.tracking_rows_full[0]
        footer_row_full = []
        if len(self.tracking_rows_full) > 2 and any("setting" in str(cell).lower() for cell in self.tracking_rows_full[-1]):
            data_rows_full = self.tracking_rows_full[1:-1]
            footer_row_full = self.tracking_rows_full[-1]
        else:
            data_rows_full = self.tracking_rows_full[1:]

        data_rows_tracking = self.tracking_rows

        # --- 1. PLACEMENT SELECTION FRAME (Top) ---
        placement_frame = ttk.Labelframe(modal, text=f"Select Placements ({len(data_rows_full)})", padding=10)
        placement_frame.pack(fill=BOTH, expand=True, padx=10, pady=10)

        # Top controls for placements (Search + Buttons)
        top_p_frame = ttk.Frame(placement_frame)
        top_p_frame.pack(fill=X, pady=(0, 5))

        # Search Bar
        ttk.Label(top_p_frame, text="Search:").pack(side=LEFT, padx=(0, 5))
        search_var = tk.StringVar()
        search_entry = ttk.Entry(top_p_frame, textvariable=search_var)
        search_entry.pack(side=LEFT, fill=X, expand=True, padx=(0, 10))
        
        def set_all(state):
            # Only affect visible (filtered) checkboxes
            for cb, var, name in cb_items:
                if cb.winfo_manager(): 
                    var.set(state)

        ttk.Button(top_p_frame, text="Deselect All", bootstyle="outline", command=lambda: set_all(False)).pack(side=RIGHT, padx=2)
        ttk.Button(top_p_frame, text="Select All", bootstyle="outline", command=lambda: set_all(True)).pack(side=RIGHT, padx=2)

        # Scrollable Canvas
        canvas_frame = ttk.Frame(placement_frame)
        canvas_frame.pack(fill=BOTH, expand=True)

        canvas = tk.Canvas(canvas_frame, highlightthickness=0)
        scrollbar = ttk.Scrollbar(canvas_frame, orient="vertical", command=canvas.yview)
        scrollable_content = ttk.Frame(canvas)

        scrollable_content.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=scrollable_content, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side=LEFT, fill=BOTH, expand=True)
        scrollbar.pack(side=RIGHT, fill=Y)

        # Placement Checkboxes & Search Logic
        checkbox_vars = []
        cb_items = [] # Store widget, var, and name for search filtering
        for idx, row in enumerate(data_rows_full):
            var = tk.BooleanVar(value=False)
            checkbox_vars.append(var)
            placement_name = row[0] if row else f"Row {idx+1}"
            cb = ttk.Checkbutton(scrollable_content, text=placement_name, variable=var)
            cb.pack(anchor="w", pady=2, padx=5)
            cb_items.append((cb, var, placement_name.lower()))

        def filter_placements(*args):
            query = search_var.get().lower()
            for cb, var, name in cb_items:
                if query in name:
                    cb.pack(anchor="w", pady=2, padx=5)
                else:
                    cb.pack_forget()

        search_var.trace_add("write", filter_placements)

        # --- 2. BOTTOM CONTAINER (Tasks + Run Button) ---
        bottom_container = ttk.Frame(modal, padding=10)
        bottom_container.pack(fill=X, side=BOTTOM)

        # Task Selection Frame
        task_frame = ttk.Labelframe(bottom_container, text="Tasks to Run", padding=10)
        task_frame.pack(fill=X, pady=(0, 10))
        
        modal_task_vars = {}
        for key, text in self.mode_mapping.items():
            var = tk.BooleanVar(value=self.task_vars[key].get())
            modal_task_vars[key] = var
            ttk.Checkbutton(task_frame, text=text, variable=var).pack(side=LEFT, padx=5)

        # Run Check Button
        def execute_selected():
            selected_tasks = [k for k, v in modal_task_vars.items() if v.get()]
            if not selected_tasks:
                self.log("[WARNING] No tasks selected in the modal.")
                return

            selected_indices = [i for i, var in enumerate(checkbox_vars) if var.get()]
            if not selected_indices:
                self.log("[WARNING] No placements selected in the modal.")
                return

            self.log(f"\n--- Checking {len(selected_indices)} placement(s) ---")
            
            filtered_cos_links = [self.cos_links[i] for i in selected_indices]
            
            filtered_tracking_rows_full = [header_row_full]
            for i in selected_indices:
                filtered_tracking_rows_full.append(data_rows_full[i])
            if footer_row_full:
                filtered_tracking_rows_full.append(footer_row_full)

            filtered_tracking_rows = []
            for i in selected_indices:
                if i < len(data_rows_tracking):
                    filtered_tracking_rows.append(data_rows_tracking[i])

            modal.destroy()
            
            threading.Thread(
                target=self._run_check_thread, 
                args=(selected_tasks, filtered_cos_links, filtered_tracking_rows, filtered_tracking_rows_full), 
                daemon=True
            ).start()

        ttk.Button(bottom_container, text="Run Check on Selected Placements", bootstyle="warning", command=execute_selected).pack(fill=X)

    def _run_check_thread(self, selected_tasks, filtered_cos_links, filtered_tracking_rows, filtered_tracking_rows_full):
        self.after(0, lambda: self.btn_retry.config(state="disabled"))
        try:
            self.execute_tasks(
                selected_tasks,
                filtered_cos_links, 
                filtered_tracking_rows, 
                self.tracking_header, 
                filtered_tracking_rows_full, 
                self.spreadsheet
            )
        except Exception as e:
            self.log(f"[ERROR] Check failed: {e}")
        finally:
            self.after(0, lambda: self.btn_retry.config(state="normal"))

if __name__ == "__main__":
    import multiprocessing
    multiprocessing.freeze_support()
    app = ttk.Window("QA TOOL by QA POOL", "cyborg")
    app.minsize(1200, 550)

    try:
        icon_path = resource_path("favicon.ico")
        if os.path.exists(icon_path):
            app.iconbitmap(icon_path)
    except Exception as e:
        print(f"Icon loading failed: {e}")

    app.update_idletasks()
    w, h = app.winfo_width(), app.winfo_height()
    x = (app.winfo_screenwidth() // 2) - (w // 2)
    y = (app.winfo_screenheight() // 2) - (h // 2)
    app.geometry(f"+{x}+{y}")

    QATool(app)
    app.mainloop()