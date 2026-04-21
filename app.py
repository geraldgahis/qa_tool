# app.py

import gc
import os
import sys
import subprocess
import ttkbootstrap as ttk
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
        # We check if the folder exists first to avoid the subprocess call entirely if possible
        if not os.path.exists(os.environ["PLAYWRIGHT_BROWSERS_PATH"]):
            log_func("First time setup: Installing browser components (this may take a minute)...")
        
        # Use 'python' instead of sys.executable to avoid restarting the GUI
        # If this still fails in EXE, Playwright usually bundles a standalone tool
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
        # Custom format: Jan. 28, 2026 04:55:03
        timestamp = datetime.now().strftime("%b. %d, %Y %H:%M:%S")

        log_sh = authorized_gc.open_by_key(PRIVATE_LOG_SHEET_ID)
        try:
            log_ws = log_sh.get_worksheet(0) 
        except:
            log_ws = log_sh.sheet1
            
        # Append: Timestamp | User | Computer | Task | Spreadsheet URL
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
        self.mode_var = ttk.StringVar(value="all")

        self.create_input_row()
        self.create_mode_row()
        self.create_results_view()

        # INSIDE QATool.__init__
        self.creds = Credentials.from_service_account_file(
            SERVICE_ACCOUNT_FILE, scopes=SCOPES
        )
        self.gc = gspread.authorize(self.creds)
        
        # Pass self.gc to the function
        # threading.Thread(target=log_to_sheet, args=(self.gc, "App Opened", "N/A"), daemon=True).start()

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

        type_lbl = ttk.Label(type_row, text="Task", width=8)
        type_lbl.pack(side=LEFT)

        self.mode_mapping = {
            "all": "All",
            "capture_assets": "Capture Assets",
            "capture_tracking": "Capture Tracking",
            "verify_cid_gam": "Verify CID & GAM",
            "code_comparison": "Code Comparison",
        }

        self.mode_buttons = []

        def log_mode_change():
            mode_text = self.mode_mapping.get(self.mode_var.get(), self.mode_var.get())
            self.log(f"[TASK] {mode_text}")

        for value, text in self.mode_mapping.items():
            opt = ttk.Radiobutton(
                master=type_row,
                text=text,
                variable=self.mode_var,
                value=value,
                command=log_mode_change
            )
            opt.pack(side=LEFT, padx=0 if value == "all" else 15)
            self.mode_buttons.append(opt)

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
        
        # --- LOG THE TASK AND THE URL ---
        selected_mode = self.mode_mapping.get(self.mode_var.get(), "Unknown")
        log_action = f"{selected_mode}"
        
        # We pass self.gc, the task name, and the URL
        threading.Thread(
            target=log_to_sheet, 
            args=(self.gc, log_action, spreadsheet_url), 
            daemon=True
        ).start()
        # --------------------------------

        threading.Thread(
            target=self.open_spreadsheet,
            args=(spreadsheet_url,),
            daemon=True
        ).start()

    def run_by_mode(self, cos_links, tracking_rows, tracking_header, tracking_rows_full, spreadsheet):
        mode = self.mode_var.get()

        try:
            if mode == "all":
                self.log("[MODE] Running ALL tasks")

                self.log("Starting Asset and Textmode extraction...")
                asyncio.run(run_assets_scraper(cos_links, self.log, spreadsheet))
                self.log("Asset extraction completed.\n")

                self.log("Starting Tracking extraction...")
                asyncio.run(run_tracking_scraper(cos_links, tracking_rows, tracking_header, self.log, spreadsheet))
                self.log("Tracking extraction completed.")

                self.log("Starting CID & GAM Verification...")
                asyncio.run(run_verification(self.cos_links, tracking_rows_full, self.log, spreadsheet))
                self.log("CID & GAM Verification completed.")
                
                self.log("\nStarting Code Comparison...")
                asyncio.run(run_code_comparison(self.cos_links, tracking_rows_full, self.log, spreadsheet))
                self.log("Code Comparison completed.")

            elif mode == "capture_assets":
                self.log("Starting Assets & Textmode extraction...")
                asyncio.run(run_assets_scraper(cos_links, self.log, spreadsheet))
                self.log("Asset extraction completed.")

            elif mode == "capture_tracking":
                self.log("Starting Tracking extraction...")
                asyncio.run(run_tracking_scraper(cos_links, tracking_rows, tracking_header, self.log, spreadsheet))
                self.log("Tracking extraction completed.")
            
            elif mode == "verify_cid_gam":
                self.log("Starting CID & GAM Verification...")
                asyncio.run(run_verification(self.cos_links, tracking_rows_full, self.log, spreadsheet))
                self.log("CID & GAM Verification completed.")
            
            elif mode == "code_comparison":
                self.log("Starting Code Comparison...")
                asyncio.run(run_code_comparison(self.cos_links, tracking_rows_full, self.log, spreadsheet))
                self.log("Code Comparison completed.")

            else:
                self.log(f"[WARN] Unsupported mode: {mode}")

        except Exception as e:
            self.log(f"Error during execution: {e}")

    def _extract_c6_rows_full(self, values, c6_items):
        """
        Extracts Header, Matching Data, and a raw Footer.
        Removes the original last column to make room for COS Links.
        Strictly verifies that all c6_items are found.
        """
        if not values or not c6_items:
            return []

        # 1. Clean input items to handle hidden spaces
        c6_items_clean = [str(item).strip() for item in c6_items if item.strip()]
        
        match_indices = []
        names_found = set()

        # 2. Find matches using stripped comparison
        for idx, row in enumerate(values):
            # Clean the row for checking
            row_str = [str(cell).strip() for cell in row]
            for target in c6_items_clean:
                if target in row_str:
                    match_indices.append(idx)
                    names_found.add(target)

        # 3. STRICT CHECK: If count doesn't match, return empty to stop execution
        if len(names_found) < len(c6_items_clean):
            # We don't log here because _extract_c6_rows already logged the missing names
            return []

        first_match = min(match_indices)
        last_match = max(match_indices)

        all_rows = []
        last_values = [None] * len(values[0])
        
        # 4. Resolve merged cells and extract range
        for idx in range(last_match + 1):
            row = values[idx]
            resolved_row = []
            
            # Use [:-1] to remove the original last column
            for i, cell in enumerate(row[:-1]):
                val = str(cell).strip()
                if val:
                    last_values[i] = val
                    resolved_row.append(val)
                else:
                    resolved_row.append(last_values[i] if i < len(last_values) else "")
            
            # Start including rows from one row above the first match (the header)
            if idx >= max(0, first_match - 1):
                all_rows.append(resolved_row)

        # 5. Handle Footer (Raw, removing last column)
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
        
        # Clean the input names from C6
        c6_items_clean = [str(item).strip() for item in c6_items if item.strip()]
        names_found_in_sheet = set()

        if not values:
            return None, []

        # 1. Identify Header
        end_index = None
        for idx, row in enumerate(values):
            # Clean the row for checking
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

        # 2. Extract Data with Merged Cell Handling
        last_values = [None] * len(values[0])
        for row in values:
            resolved_row = []
            for i, cell in enumerate(row):
                val = str(cell).strip()
                if val:
                    last_values[i] = val
                resolved_row.append(last_values[i] if i < len(last_values) else "")

            # Check for matches using the cleaned names
            for item in c6_items_clean:
                if item in resolved_row:
                    names_found_in_sheet.add(item)
                    subset = resolved_row[:end_index] if end_index else resolved_row
                    row_key = tuple(subset)
                    if row_key not in seen_keys:
                        found_rows.append(subset)
                        seen_keys.add(row_key)

        # 3. Final Validation
        missing = set(c6_items_clean) - names_found_in_sheet
        if missing:
            # We don't return rows if even one is missing (per your rule)
            return tracking_header, []

        return tracking_header, found_rows


    # ---------------- Spreadsheet Logic ---------------- #

    def open_spreadsheet(self, url):
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

            # Access the tracking sheet
            try:
                final_sheet = sh.worksheet("Final CIDs & Tracking")
                final_values = final_sheet.get_all_values()
            except Exception:
                self.log("[ERROR] Could not find sheet named 'Final CIDs & Tracking'.")
                return

            # --- UPDATED VALIDATION STEP ---
            tracking_header, tracking_rows = self._extract_c6_rows(final_values, c6_items)
            
            # 1. Check if header was found
            if not tracking_header:
                self.log("\n[✕][STOPPED] Could not find the tracking header (Small CID/CID).")
                return
            
            if not tracking_rows:
                self.log("\n[✕][STOPPED] One or more Creative Names not found. Please check if Creative Names in your DTC match exactly with those in 'Final CIDs & Tracking'.")
                return
        
            tracking_rows_full = self._extract_c6_rows_full(final_values, c6_items)
            
            self.log(f"[✓][Success] All items verified.")

            # ---------------- Run Core Tasks ---------------- #
            # Using threading to prevent UI freeze during scrapers
            threading.Thread(
                target=self.run_by_mode,
                args=(self.cos_links, tracking_rows, tracking_header, tracking_rows_full, sh),
                daemon=True
            ).start()

        except Exception as e:
            self.log(f"Error: {e}")

    # ---------------- Helpers ---------------- #

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


    # ---------------- Core Tasks ---------------- #

    def run_assets_task(self, cos_links, spreadsheet):
        """
        Run the Assets + Textmode extraction using the assets_scraper.
        This is called in a separate thread to avoid freezing the UI.
        """
        
        try:
            self.log("Starting Assets & Textmode extraction...")
            asyncio.run(run_assets_scraper(cos_links, self.log, spreadsheet))
            self.log("Assets extraction completed.")
        except Exception as e:
            self.log(f"Error in Assets extraction: {e}")

    def run_tracking_task(self, cos_links, spreadsheet):
        """
        Placeholder for future tracking scraper task.
        """
        try:
            self.log("Starting Tracking extraction...")
            asyncio.run(run_tracking_scraper(cos_links, self.log, spreadsheet))
            self.log("Tracking extraction completed.")
        except Exception as e:
            self.log(f"Error in Tracking extraction: {e}")



if __name__ == "__main__":
    import multiprocessing
    multiprocessing.freeze_support()
    app = ttk.Window("QA TOOL by QA POOL", "cyborg")
    app.minsize(1200, 550)

    try:
        icon_path = resource_path("favicon.ico")
        if os.path.exists(icon_path):
            # iconbitmap is the correct method for .ico files on Windows
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
