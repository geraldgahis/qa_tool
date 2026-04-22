# sheet_writer.py
import os
import sys
import gspread
from google.oauth2.service_account import Credentials


def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

def get_sheet(spreadsheet_id, sheet_name):
    # USE THE HELPER HERE
    json_path = resource_path("credentials.json")
    creds = Credentials.from_service_account_file(json_path, scopes=SCOPES)
    client = gspread.authorize(creds)
    spreadsheet = client.open_by_key(spreadsheet_id)

    try:
        worksheet = spreadsheet.worksheet(sheet_name)
    except gspread.WorksheetNotFound:
        worksheet = spreadsheet.add_worksheet(title=sheet_name, rows=1000, cols=20)

    return worksheet


def write_assets_to_sheet(worksheet, weather_table: dict, placement_name: str, ad_type: str, log):
    """
    Writes the asset WEATHER TABLE exactly like the old logic.
    Appends data (never overwrites).
    """
    try:
        rows_to_write = []
        rows_to_write.append([])

        # placement title
        clean_placement_name = placement_name.replace(".html", "")
        rows_to_write.append([clean_placement_name])

        has_mw_fg = False
        has_close_btn = False
        has_expand_btn = False

        if ad_type.startswith("mw-"):
            has_mw_fg = any(
                weather_table[row]["OPEN_FG"]["urls"] or
                weather_table[row]["CLOSED_FG"]["urls"]
                for row in weather_table
            )
            has_close_btn = any(
                weather_table[row]["CLOSE_BTN"]["urls"]
                for row in weather_table
            )
            has_expand_btn = any(
                weather_table[row]["EXPAND_BTN"]["urls"]
                for row in weather_table
            )

        # ---------------- HEADER ----------------
        header = ["Weather"]

        if ad_type.startswith("ma-"):
            header.extend(["BG Small", "BG Large"])

            has_fg = any(
                weather_table[row]["FG"]["SMALL"]["urls"] or
                weather_table[row]["FG"]["LARGE"]["urls"]
                for row in weather_table
            )
            if has_fg:
                header.extend(["FG Small", "FG Large"])
            
            # Only include OTHER ASSETS if they have values
            has_other = any(
                weather_table[row]["OTHER_ASSETS"]["SMALL"]["urls"] or
                weather_table[row]["OTHER_ASSETS"]["LARGE"]["urls"]
                for row in weather_table
            )
            if has_other:
                header.extend(["Other Assets - Small", "Other Assets - Large"]) 

        elif ad_type.startswith("mw-"):
            header.extend(["Open BG", "Closed BG"])

            if has_mw_fg:
                header.extend(["Open FG", "Closed FG"])
            if has_close_btn:
                header.append("Close Button")
            if has_expand_btn:
                header.append("Expand Button")

            # Only include OTHER ASSETS if they have values
            has_other = any(
                any(url.lower().endswith((".png", ".jpg", ".jpeg")) for url in weather_table[row]["OTHER_ASSETS"])
                for row in weather_table
            )
            if has_other:
                header.append("Other Assets") 

        if ad_type.startswith("ma-"):
            header.append("Textmode") 


        rows_to_write.append(header)

        # ---------------- ROWS ----------------
        for row_label, sizes in weather_table.items():
            row = [row_label]

            if ad_type.startswith("ma-"):
                # BG
                bg_small = "\n".join(sizes["BG"]["SMALL"]["urls"])
                bg_large = "\n".join(sizes["BG"]["LARGE"]["urls"])
                row.extend([bg_small, bg_large])

                # FG if present
                if "FG Small" in header:
                    fg_small = "\n".join(sizes["FG"]["SMALL"]["urls"])
                    fg_large = "\n".join(sizes["FG"]["LARGE"]["urls"])
                    row.extend([fg_small, fg_large])

                # OTHER ASSETS only, now includes all files
                other_small = "\n".join(sizes["OTHER_ASSETS"]["SMALL"]["urls"])
                other_large = "\n".join(sizes["OTHER_ASSETS"]["LARGE"]["urls"])
                if "Other Assets - Small" in header:
                    row.append(other_small)
                if "Other Assets - Large" in header:
                    row.append(other_large)


                # TEXTMODE last column
                row.append(sizes["textmode"])

            elif ad_type.startswith("mw-"):
                # MW logic remains the same
                open_bg = "\n".join(sizes["OPEN_BG"]["urls"])
                closed_bg = "\n".join(sizes["CLOSED_BG"]["urls"])
                row.extend([open_bg, closed_bg])

                if has_mw_fg:
                    open_fg = "\n".join(sizes["OPEN_FG"]["urls"])
                    closed_fg = "\n".join(sizes["CLOSED_FG"]["urls"])
                    row.extend([open_fg, closed_fg])
                if has_close_btn:
                    row.append("\n".join(sizes["CLOSE_BTN"]["urls"]))
                if has_expand_btn:
                    row.append("\n".join(sizes["EXPAND_BTN"]["urls"]))

                # ONLY include OTHER_ASSETS images
                other_assets = "\n".join(
                    url for url in sizes["OTHER_ASSETS"]
                    if url.lower().endswith((".png", ".jpg", ".jpeg"))
                )
                if "Other Assets" in header:
                    row.append(other_assets)

            rows_to_write.append(row)

        rows_to_write.append([])

        worksheet.append_rows(rows_to_write, value_input_option="RAW")

        # ---------------- APPLY COLOR TO HEADER ----------------
        # Find all cells in Column A that contain exactly "Weather"
        header_cells = worksheet.findall("Weather", in_column=1)
        
        num_cols = len(header)
        col_letter = chr(64 + num_cols) if num_cols <= 26 else "Z"

        for cell in header_cells:
            # Color the row from Column A to the end of the header columns
            header_range = f"A{cell.row}:{col_letter}{cell.row}"
            
            worksheet.format(header_range, {
                "backgroundColor": {
                    "red": 0.27, 
                    "green": 0.74, 
                    "blue": 0.78
                }, # #46bdc6
                "textFormat": {
                    "foregroundColor": {"red": 0.0, "green": 0.0, "blue": 0.0}, 
                    "bold": True
                },
                "horizontalAlignment": "CENTER"
            })

    except Exception as e:
        log(f"Error writing asset data to Google Sheet: {e}")

def write_trackings_to_sheet(
    worksheet,
    tracking_header: list,
    tracking_rows: list,
    log,
    is_row_only=False
):
    """
    Writes tracking data to Google Sheet.
    Appends data (never overwrites).
    If is_row_only is True, it appends without adding headers or spacing.
    """
    try:
        rows_to_write = []

        if not is_row_only:
            # spacing row
            rows_to_write.append([])
            rows_to_write.append(["Tracking Results"]) # Optional Title

            # header
            clean_header = [h.strip() for h in tracking_header]
            rows_to_write.append(clean_header)

        # rows
        for row in tracking_rows:
            rows_to_write.append(row)

        # If we are writing a full block or the very last row, we might want spacing
        # But for immediate per-row appending, we only add spacing in block mode
        if not is_row_only:
            rows_to_write.append([])

        if rows_to_write:
            worksheet.append_rows(rows_to_write, value_input_option="RAW")
        
        # ---------------- APPLY COLOR TO HEADER ----------------
        # Find all cells in Column A that contain exactly "Placement Name"
        header_cells = worksheet.findall("Placement Name", in_column=1)
        
        num_cols = len(tracking_header)
        col_letter = chr(64 + num_cols) if num_cols <= 26 else "Z"

        for cell in header_cells:
            # Color the row from Column A to the end of the header columns
            header_range = f"A{cell.row}:{col_letter}{cell.row}"
            
            worksheet.format(header_range, {
                "backgroundColor": {
                    "red": 0.27, 
                    "green": 0.74, 
                    "blue": 0.78
                }, # #46bdc6
                "textFormat": {
                    "foregroundColor": {"red": 0.0, "green": 0.0, "blue": 0.0}, 
                    "bold": True
                },
                "horizontalAlignment": "CENTER"
            })

    except Exception as e:
        log(f"Error writing tracking data to Google Sheet: {e}")


def write_code_comparison_to_sheet(worksheet, headers, rows, comparison_storage, log, is_header_only=False, is_tracking_table=False):
    try:
        rows_to_write = []
        is_generic_mode = "CID" in headers and "Small CID" not in headers
        
        # 1) Add new SafeFrame columns dynamically
        if is_generic_mode:
            extra_headers = ["Lab Product", "Ad Size - CID", "Ad Size - Test GAM", "SafeFrame - CID", "SafeFrame - Test GAM"]
        else:
            extra_headers = [
                "Lab Product - Small", "Lab Product - Large", "Ad Size - Small", "Ad Size - Large", 
                "SafeFrame - Small CID", "SafeFrame - Large CID", "SafeFrame - Small Test GAM", "SafeFrame - Large Test GAM"
            ]
            
        new_headers = headers + extra_headers

        if is_header_only:
            rows_to_write.append([]) 
            rows_to_write.append(["Verify CID and GAM settings"])
            
            rows_to_write.append(new_headers)
            original_footer = comparison_storage.get("footer_row", [])
            if original_footer:
                padded_footer = list(original_footer) + ([""] * len(extra_headers))
                rows_to_write.append(padded_footer)
        
        elif is_tracking_table:
            rows_to_write.append([]) 
            rows_to_write.append([]) 
            rows_to_write.append(["Verify Advanced Settings"])

            footer_id_map = comparison_storage.get("footer_id_map", {})
            footer_headers = list(footer_id_map.keys()) 
            id_to_name = {str(v): k for k, v in footer_id_map.items()}
            results = comparison_storage.get("results", [])
            has_generic_data = any(res["found_ids"].get("CID") or res["found_ids"].get("Test GAM") for res in results)
            mapping_headers = ["CID", "Test GAM"] if has_generic_data else ["Small CID", "Large CID", "Small Test GAM", "Large Test GAM"]
            rows_to_write.append(["Placement Name"] + footer_headers + mapping_headers)

            for res in results:
                found_ids = res.get("found_ids", {})
                row_map = res.get("row_map", {}) 
                row_data = [res["placement_name"]]
                for h_name in footer_headers:
                    target_id = str(footer_id_map[h_name])
                    actual_value = row_map.get(h_name, "N/A")
                    is_found = any(target_id in label_list for label_list in found_ids.values())
                    row_data.append(actual_value if is_found else "")
                for h in mapping_headers:
                    label_ids = found_ids.get(h, [])
                    translated_names = [id_to_name.get(item_id, item_id) for item_id in label_ids]
                    row_data.append("\n".join(translated_names) if translated_names else "")
                rows_to_write.append(row_data)

        else:
            for idx, row in enumerate(rows):
                data = comparison_storage.get(idx, {})
                if is_generic_mode:
                    row_extras = [
                        data.get("lab_product_generic", "N/A"),
                        data.get("ad_size_cid", "N/A"),
                        data.get("ad_size_test_gam", "N/A"),
                        data.get("sf_cid", "N/A"),
                        data.get("sf_gam", "N/A")
                    ]
                else:
                    row_extras = [
                        data.get("lab_product_small", "N/A"),
                        data.get("lab_product_large", "N/A"),
                        data.get("ad_size_small", "N/A"),
                        data.get("ad_size_large", "N/A"),
                        data.get("sf_small_cid", "N/A"),
                        data.get("sf_large_cid", "N/A"),
                        data.get("sf_small_gam", "N/A"),
                        data.get("sf_large_gam", "N/A")
                    ]
                rows_to_write.append(list(row) + row_extras)
                
            original_footer = comparison_storage.get("footer_row", [])
            if original_footer:
                padded_footer = list(original_footer) + ([""] * len(extra_headers))
                rows_to_write.append(padded_footer)
            
        if rows_to_write:
            worksheet.append_rows(rows_to_write, value_input_option="RAW")
        
        # ---------------- APPLY COLOR TO HEADER ----------------
        # Fixed: We execute the styling logic whenever we write headers
        if is_header_only or is_tracking_table:
            # Safely grab the first column header (usually 'Placement Name')
            search_term = "Placement Name" if not headers else headers[0]
            header_cells = worksheet.findall(search_term, in_column=1)
            
            # Dynamically calculate how many columns long the highlight should stretch
            target_num_cols = len(new_headers) if is_header_only else (1 + len(footer_headers) + len(mapping_headers))
            
            # Helper to convert col number to Excel-style Letter (e.g., 28 -> AB)
            col_letter = ""
            temp = target_num_cols
            while temp > 0:
                temp, remainder = divmod(temp - 1, 26)
                col_letter = chr(65 + remainder) + col_letter

            for cell in header_cells:
                header_range = f"A{cell.row}:{col_letter}{cell.row}"
                
                worksheet.format(header_range, {
                    "backgroundColor": {
                        "red": 0.27, 
                        "green": 0.74, 
                        "blue": 0.78
                    }, # #46bdc6
                    "textFormat": {
                        "foregroundColor": {"red": 0.0, "green": 0.0, "blue": 0.0}, 
                        "bold": True
                    },
                    "horizontalAlignment": "CENTER"
                })
                
    except Exception as e:
        log(f"Error writing row to sheet: {e}")