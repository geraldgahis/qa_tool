import asyncio
import os
import difflib

from playwright.async_api import async_playwright
from utils.url_parser import extract_url_parts
from sheet_writer import get_sheet, write_code_comparison_to_sheet


async def code_comparing(tcl_number, placement_name_cos, snippets, log, browser_context):
    """Generates a wide A3 landscape PDF with synced side-by-side alignment and line numbers."""
    base_results_dir = "Code Comparison Results"
    
    beautify_script = """
    async (code) => {
        if (!window.beautifier) {
            await new Promise(r => {
                const s = document.createElement('script');
                s.src = 'https://cdnjs.cloudflare.com/ajax/libs/js-beautify/1.14.7/beautify-html.min.js';
                s.onload = r;
                document.head.appendChild(s);
            });
        }
        return html_beautify(code, { indent_size: 2, wrap_line_length: 120, preserve_newlines: true });
    }
    """

    has_prefixed = any(key in snippets for key in ["Small CID", "Large CID", "Small Test GAM", "Large Test GAM"])
    if has_prefixed:
        comparison_pairs = [
            ("Small CID", "Large CID", "Small CID vs Large CID"),
            ("Small CID", "Small Test GAM", "Small CID vs Small Test GAM"),
            ("Small CID", "Large Test GAM", "Small CID vs Large Test GAM"),
            ("Small CID", "COS Link", "Small CID vs COS Link"),
            ("Large CID", "Small Test GAM", "Large CID vs Small Test GAM"),
            ("Large CID", "Large Test GAM", "Large CID vs Large Test GAM"),
            ("Large CID", "COS Link", "Large CID vs COS Link"),
            ("Small Test GAM", "Large Test GAM", "Small Test GAM vs Large Test GAM"),
            ("Small Test GAM", "COS Link", "Small Test GAM vs COS Link"),
            ("Large Test GAM", "COS Link", "Large Test GAM vs COS Link")
        ]
    else:
        comparison_pairs = [
            ("CID", "Test GAM", "CID vs Test GAM"),
            ("CID", "COS Link", "CID vs COS Link"),
            ("Test GAM", "COS Link", "Test GAM vs COS Link")
        ]

    format_page = await browser_context.new_page()
    await format_page.goto("about:blank")

    for left_key, right_key, folder_name in comparison_pairs:
        raw_left = snippets.get(left_key, "")
        raw_right = snippets.get(right_key, "")
        if not raw_left or not raw_right: continue

        left_code = await format_page.evaluate(beautify_script, raw_left)
        right_code = await format_page.evaluate(beautify_script, raw_right)

        # Generate the diff
        diff = list(difflib.unified_diff(
            left_code.splitlines(), 
            right_code.splitlines(), 
            fromfile=left_key, 
            tofile=right_key, 
            n=100000, 
            lineterm=''
        ))

        # FIX: If there is NO difference, unified_diff is empty. 
        # We must fill it manually to show the code in the PDF.
        if not diff:
            # We treat the entire left_code as "unchanged" (starting with a space)
            diff = [f"--- {left_key}", f"+++ {right_key}"] + [f" {line}" for line in left_code.splitlines()]
        
        has_real_diff = len([l for l in diff if l.startswith(('+', '-')) and not l.startswith(('+++', '---'))]) > 0
        final_folder_name = f"{folder_name} (diff found)" if has_real_diff else folder_name

        target_dir = os.path.join(base_results_dir, tcl_number, final_folder_name)
        os.makedirs(target_dir, exist_ok=True)
        
        clean_filename = placement_name_cos.replace(".html", "")
        pdf_path = os.path.join(target_dir, f"{clean_filename}.pdf")
        html_path = os.path.join(target_dir, "temp_diff.html")

        left_html_lines, right_html_lines = [], []
        l_num, r_num = 1, 1

        def make_row(num, text, cls):
            n_str = str(num) if num else ""
            # Text uses <span> instead of direct text to ensure padding-left alignment for wrapped lines
            return f'<div class="code-row {cls}"><span class="line-num">{n_str}</span><span class="code-text">{text}</span></div>'

        # Sync Logic: Group consecutive - and +
        i = 0
        while i < len(diff):
            line = diff[i]
            if line.startswith('---'):
                left_html_lines.append(make_row(None, line.replace('<','&lt;'), "header-left"))
                i += 1
            elif line.startswith('+++'):
                right_html_lines.append(make_row(None, line.replace('<','&lt;'), "header-right"))
                i += 1
            elif line.startswith('@@'):
                row = make_row("...", line.replace('<','&lt;'), "hunk")
                left_html_lines.append(row); right_html_lines.append(row)
                i += 1
            elif line.startswith('-') or line.startswith('+'):
                rems, adds = [], []
                while i < len(diff) and diff[i].startswith('-'):
                    rems.append(diff[i]); i += 1
                while i < len(diff) and diff[i].startswith('+'):
                    adds.append(diff[i]); i += 1
                
                # Pair side-by-side; fills gaps with "empty" rows automatically
                for j in range(max(len(rems), len(adds))):
                    rem_text = rems[j].replace('<','&lt;') if j < len(rems) else ""
                    add_text = adds[j].replace('<','&lt;') if j < len(adds) else ""
                    
                    left_html_lines.append(make_row(l_num if j < len(rems) else "", rem_text, "removed" if j < len(rems) else "empty"))
                    right_html_lines.append(make_row(r_num if j < len(adds) else "", add_text, "added" if j < len(adds) else "empty"))
                    
                    if j < len(rems): l_num += 1
                    if j < len(adds): r_num += 1
            else:
                txt = line[1:].replace('<','&lt;') if line.startswith(' ') else line.replace('<','&lt;')
                left_html_lines.append(make_row(l_num, txt, ""))
                right_html_lines.append(make_row(r_num, txt, ""))
                l_num += 1; r_num += 1; i += 1

        html_content = f"""
        <html>
        <head>
            <style>
                body {{ font-family: 'SFMono-Regular', Consolas, monospace; font-size: 14px; margin: 0; padding: 15px; background: #fff; }}
                .container {{ display: flex; border: 1px solid #d1d5da; align-items: stretch; }}
                .pane {{ flex: 1; width: 50%; min-width: 0; border-right: 1px solid #d1d5da; }}
                .pane:last-child {{ border-right: none; }}
                /* align-items: stretch ensures both number and text columns have the same height */
                .code-row {{ display: flex; align-items: stretch; line-height: 1.6; min-height: 1.6em; border-bottom: 0.5px solid #f0f0f0; }}
                .line-num {{ 
                    width: 45px; min-width: 45px; text-align: right; padding-right: 10px; 
                    color: #afb8c1; border-right: 1px solid #e1e4e8; background: #f6f8fa; 
                    font-size: 11px; padding-top: 2px; flex-shrink: 0;
                }}
                .code-text {{ 
                    flex: 1; padding: 0 10px; white-space: pre-wrap; word-break: break-all; 
                    text-indent: -20px; padding-left: 30px; display: block;
                }}
                .removed {{ background: #ffeef0; }} .added {{ background: #e6ffed; }} .hunk {{ background: #f1f8ff; color: #005cc5; }}
                .header-left {{ background: #ffdce0; font-weight: bold; }} .header-right {{ background: #cdffd8; font-weight: bold; }}
                .empty {{ background: #fafbfc; }}
            </style>
        </head>
        <body>
            <div style="font-weight: bold; margin-bottom: 10px;">{placement_name_cos} | {final_folder_name}</div>
            <div class="container"><div class="pane">{"".join(left_html_lines)}</div><div class="pane">{"".join(right_html_lines)}</div></div>
        </body>
        </html>
        """

        try:
            with open(html_path, "w", encoding="utf-8") as f:
                f.write(html_content)
            page = await browser_context.new_page()
            await page.goto(f"file://{os.path.abspath(html_path)}")
            await page.pdf(
                path=pdf_path, format="A3", landscape=True, print_background=True,
                margin={"top": "5mm", "bottom": "5mm", "left": "5mm", "right": "5mm"}
            )
            await page.close()
            os.remove(html_path)
            log(f"    [✓] {final_folder_name} created.")
        except Exception as e:
            log(f"    [✕] Error: {e}")
    log("\n")
    await format_page.close()

async def run_code_comparison(cos_links, tracking_rows, log, spreadsheet):
    if not tracking_rows or len(tracking_rows) < 1:
        log("No data to display.")
        return

    # log(f"Checking data: {len(tracking_rows)} rows received.") 
    
    headers = [h.strip() for h in tracking_rows[0]]
    
    # 1. Handle Footer Detection logic
    footer_row = []
    if len(tracking_rows) > 2:
        # Check if the last row actually looks like a footer
        potential_footer = tracking_rows[-1]
        has_footer_text = any("setting" in str(cell).lower() for cell in potential_footer)
        
        if has_footer_text:
            data_rows = tracking_rows[1:-1]
            footer_row = potential_footer
        else:
            # If the last row doesn't have "setting", treat it as a data row
            data_rows = tracking_rows[1:]
            footer_row = []
    else:
        # Minimal case: Header and 1 Data row
        data_rows = tracking_rows[1:]
        footer_row = [] 

    if not data_rows:
        log("No data rows found to process.")
        return
    # Identify which columns have the "*Add under advanced settings" footer
    advanced_columns = []
    footer_id_map = {} # Maps Header Name -> ID (1, 2, 3...)
    next_id = 1
    if footer_row:
        for col_idx, footer_cell in enumerate(footer_row):
            if "setting" in str(footer_cell).lower():
                advanced_columns.append(col_idx)
                header_name = headers[col_idx]
                footer_id_map[header_name] = next_id
                next_id += 1

    user_data_dir = os.path.join(os.getcwd(), "USER_SESSION")
    
    # Pre-get the worksheet and write the header once before starting the loop
    try:
        comparison_worksheet = get_sheet(spreadsheet.id, "Extracted Data")
        write_code_comparison_to_sheet(comparison_worksheet, headers, [], {},log, is_header_only=True)
    except Exception as e:
        log(f"Error initializing sheet: {e}")
        return
    
    all_row_final_results = []

    async with async_playwright() as p:

        # 1. INITIAL AUTHENTICATION CHECK
        auth_context = await p.chromium.launch_persistent_context(user_data_dir, headless=False, channel="msedge",)
        auth_page = auth_context.pages[0] if auth_context.pages else await auth_context.new_page()
        await auth_page.goto("https://admanager.google.com/")
        
        if "accounts.google.com" in auth_page.url:
            log("--- ACTION REQUIRED: Please log in to Google Ad Manager ---")
            while "accounts.google.com" in auth_page.url:
                await asyncio.sleep(2)
            log("Login successful.")
        await auth_context.close()

        # 2. ROW-BY-ROW VERIFICATION
        total_rows = len(data_rows)
        for idx, row in enumerate(data_rows, start=1):
            row_map = {headers[i]: str(row[i]).strip() for i in range(len(headers))}
            
            # GET COS LINK (1:1 mapping with placement rows)
            current_cos_url = cos_links[idx - 1] if idx - 1 < len(cos_links) else None

            # --- NEW: PARSE TCL NUMBER AND CREATE NESTED FOLDERS ---
            tcl_number = "UnknownTCL"
            placement_name_cos = f"Row_{idx}" # Default fallback
            base_results_dir = "Code Comparison Results"
            tcl_folder_path = os.path.join(base_results_dir, tcl_number)

            if current_cos_url:
                parsed_parts = extract_url_parts(current_cos_url)
                if parsed_parts:
                    tcl_number = parsed_parts.get("tcl_number", "UnknownTCL")
                    placement_name_cos = parsed_parts.get("placement_name", f"Row_{idx}")
                    tcl_folder_path = os.path.join(base_results_dir, tcl_number)
            
            # Create the nested directory structure
            os.makedirs(tcl_folder_path, exist_ok=True)
            
            def get_val(keys):
                for k in keys:
                    if k in row_map: return row_map[k]
                return None

            current_row_advanced_targets = {}
            for col_idx in advanced_columns:
                header_name = headers[col_idx]
                val = str(row[col_idx]).strip()
                if val and val != "None":
                    current_row_advanced_targets[header_name] = val

            placement_name = row_map.get("Placement Name", f"Row {idx}")
            
            row_results = {
                "lab_product": "N/A", 
                "ad_size_small": [], 
                "ad_size_large": [],
                "code_snippets": {},
                "cos_link_url": current_cos_url,
                "tcl_number": tcl_number,
                # New storage for the second table mapping
                "found_ids": {
                    "Small CID": [], "Large CID": [], 
                    "Small Test GAM": [], "Large Test GAM": [],
                    "CID": [], "Test GAM": []
                }
            }

            raw_checks = []

            # --- Case 1: Sized Columns (Small/Large) ---
            s_cid = get_val(["Small CID", "SmallCID"])
            s_gam = get_val(["Small Test GAM", "SmallTestGAM", "SmallTest GAM"])
            if s_cid: raw_checks.append((s_cid, row_map.get("Creative Name - Small"), "Small CID"))
            if s_gam: raw_checks.append((s_gam, row_map.get("Creative Name - Small"), "Small Test GAM"))
            
            l_cid = get_val(["Large CID", "LargeCID"])
            l_gam = get_val(["Large Test GAM", "LargeTestGAM", "LargeTest GAM"])
            if l_cid: raw_checks.append((l_cid, row_map.get("Creative Name - Large"), "Large CID"))
            if l_gam: raw_checks.append((l_gam, row_map.get("Creative Name - Large"), "Large Test GAM"))

            # --- Case 2: Generic Columns (No Size) ---
            g_cid = row_map.get("CID")
            g_gam = get_val(["Test GAM", "TestGAM"])
            if g_cid: raw_checks.append((g_cid, row_map.get("Creative Name"), "CID"))
            if g_gam: raw_checks.append((g_gam, row_map.get("Creative Name"), "Test GAM"))

            valid_checks = []
            current_row_data = list(row)
            for url, name, label in raw_checks:
                if url and name and url.startswith("http"):
                    processed_url = url
                    if "CID" in label:
                        if "tab=preview" in processed_url:
                            processed_url = processed_url.replace("tab=preview", "tab=settings")
                        elif "tab=" not in processed_url:
                            processed_url += "&tab=settings" if "#" in processed_url or "&" in processed_url else "#tab=settings"
                    valid_checks.append((processed_url, name, label))

            log(f"[{idx}/{total_rows}] {placement_name}")
            row_context = await p.chromium.launch_persistent_context(user_data_dir, headless=False, channel="msedge",)
            await row_context.grant_permissions(["notifications", "clipboard-read", "clipboard-write"])
            
            try:
                # --- SIMULATE CTRL+U -> SELECT ALL -> COPY FOR COS LINK ---
                if current_cos_url and current_cos_url.startswith("http"):
                    view_source_url = f"view-source:{current_cos_url}"
                    cos_page = await row_context.new_page()
                    log(f"  → Opening View-Source: {current_cos_url}")
                    try:
                        await cos_page.goto(view_source_url, timeout=45000)
                        await asyncio.sleep(2) # Ensure code loads
                        
                        # Simulate Keyboard Actions: Ctrl+A, Ctrl+C
                        await cos_page.keyboard.down("Control")
                        await cos_page.keyboard.press("a")
                        await cos_page.keyboard.press("c")
                        await cos_page.keyboard.up("Control")
                        
                        # Read from clipboard
                        cos_source = await cos_page.evaluate("navigator.clipboard.readText()")
                        row_results["code_snippets"]["COS Link"] = cos_source
                        log(f"    [✓] COS Link source code captured.")
                    except Exception as cos_e:
                        log(f"    [✕] Failed to capture COS Link source code: {cos_e}")
                    finally:
                        await cos_page.close()

                # --- PROCESS OTHER URLS (CID / GAM) ---
                for url, creative_name, label in valid_checks:
                    p_page = await row_context.new_page()
                    log(f"  → Opening {label}: Searching for '{creative_name}'")
                    
                    try:
                        await p_page.goto(url, timeout=60000)
                        await p_page.wait_for_load_state("networkidle")
                        await asyncio.sleep(2) 

                        page_title_element = p_page.locator('h1[debugid="page-title"], .page-title, h1')
                        
                        if await page_title_element.count() > 0:
                            raw_title = await page_title_element.first.inner_text()
                            actual_name_on_page = raw_title.replace("Creative:", "").strip()
                            
                            if actual_name_on_page == creative_name.strip():
                                log(f"    [✓] Creative Name match in {label}")
                            else:
                                log(f"    [✕] {label} name mismatch: Expected '{creative_name.strip()}'")
                                log(f"    [✕] Found '{actual_name_on_page}'")
                                target_header = ""
                                if "Small" in label: target_header = "Creative Name - Small"
                                elif "Large" in label: target_header = "Creative Name - Large"
                                elif label == "CID" or label == "Test GAM": target_header = "Creative Name"

                                if target_header in headers:
                                    header_idx = headers.index(target_header)
                                    current_row_data[header_idx] = "REJECTED"

                            # --- CODE EXTRACTION (CodeMirror) ---
                            try:
                                await p_page.bring_to_front()
                                await p_page.wait_for_selector("div.CodeMirror", timeout=10000)
                                await p_page.focus("div.CodeMirror")
                                await p_page.click("div.CodeMirror")
                                await p_page.keyboard.down("Control")
                                await p_page.keyboard.press("a") 
                                await p_page.keyboard.press("c") 
                                await p_page.keyboard.up("Control")

                                try:
                                    copied_clipboard = await p_page.evaluate("navigator.clipboard.readText()")
                                except:
                                    copied_clipboard = ""

                                code_text = await p_page.evaluate("""() => {
                                    const cmElem = document.querySelector("div.CodeMirror");
                                    if (!cmElem) return "";
                                    const cmInstance = cmElem.CodeMirror || (cmElem.nextSibling && cmElem.nextSibling.CodeMirror);
                                    if (cmInstance && typeof cmInstance.getValue === "function") {
                                        return cmInstance.getValue();
                                    }
                                    const lines = Array.from(cmElem.querySelectorAll(".CodeMirror-line"));
                                    return lines.map(l => l.innerText).join("\\n");
                                }""")
                                
                                final_code = copied_clipboard if copied_clipboard else code_text
                                row_results["code_snippets"][label] = final_code
                                
                            except Exception as code_e:
                                log(f"    [✕] Code extraction failed for {label}: {code_e}")

                            # --- AD SIZE EXTRACTION ---
                            ad_size_locator = p_page.locator('dynamic-component[debugid="read-only-element"] .read-only-content').first
                            if await ad_size_locator.count() > 0:
                                current_size = (await ad_size_locator.inner_text()).strip()
                                if "Small" in label:
                                    if current_size not in row_results["ad_size_small"]:
                                        row_results["ad_size_small"].append(current_size)
                                elif "Large" in label:
                                    if current_size not in row_results["ad_size_large"]:
                                        row_results["ad_size_large"].append(current_size)
                                # NEW: Handle Generic CID / Test GAM
                                elif label == "CID":
                                    row_results["ad_size_cid"] = current_size
                                elif label == "Test GAM":
                                    row_results["ad_size_test_gam"] = current_size
                                
                                log(f"    [✓] Ad Size Found: {current_size}")

                            # --- LAB PRODUCT & SETTINGS ---
                            is_test_gam = "Test GAM" in label
                            try:
                                if is_test_gam:
                                    settings_tab = p_page.locator('a.tab-button:has-text("Settings")')
                                    if await settings_tab.count() > 0:
                                        await settings_tab.first.click()
                                        await p_page.locator('h2', has_text="Settings").first.wait_for(state="visible", timeout=10000)
                                    else:
                                        await p_page.click('text="Settings"')

                                if "CID" in label:
                                    await p_page.locator('drx-form-field').first.wait_for(state="visible", timeout=7000)
                                    lab_field_locator = p_page.locator('drx-form-field').filter(has_text="Creative Lab Product").locator('.button-text').first
                                    try:
                                        await lab_field_locator.wait_for(state="visible", timeout=5000)
                                        current_lab = (await lab_field_locator.inner_text()).strip()
                                        
                                        # NEW: Store based on specific labels to avoid N/A
                                        if label == "CID":
                                            row_results["lab_product_generic"] = current_lab
                                        elif "Small" in label:
                                            row_results["lab_product_small"] = current_lab
                                        elif "Large" in label:
                                            row_results["lab_product_large"] = current_lab
                                            
                                        log(f"    [✓] Lab Product: {current_lab}")
                                    except:
                                        log(f"    [✕] Creative Lab Product field not found.")

                                if current_row_advanced_targets:
                                    advanced_toggle = p_page.locator('text="*Add under advanced settings"')
                                    if await advanced_toggle.count() > 0:
                                        await advanced_toggle.first.scroll_into_view_if_needed()
                                        await advanced_toggle.first.click()
                                        await p_page.locator('input, textarea').first.wait_for(state="attached", timeout=5000)

                                    all_page_text = await p_page.content()
                                    form_values = await p_page.evaluate("() => Array.from(document.querySelectorAll('input, textarea')).map(i => i.value).join(' ')")
                                    combined_search_area = all_page_text + " " + form_values

                                    for header_name, target_val in current_row_advanced_targets.items():
                                        if target_val in combined_search_area:
                                            found_id = str(footer_id_map[header_name])
                                            row_results["found_ids"][label].append(found_id)
                                            log(f"    [✓] Found {header_name} in {label}.")
                                        else:
                                            log(f"    [✕] NOT found {header_name} value in {label}.")
                            except Exception as tab_e:
                                log(f"    [WARNING] Settings tasks error: {tab_e}")
                        else:
                            log(f"    [ERROR] Could not find the Creative Name on page.")
                    except Exception as e:
                        log(f"    [ERROR] Failed to process {label}: {e}")
                    finally:
                        await p_page.close()
                
                # --- NEW: RUN CODE COMPARISON PDF GENERATION ---
                if row_results["code_snippets"]:
                    log(f"  \n→ Running side-by-side code comparison for {placement_name}")
                    await code_comparing(
                        tcl_number, 
                        placement_name_cos, 
                        row_results["code_snippets"], 
                        log, 
                        row_context
                    )

                # --- WRITE TO TABLE 1 (CODE COMPARISON) ---
                formatted_results = {
                    # Lab Products
                    "lab_product_generic": row_results.get("lab_product_generic", "N/A"),
                    "lab_product_small": row_results.get("lab_product_small", "N/A"),
                    "lab_product_large": row_results.get("lab_product_large", "N/A"),
                    
                    # Ad Sizes (Generic/CID Mode)
                    "ad_size_cid": row_results.get("ad_size_cid", "N/A"),
                    "ad_size_test_gam": row_results.get("ad_size_test_gam", "N/A"),
                    
                    # Ad Sizes (Sized Mode)
                    "ad_size_small": ", ".join(row_results["ad_size_small"]) if row_results.get("ad_size_small") else "N/A",
                    "ad_size_large": ", ".join(row_results["ad_size_large"]) if row_results.get("ad_size_large") else "N/A"
                }
                write_code_comparison_to_sheet(comparison_worksheet, headers, [current_row_data], {0: formatted_results}, log)
                
                # --- STORE FOR TABLE 2 (TRACKING VERIFIER) ---
                all_row_final_results.append({
                    "placement_name": placement_name,
                    "found_ids": row_results["found_ids"],
                    "row_map": row_map
                })
            
            finally:
                await row_context.close()
        if footer_row:
            write_code_comparison_to_sheet(comparison_worksheet, headers, [], {"footer_row": footer_row}, log)
    # --- FINAL STEP: WRITE THE ENTIRE TRACKING VERIFIER TABLE ---
    if all_row_final_results and footer_id_map:
        tracking_data = {
            "footer_id_map": footer_id_map,
            "results": all_row_final_results
        }
        write_code_comparison_to_sheet(comparison_worksheet, headers, [], tracking_data, log, is_tracking_table=True)
    else:
        log("\n[INFO] Advanced settings not found; skipping validation.")
