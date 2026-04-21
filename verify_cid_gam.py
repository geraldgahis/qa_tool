import asyncio
import os
from playwright.async_api import async_playwright
from utils.url_parser import extract_url_parts
from sheet_writer import get_sheet, write_code_comparison_to_sheet

async def run_verification(cos_links, tracking_rows, log, spreadsheet):
    if not tracking_rows or len(tracking_rows) < 1:
        log("No data to display.")
        return

    headers = [h.strip() for h in tracking_rows[0]]
    
    # Handle Footer Detection logic
    footer_row = []
    if len(tracking_rows) > 2:
        potential_footer = tracking_rows[-1]
        if any("setting" in str(cell).lower() for cell in potential_footer):
            data_rows = tracking_rows[1:-1]
            footer_row = potential_footer
        else:
            data_rows = tracking_rows[1:]
    else:
        data_rows = tracking_rows[1:]

    if not data_rows:
        log("No data rows found to process.")
        return

    advanced_columns = []
    footer_id_map = {}
    next_id = 1
    if footer_row:
        for col_idx, footer_cell in enumerate(footer_row):
            if "setting" in str(footer_cell).lower():
                advanced_columns.append(col_idx)
                header_name = headers[col_idx]
                footer_id_map[header_name] = next_id
                next_id += 1

    user_data_dir = os.path.join(os.getcwd(), "USER_SESSION")
    
    try:
        comparison_worksheet = get_sheet(spreadsheet.id, "Extracted Data")
        write_code_comparison_to_sheet(comparison_worksheet, headers, [], {}, log, is_header_only=True)
    except Exception as e:
        log(f"Error initializing sheet: {e}")
        return
    
    all_row_final_results = []

    async with async_playwright() as p:
        auth_context = await p.chromium.launch_persistent_context(user_data_dir, headless=False, channel="msedge")
        auth_page = auth_context.pages[0] if auth_context.pages else await auth_context.new_page()
        await auth_page.goto("https://admanager.google.com/")
        
        if "accounts.google.com" in auth_page.url:
            log("--- ACTION REQUIRED: Please log in to Google Ad Manager ---")
            while "accounts.google.com" in auth_page.url:
                await asyncio.sleep(2)
            log("Login successful.")
        await auth_context.close()

        total_rows = len(data_rows)
        for idx, row in enumerate(data_rows, start=1):
            row_map = {headers[i]: str(row[i]).strip() for i in range(len(headers))}
            current_cos_url = cos_links[idx - 1] if idx - 1 < len(cos_links) else None

            placement_name = row_map.get("Placement Name", f"Row {idx}")
            
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

            row_results = {
                "lab_product_generic": "N/A", "lab_product_small": "N/A", "lab_product_large": "N/A",
                "ad_size_cid": "N/A", "ad_size_test_gam": "N/A",
                "ad_size_small": [], "ad_size_large": [],
                "sf_cid": "N/A", "sf_gam": "N/A",
                "sf_small_cid": "N/A", "sf_large_cid": "N/A", 
                "sf_small_gam": "N/A", "sf_large_gam": "N/A",
                "found_ids": { "Small CID": [], "Large CID": [], "Small Test GAM": [], "Large Test GAM": [], "CID": [], "Test GAM": [] }
            }

            # --- COMPREHENSIVE HEADER VALIDATION ---
            raw_checks = []
            
            # 1. Small Sizes
            s_cid = get_val(["Small CID", "SmallCID"])
            s_gam = get_val(["Small Test GAM", "SmallTest GAM", "SmallTestGAM", "Small TestGAM"])
            
            if s_cid: raw_checks.append((s_cid, row_map.get("Creative Name - Small"), "Small CID"))
            if s_gam: raw_checks.append((s_gam, row_map.get("Creative Name - Small"), "Small Test GAM"))
            
            # 2. Large Sizes
            l_cid = get_val(["Large CID", "LargeCID"])
            l_gam = get_val(["Large Test GAM", "LargeTest GAM", "LargeTestGAM", "Large TestGAM"])
            
            if l_cid: raw_checks.append((l_cid, row_map.get("Creative Name - Large"), "Large CID"))
            if l_gam: raw_checks.append((l_gam, row_map.get("Creative Name - Large"), "Large Test GAM"))

            # 3. Generic Sizes (No 'Small' or 'Large' prefix)
            g_cid = get_val(["CID"]) 
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

            log(f"\n[{idx}/{total_rows}] Verifying Data for: {placement_name}")
            row_context = await p.chromium.launch_persistent_context(user_data_dir, headless=False, channel="msedge")
            
            try:
                async def pre_load_tab(target_url, c_name, t_label):
                    page = await row_context.new_page()
                    try: await page.goto(target_url, timeout=60000)
                    except: pass 
                    return (page, target_url, c_name, t_label)

                load_tasks = [pre_load_tab(url, creative_name, label) for url, creative_name, label in valid_checks]
                loaded_tabs = await asyncio.gather(*load_tasks)

                for p_page, url, creative_name, label in loaded_tabs:
                    max_retries = 3
                    for attempt in range(1, max_retries + 1):
                        try:
                            await p_page.bring_to_front()
                            if attempt > 1: await p_page.reload(timeout=60000)
                            
                            await p_page.wait_for_load_state("networkidle", timeout=30000)
                            await asyncio.sleep(2) 

                            # --- TITLE VALIDATION ---
                            page_title_element = p_page.locator('h1[debugid="page-title"], .page-title, h1')
                            if await page_title_element.count() > 0:
                                raw_title = await page_title_element.first.inner_text()
                                actual_name_on_page = raw_title.replace("Creative:", "").strip()
                                
                                if actual_name_on_page == creative_name.strip():
                                    log(f"    [✓] Creative Name match in {label}")
                                else:
                                    log(f"    [✕] {label} mismatch: Expected '{creative_name.strip()}', Found '{actual_name_on_page}'")
                                    target_header = "Creative Name - Small" if "Small" in label else "Creative Name - Large" if "Large" in label else "Creative Name"
                                    if target_header in headers:
                                        current_row_data[headers.index(target_header)] = "REJECTED"
                            else:
                                raise Exception("Creative Name title element not found.")

                            # --- AD SIZE EXTRACTION ---
                            ad_size_locator = p_page.locator('dynamic-component[debugid="read-only-element"] .read-only-content').first
                            if await ad_size_locator.count() > 0:
                                current_size = (await ad_size_locator.inner_text()).strip()
                                if "Small" in label and current_size not in row_results["ad_size_small"]: row_results["ad_size_small"].append(current_size)
                                elif "Large" in label and current_size not in row_results["ad_size_large"]: row_results["ad_size_large"].append(current_size)
                                elif label == "CID": row_results["ad_size_cid"] = current_size
                                elif label == "Test GAM": row_results["ad_size_test_gam"] = current_size
                                log(f"    [✓] Ad Size Found: {current_size}")

                            # --- LAB PRODUCT & SETTINGS NAV ---
                            is_test_gam = "Test GAM" in label
                            if is_test_gam:
                                settings_tab = p_page.locator('a.tab-button:has-text("Settings")')
                                if await settings_tab.count() > 0:
                                    await settings_tab.first.click()
                                    await p_page.locator('h2', has_text="Settings").first.wait_for(state="visible", timeout=10000)
                                else:
                                    await p_page.click('text="Settings"')

                            # --- LAB PRODUCT FIELD ---
                            if "CID" in label:
                                await p_page.locator('drx-form-field').first.wait_for(state="visible", timeout=7000)
                                lab_field_locator = p_page.locator('drx-form-field').filter(has_text="Creative Lab Product").locator('.button-text').first
                                try:
                                    await lab_field_locator.wait_for(state="visible", timeout=5000)
                                    current_lab = (await lab_field_locator.inner_text()).strip()
                                    if label == "CID": row_results["lab_product_generic"] = current_lab
                                    elif "Small" in label: row_results["lab_product_small"] = current_lab
                                    elif "Large" in label: row_results["lab_product_large"] = current_lab
                                    log(f"    [✓] Lab Product: {current_lab}")
                                except:
                                    log(f"    [✕] Creative Lab Product field not found.")

                            # --- SAFEFRAME VERIFICATION ---
                            try:
                                sf_locator = p_page.locator('material-checkbox').filter(has_text="Serve into a SafeFrame").first
                                await sf_locator.wait_for(state="attached", timeout=5000)
                                is_checked = await sf_locator.get_attribute("aria-checked")
                                
                                sf_status = "FAILED" if is_checked == "true" else "PASSED"
                                
                                if label == "Small CID": row_results["sf_small_cid"] = sf_status
                                elif label == "Large CID": row_results["sf_large_cid"] = sf_status
                                elif label == "Small Test GAM": row_results["sf_small_gam"] = sf_status
                                elif label == "Large Test GAM": row_results["sf_large_gam"] = sf_status
                                elif label == "CID": row_results["sf_cid"] = sf_status
                                elif label == "Test GAM": row_results["sf_gam"] = sf_status
                                
                                log(f"    [{'✕' if sf_status == 'FAILED' else '✓'}] SafeFrame ({label}): {sf_status}")
                            except Exception as sf_e:
                                log(f"    [✕] SafeFrame checkbox not found for {label}.")

                            # --- ADVANCED SETTINGS TARGETS ---
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
                                        row_results["found_ids"][label].append(str(footer_id_map[header_name]))
                                        log(f"    [✓] Found {header_name} in {label}.")

                            break # Success, break retry loop
                        except Exception as e:
                            if attempt == max_retries: log(f"    [ERROR] Max retries reached for {label}.")
                    
                    await p_page.close()

                # Map all the extracted data back
                formatted_results = {
                    "lab_product_generic": row_results["lab_product_generic"],
                    "lab_product_small": row_results["lab_product_small"],
                    "lab_product_large": row_results["lab_product_large"],
                    "ad_size_cid": row_results["ad_size_cid"],
                    "ad_size_test_gam": row_results["ad_size_test_gam"],
                    "ad_size_small": ", ".join(row_results["ad_size_small"]) if row_results["ad_size_small"] else "N/A",
                    "ad_size_large": ", ".join(row_results["ad_size_large"]) if row_results["ad_size_large"] else "N/A",
                    "sf_cid": row_results["sf_cid"],
                    "sf_gam": row_results["sf_gam"],
                    "sf_small_cid": row_results["sf_small_cid"],
                    "sf_large_cid": row_results["sf_large_cid"],
                    "sf_small_gam": row_results["sf_small_gam"],
                    "sf_large_gam": row_results["sf_large_gam"]
                }
                write_code_comparison_to_sheet(comparison_worksheet, headers, [current_row_data], {0: formatted_results}, log)
                
                all_row_final_results.append({
                    "placement_name": placement_name,
                    "found_ids": row_results["found_ids"],
                    "row_map": row_map
                })
            
            finally:
                await row_context.close()
                
        if footer_row: write_code_comparison_to_sheet(comparison_worksheet, headers, [], {"footer_row": footer_row}, log)
    
    if all_row_final_results and footer_id_map:
        tracking_data = { "footer_id_map": footer_id_map, "results": all_row_final_results }
        write_code_comparison_to_sheet(comparison_worksheet, headers, [], tracking_data, log, is_tracking_table=True)