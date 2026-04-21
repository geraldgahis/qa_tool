import asyncio
import os
import difflib

from playwright.async_api import async_playwright
from utils.url_parser import extract_url_parts
# Note: sheet_writer is no longer needed here since verify_cid_gam handles sheet updates!


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
        if not diff:
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
    
    headers = [h.strip() for h in tracking_rows[0]]
    
    # We only need the data rows (skipping headers and potential footer)
    data_rows = tracking_rows[1:]
    if len(tracking_rows) > 2 and any("setting" in str(cell).lower() for cell in tracking_rows[-1]):
        data_rows = tracking_rows[1:-1]

    if not data_rows:
        log("No data rows found to process.")
        return

    user_data_dir = os.path.join(os.getcwd(), "USER_SESSION")

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

        # 2. ROW-BY-ROW CODE EXTRACTION
        total_rows = len(data_rows)
        for idx, row in enumerate(data_rows, start=1):
            row_map = {headers[i]: str(row[i]).strip() for i in range(len(headers))}
            
            # GET COS LINK
            current_cos_url = cos_links[idx - 1] if idx - 1 < len(cos_links) else None

            # PARSE TCL NUMBER AND CREATE NESTED FOLDERS
            tcl_number = "UnknownTCL"
            placement_name_cos = f"Row_{idx}" 
            base_results_dir = "Code Comparison Results"
            tcl_folder_path = os.path.join(base_results_dir, tcl_number)

            if current_cos_url:
                parsed_parts = extract_url_parts(current_cos_url)
                if parsed_parts:
                    tcl_number = parsed_parts.get("tcl_number", "UnknownTCL")
                    placement_name_cos = parsed_parts.get("placement_name", f"Row_{idx}")
                    tcl_folder_path = os.path.join(base_results_dir, tcl_number)
            
            os.makedirs(tcl_folder_path, exist_ok=True)
            
            def get_val(keys):
                for k in keys:
                    if k in row_map: return row_map[k]
                return None

            placement_name = row_map.get("Placement Name", f"Row {idx}")
            
            row_results = {
                "code_snippets": {},
                "cos_link_url": current_cos_url,
                "tcl_number": tcl_number
            }

            # COMPREHENSIVE HEADER VALIDATION
            raw_checks = []
            
            s_cid = get_val(["Small CID", "SmallCID"])
            s_gam = get_val(["Small Test GAM", "SmallTest GAM", "SmallTestGAM", "Small TestGAM"])
            if s_cid: raw_checks.append((s_cid, row_map.get("Creative Name - Small"), "Small CID"))
            if s_gam: raw_checks.append((s_gam, row_map.get("Creative Name - Small"), "Small Test GAM"))
            
            l_cid = get_val(["Large CID", "LargeCID"])
            l_gam = get_val(["Large Test GAM", "LargeTest GAM", "LargeTestGAM", "Large TestGAM"])
            if l_cid: raw_checks.append((l_cid, row_map.get("Creative Name - Large"), "Large CID"))
            if l_gam: raw_checks.append((l_gam, row_map.get("Creative Name - Large"), "Large Test GAM"))

            g_cid = get_val(["CID"]) 
            g_gam = get_val(["Test GAM", "TestGAM"])
            if g_cid: raw_checks.append((g_cid, row_map.get("Creative Name"), "CID"))
            if g_gam: raw_checks.append((g_gam, row_map.get("Creative Name"), "Test GAM"))

            valid_checks = []
            for url, name, label in raw_checks:
                if url and name and url.startswith("http"):
                    processed_url = url
                    if "CID" in label:
                        if "tab=preview" in processed_url:
                            processed_url = processed_url.replace("tab=preview", "tab=settings")
                        elif "tab=" not in processed_url:
                            processed_url += "&tab=settings" if "#" in processed_url or "&" in processed_url else "#tab=settings"
                    valid_checks.append((processed_url, name, label))

            log(f"[{idx}/{total_rows}] Code Comparison for: {placement_name}")
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
                        await asyncio.sleep(2) 
                        
                        await cos_page.keyboard.down("Control")
                        await cos_page.keyboard.press("a")
                        await cos_page.keyboard.press("c")
                        await cos_page.keyboard.up("Control")
                        
                        cos_source = await cos_page.evaluate("navigator.clipboard.readText()")
                        row_results["code_snippets"]["COS Link"] = cos_source
                        log(f"    [✓] COS Link source code captured.")
                    except Exception as cos_e:
                        log(f"    [✕] Failed to capture COS Link source code: {cos_e}")
                    finally:
                        await cos_page.close()

                # --- PROCESS OTHER URLS (Code Extraction Only) ---
                async def pre_load_tab(target_url, c_name, t_label):
                    page = await row_context.new_page()
                    log(f"  → Concurrently loading {t_label} in background...")
                    try:
                        await page.goto(target_url, timeout=60000)
                    except Exception:
                        pass 
                    return (page, target_url, c_name, t_label)

                load_tasks = [pre_load_tab(url, creative_name, label) for url, creative_name, label in valid_checks]
                loaded_tabs = await asyncio.gather(*load_tasks)

                for p_page, url, creative_name, label in loaded_tabs:
                    log(f"  → Extracting code from {label}...")
                    
                    max_retries = 3
                    for attempt in range(1, max_retries + 1):
                        try:
                            await p_page.bring_to_front()
                            
                            if attempt > 1:
                                log(f"    [↻] Refreshing {label} (Attempt {attempt}/{max_retries})...")
                                await p_page.reload(timeout=60000)
                            
                            await p_page.wait_for_load_state("networkidle", timeout=30000)
                            await asyncio.sleep(2) 

                            # --- CODE EXTRACTION (CodeMirror) ---
                            await p_page.wait_for_selector("div.CodeMirror", timeout=15000)
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
                            log(f"    [✓] Source code extracted successfully.")
                            
                            # Success! Break retry loop
                            break 
                            
                        except Exception as e:
                            if "Timeout" in type(e).__name__:
                                log(f"    [✕] Timeout exceeded on {label} (Attempt {attempt}).")
                            else:
                                log(f"    [✕] Error processing {label} (Attempt {attempt}): {e}")
                                
                            if attempt == max_retries:
                                log(f"    [ERROR] Max retries reached for {label}. Skipping extraction.")
                    
                    await p_page.close()
                
                # --- RUN CODE COMPARISON PDF GENERATION ---
                if row_results["code_snippets"]:
                    log(f"  \n→ Generating side-by-side Diff PDFs for {placement_name}")
                    await code_comparing(
                        tcl_number, 
                        placement_name_cos, 
                        row_results["code_snippets"], 
                        log, 
                        row_context
                    )

            finally:
                await row_context.close()