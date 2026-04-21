# tracking_scraper.py
import asyncio
import re
from urllib.parse import quote, urlparse
from playwright.async_api import async_playwright
from sheet_writer import get_sheet, write_trackings_to_sheet


def strip_tracking_urls(tracking_header, tracking_rows, log):

    # ADD THIS GUARD
    if not tracking_header:
        log("[ERROR] No tracking headers found. Check if 'Placement Name', 'Weather/Creative' exists in the sheet.")
        return []
    
    skip_cols = {
        i for i, col in enumerate(tracking_header)
        if col.strip() in ("Placement Name", "Weather/Creative", "ClickTag", "mraid.js")
    }

    stripped_rows = []

    for row in tracking_rows:
        new_row = []

        for i, val in enumerate(row):
            if i in skip_cols:
                new_row.append(val)
                continue

            base_urls = set()

            if isinstance(val, str):
                script_match = re.findall(
                    r'<script[^>]+src=["\']([^"\']+)["\']',
                    val,
                    flags=re.IGNORECASE
                )
                urls_to_parse = script_match if script_match else re.split(r'[;,]\s*', val)

                for u in urls_to_parse:
                    if u.startswith(("http://", "https://")):
                        parsed = urlparse(u)
                        if parsed.scheme and parsed.netloc:
                            base_url = f"{parsed.scheme}://{parsed.netloc}/"
                            if base_url not in base_urls:
                                # log(f"[STRIPPED URL] {tracking_header[i].strip()} -> {base_url}")
                                base_urls.add(base_url)

            new_row.append(", ".join(sorted(base_urls)) if base_urls else val)

        stripped_rows.append(new_row)

    return stripped_rows



async def run_tracking_scraper(cos_links, tracking_rows, tracking_header, log, spreadsheet):

    worksheet = get_sheet(spreadsheet.id, "Extracted Data")
    stripped_tracking_rows = strip_tracking_urls(tracking_header, tracking_rows, log)
    
    final_rows = []
    header_list = [h.strip() for h in tracking_header]

    # --- INITIALIZE HEADER IN SHEET ---
    try:
        write_trackings_to_sheet(worksheet, tracking_header, [], log)
    except Exception as e:
        log(f"Error initializing tracking header: {e}")

    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True, channel="msedge")
        
        for idx, (cos_url, original_row) in enumerate(zip(cos_links, stripped_tracking_rows)):
            log(f"[{idx + 1}/{len(tracking_rows)}] {original_row[0]}")
            log(f"  → Opening {cos_url}")

            while len(original_row) < len(header_list):
                original_row.append("")

            context = await browser.new_context()
            page = await context.new_page()

            # --- FIX: Initialize current_row_data thoughtfully ---
            current_row_data = {}
            filters_dict = {}
            
            for i, header in enumerate(header_list):
                # Keep fixed data as-is
                if header in ("Placement Name", "Weather/Creative", "ClickTag", "mraid.js"):
                    current_row_data[header] = original_row[i]
                else:
                    # For tracking columns: initialize as empty so we don't write the filter/base URL
                    current_row_data[header] = "" 
                    
                    # Store the base URLs (filters) for the network listener
                    # Use the actual data from the stripped row for filtering
                    filters = original_row[i].split(", ") if isinstance(original_row[i], str) and original_row[i] else []
                    filters_dict[header] = filters

            # Network listener focused ONLY on this row
            def on_request(request):
                url = request.url
                for header, filters in filters_dict.items():
                    for f in filters:
                        # Only update if the filter matches AND we haven't already caught a URL for this header
                        if f and f in url and not current_row_data[header]:
                            current_row_data[header] = url
                            log(f"    [✓] Found {header}")
                            break
                        
            page.on("request", on_request)

            try:
                await page.goto(cos_url, timeout=30000)
                await page.wait_for_load_state("networkidle")
                await asyncio.sleep(2) 

                # --- DV Tag Update Logic ---
                # Fixed check to only run if we haven't found a DV Tag or if the current one is just a base URL
                if "DV Tag" in current_row_data:
                    dv_source_value = await page.evaluate("""
                        () => {
                            const scripts = Array.from(document.querySelectorAll('script'));
                            const dvScript = scripts.find(s => 
                                (s.src && s.src.includes('doubleverify')) || 
                                (s.innerHTML && s.innerHTML.includes('doubleverify'))
                            );
                            if (dvScript) {
                                return dvScript.src ? dvScript.src : dvScript.innerHTML;
                            }
                            return null;
                        }
                    """)
                    if dv_source_value:
                        current_row_data["DV Tag"] = dv_source_value

                    # --- mraid.js Logic ---
                    # Get the raw HTML content to see commented-out code
                    html_content = await page.content()

                    # Regex to find the mraid.js script, whether commented or not
                    mraid_regex = r'<!--.*?<script[^>]+src=["\']mraid\.js["\'][^>]*>.*?</script>.*?-->|<script[^>]+src=["\']mraid\.js["\'][^>]*>.*?</script>'
                    mraid_match = re.search(mraid_regex, html_content, re.IGNORECASE | re.DOTALL)

                    if mraid_match:
                        if "<!--" in mraid_match.group(0):
                            # log("  [CHECK] mraid.js found (Commented)")
                            current_row_data["mraid.js"] = "Commented"
                            # Do NOT add header if commented
                        else:
                            current_row_data["mraid.js"] = '<script src="mraid.js"></script>'
                            # log("  [CHECK] mraid.js found (Uncommented)")
                            # Add header to both lists if not already present
                            if "mraid.js" not in header_list:
                                header_list.append("mraid.js")
                            if "mraid.js" not in tracking_header:
                                tracking_header.append("mraid.js")
                    # Do nothing if mraid.js not found

                # Capture clickTag specifically for this page
                try:
                    clicktag = await page.evaluate(
                        """() => {
                            if (window.labAd?.tracking?.clickTag) return window.labAd.tracking.clickTag;
                            if (window.globalData?.tracking?.clickTag) return window.globalData.tracking.clickTag;
                            return null;
                        }"""
                    )
                    if clicktag:
                        def encode_braces(match):
                            return "$" + quote("{" + match.group(1) + "}", safe='')
                        current_row_data["ClickTag"] = re.sub(r'\$\{([^}]+)\}', encode_braces, clicktag)
                        log(f"    [✓] Found ClickTag")
                    else:
                        current_row_data["ClickTag"] = "N/A"
                except Exception:
                    current_row_data["ClickTag"] = "N/A"

            except Exception as e:
                log(f"  [ERROR] Page failed: {e}")

            await context.close()

            # --- WRITE CURRENT ROW IMMEDIATELY ---
            # Now only contains network URLs or empty strings, no base-URL filters
            row_to_append = [current_row_data.get(h, "") for h in header_list]
            
            write_trackings_to_sheet(
                worksheet=worksheet,
                tracking_header=tracking_header,
                tracking_rows=[row_to_append],
                log=log,
                is_row_only=True
            )
            
            final_rows.append(row_to_append)
