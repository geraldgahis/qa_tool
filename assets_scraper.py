#assets_scraper.py

import gspread
from playwright.async_api import async_playwright
from sheet_writer import write_assets_to_sheet, get_sheet
from utils.constants import WEATHER_PARAMS, WEATHER, DAY_NIGHT, SIZE_MAPPING
from utils.url_parser import extract_url_parts

async def run_assets_scraper(cos_links, log, spreadsheet):
    """
    cos_links: list of COS URLs
    weather_filter: function to filter network requests (req_url -> True/False)
    """
    worksheet = get_sheet(spreadsheet.id, "Extracted Data")
    total = len(cos_links)
    for idx, url in enumerate(cos_links, 1):
        parts = extract_url_parts(url)
        
        ad_type = parts.get("ad_type", "")
        placement_name = parts.get("placement_name", f"COS_{idx}")
        base_here = parts["base_url"]
        other_assets_prefix = parts["other_assets"]

        log(f"[{idx}/{total}] Getting assets: {url}")

        # Data structure for table
        weather_table = {}
        # BEFORE: row_label = f"{w_label} {dn_label}"
        for w_code, w_label in WEATHER.items():
            for dn_code, dn_label in DAY_NIGHT.items():
                row_label = f"{w_label} {dn_label}"  # now matches later row_label

                if ad_type.startswith("ma-"):
                    weather_table[row_label] = {
                        "BG": {"SMALL": {"urls": []}, "LARGE": {"urls": []}},
                        "FG": {"SMALL": {"urls": []}, "LARGE": {"urls": []}},
                        "textmode": "N/A",
                        "OTHER_ASSETS": {"SMALL": {"urls": []}, "LARGE": {"urls": []}},
                    }
                elif ad_type.startswith("mw-"):
                    weather_table[row_label] = {
                        "OPEN_BG": {"urls": []},
                        "CLOSED_BG": {"urls": []},
                        "OPEN_FG": {"urls": []},
                        "CLOSED_FG": {"urls": []},
                        "CLOSE_BTN": {"urls": []},
                        "EXPAND_BTN": {"urls": []},
                        "OTHER_ASSETS": [],
                    }

        
        async with async_playwright() as p:
            browser = await p.chromium.launch(headless=True, channel="msedge",)
            context = await browser.new_context()

            for weather_idx, weather_param in enumerate(WEATHER_PARAMS, 1):
                weather_url = f"{base_here}#{weather_param}"
                log(
                    f"→ Opening: {weather_url}"
                )

                page = await context.new_page()

                # Parse fragment for this tab
                fragment_parts = weather_param.split("+")
                weather_code = fragment_parts[1] if len(fragment_parts) > 1 else None
                daynight_code = fragment_parts[2] if len(fragment_parts) > 2 else None
                size_code = fragment_parts[3] if len(fragment_parts) > 3 else None

                # Lowercase for consistency
                weather_code_l = weather_code.lower() if weather_code else ""
                daynight_code_l = daynight_code.lower() if daynight_code else ""
                size_label = SIZE_MAPPING.get(size_code, None)

                row_label = f"{WEATHER.get(weather_code,'')} {DAY_NIGHT.get(daynight_code,'')}"

                matching_requests = []

                # Capture network requests
                def on_request(request):
                    req_url = request.url.lower()

                    # MAIM
                    if ad_type.startswith("ma-"):
                        # BG
                        if f"bg-{weather_code_l}-{daynight_code_l}" in req_url:
                            matching_requests.append(("BG", request.url))
                            log(f"      [{idx}/{total}][BG] {request.url}")
                        # FG
                        elif "-fg-" in req_url:
                            matching_requests.append(("FG", request.url))
                            log(f"      [{idx}/{total}][FG] {request.url}")
                        # OTHER ASSETS
                        elif other_assets_prefix.lower() in req_url:
                            if size_label and not req_url.endswith((".html", "js", "ts")):
                                if req_url not in weather_table[row_label]["OTHER_ASSETS"][size_label]["urls"]:
                                    weather_table[row_label]["OTHER_ASSETS"][size_label]["urls"].append(request.url)
                                    log(f"      [{idx}/{total}][OTHER ASSETS {size_label}] {request.url}")

                    # MWIM
                    elif ad_type.startswith("mw-"):
                        # OPEN BG
                        if f"open-bg-{weather_code_l}-{daynight_code_l}" in req_url:
                            if request.url not in weather_table[row_label]["OPEN_BG"]["urls"]:
                                weather_table[row_label]["OPEN_BG"]["urls"].append(request.url)
                            log(f"      [{idx}/{total}][OPEN BG] {request.url}")
                        # CLOSED BG
                        elif f"closed-bg-{weather_code_l}-{daynight_code_l}" in req_url:
                            if request.url not in weather_table[row_label]["CLOSED_BG"]["urls"]:
                                weather_table[row_label]["CLOSED_BG"]["urls"].append(request.url)
                            log(f"      [{idx}/{total}][CLOSED BG] {request.url}")
                        # OPEN FG
                        elif "open-fg" in req_url:
                            if request.url not in weather_table[row_label]["OPEN_FG"]["urls"]:
                                weather_table[row_label]["OPEN_FG"]["urls"].append(request.url)
                            log(f"      [{idx}/{total}][OPEN FG] {request.url}")
                        # CLOSED FG
                        elif "closed-fg" in req_url:
                            if request.url not in weather_table[row_label]["CLOSED_FG"]["urls"]:
                                weather_table[row_label]["CLOSED_FG"]["urls"].append(request.url)
                            log(f"      [{idx}/{total}][CLOSED FG] {request.url}")

                        # CLOSE BTN
                        elif "mw-close-btn-2x.png" in req_url:
                            if request.url not in weather_table[row_label]["CLOSE_BTN"]["urls"]:
                                weather_table[row_label]["CLOSE_BTN"]["urls"].append(request.url)
                            log(f"      [{idx}/{total}][CLOSE BTN] {request.url}")
                        # EXPAND BTN
                        elif "mw-expand-btn-2x.png" in req_url:
                            if request.url not in weather_table[row_label]["EXPAND_BTN"]["urls"]:
                                weather_table[row_label]["EXPAND_BTN"]["urls"].append(request.url)
                            log(f"      [{idx}/{total}][EXPAND BTN] {request.url}")
                        # OTHER ASSETS
                        elif other_assets_prefix.lower() in req_url:
                            if not req_url.endswith((".html", "js", "ts")) and req_url not in weather_table[row_label]["OTHER_ASSETS"]:
                                weather_table[row_label]["OTHER_ASSETS"].append(request.url)
                                log(f"      [{idx}/{total}][OTHER ASSETS] {request.url}")


                page.on("request", on_request)

                try:
                    await page.goto(weather_url, timeout=30000)
                    await page.wait_for_load_state("networkidle")

                    # Grab textMode value from page
                    if ad_type.startswith("ma-"):
                        try:
                            if daynight_code_l == "d":
                                textmode_value = await page.evaluate(
                                    "() => labAd?.selected?.bg?.textModeDay "
                                    "|| labAd?.selected?.bg?.textMode "
                                    "|| labAd?.selected?.textMode "
                                    "|| 'N/A'"
                                )
                            elif daynight_code_l == "n":
                                textmode_value = await page.evaluate(
                                    "() => labAd?.selected?.bg?.textModeNight "
                                    "|| labAd?.selected?.bg?.textMode "
                                    "|| labAd?.selected?.textMode "
                                    "|| 'N/A'"
                                )
                            else:
                                textmode_value = await page.evaluate(
                                    "() => labAd?.selected?.bg?.textMode "
                                    "|| labAd?.selected?.textMode "
                                    "|| 'N/A'"
                                )
                        except Exception:
                            textmode_value = "N/A"

                        log(
                            f"      [{idx}/{total}][TEXTMODE]: {textmode_value}"
                        )

                except Exception as e:
                    log(
                        f"      [{idx}/{total}]Failed ({weather_param}): {e}"
                    )
                    textmode_value = "N/A"

                # Store textmode
                if row_label in weather_table:
                    if ad_type.startswith("ma-") and size_label:
                        for asset_type, url in matching_requests:
                            weather_table[row_label][asset_type][size_label]["urls"].append(url)
                        weather_table[row_label]["textmode"] = textmode_value
                    elif ad_type.startswith("mw-"):
                        for asset_type, url in matching_requests:
                            if asset_type == "OPEN_BG":
                                if url not in weather_table[row_label]["OPEN_BG"]["urls"]:
                                    weather_table[row_label]["OPEN_BG"]["urls"].append(url)
                            else:
                                weather_table[row_label][asset_type][size_label]["urls"].append(url)

            # Close browser
            await browser.close()
        
        try:
            worksheet = spreadsheet.worksheet("Extracted Data")
        except gspread.WorksheetNotFound:
            worksheet = spreadsheet.add_worksheet("Extracted Data", rows=100, cols=20)


        # Write to sheet
        write_assets_to_sheet(worksheet, weather_table, placement_name, ad_type, log)
