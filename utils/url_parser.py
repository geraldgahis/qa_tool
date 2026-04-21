# utils/url_parser.py
import os
from urllib.parse import urlparse, urlunparse
from utils.constants import WEATHER, DAY_NIGHT, SIZE_MAPPING, WEATHER_ICONS

def extract_url_parts(url: str) -> dict | None:
    try:
        parsed = urlparse(url)
        parts = parsed.path.strip("/").split("/")
        if len(parts) < 4:
            return None

        dir_path = f"{parsed.scheme}://{parsed.netloc}/cl"
        dir_client = parts[1]

        # tcl_number and ad_type
        tcl_number, ad_type = (parts[2].split("-", 1) + [""])[:2]

        # Base placement_name (folder)
        placement_name = parts[3]

        # If last part (file) is not the same as folder, append it before .html
        if len(parts) > 4:
            file_name = parts[4]
            if file_name.endswith(".html"):
                file_base, _ = os.path.splitext(file_name)
                # append to placement_name if different
                if file_base not in placement_name:
                    placement_name = f"{placement_name}-{file_base}"
        base_url = urlunparse(parsed._replace(fragment=""))

        # Fragment: #adstest+weather+daynight+size
        frag = parsed.fragment.split("+")
        weather_code, daynight_code, size_code = (frag[1:4] + [None]*3)[:3]

        weather_filter = f"{weather_code}-{daynight_code}".lower() if weather_code and daynight_code else None
        weather_label = WEATHER.get(weather_code, "")
        daynight_label = DAY_NIGHT.get(daynight_code, "")
        size_label = SIZE_MAPPING.get(size_code, "")
        weather_icon = WEATHER_ICONS.get(weather_label, "")

        return {
            "dir": dir_path,
            "dirClient": dir_client,
            "tcl_number": tcl_number,
            "ad_type": ad_type,
            "placement_name": placement_name,
            "base_url": base_url,
            "weather_code": weather_code,
            "daynight_code": daynight_code,
            "size_code": size_code,
            "weather_filter": weather_filter,
            "weather_label": weather_label,
            "daynight_label": daynight_label,
            "size_label": size_label,
            "weather_icon": weather_icon,
            "other_assets": f"{dir_path}/{dir_client}/{tcl_number}-{ad_type}/",
        }
    except Exception as e:
        print(f"[URL PARSE ERROR] {e}")
        return None
