#constants.py
WEATHER_PARAMS = [
    "adstest+clr+D+sm+3+o", "adstest+clr+N+sm+3+o",
    "adstest+cld+D+sm+3+o", "adstest+cld+N+sm+3+o",
    "adstest+rain+D+sm+3+o", "adstest+rain+N+sm+3+o",
    "adstest+snow+D+sm+3+o", "adstest+snow+N+sm+3+o",
    "adstest+clr+D+lg+3+o", "adstest+clr+N+lg+3+o",
    "adstest+cld+D+lg+3+o", "adstest+cld+N+lg+3+o",
    "adstest+rain+D+lg+3+o", "adstest+rain+N+lg+3+o",
    "adstest+snow+D+lg+3+o", "adstest+snow+N+lg+3+o",
]
WEATHER_ICONS = {
    "Clear": "☀️",
    "Cloudy": "☁️",
    "Rainy": "🌧️",
    "Wintry": "❄️"
}
WEATHER = {"clr": "Clear", "cld": "Cloudy", "rain": "Rainy", "snow": "Wintry"}
SIZE_MAPPING = {"sm": "SMALL", "lg": "LARGE"}
DAY_NIGHT = {"D": "Day", "N": "Night"}

SORT_ORDER = [
    "Clear Day",
    "Clear Night",
    "Cloudy Day",
    "Cloudy Night",
    "Rainy Day",
    "Rainy Night",
    "Wintry Day",
    "Wintry Night"
]