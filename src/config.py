"""
Centralized configuration for Fuel Price Tracker.
All magic numbers, URLs, and thresholds in one place.
"""
from pathlib import Path

# ── Paths ──
BASE_DIR = Path(__file__).parent.parent
EXCEL_PATH = BASE_DIR / "fuel_tracker.xlsx"
JSON_PATH = BASE_DIR / "latest_results.json"

# ── HTTP ──
USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 Chrome/131.0.0.0 Safari/537.36"
)
HEADERS = {
    "User-Agent": USER_AGENT,
    "Accept": "text/html,application/xhtml+xml,*/*;q=0.8",
    "Accept-Language": "en-US,en;q=0.9",
}
REQUEST_TIMEOUT = 20  # seconds
FX_TIMEOUT = 30

# ── Retry settings ──
MAX_RETRIES = 3
RETRY_BACKOFF_FACTOR = 1.0  # seconds: 1, 2, 4...

# ── FX rate validation ──
FX_PLN_EUR_MIN = 3.5
FX_PLN_EUR_MAX = 6.0
FX_SEK_EUR_MIN = 9.0
FX_SEK_EUR_MAX = 13.0

# ── Price validation ranges ──
ORLEN_PL_MIN = 3000   # PLN/m³
ORLEN_PL_MAX = 10000  # PLN/m³
ORLEN_LT_MIN = 1000   # EUR/1000l (before /1000 conversion)
ORLEN_LT_MAX = 2500   # EUR/1000l
DIESEL_EUR_MIN = 0.5   # EUR/l
DIESEL_EUR_MAX = 3.5    # EUR/l
BSH_SE_MIN = 10        # SEK/l
BSH_SE_MAX = 35        # SEK/l

# ── Outlier detection: max % change from previous day ──
MAX_DAILY_CHANGE_PCT = 15.0

# ── Source URLs ──
URLS = {
    "fx_frankfurter": "https://api.frankfurter.app/latest?from=EUR&to=PLN,SEK",
    "fx_exchangerate": "https://api.exchangerate.host/latest?base=EUR&symbols=PLN,SEK",
    "fx_ecb_xml": "https://www.ecb.europa.eu/stats/eurofxref/eurofxref-daily.xml",
    "orlen_pl": "https://www.petrodom.pl/en/current-wholesale-fuel-prices-provided-by-pkn-orlen/",
    "orlen_lt_list": "https://www.orlenlietuva.lt/lt/wholesale/_layouts/f2hPriceTable/default.aspx",
    "orlen_lt_base": "https://www.orlenlietuva.lt",
    "elvis_de": "https://www.fuel-prices.eu/cheapest/",
    "bsh_se": "https://st1.se/foretag/listpris",
    "eu_bulletin_xlsx": (
        "https://energy.ec.europa.eu/document/download/"
        "264c2d0f-f161-4ea3-a777-78faae59bea0_en?"
        "filename=Weekly%20Oil%20Bulletin%20Weekly%20prices%20with%20Taxes%20-%202024-02-19.xlsx"
    ),
}

# ── EU Bulletin country mappings ──
EU_COUNTRIES = {"LT": None, "LV": None, "EE": None, "DK": None, "SE": None, "FI": None}
EU_COUNTRY_NAMES = {
    "Lithuania": "LT", "Latvia": "LV", "Estonia": "EE",
    "Denmark": "DK", "Sweden": "SE", "Finland": "FI",
}
EU_WEEKLY_COLUMNS = {"LT": 2, "LV": 3, "EE": 4, "DK": 5, "SE": 6, "FI": 7, "EU_AVG": 8}

# ── Excel formatting ──
EXCEL_FONT_FAMILY = "Aptos"
