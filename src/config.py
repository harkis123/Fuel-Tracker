"""
Centralized configuration for Fuel Price Tracker.
All magic numbers, URLs, and thresholds in one place.

NOTE: scraper.py imports from this module — keep names stable.
"""
import os
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

# ── Outlier detection: max % change from previous value ──
# Applied per-source in scraper via validate_price_change(); a larger jump
# is written but flagged SUSPECT in the Notes column instead of silently passing.
MAX_DAILY_CHANGE_PCT = 15.0

# ══════════════════════════════════════════════════════════════════
# ORLEN LT — PDF column / product selection
# ══════════════════════════════════════════════════════════════════
# The Orlen Lietuva "realizacija" PDF has 5 numeric columns per product:
#   1 Bazinė pardavimo kaina | 2 Akcizas | 3 Bazė+akcizas (be PVM)
#   4 PVM 21% | 5 Pardavimo kaina su PVM
# We parse by column position, NOT by "max number in line".
ORLEN_LT_COL_BASE = 0        # bazinė pardavimo kaina
ORLEN_LT_COL_EXCISE = 1      # akcizas
ORLEN_LT_COL_NET = 2         # bazė + akcizas, BE PVM
ORLEN_LT_COL_VAT = 3         # PVM 21%
ORLEN_LT_COL_GROSS = 4       # pardavimo kaina SU PVM

# Which price the tracker stores. The dashboard MECFS + PL VAT-parity are built
# on SU PVM (gross), so default = gross. Flip to ORLEN_LT_COL_NET for be-PVM.
ORLEN_LT_PRICE_COL = ORLEN_LT_COL_GROSS

# The product line we want = road diesel "Dyzelinas ... su RRME".
# MUST contain these, and MUST NOT contain any exclude term (agri / marine /
# heating oil), which would otherwise silently match a much cheaper product.
ORLEN_LT_REQUIRE = ["Dyzelinas", "RRME"]
ORLEN_LT_EXCLUDE = ["žemės ūkiui", "zemes ukiui", "laivų", "laivu",
                    "krosnių", "krosniu", "šild", "sild", "heat"]
# The PDF lists the same product for several terminals. Index 0 = first listed
# (Mažeikiai / Orlen Lietuva refinery). Set to None to AVERAGE all terminals.
ORLEN_LT_TERMINAL_INDEX = 0

# Seasonal EN 590 diesel class hint (used only to disambiguate when several
# diesel grades are present). Vasara C, pereinamasis E, žiema 0/1/2.
ORLEN_LT_SEASON_CLASSES = {
    (5, 6, 7, 8, 9): ["C"],
    (4, 10): ["E"],
    (11, 12, 1, 2, 3): ["2", "1", "0", "E"],
}

# ══════════════════════════════════════════════════════════════════
# ELVIS DE — German diesel reference
# ══════════════════════════════════════════════════════════════════
# The real ELVIS Dieselfloater is partner-only (BLUE.net, not public). As a
# public proxy for the German diesel price we use Tankerkönig (official MTS-K
# pump-price data, free) when an API key is provided, else the EC Oil Bulletin
# Germany diesel column (no key, weekly, already fetched anyway).
ELVIS_DE_SOURCE = "tankerkoenig"  # "tankerkoenig" | "ec_bulletin"
TANKERKOENIG_API_KEY = os.environ.get("TANKERKOENIG_API_KEY", "").strip()
# Basket of major-city coordinates → national diesel average (lat, lng).
TANKERKOENIG_CITIES = [
    ("Berlin", 52.5200, 13.4050),
    ("München", 48.1351, 11.5820),
    ("Hamburg", 53.5511, 9.9937),
    ("Köln", 50.9375, 6.9603),
    ("Frankfurt", 50.1109, 8.6821),
]
TANKERKOENIG_RADIUS_KM = 5   # max allowed by API
TANKERKOENIG_URL = "https://creativecommons.tankerkoenig.de/json/list.php"

# ── Source URLs ──
URLS = {
    "fx_frankfurter": "https://api.frankfurter.app/latest?from=EUR&to=PLN,SEK",
    "fx_exchangerate": "https://api.exchangerate.host/latest?base=EUR&symbols=PLN,SEK",
    "fx_ecb_xml": "https://www.ecb.europa.eu/stats/eurofxref/eurofxref-daily.xml",
    "orlen_pl": "https://www.petrodom.pl/en/current-wholesale-fuel-prices-provided-by-pkn-orlen/",
    "orlen_pl_alt": "https://www.petrodom.pl/oferta/aktualne-hurtowe-ceny-paliw-orlen/",
    "orlen_lt_list": "https://www.orlenlietuva.lt/lt/wholesale/_layouts/f2hPriceTable/default.aspx",
    "orlen_lt_base": "https://www.orlenlietuva.lt",
    "elvis_de_fallback": "https://www.fuel-prices.eu/cheapest/",
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
