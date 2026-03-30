"""
Fuel Price Tracker v7
Improvements: structured logging, retry with backoff, type hints,
centralized config, data validation, per-source error isolation.
"""
from __future__ import annotations

import calendar
import io
import json
import logging
import re
import sys
import time
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any, Optional

import requests
from bs4 import BeautifulSoup
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

from config import (
    BSH_SE_MAX,
    BSH_SE_MIN,
    DIESEL_EUR_MAX,
    DIESEL_EUR_MIN,
    EU_COUNTRIES,
    EU_COUNTRY_NAMES,
    EU_WEEKLY_COLUMNS,
    EXCEL_FONT_FAMILY,
    EXCEL_PATH,
    FX_PLN_EUR_MAX,
    FX_PLN_EUR_MIN,
    FX_SEK_EUR_MAX,
    FX_SEK_EUR_MIN,
    FX_TIMEOUT,
    HEADERS,
    JSON_PATH,
    MAX_DAILY_CHANGE_PCT,
    MAX_RETRIES,
    ORLEN_LT_MAX,
    ORLEN_LT_MIN,
    ORLEN_PL_MAX,
    ORLEN_PL_MIN,
    REQUEST_TIMEOUT,
    RETRY_BACKOFF_FACTOR,
    URLS,
)

# ── Logging setup ──
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
    datefmt="%H:%M:%S",
)
logger = logging.getLogger("fuel-tracker")

TODAY = datetime.now()
TODAY_STR = TODAY.strftime("%Y-%m-%d")
WDAY = TODAY.weekday()


# ── HTTP session with retry ──
def create_session() -> requests.Session:
    """Create a requests session with automatic retry and backoff."""
    session = requests.Session()
    retry = Retry(
        total=MAX_RETRIES,
        backoff_factor=RETRY_BACKOFF_FACTOR,
        status_forcelist=[429, 500, 502, 503, 504],
        allowed_methods=["GET"],
    )
    adapter = HTTPAdapter(max_retries=retry)
    session.mount("https://", adapter)
    session.mount("http://", adapter)
    session.headers.update(HEADERS)
    return session


SESSION = create_session()


# ── Utilities ──
def clean_num(s: Any) -> Optional[float]:
    """Parse numbers like '1 756.06', '6 192', '1,234.56', '1234,56'."""
    if s is None:
        return None
    s = str(s).strip()
    s = re.sub(r"[€$£\xa0]", "", s)
    # "1 756.06" — spaces as thousand sep, dot as decimal
    if re.match(r"^\d[\d ]+\.\d+$", s):
        return float(s.replace(" ", ""))
    # "1 756,06" — spaces as thousand sep, comma as decimal
    if re.match(r"^\d[\d ]+,\d+$", s):
        return float(s.replace(" ", "").replace(",", "."))
    # "1,234.56"
    if "," in s and "." in s:
        return float(s.replace(",", ""))
    # "1234,56"
    if "," in s and "." not in s:
        s = s.replace(",", ".")
    s = s.replace(" ", "")
    try:
        return float(s)
    except ValueError:
        return None


def validate_fx(pln: float, sek: float) -> bool:
    """Validate FX rates are within historical norms."""
    if not (FX_PLN_EUR_MIN <= pln <= FX_PLN_EUR_MAX):
        logger.warning("PLN/EUR rate %.4f outside expected range [%.1f, %.1f]", pln, FX_PLN_EUR_MIN, FX_PLN_EUR_MAX)
        return False
    if not (FX_SEK_EUR_MIN <= sek <= FX_SEK_EUR_MAX):
        logger.warning("SEK/EUR rate %.4f outside expected range [%.1f, %.1f]", sek, FX_SEK_EUR_MIN, FX_SEK_EUR_MAX)
        return False
    return True


def validate_price_change(
    source: str, new_val: float, prev_val: Optional[float]
) -> bool:
    """Warn if price changed more than MAX_DAILY_CHANGE_PCT from previous day."""
    if prev_val is None or prev_val == 0:
        return True
    pct = abs((new_val - prev_val) / prev_val) * 100
    if pct > MAX_DAILY_CHANGE_PCT:
        logger.warning(
            "%s: %.1f%% change (%.4f → %.4f) exceeds %.0f%% threshold",
            source, pct, prev_val, new_val, MAX_DAILY_CHANGE_PCT,
        )
        return False  # still use the value, but flag it
    return True


def get_previous_values() -> dict[str, Optional[float]]:
    """Read previous day's values from Excel for outlier detection."""
    prev: dict[str, Optional[float]] = {}
    if not EXCEL_PATH.exists():
        return prev
    try:
        wb = load_workbook(str(EXCEL_PATH), data_only=True)
        if "Daily Tracker" not in wb.sheetnames:
            return prev
        ws = wb["Daily Tracker"]
        # Row 5 is newest; row 6 is previous
        for row in range(5, min(20, ws.max_row + 1)):
            date_cell = ws.cell(row=row, column=1).value
            if date_cell is None:
                continue
            ds = date_cell.strftime("%Y-%m-%d") if hasattr(date_cell, "strftime") else str(date_cell)[:10]
            if ds == TODAY_STR:
                continue  # skip today's row if it exists
            # Found previous day
            prev["orlen_pl_pln"] = ws.cell(row=row, column=3).value
            prev["orlen_lt_eur"] = ws.cell(row=row, column=6).value
            prev["elvis_de_eur"] = ws.cell(row=row, column=9).value
            prev["bsh_se_sek"] = ws.cell(row=row, column=10).value
            break
        wb.close()
    except Exception as e:
        logger.warning("Could not read previous values: %s", e)
    return prev


# ═══════════════════════════════════════
# 1. FX RATES — with multiple fallback APIs
# ═══════════════════════════════════════
def fetch_fx() -> Optional[dict[str, float]]:
    """Fetch EUR → PLN and SEK exchange rates with triple fallback."""
    # Source 1: frankfurter.app (ECB data)
    try:
        r = SESSION.get(URLS["fx_frankfurter"], timeout=FX_TIMEOUT)
        r.raise_for_status()
        d = r.json().get("rates", {})
        if d.get("PLN") and d.get("SEK"):
            pln, sek = float(d["PLN"]), float(d["SEK"])
            if validate_fx(pln, sek):
                logger.info("FX frankfurter.app: PLN=%.4f, SEK=%.4f", pln, sek)
                return {"PLN_EUR": pln, "SEK_EUR": sek}
    except Exception as e:
        logger.warning("FX frankfurter.app failed: %s", e)

    time.sleep(2)

    # Source 2: exchangerate.host
    try:
        r = SESSION.get(URLS["fx_exchangerate"], timeout=FX_TIMEOUT)
        r.raise_for_status()
        d = r.json().get("rates", {})
        if d.get("PLN") and d.get("SEK"):
            pln, sek = float(d["PLN"]), float(d["SEK"])
            if validate_fx(pln, sek):
                logger.info("FX exchangerate.host: PLN=%.4f, SEK=%.4f", pln, sek)
                return {"PLN_EUR": pln, "SEK_EUR": sek}
    except Exception as e:
        logger.warning("FX exchangerate.host failed: %s", e)

    time.sleep(2)

    # Source 3: ECB direct XML
    try:
        r = SESSION.get(URLS["fx_ecb_xml"], timeout=FX_TIMEOUT)
        r.raise_for_status()
        pln_m = re.search(r"currency='PLN'\s+rate='([\d.]+)'", r.text)
        sek_m = re.search(r"currency='SEK'\s+rate='([\d.]+)'", r.text)
        if pln_m and sek_m:
            pln, sek = float(pln_m.group(1)), float(sek_m.group(1))
            if validate_fx(pln, sek):
                logger.info("FX ECB XML: PLN=%.4f, SEK=%.4f", pln, sek)
                return {"PLN_EUR": pln, "SEK_EUR": sek}
    except Exception as e:
        logger.warning("FX ECB XML failed: %s", e)

    logger.error("FX: all 3 sources failed")
    return None


# ═══════════════════════════════════════
# 2. ORLEN PL — via petrodom.pl
# ═══════════════════════════════════════
def fetch_orlen_pl() -> Optional[dict[str, float]]:
    """Scrape Orlen PL Ekodiesel wholesale price in PLN/m³."""
    try:
        r = SESSION.get(URLS["orlen_pl"], timeout=REQUEST_TIMEOUT)
        r.raise_for_status()
        soup = BeautifulSoup(r.text, "html.parser")

        # Method 1: Table parsing
        tables = soup.find_all("table")
        logger.info("Orlen PL: found %d tables", len(tables))
        for table in tables:
            for row in table.find_all("tr"):
                cells = [td.get_text(strip=True) for td in row.find_all(["td", "th"])]
                for i, cell in enumerate(cells):
                    if (
                        "ekodiesel" in cell.lower()
                        and "arktyczny" not in cell.lower()
                        and "grzewczy" not in cell.lower()
                    ):
                        logger.debug("Orlen PL: found Ekodiesel cell: %r", cell)
                        for j in range(i + 1, min(i + 3, len(cells))):
                            price = clean_num(cells[j])
                            if price and ORLEN_PL_MIN < price < ORLEN_PL_MAX:
                                logger.info("Orlen PL: Ekodiesel = %.2f PLN/m³", price)
                                return {"price_pln_m3": price}

        # Method 2: Text fallback
        text = soup.get_text(" ", strip=True)
        logger.info("Orlen PL: table parsing failed, trying text search")
        m = re.search(r"[Ee]kodiesel[^0-9]{0,60}?(\d[\d\s\xa0]*\d)", text)
        if m:
            price = clean_num(m.group(1))
            if price and ORLEN_PL_MIN < price < ORLEN_PL_MAX:
                logger.info("Orlen PL: Ekodiesel (text fallback) = %.2f PLN/m³", price)
                return {"price_pln_m3": price}

        logger.warning("Orlen PL: Ekodiesel not found in tables or text")
        return None
    except Exception as e:
        logger.error("Orlen PL: %s", e)
        return None


# ═══════════════════════════════════════
# 3. ORLEN LT — PDF: pardavimo kaina su PVM
# ═══════════════════════════════════════
def parse_orlen_lt_pdf(pdf_bytes: bytes) -> Optional[dict[str, float]]:
    """Extract diesel price from Orlen LT PDF protocol."""
    try:
        import pdfplumber

        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            text = pdf.pages[0].extract_text() or ""
            for line in text.split("\n"):
                if "Dyzelinas E kl. su RRME" not in line:
                    continue
                # Extract numbers like "897.69" and "1 756.06"
                # Pattern: digits (optionally with spaces), dot, exactly 2 decimal digits
                nums = re.findall(r"(\d[\d ]*\.\d{2})", line)
                cleaned = [float(n.replace(" ", "")) for n in nums]
                if cleaned:
                    # LAST number = Pardavimo kaina su PVM (EUR/1000l)
                    selling_price = cleaned[-1]
                    if ORLEN_LT_MIN < selling_price < ORLEN_LT_MAX:
                        eur_l = round(selling_price / 1000, 4)
                        logger.info(
                            "Orlen LT: %.2f EUR/1000l = %.4f EUR/l (su PVM)",
                            selling_price, eur_l,
                        )
                        return {"price_eur_l": eur_l}
                    else:
                        logger.warning(
                            "Orlen LT: price %.2f outside range [%d, %d]",
                            selling_price, ORLEN_LT_MIN, ORLEN_LT_MAX,
                        )
                break  # only check the Dyzelinas line
    except ImportError:
        logger.warning("Orlen LT: pdfplumber not installed")
    except Exception as e:
        logger.warning("Orlen LT: PDF parse error: %s", e)
    return None


def fetch_orlen_lt() -> Optional[dict[str, float]]:
    """Fetch Orlen LT diesel price from PDF protocols."""
    try:
        # Step 1: Try TODAY's PDF first (constructed URL) — it may not be in archive yet
        pdf_links: list[str] = []
        for days_back in range(3):
            d = TODAY - timedelta(days=days_back)
            pdf_links.append(
                f"{URLS['orlen_lt_base']}/LT/Wholesale/Prices/"
                f"Kainos {d.strftime('%Y %m %d')} realizacija internet.pdf"
            )

        # Step 2: Also get links from archive page
        try:
            r = SESSION.get(URLS["orlen_lt_list"], timeout=REQUEST_TIMEOUT)
            r.raise_for_status()
            soup = BeautifulSoup(r.text, "html.parser")
            for a in soup.find_all("a", href=True):
                href = a["href"]
                if ".pdf" in href.lower() and "kainos" in href.lower():
                    if not href.startswith("http"):
                        href = URLS["orlen_lt_base"] + href
                    if href not in pdf_links:
                        pdf_links.append(href)
        except Exception as e:
            logger.warning("Orlen LT: archive page failed: %s", e)

        for pdf_url in pdf_links[:7]:
            try:
                filename = pdf_url.split("/")[-1]
                logger.info("Orlen LT: trying %s", filename)
                r2 = SESSION.get(pdf_url, timeout=15)
                if r2.status_code == 200 and len(r2.content) > 500:
                    price = parse_orlen_lt_pdf(r2.content)
                    if price:
                        return price
            except Exception as e:
                logger.debug("Orlen LT: PDF %s failed: %s", filename, e)
                continue

        logger.warning("Orlen LT: no PDF parsed successfully")
        return None
    except Exception as e:
        logger.error("Orlen LT: %s", e)
        return None


# ═══════════════════════════════════════
# 4. ELVIS DE — DIESEL from fuel-prices.eu
# ═══════════════════════════════════════
def fetch_elvis_de() -> Optional[dict[str, float]]:
    """Scrape Germany diesel price from fuel-prices.eu."""
    try:
        r = SESSION.get(URLS["elvis_de"], timeout=REQUEST_TIMEOUT)
        r.raise_for_status()
        soup = BeautifulSoup(r.text, "html.parser")

        for table in soup.find_all("table"):
            if "diesel" not in table.get_text(" ", strip=True).lower():
                continue
            for row in table.find_all("tr"):
                cells = [td.get_text(strip=True) for td in row.find_all(["td", "th"])]
                row_text = " ".join(cells)
                if "germany" in row_text.lower() or " DE" in row_text:
                    for cell in cells:
                        m = re.search(r"€?(\d\.\d{3})", cell)
                        if m:
                            price = float(m.group(1))
                            if DIESEL_EUR_MIN < price < DIESEL_EUR_MAX:
                                logger.info("Elvis DE: Germany Diesel = %.3f EUR/l", price)
                                return {"price_eur_l": price}

        logger.warning("Elvis DE: Germany diesel not found")
        return None
    except Exception as e:
        logger.error("Elvis DE: %s", e)
        return None


# ═══════════════════════════════════════
# 5. BSH / ST1 SE
# ═══════════════════════════════════════
def fetch_bsh_se() -> Optional[dict[str, float]]:
    """Scrape BSH/ST1 Sweden diesel price in SEK/l."""
    try:
        r = SESSION.get(URLS["bsh_se"], timeout=REQUEST_TIMEOUT)
        r.raise_for_status()
        text = BeautifulSoup(r.text, "html.parser").get_text(" ", strip=True)

        patterns = [
            r"[Dd]iesel\s*(?:MK[123])?\s*[^0-9]{0,30}?(\d{1,2}[.,]\d{2})",
            r"(\d{1,2}[.,]\d{2})\s*(?:kr|SEK)[^0-9]{0,20}[Dd]iesel",
        ]
        for pat in patterns:
            m = re.search(pat, text)
            if m:
                p = float(m.group(1).replace(",", "."))
                if BSH_SE_MIN < p < BSH_SE_MAX:
                    logger.info("BSH SE: Diesel = %.2f SEK/l", p)
                    return {"price_sek_l": p}

        logger.warning("BSH SE: diesel price not found")
        return None
    except Exception as e:
        logger.error("BSH SE: %s", e)
        return None


# ═══════════════════════════════════════
# 6. EU BULLETIN — direct EC XLSX download
# ═══════════════════════════════════════
def fetch_eu_bulletin() -> Optional[dict[str, Any]]:
    """Download EC Weekly Oil Bulletin XLSX — diesel column only."""
    try:
        logger.info("EU Bulletin: downloading EC XLSX...")
        r = SESSION.get(URLS["eu_bulletin_xlsx"], timeout=30)
        if r.status_code != 200:
            logger.warning("EU Bulletin: EC HTTP %d", r.status_code)
            return _fetch_eu_bulletin_fallback()

        from openpyxl import load_workbook as lwb

        wb = lwb(io.BytesIO(r.content), data_only=True)
        countries: dict[str, Optional[float]] = dict(EU_COUNTRIES)
        eu_avg: Optional[float] = None
        de_diesel: Optional[float] = None
        ec_date: Optional[str] = None

        ws = wb[wb.sheetnames[0]]
        logger.info("EU Bulletin: sheet '%s' (%dr x %dc)", ws.title, ws.max_row, ws.max_column)

        # Step 1: Find the DIESEL column dynamically
        diesel_col = 3  # default: column C
        for row in range(1, 6):
            for col in range(1, min(10, ws.max_column + 1)):
                val = str(ws.cell(row=row, column=col).value or "").lower()
                if "gas oil" in val or "diesel" in val or "gasoil" in val:
                    diesel_col = col
                    logger.info("EU Bulletin: diesel column %d (row %d)", col, row)
                    break
            if diesel_col != 3:
                break

        # Step 2: Find date in first few rows
        for row in range(1, 6):
            for col in range(1, min(10, ws.max_column + 1)):
                val = ws.cell(row=row, column=col).value
                if val is None:
                    continue
                if hasattr(val, "strftime"):
                    ec_date = val.strftime("%Y-%m-%d")
                    logger.info("EU Bulletin: date (datetime) = %s", ec_date)
                elif isinstance(val, str):
                    dm = re.search(r"(\d{1,2})[/.-](\d{1,2})[/.-](\d{4})", val)
                    if dm:
                        d, mo, y = dm.group(1), dm.group(2), dm.group(3)
                        ec_date = f"{y}-{mo.zfill(2)}-{d.zfill(2)}"
                        logger.info("EU Bulletin: date (text) = %s", ec_date)

        # Step 3: Read country prices from DIESEL column
        for row in range(1, ws.max_row + 1):
            cell0 = str(ws.cell(row=row, column=1).value or "").strip()
            if not cell0:
                continue

            for cname, cc in EU_COUNTRY_NAMES.items():
                if cname.lower() in cell0.lower():
                    val = ws.cell(row=row, column=diesel_col).value
                    if val is not None:
                        try:
                            v = float(val)
                            if DIESEL_EUR_MIN < v < DIESEL_EUR_MAX:
                                countries[cc] = round(v, 4)
                            elif 500 < v < 3500:
                                countries[cc] = round(v / 1000, 4)
                            logger.info("EU Bulletin: %s = %.4f EUR/l", cname, countries[cc] or 0)
                        except (ValueError, TypeError):
                            pass

            if "germany" in cell0.lower():
                val = ws.cell(row=row, column=diesel_col).value
                if val is not None:
                    try:
                        v = float(val)
                        if DIESEL_EUR_MIN < v < DIESEL_EUR_MAX:
                            de_diesel = round(v, 4)
                        elif 500 < v < 3500:
                            de_diesel = round(v / 1000, 4)
                        logger.info("EU Bulletin: Germany = %.4f EUR/l", de_diesel or 0)
                    except (ValueError, TypeError):
                        pass

            # EU weighted average — EU27, skip Euro Area
            c0l = cell0.lower()
            if eu_avg is None and (
                "ce/ec" in c0l or "eur27" in c0l or "eu" in c0l
            ) and (
                "average" in c0l or "weighted" in c0l or "moyenne" in c0l or "durchschnitt" in c0l
            ):
                if "euro area" in c0l or "eurozone" in c0l:
                    continue
                val = ws.cell(row=row, column=diesel_col).value
                if val is not None:
                    try:
                        v = float(val)
                        if DIESEL_EUR_MIN < v < DIESEL_EUR_MAX:
                            eu_avg = round(v, 4)
                        elif 500 < v < 3500:
                            eu_avg = round(v / 1000, 4)
                        logger.info("EU Bulletin: EU27 avg = %.4f EUR/l", eu_avg or 0)
                    except (ValueError, TypeError):
                        pass

        wb.close()
        found = {k: v for k, v in countries.items() if v is not None}
        if found:
            logger.info("EU Bulletin: %d countries, avg=%s, DE=%s, date=%s", len(found), eu_avg, de_diesel, ec_date)
            result: dict[str, Any] = {**countries, "EU_AVG": eu_avg}
            if de_diesel:
                result["DE"] = de_diesel
            if ec_date:
                result["_date"] = ec_date
            return result

        logger.warning("EU Bulletin: no diesel data in EC XLSX")
        return _fetch_eu_bulletin_fallback()

    except Exception as e:
        logger.warning("EU Bulletin: EC XLSX error: %s", e)
        return _fetch_eu_bulletin_fallback()


def _fetch_eu_bulletin_fallback() -> Optional[dict[str, Any]]:
    """Fallback: scrape fuel-prices.eu for EU bulletin data."""
    try:
        r = SESSION.get(URLS["elvis_de"], timeout=REQUEST_TIMEOUT)
        r.raise_for_status()
        soup = BeautifulSoup(r.text, "html.parser")
        countries: dict[str, Optional[float]] = dict(EU_COUNTRIES)
        cc_map = {
            "lithuania": "LT", "latvia": "LV", "estonia": "EE",
            "denmark": "DK", "sweden": "SE", "finland": "FI",
        }
        eu_avg: Optional[float] = None

        for table in soup.find_all("table"):
            if "diesel" not in table.get_text(" ", strip=True).lower():
                continue
            for row in table.find_all("tr"):
                cells = [td.get_text(strip=True) for td in row.find_all(["td", "th"])]
                if len(cells) < 2:
                    continue
                rt = " ".join(cells[:2]).lower()
                for name, cc in cc_map.items():
                    if name in rt:
                        for cell in cells:
                            m = re.search(r"€?(\d\.\d{3})", cell)
                            if m:
                                val = float(m.group(1))
                                if DIESEL_EUR_MIN < val < DIESEL_EUR_MAX:
                                    countries[cc] = val
                                    break

        text = soup.get_text(" ", strip=True)
        m = re.search(r"€(\d\.\d{3})/L\s+for\s+diesel", text)
        if m:
            eu_avg = float(m.group(1))

        found = {k: v for k, v in countries.items() if v}
        if found:
            logger.info("EU Bulletin fallback: %s, avg=%s", found, eu_avg)
            return {**countries, "EU_AVG": eu_avg}
        return None
    except Exception as e:
        logger.error("EU Bulletin fallback: %s", e)
        return None


# ═══════════════════════════════════════
# EXCEL: check if date already exists
# ═══════════════════════════════════════
def date_exists_in_daily(ws: Any, target_date: datetime) -> Optional[int]:
    """Check if target date already has a row in Daily Tracker. Returns row number or None."""
    for row in range(5, min(15, ws.max_row + 1)):
        cell = ws.cell(row=row, column=1).value
        if cell is None:
            continue
        if hasattr(cell, "date"):
            if cell.date() == target_date.date():
                return row
        elif hasattr(cell, "strftime"):
            if cell.strftime("%Y-%m-%d") == target_date.strftime("%Y-%m-%d"):
                return row
        elif isinstance(cell, str) and cell[:10] == target_date.strftime("%Y-%m-%d"):
            return row
    return None


def date_exists_in_weekly(ws: Any, date_str: str) -> Optional[int]:
    """Check if this week's data already exists. Returns row number or None."""
    for row in range(4, min(ws.max_row + 1, 100)):
        cell = ws.cell(row=row, column=1).value
        if cell is None:
            continue
        if hasattr(cell, "strftime"):
            if cell.strftime("%Y-%m-%d") == date_str:
                return row
        elif hasattr(cell, "date"):
            if cell.date().strftime("%Y-%m-%d") == date_str:
                return row
        elif isinstance(cell, str) and cell[:10] == date_str:
            return row
    return None


# ═══════════════════════════════════════
# EXCEL WRITER
# ═══════════════════════════════════════
def update_excel(
    fx: Optional[dict] = None,
    orlen_pl: Optional[dict] = None,
    orlen_lt: Optional[dict] = None,
    elvis_de: Optional[dict] = None,
    bsh_se: Optional[dict] = None,
    eu_bulletin: Optional[dict] = None,
    _failures: Optional[dict] = None,
) -> bool:
    """Write scraped data to Excel workbook."""
    if _failures is None:
        _failures = {}
    if not EXCEL_PATH.exists():
        logger.error("Excel not found: %s", EXCEL_PATH)
        return False

    wb = load_workbook(str(EXCEL_PATH))

    if "Daily Tracker" in wb.sheetnames:
        ws = wb["Daily Tracker"]

        existing_row = date_exists_in_daily(ws, TODAY)
        if existing_row:
            row = existing_row
            logger.info("Excel: updating existing row %d for %s", row, TODAY_STR)
        else:
            ws.insert_rows(5)
            row = 5
            logger.info("Excel: inserting new row %d for %s", row, TODAY_STR)

        ifont = Font(name=EXCEL_FONT_FAMILY, size=10, color="1D4ED8")
        ifill = PatternFill("solid", fgColor="EFF6FF")
        dfont = Font(name=EXCEL_FONT_FAMILY, size=10, color="1F2937")
        brd = Border(
            left=Side("thin", color="D1D5DB"),
            right=Side("thin", color="D1D5DB"),
            top=Side("thin", color="D1D5DB"),
            bottom=Side("thin", color="D1D5DB"),
        )

        def wc(
            col: int,
            val: Any,
            fmt: str = "General",
            font: Font = dfont,
            fill: Optional[PatternFill] = None,
        ) -> None:
            c = ws.cell(row=row, column=col)
            if val is not None or c.value is None:
                c.value = val
            c.number_format = fmt
            c.font = font
            c.border = brd
            c.alignment = Alignment(
                horizontal="right" if col > 2 else "center", vertical="center"
            )
            if fill:
                c.fill = fill

        wc(1, TODAY, "YYYY-MM-DD", Font(name=EXCEL_FONT_FAMILY, size=10, bold=True, color="1F2937"))
        wc(2, calendar.day_abbr[WDAY], font=Font(name=EXCEL_FONT_FAMILY, size=9, color="6B7280"))
        wc(3, orlen_pl["price_pln_m3"] if orlen_pl else None, "#,##0.00", ifont, ifill)
        wc(4, fx["PLN_EUR"] if fx else None, "0.0000", ifont, ifill)

        pl_eur_l = (
            round(orlen_pl["price_pln_m3"] / fx["PLN_EUR"] / 1000, 4)
            if (orlen_pl and fx and fx.get("PLN_EUR"))
            else None
        )
        wc(5, pl_eur_l, "0.000")

        lt_val = orlen_lt["price_eur_l"] if orlen_lt else None
        wc(6, lt_val, "0.000", ifont, ifill)

        delta = round(pl_eur_l - lt_val, 4) if (pl_eur_l and lt_val) else None
        wc(7, delta, '+0.000;-0.000;"-"')
        delta_pct = round(delta / lt_val, 4) if (delta is not None and lt_val) else None
        wc(8, delta_pct, '+0.0%;-0.0%;"-"')

        wc(9, elvis_de["price_eur_l"] if elvis_de else None, "0.000", ifont, ifill)
        if eu_bulletin and eu_bulletin.get("DE"):
            wc(9, eu_bulletin["DE"], "0.000", ifont, ifill)

        wc(10, bsh_se["price_sek_l"] if bsh_se else None, "0.00", ifont, ifill)
        wc(11, fx["SEK_EUR"] if fx else None, "0.0000", ifont, ifill)

        se_eur = (
            round(bsh_se["price_sek_l"] / fx["SEK_EUR"], 4)
            if (bsh_se and fx and fx.get("SEK_EUR"))
            else None
        )
        wc(12, se_eur, "0.000")

        ok = [
            k for k, v in {
                "FX": fx, "PL": orlen_pl, "LT": orlen_lt,
                "DE": (elvis_de or (eu_bulletin and eu_bulletin.get("DE"))),
                "SE": bsh_se, "EU": eu_bulletin,
            }.items() if v
        ]
        notes = f"Auto: {','.join(ok)}"
        if _failures:
            miss = ",".join([f"{k}={v}" for k, v in _failures.items()])
            notes += f"|MISS:{miss}"
        wc(13, notes, font=Font(name=EXCEL_FONT_FAMILY, size=9, color="6B7280"))
        wc(14, "Auto", font=Font(name=EXCEL_FONT_FAMILY, size=8, color="9CA3AF"))
        logger.info("Excel: row %d: %s [%s]", row, TODAY_STR, ",".join(ok))

    # Weekly — use actual EC date, not calculated Monday
    if eu_bulletin and "Weekly Oil Bulletin" in wb.sheetnames:
        ws_w = wb["Weekly Oil Bulletin"]

        ec_date_str = eu_bulletin.get("_date")
        if ec_date_str:
            week_date = datetime.strptime(ec_date_str, "%Y-%m-%d")
            logger.info("Excel: using EC date: %s", ec_date_str)
        else:
            week_date = TODAY - timedelta(days=WDAY)
            ec_date_str = week_date.strftime("%Y-%m-%d")
            logger.info("Excel: using calculated Monday: %s", ec_date_str)

        existing = date_exists_in_weekly(ws_w, ec_date_str)
        if existing:
            w_row = existing
            logger.info("Excel: updating weekly row %d for %s", w_row, ec_date_str)
        else:
            ws_w.insert_rows(4)
            w_row = 4
            logger.info("Excel: inserting weekly row %d for %s", w_row, ec_date_str)

        ws_w.cell(row=w_row, column=1).value = week_date
        ws_w.cell(row=w_row, column=1).number_format = "YYYY-MM-DD"
        ws_w.cell(row=w_row, column=1).font = Font(name=EXCEL_FONT_FAMILY, size=10, bold=True, color="1F2937")

        for k, col in EU_WEEKLY_COLUMNS.items():
            val = eu_bulletin.get(k)
            if val is not None:
                c = ws_w.cell(row=w_row, column=col)
                c.value = val
                c.number_format = "0.000"
                c.font = Font(name=EXCEL_FONT_FAMILY, size=10, color="1D4ED8")

        lt_v = eu_bulletin.get("LT")
        eu_v = eu_bulletin.get("EU_AVG")
        ws_w.cell(row=w_row, column=9).value = (
            round((lt_v - eu_v) / eu_v, 4) if (lt_v and eu_v) else None
        )
        ws_w.cell(row=w_row, column=9).number_format = '+0.0%;-0.0%;"-"'

    wb.save(str(EXCEL_PATH))
    logger.info("Excel: saved %s", EXCEL_PATH)
    return True


# ═══════════════════════════════════════
# MAIN
# ═══════════════════════════════════════
def main() -> None:
    logger.info("=" * 50)
    logger.info("FUEL PRICE TRACKER v7 — %s", TODAY_STR)
    logger.info("=" * 50)

    R: dict[str, Any] = {}
    FAIL: dict[str, str] = {}

    # Load previous values for outlier detection
    prev = get_previous_values()

    # FX rates
    R["fx"] = fetch_fx()
    if not R["fx"]:
        FAIL["FX"] = "ALL_SOURCES_FAILED"

    if WDAY < 5:
        # Orlen PL
        R["orlen_pl"] = fetch_orlen_pl()
        if not R["orlen_pl"]:
            FAIL["PL"] = "NO_DATA"
        elif prev.get("orlen_pl_pln"):
            validate_price_change("Orlen PL", R["orlen_pl"]["price_pln_m3"], prev["orlen_pl_pln"])

        # Orlen LT
        R["orlen_lt"] = fetch_orlen_lt()
        if not R["orlen_lt"]:
            FAIL["LT"] = "NO_PDF"
        elif prev.get("orlen_lt_eur"):
            validate_price_change("Orlen LT", R["orlen_lt"]["price_eur_l"], prev["orlen_lt_eur"])

        # Elvis DE
        R["elvis_de"] = fetch_elvis_de()
        if not R["elvis_de"]:
            FAIL["DE"] = "NO_DATA"
        elif prev.get("elvis_de_eur"):
            validate_price_change("Elvis DE", R["elvis_de"]["price_eur_l"], prev["elvis_de_eur"])

        # BSH SE
        R["bsh_se"] = fetch_bsh_se()
        if not R["bsh_se"]:
            FAIL["SE"] = "NO_DATA"
        elif prev.get("bsh_se_sek"):
            validate_price_change("BSH SE", R["bsh_se"]["price_sek_l"], prev["bsh_se_sek"])
    else:
        logger.info("Weekend — skipping daily sources")
        for k in ["orlen_pl", "orlen_lt", "elvis_de", "bsh_se"]:
            R[k] = None
        FAIL = {"PL": "WEEKEND", "LT": "WEEKEND", "DE": "WEEKEND", "SE": "WEEKEND"}

    # EU bulletin
    R["eu_bulletin"] = fetch_eu_bulletin()
    if not R["eu_bulletin"]:
        FAIL["EU"] = "NO_DATA"

    # Summary
    ok = sum(1 for v in R.values() if v is not None)
    logger.info("RESULTS: %d/%d sources succeeded", ok, len(R))
    for k, v in R.items():
        status = "OK" if v else "FAILED"
        logger.info("  %s %s: %s", status, k, v if v else FAIL.get(k.upper()[:2], "unknown"))

    # Write Excel
    success = update_excel(**R, _failures=FAIL)

    # Write JSON status with per-source details
    status_data = {
        "date": TODAY_STR,
        "version": "v7",
        "sources_ok": ok,
        "sources_total": len(R),
        "results": {k: bool(v) for k, v in R.items()},
        "data": {k: v for k, v in R.items() if v},
        "failures": FAIL,
    }
    JSON_PATH.write_text(json.dumps(status_data, indent=2, default=str))
    logger.info("JSON status written to %s", JSON_PATH)

    if not success:
        logger.error("Excel update failed!")
        sys.exit(1)

    logger.info("Done!")


if __name__ == "__main__":
    main()
