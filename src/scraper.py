"""
Fuel Price Tracker — Automated Data Collection
Fetches diesel prices from all configured sources and updates the Excel tracker.
Designed to run daily via GitHub Actions.
"""

import requests
import json
import re
import os
import sys
from datetime import datetime, timedelta
from bs4 import BeautifulSoup
from openpyxl import load_workbook
from pathlib import Path

EXCEL_PATH = Path(__file__).parent.parent / "fuel_tracker.xlsx"
HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "en-US,en;q=0.9,lt;q=0.8,pl;q=0.7",
}

TODAY = datetime.now().strftime("%Y-%m-%d")
WEEKDAY = datetime.now().weekday()  # 0=Mon, 6=Sun


def log(source, msg, level="INFO"):
    print(f"[{level}] {source}: {msg}")


# ═══════════════════════════════════════════════
# 1. FX RATES — ECB via Frankfurter API
# ═══════════════════════════════════════════════
def fetch_fx_rates():
    """Fetch PLN/EUR and SEK/EUR from ECB via frankfurter.app"""
    try:
        url = "https://api.frankfurter.app/latest?from=EUR&to=PLN,SEK"
        r = requests.get(url, headers=HEADERS, timeout=15)
        r.raise_for_status()
        data = r.json()
        rates = data.get("rates", {})
        pln = rates.get("PLN")
        sek = rates.get("SEK")
        log("FX", f"PLN/EUR={pln}, SEK/EUR={sek}")
        return {"PLN_EUR": pln, "SEK_EUR": sek, "date": data.get("date", TODAY)}
    except Exception as e:
        log("FX", f"Failed: {e}", "ERROR")
        return None


# ═══════════════════════════════════════════════
# 2. ORLEN PL — Ekodiesel wholesale price
# ═══════════════════════════════════════════════
def fetch_orlen_pl():
    """Fetch Ekodiesel wholesale price from Orlen PL"""
    try:
        # Orlen PL publishes wholesale prices - try their API/page
        url = "https://www.orlen.pl/pl/dla-biznesu/hurtowe-ceny-paliw"
        r = requests.get(url, headers=HEADERS, timeout=20)
        r.raise_for_status()
        soup = BeautifulSoup(r.text, "html.parser")

        # Look for Ekodiesel price in page content
        # Orlen typically has structured data or table with prices
        text = soup.get_text()

        # Try to find price patterns near "Ekodiesel" or "EKODIESEL"
        patterns = [
            r'[Ee]kodiesel[^0-9]*?(\d[\d\s]*[\.,]\d{2})',
            r'[Ee]kodiesel.*?(\d{4,5}[\.,]\d{2})',
            r'ON\s+Ekodiesel[^0-9]*?(\d[\d\s]*[\.,]\d{2})',
        ]
        for pat in patterns:
            match = re.search(pat, text, re.DOTALL)
            if match:
                price_str = match.group(1).replace(" ", "").replace(",", ".")
                price = float(price_str)
                if 3000 < price < 10000:  # sanity: PLN/m³ should be in this range
                    log("Orlen PL", f"Ekodiesel={price} PLN/m³")
                    return {"price_pln_m3": price}

        # If page is JS-rendered, try known API endpoints
        api_urls = [
            "https://www.orlen.pl/api/fuel-prices",
            "https://www.orlen.pl/services/fuel-prices",
        ]
        for api_url in api_urls:
            try:
                r2 = requests.get(api_url, headers=HEADERS, timeout=10)
                if r2.status_code == 200:
                    data = r2.json()
                    log("Orlen PL", f"API response keys: {list(data.keys()) if isinstance(data, dict) else 'list'}")
                    # Parse based on actual API structure
                    return parse_orlen_pl_api(data)
            except:
                continue

        log("Orlen PL", "Could not extract price from page or API", "WARN")
        return None
    except Exception as e:
        log("Orlen PL", f"Failed: {e}", "ERROR")
        return None


def parse_orlen_pl_api(data):
    """Try to extract Ekodiesel price from Orlen API response"""
    if isinstance(data, dict):
        for key, val in data.items():
            if isinstance(val, list):
                for item in val:
                    if isinstance(item, dict):
                        name = str(item.get("name", "") or item.get("productName", "")).lower()
                        if "ekodiesel" in name or "diesel" in name:
                            price = item.get("price") or item.get("value") or item.get("netPrice")
                            if price:
                                log("Orlen PL", f"API: Ekodiesel={price}")
                                return {"price_pln_m3": float(price)}
    return None


# ═══════════════════════════════════════════════
# 3. ORLEN LT — Dyzelinas E wholesale
# ═══════════════════════════════════════════════
def fetch_orlen_lt():
    """Fetch diesel price from Orlen Lietuva wholesale"""
    try:
        url = "https://www.orlenlietuva.lt/LT/Wholesale/Puslapiai/Kainu-protokolai.aspx"
        r = requests.get(url, headers=HEADERS, timeout=20)
        r.raise_for_status()
        soup = BeautifulSoup(r.text, "html.parser")
        text = soup.get_text()

        # Look for diesel price pattern — typically EUR value like 1.234 or 1234.56
        # Near keywords: Dyzelinas, Diesel, E kl., RRME
        patterns = [
            r'[Dd]yzelinas\s+E[^0-9]*?(\d+[\.,]\d{2,4})',
            r'[Dd]iesel[^0-9]*?(\d+[\.,]\d{2,4})',
            r'su\s+(?:akcizu|PVM)[^0-9]*?(\d+[\.,]\d{2,4})',
        ]

        for pat in patterns:
            match = re.search(pat, text, re.DOTALL)
            if match:
                price_str = match.group(1).replace(",", ".")
                price = float(price_str)
                # Price in EUR/l should be 0.8-2.5; in EUR/t it's 800-2500
                if 0.5 < price < 3.0:
                    log("Orlen LT", f"Diesel E={price} EUR/l")
                    return {"price_eur_l": price}
                elif 500 < price < 3000:
                    price_l = price / 1000 * 1.2  # approximate t to l conversion
                    log("Orlen LT", f"Diesel E={price} EUR/t → ~{price_l:.3f} EUR/l")
                    return {"price_eur_l": price_l}

        # Try to find PDF links (kainų protokolai are often PDFs)
        pdf_links = [a["href"] for a in soup.find_all("a", href=True) if ".pdf" in a["href"].lower()]
        if pdf_links:
            log("Orlen LT", f"Found {len(pdf_links)} PDF links — would need PDF parsing", "WARN")

        log("Orlen LT", "Could not extract price", "WARN")
        return None
    except Exception as e:
        log("Orlen LT", f"Failed: {e}", "ERROR")
        return None


# ═══════════════════════════════════════════════
# 4. ELVIS FSC DE — mehr-tanken.de
# ═══════════════════════════════════════════════
def fetch_elvis_de():
    """Fetch diesel price from mehr-tanken.de"""
    try:
        url = "https://www.mehr-tanken.de/aktuelle-spritpreise/"
        r = requests.get(url, headers=HEADERS, timeout=20)
        r.raise_for_status()
        soup = BeautifulSoup(r.text, "html.parser")
        text = soup.get_text()

        # Look for diesel price — typically 1.XX or 1.XXX EUR
        patterns = [
            r'[Dd]iesel[^0-9]*?(\d[\.,]\d{2,3})\s*(?:€|EUR)',
            r'[Dd]iesel[^0-9]*?(\d[\.,]\d{2,3})',
            r'(\d[\.,]\d{2,3})\s*(?:€|EUR)[^0-9]*?[Dd]iesel',
        ]

        for pat in patterns:
            match = re.search(pat, text)
            if match:
                price_str = match.group(1).replace(",", ".")
                price = float(price_str)
                if 0.8 < price < 3.0:
                    log("Elvis DE", f"Diesel={price} EUR/l")
                    return {"price_eur_l": price}

        log("Elvis DE", "Could not extract price", "WARN")
        return None
    except Exception as e:
        log("Elvis DE", f"Failed: {e}", "ERROR")
        return None


# ═══════════════════════════════════════════════
# 5. BSH/ST1 SE — st1.se/foretag/listpris
# ═══════════════════════════════════════════════
def fetch_bsh_se():
    """Fetch diesel list price from ST1 Sweden"""
    try:
        url = "https://st1.se/foretag/listpris"
        r = requests.get(url, headers=HEADERS, timeout=20)
        r.raise_for_status()
        soup = BeautifulSoup(r.text, "html.parser")
        text = soup.get_text()

        # Look for diesel price in SEK — typically 18-25 SEK/l
        patterns = [
            r'[Dd]iesel\s*(?:MK[13])?\s*[^0-9]*?(\d{1,2}[\.,]\d{2})\s*(?:kr|SEK)',
            r'[Dd]iesel[^0-9]*?(\d{1,2}[\.,]\d{2})',
            r'(\d{1,2}[\.,]\d{2})\s*(?:kr|SEK)[^0-9]*?[Dd]iesel',
        ]

        for pat in patterns:
            match = re.search(pat, text)
            if match:
                price_str = match.group(1).replace(",", ".")
                price = float(price_str)
                if 10 < price < 35:  # SEK/l range
                    log("BSH SE", f"Diesel={price} SEK/l")
                    return {"price_sek_l": price}

        log("BSH SE", "Could not extract price", "WARN")
        return None
    except Exception as e:
        log("BSH SE", f"Failed: {e}", "ERROR")
        return None


# ═══════════════════════════════════════════════
# 6. EU WEEKLY OIL BULLETIN
# ═══════════════════════════════════════════════
def fetch_eu_bulletin():
    """Fetch latest diesel prices from EU Weekly Oil Bulletin (runs on Mondays)"""
    try:
        # The EC publishes data at this endpoint
        url = "https://energy.ec.europa.eu/data-and-analysis/weekly-oil-bulletin_en"
        r = requests.get(url, headers=HEADERS, timeout=20)
        r.raise_for_status()
        soup = BeautifulSoup(r.text, "html.parser")

        # Try to find download link for the latest data file
        links = soup.find_all("a", href=True)
        data_links = [a for a in links if any(kw in a["href"].lower() for kw in [".csv", ".xls", "download", "prices_with_taxes"])]

        if data_links:
            data_url = data_links[0]["href"]
            if not data_url.startswith("http"):
                data_url = "https://energy.ec.europa.eu" + data_url
            log("EU Bulletin", f"Found data link: {data_url}")

            # Download and parse the data file
            r2 = requests.get(data_url, headers=HEADERS, timeout=20)
            r2.raise_for_status()
            return parse_eu_bulletin_data(r2.text)

        # Fallback: try to parse prices from the HTML page itself
        text = soup.get_text()
        return parse_eu_bulletin_html(text)

    except Exception as e:
        log("EU Bulletin", f"Failed: {e}", "ERROR")
        return None


def parse_eu_bulletin_data(csv_text):
    """Parse EU bulletin CSV/text data"""
    countries = {"LT": None, "LV": None, "EE": None, "DK": None, "SE": None, "FI": None}
    eu_avg = None

    lines = csv_text.split("\n")
    for line in lines:
        parts = line.split(",") if "," in line else line.split("\t")
        for i, part in enumerate(parts):
            part_clean = part.strip().upper()
            for cc in countries:
                if cc == part_clean or f"({cc})" in part_clean:
                    # Look for numeric value in nearby columns
                    for j in range(i+1, min(i+5, len(parts))):
                        try:
                            val = float(parts[j].strip().replace(",", "."))
                            if 0.5 < val < 3.0:
                                countries[cc] = val
                                break
                        except:
                            continue

            if "EU" in part_clean and ("AVG" in part_clean or "AVERAGE" in part_clean or "WEIGHTED" in part_clean):
                for j in range(i+1, min(i+5, len(parts))):
                    try:
                        val = float(parts[j].strip().replace(",", "."))
                        if 0.5 < val < 3.0:
                            eu_avg = val
                            break
                    except:
                        continue

    if any(v for v in countries.values() if v is not None):
        log("EU Bulletin", f"Parsed: {countries}, EU avg={eu_avg}")
        return {**countries, "EU_AVG": eu_avg}

    log("EU Bulletin", "Could not parse data file", "WARN")
    return None


def parse_eu_bulletin_html(text):
    """Fallback: try to extract prices from page text"""
    countries = {"LT": None, "LV": None, "EE": None, "DK": None, "SE": None, "FI": None}
    for cc in countries:
        # Look for country code followed by a price
        patterns = [
            rf'(?:Lithuania|Latvia|Estonia|Denmark|Sweden|Finland)\s*[\s:]*(\d[\.,]\d{{2,3}})',
        ]
        country_names = {"LT": "Lithuania", "LV": "Latvia", "EE": "Estonia", "DK": "Denmark", "SE": "Sweden", "FI": "Finland"}
        pat = rf'{country_names[cc]}[\s\S]{{0,50}}?(\d[\.,]\d{{2,3}})'
        match = re.search(pat, text)
        if match:
            val = float(match.group(1).replace(",", "."))
            if 0.5 < val < 3.0:
                countries[cc] = val

    if any(v for v in countries.values() if v is not None):
        log("EU Bulletin", f"HTML parsed: {countries}")
        return countries
    return None


# ═══════════════════════════════════════════════
# EXCEL WRITER
# ═══════════════════════════════════════════════
def find_date_row(ws, target_date_str, start_row=5, date_col=1):
    """Find row matching today's date, or first empty row"""
    target = datetime.strptime(target_date_str, "%Y-%m-%d").date()
    first_empty = None

    for row in range(start_row, start_row + 65):
        cell_val = ws.cell(row=row, column=date_col).value
        if cell_val is None:
            if first_empty is None:
                first_empty = row
            continue
        if hasattr(cell_val, 'date'):
            cell_date = cell_val.date()
        elif isinstance(cell_val, str):
            try:
                cell_date = datetime.strptime(cell_val, "%Y-%m-%d").date()
            except:
                continue
        else:
            continue

        if cell_date == target:
            return row

    return first_empty


def find_weekly_row(ws, start_row=4, date_col=1):
    """Find first empty row in weekly sheet, or row matching this week"""
    today = datetime.now().date()
    monday = today - timedelta(days=today.weekday())

    for row in range(start_row, start_row + 35):
        cell_val = ws.cell(row=row, column=date_col).value
        if cell_val is None:
            return row
        if hasattr(cell_val, 'date'):
            if cell_val.date() == monday:
                return row

    return start_row


def update_excel(fx=None, orlen_pl=None, orlen_lt=None, elvis_de=None, bsh_se=None, eu_bulletin=None):
    """Write all fetched data into the Excel tracker"""
    if not EXCEL_PATH.exists():
        log("Excel", f"File not found: {EXCEL_PATH}", "ERROR")
        return False

    wb = load_workbook(str(EXCEL_PATH))

    # ─── DAILY TRACKER ───
    if "Daily Tracker" in wb.sheetnames:
        ws = wb["Daily Tracker"]
        row = find_date_row(ws, TODAY)

        if row:
            log("Excel", f"Writing daily data to row {row}")

            # Date (if empty)
            if ws.cell(row=row, column=1).value is None:
                ws.cell(row=row, column=1).value = datetime.strptime(TODAY, "%Y-%m-%d")

            # Orlen PL — col C (PLN/m³)
            if orlen_pl and orlen_pl.get("price_pln_m3"):
                ws.cell(row=row, column=3).value = orlen_pl["price_pln_m3"]
                log("Excel", f"  C{row} = {orlen_pl['price_pln_m3']} (Orlen PL)")

            # FX PLN/EUR — col D
            if fx and fx.get("PLN_EUR"):
                ws.cell(row=row, column=4).value = fx["PLN_EUR"]
                log("Excel", f"  D{row} = {fx['PLN_EUR']} (PLN/EUR)")

            # Orlen LT — col F (EUR/l)
            if orlen_lt and orlen_lt.get("price_eur_l"):
                ws.cell(row=row, column=6).value = orlen_lt["price_eur_l"]
                log("Excel", f"  F{row} = {orlen_lt['price_eur_l']} (Orlen LT)")

            # Elvis DE — col I (EUR/l)
            if elvis_de and elvis_de.get("price_eur_l"):
                ws.cell(row=row, column=9).value = elvis_de["price_eur_l"]
                log("Excel", f"  I{row} = {elvis_de['price_eur_l']} (Elvis DE)")

            # BSH SE — col J (SEK/l)
            if bsh_se and bsh_se.get("price_sek_l"):
                ws.cell(row=row, column=10).value = bsh_se["price_sek_l"]
                log("Excel", f"  J{row} = {bsh_se['price_sek_l']} (BSH SE)")

            # FX SEK/EUR — col K
            if fx and fx.get("SEK_EUR"):
                ws.cell(row=row, column=11).value = fx["SEK_EUR"]
                log("Excel", f"  K{row} = {fx['SEK_EUR']} (SEK/EUR)")
        else:
            log("Excel", "No matching date row found in Daily Tracker", "WARN")

    # ─── WEEKLY OIL BULLETIN ───
    if eu_bulletin and "Weekly Oil Bulletin" in wb.sheetnames:
        ws_w = wb["Weekly Oil Bulletin"]
        w_row = find_weekly_row(ws_w)

        if w_row:
            log("Excel", f"Writing weekly EU data to row {w_row}")
            today = datetime.now().date()
            monday = today - timedelta(days=today.weekday())
            ws_w.cell(row=w_row, column=1).value = datetime.combine(monday, datetime.min.time())

            col_map = {"LT": 2, "LV": 3, "EE": 4, "DK": 5, "SE": 6, "FI": 7, "EU_AVG": 8}
            for key, col in col_map.items():
                val = eu_bulletin.get(key)
                if val is not None:
                    ws_w.cell(row=w_row, column=col).value = val
                    log("Excel", f"  {get_column_letter_simple(col)}{w_row} = {val} ({key})")

    wb.save(str(EXCEL_PATH))
    log("Excel", f"Saved: {EXCEL_PATH}")
    return True


def get_column_letter_simple(col):
    return chr(64 + col) if col <= 26 else f"A{chr(64 + col - 26)}"


# ═══════════════════════════════════════════════
# MAIN
# ═══════════════════════════════════════════════
def main():
    print(f"\n{'='*60}")
    print(f"  FUEL PRICE TRACKER — {TODAY}")
    print(f"{'='*60}\n")

    results = {}

    # Always fetch FX rates (daily)
    print("── Fetching FX rates ──")
    results["fx"] = fetch_fx_rates()

    # Daily sources (skip weekends)
    if WEEKDAY < 5:
        print("\n── Fetching Orlen PL ──")
        results["orlen_pl"] = fetch_orlen_pl()

        print("\n── Fetching Orlen LT ──")
        results["orlen_lt"] = fetch_orlen_lt()

        print("\n── Fetching Elvis DE ──")
        results["elvis_de"] = fetch_elvis_de()

        print("\n── Fetching BSH/ST1 SE ──")
        results["bsh_se"] = fetch_bsh_se()
    else:
        print("\n── Weekend — skipping daily wholesale sources ──")

    # Weekly EU bulletin (Monday only)
    if WEEKDAY == 0:
        print("\n── Fetching EU Weekly Oil Bulletin (Monday) ──")
        results["eu_bulletin"] = fetch_eu_bulletin()
    else:
        results["eu_bulletin"] = None

    # Summary
    print(f"\n{'─'*60}")
    print("RESULTS SUMMARY:")
    succeeded = sum(1 for v in results.values() if v is not None)
    total = sum(1 for k, v in results.items() if k != "eu_bulletin" or WEEKDAY == 0)
    print(f"  {succeeded}/{total} sources fetched successfully")
    for k, v in results.items():
        status = "✅" if v else "❌"
        print(f"  {status} {k}: {v if v else 'FAILED'}")
    print(f"{'─'*60}\n")

    # Write to Excel
    print("── Updating Excel ──")
    success = update_excel(
        fx=results.get("fx"),
        orlen_pl=results.get("orlen_pl"),
        orlen_lt=results.get("orlen_lt"),
        elvis_de=results.get("elvis_de"),
        bsh_se=results.get("bsh_se"),
        eu_bulletin=results.get("eu_bulletin"),
    )

    if success:
        print("\n✅ Excel updated successfully!")
    else:
        print("\n❌ Excel update failed!")
        sys.exit(1)

    # Write results to JSON for GitHub Actions summary
    summary_path = Path(__file__).parent.parent / "latest_results.json"
    summary = {
        "date": TODAY,
        "results": {k: bool(v) for k, v in results.items()},
        "data": {k: v for k, v in results.items() if v is not None},
    }
    summary_path.write_text(json.dumps(summary, indent=2, default=str))
    print(f"Summary written to {summary_path}")


if __name__ == "__main__":
    main()
