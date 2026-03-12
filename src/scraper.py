"""
Fuel Price Tracker — Automated Data Collection v2
Inserts new rows into Excel tracker with fetched data.
"""

import requests
import json
import re
import os
import sys
from datetime import datetime, timedelta
from bs4 import BeautifulSoup
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from pathlib import Path

EXCEL_PATH = Path(__file__).parent.parent / "fuel_tracker.xlsx"
HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "en-US,en;q=0.9,lt;q=0.8,pl;q=0.7,de;q=0.6",
}

TODAY = datetime.now()
TODAY_STR = TODAY.strftime("%Y-%m-%d")
WEEKDAY = TODAY.weekday()

def log(src, msg, lvl="INFO"):
    print(f"[{lvl}] {src}: {msg}")


# ═══════════════════════════════════════
# 1. FX RATES
# ═══════════════════════════════════════
def fetch_fx():
    try:
        r = requests.get("https://api.frankfurter.app/latest?from=EUR&to=PLN,SEK", timeout=15)
        r.raise_for_status()
        data = r.json()
        rates = data.get("rates", {})
        log("FX", f"PLN={rates.get('PLN')}, SEK={rates.get('SEK')}")
        return {"PLN_EUR": rates.get("PLN"), "SEK_EUR": rates.get("SEK")}
    except Exception as e:
        log("FX", f"Failed: {e}", "ERROR")
        return None


# ═══════════════════════════════════════
# 2. ORLEN PL
# ═══════════════════════════════════════
def fetch_orlen_pl():
    try:
        # Try the archive page with XLS download links
        url = "https://www.orlen.pl/pl/dla-biznesu/hurtowe-ceny-paliw"
        r = requests.get(url, headers=HEADERS, timeout=20, allow_redirects=True)

        if r.status_code != 200:
            log("Orlen PL", f"HTTP {r.status_code}", "WARN")
            return None

        soup = BeautifulSoup(r.text, "html.parser")

        # Method 1: Look for JSON data embedded in page
        scripts = soup.find_all("script")
        for script in scripts:
            txt = script.string or ""
            if "ekodiesel" in txt.lower() or "diesel" in txt.lower():
                # Try to find price in JSON structures
                matches = re.findall(r'"(?:price|value|cena)":\s*"?(\d+[.,]\d+)"?', txt, re.IGNORECASE)
                for m in matches:
                    price = float(m.replace(",", "."))
                    if 3000 < price < 9000:
                        log("Orlen PL", f"Found in script: {price} PLN/m³")
                        return {"price_pln_m3": price}

        # Method 2: Parse visible text
        text = soup.get_text(" ", strip=True)
        # Look for Ekodiesel followed by a price
        for pat in [
            r'[Ee]kodiesel[^0-9]{0,80}?(\d[\d\s]{2,6}[.,]\d{2})',
            r'(\d{4,5}[.,]\d{2})[^0-9]{0,30}[Ee]kodiesel',
        ]:
            m = re.search(pat, text)
            if m:
                price = float(m.group(1).replace(" ", "").replace(",", "."))
                if 3000 < price < 9000:
                    log("Orlen PL", f"Parsed: {price} PLN/m³")
                    return {"price_pln_m3": price}

        # Method 3: Try API endpoints
        for api in [
            "https://www.orlen.pl/services/fuel-prices/wholesale",
            "https://www.orlen.pl/api/wholesale-fuel-prices",
            "https://www.orlen.pl/content/dam/orlenpl/fuel-prices/wholesale.json",
        ]:
            try:
                r2 = requests.get(api, headers=HEADERS, timeout=10)
                if r2.status_code == 200 and r2.headers.get("content-type", "").startswith("application/json"):
                    data = r2.json()
                    # Search recursively for diesel price
                    price = find_diesel_price(data, ["ekodiesel", "diesel", "on"])
                    if price and 3000 < price < 9000:
                        log("Orlen PL", f"API: {price} PLN/m³")
                        return {"price_pln_m3": price}
            except:
                continue

        log("Orlen PL", "Could not extract price — site likely JS-rendered", "WARN")
        return None
    except Exception as e:
        log("Orlen PL", f"Failed: {e}", "ERROR")
        return None


def find_diesel_price(obj, keywords, depth=0):
    """Recursively search JSON for diesel price"""
    if depth > 5:
        return None
    if isinstance(obj, dict):
        name_fields = [str(obj.get(k, "")).lower() for k in ["name", "productName", "product", "fuel", "label"]]
        if any(kw in nf for nf in name_fields for kw in keywords):
            for k in ["price", "value", "netPrice", "grossPrice", "cena"]:
                if k in obj and obj[k]:
                    try:
                        return float(str(obj[k]).replace(",", "."))
                    except:
                        pass
        for v in obj.values():
            result = find_diesel_price(v, keywords, depth + 1)
            if result:
                return result
    elif isinstance(obj, list):
        for item in obj:
            result = find_diesel_price(item, keywords, depth + 1)
            if result:
                return result
    return None


# ═══════════════════════════════════════
# 3. ORLEN LT
# ═══════════════════════════════════════
def fetch_orlen_lt():
    try:
        url = "https://www.orlenlietuva.lt/LT/Wholesale/Puslapiai/Kainu-protokolai.aspx"
        r = requests.get(url, headers=HEADERS, timeout=20)

        if r.status_code != 200:
            log("Orlen LT", f"HTTP {r.status_code}", "WARN")
            return None

        soup = BeautifulSoup(r.text, "html.parser")

        # Look for tables with price data
        tables = soup.find_all("table")
        for table in tables:
            text = table.get_text(" ", strip=True).lower()
            if "dyzelinas" in text or "diesel" in text:
                cells = table.find_all("td")
                for i, cell in enumerate(cells):
                    ct = cell.get_text(strip=True).lower()
                    if "dyzelinas" in ct and "rrme" in ct:
                        # Look at next cells for price
                        for j in range(i+1, min(i+6, len(cells))):
                            try:
                                val = float(cells[j].get_text(strip=True).replace(",", ".").replace(" ", ""))
                                if 0.5 < val < 3.0:
                                    log("Orlen LT", f"Table: {val} EUR/l")
                                    return {"price_eur_l": val}
                                elif 500 < val < 3000:
                                    log("Orlen LT", f"Table: {val} EUR/t")
                                    return {"price_eur_l": val / 1000 * 1.19}
                            except:
                                continue

        # Try to find PDF links and note them
        pdf_links = []
        for a in soup.find_all("a", href=True):
            href = a["href"]
            if ".pdf" in href.lower():
                pdf_links.append(href)

        if pdf_links:
            log("Orlen LT", f"Found {len(pdf_links)} PDFs but cannot parse them in this environment", "WARN")

        # Fallback: look in page text
        text = soup.get_text(" ", strip=True)
        for pat in [
            r'[Dd]yzelinas\s+E[^0-9]{0,60}?(\d+[.,]\d{2,4})',
            r'su\s+(?:akcizu|PVM)[^0-9]{0,40}?(\d+[.,]\d{2,4})',
        ]:
            m = re.search(pat, text)
            if m:
                val = float(m.group(1).replace(",", "."))
                if 0.5 < val < 3.0:
                    log("Orlen LT", f"Text: {val} EUR/l")
                    return {"price_eur_l": val}

        log("Orlen LT", "Could not extract — likely PDF-only", "WARN")
        return None
    except Exception as e:
        log("Orlen LT", f"Failed: {e}", "ERROR")
        return None


# ═══════════════════════════════════════
# 4. ELVIS DE
# ═══════════════════════════════════════
def fetch_elvis_de():
    try:
        # Try Tankerkönig API first (free, official German fuel price API)
        # Fallback to mehr-tanken.de
        url = "https://www.mehr-tanken.de/aktuelle-spritpreise/"
        r = requests.get(url, headers=HEADERS, timeout=20)
        r.raise_for_status()
        soup = BeautifulSoup(r.text, "html.parser")
        text = soup.get_text(" ", strip=True)

        for pat in [
            r'[Dd]iesel[^0-9]{0,40}?(\d[.,]\d{2,3})\s*(?:€|EUR|Euro)',
            r'[Dd]iesel[^0-9]{0,40}?(\d[.,]\d{2,3})',
            r'(\d[.,]\d{2,3})\s*(?:€|EUR)[^0-9]{0,20}[Dd]iesel',
        ]:
            m = re.search(pat, text)
            if m:
                price = float(m.group(1).replace(",", "."))
                if 0.8 < price < 3.0:
                    log("Elvis DE", f"Diesel={price} EUR/l")
                    return {"price_eur_l": price}

        # Try finding in structured data / JSON-LD
        for script in soup.find_all("script", {"type": "application/ld+json"}):
            try:
                data = json.loads(script.string)
                price = find_diesel_price(data, ["diesel"])
                if price and 0.8 < price < 3.0:
                    log("Elvis DE", f"JSON-LD: {price} EUR/l")
                    return {"price_eur_l": price}
            except:
                continue

        log("Elvis DE", "Could not extract", "WARN")
        return None
    except Exception as e:
        log("Elvis DE", f"Failed: {e}", "ERROR")
        return None


# ═══════════════════════════════════════
# 5. BSH / ST1 SE
# ═══════════════════════════════════════
def fetch_bsh_se():
    try:
        url = "https://st1.se/foretag/listpris"
        r = requests.get(url, headers=HEADERS, timeout=20)
        r.raise_for_status()
        soup = BeautifulSoup(r.text, "html.parser")
        text = soup.get_text(" ", strip=True)

        for pat in [
            r'[Dd]iesel\s*(?:MK[123])?\s*[^0-9]{0,30}?(\d{1,2}[.,]\d{2})',
            r'(\d{1,2}[.,]\d{2})\s*(?:kr|SEK)[^0-9]{0,20}[Dd]iesel',
        ]:
            m = re.search(pat, text)
            if m:
                price = float(m.group(1).replace(",", "."))
                if 10 < price < 35:
                    log("BSH SE", f"Diesel={price} SEK/l")
                    return {"price_sek_l": price}

        log("BSH SE", "Could not extract", "WARN")
        return None
    except Exception as e:
        log("BSH SE", f"Failed: {e}", "ERROR")
        return None


# ═══════════════════════════════════════
# 6. EU WEEKLY OIL BULLETIN
# ═══════════════════════════════════════
def fetch_eu_bulletin():
    try:
        # Try the direct data download endpoint
        base = "https://energy.ec.europa.eu"
        page_url = f"{base}/data-and-analysis/weekly-oil-bulletin_en"
        r = requests.get(page_url, headers=HEADERS, timeout=20)
        r.raise_for_status()
        soup = BeautifulSoup(r.text, "html.parser")

        # Find download links for CSV/XLS
        download_links = []
        for a in soup.find_all("a", href=True):
            href = a["href"].lower()
            text = a.get_text(strip=True).lower()
            if any(kw in href for kw in [".csv", ".xls", "download", "export"]):
                download_links.append(a["href"])
            elif any(kw in text for kw in ["download", "csv", "excel", "prices with taxes"]):
                download_links.append(a["href"])

        for link in download_links:
            full_url = link if link.startswith("http") else base + link
            try:
                r2 = requests.get(full_url, headers=HEADERS, timeout=20)
                if r2.status_code == 200:
                    result = parse_eu_csv(r2.text)
                    if result:
                        return result
            except:
                continue

        # Fallback: parse from HTML tables on page
        tables = soup.find_all("table")
        for table in tables:
            result = parse_eu_table(table)
            if result:
                return result

        log("EU Bulletin", "Could not find data", "WARN")
        return None
    except Exception as e:
        log("EU Bulletin", f"Failed: {e}", "ERROR")
        return None


def parse_eu_csv(text):
    countries = {"LT": None, "LV": None, "EE": None, "DK": None, "SE": None, "FI": None}
    eu_avg = None
    cc_names = {"Lithuania": "LT", "Latvia": "LV", "Estonia": "EE", "Denmark": "DK", "Sweden": "SE", "Finland": "FI"}

    for line in text.split("\n"):
        parts = [p.strip() for p in (line.split(",") if "," in line else line.split("\t"))]
        for i, part in enumerate(parts):
            clean = part.strip('"').strip()
            # Check country codes and names
            matched_cc = None
            if clean.upper() in countries:
                matched_cc = clean.upper()
            for name, cc in cc_names.items():
                if name.lower() in clean.lower():
                    matched_cc = cc
                    break

            if matched_cc:
                for j in range(i+1, min(i+8, len(parts))):
                    try:
                        val = float(parts[j].strip().strip('"').replace(",", "."))
                        if 0.5 < val < 3.5:
                            countries[matched_cc] = val
                            break
                    except:
                        continue

            if "eu" in clean.lower() and any(kw in clean.lower() for kw in ["avg", "average", "weighted", "mean"]):
                for j in range(i+1, min(i+8, len(parts))):
                    try:
                        val = float(parts[j].strip().strip('"').replace(",", "."))
                        if 0.5 < val < 3.5:
                            eu_avg = val
                            break
                    except:
                        continue

    if any(v for v in countries.values() if v is not None):
        log("EU Bulletin", f"CSV: {countries}, avg={eu_avg}")
        return {**countries, "EU_AVG": eu_avg}
    return None


def parse_eu_table(table):
    countries = {"LT": None, "LV": None, "EE": None, "DK": None, "SE": None, "FI": None}
    cc_names = {"Lithuania": "LT", "Latvia": "LV", "Estonia": "EE", "Denmark": "DK", "Sweden": "SE", "Finland": "FI"}

    rows = table.find_all("tr")
    for row in rows:
        cells = [td.get_text(strip=True) for td in row.find_all(["td", "th"])]
        for i, cell in enumerate(cells):
            matched_cc = None
            if cell.upper() in countries:
                matched_cc = cell.upper()
            for name, cc in cc_names.items():
                if name.lower() in cell.lower():
                    matched_cc = cc
                    break

            if matched_cc:
                for j in range(i+1, min(i+6, len(cells))):
                    try:
                        val = float(cells[j].replace(",", ".").replace(" ", ""))
                        if 0.5 < val < 3.5:
                            countries[matched_cc] = val
                            break
                    except:
                        continue

    if any(v for v in countries.values() if v is not None):
        log("EU Bulletin", f"HTML table: {countries}")
        return countries
    return None


# ═══════════════════════════════════════
# EXCEL WRITER — INSERT ROW
# ═══════════════════════════════════════
def update_excel(fx=None, orlen_pl=None, orlen_lt=None, elvis_de=None, bsh_se=None, eu_bulletin=None):
    if not EXCEL_PATH.exists():
        log("Excel", f"Not found: {EXCEL_PATH}", "ERROR")
        return False

    wb = load_workbook(str(EXCEL_PATH))

    # ─── DAILY TRACKER: insert new row at position 5 ───
    if "Daily Tracker" in wb.sheetnames:
        ws = wb["Daily Tracker"]

        # Shift rows down: insert a blank row at position 5
        ws.insert_rows(5)

        # Date
        ws.cell(row=5, column=1).value = TODAY
        ws.cell(row=5, column=1).number_format = 'YYYY-MM-DD'
        ws.cell(row=5, column=1).font = Font(name="Aptos", size=10, bold=True, color="1F2937")
        ws.cell(row=5, column=1).alignment = Alignment(horizontal="center", vertical="center")

        # Day name
        import calendar
        ws.cell(row=5, column=2).value = calendar.day_abbr[WEEKDAY]
        ws.cell(row=5, column=2).font = Font(name="Aptos", size=9, color="6B7280")
        ws.cell(row=5, column=2).alignment = Alignment(horizontal="center", vertical="center")

        input_font = Font(name="Aptos", size=10, color="1D4ED8")
        input_fill = PatternFill("solid", fgColor="EFF6FF")
        data_font = Font(name="Aptos", size=10, color="1F2937")
        brd = Border(
            left=Side("thin", color="D1D5DB"), right=Side("thin", color="D1D5DB"),
            top=Side("thin", color="D1D5DB"), bottom=Side("thin", color="D1D5DB"))

        # Orlen PL — col C (PLN/m³)
        if orlen_pl and orlen_pl.get("price_pln_m3"):
            ws.cell(row=5, column=3).value = orlen_pl["price_pln_m3"]
            ws.cell(row=5, column=3).number_format = '#,##0.00'
        ws.cell(row=5, column=3).font = input_font
        ws.cell(row=5, column=3).fill = input_fill
        ws.cell(row=5, column=3).border = brd

        # PLN/EUR — col D
        if fx and fx.get("PLN_EUR"):
            ws.cell(row=5, column=4).value = fx["PLN_EUR"]
        ws.cell(row=5, column=4).number_format = '0.0000'
        ws.cell(row=5, column=4).font = input_font
        ws.cell(row=5, column=4).fill = input_fill
        ws.cell(row=5, column=4).border = brd

        # EUR/l formula — col E
        ws.cell(row=5, column=5).value = '=IF(AND(C5<>"",D5<>"",D5<>0),C5/D5/1000,"")'
        ws.cell(row=5, column=5).number_format = '0.000'
        ws.cell(row=5, column=5).font = data_font
        ws.cell(row=5, column=5).border = brd

        # Orlen LT — col F
        if orlen_lt and orlen_lt.get("price_eur_l"):
            ws.cell(row=5, column=6).value = orlen_lt["price_eur_l"]
        ws.cell(row=5, column=6).number_format = '0.000'
        ws.cell(row=5, column=6).font = input_font
        ws.cell(row=5, column=6).fill = input_fill
        ws.cell(row=5, column=6).border = brd

        # Delta EUR/l — col G
        ws.cell(row=5, column=7).value = '=IF(AND(E5<>"",F5<>""),E5-F5,"")'
        ws.cell(row=5, column=7).number_format = '+0.000;-0.000;"-"'
        ws.cell(row=5, column=7).font = data_font
        ws.cell(row=5, column=7).border = brd

        # Delta % — col H
        ws.cell(row=5, column=8).value = '=IF(AND(E5<>"",F5<>"",F5<>0),(E5-F5)/F5,"")'
        ws.cell(row=5, column=8).number_format = '+0.0%;-0.0%;"-"'
        ws.cell(row=5, column=8).font = data_font
        ws.cell(row=5, column=8).border = brd

        # Elvis DE — col I
        if elvis_de and elvis_de.get("price_eur_l"):
            ws.cell(row=5, column=9).value = elvis_de["price_eur_l"]
        ws.cell(row=5, column=9).number_format = '0.000'
        ws.cell(row=5, column=9).font = input_font
        ws.cell(row=5, column=9).fill = input_fill
        ws.cell(row=5, column=9).border = brd

        # BSH SE SEK — col J
        if bsh_se and bsh_se.get("price_sek_l"):
            ws.cell(row=5, column=10).value = bsh_se["price_sek_l"]
        ws.cell(row=5, column=10).number_format = '0.00'
        ws.cell(row=5, column=10).font = input_font
        ws.cell(row=5, column=10).fill = input_fill
        ws.cell(row=5, column=10).border = brd

        # SEK/EUR — col K
        if fx and fx.get("SEK_EUR"):
            ws.cell(row=5, column=11).value = fx["SEK_EUR"]
        ws.cell(row=5, column=11).number_format = '0.0000'
        ws.cell(row=5, column=11).font = input_font
        ws.cell(row=5, column=11).fill = input_fill
        ws.cell(row=5, column=11).border = brd

        # BSH EUR formula — col L
        ws.cell(row=5, column=12).value = '=IF(AND(J5<>"",K5<>"",K5<>0),J5/K5,"")'
        ws.cell(row=5, column=12).number_format = '0.000'
        ws.cell(row=5, column=12).font = data_font
        ws.cell(row=5, column=12).border = brd

        # Notes — col M
        sources_ok = []
        if fx: sources_ok.append("FX")
        if orlen_pl: sources_ok.append("PL")
        if orlen_lt: sources_ok.append("LT")
        if elvis_de: sources_ok.append("DE")
        if bsh_se: sources_ok.append("SE")
        ws.cell(row=5, column=13).value = f"Auto: {','.join(sources_ok)}"
        ws.cell(row=5, column=13).font = Font(name="Aptos", size=9, color="6B7280")
        ws.cell(row=5, column=13).border = brd

        ws.cell(row=5, column=14).value = "Auto"
        ws.cell(row=5, column=14).font = Font(name="Aptos", size=8, color="9CA3AF")
        ws.cell(row=5, column=14).alignment = Alignment(horizontal="center")
        ws.cell(row=5, column=14).border = brd

        log("Excel", f"Inserted daily row 5: {TODAY_STR}")

    # ─── WEEKLY OIL BULLETIN: insert on Mondays ───
    if eu_bulletin and WEEKDAY == 0 and "Weekly Oil Bulletin" in wb.sheetnames:
        ws_w = wb["Weekly Oil Bulletin"]
        ws_w.insert_rows(4)

        monday = TODAY - timedelta(days=TODAY.weekday())
        ws_w.cell(row=4, column=1).value = monday
        ws_w.cell(row=4, column=1).number_format = 'YYYY-MM-DD'
        ws_w.cell(row=4, column=1).font = Font(name="Aptos", size=10, bold=True, color="1F2937")

        col_map = {"LT": 2, "LV": 3, "EE": 4, "DK": 5, "SE": 6, "FI": 7, "EU_AVG": 8}
        for key, col in col_map.items():
            val = eu_bulletin.get(key)
            if val is not None:
                ws_w.cell(row=4, column=col).value = val
                ws_w.cell(row=4, column=col).number_format = '0.000'
                ws_w.cell(row=4, column=col).font = Font(name="Aptos", size=10, color="1D4ED8")

        # LT vs EU formula
        ws_w.cell(row=4, column=9).value = '=IF(AND(B4<>"",H4<>"",H4<>0),(B4-H4)/H4,"")'
        ws_w.cell(row=4, column=9).number_format = '+0.0%;-0.0%;"-"'

        log("Excel", f"Inserted weekly row 4: {monday.strftime('%Y-%m-%d')}")

    wb.save(str(EXCEL_PATH))
    log("Excel", f"Saved: {EXCEL_PATH}")
    return True


# ═══════════════════════════════════════
# MAIN
# ═══════════════════════════════════════
def main():
    print(f"\n{'='*60}")
    print(f"  FUEL PRICE TRACKER — {TODAY_STR}")
    print(f"{'='*60}\n")

    results = {}

    print("── FX Rates ──")
    results["fx"] = fetch_fx()

    if WEEKDAY < 5:
        print("\n── Orlen PL ──")
        results["orlen_pl"] = fetch_orlen_pl()
        print("\n── Orlen LT ──")
        results["orlen_lt"] = fetch_orlen_lt()
        print("\n── Elvis DE ──")
        results["elvis_de"] = fetch_elvis_de()
        print("\n── BSH/ST1 SE ──")
        results["bsh_se"] = fetch_bsh_se()
    else:
        print("\n── Weekend — skipping daily sources ──")
        for k in ["orlen_pl", "orlen_lt", "elvis_de", "bsh_se"]:
            results[k] = None

    if WEEKDAY == 0:
        print("\n── EU Oil Bulletin (Monday) ──")
        results["eu_bulletin"] = fetch_eu_bulletin()
    else:
        results["eu_bulletin"] = None

    # Summary
    print(f"\n{'─'*60}")
    succeeded = sum(1 for v in results.values() if v is not None)
    total = sum(1 for k in results if k != "eu_bulletin" or WEEKDAY == 0)
    print(f"RESULTS: {succeeded}/{total} sources OK")
    for k, v in results.items():
        print(f"  {'✅' if v else '❌'} {k}")
    print(f"{'─'*60}\n")

    print("── Updating Excel ──")
    ok = update_excel(**results)

    # Write JSON summary
    summary = {
        "date": TODAY_STR,
        "results": {k: bool(v) for k, v in results.items()},
        "data": {k: v for k, v in results.items() if v is not None},
    }
    Path(EXCEL_PATH.parent / "latest_results.json").write_text(json.dumps(summary, indent=2, default=str))

    if ok:
        print("\n✅ Done!")
    else:
        print("\n❌ Failed!")
        sys.exit(1)


if __name__ == "__main__":
    main()
