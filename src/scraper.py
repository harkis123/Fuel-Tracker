"""
Fuel Price Tracker v3 — All sources working
"""
import requests, json, re, os, sys, io, calendar
from datetime import datetime, timedelta
from bs4 import BeautifulSoup
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from pathlib import Path

EXCEL_PATH = Path(__file__).parent.parent / "fuel_tracker.xlsx"
UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/131.0.0.0 Safari/537.36"
H = {"User-Agent": UA, "Accept": "text/html,application/xhtml+xml,*/*;q=0.8", "Accept-Language": "en-US,en;q=0.9,lt;q=0.8,pl;q=0.7"}
TODAY = datetime.now()
TODAY_STR = TODAY.strftime("%Y-%m-%d")
WDAY = TODAY.weekday()

def log(s, m, l="INFO"): print(f"[{l}] {s}: {m}")

# ═══════════════════════════════════════
# 1. FX RATES — frankfurter.app
# ═══════════════════════════════════════
def fetch_fx():
    try:
        r = requests.get("https://api.frankfurter.app/latest?from=EUR&to=PLN,SEK", timeout=15)
        r.raise_for_status()
        d = r.json().get("rates", {})
        log("FX", f"PLN={d.get('PLN')}, SEK={d.get('SEK')}")
        return {"PLN_EUR": d.get("PLN"), "SEK_EUR": d.get("SEK")}
    except Exception as e:
        log("FX", str(e), "ERROR")
        return None

# ═══════════════════════════════════════
# 2. ORLEN PL — via petrodom.pl (mirrors Orlen wholesale)
# ═══════════════════════════════════════
def fetch_orlen_pl():
    try:
        url = "https://www.petrodom.pl/en/current-wholesale-fuel-prices-provided-by-pkn-orlen/"
        r = requests.get(url, headers=H, timeout=20)
        r.raise_for_status()
        soup = BeautifulSoup(r.text, "html.parser")
        
        # Find table with fuel prices
        for table in soup.find_all("table"):
            rows = table.find_all("tr")
            for row in rows:
                cells = [td.get_text(strip=True) for td in row.find_all(["td", "th"])]
                for i, cell in enumerate(cells):
                    if "ekodiesel" in cell.lower():
                        # Next cell should be the price
                        for j in range(i+1, min(i+3, len(cells))):
                            price_str = cells[j].replace(" ", "").replace(",", ".").replace("\xa0", "")
                            try:
                                price = float(re.sub(r'[^\d.]', '', price_str))
                                if 3000 < price < 10000:
                                    log("Orlen PL", f"Ekodiesel = {price} PLN/m³ (via petrodom.pl)")
                                    return {"price_pln_m3": price}
                            except:
                                continue
        
        log("Orlen PL", "Could not find Ekodiesel in table", "WARN")
        return None
    except Exception as e:
        log("Orlen PL", str(e), "ERROR")
        return None

# ═══════════════════════════════════════
# 3. ORLEN LT — PDF price protocol parsing
# ═══════════════════════════════════════
def fetch_orlen_lt():
    try:
        # First get the list of available protocols
        list_url = "https://www.orlenlietuva.lt/lt/wholesale/_layouts/f2hPriceTable/default.aspx"
        r = requests.get(list_url, headers=H, timeout=20)
        r.raise_for_status()
        soup = BeautifulSoup(r.text, "html.parser")
        
        # Find PDF links
        pdf_links = []
        for a in soup.find_all("a", href=True):
            href = a["href"]
            if ".pdf" in href.lower() and "kainos" in href.lower():
                if not href.startswith("http"):
                    href = "https://www.orlenlietuva.lt" + href
                pdf_links.append(href)
        
        if not pdf_links:
            # Try constructing URL from today's date
            for days_back in range(5):
                d = TODAY - timedelta(days=days_back)
                ds = d.strftime("%Y %m %d")
                pdf_url = f"https://www.orlenlietuva.lt/LT/Wholesale/Prices/Kainos {ds} realizacija internet.pdf"
                pdf_links.append(pdf_url)
        
        # Try to download and parse the latest PDF
        for pdf_url in pdf_links[:3]:
            try:
                r2 = requests.get(pdf_url, headers=H, timeout=15)
                if r2.status_code == 200 and len(r2.content) > 500:
                    price = parse_orlen_lt_pdf(r2.content)
                    if price:
                        return price
            except:
                continue
        
        log("Orlen LT", "Could not download/parse any PDF", "WARN")
        return None
    except Exception as e:
        log("Orlen LT", str(e), "ERROR")
        return None


def parse_orlen_lt_pdf(pdf_bytes):
    """Extract diesel price from Orlen LT PDF protocol"""
    try:
        # Try using pdfplumber if available
        try:
            import pdfplumber
            with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
                for page in pdf.pages:
                    text = page.extract_text() or ""
                    price = find_lt_diesel_price(text)
                    if price:
                        return price
                    # Also try tables
                    tables = page.extract_tables()
                    for table in tables:
                        for row in table:
                            row_text = " ".join([str(c) for c in row if c])
                            price = find_lt_diesel_price(row_text)
                            if price:
                                return price
        except ImportError:
            pass
        
        # Fallback: extract text from PDF bytes directly (basic)
        text = pdf_bytes.decode("latin-1", errors="ignore")
        price = find_lt_diesel_price(text)
        if price:
            return price
        
        log("Orlen LT", "PDF downloaded but could not extract price", "WARN")
        return None
    except Exception as e:
        log("Orlen LT", f"PDF parse error: {e}", "WARN")
        return None


def find_lt_diesel_price(text):
    """Find diesel price in Orlen LT protocol text"""
    text_lower = text.lower()
    
    # Look for "Dyzelinas E" with RRME, then find price
    patterns = [
        # Dyzelinas E kl. su RRME ... price with PVM
        r'dyzelinas\s+e[^0-9]{0,100}?su\s+(?:akcizu|pvm)[^0-9]{0,50}?(\d+[.,]\d{2,4})',
        r'dyzelinas\s+e[^0-9]{0,100}?(\d+[.,]\d{2,4})\s*(?:eur|€)',
        # General diesel price
        r'dyzelinas[^0-9]{0,80}?(\d+[.,]\d{2,4})',
        # Price near "su PVM" or "su akcizu"
        r'su\s+pvm[^0-9]{0,30}?(\d+[.,]\d{2,4})',
        # Any price that looks like EUR/l in the diesel section
        r'(\d{1}[.,]\d{3,4})\s',
    ]
    
    for pat in patterns:
        matches = re.findall(pat, text_lower)
        for m in matches:
            try:
                val = float(m.replace(",", "."))
                # EUR/l should be 0.8-3.0
                if 0.8 < val < 3.0:
                    log("Orlen LT", f"Diesel E = {val} EUR/l (from PDF)")
                    return {"price_eur_l": val}
                # EUR/t would be 800-3000
                elif 800 < val < 3000:
                    eur_l = val / 1000 * 1.19  # rough density conversion
                    log("Orlen LT", f"Diesel E = {val} EUR/t → {eur_l:.3f} EUR/l")
                    return {"price_eur_l": round(eur_l, 4)}
            except:
                continue
    return None

# ═══════════════════════════════════════
# 4. ELVIS FSC DE — mehr-tanken.de
# ═══════════════════════════════════════
def fetch_elvis_de():
    try:
        r = requests.get("https://www.mehr-tanken.de/aktuelle-spritpreise/", headers=H, timeout=20)
        r.raise_for_status()
        soup = BeautifulSoup(r.text, "html.parser")
        text = soup.get_text(" ", strip=True)
        for pat in [
            r'[Dd]iesel[^0-9]{0,40}?(\d[.,]\d{2,3})\s*(?:€|EUR|Euro)',
            r'[Dd]iesel[^0-9]{0,40}?(\d[.,]\d{2,3})',
        ]:
            m = re.search(pat, text)
            if m:
                p = float(m.group(1).replace(",", "."))
                if 0.8 < p < 3.0:
                    log("Elvis DE", f"Diesel={p} EUR/l")
                    return {"price_eur_l": p}
        log("Elvis DE", "Not found", "WARN")
        return None
    except Exception as e:
        log("Elvis DE", str(e), "ERROR")
        return None

# ═══════════════════════════════════════
# 5. BSH / ST1 SE
# ═══════════════════════════════════════
def fetch_bsh_se():
    try:
        r = requests.get("https://st1.se/foretag/listpris", headers=H, timeout=20)
        r.raise_for_status()
        soup = BeautifulSoup(r.text, "html.parser")
        text = soup.get_text(" ", strip=True)
        for pat in [
            r'[Dd]iesel\s*(?:MK[123])?\s*[^0-9]{0,30}?(\d{1,2}[.,]\d{2})',
            r'(\d{1,2}[.,]\d{2})\s*(?:kr|SEK)[^0-9]{0,20}[Dd]iesel',
        ]:
            m = re.search(pat, text)
            if m:
                p = float(m.group(1).replace(",", "."))
                if 10 < p < 35:
                    log("BSH SE", f"Diesel={p} SEK/l")
                    return {"price_sek_l": p}
        log("BSH SE", "Not found", "WARN")
        return None
    except Exception as e:
        log("BSH SE", str(e), "ERROR")
        return None

# ═══════════════════════════════════════
# 6. EU WEEKLY OIL BULLETIN — via fuel-prices.eu
# ═══════════════════════════════════════
def fetch_eu_bulletin():
    try:
        # fuel-prices.eu provides EC data in accessible format
        r = requests.get("https://www.fuel-prices.eu/", headers=H, timeout=20)
        r.raise_for_status()
        soup = BeautifulSoup(r.text, "html.parser")
        text = soup.get_text(" ", strip=True)
        
        countries = {"LT": None, "LV": None, "EE": None, "DK": None, "SE": None, "FI": None}
        cc_names = {
            "Lithuania": "LT", "Latvia": "LV", "Estonia": "EE",
            "Denmark": "DK", "Sweden": "SE", "Finland": "FI"
        }
        eu_avg = None
        
        # Try to find diesel prices in tables
        tables = soup.find_all("table")
        for table in tables:
            rows = table.find_all("tr")
            # Check if this table has diesel data
            table_text = table.get_text(" ", strip=True).lower()
            if "diesel" not in table_text and "gasoil" not in table_text:
                continue
            
            for row in rows:
                cells = [td.get_text(strip=True) for td in row.find_all(["td", "th"])]
                if len(cells) < 2:
                    continue
                
                row_text = cells[0]
                for name, cc in cc_names.items():
                    if name.lower() in row_text.lower() or cc == row_text.strip().upper():
                        # Find diesel price in subsequent cells
                        for c in cells[1:]:
                            try:
                                val = float(c.replace(",", ".").replace("€", "").replace(" ", "").strip())
                                if 0.5 < val < 3.5:
                                    countries[cc] = val
                                    break
                            except:
                                continue
        
        # Also try regex on full page text for country prices
        for name, cc in cc_names.items():
            if countries[cc] is None:
                # Pattern: "Lithuania ... 1.567" or "Lithuania €1.567"
                pat = rf'{name}[^0-9]{{0,80}}?[€]?\s*(\d[.,]\d{{2,3}})'
                m = re.search(pat, text, re.IGNORECASE)
                if m:
                    val = float(m.group(1).replace(",", "."))
                    if 0.5 < val < 3.5:
                        countries[cc] = val
        
        # EU average
        for pat in [
            r'EU\s*(?:27|avg|average)[^0-9]{0,30}?[€]?\s*(\d[.,]\d{2,3})',
            r'average[^0-9]{0,30}?[€]?\s*(\d[.,]\d{2,3})',
            r'EU\s*avg[^0-9]{0,10}?(\d[.,]\d{2,3})',
        ]:
            m = re.search(pat, text, re.IGNORECASE)
            if m:
                val = float(m.group(1).replace(",", "."))
                if 0.5 < val < 3.5:
                    eu_avg = val
                    break
        
        found = {k: v for k, v in countries.items() if v is not None}
        if found:
            log("EU Bulletin", f"Found {len(found)} countries: {found}, EU avg={eu_avg}")
            return {**countries, "EU_AVG": eu_avg}
        
        # Fallback: try EC direct PDF
        return fetch_eu_bulletin_pdf()
        
    except Exception as e:
        log("EU Bulletin", str(e), "ERROR")
        return fetch_eu_bulletin_pdf()


def fetch_eu_bulletin_pdf():
    """Fallback: try EC's latest prices PDF"""
    try:
        pdf_url = "https://ec.europa.eu/energy/observatory/reports/latest_prices_with_taxes.pdf"
        r = requests.get(pdf_url, headers=H, timeout=20)
        if r.status_code != 200:
            log("EU Bulletin PDF", f"HTTP {r.status_code}", "WARN")
            return None
        
        # Try pdfplumber
        try:
            import pdfplumber
            countries = {"LT": None, "LV": None, "EE": None, "DK": None, "SE": None, "FI": None}
            cc_names = {"Lithuania": "LT", "Latvia": "LV", "Estonia": "EE", "Denmark": "DK", "Sweden": "SE", "Finland": "FI"}
            
            with pdfplumber.open(io.BytesIO(r.content)) as pdf:
                for page in pdf.pages:
                    tables = page.extract_tables()
                    for table in tables:
                        for row in table:
                            if not row or not row[0]:
                                continue
                            cell0 = str(row[0]).strip()
                            for name, cc in cc_names.items():
                                if name.lower() in cell0.lower():
                                    # Look for diesel price (usually 4th-6th column)
                                    for c in row[1:]:
                                        if c is None:
                                            continue
                                        try:
                                            val = float(str(c).replace(",", ".").replace(" ", ""))
                                            if 0.5 < val < 3.5:
                                                countries[cc] = val
                                                break
                                            elif 500 < val < 3500:
                                                countries[cc] = val / 1000
                                                break
                                        except:
                                            continue
            
            found = {k: v for k, v in countries.items() if v is not None}
            if found:
                log("EU Bulletin PDF", f"Found: {found}")
                return countries
        except ImportError:
            log("EU Bulletin PDF", "pdfplumber not installed", "WARN")
        
        return None
    except Exception as e:
        log("EU Bulletin PDF", str(e), "ERROR")
        return None

# ═══════════════════════════════════════
# EXCEL WRITER
# ═══════════════════════════════════════
def update_excel(fx=None, orlen_pl=None, orlen_lt=None, elvis_de=None, bsh_se=None, eu_bulletin=None):
    if not EXCEL_PATH.exists():
        log("Excel", f"Not found: {EXCEL_PATH}", "ERROR")
        return False
    wb = load_workbook(str(EXCEL_PATH))

    if "Daily Tracker" in wb.sheetnames:
        ws = wb["Daily Tracker"]
        ws.insert_rows(5)
        
        ifont = Font(name="Aptos", size=10, color="1D4ED8")
        ifill = PatternFill("solid", fgColor="EFF6FF")
        dfont = Font(name="Aptos", size=10, color="1F2937")
        brd = Border(left=Side("thin",color="D1D5DB"),right=Side("thin",color="D1D5DB"),top=Side("thin",color="D1D5DB"),bottom=Side("thin",color="D1D5DB"))
        
        def wc(col, val, fmt='General', font=dfont, fill=None):
            c = ws.cell(row=5, column=col)
            c.value = val; c.number_format = fmt; c.font = font; c.border = brd
            c.alignment = Alignment(horizontal="right" if col > 2 else "center", vertical="center")
            if fill: c.fill = fill
        
        wc(1, TODAY, 'YYYY-MM-DD', Font(name="Aptos",size=10,bold=True,color="1F2937"))
        wc(2, calendar.day_abbr[WDAY], font=Font(name="Aptos",size=9,color="6B7280"))
        
        # Orlen PL PLN/m³
        wc(3, orlen_pl["price_pln_m3"] if orlen_pl else None, '#,##0.00', ifont, ifill)
        # PLN/EUR
        wc(4, fx["PLN_EUR"] if fx else None, '0.0000', ifont, ifill)
        # EUR/l formula
        wc(5, '=IF(AND(C5<>"",D5<>"",D5<>0),C5/D5/1000,"")', '0.000')
        # Orlen LT EUR/l
        wc(6, orlen_lt["price_eur_l"] if orlen_lt else None, '0.000', ifont, ifill)
        # Delta EUR/l
        wc(7, '=IF(AND(E5<>"",F5<>""),E5-F5,"")', '+0.000;-0.000;"-"')
        # Delta %
        wc(8, '=IF(AND(E5<>"",F5<>"",F5<>0),(E5-F5)/F5,"")', '+0.0%;-0.0%;"-"')
        # Elvis DE
        wc(9, elvis_de["price_eur_l"] if elvis_de else None, '0.000', ifont, ifill)
        # BSH SEK
        wc(10, bsh_se["price_sek_l"] if bsh_se else None, '0.00', ifont, ifill)
        # SEK/EUR
        wc(11, fx["SEK_EUR"] if fx else None, '0.0000', ifont, ifill)
        # BSH EUR formula
        wc(12, '=IF(AND(J5<>"",K5<>"",K5<>0),J5/K5,"")', '0.000')
        
        ok = [k for k,v in {"FX":fx,"PL":orlen_pl,"LT":orlen_lt,"DE":elvis_de,"SE":bsh_se}.items() if v]
        wc(13, f"Auto: {','.join(ok)}", font=Font(name="Aptos",size=9,color="6B7280"))
        wc(14, "Auto", font=Font(name="Aptos",size=8,color="9CA3AF"))
        
        log("Excel", f"Daily row 5: {TODAY_STR} [{','.join(ok)}]")

    # Weekly
    if eu_bulletin and WDAY == 0 and "Weekly Oil Bulletin" in wb.sheetnames:
        ws_w = wb["Weekly Oil Bulletin"]
        ws_w.insert_rows(4)
        monday = TODAY - timedelta(days=WDAY)
        ws_w.cell(row=4,column=1).value = monday
        ws_w.cell(row=4,column=1).number_format = 'YYYY-MM-DD'
        ws_w.cell(row=4,column=1).font = Font(name="Aptos",size=10,bold=True,color="1F2937")
        
        col_map = {"LT":2,"LV":3,"EE":4,"DK":5,"SE":6,"FI":7,"EU_AVG":8}
        for k,col in col_map.items():
            val = eu_bulletin.get(k)
            if val is not None:
                c = ws_w.cell(row=4,column=col)
                c.value = val; c.number_format = '0.000'
                c.font = Font(name="Aptos",size=10,color="1D4ED8")
        
        ws_w.cell(row=4,column=9).value = '=IF(AND(B4<>"",H4<>"",H4<>0),(B4-H4)/H4,"")'
        ws_w.cell(row=4,column=9).number_format = '+0.0%;-0.0%;"-"'
        log("Excel", f"Weekly row 4: {monday.strftime('%Y-%m-%d')}")

    wb.save(str(EXCEL_PATH))
    log("Excel", f"Saved: {EXCEL_PATH}")
    return True

# ═══════════════════════════════════════
# MAIN
# ═══════════════════════════════════════
def main():
    print(f"\n{'='*60}\n  FUEL PRICE TRACKER v3 — {TODAY_STR}\n{'='*60}\n")
    R = {}
    
    print("── FX Rates ──")
    R["fx"] = fetch_fx()
    
    if WDAY < 5:
        print("\n── Orlen PL (via petrodom.pl) ──")
        R["orlen_pl"] = fetch_orlen_pl()
        print("\n── Orlen LT (PDF protocol) ──")
        R["orlen_lt"] = fetch_orlen_lt()
        print("\n── Elvis DE ──")
        R["elvis_de"] = fetch_elvis_de()
        print("\n── BSH/ST1 SE ──")
        R["bsh_se"] = fetch_bsh_se()
    else:
        print("\n── Weekend — skipping ──")
        for k in ["orlen_pl","orlen_lt","elvis_de","bsh_se"]: R[k] = None
    
    if WDAY == 0:
        print("\n── EU Oil Bulletin (Monday) ──")
        R["eu_bulletin"] = fetch_eu_bulletin()
    else:
        R["eu_bulletin"] = None
    
    print(f"\n{'─'*60}")
    ok = sum(1 for v in R.values() if v is not None)
    total = sum(1 for k in R if k != "eu_bulletin" or WDAY == 0)
    print(f"RESULTS: {ok}/{total}")
    for k,v in R.items(): print(f"  {'✅' if v else '❌'} {k}: {v if v else 'FAILED'}")
    print(f"{'─'*60}\n")
    
    print("── Updating Excel ──")
    success = update_excel(**R)
    
    Path(EXCEL_PATH.parent / "latest_results.json").write_text(
        json.dumps({"date":TODAY_STR,"results":{k:bool(v) for k,v in R.items()},"data":{k:v for k,v in R.items() if v}}, indent=2, default=str))
    
    print(f"\n{'✅ Done!' if success else '❌ Failed!'}")
    if not success: sys.exit(1)

if __name__ == "__main__":
    main()
