"""
Fuel Price Tracker v4 — FINAL — All corrections applied
- Orlen LT: pardavimo kaina su PVM (EUR/1000l ÷ 1000)
- Elvis DE: Diesel from fuel-prices.eu (not Super E5 from mehr-tanken)
- EU Bulletin: All country diesel prices from fuel-prices.eu/cheapest/
- BSH SE: st1.se diesel SEK/l
"""
import requests, json, re, os, sys, io, calendar
from datetime import datetime, timedelta
from bs4 import BeautifulSoup
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from pathlib import Path

EXCEL_PATH = Path(__file__).parent.parent / "fuel_tracker.xlsx"
UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/131.0.0.0 Safari/537.36"
H = {"User-Agent": UA, "Accept": "text/html,application/xhtml+xml,*/*;q=0.8", "Accept-Language": "en-US,en;q=0.9"}
TODAY = datetime.now()
TODAY_STR = TODAY.strftime("%Y-%m-%d")
WDAY = TODAY.weekday()

def log(s, m, l="INFO"): print(f"[{l}] {s}: {m}")

# ═══════════════════════════════════════
# 1. FX RATES
# ═══════════════════════════════════════
def fetch_fx():
    try:
        r = requests.get("https://api.frankfurter.app/latest?from=EUR&to=PLN,SEK", timeout=15)
        r.raise_for_status()
        d = r.json().get("rates", {})
        log("FX", f"PLN={d.get('PLN')}, SEK={d.get('SEK')}")
        return {"PLN_EUR": d.get("PLN"), "SEK_EUR": d.get("SEK")}
    except Exception as e:
        log("FX", str(e), "ERROR"); return None

# ═══════════════════════════════════════
# 2. ORLEN PL — via petrodom.pl
# ═══════════════════════════════════════
def fetch_orlen_pl():
    try:
        r = requests.get("https://www.petrodom.pl/en/current-wholesale-fuel-prices-provided-by-pkn-orlen/", headers=H, timeout=20)
        r.raise_for_status()
        soup = BeautifulSoup(r.text, "html.parser")
        for table in soup.find_all("table"):
            for row in table.find_all("tr"):
                cells = [td.get_text(strip=True) for td in row.find_all(["td","th"])]
                for i, cell in enumerate(cells):
                    if "ekodiesel" in cell.lower():
                        for j in range(i+1, min(i+3, len(cells))):
                            try:
                                price = float(re.sub(r'[^\d.]', '', cells[j].replace(" ","").replace(",",".").replace("\xa0","")))
                                if 3000 < price < 10000:
                                    log("Orlen PL", f"Ekodiesel = {price} PLN/m³")
                                    return {"price_pln_m3": price}
                            except: continue
        log("Orlen PL", "Not found", "WARN"); return None
    except Exception as e:
        log("Orlen PL", str(e), "ERROR"); return None

# ═══════════════════════════════════════
# 3. ORLEN LT — PDF: pardavimo kaina su PVM
# ═══════════════════════════════════════
def fetch_orlen_lt():
    try:
        list_url = "https://www.orlenlietuva.lt/lt/wholesale/_layouts/f2hPriceTable/default.aspx"
        r = requests.get(list_url, headers=H, timeout=20)
        r.raise_for_status()
        soup = BeautifulSoup(r.text, "html.parser")
        
        pdf_links = []
        for a in soup.find_all("a", href=True):
            href = a["href"]
            if ".pdf" in href.lower() and ("kainos" in href.lower() or "realizacija" in href.lower()):
                if not href.startswith("http"):
                    href = "https://www.orlenlietuva.lt" + href
                pdf_links.append(href.replace(" ", "%20")) # Sutvarkome tarpus URL
        
        # Jei neradome nuorodų puslapyje, bandom sugeneruoti pagal datą (kaip fallback)
        if not pdf_links:
            for days_back in range(7):
                d = TODAY - timedelta(days=days_back)
                pdf_links.append(f"https://www.orlenlietuva.lt/LT/Wholesale/Prices/Kainos%20{d.strftime('%Y%%20%m%%20%d')}%%20realizacija%%20internet.pdf")

        for pdf_url in pdf_links[:5]:
            try:
                log("Orlen LT", f"Tikrinamas PDF: {pdf_url}")
                r2 = requests.get(pdf_url, headers=H, timeout=15)
                if r2.status_code == 200 and len(r2.content) > 500:
                    price = parse_orlen_lt_pdf(r2.content)
                    if price: return price
            except: continue
            
        log("Orlen LT", "Kaina nerasta nei viename PDF", "WARN")
        return None
    except Exception as e:
        log("Orlen LT", str(e), "ERROR"); return None

def parse_orlen_lt_pdf(pdf_bytes):
    try:
        import pdfplumber
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            for page in pdf.pages:
                tables = page.extract_tables()
                for table in tables:
                    for row in table:
                        if not row: continue
                        # Sujungiam visas eilutės ląsteles į vieną tekstą paieškai
                        row_text = " ".join([str(c) for c in row if c]).lower()
                        
                        if "dyzelinas" in row_text and ("e kl" in row_text or "rrme" in row_text):
                            # Išvalome langelius ir sudedame skaičius į sąrašą
                            raw_values = []
                            for cell in row:
                                if cell:
                                    # Pašaliname viską išskyrus skaičius ir tašką/kablelį
                                    clean = re.sub(r'[^\d.,]', '', str(cell).replace(",", "."))
                                    if clean and clean != ".":
                                        raw_values.append(clean)
                            
                            # Logika: Jei kaina "1 801.40" suskilo į "1" ir "801.40", jas sujungiame
                            full_numbers = []
                            i = 0
                            while i < len(raw_values):
                                val = raw_values[i]
                                # Jei skaičius yra "1" arba "2" ir po jo eina skaičius su šimtais (pvz "801.40")
                                if (val in ["1", "2"]) and (i + 1 < len(raw_values)) and ("." in raw_values[i+1]):
                                    full_numbers.append(float(val + raw_values[i+1]))
                                    i += 2
                                else:
                                    try: full_numbers.append(float(val))
                                    except: pass
                                    i += 1
                            
                            # Ieškome didžiausios vertės (tai bus kaina su PVM)
                            valid_prices = [p for p in full_numbers if 1000 < p < 3000]
                            if valid_prices:
                                final_price = max(valid_prices)
                                eur_l = final_price / 1000
                                log("Orlen LT", f"SĖKMĖ! Rasta kaina su PVM: {final_price} EUR/1000l")
                                return {"price_eur_l": round(eur_l, 4)}
                                
    except Exception as e:
        log("Orlen LT", f"PDF klaida: {e}", "WARN")
    return None

def find_lt_diesel_in_text(text):
    """Fallback text search for Orlen LT diesel price"""
    # Look for pattern: Dyzelinas E ... number around 1500-2000 (EUR/1000l with VAT)
    patterns = [
        r'[Dd]yzelinas\s+E[^0-9]{0,150}?(\d{1,2}[\s\xa0]?\d{3}[.,]\d{2})',
        r'su\s+PVM[^0-9]{0,50}?(\d{1,2}[\s\xa0]?\d{3}[.,]\d{2})',
        r'(\d{1,2}[\s\xa0]?\d{3}[.,]\d{2})[^0-9]{0,30}su\s+PVM',
    ]
    for pat in patterns:
        matches = re.findall(pat, text, re.IGNORECASE)
        for m in matches:
            clean_val = re.sub(r'[^\d,.]', '', m).replace(",", ".")
            val = float(clean_val)
            if 1000 < val < 2500:  # EUR/1000l range
                eur_l = val / 1000
                log("Orlen LT", f"Text: {val} EUR/1000l → {eur_l:.3f} EUR/l")
                return {"price_eur_l": round(eur_l, 4)}
    return None

# ═══════════════════════════════════════
# 4. ELVIS DE — Diesel from fuel-prices.eu (NOT Super E5!)
# ═══════════════════════════════════════
def fetch_elvis_de():
    """Get Germany DIESEL price from fuel-prices.eu (EC Oil Bulletin data)"""
    try:
        r = requests.get("https://www.fuel-prices.eu/cheapest/", headers=H, timeout=20)
        r.raise_for_status()
        soup = BeautifulSoup(r.text, "html.parser")
        
        # Find the diesel table (second table on the page)
        tables = soup.find_all("table")
        for table in tables:
            table_text = table.get_text(" ", strip=True).lower()
            if "diesel" not in table_text: continue
            
            for row in table.find_all("tr"):
                cells = [td.get_text(strip=True) for td in row.find_all(["td","th"])]
                row_text = " ".join(cells).lower()
                if "germany" in row_text or "DE" in " ".join(cells):
                    for cell in cells:
                        # Price format: €1.812 or 1.812
                        m = re.search(r'€?(\d\.\d{3})', cell)
                        if m:
                            price = float(m.group(1))
                            if 0.8 < price < 3.5:
                                log("Elvis DE", f"Germany Diesel = {price} EUR/l (EC Oil Bulletin)")
                                return {"price_eur_l": price}
        
        log("Elvis DE", "Germany Diesel not found in table", "WARN")
        return None
    except Exception as e:
        log("Elvis DE", str(e), "ERROR"); return None

# ═══════════════════════════════════════
# 5. BSH / ST1 SE
# ═══════════════════════════════════════
def fetch_bsh_se():
    try:
        r = requests.get("https://st1.se/foretag/listpris", headers=H, timeout=20)
        r.raise_for_status()
        soup = BeautifulSoup(r.text, "html.parser")
        text = soup.get_text(" ", strip=True)
        for pat in [r'[Dd]iesel\s*(?:MK[123])?\s*[^0-9]{0,30}?(\d{1,2}[.,]\d{2})', r'(\d{1,2}[.,]\d{2})\s*(?:kr|SEK)[^0-9]{0,20}[Dd]iesel']:
            m = re.search(pat, text)
            if m:
                p = float(m.group(1).replace(",", "."))
                if 10 < p < 35:
                    log("BSH SE", f"Diesel = {p} SEK/l")
                    return {"price_sek_l": p}
        log("BSH SE", "Not found", "WARN"); return None
    except Exception as e:
        log("BSH SE", str(e), "ERROR"); return None

# ═══════════════════════════════════════
# 6. EU BULLETIN — ALL countries from fuel-prices.eu
# ═══════════════════════════════════════
def fetch_eu_bulletin():
    """Fetch diesel prices for LT/LV/EE/DK/SE/FI + EU avg from fuel-prices.eu"""
    try:
        r = requests.get("https://www.fuel-prices.eu/cheapest/", headers=H, timeout=20)
        r.raise_for_status()
        soup = BeautifulSoup(r.text, "html.parser")
        
        countries = {"LT": None, "LV": None, "EE": None, "DK": None, "SE": None, "FI": None}
        cc_map = {
            "lithuania": "LT", "latvia": "LV", "estonia": "EE",
            "denmark": "DK", "sweden": "SE", "finland": "FI"
        }
        eu_avg = None
        
        tables = soup.find_all("table")
        for table in tables:
            table_text = table.get_text(" ", strip=True).lower()
            if "diesel" not in table_text: continue
            
            for row in table.find_all("tr"):
                cells = [td.get_text(strip=True) for td in row.find_all(["td","th"])]
                if len(cells) < 2: continue
                row_text = cells[0].lower() + " " + cells[1].lower()
                
                for name, cc in cc_map.items():
                    if name in row_text:
                        for cell in cells:
                            m = re.search(r'€?(\d\.\d{3})', cell)
                            if m:
                                val = float(m.group(1))
                                if 0.5 < val < 3.5:
                                    countries[cc] = val
                                    break
        
        # EU average from page text
        text = soup.get_text(" ", strip=True)
        m = re.search(r'€(\d\.\d{3})/L\s+for\s+diesel', text)
        if m:
            eu_avg = float(m.group(1))
        
        found = {k: v for k, v in countries.items() if v is not None}
        if found:
            log("EU Bulletin", f"Diesel prices: {found}, EU avg={eu_avg}")
            return {**countries, "EU_AVG": eu_avg}
        
        log("EU Bulletin", "No countries found", "WARN")
        return None
    except Exception as e:
        log("EU Bulletin", str(e), "ERROR"); return None

# ═══════════════════════════════════════
# EXCEL WRITER — CALCULATED values (not formulas)
# ═══════════════════════════════════════
def update_excel(fx=None, orlen_pl=None, orlen_lt=None, elvis_de=None, bsh_se=None, eu_bulletin=None):
    if not EXCEL_PATH.exists():
        log("Excel", f"Not found: {EXCEL_PATH}", "ERROR"); return False
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

        # C: Orlen PL PLN/m³
        wc(3, orlen_pl["price_pln_m3"] if orlen_pl else None, '#,##0.00', ifont, ifill)
        # D: PLN/EUR
        wc(4, fx["PLN_EUR"] if fx else None, '0.0000', ifont, ifill)
        # E: Orlen PL EUR/l — CALCULATED
        pl_eur_l = orlen_pl["price_pln_m3"] / fx["PLN_EUR"] / 1000 if (orlen_pl and fx and fx.get("PLN_EUR")) else None
        wc(5, round(pl_eur_l, 4) if pl_eur_l else None, '0.000')
        # F: Orlen LT EUR/l
        lt_val = orlen_lt["price_eur_l"] if orlen_lt else None
        wc(6, lt_val, '0.000', ifont, ifill)
        # G: Delta — CALCULATED
        delta = round(pl_eur_l - lt_val, 4) if (pl_eur_l and lt_val) else None
        wc(7, delta, '+0.000;-0.000;"-"')
        # H: Delta % — CALCULATED
        delta_pct = round(delta / lt_val, 4) if (delta is not None and lt_val) else None
        wc(8, delta_pct, '+0.0%;-0.0%;"-"')
        # I: Elvis DE (Diesel!)
        wc(9, elvis_de["price_eur_l"] if elvis_de else None, '0.000', ifont, ifill)
        # J: BSH SEK
        wc(10, bsh_se["price_sek_l"] if bsh_se else None, '0.00', ifont, ifill)
        # K: SEK/EUR
        wc(11, fx["SEK_EUR"] if fx else None, '0.0000', ifont, ifill)
        # L: BSH EUR — CALCULATED
        se_eur = round(bsh_se["price_sek_l"] / fx["SEK_EUR"], 4) if (bsh_se and fx and fx.get("SEK_EUR")) else None
        wc(12, se_eur, '0.000')

        ok = [k for k,v in {"FX":fx,"PL":orlen_pl,"LT":orlen_lt,"DE":elvis_de,"SE":bsh_se}.items() if v]
        wc(13, f"Auto: {','.join(ok)}", font=Font(name="Aptos",size=9,color="6B7280"))
        wc(14, "Auto", font=Font(name="Aptos",size=8,color="9CA3AF"))
        log("Excel", f"Daily row 5: {TODAY_STR} [{','.join(ok)}]")

    # Weekly — runs every day now (not just Monday) since fuel-prices.eu updates weekly
    if eu_bulletin and "Weekly Oil Bulletin" in wb.sheetnames:
        ws_w = wb["Weekly Oil Bulletin"]
        # Check if this week already has data
        existing_date = ws_w.cell(row=4, column=1).value
        monday = TODAY - timedelta(days=WDAY)
        already_exists = False
        if existing_date:
            if hasattr(existing_date, 'date'):
                already_exists = existing_date.date() == monday.date()
            elif isinstance(existing_date, str):
                already_exists = existing_date[:10] == monday.strftime("%Y-%m-%d")
        
        if not already_exists:
            ws_w.insert_rows(4)
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
            lt_v = eu_bulletin.get("LT"); eu_v = eu_bulletin.get("EU_AVG")
            ws_w.cell(row=4,column=9).value = round((lt_v-eu_v)/eu_v,4) if (lt_v and eu_v) else None
            ws_w.cell(row=4,column=9).number_format = '+0.0%;-0.0%;"-"'
            log("Excel", f"Weekly row 4: {monday.strftime('%Y-%m-%d')}")
        else:
            log("Excel", f"Weekly row for {monday.strftime('%Y-%m-%d')} already exists, skipping")

    wb.save(str(EXCEL_PATH))
    log("Excel", f"Saved: {EXCEL_PATH}")
    return True

# ═══════════════════════════════════════
# MAIN
# ═══════════════════════════════════════
def main():
    print(f"\n{'='*60}\n  FUEL PRICE TRACKER v4 — {TODAY_STR}\n{'='*60}\n")
    R = {}
    print("── FX Rates ──"); R["fx"] = fetch_fx()
    if WDAY < 5:
        print("\n── Orlen PL ──"); R["orlen_pl"] = fetch_orlen_pl()
        print("\n── Orlen LT (PDF) ──"); R["orlen_lt"] = fetch_orlen_lt()
        print("\n── Elvis DE (Diesel) ──"); R["elvis_de"] = fetch_elvis_de()
        print("\n── BSH/ST1 SE ──"); R["bsh_se"] = fetch_bsh_se()
    else:
        print("\n── Weekend ──")
        for k in ["orlen_pl","orlen_lt","elvis_de","bsh_se"]: R[k] = None
    print("\n── EU Oil Bulletin ──"); R["eu_bulletin"] = fetch_eu_bulletin()
    
    print(f"\n{'─'*60}")
    ok = sum(1 for v in R.values() if v is not None)
    print(f"RESULTS: {ok}/{len(R)}")
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
