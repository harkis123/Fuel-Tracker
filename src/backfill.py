"""
Backfill historical fuel prices from 2026-02-01 to today.
Run ONCE to populate history, then delete or disable.
Usage: python src/backfill.py
"""
import requests, json, re, io, sys
from datetime import datetime, timedelta
from bs4 import BeautifulSoup
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from pathlib import Path

EXCEL_PATH = Path(__file__).parent.parent / "fuel_tracker.xlsx"
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/131.0.0.0 Safari/537.36"}

def log(m): print(m)

def clean_num(s):
    if s is None: return None
    s = str(s).strip()
    s = re.sub(r'[€$£\xa0]', '', s)
    if re.match(r'^\d[\d ]+\.\d+$', s): s = s.replace(' ', ''); return float(s)
    if re.match(r'^\d[\d ]+,\d+$', s): s = s.replace(' ', '').replace(',', '.'); return float(s)
    if ',' in s and '.' in s: s = s.replace(',', ''); return float(s)
    if ',' in s: s = s.replace(',', '.')
    s = s.replace(' ', '')
    try: return float(s)
    except: return None

# ═══════════════════════════════════════
# FETCH HISTORICAL FX RATES
# ═══════════════════════════════════════
def fetch_fx_history(start="2026-02-01", end=None):
    """Get daily PLN/EUR and SEK/EUR from frankfurter.app"""
    if end is None: end = datetime.now().strftime("%Y-%m-%d")
    log(f"Fetching FX rates {start} to {end}...")
    r = requests.get(f"https://api.frankfurter.app/{start}..{end}?from=EUR&to=PLN,SEK", timeout=30)
    r.raise_for_status()
    data = r.json()
    rates = {}
    for date_str, vals in data.get("rates", {}).items():
        rates[date_str] = {"PLN_EUR": vals.get("PLN"), "SEK_EUR": vals.get("SEK")}
    log(f"  Got {len(rates)} days of FX data")
    return rates

# ═══════════════════════════════════════
# FETCH HISTORICAL EU BULLETIN
# ═══════════════════════════════════════
def fetch_eu_history():
    """Get EU diesel prices from fuel-prices.eu"""
    log("Fetching EU Bulletin diesel prices...")
    countries = {"LT": None, "LV": None, "EE": None, "DK": None, "SE": None, "FI": None}
    cc_map = {"lithuania":"LT","latvia":"LV","estonia":"EE","denmark":"DK","sweden":"SE","finland":"FI"}
    
    try:
        r = requests.get("https://www.fuel-prices.eu/cheapest/", headers=H, timeout=20)
        r.raise_for_status()
        soup = BeautifulSoup(r.text, "html.parser")
        eu_avg = None
        
        for table in soup.find_all("table"):
            if "diesel" not in table.get_text(" ", strip=True).lower(): continue
            for row in table.find_all("tr"):
                cells = [td.get_text(strip=True) for td in row.find_all(["td","th"])]
                if len(cells) < 2: continue
                rt = " ".join(cells[:2]).lower()
                for name, cc in cc_map.items():
                    if name in rt:
                        for cell in cells:
                            m = re.search(r'€?(\d\.\d{3})', cell)
                            if m:
                                val = float(m.group(1))
                                if 0.5 < val < 3.5: countries[cc] = val; break
        
        text = soup.get_text(" ", strip=True)
        m = re.search(r'€(\d\.\d{3})/L\s+for\s+diesel', text)
        if m: eu_avg = float(m.group(1))
        
        # Also get Germany for Elvis DE
        de_diesel = None
        for table in soup.find_all("table"):
            if "diesel" not in table.get_text(" ", strip=True).lower(): continue
            for row in table.find_all("tr"):
                cells = [td.get_text(strip=True) for td in row.find_all(["td","th"])]
                if "germany" in " ".join(cells).lower():
                    for cell in cells:
                        m = re.search(r'€?(\d\.\d{3})', cell)
                        if m:
                            de_diesel = float(m.group(1))
                            break
        
        found = {k:v for k,v in countries.items() if v}
        log(f"  EU: {found}, avg={eu_avg}, DE={de_diesel}")
        return countries, eu_avg, de_diesel
    except Exception as e:
        log(f"  EU Bulletin error: {e}")
        return countries, None, None

# ═══════════════════════════════════════
# FETCH ORLEN LT HISTORY (PDF archive)
# ═══════════════════════════════════════
def fetch_orlen_lt_history():
    """Fetch multiple Orlen LT PDF protocols from archive"""
    log("Fetching Orlen LT PDF archive...")
    
    try:
        r = requests.get("https://www.orlenlietuva.lt/lt/wholesale/_layouts/f2hPriceTable/default.aspx", headers=H, timeout=20)
        r.raise_for_status()
        soup = BeautifulSoup(r.text, "html.parser")
        
        results = {}
        pdf_links = []
        for a in soup.find_all("a", href=True):
            href = a["href"]
            if ".pdf" not in href.lower() or "kainos" not in href.lower(): continue
            if not href.startswith("http"):
                href = "https://www.orlenlietuva.lt" + href
            # Extract date from URL
            date_match = re.search(r'(\d{4})\s*(\d{2})\s*(\d{2})', href)
            if date_match:
                date_str = f"{date_match.group(1)}-{date_match.group(2)}-{date_match.group(3)}"
                pdf_links.append((date_str, href))
        
        log(f"  Found {len(pdf_links)} PDF links")
        
        for date_str, pdf_url in pdf_links:
            if date_str < "2026-02-01": continue
            try:
                r2 = requests.get(pdf_url, headers=H, timeout=15)
                if r2.status_code != 200 or len(r2.content) < 500: continue
                
                import pdfplumber
                with pdfplumber.open(io.BytesIO(r2.content)) as pdf:
                    for page in pdf.pages:
                        for table in page.extract_tables():
                            for row in table:
                                if not row: continue
                                row_str = " ".join([str(c) for c in row if c])
                                if "dyzelinas" not in row_str.lower(): continue
                                if "rrme" not in row_str.lower() and "e kl" not in row_str.lower(): continue
                                
                                prices = []
                                for cell in row:
                                    val = clean_num(cell)
                                    if val and val > 0: prices.append(val)
                                
                                eur_1000l = [p for p in prices if 1000 < p < 2500]
                                if eur_1000l:
                                    selling_price = max(eur_1000l)
                                    results[date_str] = round(selling_price / 1000, 4)
                                    log(f"  {date_str}: {selling_price} EUR/1000l → {results[date_str]} EUR/l")
                                    break
            except Exception as e:
                log(f"  {date_str}: error - {e}")
                continue
        
        log(f"  Got {len(results)} Orlen LT prices")
        return results
    except Exception as e:
        log(f"  Orlen LT history error: {e}")
        return {}

# ═══════════════════════════════════════
# FETCH ORLEN PL (current only)
# ═══════════════════════════════════════
def fetch_orlen_pl_current():
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
                            price = clean_num(cells[j])
                            if price and 3000 < price < 10000:
                                return price
    except: pass
    return None

# ═══════════════════════════════════════
# FETCH BSH SE (current only)
# ═══════════════════════════════════════
def fetch_bsh_se_current():
    try:
        r = requests.get("https://st1.se/foretag/listpris", headers=H, timeout=20)
        r.raise_for_status()
        text = BeautifulSoup(r.text, "html.parser").get_text(" ", strip=True)
        for pat in [r'[Dd]iesel\s*(?:MK[123])?\s*[^0-9]{0,30}?(\d{1,2}[.,]\d{2})']:
            m = re.search(pat, text)
            if m:
                p = float(m.group(1).replace(",", "."))
                if 10 < p < 35: return p
    except: pass
    return None

# ═══════════════════════════════════════
# MAIN BACKFILL
# ═══════════════════════════════════════
def main():
    log("="*60)
    log("  BACKFILL: 2026-02-01 to today")
    log("="*60)
    
    # 1. Fetch all data
    fx_history = fetch_fx_history("2026-02-01")
    orlen_lt_history = fetch_orlen_lt_history()
    eu_countries, eu_avg, de_diesel = fetch_eu_history()
    orlen_pl = fetch_orlen_pl_current()
    bsh_sek = fetch_bsh_se_current()
    
    log(f"\nOrlen PL current: {orlen_pl} PLN/m³")
    log(f"BSH SE current: {bsh_sek} SEK/l")
    log(f"Elvis DE (EU Bulletin): {de_diesel} EUR/l")
    
    # 2. Open Excel
    if not EXCEL_PATH.exists():
        log(f"ERROR: {EXCEL_PATH} not found"); sys.exit(1)
    wb = load_workbook(str(EXCEL_PATH))
    ws = wb["Daily Tracker"]
    
    # 3. Clear old data rows (keep headers rows 1-4)
    if ws.max_row > 4:
        ws.delete_rows(5, ws.max_row - 4)
    
    # 4. Generate dates from today back to 2026-02-01
    start = datetime(2026, 2, 1)
    end = datetime.now()
    dates = []
    d = end
    while d >= start:
        if d.weekday() < 5:  # Skip weekends
            dates.append(d)
        d -= timedelta(days=1)
    
    log(f"\nWriting {len(dates)} business days...")
    
    ifont = Font(name="Aptos", size=10, color="1D4ED8")
    ifill = PatternFill("solid", fgColor="EFF6FF")
    dfont = Font(name="Aptos", size=10, color="1F2937")
    brd = Border(left=Side("thin",color="D1D5DB"),right=Side("thin",color="D1D5DB"),top=Side("thin",color="D1D5DB"),bottom=Side("thin",color="D1D5DB"))
    
    import calendar
    
    for i, date in enumerate(dates):
        row = 5 + i
        ds = date.strftime("%Y-%m-%d")
        fx = fx_history.get(ds, {})
        pln_eur = fx.get("PLN_EUR")
        sek_eur = fx.get("SEK_EUR")
        lt_price = orlen_lt_history.get(ds)
        
        def wc(col, val, fmt='General', font=dfont, fill=None):
            c = ws.cell(row=row, column=col)
            c.value = val; c.number_format = fmt; c.font = font; c.border = brd
            c.alignment = Alignment(horizontal="right" if col > 2 else "center", vertical="center")
            if fill: c.fill = fill
        
        # Only current Orlen PL price for recent days (last 3 days)
        pl_m3 = orlen_pl if (end - date).days <= 3 else None
        pl_eur = round(pl_m3 / pln_eur / 1000, 4) if (pl_m3 and pln_eur) else None
        
        # Only current BSH for recent days
        bsh = bsh_sek if (end - date).days <= 3 else None
        bsh_eur = round(bsh / sek_eur, 4) if (bsh and sek_eur) else None
        
        # Only current DE diesel for recent days
        de = de_diesel if (end - date).days <= 7 else None
        
        delta = round(pl_eur - lt_price, 4) if (pl_eur and lt_price) else None
        delta_pct = round(delta / lt_price, 4) if (delta is not None and lt_price) else None
        
        wc(1, date, 'YYYY-MM-DD', Font(name="Aptos",size=10,bold=True,color="1F2937"))
        wc(2, calendar.day_abbr[date.weekday()], font=Font(name="Aptos",size=9,color="6B7280"))
        wc(3, pl_m3, '#,##0.00', ifont, ifill)
        wc(4, pln_eur, '0.0000', ifont, ifill)
        wc(5, pl_eur, '0.000')
        wc(6, lt_price, '0.000', ifont, ifill)
        wc(7, delta, '+0.000;-0.000;"-"')
        wc(8, delta_pct, '+0.0%;-0.0%;"-"')
        wc(9, de, '0.000', ifont, ifill)
        wc(10, bsh, '0.00', ifont, ifill)
        wc(11, sek_eur, '0.0000', ifont, ifill)
        wc(12, bsh_eur, '0.000')
        
        sources = []
        if pln_eur: sources.append("FX")
        if pl_m3: sources.append("PL")
        if lt_price: sources.append("LT")
        if de: sources.append("DE")
        if bsh: sources.append("SE")
        wc(13, f"{'Auto' if sources else 'FX only'}: {','.join(sources)}" if sources else "FX only", font=Font(name="Aptos",size=9,color="6B7280"))
        wc(14, "Backfill", font=Font(name="Aptos",size=8,color="9CA3AF"))
    
    # 5. EU Bulletin — add current week
    if eu_countries and any(v for v in eu_countries.values() if v):
        ws_w = wb["Weekly Oil Bulletin"]
        # Clear old weekly data
        if ws_w.max_row > 3:
            ws_w.delete_rows(4, ws_w.max_row - 3)
        
        monday = end - timedelta(days=end.weekday())
        ws_w.cell(row=4,column=1).value = monday
        ws_w.cell(row=4,column=1).number_format = 'YYYY-MM-DD'
        ws_w.cell(row=4,column=1).font = Font(name="Aptos",size=10,bold=True,color="1F2937")
        for k,col in {"LT":2,"LV":3,"EE":4,"DK":5,"SE":6,"FI":7,"EU_AVG":8}.items():
            val = eu_countries.get(k) if k != "EU_AVG" else eu_avg
            if val:
                c = ws_w.cell(row=4,column=col)
                c.value = val; c.number_format = '0.000'
                c.font = Font(name="Aptos",size=10,color="1D4ED8")
        lt_v = eu_countries.get("LT")
        ws_w.cell(row=4,column=9).value = round((lt_v-eu_avg)/eu_avg,4) if (lt_v and eu_avg) else None
        ws_w.cell(row=4,column=9).number_format = '+0.0%;-0.0%;"-"'
    
    wb.save(str(EXCEL_PATH))
    log(f"\n✅ Backfill complete! {len(dates)} days written to {EXCEL_PATH}")
    log("Now run the regular scraper to get today's data on top.")

if __name__ == "__main__":
    main()
