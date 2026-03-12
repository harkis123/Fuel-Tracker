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
from openpyxl.styles import Font, PatternFill, Alignment
from pathlib import Path

# --- KONFIGŪRACIJA ---
EXCEL_PATH = Path(__file__).parent.parent / "fuel_tracker.xlsx"
UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
H = {"User-Agent": UA}
TODAY = datetime.now()

def log(s, m, l="INFO"): print(f"[{l}] {s}: {m}")

# ═══════════════════════════════════════
# 1. VALIUTŲ KURSAI
# ═══════════════════════════════════════
def fetch_fx():
    try:
        r = requests.get("https://api.frankfurter.app/latest?from=EUR&to=PLN,SEK", timeout=15)
        d = r.json().get("rates", {})
        return {"PLN_EUR": d.get("PLN"), "SEK_EUR": d.get("SEK")}
    except: return None

# ═══════════════════════════════════════
# 2. ORLEN LIETUVA (Iš oficialaus PDF)
# ═══════════════════════════════════════
def fetch_orlen_lt():
    try:
        import pdfplumber
        # Nueiname į kainų puslapį rasti naujausią PDF
        list_url = "https://www.orlenlietuva.lt/LT/Wholesale/Prices/Pages/default.aspx"
        r = requests.get(list_url, headers=H, timeout=20)
        soup = BeautifulSoup(r.text, "html.parser")
        
        pdf_url = None
        for a in soup.find_all("a", href=True):
            if ".pdf" in a["href"].lower() and "realizacija" in a["href"].lower():
                href = a["href"]
                pdf_url = href if href.startswith("http") else "https://www.orlenlietuva.lt" + href
                break
        
        if not pdf_url: return None
        log("Orlen LT", f"Atsisiunčiama: {pdf_url}")
        
        pdf_res = requests.get(pdf_url.replace(" ", "%20"), headers=H, timeout=20)
        with pdfplumber.open(io.BytesIO(pdf_res.content)) as pdf:
            text = pdf.pages[0].extract_text()
            
            for line in text.split('\n'):
                if "dyzelinas" in line.lower() and "e kl" in line.lower():
                    # Šis regex ras bet kokį skaičių "X XXX.XX" formatu (kiekvieną dieną kitokį)
                    # Jis ignoruoja tarpus tarp tūkstančių ir šimtų
                    matches = re.findall(r'(\d)[\s\xa0]?(\d{3}[.,]\d{2})', line)
                    if matches:
                        # Paskutinis skaičius eilutėje = kaina su PVM
                        last_m = matches[-1]
                        clean_val = float(last_m[0] + last_m[1].replace(",", "."))
                        log("Orlen LT", f"Rasta dienos kaina: {clean_val}")
                        return {"price_eur_l": round(clean_val / 1000, 4)}
        return None
    except Exception as e:
        log("Orlen LT", f"Klaida: {e}", "ERROR"); return None

# ═══════════════════════════════════════
# 3. ORLEN LENKIJA
# ═══════════════════════════════════════
def fetch_orlen_pl():
    try:
        r = requests.get("https://www.petrodom.pl/en/current-wholesale-fuel-prices-provided-by-pkn-orlen/", headers=H, timeout=20)
        soup = BeautifulSoup(r.text, "html.parser")
        for row in soup.find_all("tr"):
            txt = row.get_text().lower()
            if "ekodiesel" in txt:
                nums = re.findall(r'(\d{4}[.,]\d{2})', txt.replace(" ", ""))
                if nums: return {"price_pln_m3": float(nums[0].replace(",", "."))}
        return None
    except: return None

# ═══════════════════════════════════════
# 4. KITI (DE, SE)
# ═══════════════════════════════════════
def fetch_others():
    res = {"de": None, "se": None}
    try:
        r = requests.get("https://www.fuel-prices.eu/cheapest/", headers=H, timeout=20)
        m = re.search(r'Germany.*?(\d\.\d{3})', r.text, re.S)
        if m: res["de"] = float(m.group(1))
        
        r2 = requests.get("https://st1.se/foretag/listpris", headers=H, timeout=20)
        m2 = re.search(r'Diesel.*?(\d{2}[.,]\d{2})', r2.text)
        if m2: res["se"] = float(m2.group(1).replace(",", "."))
    except: pass
    return res

# ═══════════════════════════════════════
# 5. EXCEL PILDYMAS
# ═══════════════════════════════════════
def update_excel(fx, pl, lt, others):
    if not EXCEL_PATH.exists(): 
        log("Excel", f"Nerastas failas: {EXCEL_PATH}", "ERROR")
        return False
        
    wb = load_workbook(str(EXCEL_PATH))
    ws = wb["Daily Tracker"]
    ws.insert_rows(5)
    
    # Stiliai
    f_bold = Font(bold=True)
    f_blue = Font(color="1D4ED8")
    fill_blue = PatternFill("solid", fgColor="EFF6FF")
    
    def write(col, val, fmt='General', font=None, fill=None):
        c = ws.cell(row=5, column=col)
        c.value = val
        c.number_format = fmt
        if font: c.font = font
        if fill: c.fill = fill
        c.alignment = Alignment(horizontal="right")

    # Pildymas pagal tavo Excel struktūrą
    write(1, TODAY, 'YYYY-MM-DD', f_bold)
    write(2, calendar.day_abbr[TODAY.weekday()])
    
    if pl: write(3, pl["price_pln_m3"], '#,##0.00', f_blue, fill_blue)
    if fx: write(4, fx["PLN_EUR"], '0.0000', f_blue, fill_blue)
    
    pl_eur = (pl["price_pln_m3"] / fx["PLN_EUR"] / 1000) if (pl and fx) else None
    if pl_eur: write(5, pl_eur, '0.000')
    
    lt_val = lt["price_eur_l"] if lt else None
    if lt_val: write(6, lt_val, '0.000', f_blue, fill_blue)
    
    if pl_eur and lt_val:
        write(7, pl_eur - lt_val, '+0.000;-0.000')
        write(8, (pl_eur - lt_val) / lt_val, '0.0%')
        
    if others["de"]: write(9, others["de"], '0.000', f_blue, fill_blue)
    if others["se"]: write(10, others["se"], '0.00', f_blue, fill_blue)
    if fx: write(11, fx["SEK_EUR"], '0.0000', f_blue, fill_blue)
    if others["se"] and fx: write(12, others["se"]/fx["SEK_EUR"], '0.000')

    # Statuso žymėjimas (Notes stulpelis)
    notes = "Auto: FX"
    if pl: notes += ", PL"
    if lt: notes += ", LT"
    if others["de"]: notes += ", DE"
    if others["se"]: notes += ", SE"
    write(13, notes)

    wb.save(str(EXCEL_PATH))
    return True

# ═══════════════════════════════════════
# PALEIDIMAS
# ═══════════════════════════════════════
if __name__ == "__main__":
    print("--- STARTING TRACKER V5 ---")
    fx_data = fetch_fx()
    pl_data = fetch_orlen_pl()
    lt_data = fetch_orlen_lt() # ČIA BUVO KLAIDA, DABAR APIBRĖŽTA
    other_data = fetch_others()
    
    status = update_excel(fx_data, pl_data, lt_data, other_data)
    if status:
        print("✅ SUCCESS: Excel updated.")
    else:
        print("❌ FAILED: Check logs.")
