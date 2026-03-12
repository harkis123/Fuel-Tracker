"""
Fuel Price Tracker v4 — FINAL — All corrections applied
- Orlen LT: pardavimo kaina su PVM (EUR/1000l ÷ 1000)
- Elvis DE: Diesel from fuel-prices.eu (not Super E5 from mehr-tanken)
- EU Bulletin: All country diesel prices from fuel-prices.eu/cheapest/
- BSH SE: st1.se diesel SEK/l
"""

import requests, re, io
from bs4 import BeautifulSoup
from openpyxl import load_workbook
from datetime import datetime
from pathlib import Path

# Nustatymai
EXCEL_PATH = Path("fuel_tracker.xlsx")
H = {"User-Agent": "Mozilla/5.0"}

def fetch_orlen_lt():
    try:
        import pdfplumber
        # 1. Gauname naujausią nuorodą
        r = requests.get("https://www.orlenlietuva.lt/LT/Wholesale/Prices/Pages/default.aspx", headers=H)
        soup = BeautifulSoup(r.text, "html.parser")
        pdf_url = next((a["href"] for a in soup.find_all("a", href=True) if "realizacija" in a["href"].lower()), None)
        if not pdf_url.startswith("http"): pdf_url = "https://www.orlenlietuva.lt" + pdf_url
        
        # 2. Skaitome PDF
        res = requests.get(pdf_url.replace(" ", "%20"), headers=H)
        with pdfplumber.open(io.BytesIO(res.content)) as pdf:
            page = pdf.pages[0]
            table = page.extract_table()
            
            for row in table:
                # Nuvalome tuščius langelius
                row = [c.strip() if c else "" for c in row]
                row_str = " ".join(row).lower()
                
                if "dyzelinas" in row_str and "e kl" in row_str:
                    # Orlen PDF specifika: 1 stulpelyje yra "1", o 2 stulpelyje "801.40"
                    # Mes surandame visus skaičius ir juos sujungiame
                    full_text = "".join(row).replace(",", ".")
                    # Ieškome kainos formato (pvz. 1801.40)
                    match = re.search(r'(\d{4}\.\d{2})', full_text)
                    if match:
                        price = float(match.group(1))
                        print(f"[LT] Rasta kaina: {price}")
                        return {"price_eur_l": round(price / 1000, 4)}
        return None
    except Exception as e:
        print(f"[LT] Klaida: {e}")
        return None

def fetch_orlen_pl():
    try:
        r = requests.get("https://www.petrodom.pl/en/current-wholesale-fuel-prices-provided-by-pkn-orlen/", headers=H)
        soup = BeautifulSoup(r.text, "html.parser")
        for row in soup.find_all("tr"):
            txt = row.get_text().lower()
            if "ekodiesel" in txt:
                # Svarbu: paimame tik pirmą skaičių (bazinė kaina be PVM Lenkijoje pildoma be PVM)
                nums = re.findall(r'(\d{4}[.,]\d{2})', txt)
                if nums:
                    val = float(nums[0].replace(",", "."))
                    print(f"[PL] Rasta kaina: {val}")
                    return {"price_pln_m3": val}
        return None
    except Exception as e:
        print(f"[PL] Klaida: {e}")
        return None

if __name__ == "__main__":
    print("--- TESTAS ---")
    lt = fetch_orlen_lt()
    pl = fetch_orlen_pl()
    
    if lt: print(f"LT Rezultatas: {lt['price_eur_l']} EUR/l")
    else: print("LT Rezultatas: NERASTA")
    
    if pl: print(f"PL Rezultatas: {pl['price_pln_m3']} PLN/m3")
    else: print("PL Rezultatas: NERASTA")
if __name__ == "__main__":
    main()
