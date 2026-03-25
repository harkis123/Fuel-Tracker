"""
Backfill missing data in fuel_tracker.xlsx
Run once: python src/backfill.py
"""
import requests, re, io
from datetime import datetime
from openpyxl import load_workbook
from pathlib import Path

EXCEL_PATH = Path(__file__).parent.parent / "fuel_tracker.xlsx"
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/131.0.0.0 Safari/537.36"}

def log(m): print(f"  {m}")

def main():
    print("="*60)
    print("  BACKFILL — filling all gaps")
    print("="*60)

    # 1. FX history
    print("\n[1/4] FX rates...")
    fx = {}
    try:
        r = requests.get("https://api.frankfurter.app/2026-02-01..2026-03-25?from=EUR&to=PLN,SEK", timeout=30)
        r.raise_for_status()
        fx = r.json().get("rates", {})
        log(f"frankfurter.app: {len(fx)} days")
    except Exception as e:
        log(f"frankfurter failed: {e}")
        try:
            r = requests.get("https://www.ecb.europa.eu/stats/eurofxref/eurofxref-hist-90d.xml", timeout=30)
            r.raise_for_status()
            import xml.etree.ElementTree as ET
            root = ET.fromstring(r.text)
            ns = {'eurofx': 'http://www.ecb.int/vocabulary/2002-08-01/eurofxref'}
            for cube in root.findall('.//{http://www.ecb.int/vocabulary/2002-08-01/eurofxref}Cube[@time]'):
                date = cube.get('time')
                rates = {}
                for rate in cube:
                    if rate.get('currency') in ('PLN', 'SEK'):
                        rates[rate.get('currency')] = float(rate.get('rate'))
                if 'PLN' in rates and 'SEK' in rates:
                    fx[date] = rates
            log(f"ECB XML: {len(fx)} days")
        except Exception as e2:
            log(f"ECB also failed: {e2}")

    # 2. Orlen LT PDFs
    print("\n[2/4] Orlen LT PDFs...")
    lt_prices = {}
    try:
        from bs4 import BeautifulSoup
        import pdfplumber
        r = requests.get("https://www.orlenlietuva.lt/lt/wholesale/_layouts/f2hPriceTable/default.aspx", headers=H, timeout=20)
        r.raise_for_status()
        soup = BeautifulSoup(r.text, "html.parser")
        for a in soup.find_all("a", href=True):
            href = a["href"]
            if ".pdf" not in href.lower() or "kainos" not in href.lower(): continue
            if not href.startswith("http"): href = "https://www.orlenlietuva.lt" + href
            dm = re.search(r'(\d{4})\s*(\d{2})\s*(\d{2})', href)
            if not dm: continue
            ds = f"{dm.group(1)}-{dm.group(2)}-{dm.group(3)}"
            if ds < "2026-02-01": continue
            try:
                r2 = requests.get(href, headers=H, timeout=15)
                if r2.status_code != 200 or len(r2.content) < 500: continue
                with pdfplumber.open(io.BytesIO(r2.content)) as pdf:
                    text = pdf.pages[0].extract_text() or ""
                    for line in text.split('\n'):
                        if 'Dyzelinas E kl. su RRME' not in line: continue
                        nums = re.findall(r'(\d[\d ]*\.\d{2})', line)
                        cleaned = [float(n.replace(' ', '')) for n in nums]
                        if cleaned:
                            sp = cleaned[-1]
                            if 1000 < sp < 2500:
                                lt_prices[ds] = round(sp / 1000, 4)
                                log(f"LT {ds}: {sp} -> {lt_prices[ds]}")
                        break
            except: pass
        log(f"Got {len(lt_prices)} LT prices")
    except Exception as e:
        log(f"LT error: {e}")

    # 3. EC XLSX for DE
    print("\n[3/4] EC XLSX for DE diesel...")
    de_price = None
    try:
        ec_url = "https://energy.ec.europa.eu/document/download/264c2d0f-f161-4ea3-a777-78faae59bea0_en?filename=Weekly%20Oil%20Bulletin%20Weekly%20prices%20with%20Taxes%20-%202024-02-19.xlsx"
        r = requests.get(ec_url, headers=H, timeout=30)
        ec_wb = load_workbook(io.BytesIO(r.content), data_only=True)
        ec_ws = ec_wb[ec_wb.sheetnames[0]]
        diesel_col = 3
        for row in range(1, 6):
            for col in range(1, 10):
                val = str(ec_ws.cell(row=row, column=col).value or "").lower()
                if "gas oil" in val or "diesel" in val:
                    diesel_col = col; break
        for row in range(1, ec_ws.max_row + 1):
            c0 = str(ec_ws.cell(row=row, column=1).value or "").lower()
            if "germany" in c0:
                v = ec_ws.cell(row=row, column=diesel_col).value
                if v:
                    fv = float(v)
                    de_price = round(fv / 1000, 4) if fv > 100 else round(fv, 4)
                    log(f"DE diesel: {de_price}")
                    break
    except Exception as e:
        log(f"EC error: {e}")

    # 4. Fill Excel
    print("\n[4/4] Filling Excel...")
    if not EXCEL_PATH.exists():
        print(f"ERROR: {EXCEL_PATH} not found"); return

    wb = load_workbook(str(EXCEL_PATH))
    ws = wb['Daily Tracker']
    filled = 0

    for r in range(5, ws.max_row + 1):
        dt = ws.cell(row=r, column=1).value
        if dt is None: continue
        ds = dt.strftime('%Y-%m-%d') if hasattr(dt, 'strftime') else str(dt)[:10]
        changed = False

        fx_day = fx.get(ds, {})
        pln_eur = fx_day.get("PLN")
        sek_eur = fx_day.get("SEK")

        if ws.cell(row=r, column=4).value is None and pln_eur:
            ws.cell(row=r, column=4).value = pln_eur
            ws.cell(row=r, column=4).number_format = '0.0000'
            changed = True

        if ws.cell(row=r, column=11).value is None and sek_eur:
            ws.cell(row=r, column=11).value = sek_eur
            ws.cell(row=r, column=11).number_format = '0.0000'
            changed = True

        if ws.cell(row=r, column=6).value is None and ds in lt_prices:
            ws.cell(row=r, column=6).value = lt_prices[ds]
            ws.cell(row=r, column=6).number_format = '0.000'
            changed = True

        if ws.cell(row=r, column=9).value is None and de_price:
            ws.cell(row=r, column=9).value = de_price
            ws.cell(row=r, column=9).number_format = '0.000'
            changed = True

        # Recalculate PL EUR/l
        plm3 = ws.cell(row=r, column=3).value
        plneur = ws.cell(row=r, column=4).value or pln_eur
        if plm3 and plneur:
            ws.cell(row=r, column=5).value = round(plm3 / plneur / 1000, 4)
            ws.cell(row=r, column=5).number_format = '0.000'

        # Recalculate delta
        pleurl = ws.cell(row=r, column=5).value
        lteurl = ws.cell(row=r, column=6).value
        if pleurl and lteurl:
            ws.cell(row=r, column=7).value = round(pleurl - lteurl, 4)
            ws.cell(row=r, column=7).number_format = '+0.000;-0.000;"-"'
            ws.cell(row=r, column=8).value = round((pleurl - lteurl) / lteurl, 4)
            ws.cell(row=r, column=8).number_format = '+0.0%;-0.0%;"-"'

        # Recalculate SE EUR/l
        sesek = ws.cell(row=r, column=10).value
        sekeur = ws.cell(row=r, column=11).value or sek_eur
        if sesek and sekeur:
            ws.cell(row=r, column=12).value = round(sesek / sekeur, 4)
            ws.cell(row=r, column=12).number_format = '0.000'

        ok = []
        if ws.cell(row=r, column=4).value: ok.append("FX")
        if ws.cell(row=r, column=3).value: ok.append("PL")
        if ws.cell(row=r, column=6).value: ok.append("LT")
        if ws.cell(row=r, column=9).value: ok.append("DE")
        if ws.cell(row=r, column=10).value: ok.append("SE")
        ws.cell(row=r, column=13).value = f"Auto: {','.join(ok)}"

        if changed:
            filled += 1
            log(f"Filled: {ds}")

    wb.save(str(EXCEL_PATH))
    print(f"\n  DONE! Filled {filled} rows")

if __name__ == "__main__":
    main()
