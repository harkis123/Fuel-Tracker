# ⛽ Fuel Price Tracker

Automatinis kuro kainų sekimas iš 6 šaltinių. Veikia per GitHub Actions — jokio serverio, jokio rankinio darbo.

## Ką daro

Kas darbo dieną 10:30 LT laiku automatiškai:

| Šaltinis | Duomenys | Dažnumas |
|----------|----------|----------|
| ECB (frankfurter.app) | PLN/EUR, SEK/EUR kursai | Kasdien |
| Orlen PL | Ekodiesel hurtinė kaina PLN/m³ | Kasdien |
| Orlen LT | Dyzelinas E su PVM, EUR/l | Kasdien |
| mehr-tanken.de | Elvis FSC Diesel DE, EUR/l | Kasdien |
| ST1.se | BSH/ST1 Diesel SE, SEK/l | Kasdien |
| EC Oil Bulletin | LT/LV/EE/DK/SE/FI diesel kainos | Pirmadienis |

Surinkti duomenys automatiškai įrašomi į `fuel_tracker.xlsx` ir commit'inami atgal į repo.

## Pradžia (setup ~5 min)

### 1. Fork arba sukurk repo

```bash
# Variantas A: Sukurk naują repo GitHub.com → "New repository" → "fuel-tracker"
# Variantas B: Arba komandinėje eilutėje:
git clone <this-repo>
cd fuel-tracker
git remote set-url origin https://github.com/TAVO-USERNAME/fuel-tracker.git
git push -u origin main
```

### 2. Įkelk failus

Įkelk visus šio repo failus į savo GitHub repo:
- `src/scraper.py` — pagrindinis scraper'is
- `.github/workflows/update-prices.yml` — GitHub Actions workflow
- `requirements.txt` — Python dependencies
- `fuel_tracker.xlsx` — Excel tracker šablonas

### 3. Įjunk GitHub Actions

1. Eik į savo repo → **Settings** → **Actions** → **General**
2. Pažymėk **"Allow all actions and reusable workflows"**
3. Apačioje **"Workflow permissions"** → pažymėk **"Read and write permissions"**
4. Spausk **Save**

### 4. Paleisk pirmą kartą

1. Eik į repo → **Actions** tab
2. Kairėje pasirink **"Fuel Price Tracker — Daily Update"**
3. Spausk **"Run workflow"** → **"Run workflow"**
4. Palauk ~1 min kol pabaigs
5. Patikrink `fuel_tracker.xlsx` — turėtų būti nauji duomenys!

### 5. SharePoint sinchronizacija

Kad Excel automatiškai atsinaujintų SharePoint'e, yra keli variantai:

**Variantas A (paprasčiausias):** Rankiniu būdu atsisiųsk `fuel_tracker.xlsx` iš GitHub ir įkelk į SharePoint kai reikia.

**Variantas B (automatinis):** Pridėk dar vieną GitHub Actions step kuris naudoja Microsoft Graph API siųsti failą į SharePoint. Tam reikės Azure App Registration — žr. `docs/sharepoint-sync.md` (TODO).

**Variantas C:** Naudok OneDrive sync klientą kuris sinchronizuoja local folder'į su SharePoint. Tada GitHub Actions rašo į tą folder'į.

## Kaip veikia

```
GitHub Actions (cron: 08:30 UTC, Mon-Fri)
    │
    ├─→ Fetch FX rates (frankfurter.app API)
    ├─→ Scrape Orlen PL (orlen.pl)
    ├─→ Scrape Orlen LT (orlenlietuva.lt)
    ├─→ Scrape Elvis DE (mehr-tanken.de)
    ├─→ Scrape BSH/ST1 SE (st1.se)
    ├─→ [Mondays] Fetch EU Oil Bulletin (energy.ec.europa.eu)
    │
    ├─→ Write all data to fuel_tracker.xlsx
    └─→ Git commit + push
```

## Troubleshooting

**Flow failina?** Eik į Actions → paskutinis run → pažiūrėk logs. Dažniausios priežastys:
- Šaltinio puslapio struktūra pasikeitė → reikia atnaujinti regex patterns `scraper.py`
- Puslapio blokuoja GitHub Actions IP → gali reikėti proxy

**Trūksta duomenų iš konkretaus šaltinio?** Kai kurie puslapiai naudoja JavaScript renderinimą ir paprastas HTTP GET negauna duomenų. Tokiu atveju `scraper.py` logina WARN — patikrink Actions logs.

**Excel nerodo naujausių duomenų?** Patikrink ar data eilutė sutampa su šiandienos data. Scraper'is ieško eilutės pagal datą.

## Failų struktūra

```
fuel-tracker/
├── .github/workflows/
│   └── update-prices.yml    ← GitHub Actions workflow (pinned SHAs)
├── src/
│   ├── config.py            ← Centralizuota konfigūracija (URLs, ribos, konstantos)
│   ├── scraper.py           ← Pagrindinis scraper (v7: logging, retries, type hints)
│   └── backfill.py          ← Istorinių duomenų užpildymas
├── tests/
│   └── test_scraper.py      ← Unit testai (pytest)
├── fuel_tracker.xlsx         ← Excel failas (atnaujinamas automatiškai)
├── latest_results.json       ← Paskutinio run rezultatai + šaltinių statusas
├── index.html                ← Dashboard (vanilla JS, offline cache)
├── requirements.txt          ← Python dependencies (pinned versions)
├── CONTRIBUTING.md           ← Kaip pridėti naujus šaltinius
├── .gitignore                ← Git ignore rules
└── README.md                 ← Šis failas
```
