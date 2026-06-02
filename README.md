# ⛽ Fuel Price Tracker

Automatinis kuro kainų sekimas iš 6 šaltinių. Veikia per GitHub Actions — jokio serverio, jokio rankinio darbo.

## Ką daro

Kas darbo dieną ~16:00 LT (13:00 UTC) laiku automatiškai:

| Šaltinis | Duomenys | Dažnumas |
|----------|----------|----------|
| ECB (frankfurter.app) | PLN/EUR, SEK/EUR kursai | Kasdien |
| Orlen PL | Ekodiesel hurtinė kaina PLN/m³ | Kasdien |
| Orlen LT | Dyzelinas C/RRME, **pardavimo kaina su PVM** (1-as term., Mažeikiai), EUR/l | Kasdien |
| Tankerkönig → EC Bulletin | VK dyzelinas (ELVIS DE proxy), EUR/l | Kasdien |
| ST1.se | BSH/ST1 Diesel SE, SEK/l | Kasdien |
| EC Oil Bulletin | LT/LV/EE/DK/SE/FI diesel kainos | Savaitinis |

Surinkti duomenys automatiškai įrašomi į `fuel_tracker.xlsx` ir commit'inami atgal į repo.

### Pastabos dėl šaltinių

- **Orlen LT** — imama „Pardavimo kaina su PVM" (5-as PDF stulpelis), tik kelių dyzelinas `Dyzelinas … su RRME` (žemės ūkio / laivų / krosnių eilutės atmetamos), 1-as terminalas (Mažeikiai). Scraper'is tikrina PDF vidinę datą („galioja nuo"): jei šiandienos PDF dar nepaskelbtas, reikšmė įrašoma į **tikros PDF datos** eilutę (taip išvengiama anksčiau buvusio „off-by-one"). Be/su PVM pasirenkama `config.py` → `ORLEN_LT_PRICE_COL`.
- **ELVIS DE** — tikras ELVIS Dieselfloater yra **tik partneriams** (BLUE.net), viešai neskelbiamas. Kaip viešas pakaitalas naudojamas **Tankerkönig** (oficialūs MTS-K degalinių dyzelino duomenys) — reikia nemokamo API rakto kaip GitHub Secret `TANKERKOENIG_API_KEY` (gauti: https://creativecommons.tankerkoenig.de/). Jei rakto nėra, naudojama EC Oil Bulletin Vokietijos dyzelino kaina.
- **Outlier apsauga** — jei kaina pasikeičia >15 % per dieną (`config.py` → `MAX_DAILY_CHANGE_PCT`), reikšmė vis tiek įrašoma, bet pažymima `SUSP:` Notes stulpelyje.
- **Istorijos taisymas** — `python src/backfill.py` (arba Actions → Run workflow → backfill=true) iš naujo atkuria teisingą Orlen LT kiekvienai datai pagal jos pačios PDF.

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
GitHub Actions (cron: 13:00 UTC ≈ 16:00 LT, Mon-Fri)
    │
    ├─→ Fetch FX rates (frankfurter.app API)
    ├─→ Scrape Orlen PL (petrodom.pl)
    ├─→ Scrape Orlen LT (orlenlietuva.lt — PDF data tikrinama)
    ├─→ Fetch DE diesel (Tankerkönig/MTS-K → EC Bulletin)
    ├─→ Scrape BSH/ST1 SE (st1.se)
    ├─→ Fetch EU Oil Bulletin (energy.ec.europa.eu, savaitinis)
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
