# Contributing to Fuel Price Tracker

## Adding a New Country/Source

1. **Create a fetch function** in `src/scraper.py`:
   ```python
   def fetch_new_source() -> Optional[dict[str, float]]:
       """Scrape [source name] for [country] diesel price."""
       try:
           r = SESSION.get(url, timeout=REQUEST_TIMEOUT)
           r.raise_for_status()
           # ... parse price ...
           logger.info("New Source: price = %.3f EUR/l", price)
           return {"price_eur_l": price}
       except Exception as e:
           logger.error("New Source: %s", e)
           return None
   ```

2. **Add configuration** in `src/config.py`:
   - URL to `URLS` dict
   - Price validation range constants
   - Any country mappings needed

3. **Wire it into `main()`** in `src/scraper.py`:
   - Call your function and store the result
   - Add failure tracking
   - Add outlier detection if applicable

4. **Update Excel writer** if adding a new column:
   - Add the column in `update_excel()`
   - Update the dashboard `index.html` to display it

5. **Add tests** in `tests/test_scraper.py`

6. **Update `README.md`** source table

## Running Tests

```bash
pip install pytest
pytest tests/ -v
```

## Project Structure

```
src/
  config.py      # All URLs, thresholds, magic numbers
  scraper.py     # Main scraper with fetch_*() functions
  backfill.py    # One-time historical data filler
tests/
  test_scraper.py  # Unit tests
index.html         # Dashboard (vanilla JS, reads from GitHub)
```

## Excel Sheet Structure

### Daily Tracker (columns A-N)
| Col | Field          | Source   | Unit     |
|-----|----------------|----------|----------|
| A   | Date           | Auto     | YYYY-MM-DD |
| B   | Day            | Auto     | Mon-Sun  |
| C   | Orlen PL       | Input    | PLN/m³   |
| D   | PLN/EUR        | Input    | Rate     |
| E   | Orlen PL EUR/l | Calc     | EUR/l    |
| F   | Orlen LT       | Input    | EUR/l    |
| G   | Delta EUR/l    | Calc     | EUR/l    |
| H   | Delta %        | Calc     | %        |
| I   | Elvis DE       | Input    | EUR/l    |
| J   | BSH SE SEK/l   | Input    | SEK/l    |
| K   | SEK/EUR        | Input    | Rate     |
| L   | BSH SE EUR/l   | Calc     | EUR/l    |
| M   | Notes          | Auto     | Sources  |
| N   | Status         | Auto     | Auto     |

### Weekly Oil Bulletin (columns A-I)
| Col | Field      | Unit  |
|-----|------------|-------|
| A   | Week date  | YYYY-MM-DD |
| B-G | LT,LV,EE,DK,SE,FI | EUR/l |
| H   | EU Average | EUR/l |
| I   | LT vs EU   | %     |
