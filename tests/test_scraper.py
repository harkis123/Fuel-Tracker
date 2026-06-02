"""Unit tests for Fuel Price Tracker scraper."""
import sys
from pathlib import Path

import pytest

# Add src to path so we can import modules
sys.path.insert(0, str(Path(__file__).parent.parent / "src"))

from config import (
    BSH_SE_MAX,
    BSH_SE_MIN,
    DIESEL_EUR_MAX,
    DIESEL_EUR_MIN,
    FX_PLN_EUR_MAX,
    FX_PLN_EUR_MIN,
    FX_SEK_EUR_MAX,
    FX_SEK_EUR_MIN,
    ORLEN_PL_MAX,
    ORLEN_PL_MIN,
)
from scraper import clean_num, validate_fx, validate_price_change, _find_orlen_lt_prices


class TestCleanNum:
    def test_simple_integer(self):
        assert clean_num("1234") == 1234.0

    def test_simple_float(self):
        assert clean_num("12.34") == 12.34

    def test_spaces_as_thousand_sep_dot_decimal(self):
        assert clean_num("1 756.06") == 1756.06

    def test_spaces_as_thousand_sep_comma_decimal(self):
        assert clean_num("1 756,06") == 1756.06

    def test_comma_and_dot(self):
        assert clean_num("1,234.56") == 1234.56

    def test_comma_as_decimal(self):
        assert clean_num("1234,56") == 1234.56

    def test_currency_euro(self):
        assert clean_num("€1.234") == 1.234

    def test_currency_dollar(self):
        assert clean_num("$99.99") == 99.99

    def test_none_input(self):
        assert clean_num(None) is None

    def test_empty_string(self):
        assert clean_num("") is None

    def test_non_numeric(self):
        assert clean_num("abc") is None

    def test_nbsp_removal(self):
        assert clean_num("1\xa0756.06") == 1756.06

    def test_large_number(self):
        assert clean_num("6 870.00") == 6870.0

    def test_integer_with_spaces(self):
        assert clean_num("6 192") == 6192.0


class TestValidateFx:
    def test_valid_rates(self):
        assert validate_fx(4.27, 10.84) is True

    def test_pln_too_low(self):
        assert validate_fx(2.0, 10.84) is False

    def test_pln_too_high(self):
        assert validate_fx(7.0, 10.84) is False

    def test_sek_too_low(self):
        assert validate_fx(4.27, 8.0) is False

    def test_sek_too_high(self):
        assert validate_fx(4.27, 14.0) is False

    def test_boundary_values(self):
        assert validate_fx(FX_PLN_EUR_MIN, FX_SEK_EUR_MIN) is True
        assert validate_fx(FX_PLN_EUR_MAX, FX_SEK_EUR_MAX) is True


class TestValidatePriceChange:
    def test_small_change(self):
        assert validate_price_change("Test", 100.0, 99.0) is True

    def test_large_change(self):
        assert validate_price_change("Test", 120.0, 100.0) is False

    def test_no_previous(self):
        assert validate_price_change("Test", 100.0, None) is True

    def test_zero_previous(self):
        assert validate_price_change("Test", 100.0, 0) is True

    def test_exact_threshold(self):
        # 15% change
        assert validate_price_change("Test", 115.0, 100.0) is False

    def test_negative_change(self):
        assert validate_price_change("Test", 80.0, 100.0) is False


class TestPriceRanges:
    """Verify configuration ranges are sensible."""

    def test_orlen_pl_range(self):
        assert ORLEN_PL_MIN > 0
        assert ORLEN_PL_MAX > ORLEN_PL_MIN

    def test_diesel_eur_range(self):
        assert DIESEL_EUR_MIN > 0
        assert DIESEL_EUR_MAX > DIESEL_EUR_MIN
        assert DIESEL_EUR_MAX < 5.0  # sanity

    def test_bsh_se_range(self):
        assert BSH_SE_MIN > 0
        assert BSH_SE_MAX > BSH_SE_MIN

    def test_fx_ranges(self):
        assert FX_PLN_EUR_MIN > 0
        assert FX_PLN_EUR_MAX > FX_PLN_EUR_MIN


# Real 2026-06-01 PDF excerpt: 2 road-diesel terminals + agri/marine decoys
ORLEN_LT_SAMPLE = (
    "Kainos galioja nuo 2026-06-01 9:00 val\n"
    "Dyzelinas C kl. su RRME 836.77 503.60 1 340.37 281.48 1 621.85\n"
    "Dyzelinas žemės ūkiui C kl. su RRME 836.77 35.00 871.77 183.07 1 054.84\n"
    "Dyzelinas laivų atsargoms C kl. su RRME 836.77 0.00 836.77 175.72 1 012.49\n"
    "Dyzelinas C kl. su RRME 841.44 503.60 1 345.04 282.46 1 627.50\n"
)


class TestOrlenLTParser:
    def test_selects_gross_su_pvm_first_terminal(self):
        r = _find_orlen_lt_prices(ORLEN_LT_SAMPLE)
        assert r is not None
        assert r["price_eur_1000l_su_pvm"] == 1621.85   # 5th column, terminal 0
        assert r["price_eur_1000l_be_pvm"] == 1340.37    # 3rd column (net)
        assert abs(r["price_eur_l"] - 1.6219) < 0.0005

    def test_excludes_agri_and_marine(self):
        r = _find_orlen_lt_prices(ORLEN_LT_SAMPLE)
        # must never pick the cheaper agri (1054.84) or marine (1012.49) lines
        assert r["price_eur_1000l_su_pvm"] not in (1054.84, 1012.49)

    def test_extracts_pdf_date(self):
        r = _find_orlen_lt_prices(ORLEN_LT_SAMPLE)
        assert r["pdf_date"] == "2026-06-01"

    def test_no_match_returns_none(self):
        assert _find_orlen_lt_prices("Automobilinis 95 benzinas 700.00 1 200.00") is None
