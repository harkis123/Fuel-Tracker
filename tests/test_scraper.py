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
    ORLEN_PL_MAX,
    ORLEN_PL_MIN,
)
from scraper import clean_num, validate_fx, validate_price_change


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
