import re
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from typing import Any, Union

TWO_PLACES = Decimal("0.01")


def parse_decimal(value: Any, default: Decimal = Decimal("0.00")) -> Decimal:
    """
    Safely converts input value (str, int, float, Decimal, list, dict) to Decimal with 2 decimal places.
    Handles Tally multi-currency strings (e.g. '-$265.00 @ ₹ 92.10/$ = -₹ 24406.50'),
    currency symbols, Dr/Cr indicators, parentheses, commas, and Tally metadata arrays.
    """
    if value is None:
        return default
    if isinstance(value, list):
        for item in value:
            if not isinstance(item, dict):
                return parse_decimal(item, default)
            elif "amount" in item:
                return parse_decimal(item["amount"], default)
            elif "value" in item:
                return parse_decimal(item["value"], default)
        return default
    if isinstance(value, dict):
        amt = value.get("amount") or value.get("value") or value.get("amt")
        if amt is not None:
            return parse_decimal(amt, default)
        return default
    if isinstance(value, Decimal):
        return value.quantize(TWO_PLACES, rounding=ROUND_HALF_UP)
    if isinstance(value, (int, float)):
        return Decimal(str(value)).quantize(TWO_PLACES, rounding=ROUND_HALF_UP)
    if isinstance(value, str):
        val_str = value.strip()
        if not val_str:
            return default

        # 1. If Tally forex equation format is used (e.g. '... = ...'), take the converted base currency amount (rightmost part after '=')
        if "=" in val_str:
            val_str = val_str.split("=")[-1].strip()

        # 2. Strip '/-' or '/=' currency suffixes (e.g., '10000/-')
        val_str = re.sub(r'/[=\-]\s*$', '', val_str).strip()

        # 3. Check for Dr / Cr suffixes
        is_negative = False
        val_upper = val_str.upper()
        if val_upper.endswith(" CR") or val_upper.endswith("CR"):
            val_str = re.sub(r'(?i)\s*cr$', '', val_str).strip()
        elif val_upper.endswith(" DR") or val_upper.endswith("DR"):
            val_str = re.sub(r'(?i)\s*dr$', '', val_str).strip()

        # 4. Check for accounting parentheses: (123.45) -> -123.45
        if val_str.startswith("(") and val_str.endswith(")"):
            is_negative = True
            val_str = val_str[1:-1].strip()

        # 5. Detect negative sign
        if "-" in val_str:
            is_negative = True

        # 6. Remove currency symbols, @ rate annotations, commas, spaces
        val_str = re.sub(r'[₹$€£¥\s/,\-+]', '', val_str)
        val_str = re.sub(r'(?i)\b(inr|usd|eur|gbp|aed|rs\.?)\b', '', val_str).strip()

        if not val_str:
            return default

        try:
            dec = Decimal(val_str)
            if is_negative:
                dec = -dec
            return dec.quantize(TWO_PLACES, rounding=ROUND_HALF_UP)
        except InvalidOperation:
            # Fallback regex extraction of numeric float from the string
            match = re.search(r'\d*\.?\d+', val_str)
            if match:
                try:
                    dec = Decimal(match.group())
                    if is_negative and dec > 0:
                        dec = -dec
                    return dec.quantize(TWO_PLACES, rounding=ROUND_HALF_UP)
                except InvalidOperation:
                    pass
            raise ValueError(f"Invalid monetary amount: '{value}'")

    try:
        return Decimal(str(value)).quantize(TWO_PLACES, rounding=ROUND_HALF_UP)
    except (InvalidOperation, TypeError):
        raise ValueError(f"Cannot parse value '{value}' as Decimal")


def format_currency(amount: Union[Decimal, int, float, str, None]) -> str:
    """
    Formats a Decimal into standard currency representation with 2 decimal places.
    """
    dec = parse_decimal(amount)
    return f"{dec:.2f}"
