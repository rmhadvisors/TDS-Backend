from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from typing import Any, Union

TWO_PLACES = Decimal("0.01")


def parse_decimal(value: Any, default: Decimal = Decimal("0.00")) -> Decimal:
    """
    Safely converts input value (str, int, float, Decimal) to Decimal with 2 decimal places.
    Throws ValueError if invalid monetary format.
    """
    if value is None:
        return default
    if isinstance(value, Decimal):
        return value.quantize(TWO_PLACES, rounding=ROUND_HALF_UP)
    if isinstance(value, (int, float)):
        # Convert float via str to prevent binary float inaccuracies
        return Decimal(str(value)).quantize(TWO_PLACES, rounding=ROUND_HALF_UP)
    if isinstance(value, str):
        cleaned = value.strip().replace(",", "").replace("₹", "").replace("$", "")
        if not cleaned:
            return default
        try:
            return Decimal(cleaned).quantize(TWO_PLACES, rounding=ROUND_HALF_UP)
        except InvalidOperation:
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
