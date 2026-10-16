import re
from datetime import datetime
from typing import Any, Optional

# Supported date formats
DATE_FORMATS = [
    "%Y%m%d",            # 20260804
    "%Y-%m-%d",          # 2026-08-04
    "%d-%m-%Y",          # 04-08-2026
    "%d-%m-%y",          # 04-08-26
    "%d/%m/%Y",          # 04/08/2026
    "%d/%m/%y",          # 04/08/26
    "%d.%m.%Y",          # 04.08.2026
    "%d.%m.%y",          # 04.08.26
    "%d-%b-%Y",          # 04-Aug-2026
    "%d-%b-%y",          # 04-Aug-26
    "%d-%B-%Y",          # 04-August-2026
    "%d %b %Y",          # 04 Aug 2026
    "%d %B %Y",          # 04 August 2026
    "%d %b %y",          # 04 Aug 26
    "%d/%b/%Y",          # 04/Aug/2026
    "%Y/%m/%d",          # 2026/08/04
    "%Y.%m.%d",          # 2026.08.04
    "%m/%d/%Y",          # 08/04/2026
    "%m-%d-%Y",          # 08-04-2026
    "%Y-%m-%dT%H:%M:%S", # ISO
    "%Y-%m-%dT%H:%M:%S.%f",
]

TECHNICAL_FIELDS = {
    "alterid",
    "masterid",
    "vouchermodifieddate",
    "updateddatetime",
    "createddatetime",
    "remoteid",
    "vchkey",
    "guid",
    "_id",
    "audit_trail",
    "timestamp",
    "reservedname",
    "persistedview",
}


def normalize_date(date_val: Any) -> str:
    """
    Normalizes a date value into canonical 'YYYY-MM-DD' string.
    Raises ValueError if date cannot be parsed.
    """
    if date_val is None:
        raise ValueError("Missing date value")

    if isinstance(date_val, datetime):
        return date_val.strftime("%Y-%m-%d")

    val_str = str(date_val).strip()
    if not val_str:
        raise ValueError("Empty date value")

    # If already in ISO format with T, take the date part
    if "T" in val_str:
        val_str = val_str.split("T")[0]

    for fmt in DATE_FORMATS:
        try:
            dt = datetime.strptime(val_str, fmt)
            return dt.strftime("%Y-%m-%d")
        except ValueError:
            continue

    raise ValueError(f"Unrecognized date format: '{date_val}'. Expected YYYY-MM-DD, DD-MM-YYYY, DD-Mon-YYYY, or YYYYMMDD.")


def normalize_string(val: Optional[str]) -> str:
    """
    Normalizes string by stripping leading/trailing whitespace and collapsing multiple spaces.
    """
    if val is None:
        return ""
    return re.sub(r"\s+", " ", str(val).strip())


def is_technical_field(field_name: str) -> bool:
    """
    Returns True if the field name is known technical metadata.
    """
    return field_name.strip().lower() in TECHNICAL_FIELDS
