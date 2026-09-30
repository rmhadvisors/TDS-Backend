import json
from typing import Any, Dict, List, Tuple, Union


class ValidationError(Exception):
    def __init__(self, message: str, details: List[str] = None):
        super().__init__(message)
        self.message = message
        self.details = details or [message]


def decode_multiformat_bytes(content: Union[bytes, bytearray]) -> str:
    """
    Decodes raw bytes into a string by detecting BOM headers, JSON null-byte
    signatures (for UTF-16/UTF-32 without BOM), and trying fallback encodings.
    """
    if isinstance(content, str):
        return content

    # 1. BOM-based Sniffing (Fast & 100% Accurate)
    if content.startswith(b"\xef\xbb\xbf"):
        return content.decode("utf-8-sig")
    elif content.startswith(b"\xff\xfe\x00\x00") or content.startswith(b"\x00\x00\xfe\xff"):
        return content.decode("utf-32")
    elif content.startswith(b"\xff\xfe") or content.startswith(b"\xfe\xff"):
        return content.decode("utf-16")

    # 2. Heuristic for UTF-16 / UTF-32 without BOM
    # In JSON text (starting with '{' or '[' or whitespace), UTF-16/32 contains alternating 0x00 bytes.
    if len(content) >= 4 and (content[1] == 0 and content[2] == 0 and content[3] == 0):
        try:
            return content.decode("utf-32-le")
        except UnicodeDecodeError:
            pass
    if len(content) >= 4 and (content[0] == 0 and content[1] == 0 and content[2] == 0):
        try:
            return content.decode("utf-32-be")
        except UnicodeDecodeError:
            pass
    if len(content) >= 2 and content[1] == 0:
        try:
            return content.decode("utf-16-le")
        except UnicodeDecodeError:
            pass
    if len(content) >= 2 and content[0] == 0:
        try:
            return content.decode("utf-16-be")
        except UnicodeDecodeError:
            pass

    # 3. Sequential Candidate Fallback (for standard UTF-8, Windows-1252, Latin-1)
    encodings_to_try = ["utf-8", "cp1252", "latin-1"]
    for enc in encodings_to_try:
        try:
            return content.decode(enc)
        except (UnicodeDecodeError, LookupError):
            continue

    raise ValidationError("File encoding error: Unable to decode file content with supported encodings (UTF-8, UTF-16, Latin-1, etc.)")


def parse_json_content(content: Union[str, bytes, bytearray]) -> Any:
    """
    Parses JSON string or bytes, raising ValidationError if corrupted.
    Supports UTF-8, UTF-16 LE/BE, UTF-32, Latin-1/CP1252.
    """
    if not content:
        raise ValidationError("File content is empty")

    if isinstance(content, (bytes, bytearray)):
        content = decode_multiformat_bytes(content)

    try:
        return json.loads(content)
    except json.JSONDecodeError as e:
        raise ValidationError(f"Invalid JSON format: {str(e)} at line {e.lineno}, col {e.colno}")


def extract_records(data: Any, record_type: str = "transaction") -> List[Dict[str, Any]]:
    """
    Extracts a list of record dictionaries from raw parsed JSON.
    Supports:
    1. Direct list: [{...}, {...}]
    2. Dict with envelope/keys: {"transactions": [...]}, {"vouchers": [...]}, {"masters": [...]}, {"tallymessage": [...]}
    3. Tally XML-to-JSON format: {"ENVELOPE": {"BODY": {"IMPORTDATA": {"REQUESTDATA": {"TALLYMESSAGE": [...]}}}}}
    """
    if isinstance(data, list):
        items = [item for item in data if isinstance(item, dict)]
    elif isinstance(data, dict):
        lower_map = {k.lower(): v for k, v in data.items()}
        found_list = None
        for key in ["transactions", "vouchers", "data", "records", "items", "entries", "master", "masters", "ledgers", "tallymessage"]:
            if key in lower_map and isinstance(lower_map[key], list):
                found_list = lower_map[key]
                break

        if not found_list:
            req_data = (
                lower_map.get("envelope", {})
                .get("body", {})
                .get("importdata", {})
                .get("requestdata", {})
            )
            if isinstance(req_data, dict):
                tmsg = req_data.get("tallymessage") or req_data.get("TALLYMESSAGE")
                if isinstance(tmsg, list):
                    found_list = tmsg
                elif isinstance(tmsg, dict):
                    found_list = [tmsg]

        if found_list is not None:
            items = found_list
        elif any(k.upper() in ["VOUCHERNUMBER", "GUID", "VCHKEY", "PARTYLEDGERNAME", "NAME"] for k in data.keys()):
            items = [data]
        else:
            raise ValidationError(f"Unable to locate '{record_type}' records in JSON structure")
    else:
        raise ValidationError(f"Expected JSON object or array, got {type(data).__name__}")

    records = []
    for item in items:
        if not isinstance(item, dict):
            continue
        if record_type == "transaction":
            vch = item.get("VOUCHER") or item.get("voucher")
            if vch and isinstance(vch, dict):
                records.append(vch)
            elif any(k.lower() in ["vouchernumber", "vouchertypename", "voucherkey", "allledgerentries", "partyledgername"] for k in item.keys()):
                records.append(item)
            elif "guid" in item and "date" in item:
                records.append(item)
        elif record_type == "master":
            led = item.get("LEDGER") or item.get("ledger")
            if led and isinstance(led, dict):
                records.append(led)
            elif any(k.lower() in ["parent", "languagename", "ledgername", "partyname", "incometaxnumber", "taxclassificationname", "gstapplicable"] for k in item.keys()):
                records.append(item)
            elif "name" in item or "NAME" in item:
                records.append(item)

    if not records and items:
        records = [it for it in items if isinstance(it, dict)]

    if not records:
        raise ValidationError(f"Unable to locate '{record_type}' records in JSON structure")

    return records


def validate_and_extract_transactions(raw_content: Union[str, bytes]) -> List[Dict[str, Any]]:
    """
    Parses and validates raw transaction JSON content.
    """
    parsed = parse_json_content(raw_content)
    records = extract_records(parsed, record_type="transaction")
    if not records:
        raise ValidationError("No transaction records found in transaction file")
    return records


def validate_and_extract_masters(raw_content: Union[str, bytes]) -> List[Dict[str, Any]]:
    """
    Parses and validates raw master JSON content.
    """
    parsed = parse_json_content(raw_content)
    records = extract_records(parsed, record_type="master")
    if not records:
        raise ValidationError("No master records found in master file")
    return records
