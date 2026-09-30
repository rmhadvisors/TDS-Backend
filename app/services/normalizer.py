from decimal import Decimal
from typing import Any, Dict, List, Optional, Set
import re

from app.models.comparison_models import (
    NormalizedLedger,
    NormalizedMaster,
    NormalizedTransaction,
)
from app.services.parser import ValidationError
from app.utils.amounts import parse_decimal
from app.utils.normalization import is_technical_field, normalize_date, normalize_string


def _get_first_key(d: Dict[str, Any], keys: List[str]) -> Optional[Any]:
    """Helper to find first matching key in dictionary case-insensitively."""
    lower_map = {k.lower(): v for k, v in d.items()}
    for k in keys:
        if k.lower() in lower_map and lower_map[k.lower()] is not None:
            return lower_map[k.lower()]
    return None


def is_tds_ledger(name: str, entry_dict: Dict[str, Any]) -> bool:
    """
    Determines if a ledger entry represents a TDS deduction ledger.
    """
    if not name:
        return False
    name_upper = name.upper()
    if "TDS" in name_upper or "TAX DEDUCTED" in name_upper:
        return True
    
    # Check explicit keys in raw entry
    for key in ["istds", "is_tds", "tdsapplicable", "tds_section", "tdssection"]:
        val = _get_first_key(entry_dict, [key])
        if val is not None and str(val).lower() in ["yes", "true", "1"]:
            return True
        if key in ["tds_section", "tdssection"] and val:
            return True
            
    return False


def normalize_transaction(raw: Dict[str, Any], index: int) -> NormalizedTransaction:
    """
    Converts raw transaction JSON dictionary into canonical NormalizedTransaction.
    Raises ValidationError with exact index and reason if required data is corrupted.
    """
    # 1. GUID
    guid = _get_first_key(raw, ["guid", "voucherkey", "vchkey", "id", "remoteid", "voucher_id"])
    if not guid:
        # If no explicit guid, create stable fallback from voucher details if possible, or flag error
        vch_num = _get_first_key(raw, ["vouchernumber", "voucher_no", "vchno", "number", "vch_no"])
        vch_date = _get_first_key(raw, ["date", "voucher_date", "vchdate"])
        if vch_num and vch_date:
            guid = f"VCH-{vch_num}-{vch_date}"
        else:
            raise ValidationError(f"Transaction at index {index} has no GUID or Voucher Identifier")
    guid = str(guid).strip()

    # 2. Voucher Number
    vch_number = _get_first_key(raw, ["vouchernumber", "voucher_no", "vchno", "number", "vch_no", "voucher_num"])
    if vch_number is None:
        raise ValidationError(f"Transaction (GUID: {guid}, index: {index}) has missing voucher number")
    vch_number = str(vch_number).strip()

    # 3. Voucher Type
    vch_type = _get_first_key(raw, ["vouchertypename", "voucher_type", "type", "vchtype", "voucher_typ_name"])
    vch_type = normalize_string(vch_type or "Journal")

    # 4. Date
    raw_date = _get_first_key(raw, ["date", "voucher_date", "vchdate", "vch_date"])
    if not raw_date:
        raise ValidationError(f"Transaction (GUID: {guid}, index: {index}) has no voucher date")
    try:
        norm_date = normalize_date(raw_date)
    except ValueError as e:
        raise ValidationError(f"Transaction (GUID: {guid}, index: {index}) has invalid date: {str(e)}")

    # 5. Party Name
    party_name = _get_first_key(raw, ["partyledgername", "party_name", "party", "partyname", "party_ledger"])
    party_name = normalize_string(party_name or "")

    # 6. Narration
    narration = _get_first_key(raw, ["narration", "remarks", "description"])
    narration = normalize_string(narration) if narration else None

    # 7. Ledgers
    raw_ledgers = _get_first_key(
        raw, 
        ["allledgerentries.list", "ledgerentries.list", "ledgers", "ledger_entries", "allledgerentries", "all_ledger_entries"]
    )

    ledgers: List[NormalizedLedger] = []
    if raw_ledgers:
        if isinstance(raw_ledgers, dict):
            raw_ledgers = [raw_ledgers]
        elif not isinstance(raw_ledgers, list):
            raw_ledgers = []

        for l_idx, l_raw in enumerate(raw_ledgers):
            if not isinstance(l_raw, dict):
                continue
            l_name = _get_first_key(l_raw, ["ledgername", "ledger_name", "name", "ledger"])
            if not l_name:
                continue
            l_name = normalize_string(l_name)

            raw_amt = _get_first_key(l_raw, ["amount", "amt", "ledger_amount", "ledgeramount"])
            try:
                l_amount = parse_decimal(raw_amt, Decimal("0.00"))
            except ValueError as e:
                raise ValidationError(f"Transaction (GUID: {guid}, index: {index}) ledger '{l_name}' has invalid amount: {str(e)}")

            # In Tally: ISDEEMEDPOSITIVE == "Yes" -> Debit, "No" -> Credit
            # If negative amount is supplied, Tally uses negative amount for opposite sign
            is_deemed_positive = _get_first_key(l_raw, ["isdeemedpositive", "is_deemed_positive", "isdebit", "is_debit"])
            if is_deemed_positive is not None:
                is_debit = str(is_deemed_positive).strip().lower() in ["yes", "true", "1"]
                is_credit = not is_debit
            else:
                # Infer from sign of amount if not explicitly given
                if l_amount < 0:
                    is_debit = False
                    is_credit = True
                    l_amount = abs(l_amount)
                else:
                    is_debit = True
                    is_credit = False

            # Positive magnitude for ledger amount
            l_amount = abs(l_amount)

            is_tds = is_tds_ledger(l_name, l_raw)
            tds_sec = _get_first_key(l_raw, ["tdssection", "tds_section", "section"])
            tds_rate_val = _get_first_key(l_raw, ["tdsrate", "tds_rate", "rate"])
            tds_rate = parse_decimal(tds_rate_val, None) if tds_rate_val is not None else None

            ledgers.append(
                NormalizedLedger(
                    ledger_name=l_name,
                    amount=l_amount,
                    is_debit=is_debit,
                    is_credit=is_credit,
                    ledger_id=_get_first_key(l_raw, ["ledgerid", "ledger_id", "id"]),
                    is_tds=is_tds,
                    tds_rate=tds_rate,
                    tds_section=str(tds_sec).strip() if tds_sec else None,
                )
            )

    # If party name is still empty, derive from non-TDS ledger
    if not party_name and ledgers:
        non_tds = [l for l in ledgers if not l.is_tds]
        if non_tds:
            party_name = non_tds[0].ledger_name
        else:
            party_name = ledgers[0].ledger_name

    # Calculate Total Amount & TDS Amount
    # Total amount = party amount or expense base amount (gross)
    tds_amount = sum((l.amount for l in ledgers if l.is_tds), Decimal("0.00"))

    # Determine base/total amount
    # If explicit voucher amount given:
    vch_amt_val = _get_first_key(raw, ["amount", "total_amount", "voucher_amount", "vchamount"])
    if vch_amt_val is not None:
        try:
            total_amount = abs(parse_decimal(vch_amt_val))
        except ValueError:
            total_amount = Decimal("0.00")
    else:
        # Derive from party ledger or non-TDS ledgers
        party_ledgers = [l for l in ledgers if l.ledger_name.strip().lower() == party_name.strip().lower()]
        if party_ledgers:
            # Party ledger amount is net payable/receivable; gross is party + tds (if TDS deducted)
            party_amt = sum(l.amount for l in party_ledgers)
            # If TDS was deducted, gross total is party + TDS or non-TDS expense ledgers
            expense_ledgers = [l for l in ledgers if not l.is_tds and l.ledger_name.strip().lower() != party_name.strip().lower()]
            if expense_ledgers:
                total_amount = sum(l.amount for l in expense_ledgers)
            else:
                total_amount = party_amt + tds_amount if tds_amount > 0 else party_amt
        else:
            non_tds_amt = sum((l.amount for l in ledgers if not l.is_tds), Decimal("0.00"))
            total_amount = non_tds_amt if non_tds_amt > 0 else sum((l.amount for l in ledgers), Decimal("0.00"))

    # Extract transaction-level TDS section and rate if available
    tx_tds_sec = _get_first_key(raw, ["tdssection", "tds_section", "section"])
    tx_tds_rate = _get_first_key(raw, ["tdsrate", "tds_rate", "rate"])

    # Sort ledgers deterministically for order-independent equality
    sorted_ledgers = sorted(ledgers, key=lambda l: l.sort_key())

    return NormalizedTransaction(
        guid=guid,
        voucher_number=vch_number,
        voucher_type=vch_type,
        date=norm_date,
        party_name=party_name,
        party_id=_get_first_key(raw, ["partyid", "party_id"]),
        narration=narration,
        ledgers=sorted_ledgers,
        total_amount=total_amount,
        tds_amount=tds_amount,
        tds_rate=parse_decimal(tx_tds_rate, None) if tx_tds_rate is not None else None,
        tds_section=str(tx_tds_sec).strip() if tx_tds_sec else None,
        raw_reference=raw,
    )


def _extract_master_name(raw: Dict[str, Any]) -> Optional[str]:
    """Helper to extract ledger/master name from standard keys or Tally languagename structure."""
    val = _get_first_key(raw, ["name", "ledgername", "partyname", "ledger_name", "party_name"])
    if val:
        if isinstance(val, list):
            strs = [x for x in val if isinstance(x, str)]
            if strs:
                return strs[0]
        elif isinstance(val, str):
            return val

    lang = raw.get("languagename") or raw.get("LANGUAGENAME")
    if lang:
        if isinstance(lang, dict):
            lang = [lang]
        if isinstance(lang, list) and len(lang) > 0:
            for l_item in lang:
                if isinstance(l_item, dict) and "name" in l_item:
                    n = l_item["name"]
                    if isinstance(n, list):
                        strs = [x for x in n if isinstance(x, str)]
                        if strs:
                            return strs[0]
                    elif isinstance(n, str):
                        return n

    val = _get_first_key(raw, ["mailingname", "originalname", "guid"])
    if val:
        return str(val)
    return None


def normalize_master(raw: Dict[str, Any], index: int) -> NormalizedMaster:
    """
    Converts raw master JSON dictionary into canonical NormalizedMaster.
    """
    name = _extract_master_name(raw)
    if not name:
        raise ValidationError(f"Master record at index {index} has no Name")
    name = normalize_string(name)

    guid = _get_first_key(raw, ["guid", "masterid", "id", "ledger_id", "remoteid"])
    if not guid:
        guid = f"MST-{name.upper()}"
    guid = str(guid).strip()

    parent = _get_first_key(raw, ["parent", "group", "parentgroup", "parent_group"])
    pan = _get_first_key(raw, ["incometaxnumber", "pan", "pannumber", "pan_no", "it_number"])
    gstin = _get_first_key(raw, ["partygstin", "gstin", "gst_number", "gstin_uin"])
    is_tds = _get_first_key(raw, ["istdsapplicable", "is_tds_applicable", "tdsapplicable"])
    tds_sec = _get_first_key(raw, ["tdssection", "tds_section", "section"])
    tds_rate_name = _get_first_key(raw, ["tdsratename", "tds_rate_name", "rate_name"])

    # Filter out technical fields from raw_fields
    clean_raw = {k: v for k, v in raw.items() if not is_technical_field(k)}

    return NormalizedMaster(
        guid=guid,
        name=name,
        parent_group=normalize_string(parent) if parent else None,
        pan=str(pan).strip().upper() if pan else None,
        gstin=str(gstin).strip().upper() if gstin else None,
        is_tds_applicable=str(is_tds).strip().lower() in ["yes", "true", "1"] if is_tds is not None else None,
        tds_section=str(tds_sec).strip() if tds_sec else None,
        tds_rate_name=str(tds_rate_name).strip() if tds_rate_name else None,
        raw_fields=clean_raw,
    )
