from decimal import Decimal
from typing import List, Tuple
from app.models.comparison_models import FieldDifference, NormalizedLedger, NormalizedTransaction
from app.utils.amounts import parse_decimal


def are_ledgers_equivalent(l1: List[NormalizedLedger], l2: List[NormalizedLedger]) -> bool:
    """
    Compares two lists of normalized ledgers. Order does not matter because they are sorted canonically.
    Duplicate ledger entries are preserved and must match in multiplicity and values.
    """
    if len(l1) != len(l2):
        return False
    for a, b in zip(l1, l2):
        if not a.is_equivalent(b):
            return False
    return True


def compare_transactions(old_tx: NormalizedTransaction, new_tx: NormalizedTransaction) -> List[FieldDifference]:
    """
    Compares business fields of two transactions and returns a list of field differences.
    Monetary differences include calculated delta (New - Old).
    """
    diffs: List[FieldDifference] = []

    # 1. Date
    if old_tx.date != new_tx.date:
        diffs.append(FieldDifference(field="Date", old_value=old_tx.date, new_value=new_tx.date))

    # 2. Voucher Number
    if old_tx.voucher_number != new_tx.voucher_number:
        diffs.append(FieldDifference(field="Voucher Number", old_value=old_tx.voucher_number, new_value=new_tx.voucher_number))

    # 3. Voucher Type
    if old_tx.voucher_type.strip().lower() != new_tx.voucher_type.strip().lower():
        diffs.append(FieldDifference(field="Voucher Type", old_value=old_tx.voucher_type, new_value=new_tx.voucher_type))

    # 4. Party Name
    if old_tx.party_name.strip().lower() != new_tx.party_name.strip().lower():
        diffs.append(FieldDifference(field="Party", old_value=old_tx.party_name, new_value=new_tx.party_name))

    # 5. Total Amount
    if old_tx.total_amount != new_tx.total_amount:
        diffs.append(
            FieldDifference(
                field="Total Amount",
                old_value=old_tx.total_amount,
                new_value=new_tx.total_amount,
                difference=new_tx.total_amount - old_tx.total_amount,
            )
        )

    # 6. TDS Amount
    if old_tx.tds_amount != new_tx.tds_amount:
        diffs.append(
            FieldDifference(
                field="TDS Amount",
                old_value=old_tx.tds_amount,
                new_value=new_tx.tds_amount,
                difference=new_tx.tds_amount - old_tx.tds_amount,
            )
        )

    # 7. TDS Section
    if (old_tx.tds_section or "").strip() != (new_tx.tds_section or "").strip():
        diffs.append(
            FieldDifference(
                field="TDS Section",
                old_value=old_tx.tds_section or "N/A",
                new_value=new_tx.tds_section or "N/A",
            )
        )

    # 8. TDS Rate
    if old_tx.tds_rate != new_tx.tds_rate:
        diffs.append(
            FieldDifference(
                field="TDS Rate",
                old_value=f"{old_tx.tds_rate}%" if old_tx.tds_rate is not None else "N/A",
                new_value=f"{new_tx.tds_rate}%" if new_tx.tds_rate is not None else "N/A",
            )
        )

    # 9. Narration
    if (old_tx.narration or "").strip().lower() != (new_tx.narration or "").strip().lower():
        diffs.append(
            FieldDifference(
                field="Narration",
                old_value=old_tx.narration or "",
                new_value=new_tx.narration or "",
            )
        )

    # 10. Ledger Structure & Line-item amounts
    if not are_ledgers_equivalent(old_tx.ledgers, new_tx.ledgers):
        old_ledger_summary = ", ".join([f"{l.ledger_name}: ₹{l.amount}" for l in old_tx.ledgers])
        new_ledger_summary = ", ".join([f"{l.ledger_name}: ₹{l.amount}" for l in new_tx.ledgers])
        diffs.append(
            FieldDifference(
                field="Ledger Entries",
                old_value=old_ledger_summary,
                new_value=new_ledger_summary,
            )
        )

    return diffs
