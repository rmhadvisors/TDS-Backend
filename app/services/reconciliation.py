from decimal import Decimal
from typing import List

from app.models.comparison_models import (
    ClassificationStatus,
    NormalizedTransaction,
    ReconciliationResult,
    TransactionComparison,
)


def perform_reconciliation(
    original_txs: List[NormalizedTransaction],
    updated_txs: List[NormalizedTransaction],
    comparisons: List[TransactionComparison],
) -> ReconciliationResult:
    """
    Mandatory mathematical validation of transaction counts, total amounts, and TDS amounts.
    Equations:
    Original total count == Unchanged + Modified + Missing + Recreated + Unresolved (old side)
    Updated total count  == Unchanged + Modified + New + Recreated + Unresolved (new side)
    """
    errors: List[str] = []

    # 1. Count checks
    orig_total_count = len(original_txs)
    upd_total_count = len(updated_txs)

    classified_orig_count = sum(1 for c in comparisons if c.old_transaction is not None)
    classified_upd_count = sum(1 for c in comparisons if c.new_transaction is not None)

    count_status = "PASS"
    if orig_total_count != classified_orig_count:
        count_status = "FAIL"
        errors.append(
            f"Original transaction count mismatch: Input has {orig_total_count} records, but categorized {classified_orig_count} records."
        )

    if upd_total_count != classified_upd_count:
        count_status = "FAIL"
        errors.append(
            f"Updated transaction count mismatch: Input has {upd_total_count} records, but categorized {classified_upd_count} records."
        )

    # 2. Total Amount checks
    raw_orig_amount = sum((tx.total_amount for tx in original_txs), Decimal("0.00"))
    raw_upd_amount = sum((tx.total_amount for tx in updated_txs), Decimal("0.00"))

    cat_orig_amount = sum((c.old_transaction.total_amount for c in comparisons if c.old_transaction is not None), Decimal("0.00"))
    cat_upd_amount = sum((c.new_transaction.total_amount for c in comparisons if c.new_transaction is not None), Decimal("0.00"))

    amt_status = "PASS"
    if raw_orig_amount != cat_orig_amount:
        amt_status = "FAIL"
        errors.append(
            f"Original amount mismatch: Input sum is ₹{raw_orig_amount}, but categorized sum is ₹{cat_orig_amount}."
        )

    if raw_upd_amount != cat_upd_amount:
        amt_status = "FAIL"
        errors.append(
            f"Updated amount mismatch: Input sum is ₹{raw_upd_amount}, but categorized sum is ₹{cat_upd_amount}."
        )

    # 3. TDS Amount checks
    raw_orig_tds = sum((tx.tds_amount for tx in original_txs), Decimal("0.00"))
    raw_upd_tds = sum((tx.tds_amount for tx in updated_txs), Decimal("0.00"))

    cat_orig_tds = sum((c.old_transaction.tds_amount for c in comparisons if c.old_transaction is not None), Decimal("0.00"))
    cat_upd_tds = sum((c.new_transaction.tds_amount for c in comparisons if c.new_transaction is not None), Decimal("0.00"))

    tds_status = "PASS"
    if raw_orig_tds != cat_orig_tds:
        tds_status = "FAIL"
        errors.append(
            f"Original TDS mismatch: Input sum is ₹{raw_orig_tds}, but categorized sum is ₹{cat_orig_tds}."
        )

    if raw_upd_tds != cat_upd_tds:
        tds_status = "FAIL"
        errors.append(
            f"Updated TDS mismatch: Input sum is ₹{raw_upd_tds}, but categorized sum is ₹{cat_upd_tds}."
        )

    is_valid = (count_status == "PASS" and amt_status == "PASS" and tds_status == "PASS")

    return ReconciliationResult(
        transaction_count=count_status,
        amount=amt_status,
        tds=tds_status,
        is_valid=is_valid,
        details={
            "orig_total_count": orig_total_count,
            "classified_orig_count": classified_orig_count,
            "upd_total_count": upd_total_count,
            "classified_upd_count": classified_upd_count,
            "raw_orig_amount": str(raw_orig_amount),
            "cat_orig_amount": str(cat_orig_amount),
            "raw_upd_amount": str(raw_upd_amount),
            "cat_upd_amount": str(cat_upd_amount),
            "raw_orig_tds": str(raw_orig_tds),
            "cat_orig_tds": str(cat_orig_tds),
            "raw_upd_tds": str(raw_upd_tds),
            "cat_upd_tds": str(cat_upd_tds),
        },
        errors=errors,
    )
