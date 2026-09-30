from decimal import Decimal
import pytest
from app.models.comparison_models import ClassificationStatus, NormalizedTransaction, TransactionComparison
from app.services.reconciliation import perform_reconciliation


def test_reconciliation_pass():
    t1 = NormalizedTransaction(
        guid="G1", voucher_number="1", voucher_type="Payment", date="2026-08-01",
        party_name="Party A", total_amount=Decimal("100000.00"), tds_amount=Decimal("10000.00")
    )
    t2 = NormalizedTransaction(
        guid="G1", voucher_number="1", voucher_type="Payment", date="2026-08-01",
        party_name="Party A", total_amount=Decimal("100000.00"), tds_amount=Decimal("10000.00")
    )
    comp = TransactionComparison(
        status=ClassificationStatus.UNCHANGED,
        old_transaction=t1,
        new_transaction=t2,
    )

    rec = perform_reconciliation([t1], [t2], [comp])
    assert rec.is_valid is True
    assert rec.transaction_count == "PASS"
    assert rec.amount == "PASS"
    assert rec.tds == "PASS"
    assert len(rec.errors) == 0


# Test 13: Reconciliation failure
def test_reconciliation_failure_detected():
    t1 = NormalizedTransaction(
        guid="G1", voucher_number="1", voucher_type="Payment", date="2026-08-01",
        party_name="Party A", total_amount=Decimal("100000.00"), tds_amount=Decimal("10000.00")
    )
    t2 = NormalizedTransaction(
        guid="G2", voucher_number="2", voucher_type="Payment", date="2026-08-01",
        party_name="Party B", total_amount=Decimal("200000.00"), tds_amount=Decimal("20000.00")
    )

    # Corrupted comparison: only 1 transaction is categorized instead of both
    comp = TransactionComparison(
        status=ClassificationStatus.MISSING,
        old_transaction=t1,
        new_transaction=None,
    )

    rec = perform_reconciliation([t1, t2], [], [comp])
    assert rec.is_valid is False
    assert rec.transaction_count == "FAIL"
    assert rec.amount == "FAIL"
    assert len(rec.errors) > 0
