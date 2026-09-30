from decimal import Decimal
import pytest
from app.models.comparison_models import ClassificationStatus, NormalizedLedger, NormalizedTransaction
from app.services.transaction_matcher import match_and_compare_transactions


def make_sample_tx(
    guid: str,
    vch_no: str = "100",
    date: str = "2026-08-10",
    vch_type: str = "Payment",
    party: str = "ABC Ltd",
    total_amount: Decimal = Decimal("100000.00"),
    tds_amount: Decimal = Decimal("10000.00"),
    ledgers: list = None,
) -> NormalizedTransaction:
    if ledgers is None:
        ledgers = [
            NormalizedLedger(ledger_name=party, amount=total_amount - tds_amount, is_debit=False, is_credit=True),
            NormalizedLedger(ledger_name="TDS on Professional Fees", amount=tds_amount, is_debit=False, is_credit=True, is_tds=True),
            NormalizedLedger(ledger_name="Professional Fees Expense", amount=total_amount, is_debit=True, is_credit=False),
        ]
        ledgers = sorted(ledgers, key=lambda l: l.sort_key())

    return NormalizedTransaction(
        guid=guid,
        voucher_number=vch_no,
        voucher_type=vch_type,
        date=date,
        party_name=party,
        ledgers=ledgers,
        total_amount=total_amount,
        tds_amount=tds_amount,
    )


# Test 1: Same GUID + identical data -> UNCHANGED
def test_same_guid_identical_data():
    t1 = make_sample_tx("GUID-A")
    t2 = make_sample_tx("GUID-A")
    results = match_and_compare_transactions([t1], [t2])
    assert len(results) == 1
    assert results[0].status == ClassificationStatus.UNCHANGED
    assert results[0].amount_difference == Decimal("0.00")
    assert results[0].tds_difference == Decimal("0.00")


# Test 2: Same GUID + amount changed -> MODIFIED
def test_same_guid_amount_changed():
    t1 = make_sample_tx("GUID-A", total_amount=Decimal("100000.00"), tds_amount=Decimal("10000.00"))
    t2 = make_sample_tx("GUID-A", total_amount=Decimal("120000.00"), tds_amount=Decimal("12000.00"))
    results = match_and_compare_transactions([t1], [t2])
    assert len(results) == 1
    assert results[0].status == ClassificationStatus.MODIFIED
    assert results[0].amount_difference == Decimal("20000.00")
    assert results[0].tds_difference == Decimal("2000.00")
    assert any(d.field == "Total Amount" for d in results[0].field_changes)


# Test 3: Old-only transaction -> MISSING
def test_old_only_missing():
    t1 = make_sample_tx("GUID-A", total_amount=Decimal("50000.00"), tds_amount=Decimal("5000.00"))
    results = match_and_compare_transactions([t1], [])
    assert len(results) == 1
    assert results[0].status == ClassificationStatus.MISSING
    assert results[0].amount_difference == Decimal("-50000.00")
    assert results[0].tds_difference == Decimal("-5000.00")


# Test 4: New-only transaction -> NEW
def test_new_only_new():
    t2 = make_sample_tx("GUID-B", total_amount=Decimal("75000.00"), tds_amount=Decimal("7500.00"))
    results = match_and_compare_transactions([], [t2])
    assert len(results) == 1
    assert results[0].status == ClassificationStatus.NEW
    assert results[0].amount_difference == Decimal("75000.00")
    assert results[0].tds_difference == Decimal("7500.00")


# Test 5: Different GUID + same business transaction -> RECREATED
def test_different_guid_same_business_tx():
    t1 = make_sample_tx("GUID-A", vch_no="100", date="2026-08-10")
    t2 = make_sample_tx("GUID-B", vch_no="100", date="2026-08-10")
    results = match_and_compare_transactions([t1], [t2])
    assert len(results) == 1
    assert results[0].status == ClassificationStatus.RECREATED
    assert results[0].old_transaction.guid == "GUID-A"
    assert results[0].new_transaction.guid == "GUID-B"


# Test 6: Voucher number changed but rest matches -> RECREATED
def test_voucher_number_changed_rest_matches():
    t1 = make_sample_tx("GUID-A", vch_no="50", date="2026-08-04", total_amount=Decimal("50000.00"), tds_amount=Decimal("5000.00"))
    t2 = make_sample_tx("GUID-B", vch_no="55", date="2026-08-04", total_amount=Decimal("50000.00"), tds_amount=Decimal("5000.00"))
    results = match_and_compare_transactions([t1], [t2])
    assert len(results) == 1
    assert results[0].status == ClassificationStatus.RECREATED
    assert "Voucher number changed" in results[0].match_reason


# Test 7: Ledger order changed -> UNCHANGED
def test_ledger_order_changed_unchanged():
    l1 = [
        NormalizedLedger("ABC Ltd", Decimal("90000.00"), False, True),
        NormalizedLedger("TDS Payable", Decimal("10000.00"), False, True, is_tds=True),
        NormalizedLedger("Rent Expense", Decimal("100000.00"), True, False),
    ]
    # Pass 2 has reversed initial order
    l2 = [
        NormalizedLedger("Rent Expense", Decimal("100000.00"), True, False),
        NormalizedLedger("ABC Ltd", Decimal("90000.00"), False, True),
        NormalizedLedger("TDS Payable", Decimal("10000.00"), False, True, is_tds=True),
    ]
    t1 = NormalizedTransaction(
        guid="GUID-1", voucher_number="101", voucher_type="Payment", date="2026-08-01",
        party_name="ABC Ltd", ledgers=sorted(l1, key=lambda l: l.sort_key()), total_amount=Decimal("100000.00"), tds_amount=Decimal("10000.00")
    )
    t2 = NormalizedTransaction(
        guid="GUID-1", voucher_number="101", voucher_type="Payment", date="2026-08-01",
        party_name="ABC Ltd", ledgers=sorted(l2, key=lambda l: l.sort_key()), total_amount=Decimal("100000.00"), tds_amount=Decimal("10000.00")
    )
    results = match_and_compare_transactions([t1], [t2])
    assert len(results) == 1
    assert results[0].status == ClassificationStatus.UNCHANGED


# Test 8: Amount changed -> MODIFIED
def test_amount_changed():
    t1 = make_sample_tx("GUID-1", total_amount=Decimal("10000.00"), tds_amount=Decimal("1000.00"))
    t2 = make_sample_tx("GUID-1", total_amount=Decimal("15000.00"), tds_amount=Decimal("1000.00"))
    results = match_and_compare_transactions([t1], [t2])
    assert len(results) == 1
    assert results[0].status == ClassificationStatus.MODIFIED
    assert results[0].amount_difference == Decimal("5000.00")


# Test 9: TDS changed -> MODIFIED and exact TDS difference
def test_tds_changed():
    t1 = make_sample_tx("GUID-1", total_amount=Decimal("100000.00"), tds_amount=Decimal("10000.00"))
    t2 = make_sample_tx("GUID-1", total_amount=Decimal("100000.00"), tds_amount=Decimal("20000.00"))
    results = match_and_compare_transactions([t1], [t2])
    assert len(results) == 1
    assert results[0].status == ClassificationStatus.MODIFIED
    assert results[0].tds_difference == Decimal("10000.00")
    assert results[0].amount_difference == Decimal("0.00")


# Test 10: Ambiguous candidate -> UNRESOLVED
def test_ambiguous_candidates_unresolved():
    # 1 old transaction vs 2 potential matching new vouchers with different voucher numbers and GUIDs
    t_old = make_sample_tx("GUID-OLD", vch_no="V-1", date="2026-08-10", party="Ambiguous Ltd", total_amount=Decimal("50000.00"))
    t_new1 = make_sample_tx("GUID-NEW-1", vch_no="V-2", date="2026-08-10", party="Ambiguous Ltd", total_amount=Decimal("50000.00"))
    t_new2 = make_sample_tx("GUID-NEW-2", vch_no="V-3", date="2026-08-10", party="Ambiguous Ltd", total_amount=Decimal("50000.00"))

    results = match_and_compare_transactions([t_old], [t_new1, t_new2])
    unres = [r for r in results if r.status == ClassificationStatus.UNRESOLVED]
    assert len(unres) >= 1
    assert unres[0].old_transaction.guid == "GUID-OLD"
    assert "GUID-NEW-1" in unres[0].candidates
    assert "GUID-NEW-2" in unres[0].candidates
