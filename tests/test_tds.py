from decimal import Decimal
import pytest
from app.models.comparison_models import ClassificationStatus, NormalizedTransaction, TransactionComparison
from app.services.tds_calculator import build_comparison_summary, calculate_party_summaries


def test_tds_party_aggregation():
    t_old = NormalizedTransaction(
        guid="G1", voucher_number="1", voucher_type="Payment", date="2026-08-01",
        party_name="Aayat Enterprises", total_amount=Decimal("500000.00"), tds_amount=Decimal("50000.00")
    )
    t_new = NormalizedTransaction(
        guid="G1", voucher_number="1", voucher_type="Payment", date="2026-08-01",
        party_name="Aayat Enterprises", total_amount=Decimal("550000.00"), tds_amount=Decimal("55000.00")
    )
    comp = TransactionComparison(
        status=ClassificationStatus.MODIFIED,
        old_transaction=t_old,
        new_transaction=t_new,
        amount_difference=Decimal("50000.00"),
        tds_difference=Decimal("5000.00"),
    )
    summaries = calculate_party_summaries([comp])
    assert len(summaries) == 1
    p = summaries[0]
    assert p.party == "Aayat Enterprises"
    assert p.affected_entries == 1
    assert p.modified_count == 1
    assert p.original_amount == Decimal("500000.00")
    assert p.updated_amount == Decimal("550000.00")
    assert p.amount_difference == Decimal("50000.00")
    assert p.original_tds == Decimal("50000.00")
    assert p.updated_tds == Decimal("55000.00")
    assert p.tds_difference == Decimal("5000.00")


def test_build_summary():
    t_old = NormalizedTransaction(
        guid="G1", voucher_number="1", voucher_type="Payment", date="2026-08-01",
        party_name="XYZ Corp", total_amount=Decimal("100000.00"), tds_amount=Decimal("10000.00")
    )
    t_new = NormalizedTransaction(
        guid="G1", voucher_number="1", voucher_type="Payment", date="2026-08-01",
        party_name="XYZ Corp", total_amount=Decimal("90000.00"), tds_amount=Decimal("9000.00")
    )
    comp = TransactionComparison(
        status=ClassificationStatus.MODIFIED,
        old_transaction=t_old,
        new_transaction=t_new,
        amount_difference=Decimal("-10000.00"),
        tds_difference=Decimal("-1000.00"),
    )
    summary = build_comparison_summary([comp], 1, 1)
    assert summary.original_transactions == 1
    assert summary.updated_transactions == 1
    assert summary.modified == 1
    assert summary.original_amount == "100000.00"
    assert summary.updated_amount == "90000.00"
    assert summary.amount_difference == "-10000.00"
    assert summary.original_tds == "10000.00"
    assert summary.updated_tds == "9000.00"
    assert summary.tds_difference == "-1000.00"


def test_parse_tally_forex_and_complex_amounts():
    from app.utils.amounts import parse_decimal

    # Tally multi-currency equation string
    res1 = parse_decimal("-$265.00 @ ₹ 92.10/$ = -₹ 24406.50")
    assert res1 == Decimal("-24406.50")

    res2 = parse_decimal("$265.00 @ 92.10 = ₹ 24406.50")
    assert res2 == Decimal("24406.50")

    # Dr / Cr suffixes
    res3 = parse_decimal("5000.00 Cr")
    assert res3 == Decimal("5000.00")

    # Accounting parentheses
    res4 = parse_decimal("(12500.50)")
    assert res4 == Decimal("-12500.50")

    # Standard symbols
    res5 = parse_decimal("₹ 1,50,000.00 /-")
    assert res5 == Decimal("150000.00")


def test_normalize_transaction_with_missing_voucher_number():
    from app.services.normalizer import normalize_transaction

    raw_tx = {
        "GUID": "3f3fa0de-cff8-43b7-a25c-63ec831b64f1-00000002",
        "DATE": "2025-05-10",
        "VOUCHERTYPENAME": "Journal",
        "PARTYLEDGERNAME": "Foreign Bank",
        "ALLLEDGERENTRIES.LIST": [
            {
                "LEDGERNAME": "FOREIGN BANK CHARGES",
                "AMOUNT": "-$265.00 @ ₹ 92.10/$ = -₹ 24406.50",
                "ISDEEMEDPOSITIVE": "Yes",
            },
            {
                "LEDGERNAME": "Foreign Bank",
                "AMOUNT": "$265.00 @ ₹ 92.10/$ = ₹ 24406.50",
                "ISDEEMEDPOSITIVE": "No",
            }
        ]
    }

    norm = normalize_transaction(raw_tx, index=1)
    assert norm.guid == "3f3fa0de-cff8-43b7-a25c-63ec831b64f1-00000002"
    assert norm.voucher_number == ""  # Gracefully empty
    assert len(norm.ledgers) == 2
    assert norm.ledgers[0].amount == Decimal("24406.50")
    assert norm.total_amount == Decimal("24406.50")


def test_is_tds_viable_transaction():
    from app.models.comparison_models import NormalizedLedger
    from app.services.normalizer import is_tds_viable_transaction

    # 1. Non-TDS transaction (should be excluded)
    tx_non_tds = NormalizedTransaction(
        guid="G-NON-TDS",
        voucher_number="V-1",
        voucher_type="Payment",
        date="2026-05-10",
        party_name="Retail Store",
        total_amount=Decimal("5000.00"),
        tds_amount=Decimal("0.00"),
        ledgers=[NormalizedLedger(ledger_name="Retail Store", amount=Decimal("5000.00"), is_debit=True, is_credit=False)]
    )
    assert is_tds_viable_transaction(tx_non_tds, tds_party_names=set()) is False

    # 2. Transaction with non-zero TDS amount (should be included)
    tx_tds_amount = NormalizedTransaction(
        guid="G-TDS-1",
        voucher_number="V-2",
        voucher_type="Payment",
        date="2026-05-10",
        party_name="Contractor Corp",
        total_amount=Decimal("100000.00"),
        tds_amount=Decimal("2000.00"),
    )
    assert is_tds_viable_transaction(tx_tds_amount, tds_party_names=set()) is True

    # 3. Transaction with TDS line ledger (should be included)
    tx_tds_ledger = NormalizedTransaction(
        guid="G-TDS-2",
        voucher_number="V-3",
        voucher_type="Journal",
        date="2026-05-10",
        party_name="Consultant LLC",
        total_amount=Decimal("50000.00"),
        tds_amount=Decimal("0.00"),
        ledgers=[
            NormalizedLedger(ledger_name="TDS Payble", amount=Decimal("5000.00"), is_debit=False, is_credit=True, is_tds=True)
        ]
    )
    assert is_tds_viable_transaction(tx_tds_ledger, tds_party_names=set()) is True

    # 4. Transaction with party marked as TDS applicable in Master records
    tx_tds_party = NormalizedTransaction(
        guid="G-TDS-3",
        voucher_number="V-4",
        voucher_type="Purchase",
        date="2026-05-10",
        party_name="Sky Dev Corp",
        total_amount=Decimal("75000.00"),
        tds_amount=Decimal("0.00"),
    )
    assert is_tds_viable_transaction(tx_tds_party, tds_party_names={"sky dev corp"}) is True
