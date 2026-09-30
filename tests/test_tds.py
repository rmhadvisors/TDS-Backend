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
