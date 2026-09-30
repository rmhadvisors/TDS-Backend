from collections import defaultdict
from decimal import Decimal
from typing import Dict, List, Tuple

from app.models.comparison_models import (
    ClassificationStatus,
    ComparisonSummarySchema,
    PartySummary,
    TransactionComparison,
)


def calculate_party_summaries(
    comparisons: List[TransactionComparison],
) -> List[PartySummary]:
    """
    Groups transaction comparison results by party and computes monetary and count aggregates.
    """
    party_data: Dict[str, Dict[str, Any]] = defaultdict(
        lambda: {
            "affected_entries": 0,
            "missing_count": 0,
            "new_count": 0,
            "modified_count": 0,
            "recreated_count": 0,
            "original_amount": Decimal("0.00"),
            "updated_amount": Decimal("0.00"),
            "original_tds": Decimal("0.00"),
            "updated_tds": Decimal("0.00"),
        }
    )

    for item in comparisons:
        party = "Unknown Party"
        if item.old_transaction and item.old_transaction.party_name:
            party = item.old_transaction.party_name
        elif item.new_transaction and item.new_transaction.party_name:
            party = item.new_transaction.party_name

        rec = party_data[party]

        if item.status == ClassificationStatus.MISSING:
            rec["affected_entries"] += 1
            rec["missing_count"] += 1
            rec["original_amount"] += item.old_transaction.total_amount
            rec["original_tds"] += item.old_transaction.tds_amount

        elif item.status == ClassificationStatus.NEW:
            rec["affected_entries"] += 1
            rec["new_count"] += 1
            rec["updated_amount"] += item.new_transaction.total_amount
            rec["updated_tds"] += item.new_transaction.tds_amount

        elif item.status == ClassificationStatus.MODIFIED:
            rec["affected_entries"] += 1
            rec["modified_count"] += 1
            rec["original_amount"] += item.old_transaction.total_amount
            rec["updated_amount"] += item.new_transaction.total_amount
            rec["original_tds"] += item.old_transaction.tds_amount
            rec["updated_tds"] += item.new_transaction.tds_amount

        elif item.status == ClassificationStatus.RECREATED:
            rec["affected_entries"] += 1
            rec["recreated_count"] += 1
            rec["original_amount"] += item.old_transaction.total_amount
            rec["updated_amount"] += item.new_transaction.total_amount
            rec["original_tds"] += item.old_transaction.tds_amount
            rec["updated_tds"] += item.new_transaction.tds_amount

        elif item.status == ClassificationStatus.UNRESOLVED:
            rec["affected_entries"] += 1
            if item.old_transaction:
                rec["original_amount"] += item.old_transaction.total_amount
                rec["original_tds"] += item.old_transaction.tds_amount

        elif item.status == ClassificationStatus.UNCHANGED:
            rec["original_amount"] += item.old_transaction.total_amount
            rec["updated_amount"] += item.new_transaction.total_amount
            rec["original_tds"] += item.old_transaction.tds_amount
            rec["updated_tds"] += item.new_transaction.tds_amount

    summaries: List[PartySummary] = []
    for party, data in sorted(party_data.items(), key=lambda x: x[0].lower()):
        amt_diff = data["updated_amount"] - data["original_amount"]
        tds_diff = data["updated_tds"] - data["original_tds"]
        summaries.append(
            PartySummary(
                party=party,
                affected_entries=data["affected_entries"],
                missing_count=data["missing_count"],
                new_count=data["new_count"],
                modified_count=data["modified_count"],
                recreated_count=data["recreated_count"],
                original_amount=data["original_amount"],
                updated_amount=data["updated_amount"],
                amount_difference=amt_diff,
                original_tds=data["original_tds"],
                updated_tds=data["updated_tds"],
                tds_difference=tds_diff,
            )
        )

    return summaries


def build_comparison_summary(
    comparisons: List[TransactionComparison],
    total_old_count: int,
    total_new_count: int,
) -> ComparisonSummarySchema:
    """
    Computes high-level summary metrics across all comparisons.
    """
    counts = {
        ClassificationStatus.UNCHANGED: 0,
        ClassificationStatus.MODIFIED: 0,
        ClassificationStatus.NEW: 0,
        ClassificationStatus.MISSING: 0,
        ClassificationStatus.RECREATED: 0,
        ClassificationStatus.UNRESOLVED: 0,
    }

    orig_total_amt = Decimal("0.00")
    upd_total_amt = Decimal("0.00")
    orig_total_tds = Decimal("0.00")
    upd_total_tds = Decimal("0.00")

    for item in comparisons:
        counts[item.status] = counts.get(item.status, 0) + 1

        if item.old_transaction:
            orig_total_amt += item.old_transaction.total_amount
            orig_total_tds += item.old_transaction.tds_amount

        if item.new_transaction:
            upd_total_amt += item.new_transaction.total_amount
            upd_total_tds += item.new_transaction.tds_amount

    amt_diff = upd_total_amt - orig_total_amt
    tds_diff = upd_total_tds - orig_total_tds

    return ComparisonSummarySchema(
        original_transactions=total_old_count,
        updated_transactions=total_new_count,
        unchanged=counts[ClassificationStatus.UNCHANGED],
        modified=counts[ClassificationStatus.MODIFIED],
        new=counts[ClassificationStatus.NEW],
        missing=counts[ClassificationStatus.MISSING],
        recreated=counts[ClassificationStatus.RECREATED],
        unresolved=counts[ClassificationStatus.UNRESOLVED],
        original_amount=f"{orig_total_amt:.2f}",
        updated_amount=f"{upd_total_amt:.2f}",
        amount_difference=f"{amt_diff:.2f}",
        original_tds=f"{orig_total_tds:.2f}",
        updated_tds=f"{upd_total_tds:.2f}",
        tds_difference=f"{tds_diff:.2f}",
    )
