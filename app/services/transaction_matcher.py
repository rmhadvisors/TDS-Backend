from collections import defaultdict
from decimal import Decimal
from typing import Dict, List, Set, Tuple

from app.models.comparison_models import (
    ClassificationStatus,
    FieldDifference,
    NormalizedTransaction,
    TransactionComparison,
)
from app.services.parser import ValidationError
from app.services.transaction_comparator import are_ledgers_equivalent, compare_transactions


def check_duplicate_guids(transactions: List[NormalizedTransaction], source_name: str) -> None:
    """
    Checks for duplicate GUIDs within a single dataset.
    Raises ValidationError if duplicates exist.
    """
    seen: Dict[str, int] = {}
    for tx in transactions:
        seen[tx.guid] = seen.get(tx.guid, 0) + 1

    duplicates = [f"GUID '{guid}' occurs {count} times" for guid, count in seen.items() if count > 1]
    if duplicates:
        raise ValidationError(
            f"Duplicate transaction GUIDs detected in {source_name}",
            details=duplicates
        )


def match_and_compare_transactions(
    old_txs: List[NormalizedTransaction],
    new_txs: List[NormalizedTransaction],
) -> List[TransactionComparison]:
    """
    Multi-pass deterministic transaction matching algorithm.
    Every transaction is classified into exactly one status:
    - UNCHANGED
    - MODIFIED
    - RECREATED
    - UNRESOLVED
    - MISSING
    - NEW
    """
    # 0. Duplicate check
    check_duplicate_guids(old_txs, "Original Snapshot")
    check_duplicate_guids(new_txs, "Updated Snapshot")

    results: List[TransactionComparison] = []

    old_by_guid: Dict[str, NormalizedTransaction] = {tx.guid: tx for tx in old_txs}
    new_by_guid: Dict[str, NormalizedTransaction] = {tx.guid: tx for tx in new_txs}

    matched_old_guids: Set[str] = set()
    matched_new_guids: Set[str] = set()

    # =========================================================================
    # PASS 1 — Exact GUID Match
    # =========================================================================
    for guid, old_tx in old_by_guid.items():
        if guid in new_by_guid:
            new_tx = new_by_guid[guid]
            matched_old_guids.add(guid)
            matched_new_guids.add(guid)

            diffs = compare_transactions(old_tx, new_tx)
            if not diffs:
                results.append(
                    TransactionComparison(
                        status=ClassificationStatus.UNCHANGED,
                        old_transaction=old_tx,
                        new_transaction=new_tx,
                        match_reason="Exact GUID match; identical business fields",
                        amount_difference=Decimal("0.00"),
                        tds_difference=Decimal("0.00"),
                    )
                )
            else:
                amt_diff = new_tx.total_amount - old_tx.total_amount
                tds_diff = new_tx.tds_amount - old_tx.tds_amount
                results.append(
                    TransactionComparison(
                        status=ClassificationStatus.MODIFIED,
                        old_transaction=old_tx,
                        new_transaction=new_tx,
                        match_reason="Exact GUID match; business fields modified",
                        field_changes=diffs,
                        amount_difference=amt_diff,
                        tds_difference=tds_diff,
                    )
                )

    # Remaining unmatched sets
    unmatched_old = [tx for tx in old_txs if tx.guid not in matched_old_guids]
    unmatched_new = [tx for tx in new_txs if tx.guid not in matched_new_guids]

    # =========================================================================
    # PASS 2 — Business Key Candidate Match (Voucher No + Date + Type + Party)
    # =========================================================================
    # Group unmatched by candidate key
    def make_biz_key(tx: NormalizedTransaction) -> Tuple[str, str, str, str]:
        return (
            tx.voucher_number.strip().lower(),
            tx.date,
            tx.voucher_type.strip().lower(),
            tx.party_name.strip().lower(),
        )

    old_biz_map: Dict[Tuple[str, str, str, str], List[NormalizedTransaction]] = defaultdict(list)
    for tx in unmatched_old:
        old_biz_map[make_biz_key(tx)].append(tx)

    new_biz_map: Dict[Tuple[str, str, str, str], List[NormalizedTransaction]] = defaultdict(list)
    for tx in unmatched_new:
        new_biz_map[make_biz_key(tx)].append(tx)

    pass2_matched_old: Set[str] = set()
    pass2_matched_new: Set[str] = set()

    for biz_key, old_candidates in old_biz_map.items():
        if biz_key in new_biz_map:
            new_candidates = new_biz_map[biz_key]
            # Unique 1-to-1 match with identical ledger structure
            if len(old_candidates) == 1 and len(new_candidates) == 1:
                o_tx = old_candidates[0]
                n_tx = new_candidates[0]
                if are_ledgers_equivalent(o_tx.ledgers, n_tx.ledgers):
                    pass2_matched_old.add(o_tx.guid)
                    pass2_matched_new.add(n_tx.guid)
                    amt_diff = n_tx.total_amount - o_tx.total_amount
                    tds_diff = n_tx.tds_amount - o_tx.tds_amount
                    results.append(
                        TransactionComparison(
                            status=ClassificationStatus.RECREATED,
                            old_transaction=o_tx,
                            new_transaction=n_tx,
                            match_reason="Business key match: voucher number + date + voucher type + party + identical ledger structure (GUID changed)",
                            amount_difference=amt_diff,
                            tds_difference=tds_diff,
                        )
                    )

    matched_old_guids.update(pass2_matched_old)
    matched_new_guids.update(pass2_matched_new)
    unmatched_old = [tx for tx in unmatched_old if tx.guid not in matched_old_guids]
    unmatched_new = [tx for tx in unmatched_new if tx.guid not in matched_new_guids]

    # =========================================================================
    # PASS 3 — Recreated with Changed Voucher Number
    # =========================================================================
    def make_content_key(tx: NormalizedTransaction) -> Tuple[str, str, str, Decimal, Decimal]:
        return (
            tx.date,
            tx.voucher_type.strip().lower(),
            tx.party_name.strip().lower(),
            tx.total_amount,
            tx.tds_amount,
        )

    old_content_map: Dict[Tuple[str, str, str, Decimal, Decimal], List[NormalizedTransaction]] = defaultdict(list)
    for tx in unmatched_old:
        old_content_map[make_content_key(tx)].append(tx)

    new_content_map: Dict[Tuple[str, str, str, Decimal, Decimal], List[NormalizedTransaction]] = defaultdict(list)
    for tx in unmatched_new:
        new_content_map[make_content_key(tx)].append(tx)

    pass3_matched_old: Set[str] = set()
    pass3_matched_new: Set[str] = set()

    for c_key, old_candidates in old_content_map.items():
        if c_key in new_content_map:
            new_candidates = new_content_map[c_key]
            if len(old_candidates) == 1 and len(new_candidates) == 1:
                o_tx = old_candidates[0]
                n_tx = new_candidates[0]
                if are_ledgers_equivalent(o_tx.ledgers, n_tx.ledgers):
                    pass3_matched_old.add(o_tx.guid)
                    pass3_matched_new.add(n_tx.guid)
                    amt_diff = n_tx.total_amount - o_tx.total_amount
                    tds_diff = n_tx.tds_amount - o_tx.tds_amount
                    results.append(
                        TransactionComparison(
                            status=ClassificationStatus.RECREATED,
                            old_transaction=o_tx,
                            new_transaction=n_tx,
                            match_reason=f"Voucher number changed: '{o_tx.voucher_number}' -> '{n_tx.voucher_number}'; identical date, party, amounts and ledgers",
                            amount_difference=amt_diff,
                            tds_difference=tds_diff,
                        )
                    )

    matched_old_guids.update(pass3_matched_old)
    matched_new_guids.update(pass3_matched_new)
    unmatched_old = [tx for tx in unmatched_old if tx.guid not in matched_old_guids]
    unmatched_new = [tx for tx in unmatched_new if tx.guid not in matched_new_guids]

    # =========================================================================
    # PASS 4 — Detect Ambiguous Candidates (UNRESOLVED)
    # =========================================================================
    # An unmatched old transaction is ambiguous if multiple candidates in updated data
    # match party + date + amount or similar strong attributes.
    ambiguous_old: Dict[str, List[NormalizedTransaction]] = defaultdict(list)

    for o_tx in unmatched_old:
        candidates = [
            n_tx for n_tx in unmatched_new
            if n_tx.party_name.strip().lower() == o_tx.party_name.strip().lower()
            and n_tx.date == o_tx.date
            and n_tx.total_amount == o_tx.total_amount
        ]
        if len(candidates) > 1:
            ambiguous_old[o_tx.guid] = candidates

    pass4_old_guids: Set[str] = set()

    for o_guid, c_list in ambiguous_old.items():
        o_tx = old_by_guid[o_guid]
        pass4_old_guids.add(o_guid)

        cand_info = [f"GUID: {c.guid}, Voucher: {c.voucher_number}, Amount: ₹{c.total_amount}" for c in c_list]
        results.append(
            TransactionComparison(
                status=ClassificationStatus.UNRESOLVED,
                old_transaction=o_tx,
                new_transaction=None,
                match_reason=f"Ambiguous candidates ({len(c_list)} potential matches in updated snapshot)",
                candidates=[c.guid for c in c_list],
                unresolved_reason=f"Multiple candidate vouchers found for party '{o_tx.party_name}' on date {o_tx.date}: " + "; ".join(cand_info),
                amount_difference=-o_tx.total_amount,
                tds_difference=-o_tx.tds_amount,
            )
        )

    matched_old_guids.update(pass4_old_guids)
    unmatched_old = [tx for tx in unmatched_old if tx.guid not in matched_old_guids]

    # =========================================================================
    # PASS 5 — Remaining Old (MISSING) & Remaining New (NEW)
    # =========================================================================
    for o_tx in unmatched_old:
        results.append(
            TransactionComparison(
                status=ClassificationStatus.MISSING,
                old_transaction=o_tx,
                new_transaction=None,
                match_reason="Unmatched old transaction; missing in updated data",
                amount_difference=-o_tx.total_amount,
                tds_difference=-o_tx.tds_amount,
            )
        )

    for n_tx in unmatched_new:
        results.append(
            TransactionComparison(
                status=ClassificationStatus.NEW,
                old_transaction=None,
                new_transaction=n_tx,
                match_reason="Unmatched new transaction; newly added in updated snapshot",
                amount_difference=n_tx.total_amount,
                tds_difference=n_tx.tds_amount,
            )
        )

    return results
