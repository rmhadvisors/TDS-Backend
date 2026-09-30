from typing import Dict, List, Set, Tuple
from app.models.comparison_models import MasterChange, MasterStatus, NormalizedMaster


def compare_masters(
    old_masters: List[NormalizedMaster],
    new_masters: List[NormalizedMaster],
) -> List[MasterChange]:
    """
    Compares original vs updated master records.
    Filters out technical metadata and identifies business-relevant changes.
    """
    changes: List[MasterChange] = []

    # Map by GUID or by Name fallback
    old_by_guid: Dict[str, NormalizedMaster] = {m.guid: m for m in old_masters}
    new_by_guid: Dict[str, NormalizedMaster] = {m.guid: m for m in new_masters}

    matched_old: Set[str] = set()
    matched_new: Set[str] = set()

    for guid, old_m in old_by_guid.items():
        if guid in new_by_guid:
            new_m = new_by_guid[guid]
            matched_old.add(guid)
            matched_new.add(guid)

            # Compare business fields
            # 1. Name
            if old_m.name.lower() != new_m.name.lower():
                changes.append(
                    MasterChange(
                        master_id=guid,
                        party_or_ledger=new_m.name,
                        field="Name",
                        old_value=old_m.name,
                        new_value=new_m.name,
                        change_type=MasterStatus.MASTER_MODIFIED,
                    )
                )

            # 2. Parent / Group
            if (old_m.parent_group or "").lower() != (new_m.parent_group or "").lower():
                changes.append(
                    MasterChange(
                        master_id=guid,
                        party_or_ledger=old_m.name,
                        field="Parent Group",
                        old_value=old_m.parent_group or "N/A",
                        new_value=new_m.parent_group or "N/A",
                        change_type=MasterStatus.MASTER_MODIFIED,
                    )
                )

            # 3. PAN
            if (old_m.pan or "").upper() != (new_m.pan or "").upper():
                changes.append(
                    MasterChange(
                        master_id=guid,
                        party_or_ledger=old_m.name,
                        field="PAN",
                        old_value=old_m.pan or "N/A",
                        new_value=new_m.pan or "N/A",
                        change_type=MasterStatus.MASTER_MODIFIED,
                    )
                )

            # 4. GSTIN
            if (old_m.gstin or "").upper() != (new_m.gstin or "").upper():
                changes.append(
                    MasterChange(
                        master_id=guid,
                        party_or_ledger=old_m.name,
                        field="GSTIN",
                        old_value=old_m.gstin or "N/A",
                        new_value=new_m.gstin or "N/A",
                        change_type=MasterStatus.MASTER_MODIFIED,
                    )
                )

            # 5. TDS Applicable
            if old_m.is_tds_applicable != new_m.is_tds_applicable:
                changes.append(
                    MasterChange(
                        master_id=guid,
                        party_or_ledger=old_m.name,
                        field="TDS Applicable",
                        old_value=str(old_m.is_tds_applicable) if old_m.is_tds_applicable is not None else "N/A",
                        new_value=str(new_m.is_tds_applicable) if new_m.is_tds_applicable is not None else "N/A",
                        change_type=MasterStatus.MASTER_MODIFIED,
                    )
                )

            # 6. TDS Section
            if (old_m.tds_section or "") != (new_m.tds_section or ""):
                changes.append(
                    MasterChange(
                        master_id=guid,
                        party_or_ledger=old_m.name,
                        field="TDS Section",
                        old_value=old_m.tds_section or "N/A",
                        new_value=new_m.tds_section or "N/A",
                        change_type=MasterStatus.MASTER_MODIFIED,
                    )
                )

            # 7. TDS Rate Name
            if (old_m.tds_rate_name or "") != (new_m.tds_rate_name or ""):
                changes.append(
                    MasterChange(
                        master_id=guid,
                        party_or_ledger=old_m.name,
                        field="TDS Rate Name",
                        old_value=old_m.tds_rate_name or "N/A",
                        new_value=new_m.tds_rate_name or "N/A",
                        change_type=MasterStatus.MASTER_MODIFIED,
                    )
                )

    # Missing Masters in updated
    for old_m in old_masters:
        if old_m.guid not in matched_old:
            changes.append(
                MasterChange(
                    master_id=old_m.guid,
                    party_or_ledger=old_m.name,
                    field="Record Status",
                    old_value="Present in Original Master",
                    new_value="Missing in Updated Master",
                    change_type=MasterStatus.MASTER_MISSING,
                )
            )

    # New Masters in updated
    for new_m in new_masters:
        if new_m.guid not in matched_new:
            changes.append(
                MasterChange(
                    master_id=new_m.guid,
                    party_or_ledger=new_m.name,
                    field="Record Status",
                    old_value="None",
                    new_value="Newly added in Updated Master",
                    change_type=MasterStatus.MASTER_NEW,
                )
            )

    return changes
