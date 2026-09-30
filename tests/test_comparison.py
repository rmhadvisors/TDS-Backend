import pytest
from app.models.comparison_models import MasterStatus, NormalizedMaster
from app.services.master_comparator import compare_masters
from app.utils.normalization import is_technical_field


# Test 11: Master field changed -> MASTER_MODIFIED
def test_master_field_changed():
    m1 = NormalizedMaster(
        guid="MST-1",
        name="ABC Traders",
        parent_group="Sundry Creditors",
        pan="AAAAA1111A",
        gstin="27AAAAA1111A1Z1",
        is_tds_applicable=True,
        tds_section="194C",
    )
    m2 = NormalizedMaster(
        guid="MST-1",
        name="ABC Traders",
        parent_group="Sundry Creditors",
        pan="BBBBB2222B",  # Changed PAN
        gstin="27AAAAA1111A1Z1",
        is_tds_applicable=True,
        tds_section="194J",  # Changed Section
    )
    changes = compare_masters([m1], [m2])
    assert len(changes) == 2
    assert all(c.change_type == MasterStatus.MASTER_MODIFIED for c in changes)
    fields = [c.field for c in changes]
    assert "PAN" in fields
    assert "TDS Section" in fields


# Test 12: Technical alterid changed only -> MASTER_UNCHANGED at business level
def test_technical_alterid_ignored():
    assert is_technical_field("alterid") is True
    assert is_technical_field("updateddatetime") is True
    assert is_technical_field("pan") is False

    m1 = NormalizedMaster(
        guid="MST-1",
        name="ABC Traders",
        raw_fields={"name": "ABC Traders"}
    )
    m2 = NormalizedMaster(
        guid="MST-1",
        name="ABC Traders",
        raw_fields={"name": "ABC Traders"}
    )
    changes = compare_masters([m1], [m2])
    assert len(changes) == 0  # No business changes
