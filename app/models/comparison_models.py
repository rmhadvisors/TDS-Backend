from dataclasses import dataclass, field
from decimal import Decimal
from enum import Enum
from typing import Any, Dict, List, Optional
from pydantic import BaseModel, ConfigDict


class ClassificationStatus(str, Enum):
    UNCHANGED = "UNCHANGED"
    MODIFIED = "MODIFIED"
    NEW = "NEW"
    MISSING = "MISSING"  # Missing in Updated Data
    RECREATED = "RECREATED"
    UNRESOLVED = "UNRESOLVED"


class MasterStatus(str, Enum):
    MASTER_UNCHANGED = "MASTER_UNCHANGED"
    MASTER_MODIFIED = "MASTER_MODIFIED"
    MASTER_NEW = "MASTER_NEW"
    MASTER_MISSING = "MASTER_MISSING"


@dataclass
class NormalizedLedger:
    ledger_name: str
    amount: Decimal
    is_debit: bool
    is_credit: bool
    ledger_id: Optional[str] = None
    is_tds: bool = False
    tds_rate: Optional[Decimal] = None
    tds_section: Optional[str] = None

    def sort_key(self):
        return (
            self.ledger_name.strip().lower(),
            self.amount,
            1 if self.is_debit else 0,
            1 if self.is_credit else 0,
            1 if self.is_tds else 0
        )

    def is_equivalent(self, other: "NormalizedLedger") -> bool:
        return (
            self.ledger_name.strip().lower() == other.ledger_name.strip().lower()
            and self.amount == other.amount
            and self.is_debit == other.is_debit
            and self.is_credit == other.is_credit
        )


@dataclass
class NormalizedTransaction:
    guid: str
    voucher_number: str
    voucher_type: str
    date: str  # Normalized YYYY-MM-DD
    party_name: str
    party_id: Optional[str] = None
    narration: Optional[str] = None
    ledgers: List[NormalizedLedger] = field(default_factory=list)
    total_amount: Decimal = Decimal("0.00")
    tds_amount: Decimal = Decimal("0.00")
    tds_rate: Optional[Decimal] = None
    tds_section: Optional[str] = None
    raw_reference: Dict[str, Any] = field(default_factory=dict)
    is_optional: bool = False    # Tally "optional" (memo / provisional) voucher — not part of the books
    is_cancelled: bool = False


@dataclass
class NormalizedMaster:
    guid: str
    name: str
    parent_group: Optional[str] = None
    pan: Optional[str] = None
    gstin: Optional[str] = None
    is_tds_applicable: Optional[bool] = None
    tds_section: Optional[str] = None
    tds_rate_name: Optional[str] = None
    raw_fields: Dict[str, Any] = field(default_factory=dict)


@dataclass
class FieldDifference:
    field: str
    old_value: Any
    new_value: Any
    difference: Optional[Decimal] = None


@dataclass
class TransactionComparison:
    status: ClassificationStatus
    old_transaction: Optional[NormalizedTransaction] = None
    new_transaction: Optional[NormalizedTransaction] = None
    match_reason: str = ""
    field_changes: List[FieldDifference] = field(default_factory=list)
    amount_difference: Decimal = Decimal("0.00")
    tds_difference: Decimal = Decimal("0.00")
    candidates: Optional[List[str]] = None
    unresolved_reason: Optional[str] = None


@dataclass
class MasterChange:
    master_id: str
    party_or_ledger: str
    field: str
    old_value: Any
    new_value: Any
    change_type: MasterStatus


@dataclass
class PartySummary:
    party: str
    affected_entries: int
    missing_count: int
    new_count: int
    modified_count: int
    recreated_count: int
    original_amount: Decimal
    updated_amount: Decimal
    amount_difference: Decimal
    original_tds: Decimal
    updated_tds: Decimal
    tds_difference: Decimal


@dataclass
class ReconciliationResult:
    transaction_count: str  # "PASS" or "FAIL"
    amount: str             # "PASS" or "FAIL"
    tds: str                # "PASS" or "FAIL"
    is_valid: bool
    details: Dict[str, Any] = field(default_factory=dict)
    errors: List[str] = field(default_factory=list)


# Pydantic Schemas for API Serialization
class ComparisonSummarySchema(BaseModel):
    model_config = ConfigDict(extra="ignore")
    original_transactions: int
    updated_transactions: int
    unchanged: int
    modified: int
    new: int
    missing: int
    recreated: int
    unresolved: int
    original_amount: str
    updated_amount: str
    amount_difference: str
    original_tds: str
    updated_tds: str
    tds_difference: str


class ReconciliationSchema(BaseModel):
    model_config = ConfigDict(extra="ignore")
    transaction_count: str
    amount: str
    tds: str
    is_valid: bool
    errors: List[str] = []


class ReportInfoSchema(BaseModel):
    model_config = ConfigDict(extra="ignore")
    filename: str
    download_id: str
    party_name: Optional[str] = None
    period: Optional[str] = None
    saved_to_history: bool = True


class UnresolvedSectionSchema(BaseModel):
    """TDS Working rows whose section was only guessed: the ledgers / party a user keyword can be taken from."""
    model_config = ConfigDict(extra="ignore")
    names: List[str]
    guessed_section: str
    deductee: str
    rows: int
    tds: str
    vouchers: List[str] = []


class ComparisonResponse(BaseModel):
    model_config = ConfigDict(extra="ignore")
    success: bool
    summary: ComparisonSummarySchema
    reconciliation: ReconciliationSchema
    report: ReportInfoSchema
    unresolved_sections: List[UnresolvedSectionSchema] = []
    # Original-voucher parties missing from the latest master (renamed/deleted after filing); empty when the
    # original master was uploaded
    missing_master_ledgers: List[str] = []


class UserKeywordSchema(BaseModel):
    keyword: str
    section: str
