import time
import uuid
from typing import Dict, Tuple
from fastapi import APIRouter, File, HTTPException, Response, UploadFile, status
from fastapi.responses import JSONResponse

from app.models.comparison_models import (
    ComparisonResponse,
    ComparisonSummarySchema,
    ReconciliationSchema,
    ReportInfoSchema,
)
from app.services.excel_report import generate_excel_report
from app.services.master_comparator import compare_masters
from app.services.normalizer import normalize_master, normalize_transaction
from app.services.parser import (
    ValidationError,
    validate_and_extract_masters,
    validate_and_extract_transactions,
)
from app.services.reconciliation import perform_reconciliation
from app.services.tds_calculator import (
    build_comparison_summary,
    calculate_party_summaries,
)
from app.services.transaction_matcher import match_and_compare_transactions

router = APIRouter(prefix="/api", tags=["comparison"])

# In-memory storage for reports with timestamp for automatic cleanup
REPORT_STORAGE: Dict[str, Tuple[bytes, float]] = {}
REPORT_TTL_SECONDS = 3600  # 1 hour


def _cleanup_old_reports():
    now = time.time()
    expired = [k for k, (_, ts) in REPORT_STORAGE.items() if now - ts > REPORT_TTL_SECONDS]
    for k in expired:
        REPORT_STORAGE.pop(k, None)


@router.post("/compare", response_model=ComparisonResponse)
async def compare_snapshots(
    master_old: UploadFile = File(..., description="Original Master JSON"),
    transactions_old: UploadFile = File(..., description="Original Transactions JSON"),
    master_new: UploadFile = File(..., description="Updated Master JSON"),
    transactions_new: UploadFile = File(..., description="Updated Transactions JSON"),
):
    _cleanup_old_reports()

    # 1. Read files
    try:
        content_m_old = await master_old.read()
        content_t_old = await transactions_old.read()
        content_m_new = await master_new.read()
        content_t_new = await transactions_new.read()
    except Exception as e:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail={"error": "File Read Error", "details": [f"Could not read uploaded files: {str(e)}"]}
        )

    # 2. Parse & Extract
    try:
        raw_m_old = validate_and_extract_masters(content_m_old)
        raw_t_old = validate_and_extract_transactions(content_t_old)
        raw_m_new = validate_and_extract_masters(content_m_new)
        raw_t_new = validate_and_extract_transactions(content_t_new)
    except ValidationError as e:
        return JSONResponse(
            status_code=status.HTTP_400_BAD_REQUEST,
            content={"error": "JSON Structure Validation Error", "details": e.details}
        )
    except Exception as e:
        return JSONResponse(
            status_code=status.HTTP_400_BAD_REQUEST,
            content={"error": "Unexpected JSON Error", "details": [str(e)]}
        )

    # 3. Normalize Records
    errors = []
    norm_m_old = []
    for idx, r in enumerate(raw_m_old):
        try:
            norm_m_old.append(normalize_master(r, idx))
        except ValidationError as e:
            errors.append(f"Original Master: {e.message}")

    norm_m_new = []
    for idx, r in enumerate(raw_m_new):
        try:
            norm_m_new.append(normalize_master(r, idx))
        except ValidationError as e:
            errors.append(f"Updated Master: {e.message}")

    norm_t_old = []
    for idx, r in enumerate(raw_t_old):
        try:
            norm_t_old.append(normalize_transaction(r, idx))
        except ValidationError as e:
            errors.append(f"Original Transactions: {e.message}")

    norm_t_new = []
    for idx, r in enumerate(raw_t_new):
        try:
            norm_t_new.append(normalize_transaction(r, idx))
        except ValidationError as e:
            errors.append(f"Updated Transactions: {e.message}")

    if errors:
        return JSONResponse(
            status_code=status.HTTP_400_BAD_REQUEST,
            content={"error": "Data Normalization Failed", "details": errors}
        )

    # 4. Compare Master & Transactions
    try:
        master_changes = compare_masters(norm_m_old, norm_m_new)
        comparisons = match_and_compare_transactions(norm_t_old, norm_t_new)
    except ValidationError as e:
        return JSONResponse(
            status_code=status.HTTP_400_BAD_REQUEST,
            content={"error": "Comparison Failed", "details": e.details}
        )
    except Exception as e:
        return JSONResponse(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            content={"error": "Comparison Engine Error", "details": [str(e)]}
        )

    # 5. TDS Calculations & Aggregations
    party_summaries = calculate_party_summaries(comparisons)
    summary_data = build_comparison_summary(comparisons, len(norm_t_old), len(norm_t_new))

    # 6. Reconciliation Checks
    reconciliation = perform_reconciliation(norm_t_old, norm_t_new, comparisons)

    if not reconciliation.is_valid:
        return JSONResponse(
            status_code=status.HTTP_422_UNPROCESSABLE_ENTITY,
            content={
                "error": "Comparison Validation Failed (Mathematical Reconciliation Error)",
                "details": reconciliation.errors,
                "reconciliation": {
                    "transaction_count": reconciliation.transaction_count,
                    "amount": reconciliation.amount,
                    "tds": reconciliation.tds,
                    "is_valid": False,
                    "errors": reconciliation.errors,
                }
            }
        )

    # 7. Generate Excel Report
    excel_bytes = generate_excel_report(
        summary=summary_data,
        reconciliation=reconciliation,
        party_summaries=party_summaries,
        comparisons=comparisons,
        master_changes=master_changes,
    )

    download_id = str(uuid.uuid4())
    REPORT_STORAGE[download_id] = (excel_bytes, time.time())

    return ComparisonResponse(
        success=True,
        summary=summary_data,
        reconciliation=ReconciliationSchema(
            transaction_count=reconciliation.transaction_count,
            amount=reconciliation.amount,
            tds=reconciliation.tds,
            is_valid=reconciliation.is_valid,
            errors=reconciliation.errors,
        ),
        report=ReportInfoSchema(
            filename="TDS_Change_Comparison_Report.xlsx",
            download_id=download_id,
        )
    )


@router.get("/download/{download_id}")
async def download_report(download_id: str):
    _cleanup_old_reports()
    item = REPORT_STORAGE.get(download_id)
    if not item:
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail="Report not found or has expired. Please run the comparison again."
        )

    excel_bytes, _ = item
    return Response(
        content=excel_bytes,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": 'attachment; filename="TDS_Change_Comparison_Report.xlsx"'},
    )


@router.get("/health")
async def health_check():
    return {"status": "ok", "service": "TDS Comparison Engine"}
