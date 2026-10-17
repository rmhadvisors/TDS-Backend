import logging
import re
import time
import uuid
from datetime import datetime, timezone
from decimal import Decimal
from typing import Dict, Optional, Tuple
from fastapi import APIRouter, File, Form, HTTPException, Query, Response, UploadFile, status
from fastapi.responses import JSONResponse, RedirectResponse

from app.models.comparison_models import (
    ComparisonResponse,
    ComparisonSummarySchema,
    ReconciliationSchema,
    ReportInfoSchema,
    UnresolvedSectionSchema,
    UserKeywordSchema,
)
from app.services.company_check import PARTIES_MISMATCH_ERROR, find_company_mismatch
from app.services.excel_report import derive_period_and_party, find_unresolved_sections, generate_excel_report
from app.services.master_comparator import compare_masters
from app.services.normalizer import (
    discover_tds_parties,
    is_tds_viable_transaction,
    normalize_master,
    normalize_transaction,
)
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
from app.services.user_keywords import add_user_keyword, list_user_keywords, remove_user_keyword, section_options
from app.services import report_store

logger = logging.getLogger(__name__)

router = APIRouter(prefix="/api", tags=["comparison"])

# Reports just generated stay in memory for an hour (fast download straight after a comparison, and still
# downloadable if saving to the history store failed); every report is also kept in the history store.
REPORT_STORAGE: Dict[str, Tuple[bytes, float, str]] = {}
REPORT_TTL_SECONDS = 3600  # 1 hour


def _cleanup_old_reports():
    now = time.time()
    expired = [k for k, (_, ts, _) in REPORT_STORAGE.items() if now - ts > REPORT_TTL_SECONDS]
    for k in expired:
        REPORT_STORAGE.pop(k, None)


@router.post("/compare", response_model=ComparisonResponse)
async def compare_snapshots(
    transactions_old: UploadFile = File(..., description="Original Transactions JSON"),
    master_new: UploadFile = File(..., description="Latest Master JSON (used for both snapshots unless master_old is given)"),
    transactions_new: UploadFile = File(..., description="Updated Transactions JSON"),
    party_name: Optional[str] = Form(None, description="Optional Party Name"),
    period: Optional[str] = Form(None, description="Optional Period / Date Range"),
    period_from: Optional[str] = Form(None, description="Optional report period start, YYYY-MM-DD (earliest voucher date if not given)"),
    period_to: Optional[str] = Form(None, description="Optional report period end, YYYY-MM-DD (latest voucher date if not given)"),
    tds_rows: Optional[str] = Form("bill", description="TDS Working rows: 'bill' (one row per bill, datewise as in the vouchers) or 'month' (per deductee per month, as in the 26Q return)"),
    master_old: Optional[UploadFile] = File(None, description="Optional Original Master JSON - only needed when ledgers were renamed or deleted after filing"),
):
    _cleanup_old_reports()

    # 1. Read files. The master rarely changes after filing, so normally only the latest one is uploaded and it
    # stands in for the original too - one large file less to upload and parse.
    try:
        content_m_old = await master_old.read() if master_old is not None else None
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
        raw_m_old = validate_and_extract_masters(content_m_old) if content_m_old is not None else []
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

    original_master_given = content_m_old is not None
    if not original_master_given:
        norm_m_old = norm_m_new

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

    # All four files must come from the same party (Tally company)
    mismatch = find_company_mismatch(
        ([("Original Master JSON", (m.guid for m in norm_m_old))] if original_master_given else []) + [
        ("Original Transaction JSON", (t.guid for t in norm_t_old)),
        ("Updated Master JSON", (m.guid for m in norm_m_new)),
        ("Updated Transaction JSON", (t.guid for t in norm_t_new)),
    ])
    if mismatch:
        return JSONResponse(
            status_code=status.HTTP_400_BAD_REQUEST,
            content={"error": PARTIES_MISMATCH_ERROR, "details": [mismatch]}
        )

    # 4. Filter to only TDS-viable transactions using automated two-pass discovery
    tds_parties = discover_tds_parties(norm_m_old, norm_m_new, norm_t_old, norm_t_new)
    # unfiltered vouchers of the snapshot the TDS Working is built from: TDS challans give the deposit dates
    all_t_working = norm_t_new or norm_t_old
    norm_t_old = [tx for tx in norm_t_old if is_tds_viable_transaction(tx, tds_parties)]
    norm_t_new = [tx for tx in norm_t_new if is_tds_viable_transaction(tx, tds_parties)]

    # Ledgers posted in the original TDS vouchers that the latest master no longer has were renamed or deleted after
    # filing; their details (group, PAN, section) are only in the original master, so the user is asked for it.
    # The posted ledger names are checked, not the voucher's party name: that is display text and may differ from the
    # ledger ("Ugamraj Singhvi" on a voucher posted to "Ugamraj Singhvi(Labour)"). Names match ignoring case, spaces
    # and punctuation, as vouchers can spell a ledger differently ("Input Cgst @2.5%" posted to "INPUT CGST@2.5%").
    missing_master_ledgers = []
    if not original_master_given:
        key = lambda name: re.sub(r"[^a-z0-9]", "", name.lower())
        master_names = {key(m.name) for m in norm_m_new if m.name}
        missing_master_ledgers = sorted({l.ledger_name.strip() for tx in norm_t_old for l in tx.ledgers
                                         if l.ledger_name and key(l.ledger_name) not in master_names})

    # 5. Compare Master & Transactions
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
    derived_party, derived_period = derive_period_and_party(
        norm_t_old=norm_t_old,
        norm_t_new=norm_t_new,
        norm_m_old=norm_m_old,
        norm_m_new=norm_m_new,
        party_name=party_name,
        period=period,
        period_from=period_from,
        period_to=period_to,
    )

    all_masters = (norm_m_new or []) + (norm_m_old or [])
    excel_bytes = generate_excel_report(
        summary=summary_data,
        reconciliation=reconciliation,
        party_summaries=party_summaries,
        comparisons=comparisons,
        master_changes=master_changes,
        party_name=derived_party,
        period=derived_period,
        masters=all_masters,
        tds_rows="month" if (tds_rows or "").strip().lower() == "month" else "bill",
        all_transactions=all_t_working,
    )

    download_id = str(uuid.uuid4())
    created = datetime.now(timezone.utc)
    filename = report_store.download_filename(derived_party, derived_period, created)
    REPORT_STORAGE[download_id] = (excel_bytes, time.time(), filename)
    try:
        report_store.get_store().save(
            report_store.make_object_name(download_id, derived_party, derived_period, created), excel_bytes
        )
        saved_to_history = True
    except Exception:
        logger.exception("Could not save report %s to the history store", download_id)
        saved_to_history = False
    unresolved = find_unresolved_sections(comparisons, all_masters, all_t_working)

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
            filename=filename,
            download_id=download_id,
            saved_to_history=saved_to_history,
            party_name=derived_party,
            period=derived_period,
        ),
        unresolved_sections=[UnresolvedSectionSchema(**u) for u in unresolved],
        missing_master_ledgers=missing_master_ledgers,
    )


@router.get("/download/{download_id}")
async def download_report(download_id: str):
    _cleanup_old_reports()
    item = REPORT_STORAGE.get(download_id)
    if item:
        excel_bytes, _, filename = item
    else:
        try:
            stored = report_store.find_report(download_id)
        except Exception:
            logger.exception("Could not look up report %s in the history store", download_id)
            raise HTTPException(status_code=status.HTTP_502_BAD_GATEWAY, detail="Report storage is unavailable.")
        if not stored:
            raise HTTPException(
                status_code=status.HTTP_404_NOT_FOUND,
                detail="Report not found. Please run the comparison again."
            )
        store = report_store.get_store()
        if store.kind == "s3":
            # The browser fetches the file straight from the private bucket with a link valid for a few minutes.
            return RedirectResponse(store.presigned_url(stored), status_code=status.HTTP_307_TEMPORARY_REDIRECT)
        excel_bytes, filename = store.read(stored), stored.filename

    return Response(
        content=excel_bytes,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'},
    )


@router.get("/history")
async def get_report_history(
    search: str = Query("", description="Filter by party name or period"),
    limit: int = Query(200, ge=1, le=1000),
):
    """Every report generated so far (shared by all users of the app), newest first."""
    try:
        reports = report_store.list_reports(search, limit)
    except Exception:
        logger.exception("Could not list the history store")
        raise HTTPException(status_code=status.HTTP_502_BAD_GATEWAY, detail="Report history is unavailable.")
    return {"storage": report_store.get_store().kind, "reports": [r.as_dict() for r in reports]}


@router.delete("/history/{report_id}")
async def delete_report(report_id: str):
    """Removes a report for everyone: from the history store (the S3 object itself) and the 1-hour memory cache."""
    try:
        stored = report_store.find_report(report_id)
        if stored:
            report_store.get_store().delete(stored)
    except Exception:
        logger.exception("Could not delete report %s from the history store", report_id)
        raise HTTPException(status_code=status.HTTP_502_BAD_GATEWAY, detail="Could not delete the report from storage.")
    cached = REPORT_STORAGE.pop(report_id, None)
    if not stored and not cached:
        raise HTTPException(status_code=status.HTTP_404_NOT_FOUND, detail="Report not found.")
    return {"deleted": report_id}


@router.get("/keywords")
async def get_user_keywords():
    """Keywords users added for natures of payment the statutory list does not know, and the sections to pick from."""
    return {"keywords": list_user_keywords(), "sections": section_options()}


@router.post("/keywords")
async def post_user_keyword(item: UserKeywordSchema):
    try:
        saved = add_user_keyword(item.keyword, item.section)
    except ValueError as e:
        return JSONResponse(status_code=status.HTTP_400_BAD_REQUEST, content={"error": "Invalid keyword", "details": [str(e)]})
    return {"saved": saved, "keywords": list_user_keywords()}


@router.delete("/keywords/{keyword}")
async def delete_user_keyword(keyword: str):
    if not remove_user_keyword(keyword):
        raise HTTPException(status_code=status.HTTP_404_NOT_FOUND, detail=f"Keyword '{keyword}' not found.")
    return {"keywords": list_user_keywords()}


@router.get("/health")
async def health_check():
    return {"status": "ok", "service": "TDS Comparison Engine"}
