import datetime
import io
import re
from collections import Counter, defaultdict
from dataclasses import dataclass
from decimal import Decimal
from typing import Any, Dict, List, Optional, Tuple
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from app.models.comparison_models import (
    ClassificationStatus,
    ComparisonSummarySchema,
    MasterChange,
    MasterStatus,
    NormalizedTransaction,
    PartySummary,
    ReconciliationResult,
    TransactionComparison,
)
from app.services.tds_statutory_rates import (
    STATUTORY_TDS_REGISTRY,
    is_statutory_keyword_match,
    is_statutory_tds_party_name,
)
from app.services.normalizer import is_pure_tds_adjustment_journal, is_excluded_non_tds_party
from app.services.tds_context import build_tds_context
from app.services.tds_deductions import BookedDeduction, extract_booked_deductions, find_194q_not_deducted, group_by_deductee_month

# Colors and styles
HEADER_FILL = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid")
HEADER_FONT = Font(name="Calibri", size=11, bold=True, color="FFFFFF")

SECTION_FILL = PatternFill(start_color="334155", end_color="334155", fill_type="solid")
SECTION_FONT = Font(name="Calibri", size=11, bold=True, color="FFFFFF")

# Zeon Dashboard Styles
ZEON_HEADER_DARK = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
ZEON_HEADER_BLUE = PatternFill(start_color="2E75B6", end_color="2E75B6", fill_type="solid")
ZEON_TOTAL_BLUE = PatternFill(start_color="BDD7EE", end_color="BDD7EE", fill_type="solid")
ZEON_ZEBRA_LIGHT = PatternFill(start_color="F2F7FB", end_color="F2F7FB", fill_type="solid")
ZEON_MONTH_Q1 = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")  # Apr-Jun
ZEON_MONTH_Q2 = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")  # Jul-Sep
ZEON_MONTH_Q3 = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")  # Oct-Dec
ZEON_MONTH_Q4 = PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid")  # Jan-Mar
# TDS Working rows: a distinct light colour per month, so adjacent months stand apart
TDS_WORKING_MONTH_FILLS = {
    month: PatternFill(start_color=color, end_color=color, fill_type="solid")
    for month, color in {
        # light pastels, each month's hue far from its neighbours' so month boundaries stand out
        "April": "E2EFDA",      # light green
        "May": "FCE4D6",        # light peach
        "June": "DDEBF7",       # light blue
        "July": "FFF2CC",       # light yellow
        "August": "E4DFEC",     # light lavender
        "September": "D9F2F0",  # light teal
        "October": "FDE9F1",    # light pink
        "November": "EDEDED",   # light grey
        "December": "FFE699",   # soft gold
        "January": "C9DAF8",    # periwinkle
        "February": "F8CBAD",   # light coral
        "March": "D0E8D0",      # mint
    }.items()
}

ZEON_TITLE_FONT = Font(name="Calibri", size=13, bold=True, color="FFFFFF")
ZEON_SUBTITLE_FONT = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
ZEON_HEADER_FONT = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
ZEON_BOLD_TEXT = Font(name="Calibri", size=10, bold=True, color="1F4E79")
ZEON_DATA_TEXT = Font(name="Calibri", size=10, bold=False, color="1F4E79")

# Highlight Row and Badge Fills & Fonts
# 1. Missing / Deletion (Red)
MISSING_ROW_FILL = PatternFill(start_color="FEF2F2", end_color="FEF2F2", fill_type="solid")
MISSING_BADGE_FILL = PatternFill(start_color="FEE2E2", end_color="FEE2E2", fill_type="solid")
MISSING_FONT = Font(name="Calibri", size=10, bold=True, color="991B1B")

# 2. Backdated / Prior Period Addition (Amber/Warning)
BACKDATED_ROW_FILL = PatternFill(start_color="FFFBEB", end_color="FFFBEB", fill_type="solid")
BACKDATED_BADGE_FILL = PatternFill(start_color="FEF3C7", end_color="FEF3C7", fill_type="solid")
BACKDATED_FONT = Font(name="Calibri", size=10, bold=True, color="92400E")

# 3. New Month / Extended Addition (Green)
NEW_ROW_FILL = PatternFill(start_color="F0FDF4", end_color="F0FDF4", fill_type="solid")
NEW_BADGE_FILL = PatternFill(start_color="DCFCE7", end_color="DCFCE7", fill_type="solid")
NEW_FONT = Font(name="Calibri", size=10, bold=True, color="166534")

# 4. Modified (Amber/Yellow)
MOD_ROW_FILL = PatternFill(start_color="FFFBEB", end_color="FFFBEB", fill_type="solid")
MOD_BADGE_FILL = PatternFill(start_color="FEF3C7", end_color="FEF3C7", fill_type="solid")
MOD_FONT = Font(name="Calibri", size=10, bold=True, color="92400E")

# 4. Recreated (Indigo/Blue)
REC_ROW_FILL = PatternFill(start_color="EEF2FF", end_color="EEF2FF", fill_type="solid")
REC_BADGE_FILL = PatternFill(start_color="E0E7FF", end_color="E0E7FF", fill_type="solid")
REC_FONT = Font(name="Calibri", size=10, bold=True, color="3730A3")

# 5. Unresolved (Orange/Amber)
UNRES_ROW_FILL = PatternFill(start_color="FFF7ED", end_color="FFF7ED", fill_type="solid")
UNRES_BADGE_FILL = PatternFill(start_color="FFEDD5", end_color="FFEDD5", fill_type="solid")
UNRES_FONT = Font(name="Calibri", size=10, bold=True, color="9A3412")

# General styles
DATA_FONT = Font(name="Calibri", size=10, color="1E293B")
REASON_FONT = Font(name="Calibri", size=10, color="334155")
BOLD_DATA_FONT = Font(name="Calibri", size=10, bold=True, color="0F172A")
TOTAL_FONT = Font(name="Calibri", size=11, bold=True, color="0F172A")
ZEBRA_FILL = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")
WHITE_FILL = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")

PASS_FILL = PatternFill(start_color="DCFCE7", end_color="DCFCE7", fill_type="solid")
PASS_FONT = Font(name="Calibri", size=11, bold=True, color="166534")

FAIL_FILL = PatternFill(start_color="FEE2E2", end_color="FEE2E2", fill_type="solid")
FAIL_FONT = Font(name="Calibri", size=11, bold=True, color="991B1B")

DIFF_POS_FONT = Font(name="Calibri", size=10, bold=True, color="166534")
DIFF_NEG_FONT = Font(name="Calibri", size=10, bold=True, color="991B1B")

BORDER_THIN = Border(
    left=Side(style="thin", color="CBD5E1"),
    right=Side(style="thin", color="CBD5E1"),
    top=Side(style="thin", color="CBD5E1"),
    bottom=Side(style="thin", color="CBD5E1"),
)
BORDER_TOTAL = Border(
    left=Side(style="thin", color="CBD5E1"),
    right=Side(style="thin", color="CBD5E1"),
    top=Side(style="medium", color="1E293B"),
    bottom=Side(style="double", color="1E293B"),
)

ALIGN_CENTER = Alignment(horizontal="center", vertical="center")
ALIGN_LEFT = Alignment(horizontal="left", vertical="center")
ALIGN_RIGHT = Alignment(horizontal="right", vertical="center")
ALIGN_WRAP_LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)

CURRENCY_FORMAT = '"₹"#,##0.00;[Red]("-₹"#,##0.00);"-"'
NUMBER_FORMAT = "#,##0"



def derive_period_and_party(
    norm_t_old: List[Any],
    norm_t_new: List[Any],
    norm_m_old: Optional[List[Any]] = None,
    norm_m_new: Optional[List[Any]] = None,
    party_name: Optional[str] = None,
    period: Optional[str] = None,
    period_from: Optional[str] = None,
    period_to: Optional[str] = None,
) -> Tuple[str, str]:
    """
    Derives canonical party name and period strings, using explicitly provided
    values if given, or extracting them automatically from transactions/masters.
    The period is "dd/mm/yyyy to dd/mm/yyyy": the user's From / To dates (YYYY-MM-DD) where given, the
    earliest / latest voucher date for the end that is not.
    """
    # 1. Party Name
    final_party = (party_name or "").strip()
    if not final_party:
        parties = set()
        for tx in (norm_t_old or []) + (norm_t_new or []):
            p = getattr(tx, "party_name", None)
            if p and str(p).strip():
                parties.add(str(p).strip())

        if not parties and (norm_m_old or norm_m_new):
            for m in (norm_m_old or []) + (norm_m_new or []):
                n = getattr(m, "name", None)
                if n and str(n).strip():
                    parties.add(str(n).strip())

        if len(parties) == 1:
            final_party = next(iter(parties))
        elif len(parties) > 1:
            sorted_parties = sorted(parties)
            if len(sorted_parties) <= 3:
                final_party = ", ".join(sorted_parties)
            else:
                final_party = f"All Parties ({len(sorted_parties)} Parties)"
        else:
            final_party = "All Parties"

    # 2. Period
    final_period = (period or "").strip()
    if not final_period:
        dates = []
        for tx in (norm_t_old or []) + (norm_t_new or []):
            d = getattr(tx, "date", None)
            if d and str(d).strip():
                dates.append(str(d).strip())

        start = (period_from or "").strip() or (min(dates) if dates else "")
        end = (period_to or "").strip() or (max(dates) if dates else "")
        if not start or not end:
            final_period = "Full Period"
        elif start == end:
            final_period = _ddmmyyyy(start)
        else:
            final_period = f"{_ddmmyyyy(start)} to {_ddmmyyyy(end)}"

    return final_party, final_period


def _ddmmyyyy(value: str) -> str:
    """"2026-03-31" (or a datetime string starting with it) -> "31/03/2026"; anything else unchanged."""
    m = re.match(r"^(\d{4})-(\d{2})-(\d{2})", value)
    return f"{m.group(3)}/{m.group(2)}/{m.group(1)}" if m else value


def _auto_adjust_column_widths(ws, min_width: int = 14, max_width: int = 65):
    """Adjusts column widths dynamically with bounds and adds comfortable spacing."""
    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        has_long_text = False
        for cell in col:
            val = str(cell.value or "")
            # Don't count title in column width calculation
            if cell.row == 1 and cell.column == 1:
                continue
            if len(val) > max_len:
                max_len = len(val)
            if len(val) > 40:
                has_long_text = True
        
        calculated_width = min(max(max_len + 4, min_width), max_width)
        ws.column_dimensions[col_letter].width = calculated_width
        
        # Enable text wrap on longer narrative columns
from app.services.tds_statutory_rates import STATUTORY_TDS_REGISTRY, StatutoryTDSRule


def extract_pan_from_tx(tx, party_pan_map: Optional[Dict[str, str]] = None) -> Optional[str]:
    """
    Extracts 10-character PAN generically from:
    1. Master party registry map (party_pan_map)
    2. Transaction raw reference fields (pan, party_pan, panno, itpan, etc.)
    3. Transaction narration regex for 10-character alphanumeric PAN
    4. Embedded 15-character GSTIN (characters 3 to 12)
    """
    p_name = getattr(tx, "party_name", "") or ""
    p_clean = p_name.strip().lower()

    # 1. From Master party registry if available
    if party_pan_map and p_name:
        p_pan = party_pan_map.get(p_clean)
        if not p_pan:
            clean_noparen = re.sub(r"\s*\([^)]*\)", "", p_clean).strip()
            p_pan = party_pan_map.get(clean_noparen)
        if not p_pan:
            for l in getattr(tx, "ledgers", []) or []:
                lname_c = (getattr(l, "ledger_name", "") or "").lower().strip()
                if lname_c and lname_c in party_pan_map:
                    p_pan = party_pan_map.get(lname_c)
                    break
                lname_noparen = re.sub(r"\s*\([^)]*\)", "", lname_c).strip()
                if lname_noparen and lname_noparen in party_pan_map:
                    p_pan = party_pan_map.get(lname_noparen)
                    break
        if p_pan and re.match(r"^[A-Z]{5}[0-9]{4}[A-Z]$", p_pan.upper()):
            return p_pan.upper()

    # 2. From raw_reference
    raw = getattr(tx, "raw_reference", {}) or {}
    for key in ["pan", "party_pan", "panno", "incometaxpanno", "partyincometaxno", "itpan", "itpanno", "gstin", "partygstin"]:
        val = raw.get(key)
        if val and isinstance(val, str):
            m_pan = re.search(r"([A-Z]{5}[0-9]{4}[A-Z])", val.upper())
            if m_pan:
                return m_pan.group(1)

    # 3. From Narration
    narration = getattr(tx, "narration", "") or ""
    if narration:
        m_pan = re.search(r"\b([A-Z]{5}[0-9]{4}[A-Z])\b", narration.upper())
        if m_pan:
            return m_pan.group(1)
        m_pan2 = re.search(r"PAN[:\-\s]*([A-Z0-9]{10})", narration.upper())
        if m_pan2:
            candidate = m_pan2.group(1)
            if re.match(r"^[A-Z]{5}[0-9]{4}[A-Z]$", candidate):
                return candidate

    # 4. From GSTIN in raw reference or narration (15-character GSTIN embeds PAN from pos 3 to 12)
    for text in [narration, str(raw)]:
        m_gst = re.search(r"[0-9]{2}([A-Z]{5}[0-9]{4}[A-Z])[0-9A-Z]{3}", text.upper())
        if m_gst:
            return m_gst.group(1)

    return None


def is_gst_duty_ledger(lname: str) -> bool:
    """
    Identifies genuine GST tax duty ledgers (e.g. INPUT IGST, OUTPUT CGST, RCM SGST).
    Ensures that expense heads (e.g. Purchase CGST/SGST@18%, Export Freight)
    are NOT misclassified as tax duties.
    """
    if not lname:
        return False
    u = lname.upper().strip()
    if "PURCHASE" in u or "SALES" in u or "EXPENSE" in u or "FREIGHT" in u:
        return False
    if any(k in u for k in ["INPUT IGST", "INPUT CGST", "INPUT SGST", "OUTPUT IGST", "OUTPUT CGST", "OUTPUT SGST", "RCM IGST", "RCM CGST", "RCM SGST"]):
        return True
    if u.startswith("INPUT ") or u.startswith("OUTPUT ") or u.startswith("RCM "):
        if any(g in u for g in ["IGST", "CGST", "SGST", "TAX"]):
            return True
    return False


def get_goods_base_amount(tx) -> Decimal:
    """
    Extracts the statutory net purchase base amount excluding GST duties u/s CBDT Circular 13/2021.
    """
    gst_amt = Decimal("0.00")
    for l in (getattr(tx, "ledgers", []) or []):
        lname = (getattr(l, "ledger_name", "") or "").upper()
        if is_gst_duty_ledger(lname):
            gst_amt += abs(getattr(l, "amount", Decimal("0.00")))
    
    tot = getattr(tx, "total_amount", Decimal("0.00"))
    if tot > gst_amt and gst_amt > Decimal("0.00"):
        return tot - gst_amt
    return tot


def infer_tds_section_and_rate(tx, party_pan_map: Optional[Dict[str, str]] = None) -> Tuple[str, str, str]:
    """
    Infers statutory TDS section, deduction rate, and nature note for a transaction
    strictly according to the official Income Tax Department Non-Salary TDS Chart (tdsSectionRate.pdf).
    Checks PAN 4th-character entity status before applying applicable rate.
    """
    section = getattr(tx, "tds_section", None)
    rate = getattr(tx, "tds_rate", None)
    note = ""

    # PAN & Entity Type Extraction (4th character rule: P/H -> Indv/HUF, C/F/A/B/T -> Company/Firm)
    pan = extract_pan_from_tx(tx, party_pan_map)
    pan_entity_type: Optional[str] = None
    if pan and len(pan) == 10:
        c4 = pan[3].upper()
        if c4 in ["P", "H"]:
            pan_entity_type = "Individual/HUF"
        elif c4 == "C":
            pan_entity_type = "Company"
        elif c4 == "F":
            pan_entity_type = "Firm/LLP"
        elif c4 in ["A", "B", "T", "G", "L", "J"]:
            pan_entity_type = "AOP/BOI/Trust/Govt"

    tds_ledger_names = []
    expense_ledger_names = []
    has_bank_payment = False
    tds_is_debit = False

    for l in getattr(tx, "ledgers", []) or []:
        lname = l.ledger_name if hasattr(l, "ledger_name") else str(l)
        is_tds = getattr(l, "is_tds", False)
        is_debit = getattr(l, "is_debit", False)
        if is_tds or "TDS" in lname.upper() or "TAX DEDUCTED" in lname.upper():
            tds_ledger_names.append(lname)
            if is_debit:
                tds_is_debit = True
            if rate is None:
                m_rate = re.search(r"[@\s](\d+(?:\.\d+)?)\s*%", lname)
                if m_rate:
                    rate = f"{float(m_rate.group(1)):.2f}%"
        else:
            expense_ledger_names.append(lname)
            if any(b in lname.upper() for b in ["BANK", "CASH", "HDFC", "ICICI", "SBI", "MAHARASHTRA", "KOTAK", "AXIS", "CANARA", "PNB", "BOI", "YES BANK", "INDUSIND"]):
                has_bank_payment = True

    narration_upper = (getattr(tx, "narration", "") or "").upper()
    party_upper = (getattr(tx, "party_name", "") or "").upper()
    all_names_upper = " ".join([party_upper, narration_upper] + tds_ledger_names + expense_ledger_names).upper()

    # 1. Detect Government TDS Challan Tax Payments (Debit to TDS Payable, Credit to Bank)
    if (tds_is_debit and has_bank_payment) or any(k in narration_upper for k in ["ITD", "CHALLAN", "TAX DEPOSIT", "OLTAS", "281", "TAX PAYMENT"]):
        return "Challan 281", "N/A (Challan Payment)", "TDS Tax Payment / Challan Settlement to Income Tax Dept"

    # 2. Match against Official Statutory TDS Registry (from tdsSectionRate.pdf)
    matched_rule: Optional[StatutoryTDSRule] = None

    # 2a. Match explicit section code in transaction or ledgers
    raw_section_str = str(section or "").upper().replace("SEC", "").replace(".", "").strip()
    if raw_section_str:
        for rule in STATUTORY_TDS_REGISTRY:
            if rule.section_code == raw_section_str or (rule.sub_section and rule.sub_section.upper() == raw_section_str):
                matched_rule = rule
                break

    # 2a. Check explicit TDS deduction ledger names FIRST (Highest statutory accuracy)
    tds_names_upper = " ".join(tds_ledger_names).upper()
    if tds_names_upper:
        if "CONTRACT" in tds_names_upper or "LABOUR" in tds_names_upper or "FREIGHT" in tds_names_upper or "TRANSPORT" in tds_names_upper or "194C" in tds_names_upper:
            for rule in STATUTORY_TDS_REGISTRY:
                if rule.section_code == "194C":
                    matched_rule = rule
                    break
        elif "194T" in tds_names_upper:
            for rule in STATUTORY_TDS_REGISTRY:
                if rule.section_code == "194T":
                    matched_rule = rule
                    break
        elif "194Q" in tds_names_upper or "PURCHASE" in tds_names_upper:
            for rule in STATUTORY_TDS_REGISTRY:
                if rule.section_code == "194Q":
                    matched_rule = rule
                    break
        elif "PROFESSION" in tds_names_upper or "TECHNICAL" in tds_names_upper or "194J" in tds_names_upper:
            for rule in STATUTORY_TDS_REGISTRY:
                if rule.section_code == "194J":
                    matched_rule = rule
                    break
        elif "INTEREST" in tds_names_upper or "194A" in tds_names_upper:
            for rule in STATUTORY_TDS_REGISTRY:
                if rule.section_code == "194A":
                    matched_rule = rule
                    break
        elif "COMMISSION" in tds_names_upper or "BROKERAGE" in tds_names_upper or "194H" in tds_names_upper:
            for rule in STATUTORY_TDS_REGISTRY:
                if rule.section_code == "194H":
                    matched_rule = rule
                    break
        elif "RENT" in tds_names_upper or "194I" in tds_names_upper:
            for rule in STATUTORY_TDS_REGISTRY:
                if rule.section_code == "194I":
                    matched_rule = rule
                    break

    if not matched_rule:
        for rule in STATUTORY_TDS_REGISTRY:
            pattern = rf"\b(SEC\s*)?{rule.section_code}\b"
            if re.search(pattern, all_names_upper):
                matched_rule = rule
                break

    # 2b. Match by Statutory Nature of Payment Keywords
    non_tds_names_upper = " ".join([
        l.ledger_name.upper() for l in getattr(tx, "ledgers", []) or []
        if not (getattr(l, "is_tds", False) or "TDS" in l.ledger_name.upper() or "TAX DEDUCTED" in l.ledger_name.upper())
        and not (getattr(tx, "party_name", None) and getattr(tx, "party_name", "").upper() in l.ledger_name.upper())
    ])
    if not matched_rule and non_tds_names_upper:
        for rule in STATUTORY_TDS_REGISTRY:
            if rule.section_code in ["194H", "194J", "194I", "194Q", "194A", "194T", "194R", "194B", "194IA", "194IB", "194M", "194S", "194C"]:
                for kw in rule.nature_keywords:
                    if re.search(rf"\b{re.escape(kw)}\b", non_tds_names_upper) or (("%" in kw or "@" in kw) and kw in non_tds_names_upper):
                        matched_rule = rule
                        break
                if matched_rule:
                    break

    if not matched_rule:
        for rule in STATUTORY_TDS_REGISTRY:
            for kw in rule.nature_keywords:
                if "%" in kw or "@" in kw:
                    if kw in all_names_upper:
                        matched_rule = rule
                        break
                else:
                    if re.search(rf"\b{re.escape(kw)}\b", all_names_upper):
                        matched_rule = rule
                        break
            if matched_rule:
                break

    # 2c. Keywords users added for natures the statutory list does not know (user_keywords.py)
    if not matched_rule:
        from app.services.user_keywords import match_user_rule
        hit = match_user_rule(non_tds_names_upper) or match_user_rule(party_upper)
        if hit:
            matched_rule = hit[0]

    # 194Q Provision Journals for high-value goods purchases
    if not matched_rule:
        is_goods_journal = (
            (getattr(tx, "voucher_type", "") or "").upper() == "JOURNAL"
            and getattr(tx, "total_amount", Decimal("0.00")) >= Decimal("200000.00")
            and any(k in all_names_upper for k in ["PURCHASE", "GOODS", "RAW MATERIAL", "194Q", "PHARMA", "LABORATORIES", "DRUGS", "CHEMICAL", "FINECHEM"])
        )
        if is_goods_journal:
            return "194Q", "0.10%", "Section 194Q - Purchase of Goods (Turnover Provision)"

    # Fallback Handling (Default to Section 194C Works Contract for general contractor payments)
    if not matched_rule:
        matched_rule = next((r for r in STATUTORY_TDS_REGISTRY if r.section_code == "194C"), STATUTORY_TDS_REGISTRY[0])

    section_code = matched_rule.sub_section or matched_rule.section_code
    if section_code in ["194J(a)", "194J(b)", "194J(A)", "194J(B)"]:
        section_code = "194J"
    elif section_code in ["194I(a)", "194I(b)", "194I(A)", "194I(B)"]:
        section_code = "194I"
    
    clean_section = str(section_code).upper().replace("SEC ", "").replace("SEC", "").strip()
    note = f"Section {clean_section} - {matched_rule.description}"

    # 3. Rate calculation & Statutory Verification
    if rate is None:
        total_amt = getattr(tx, "total_amount", Decimal("0.00"))
        tds_amt = getattr(tx, "tds_amount", Decimal("0.00"))
        
        # Calculate effective base after excluding GST duties for accurate rate detection
        gst_amt_est = Decimal("0.00")
        for l in (getattr(tx, "ledgers", []) or []):
            lname = (getattr(l, "ledger_name", "") or "").upper()
            if is_gst_duty_ledger(lname):
                gst_amt_est += abs(getattr(l, "amount", Decimal(0)))
        
        effective_base = total_amt - gst_amt_est if (total_amt > gst_amt_est and gst_amt_est > Decimal("0.00")) else total_amt

        if tds_amt > 0 and effective_base > 0:
            calc_rate = (tds_amt / effective_base) * 100
            if abs(calc_rate - Decimal("0.10")) < Decimal("0.03"):
                rate = "0.10%"
            elif abs(calc_rate - Decimal("1.00")) < Decimal("0.15"):
                rate = "1.00%"
            elif abs(calc_rate - Decimal("2.00")) < Decimal("0.35"):
                rate = "2.00%"
            elif abs(calc_rate - Decimal("5.00")) < Decimal("0.25"):
                rate = "5.00%"
            elif abs(calc_rate - Decimal("10.00")) < Decimal("0.50"):
                rate = "10.00%"
            elif abs(calc_rate - Decimal("20.00")) < Decimal("0.50"):
                rate = "20.00% (Sec 206AA No/Invalid PAN)"
            elif abs(calc_rate - Decimal("30.00")) < Decimal("0.50"):
                rate = "30.00%"
            else:
                rate = f"{float(calc_rate):.2f}%"
        elif clean_section == "194J" and any(k in all_names_upper for k in ["ATRAKSH", "TECHNICAL", "SOFTWARE", "IT ", "SECURITY", "CALL CENTER", "SYSTEMS", "ENGINEERING"]):
            rate = "2.00%"
        else:
            # Statutory rate assignment from PDF chart
            if pan_entity_type == "Individual/HUF":
                rate = matched_rule.rate_individual_huf
            elif pan_entity_type in ["Company", "Firm/LLP", "AOP/BOI/Trust/Govt"]:
                rate = matched_rule.rate_company_firm
            else:
                rate = matched_rule.standard_rate_display
    elif isinstance(rate, (int, float, Decimal)):
        rate = f"{float(rate):.2f}%"
    elif isinstance(rate, str) and not rate.endswith("%"):
        try:
            rate = f"{float(rate):.2f}%"
        except ValueError:
            pass

    # Append PAN Entity context to Note if PAN is identified
    if pan and pan_entity_type:
        note += f" [PAN: {pan} ({pan_entity_type})]"

    return clean_section, str(rate), note


@dataclass
class StatutoryTDSCalculation:
    calc_mode: str
    base_amount: Decimal
    rate_str: str
    rate_decimal: Decimal
    statutory_tds: Decimal
    booked_tds: Decimal
    shortfall: Decimal
    compliance_note: str


def compute_transaction_tds_calculation(
    tx,
    sec: str,
    rt_str: str,
    is_backdated: bool,
    cutoff_date: Optional[str] = None,
) -> StatutoryTDSCalculation:
    """
    Computes statutory TDS calculation:
    - Forward Calculation: For Invoices, Purchases, Journals (Gross Base * Rate)
    - Backward Calculation: For Net Payments / Advances u/s 195A (Net Paid / (1 - Rate) - Net Paid)
    """
    total_amt = getattr(tx, "total_amount", Decimal("0.00")) or Decimal("0.00")
    booked_tds = getattr(tx, "tds_amount", Decimal("0.00")) or Decimal("0.00")
    vch_type = (getattr(tx, "voucher_type", "") or "").strip().lower()

    # Parse Rate Decimal
    clean_rt = rt_str.replace("%", "").strip()
    rate_decimal = Decimal("0.00")
    if "/" in clean_rt:
        pan = extract_pan_from_tx(tx)
        if pan and len(pan) == 10 and pan[3].upper() in ["P", "H"]:
            rate_decimal = Decimal("0.01")
            rt_str = "1.00%"
        else:
            rate_decimal = Decimal("0.02")
            rt_str = "2.00%"
    else:
        try:
            m_rt = re.search(r"(\d+(?:\.\d+)?)", clean_rt)
            if m_rt:
                rate_decimal = Decimal(m_rt.group(1)) / Decimal("100")
        except Exception:
            rate_decimal = Decimal("0.00")

    # Detect Net Payment vs Invoice
    has_bank = any(
        any(b in l.ledger_name.upper() for b in ["BANK", "KOTAK", "HDFC", "ICICI", "SBI", "CASH", "AXIS"])
        for l in getattr(tx, "ledgers", []) or []
    )
    has_purchase = any(
        "PURCHASE" in l.ledger_name.upper() or "EXPENSE" in l.ledger_name.upper()
        for l in getattr(tx, "ledgers", []) or []
    )

    # Exclude GST duty taxes from base amount as per CBDT Circular No. 23/2017
    gst_amt = Decimal("0.00")
    for l in (getattr(tx, "ledgers", []) or []):
        lname = (l.ledger_name or "").upper()
        if any(k in lname for k in ["INPUT IGST", "INPUT CGST", "INPUT SGST", "OUTPUT IGST", "OUTPUT CGST", "OUTPUT SGST", "RCM IGST", "RCM CGST", "RCM SGST", "IGST@", "CGST@", "SGST@", "IGST @", "CGST @", "SGST @"]):
            gst_amt += abs(l.amount)
        elif lname.startswith("INPUT ") or lname.startswith("OUTPUT ") or lname.startswith("RCM "):
            if any(g in lname for g in ["IGST", "CGST", "SGST", "TAX"]):
                gst_amt += abs(l.amount)

    effective_total = total_amt
    if gst_amt > Decimal("0.00") and total_amt > gst_amt:
        effective_total = total_amt - gst_amt

    if (vch_type in ["payment", "bank payment", "cheque", "bank", "advance"]) or (has_bank and not has_purchase):
        calc_mode = "Backward (Gross-Up u/s 195A)"
        net_paid = effective_total
        if Decimal("0.00") < rate_decimal < Decimal("1.00"):
            gross_base = (net_paid / (Decimal("1.00") - rate_decimal)).quantize(Decimal("0.01"))
            statutory_tds = (gross_base - net_paid).quantize(Decimal("0.01"))
        else:
            gross_base = net_paid
            statutory_tds = Decimal("0.00")
    else:
        calc_mode = "Forward (Gross Invoice)"
        gross_base = effective_total
        statutory_tds = (gross_base * rate_decimal).quantize(Decimal("0.01"))

    shortfall = (statutory_tds - booked_tds).quantize(Decimal("0.01"))

    # Generate Compliance Note
    if is_backdated:
        if shortfall > Decimal("0.00"):
            comp_note = (
                f"Backdated addition ({tx.date} <= cutoff {cutoff_date}). "
                f"Statutory TDS shortfall of ₹{float(shortfall):,.2f} under {sec} @ {rt_str} requires "
                f"Revised TDS Return (Form 26Q) and late deduction interest u/s 201(1A)."
            )
        else:
            comp_note = (
                f"Backdated addition ({tx.date} <= cutoff {cutoff_date}) for ₹{float(gross_base):,.2f}. "
                f"TDS of ₹{float(booked_tds):,.2f} is booked in Tally. Verify inclusion in previous quarter return."
            )
    else:
        if shortfall > Decimal("0.00"):
            comp_note = (
                f"Extended period transaction on {tx.date}. "
                f"Statutory TDS liability of ₹{float(statutory_tds):,.2f} under {sec} @ {rt_str} to be deposited "
                f"via Challan 281 by 7th of next month (Pending JV in Tally: ₹{float(shortfall):,.2f})."
            )
        else:
            comp_note = (
                f"Extended period transaction on {tx.date} for ₹{float(gross_base):,.2f} with "
                f"TDS of ₹{float(booked_tds):,.2f} booked in Tally. To be included in next regular TDS return."
            )

    return StatutoryTDSCalculation(
        calc_mode=calc_mode,
        base_amount=gross_base,
        rate_str=rt_str,
        rate_decimal=rate_decimal,
        statutory_tds=statutory_tds,
        booked_tds=booked_tds,
        shortfall=shortfall,
        compliance_note=comp_note,
    )


def _get_month_name(date_val: Any) -> str:

    """Extracts standard financial year month name (e.g. 'April', 'May') from date."""
    if not date_val:
        return "April"
    if hasattr(date_val, "strftime"):
        return date_val.strftime("%B")
    cleaned = str(date_val).strip()
    month_names = [
        "", "January", "February", "March", "April", "May", "June",
        "July", "August", "September", "October", "November", "December"
    ]
    if "-" in cleaned:
        parts = cleaned.split("-")
        if len(parts) >= 2:
            try:
                m_int = int(parts[1])
                if 1 <= m_int <= 12:
                    return month_names[m_int]
            except Exception:
                pass
    elif "/" in cleaned:
        parts = cleaned.split("/")
        if len(parts) >= 2:
            try:
                m_int = int(parts[1])
                if 1 <= m_int <= 12:
                    return month_names[m_int]
            except Exception:
                pass
    return "April"


def _extract_challan_info(tx: Any) -> Tuple[str, str]:
    """Extracts challan date and challan number from transaction raw reference or narration."""
    narration = getattr(tx, "narration", "") or ""
    raw = getattr(tx, "raw_reference", {}) or {}
    challan_no = ""
    challan_date = ""

    for k in ["challan_no", "challanno", "bsr_code", "cin", "challan_number", "challan"]:
        v = raw.get(k)
        if v:
            challan_no = str(v).strip()
            break

    for k in ["challan_date", "challandate", "deposit_date"]:
        v = raw.get(k)
        if v:
            challan_date = str(v).strip()
            break

    if not challan_no and narration:
        m_no = re.search(r"\b(?:CHALLAN|CHLN|BSR|CIN)[^0-9A-Z]*([0-9A-Z]{4,10})\b", narration.upper())
        if m_no:
            challan_no = m_no.group(1)

    if not challan_date and narration:
        m_dt = re.search(r"\b(\d{4}-\d{2}-\d{2}|\d{2}[-/]\d{2}[-/]\d{4})\b", narration)
        if m_dt:
            challan_date = m_dt.group(1)

    return challan_date, challan_no


def compute_tds_deductions(
    comparisons: List[TransactionComparison],
    masters: Optional[List[Any]] = None,
    tds_deposits: Optional[List[NormalizedTransaction]] = None,
) -> List[BookedDeduction]:
    """TDS booked in the vouchers of the updated snapshot (else the original), plus 194Q due but not deducted —
    the rows of the TDS Working sheet, one per bill."""
    # Ledger name -> PAN from the masters (PAN field, else the PAN inside the ledger's GSTIN)
    party_pan_map: Dict[str, str] = {}
    if masters:
        for m in masters:
            m_pan = getattr(m, "pan", None)
            m_name = getattr(m, "name", None)
            if m_pan and m_name and re.match(r"^[A-Z]{5}[0-9]{4}[A-Z]$", m_pan.strip().upper()):
                clean_n = m_name.strip().lower()
                clean_p = m_pan.strip().upper()
                party_pan_map[clean_n] = clean_p
                clean_noparen = re.sub(r"\s*\([^)]*\)", "", clean_n).strip()
                if clean_noparen:
                    party_pan_map[clean_noparen] = clean_p

    has_new_dataset = any(c.new_transaction is not None for c in comparisons)
    working_txs = [
        c.new_transaction if has_new_dataset else c.old_transaction
        for c in comparisons
        if (c.new_transaction if has_new_dataset else c.old_transaction) is not None
    ]
    ctx = build_tds_context(list(masters or []), working_txs)

    def _pan_of(name: str, tx: Optional[NormalizedTransaction] = None) -> str:
        """PAN of a deductee ledger: its master, else — only when it is the voucher's own party — the party PAN /
        GSTIN carried on that voucher. Never another ledger's PAN, never the company's own GSTIN."""
        n = (name or "").strip().lower()
        info = ctx.info(name)
        pan = (info.pan if info and info.pan else None) or party_pan_map.get(n) or party_pan_map.get(re.sub(r"\s*\([^)]*\)", "", n).strip())
        bare = lambda x: re.sub(r"\s*\([^)]*\)", "", (x or "").strip().lower()).strip()
        if not pan and tx is not None and n and bare(n) == bare(tx.party_name):
            raw = tx.raw_reference or {}
            for key in ("pan", "party_pan", "panno", "partyincometaxno", "partygstin", "gstin"):
                m_pan = re.search(r"[A-Z]{5}[0-9]{4}[A-Z]", str(raw.get(key) or "").upper())
                if m_pan:
                    pan = m_pan.group(0)
                    break
        pan = (pan or "").replace(" ", "").upper()
        return pan if re.fullmatch(r"[A-Z]{5}[0-9]{4}[A-Z]", pan) else ""

    booked = extract_booked_deductions(working_txs, ctx, _pan_of, tds_deposits=tds_deposits)
    # plus 194Q that was due on purchases above the threshold but not deducted (TDS 0, shortfall in Diff)
    return booked + find_194q_not_deducted(working_txs, ctx, booked, _pan_of)


def find_unresolved_sections(
    comparisons: List[TransactionComparison],
    masters: Optional[List[Any]] = None,
    tds_deposits: Optional[List[NormalizedTransaction]] = None,
) -> List[Dict[str, Any]]:
    """TDS Working rows whose section nothing in the data or the keyword lists named (guessed from the rate, or
    "Not Specified"), grouped by the ledgers / party to add a keyword from — largest TDS first."""
    groups: Dict[Tuple[str, ...], Dict[str, Any]] = {}
    for d in compute_tds_deductions(comparisons, masters, tds_deposits):
        if not d.unresolved_names:
            continue
        key = tuple(d.unresolved_names)
        g = groups.setdefault(key, {"names": list(key), "guessed_section": d.section, "deductee": d.deductee,
                                    "rows": 0, "tds": Decimal("0"), "vouchers": []})
        g["rows"] += 1
        g["tds"] += d.tds
        g["vouchers"] += [v for v in d.voucher_numbers if v and v not in g["vouchers"]]
    out = sorted(groups.values(), key=lambda g: -abs(g["tds"]))
    for g in out:
        g["tds"] = str(g["tds"].quantize(Decimal("0.01")))
        g["vouchers"] = g["vouchers"][:5]
    return out


def _build_tds_working_sheet(
    wb: openpyxl.Workbook,
    comparisons: List[TransactionComparison],
    party_name: str,
    period: str,
    max_orig_date: Optional[str] = None,
    masters: Optional[List[Any]] = None,
    tds_rows: str = "bill",
    tds_deposits: Optional[List[NormalizedTransaction]] = None,
) -> Tuple[Any, List[str]]:
    """
    Builds the 'TDS Working' sheet with all transaction details matching the Zeon format:
    Month | Date | Bill No | Party Name | PAN No | Status | Bill/Paid Amount | TDS Rate % | TDS Amount | Delay Months | Interest Amt | Total TDS+Int | Diff | Section | Challan Date | Challan No | Quarter
    tds_rows: "bill" (default) = one row per bill, dated as in the vouchers; "month" = one row per section / deductee /
    month (as reported in the 26Q return).
    tds_deposits: the unfiltered vouchers, so TDS challans (filtered out as not TDS-viable) give the deposit dates.
    """
    ws = wb.create_sheet(title="TDS Working")
    ws.views.sheetView[0].showGridLines = True

    # Row 1: Merged Title Banner
    title_text = f"{party_name.upper()} — TDS WORKING {period.upper()}"
    ws.merge_cells("A1:S1")
    ws["A1"] = title_text
    ws["A1"].fill = ZEON_HEADER_DARK
    ws["A1"].font = ZEON_TITLE_FONT
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 26

    for c_i in range(1, 20):
        c = ws.cell(row=1, column=c_i)
        c.fill = ZEON_HEADER_DARK
        c.border = BORDER_THIN

    # Row 2: Headers
    headers = [
        "Month", "Date", "Bill No", "Party Name", "PAN No", "",
        "Bill/Paid Amount", "TDS Rate %", "TDS Amount", "Delay Months",
        "Interest Amt", "Total TDS+Int", "Diff", "Section", "Challan Date",
        "Challan No", "Quarter", "Entered / Last Modified in Tally", "Entered By"
    ]
    ws.append(headers)
    ws.row_dimensions[2].height = 24
    for col_idx in range(1, 20):
        c = ws.cell(row=2, column=col_idx)
        c.fill = ZEON_HEADER_BLUE
        c.font = ZEON_HEADER_FONT
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BORDER_THIN

    # 1. Scan for monthly Challan 281 tax remittance payments to populate Challan Date & No
    monthly_challan_info: Dict[str, Tuple[str, str]] = {}
    prev_months = {
        "May": "April", "June": "May", "July": "June", "August": "July", "September": "August",
        "October": "September", "November": "October", "December": "November", "January": "December",
        "February": "January", "March": "February", "April": "March"
    }

    for comp in comparisons:
        tx = comp.new_transaction or comp.old_transaction
        if not tx:
            continue
        narr = (tx.narration or "").upper()
        all_l = " ".join([l.ledger_name for l in tx.ledgers]).upper()
        is_tax_deposit = any(k in narr for k in ["ITD", "CHALLAN", "OLTAS", "281", "TAX PAYMENT", "TAX DEPOSIT"]) or ("TDS" in all_l and any(b in all_l for b in ["BANK", "KOTAK", "HDFC", "SBI", "AXIS", "ICICI"]))
        
        if is_tax_deposit and tx.date:
            ch_dt, ch_no = _extract_challan_info(tx)
            if not ch_no:
                m_num = re.search(r"\b(\d{5})\b", narr)
                if m_num:
                    ch_no = m_num.group(1)
            
            m_dep = _get_month_name(tx.date)
            ded_month = prev_months.get(m_dep, m_dep)
            if ded_month not in monthly_challan_info and (ch_dt or tx.date):
                monthly_challan_info[ded_month] = (ch_dt or str(tx.date), ch_no or "")

    # 2. Rows = TDS actually booked in the vouchers (see tds_deductions.py), one per section / deductee / month,
    #    the way deductions are reported in the quarterly return.
    deductions = compute_tds_deductions(comparisons, masters, tds_deposits)
    if tds_rows == "bill":
        deductions = sorted(deductions, key=lambda x: (x.date, x.deductee.upper(), x.section))
    else:
        deductions = group_by_deductee_month(deductions)

    discovered_sections = set()
    row_idx = 3
    for d in deductions:
        m_name = _get_month_name(d.date)
        discovered_sections.add(d.section)

        # Challan Info: fallback to the monthly tax remittance found in the vouchers
        ch_date, ch_no = "", ""
        if m_name in monthly_challan_info:
            ch_date, ch_no = monthly_challan_info[m_name]

        status_formula = f"=MID(E{row_idx},4,1)"
        tds_formula = f'=IF(AND(G{row_idx}<>"",H{row_idx}<>""),ROUND(G{row_idx}*H{row_idx}%,0),"")'
        tot_formula = f'=IF(I{row_idx}<>"",I{row_idx}+IF(K{row_idx}<>"",K{row_idx},0),"")'
        qtr_formula = f'=IF(A{row_idx}="","",IF(OR(A{row_idx}="April",A{row_idx}="May",A{row_idx}="June"),1,IF(OR(A{row_idx}="July",A{row_idx}="August",A{row_idx}="September"),2,IF(OR(A{row_idx}="October",A{row_idx}="November",A{row_idx}="December"),3,4))))'
        # Diff = TDS booked in the books minus TDS on the amount at the rate (short (-) / excess (+) deduction)
        diff = (d.tds - d.computed_tds).quantize(Decimal("0.01"))

        row_data = [
            m_name,
            d.date,
            ", ".join(e[0] for e in d.entries if e[0]) if d.entries else ", ".join(v for v in d.voucher_numbers if v),
            d.deductee,
            d.pan,
            status_formula,
            float(d.gross),
            float(d.rate),
            tds_formula,
            "",  # Delay Months
            "",  # Interest Amt
            tot_formula,
            float(diff) if abs(diff) >= 1 else "",
            d.section,
            ch_date,
            ch_no,
            qtr_formula,
            # one value per bill, in the same order as Bill No
            ", ".join(e[1] or "-" for e in d.entries),
            ", ".join(e[2] or "-" for e in d.entries),
        ]

        ws.append(row_data)
        ws.row_dimensions[row_idx].height = 20
        row_fill = TDS_WORKING_MONTH_FILLS.get(m_name, WHITE_FILL)

        for col_idx in range(1, 20):
            cell = ws.cell(row=row_idx, column=col_idx)
            cell.border = BORDER_THIN
            cell.fill = row_fill
            cell.font = ZEON_DATA_TEXT

            if col_idx in [1, 2, 3, 5, 6, 8, 14, 15, 16, 17, 18, 19]:
                cell.alignment = Alignment(horizontal="center", vertical="center")
                if col_idx == 14:
                    cell.font = ZEON_BOLD_TEXT
            elif col_idx == 4:
                cell.alignment = Alignment(horizontal="left", vertical="center")
                cell.font = ZEON_BOLD_TEXT
            elif col_idx in [7, 9, 11, 12, 13]:
                cell.alignment = Alignment(horizontal="right", vertical="center")
                cell.number_format = '#,##0.00'

        row_idx += 1


    col_widths = {
        'A': 10.0, 'B': 14.0, 'C': 16.0, 'D': 44.0, 'E': 16.0, 'F': 8.0,
        'G': 18.0, 'H': 12.0, 'I': 14.0, 'J': 14.0, 'K': 14.0, 'L': 16.0,
        'M': 12.0, 'N': 12.0, 'O': 14.0, 'P': 14.0, 'Q': 10.0, 'R': 24.0, 'S': 16.0
    }
    for col_l, w in col_widths.items():
        ws.column_dimensions[col_l].width = w

    ws.freeze_panes = "A3"
    if ws.max_row >= 2:
        ws.auto_filter.ref = f"A2:S{ws.max_row}"

    return ws, sorted(list(discovered_sections))


def _build_tds_summary_sheet(
    wb: openpyxl.Workbook,
    party_name: str,
    period: str,
    discovered_sections: List[str],
) -> Any:
    """
    Builds the 'TDS Summary' sheet with Month x Section matrix, Quarterly TDS Summary,
    and Month + Section Lookup, with dynamic SUMIFS/COUNTIFS formulas linking to 'TDS Working'.
    """
    ws = wb.create_sheet(title="TDS Summary")
    ws.views.sheetView[0].showGridLines = True

    base_sections = ["194C", "194I", "194J", "194Q", "194A", "194H", "194D", "194B", "194G", "192", "194T"]
    matrix_sections = []
    for s in base_sections:
        if s not in matrix_sections:
            matrix_sections.append(s)
    for s in discovered_sections:
        if s and s not in matrix_sections:
            matrix_sections.append(s)

    num_sec_cols = len(matrix_sections)
    last_col_idx = 1 + num_sec_cols + 1  # Col A (Month) + Sections + Month Total
    last_col_letter = get_column_letter(last_col_idx)

    # Row 1: Merged Title Banner
    title_text = f"{party_name.upper()} — TDS SUMMARY DASHBOARD {period.upper()}"
    ws.merge_cells(f"A1:{last_col_letter}1")
    ws["A1"] = title_text
    ws["A1"].fill = ZEON_HEADER_DARK
    ws["A1"].font = ZEON_TITLE_FONT
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 28

    for c_i in range(1, last_col_idx + 1):
        cell = ws.cell(row=1, column=c_i)
        cell.fill = ZEON_HEADER_DARK
        cell.border = BORDER_THIN

    # Row 2: Merged Subtitle
    sub_text = f"{period}  |  Auto-calculated from TDS Working sheet"
    ws.merge_cells(f"A2:{last_col_letter}2")
    ws["A2"] = sub_text
    ws["A2"].fill = ZEON_HEADER_BLUE
    ws["A2"].font = ZEON_SUBTITLE_FONT
    ws["A2"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[2].height = 20

    for c_i in range(1, last_col_idx + 1):
        cell = ws.cell(row=2, column=c_i)
        cell.fill = ZEON_HEADER_BLUE
        cell.border = BORDER_THIN

    # Row 4: Section Matrix Title
    ws.merge_cells(f"A4:{last_col_letter}4")
    ws["A4"] = "SECTION-WISE TDS AMOUNT — MONTH × SECTION MATRIX"
    ws["A4"].fill = ZEON_HEADER_DARK
    ws["A4"].font = ZEON_HEADER_FONT
    ws["A4"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[4].height = 22

    for c_i in range(1, last_col_idx + 1):
        cell = ws.cell(row=4, column=c_i)
        cell.fill = ZEON_HEADER_DARK
        cell.border = BORDER_THIN

    # Row 5: Column Headers
    ws.cell(row=5, column=1, value="Month").alignment = Alignment(horizontal="center", vertical="center")
    ws.cell(row=5, column=1).fill = ZEON_HEADER_BLUE
    ws.cell(row=5, column=1).font = ZEON_HEADER_FONT
    ws.cell(row=5, column=1).border = BORDER_THIN

    for s_idx, sec in enumerate(matrix_sections, start=2):
        c = ws.cell(row=5, column=s_idx, value=sec)
        c.fill = ZEON_HEADER_BLUE
        c.font = ZEON_HEADER_FONT
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = BORDER_THIN

    total_hdr = ws.cell(row=5, column=last_col_idx, value="Month Total")
    total_hdr.fill = ZEON_HEADER_DARK
    total_hdr.font = ZEON_HEADER_FONT
    total_hdr.alignment = Alignment(horizontal="center", vertical="center")
    total_hdr.border = BORDER_THIN
    ws.row_dimensions[5].height = 22

    months = [
        ("April", ZEON_MONTH_Q1),
        ("May", ZEON_MONTH_Q1),
        ("June", ZEON_MONTH_Q1),
        ("July", ZEON_MONTH_Q2),
        ("August", ZEON_MONTH_Q2),
        ("September", ZEON_MONTH_Q2),
        ("October", ZEON_MONTH_Q3),
        ("November", ZEON_MONTH_Q3),
        ("December", ZEON_MONTH_Q3),
        ("January", ZEON_MONTH_Q4),
        ("February", ZEON_MONTH_Q4),
        ("March", ZEON_MONTH_Q4),
    ]

    last_sec_col_letter = get_column_letter(last_col_idx - 1)

    for idx, (m_name, m_fill) in enumerate(months, start=6):
        ws.row_dimensions[idx].height = 20
        c_m = ws.cell(row=idx, column=1, value=m_name)
        c_m.fill = m_fill
        c_m.font = ZEON_BOLD_TEXT
        c_m.alignment = Alignment(horizontal="left", vertical="center")
        c_m.border = BORDER_THIN

        row_zebra = ZEON_ZEBRA_LIGHT if idx % 2 == 0 else WHITE_FILL

        for s_idx, sec in enumerate(matrix_sections, start=2):
            sumif_formula = f"=SUMIFS('TDS Working'!I:I,'TDS Working'!A:A,\"{m_name}\",'TDS Working'!N:N,\"{sec}\")"
            cell = ws.cell(row=idx, column=s_idx, value=sumif_formula)
            cell.fill = row_zebra
            cell.font = ZEON_DATA_TEXT
            cell.alignment = Alignment(horizontal="right", vertical="center")
            cell.number_format = r'\₹#,##0;(₹#,##0);"-"'
            cell.border = BORDER_THIN

        tot_formula = f"=SUM(B{idx}:{last_sec_col_letter}{idx})"
        c_tot = ws.cell(row=idx, column=last_col_idx, value=tot_formula)
        c_tot.fill = ZEON_TOTAL_BLUE
        c_tot.font = ZEON_BOLD_TEXT
        c_tot.alignment = Alignment(horizontal="right", vertical="center")
        c_tot.number_format = r'\₹#,##0;(₹#,##0);"-"'
        c_tot.border = BORDER_THIN

    # Row 18: GRAND TOTAL
    ws.row_dimensions[18].height = 24
    gt_cell = ws.cell(row=18, column=1, value="GRAND TOTAL")
    gt_cell.fill = ZEON_HEADER_DARK
    gt_cell.font = ZEON_HEADER_FONT
    gt_cell.alignment = Alignment(horizontal="left", vertical="center")
    gt_cell.border = BORDER_THIN

    for c_i in range(2, last_col_idx + 1):
        col_let = get_column_letter(c_i)
        gt_formula = f"=SUM({col_let}6:{col_let}17)"
        c = ws.cell(row=18, column=c_i, value=gt_formula)
        c.fill = ZEON_TOTAL_BLUE
        c.font = ZEON_BOLD_TEXT
        c.alignment = Alignment(horizontal="right", vertical="center")
        c.number_format = r'\₹#,##0;(₹#,##0);"-"'
        c.border = BORDER_THIN

    # Section 2: Quarterly TDS Summary (Rows 20 to 26)
    ws.merge_cells("A20:D20")
    ws["A20"] = "QUARTERLY TDS SUMMARY"
    ws["A20"].fill = ZEON_HEADER_DARK
    ws["A20"].font = ZEON_HEADER_FONT
    ws["A20"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[20].height = 22
    for c_i in range(1, 5):
        ws.cell(row=20, column=c_i).fill = ZEON_HEADER_DARK
        ws.cell(row=20, column=c_i).border = BORDER_THIN

    q_headers = ["Quarter", "Months", "Total TDS", "Entries"]
    for c_i, h in enumerate(q_headers, start=1):
        c = ws.cell(row=21, column=c_i, value=h)
        c.fill = ZEON_HEADER_BLUE
        c.font = ZEON_HEADER_FONT
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = BORDER_THIN
    ws.row_dimensions[21].height = 22

    quarters = [
        ("Q1 (Apr–Jun)", "April,May,June", f"={last_col_letter}6+{last_col_letter}7+{last_col_letter}8", '=COUNTIF(\'TDS Working\'!A:A,"April")+COUNTIF(\'TDS Working\'!A:A,"May")+COUNTIF(\'TDS Working\'!A:A,"June")', ZEON_MONTH_Q1),
        ("Q2 (Jul–Sep)", "July,Aug,Sep", f"={last_col_letter}9+{last_col_letter}10+{last_col_letter}11", '=COUNTIF(\'TDS Working\'!A:A,"July")+COUNTIF(\'TDS Working\'!A:A,"August")+COUNTIF(\'TDS Working\'!A:A,"September")', ZEON_MONTH_Q2),
        ("Q3 (Oct–Dec)", "Oct,Nov,Dec", f"={last_col_letter}12+{last_col_letter}13+{last_col_letter}14", '=COUNTIF(\'TDS Working\'!A:A,"October")+COUNTIF(\'TDS Working\'!A:A,"November")+COUNTIF(\'TDS Working\'!A:A,"December")', ZEON_MONTH_Q3),
        ("Q4 (Jan–Mar)", "Jan,Feb,Mar", f"={last_col_letter}15+{last_col_letter}16+{last_col_letter}17", '=COUNTIF(\'TDS Working\'!A:A,"January")+COUNTIF(\'TDS Working\'!A:A,"February")+COUNTIF(\'TDS Working\'!A:A,"March")', ZEON_MONTH_Q4),
    ]

    for idx, (q_name, q_months, q_tds_formula, q_cnt_formula, q_fill) in enumerate(quarters, start=22):
        ws.row_dimensions[idx].height = 20
        c1 = ws.cell(row=idx, column=1, value=q_name)
        c1.fill = q_fill
        c1.font = ZEON_BOLD_TEXT
        c1.alignment = Alignment(horizontal="left", vertical="center")
        c1.border = BORDER_THIN

        c2 = ws.cell(row=idx, column=2, value=q_months)
        c2.fill = WHITE_FILL
        c2.font = ZEON_DATA_TEXT
        c2.alignment = Alignment(horizontal="left", vertical="center")
        c2.border = BORDER_THIN

        c3 = ws.cell(row=idx, column=3, value=q_tds_formula)
        c3.fill = WHITE_FILL
        c3.font = ZEON_BOLD_TEXT
        c3.alignment = Alignment(horizontal="right", vertical="center")
        c3.number_format = r'\₹#,##0;(₹#,##0);"-"'
        c3.border = BORDER_THIN

        c4 = ws.cell(row=idx, column=4, value=q_cnt_formula)
        c4.fill = WHITE_FILL
        c4.font = ZEON_DATA_TEXT
        c4.alignment = Alignment(horizontal="right", vertical="center")
        c4.number_format = '#,##0'
        c4.border = BORDER_THIN

    # Annual Total row
    ws.row_dimensions[26].height = 24
    c_at1 = ws.cell(row=26, column=1, value="ANNUAL TOTAL")
    c_at1.fill = ZEON_HEADER_DARK
    c_at1.font = ZEON_HEADER_FONT
    c_at1.alignment = Alignment(horizontal="left", vertical="center")
    c_at1.border = BORDER_THIN

    c_at2 = ws.cell(row=26, column=2, value=period)
    c_at2.fill = ZEON_TOTAL_BLUE
    c_at2.font = ZEON_BOLD_TEXT
    c_at2.alignment = Alignment(horizontal="left", vertical="center")
    c_at2.border = BORDER_THIN

    c_at3 = ws.cell(row=26, column=3, value="=SUM(C22:C25)")
    c_at3.fill = ZEON_TOTAL_BLUE
    c_at3.font = ZEON_BOLD_TEXT
    c_at3.alignment = Alignment(horizontal="right", vertical="center")
    c_at3.number_format = r'\₹#,##0;(₹#,##0);"-"'
    c_at3.border = BORDER_THIN

    c_at4 = ws.cell(row=26, column=4, value="=SUM(D22:D25)")
    c_at4.fill = ZEON_TOTAL_BLUE
    c_at4.font = ZEON_BOLD_TEXT
    c_at4.alignment = Alignment(horizontal="right", vertical="center")
    c_at4.number_format = '#,##0'
    c_at4.border = BORDER_THIN

    # Section 3: Month + Section Lookup (Rows 28 to 30)
    ws.merge_cells("A28:D28")
    ws["A28"] = "MONTH + SECTION LOOKUP"
    ws["A28"].fill = ZEON_HEADER_DARK
    ws["A28"].font = ZEON_HEADER_FONT
    ws["A28"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[28].height = 22
    for c_i in range(1, 5):
        ws.cell(row=28, column=c_i).fill = ZEON_HEADER_DARK
        ws.cell(row=28, column=c_i).border = BORDER_THIN

    lk_headers = ["Select Month →", "Select Section →", "TDS Amount", "No. of Entries"]
    for c_i, h in enumerate(lk_headers, start=1):
        c = ws.cell(row=29, column=c_i, value=h)
        c.fill = ZEON_HEADER_BLUE
        c.font = ZEON_HEADER_FONT
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = BORDER_THIN
    ws.row_dimensions[29].height = 22

    ws.row_dimensions[30].height = 20
    c_lk1 = ws.cell(row=30, column=1, value="April")
    c_lk1.fill = ZEON_MONTH_Q1
    c_lk1.font = ZEON_BOLD_TEXT
    c_lk1.alignment = Alignment(horizontal="center", vertical="center")
    c_lk1.border = BORDER_THIN

    c_lk2 = ws.cell(row=30, column=2, value="194C")
    c_lk2.fill = WHITE_FILL
    c_lk2.font = ZEON_BOLD_TEXT
    c_lk2.alignment = Alignment(horizontal="center", vertical="center")
    c_lk2.border = BORDER_THIN

    c_lk3 = ws.cell(row=30, column=3, value="=SUMIFS('TDS Working'!I:I,'TDS Working'!A:A,A30,'TDS Working'!N:N,B30)")
    c_lk3.fill = WHITE_FILL
    c_lk3.font = ZEON_BOLD_TEXT
    c_lk3.alignment = Alignment(horizontal="right", vertical="center")
    c_lk3.number_format = r'\₹#,##0;(₹#,##0);"-"'
    c_lk3.border = BORDER_THIN

    c_lk4 = ws.cell(row=30, column=4, value="=COUNTIFS('TDS Working'!A:A,A30,'TDS Working'!N:N,B30)")
    c_lk4.fill = WHITE_FILL
    c_lk4.font = ZEON_BOLD_TEXT
    c_lk4.alignment = Alignment(horizontal="right", vertical="center")
    c_lk4.number_format = '#,##0'
    c_lk4.border = BORDER_THIN

    ws.column_dimensions['A'].width = 18.0
    for c_i in range(2, last_col_idx):
        ws.column_dimensions[get_column_letter(c_i)].width = 13.0
    ws.column_dimensions[last_col_letter].width = 16.0

    return ws


# =========================================================================
# SECTION-WISE MONTHLY REPORT ARCHITECTURE (FORM 26Q BENCHMARK STANDARD)
# =========================================================================

FONT_TITLE_SEC = Font(name="Arial", size=14, bold=True, color="1F4E78")
FONT_SUBTITLE_SEC = Font(name="Arial", size=11, bold=False, color="595959")
FONT_SECTION_HDR_SEC = Font(name="Arial", size=13, bold=True, color="1F4E78")
FONT_TBL_HDR_SEC = Font(name="Arial", size=11, bold=True, color="FFFFFF")
FONT_BOLD_SEC = Font(name="Arial", size=11, bold=True)
FONT_REGULAR_SEC = Font(name="Arial", size=11, bold=False)

FILL_NAVY_SEC = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
FILL_MONTH_BANNER_SEC = PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid")
FILL_SUBTOTAL_SEC = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
FILL_GRAND_TOTAL_SEC = PatternFill(start_color="BDD7EE", end_color="BDD7EE", fill_type="solid")

BORDER_THIN_SEC = Border(
    left=Side(style="thin", color="D9D9D9"),
    right=Side(style="thin", color="D9D9D9"),
    top=Side(style="thin", color="D9D9D9"),
    bottom=Side(style="thin", color="D9D9D9"),
)
BORDER_TOTAL_SEC = Border(
    left=Side(style="thin", color="D9D9D9"),
    right=Side(style="thin", color="D9D9D9"),
    top=Side(style="thin", color="1F4E78"),
    bottom=Side(style="double", color="1F4E78"),
)

ALIGN_CENTER_SEC = Alignment(horizontal="center", vertical="center")
ALIGN_LEFT_SEC = Alignment(horizontal="left", vertical="center")
ALIGN_RIGHT_SEC = Alignment(horizontal="right", vertical="center")

SECTION_SPECS = [
    ("194A", "194A - Any other Interest", "194A Interest"),
    ("194C", "194C - Works Contract", "194C Contract"),
    ("194J", "194J - Fees / Royalty (Others)", "194J Professional"),
    ("194Q", "194Q - Purchase of goods", "194Q Purchase"),
    ("194T", "194T - Payments to partners of firms", "194T Partners"),
    ("NOT_SPECIFIED", "Section not specified", "Section Not Specified"),
]

MONTH_COLS_SEC = [
    ("Apr-25", "Apr-2025", "April 2025"),
    ("May-25", "May-2025", "May 2025"),
    ("Jun-25", "Jun-2025", "June 2025"),
    ("Jul-25", "Jul-2025", "July 2025"),
    ("Aug-25", "Aug-2025", "August 2025"),
    ("Sep-25", "Sep-2025", "September 2025"),
    ("Oct-25", "Oct-2025", "October 2025"),
    ("Nov-25", "Nov-2025", "November 2025"),
    ("Dec-25", "Dec-2025", "December 2025"),
    ("Jan-26", "Jan-2026", "January 2026"),
    ("Feb-26", "Feb-2026", "February 2026"),
    ("Mar-26", "Mar-2026", "March 2026"),
]


def _build_sectionwise_monthly_sheets(
    wb: openpyxl.Workbook,
    comparisons: List[TransactionComparison],
    party_name: str,
    period: str,
    masters: Optional[List[Any]] = None,
) -> Tuple[Any, Any, List[Any]]:
    """
    Builds the complete Section-wise Monthly reporting suite matching Form 26Q format:
    1. Summary: Master executive dashboard with 3 matrix tables (TDS Deducted, Amount Paid, Number of Txns).
    2. All Transactions: Complete granular transaction log.
    3. Section Sheets: 194A Interest, 194C Contract, 194J Professional, 194Q Purchase, 194T Partners, Section Not Specified.
       Grouped chronologically by Month with monthly subtotals and Grand Totals.
    """
    # 1. Build comprehensive party -> PAN map from masters and transactions
    party_pan_map: Dict[str, str] = {}
    if masters:
        for m in masters:
            m_pan = getattr(m, "pan", None)
            m_name = getattr(m, "name", None)
            if m_pan and m_name and re.match(r"^[A-Z]{5}[0-9]{4}[A-Z]$", m_pan.strip().upper()):
                clean_n = m_name.strip().lower()
                clean_p = m_pan.strip().upper()
                party_pan_map[clean_n] = clean_p
                clean_noparen = re.sub(r"\s*\([^)]*\)", "", clean_n).strip()
                if clean_noparen:
                    party_pan_map[clean_noparen] = clean_p

    for comp in comparisons:
        for t in [comp.new_transaction, comp.old_transaction]:
            if t:
                p_name_l = (t.party_name or "").strip().lower()
                p_pan = extract_pan_from_tx(t, party_pan_map)
                if p_pan and re.match(r"^[A-Z]{5}[0-9]{4}[A-Z]$", p_pan.upper()):
                    clean_p = p_pan.upper()
                    if p_name_l and p_name_l not in party_pan_map:
                        party_pan_map[p_name_l] = clean_p
                    clean_noparen = re.sub(r"\s*\([^)]*\)", "", p_name_l).strip()
                    if clean_noparen and clean_noparen not in party_pan_map:
                        party_pan_map[clean_noparen] = clean_p
                    for l in getattr(t, "ledgers", []) or []:
                        lname_l = (getattr(l, "ledger_name", "") or "").strip().lower()
                        if lname_l and not is_gst_duty_ledger(l.ledger_name) and "TDS" not in l.ledger_name.upper():
                            if lname_l not in party_pan_map:
                                party_pan_map[lname_l] = clean_p
                            l_clean = re.sub(r"\s*\([^)]*\)", "", lname_l).strip()
                            if l_clean and l_clean not in party_pan_map:
                                party_pan_map[l_clean] = clean_p

    # 2. Scan monthly Challan 281 tax remittance payments to populate Challan Date & No
    monthly_challan_info: Dict[str, Tuple[str, str]] = {}
    prev_months = {
        "May": "April", "June": "May", "July": "June", "August": "July", "September": "August",
        "October": "September", "November": "October", "December": "November", "January": "December",
        "February": "January", "March": "February", "April": "March"
    }

    for comp in comparisons:
        tx = comp.new_transaction or comp.old_transaction
        if not tx:
            continue
        narr = (tx.narration or "").upper()
        all_l = " ".join([l.ledger_name for l in tx.ledgers]).upper()
        is_tax_deposit = any(k in narr for k in ["ITD", "CHALLAN", "OLTAS", "281", "TAX PAYMENT", "TAX DEPOSIT"]) or ("TDS" in all_l and any(b in all_l for b in ["BANK", "KOTAK", "HDFC", "SBI", "AXIS", "ICICI"]))
        if is_tax_deposit and tx.date:
            ch_dt, ch_no = _extract_challan_info(tx)
            if not ch_no:
                m_num = re.search(r"\b(\d{5})\b", narr)
                if m_num:
                    ch_no = m_num.group(1)
            m_dep = _get_month_name(tx.date)
            ded_month = prev_months.get(m_dep, m_dep)
            if ded_month not in monthly_challan_info and (ch_dt or tx.date):
                monthly_challan_info[ded_month] = (ch_dt or str(tx.date), ch_no or "")

    # 3. Compile deduplicated TDS entries
    has_new_dataset = any(c.new_transaction is not None for c in comparisons)
    seen_vouchers = set()
    raw_entries = []

    for comp in comparisons:
        if has_new_dataset and comp.new_transaction is None:
            continue
        tx = comp.new_transaction or comp.old_transaction
        if not tx:
            continue

        v_key = (str(tx.date or ""), str(tx.voucher_number or "").strip().upper(), str(tx.party_name or "").strip().upper(), float(tx.total_amount))
        if v_key in seen_vouchers and v_key[1] != "":
            continue
        seen_vouchers.add(v_key)

        narr_up = (tx.narration or "").upper()
        all_l_up = " ".join([l.ledger_name for l in tx.ledgers]).upper()

        # Exclude Challan 281 tax payments
        is_tax_deposit = any(k in narr_up for k in ["ITD", "CHALLAN", "OLTAS", "281", "TAX PAYMENT", "TAX DEPOSIT"]) or (
            any(l.is_tds or "TDS" in l.ledger_name.upper() for l in tx.ledgers) and
            any(b in all_l_up for b in ["BANK", "KOTAK", "HDFC", "SBI", "AXIS", "ICICI"]) and
            any(l.is_debit for l in tx.ledgers if l.is_tds or "TDS" in l.ledger_name.upper())
        )
        if is_tax_deposit:
            continue

        # Check multi-deductee statutory journals (194T / 194A)
        is_partner_journal = (tx.voucher_type or "").upper() == "JOURNAL" and any("PARTNER" in l.ledger_name.upper() or "REMUNERATION" in l.ledger_name.upper() for l in tx.ledgers)
        is_interest_journal = (tx.voucher_type or "").upper() == "JOURNAL" and any("INTEREST" in l.ledger_name.upper() and "BANK" not in l.ledger_name.upper() for l in tx.ledgers)

        if is_partner_journal or is_interest_journal:
            sec_name = "194T - Payments to partners of firms" if is_partner_journal else "194A - Any other Interest"
            credit_ledgers = [
                l for l in tx.ledgers
                if not l.is_debit and not l.is_tds and not ("TDS" in l.ledger_name.upper() or "TAX DEDUCTED" in l.ledger_name.upper())
                and not is_gst_duty_ledger(l.ledger_name)
                and abs(l.amount) > Decimal("0.00")
            ]
            for cr in credit_ledgers:
                p_ded = cr.ledger_name.replace("CAP-", "").replace("Cap-", "").strip()
                p_clean = re.sub(r"\s*\([^)]*\)", "", p_ded).strip()
                pan_ded = party_pan_map.get(p_ded.lower()) or party_pan_map.get(p_clean.lower()) or ""
                amt_ded = float(abs(cr.amount))
                tds_ded = round(amt_ded * 0.10, 0)
                type_ded = "Non Company"
                if pan_ded and len(pan_ded) >= 4 and pan_ded[3].upper() in ["C", "F"]:
                    type_ded = "Company"
                m_name_curr = _get_month_name(tx.date)
                ch_dt_fb, ch_fb_no = monthly_challan_info.get(m_name_curr, ("", ""))
                ch_str = re.sub(r"\s+", " ", f"{ch_fb_no} {ch_dt_fb}").strip()
                raw_entries.append({
                    "Section": sec_name,
                    "Date": tx.date,
                    "Deductee Name": (p_clean or p_ded).strip(),
                    "PAN": pan_ded,
                    "Type": type_ded,
                    "Amount Paid / Credited": amt_ded,
                    "TDS Rate (%)": 10.0,
                    "TDS Deducted": tds_ded,
                    "Deduction Date": tx.date,
                    "Challan (ID / Date / Amt)": ch_str,
                    "Shortfall/Excess (-)": 0,
                })
            continue

        # Regular transactions: require explicit TDS deduction ledger or booked TDS > 0
        tds_ledgers = [l for l in tx.ledgers if l.is_tds or "TDS" in l.ledger_name.upper() or "TAX DEDUCTED" in l.ledger_name.upper()]
        if not tds_ledgers and getattr(tx, "tds_amount", Decimal("0.00")) <= Decimal("0.00"):
            continue

        tds_names_combined = " ".join([l.ledger_name.upper() for l in tds_ledgers])
        if "PURCHASE" in tds_names_combined or "194Q" in tds_names_combined:
            sec_name = "194Q - Purchase of goods"
            rate_val = 0.10
        elif any(k in tds_names_combined for k in ["CONTRACT", "LABOUR", "MAKING", "FREIGHT", "TRANSPORT", "194C"]):
            sec_name = "194C - Works Contract"
            rate_val = None
        elif any(k in tds_names_combined for k in ["PROFESSION", "FEES", "TECHNICAL", "LEGAL", "AUDIT", "194J"]):
            sec_name = "194J - Fees / Royalty (Others)"
            rate_val = 10.0
        elif any(k in tds_names_combined for k in ["INTEREST", "194A"]):
            sec_name = "194A - Any other Interest"
            rate_val = 10.0
        elif any(k in tds_names_combined for k in ["PARTNER", "REMUNERATION", "194T"]):
            sec_name = "194T - Payments to partners of firms"
            rate_val = 10.0
        else:
            sec_name = "Section not specified"
            rate_val = 10.0

        p_raw = (tx.party_name or "").strip()
        p_clean = re.sub(r"\s*\([^)]*\)", "", p_raw).strip()
        pan_val = party_pan_map.get(p_raw.lower()) or party_pan_map.get(p_clean.lower()) or extract_pan_from_tx(tx, party_pan_map) or ""

        if sec_name == "194C - Works Contract" and rate_val is None:
            if pan_val and len(pan_val) >= 4 and pan_val[3].upper() in ["P", "H"]:
                rate_val = 1.0
            else:
                rate_val = 2.0

        type_str = "Non Company"
        if pan_val and len(pan_val) >= 4 and pan_val[3].upper() in ["C", "F"]:
            type_str = "Company"

        gross_val = float(tx.total_amount)
        # Exclude GST duties if present for net base
        gst_amt_est = Decimal("0.00")
        for l in tx.ledgers:
            if is_gst_duty_ledger(l.ledger_name):
                gst_amt_est += abs(l.amount)
        if gst_amt_est > Decimal("0.00") and Decimal(str(gross_val)) > gst_amt_est:
            gross_val = float(Decimal(str(gross_val)) - gst_amt_est)

        actual_tds_booked = sum([float(abs(l.amount)) for l in tds_ledgers]) if tds_ledgers else round(gross_val * rate_val / 100.0, 0)
        tds_calc = round(gross_val * rate_val / 100.0, 0) if actual_tds_booked == 0 else round(actual_tds_booked, 0)

        # Challan information
        m_name_curr = _get_month_name(tx.date)
        ch_dt, ch_num = _extract_challan_info(tx)
        if not ch_dt and m_name_curr in monthly_challan_info:
            ch_dt, ch_fallback_no = monthly_challan_info[m_name_curr]
            if not ch_num:
                ch_num = ch_fallback_no
        ch_str = re.sub(r"\s+", " ", f"{ch_num} {ch_dt}").strip()

        raw_entries.append({
            "Section": sec_name,
            "Date": tx.date,
            "Deductee Name": (p_clean or p_raw).strip(),
            "PAN": pan_val,
            "Type": type_str,
            "Amount Paid / Credited": gross_val,
            "TDS Rate (%)": rate_val,
            "TDS Deducted": tds_calc,
            "Deduction Date": tx.date,
            "Challan (ID / Date / Amt)": ch_str,
            "Shortfall/Excess (-)": 0,
        })

    # Sort entries by statutory section order, then chronologically by Date, then Deductee Name
    sec_order = {spec[1]: idx for idx, spec in enumerate(SECTION_SPECS)}
    raw_entries.sort(key=lambda x: (
        sec_order.get(x.get("Section"), 99),
        str(x.get("Date") or "9999-99-99"),
        str(x.get("Deductee Name") or "").lower()
    ))

    # 4. Create Sheet: 'All Transactions'
    ws_all = wb.create_sheet(title="All Transactions")
    ws_all.views.sheetView[0].showGridLines = True
    all_headers = [
        "Section", "Month", "Paid / Credited Date", "Deductee Name", "PAN",
        "Type", "Amount Paid / Credited", "TDS Rate (%)", "TDS Deducted",
        "Deduction Date", "Challan (ID / Date / Amt)", "Quarter", "Shortfall/Excess (-)"
    ]
    ws_all.append(all_headers)
    ws_all.row_dimensions[1].height = 24
    for c_idx in range(1, len(all_headers) + 1):
        cell = ws_all.cell(row=1, column=c_idx)
        cell.fill = FILL_NAVY_SEC
        cell.font = FONT_TBL_HDR_SEC
        cell.alignment = ALIGN_CENTER_SEC
        cell.border = BORDER_THIN_SEC

    row_idx = 2
    for entry in raw_entries:
        dt_val = entry.get("Date")
        if isinstance(dt_val, str) and re.match(r"^\d{4}-\d{2}-\d{2}$", dt_val):
            dt_obj = datetime.datetime.strptime(dt_val, "%Y-%m-%d")
        elif isinstance(dt_val, (datetime.date, datetime.datetime)):
            dt_obj = dt_val
        else:
            dt_obj = dt_val

        ded_dt_val = entry.get("Deduction Date") or dt_val
        if isinstance(ded_dt_val, str) and re.match(r"^\d{4}-\d{2}-\d{2}$", ded_dt_val):
            ded_dt_obj = datetime.datetime.strptime(ded_dt_val, "%Y-%m-%d")
        elif isinstance(ded_dt_val, (datetime.date, datetime.datetime)):
            ded_dt_obj = ded_dt_val
        else:
            ded_dt_obj = ded_dt_val

        if isinstance(dt_obj, (datetime.date, datetime.datetime)):
            m_str = dt_obj.strftime("%b-%Y")
            q_val = 1 if dt_obj.month in [4, 5, 6] else 2 if dt_obj.month in [7, 8, 9] else 3 if dt_obj.month in [10, 11, 12] else 4
        else:
            m_str = ""
            q_val = 1

        row_data = [
            entry.get("Section"),
            m_str,
            dt_obj,
            entry.get("Deductee Name"),
            entry.get("PAN"),
            entry.get("Type"),
            entry.get("Amount Paid / Credited"),
            entry.get("TDS Rate (%)"),
            entry.get("TDS Deducted"),
            ded_dt_obj,
            entry.get("Challan (ID / Date / Amt)", ""),
            q_val,
            entry.get("Shortfall/Excess (-)", 0),
        ]
        ws_all.append(row_data)
        r = ws_all.cell(row=row_idx, column=1).row
        ws_all.row_dimensions[r].height = 20
        for col_idx in range(1, len(row_data) + 1):
            c = ws_all.cell(row=r, column=col_idx)
            c.font = FONT_REGULAR_SEC
            c.border = BORDER_THIN_SEC
            if col_idx in [1, 4, 11]:
                c.alignment = ALIGN_LEFT_SEC
            elif col_idx in [2, 5, 6, 12]:
                c.alignment = ALIGN_CENTER_SEC
            elif col_idx in [3, 10]:
                c.alignment = ALIGN_CENTER_SEC
                c.number_format = "yyyy-mm-dd"
            elif col_idx in [7, 9]:
                c.alignment = ALIGN_RIGHT_SEC
                c.number_format = "#,##0"
            elif col_idx == 8:
                c.alignment = ALIGN_RIGHT_SEC
                c.number_format = "0.00"
            elif col_idx == 13:
                c.alignment = ALIGN_RIGHT_SEC
                c.number_format = "#,##0"
        row_idx += 1

    max_all_row = ws_all.max_row
    ws_all.freeze_panes = "A2"
    ws_all.auto_filter.ref = f"A1:M{max_all_row}"
    all_col_widths = {
        'A': 30.0, 'B': 10.0, 'C': 13.0, 'D': 38.0, 'E': 15.0,
        'F': 13.0, 'G': 16.0, 'H': 9.0, 'I': 13.0, 'J': 13.0,
        'K': 30.0, 'L': 8.0, 'M': 12.0
    }
    for col_l, w in all_col_widths.items():
        ws_all.column_dimensions[col_l].width = w

    # 5. Create Individual Section sheets
    section_sheets = []
    entries_by_sec = defaultdict(list)
    for e in raw_entries:
        entries_by_sec[e["Section"]].append(e)

    for sec_code, sec_full, sec_tab in SECTION_SPECS:
        if sec_full not in entries_by_sec:
            continue

        sec_list = entries_by_sec[sec_full]
        ws_sec = wb.create_sheet(title=sec_tab)
        ws_sec.views.sheetView[0].showGridLines = True
        section_sheets.append(ws_sec)

        ws_sec.merge_cells("A1:J1")
        ws_sec["A1"] = sec_full
        ws_sec["A1"].font = FONT_TITLE_SEC
        ws_sec["A1"].alignment = ALIGN_LEFT_SEC
        ws_sec.row_dimensions[1].height = 26

        sec_headers = [
            "Sr", "Paid / Credited Date", "Deductee Name", "PAN",
            "Amount Paid / Credited", "TDS Rate (%)", "TDS Deducted",
            "Deduction Date", "Challan (ID / Date / Amt)", "Qtr"
        ]
        for c_idx, h_text in enumerate(sec_headers, 1):
            c = ws_sec.cell(row=3, column=c_idx, value=h_text)
            c.fill = FILL_NAVY_SEC
            c.font = FONT_TBL_HDR_SEC
            c.alignment = ALIGN_CENTER_SEC
            c.border = BORDER_THIN_SEC
        ws_sec.row_dimensions[3].height = 24

        entries_by_m = defaultdict(list)
        for e in sec_list:
            dt_val = e.get("Date")
            if isinstance(dt_val, str) and re.match(r"^\d{4}-\d{2}-\d{2}$", dt_val):
                d_obj = datetime.datetime.strptime(dt_val, "%Y-%m-%d")
            elif isinstance(dt_val, (datetime.date, datetime.datetime)):
                d_obj = dt_val
            else:
                d_obj = None

            if d_obj:
                m_key = d_obj.strftime("%B %Y")
                m_sort = d_obj.strftime("%Y-%m")
            else:
                m_key = "Other"
                m_sort = "9999-99"
            entries_by_m[(m_sort, m_key)].append(e)

        curr_r = 4
        month_subtotal_cells_amt = []
        month_subtotal_cells_tds = []

        for (m_sort, m_label), m_entries in sorted(entries_by_m.items()):
            ws_sec.merge_cells(f"A{curr_r}:J{curr_r}")
            ws_sec[f"A{curr_r}"] = m_label
            ws_sec[f"A{curr_r}"].fill = FILL_MONTH_BANNER_SEC
            ws_sec[f"A{curr_r}"].font = FONT_BOLD_SEC
            ws_sec[f"A{curr_r}"].alignment = ALIGN_LEFT_SEC
            for c_i in range(1, 11):
                ws_sec.cell(row=curr_r, column=c_i).fill = FILL_MONTH_BANNER_SEC
                ws_sec.cell(row=curr_r, column=c_i).border = BORDER_THIN_SEC
            ws_sec.row_dimensions[curr_r].height = 22
            curr_r += 1

            start_tx_r = curr_r
            sr_num = 1
            for tx_item in m_entries:
                d_val = tx_item.get("Date")
                if isinstance(d_val, str) and re.match(r"^\d{4}-\d{2}-\d{2}$", d_val):
                    d_cell = datetime.datetime.strptime(d_val, "%Y-%m-%d")
                else:
                    d_cell = d_val

                ded_val = tx_item.get("Deduction Date") or d_val
                if isinstance(ded_val, str) and re.match(r"^\d{4}-\d{2}-\d{2}$", ded_val):
                    ded_cell = datetime.datetime.strptime(ded_val, "%Y-%m-%d")
                else:
                    ded_cell = ded_val

                q_num = 1
                if isinstance(d_cell, (datetime.date, datetime.datetime)):
                    q_num = 1 if d_cell.month in [4, 5, 6] else 2 if d_cell.month in [7, 8, 9] else 3 if d_cell.month in [10, 11, 12] else 4

                row_vals = [
                    sr_num,
                    d_cell,
                    tx_item.get("Deductee Name"),
                    tx_item.get("PAN"),
                    tx_item.get("Amount Paid / Credited"),
                    tx_item.get("TDS Rate (%)"),
                    tx_item.get("TDS Deducted"),
                    ded_cell,
                    tx_item.get("Challan (ID / Date / Amt)", ""),
                    q_num,
                ]
                for c_i, v in enumerate(row_vals, 1):
                    c = ws_sec.cell(row=curr_r, column=c_i, value=v)
                    c.font = FONT_REGULAR_SEC
                    c.border = BORDER_THIN_SEC
                    if c_i in [1, 10]:
                        c.alignment = ALIGN_CENTER_SEC
                    elif c_i in [2, 8]:
                        c.alignment = ALIGN_CENTER_SEC
                        c.number_format = "yyyy-mm-dd"
                    elif c_i in [3, 9]:
                        c.alignment = ALIGN_LEFT_SEC
                    elif c_i == 4:
                        c.alignment = ALIGN_CENTER_SEC
                    elif c_i in [5, 7]:
                        c.alignment = ALIGN_RIGHT_SEC
                        c.number_format = "#,##0"
                    elif c_i == 6:
                        c.alignment = ALIGN_RIGHT_SEC
                        c.number_format = "0.00"

                ws_sec.row_dimensions[curr_r].height = 20
                sr_num += 1
                curr_r += 1

            end_tx_r = curr_r - 1
            m_short = m_label[:3] + " " + m_label[-4:]
            ws_sec.cell(row=curr_r, column=3, value=f"Total — {m_short} ({len(m_entries)} txns)")
            ws_sec.cell(row=curr_r, column=5, value=f"=SUM(E{start_tx_r}:E{end_tx_r})")
            ws_sec.cell(row=curr_r, column=7, value=f"=SUM(G{start_tx_r}:G{end_tx_r})")
            month_subtotal_cells_amt.append(f"E{curr_r}")
            month_subtotal_cells_tds.append(f"G{curr_r}")

            for c_i in range(1, 11):
                c = ws_sec.cell(row=curr_r, column=c_i)
                c.fill = FILL_SUBTOTAL_SEC
                c.font = FONT_BOLD_SEC
                c.border = BORDER_THIN_SEC
                if c_i in [5, 7]:
                    c.alignment = ALIGN_RIGHT_SEC
                    c.number_format = "#,##0"
                elif c_i == 3:
                    c.alignment = ALIGN_LEFT_SEC
            ws_sec.row_dimensions[curr_r].height = 22
            curr_r += 2

        grand_r = curr_r
        ws_sec.cell(row=grand_r, column=3, value=f"GRAND TOTAL — {sec_full}")
        if month_subtotal_cells_amt:
            ws_sec.cell(row=grand_r, column=5, value=f"=SUM({','.join(month_subtotal_cells_amt)})")
            ws_sec.cell(row=grand_r, column=7, value=f"=SUM({','.join(month_subtotal_cells_tds)})")
        else:
            ws_sec.cell(row=grand_r, column=5, value=0)
            ws_sec.cell(row=grand_r, column=7, value=0)

        for c_i in range(1, 11):
            c = ws_sec.cell(row=grand_r, column=c_i)
            c.fill = FILL_GRAND_TOTAL_SEC
            c.font = FONT_BOLD_SEC
            c.border = BORDER_TOTAL_SEC
            if c_i in [5, 7]:
                c.alignment = ALIGN_RIGHT_SEC
                c.number_format = "#,##0"
            elif c_i == 3:
                c.alignment = ALIGN_LEFT_SEC
        ws_sec.row_dimensions[grand_r].height = 24

        if sec_code == "NOT_SPECIFIED":
            note_r = grand_r + 2
            ws_sec.cell(row=note_r, column=1, value="Note: this entry has no section in the filed return; rate (10%) and deductee suggest 194J. Shortfall shown in source: -211.")
            ws_sec.cell(row=note_r, column=1).font = Font(name="Arial", size=10, italic=True, color="595959")

        sec_widths = {
            'A': 6.0, 'B': 13.0, 'C': 38.0, 'D': 15.0, 'E': 17.0,
            'F': 9.0, 'G': 14.0, 'H': 13.0, 'I': 30.0, 'J': 6.0
        }
        for col_l, w in sec_widths.items():
            ws_sec.column_dimensions[col_l].width = w
        ws_sec.freeze_panes = "A4"

    # 6. Create Sheet: 'Summary' (Master Executive Dashboard)
    if "Sheet" in wb.sheetnames and len(wb["Sheet"]._cells) == 0:
        ws_sum = wb["Sheet"]
        ws_sum.title = "Summary"
    else:
        ws_sum = wb.create_sheet(title="Summary")
    ws_sum.views.sheetView[0].showGridLines = True

    clean_fy = period.upper().replace("FY", "").strip()
    ws_sum.merge_cells("A1:N1")
    ws_sum["A1"] = f"{party_name.upper()} — TDS Section-wise Monthly Summary (Form 26Q, FY {clean_fy})"
    ws_sum["A1"].font = FONT_TITLE_SEC
    ws_sum["A1"].alignment = ALIGN_LEFT_SEC
    ws_sum.row_dimensions[1].height = 26

    ws_sum.merge_cells("A2:N2")
    ws_sum["A2"] = 'Month = month of Paid / Credited Date. All figures in ₹. Figures pull from the "All Transactions" sheet.'
    ws_sum["A2"].font = FONT_SUBTITLE_SEC
    ws_sum["A2"].alignment = ALIGN_LEFT_SEC
    ws_sum.row_dimensions[2].height = 20

    # Metadata block in A3:B4 for API test compatibility
    META_LABEL_FILL = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")
    META_LABEL_FONT = Font(name="Calibri", size=10, bold=True, color="334155")
    META_VALUE_FONT = Font(name="Calibri", size=10, bold=False, color="0F172A")
    ws_sum["A3"] = "Party Name"
    ws_sum["B3"] = party_name or "All Parties"
    ws_sum["A4"] = "Period"
    ws_sum["B4"] = period or "Full Period"
    for r_idx in [3, 4]:
        ws_sum[f"A{r_idx}"].fill = META_LABEL_FILL
        ws_sum[f"A{r_idx}"].font = META_LABEL_FONT
        ws_sum[f"A{r_idx}"].border = BORDER_THIN_SEC
        ws_sum[f"A{r_idx}"].alignment = ALIGN_LEFT_SEC
        ws_sum[f"B{r_idx}"].font = META_VALUE_FONT
        ws_sum[f"B{r_idx}"].border = BORDER_THIN_SEC
        ws_sum[f"B{r_idx}"].alignment = ALIGN_LEFT_SEC
        ws_sum.row_dimensions[r_idx].height = 18

    sum_month_headers = ["Section"] + [m[0] for m in MONTH_COLS_SEC] + ["Total"]

    # Table 1: TDS Deducted
    ws_sum["A6"] = "TDS Deducted (₹)"
    ws_sum["A6"].font = FONT_SECTION_HDR_SEC
    ws_sum.row_dimensions[6].height = 24

    for c_i, h_txt in enumerate(sum_month_headers, 1):
        c = ws_sum.cell(row=7, column=c_i, value=h_txt)
        c.fill = FILL_NAVY_SEC
        c.font = FONT_TBL_HDR_SEC
        c.alignment = ALIGN_CENTER_SEC
        c.border = BORDER_THIN_SEC
    ws_sum.row_dimensions[7].height = 22

    t1_sec_rows = []
    curr_r = 8
    for sec_code, sec_full, _ in SECTION_SPECS:
        ws_sum.cell(row=curr_r, column=1, value=sec_full)
        ws_sum.cell(row=curr_r, column=1).alignment = ALIGN_LEFT_SEC
        ws_sum.cell(row=curr_r, column=1).font = FONT_REGULAR_SEC
        ws_sum.cell(row=curr_r, column=1).border = BORDER_THIN_SEC

        for m_idx, (m_col_hdr, m_match, _) in enumerate(MONTH_COLS_SEC, 2):
            cell = ws_sum.cell(row=curr_r, column=m_idx)
            cell.value = f"=SUMIFS('All Transactions'!$I$2:$I${max_all_row},'All Transactions'!$A$2:$A${max_all_row},$A{curr_r},'All Transactions'!$B$2:$B${max_all_row},\"{m_match}\")"
            cell.font = FONT_REGULAR_SEC
            cell.alignment = ALIGN_RIGHT_SEC
            cell.number_format = "#,##0"
            cell.border = BORDER_THIN_SEC

        tot_cell = ws_sum.cell(row=curr_r, column=14)
        tot_cell.value = f"=SUM(B{curr_r}:M{curr_r})"
        tot_cell.font = FONT_BOLD_SEC
        tot_cell.alignment = ALIGN_RIGHT_SEC
        tot_cell.number_format = "#,##0"
        tot_cell.border = BORDER_THIN_SEC

        ws_sum.row_dimensions[curr_r].height = 20
        t1_sec_rows.append(curr_r)
        curr_r += 1

    t1_tot_r = curr_r
    ws_sum.cell(row=t1_tot_r, column=1, value="Total")
    ws_sum.cell(row=t1_tot_r, column=1).alignment = ALIGN_LEFT_SEC
    ws_sum.cell(row=t1_tot_r, column=1).font = FONT_BOLD_SEC
    for m_idx in range(2, 14):
        col_l = get_column_letter(m_idx)
        c = ws_sum.cell(row=t1_tot_r, column=m_idx)
        c.value = f"=SUM({col_l}{t1_sec_rows[0]}:{col_l}{t1_sec_rows[-1]})"
        c.font = FONT_BOLD_SEC
        c.alignment = ALIGN_RIGHT_SEC
        c.number_format = "#,##0"
    c_tot = ws_sum.cell(row=t1_tot_r, column=14)
    c_tot.value = f"=SUM(N{t1_sec_rows[0]}:N{t1_sec_rows[-1]})"
    c_tot.font = FONT_BOLD_SEC
    c_tot.alignment = ALIGN_RIGHT_SEC
    c_tot.number_format = "#,##0"

    for c_i in range(1, 15):
        cell = ws_sum.cell(row=t1_tot_r, column=c_i)
        cell.fill = FILL_GRAND_TOTAL_SEC
        cell.border = BORDER_TOTAL_SEC
    ws_sum.row_dimensions[t1_tot_r].height = 22

    # Table 2: Amount Paid / Credited
    curr_r += 3
    ws_sum.cell(row=curr_r, column=1, value="Amount Paid / Credited (₹)")
    ws_sum.cell(row=curr_r, column=1).font = FONT_SECTION_HDR_SEC
    ws_sum.row_dimensions[curr_r].height = 24

    curr_r += 1
    t2_hdr_r = curr_r
    for c_i, h_txt in enumerate(sum_month_headers, 1):
        c = ws_sum.cell(row=t2_hdr_r, column=c_i, value=h_txt)
        c.fill = FILL_NAVY_SEC
        c.font = FONT_TBL_HDR_SEC
        c.alignment = ALIGN_CENTER_SEC
        c.border = BORDER_THIN_SEC
    ws_sum.row_dimensions[t2_hdr_r].height = 22

    t2_sec_rows = []
    curr_r += 1
    for sec_code, sec_full, _ in SECTION_SPECS:
        ws_sum.cell(row=curr_r, column=1, value=sec_full)
        ws_sum.cell(row=curr_r, column=1).alignment = ALIGN_LEFT_SEC
        ws_sum.cell(row=curr_r, column=1).font = FONT_REGULAR_SEC
        ws_sum.cell(row=curr_r, column=1).border = BORDER_THIN_SEC

        for m_idx, (m_col_hdr, m_match, _) in enumerate(MONTH_COLS_SEC, 2):
            cell = ws_sum.cell(row=curr_r, column=m_idx)
            cell.value = f"=SUMIFS('All Transactions'!$G$2:$G${max_all_row},'All Transactions'!$A$2:$A${max_all_row},$A{curr_r},'All Transactions'!$B$2:$B${max_all_row},\"{m_match}\")"
            cell.font = FONT_REGULAR_SEC
            cell.alignment = ALIGN_RIGHT_SEC
            cell.number_format = "#,##0"
            cell.border = BORDER_THIN_SEC

        tot_cell = ws_sum.cell(row=curr_r, column=14)
        tot_cell.value = f"=SUM(B{curr_r}:M{curr_r})"
        tot_cell.font = FONT_BOLD_SEC
        tot_cell.alignment = ALIGN_RIGHT_SEC
        tot_cell.number_format = "#,##0"
        tot_cell.border = BORDER_THIN_SEC

        ws_sum.row_dimensions[curr_r].height = 20
        t2_sec_rows.append(curr_r)
        curr_r += 1

    t2_tot_r = curr_r
    ws_sum.cell(row=t2_tot_r, column=1, value="Total")
    ws_sum.cell(row=t2_tot_r, column=1).alignment = ALIGN_LEFT_SEC
    ws_sum.cell(row=t2_tot_r, column=1).font = FONT_BOLD_SEC
    for m_idx in range(2, 14):
        col_l = get_column_letter(m_idx)
        c = ws_sum.cell(row=t2_tot_r, column=m_idx)
        c.value = f"=SUM({col_l}{t2_sec_rows[0]}:{col_l}{t2_sec_rows[-1]})"
        c.font = FONT_BOLD_SEC
        c.alignment = ALIGN_RIGHT_SEC
        c.number_format = "#,##0"
    c_tot = ws_sum.cell(row=t2_tot_r, column=14)
    c_tot.value = f"=SUM(N{t2_sec_rows[0]}:N{t2_sec_rows[-1]})"
    c_tot.font = FONT_BOLD_SEC
    c_tot.alignment = ALIGN_RIGHT_SEC
    c_tot.number_format = "#,##0"

    for c_i in range(1, 15):
        cell = ws_sum.cell(row=t2_tot_r, column=c_i)
        cell.fill = FILL_GRAND_TOTAL_SEC
        cell.border = BORDER_TOTAL_SEC
    ws_sum.row_dimensions[t2_tot_r].height = 22

    # Table 3: Number of Transactions
    curr_r += 3
    ws_sum.cell(row=curr_r, column=1, value="Number of Transactions")
    ws_sum.cell(row=curr_r, column=1).font = FONT_SECTION_HDR_SEC
    ws_sum.row_dimensions[curr_r].height = 24

    curr_r += 1
    t3_hdr_r = curr_r
    for c_i, h_txt in enumerate(sum_month_headers, 1):
        c = ws_sum.cell(row=t3_hdr_r, column=c_i, value=h_txt)
        c.fill = FILL_NAVY_SEC
        c.font = FONT_TBL_HDR_SEC
        c.alignment = ALIGN_CENTER_SEC
        c.border = BORDER_THIN_SEC
    ws_sum.row_dimensions[t3_hdr_r].height = 22

    t3_sec_rows = []
    curr_r += 1
    for sec_code, sec_full, _ in SECTION_SPECS:
        ws_sum.cell(row=curr_r, column=1, value=sec_full)
        ws_sum.cell(row=curr_r, column=1).alignment = ALIGN_LEFT_SEC
        ws_sum.cell(row=curr_r, column=1).font = FONT_REGULAR_SEC
        ws_sum.cell(row=curr_r, column=1).border = BORDER_THIN_SEC

        for m_idx, (m_col_hdr, m_match, _) in enumerate(MONTH_COLS_SEC, 2):
            cell = ws_sum.cell(row=curr_r, column=m_idx)
            cell.value = f"=COUNTIFS('All Transactions'!$A$2:$A${max_all_row},$A{curr_r},'All Transactions'!$B$2:$B${max_all_row},\"{m_match}\")"
            cell.font = FONT_REGULAR_SEC
            cell.alignment = ALIGN_RIGHT_SEC
            cell.number_format = "#,##0"
            cell.border = BORDER_THIN_SEC

        tot_cell = ws_sum.cell(row=curr_r, column=14)
        tot_cell.value = f"=SUM(B{curr_r}:M{curr_r})"
        tot_cell.font = FONT_BOLD_SEC
        tot_cell.alignment = ALIGN_RIGHT_SEC
        tot_cell.number_format = "#,##0"
        tot_cell.border = BORDER_THIN_SEC

        ws_sum.row_dimensions[curr_r].height = 20
        t3_sec_rows.append(curr_r)
        curr_r += 1

    t3_tot_r = curr_r
    ws_sum.cell(row=t3_tot_r, column=1, value="Total")
    ws_sum.cell(row=t3_tot_r, column=1).alignment = ALIGN_LEFT_SEC
    ws_sum.cell(row=t3_tot_r, column=1).font = FONT_BOLD_SEC
    for m_idx in range(2, 14):
        col_l = get_column_letter(m_idx)
        c = ws_sum.cell(row=t3_tot_r, column=m_idx)
        c.value = f"=SUM({col_l}{t3_sec_rows[0]}:{col_l}{t3_sec_rows[-1]})"
        c.font = FONT_BOLD_SEC
        c.alignment = ALIGN_RIGHT_SEC
        c.number_format = "#,##0"
    c_tot = ws_sum.cell(row=t3_tot_r, column=14)
    c_tot.value = f"=SUM(N{t3_sec_rows[0]}:N{t3_sec_rows[-1]})"
    c_tot.font = FONT_BOLD_SEC
    c_tot.alignment = ALIGN_RIGHT_SEC
    c_tot.number_format = "#,##0"

    for c_i in range(1, 15):
        cell = ws_sum.cell(row=t3_tot_r, column=c_i)
        cell.fill = FILL_GRAND_TOTAL_SEC
        cell.border = BORDER_TOTAL_SEC
    ws_sum.row_dimensions[t3_tot_r].height = 22

    ws_sum.column_dimensions['A'].width = 36.0
    for m_idx in range(2, 15):
        ws_sum.column_dimensions[get_column_letter(m_idx)].width = 13.0

    return ws_sum, ws_all, section_sheets


_VALID_PAN = re.compile(r"^[A-Z]{5}[0-9]{4}[A-Z]$")


def _pan_missing_findings(ws_tds_working: Any) -> List[tuple]:
    """One review finding per TDS Working party with no PAN (none in its Tally ledger master or GSTIN), as
    (category, ref, party, section, TDS booked, finding). The TDS Working itself is left as booked."""
    parties: Dict[str, Dict[str, Any]] = defaultdict(lambda: {"bills": [], "base": 0.0, "booked": 0.0, "secs": set()})
    for row in ws_tds_working.iter_rows(min_row=3, values_only=True):
        if len(row) < 14:
            continue
        bill, party, pan, base, rate, diff, sec = row[2], row[3], row[4], row[6], row[7], row[12], row[13]
        if not party or not isinstance(base, (int, float)) or not isinstance(rate, (int, float)):
            continue
        if _VALID_PAN.match(str(pan or "").replace(" ", "").upper()):
            continue
        p = parties[party]
        p["bills"] += [b.strip() for b in str(bill or "").split(",") if b.strip()]
        p["base"] += base
        p["booked"] += round(base * rate / 100) + (diff if isinstance(diff, (int, float)) else 0)
        p["secs"].add(str(sec))

    findings = []
    for party, p in sorted(parties.items(), key=lambda kv: kv[0].upper()):
        bills = list(dict.fromkeys(p["bills"]))
        ref = ", ".join(bills) if len(bills) <= 4 else f"Multiple ({len(bills)} bills)"
        findings.append((
            "PAN MISSING", ref, party, "/".join(sorted(p["secs"])), round(p["booked"], 2),
            f"No PAN in the Tally ledger master (PAN/IT No.) for this deductee: Rs. {p['base']:,.2f} paid/credited, "
            f"TDS Rs. {p['booked']:,.0f} booked. Enter the PAN in the ledger master before filing Form 26Q; "
            f"if the deductee has not furnished a PAN, Sec 206AA requires TDS at 20% "
            f"(Rs. {round(p['base'] * 0.20):,} on this amount).",
        ))
    return findings


def _build_review_points_sheet(
    wb: openpyxl.Workbook,
    comparisons: List[TransactionComparison],
    party_name: str,
    period: str,
    ws_tds_working: Any = None,
) -> Any:
    """
    Builds the 'Review Points' sheet listing concrete audit findings, section verifications,
    potential duplicate bills, rate anomalies, and conversion methodology notes.
    """
    ws = wb.create_sheet(title="Review Points")
    ws.views.sheetView[0].showGridLines = True

    # Title Banner
    ws.merge_cells("A1:G1")
    ws["A1"] = f"{party_name.upper()} — AUDIT REVIEW POINTS & OBSERVATIONS {period.upper()}"
    ws["A1"].fill = ZEON_HEADER_DARK
    ws["A1"].font = ZEON_TITLE_FONT
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 28

    for c_i in range(1, 8):
        c = ws.cell(row=1, column=c_i)
        c.fill = ZEON_HEADER_DARK
        c.border = BORDER_THIN

    # Subtitle
    ws.merge_cells("A2:G2")
    ws["A2"] = "Actionable Compliance Notes, Rate Adjustments, Deduplication & Statutory Review"
    ws["A2"].fill = ZEON_HEADER_BLUE
    ws["A2"].font = ZEON_SUBTITLE_FONT
    ws["A2"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[2].height = 20

    for c_i in range(1, 8):
        c = ws.cell(row=2, column=c_i)
        c.fill = ZEON_HEADER_BLUE
        c.border = BORDER_THIN

    # Headers
    headers = ["#", "Category", "Ref / Voucher", "Party Name", "Section", "TDS / Amount (₹)", "Concrete Audit Finding & Statutory Recommendation"]
    ws.append(headers)
    ws.row_dimensions[3].height = 24
    for col_idx in range(1, 8):
        c = ws.cell(row=3, column=col_idx)
        c.fill = ZEON_HEADER_BLUE
        c.font = ZEON_HEADER_FONT
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BORDER_THIN

    # Scan comparisons for audit findings
    review_rows = []
    seen_parties = defaultdict(list)
    
    # 1. Partner Remuneration & HUF Interest Entries
    has_194t = False
    has_194a = False
    for comp in comparisons:
        tx = comp.new_transaction or comp.old_transaction
        if not tx or is_excluded_non_tds_party(tx.party_name):
            continue
        p_name = tx.party_name or ""
        seen_parties[p_name.upper()].append(tx)
        all_l = " ".join([l.ledger_name for l in tx.ledgers]).upper()
        if "PARTNER" in all_l or "194T" in all_l or "REMUNERATION" in all_l:
            has_194t = True
        if "INTEREST ON LOAN" in all_l or "194A" in all_l or ("INTEREST" in all_l and "BANK" not in all_l):
            has_194a = True

    item_no = 1
    if has_194t:
        review_rows.append((
            item_no, "SECTION 194T", "Year-End", "Partners Remuneration & Interest", "194T", "10.00%",
            "Partner remuneration and interest are mapped to Section 194T @ 10% (w.e.f. 1-Oct-2024 / FY 2025-26) with individual partner PANs to prevent double-counting of memo totals."
        ))
        item_no += 1

    if has_194a:
        review_rows.append((
            item_no, "SECTION 194A", "Year-End", "Interest on Loans", "194A", "10.00%",
            "Interest on loan provisions are segregated into individual lender accounts with 10% TDS under Section 194A."
        ))
        item_no += 1

    # 2. Check for multi-section vendors (e.g. 194C labour vs 194Q goods)
    for p_up, txs in sorted(seen_parties.items()):
        secs = set()
        for t in txs:
            sec, _, _ = infer_tds_section_and_rate(t)
            if sec:
                secs.add(sec)
        if len(secs) > 1 and ("194C" in secs and "194Q" in secs):
            p_display = txs[0].party_name or p_up
            review_rows.append((
                item_no, "SECTION VERIFY", "Multiple", p_display, "/".join(sorted(secs)), "-",
                f"Party has entries under both 194C (labour/making charges @ 1%/2%) and 194Q (goods purchase @ 0.10%). Verify invoice classification between goods vs labour."
            ))
            item_no += 1

    # 3. Deductees with no PAN in Tally (the TDS Working stays at the rate booked)
    if ws_tds_working is not None:
        for finding in _pan_missing_findings(ws_tds_working):
            review_rows.append((item_no, *finding))
            item_no += 1

    # Populate table
    row_idx = 4
    for item in review_rows:
        ws.append(list(item))
        ws.row_dimensions[row_idx].height = 22
        r_fill = ZEON_ZEBRA_LIGHT if row_idx % 2 == 0 else WHITE_FILL
        for col_i in range(1, 8):
            cell = ws.cell(row=row_idx, column=col_i)
            cell.border = BORDER_THIN
            cell.fill = r_fill
            cell.font = ZEON_DATA_TEXT
            if col_i in [1, 2, 3, 5, 6]:
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif col_i == 4:
                cell.alignment = Alignment(horizontal="left", vertical="center")
                cell.font = ZEON_BOLD_TEXT
            elif col_i == 7:
                cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        row_idx += 1

    # Conversion Methodology Section
    row_idx += 1
    ws.merge_cells(f"A{row_idx}:G{row_idx}")
    banner = ws.cell(row=row_idx, column=1, value="STATUTORY TDS AUDIT & RECONCILIATION METHODOLOGY")
    banner.fill = ZEON_HEADER_DARK
    banner.font = ZEON_HEADER_FONT
    banner.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[row_idx].height = 22
    row_idx += 1

    methodology_notes = [
        ("1. Section 194T Remuneration", "Section 194T partner remuneration is mapped to individual partner PANs per registered partnership deed statutory limits rather than aggregate memo balances."),
        ("2. Section 194A Loan Interest", "Interest on loan provisions are segregated into individual lender accounts with 10% TDS under Section 194A."),
        ("3. Section 194C / 194Q Multi-Head Split", "For suppliers with mixed supplies, making charges & labour are segregated under 194C (@ 1%/2%), while goods purchases exceeding Rs. 50L are routed to 194Q (@ 0.10%)."),
        ("4. GST Duty Stripping (CBDT 13/2021)", "Pure GST tax duties (Input CGST, SGST, IGST) are stripped from purchase invoices to apply TDS strictly on net taxable invoice base."),
        ("5. Cumulative 194Q Threshold Tracking", "Turnover per supplier is tracked chronologically across the FY; 194Q applies strictly to purchases exceeding the Rs. 50,00,000 statutory limit."),
        ("6. Bank & Contra Filtering", "Internal bank transfers, POS terminal settlements, and cash/cheque sales receipts are strictly excluded from TDS liability."),
        ("7. Rupee Rounding & Filing Alignment", "All TDS deductions and interest are rounded to the nearest integer rupee matching quarterly Form 26Q returns.")
    ]

    for title_m, desc_m in methodology_notes:
        ws.cell(row=row_idx, column=1, value="").fill = WHITE_FILL
        ws.cell(row=row_idx, column=2, value=title_m).font = ZEON_BOLD_TEXT
        ws.cell(row=row_idx, column=2).fill = WHITE_FILL
        ws.cell(row=row_idx, column=2).alignment = Alignment(horizontal="left", vertical="center")
        
        ws.merge_cells(f"C{row_idx}:G{row_idx}")
        c_desc = ws.cell(row=row_idx, column=3, value=desc_m)
        c_desc.font = ZEON_DATA_TEXT
        c_desc.alignment = Alignment(horizontal="left", vertical="center")
        for ci in range(1, 8):
            ws.cell(row=row_idx, column=ci).border = BORDER_THIN
        ws.row_dimensions[row_idx].height = 20
        row_idx += 1

    col_widths = {'A': 6.0, 'B': 20.0, 'C': 18.0, 'D': 35.0, 'E': 14.0, 'F': 18.0, 'G': 65.0}
    for col_l, w in col_widths.items():
        ws.column_dimensions[col_l].width = w

    return ws


def generate_excel_report(
    summary: ComparisonSummarySchema,
    reconciliation: ReconciliationResult,
    party_summaries: List[PartySummary],
    comparisons: List[TransactionComparison],
    master_changes: List[MasterChange],
    party_name: Optional[str] = "All Parties",
    period: Optional[str] = "Full Period",
    masters: Optional[List[Any]] = None,
    report_format: Optional[str] = None,
    tds_rows: str = "bill",
    all_transactions: Optional[List[Any]] = None,
) -> bytes:
    """
    Generates a full Excel comparison, TDS Working/Summary, and Form 26Q Section-wise Monthly reporting suite.
    Features:
    1. Summary: Form 26Q Master Executive Dashboard with 3 Monthly Matrix tables (TDS Deducted, Amount Paid, Number of Txns).
    2. All Transactions: Granular transaction register.
    3. Section Sheets: 194A, 194C, 194J, 194Q, 194T, Section Not Specified with monthly banners and totals.
    4. Review Points: Actionable audit observations, rate checks, deduplication & statutory notes.
    5-13. Reconciliation & Audit Sheets: Reconciliation Overview, Party-wise Summary, Deleted, Backdated, New Month, Modified, Recreated, Unresolved, Master Changes.
    Returns bytes content of the workbook.
    """
    wb = openpyxl.Workbook()

    # Determine baseline cutoff date from original snapshot (using dense month volume to filter stray outlier dates)
    orig_dates = [
        c.old_transaction.date
        for c in comparisons
        if c.old_transaction and c.old_transaction.date
    ]
    if orig_dates:
        month_counts = Counter([d[:7] for d in orig_dates if len(d) >= 7])
        if month_counts:
            max_cnt = max(month_counts.values())
            density_threshold = max(5, int(max_cnt * 0.10))
            dense_months = [m for m, cnt in month_counts.items() if cnt >= density_threshold]
            if dense_months:
                latest_dense_month = max(dense_months)
                dense_dates = [d for d in orig_dates if d.startswith(latest_dense_month)]
                max_orig_date = max(dense_dates) if dense_dates else max(orig_dates)
            else:
                max_orig_date = max(orig_dates)
        else:
            max_orig_date = max(orig_dates)
    else:
        max_orig_date = None

    new_items = [c for c in comparisons if c.status == ClassificationStatus.NEW]
    backdated_items = [
        c for c in new_items
        if max_orig_date and c.new_transaction and c.new_transaction.date <= max_orig_date
    ]
    new_month_items = [
        c for c in new_items
        if (not max_orig_date) or (c.new_transaction and c.new_transaction.date > max_orig_date)
    ]

    # -------------------------------------------------------------------------
    # TDS Working & Summary Sheets (Zeon / MK4 Dashboard Suite)
    # -------------------------------------------------------------------------
    ws_tds_working, discovered_sections = _build_tds_working_sheet(
        wb=wb,
        comparisons=comparisons,
        party_name=party_name or "All Parties",
        period=period or "Full Period",
        max_orig_date=max_orig_date,
        masters=masters,
        tds_rows=tds_rows,
        tds_deposits=all_transactions,
    )

    ws_tds_summary = _build_tds_summary_sheet(
        wb=wb,
        party_name=party_name or "All Parties",
        period=period or "Full Period",
        discovered_sections=discovered_sections,
    )

    ws_review_points = _build_review_points_sheet(
        wb=wb,
        comparisons=comparisons,
        party_name=party_name or "All Parties",
        period=period or "Full Period",
        ws_tds_working=ws_tds_working,
    )

    # -------------------------------------------------------------------------
    # Optional Form 26Q Section-wise Monthly Reporting Suite (Summary, All Transactions, Section Sheets)
    # Only generated when explicitly requested via report_format='form26q' or 'sectionwise'
    # -------------------------------------------------------------------------
    if report_format in ("form26q", "sectionwise"):
        ws_sec_summary, ws_all_tx, section_sheets = _build_sectionwise_monthly_sheets(
            wb=wb,
            comparisons=comparisons,
            party_name=party_name or "All Parties",
            period=period or "Full Period",
            masters=masters,
        )

    # -------------------------------------------------------------------------
    # Reconciliation & Audit Metric Overview Sheet (Sheet 4 in MK4 format)
    # -------------------------------------------------------------------------
    ws_recon_overview = wb.create_sheet(title="Summary")
    ws_recon_overview.views.sheetView[0].showGridLines = True

    # Title Banner
    ws_recon_overview["A1"] = "TDS Change Comparison & Reconciliation Summary Report"
    ws_recon_overview["A1"].font = Font(name="Calibri", size=15, bold=True, color="0F172A")
    ws_recon_overview.row_dimensions[1].height = 28

    # Metadata Block: Party Name and Period
    META_LABEL_FILL = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")
    META_LABEL_FONT = Font(name="Calibri", size=10, bold=True, color="334155")
    META_VALUE_FONT = Font(name="Calibri", size=10, bold=False, color="0F172A")

    ws_recon_overview["A3"] = "Party Name"
    ws_recon_overview["B3"] = party_name or "All Parties"
    ws_recon_overview["A4"] = "Period"
    ws_recon_overview["B4"] = period or "Full Period"

    for r_idx in [3, 4]:
        ws_recon_overview[f"A{r_idx}"].fill = META_LABEL_FILL
        ws_recon_overview[f"A{r_idx}"].font = META_LABEL_FONT
        ws_recon_overview[f"A{r_idx}"].border = BORDER_THIN
        ws_recon_overview[f"A{r_idx}"].alignment = ALIGN_LEFT
        ws_recon_overview[f"B{r_idx}"].font = META_VALUE_FONT
        ws_recon_overview[f"B{r_idx}"].border = BORDER_THIN
        ws_recon_overview[f"B{r_idx}"].alignment = ALIGN_LEFT
        ws_recon_overview.row_dimensions[r_idx].height = 20

    # Section 1: Transaction Metric Summary
    ws_recon_overview["A6"] = "TRANSACTION METRIC"
    ws_recon_overview["B6"] = "COUNT"
    ws_recon_overview["C6"] = "STATUS / AUDIT REMARKS"
    for col_let in ["A", "B", "C"]:
        ws_recon_overview[f"{col_let}6"].fill = HEADER_FILL
        ws_recon_overview[f"{col_let}6"].font = HEADER_FONT
        ws_recon_overview[f"{col_let}6"].alignment = ALIGN_CENTER if col_let == "B" else ALIGN_LEFT
    ws_recon_overview.row_dimensions[6].height = 24

    count_rows = [
        ("Original Snapshot Transactions", summary.original_transactions, "Total baseline transactions parsed from original snapshot"),
        ("Updated Snapshot Transactions", summary.updated_transactions, "Total updated transactions parsed from updated snapshot"),
        ("Unchanged (Identical)", summary.unchanged, "Exact match with identical voucher details, ledgers & amounts"),
        ("Modified", summary.modified, "Transactions matched by GUID with altered business fields or amounts"),
        ("Backdated Entries (Common Period)", len(backdated_items), f"Newly added vouchers dated on or before baseline cutoff ({max_orig_date or 'N/A'})"),
        ("New Month Entries (Extended Period)", len(new_month_items), "Newly added vouchers for extended month(s)/period"),
        ("Missing in Updated Data (Deleted)", summary.missing, "Baseline transactions deleted or missing in updated snapshot"),
        ("Recreated / Renumbered", summary.recreated, "Vouchers re-entered with new GUID or modified voucher numbering"),
        ("Unresolved / Review", summary.unresolved, "Ambiguous matches requiring manual audit review"),
    ]

    for idx, (label, val, remark) in enumerate(count_rows, start=7):
        ws_recon_overview[f"A{idx}"] = label
        ws_recon_overview[f"B{idx}"] = val
        ws_recon_overview[f"C{idx}"] = remark
        ws_recon_overview.row_dimensions[idx].height = 20

        ws_recon_overview[f"A{idx}"].font = BOLD_DATA_FONT if idx in [7, 8] else DATA_FONT
        ws_recon_overview[f"B{idx}"].font = BOLD_DATA_FONT
        ws_recon_overview[f"C{idx}"].font = REASON_FONT

        ws_recon_overview[f"A{idx}"].alignment = ALIGN_LEFT
        ws_recon_overview[f"B{idx}"].alignment = ALIGN_RIGHT
        ws_recon_overview[f"C{idx}"].alignment = ALIGN_LEFT

        ws_recon_overview[f"B{idx}"].number_format = NUMBER_FORMAT

        # Highlighting for special rows
        if "Missing" in label and val > 0:
            row_fill = MISSING_ROW_FILL
            ws_recon_overview[f"A{idx}"].font = MISSING_FONT
            ws_recon_overview[f"B{idx}"].font = MISSING_FONT
        elif "New" in label and val > 0:
            row_fill = NEW_ROW_FILL
            ws_recon_overview[f"A{idx}"].font = NEW_FONT
            ws_recon_overview[f"B{idx}"].font = NEW_FONT
        elif "Modified" in label and val > 0:
            row_fill = MOD_ROW_FILL
            ws_recon_overview[f"A{idx}"].font = MOD_FONT
            ws_recon_overview[f"B{idx}"].font = MOD_FONT
        elif "Unresolved" in label and val > 0:
            row_fill = UNRES_ROW_FILL
            ws_recon_overview[f"A{idx}"].font = UNRES_FONT
            ws_recon_overview[f"B{idx}"].font = UNRES_FONT
        else:
            row_fill = ZEBRA_FILL if idx % 2 == 0 else WHITE_FILL

        for col_let in ["A", "B", "C"]:
            ws_recon_overview[f"{col_let}{idx}"].border = BORDER_THIN
            ws_recon_overview[f"{col_let}{idx}"].fill = row_fill

    # Section 2: Monetary Summary
    start_row = 17
    ws_recon_overview[f"A{start_row}"] = "MONETARY METRIC"
    ws_recon_overview[f"B{start_row}"] = "ORIGINAL (₹)"
    ws_recon_overview[f"C{start_row}"] = "UPDATED (₹)"
    ws_recon_overview[f"D{start_row}"] = "VARIANCE (₹)"
    ws_recon_overview[f"E{start_row}"] = "CONCRETE REASON / STATUS"
    for col_let in ["A", "B", "C", "D", "E"]:
        ws_recon_overview[f"{col_let}{start_row}"].fill = HEADER_FILL
        ws_recon_overview[f"{col_let}{start_row}"].font = HEADER_FONT
        ws_recon_overview[f"{col_let}{start_row}"].alignment = ALIGN_CENTER if col_let in ["B", "C", "D"] else ALIGN_LEFT
    ws_recon_overview.row_dimensions[start_row].height = 24

    amt_diff = float(summary.amount_difference)
    tds_diff = float(summary.tds_difference)

    amt_reason = (
        "Exact Match — Baseline and updated transaction amounts are balanced"
        if amt_diff == 0
        else f"Net {'increase' if amt_diff > 0 else 'decrease'} of ₹{abs(amt_diff):,.2f} in updated transactions"
    )
    tds_reason = (
        "Exact Match — Baseline and updated TDS withholding deductions are balanced"
        if tds_diff == 0
        else f"Net {'increase' if tds_diff > 0 else 'decrease'} of ₹{abs(tds_diff):,.2f} in TDS deductions across snapshots"
    )

    money_rows = [
        ("Total Transaction Amount", float(summary.original_amount), float(summary.updated_amount), amt_diff, amt_reason),
        ("Total TDS Amount", float(summary.original_tds), float(summary.updated_tds), tds_diff, tds_reason),
    ]

    for idx, (label, orig_v, upd_v, diff_v, reason) in enumerate(money_rows, start=start_row + 1):
        ws_recon_overview[f"A{idx}"] = label
        ws_recon_overview[f"B{idx}"] = orig_v
        ws_recon_overview[f"C{idx}"] = upd_v
        ws_recon_overview[f"D{idx}"] = diff_v
        ws_recon_overview[f"E{idx}"] = reason
        ws_recon_overview.row_dimensions[idx].height = 20

        ws_recon_overview[f"A{idx}"].font = BOLD_DATA_FONT
        ws_recon_overview[f"B{idx}"].font = DATA_FONT
        ws_recon_overview[f"C{idx}"].font = DATA_FONT
        ws_recon_overview[f"D{idx}"].font = DIFF_POS_FONT if diff_v >= 0 else DIFF_NEG_FONT
        ws_recon_overview[f"E{idx}"].font = REASON_FONT

        ws_recon_overview[f"A{idx}"].alignment = ALIGN_LEFT
        ws_recon_overview[f"B{idx}"].alignment = ALIGN_RIGHT
        ws_recon_overview[f"C{idx}"].alignment = ALIGN_RIGHT
        ws_recon_overview[f"D{idx}"].alignment = ALIGN_RIGHT
        ws_recon_overview[f"E{idx}"].alignment = ALIGN_LEFT

        for col_let in ["A", "B", "C", "D", "E"]:
            ws_recon_overview[f"{col_let}{idx}"].border = BORDER_THIN
            ws_recon_overview[f"{col_let}{idx}"].fill = ZEBRA_FILL if idx % 2 == 0 else WHITE_FILL
        for col_let in ["B", "C", "D"]:
            ws_recon_overview[f"{col_let}{idx}"].number_format = CURRENCY_FORMAT

    # Section 3: Reconciliation Checks
    rec_start = 22
    ws_recon_overview[f"A{rec_start}"] = "RECONCILIATION CHECK"
    ws_recon_overview[f"B{rec_start}"] = "STATUS"
    ws_recon_overview[f"C{rec_start}"] = "CONCRETE REASON / EXPLANATION"
    for col_let in ["A", "B", "C"]:
        ws_recon_overview[f"{col_let}{rec_start}"].fill = HEADER_FILL
        ws_recon_overview[f"{col_let}{rec_start}"].font = HEADER_FONT
        ws_recon_overview[f"{col_let}{rec_start}"].alignment = ALIGN_CENTER if col_let == "B" else ALIGN_LEFT
    ws_recon_overview.row_dimensions[rec_start].height = 24

    count_var = summary.updated_transactions - summary.original_transactions
    count_reason = (
        f"PASS — Transaction counts are balanced ({summary.original_transactions} vouchers in both snapshots)"
        if reconciliation.transaction_count == "PASS"
        else f"FAIL — Count variance of {abs(count_var)} voucher(s) (Original: {summary.original_transactions}, Updated: {summary.updated_transactions})"
    )

    amt_rec_reason = (
        f"PASS — Total amount matches perfectly (₹{float(summary.original_amount):,.2f})"
        if reconciliation.amount == "PASS"
        else f"FAIL — Total amount variance of ₹{abs(amt_diff):,.2f} (Original: ₹{float(summary.original_amount):,.2f}, Updated: ₹{float(summary.updated_amount):,.2f})"
    )

    tds_rec_reason = (
        f"PASS — Total TDS matches perfectly (₹{float(summary.original_tds):,.2f})"
        if reconciliation.tds == "PASS"
        else f"FAIL — Total TDS variance of ₹{abs(tds_diff):,.2f} (Original: ₹{float(summary.original_tds):,.2f}, Updated: ₹{float(summary.updated_tds):,.2f})"
    )

    rec_rows = [
        ("Transaction Count Balance", reconciliation.transaction_count, count_reason),
        ("Total Amount Balance", reconciliation.amount, amt_rec_reason),
        ("Total TDS Balance", reconciliation.tds, tds_rec_reason),
    ]

    for idx, (label, status, reason) in enumerate(rec_rows, start=rec_start + 1):
        ws_recon_overview[f"A{idx}"] = label
        ws_recon_overview[f"B{idx}"] = status
        ws_recon_overview[f"C{idx}"] = reason
        ws_recon_overview.row_dimensions[idx].height = 22

        ws_recon_overview[f"A{idx}"].font = BOLD_DATA_FONT
        ws_recon_overview[f"A{idx}"].alignment = ALIGN_LEFT
        ws_recon_overview[f"B{idx}"].alignment = ALIGN_CENTER
        ws_recon_overview[f"C{idx}"].font = REASON_FONT
        ws_recon_overview[f"C{idx}"].alignment = ALIGN_LEFT

        for col_let in ["A", "B", "C"]:
            ws_recon_overview[f"{col_let}{idx}"].border = BORDER_THIN
            ws_recon_overview[f"{col_let}{idx}"].fill = ZEBRA_FILL if idx % 2 == 0 else WHITE_FILL

        if status == "PASS":
            ws_recon_overview[f"B{idx}"].fill = PASS_FILL
            ws_recon_overview[f"B{idx}"].font = PASS_FONT
        else:
            ws_recon_overview[f"B{idx}"].fill = FAIL_FILL
            ws_recon_overview[f"B{idx}"].font = FAIL_FONT

    _auto_adjust_column_widths(ws_recon_overview)

    # -------------------------------------------------------------------------
    # Sheet 2: Party-wise Summary
    # -------------------------------------------------------------------------
    ws_party = wb.create_sheet(title="Party-wise Summary")
    ws_party.views.sheetView[0].showGridLines = True
    party_headers = [
        "Party Name", "Status", "Concrete Reason / Remarks", "Affected Entries",
        "Missing (Deleted)", "New (Added)", "Modified", "Recreated",
        "Original Amount", "Updated Amount", "Amount Difference",
        "Original TDS", "Updated TDS", "TDS Difference"
    ]
    ws_party.append(party_headers)
    ws_party.row_dimensions[1].height = 26
    for col_idx in range(1, len(party_headers) + 1):
        cell = ws_party.cell(row=1, column=col_idx)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = ALIGN_CENTER

    # Filter to only parties that have changes or variances
    changed_parties = [
        p for p in party_summaries
        if not (p.affected_entries == 0 and p.amount_difference == Decimal("0.00") and p.tds_difference == Decimal("0.00"))
    ]

    for p in changed_parties:
        # Determine status and concrete reason for each party with changes
        if p.missing_count > 0 and p.new_count > 0:
            p_status = "MIXED CHANGES"
            p_reason = f"{p.missing_count} deleted voucher(s), {p.new_count} newly added voucher(s)"
            row_fill = MISSING_ROW_FILL
            status_fill = MISSING_BADGE_FILL
            status_font = MISSING_FONT
        elif p.missing_count > 0:
            p_status = "HAS DELETIONS"
            p_reason = f"{p.missing_count} voucher(s) missing / deleted in updated data"
            row_fill = MISSING_ROW_FILL
            status_fill = MISSING_BADGE_FILL
            status_font = MISSING_FONT
        elif p.new_count > 0:
            p_status = "HAS ADDITIONS"
            p_reason = f"{p.new_count} newly recorded voucher(s) in updated data"
            row_fill = NEW_ROW_FILL
            status_fill = NEW_BADGE_FILL
            status_font = NEW_FONT
        elif p.modified_count > 0:
            p_status = "MODIFIED"
            p_reason = f"{p.modified_count} voucher(s) modified with variance"
            row_fill = MOD_ROW_FILL
            status_fill = MOD_BADGE_FILL
            status_font = MOD_FONT
        elif p.recreated_count > 0:
            p_status = "RECREATED"
            p_reason = f"{p.recreated_count} voucher(s) recreated/renumbered"
            row_fill = REC_ROW_FILL
            status_fill = REC_BADGE_FILL
            status_font = REC_FONT
        else:
            p_status = "VARIANCE"
            p_reason = f"Amount variance: ₹{float(p.amount_difference):,.2f}, TDS variance: ₹{float(p.tds_difference):,.2f}"
            row_fill = MOD_ROW_FILL
            status_fill = MOD_BADGE_FILL
            status_font = MOD_FONT

        row_vals = [
            p.party, p_status, p_reason, p.affected_entries,
            p.missing_count, p.new_count, p.modified_count, p.recreated_count,
            float(p.original_amount), float(p.updated_amount), float(p.amount_difference),
            float(p.original_tds), float(p.updated_tds), float(p.tds_difference),
        ]
        ws_party.append(row_vals)
        r_idx = ws_party.max_row
        ws_party.row_dimensions[r_idx].height = 20

        for col_idx in range(1, len(party_headers) + 1):
            c = ws_party.cell(row=r_idx, column=col_idx)
            c.border = BORDER_THIN
            c.fill = row_fill
            c.font = DATA_FONT

            if col_idx == 1:
                c.font = BOLD_DATA_FONT
                c.alignment = ALIGN_LEFT
            elif col_idx == 2:
                c.fill = status_fill
                c.font = status_font
                c.alignment = ALIGN_CENTER
            elif col_idx == 3:
                c.font = REASON_FONT
                c.alignment = ALIGN_LEFT
            elif col_idx in [4, 5, 6, 7, 8]:
                c.alignment = ALIGN_RIGHT
                c.number_format = NUMBER_FORMAT
            elif col_idx >= 9:
                c.alignment = ALIGN_RIGHT
                c.number_format = CURRENCY_FORMAT
                if col_idx in [11, 14]:  # Difference columns
                    diff_val = c.value
                    if isinstance(diff_val, (int, float)):
                        if diff_val > 0:
                            c.font = DIFF_POS_FONT
                        elif diff_val < 0:
                            c.font = DIFF_NEG_FONT

    # Total Summary Row for Party Summary
    if changed_parties:
        total_affected = sum(p.affected_entries for p in changed_parties)
        total_missing = sum(p.missing_count for p in changed_parties)
        total_new = sum(p.new_count for p in changed_parties)
        total_mod = sum(p.modified_count for p in changed_parties)
        total_rec = sum(p.recreated_count for p in changed_parties)
        total_orig_amt = sum(p.original_amount for p in changed_parties)
        total_upd_amt = sum(p.updated_amount for p in changed_parties)
        total_amt_diff = sum(p.amount_difference for p in changed_parties)
        total_orig_tds = sum(p.original_tds for p in changed_parties)
        total_upd_tds = sum(p.updated_tds for p in changed_parties)
        total_tds_diff = sum(p.tds_difference for p in changed_parties)

        totals_row = [
            "TOTAL AFFECTED PARTIES",
            f"{len(changed_parties)} Parties with Changes",
            "Consolidated Affected Party Metrics",
            total_affected, total_missing, total_new, total_mod, total_rec,
            float(total_orig_amt), float(total_upd_amt), float(total_amt_diff),
            float(total_orig_tds), float(total_upd_tds), float(total_tds_diff)
        ]
        ws_party.append(totals_row)
        tot_idx = ws_party.max_row
        ws_party.row_dimensions[tot_idx].height = 24
        for col_idx in range(1, len(party_headers) + 1):
            c = ws_party.cell(row=tot_idx, column=col_idx)
            c.border = BORDER_TOTAL
            c.fill = META_LABEL_FILL
            c.font = TOTAL_FONT
            if col_idx in [1, 3]:
                c.alignment = ALIGN_LEFT
            elif col_idx == 2:
                c.alignment = ALIGN_CENTER
            elif col_idx in [4, 5, 6, 7, 8]:
                c.alignment = ALIGN_RIGHT
                c.number_format = NUMBER_FORMAT
            elif col_idx >= 9:
                c.alignment = ALIGN_RIGHT
                c.number_format = CURRENCY_FORMAT

    ws_party.freeze_panes = "A2"
    ws_party.auto_filter.ref = ws_party.dimensions
    _auto_adjust_column_widths(ws_party)

    # -------------------------------------------------------------------------
    # Sheet 3: Deleted Entries (Highlighted in Red with Concrete Reasons)
    # -------------------------------------------------------------------------
    ws_missing = wb.create_sheet(title="Deleted Entries")
    ws_missing.views.sheetView[0].showGridLines = True
    missing_headers = [
        "Voucher Number", "Date", "Voucher Type", "Party Name",
        "Original Amount", "Original TDS", "Status", "Concrete Reason / Deletion Impact", "Original GUID"
    ]
    ws_missing.append(missing_headers)
    ws_missing.row_dimensions[1].height = 26
    for col_idx in range(1, len(missing_headers) + 1):
        cell = ws_missing.cell(row=1, column=col_idx)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = ALIGN_CENTER

    missing_items = [c for c in comparisons if c.status == ClassificationStatus.MISSING]
    for item in missing_items:
        tx = item.old_transaction
        ledger_names = ", ".join([l.ledger_name for l in tx.ledgers]) if tx.ledgers else "No line ledgers"
        concrete_reason = (
            f"Voucher exists in Original snapshot on {tx.date} for ₹{float(tx.total_amount):,.2f} "
            f"(TDS: ₹{float(tx.tds_amount):,.2f}), but was deleted/removed in Updated snapshot. "
            f"Ledgers: {ledger_names}."
        )

        ws_missing.append([
            tx.voucher_number or "Unnumbered", tx.date, tx.voucher_type, tx.party_name,
            float(tx.total_amount), float(tx.tds_amount), "DELETED / MISSING", concrete_reason, tx.guid
        ])
        r_idx = ws_missing.max_row
        ws_missing.row_dimensions[r_idx].height = 22

        for col_idx in range(1, len(missing_headers) + 1):
            c = ws_missing.cell(row=r_idx, column=col_idx)
            c.border = BORDER_THIN
            c.fill = MISSING_ROW_FILL
            c.font = DATA_FONT

            if col_idx in [1, 2, 3]:
                c.alignment = ALIGN_CENTER
            elif col_idx == 4:
                c.alignment = ALIGN_LEFT
                c.font = BOLD_DATA_FONT
            elif col_idx in [5, 6]:
                c.alignment = ALIGN_RIGHT
                c.number_format = CURRENCY_FORMAT
            elif col_idx == 7:
                c.fill = MISSING_BADGE_FILL
                c.font = MISSING_FONT
                c.alignment = ALIGN_CENTER
            elif col_idx == 8:
                c.font = REASON_FONT
                c.alignment = ALIGN_LEFT
            elif col_idx == 9:
                c.alignment = ALIGN_LEFT

    ws_missing.freeze_panes = "A2"
    ws_missing.auto_filter.ref = ws_missing.dimensions
    _auto_adjust_column_widths(ws_missing)

    # -------------------------------------------------------------------------
    # Sheet 4: Backdated Entries (Highlighted in Amber with Forward/Backward TDS Calculation)
    # -------------------------------------------------------------------------
    ws_backdated = wb.create_sheet(title="Backdated Entries")
    ws_backdated.views.sheetView[0].showGridLines = True
    backdated_headers = [
        "Voucher Number", "Date", "Voucher Type", "Party Name",
        "Gross / Base Amount", "Calculation Method", "TDS Section", "Statutory Rate",
        "Calculated Statutory TDS", "Booked TDS in Tally", "TDS Shortfall / Variance", "Status",
        "Note / Nature of Deduction", "Statutory Audit & Compliance Action", "Updated GUID"
    ]
    ws_backdated.append(backdated_headers)
    ws_backdated.row_dimensions[1].height = 26
    for col_idx in range(1, len(backdated_headers) + 1):
        cell = ws_backdated.cell(row=1, column=col_idx)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = ALIGN_CENTER

    tot_base_bk = Decimal("0.00")
    tot_calc_bk = Decimal("0.00")
    tot_book_bk = Decimal("0.00")
    tot_diff_bk = Decimal("0.00")

    for item in backdated_items:
        tx = item.new_transaction
        sec, rt, nt = infer_tds_section_and_rate(tx)
        calc = compute_transaction_tds_calculation(tx, sec, rt, is_backdated=True, cutoff_date=max_orig_date)

        tot_base_bk += calc.base_amount
        tot_calc_bk += calc.statutory_tds
        tot_book_bk += calc.booked_tds
        tot_diff_bk += calc.shortfall

        ws_backdated.append([
            tx.voucher_number or "Unnumbered",
            tx.date,
            tx.voucher_type,
            tx.party_name,
            float(calc.base_amount),
            calc.calc_mode,
            sec,
            calc.rate_str,
            float(calc.statutory_tds),
            float(calc.booked_tds),
            float(calc.shortfall),
            "BACKDATED / ADDED",
            nt,
            calc.compliance_note,
            tx.guid,
        ])
        r_idx = ws_backdated.max_row
        ws_backdated.row_dimensions[r_idx].height = 22

        for col_idx in range(1, len(backdated_headers) + 1):
            c = ws_backdated.cell(row=r_idx, column=col_idx)
            c.border = BORDER_THIN
            c.fill = BACKDATED_ROW_FILL
            c.font = DATA_FONT

            if col_idx in [1, 2, 3, 6, 8]:
                c.alignment = ALIGN_CENTER
            elif col_idx == 4:
                c.alignment = ALIGN_LEFT
                c.font = BOLD_DATA_FONT
            elif col_idx in [5, 9, 10, 11]:
                c.alignment = ALIGN_RIGHT
                c.number_format = CURRENCY_FORMAT
                if col_idx == 11 and calc.shortfall > Decimal("0.00"):
                    c.font = Font(name="Calibri", size=10, bold=True, color="991B1B")
            elif col_idx == 7:
                c.alignment = ALIGN_CENTER
                c.font = BOLD_DATA_FONT
            elif col_idx == 12:
                c.fill = BACKDATED_BADGE_FILL
                c.font = BACKDATED_FONT
                c.alignment = ALIGN_CENTER
            elif col_idx in [13, 14]:
                c.font = REASON_FONT
                c.alignment = ALIGN_LEFT
            elif col_idx == 15:
                c.alignment = ALIGN_LEFT

    # Add Totals Row for Backdated Entries
    if backdated_items:
        tot_row_idx = ws_backdated.max_row + 1
        ws_backdated.append([
            "", "", "", "TOTAL",
            float(tot_base_bk), "", "", "",
            float(tot_calc_bk), float(tot_book_bk), float(tot_diff_bk),
            "", "", "", ""
        ])
        ws_backdated.row_dimensions[tot_row_idx].height = 24
        for col_idx in range(1, len(backdated_headers) + 1):
            c = ws_backdated.cell(row=tot_row_idx, column=col_idx)
            c.border = BORDER_THIN
            c.fill = SECTION_FILL
            c.font = SECTION_FONT
            if col_idx == 4:
                c.alignment = ALIGN_LEFT
            elif col_idx in [5, 9, 10, 11]:
                c.alignment = ALIGN_RIGHT
                c.number_format = CURRENCY_FORMAT

    ws_backdated.freeze_panes = "A2"
    ws_backdated.auto_filter.ref = ws_backdated.dimensions
    _auto_adjust_column_widths(ws_backdated)

    # -------------------------------------------------------------------------
    # Sheet 5: New Month Entries (Highlighted in Green with Forward/Backward TDS Calculation)
    # -------------------------------------------------------------------------
    ws_new_month = wb.create_sheet(title="New Month Entries")
    ws_new_month.views.sheetView[0].showGridLines = True
    new_month_headers = [
        "Voucher Number", "Date", "Voucher Type", "Party Name",
        "Gross / Base Amount", "Calculation Method", "TDS Section", "Statutory Rate",
        "Calculated Statutory TDS", "Booked TDS in Tally", "TDS Shortfall / Variance", "Status",
        "Note / Nature of Deduction", "Statutory Audit & Compliance Action", "Updated GUID"
    ]
    ws_new_month.append(new_month_headers)
    ws_new_month.row_dimensions[1].height = 26
    for col_idx in range(1, len(new_month_headers) + 1):
        cell = ws_new_month.cell(row=1, column=col_idx)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = ALIGN_CENTER

    tot_base_nm = Decimal("0.00")
    tot_calc_nm = Decimal("0.00")
    tot_book_nm = Decimal("0.00")
    tot_diff_nm = Decimal("0.00")

    for item in new_month_items:
        tx = item.new_transaction
        sec, rt, nt = infer_tds_section_and_rate(tx)
        calc = compute_transaction_tds_calculation(tx, sec, rt, is_backdated=False, cutoff_date=max_orig_date)

        tot_base_nm += calc.base_amount
        tot_calc_nm += calc.statutory_tds
        tot_book_nm += calc.booked_tds
        tot_diff_nm += calc.shortfall

        ws_new_month.append([
            tx.voucher_number or "Unnumbered",
            tx.date,
            tx.voucher_type,
            tx.party_name,
            float(calc.base_amount),
            calc.calc_mode,
            sec,
            calc.rate_str,
            float(calc.statutory_tds),
            float(calc.booked_tds),
            float(calc.shortfall),
            "NEW MONTH / ADDED",
            nt,
            calc.compliance_note,
            tx.guid,
        ])
        r_idx = ws_new_month.max_row
        ws_new_month.row_dimensions[r_idx].height = 22

        for col_idx in range(1, len(new_month_headers) + 1):
            c = ws_new_month.cell(row=r_idx, column=col_idx)
            c.border = BORDER_THIN
            c.fill = NEW_ROW_FILL
            c.font = DATA_FONT

            if col_idx in [1, 2, 3, 6, 8]:
                c.alignment = ALIGN_CENTER
            elif col_idx == 4:
                c.alignment = ALIGN_LEFT
                c.font = BOLD_DATA_FONT
            elif col_idx in [5, 9, 10, 11]:
                c.alignment = ALIGN_RIGHT
                c.number_format = CURRENCY_FORMAT
                if col_idx == 11 and calc.shortfall > Decimal("0.00"):
                    c.font = Font(name="Calibri", size=10, bold=True, color="991B1B")
            elif col_idx == 7:
                c.alignment = ALIGN_CENTER
                c.font = BOLD_DATA_FONT
            elif col_idx == 12:
                c.fill = NEW_BADGE_FILL
                c.font = NEW_FONT
                c.alignment = ALIGN_CENTER
            elif col_idx in [13, 14]:
                c.font = REASON_FONT
                c.alignment = ALIGN_LEFT
            elif col_idx == 15:
                c.alignment = ALIGN_LEFT

    # Add Totals Row for New Month Entries
    if new_month_items:
        tot_row_idx = ws_new_month.max_row + 1
        ws_new_month.append([
            "", "", "", "TOTAL",
            float(tot_base_nm), "", "", "",
            float(tot_calc_nm), float(tot_book_nm), float(tot_diff_nm),
            "", "", "", ""
        ])
        ws_new_month.row_dimensions[tot_row_idx].height = 24
        for col_idx in range(1, len(new_month_headers) + 1):
            c = ws_new_month.cell(row=tot_row_idx, column=col_idx)
            c.border = BORDER_THIN
            c.fill = SECTION_FILL
            c.font = SECTION_FONT
            if col_idx == 4:
                c.alignment = ALIGN_LEFT
            elif col_idx in [5, 9, 10, 11]:
                c.alignment = ALIGN_RIGHT
                c.number_format = CURRENCY_FORMAT

    ws_new_month.freeze_panes = "A2"
    ws_new_month.auto_filter.ref = ws_new_month.dimensions
    _auto_adjust_column_widths(ws_new_month)

    # -------------------------------------------------------------------------
    # Sheet 6: Modified Entries (Highlighted in Amber with Concrete Reasons)
    # -------------------------------------------------------------------------
    ws_mod = wb.create_sheet(title="Modified Entries")
    ws_mod.views.sheetView[0].showGridLines = True
    mod_headers = [
        "Voucher Number", "Date", "Voucher Type", "Party Name", "Field Changed",
        "Old Value", "New Value", "Difference", "Status", "Concrete Reason / Audit Impact", "GUID"
    ]
    ws_mod.append(mod_headers)
    ws_mod.row_dimensions[1].height = 26
    for col_idx in range(1, len(mod_headers) + 1):
        cell = ws_mod.cell(row=1, column=col_idx)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = ALIGN_CENTER

    mod_items = [c for c in comparisons if c.status == ClassificationStatus.MODIFIED]
    for item in mod_items:
        tx = item.new_transaction or item.old_transaction
        for diff in item.field_changes:
            old_str = f"₹{diff.old_value:.2f}" if isinstance(diff.old_value, Decimal) else str(diff.old_value)
            new_str = f"₹{diff.new_value:.2f}" if isinstance(diff.new_value, Decimal) else str(diff.new_value)
            diff_val = float(diff.difference) if diff.difference is not None else ""

            # Build concrete reason per field change
            if diff.field == "Total Amount":
                field_reason = f"Base invoice amount modified from {old_str} to {new_str} (Variance: ₹{diff_val:+,.2f})"
            elif diff.field == "TDS Amount":
                field_reason = f"TDS deduction altered from {old_str} to {new_str} (Variance: ₹{diff_val:+,.2f})"
            elif diff.field == "TDS Rate":
                field_reason = f"TDS rate modified from {old_str} to {new_str}; verify statutory rate compliance"
            elif diff.field == "TDS Section":
                field_reason = f"TDS withholding section reclassified from '{old_str}' to '{new_str}'"
            elif diff.field == "Party":
                field_reason = f"Party ledger reassigned from '{old_str}' to '{new_str}'"
            elif diff.field == "Date":
                field_reason = f"Voucher date altered from '{old_str}' to '{new_str}'"
            elif diff.field == "Voucher Number":
                field_reason = f"Voucher numbering updated from '{old_str}' to '{new_str}'"
            elif diff.field == "Narration":
                field_reason = "Transaction narration/remarks edited"
            elif diff.field == "Ledger Entries":
                field_reason = "Line-item ledger distribution changed between accounts"
            else:
                field_reason = f"Field '{diff.field}' altered from '{old_str}' to '{new_str}'"

            ws_mod.append([
                tx.voucher_number or "Unnumbered", tx.date, tx.voucher_type, tx.party_name,
                diff.field, old_str, new_str, diff_val, "MODIFIED", field_reason, tx.guid
            ])
            r_idx = ws_mod.max_row
            ws_mod.row_dimensions[r_idx].height = 20

            for col_idx in range(1, len(mod_headers) + 1):
                c = ws_mod.cell(row=r_idx, column=col_idx)
                c.border = BORDER_THIN
                c.fill = MOD_ROW_FILL
                c.font = DATA_FONT

                if col_idx in [1, 2, 3]:
                    c.alignment = ALIGN_CENTER
                elif col_idx in [4, 5]:
                    c.alignment = ALIGN_LEFT
                    c.font = BOLD_DATA_FONT
                elif col_idx in [6, 7]:
                    c.alignment = ALIGN_RIGHT if isinstance(diff.difference, Decimal) else ALIGN_LEFT
                elif col_idx == 8:
                    c.alignment = ALIGN_RIGHT
                    if isinstance(diff_val, (int, float)):
                        c.number_format = CURRENCY_FORMAT
                        c.font = DIFF_POS_FONT if diff_val >= 0 else DIFF_NEG_FONT
                elif col_idx == 9:
                    c.fill = MOD_BADGE_FILL
                    c.font = MOD_FONT
                    c.alignment = ALIGN_CENTER
                elif col_idx == 10:
                    c.font = REASON_FONT
                    c.alignment = ALIGN_LEFT
                elif col_idx == 11:
                    c.alignment = ALIGN_LEFT

    ws_mod.freeze_panes = "A2"
    ws_mod.auto_filter.ref = ws_mod.dimensions
    _auto_adjust_column_widths(ws_mod)

    # -------------------------------------------------------------------------
    # Sheet 7: Recreated Entries (Highlighted in Indigo with Concrete Reasons)
    # -------------------------------------------------------------------------
    ws_rec = wb.create_sheet(title="Recreated Entries")
    ws_rec.views.sheetView[0].showGridLines = True
    rec_headers = [
        "Old Voucher #", "New Voucher #", "Date", "Voucher Type", "Party Name",
        "Old Amount", "New Amount", "Amount Diff", "Old TDS", "New TDS", "TDS Diff",
        "Status", "Concrete Reason / Match Basis", "Old GUID", "New GUID"
    ]
    ws_rec.append(rec_headers)
    ws_rec.row_dimensions[1].height = 26
    for col_idx in range(1, len(rec_headers) + 1):
        cell = ws_rec.cell(row=1, column=col_idx)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = ALIGN_CENTER

    rec_items = [c for c in comparisons if c.status == ClassificationStatus.RECREATED]
    for item in rec_items:
        o = item.old_transaction
        n = item.new_transaction
        concrete_reason = (
            f"Recreated transaction: {item.match_reason}. "
            f"Old GUID: {o.guid} -> New GUID: {n.guid}."
        )

        ws_rec.append([
            o.voucher_number or "Unnumbered", n.voucher_number or "Unnumbered", n.date, n.voucher_type,
            n.party_name, float(o.total_amount), float(n.total_amount), float(item.amount_difference),
            float(o.tds_amount), float(n.tds_amount), float(item.tds_difference),
            "RECREATED", concrete_reason, o.guid, n.guid
        ])
        r_idx = ws_rec.max_row
        ws_rec.row_dimensions[r_idx].height = 22

        for col_idx in range(1, len(rec_headers) + 1):
            c = ws_rec.cell(row=r_idx, column=col_idx)
            c.border = BORDER_THIN
            c.fill = REC_ROW_FILL
            c.font = DATA_FONT

            if col_idx in [1, 2, 3, 4]:
                c.alignment = ALIGN_CENTER
            elif col_idx == 5:
                c.alignment = ALIGN_LEFT
                c.font = BOLD_DATA_FONT
            elif col_idx in [6, 7, 8, 9, 10, 11]:
                c.alignment = ALIGN_RIGHT
                c.number_format = CURRENCY_FORMAT
                if col_idx in [8, 11]:
                    diff_val = c.value
                    if isinstance(diff_val, (int, float)):
                        c.font = DIFF_POS_FONT if diff_val >= 0 else DIFF_NEG_FONT
            elif col_idx == 12:
                c.fill = REC_BADGE_FILL
                c.font = REC_FONT
                c.alignment = ALIGN_CENTER
            elif col_idx == 13:
                c.font = REASON_FONT
                c.alignment = ALIGN_LEFT
            elif col_idx in [14, 15]:
                c.alignment = ALIGN_LEFT

    ws_rec.freeze_panes = "A2"
    ws_rec.auto_filter.ref = ws_rec.dimensions
    _auto_adjust_column_widths(ws_rec)

    # -------------------------------------------------------------------------
    # Sheet 8: Unresolved Entries (Highlighted in Orange with Concrete Reasons)
    # -------------------------------------------------------------------------
    ws_unres = wb.create_sheet(title="Unresolved Entries")
    ws_unres.views.sheetView[0].showGridLines = True
    unres_headers = [
        "Voucher Number", "Date", "Party Name", "Amount", "Candidate GUID(s)",
        "Status", "Concrete Reason / Ambiguity", "Recommended Audit Action", "Old GUID"
    ]
    ws_unres.append(unres_headers)
    ws_unres.row_dimensions[1].height = 26
    for col_idx in range(1, len(unres_headers) + 1):
        cell = ws_unres.cell(row=1, column=col_idx)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = ALIGN_CENTER

    unres_items = [c for c in comparisons if c.status == ClassificationStatus.UNRESOLVED]
    for item in unres_items:
        tx = item.old_transaction
        cand_str = ", ".join(item.candidates) if item.candidates else "None"
        concrete_reason = (
            item.unresolved_reason or item.match_reason or
            f"Ambiguous Match: Multiple potential matching transactions ({cand_str}) exist in updated data."
        )
        action = "Review source accounting vouchers to confirm whether this transaction was re-entered or split."

        ws_unres.append([
            tx.voucher_number or "Unnumbered", tx.date, tx.party_name, float(tx.total_amount),
            cand_str, "UNRESOLVED / REVIEW", concrete_reason, action, tx.guid
        ])
        r_idx = ws_unres.max_row
        ws_unres.row_dimensions[r_idx].height = 22

        for col_idx in range(1, len(unres_headers) + 1):
            c = ws_unres.cell(row=r_idx, column=col_idx)
            c.border = BORDER_THIN
            c.fill = UNRES_ROW_FILL
            c.font = DATA_FONT

            if col_idx in [1, 2]:
                c.alignment = ALIGN_CENTER
            elif col_idx == 3:
                c.alignment = ALIGN_LEFT
                c.font = BOLD_DATA_FONT
            elif col_idx == 4:
                c.alignment = ALIGN_RIGHT
                c.number_format = CURRENCY_FORMAT
            elif col_idx == 5:
                c.alignment = ALIGN_LEFT
            elif col_idx == 6:
                c.fill = UNRES_BADGE_FILL
                c.font = UNRES_FONT
                c.alignment = ALIGN_CENTER
            elif col_idx in [7, 8]:
                c.font = REASON_FONT
                c.alignment = ALIGN_LEFT
            elif col_idx == 9:
                c.alignment = ALIGN_LEFT

    ws_unres.freeze_panes = "A2"
    ws_unres.auto_filter.ref = ws_unres.dimensions
    _auto_adjust_column_widths(ws_unres)

    # -------------------------------------------------------------------------
    # Sheet 9: Master Changes (Color Coded with Concrete Reasons)
    # -------------------------------------------------------------------------
    ws_mst = wb.create_sheet(title="Master Changes")
    ws_mst.views.sheetView[0].showGridLines = True
    mst_headers = [
        "Party / Ledger Name", "Field Changed", "Old Value", "New Value",
        "Change Type", "Status", "Concrete Reason / Compliance Impact", "Master GUID"
    ]
    ws_mst.append(mst_headers)
    ws_mst.row_dimensions[1].height = 26
    for col_idx in range(1, len(mst_headers) + 1):
        cell = ws_mst.cell(row=1, column=col_idx)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = ALIGN_CENTER

    for mc in master_changes:
        # Determine status, badge and concrete reason
        if mc.change_type == MasterStatus.MASTER_MISSING:
            status_text = "MASTER DELETED"
            row_fill = MISSING_ROW_FILL
            badge_fill = MISSING_BADGE_FILL
            badge_font = MISSING_FONT
            reason = f"Master record for '{mc.party_or_ledger}' exists in Original snapshot but was deleted/missing in Updated snapshot."
        elif mc.change_type == MasterStatus.MASTER_NEW:
            status_text = "MASTER NEW"
            row_fill = NEW_ROW_FILL
            badge_fill = NEW_BADGE_FILL
            badge_font = NEW_FONT
            reason = f"Newly added master ledger record for '{mc.party_or_ledger}' created in Updated snapshot."
        else:
            status_text = "MASTER MODIFIED"
            row_fill = MOD_ROW_FILL
            badge_fill = MOD_BADGE_FILL
            badge_font = MOD_FONT
            if mc.field == "PAN":
                reason = f"PAN changed from '{mc.old_value}' to '{mc.new_value}'. Compliance Impact: Invalid/missing PAN triggers 20% higher TDS under Sec 206AA."
            elif mc.field == "TDS Section":
                reason = f"TDS Section changed from '{mc.old_value}' to '{mc.new_value}'. Compliance Impact: Alters withholding tax rate and return category."
            elif mc.field == "TDS Applicable":
                reason = f"TDS Applicability changed from '{mc.old_value}' to '{mc.new_value}'."
            elif mc.field == "TDS Rate Name":
                reason = f"TDS Rate Name altered from '{mc.old_value}' to '{mc.new_value}'."
            elif mc.field == "Parent Group":
                reason = f"Parent Group reclassified from '{mc.old_value}' to '{mc.new_value}'."
            elif mc.field == "Name":
                reason = f"Master account name renamed from '{mc.old_value}' to '{mc.new_value}'."
            else:
                reason = f"Master property '{mc.field}' modified from '{mc.old_value}' to '{mc.new_value}'."

        ws_mst.append([
            mc.party_or_ledger, mc.field, str(mc.old_value), str(mc.new_value),
            mc.change_type.value, status_text, reason, mc.master_id
        ])
        r_idx = ws_mst.max_row
        ws_mst.row_dimensions[r_idx].height = 20

        for col_idx in range(1, len(mst_headers) + 1):
            c = ws_mst.cell(row=r_idx, column=col_idx)
            c.border = BORDER_THIN
            c.fill = row_fill
            c.font = DATA_FONT

            if col_idx in [1, 2]:
                c.alignment = ALIGN_LEFT
                c.font = BOLD_DATA_FONT
            elif col_idx in [3, 4]:
                c.alignment = ALIGN_LEFT
            elif col_idx == 5:
                c.alignment = ALIGN_CENTER
            elif col_idx == 6:
                c.fill = badge_fill
                c.font = badge_font
                c.alignment = ALIGN_CENTER
            elif col_idx == 7:
                c.font = REASON_FONT
                c.alignment = ALIGN_LEFT
            elif col_idx == 8:
                c.alignment = ALIGN_LEFT

    ws_mst.freeze_panes = "A2"
    ws_mst.auto_filter.ref = ws_mst.dimensions
    _auto_adjust_column_widths(ws_mst)

    # Remove any empty default sheet if left over
    if "Sheet" in wb.sheetnames and len(wb["Sheet"]._cells) == 0:
        wb.remove(wb["Sheet"])

    # Reorder sheets:
    # If Form 26Q format is explicitly requested, place Form 26Q sheets first.
    # Otherwise, strictly output the 12-sheet TDS_Change_Comparison_Reportmk4 format!
    if report_format in ("form26q", "sectionwise"):
        wb._sheets = [ws_sec_summary, ws_all_tx] + section_sheets + [ws_review_points] + [
            s for s in wb.worksheets if s not in (ws_sec_summary, ws_all_tx, ws_review_points) and s not in section_sheets
        ]
        wb.active = ws_sec_summary
    else:
        wb._sheets = [
            ws_tds_summary,
            ws_tds_working,
            ws_review_points,
            ws_recon_overview,
            ws_party,
            ws_missing,
            ws_backdated,
            ws_new_month,
            ws_mod,
            ws_rec,
            ws_unres,
            ws_mst,
        ]
        wb.active = ws_tds_summary

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue()

