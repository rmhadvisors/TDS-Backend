import io
from decimal import Decimal
from typing import List, Optional
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from app.models.comparison_models import (
    ClassificationStatus,
    ComparisonSummarySchema,
    MasterChange,
    PartySummary,
    ReconciliationResult,
    TransactionComparison,
)

# Colors and styles
HEADER_FILL = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid")
HEADER_FONT = Font(name="Calibri", size=11, bold=True, color="FFFFFF")

SECTION_FILL = PatternFill(start_color="334155", end_color="334155", fill_type="solid")
SECTION_FONT = Font(name="Calibri", size=11, bold=True, color="FFFFFF")

PASS_FILL = PatternFill(start_color="DCFCE7", end_color="DCFCE7", fill_type="solid")
PASS_FONT = Font(name="Calibri", size=11, bold=True, color="166534")

FAIL_FILL = PatternFill(start_color="FEE2E2", end_color="FEE2E2", fill_type="solid")
FAIL_FONT = Font(name="Calibri", size=11, bold=True, color="991B1B")

DIFF_POS_FILL = PatternFill(start_color="F0FDF4", end_color="F0FDF4", fill_type="solid")
DIFF_NEG_FILL = PatternFill(start_color="FEF2F2", end_color="FEF2F2", fill_type="solid")

BORDER_THIN = Border(
    left=Side(style="thin", color="E2E8F0"),
    right=Side(style="thin", color="E2E8F0"),
    top=Side(style="thin", color="E2E8F0"),
    bottom=Side(style="thin", color="E2E8F0"),
)

CURRENCY_FORMAT = '"₹"#,##0.00;[Red]("-₹"#,##0.00);"-"'
NUMBER_FORMAT = "#,##0"


def _auto_adjust_column_widths(ws):
    """Adjusts column widths dynamically and adds auto-filters."""
    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            val = str(cell.value or "")
            if len(val) > max_len:
                max_len = len(val)
        ws.column_dimensions[col_letter].width = max(max_len + 4, 14)


def generate_excel_report(
    summary: ComparisonSummarySchema,
    reconciliation: ReconciliationResult,
    party_summaries: List[PartySummary],
    comparisons: List[TransactionComparison],
    master_changes: List[MasterChange],
) -> bytes:
    """
    Generates a full 8-sheet Excel comparison report using openpyxl.
    Returns bytes content of the workbook.
    """
    wb = openpyxl.Workbook()

    # -------------------------------------------------------------------------
    # Sheet 1: Summary
    # -------------------------------------------------------------------------
    ws_summary = wb.active
    ws_summary.title = "Summary"
    ws_summary.views.sheetView[0].showGridLines = True

    # Title
    ws_summary["A1"] = "TDS Change Comparison Summary Report"
    ws_summary["A1"].font = Font(name="Calibri", size=16, bold=True, color="0F172A")
    ws_summary.row_dimensions[1].height = 30

    # Section 1: Transaction Counts
    ws_summary["A3"] = "TRANSACTION METRIC"
    ws_summary["B3"] = "COUNT"
    ws_summary["A3"].fill = HEADER_FILL
    ws_summary["A3"].font = HEADER_FONT
    ws_summary["B3"].fill = HEADER_FILL
    ws_summary["B3"].font = HEADER_FONT

    count_rows = [
        ("Original Transactions", summary.original_transactions),
        ("Updated Transactions", summary.updated_transactions),
        ("Unchanged", summary.unchanged),
        ("Modified", summary.modified),
        ("New", summary.new),
        ("Missing in Updated Data", summary.missing),
        ("Recreated", summary.recreated),
        ("Unresolved / Review", summary.unresolved),
    ]

    for idx, (label, val) in enumerate(count_rows, start=4):
        ws_summary[f"A{idx}"] = label
        ws_summary[f"B{idx}"] = val
        ws_summary[f"B{idx}"].number_format = NUMBER_FORMAT
        ws_summary[f"A{idx}"].border = BORDER_THIN
        ws_summary[f"B{idx}"].border = BORDER_THIN

    # Section 2: Monetary Summary
    start_row = 13
    ws_summary[f"A{start_row}"] = "MONETARY METRIC"
    ws_summary[f"B{start_row}"] = "ORIGINAL (₹)"
    ws_summary[f"C{start_row}"] = "UPDATED (₹)"
    ws_summary[f"D{start_row}"] = "DIFFERENCE (₹)"
    for col_let in ["A", "B", "C", "D"]:
        ws_summary[f"{col_let}{start_row}"].fill = HEADER_FILL
        ws_summary[f"{col_let}{start_row}"].font = HEADER_FONT

    money_rows = [
        ("Total Transaction Amount", float(summary.original_amount), float(summary.updated_amount), float(summary.amount_difference)),
        ("Total TDS Amount", float(summary.original_tds), float(summary.updated_tds), float(summary.tds_difference)),
    ]

    for idx, (label, orig_v, upd_v, diff_v) in enumerate(money_rows, start=start_row + 1):
        ws_summary[f"A{idx}"] = label
        ws_summary[f"B{idx}"] = orig_v
        ws_summary[f"C{idx}"] = upd_v
        ws_summary[f"D{idx}"] = diff_v

        for col_let in ["A", "B", "C", "D"]:
            ws_summary[f"{col_let}{idx}"].border = BORDER_THIN
        for col_let in ["B", "C", "D"]:
            ws_summary[f"{col_let}{idx}"].number_format = CURRENCY_FORMAT

    # Section 3: Reconciliation Checks
    rec_start = 17
    ws_summary[f"A{rec_start}"] = "RECONCILIATION CHECK"
    ws_summary[f"B{rec_start}"] = "STATUS"
    ws_summary[f"A{rec_start}"].fill = HEADER_FILL
    ws_summary[f"A{rec_start}"].font = HEADER_FONT
    ws_summary[f"B{rec_start}"].fill = HEADER_FILL
    ws_summary[f"B{rec_start}"].font = HEADER_FONT

    rec_rows = [
        ("Transaction Count Balance", reconciliation.transaction_count),
        ("Total Amount Balance", reconciliation.amount),
        ("Total TDS Balance", reconciliation.tds),
    ]

    for idx, (label, status) in enumerate(rec_rows, start=rec_start + 1):
        ws_summary[f"A{idx}"] = label
        ws_summary[f"B{idx}"] = status
        ws_summary[f"A{idx}"].border = BORDER_THIN
        ws_summary[f"B{idx}"].border = BORDER_THIN
        if status == "PASS":
            ws_summary[f"B{idx}"].fill = PASS_FILL
            ws_summary[f"B{idx}"].font = PASS_FONT
        else:
            ws_summary[f"B{idx}"].fill = FAIL_FILL
            ws_summary[f"B{idx}"].font = FAIL_FONT

    _auto_adjust_column_widths(ws_summary)

    # -------------------------------------------------------------------------
    # Sheet 2: Party-wise Summary
    # -------------------------------------------------------------------------
    ws_party = wb.create_sheet(title="Party-wise Summary")
    ws_party.views.sheetView[0].showGridLines = True
    party_headers = [
        "Party", "Affected Entries", "Missing", "New", "Modified", "Recreated",
        "Original Amount", "Updated Amount", "Amount Difference",
        "Original TDS", "Updated TDS", "TDS Difference"
    ]
    ws_party.append(party_headers)
    for col_idx in range(1, len(party_headers) + 1):
        cell = ws_party.cell(row=1, column=col_idx)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center")

    for p in party_summaries:
        row_vals = [
            p.party, p.affected_entries, p.missing_count, p.new_count, p.modified_count, p.recreated_count,
            float(p.original_amount), float(p.updated_amount), float(p.amount_difference),
            float(p.original_tds), float(p.updated_tds), float(p.tds_difference),
        ]
        ws_party.append(row_vals)
        r_idx = ws_party.max_row
        for col_idx in range(1, len(party_headers) + 1):
            c = ws_party.cell(row=r_idx, column=col_idx)
            c.border = BORDER_THIN
            if col_idx in [2, 3, 4, 5, 6]:
                c.number_format = NUMBER_FORMAT
            elif col_idx >= 7:
                c.number_format = CURRENCY_FORMAT

    ws_party.freeze_panes = "A2"
    ws_party.auto_filter.ref = ws_party.dimensions
    _auto_adjust_column_widths(ws_party)

    # -------------------------------------------------------------------------
    # Sheet 3: Missing Entries
    # -------------------------------------------------------------------------
    ws_missing = wb.create_sheet(title="Missing Entries")
    ws_missing.views.sheetView[0].showGridLines = True
    missing_headers = ["Voucher Number", "Date", "Voucher Type", "Party", "Original Amount", "Original TDS", "Original GUID", "Status"]
    ws_missing.append(missing_headers)
    for col_idx in range(1, len(missing_headers) + 1):
        cell = ws_missing.cell(row=1, column=col_idx)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT

    missing_items = [c for c in comparisons if c.status == ClassificationStatus.MISSING]
    for item in missing_items:
        tx = item.old_transaction
        ws_missing.append([
            tx.voucher_number, tx.date, tx.voucher_type, tx.party_name,
            float(tx.total_amount), float(tx.tds_amount), tx.guid, "Missing in Updated Data"
        ])
        r_idx = ws_missing.max_row
        for col_idx in range(1, len(missing_headers) + 1):
            c = ws_missing.cell(row=r_idx, column=col_idx)
            c.border = BORDER_THIN
            if col_idx in [5, 6]:
                c.number_format = CURRENCY_FORMAT

    ws_missing.freeze_panes = "A2"
    ws_missing.auto_filter.ref = ws_missing.dimensions
    _auto_adjust_column_widths(ws_missing)

    # -------------------------------------------------------------------------
    # Sheet 4: New Entries
    # -------------------------------------------------------------------------
    ws_new = wb.create_sheet(title="New Entries")
    ws_new.views.sheetView[0].showGridLines = True
    new_headers = ["Voucher Number", "Date", "Voucher Type", "Party", "Updated Amount", "Updated TDS", "Updated GUID", "Status"]
    ws_new.append(new_headers)
    for col_idx in range(1, len(new_headers) + 1):
        cell = ws_new.cell(row=1, column=col_idx)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT

    new_items = [c for c in comparisons if c.status == ClassificationStatus.NEW]
    for item in new_items:
        tx = item.new_transaction
        ws_new.append([
            tx.voucher_number, tx.date, tx.voucher_type, tx.party_name,
            float(tx.total_amount), float(tx.tds_amount), tx.guid, "Newly Added in Updated Data"
        ])
        r_idx = ws_new.max_row
        for col_idx in range(1, len(new_headers) + 1):
            c = ws_new.cell(row=r_idx, column=col_idx)
            c.border = BORDER_THIN
            if col_idx in [5, 6]:
                c.number_format = CURRENCY_FORMAT

    ws_new.freeze_panes = "A2"
    ws_new.auto_filter.ref = ws_new.dimensions
    _auto_adjust_column_widths(ws_new)

    # -------------------------------------------------------------------------
    # Sheet 5: Modified Entries (one row per changed field)
    # -------------------------------------------------------------------------
    ws_mod = wb.create_sheet(title="Modified Entries")
    ws_mod.views.sheetView[0].showGridLines = True
    mod_headers = ["Voucher Number", "Date", "Voucher Type", "Party", "GUID", "Field", "Old Value", "New Value", "Difference"]
    ws_mod.append(mod_headers)
    for col_idx in range(1, len(mod_headers) + 1):
        cell = ws_mod.cell(row=1, column=col_idx)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT

    mod_items = [c for c in comparisons if c.status == ClassificationStatus.MODIFIED]
    for item in mod_items:
        tx = item.new_transaction or item.old_transaction
        for diff in item.field_changes:
            old_str = f"₹{diff.old_value:.2f}" if isinstance(diff.old_value, Decimal) else str(diff.old_value)
            new_str = f"₹{diff.new_value:.2f}" if isinstance(diff.new_value, Decimal) else str(diff.new_value)
            diff_val = float(diff.difference) if diff.difference is not None else ""

            ws_mod.append([
                tx.voucher_number, tx.date, tx.voucher_type, tx.party_name, tx.guid,
                diff.field, old_str, new_str, diff_val
            ])
            r_idx = ws_mod.max_row
            for col_idx in range(1, len(mod_headers) + 1):
                c = ws_mod.cell(row=r_idx, column=col_idx)
                c.border = BORDER_THIN
                if col_idx == 9 and isinstance(diff_val, (int, float)):
                    c.number_format = CURRENCY_FORMAT

    ws_mod.freeze_panes = "A2"
    ws_mod.auto_filter.ref = ws_mod.dimensions
    _auto_adjust_column_widths(ws_mod)

    # -------------------------------------------------------------------------
    # Sheet 6: Recreated Entries
    # -------------------------------------------------------------------------
    ws_rec = wb.create_sheet(title="Recreated Entries")
    ws_rec.views.sheetView[0].showGridLines = True
    rec_headers = [
        "Old GUID", "New GUID", "Old Voucher Number", "New Voucher Number", "Date", "Voucher Type",
        "Party", "Old Amount", "New Amount", "Amount Difference",
        "Old TDS", "New TDS", "TDS Difference", "Match Reason"
    ]
    ws_rec.append(rec_headers)
    for col_idx in range(1, len(rec_headers) + 1):
        cell = ws_rec.cell(row=1, column=col_idx)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT

    rec_items = [c for c in comparisons if c.status == ClassificationStatus.RECREATED]
    for item in rec_items:
        o = item.old_transaction
        n = item.new_transaction
        ws_rec.append([
            o.guid, n.guid, o.voucher_number, n.voucher_number, n.date, n.voucher_type,
            n.party_name, float(o.total_amount), float(n.total_amount), float(item.amount_difference),
            float(o.tds_amount), float(n.tds_amount), float(item.tds_difference), item.match_reason
        ])
        r_idx = ws_rec.max_row
        for col_idx in range(1, len(rec_headers) + 1):
            c = ws_rec.cell(row=r_idx, column=col_idx)
            c.border = BORDER_THIN
            if col_idx in [8, 9, 10, 11, 12, 13]:
                c.number_format = CURRENCY_FORMAT

    ws_rec.freeze_panes = "A2"
    ws_rec.auto_filter.ref = ws_rec.dimensions
    _auto_adjust_column_widths(ws_rec)

    # -------------------------------------------------------------------------
    # Sheet 7: Unresolved Entries
    # -------------------------------------------------------------------------
    ws_unres = wb.create_sheet(title="Unresolved Entries")
    ws_unres.views.sheetView[0].showGridLines = True
    unres_headers = ["Old GUID", "New Candidate GUID(s)", "Voucher Number", "Date", "Party", "Amount", "Reason"]
    ws_unres.append(unres_headers)
    for col_idx in range(1, len(unres_headers) + 1):
        cell = ws_unres.cell(row=1, column=col_idx)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT

    unres_items = [c for c in comparisons if c.status == ClassificationStatus.UNRESOLVED]
    for item in unres_items:
        tx = item.old_transaction
        cand_str = ", ".join(item.candidates) if item.candidates else "None"
        ws_unres.append([
            tx.guid, cand_str, tx.voucher_number, tx.date, tx.party_name,
            float(tx.total_amount), item.unresolved_reason or item.match_reason
        ])
        r_idx = ws_unres.max_row
        for col_idx in range(1, len(unres_headers) + 1):
            c = ws_unres.cell(row=r_idx, column=col_idx)
            c.border = BORDER_THIN
            if col_idx == 6:
                c.number_format = CURRENCY_FORMAT

    ws_unres.freeze_panes = "A2"
    ws_unres.auto_filter.ref = ws_unres.dimensions
    _auto_adjust_column_widths(ws_unres)

    # -------------------------------------------------------------------------
    # Sheet 8: Master Changes
    # -------------------------------------------------------------------------
    ws_mst = wb.create_sheet(title="Master Changes")
    ws_mst.views.sheetView[0].showGridLines = True
    mst_headers = ["Master ID/GUID", "Party/Ledger", "Field", "Old Value", "New Value", "Change Type"]
    ws_mst.append(mst_headers)
    for col_idx in range(1, len(mst_headers) + 1):
        cell = ws_mst.cell(row=1, column=col_idx)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT

    for mc in master_changes:
        ws_mst.append([
            mc.master_id, mc.party_or_ledger, mc.field, str(mc.old_value), str(mc.new_value), mc.change_type.value
        ])
        r_idx = ws_mst.max_row
        for col_idx in range(1, len(mst_headers) + 1):
            c = ws_mst.cell(row=r_idx, column=col_idx)
            c.border = BORDER_THIN

    ws_mst.freeze_panes = "A2"
    ws_mst.auto_filter.ref = ws_mst.dimensions
    _auto_adjust_column_widths(ws_mst)

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue()
