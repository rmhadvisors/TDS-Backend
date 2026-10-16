import io
from decimal import Decimal
import openpyxl
import pytest

from app.models.comparison_models import (
    ClassificationStatus,
    ComparisonSummarySchema,
    FieldDifference,
    MasterChange,
    MasterStatus,
    NormalizedLedger,
    NormalizedTransaction,
    PartySummary,
    ReconciliationResult,
    TransactionComparison,
)
from app.services.excel_report import generate_excel_report
from app.services.normalizer import normalize_transaction


def test_excel_report_concrete_reasons_and_highlights():
    # Setup summary & reconciliation schemas
    summary = ComparisonSummarySchema(
        original_transactions=5,
        updated_transactions=5,
        unchanged=1,
        modified=1,
        new=1,
        missing=1,
        recreated=1,
        unresolved=0,
        original_amount="250000.00",
        updated_amount="275000.00",
        amount_difference="25000.00",
        original_tds="5000.00",
        updated_tds="5500.00",
        tds_difference="500.00",
    )

    reconciliation = ReconciliationResult(
        transaction_count="PASS",
        amount="FAIL",
        tds="FAIL",
        is_valid=False,
    )

    party_summaries = [
        PartySummary(
            party="Vendor Alpha",
            affected_entries=2,
            missing_count=1,
            new_count=1,
            modified_count=0,
            recreated_count=0,
            original_amount=Decimal("100000.00"),
            updated_amount=Decimal("125000.00"),
            amount_difference=Decimal("25000.00"),
            original_tds=Decimal("2000.00"),
            updated_tds=Decimal("2500.00"),
            tds_difference=Decimal("500.00"),
        )
    ]

    t_old_missing = NormalizedTransaction(
        guid="GUID-OLD-1",
        voucher_number="VCH-100",
        voucher_type="Payment",
        date="2026-05-10",
        party_name="Vendor Alpha",
        total_amount=Decimal("50000.00"),
        tds_amount=Decimal("1000.00"),
        ledgers=[NormalizedLedger(ledger_name="Vendor Alpha", amount=Decimal("50000.00"), is_debit=True, is_credit=False)]
    )

    t_new_added = NormalizedTransaction(
        guid="GUID-NEW-2",
        voucher_number="VCH-200",
        voucher_type="Payment",
        date="2026-05-15",
        party_name="Vendor Alpha",
        total_amount=Decimal("75000.00"),
        tds_amount=Decimal("1500.00"),
        ledgers=[NormalizedLedger(ledger_name="Vendor Alpha", amount=Decimal("75000.00"), is_debit=True, is_credit=False)]
    )

    t_mod_old = NormalizedTransaction(
        guid="GUID-MOD-3",
        voucher_number="VCH-300",
        voucher_type="Journal",
        date="2026-05-20",
        party_name="Vendor Alpha",
        total_amount=Decimal("20000.00"),
        tds_amount=Decimal("400.00"),
    )
    t_mod_new = NormalizedTransaction(
        guid="GUID-MOD-3",
        voucher_number="VCH-300",
        voucher_type="Journal",
        date="2026-05-20",
        party_name="Vendor Alpha",
        total_amount=Decimal("25000.00"),
        tds_amount=Decimal("500.00"),
    )

    t_new_month = NormalizedTransaction(
        guid="GUID-NEW-3",
        voucher_number="VCH-201",
        voucher_type="Payment",
        date="2026-06-05",
        party_name="Vendor Alpha",
        total_amount=Decimal("30000.00"),
        tds_amount=Decimal("600.00"),
        ledgers=[NormalizedLedger(ledger_name="Vendor Alpha", amount=Decimal("30000.00"), is_debit=True, is_credit=False)]
    )

    comparisons = [
        TransactionComparison(
            status=ClassificationStatus.MISSING,
            old_transaction=t_old_missing,
            new_transaction=None,
            amount_difference=Decimal("-50000.00"),
            tds_difference=Decimal("-1000.00"),
        ),
        TransactionComparison(
            status=ClassificationStatus.NEW,
            old_transaction=None,
            new_transaction=t_new_added,
            amount_difference=Decimal("75000.00"),
            tds_difference=Decimal("1500.00"),
        ),
        TransactionComparison(
            status=ClassificationStatus.NEW,
            old_transaction=None,
            new_transaction=t_new_month,
            amount_difference=Decimal("30000.00"),
            tds_difference=Decimal("600.00"),
        ),
        TransactionComparison(
            status=ClassificationStatus.MODIFIED,
            old_transaction=t_mod_old,
            new_transaction=t_mod_new,
            field_changes=[
                FieldDifference(
                    field="Total Amount",
                    old_value=Decimal("20000.00"),
                    new_value=Decimal("25000.00"),
                    difference=Decimal("5000.00")
                )
            ],
            amount_difference=Decimal("5000.00"),
            tds_difference=Decimal("100.00"),
        )
    ]

    master_changes = [
        MasterChange(
            master_id="MST-1",
            party_or_ledger="Vendor Alpha",
            field="PAN",
            old_value="N/A",
            new_value="ABCDE1234F",
            change_type=MasterStatus.MASTER_MODIFIED,
        )
    ]

    report_bytes = generate_excel_report(
        summary=summary,
        reconciliation=reconciliation,
        party_summaries=party_summaries,
        comparisons=comparisons,
        master_changes=master_changes,
        party_name="Vendor Alpha",
        period="2026-05-01 to 2026-05-31",
    )

    assert len(report_bytes) > 0
    wb = openpyxl.load_workbook(io.BytesIO(report_bytes))

    # 1. Check Sheet Names (9 sheets)
    expected_sheets = [
        "Summary", "Party-wise Summary", "Deleted Entries", "Backdated Entries", "New Month Entries",
        "Modified Entries", "Recreated Entries", "Unresolved Entries", "Master Changes"
    ]
    for s in expected_sheets:
        assert s in wb.sheetnames

    # 2. Check Deleted Entries Sheet (Highlight + Concrete Reason)
    ws_missing = wb["Deleted Entries"]
    assert ws_missing["A1"].value == "Voucher Number"
    assert ws_missing["H1"].value == "Concrete Reason / Deletion Impact"
    assert ws_missing["A2"].value == "VCH-100"
    assert ws_missing["G2"].value == "DELETED / MISSING"
    assert "deleted/removed in Updated snapshot" in ws_missing["H2"].value
    assert ws_missing["A2"].fill.start_color.rgb in ["00FEF2F2", "FEF2F2", None]

    # 3. Check Backdated Entries Sheet (Forward/Backward TDS Calculation + Highlights)
    ws_backdated = wb["Backdated Entries"]
    assert ws_backdated["A1"].value == "Voucher Number"
    assert ws_backdated["E1"].value == "Gross / Base Amount"
    assert ws_backdated["F1"].value == "Calculation Method"
    assert ws_backdated["G1"].value == "TDS Section"
    assert ws_backdated["H1"].value == "Statutory Rate"
    assert ws_backdated["I1"].value == "Calculated Statutory TDS"
    assert ws_backdated["J1"].value == "Booked TDS in Tally"
    assert ws_backdated["K1"].value == "TDS Shortfall / Variance"
    assert ws_backdated["L1"].value == "Status"
    assert ws_backdated["M1"].value == "Note / Nature of Deduction"
    assert ws_backdated["N1"].value == "Statutory Audit & Compliance Action"
    assert ws_backdated["A2"].value == "VCH-200"
    assert "194C" in ws_backdated["G2"].value
    assert ws_backdated["L2"].value == "BACKDATED / ADDED"
    assert "Contractors" in ws_backdated["M2"].value
    assert "Backdated addition" in ws_backdated["N2"].value
    assert ws_backdated["A2"].fill.start_color.rgb in ["00FFFBEB", "FFFBEB", None]

    # 4. Check New Month Entries Sheet (Forward/Backward TDS Calculation + Highlights)
    ws_new_month = wb["New Month Entries"]
    assert ws_new_month["A1"].value == "Voucher Number"
    assert ws_new_month["E1"].value == "Gross / Base Amount"
    assert ws_new_month["F1"].value == "Calculation Method"
    assert ws_new_month["G1"].value == "TDS Section"
    assert ws_new_month["H1"].value == "Statutory Rate"
    assert ws_new_month["I1"].value == "Calculated Statutory TDS"
    assert ws_new_month["J1"].value == "Booked TDS in Tally"
    assert ws_new_month["K1"].value == "TDS Shortfall / Variance"
    assert ws_new_month["L1"].value == "Status"
    assert ws_new_month["M1"].value == "Note / Nature of Deduction"
    assert ws_new_month["N1"].value == "Statutory Audit & Compliance Action"
    assert ws_new_month["A2"].value == "VCH-201"
    assert "194C" in ws_new_month["G2"].value
    assert ws_new_month["L2"].value == "NEW MONTH / ADDED"
    assert "Contractors" in ws_new_month["M2"].value
    assert "Extended period transaction" in ws_new_month["N2"].value
    assert ws_new_month["A2"].fill.start_color.rgb in ["00F0FDF4", "F0FDF4", None]

    # 5. Check Modified Entries Sheet
    ws_mod = wb["Modified Entries"]
    assert ws_mod["A2"].value == "VCH-300"
    assert ws_mod["E2"].value == "Total Amount"
    assert "Base invoice amount modified" in ws_mod["J2"].value

    # 6. Check Master Changes Sheet
    ws_mst = wb["Master Changes"]
    assert ws_mst["A2"].value == "Vendor Alpha"
    assert ws_mst["B2"].value == "PAN"
    assert "Sec 206AA" in ws_mst["G2"].value


def test_transaction_normalization_with_inventory_and_ledgerentries():
    # Test normalization when voucher has 'ledgerentries' and 'allinventoryentries'
    raw_vch = {
        "guid": "TEST-VCH-PURCHASE-1",
        "vouchernumber": "PUR-01",
        "vouchertypename": "Purchase",
        "date": "2026-05-16",
        "partyledgername": "Classic Tiles Studio",
        "ledgerentries": [
            {
                "ledgername": "Classic Tiles Studio",
                "amount": "704965.00",
                "isdeemedpositive": False,
                "ispartyledger": True,
            },
            {
                "ledgername": "Input Cgst @9%",
                "amount": "-53768.53",
                "isdeemedpositive": True,
            },
            {
                "ledgername": "Input Sgst @9%",
                "amount": "-53768.53",
                "isdeemedpositive": True,
            },
            {
                "ledgername": "Round Off",
                "amount": "0.21",
                "isdeemedpositive": True,
            }
        ],
        "allinventoryentries": [
            {
                "stockitemname": "Tiles @18%",
                "amount": "-597428.15",
                "accountingallocations": [
                    {
                        "ledgername": "Tiles",
                        "amount": "-597428.15",
                        "isdeemedpositive": True,
                    }
                ]
            }
        ]
    }

    norm_tx = normalize_transaction(raw_vch, 0)
    assert norm_tx.guid == "TEST-VCH-PURCHASE-1"
    assert norm_tx.voucher_number == "PUR-01"
    assert norm_tx.party_name == "Classic Tiles Studio"
    assert norm_tx.total_amount == Decimal("704965.00")
    assert len(norm_tx.ledgers) == 5  # Party + CGST + SGST + RoundOff + Tiles


def test_pan_4th_character_rate_inference():
    from app.services.excel_report import infer_tds_section_and_rate, extract_pan_from_tx

    # 1. Individual PAN (4th char 'P') under 194C -> 1.00%
    tx_indv = NormalizedTransaction(
        guid="G-1",
        voucher_number="V-1",
        voucher_type="Payment",
        date="2026-05-10",
        party_name="Abhinav Cargo Movers",
        narration="Bill 101 PAN: CAHPP4138H Transport Charges",
        total_amount=Decimal("0.00"),
        tds_amount=Decimal("0.00"),
    )
    assert extract_pan_from_tx(tx_indv) == "CAHPP4138H"
    sec, rt, nt = infer_tds_section_and_rate(tx_indv)
    assert sec == "194C"
    assert rt == "1.00%"
    assert "Individual/HUF" in nt

    # 2. Company PAN (4th char 'C') under 194C -> 2.00%
    tx_comp = NormalizedTransaction(
        guid="G-2",
        voucher_number="V-2",
        voucher_type="Payment",
        date="2026-05-10",
        party_name="DHL Express India Pvt Ltd",
        raw_reference={"pan": "AABCD3611Q"},
        total_amount=Decimal("0.00"),
        tds_amount=Decimal("0.00"),
    )
    assert extract_pan_from_tx(tx_comp) == "AABCD3611Q"
    sec, rt, nt = infer_tds_section_and_rate(tx_comp)
    assert sec == "194C"
    assert rt == "2.00%"
    assert "Company" in nt

    # 3. Partnership Firm PAN (4th char 'F') under 194C -> 2.00%
    tx_firm = NormalizedTransaction(
        guid="G-3",
        voucher_number="V-3",
        voucher_type="Payment",
        date="2026-05-10",
        party_name="Easter Shipping & Logistics",
        narration="Freight invoice pan-AACFE9624L",
        total_amount=Decimal("0.00"),
        tds_amount=Decimal("0.00"),
    )
    assert extract_pan_from_tx(tx_firm) == "AACFE9624L"
    sec, rt, nt = infer_tds_section_and_rate(tx_firm)
    assert sec == "194C"
    assert rt == "2.00%"
    assert "Firm/LLP" in nt


def test_statutory_tds_rate_pdf_chart_coverage():
    from app.services.excel_report import infer_tds_section_and_rate

    # 1. 194J - Technical Services (2%)
    tx_tech = NormalizedTransaction(
        guid="G-T1", voucher_number="V-T1", voucher_type="Journal", date="2026-05-10",
        party_name="Tech Cloud Systems", narration="Software development IT service charges",
        total_amount=Decimal("100000.00"), tds_amount=Decimal("2000.00")
    )
    sec, rt, nt = infer_tds_section_and_rate(tx_tech)
    assert sec == "194J"
    assert rt == "2.00%"

    # 2. 194J - Professional Services (10%)
    tx_prof = NormalizedTransaction(
        guid="G-T2", voucher_number="V-T2", voucher_type="Journal", date="2026-05-10",
        party_name="Apex Legal Consultants", narration="Statutory audit and CA consultation fee",
        total_amount=Decimal("100000.00"), tds_amount=Decimal("10000.00")
    )
    sec, rt, nt = infer_tds_section_and_rate(tx_prof)
    assert sec == "194J"
    assert rt == "10.00%"

    # 3. 194H - Commission / Brokerage (2% w.e.f. 1-Oct-2024 onwards)
    tx_comm = NormalizedTransaction(
        guid="G-T3", voucher_number="V-T3", voucher_type="Payment", date="2026-05-10",
        party_name="Prime Brokerage Agency", narration="Channel partner incentive and commission",
        total_amount=Decimal("50000.00"), tds_amount=Decimal("1000.00")
    )
    sec, rt, nt = infer_tds_section_and_rate(tx_comm)
    assert sec == "194H"
    assert rt == "2.00%"

    # 4. 194I - Rent on Plant / Machinery / Equipment (2%)
    tx_plant_rent = NormalizedTransaction(
        guid="G-T4", voucher_number="V-T4", voucher_type="Journal", date="2026-05-10",
        party_name="Industrial Crane Hires", narration="Monthly plant & machinery crane hire rent",
        total_amount=Decimal("60000.00"), tds_amount=Decimal("1200.00")
    )
    sec, rt, nt = infer_tds_section_and_rate(tx_plant_rent)
    assert sec == "194I"
    assert rt == "2.00%"

    # 5. 194I - Rent on Land / Building (10%)
    tx_bldg_rent = NormalizedTransaction(
        guid="G-T5", voucher_number="V-T5", voucher_type="Journal", date="2026-05-10",
        party_name="Corporate Tower Properties", narration="Office premises monthly rent",
        total_amount=Decimal("80000.00"), tds_amount=Decimal("8000.00")
    )
    sec, rt, nt = infer_tds_section_and_rate(tx_bldg_rent)
    assert sec == "194I"
    assert rt == "10.00%"

    # 6. 194Q - Purchase of Goods (0.10%)
    tx_goods = NormalizedTransaction(
        guid="G-T6", voucher_number="V-T6", voucher_type="Purchase", date="2026-05-10",
        party_name="Steel Raw Materials Corp", narration="Purchase of Goods exceeding 50 Lakhs",
        total_amount=Decimal("6000000.00"), tds_amount=Decimal("6000.00")
    )
    sec, rt, nt = infer_tds_section_and_rate(tx_goods)
    assert sec == "194Q"
    assert rt == "0.10%"

    # 7. 194R - Business Perquisite (10%)
    tx_perq = NormalizedTransaction(
        guid="G-T7", voucher_number="V-T7", voucher_type="Journal", date="2026-05-10",
        party_name="Dealer Incentive Scheme", narration="Dealer business promotion perquisite",
        total_amount=Decimal("50000.00"), tds_amount=Decimal("5000.00")
    )
    sec, rt, nt = infer_tds_section_and_rate(tx_perq)
    assert sec == "194R"
    assert rt == "10.00%"

    # 8. 194T - Partner Remuneration (10%)
    tx_partner = NormalizedTransaction(
        guid="G-T8", voucher_number="V-T8", voucher_type="Payment", date="2026-05-10",
        party_name="Managing Partner", narration="Partner remuneration and interest on capital",
        total_amount=Decimal("100000.00"), tds_amount=Decimal("10000.00")
    )
    sec, rt, nt = infer_tds_section_and_rate(tx_partner)
    assert sec == "194T"
    assert rt == "10.00%"

    # 9. 194B - Winnings (30%)
    tx_win = NormalizedTransaction(
        guid="G-T9", voucher_number="V-T9", voucher_type="Payment", date="2026-05-10",
        party_name="Contest Winner", narration="Lottery winning payout 194B",
        total_amount=Decimal("50000.00"), tds_amount=Decimal("15000.00")
    )
    sec, rt, nt = infer_tds_section_and_rate(tx_win)
    assert sec == "194B"
    assert rt == "30.00%"


def test_tds_summary_and_tds_working_sheets():
    """Validates that TDS Summary and TDS Working sheets match the Zeon dashboard structure."""
    summary = ComparisonSummarySchema(
        original_transactions=2,
        updated_transactions=2,
        unchanged=2,
        modified=0,
        new=0,
        missing=0,
        recreated=0,
        unresolved=0,
        original_amount="100000.00",
        updated_amount="100000.00",
        amount_difference="0.00",
        original_tds="2000.00",
        updated_tds="2000.00",
        tds_difference="0.00",
    )

    reconciliation = ReconciliationResult(
        transaction_count="PASS",
        amount="PASS",
        tds="PASS",
        is_valid=True,
    )

    tx1 = NormalizedTransaction(
        guid="TX-01",
        voucher_number="VCH-01",
        voucher_type="Payment",
        date="2026-04-10",
        party_name="DHL Express India Pvt Ltd",
        total_amount=Decimal("50000.00"),
        tds_amount=Decimal("1000.00"),
        raw_reference={"pan": "AABCD3611Q"}
    )
    tx2 = NormalizedTransaction(
        guid="TX-02",
        voucher_number="VCH-02",
        voucher_type="Payment",
        date="2026-05-15",
        party_name="Easter Shipping & Logistics",
        total_amount=Decimal("50000.00"),
        tds_amount=Decimal("1000.00"),
        raw_reference={"pan": "AACFE9624L"}
    )

    comparisons = [
        TransactionComparison(status=ClassificationStatus.UNCHANGED, old_transaction=tx1, new_transaction=tx1),
        TransactionComparison(status=ClassificationStatus.UNCHANGED, old_transaction=tx2, new_transaction=tx2),
    ]

    report_bytes = generate_excel_report(
        summary=summary,
        reconciliation=reconciliation,
        party_summaries=[],
        comparisons=comparisons,
        master_changes=[],
        party_name="ZEON PHARMA INDUSTRIES",
        period="2026-27",
    )

    wb = openpyxl.load_workbook(io.BytesIO(report_bytes), data_only=False)

    # 1. Verify Sheet Positioning (TDS Summary & TDS Working as first two sheets)
    assert wb.sheetnames[0] == "TDS Summary"
    assert wb.sheetnames[1] == "TDS Working"

    # 2. Verify TDS Summary Sheet Matrix & Lookups
    ws_sum = wb["TDS Summary"]
    assert "TDS SUMMARY DASHBOARD" in str(ws_sum["A1"].value)
    assert ws_sum["A4"].value == "SECTION-WISE TDS AMOUNT — MONTH × SECTION MATRIX"
    assert ws_sum["A5"].value == "Month"
    assert ws_sum["A6"].value == "April"
    assert ws_sum["A18"].value == "GRAND TOTAL"
    assert "SUMIFS" in str(ws_sum["B6"].value)
    assert ws_sum["A20"].value == "QUARTERLY TDS SUMMARY"
    assert ws_sum["A28"].value == "MONTH + SECTION LOOKUP"

    # 3. Verify TDS Working Sheet Structure & Formulas
    ws_work = wb["TDS Working"]
    assert "TDS WORKING" in str(ws_work["A1"].value)
    assert ws_work["A2"].value == "Month"
    assert ws_work["B2"].value == "Date"
    assert ws_work["D2"].value == "Party Name"
    assert ws_work["E2"].value == "PAN No"
    assert ws_work["G2"].value == "Bill/Paid Amount"
    assert ws_work["H2"].value == "TDS Rate %"
    assert ws_work["I2"].value == "TDS Amount"
    assert ws_work["N2"].value == "Section"
    assert ws_work["Q2"].value == "Quarter"

    # Check row 3 formulas
    assert ws_work["A3"].value == "April"
    assert "=MID(E3,4,1)" in str(ws_work["F3"].value)
    assert "ROUND" in str(ws_work["I3"].value)
    assert "April" in str(ws_work["Q3"].value)


def test_review_point_lists_deductees_without_pan():
    import openpyxl
    from app.services.excel_report import _pan_missing_findings

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["title"])
    ws.append(["Month", "Date", "Bill No", "Party Name", "PAN No", "", "Bill/Paid Amount", "TDS Rate %", "TDS Amount",
               "", "", "", "Diff", "Section"])
    row = lambda bill, party, pan, amt, rate, diff, sec: ["April", "2025-04-01", bill, party, pan, "", amt, rate, "", "", "", "", diff, sec]
    ws.append(row("1", "No Pan Agent", None, 100000, 2, None, "194H"))
    ws.append(row("2, 3", "No Pan Agent", "", 50000, 2, -500, "194H"))     # TDS booked 500 instead of 1,000
    ws.append(row("4", "Acme Pvt Ltd", "AABCA 1234 F", 50000, 2, None, "194C"))   # PAN with spaces: fine

    findings = _pan_missing_findings(ws)
    assert [(f[0], f[1], f[2], f[3], f[4]) for f in findings] == [("PAN MISSING", "1, 2, 3", "No Pan Agent", "194H", 2500)]
    assert "Rs. 30,000" in findings[0][5]   # 20% of 1,50,000 under Sec 206AA
