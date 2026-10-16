from decimal import Decimal

from app.services.normalizer import is_tds_viable_transaction, normalize_master, normalize_transaction
from app.services.tds_context import build_tds_context
from app.services.tds_deductions import (
    extract_booked_deductions,
    find_194q_not_deducted,
    group_by_deductee_month,
)


def _group(name, parent="", reserved="", **flags):
    return {"metadata": {"type": "Group", "name": name, "reservedname": reserved}, "name": name, "parent": parent, **flags}


def _ledger(name, parent, **extra):
    return {"metadata": {"type": "Ledger", "name": name}, "name": name, "parent": parent, **extra}


MASTERS = [normalize_master(m, i) for i, m in enumerate([
    _group("Current Liabilities", reserved="Current Liabilities"),
    _group("Duties & Taxes", "Current Liabilities", "Duties & Taxes"),
    _group("TDS", "Duties & Taxes"),
    _group("Sundry Creditors", "Current Liabilities", "Sundry Creditors"),
    _group("Purchase Accounts", reserved="Purchase Accounts"),
    _group("GOLD PURCHASE", "Purchase Accounts"),
    _group("Other Expenses", reserved="Indirect Expenses"),          # renamed primary group
    _group("Employee Benefit Expenses", isrevenue=True, isdeemedpositive=True),   # user-created primary group
    _group("Salary", "Employee Benefit Expenses"),
    _group("Capital Account", reserved="Capital Account"),
    _group("Sales Accounts", reserved="Sales Accounts"),
    _group("Current Assets", reserved="Current Assets"),
    _group("SUNDRY DR", "Current Assets", "Sundry Debtors"),
    _group("Bank Accounts", "Current Assets", "Bank Accounts"),
    _ledger("TDS ON PURCHASE OF GOODS @ 0.10%", "TDS", tdsratename="TDS on Purchase of Goods"),
    _ledger("TDS ON CONTRACT @ 1%", "TDS"),
    _ledger("Tds 194T", "TDS"),
    _ledger("tds Payable", "TDS"),
    _ledger("Purchase Gold GST @ 3%", "GOLD PURCHASE"),
    _ledger("Input CGST @ 1.5%", "Duties & Taxes"),
    _ledger("Labour Charges", "Other Expenses"),
    _ledger("Round Off", "Other Expenses", vatdealernature="Invoice Rounding"),
    _ledger("Aarav Salary", "Salary"),
    _ledger("Gold Supplier", "Sundry Creditors", incometaxnumber="AAAPG1234C"),
    _ledger("Karigar", "Sundry Creditors", incometaxnumber="ABCPK1234D"),
    _ledger("CAP- Partner One", "Capital Account"),
    _ledger("CAP- Partner Two", "Capital Account"),
    _ledger("Gold Sale", "Sales Accounts"),
    _ledger("Walk-in Customer", "SUNDRY DR"),
    _ledger("Main Bank", "Bank Accounts"),
    # per-party TDS ledgers that name no section, bills booked separately from the TDS journal
    _ledger("TDS on Agent", "TDS"),
    _ledger("TDS on Vendor", "TDS"),
    _ledger("Commission Expenses", "Other Expenses"),
    _ledger("Application Assistance Charges", "Other Expenses"),
    _ledger("Input Igst @18%", "Duties & Taxes"),
    _ledger("Commission Agent", "Sundry Creditors", incometaxnumber="AALFP8511C"),
    _ledger("Service Vendor", "Sundry Creditors"),
])]


def _vch(guid, date, vtype, vno, party, lines, inv=None, optional=False):
    """lines: (ledger, signed Tally amount, isdeemedpositive). Tally: negative amount = debit."""
    raw = {
        "guid": guid, "date": date, "vouchertypename": vtype, "vouchernumber": vno, "partyledgername": party,
        "isoptional": optional,
        "ledgerentries": [{"ledgername": n, "amount": a, "isdeemedpositive": dp} for n, a, dp in lines],
    }
    if inv:
        raw["allinventoryentries"] = [{"amount": a, "isdeemedpositive": True,
                                       "accountingallocations": [{"ledgername": n, "amount": a, "isdeemedpositive": True}]}
                                      for n, a in inv]
    return normalize_transaction(raw, 0)


def _purchase(guid, date, vno, goods, tds=None, optional=False):
    gst = round(goods * 0.015, 2)
    lines = [("Gold Supplier", f"{goods + 2 * gst - (tds or 0):.2f}", False),
             ("Input CGST @ 1.5%", f"-{gst:.2f}", True), ("Input CGST @ 1.5%", f"-{gst:.2f}", True)]
    if tds:
        # Tally invoice mode: the TDS line says ISDEEMEDPOSITIVE=Yes although it is a credit
        lines.append(("TDS ON PURCHASE OF GOODS @ 0.10%", f"{tds:.2f}", True))
    return _vch(guid, date, "Purchase", vno, "Gold Supplier", lines, inv=[("Purchase Gold GST @ 3%", f"-{goods:.2f}")], optional=optional)


def _run(txs):
    ctx = build_tds_context(MASTERS, txs)
    pan = lambda name, tx: (ctx.info(name).pan if ctx.info(name) and ctx.info(name).pan else "") or ""
    return ctx, extract_booked_deductions(txs, ctx, pan), pan


def test_tds_line_in_invoice_is_a_credit_despite_isdeemedpositive():
    tx = _purchase("P1", "20250402", "20", 896329.0, tds=896.33)
    tds_line = next(l for l in tx.ledgers if "TDS" in l.ledger_name)
    assert tds_line.is_credit
    _, deds, _ = _run([tx])
    assert len(deds) == 1
    d = deds[0]
    assert (d.section, d.deductee, d.pan) == ("194Q", "Gold Supplier", "AAAPG1234C")
    # the purchase ledger name contains "GST" but its group makes it the purchase value, not a duty
    assert d.gross == Decimal("896329.00") and d.rate == Decimal("0.10") and d.tds == Decimal("896.33")


def test_optional_voucher_is_not_a_deduction():
    tx = _purchase("P2", "20250402", "21", 3954560.0, tds=3954.56, optional=True)
    assert tx.is_optional and tx.tds_amount == Decimal("0.00")
    assert _run([tx])[1] == []


def test_partner_journal_is_split_per_partner():
    tx = _vch("J1", "20260331", "Journal", "1418", "Aarav Salary", [
        ("CAP- Partner One", "-1003200.00", True), ("CAP- Partner Two", "-250800.00", True), ("Tds 194T", "1254000.00", False)])
    ctx, deds, _ = _run([tx])
    assert is_tds_viable_transaction(tx, None, ctx)
    got = sorted((d.deductee, d.gross, d.tds) for d in deds)
    assert got == [("Partner One", Decimal("10032000.00"), Decimal("1003200.00")),
                   ("Partner Two", Decimal("2508000.00"), Decimal("250800.00"))]
    assert all(d.section == "194T" and d.rate == Decimal("10") for d in deds)


def test_tds_charged_to_party_without_bill():
    tx = _vch("J2", "20251231", "Journal", "233", "Karigar", [("Karigar", "-2051.00", True), ("TDS ON CONTRACT @ 1%", "2051.00", False)])
    ctx, deds, _ = _run([tx])
    assert is_tds_viable_transaction(tx, None, ctx)
    assert [(d.section, d.deductee, d.gross, d.tds) for d in deds] == [("194C", "Karigar", Decimal("205100.00"), Decimal("2051.00"))]


def test_sectionless_tds_ledger_takes_section_from_bill():
    tx = _vch("E1", "20250428", "GST Expenses", "187", "Karigar", [
        ("Karigar", "32581.00", False), ("Labour Charges", "-32910.00", True), ("tds Payable", "329.00", True)])
    deds = _run([tx])[1]
    assert [(d.section, d.gross, d.tds) for d in deds] == [("194C", Decimal("32910.00"), Decimal("329.00"))]


def test_round_off_does_not_make_a_sale_a_tds_event():
    tx = _vch("S1", "20250405", "Sales", "1372", "Walk-in Customer", [
        ("Walk-in Customer", "-33950.00", True), ("Gold Sale", "33951.00", False), ("Round Off", "-1.00", False)])
    ctx = build_tds_context(MASTERS, [_purchase("P3", "20250402", "1", 100000.0, tds=100.0), tx])
    assert ctx.category("Round Off") == "ROUND_OFF"
    assert ctx.category("Labour Charges") == "EXPENSE"          # via the reserved name of "Other Expenses"
    assert ctx.category("Aarav Salary") == "EXPENSE"            # via the revenue flags of a custom primary group
    assert not is_tds_viable_transaction(tx, None, ctx)


def test_rows_grouped_per_deductee_per_month():
    txs = [_purchase("P4", "20250709", "103", 971310.0, tds=971.31), _purchase("P5", "20250717", "104", 975630.0, tds=975.63)]
    rows = group_by_deductee_month(_run(txs)[1])
    assert len(rows) == 1
    r = rows[0]
    assert (r.date, r.voucher_numbers, r.gross, r.computed_tds) == ("2025-07-17", ["103", "104"], Decimal("1946940.00"), Decimal("1947"))


def test_194q_due_above_threshold_but_not_deducted():
    txs = [_purchase("Q1", "20260414", "1", 4000000.0), _purchase("Q2", "20260420", "2", 3000000.0)]
    ctx, booked, pan = _run(txs)
    due = find_194q_not_deducted(txs, ctx, booked, pan)
    # only the purchase value above Rs 50 lakh in the FY: 70L - 50L = 20L @ 0.1%
    assert [(d.voucher_numbers, d.gross, d.tds, d.computed_tds) for d in due] == [(["2"], Decimal("2000000.00"), Decimal("0.00"), Decimal("2000"))]


def test_pan_from_gstin_in_ledger_gst_registration_details():
    m = normalize_master({"metadata": {"type": "Ledger", "name": "Innowit"}, "name": "Innowit", "parent": "Sundry Creditors",
                          "ledgstregdetails": [{"applicablefrom": "20200401", "gstin": "27AAAAA0000A1Z5"},
                                               {"applicablefrom": "20230401", "gstin": "27BQQPR5793E1Z2"}]}, 0)
    assert m.pan == "BQQPR5793E"


def test_base_is_the_lines_the_booked_tds_was_worked_out_on():
    # gold 2,00,925 + stone 5,064 on one bill, TDS booked 201 = 0.1% of the gold line only
    tx = _vch("P6", "20250709", "Purchase", "98", "Gold Supplier", [
        ("Gold Supplier", "211767.00", False), ("Input CGST @ 1.5%", "-6179.68", True),
        ("TDS ON PURCHASE OF GOODS @ 0.10%", "201.00", True)],
        inv=[("Purchase Gold GST @ 3%", "-200925.00"), ("Purchase Gold GST @ 3%", "-5064.00")])
    d = _run([tx])[1][0]
    assert d.gross == Decimal("200925.00") and d.computed_tds == Decimal("201")


def test_entered_date_and_user_listed_per_bill_in_bill_order():
    a = _purchase("E1", "20250709", "103", 971310.0, tds=971.31)
    b = _purchase("E2", "20250717", "104", 975630.0, tds=975.63)
    a.raw_reference.update({"updateddatetime": "20250826103015000", "enteredby": "kiran"})
    b.raw_reference.update({"updateddatetime": "20250718091500000", "enteredby": "RAMZAN"})
    row = group_by_deductee_month(_run([a, b])[1])[0]
    assert row.entries == [("103", "26-08-2025 10:30", "kiran"), ("104", "18-07-2025 09:15", "RAMZAN")]


def _bill(guid, date, vno, party, expense, base, gst=0.0):
    lines = [(party, f"{base + gst:.2f}", False), (expense, f"-{base:.2f}", True)]
    if gst:
        lines.append(("Input Igst @18%", f"-{gst:.2f}", True))
    return _vch(guid, date, "Gst Expenses", vno, party, lines)


def _tds_journal(guid, date, vno, party, tds_ledger, tds):
    return _vch(guid, date, "Journal", vno, party, [(party, f"-{tds:.2f}", True), (tds_ledger, f"{tds:.2f}", False)])


def test_tds_journal_is_linked_to_the_separate_bill_it_was_deducted_on():
    # bill on 30 June, "Being TDS recorded" journal on 2 July to a TDS ledger that names no section
    txs = [_bill("B1", "20250630", "EDUD001", "Commission Agent", "Commission Expenses", 49346.0),
           _tds_journal("J1", "20250702", "65", "Commission Agent", "TDS on Agent", 987.0)]
    ctx, deds, _ = _run(txs)
    assert is_tds_viable_transaction(txs[1], None, ctx)
    assert [(d.section, d.date, d.gross, d.rate, d.tds) for d in deds] == \
        [("194H", "2025-06-30", Decimal("49346.00"), Decimal("2"), Decimal("987.00"))]   # TDS is due on credit: June
    assert deds[0].voucher_numbers == ["65", "EDUD001"]


def test_tds_journal_covering_several_bills_takes_their_total_ex_gst():
    txs = [_bill("B2", "20250411", "HO-21", "Service Vendor", "Application Assistance Charges", 25423.73, gst=4576.27),
           _bill("B3", "20250423", "HO-37", "Service Vendor", "Application Assistance Charges", 25423.73, gst=4576.27),
           _tds_journal("J2", "20250423", "24", "Service Vendor", "TDS on Vendor", 1016.0)]
    deds = _run(txs)[1]
    assert [(d.gross, d.rate, d.date) for d in deds] == [(Decimal("50847.46"), Decimal("2"), "2025-04-23")]


def test_tds_journal_without_its_bill_is_still_reported():
    # bill is in last year's books; the section comes from the party's other bills, amount = TDS / rate
    txs = [_tds_journal("J3", "20250418", "10", "Commission Agent", "TDS on Agent", 40.0),
           _bill("B4", "20250620", "VS/5087", "Commission Agent", "Commission Expenses", 4125.0)]
    deds = [d for d in _run(txs)[1] if d.guids[0] == "J3"]
    assert [(d.section, d.gross, d.rate, d.date) for d in deds] == [("194H", Decimal("2000.00"), Decimal("2"), "2025-04-18")]


def test_booked_rate_is_snapped_to_the_standard_rate_when_pan_is_missing():
    # 1% of 17,207 = 172.07, booked rounded up to 173: the rate is 1%, not 1.01%
    tx = _vch("C1", "20250709", "Journal", "118", "Service Vendor", [
        ("Service Vendor", "17034.00", False), ("Labour Charges", "-17207.00", True), ("tds Payable", "173.00", False)])
    d = _run([tx])[1][0]
    assert (d.section, d.pan, d.rate, d.gross) == ("194C", "", Decimal("1"), Decimal("17207.00"))


def test_194q_journal_is_not_linked_to_purchase_bills():
    # 194Q is on the FY excess over Rs 50 lakh, so a month-end 194Q journal is not tied to particular bills
    txs = [_purchase("Q3", "20250510", "7", 1000000.0),
           _vch("J4", "20250531", "Journal", "8", "Gold Supplier",
                [("Gold Supplier", "-1000.00", True), ("TDS ON PURCHASE OF GOODS @ 0.10%", "1000.00", False)])]
    deds = _run(txs)[1]
    assert [(d.section, d.date, d.gross, d.voucher_numbers) for d in deds] == [("194Q", "2025-05-31", Decimal("1000000.00"), ["8"])]


def _challan(guid, date, vno, tds_ledger, amount):
    return _vch(guid, date, "Payment", vno, "Main Bank", [(tds_ledger, f"-{amount:.2f}", True), ("Main Bank", f"{amount:.2f}", False)])


def test_tds_on_earlier_months_bill_deposited_on_time_stays_in_the_bill_month():
    txs = [_bill("B5", "20250923", "VS/9832", "Commission Agent", "Commission Expenses", 193500.0),
           _tds_journal("J5", "20251006", "172", "Commission Agent", "TDS on Agent", 3870.0),
           _challan("C5", "20251006", "621", "TDS on Agent", 3870.0)]           # by 7 October: on time
    assert [d.date for d in _run(txs)[1] if d.guids[0] == "J5"] == ["2025-09-23"]


def test_tds_on_earlier_months_bill_deposited_late_is_reported_when_deducted():
    txs = [_bill("B6", "20251015", "VS/11330", "Commission Agent", "Commission Expenses", 44713.54),
           _tds_journal("J6", "20251108", "85", "Commission Agent", "TDS on Agent", 894.0),
           _challan("C6", "20251108", "762", "TDS on Agent", 894.0)]            # after 7 November: late
    assert [d.date for d in _run(txs)[1] if d.guids[0] == "J6"] == ["2025-11-08"]


def test_deposit_due_dates():
    from app.services.tds_deductions import _deposit_due_date
    assert [_deposit_due_date(d) for d in ("2025-06-30", "2025-12-15", "2026-03-31")] == ["2025-07-07", "2026-01-07", "2026-04-30"]
