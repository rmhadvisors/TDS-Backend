import pytest
from app.services.tds_statutory_rates import find_matching_rules, suggest_statutory_rule


def _section(text):
    r = suggest_statutory_rule(text)
    return (r.sub_section or r.section_code) if r else None


@pytest.mark.parametrize("text, expected", [
    # "TDS on ..." in a narration no longer hides the nature of the payment
    ("Payment for contract work, TDS on contract deducted", "194C"),
    ("Being TDS on professional fees deducted", "194J(b)"),
    # a TDS account itself is never an expense
    ("TDS on Contract", None),
    ("TDS on Purchase of Goods", None),
    ("TDS - Professional", None),
    ("TDS Payable", None),
    ("Tax Deducted at Source", None),
    ("Interest on TDS", None),
    # names that only start like a keyword
    ("Eduguide Overseas Studies Private Limited", None),
    ("TCS Ltd software services", None),
    ("Rentokil Pest Control", "194C"),
    ("Auditorium hire", None),
    ("AMCO Batteries", None),
    ("Guardian Pharmacy", None),
    # real keywords with their usual endings still match
    ("Office Rent", "194I(b)"),
    ("Rental of premises", "194I(b)"),
    ("Statutory Auditor fees", "194J(b)"),
    ("Audit fees", "194J(b)"),
    ("AMC charges", "194C"),
    ("Security Guards", "194C"),
    ("Logistics charges", "194C"),
    ("Consultancy charges", "194J(b)"),
    # priority
    ("Commission to Director", "194J(b)"),
    ("Sales Commission", "194H"),
    ("Interest on Capital", "194T"),
    # 195 / 206C only on clear wording
    ("Foreign Remittance to vendor", "195"),
    ("Payment to Non-Resident", "195"),
    ("TCS on sale of scrap", "206C"),
    # non-resident / fund sections: the specific section beats 195 / 194A / 194K
    ("Fees to Non-Resident Sportsman", "194E"),
    ("Payment u/s 194EE NSS", None),
    ("Interest on ECB", "194LC"),
    ("Distribution from InvIT", "194LBA"),
    ("Reiterated claim", None),
    ("AIF Distribution", "194LBB"),
    ("Income to Offshore Fund", "196B"),
    ("Dividend on GDRs", "196C"),
    ("FPI income from securities", "196D"),
    # real ledger / narration text from client data
    ("Professional Services 18%", "194J(b)"),
    ("Consulting and Management Services 18%", "194J(b)"),
    ("Software Exp 18%", "194J(a)"),
    ("Vinod Kumar Goa Rent", "194I(b)"),
    ("IMPS-518611131291-EQL FINANCIAL TECHNOLOGIES PRIVATE LIMIT-ICIC-xxxxxxxx7853-India Filings Accounting Contract work", "194C"),
    ("Being TDS Recorded", None),
    ("HDFC Bank Charges", None),
    ("Round Off", None),
])
def test_statutory_section_for_text(text, expected):
    assert _section(text) == expected


def test_ambiguous_text_lists_every_matching_section():
    sections = [r.sub_section or r.section_code for r, _ in find_matching_rules("Rent for transport vehicle")]
    assert "194C" in sections
