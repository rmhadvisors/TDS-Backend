"""
Statutory TDS Section and Rate Registry.

Source: Income Tax Department (TRACES) Non-Salary TDS Rate Chart, Financial Year 2025-26
        (backend/pdf/tdsSectionRate.pdf, "Last reviewed and updated on 31-March-2026").

IMPORTANT – FY 2026-27 onwards:
    From 1 April 2026 the Income-tax Act, 2025 replaces the 1961 Act. Non-salary TDS moves to
    section 393 (salary: 392, TCS: 394) and returns move to new forms (26Q -> Form 140,
    24Q -> Form 138). The section codes below are the 1961-Act codes (194C, 194J ...), which are
    still the common reference in Tally and in practice, but rates/thresholds for FY 2026-27 must be
    re-checked against the 2025 Act / a FY 2026-27 chart — this PDF does not cover that year.

Matching rules:
    * Keywords match at the START of a word ("LEASE" does not match "RELEASE"). Multi-word keywords match as phrases.
      Short keywords in _WHOLE_WORD_ENDINGS must also END the word, with only the listed endings ("RENT" matches
      "RENTAL" but not "RENTOKIL"; "AUDIT" matches "AUDITOR" but not "AUDITORIUM").
    * A text that is a TDS account itself ("TDS on Contract", "TDS Payable") matches nothing; "TDS on ..." inside a
      narration does not block the match.
    * Each rule can carry exclude_keywords; a text that hits an exclusion is not treated as that section.
    * No client/party-specific names in keywords — map known parties separately.
"""

from dataclasses import dataclass, field
from decimal import Decimal
from typing import Dict, List, Optional, Tuple
import re


@dataclass
class StatutoryTDSRule:
    section_code: str
    sub_section: Optional[str]
    description: str
    nature_keywords: List[str]
    threshold_inr: Optional[Decimal]
    rate_company_firm: str
    rate_individual_huf: str
    rate_no_pan: str
    standard_rate_display: str
    notes: str = ""
    exclude_keywords: List[str] = field(default_factory=list)
    threshold_basis: str = ""          # how threshold_inr is applied (per bill / per FY / per month ...)
    priority: int = 50                 # lower = checked first when several sections match


# Words that mean the ledger is NOT a TDS-able expense payment, whatever else it contains.
# ("TDS on <nature>" is not here: it names a TDS account only at the start of the text — see _TDS_ACCOUNT — while
#  a narration such as "contract work, TDS on contract deducted" still tells the nature of the payment.)
GLOBAL_EXCLUDE_KEYWORDS: List[str] = [
    "TDS PAYABLE", "TDS RECEIVABLE", "INTEREST ON TDS", "LATE FEE", "234E",
    "INTEREST ON GST", "GST LATE", "INTEREST ON INCOME TAX", "INTEREST RECEIVED", "INTEREST INCOME",
    "BANK CHARGES", "OD FACILITY", "CC FACILITY", "CASH CREDIT", "OVERDRAFT",
    "CURRENT ACCOUNT", "SAVINGS ACCOUNT", "CURRENT ASSETS", "CURRENT LIABILITIES",
    "PROFIT ON", "LOSS ON", "DISCOUNT RECEIVED", "ROUND OFF", "ROUNDING OFF",
]


STATUTORY_TDS_REGISTRY: List[StatutoryTDSRule] = [
    # ---------------------------------------------------------------- 194T (before 194A: partner interest)
    StatutoryTDSRule(
        section_code="194T", sub_section=None,
        description="Payment of Salary, Remuneration, Commission, Bonus or Interest to Partner of Firm",
        nature_keywords=[
            "194T", "PARTNER REMUNERATION", "PARTNERS REMUNERATION", "REMUNERATION TO PARTNER",
            "INTEREST ON CAPITAL", "INTEREST ON PARTNER", "INTEREST TO PARTNER", "PARTNER INTEREST",
            "PARTNERS INTEREST", "PARTNER SALARY", "PARTNERS SALARY", "SALARY TO PARTNER",
            "PARTNER BONUS", "PARTNER COMMISSION", "COMMISSION TO PARTNER",
        ],
        threshold_inr=Decimal("20000.00"), threshold_basis="aggregate per partner per FY",
        rate_company_firm="10.00%", rate_individual_huf="10.00%", rate_no_pan="20.00%",
        standard_rate_display="10.00%",
        notes="10% on salary/remuneration/commission/bonus/interest to partners u/s 194T (w.e.f. 1-Apr-2025). Drawings / capital withdrawals are not covered.",
        exclude_keywords=["DRAWING", "CAPITAL WITHDRAWAL"], priority=5,
    ),
    # ---------------------------------------------------------------- 194Q
    StatutoryTDSRule(
        section_code="194Q", sub_section=None,
        description="TDS on Purchase of Goods exceeding Rs. 50 Lakhs",
        nature_keywords=["PURCHASE OF GOODS", "PURCHASE GOODS", "TDS ON PURCHASE OF GOODS", "TDS ON GOODS",
                         "TDS ON PURCHASE", "TDS-194Q", "TDS 194Q", "194Q", "TDS ON RAW MATERIAL"],
        threshold_inr=Decimal("5000000.00"), threshold_basis="on value in excess of Rs 50L per seller per FY (buyer turnover > Rs 10 Cr in preceding FY)",
        rate_company_firm="0.10%", rate_individual_huf="0.10%", rate_no_pan="5.00%",
        standard_rate_display="0.10%",
        notes="0.10% on value exceeding Rs 50L (5% if no PAN — 2nd proviso to s.206AA)", priority=10,
    ),
    # ---------------------------------------------------------------- 194H
    StatutoryTDSRule(
        section_code="194H", sub_section=None,
        description="Commission / Brokerage Charges",
        nature_keywords=["COMMISSION", "COMISSION", "COMMISION", "BROKERAGE", "BROKER", "DEL CREDERE",
                         "CHANNEL PARTNER", "SALES COMMISSION", "AGENT COMMISSION", "C&F COMMISSION",
                         "CNF COMMISSION", "CARRYING AND FORWARDING AGENT"],
        threshold_inr=Decimal("20000.00"), threshold_basis="aggregate per payee per FY",
        rate_company_firm="2.00%", rate_individual_huf="2.00%", rate_no_pan="20.00%",
        standard_rate_display="2.00%",
        notes="2% w.e.f. 1st Oct 2024 onwards (5% prior to Oct 2024). Commission to a non-resident / overseas agent is s.195, not 194H. Insurance commission is 194D, lottery 194G.",
        exclude_keywords=["INSURANCE COMMISSION", "LOTTERY", "OVERSEAS", "FOREIGN", "EXPORT COMMISSION", "PARTNER",
                          "DIRECTOR"],   # commission to a director is 194J(b)
        priority=20,
    ),
    # ---------------------------------------------------------------- 194J(b) before 194J(a)
    StatutoryTDSRule(
        section_code="194J", sub_section="194J(b)",
        description="Fees for Professional Services / Royalty / Director fees / Non-compete",
        nature_keywords=[
            "PROFESSIONAL", "LEGAL", "ADVOCATE", "LAWYER", "SOLICITOR", "CONSULT", "AUDIT", "TAX AUDIT",
            "CHARTERED ACCOUNTANT", "CA FEES", "COMPANY SECRETARY", "CS FEES", "COST ACCOUNTANT",
            "ACCOUNTING CHARGES", "BOOKKEEPING", "RETAINER", "RETAINERSHIP", "ARCHITECT", "INTERIOR DESIGN",
            "VALUATION", "VALUER", "SURVEYOR", "DOCTOR", "MEDICAL SERVICE", "MEDICAL PROFESSIONAL",
            "DIRECTOR SITTING", "SITTING FEE", "DIRECTOR REMUNERATION", "DIRECTOR FEE", "COMMISSION TO DIRECTOR",
            "ROYALTY", "SOFTWARE LICENSE", "SOFTWARE LICENCE", "LICENSE FEE", "LICENCE FEE", "NON-COMPETE",
            "NON COMPETE", "FILM ARTIST", "ARTIST FEE",
        ],
        threshold_inr=Decimal("50000.00"), threshold_basis="aggregate per category per payee per FY (no threshold for director fees)",
        rate_company_firm="10.00%", rate_individual_huf="10.00%", rate_no_pan="20.00%",
        standard_rate_display="10.00%",
        notes="10% for Professional services / Royalty (incl. software licence) / Director fees u/s 194J(b). Director salary under employment is s.192.",
        exclude_keywords=["DIRECTOR SALARY", "LEAVE AND LICENSE", "LEAVE & LICENSE", "LEAVE AND LICENCE", "LEAVE & LICENCE"],
        priority=25,
    ),
    StatutoryTDSRule(
        section_code="194J", sub_section="194J(a)",
        description="Fees for Technical Services / Call Center Services",
        nature_keywords=["TECHNICAL SERVICE", "TECHNICAL FEE", "TECHNICAL CHARGE", "TECHNICAL SUPPORT", "CALL CENTER",
                         "CALL CENTRE", "IT SERVICE", "IT SUPPORT", "IT CHARGES", "SOFTWARE DEVELOPMENT",
                         "SOFTWARE SUPPORT", "SOFTWARE MAINTENANCE", "DEVELOPMENT CHARGE", "WEBSITE DEVELOPMENT",
                         "APP DEVELOPMENT", "ENGINEERING SERVICE", "ENGINEERING CHARGE", "INFORMATION SECURITY", "CYBER",
                         "TESTING CHARGE", "TESTING FEE", "LAB TESTING", "ANALYTICAL CHARGE", "ANALYSIS CHARGE",
                         "CALIBRATION", "VALIDATION CHARGE", "CLINICAL TRIAL", "BIOEQUIVALENCE", "BA/BE",
                         "STABILITY STUDY", "R&D SERVICE", "CERTIFICATION CHARGE",
                         "SOFTWARE SUBSCRIPTION", "WEB HOSTING", "HOSTING CHARGE", "HOSTING SERVICE",
                         "SOFTWARE EXP", "SOFTWARE EXPENSE", "SOFTWARE EXPENSES", "SOFTWARE CHARGE", "SOFTWARE CHARGES",
                         "SAAS", "CLOUD SERVICE", "CLOUD SUBSCRIPTION"],
        threshold_inr=Decimal("50000.00"), threshold_basis="aggregate per category per payee per FY",
        rate_company_firm="2.00%", rate_individual_huf="2.00%", rate_no_pan="20.00%",
        standard_rate_display="2.00%",
        notes="2% for Technical services / call centre / software & cloud subscriptions u/s 194J(a). Software LICENCE is royalty -> 194J(b); software PURCHASE as goods -> 194Q.",
        priority=30,
    ),
    # ---------------------------------------------------------------- 194I(a) before 194I(b)
    StatutoryTDSRule(
        section_code="194I", sub_section="194I(a)",
        description="Rent on Plant / Machinery / Equipment Hire",
        nature_keywords=["EQUIPMENT HIRE", "MACHINERY HIRE", "MACHINE HIRE", "MACHINE RENT", "MACHINERY RENT",
                         "PLANT & MACHINERY", "PLANT AND MACHINERY", "CRANE HIRE", "FORKLIFT HIRE", "EQUIPMENT RENT",
                         "EQUIPMENT LEASE", "GENERATOR HIRE", "GENERATOR RENT", "DG SET HIRE", "DG SET RENT",
                         "SELF DRIVE", "WITHOUT DRIVER", "LAPTOP RENT", "COMPUTER RENT"],
        threshold_inr=Decimal("50000.00"), threshold_basis="per month or part of a month (w.e.f. 1-Apr-2025)",
        rate_company_firm="2.00%", rate_individual_huf="2.00%", rate_no_pan="20.00%",
        standard_rate_display="2.00%",
        notes="2% for Plant, Machinery, Equipment rent u/s 194I(a). Vehicle hire WITH driver / transport service is 194C.",
        priority=35,
    ),
    StatutoryTDSRule(
        section_code="194I", sub_section="194I(b)",
        description="Rent on Land / Building / Furniture / Fittings",
        nature_keywords=["RENT", "LEASE RENT", "LEASE RENTAL", "PREMISES", "OFFICE RENT", "WAREHOUSE RENT",
                         "GODOWN RENT", "GODOWN", "WAREHOUSING", "STORAGE CHARGE", "FURNITURE RENT", "BUILDING RENT",
                         "SHOP RENT", "FACTORY RENT", "GUEST HOUSE RENT", "LEAVE AND LICENSE", "LEAVE & LICENSE",
                         "LEAVE AND LICENCE", "LEAVE & LICENCE", "LAND LEASE"],
        threshold_inr=Decimal("50000.00"), threshold_basis="per month or part of a month (w.e.f. 1-Apr-2025)",
        rate_company_firm="10.00%", rate_individual_huf="10.00%", rate_no_pan="20.00%",
        standard_rate_display="10.00%",
        notes="10% for Land, Building, Furniture rent u/s 194I(b). Rent paid by individuals/HUF not under audit -> 194IB.",
        exclude_keywords=["RENT A CAB", "RENT-A-CAB", "CAB RENT", "CAR RENT", "VEHICLE RENT", "TAXI", "RENT RECEIVED",
                          "RENTAL INCOME", "RENT INCOME"],
        priority=40,
    ),
    # ---------------------------------------------------------------- 194C (broad, after the specific ones)
    StatutoryTDSRule(
        section_code="194C", sub_section=None,
        description="Payment to Contractors / Transporters / Manpower / Jobwork",
        nature_keywords=[
            "CONTRACT", "SUB-CONTRACT", "SUBCONTRACT", "WORKS CONTRACT", "CIVIL WORK", "CONSTRUCTION",
            "TRANSPORT", "FREIGHT", "CARTAGE", "CARRIAGE", "CARTING", "TEMPO", "TRUCK", "LORRY", "TRAILER",
            "LOGISTIC", "COURIER", "SHIPPING", "CLEARING", "FORWARDING", "CHA CHARGES", "CUSTOM CLEARING",
            "DELIVERY CHARGE", "PICKUP CHARGE", "VEHICLE HIRE", "VEHICLE RENT", "CAR HIRE", "CAR RENT", "CAB HIRE", "CAB RENT", "BUS HIRE", "TAXI",
            "RENT A CAB", "LOADING", "UNLOADING", "HAMALI", "HAMMALI", "PACKING CHARGE", "PACKAGING CHARGE",
            "SECURITY GUARD", "SECURITY SERVICE", "SECURITY AGENCY", "SECURITY CHARGE", "WATCH AND WARD", "GUARD",
            "MANPOWER", "MAN POWER", "CONTRACT LABOUR", "HOUSEKEEPING", "HOUSE KEEPING", "CLEANING", "PEST CONTROL",
            "GARDENING", "FACILITY MANAGEMENT", "FACILITY SERVICE", "LABOUR", "LABOR", "JOB WORK", "JOBWORK",
            "JOB CHARGE", "PROCESSING CHARGE", "CONVERSION CHARGE", "MAKING CHARGE", "MAKING CHARGES",
            "POLISHING", "PLATING", "STITCHING", "TAILORING", "DYEING", "PRINTING CHARGE", "PRINTING",
            "FABRICATION", "REPAIR", "MAINTENANCE", "AMC", "ADVERTISING", "ADVERTISEMENT", "PUBLICITY",
            "HOARDING", "BROADCASTING", "TELECASTING", "CATERING", "CANTEEN", "EVENT", "EVENT MANAGEMENT",
        ],
        threshold_inr=Decimal("100000.00"), threshold_basis="Rs 30,000 single bill OR Rs 1,00,000 aggregate per payee per FY",
        rate_company_firm="2.00%", rate_individual_huf="1.00%", rate_no_pan="20.00%",
        standard_rate_display="1.00% / 2.00%",
        notes=("1% for Individual/HUF (PAN 4th char P/H), 2% for others (C/F/A/B/T/G/L/J). "
               "Single-bill limit Rs 30,000 u/s 194C(5) is not shown on the TRACES chart but still applies. "
               "No TDS for transporters owning <=10 goods carriages who give PAN + declaration (194C(6)). "
               "Software/IT 'maintenance' and AMC of software -> 194J; check before applying."),
        exclude_keywords=["SAFEGUARD", "LABORATOR", "LABOURATOR", "SOFTWARE MAINTENANCE", "SOFTWARE AMC", "PRINTING & STATIONERY",
                          "PRINTING AND STATIONERY", "STATIONERY", "FREIGHT INWARD ON PURCHASE INCLUDED",
                          "OVERSEAS", "FOREIGN", "OCEAN FREIGHT - NON RESIDENT"],
        priority=60,
    ),
    # ---------------------------------------------------------------- 194R
    StatutoryTDSRule(
        section_code="194R", sub_section=None,
        description="Benefit or Perquisite in respect of Business / Profession",
        nature_keywords=["PERQUISITE", "BUSINESS BENEFIT", "DEALER INCENTIVE", "GIFT TO DEALER", "DEALER GIFT",
                         "TRADE INCENTIVE", "INCENTIVE TOUR", "DEALER TOUR", "FREE SAMPLE", "PHYSICIAN SAMPLE",
                         "DOCTOR SAMPLE", "SPONSORSHIP TO DOCTOR", "CONFERENCE SPONSORSHIP", "194R"],
        threshold_inr=Decimal("20000.00"), threshold_basis="aggregate value per recipient per FY",
        rate_company_firm="10.00%", rate_individual_huf="10.00%", rate_no_pan="20.00%",
        standard_rate_display="10.00%",
        notes="10% of value of business benefit/perquisite u/s 194R (incl. free samples to doctors, dealer gifts/tours). Employee benefits are s.192, not 194R.",
        exclude_keywords=["EMPLOYEE", "STAFF WELFARE", "GRATUITY", "PF ", "ESIC"], priority=45,
    ),
    # ---------------------------------------------------------------- 194A (after 194T)
    StatutoryTDSRule(
        section_code="194A", sub_section=None,
        description="Interest other than Interest on Securities",
        nature_keywords=["INTEREST ON LOAN", "INTEREST ON UNSECURED LOAN", "INTEREST ON DEPOSIT", "INTEREST ON ICD",
                         "ICD INTEREST", "LOAN INTEREST", "INTEREST PAID", "INTEREST PAYABLE", "NBFC INTEREST",
                         "FINANCE CHARGE", "INTEREST ON BORROWING", "INTEREST EXPENSE", "INTEREST"],
        threshold_inr=Decimal("10000.00"),
        threshold_basis="Rs 10,000 per payee per FY (others); Rs 50,000 if payer is bank/co-op bank/post office; Rs 1,00,000 to senior citizens by such payers",
        rate_company_firm="10.00%", rate_individual_huf="10.00%", rate_no_pan="20.00%",
        standard_rate_display="10.00%",
        notes="10% u/s 194A. No TDS on interest paid TO banks/financial institutions (194A(3)(iii)). Interest to partners is 194T.",
        exclude_keywords=["BANK INTEREST", "INTEREST TO BANK", "INTEREST ON CC", "INTEREST ON OD", "INTEREST ON CAR LOAN",
                          "INTEREST ON VEHICLE LOAN", "INTEREST ON TERM LOAN", "INTEREST ON SECURITIES", "INTEREST ON CAPITAL",
                          "PARTNER", "DEBENTURE", "BOND INTEREST"],
        priority=70,
    ),
    # ---------------------------------------------------------------- others from the chart
    StatutoryTDSRule("194S", None, "Payment on Transfer of Virtual Digital Asset (VDA / Crypto)",
                     ["VIRTUAL DIGITAL ASSET", "CRYPTO", "VDA"], Decimal("10000.00"), "1.00%", "1.00%", "20.00%", "1.00%",
                     "1% u/s 194S (threshold Rs 50,000 for specified persons)", threshold_basis="aggregate per FY", priority=80),
    StatutoryTDSRule("194IA", None, "Transfer of Immovable Property other than agricultural land",
                     ["IMMOVABLE PROPERTY PURCHASE", "PROPERTY TRANSFER", "PURCHASE OF PROPERTY", "PURCHASE OF LAND",
                      "PURCHASE OF FLAT", "PURCHASE OF BUILDING"], Decimal("5000000.00"), "1.00%", "1.00%", "20.00%", "1.00%",
                     "1% on consideration >= Rs 50L u/s 194-IA", threshold_basis="per property (consideration or stamp value)", priority=80),
    StatutoryTDSRule("194IB", None, "Payment of Rent by Individuals / HUF (> Rs 50,000 per month)",
                     ["INDIVIDUAL RENT", "HUF RENT"], Decimal("50000.00"), "2.00%", "2.00%", "20.00%", "2.00%",
                     "2% w.e.f. 1-Oct-2024 u/s 194-IB (payer: individual/HUF not liable to audit)", threshold_basis="per month", priority=80),
    StatutoryTDSRule("194IC", None, "Payment under Specified Agreement (Joint Development Agreement)",
                     ["JOINT DEVELOPMENT", "JDA", "SPECIFIED AGREEMENT", "194IC"], None, "10.00%", "10.00%", "20.00%", "10.00%",
                     "10% u/s 194-IC, no threshold", priority=80),
    StatutoryTDSRule("194LA", None, "Compensation on Compulsory Acquisition of Immovable Property",
                     ["COMPULSORY ACQUISITION", "LAND ACQUISITION", "ACQUISITION COMPENSATION"], Decimal("500000.00"),
                     "10.00%", "10.00%", "20.00%", "10.00%", "10% u/s 194LA", threshold_basis="aggregate per FY", priority=80),
    StatutoryTDSRule("194M", None, "Payment of certain sums by Individuals or HUF",
                     ["194M", "PERSONAL CONTRACT"], Decimal("5000000.00"), "2.00%", "2.00%", "20.00%", "2.00%",
                     "2% w.e.f. 1-Oct-2024 u/s 194M", threshold_basis="aggregate per payee per FY", priority=80),
    StatutoryTDSRule("194N", None, "Cash Withdrawal exceeding Rs 1 Crore",
                     ["CASH WITHDRAWAL", "194N"], Decimal("10000000.00"), "2.00%", "2.00%", "20.00%", "2.00%",
                     "2% on withdrawal > Rs 1 Cr (Rs 3 Cr for co-op societies); non-filers: 2% above Rs 20L, 5% above Rs 1 Cr (194NF)",
                     threshold_basis="aggregate per bank account per FY", priority=80),
    StatutoryTDSRule("194O", None, "TDS on E-commerce Participants",
                     ["E-COMMERCE", "ECOMMERCE", "MARKETPLACE SETTLEMENT", "AMAZON SELLER", "FLIPKART SETTLEMENT"],
                     Decimal("500000.00"), "0.10%", "0.10%", "5.00%", "0.10%",
                     "0.10% w.e.f. 1-Oct-2024 (5% if no PAN); Rs 5L threshold only for individual/HUF participants",
                     threshold_basis="gross sales per FY (individual/HUF only)", priority=80),
    StatutoryTDSRule("194D", None, "Insurance Commission",
                     ["INSURANCE COMMISSION"], Decimal("20000.00"), "10.00%", "2.00%", "20.00%", "2.00% / 10.00%",
                     "2% for Individual/HUF, 10% for others u/s 194D", threshold_basis="aggregate per FY", priority=15),
    StatutoryTDSRule("194DA", None, "Payment in respect of Life Insurance Policy",
                     ["LIFE INSURANCE POLICY", "LIC MATURITY", "POLICY MATURITY"], Decimal("100000.00"), "2.00%", "2.00%", "20.00%", "2.00%",
                     "2% on income component w.e.f. 1-Oct-2024 u/s 194DA", priority=80),
    StatutoryTDSRule("194G", None, "Commission on Lottery Tickets",
                     ["LOTTERY COMMISSION"], Decimal("20000.00"), "2.00%", "2.00%", "20.00%", "2.00%",
                     "2% w.e.f. 1-Oct-2024 u/s 194G", priority=15),
    StatutoryTDSRule("194B", None, "Winnings from Lotteries, Crossword Puzzles",
                     ["LOTTERY WINNING", "CROSSWORD", "PRIZE MONEY", "LUCKY DRAW"], Decimal("10000.00"), "30.00%", "30.00%", "30.00%", "30.00%",
                     "30% u/s 194B (Rs 10,000 per transaction)", priority=80),
    StatutoryTDSRule("194BA", None, "Winnings from Online Games",
                     ["ONLINE GAME", "194BA"], None, "30.00%", "30.00%", "30.00%", "30.00%",
                     "30% on net winnings u/s 194BA, no threshold", priority=80),
    StatutoryTDSRule("194BB", None, "Winnings from Horse Races",
                     ["HORSE RACE"], Decimal("10000.00"), "30.00%", "30.00%", "30.00%", "30.00%", "30% u/s 194BB", priority=80),
    StatutoryTDSRule("194EE", None, "Payments out of deposits under NSS",
                     ["NSS WITHDRAWAL", "NATIONAL SAVINGS SCHEME"], Decimal("2500.00"), "10.00%", "10.00%", "20.00%", "10.00%",
                     "10% u/s 194EE", priority=80),
    StatutoryTDSRule("194K", None, "Income in respect of units of Mutual Funds",
                     ["MUTUAL FUND DIVIDEND", "MF DIVIDEND", "INCOME FROM UNITS"], Decimal("10000.00"), "10.00%", "10.00%", "20.00%", "10.00%",
                     "10% u/s 194K", priority=80),
    StatutoryTDSRule("193", None, "Interest on Securities",
                     ["INTEREST ON SECURITIES", "DEBENTURE INTEREST", "INTEREST ON DEBENTURE", "BOND INTEREST"],
                     Decimal("10000.00"), "10.00%", "10.00%", "20.00%", "10.00%", "10% u/s 193", priority=65),
    StatutoryTDSRule("194", None, "Dividends",
                     ["DIVIDEND"], Decimal("10000.00"), "10.00%", "10.00%", "20.00%", "10.00%", "10% u/s 194",
                     exclude_keywords=["MUTUAL FUND", "DIVIDEND RECEIVED", "DIVIDEND INCOME"], priority=80),
    StatutoryTDSRule("192A", None, "Payment of accumulated PF balance to an employee",
                     ["PF WITHDRAWAL", "ACCUMULATED BALANCE", "192A"], Decimal("50000.00"), "-", "10.00%", "30.00%", "10.00%",
                     "10% u/s 192A (30% if no PAN per chart)", priority=80),
    StatutoryTDSRule("192", None, "Salaries / Payroll TDS",
                     ["PAYROLL TDS", "TDS ON SALARY", "DIRECTOR SALARY", "SALARY", "WAGES"], None,
                     "Average Slab Rate", "Average Slab Rate", "20.00%", "Slab Rate",
                     "Average slab rate u/s 192 (NOT in the TRACES non-salary chart). Partner salary is 194T.",
                     exclude_keywords=["PARTNER"], priority=8),
    StatutoryTDSRule("195", None, "Payment to Non-Resident / Foreign Remittance",
                     # not bare "OVERSEAS"/"FOREIGN": they are in many Indian names ("XYZ Overseas Studies Pvt Ltd")
                     ["NON-RESIDENT", "NON RESIDENT", "FOREIGN REMITTANCE", "FOREIGN PAYMENT", "FOREIGN VENDOR",
                      "FOREIGN PARTY", "FOREIGN SUPPLIER", "FOREIGN SERVICE", "OVERSEAS REMITTANCE", "OVERSEAS PAYMENT",
                      "OVERSEAS VENDOR", "OVERSEAS SUPPLIER", "OUTWARD REMITTANCE", "IMPORT OF SERVICE", "OFFSHORE"], None,
                     "Rates in Force / DTAA", "Rates in Force / DTAA", "20/30/40%", "DTAA / Rates in Force",
                     "Average/applicable rates or DTAA u/s 195; no-PAN rate per chart 20/30/40 (s.206AA(7) relief if TRC etc. furnished)",
                     priority=1),
    StatutoryTDSRule("206C", None, "Tax Collected at Source (TCS)",
                     # not bare "TCS" (also a company name): only as a tax line
                     ["206C", "TCS ON", "TCS PAYABLE", "TCS RECEIVABLE", "TCS COLLECTED", "TCS @", "TAX COLLECTED AT SOURCE"], None,
                     "0.10% / 1.00% / 5.00%", "0.10% / 1.00% / 5.00%", "5.00%", "0.10% - 5.00%",
                     "TCS u/s 206C — not part of the TDS chart. 206C(1H) on sale of goods omitted w.e.f. 1-Apr-2025.", priority=2),
    # ---------------------------------------------------------------- non-resident / fund sections
    # Only on specific wording; priority 0 where the text also holds a 195 / 194A / 194K word ("NON-RESIDENT",
    # "OFFSHORE", "INTEREST ON", "INCOME FROM UNITS") so the specific section wins.
    # Not added: 194LD (omitted w.e.f. 1-Jul-2023), 194F (omitted w.e.f. 1-Oct-2024), 194P (deducted by the bank).
    StatutoryTDSRule("194E", None, "Payment to Non-Resident Sportsmen / Sports Associations / Entertainers",
                     ["NON-RESIDENT SPORTSMAN", "NON RESIDENT SPORTSMAN", "NON-RESIDENT SPORTSMEN", "NON RESIDENT SPORTSMEN",
                      "NON-RESIDENT ENTERTAINER", "NON RESIDENT ENTERTAINER", "NON-RESIDENT ATHLETE", "NON RESIDENT ATHLETE",
                      "NON-RESIDENT SPORTS ASSOCIATION", "NON RESIDENT SPORTS ASSOCIATION", "FOREIGN SPORTSMAN",
                      "FOREIGN SPORTSMEN", "FOREIGN ATHLETE", "FOREIGN PLAYER", "FOREIGN ENTERTAINER", "OVERSEAS PLAYER",
                      "194E"], None, "20.00%", "20.00%", "20.00%", "20.00%",
                     "20% (plus surcharge & cess) u/s 194E, no threshold", priority=0),
    StatutoryTDSRule("194LB", None, "Interest from Infrastructure Debt Fund to Non-Resident",
                     ["INFRASTRUCTURE DEBT FUND", "INFRA DEBT FUND", "194LB"], None, "5.00%", "5.00%", "20.00%", "5.00%",
                     "5% u/s 194LB (payee non-resident)", priority=0),
    StatutoryTDSRule("194LBA", None, "Income from units of a Business Trust (REIT / InvIT)",
                     ["BUSINESS TRUST", "REAL ESTATE INVESTMENT TRUST", "INFRASTRUCTURE INVESTMENT TRUST", "REIT", "INVIT",
                      "194LBA"], None, "10.00%", "10.00%", "20.00%", "10.00%",
                     "Resident: 10% u/s 194LBA(1). Non-resident: interest 5%, dividend 10%, rent at rates in force.",
                     priority=0),
    StatutoryTDSRule("194LBB", None, "Income in respect of units of Investment Fund (AIF)",
                     ["ALTERNATIVE INVESTMENT FUND", "AIF INCOME", "AIF DISTRIBUTION", "INVESTMENT FUND UNITS", "194LBB"], None,
                     "10.00%", "10.00%", "20.00%", "10.00%", "Resident: 10% u/s 194LBB; non-resident: rates in force",
                     priority=0),
    StatutoryTDSRule("194LBC", None, "Income in respect of investment in Securitisation Trust",
                     ["SECURITISATION TRUST", "SECURITIZATION TRUST", "PASS THROUGH CERTIFICATE", "PASS-THROUGH CERTIFICATE",
                      "PTC INCOME", "194LBC"], None, "10.00%", "10.00%", "20.00%", "10.00%",
                     "Resident: 10% u/s 194LBC (w.e.f. 1-Apr-2023); non-resident: rates in force", priority=0),
    StatutoryTDSRule("194LC", None, "Interest from Indian company on foreign currency borrowings / bonds (ECB)",
                     ["EXTERNAL COMMERCIAL BORROWING", "ECB INTEREST", "INTEREST ON ECB", "FOREIGN CURRENCY LOAN",
                      "FOREIGN CURRENCY BORROWING", "MASALA BOND", "RUPEE DENOMINATED BOND", "194LC"], None,
                     "5.00%", "5.00%", "20.00%", "5.00%",
                     "5% u/s 194LC (4% on bonds listed on an IFSC exchange); payee non-resident", priority=0),
    StatutoryTDSRule("196A", None, "Income in respect of units of Non-Residents",
                     ["NON-RESIDENT UNITHOLDER", "NON RESIDENT UNITHOLDER", "NON-RESIDENT UNIT HOLDER",
                      "NON RESIDENT UNIT HOLDER", "UNITS TO NON-RESIDENT", "UNITS TO NON RESIDENT", "196A"], None,
                     "20.00%", "20.00%", "20.00%", "20.00%", "20% u/s 196A (or DTAA rate)", priority=0),
    StatutoryTDSRule("196B", None, "Income from units / LTCG on units to an Offshore Fund",
                     ["OFFSHORE FUND", "196B"], None, "10.00%", "10.00%", "20.00%", "10.00% / 12.50%",
                     "10% on income, 12.5% on LTCG (w.e.f. 23-Jul-2024) u/s 196B", priority=0),
    StatutoryTDSRule("196C", None, "Income from foreign currency bonds / GDRs",
                     ["GLOBAL DEPOSITORY RECEIPT", "GLOBAL DEPOSITARY RECEIPT", "GDR", "FOREIGN CURRENCY BOND",
                      "FOREIGN CURRENCY CONVERTIBLE BOND", "FCCB", "196C"], None, "10.00%", "10.00%", "20.00%", "10.00% / 12.50%",
                     "10% on interest/dividend, 12.5% on LTCG (w.e.f. 23-Jul-2024) u/s 196C", priority=0),
    StatutoryTDSRule("196D", None, "Income of Foreign Institutional / Portfolio Investors from securities",
                     ["FOREIGN INSTITUTIONAL INVESTOR", "FOREIGN PORTFOLIO INVESTOR", "FII", "FPI", "196D"], None,
                     "20.00%", "20.00%", "20.00%", "20.00%",
                     "20% u/s 196D (10% for specified funds u/s 196D(1A)); no TDS on capital gains", priority=0),
]

# Bare section code -> default sub-section when only "194J"/"194I" is known (the common case, higher rate).
DEFAULT_SUB_SECTION: Dict[str, str] = {"194J": "194J(b)", "194I": "194I(b)"}


# --------------------------------------------------------------------------- matching helpers
# Short keywords that are also the start of unrelated words / names must end the word, allowing only these endings
# ("RENT" -> RENT, RENTS, RENTAL, not RENTOKIL; "AUDIT" -> AUDITOR, not AUDITORIUM; "AMC" not AMCO; "GUARD" not GUARDIAN).
_WHOLE_WORD_ENDINGS: Dict[str, str] = {
    "RENT": "S|AL|ALS|ED|ING", "AUDIT": "S|OR|ORS|ING", "TAX AUDIT": "S|OR|ORS", "AMC": "S", "GUARD": "S|ING",
    "VDA": "S", "JDA": "S", "SAAS": "", "TAXI": "S", "TEMPO": "S", "TRUCK": "S|ING", "EVENT": "S",
    "DOCTOR": "S", "CRYPTO": "S|CURRENCY|CURRENCIES",
    # codes that are the start of a longer code ("194E" in "194EE", "194LB" in "194LBA")
    "194E": "", "194LB": "", "194LBA": "", "194LBB": "", "194LBC": "", "194LC": "",
    "196A": "", "196B": "", "196C": "", "196D": "",
    "REIT": "S", "INVIT": "S", "GDR": "S", "FCCB": "S", "FII": "S", "FPI": "S",
}


def _kw_regex(kw: str) -> "re.Pattern[str]":
    kw = kw.strip().upper()
    # keyword must start at a word boundary; a trailing boundary is not required so "LOGISTIC" hits "LOGISTICS"
    endings = _WHOLE_WORD_ENDINGS.get(kw)
    if endings is not None:
        return re.compile(r"(?<![A-Z0-9])" + re.escape(kw) + (f"(?:{endings})?" if endings else "") + r"(?![A-Z0-9])")
    return re.compile(r"(?<![A-Z0-9])" + re.escape(kw))

_COMPILED: Dict[int, Tuple[List["re.Pattern[str]"], List["re.Pattern[str]"]]] = {
    id(r): ([_kw_regex(k) for k in r.nature_keywords], [_kw_regex(k) for k in r.exclude_keywords])
    for r in STATUTORY_TDS_REGISTRY
}
_GLOBAL_EXCL = [_kw_regex(k) for k in GLOBAL_EXCLUDE_KEYWORDS]
# a TDS / TCS account itself ("TDS on Contract", "TDS - Professional", "Tax Deducted at Source"): not an expense
_TDS_ACCOUNT = re.compile(r"^\W*(TDS|TAX\s+DEDUCTED)(?![A-Z0-9])")


def _hits(patterns, text_upper) -> List[str]:
    return [p.pattern for p in patterns if p.search(text_upper)]


def find_matching_rules(text: Optional[str]) -> List[Tuple[StatutoryTDSRule, str]]:
    """
    All rules whose keywords match `text` (party / ledger / narration) and are not excluded,
    ordered by rule priority. Returns [(rule, matched_keyword), ...]. Empty if a global exclusion hits.
    """
    if not text:
        return []
    u = text.upper()
    if _TDS_ACCOUNT.match(u) or any(p.search(u) for p in _GLOBAL_EXCL):
        return []
    out = []
    for rule in sorted(STATUTORY_TDS_REGISTRY, key=lambda r: r.priority):
        inc, exc = _COMPILED[id(rule)]
        if any(p.search(u) for p in exc):
            continue
        for kw, p in zip(rule.nature_keywords, inc):
            if p.search(u):
                out.append((rule, kw.strip()))
                break
    if not out:
        # keywords users added for natures this list does not know (user_keywords.py) — only where none matched
        from app.services.user_keywords import match_user_rule
        hit = match_user_rule(text)
        if hit:
            out.append(hit)
    return out


def suggest_statutory_rule(text: Optional[str]) -> Optional[StatutoryTDSRule]:
    """Best single rule for `text` (highest priority match), or None."""
    m = find_matching_rules(text)
    return m[0][0] if m else None


def is_statutory_keyword_match(text: Optional[str]) -> bool:
    """
    Checks if a given string (party name, ledger name, narration) matches any statutory TDS nature keywords.
    (Same signature as before; now word-boundary aware and honours exclusions.)
    """
    return bool(find_matching_rules(text))


def is_statutory_tds_party_name(party_name: Optional[str]) -> bool:
    """
    Checks if a party/vendor name matches statutory TDS nature patterns or keywords.
    """
    if not party_name:
        return False
    return is_statutory_keyword_match(party_name)


def get_statutory_rule_by_section(section_code: str) -> Optional[StatutoryTDSRule]:
    """
    Finds a statutory rule by section code (e.g. '194C', '194J', '194J(a)', '194-IA').
    A bare '194J' / '194I' returns 194J(b) / 194I(b) (10%) — previously it returned the 2% sub-section.
    """
    if not section_code:
        return None
    cleaned = section_code.strip().upper().replace("-", "").replace(" ", "")
    cleaned = DEFAULT_SUB_SECTION.get(cleaned, cleaned).upper()
    for rule in STATUTORY_TDS_REGISTRY:
        if rule.sub_section and rule.sub_section.upper() == cleaned:
            return rule
    for rule in STATUTORY_TDS_REGISTRY:
        if rule.section_code.upper() == cleaned:
            return rule
    return None


def rate_for_pan(rule: StatutoryTDSRule, pan: Optional[str]) -> str:
    """Rate string for a payee PAN: no/invalid PAN -> rate_no_pan; 4th char P/H -> individual/HUF; else company/firm."""
    p = (pan or "").strip().upper()
    if not re.fullmatch(r"[A-Z]{5}[0-9]{4}[A-Z]", p):
        return rule.rate_no_pan
    return rule.rate_individual_huf if p[3] in "PH" else rule.rate_company_firm
