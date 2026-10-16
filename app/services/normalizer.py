from collections import defaultdict
from decimal import Decimal
from typing import Any, Dict, List, Optional, Set, Tuple
import re

from app.models.comparison_models import (
    NormalizedLedger,
    NormalizedMaster,
    NormalizedTransaction,
)
from app.services.parser import ValidationError
from app.services.tds_statutory_rates import (
    STATUTORY_TDS_REGISTRY,
    is_statutory_keyword_match,
    is_statutory_tds_party_name,
)
from app.services.tds_context import TDSContext, build_tds_context
from app.utils.amounts import parse_decimal
from app.utils.normalization import is_technical_field, normalize_date, normalize_string


# =============================================================================
# Generic & Client-Configurable Normalizer Rules (No Hardcoded Parties)
# =============================================================================
DEFAULT_PARTY_ALIASES: List[Tuple[str, str, Optional[str]]] = []
DEFAULT_PARTNER_HINTS: List[str] = []

# Party aliases: (regex on party/ledger text, canonical party name, default section if none booked)
PARTY_ALIASES: List[Tuple[str, str, Optional[str]]] = list(DEFAULT_PARTY_ALIASES)

# Extra regexes that pull a supplier-bill reference out of the narration (used to drop re-entered bills).
DUPLICATE_BILL_REF_PATTERNS: List[str] = []

# Optional: partner names, only used if supplied via client configuration
PARTNER_NAME_HINTS: List[str] = list(DEFAULT_PARTNER_HINTS)


def configure_client(
    party_aliases: Optional[List[Tuple[str, str, Optional[str]]]] = None,
    bill_ref_patterns: Optional[List[str]] = None,
    partner_name_hints: Optional[List[str]] = None,
) -> None:
    """Load (or clear/reset) client rules."""
    global PARTY_ALIASES, DUPLICATE_BILL_REF_PATTERNS, PARTNER_NAME_HINTS, _ALIAS_RES, _BILL_REF_RES
    PARTY_ALIASES = list(party_aliases if party_aliases is not None else DEFAULT_PARTY_ALIASES)
    DUPLICATE_BILL_REF_PATTERNS = list(bill_ref_patterns or [])
    PARTNER_NAME_HINTS = [h.upper() for h in (partner_name_hints if partner_name_hints is not None else DEFAULT_PARTNER_HINTS)]
    _ALIAS_RES = [(re.compile(p, re.I), c, d) for p, c, d in PARTY_ALIASES]
    _BILL_REF_RES = [re.compile(p, re.I) for p in DUPLICATE_BILL_REF_PATTERNS]


_ALIAS_RES: List[Tuple["re.Pattern[str]", str, Optional[str]]] = []
_BILL_REF_RES: List["re.Pattern[str]"] = []


def resolve_party_alias(name: Optional[str]) -> Tuple[Optional[str], Optional[str]]:
    """Resolve raw party or ledger name to canonical name and default section."""
    if not name:
        return name, None
    for pat, canon, sec in _ALIAS_RES:
        if pat.search(name):
            return canon, sec
    return name, None


# Dataset-learned context (built by discover_tds_parties / build_tds_context). None = rule-only fallback.
_ACTIVE_CTX: Optional[TDSContext] = None


def set_tds_context(ctx: Optional[TDSContext]) -> None:
    global _ACTIVE_CTX
    _ACTIVE_CTX = ctx


def get_tds_context() -> Optional[TDSContext]:
    return _ACTIVE_CTX


# =============================================================================
# Generic helpers
# =============================================================================
def _get_first_key(d: Dict[str, Any], keys: List[str]) -> Optional[Any]:
    """Helper to find first matching key in dictionary case-insensitively."""
    lower_map = {k.lower(): v for k, v in d.items()}
    for k in keys:
        if k.lower() in lower_map and lower_map[k.lower()] is not None:
            return lower_map[k.lower()]
    return None


def _word_re(words: List[str]) -> "re.Pattern[str]":
    """One regex that matches any of `words` as whole words/phrases."""
    alts = sorted({re.escape(w.strip().upper()) for w in words if w.strip()}, key=len, reverse=True)
    return re.compile(r"(?<![A-Z0-9])(?:" + "|".join(alts) + r")(?![A-Z0-9])")


def _safe_decimal(v: Any) -> Optional[Decimal]:
    """Parse a TDS rate; ignore non-numeric values such as an item rate '500.00/Nos'."""
    if v is None:
        return None
    try:
        return parse_decimal(v, None)
    except (ValueError, ArithmeticError):
        return None


_SECTION_VALUE_RE = re.compile(r"^\s*(SEC(TION)?\s*)?19[2-6][A-Z]{0,3}(\([AB]\))?\s*$", re.I)


def _clean_section(v: Any) -> Optional[str]:
    """Accept only values that look like a TDS section (192..196x)."""
    if v is None:
        return None
    s = str(v).strip()
    return s if s and _SECTION_VALUE_RE.match(s) else None


def _alnum(s: Optional[str]) -> str:
    return re.sub(r"[^A-Z0-9]", "", (s or "").upper())


# =============================================================================
# TDS ledger detection
# =============================================================================
_TDS_NAME_RE = re.compile(r"(?<![A-Z])TDS(?![A-Z])|TAX DEDUCTED")
# TDS-looking ledgers that are NOT a deduction made by us
_TDS_NOT_DEDUCTION_RE = _word_re([
    "TDS RECEIVABLE", "TDS RECOVERABLE", "TDS ON SALES", "TDS BY CUSTOMER", "TDS DEDUCTED BY",
    "INTEREST ON TDS", "LATE FEE", "LATE FILING", "234E", "201(1A)", "TCS", "TDS REFUND",
])


def is_tds_ledger(name: str, entry_dict: Dict[str, Any]) -> bool:
    """
    Determines if a ledger entry represents a TDS deduction ledger (TDS we deduct from a payee).
    'TDS Receivable' (deducted by our customers), interest/late fee on TDS and TCS are excluded.
    """
    if not name:
        return False
    name_upper = name.upper()
    if _TDS_NOT_DEDUCTION_RE.search(name_upper):
        return False
    if _TDS_NAME_RE.search(name_upper):
        return True

    for key in ["istds", "is_tds", "tdsapplicable"]:
        val = _get_first_key(entry_dict, [key])
        if val is not None and str(val).lower() in ["yes", "true", "1"]:
            return True
    if _clean_section(_get_first_key(entry_dict, ["tds_section", "tdssection"])):
        return True
    return False


# =============================================================================
# Non-TDS parties (banks, cash, gateways, sales, duties)
# =============================================================================
# Bank names that are unambiguous on their own (cannot be a vendor name).
_UNAMBIGUOUS_BANK_TOKENS = [
    "HDFC", "ICICI", "KOTAK", "SBI", "IDBI", "INDUSIND", "RBL", "HSBC", "DEUTSCHE", "BCCB", "BASSEIN CATHOLIC",
    "TJSB", "NKGSB", "SARASWAT", "ABHYUDAYA", "SHAMRAO", "SVC BANK", "YES BANK", "STANDARD CHARTERED", "CITIBANK",
    "BANDHAN", "AXIS BANK", "CANARA BANK", "PNB", "BANK OF BARODA", "BANK OF INDIA", "BANK OF MAHARASHTRA",
    "UNION BANK", "INDIAN BANK", "INDIAN OVERSEAS BANK", "UCO BANK", "CENTRAL BANK", "PUNJAB NATIONAL BANK",
    "FEDERAL BANK", "COSMOS BANK", "IDFC FIRST",
]
_BANK_WORD_RE = re.compile(r"(?<![A-Z])BANK(?![A-Z])")  # "BANK" as a word — not BANKE / BANKURA
_NON_TDS_RE = _word_re(_UNAMBIGUOUS_BANK_TOKENS + [
    "CO-OP BANK", "CO.OP BANK", "COOPERATIVE BANK", "CO-OPERATIVE BANK", "SAHAKARI BANK", "CREDIT SOCIETY",
    "MSWIPE", "RAZORPAY", "PAYTM", "PAYU", "PINELABS", "PINE LABS", "CARD SWIPE", "BILLDESK", "CASHFREE",
    "CCAVENUE", "PHONEPE", "GPAY", "GOOGLE PAY", "GOOGLEPAY",
    "CHEQUE SALES", "CASH SALES", "COUNTER SALES", "DIRECT SALES", "SALES A/C", "SALES ACCOUNT",
    "GOLD SALES", "SILVER SALES", "DIAMOND SALES", "PLATINUM SALES",
    "DUTIES", "DUTIES & TAXES", "GST", "IGST", "CGST", "SGST", "UTGST", "CESS",
    "INTEREST RECEIVED", "DISCOUNT RECEIVED", "DISCOUNT ALLOWED", "EXCHANGE GAIN", "EXCHANGE LOSS",
    "FOREX GAIN", "FOREX LOSS",
])

NOMINAL_EXPENSE_EXCLUSIONS = {
    "LABOUR CHARGES", "LABOUR CHARGE", "FREIGHT CHARGES", "FREIGHT CHARGE", "TRANSPORT CHARGES", "DESIGNING CHARGES",
    "INTEREST CHARGES", "SALARY PAYABLE", "CAPITAL A/C", "PARTNER DRAWINGS", "EXPENSE A/C", "CASH", "PETTY CASH",
    "ROUND OFF", "ROUNDING", "ROUNDED OFF",
}


def is_excluded_non_tds_party(name: Optional[str]) -> bool:
    """Returns True if the name represents a bank, cash, POS gateway, sales receipt, internal transfer, or tax duty account."""
    if not name:
        return True
    n_up = " ".join(name.strip().upper().split())
    if n_up in NOMINAL_EXPENSE_EXCLUSIONS:
        return True
    if _BANK_WORD_RE.search(n_up) or _NON_TDS_RE.search(n_up):
        return True
    if re.search(r"\b(CD|OD|CC|CA|SB)\s*[/-]?\s*\d{3,}", n_up) or re.search(r"\b(A/C|ACC|ACCOUNT)\s*(NO)?\.?\s*\d{4,}", n_up):
        return True
    return False


# =============================================================================
# Voucher-type helpers
# =============================================================================
def _vtype(tx_or_str: Any) -> str:
    v = tx_or_str if isinstance(tx_or_str, str) else (getattr(tx_or_str, "voucher_type", "") or "")
    return v.strip().upper()


def _is_outward_voucher(vt: str) -> bool:
    """Sales / receipts / credit notes: we are the payee, never the TDS deductor."""
    return ("SALE" in vt and "PURCHASE" not in vt) or vt.startswith("RECEIPT") or vt.startswith("CREDIT NOTE") \
        or vt.startswith("DELIVERY NOTE") or vt.startswith("SALES ORDER")


_DUTY_LEDGER_RE = _word_re(["GST", "IGST", "CGST", "SGST", "UTGST", "CESS", "INPUT", "OUTPUT", "RCM", "ROUND OFF",
                            "ROUNDOFF", "ROUNDING", "ROUNDED OFF", "TCS"])


def compute_tds_base(tx: NormalizedTransaction) -> Decimal:
    """
    TDS base for a bill/expense voucher: debit-side expense/purchase ledgers, EXCLUDING GST/RCM/round-off,
    TDS and the party ledger (CBDT Circular 23/2017: TDS on amount excluding GST where GST is shown separately).
    Falls back to total_amount when no such ledger exists.
    """
    party = _alnum(tx.party_name)
    base = Decimal("0.00")
    for l in tx.ledgers or []:
        if not l.is_debit or l.is_tds:
            continue
        ln = (l.ledger_name or "").upper()
        if _DUTY_LEDGER_RE.search(ln):
            continue
        if party and (_alnum(ln) == party):
            continue
        base += l.amount
    return base if base > 0 else tx.total_amount


# =============================================================================
# Party discovery
# =============================================================================
def _fy_of(date_val: Any) -> str:
    """'2026-27' style FY label from a normalized date (date/datetime or 'YYYY-MM-DD'/'YYYYMMDD' string)."""
    s = str(date_val or "")
    m = re.match(r"(\d{4})-?(\d{2})", s)
    if not m:
        return "?"
    y, mo = int(m.group(1)), int(m.group(2))
    start = y if mo >= 4 else y - 1
    return f"{start}-{str(start + 1)[-2:]}"


def discover_tds_parties(
    norm_m_old: List[NormalizedMaster],
    norm_m_new: List[NormalizedMaster],
    norm_t_old: List[NormalizedTransaction],
    norm_t_new: List[NormalizedTransaction],
) -> Set[str]:
    """
    Discovers TDS-relevant deductees from THIS dataset. Builds and activates a TDSContext learned from the
    company's own groups, ledger masters and past TDS deductions (see tds_context.py), so nothing is tied
    to a particular client. Statutory keywords are used only where the data has no signal.
    """
    ctx = build_tds_context(list(norm_m_new) + list(norm_m_old), list(norm_t_new) + list(norm_t_old))
    set_tds_context(ctx)
    discovered: Set[str] = set()

    # 1. Masters
    for m in (norm_m_old + norm_m_new):
        if not m.name:
            continue
        non_ded = ctx.is_non_deductee(m.name)
        if non_ded or (non_ded is None and is_excluded_non_tds_party(m.name)):
            continue
        info = ctx.info(m.name)
        cat = info.category if info else None
        if cat in ("EXPENSE", "PURCHASE", "FIXED_ASSET"):
            continue                                  # expense/purchase ledgers are natures, not deductees
        n = m.name.strip().lower()
        explicit = m.is_tds_applicable is True or bool(m.tds_section) or bool(getattr(m, "tds_rate_name", None)) \
            or bool(info and (info.nature_section or info.deductee_type))
        if explicit:
            discovered.add(n)
        elif cat == "DEBTOR":
            continue
        elif ctx.learned_party_section(m.name):
            discovered.add(n)
        elif cat is None and is_statutory_tds_party_name(m.name):
            discovered.add(n)

    # 2. Vouchers: any voucher whose strongest signal gives a section makes its party a deductee
    seen: Set[str] = set()
    for tx in (norm_t_new + norm_t_old):
        if not tx.party_name or tx.guid in seen:
            continue
        seen.add(tx.guid)
        if is_tds_viable_transaction(tx, None, ctx) and _deductee_name(tx, ctx):
            discovered.add(_deductee_name(tx, ctx).strip().lower())

    # 3. Section 194Q: purchases of goods > Rs 50 lakh per seller per FY, on value excluding GST.
    THRESHOLD_194Q = Decimal("5000000.00")
    cumulative: Dict[Tuple[str, str], Decimal] = defaultdict(Decimal)
    seen_g: Set[str] = set()
    seen_k: Set[Tuple[str, str, str, str]] = set()
    for tx in (norm_t_new + norm_t_old):
        if not tx.party_name:
            continue
        key = (tx.party_name.strip().lower(), tx.voucher_number or "", str(tx.date), str(tx.total_amount))
        if tx.guid in seen_g or key in seen_k:
            continue
        seen_g.add(tx.guid); seen_k.add(key)
        if _is_outward(tx, ctx) or _non_deductee(tx.party_name, ctx):
            continue
        goods = Decimal("0.00")
        for l in tx.ledgers:
            if not l.is_debit or l.is_tds:
                continue
            cat = ctx.category(l.ledger_name)
            if cat == "PURCHASE" or (cat is None and "PURCHASE" in l.ledger_name.upper()):
                goods += l.amount
        if goods > 0:
            cumulative[(key[0], _fy_of(tx.date))] += goods
    for (p_name, _fy), tot in cumulative.items():
        if tot > THRESHOLD_194Q:
            discovered.add(p_name)

    return discovered


def deduplicate_provisional_and_identical_purchases(tx_list: List[NormalizedTransaction]) -> List[NormalizedTransaction]:
    """
    Deduplicates provisional placeholder entries (such as GSTR-2A entries in Tally)
    when the actual supplier invoice for the same amount and party is already present.
    Also deduplicates duplicate re-entries of the exact same supplier bill reference.
    """
    by_party_amt: Dict[Tuple[str, Decimal], List[NormalizedTransaction]] = defaultdict(list)
    for tx in tx_list:
        p = (tx.party_name or "").upper().strip()
        amt = tx.total_amount
        vtype = _vtype(tx)
        if (vtype.startswith("PURCHASE") or vtype.startswith("JOURNAL")) and amt > Decimal("0.00"):
            by_party_amt[(p, amt)].append(tx)

    gstr_re = re.compile(r"GSTR\s*-?\s*2A|GSTRA2A")
    ref_res = _BILL_REF_RES
    suppressed_guids: Set[str] = set()
    for (p, amt), group in by_party_amt.items():
        if len(group) <= 1:
            continue
        gstr_entries = [t for t in group if gstr_re.search((t.narration or "").upper())]
        non_gstr_entries = [t for t in group if t not in gstr_entries]

        if gstr_entries and non_gstr_entries:
            for t in gstr_entries:
                suppressed_guids.add(t.guid)
        else:
            seen_refs: Dict[str, NormalizedTransaction] = {}
            for t in group:
                narr = (t.narration or "").upper()
                m = next((r.search(narr) for r in ref_res if r.search(narr)), None)
                ref = m.group(0) if m else _supplier_ref(t)
                if ref:
                    if ref in seen_refs:
                        suppressed_guids.add(t.guid)
                    else:
                        seen_refs[ref] = t

    return [t for t in tx_list if t.guid not in suppressed_guids]


def _supplier_ref(tx: NormalizedTransaction) -> Optional[str]:
    """Generic: the supplier invoice no. Tally stores on the voucher (REFERENCE / supplier invoice no.)."""
    raw = getattr(tx, "raw_reference", None) or {}
    if not isinstance(raw, dict):
        return None
    v = _get_first_key(raw, ["reference", "referenceno", "supplierinvoiceno", "supplier_invoice_no", "billno", "bill_no"])
    v = _alnum(str(v)) if v else ""
    return v if len(v) >= 3 else None


def is_pure_tds_adjustment_journal(tx: NormalizedTransaction) -> bool:
    """
    Checks if a Journal transaction is merely an internal accounting entry
    booking TDS liability (e.g. Dr Party, Cr TDS Payable / TDS ON PURCHASE 194Q) rather than an actual bill or expense.
    """
    if not _vtype(tx).startswith("JOURNAL"):
        return False
    ledgers = getattr(tx, "ledgers", []) or []
    if len(ledgers) < 2:
        return False
    party = _alnum(getattr(tx, "party_name", ""))
    if not party:
        return False
    has_tds = has_party = False
    for l in ledgers:
        lname = (getattr(l, "ledger_name", "") or "").upper()
        if l.is_tds:
            has_tds = True
        elif _alnum(lname) == party:
            has_party = True
        elif re.search(r"ROUND\s*(ED\s*)?OFF|ROUNDING", lname):
            continue
        else:
            return False  # Contains actual expense or purchase ledger
    return has_tds and has_party


# =============================================================================
# Viability
# =============================================================================
_PARTNER_EXPENSE_RE = _word_re(["INTEREST ON CAPITAL", "INTEREST TO PARTNER", "PARTNER INTEREST", "REMUNERATION",
                                "SALARY TO PARTNER", "PARTNER SALARY", "COMMISSION TO PARTNER", "BONUS TO PARTNER", "194T"])
_CAPITAL_RE = re.compile(r"\bCAP\s*-|\bCAPITAL\b|\bDRAWINGS?\b")


_CAP_LEDGER_RE = re.compile(r"^\s*CAP(ITAL)?\s*[-:.]|\bCAPITAL\s*(A/C|ACCOUNT|ACC)\b|\bPARTNER'?S?\s*(CAPITAL|CURRENT)\b|\bCURRENT\s*(A/C|ACCOUNT)\s*[-:]", re.I)


def _is_partner_capital_ledger(name: str) -> bool:
    """Generic partner capital/current account: 'CAP- X', 'Capital A/c - X', 'X Capital Account', 'Partner Current A/c X'.
    Bank 'Current Account' ledgers are excluded."""
    n = (name or "").upper()
    if is_excluded_non_tds_party(n) and not _CAP_LEDGER_RE.search(n):
        return False
    if _BANK_WORD_RE.search(n):
        return False
    return bool(_CAP_LEDGER_RE.search(n)) or (bool(PARTNER_NAME_HINTS) and any(h in n for h in PARTNER_NAME_HINTS) and "CAPITAL" in n)


def _partner_name_from_capital(name: str) -> str:
    n = re.sub(r"^\s*CAP(ITAL)?\s*[-:.]\s*", "", name, flags=re.I)
    n = re.sub(r"\b(PARTNER'?S?\s*)?(CAPITAL|CURRENT)\s*(A/C|ACCOUNT|ACC)?\b\s*[-:]?", "", n, flags=re.I)
    return " ".join(n.split()).strip(" -:") or name


def _section_regexes() -> List["re.Pattern[str]"]:
    pats = []
    for rule in STATUTORY_TDS_REGISTRY:
        for code in {rule.section_code, rule.sub_section or rule.section_code}:
            c = re.escape(code.upper())
            if re.fullmatch(r"\d+", code):
                # bare numbers (192/193/194/195) only count with a 'u/s', 'sec' or 'TDS' prefix — not "BILL NO 195"
                pats.append(re.compile(rf"(?:\bU/?S\.?|\bSEC(?:TION)?\.?|\bTDS)\s*[-:]?\s*{c}(?![0-9A-Z])"))
            else:
                pats.append(re.compile(rf"(?<![0-9A-Z]){c}(?![0-9A-Z])"))
    return pats


_SECTION_RES = _section_regexes()


def _is_outward(tx: NormalizedTransaction, ctx: Optional[TDSContext]) -> bool:
    if ctx is not None:
        o = ctx.is_outward(tx)
        if o is not None:
            return o
    return _is_outward_voucher(_vtype(tx))


def _non_deductee(name: Optional[str], ctx: Optional[TDSContext]) -> bool:
    if not name:
        return True
    if ctx is not None:
        r = ctx.is_non_deductee(name)
        if r is not None:
            return r
    return is_excluded_non_tds_party(name)


def _deductee_name(tx: NormalizedTransaction, ctx: Optional[TDSContext]) -> Optional[str]:
    """The payee: the partner for 194T (capital ledger), else the voucher party."""
    if ctx is not None:
        cap = ctx.partner_capital_ledger(tx)
        if cap is not None:
            return _partner_name_from_capital(cap.ledger_name)
    return tx.party_name


def _is_tds_line(l: NormalizedLedger, ctx: Optional[TDSContext]) -> bool:
    return ctx.is_tds_line(l) if ctx is not None else bool(l.is_tds)


def _is_cost(name: str, party: Optional[str], ctx: Optional[TDSContext]) -> bool:
    if ctx is not None:
        cat = ctx.category(name)
        if cat is not None:
            return cat in ("EXPENSE", "PURCHASE", "FIXED_ASSET") and _alnum(name) != _alnum(party)
    n = (name or "").upper()
    return _alnum(name) != _alnum(party) and not _BANK_WORD_RE.search(n) and not _NON_TDS_RE.search(n) \
        and n.strip() not in ("CASH", "PETTY CASH") \
        and not re.search(r"(?<![A-Z])(PAYABLE|OUTSTANDING|PROVISION|ADVANCE|DEPOSIT|CAPITAL|DRAWINGS?)(?![A-Z])", n)


def suggest_tds_section(tx: NormalizedTransaction, ctx: Optional[TDSContext] = None) -> Tuple[Optional[str], str]:
    """(section, reason) from the dataset-learned context; keyword fallback when the data has no signal."""
    ctx = ctx or _ACTIVE_CTX
    if ctx is not None:
        sec, why = ctx.suggest_section(tx)
        if sec:
            return sec, why
    for l in tx.ledgers:
        if l.is_debit and _is_cost(l.ledger_name, tx.party_name, ctx):
            from app.services.tds_statutory_rates import suggest_statutory_rule
            r = suggest_statutory_rule(l.ledger_name)
            if r:
                return (r.sub_section or r.section_code), f"keyword on ledger '{l.ledger_name}'"
    if tx.party_name:
        from app.services.tds_statutory_rates import suggest_statutory_rule
        r = suggest_statutory_rule(tx.party_name)
        if r:
            return (r.sub_section or r.section_code), f"keyword in party name '{tx.party_name}'"
    return None, ""


def is_tds_viable_transaction(
    tx: NormalizedTransaction,
    tds_party_names: Optional[Set[str]] = None,
    ctx: Optional[TDSContext] = None,
) -> bool:
    """
    Is this voucher a TDS event (deduction made, or deduction due)?
    Uses the dataset-learned context when available (groups, masters, past deductions), else rule fallbacks.
    """
    ctx = ctx or _ACTIVE_CTX
    vt = _vtype(tx)

    if vt.startswith("CONTRA"):
        return False
    # TDS actually booked (our TDS ledger credited) makes us the payer, so it is a deduction whatever the party
    # ledger looks like — also a journal that only charges the TDS to the party (Dr Party / Cr TDS) for a bill
    # booked earlier, or a partners' journal (Dr Capital / Cr TDS).
    if any(l.is_credit and _is_tds_line(l, ctx) for l in tx.ledgers):
        return True
    if _is_outward(tx, ctx):
        return False
    if _non_deductee(tx.party_name, ctx) and not (ctx is not None and ctx.partner_capital_ledger(tx)):
        return False
    if is_pure_tds_adjustment_journal(tx):
        return False

    has_194q_ledger = any("194Q" in (l.ledger_name or "").upper() for l in tx.ledgers)
    if tx.tds_amount != Decimal("0.00") or has_194q_ledger:
        return True

    cost_debits = [l for l in tx.ledgers if l.is_debit and not _is_tds_line(l, ctx) and _is_cost(l.ledger_name, tx.party_name, ctx)]

    # A payment that only settles a bill (Dr party / Cr bank) is not a new TDS event; a direct expense payment is.
    if vt.startswith("PAYMENT") and not cost_debits:
        return False

    # Partner: Cr partner Capital + Dr interest/remuneration expense = 194T. Capital/drawings moves alone are not.
    if ctx is not None and ctx.partner_capital_ledger(tx) is not None:
        return bool(cost_debits)
    all_l_up = " ".join(l.ledger_name for l in tx.ledgers).upper()
    if vt.startswith("JOURNAL") and ("SALARY PAYABLE" in all_l_up or _CAPITAL_RE.search(all_l_up)) \
            and not _PARTNER_EXPENSE_RE.search(all_l_up):
        return False

    if tx.tds_section or tx.tds_rate:
        return True

    # Dataset-learned signals: master Nature of Payment / earlier deductions on this ledger or party
    if ctx is not None and ctx.suggest_section(tx)[0]:
        return True

    if tds_party_names and tx.party_name and tx.party_name.strip().lower() in tds_party_names:
        return True

    # Fallback only where the data gave no answer: statutory keywords on the cost ledgers and party name
    for l in cost_debits:
        if is_statutory_keyword_match(l.ledger_name) and not (ctx is not None and ctx.info(l.ledger_name) and ctx.info(l.ledger_name).tds_applicable is False):
            return True
    if tx.party_name and is_statutory_tds_party_name(tx.party_name):
        pi = ctx.info(tx.party_name) if ctx is not None else None
        if not (pi and pi.tds_applicable is False):
            return True

    all_text = " ".join([tx.party_name or "", (tx.narration or "").upper()] + [l.ledger_name for l in tx.ledgers]).upper()
    return any(p.search(all_text) for p in _SECTION_RES)


# =============================================================================
# Normalization
# =============================================================================
def normalize_transaction(raw: Dict[str, Any], index: int) -> NormalizedTransaction:
    """
    Converts raw transaction JSON dictionary into canonical NormalizedTransaction.
    Raises ValidationError with exact index and reason if required data is corrupted.
    """
    # 1. GUID
    guid = _get_first_key(raw, ["guid", "voucherkey", "vchkey", "id", "remoteid", "voucher_id"])
    if not guid:
        vch_num = _get_first_key(raw, ["vouchernumber", "voucher_no", "vchno", "number", "vch_no"])
        vch_date = _get_first_key(raw, ["date", "voucher_date", "vchdate"])
        if vch_num and vch_date:
            guid = f"VCH-{vch_num}-{vch_date}"
        else:
            raise ValidationError(f"Transaction at index {index} has no GUID or Voucher Identifier")
    guid = str(guid).strip()

    # 2. Voucher Number (Allow unnumbered/blank vouchers gracefully)
    vch_number = _get_first_key(raw, [
        "vouchernumber", "voucher_no", "vchno", "number", "vch_no", "voucher_num",
        "reference", "referenceno", "ref_no", "refno"
    ])
    vch_number = str(vch_number).strip() if vch_number is not None else ""

    # 3. Voucher Type
    vch_type = _get_first_key(raw, ["vouchertypename", "voucher_type", "type", "vchtype", "voucher_typ_name"])
    vch_type = normalize_string(vch_type or "Journal")

    # 4. Date
    raw_date = _get_first_key(raw, ["date", "voucher_date", "vchdate", "vch_date"])
    if not raw_date:
        raise ValidationError(f"Transaction (GUID: {guid}, index: {index}) has no voucher date")
    try:
        norm_date = normalize_date(raw_date)
    except ValueError as e:
        raise ValidationError(f"Transaction (GUID: {guid}, index: {index}) has invalid date: {str(e)}")

    # 5. Party Name
    party_name = _get_first_key(raw, ["partyledgername", "party_name", "party", "partyname", "party_ledger"])
    party_name = normalize_string(party_name or "")

    # 6. Narration
    narration = _get_first_key(raw, ["narration", "remarks", "description"])
    narration = normalize_string(narration) if narration else None

    # 7. Ledgers
    raw_ledgers = _get_first_key(
        raw,
        [
            "allledgerentries.list", "allledgerentries", "ledgerentries.list", "ledgerentries",
            "ledgers", "ledger_entries", "all_ledger_entries", "allledgerentry", "ledgerentry",
            "accountingentries", "accounting_entries"
        ]
    )
    if isinstance(raw_ledgers, dict):
        raw_ledgers = [raw_ledgers]
    elif not isinstance(raw_ledgers, list):
        raw_ledgers = []
    raw_ledgers = list(raw_ledgers)

    raw_inv = _get_first_key(
        raw,
        ["allinventoryentries.list", "allinventoryentries", "inventoryentries.list", "inventoryentries", "inventory_entries"]
    )
    if isinstance(raw_inv, dict):
        raw_inv = [raw_inv]
    if isinstance(raw_inv, list):
        for inv_item in raw_inv:
            if isinstance(inv_item, dict):
                allocs = _get_first_key(inv_item, ["accountingallocations.list", "accountingallocations", "accounting_allocations", "ledgerentries", "ledgers"])
                if isinstance(allocs, dict):
                    allocs = [allocs]
                if isinstance(allocs, list):
                    raw_ledgers.extend(a for a in allocs if isinstance(a, dict))

    if not raw_ledgers and isinstance(raw_inv, list):
        for inv_item in raw_inv:
            if isinstance(inv_item, dict):
                item_name = _get_first_key(inv_item, ["stockitemname", "itemname", "name"]) or "Inventory Item"
                item_amt = _get_first_key(inv_item, ["amount", "amt", "item_amount"])
                if item_amt is not None:
                    raw_ledgers.append({
                        "ledgername": item_name,
                        "amount": item_amt,
                        "isdeemedpositive": inv_item.get("isdeemedpositive", True),
                    })

    if not party_name and raw_ledgers:
        for l_raw in raw_ledgers:
            if isinstance(l_raw, dict) and str(l_raw.get("ispartyledger", "")).lower() in ["true", "yes", "1"]:
                p_name_cand = _get_first_key(l_raw, ["ledgername", "ledger_name", "name"])
                if p_name_cand:
                    party_name = normalize_string(p_name_cand)
                    break

    parsed_ledgers: List[Tuple[Dict[str, Any], str, Decimal]] = []
    for l_raw in raw_ledgers:
        if not isinstance(l_raw, dict):
            continue
        l_name = _get_first_key(l_raw, ["ledgername", "ledger_name", "name", "ledger"])
        if not l_name:
            continue
        l_name = normalize_string(l_name)

        raw_amt = _get_first_key(l_raw, ["amount", "amt", "ledger_amount", "ledgeramount", "billamount"])
        try:
            l_amount = parse_decimal(raw_amt, Decimal("0.00"))
        except ValueError as e:
            raise ValidationError(f"Transaction (GUID: {guid}, index: {index}) ledger '{l_name}' has invalid amount: {str(e)}")
        parsed_ledgers.append((l_raw, l_name, l_amount))

    # When amounts are signed, the sign decides the side. ISDEEMEDPOSITIVE is not reliable line by line: in Tally
    # invoice-mode vouchers the TDS / round-off lines carry ISDEEMEDPOSITIVE=Yes with a positive (credit) amount.
    # The sign convention itself (Tally: NEGATIVE = debit; some exports: positive = debit) is read from the lines
    # of this voucher on which ISDEEMEDPOSITIVE and the sign agree.
    def _deemed(l_raw: Dict[str, Any]) -> Optional[bool]:
        v = _get_first_key(l_raw, ["isdeemedpositive", "is_deemed_positive", "isdebit", "is_debit"])
        return None if v is None else str(v).strip().lower() in ["yes", "true", "1"]

    amounts_signed = any(a < 0 for _, _, a in parsed_ledgers)
    votes = [(a < 0) == _deemed(r) for r, _, a in parsed_ledgers if a != 0 and _deemed(r) is not None]
    negative_is_debit = not votes or sum(votes) * 2 >= len(votes)

    ledgers: List[NormalizedLedger] = []
    for l_raw, l_name, l_amount in parsed_ledgers:
        deemed = _deemed(l_raw)
        if amounts_signed and l_amount != 0:
            is_debit = (l_amount < 0) if negative_is_debit else (l_amount > 0)
        elif deemed is not None:
            is_debit = deemed
        else:
            is_debit = l_amount < 0
        is_credit = not is_debit
        l_amount = abs(l_amount)

        is_tds = is_tds_ledger(l_name, l_raw)
        tds_sec = _clean_section(_get_first_key(l_raw, ["tdssection", "tds_section", "section"]))
        tds_rate = _safe_decimal(_get_first_key(l_raw, ["tdsrate", "tds_rate"]))

        ledgers.append(
            NormalizedLedger(
                ledger_name=l_name,
                amount=l_amount,
                is_debit=is_debit,
                is_credit=is_credit,
                ledger_id=_get_first_key(l_raw, ["ledgerid", "ledger_id", "id"]),
                is_tds=is_tds,
                tds_rate=tds_rate,
                tds_section=tds_sec,
            )
        )

    if not party_name and ledgers:
        non_tds = [l for l in ledgers if not l.is_tds]
        party_name = (non_tds or ledgers)[0].ledger_name

    # TDS deducted:
    # In invoice / expense / journal / purchase vouchers, any TDS withholding ledger represents
    # tax deducted at source. In Tally exports, tax entries within invoices often carry
    # isdeemedpositive: True. Thus for non-payment/non-debit note vouchers, TDS deduction is the sum of TDS ledger amounts.
    vt_clean = _vtype(vch_type)
    is_optional = _is_yes(_get_first_key(raw, ["isoptional", "is_optional"]))
    is_cancelled = _is_yes(_get_first_key(raw, ["iscancelled", "is_cancelled", "isvoid"]))
    if vt_clean.startswith("PAYMENT") or vt_clean.startswith("DEBIT NOTE") or is_optional or is_cancelled:
        # optional (memo) / cancelled vouchers are not in the books: no TDS was deducted by them
        tds_amount = Decimal("0.00")
    else:
        tds_amount = sum((l.amount for l in ledgers if l.is_tds), Decimal("0.00"))

    vch_amt_val = _get_first_key(raw, ["amount", "total_amount", "voucher_amount", "vchamount", "vch_amount", "netamount", "grossamount", "totalamount"])
    total_amount = Decimal("0.00")
    if vch_amt_val is not None:
        try:
            total_amount = abs(parse_decimal(vch_amt_val))
        except ValueError:
            total_amount = Decimal("0.00")

    if total_amount == Decimal("0.00") and ledgers:
        p_key = _alnum(party_name)
        party_ledgers = [l for l in ledgers if _alnum(l.ledger_name) == p_key]
        if party_ledgers:
            party_amt = sum(l.amount for l in party_ledgers)
            total_amount = party_amt + tds_amount if tds_amount > 0 else party_amt
        else:
            # one side only — summing debits AND credits doubles the voucher value
            debit_side = sum((l.amount for l in ledgers if l.is_debit and not l.is_tds), Decimal("0.00"))
            total_amount = debit_side if debit_side > 0 else sum((l.amount for l in ledgers if l.is_credit), Decimal("0.00"))

    tx_tds_sec = _clean_section(_get_first_key(raw, ["tdssection", "tds_section", "section"]))
    tx_tds_rate = _safe_decimal(_get_first_key(raw, ["tdsrate", "tds_rate"]))

    # 5b. Partner (194T) and configured party aliases
    all_l_names = [l.ledger_name for l in ledgers]
    joined = " ".join(all_l_names).upper()
    partner_cap = [l for l in ledgers if _is_partner_capital_ledger(l.ledger_name)]
    if "194T" in joined or (partner_cap and _PARTNER_EXPENSE_RE.search(joined)):
        if partner_cap:
            # keep the INDIVIDUAL partner (194T threshold/PAN is per partner) — strip the "CAP-" prefix
            party_name = _partner_name_from_capital(partner_cap[0].ledger_name)
        tx_tds_sec = tx_tds_sec or "194T"
        # base = the interest/remuneration expense actually debited (not TDS x 10)
        exp = sum((l.amount for l in ledgers if l.is_debit and _PARTNER_EXPENSE_RE.search(l.ledger_name.upper())), Decimal("0.00"))
        if exp > 0:
            total_amount = exp
    else:
        text = f"{party_name} | {joined}".upper()
        for rx, canonical, default_sec in _ALIAS_RES:
            if rx.search(text):
                party_name = canonical
                tx_tds_sec = tx_tds_sec or default_sec
                break

    if not tx_tds_sec:
        for l in ledgers:
            if l.is_tds or "TDS" in l.ledger_name.upper():
                ln_up = l.ledger_name.upper()
                if "194T" in ln_up or "PARTNER" in ln_up:
                    tx_tds_sec = "194T"
                    break
                elif "INTEREST" in ln_up:
                    tx_tds_sec = "194A"
                    break
                elif any(k in ln_up for k in ["PROFESSIONAL", "FEES", "TECHNICAL", "194J"]):
                    tx_tds_sec = "194J"
                    break
                elif any(k in ln_up for k in ["CONTRACT", "LABOUR", "WORK", "194C"]):
                    tx_tds_sec = "194C"
                    break
                elif any(k in ln_up for k in ["PURCHASE", "GOODS", "194Q"]):
                    tx_tds_sec = "194Q"
                    break
                elif any(k in ln_up for k in ["RENT", "194I"]):
                    tx_tds_sec = "194I"
                    break
                elif any(k in ln_up for k in ["COMMISSION", "BROKERAGE", "194H"]):
                    tx_tds_sec = "194H"
                    break

    if party_name:
        party_name = re.sub(r"\s*\((?:LABOUR|LABOUR\s+CHARGES?|CREDITOR)\)", "", party_name, flags=re.I).strip()

    sorted_ledgers = sorted(ledgers, key=lambda l: l.sort_key())

    return NormalizedTransaction(
        guid=guid,
        voucher_number=vch_number,
        voucher_type=vch_type,
        date=norm_date,
        party_name=party_name,
        party_id=_get_first_key(raw, ["partyid", "party_id"]),
        narration=narration,
        ledgers=sorted_ledgers,
        total_amount=total_amount,
        tds_amount=tds_amount,
        tds_rate=tx_tds_rate,
        tds_section=tx_tds_sec,
        raw_reference=raw,
        is_optional=is_optional,
        is_cancelled=is_cancelled,
    )


def _is_yes(v: Any) -> bool:
    return str(v).strip().lower() in ("yes", "true", "1") if v is not None else False


def _extract_master_name(raw: Dict[str, Any]) -> Optional[str]:
    """Helper to extract ledger/master name from standard keys or Tally languagename structure."""
    val = _get_first_key(raw, ["name", "ledgername", "partyname", "ledger_name", "party_name"])
    if val:
        if isinstance(val, list):
            strs = [x for x in val if isinstance(x, str)]
            if strs:
                return strs[0]
        elif isinstance(val, str):
            return val

    lang = raw.get("languagename") or raw.get("LANGUAGENAME")
    if lang:
        if isinstance(lang, dict):
            lang = [lang]
        if isinstance(lang, list):
            for l_item in lang:
                if isinstance(l_item, dict) and "name" in l_item:
                    n = l_item["name"]
                    if isinstance(n, list):
                        strs = [x for x in n if isinstance(x, str)]
                        if strs:
                            return strs[0]
                    elif isinstance(n, str):
                        return n

    val = _get_first_key(raw, ["mailingname", "originalname", "guid"])
    if val:
        return str(val)
    return None


def normalize_master(raw: Dict[str, Any], index: int) -> NormalizedMaster:
    """
    Converts raw master JSON dictionary into canonical NormalizedMaster.
    """
    name = _extract_master_name(raw)
    if not name:
        raise ValidationError(f"Master record at index {index} has no Name")
    name = normalize_string(name)

    guid = _get_first_key(raw, ["guid", "masterid", "id", "ledger_id", "remoteid"])
    if not guid:
        guid = f"MST-{name.upper()}"
    guid = str(guid).strip()

    parent = _get_first_key(raw, ["parent", "group", "parentgroup", "parent_group"])
    pan = _get_first_key(raw, ["incometaxnumber", "pan", "pannumber", "pan_no", "it_number"])
    gstin = _get_first_key(raw, ["partygstin", "gstin", "gst_number", "gstin_uin"])
    is_tds = _get_first_key(raw, ["istdsapplicable", "is_tds_applicable", "tdsapplicable"])
    tds_sec = _get_first_key(raw, ["tdssection", "tds_section", "section"])
    tds_rate_name = _get_first_key(raw, ["tdsratename", "tds_rate_name", "rate_name"])

    if not gstin:
        # Tally keeps a ledger's GST registration history in LEDGSTREGDETAILS; the latest one holds the GSTIN
        regs = _get_first_key(raw, ["ledgstregdetails", "ledgstregdetails.list"])
        regs = [regs] if isinstance(regs, dict) else (regs if isinstance(regs, list) else [])
        regs = [r for r in regs if isinstance(r, dict) and _get_first_key(r, ["gstin"])]
        if regs:
            latest = max(regs, key=lambda r: str(_get_first_key(r, ["applicablefrom"]) or ""))
            gstin = _get_first_key(latest, ["gstin"])
    pan_clean = re.sub(r"\s+", "", str(pan)).upper() if pan else None
    gstin_clean = re.sub(r"\s+", "", str(gstin)).upper() if gstin else None
    # Tally often has PAN blank but GSTIN filled — PAN is characters 3-12 of the GSTIN
    if (not pan_clean or not re.fullmatch(r"[A-Z]{5}\d{4}[A-Z]", pan_clean)) and gstin_clean \
            and re.fullmatch(r"\d{2}[A-Z]{5}\d{4}[A-Z][0-9A-Z]{3}", gstin_clean):
        pan_clean = gstin_clean[2:12]

    clean_raw = {k: v for k, v in raw.items() if not is_technical_field(k)}

    return NormalizedMaster(
        guid=guid,
        name=name,
        parent_group=normalize_string(parent) if parent else None,
        pan=pan_clean,
        gstin=gstin_clean,
        is_tds_applicable=str(is_tds).strip().lower() in ["yes", "true", "1"] if is_tds is not None else None,
        tds_section=str(tds_sec).strip() if tds_sec else None,
        tds_rate_name=str(tds_rate_name).strip() if tds_rate_name else None,
        raw_fields=clean_raw,
    )
