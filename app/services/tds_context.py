"""
Dataset-driven TDS context — learns everything it needs from the company's own Tally data, so the
normalizer does not depend on any client, party list or fixed bank/keyword list.

Signals, strongest first:
  1. Tally's reserved (primary) groups, which are identical in every Tally company
     (Bank Accounts, Sundry Debtors, Capital Account, Duties & Taxes, Sales Accounts, ...).
     Each ledger is resolved up its group chain to one of these.
  2. Ledger masters: "Is TDS Applicable", "Nature of Payment" / TDS category, tax type TDS on duty ledgers,
     deductee type, PAN.
  3. History in the vouchers themselves: where TDS WAS deducted, the expense ledger and the party
     "vote" for that section; later vouchers on the same ledger/party without TDS are then flagged.
  4. Only when 1–3 say nothing: statutory keywords (tds_statutory_rates) as a fallback.

Build once per dataset (discover_tds_parties() in the normalizer does this automatically):
    ctx = build_tds_context(masters, transactions)
"""
from collections import Counter, defaultdict
from dataclasses import dataclass, field
from decimal import Decimal
from typing import Any, Dict, Iterable, List, Optional, Set, Tuple
import re

# --------------------------------------------------------------------------- Tally reserved groups
# key = normalized group name -> category used by the rules
RESERVED_GROUPS: Dict[str, str] = {
    "BANK ACCOUNTS": "BANK", "BANK OD A/C": "BANK", "BANK OCC A/C": "BANK", "BANK OD ACCOUNT": "BANK",
    "CASH-IN-HAND": "CASH", "CASH IN HAND": "CASH",
    "DUTIES & TAXES": "DUTY", "DUTIES AND TAXES": "DUTY",
    "CAPITAL ACCOUNT": "CAPITAL", "RESERVES & SURPLUS": "CAPITAL",
    "SALES ACCOUNTS": "SALES", "DIRECT INCOMES": "INCOME", "INDIRECT INCOMES": "INCOME",
    "INCOME (DIRECT)": "INCOME", "INCOME (INDIRECT)": "INCOME",
    "PURCHASE ACCOUNTS": "PURCHASE",
    "DIRECT EXPENSES": "EXPENSE", "INDIRECT EXPENSES": "EXPENSE",
    "EXPENSES (DIRECT)": "EXPENSE", "EXPENSES (INDIRECT)": "EXPENSE",
    "SUNDRY DEBTORS": "DEBTOR", "SUNDRY CREDITORS": "CREDITOR",
    "LOANS (LIABILITY)": "LOAN", "SECURED LOANS": "LOAN", "UNSECURED LOANS": "LOAN",
    "LOANS & ADVANCES (ASSET)": "ADVANCE", "DEPOSITS (ASSET)": "ADVANCE",
    "CURRENT ASSETS": "CUR_ASSET", "CURRENT LIABILITIES": "CUR_LIAB", "PROVISIONS": "CUR_LIAB",
    "FIXED ASSETS": "FIXED_ASSET", "INVESTMENTS": "INVESTMENT", "STOCK-IN-HAND": "STOCK",
    "SUSPENSE A/C": "SUSPENSE", "BRANCH / DIVISIONS": "BRANCH", "MISC. EXPENSES (ASSET)": "MISC_ASSET",
}
# groups whose ledgers can never be a TDS deductee
NON_DEDUCTEE_CATEGORIES = {"BANK", "CASH", "DUTY", "SALES", "INCOME", "STOCK", "SUSPENSE", "BRANCH", "ROUND_OFF"}
# Invoice rounding ledgers: usually kept under Indirect Expenses, but never a cost a deductee is paid for.
# Tally marks them on the master (VATDEALERNATURE / rounding method "Invoice Rounding"); the name is the fallback.
_ROUND_OFF_RE = re.compile(r"(?<![A-Z])ROUND(ED|ING)?\s*-?\s*OFF(?![A-Z])|(?<![A-Z])ROUNDING(?![A-Z])")
# groups whose DEBIT means "an expense / cost we pay for"
COST_CATEGORIES = {"EXPENSE", "PURCHASE", "FIXED_ASSET"}

# Tally "Nature of Payment" (TDS category) names -> section. These are Tally's built-in list, same everywhere.
NATURE_TO_SECTION: List[Tuple[str, str]] = [
    # spellings seen in ledger names are tolerated: PROFESSION/PROFESSIONAL, COMISSION/COMMISION/COMMISSION
    (r"CONTRACT|TRANSPORT|ADVERTIS", "194C"),
    (r"PROFESSION|ROYALTY|DIRECTOR|NON.?COMPETE", "194J(b)"),
    (r"TECHNICAL|CALL CENT", "194J(a)"),
    (r"COMM?ISS?ION.*(INSURANCE)|INSURANCE COMM?ISS?ION", "194D"),
    (r"LOTTERY.*COMM?ISS?ION|COMM?ISS?ION.*LOTTERY", "194G"),
    (r"COMM?ISS?ION|BROKERAGE", "194H"),
    (r"RENT.*(PLANT|MACHIN|EQUIP)|(PLANT|MACHIN|EQUIP).*RENT", "194I(a)"),
    (r"RENT", "194I(b)"),
    (r"PURCHASE OF GOODS|194Q", "194Q"),
    (r"PARTNER|194T", "194T"),
    (r"INTEREST ON SECURIT", "193"),
    (r"INTEREST", "194A"),
    (r"DIVIDEND", "194"),
    (r"IMMOVABLE PROPERTY", "194IA"),
    (r"PERQUISITE|BENEFIT", "194R"),
    (r"VIRTUAL DIGITAL", "194S"),
    (r"E.?COMMERCE", "194O"),
    (r"NON.?RESIDENT|FOREIGN", "195"),
]

_SEC_IN_TEXT = re.compile(r"(?<![0-9A-Z])(19[2-6](?:IA|IB|IC|LA|DA|BA|BB|EE|[A-Z])?(?:\([AB]\))?)(?![0-9A-Z])")
_RATE_IN_TEXT = re.compile(r"@\s*(\d+(?:\.\d+)?)\s*%")


def _norm(s: Any) -> str:
    return " ".join(str(s or "").upper().split())


def _yes(v: Any) -> bool:
    return str(v).strip().lower() in ("yes", "true", "1") if v is not None else False


def _first(d: Dict[str, Any], keys: Iterable[str]) -> Optional[Any]:
    if not isinstance(d, dict):
        return None
    low = {str(k).lower(): v for k, v in d.items()}
    for k in keys:
        v = low.get(k.lower())
        if v not in (None, "", [], {}):
            return v
    return None


def section_from_text(text: str) -> Optional[str]:
    """'TDS ON CONTRACT @2%' -> 194C; 'TDS 194J' -> 194J; 'Payment to Contractors' -> 194C."""
    t = _norm(text)
    if not t:
        return None
    m = _SEC_IN_TEXT.search(t)
    if m:
        return m.group(1)
    for pat, sec in NATURE_TO_SECTION:
        if re.search(pat, t):
            return sec
    return None


# --------------------------------------------------------------------------- context
@dataclass
class LedgerInfo:
    name: str
    parent: Optional[str] = None
    category: Optional[str] = None          # resolved reserved-group category
    is_tds_duty: bool = False               # a TDS duty ledger (tax type TDS / name says TDS, under Duties & Taxes)
    tds_applicable: Optional[bool] = None
    nature_section: Optional[str] = None    # from master "Nature of Payment" / TDS category / section
    pan: Optional[str] = None
    deductee_type: Optional[str] = None


@dataclass
class TDSContext:
    ledgers: Dict[str, LedgerInfo] = field(default_factory=dict)          # key: normalized name
    ledger_votes: Dict[str, Counter] = field(default_factory=lambda: defaultdict(Counter))   # expense ledger -> sections seen with TDS
    party_votes: Dict[str, Counter] = field(default_factory=lambda: defaultdict(Counter))    # party -> sections seen with TDS
    min_votes: int = 1
    min_share: float = 0.6
    has_group_data: bool = False

    # ---- lookups
    def info(self, ledger_name: Optional[str]) -> Optional[LedgerInfo]:
        return self.ledgers.get(_norm(ledger_name))

    def category(self, ledger_name: Optional[str]) -> Optional[str]:
        i = self.info(ledger_name)
        if i and i.category:
            return i.category
        if _ROUND_OFF_RE.search(_norm(ledger_name)):
            return "ROUND_OFF"
        return _category_from_group_name(ledger_name)   # ledger named like a reserved group, e.g. "Cash"

    def _winner(self, votes: Counter) -> Optional[str]:
        if not votes:
            return None
        sec, n = votes.most_common(1)[0]
        total = sum(votes.values())
        return sec if n >= self.min_votes and n / total >= self.min_share else None

    def learned_ledger_section(self, ledger_name: str) -> Optional[str]:
        return self._winner(self.ledger_votes.get(_norm(ledger_name), Counter()))

    def learned_party_section(self, party: str) -> Optional[str]:
        return self._winner(self.party_votes.get(_norm(party), Counter()))

    def summary(self) -> Dict[str, Any]:
        cats = Counter(i.category for i in self.ledgers.values())
        return {
            "ledgers_known": len(self.ledgers),
            "group_data_available": self.has_group_data,
            "ledgers_by_category": dict(cats),
            "tds_duty_ledgers": sorted(i.name for i in self.ledgers.values() if i.is_tds_duty),
            "expense_ledgers_with_nature": {i.name: i.nature_section for i in self.ledgers.values() if i.nature_section and i.category in COST_CATEGORIES},
            "learned_expense_ledgers": {k: self.learned_ledger_section(k) for k in self.ledger_votes if self.learned_ledger_section(k)},
            "learned_parties": {k: self.learned_party_section(k) for k in self.party_votes if self.learned_party_section(k)},
        }


def _category_from_group_name(name: Optional[str]) -> Optional[str]:
    n = _norm(name)
    if not n:
        return None
    if n in RESERVED_GROUPS:
        return RESERVED_GROUPS[n]
    # user sub-groups usually keep the reserved name inside ("Sundry Creditors- Comission", "Bank OD A/c - HDFC")
    for g, cat in sorted(RESERVED_GROUPS.items(), key=lambda kv: -len(kv[0])):
        if g in n:
            return cat
    return None


# --------------------------------------------------------------------------- builder
def build_tds_context(masters: Iterable[Any], transactions: Iterable[Any], min_votes: int = 1, min_share: float = 0.6) -> TDSContext:
    """
    masters: NormalizedMaster objects (ledgers AND, if exported, groups) from all snapshots.
    transactions: NormalizedTransaction objects from all snapshots.
    """
    ctx = TDSContext(min_votes=min_votes, min_share=min_share)
    parent_of: Dict[str, str] = {}
    # Typed Tally group masters. Kept apart from ledgers: a group and a ledger may share a name
    # ("PROFESSIONAL FEES" group holding a "Professional Fees" ledger).
    group_parent: Dict[str, str] = {}
    group_category: Dict[str, str] = {}
    rounding_ledgers: Set[str] = set()

    for m in masters:
        name = _norm(getattr(m, "name", None))
        if not name:
            continue
        raw = getattr(m, "raw_fields", None) or {}
        parent = _norm(getattr(m, "parent_group", None) or _first(raw, ["parent", "group"]))
        meta = raw.get("metadata") if isinstance(raw.get("metadata"), dict) else {}
        if _norm(meta.get("type")) == "GROUP":
            if parent:
                group_parent[name] = parent
            # A renamed primary group keeps Tally's reserved name ("Other Expenses" -> Indirect Expenses);
            # a user-created primary group still says whether it is revenue and which side it sits on.
            cat = RESERVED_GROUPS.get(_norm(meta.get("reservedname") or raw.get("reservedname")))
            if not cat and not parent and _yes(raw.get("isrevenue")):
                cat = "EXPENSE" if _yes(raw.get("isdeemedpositive")) else "INCOME"
            if cat:
                group_category[name] = cat
            continue
        if parent:
            parent_of[name] = parent
        li = ctx.ledgers.get(name) or LedgerInfo(name=getattr(m, "name", name))
        li.parent = parent or li.parent
        if getattr(m, "pan", None):
            li.pan = m.pan
        if getattr(m, "is_tds_applicable", None) is not None:
            li.tds_applicable = m.is_tds_applicable
        taxtype = _norm(_first(raw, ["taxtype", "tax_type", "typeofdutytax", "dutytype"]))
        nature = (getattr(m, "tds_section", None) or getattr(m, "tds_rate_name", None)
                  or _first(raw, ["tdscategoryname", "tdsnatureofpayment", "natureofpayment", "nature_of_payment",
                                  "tdsnature", "tdsratename", "tdscategory"]))
        sec = section_from_text(str(nature)) if nature else None
        if sec:
            li.nature_section = sec
        li.deductee_type = _first(raw, ["tdsdeducteetype", "deducteetype", "deductee_type"]) or li.deductee_type
        if taxtype == "TDS":
            li.is_tds_duty = True
        if "ROUNDING" in _norm(_first(raw, ["vatdealernature", "roundingmethod", "gstappropriateto"])):
            rounding_ledgers.add(name)
        ctx.ledgers[name] = li

    # resolve every ledger to its reserved group by walking the parent chain
    ctx.has_group_data = bool(parent_of or group_parent)
    typed_groups = bool(group_parent or group_category)

    def _resolve_group(start: Optional[str]) -> Optional[str]:
        seen, cur = set(), start
        while cur and cur not in seen:
            seen.add(cur)
            if cur in group_category:
                return group_category[cur]
            if cur in RESERVED_GROUPS:
                return RESERVED_GROUPS[cur]
            nxt = group_parent.get(cur) if typed_groups else parent_of.get(cur)
            if not nxt:
                return _category_from_group_name(cur)
            cur = nxt
        return None

    for name, li in ctx.ledgers.items():
        li.category = (_resolve_group(li.parent) or _category_from_group_name(li.parent)) if li.parent else None
        if name in rounding_ledgers or _ROUND_OFF_RE.search(name):
            li.category = "ROUND_OFF"
        if not li.is_tds_duty and li.category in ("DUTY", "CUR_LIAB", None) and re.search(r"(?<![A-Z])TDS(?![A-Z])|TAX DEDUCTED", name) \
                and not re.search(r"RECEIVABLE|RECOVERABLE|REFUND|INTEREST ON|LATE FEE", name):
            li.is_tds_duty = True
        if li.is_tds_duty and not li.nature_section:
            li.nature_section = section_from_text(name)

    # learn from vouchers where TDS was actually deducted
    seen_guid: Set[str] = set()
    for tx in transactions:
        g = getattr(tx, "guid", None)
        if g in seen_guid:
            continue
        seen_guid.add(g)
        tds_lines = [l for l in tx.ledgers if l.is_credit and ctx.is_tds_line(l)]
        if not tds_lines:
            continue
        sec = (getattr(tx, "tds_section", None)
               or next((section_from_text(l.ledger_name) or (ctx.info(l.ledger_name).nature_section if ctx.info(l.ledger_name) else None)
                        for l in tds_lines if section_from_text(l.ledger_name) or (ctx.info(l.ledger_name) and ctx.info(l.ledger_name).nature_section)), None))
        if not sec:
            continue
        sec = _canon(sec)
        for l in tx.ledgers:
            if l.is_debit and not ctx.is_tds_line(l) and ctx.is_cost_ledger(l.ledger_name, tx.party_name):
                ctx.ledger_votes[_norm(l.ledger_name)][sec] += 1
        if getattr(tx, "party_name", None):
            ctx.party_votes[_norm(tx.party_name)][sec] += 1
    return ctx


def _canon(sec: str) -> str:
    s = _norm(sec).replace("SECTION", "").replace("SEC", "").replace(" ", "").replace("-", "")
    return re.sub(r"\(([AB])\)$", lambda m: f"({m.group(1).lower()})", s)


# --------------------------------------------------------------------------- rule helpers (methods)
def _is_tds_line(self: TDSContext, l: Any) -> bool:
    i = self.info(l.ledger_name)
    if i is not None and (i.category is not None or i.is_tds_duty):
        if i.is_tds_duty:
            return True
        if i.category in ("CUR_ASSET", "ADVANCE"):      # TDS Receivable lives under assets
            return False
    return bool(getattr(l, "is_tds", False))


def _is_cost_ledger(self: TDSContext, ledger_name: str, party_name: Optional[str] = None) -> bool:
    if _norm(ledger_name) == _norm(party_name):
        return False
    cat = self.category(ledger_name)
    if cat is not None:
        return cat in COST_CATEGORIES
    return True   # unknown group: let downstream checks decide


def _is_non_deductee(self: TDSContext, name: Optional[str]) -> Optional[bool]:
    """True/False when group data decides it, None when unknown."""
    cat = self.category(name)
    if cat is None:
        return None
    return cat in NON_DEDUCTEE_CATEGORIES


def _is_outward(self: TDSContext, tx: Any) -> Optional[bool]:
    """A sale / receipt: a SALES/INCOME ledger is credited and no cost ledger is debited. None if unknown."""
    cats_cr = [self.category(l.ledger_name) for l in tx.ledgers if l.is_credit]
    cats_dr = [self.category(l.ledger_name) for l in tx.ledgers if l.is_debit]
    if all(c is None for c in cats_cr + cats_dr):
        return None
    if any(c in ("SALES", "INCOME") for c in cats_cr) and not any(c in COST_CATEGORIES for c in cats_dr):
        return True
    if self.category(getattr(tx, "party_name", None)) == "DEBTOR" and not any(c in COST_CATEGORIES for c in cats_dr):
        return True
    return False


def _partner_capital_ledger(self: TDSContext, tx: Any) -> Optional[Any]:
    for l in tx.ledgers:
        if self.category(l.ledger_name) == "CAPITAL" and l.is_credit:
            return l
    return None


def _suggest_section(self: TDSContext, tx: Any) -> Tuple[Optional[str], str]:
    """(section, reason) for a voucher, from the strongest available signal."""
    party = getattr(tx, "party_name", None)
    for l in tx.ledgers:
        if l.is_credit and self.is_tds_line(l):
            i = self.info(l.ledger_name)
            s = section_from_text(l.ledger_name) or (i.nature_section if i else None)
            if s:
                return _canon(s), f"TDS ledger '{l.ledger_name}'"
    if getattr(tx, "tds_section", None):
        return _canon(tx.tds_section), "section on voucher"
    if self.partner_capital_ledger(tx) and any(self.is_cost_ledger(l.ledger_name, party) for l in tx.ledgers if l.is_debit):
        return "194T", "credited to a partner's Capital Account"
    for l in tx.ledgers:
        if l.is_debit and self.is_cost_ledger(l.ledger_name, party):
            i = self.info(l.ledger_name)
            if i and i.nature_section:
                return _canon(i.nature_section), f"Nature of Payment on ledger '{l.ledger_name}'"
    for l in tx.ledgers:
        if l.is_debit and self.is_cost_ledger(l.ledger_name, party):
            s = self.learned_ledger_section(l.ledger_name)
            if s:
                return s, f"TDS deducted before on ledger '{l.ledger_name}'"
    pi = self.info(party)
    if pi and pi.nature_section:
        return _canon(pi.nature_section), "Nature of Payment on party master"
    if party:
        s = self.learned_party_section(party)
        if s:
            return s, f"TDS deducted before for party '{party}'"
    return None, ""


TDSContext.is_tds_line = _is_tds_line
TDSContext.is_cost_ledger = _is_cost_ledger
TDSContext.is_non_deductee = _is_non_deductee
TDSContext.is_outward = _is_outward
TDSContext.partner_capital_ledger = _partner_capital_ledger
TDSContext.suggest_section = _suggest_section
