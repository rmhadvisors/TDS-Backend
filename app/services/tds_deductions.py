"""
Booked TDS deductions — the TDS the company actually deducted, read from its own vouchers.

This is what a Form 26Q return reports, so the TDS Working sheet is built from it. Everything is decided from
the structure of each voucher and the dataset's own masters (see tds_context.py); there are no party names,
bank names or client-specific keywords here.

Rules, per voucher:
  1. Optional (memo / provisional) and cancelled vouchers are not part of the books -> skipped.
  2. A TDS line is a TDS duty ledger (under Duties & Taxes). Its section is read from its name or its master
     (Nature of Payment / TDS rate name). A TDS ledger with no section (a single "TDS Payable") counts only when
     it is credited on a bill; the section then comes from the bill (TDS booked before on the same expense
     ledger or party, the ledger's Nature of Payment, statutory keywords, then the rate booked). Sectionless
     debits (year-end transfers, write-offs) are not deductions.
  3. Credit to a TDS ledger = deduction; debit = reversal. In a voucher that also moves bank/cash, a TDS debit
     is a deposit of tax already deducted (challan), not a reversal, and is ignored.
  4. Deductee and amount paid/credited:
       a. the voucher debits expense / purchase / fixed-asset ledgers (a bill): the deductee is the payee
          credited in the voucher, and the amount is those cost lines, excluding GST and round-off. When one bill
          carries two sections, purchase ledgers go to 194Q and other cost ledgers to the other section.
       b. no cost line: the TDS is charged straight to the deductee accounts (Dr partners' capital / Dr party,
          Cr TDS). Each such account is a deductee, its debit is its TDS and amount = TDS / rate.
       c. b. for a party whose bill was booked in a separate voucher with no TDS line (bill on day 1, "Being TDS
          recorded" journal Dr Party / Cr TDS later): the journal is linked to the latest unlinked bill(s) of
          that party, on or before the journal, whose amount (cost lines, ex GST) at a standard rate gives the
          TDS. Amount = those bills, section = from the bill when the TDS ledger names none (per-party ledgers
          such as "TDS on <party>"; when that section's rate is not the rate booked, the section the party's
          payment narrations between bill and journal name, if its rate is). Date = the last bill (TDS is due on credit), unless the bill is in an earlier
          month and the TDS was deposited after its due date (7th of the next month, 30 April for March; the
          deposit is the first bank payment debiting that TDS ledger on or after the journal): then the TDS is
          reported in the month it was actually deducted, the journal's.
     A TDS line whose ledger names no section is never dropped: it takes the section of the bill it is linked to
     (or the party's latest bill), else "Not Specified".
  5. Rate: the rate in the TDS ledger name ("TDS ON CONTRACT @ 1%"); else the statutory rate for the PAN;
     else the rate actually booked, snapped to the standard rate it was worked out at (TDS is rounded to the rupee).
"""
from collections import defaultdict
from dataclasses import dataclass, field
from decimal import ROUND_HALF_UP, Decimal
from typing import Callable, Dict, Iterable, List, Optional, Tuple
import re

from app.models.comparison_models import NormalizedLedger, NormalizedTransaction
from app.services.tds_context import COST_CATEGORIES, TDSContext, section_from_text
from app.services.tds_statutory_rates import get_statutory_rule_by_section, rate_for_pan

_ZERO = Decimal("0.00")
_RATE_IN_NAME = re.compile(r"(\d+(?:\.\d+)?)\s*%")
_SECTION_CODE = re.compile(r"^(19\d[A-Z]*)")
_ROUND_OFF = re.compile(r"(?<![A-Z])ROUND(ED|ING)?\s*(OFF)?(?![A-Z])")
_DUTY_NAME = re.compile(r"(?<![A-Z])(GST|IGST|CGST|SGST|UTGST|CESS|RCM|TCS)(?![A-Z])")
_PAN_RE = re.compile(r"^[A-Z]{5}[0-9]{4}[A-Z]$")
_CAPITAL_PREFIX = re.compile(r"^\s*CAP(ITAL)?\s*[-:.]\s*", re.I)
_PAYEE_SUFFIX = re.compile(r"\s*\((?:LABOUR|LABOUR\s+CHARGES?|CREDITOR)\)", re.I)
_MONEY_CATS = ("BANK", "CASH")


@dataclass
class BookedDeduction:
    section: str
    date: str
    deductee: str
    pan: str
    gross: Decimal
    rate: Decimal
    tds: Decimal
    voucher_numbers: List[str] = field(default_factory=list)
    guids: List[str] = field(default_factory=list)
    ledger: str = ""
    bills: int = 1
    # per bill, in the order the bills were added: (voucher no, last saved in Tally, entered by)
    entries: List[Tuple[str, str, str]] = field(default_factory=list)
    # set when nothing named the section and it was guessed from the rate (or left "Not Specified"): the bill's cost
    # ledgers and party, for the user to add a keyword from (user_keywords.py)
    unresolved_names: List[str] = field(default_factory=list)

    @property
    def computed_tds(self) -> Decimal:
        return (self.gross * self.rate / Decimal("100")).quantize(Decimal("1"), rounding=ROUND_HALF_UP)


def _entry_of(tx: NormalizedTransaction) -> Tuple[str, str, str]:
    """Voucher no, when the voucher was last saved in Tally (UPDATEDDATETIME, 'DD-MM-YYYY HH:MM') and who entered it."""
    raw = tx.raw_reference or {}
    stamp = re.sub(r"\D", "", str(raw.get("updateddatetime") or raw.get("UPDATEDDATETIME") or ""))
    saved = f"{stamp[6:8]}-{stamp[4:6]}-{stamp[0:4]} {stamp[8:10]}:{stamp[10:12]}" if len(stamp) >= 12 else ""
    by = str(raw.get("enteredby") or raw.get("ENTEREDBY") or "").strip()
    return (tx.voucher_number or "", saved, by)


def _section_of(l: NormalizedLedger, ctx: Optional[TDSContext]) -> Optional[str]:
    """Section of a TDS duty line; "" for a TDS line whose ledger names no section; None for other lines."""
    if ctx is not None:
        if not ctx.is_tds_line(l):
            return None
        info = ctx.info(l.ledger_name)
        sec = section_from_text(l.ledger_name) or (info.nature_section if info else None)
    else:
        if not l.is_tds:
            return None
        sec = l.tds_section or section_from_text(l.ledger_name)
    m = _SECTION_CODE.match((sec or "").upper().replace(" ", ""))
    return m.group(1) if m else ""


def _section_from_bill(tx: NormalizedTransaction, ctx: Optional[TDSContext], booked: Decimal, gross: Optional[Decimal]) -> Tuple[str, List[str]]:
    """Section for TDS booked to a sectionless TDS ledger, from the bill itself, and — when nothing named it and it
    was guessed from the rate — the bill's cost ledgers and party (else [])."""
    from app.services.normalizer import suggest_tds_section

    m = _SECTION_CODE.match((suggest_tds_section(tx, ctx)[0] or "").upper().replace(" ", ""))
    if m:
        return m.group(1), []
    names = [l.ledger_name for l in tx.ledgers if l.is_debit and _is_cost(l.ledger_name, ctx)] + [tx.party_name or ""]
    names = list(dict.fromkeys(n for n in names if n))
    if gross:
        pct = abs(booked) * Decimal("100") / abs(gross)
        for lo, hi, sec in _RATE_BANDS:
            if lo <= pct <= hi:
                return sec, names
    return "Not Specified", names


# booked rate -> most common section at that rate (used only when nothing else names the section)
_RATE_BANDS = [
    (Decimal("0.05"), Decimal("0.15"), "194Q"),
    (Decimal("0.8"), Decimal("2.2"), "194C"),
    (Decimal("4.5"), Decimal("5.5"), "194H"),
    (Decimal("9"), Decimal("11"), "194J"),
]


def _category(name: str, ctx: Optional[TDSContext]) -> Optional[str]:
    return ctx.category(name) if ctx is not None else None


def _is_duty_or_round_off(name: str, ctx: Optional[TDSContext]) -> bool:
    n = (name or "").upper()
    if _ROUND_OFF.search(n):
        return True
    cat = _category(name, ctx)
    if cat is not None:                 # the group decides: "Purchase Gold GST @ 3%" is a purchase, not a duty
        return cat == "DUTY"
    return bool(_DUTY_NAME.search(n))


def _is_cost(name: str, ctx: Optional[TDSContext]) -> bool:
    cat = _category(name, ctx)
    if cat is not None:
        return cat in COST_CATEGORIES
    # no group data for this ledger: anything that is not a party/money account counts as the cost
    return ctx is None or not ctx.has_group_data


_STANDARD_RATES = [Decimal(r) for r in ("0.1", "1", "2", "5", "10", "20")]


def _snap_rate(booked: Decimal, gross: Decimal) -> Optional[Decimal]:
    """The standard rate the booked TDS was worked out at, if one gives it within the rupee rounding."""
    fits = [(abs(gross * r / Decimal("100") - booked), r) for r in _STANDARD_RATES]
    diff, rate = min(fits)
    return rate if diff <= Decimal("1") else None


def _rate_for(tds_ledger: str, section: str, pan: str, booked: Decimal, gross: Optional[Decimal]) -> Decimal:
    m = _RATE_IN_NAME.search(tds_ledger or "")
    if m:
        return Decimal(m.group(1))
    rule = get_statutory_rule_by_section(section)
    if rule and _PAN_RE.match(pan or ""):
        return Decimal(rate_for_pan(rule, pan).rstrip("%"))
    if gross:
        return _snap_rate(abs(booked), abs(gross)) or \
            (abs(booked) * Decimal("100") / abs(gross)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    return Decimal(rule.rate_company_firm.rstrip("%")) if rule else Decimal("10")


# A "Dr Party / Cr TDS" journal is looked for among the party's bills of the last _LINK_DAYS days (up to _LINK_BILLS).
_LINK_DAYS = 185
_LINK_BILLS = 12


@dataclass
class _Bill:
    tx: NormalizedTransaction
    base: Decimal
    lines: List[Decimal] = field(default_factory=list)   # the cost lines making up `base`


def _index_separate_bills(transactions: Iterable[NormalizedTransaction], ctx: Optional[TDSContext]) -> Dict[str, List[_Bill]]:
    """Bills with no TDS line of their own, by the account they credit, oldest first. Needs the group data to tell
    a cost ledger from a party ledger, so without it nothing is indexed (and no journal is linked)."""
    out: Dict[str, List[_Bill]] = defaultdict(list)
    if ctx is None or not ctx.has_group_data:
        return out
    seen = set()
    for tx in transactions:
        if tx is None or tx.guid in seen or tx.is_optional or tx.is_cancelled:
            continue
        seen.add(tx.guid)
        if any(_section_of(l, ctx) is not None for l in tx.ledgers):
            continue
        # goods purchases are left to 194Q, where TDS is on the FY excess over Rs 50 lakh, not bill by bill
        cost = [l for l in tx.ledgers if l.is_debit and l.amount and _category(l.ledger_name, ctx) in ("EXPENSE", "FIXED_ASSET")
                and not _is_duty_or_round_off(l.ledger_name, ctx)]
        credited = [l for l in tx.ledgers if l.is_credit and l.amount and _category(l.ledger_name, ctx) not in _MONEY_CATS
                    and _category(l.ledger_name, ctx) not in COST_CATEGORIES and not _is_duty_or_round_off(l.ledger_name, ctx)]
        if not cost or not credited:
            continue
        payee = max(credited, key=lambda l: l.amount).ledger_name
        out[_bare(payee)].append(_Bill(tx, sum((l.amount for l in cost), _ZERO), [l.amount for l in cost]))
    for bills in out.values():
        bills.sort(key=lambda b: (b.tx.date or "", b.tx.voucher_number or ""))
    return out


def _index_tds_deposits(transactions: Iterable[NormalizedTransaction], ctx: Optional[TDSContext]) -> Dict[str, List[str]]:
    """Dates TDS was deposited (challan: a bank/cash payment debiting the TDS ledger), by TDS ledger, oldest first."""
    out: Dict[str, List[str]] = defaultdict(list)
    seen = set()
    for tx in transactions:
        if tx is None or tx.guid in seen or tx.is_optional or tx.is_cancelled:
            continue
        seen.add(tx.guid)
        if not any(_category(l.ledger_name, ctx) in _MONEY_CATS for l in tx.ledgers):
            continue
        for l in tx.ledgers:
            if l.is_debit and _section_of(l, ctx) is not None:
                out[_bare(l.ledger_name)].append(tx.date or "")
    for dates in out.values():
        dates.sort()
    return out


def _deposit_due_date(deduction_date: str) -> str:
    """Rule 30: TDS deducted in a month is due by the 7th of the next month; deducted in March, by 30 April."""
    y, m = int(deduction_date[:4]), int(deduction_date[5:7])
    if m == 3:
        return f"{y}-04-30"
    return f"{y + (m == 12)}-{m % 12 + 1:02d}-07"


def _reporting_date(bill_date: str, journal_date: str, deposits: List[str]) -> str:
    """Bill date, unless the TDS on an earlier month's bill was deposited after that month's due date: then the
    journal date. With no deposit in the data the bill date stays."""
    if (bill_date or "")[:7] == (journal_date or "")[:7]:
        return bill_date
    deposited = next((d for d in deposits if d >= journal_date), None)
    return journal_date if deposited and deposited > _deposit_due_date(bill_date) else bill_date


def _days_between(earlier: str, later: str) -> int:
    from datetime import date
    try:
        return (date.fromisoformat(str(later)[:10]) - date.fromisoformat(str(earlier)[:10])).days
    except ValueError:
        return 0


def _link_bills(bills: List[_Bill], used: set, journal_date: str, tds: Decimal, rates: List[Decimal]) -> Optional[Tuple[List[_Bill], Decimal]]:
    """The bill(s) a TDS journal was worked out on: the latest unlinked bill whose amount at one of `rates` gives the
    TDS, else the latest one where some of its cost lines do (TDS on the fee line of a bill that also carries a
    reimbursement / sundry line; the bill's amount is then those lines), else the combination of recent bills that
    does (the one reaching least far back, then the fewest bills)."""
    cands = [b for b in bills if b.tx.guid not in used and (b.tx.date or "") <= journal_date
             and _days_between(b.tx.date, journal_date) <= _LINK_DAYS][-_LINK_BILLS:][::-1]   # latest first

    def rate_fitting(total: Decimal, n_bills: int) -> Optional[Decimal]:
        tol = max(Decimal("1"), Decimal("0.5") * n_bills)          # TDS rounded to the rupee per bill
        for r in rates:
            if abs(total * r / Decimal("100") - tds) <= tol:
                return r
        return None

    for b in cands:
        r = rate_fitting(b.base, 1)
        if r:
            return [b], r
    for b in cands:
        if not 2 <= len(b.lines) <= 10:
            continue
        for mask in range(1, (1 << len(b.lines)) - 1):
            part = sum((a for i, a in enumerate(b.lines) if mask >> i & 1), _ZERO)
            r = rate_fitting(part, 1)
            if r:
                return [_Bill(b.tx, part, b.lines)], r
    best = None
    for mask in range(1, 1 << len(cands)):
        if mask & (mask - 1) == 0:
            continue                                               # single bills were tried above
        chosen = [i for i in range(len(cands)) if mask >> i & 1]
        r = rate_fitting(sum((cands[i].base for i in chosen), _ZERO), len(chosen))
        if r and (best is None or (chosen[-1], len(chosen)) < best[0]):
            best = ((chosen[-1], len(chosen)), [cands[i] for i in chosen], r)
    return (best[1], best[2]) if best else None


def _section_without_bill(name: str, party_bills: List[_Bill], tx: NormalizedTransaction, ctx: Optional[TDSContext], tds: Decimal) -> Tuple[str, List[str]]:
    """Section for TDS charged to a party by a journal with no matching bill (e.g. the bill is in the previous year's
    books): the party's latest earlier bill, else its first later one, else what earlier deductions for that party
    were, else a user-added keyword in the party's name, else "Not Specified" (the row is still reported). Also returns the
    names to ask the user about when the section was not found (see _section_from_bill)."""
    prior = [b for b in party_bills if (b.tx.date or "") <= (tx.date or "")]
    nearest = prior[-1] if prior else (party_bills[0] if party_bills else None)
    if nearest is not None:
        return _section_from_bill(nearest.tx, ctx, tds, None)
    m = _SECTION_CODE.match(((ctx.learned_party_section(name) if ctx is not None else None) or "").upper())
    if m:
        return m.group(1), []
    from app.services.user_keywords import match_user_rule
    hit = match_user_rule(name)
    m = _SECTION_CODE.match(((hit[0].sub_section or hit[0].section_code) if hit else "").upper())
    return (m.group(1), []) if m else ("Not Specified", [name])


def _index_narrations(transactions: Iterable[NormalizedTransaction]) -> Dict[str, List[Tuple[str, str]]]:
    """(date, narration) of every voucher, by each account it posts to."""
    out: Dict[str, List[Tuple[str, str]]] = defaultdict(list)
    seen = set()
    for tx in transactions:
        if tx is None or tx.guid in seen or not tx.narration:
            continue
        seen.add(tx.guid)
        for name in {_bare(l.ledger_name) for l in tx.ledgers}:
            out[name].append((tx.date or "", tx.narration))
    return out


def _section_fitting_rate(sec: str, rate: Decimal, pan: str, notes: List[Tuple[str, str]], start: str, end: str) -> str:
    """Section taken from a bill whose statutory rate for the PAN is not the rate booked (a "Professional Services"
    bill with TDS at 2%): when the party's vouchers between the bill and the TDS journal name the nature of the
    payment ("... Accounting Contract work") and that section's rate is the one booked, it is that section."""
    from app.services.tds_statutory_rates import suggest_statutory_rule

    def statutory(rule) -> Optional[Decimal]:
        try:
            return Decimal(rate_for_pan(rule, pan).rstrip("%"))
        except ArithmeticError:          # "Rates in Force / DTAA", "20/30/40%"
            return None

    rule = get_statutory_rule_by_section(sec)
    if rule is None or not _PAN_RE.match(pan or "") or statutory(rule) in (None, rate):
        return sec
    for date, text in notes:
        if not start <= date <= end:
            continue
        r = suggest_statutory_rule(text)
        if r and r.section_code != sec and statutory(r) == rate:
            return r.section_code
    return sec


def _rates_to_try(tds_ledger: str, section: str, pan: str) -> List[Decimal]:
    """Rates a journal's TDS may have been worked out at: the one its ledger / section and PAN give first."""
    m = _RATE_IN_NAME.search(tds_ledger or "")
    if m:
        return [Decimal(m.group(1))]
    first: List[Decimal] = []
    rule = get_statutory_rule_by_section(section) if section else None
    if rule:
        rates = [rate_for_pan(rule, pan)] if _PAN_RE.match(pan or "") else [rule.rate_company_firm, rule.rate_individual_huf]
        for s in rates:
            try:
                first.append(Decimal(s.rstrip("%")))
            except ArithmeticError:
                pass
    return first + [r for r in _STANDARD_RATES if r not in first]


def _base_used_for_tds(pool: List[NormalizedLedger], gross: Decimal, rate: Decimal, booked: Decimal) -> Decimal:
    """
    The amount the booked TDS was worked out on. Normally all cost lines of the bill; when TDS at the rate on
    all of them does not give the TDS booked (> Rs 1 apart) but a subset of the lines does — e.g. TDS taken on the
    gold line of a bill that also has a stone line — that subset is the amount paid/credited.
    """
    if not rate or len(pool) < 2 or len(pool) > 10:
        return gross
    tds_on = lambda amt: amt * rate / Decimal("100")
    if abs(tds_on(gross) - booked) <= Decimal("1"):
        return gross
    best: Optional[Decimal] = None
    amounts = [l.amount for l in pool]
    for mask in range(1, (1 << len(amounts)) - 1):
        s = sum((a for i, a in enumerate(amounts) if mask >> i & 1), _ZERO)
        if abs(tds_on(s) - booked) <= Decimal("0.5") and (best is None or s > best):
            best = s
    return best if best is not None else gross


def _display_name(ledger_name: str, ctx: Optional[TDSContext]) -> str:
    name = ledger_name or ""
    if _category(name, ctx) == "CAPITAL" or _CAPITAL_PREFIX.match(name):
        name = _CAPITAL_PREFIX.sub("", name)
    return _PAYEE_SUFFIX.sub("", name).strip() or ledger_name


def extract_booked_deductions(
    transactions: Iterable[NormalizedTransaction],
    ctx: Optional[TDSContext],
    pan_of: Callable[[str, NormalizedTransaction], str],
    tds_deposits: Optional[Iterable[NormalizedTransaction]] = None,
) -> List[BookedDeduction]:
    """tds_deposits: vouchers to read TDS challans from (default: `transactions`), for when those were filtered out."""
    out: List[BookedDeduction] = []
    transactions = [t for t in transactions if t is not None]
    separate_bills = _index_separate_bills(transactions, ctx)
    deposits = _index_tds_deposits(transactions if tds_deposits is None else tds_deposits, ctx)
    narrations = _index_narrations(transactions if tds_deposits is None else tds_deposits)
    linked: set = set()
    seen = set()
    for tx in transactions:
        if tx is None or tx.guid in seen or tx.is_optional or tx.is_cancelled:
            continue
        seen.add(tx.guid)

        tds_lines: List[Tuple[NormalizedLedger, str]] = []
        rest: List[NormalizedLedger] = []
        for l in tx.ledgers:
            if l.amount == _ZERO:
                continue
            sec = _section_of(l, ctx)
            if sec is not None:
                tds_lines.append((l, sec))
            elif not _is_duty_or_round_off(l.ledger_name, ctx):
                rest.append(l)
        if not tds_lines:
            # Exports without ledger detail carry the deduction on the voucher itself.
            if not tx.ledgers and tx.tds_amount and tx.tds_amount > _ZERO:
                out.append(_voucher_level_deduction(tx, ctx, pan_of))
            continue

        unresolved: Dict[str, List[str]] = {}
        moves_money = any(_category(l.ledger_name, ctx) in _MONEY_CATS for l in rest)
        net: Dict[str, Decimal] = defaultdict(Decimal)
        ledger_of: Dict[str, str] = {}
        for l, sec in tds_lines:
            if l.is_debit and moves_money:
                continue                       # tax deposited (challan), not a reversal
            net[sec] += l.amount if l.is_credit else -l.amount
            ledger_of.setdefault(sec, l.ledger_name)

        payees = [l for l in rest if _category(l.ledger_name, ctx) not in _MONEY_CATS]
        cost = [l for l in payees if l.is_debit and _is_cost(l.ledger_name, ctx)]
        party_side = [l for l in payees if l not in cost]

        if "" in net:
            unnamed = net.pop("")
            if cost and unnamed > _ZERO:
                sec, names = _section_from_bill(tx, ctx, unnamed, sum((l.amount for l in cost), _ZERO))
                if names:
                    unresolved[sec] = names
                net[sec] += unnamed
                ledger_of.setdefault(sec, ledger_of[""])
            elif not cost and unnamed > _ZERO:
                net[""] = unnamed      # Dr Party / Cr "TDS on <party>": section comes from the party's bill below

        for sec, booked in net.items():
            if booked == _ZERO:
                continue
            sign = Decimal("1") if booked > 0 else Decimal("-1")
            tds_ledger = ledger_of[sec]

            if cost:
                credited = [l for l in party_side if l.is_credit]
                payee = max(credited, key=lambda l: l.amount).ledger_name if credited else (tx.party_name or "")
                pool = cost
                if len(net) > 1:
                    pool = [l for l in cost if (_category(l.ledger_name, ctx) == "PURCHASE") == (sec == "194Q")]
                gross = sum((l.amount for l in pool), _ZERO) or None
                pan = pan_of(payee, tx)
                rate = _rate_for(tds_ledger, sec, pan, booked, gross)
                if gross is None:
                    gross = abs(booked) * Decimal("100") / rate if rate else _ZERO
                else:
                    gross = _base_used_for_tds(pool, gross, rate, abs(booked))
                out.append(BookedDeduction(
                    section=sec, date=tx.date, deductee=_display_name(payee, ctx), pan=pan,
                    gross=(gross * sign).quantize(Decimal("0.01")), rate=rate, tds=booked,
                    voucher_numbers=[tx.voucher_number or ""], guids=[tx.guid], ledger=tds_ledger, entries=[_entry_of(tx)],
                    unresolved_names=list(unresolved.get(sec, [])),
                ))
                continue

            # No bill in this voucher: TDS charged straight to the deductee account(s).
            side = [l for l in party_side if (l.is_debit if booked > 0 else l.is_credit)]
            if not side:
                side_amounts = [(tx.party_name or "", abs(booked))]
            else:
                total = sum((l.amount for l in side), _ZERO)
                side_amounts = [(l.ledger_name, l.amount * abs(booked) / total) for l in side]
            for name, share in side_amounts:
                pan = pan_of(name, tx)
                d_sec, date, bills, names = sec, tx.date, [], []
                party_bills = separate_bills.get(_bare(name), []) \
                    if booked > 0 and sec != "194Q" and _category(name, ctx) != "CAPITAL" else []
                if party_bills:
                    hit = _link_bills(party_bills, linked, tx.date or "", share, _rates_to_try(tds_ledger, sec, pan))
                    if hit:
                        bills, rate = hit
                        linked.update(b.tx.guid for b in bills)
                        gross = sum((b.base for b in bills), _ZERO)
                        date = _reporting_date(max(b.tx.date for b in bills), tx.date or "", deposits.get(_bare(tds_ledger), []))
                        if not d_sec:
                            d_sec, names = _section_from_bill(bills[-1].tx, ctx, share, gross)
                            fitted = _section_fitting_rate(d_sec, rate, pan, narrations.get(_bare(name), []),
                                                           min(b.tx.date or "" for b in bills), tx.date or "")
                            if fitted != d_sec:
                                d_sec, names = fitted, []
                if not bills:
                    if not d_sec:
                        d_sec, names = _section_without_bill(name, party_bills, tx, ctx, share)
                    rate = _rate_for(tds_ledger, d_sec, pan, share, None)
                    gross = share * Decimal("100") / rate if rate else _ZERO
                out.append(BookedDeduction(
                    section=d_sec, date=date, deductee=_display_name(name, ctx), pan=pan,
                    gross=(gross * sign).quantize(Decimal("0.01")), rate=rate,
                    tds=(share * sign).quantize(Decimal("0.01")),
                    voucher_numbers=[tx.voucher_number or ""] + [b.tx.voucher_number or "" for b in bills],
                    guids=[tx.guid] + [b.tx.guid for b in bills], ledger=tds_ledger, bills=max(1, len(bills)),
                    entries=[_entry_of(tx)] + [_entry_of(b.tx) for b in bills], unresolved_names=list(names),
                ))
    return out


def _voucher_level_deduction(tx: NormalizedTransaction, ctx: Optional[TDSContext], pan_of: Callable[[str, NormalizedTransaction], str]) -> BookedDeduction:
    from app.services.normalizer import suggest_tds_section

    sec = tx.tds_section or suggest_tds_section(tx, ctx)[0] or ""
    m = _SECTION_CODE.match(sec.upper().replace(" ", ""))
    sec = m.group(1) if m else "Not Specified"
    names = [] if m else [n for n in [tx.party_name or ""] if n]
    pan = pan_of(tx.party_name or "", tx) or str((tx.raw_reference or {}).get("pan") or "").replace(" ", "").upper()
    gross = tx.total_amount or _ZERO
    rate = tx.tds_rate if tx.tds_rate else _rate_for("", sec, pan, tx.tds_amount, gross)
    return BookedDeduction(
        section=sec, date=tx.date, deductee=tx.party_name or "", pan=pan, gross=gross, rate=rate,
        tds=tx.tds_amount, voucher_numbers=[tx.voucher_number or ""], guids=[tx.guid], entries=[_entry_of(tx)],
        unresolved_names=names,
    )


THRESHOLD_194Q = Decimal("5000000.00")


def _goods_bills_194q(
    transactions: Iterable[NormalizedTransaction], ctx: Optional[TDSContext],
) -> List[Tuple[NormalizedTransaction, Optional[str], Decimal, Decimal]]:
    """
    (bill, seller, goods value, part above Rs 50 lakh) for every goods bill, in date order. Goods value is the
    purchase lines ex-GST plus, on a purchase voucher, the freight/packing etc. the supplier charges on the same
    invoice. The Rs 50 lakh is per seller per FY; seller is None when the bill names none (nothing is taxable).
    """
    from app.services.normalizer import _fy_of

    txs = sorted((t for t in transactions if t is not None and not t.is_optional and not t.is_cancelled),
                 key=lambda t: (t.date or "", t.voucher_number or ""))
    cumulative: Dict[Tuple[str, str], Decimal] = defaultdict(Decimal)
    out: List[Tuple[NormalizedTransaction, Optional[str], Decimal, Decimal]] = []
    for tx in txs:
        goods = sum(((l.amount if l.is_debit else -l.amount) for l in tx.ledgers
                     if _category(l.ledger_name, ctx) == "PURCHASE"), _ZERO)
        if goods == _ZERO:
            continue
        if "PURCHASE" in (tx.voucher_type or "").upper():
            # freight, packing etc. on the supplier's goods invoice are part of the purchase value (ex-GST)
            goods += sum(((l.amount if l.is_debit else -l.amount) for l in tx.ledgers
                          if _category(l.ledger_name, ctx) == "EXPENSE" and not _is_duty_or_round_off(l.ledger_name, ctx)), _ZERO)
        sellers = [l for l in tx.ledgers if _category(l.ledger_name, ctx) not in ("PURCHASE", "DUTY", "ROUND_OFF", "EXPENSE", "BANK", "CASH")
                   and _section_of(l, ctx) is None]
        if not sellers:
            out.append((tx, None, goods, _ZERO))
            continue
        seller = max(sellers, key=lambda l: l.amount).ledger_name
        key = (seller.upper(), _fy_of(tx.date))
        before = cumulative[key]
        cumulative[key] += goods
        taxable = min(max(cumulative[key] - max(before, THRESHOLD_194Q), _ZERO), goods)
        out.append((tx, seller, goods, taxable))
    return out


def taxable_194q_by_bill(transactions: Iterable[NormalizedTransaction], ctx: Optional[TDSContext]) -> Dict[str, Decimal]:
    """guid -> the part of that goods bill above the seller's Rs 50 lakh for the FY (0 while under it)."""
    return {tx.guid: taxable for tx, _, _, taxable in _goods_bills_194q(transactions, ctx)}


def find_194q_not_deducted(
    transactions: Iterable[NormalizedTransaction],
    ctx: Optional[TDSContext],
    booked: List[BookedDeduction],
    pan_of: Callable[[str, NormalizedTransaction], str],
) -> List[BookedDeduction]:
    """
    Section 194Q TDS that was due but not booked: purchase-of-goods value (ex-GST) from a supplier above
    Rs 50 lakh in the FY, on bills that carry no TDS line. A month-end journal deducting 194Q from that supplier
    (Dr Supplier / Cr TDS, no bill) covers that month's bills first. Rows carry tds = 0 so the working shows the
    TDS due and the shortfall.
    """
    bill_guids: set = set()
    candidates: List[Tuple[NormalizedTransaction, str, Decimal]] = []
    for tx, seller, goods, taxable in _goods_bills_194q(transactions, ctx):
        bill_guids.add(tx.guid)
        if seller is None:
            continue
        if any(l.is_credit and _section_of(l, ctx) is not None for l in tx.ledgers):
            continue                              # TDS booked on the bill itself
        if taxable > _ZERO:
            candidates.append((tx, seller, taxable))

    # 194Q deducted by journal (no bill in the voucher) for a supplier in a month covers that month's bills
    cover: Dict[Tuple[str, str], Decimal] = defaultdict(Decimal)
    for d in booked:
        if d.section == "194Q" and d.guids and d.guids[0] not in bill_guids:
            cover[(_bare(d.deductee), d.date[:7])] += d.tds

    out: List[BookedDeduction] = []
    for tx, seller, taxable in candidates:
        pan = pan_of(seller, tx)
        rate = _rate_for("", "194Q", pan, _ZERO, None)
        due = taxable * rate / Decimal("100")
        k = (_bare(_display_name(seller, ctx)), tx.date[:7])
        used = min(cover[k], due)
        cover[k] -= used
        if due - used < Decimal("0.5"):
            continue
        out.append(BookedDeduction(
            section="194Q", date=tx.date, deductee=_display_name(seller, ctx), pan=pan,
            gross=(taxable * (due - used) / due).quantize(Decimal("0.01")), rate=rate, tds=_ZERO,
            voucher_numbers=[tx.voucher_number or ""], guids=[tx.guid], ledger="", entries=[_entry_of(tx)],
        ))
    return out


def _bare(name: str) -> str:
    return re.sub(r"[^A-Z0-9]", "", (name or "").upper())


def _deductee_key(d: BookedDeduction) -> str:
    return d.pan or re.sub(r"[^A-Z0-9]", "", d.deductee.upper())


def group_by_deductee_month(deductions: Iterable[BookedDeduction]) -> List[BookedDeduction]:
    """One row per section, deductee, month and rate, as deductions are reported in a 26Q return.
    The row is dated on the deductee's last bill of the month."""
    grouped: Dict[Tuple[str, str, str, Decimal], BookedDeduction] = {}
    for d in sorted(deductions, key=lambda x: (x.date, x.deductee.upper())):
        key = (d.section, _deductee_key(d), d.date[:7], d.rate)
        g = grouped.get(key)
        if g is None:
            grouped[key] = BookedDeduction(**{**d.__dict__, "voucher_numbers": list(d.voucher_numbers), "guids": list(d.guids),
                                              "entries": list(d.entries), "unresolved_names": list(d.unresolved_names)})
            continue
        g.gross += d.gross
        g.tds += d.tds
        g.date = max(g.date, d.date)
        g.bills += d.bills
        g.voucher_numbers += [v for v in d.voucher_numbers if v and v not in g.voucher_numbers]
        g.entries += [e for gid, e in zip(d.guids, d.entries) if gid not in g.guids]
        g.guids += d.guids
        if not g.pan and d.pan:
            g.pan = d.pan
        g.unresolved_names += [n for n in d.unresolved_names if n not in g.unresolved_names]
    # a deduction fully reversed in the same month (bill + "TDS reversed" journal) leaves no row
    rows = [g for g in grouped.values()
            if abs(g.tds) >= Decimal("0.5") or (g.ledger == "" and g.computed_tds != _ZERO)]
    return sorted(rows, key=lambda x: (x.date, x.deductee.upper(), x.section))
