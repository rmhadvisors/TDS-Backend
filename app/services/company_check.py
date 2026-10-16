import re
from collections import Counter
from typing import Iterable, Optional, Sequence, Tuple

# Tally prefixes every master GUID and voucher GUID/remote id with the GUID of the company it belongs to,
# e.g. "8c83f0f0-128a-463c-9f83-c11dcb0d90d0-00001e77". That prefix is the same in every export of the same
# company (original or updated) and differs between companies.
_TALLY_ID = re.compile(r"^([0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12})-", re.IGNORECASE)

PARTIES_MISMATCH_ERROR = "Parties do not match. Please upload matching parties data."


def company_id(guids: Iterable[str]) -> Optional[str]:
    """The Tally company GUID most of these ids carry, or None when none of them is a Tally id."""
    counts = Counter(m.group(1).lower() for g in guids if g and (m := _TALLY_ID.match(str(g))))
    return counts.most_common(1)[0][0] if counts else None


def find_company_mismatch(files: Sequence[Tuple[str, Iterable[str]]]) -> Optional[str]:
    """Given (file label, ids in that file) for each upload, names the first file whose company differs from the
    first file that has one; None when they all match. Files without Tally ids are not checked."""
    reference: Optional[Tuple[str, str]] = None
    for label, guids in files:
        cid = company_id(guids)
        if cid is None:
            continue
        if reference is None:
            reference = (label, cid)
        elif cid != reference[1]:
            return f"The {label} is from a different party than the {reference[0]}."
    return None
