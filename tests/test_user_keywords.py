import pytest

from app.services.tds_statutory_rates import find_matching_rules, suggest_statutory_rule
from app.services.user_keywords import add_user_keyword, list_user_keywords, match_user_rule, remove_user_keyword


def test_no_user_keywords_by_default():
    assert list_user_keywords() == []
    assert match_user_rule("Landscaping Charges") is None


def test_user_keyword_matches_other_ledgers_with_the_word():
    add_user_keyword("landscap", "194C")
    rule, kw = match_user_rule("Landscaping Charges A/c")
    assert (rule.section_code, kw) == ("194C", "LANDSCAP")
    assert match_user_rule("Office Landscaping") is not None
    # start of a word only
    assert match_user_rule("Warelandscaping") is None
    assert suggest_statutory_rule("Landscaping Charges").section_code == "194C"


def test_statutory_keyword_wins_over_user_keyword():
    builtin = suggest_statutory_rule("Professional Fees")
    assert builtin is not None
    add_user_keyword("PROFESSIONAL", "194C")
    assert "194C" not in {r.section_code for r, _ in find_matching_rules("Professional Fees")}
    assert suggest_statutory_rule("Professional Fees").section_code == builtin.section_code


def test_rejects_unusable_keywords():
    with pytest.raises(ValueError):
        add_user_keyword("ab", "194C")
    with pytest.raises(ValueError):
        add_user_keyword("TDS Payable", "194C")
    with pytest.raises(ValueError):
        add_user_keyword("LANDSCAPING", "999Z")


def test_repoint_and_remove():
    add_user_keyword("Landscaping", "194C")
    add_user_keyword("landscaping", "194J")
    assert list_user_keywords() == [{"keyword": "LANDSCAPING", "section": "194J(b)"}]
    assert remove_user_keyword("LANDSCAPING")
    assert not remove_user_keyword("LANDSCAPING")
    assert match_user_rule("Landscaping") is None
