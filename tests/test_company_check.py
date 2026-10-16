import io
import json

from fastapi.testclient import TestClient

from app.main import app
from app.services.company_check import PARTIES_MISMATCH_ERROR, company_id, find_company_mismatch

client = TestClient(app)

EDU = "8c83f0f0-128a-463c-9f83-c11dcb0d90d0"
ZEON = "3f3fa0de-cff8-43b7-a25c-63ec831b64f1"


def test_company_id_is_the_tally_guid_prefix():
    assert company_id([f"{EDU}-00001e77", f"{EDU.upper()}-0000001D"]) == EDU
    assert company_id(["M-1", "TX-101", ""]) is None


def test_matching_files_pass():
    files = [(label, [f"{EDU}-0000{i}"]) for i, label in enumerate(["A", "B", "C", "D"])]
    assert find_company_mismatch(files) is None


def test_mismatch_names_the_odd_file():
    files = [
        ("Original Master JSON", [f"{EDU}-0001"]),
        ("Original Transaction JSON", [f"{EDU}-0002"]),
        ("Updated Master JSON", [f"{EDU}-0003"]),
        ("Updated Transaction JSON", [f"{ZEON}-0004"]),
    ]
    assert find_company_mismatch(files) == "The Updated Transaction JSON is from a different party than the Original Master JSON."


def test_files_without_tally_ids_are_not_checked():
    assert find_company_mismatch([("A", ["M-1"]), ("B", [f"{ZEON}-0001"]), ("C", ["TX-1"])]) is None


def _upload(prefix_old, prefix_new):
    def master(prefix):
        return json.dumps([
            {"GUID": f"{prefix}-0001", "NAME": "Aayat Enterprises", "PARENT": "Sundry Creditors", "PAN": "ABCDE1234F"}
        ]).encode("utf-8")

    def tx(prefix):
        return json.dumps([{
            "GUID": f"{prefix}-0101",
            "VOUCHERNUMBER": "101",
            "VOUCHERTYPENAME": "Payment",
            "DATE": "2026-08-01",
            "PARTYLEDGERNAME": "Aayat Enterprises",
            "ALLLEDGERENTRIES.LIST": [
                {"LEDGERNAME": "Aayat Enterprises", "AMOUNT": "-90000.00", "ISDEEMEDPOSITIVE": "No"},
                {"LEDGERNAME": "TDS on Contractors", "AMOUNT": "-10000.00", "ISDEEMEDPOSITIVE": "No", "IS_TDS": "Yes"},
                {"LEDGERNAME": "Contractor Charges", "AMOUNT": "100000.00", "ISDEEMEDPOSITIVE": "Yes"},
            ],
        }]).encode("utf-8")

    files = {
        "master_old": ("Master.json", io.BytesIO(master(prefix_old)), "application/json"),
        "transactions_old": ("Transactions.json", io.BytesIO(tx(prefix_old)), "application/json"),
        "master_new": ("Master1.json", io.BytesIO(master(prefix_new)), "application/json"),
        "transactions_new": ("Transactions1.json", io.BytesIO(tx(prefix_new)), "application/json"),
    }
    return client.post("/api/compare", files=files)


def test_compare_rejects_different_parties():
    res = _upload(EDU, ZEON)
    assert res.status_code == 400
    assert res.json()["error"] == PARTIES_MISMATCH_ERROR


def test_compare_accepts_same_party():
    res = _upload(EDU, EDU)
    assert res.status_code == 200, res.text
