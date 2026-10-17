import io
import json
from datetime import datetime, timezone

from fastapi.testclient import TestClient

from app.api import comparison
from app.main import app
from app.services import report_store

client = TestClient(app)

MASTER = json.dumps([
    {"GUID": "M-1", "NAME": "Aayat Enterprises", "PARENT": "Sundry Creditors", "PAN": "ABCDE1234F", "ISTDSAPPLICABLE": "Yes", "TDSSECTION": "194C"}
]).encode("utf-8")
TRANSACTIONS = json.dumps([
    {
        "GUID": "TX-101",
        "VOUCHERNUMBER": "101",
        "VOUCHERTYPENAME": "Payment",
        "DATE": "2026-08-01",
        "PARTYLEDGERNAME": "Aayat Enterprises",
        "ALLLEDGERENTRIES.LIST": [
            {"LEDGERNAME": "Aayat Enterprises", "AMOUNT": "-90000.00", "ISDEEMEDPOSITIVE": "No"},
            {"LEDGERNAME": "TDS on Contractors", "AMOUNT": "-10000.00", "ISDEEMEDPOSITIVE": "No", "IS_TDS": "Yes"},
            {"LEDGERNAME": "Contractor Charges", "AMOUNT": "100000.00", "ISDEEMEDPOSITIVE": "Yes"},
        ],
    }
]).encode("utf-8")


def _compare():
    files = {
        "master_old": ("Master.json", io.BytesIO(MASTER), "application/json"),
        "transactions_old": ("Transactions.json", io.BytesIO(TRANSACTIONS), "application/json"),
        "master_new": ("Master1.json", io.BytesIO(MASTER), "application/json"),
        "transactions_new": ("Transactions1.json", io.BytesIO(TRANSACTIONS), "application/json"),
    }
    res = client.post("/api/compare", files=files)
    assert res.status_code == 200, res.text
    return res.json()["report"]


def test_report_is_kept_in_history_and_downloadable_after_memory_expires():
    report = _compare()
    assert report["saved_to_history"] is True
    assert report["filename"].startswith("TDS_Report_Aayat-Enterprises_01-08-2026_")

    history = client.get("/api/history").json()
    assert history["storage"] == "local"
    assert [r["report_id"] for r in history["reports"]] == [report["download_id"]]
    assert history["reports"][0]["party_name"] == "Aayat Enterprises"
    assert history["reports"][0]["period"] == "01/08/2026"
    assert history["reports"][0]["filename"] == report["filename"]

    comparison.REPORT_STORAGE.clear()  # as after the 1-hour memory cache or a server restart
    dl = client.get(f"/api/download/{report['download_id']}")
    assert dl.status_code == 200
    assert report["filename"] in dl.headers["content-disposition"]
    assert dl.content[:2] == b"PK"


def test_history_search_and_newest_first():
    first = _compare()["download_id"]
    second = _compare()["download_id"]
    ids = [r["report_id"] for r in client.get("/api/history").json()["reports"]]
    assert set(ids) == {first, second}
    assert client.get("/api/history", params={"search": "aayat"}).json()["reports"]
    assert client.get("/api/history", params={"search": "someone else"}).json()["reports"] == []


def test_unknown_report_is_404():
    assert client.get("/api/download/00000000-0000-0000-0000-000000000000").status_code == 404
    assert client.get("/api/download/not-an-id").status_code == 404


def test_object_name_round_trips_any_party_and_period():
    created = datetime(2026, 10, 15, 11, 42, 0, tzinfo=timezone.utc)
    rid = "1b9d6bcd-bbfd-4b2d-9b5d-ab8dfbbd4bed"
    for party, period in [("M/s. A__B & Co (Pune)", "Apr–Sep 2026"), (None, None), ("Zeon_Life", "Q2/FY26-27")]:
        name = report_store.make_object_name(rid, party, period, created)
        assert "/" not in name
        parsed = report_store.parse_object_name(name, 10, name)
        assert (parsed.report_id, parsed.party_name, parsed.period) == (rid, party, period)
        assert parsed.created_at == "2026-10-15T11:42:00Z"
    assert report_store.parse_object_name("notes.txt", 1, "notes.txt") is None


def test_s3_store_lists_by_prefix_and_signs_attachment_download(monkeypatch):
    created = datetime(2026, 10, 15, 11, 42, 0, tzinfo=timezone.utc)
    name = report_store.make_object_name("1b9d6bcd-bbfd-4b2d-9b5d-ab8dfbbd4bed", "MK Shah", "FY 2026-27", created)

    class FakeS3:
        def __init__(self):
            self.put = []

        def put_object(self, **kw):
            self.put.append(kw)

        def get_paginator(self, _):
            class P:
                def paginate(_, Bucket, Prefix):
                    return [{"Contents": [{"Key": Prefix + name, "Size": 2048}, {"Key": Prefix + "junk.txt", "Size": 1}]}]
            return P()

        def generate_presigned_url(self, op, Params, ExpiresIn):
            return f"https://s3.example/{Params['Key']}?{Params['ResponseContentDisposition']}&{ExpiresIn}"

    store = report_store.S3ReportStore.__new__(report_store.S3ReportStore)
    store.bucket, store.prefix, store.client = "bucket", "reports/", FakeS3()
    store.save(name, b"PK")
    assert store.client.put[0]["Key"] == "reports/" + name
    assert store.client.put[0]["ServerSideEncryption"] == "AES256"

    [listed] = store.list()
    assert (listed.party_name, listed.size, listed.key) == ("MK Shah", 2048, "reports/" + name)
    url = store.presigned_url(listed)
    assert 'filename="TDS_Report_MK-Shah_FY-2026-27_2026-10-15_1142.xlsx"' in url


def test_delete_removes_report_from_history_and_downloads():
    rid = _compare()["download_id"]
    res = client.delete(f"/api/history/{rid}")
    assert res.status_code == 200 and res.json() == {"deleted": rid}
    assert client.get("/api/history").json()["reports"] == []
    assert client.get(f"/api/download/{rid}").status_code == 404  # gone from the memory cache too
    assert client.delete(f"/api/history/{rid}").status_code == 404


def test_s3_delete_removes_the_object():
    deleted = []

    class FakeS3:
        def delete_object(self, **kw):
            deleted.append(kw)

    store = report_store.S3ReportStore.__new__(report_store.S3ReportStore)
    store.bucket, store.client = "bucket", FakeS3()
    created = datetime(2026, 10, 15, 11, 42, 0, tzinfo=timezone.utc)
    name = report_store.make_object_name("1b9d6bcd-bbfd-4b2d-9b5d-ab8dfbbd4bed", "MK Shah", None, created)
    store.delete(report_store.parse_object_name(name, 1, "reports/" + name))
    assert deleted == [{"Bucket": "bucket", "Key": "reports/" + name}]
