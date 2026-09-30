import io
import json
import os
from decimal import Decimal
import pytest
from fastapi.testclient import TestClient
from app.main import app

client = TestClient(app)


def test_health_endpoint():
    res = client.get("/api/health")
    assert res.status_code == 200
    assert res.json()["status"] == "ok"


def test_compare_endpoint_flow():
    # Sample original data
    master_old = json.dumps([
        {"GUID": "M-1", "NAME": "Aayat Enterprises", "PARENT": "Sundry Creditors", "PAN": "ABCDE1234F", "ISTDSAPPLICABLE": "Yes", "TDSSECTION": "194C"}
    ]).encode("utf-8")

    tx_old = json.dumps([
        {
            "GUID": "TX-101",
            "VOUCHERNUMBER": "101",
            "VOUCHERTYPENAME": "Payment",
            "DATE": "2026-08-01",
            "PARTYLEDGERNAME": "Aayat Enterprises",
            "ALLLEDGERENTRIES.LIST": [
                {"LEDGERNAME": "Aayat Enterprises", "AMOUNT": "-90000.00", "ISDEEMEDPOSITIVE": "No"},
                {"LEDGERNAME": "TDS on Contractors", "AMOUNT": "-10000.00", "ISDEEMEDPOSITIVE": "No", "IS_TDS": "Yes"},
                {"LEDGERNAME": "Contractor Charges", "AMOUNT": "100000.00", "ISDEEMEDPOSITIVE": "Yes"}
            ]
        }
    ]).encode("utf-8")

    master_new = master_old
    tx_new = tx_old

    files = {
        "master_old": ("Master.json", io.BytesIO(master_old), "application/json"),
        "transactions_old": ("Transactions.json", io.BytesIO(tx_old), "application/json"),
        "master_new": ("Master1.json", io.BytesIO(master_new), "application/json"),
        "transactions_new": ("Transactions1.json", io.BytesIO(tx_new), "application/json"),
    }

    res = client.post("/api/compare", files=files)
    assert res.status_code == 200, res.text
    data = res.json()
    assert data["success"] is True
    assert data["summary"]["original_transactions"] == 1
    assert data["summary"]["updated_transactions"] == 1
    assert data["summary"]["unchanged"] == 1
    assert data["reconciliation"]["is_valid"] is True
    assert data["reconciliation"]["transaction_count"] == "PASS"

    download_id = data["report"]["download_id"]
    dl_res = client.get(f"/api/download/{download_id}")
    assert dl_res.status_code == 200
    assert dl_res.headers["content-type"] == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    assert len(dl_res.content) > 1000


def test_compare_with_sample_files():
    sample_dir = os.path.join(os.path.dirname(__file__), "..", "..", "sample_data")
    with open(os.path.join(sample_dir, "Master.json"), "rb") as f:
        m_old = f.read()
    with open(os.path.join(sample_dir, "Transactions.json"), "rb") as f:
        t_old = f.read()
    with open(os.path.join(sample_dir, "Master1.json"), "rb") as f:
        m_new = f.read()
    with open(os.path.join(sample_dir, "Transactions1.json"), "rb") as f:
        t_new = f.read()

    files = {
        "master_old": ("Master.json", io.BytesIO(m_old), "application/json"),
        "transactions_old": ("Transactions.json", io.BytesIO(t_old), "application/json"),
        "master_new": ("Master1.json", io.BytesIO(m_new), "application/json"),
        "transactions_new": ("Transactions1.json", io.BytesIO(t_new), "application/json"),
    }

    res = client.post("/api/compare", files=files)
    assert res.status_code == 200, res.text
    data = res.json()
    assert data["success"] is True
    summary = data["summary"]
    assert summary["original_transactions"] == 133
    assert summary["updated_transactions"] == 117
    assert summary["unchanged"] == 117
    assert summary["missing"] == 16
    assert summary["modified"] == 0
    assert summary["recreated"] == 0
    assert summary["new"] == 0
    assert summary["unresolved"] == 0

    assert data["reconciliation"]["is_valid"] is True
    assert data["reconciliation"]["transaction_count"] == "PASS"
    assert data["reconciliation"]["amount"] == "PASS"
    assert data["reconciliation"]["tds"] == "PASS"


def test_compare_with_utf16_and_bom_encodings():
    master_json = json.dumps([
        {"GUID": "M-1", "NAME": "Aayat Enterprises", "PARENT": "Sundry Creditors", "PAN": "ABCDE1234F", "ISTDSAPPLICABLE": "Yes", "TDSSECTION": "194C"}
    ])

    tx_json = json.dumps([
        {
            "GUID": "TX-101",
            "VOUCHERNUMBER": "101",
            "VOUCHERTYPENAME": "Payment",
            "DATE": "2026-08-01",
            "PARTYLEDGERNAME": "Aayat Enterprises",
            "ALLLEDGERENTRIES.LIST": [
                {"LEDGERNAME": "Aayat Enterprises", "AMOUNT": "-90000.00", "ISDEEMEDPOSITIVE": "No"},
                {"LEDGERNAME": "TDS on Contractors", "AMOUNT": "-10000.00", "ISDEEMEDPOSITIVE": "No", "IS_TDS": "Yes"},
                {"LEDGERNAME": "Contractor Charges", "AMOUNT": "100000.00", "ISDEEMEDPOSITIVE": "Yes"}
            ]
        }
    ])

    # Test UTF-16 LE (with BOM), UTF-16 BE, and UTF-8 with BOM
    files = {
        "master_old": ("Master.json", io.BytesIO(master_json.encode("utf-16")), "application/json"),
        "transactions_old": ("Transactions.json", io.BytesIO(tx_json.encode("utf-16-le")), "application/json"),
        "master_new": ("Master1.json", io.BytesIO(master_json.encode("utf-8-sig")), "application/json"),
        "transactions_new": ("Transactions1.json", io.BytesIO(tx_json.encode("utf-16-be")), "application/json"),
    }

    res = client.post("/api/compare", files=files)
    assert res.status_code == 200, res.text
    data = res.json()
    assert data["success"] is True
    assert data["summary"]["original_transactions"] == 1


def test_corrupted_json_error():
    files = {
        "master_old": ("Master.json", io.BytesIO(b"INVALID JSON {{{{"), "application/json"),
        "transactions_old": ("Transactions.json", io.BytesIO(b"[]"), "application/json"),
        "master_new": ("Master1.json", io.BytesIO(b"[]"), "application/json"),
        "transactions_new": ("Transactions1.json", io.BytesIO(b"[]"), "application/json"),
    }
    res = client.post("/api/compare", files=files)
    assert res.status_code == 400
    assert "error" in res.json()
