"""
Regression guard: the reports for the client datasets that are already verified must not change.

Each dataset in DATASETS is run through the full /api/compare pipeline and every sheet of the generated workbook is
compared cell by cell with the approved copy in tests/baselines/<name>.json. Any difference fails the test and lists
the rows that changed.

When a change is intended (a fix that makes a report more correct), review the listed differences and, once they
are right, approve them by re-saving the baselines:

    venv/Scripts/python.exe -m tests.test_regression_baselines --update
"""
import io
import json
import sys
from pathlib import Path

import openpyxl
import pytest
from fastapi.testclient import TestClient

from app.main import app

BACKEND = Path(__file__).resolve().parent.parent
BASELINES = Path(__file__).resolve().parent / "baselines"

# name -> (original master, original transactions, updated master, updated transactions), relative to backend/
DATASETS = {
    "eduguide": ("originalData/eduguide/Eduguide 25-26Master.json", "originalData/eduguide/Edu apr 25 to feb 26Transactions.json",
                 "changedData/eduguide/Eduguide 25-26Master.json", "changedData/eduguide/Eduguide 25-26Transactions.json"),
    "mk_shah": ("originalData/mk shah/Master.json", "originalData/mk shah/febtrsnsaction.json",
                "changedData/mk shah/Master.json", "changedData/mk shah/Transactions.json"),
    "zeon": ("originalData/zeon/old.json", "originalData/zeon/transactions old.json",
             "changedData/zeon/Master.json", "changedData/zeon/Transactionnew.json"),
    "rmh": ("originalData/rmh/RMH advisors 25-26Master.json", "originalData/rmh/RMH advisorsapr25 to feb26Transactions.json",
            "changedData/rmh/RMH advisors 25-26Master.json", "changedData/rmh/RMH advisors 25-26Transactions.json"),
}


def build_report(name: str) -> dict:
    """{sheet: [row, ...]} of the workbook generated for a dataset, every cell as text, empty rows left out."""
    paths = [BACKEND / p for p in DATASETS[name]]
    client = TestClient(app)
    with paths[0].open("rb") as m_old, paths[1].open("rb") as t_old, paths[2].open("rb") as m_new, paths[3].open("rb") as t_new:
        r = client.post("/api/compare", files={"master_old": m_old, "transactions_old": t_old,
                                               "master_new": m_new, "transactions_new": t_new})
    assert r.status_code == 200, r.text
    content = client.get(f"/api/download/{r.json()['report']['download_id']}").content
    wb = openpyxl.load_workbook(io.BytesIO(content))
    return {ws.title: [["" if v is None else str(v) for v in row] for row in ws.iter_rows(values_only=True)
                       if any(v is not None for v in row)]
            for ws in wb}


def _differences(expected: dict, actual: dict, limit: int = 15) -> list:
    out = []
    for sheet in sorted(set(expected) | set(actual)):
        if sheet not in actual:
            out.append(f"[{sheet}] sheet missing")
            continue
        if sheet not in expected:
            out.append(f"[{sheet}] new sheet")
            continue
        old, new = expected[sheet], actual[sheet]
        if old == new:
            continue
        removed = [r for r in old if r not in new]
        added = [r for r in new if r not in old]
        out.append(f"[{sheet}] {len(old)} rows approved, {len(new)} now; {len(removed)} changed/removed, {len(added)} changed/added"
                   + ("; same rows in a different order" if not removed and not added else ""))
        out += [f"   - {[v for v in r if v][:10]}" for r in removed[:limit]]
        out += [f"   + {[v for v in r if v][:10]}" for r in added[:limit]]
    return out


@pytest.mark.parametrize("name", list(DATASETS))
def test_report_matches_approved_baseline(name):
    if not all((BACKEND / p).exists() for p in DATASETS[name]):
        pytest.skip(f"input files for {name} not present")
    baseline = BASELINES / f"{name}.json"
    assert baseline.exists(), f"no baseline for {name}: run  python -m tests.test_regression_baselines --update"
    diffs = _differences(json.loads(baseline.read_text(encoding="utf-8")), build_report(name))
    assert not diffs, f"{name} report differs from the approved baseline:\n" + "\n".join(diffs)


if __name__ == "__main__":
    if "--update" not in sys.argv:
        sys.exit("usage: python -m tests.test_regression_baselines --update [dataset ...]")
    names = [a for a in sys.argv[1:] if a != "--update"] or list(DATASETS)
    BASELINES.mkdir(exist_ok=True)
    for n in names:
        (BASELINES / f"{n}.json").write_text(json.dumps(build_report(n), ensure_ascii=False, indent=0), encoding="utf-8")
        print(f"saved baseline: {n}")
