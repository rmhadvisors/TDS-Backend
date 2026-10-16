"""Download history: every generated Excel report is kept so any user of the app can download it again later.

Where the reports go:
  * S3, when TDS_S3_BUCKET is set (optional: TDS_S3_PREFIX, default "reports/"; AWS_REGION / AWS credentials the
    usual boto3 way - an IAM role on EC2/ECS, or AWS_ACCESS_KEY_ID / AWS_SECRET_ACCESS_KEY env vars).
  * Otherwise a local folder, backend/data/reports (override with TDS_REPORTS_DIR), so the app works without AWS.

There is no database: the history is the list of stored files. Everything the History table shows is in the file
name itself, e.g.
    2026-10-15T11-42-00Z__MK%20Shah__Apr-Sep%202026__1b9d6bcd-bbfd-4b2d-9b5d-ab8dfbbd4bed.xlsx
     created (UTC)        party       period          report id (= download_id)
Party and period are percent-encoded (and "_" too) so any name splits back exactly on "__".
"""
import os
import re
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path
from typing import List, Optional
from urllib.parse import quote, unquote

_DEFAULT_DIR = Path(__file__).resolve().parents[2] / "data" / "reports"
_STAMP = "%Y-%m-%dT%H-%M-%SZ"
_ID_RE = re.compile(r"^[0-9a-fA-F-]{36}$")
PRESIGNED_URL_SECONDS = 300


@dataclass
class StoredReport:
    report_id: str
    created_at: str  # ISO 8601, UTC
    party_name: Optional[str]
    period: Optional[str]
    size: int
    filename: str
    key: str

    def as_dict(self) -> dict:
        return {
            "report_id": self.report_id,
            "created_at": self.created_at,
            "party_name": self.party_name,
            "period": self.period,
            "size": self.size,
            "filename": self.filename,
        }


def _encode(value: Optional[str]) -> str:
    return quote((value or "").strip(), safe=" -.,()&").replace("_", "%5F") or "-"


def _decode(value: str) -> Optional[str]:
    return None if value == "-" else unquote(value)


def download_filename(party_name: Optional[str], period: Optional[str], created: datetime) -> str:
    """A readable, ASCII-only name, so downloads from the history do not all overwrite each other."""
    named = [re.sub(r"[^A-Za-z0-9.-]+", "-", p or "").strip("-") for p in (party_name, period)]
    return "_".join(["TDS_Report", *(p for p in named if p), created.strftime("%Y-%m-%d_%H%M")]) + ".xlsx"


def make_object_name(report_id: str, party_name: Optional[str], period: Optional[str], created: datetime) -> str:
    return f"{created.strftime(_STAMP)}__{_encode(party_name)}__{_encode(period)}__{report_id}.xlsx"


def parse_object_name(name: str, size: int, key: str) -> Optional[StoredReport]:
    if not name.endswith(".xlsx"):
        return None
    pieces = name[: -len(".xlsx")].split("__")
    if len(pieces) != 4 or not _ID_RE.match(pieces[3]):
        return None
    try:
        created = datetime.strptime(pieces[0], _STAMP).replace(tzinfo=timezone.utc)
    except ValueError:
        return None
    party, period = _decode(pieces[1]), _decode(pieces[2])
    return StoredReport(
        report_id=pieces[3],
        created_at=created.isoformat().replace("+00:00", "Z"),
        party_name=party,
        period=period,
        size=size,
        filename=download_filename(party, period, created),
        key=key,
    )


class LocalReportStore:
    kind = "local"

    def __init__(self, directory: Path):
        self.directory = directory

    def save(self, name: str, data: bytes) -> None:
        self.directory.mkdir(parents=True, exist_ok=True)
        (self.directory / name).write_bytes(data)

    def list(self) -> List[StoredReport]:
        if not self.directory.is_dir():
            return []
        found = (parse_object_name(p.name, p.stat().st_size, p.name) for p in self.directory.iterdir() if p.is_file())
        return [r for r in found if r]

    def read(self, report: StoredReport) -> bytes:
        return (self.directory / report.key).read_bytes()

    def delete(self, report: StoredReport) -> None:
        (self.directory / report.key).unlink(missing_ok=True)


class S3ReportStore:
    kind = "s3"

    def __init__(self, bucket: str, prefix: str):
        import boto3  # only needed when S3 is configured

        self.bucket = bucket
        self.prefix = prefix
        self.client = boto3.client("s3")

    def save(self, name: str, data: bytes) -> None:
        self.client.put_object(
            Bucket=self.bucket,
            Key=self.prefix + name,
            Body=data,
            ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            ServerSideEncryption="AES256",
        )

    def list(self) -> List[StoredReport]:
        reports = []
        for page in self.client.get_paginator("list_objects_v2").paginate(Bucket=self.bucket, Prefix=self.prefix):
            for obj in page.get("Contents", []):
                report = parse_object_name(obj["Key"][len(self.prefix):], obj["Size"], obj["Key"])
                if report:
                    reports.append(report)
        return reports

    def delete(self, report: StoredReport) -> None:
        self.client.delete_object(Bucket=self.bucket, Key=report.key)

    def presigned_url(self, report: StoredReport) -> str:
        return self.client.generate_presigned_url(
            "get_object",
            Params={
                "Bucket": self.bucket,
                "Key": report.key,
                "ResponseContentDisposition": f'attachment; filename="{report.filename}"',
            },
            ExpiresIn=PRESIGNED_URL_SECONDS,
        )


_STORE = {"config": None, "store": None}


def get_store():
    """The configured store; rebuilt when the env vars change (tests point it at a temp folder)."""
    bucket = os.environ.get("TDS_S3_BUCKET", "").strip()
    config = (bucket, os.environ.get("TDS_S3_PREFIX", "reports/"), os.environ.get("TDS_REPORTS_DIR", ""))
    if _STORE["config"] != config:
        if bucket:
            prefix = config[1].strip("/") + "/" if config[1].strip("/") else ""
            store = S3ReportStore(bucket, prefix)
        else:
            store = LocalReportStore(Path(config[2]) if config[2] else _DEFAULT_DIR)
        _STORE.update(config=config, store=store)
    return _STORE["store"]


def list_reports(search: str = "", limit: int = 200) -> List[StoredReport]:
    """Newest first; `search` matches party name or period, case-insensitive."""
    needle = search.strip().lower()
    reports = [
        r for r in get_store().list()
        if not needle or needle in (r.party_name or "").lower() or needle in (r.period or "").lower()
    ]
    reports.sort(key=lambda r: r.created_at, reverse=True)
    return reports[:limit]


def find_report(report_id: str) -> Optional[StoredReport]:
    if not _ID_RE.match(report_id):
        return None
    return next((r for r in get_store().list() if r.report_id == report_id), None)
