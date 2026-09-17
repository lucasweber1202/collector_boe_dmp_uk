"""Selected aggregate Decision Maker Panel series from the official BoE workbook."""

from __future__ import annotations

import hashlib
import io
import re
from dataclasses import dataclass
from datetime import UTC, date, datetime, time, timedelta
from typing import Any
from urllib.parse import urljoin

import httpx
import openpyxl

from scripts.config import MAX_DOWNLOAD_BYTES, REQUEST_TIMEOUT, USER_AGENT
from scripts.snapshots import Snapshot, build_snapshot
from scripts.time_series import Observation

RELEASE_ROOT = "https://www.bankofengland.co.uk/decision-maker-panel"
SELECTION = {
    "Price growth": {2: "REALISED_PRICE_3M", 13: "EXPECTED_PRICE_3M"},
    "Wage growth": {2: "REALISED_WAGE_3M", 5: "EXPECTED_WAGE_3M"},
    "Employment growth": {2: "REALISED_EMPLOYMENT_3M", 13: "EXPECTED_EMPLOYMENT_3M"},
    "Unit cost growth": {2: "REALISED_UNIT_COST_3M", 5: "EXPECTED_UNIT_COST_3M"},
    "CPI expectations": {2: "CURRENT_CPI_3M", 4: "CPI_1Y_3M", 6: "CPI_3Y_3M"},
}


@dataclass(frozen=True)
class ExtractedData:
    observations: list[Observation]
    snapshots: list[Snapshot]
    catalog: dict[str, dict[str, Any]]
    releases: list[datetime]
    availability_by_key: dict[tuple[str, date], tuple[datetime, str, date | None]]
    min_lag_days: int = 0
    max_lag_days: int = 0
    inferred_lag_days: int | None = None


def _month(value: object) -> date | None:
    if isinstance(value, datetime):
        return value.date().replace(day=1)
    if isinstance(value, str):
        try:
            return (
                datetime.strptime(value.strip(), "%b-%y").replace(tzinfo=UTC).date().replace(day=1)
            )
        except ValueError:
            return None
    return None


def parse_xlsx(
    body: bytes, snapshot_id: str, source_url: str, released: datetime, collected: datetime
) -> tuple[
    list[Observation],
    dict[str, dict[str, Any]],
    dict[tuple[str, date], tuple[datetime, str, date | None]],
]:
    book = openpyxl.load_workbook(io.BytesIO(body), data_only=True, read_only=True)
    observations = []
    catalog = {}
    availability = {}
    keys = set()
    for sheet_name, columns in SELECTION.items():
        if sheet_name not in book.sheetnames:
            raise ValueError(f"DMP workbook missing {sheet_name}")
        sheet = book[sheet_name]
        for column, label in columns.items():
            series_id = f"BOE_DMP_{label}"
            count = 0
            for row in range(1, sheet.max_row + 1):
                reference = _month(sheet.cell(row, 1).value)
                raw = sheet.cell(row, column).value
                if reference is None or not isinstance(raw, (int, float)):
                    continue
                key = (series_id, reference)
                if key in keys:
                    raise ValueError(f"Duplicate DMP key {key}")
                keys.add(key)
                count += 1
                observations.append(Observation(series_id, reference, float(raw), snapshot_id))
            if count < 5:
                raise ValueError(f"DMP series {series_id} unexpectedly short")
            catalog[series_id] = {
                "source_id": "boe_dmp_monthly",
                "name": label.replace("_", " ").title(),
                "description": f"Raw weighted DMP aggregate from sheet {sheet_name}; no collector-side transformation.",
                "frequency": "monthly",
                "unit": "percent",
                "eco_group": "surveys",
                "source_url": source_url,
                "last_publish_date": released.date(),
            }
    latest_by_series = {
        series_id: max(o.reference_date for o in observations if o.series_id == series_id)
        for series_id in catalog
    }
    for observation in observations:
        is_latest = observation.reference_date == latest_by_series[observation.series_id]
        availability[(observation.series_id, observation.reference_date)] = (
            released if is_latest else collected,
            "official_date" if is_latest else "first_seen",
            released.date() if is_latest else None,
        )
    if len(catalog) != 11 or len(observations) < 500:
        raise ValueError("DMP selected dataset unexpectedly short")
    return observations, catalog, availability


def collect() -> ExtractedData:
    fetched = datetime.now(UTC)
    with httpx.Client(
        timeout=REQUEST_TIMEOUT, headers={"User-Agent": USER_AGENT}, follow_redirects=True
    ) as client:
        page = None
        page_url = ""
        candidate = fetched.date().replace(day=1)
        for _ in range(4):
            page_url = f"{RELEASE_ROOT}/{candidate.year}/{candidate.strftime('%B').lower()}-{candidate.year}"
            response_page = client.get(page_url)
            if (
                response_page.status_code == 200
                and "monthly-dmp-data" in response_page.text.lower()
            ):
                page = response_page
                break
            candidate = (candidate.replace(day=1) - timedelta(days=1)).replace(day=1)
        if page is None:
            raise ValueError(
                "No published DMP monthly workbook discovered in the latest four months"
            )
        date_match = re.search(r"Published on\s+(\d{1,2} \w+ \d{4})", page.text, re.IGNORECASE)
        links = re.findall(
            r'href=["\']([^"\']+monthly-dmp-data[^"\']+\.xlsx)["\']', page.text, re.IGNORECASE
        )
        if not date_match or len(links) != 1:
            raise ValueError("DMP release page structure drifted")
        released = datetime.combine(
            datetime.strptime(date_match.group(1), "%d %B %Y").replace(tzinfo=UTC).date(),
            time(7),
            tzinfo=UTC,
        )
        url = urljoin(page_url, links[0])
        response = client.get(url)
        response.raise_for_status()
    body = response.content
    if not body or len(body) > MAX_DOWNLOAD_BYTES:
        raise ValueError(f"Invalid DMP artifact size {len(body)}")
    digest = hashlib.sha256(body).hexdigest()
    observations, catalog, availability = parse_xlsx(body, digest, url, released, fetched)
    snapshot = build_snapshot(
        "boe_dmp_monthly",
        url,
        "boe_dmp_monthly.xlsx",
        body,
        digest,
        response.headers.get("etag"),
        response.headers.get("last-modified"),
        fetched,
        released.date(),
    )
    return ExtractedData(observations, [snapshot], catalog, [released], availability)
