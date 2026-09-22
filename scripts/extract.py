"""Selected aggregate Decision Maker Panel series from the official BoE workbook."""

from __future__ import annotations

import hashlib
import io
import logging
import math
import re
from dataclasses import dataclass
from datetime import UTC, date, datetime, time, timedelta
from typing import Any
from urllib.parse import urljoin

import httpx
import openpyxl

from scripts.config import (
    MAX_DOWNLOAD_BYTES,
    MAX_STALE_MONTHS,
    MIN_HISTORY_YEARS,
    REQUEST_TIMEOUT,
    USER_AGENT,
)
from scripts.snapshots import Snapshot, build_snapshot
from scripts.time_series import Observation

logger = logging.getLogger(__name__)


# -- series_id contract (GUIDELINES.md 4) ---------------------------------
# series_id is uppercase, underscore-separated and ordered coarse -> fine. The
# pair below is the canonical public surface: parse splits an id into its
# components, build rejoins them, and build(*parse(sid)) == sid for every id
# this collector emits. Only economic identity is encoded -- never a delivery
# provider or any other detail of how the value reached us.


def parse_series_id(series_id: str) -> tuple[str, ...]:
    """Split a series_id into its underscore-delimited components.

    Raises ValueError on anything this collector would not have produced:
    lowercase, empty components, or an id with no structure at all.
    """
    if not series_id or series_id != series_id.upper():
        raise ValueError(f"series_id must be uppercase: {series_id!r}")
    components = tuple(series_id.split("_"))
    if any(not component for component in components):
        raise ValueError(f"series_id has an empty component: {series_id!r}")
    return components


def build_series_id(*components: str) -> str:
    """Rejoin the tuple parse_series_id returned into the original id."""
    if not components:
        raise ValueError("series_id needs at least one component")
    if any(not component or component != component.upper() for component in components):
        raise ValueError(f"invalid series_id components: {components!r}")
    return "_".join(components)


# -- 5.1 usable-series filtering ------------------------------------------


@dataclass(frozen=True)
class UsabilityReport:
    """What the filter removed, for logging and for tests to assert on."""

    kept: tuple[str, ...]
    stale: tuple[str, ...]
    short_history: tuple[str, ...]
    empty: tuple[str, ...]

    @property
    def dropped(self) -> tuple[str, ...]:
        return tuple(sorted(set(self.stale) | set(self.short_history) | set(self.empty)))


def _months_between(earlier: date, later: date) -> int:
    """Whole months from ``earlier`` to ``later``, day-of-month aware."""
    months = (later.year - earlier.year) * 12 + (later.month - earlier.month)
    if later.day < earlier.day:
        months -= 1
    return months


def _is_valid(value: Any) -> bool:
    """A real observation: present, numeric and finite."""
    if value is None:
        return False
    try:
        numeric = float(value)
    except (TypeError, ValueError):
        return False
    return math.isfinite(numeric)


def assess_series(
    reference_dates: list[date],
    today: date,
    max_stale_months: int = MAX_STALE_MONTHS,
    min_history_years: float = MIN_HISTORY_YEARS,
) -> str:
    """Classify one series from the reference dates of its valid observations.

    Returns ``"keep"``, ``"empty"``, ``"stale"`` or ``"short_history"``.
    Recency is judged at the period end and over non-null values only: a source
    that keeps listing a discontinued series with empty recent cells must not
    look live because of those blanks.
    """
    if not reference_dates:
        return "empty"
    first, last = min(reference_dates), max(reference_dates)
    if _months_between(last, today) > max_stale_months:
        return "stale"
    if _months_between(first, last) < round(min_history_years * 12):
        return "short_history"
    return "keep"


def filter_usable_series(
    observations: list[Any],
    catalog: dict[str, dict[str, Any]],
    today: date,
    max_stale_months: int = MAX_STALE_MONTHS,
    min_history_years: float = MIN_HISTORY_YEARS,
) -> tuple[list[Any], dict[str, dict[str, Any]], UsabilityReport]:
    """Drop obsolete and history-less series before anything is persisted.

    Runs after parsing and before the time_series / metadata upsert, so the
    standardized tables never carry a dead or stub series, and prunes the
    catalog alongside the observations so metadata can never describe a series
    the database does not hold (GUIDELINES.md 5.1).
    """
    valid_dates: dict[str, list[date]] = {}
    for observation in observations:
        if _is_valid(observation.value):
            valid_dates.setdefault(observation.series_id, []).append(observation.reference_date)

    verdicts: dict[str, str] = {}
    for series_id in set(catalog) | {o.series_id for o in observations}:
        verdicts[series_id] = assess_series(
            valid_dates.get(series_id, []), today, max_stale_months, min_history_years
        )

    keep = {series_id for series_id, verdict in verdicts.items() if verdict == "keep"}
    report = UsabilityReport(
        kept=tuple(sorted(keep)),
        stale=tuple(sorted(s for s, v in verdicts.items() if v == "stale")),
        short_history=tuple(sorted(s for s, v in verdicts.items() if v == "short_history")),
        empty=tuple(sorted(s for s, v in verdicts.items() if v == "empty")),
    )

    if report.dropped:
        logger.info(
            "Usable-series filter: kept %d, dropped %d "
            "(stale=%d short_history=%d empty=%d; max_stale_months=%d min_history_years=%s)",
            len(report.kept),
            len(report.dropped),
            len(report.stale),
            len(report.short_history),
            len(report.empty),
            max_stale_months,
            min_history_years,
        )
        for series_id in report.stale:
            logger.info(
                "Dropped %s: last valid observation older than %d months",
                series_id,
                max_stale_months,
            )
        for series_id in report.short_history:
            logger.info(
                "Dropped %s: valid history shorter than %s years", series_id, min_history_years
            )
        for series_id in report.empty:
            logger.info("Dropped %s: no valid observations", series_id)
    else:
        logger.info("Usable-series filter: all %d series usable", len(report.kept))

    kept_observations = [o for o in observations if o.series_id in keep]
    kept_catalog = {sid: fields for sid, fields in catalog.items() if sid in keep}
    return kept_observations, kept_catalog, report


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
