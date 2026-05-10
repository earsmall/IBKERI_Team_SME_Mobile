#!/usr/bin/env python3
"""Fetch dashboard data from URLs listed in the markdown file and build a static bundle."""

from __future__ import annotations

import json
import re
import ssl
import sys
import time
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Any, Callable
from urllib.error import HTTPError, URLError
from urllib.parse import parse_qsl, quote, urlencode, urlparse, urlunparse
from urllib.request import urlopen


ROOT = Path(__file__).resolve().parent
DATA_DIR = ROOT / "data"
URLS_PATH = ROOT / "Data Update OpenAPI URLs.md"
INDEX_PATH = ROOT / "index.html"
STATIC_BUNDLE_PATH = DATA_DIR / "dashboard-data.js"
SSL_CONTEXT = ssl._create_unverified_context()
FETCH_TIMEOUT_SECONDS = 120
FETCH_RETRY_COUNT = 3


@dataclass(frozen=True)
class UrlRecord:
    url: str
    params: dict[str, str]

    @property
    def tbl_id(self) -> str:
        return self.params.get("tblId", "")


@dataclass(frozen=True)
class Dataset:
    tab: str
    label: str
    filename: str
    key: str
    tbl_id: str
    match: Callable[[UrlRecord], bool] = lambda _record: True


@dataclass(frozen=True)
class SheetDataset:
    tab: str
    label: str
    filename: str
    key: str
    sheet_name: str


@dataclass
class Status:
    tab: str
    label: str
    latest: str
    changed: bool


SME_METRICS = {
    "DT_BR_A001": {"title": "기업수", "unit": "개", "color": "#2c7be5"},
    "DT_BR_B001": {"title": "종사자수", "unit": "명", "color": "#4a9bff"},
    "DT_BR_C001": {"title": "매출액", "unit": "백만원", "color": "#7fb8ff"},
}

EXPECTED_PERIODS = {
    "DT_BR_A001": "Y",
    "DT_BR_B001": "Y",
    "DT_BR_C001": "Y",
    "DT_303005_CI001": "M",
    "DT_1F02007": "Q",
    "DT_1KC2022": "Q",
    "DT_D10125": "M",
    "DT_512Y013": "M",
    "DT_512Y014": "M",
    "DT_501Y005": "Y",
    "DT_501Y006": "Y",
    "DT_501Y007": "Y",
    "DT_1TEC_P116": "Y",
    "DT_1TEC_P227": "Y",
    "DT_142N_F201": "Y",
}

DATASETS = (
    Dataset("실물경기", "경기동행종합지수", "business.json", "businessIndexRows", "DT_303005_CI001"),
    Dataset("실물경기", "제조업 생산지수", "business.json", "productionRows", "DT_1F02007"),
    Dataset("실물경기", "서비스업 생산지수", "business.json", "serviceProductionRows", "DT_1KC2022"),
    Dataset("실물경기", "중소제조업 평균가동률", "business.json", "operationRows", "DT_D10125"),
    Dataset("체감경기", "BSI 실적", "feeling.json", "actualRows", "DT_512Y013"),
    Dataset("체감경기", "BSI 전망", "feeling.json", "outlookRows", "DT_512Y014"),
    Dataset("경영지표", "성장성", "management.json", "growthRows", "DT_501Y005"),
    Dataset("경영지표", "수익성", "management.json", "profitRows", "DT_501Y006"),
    Dataset("경영지표", "안정성", "management.json", "stabilityRows", "DT_501Y007"),
    Dataset("수출", "중소기업 수출", "export.json", "rows", "DT_1TEC_P116"),
    Dataset("수출", "국가별 수출", "export.json", "countryRows", "DT_1TEC_P227"),
    Dataset(
        "창업",
        "창업기업 수",
        "startup.json",
        "rows",
        "DT_142N_F201",
        lambda record: {"A1", "A11"}.issubset(set(record.params.get("objL1", "").split())),
    ),
    Dataset(
        "창업",
        "업종별 창업",
        "startup.json",
        "rows",
        "DT_142N_F201",
        lambda record: {"B1", "C1"}.issubset(set(record.params.get("objL1", "").split())),
    ),
)

SHEET_DATASETS = (
    SheetDataset("대출", "대출잔액 및 순증", "loan.json", "loanRows", ""),
    SheetDataset("대출", "연체율", "loan.json", "delinquencyRows", "연체율"),
    SheetDataset("투자", "투자 총괄", "investment.json", "investmentRows", "투자"),
    SheetDataset("투자", "업력별 투자", "investment.json", "investmentStageRows", "업력별투자"),
    SheetDataset("투자", "업종별 투자", "investment.json", "investmentSectorRows", "업종별투자"),
    SheetDataset("투자", "출자자별 투자", "investment.json", "investmentSourceRows", "출자자별"),
)


def canonical_json(value: Any) -> str:
    return json.dumps(value, ensure_ascii=False, sort_keys=True, separators=(",", ":"))


def read_json(path: Path) -> dict[str, Any]:
    if not path.exists():
        return {}
    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
    except json.JSONDecodeError:
        return {}
    return payload if isinstance(payload, dict) else {}


def write_json_if_changed(path: Path, payload: dict[str, Any]) -> bool:
    existing = read_json(path)
    changed = canonical_json(existing) != canonical_json(payload)
    if changed:
        path.write_text(json.dumps(payload, ensure_ascii=False, separators=(",", ":")), encoding="utf-8")
    return changed


def read_markdown_text() -> str:
    for encoding in ("utf-8", "cp949", "euc-kr"):
        try:
            return URLS_PATH.read_text(encoding=encoding)
        except UnicodeDecodeError:
            continue
    return URLS_PATH.read_text(encoding="utf-8", errors="replace")


def parse_url_records() -> list[UrlRecord]:
    text = read_markdown_text()
    urls = re.findall(r"https://[^\s)>\]]+", text)
    records = []
    for url in urls:
        params = dict(parse_qsl(urlparse(url).query, keep_blank_values=True))
        if params.get("tblId"):
            records.append(UrlRecord(url=url, params=params))
    return records


def parse_google_sheet_doc_id() -> str:
    text = read_markdown_text()
    match = re.search(r"docs\.google\.com/spreadsheets/d/([^/\s]+)", text)
    if match:
        return match.group(1)
    match = re.search(r"([A-Za-z0-9_-]{30,})", text)
    if match:
        return match.group(1)
    raise ValueError("Google Sheet 문서 ID를 md 파일에서 찾지 못했습니다.")


def ensure_output_fields(url: str, required: tuple[str, ...] = ("TBL_ID", "NM", "PRD_SE", "PRD_DE", "LST_CHN_DE", "DT")) -> str:
    parsed = urlparse(url)
    params = dict(parse_qsl(parsed.query, keep_blank_values=True))
    fields = [field for field in re.split(r"[+\s]+", params.get("outputFields", "")) if field]
    for field in required:
        if field not in fields:
            fields.append(field)
    params["outputFields"] = "+".join(fields)
    return urlunparse(parsed._replace(query=urlencode(params, safe="+.")))


def fetch_url_bytes(url: str, label: str) -> bytes:
    last_error: Exception | None = None
    for attempt in range(1, FETCH_RETRY_COUNT + 1):
        try:
            with urlopen(url, timeout=FETCH_TIMEOUT_SECONDS, context=SSL_CONTEXT) as response:
                return response.read()
        except HTTPError:
            raise
        except Exception as error:
            last_error = error
            if attempt < FETCH_RETRY_COUNT:
                print(f"- {label}: 응답 지연/연결 실패, 재시도 중 ({attempt}/{FETCH_RETRY_COUNT})...")
                time.sleep(3 * attempt)
    raise URLError(f"{label} 호출 실패: {last_error}")


def fetch_json_url(url: str, label: str) -> Any:
    return json.loads(fetch_url_bytes(url, label).decode("utf-8", errors="replace"))


def fetch_rows(record: UrlRecord, label: str) -> list[dict[str, Any]]:
    payload = fetch_json_url(ensure_output_fields(record.url), label)
    if isinstance(payload, dict) and payload.get("err"):
        raise ValueError(payload.get("errMsg") or f"{label} API 오류")
    if not isinstance(payload, list):
        raise ValueError(f"{label} 응답이 JSON 배열이 아닙니다.")
    rows = [
        row for row in payload
        if isinstance(row, dict)
        and str(row.get("PRD_DE", "")).strip()
        and str(row.get("DT", "")).strip()
    ]
    if not rows:
        raise ValueError(f"{label} 응답에 PRD_DE/DT 값이 없습니다.")
    return rows


def extract_json_payload(text: str) -> dict[str, Any]:
    start = text.find("{")
    end = text.rfind("}")
    if start < 0 or end < start:
        raise ValueError("Google GViz JSON 응답을 찾지 못했습니다.")
    payload = json.loads(text[start:end + 1])
    if not isinstance(payload, dict):
        raise ValueError("Google GViz 응답 형식이 올바르지 않습니다.")
    return payload


def map_gviz_payload(payload: dict[str, Any]) -> list[dict[str, Any]]:
    cols = payload.get("table", {}).get("cols", [])
    rows = payload.get("table", {}).get("rows", [])
    mapped_rows = []
    for row in rows:
        cells = row.get("c", []) if isinstance(row, dict) else []
        record = {}
        for index, col in enumerate(cols):
            key = str(col.get("label") or col.get("id") or f"col_{index}").strip()
            cell = cells[index] if index < len(cells) else None
            record[key] = cell.get("v") if isinstance(cell, dict) else None
        if any(value not in (None, "") for value in record.values()):
            mapped_rows.append(record)
    if not mapped_rows:
        raise ValueError("Google Sheet에 읽을 행이 없습니다.")
    return mapped_rows


def fetch_google_sheet_rows(doc_id: str, sheet_name: str) -> list[dict[str, Any]]:
    tqx = quote("out:json")
    sheet_part = f"&sheet={quote(sheet_name)}" if sheet_name else ""
    url = f"https://docs.google.com/spreadsheets/d/{doc_id}/gviz/tq?tqx={tqx}{sheet_part}"
    label = f"Google Sheet {sheet_name or '기본 시트'}"
    text = fetch_url_bytes(url, label).decode("utf-8", errors="replace")
    return map_gviz_payload(extract_json_payload(text))


def find_record(records: list[UrlRecord], tbl_id: str, match: Callable[[UrlRecord], bool] = lambda _record: True) -> UrlRecord:
    for record in records:
        if record.tbl_id == tbl_id and match(record):
            return record
    raise ValueError(f"md 파일에서 {tbl_id} OPENAPI URL을 찾지 못했습니다.")


def latest_period(rows: list[dict[str, Any]]) -> str:
    values = [str(row.get("PRD_DE", "")).strip() for row in rows if str(row.get("PRD_DE", "")).strip()]
    return max(values) if values else ""


def last_changed(rows: list[dict[str, Any]]) -> str:
    values = [str(row.get("LST_CHN_DE", "")).strip() for row in rows if str(row.get("LST_CHN_DE", "")).strip()]
    return max(values) if values else ""


def row_identity(row: dict[str, Any]) -> tuple[str, ...]:
    return (
        str(row.get("TBL_ID") or row.get("TBL_NM") or "").strip(),
        str(row.get("ITM_ID") or row.get("ITM_NM") or "").strip(),
        str(row.get("C1") or row.get("C1_NM") or row.get("C1_OBJ_NM") or "").strip(),
        str(row.get("C2") or row.get("C2_NM") or "").strip(),
        str(row.get("C3") or row.get("C3_NM") or "").strip(),
        str(row.get("PRD_DE") or "").strip(),
    )


def merge_rows(existing_rows: list[dict[str, Any]], new_rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    merged: dict[tuple[str, ...], dict[str, Any]] = {}
    order: list[tuple[str, ...]] = []
    for row in existing_rows + new_rows:
        if not isinstance(row, dict):
            continue
        identity = row_identity(row)
        if identity not in merged:
            order.append(identity)
        merged[identity] = row
    return [merged[identity] for identity in order]


def parse_number(value: Any) -> int | float | None:
    try:
        number = float(str(value).replace(",", "").strip())
    except (TypeError, ValueError):
        return None
    return int(number) if number.is_integer() else number


def build_sme_years(rows: list[dict[str, Any]]) -> dict[str, dict[str, dict[str, int | float | None]]]:
    years: dict[str, dict[str, dict[str, int | float | None]]] = {}
    for row in rows:
        year = str(row.get("PRD_DE", "")).strip()
        value = parse_number(row.get("DT"))
        industry = str(row.get("C1_NM") or row.get("NM") or "전산업").strip() or "전산업"
        region = str(row.get("C2_NM") or "").strip()
        company_type = str(row.get("C3_NM") or row.get("ITM_NM") or "").strip()
        if not year or value is None:
            continue
        if region and region not in {"전국", "전체"}:
            continue
        if "중소" in company_type:
            bucket_key = "sme"
        elif "전체" in company_type or "전산업" in company_type or not company_type:
            bucket_key = "total"
        else:
            continue
        years.setdefault(year, {}).setdefault(industry, {"total": None, "sme": None})[bucket_key] = value
    return years


def update_sme_profile(records: list[UrlRecord]) -> tuple[bool, list[Status]]:
    path = DATA_DIR / "sme_profile.json"
    existing = read_json(path)
    datasets = existing.get("nextData", []) if isinstance(existing.get("nextData"), list) else []
    by_title = {str(item.get("title", "")): dict(item) for item in datasets if isinstance(item, dict)}
    statuses = []
    last_updates = []

    for tbl_id, metric in SME_METRICS.items():
        rows = fetch_rows(find_record(records, tbl_id), metric["title"])
        fetched_years = build_sme_years(rows)
        old_item = by_title.get(metric["title"], {"title": metric["title"], "years": {}})
        old_years = old_item.get("years", {}) if isinstance(old_item.get("years"), dict) else {}
        merged_years = {**old_years, **fetched_years}
        old_item.update({"unit": metric["unit"], "color": metric["color"], "years": dict(sorted(merged_years.items()))})
        by_title[metric["title"]] = old_item
        last_updates.append(last_changed(rows))
        statuses.append(Status("위상", metric["title"], latest_period(rows), canonical_json(old_years) != canonical_json(merged_years)))

    next_data = [by_title[metric["title"]] for metric in SME_METRICS.values()]
    payload = {
        "source": "kosis_snapshot",
        "generatedAt": date.today().isoformat(),
        "lastUpdated": max([value for value in last_updates if value], default=existing.get("lastUpdated", "")),
        "nextYears": sorted({year for item in next_data for year in item.get("years", {})}),
        "nextData": next_data,
    }
    return write_json_if_changed(path, payload), statuses


def split_business_index(rows: list[dict[str, Any]]) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    composite = []
    cycle = []
    for row in rows:
        label = " ".join(str(row.get(key, "")) for key in ("NM", "C1_NM", "C1_OBJ_NM", "OBJ_NM", "ITM_NM"))
        if "순환" in label:
            cycle.append(row)
        else:
            composite.append(row)
    return composite, cycle


def update_json_file(filename: str, payload_rows: dict[str, list[dict[str, Any]]], source: str) -> tuple[bool, dict[str, bool]]:
    path = DATA_DIR / filename
    existing = read_json(path)
    payload = dict(existing) if existing else {}
    latest_periods = dict(payload.get("latestPeriods", {})) if isinstance(payload.get("latestPeriods"), dict) else {}
    last_updated = dict(payload.get("lastUpdated", {})) if isinstance(payload.get("lastUpdated"), dict) else {}
    changed_by_key = {}

    for key, rows in payload_rows.items():
        old_rows = payload.get(key, []) if isinstance(payload.get(key), list) else []
        merged = merge_rows(old_rows, rows)
        changed_by_key[key] = canonical_json(old_rows) != canonical_json(merged)
        payload[key] = merged
        latest_periods[key] = latest_period(merged)
        last_updated[key] = max(last_updated.get(key, ""), last_changed(rows))

    payload["source"] = source
    payload["generatedAt"] = date.today().isoformat()
    payload["latestPeriods"] = latest_periods
    payload["lastUpdated"] = last_updated
    return write_json_if_changed(path, payload), changed_by_key


def update_regular_datasets(records: list[UrlRecord]) -> tuple[bool, list[Status]]:
    rows_by_file: dict[str, dict[str, list[dict[str, Any]]]] = {}
    specs: list[tuple[str, str, str, str]] = []

    for dataset in DATASETS:
        rows = fetch_rows(find_record(records, dataset.tbl_id, dataset.match), dataset.label)
        if dataset.key == "businessIndexRows":
            composite, cycle = split_business_index(rows)
            rows_by_file.setdefault(dataset.filename, {}).setdefault("businessCompositeRows", []).extend(composite)
            rows_by_file.setdefault(dataset.filename, {}).setdefault("businessCycleRows", []).extend(cycle)
            specs.append((dataset.filename, "businessCompositeRows", dataset.tab, "경기동행종합지수"))
            specs.append((dataset.filename, "businessCycleRows", dataset.tab, "경기동행지수 순환변동치"))
        else:
            rows_by_file.setdefault(dataset.filename, {}).setdefault(dataset.key, []).extend(rows)
            specs.append((dataset.filename, dataset.key, dataset.tab, dataset.label))

    any_changed = False
    changed_by_file_key = {}
    for filename, payload_rows in rows_by_file.items():
        changed, changed_by_key = update_json_file(filename, payload_rows, "kosis_snapshot")
        any_changed = any_changed or changed
        changed_by_file_key.update({(filename, key): value for key, value in changed_by_key.items()})

    statuses = [
        Status(tab, label, latest_period(rows_by_file.get(filename, {}).get(key, [])), changed_by_file_key.get((filename, key), False))
        for filename, key, tab, label in specs
    ]
    return any_changed, statuses


def parse_sheet_period(value: Any) -> tuple[str, str]:
    text = str(value or "").strip()
    if not text:
        return "", ""
    match = re.match(r"^Date\((\d+),(\d+),(\d+)\)$", text)
    if match:
        year = int(match.group(1))
        month = int(match.group(2)) + 1
        day = int(match.group(3))
        return f"{year:04d}{month:02d}{day:02d}", f"{year:04d}-{month:02d}-{day:02d}"
    compact = re.sub(r"\D", "", text)
    if len(compact) >= 4:
        return compact, text
    return text, text


def sheet_period_value(row: dict[str, Any]) -> Any:
    for key in ("시점", "날짜", "", "A"):
        value = row.get(key)
        if value not in (None, ""):
            return value
    return None


def latest_sheet_period(rows: list[dict[str, Any]]) -> str:
    values = [parse_sheet_period(sheet_period_value(row)) for row in rows]
    values = [value for value in values if value[0]]
    return max(values, key=lambda value: value[0])[1] if values else ""


def sort_sheet_rows(rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    return sorted(rows, key=lambda row: parse_sheet_period(sheet_period_value(row))[0] or canonical_json(row))


def update_google_sheet_datasets() -> tuple[bool, list[Status]]:
    doc_id = parse_google_sheet_doc_id()
    rows_by_file: dict[str, dict[str, list[dict[str, Any]]]] = {}
    statuses = []
    latest_by_file_key = {}

    for dataset in SHEET_DATASETS:
        rows = sort_sheet_rows(fetch_google_sheet_rows(doc_id, dataset.sheet_name))
        rows_by_file.setdefault(dataset.filename, {})[dataset.key] = rows
        latest_by_file_key[(dataset.filename, dataset.key)] = latest_sheet_period(rows)

    any_changed = False
    changed_by_file_key = {}
    for filename, payload_rows in rows_by_file.items():
        path = DATA_DIR / filename
        existing = read_json(path)
        payload = dict(existing) if existing else {}
        latest_periods = dict(payload.get("latestPeriods", {})) if isinstance(payload.get("latestPeriods"), dict) else {}
        for key, rows in payload_rows.items():
            old_rows = payload.get(key, []) if isinstance(payload.get(key), list) else []
            changed_by_file_key[(filename, key)] = canonical_json(old_rows) != canonical_json(rows)
            payload[key] = rows
            latest_periods[key] = latest_by_file_key.get((filename, key), "")
        payload["source"] = "google_sheet_snapshot"
        payload["generatedAt"] = date.today().isoformat()
        payload["latestPeriods"] = latest_periods
        changed = write_json_if_changed(path, payload)
        any_changed = any_changed or changed

    for dataset in SHEET_DATASETS:
        statuses.append(
            Status(
                dataset.tab,
                dataset.label,
                latest_by_file_key.get((dataset.filename, dataset.key), ""),
                changed_by_file_key.get((dataset.filename, dataset.key), False),
            )
        )
    return any_changed, statuses


def write_static_bundle() -> None:
    bundle = {path.name: read_json(path) for path in sorted(DATA_DIR.glob("*.json"))}
    STATIC_BUNDLE_PATH.write_text(
        "window.__DASHBOARD_STATIC_JSON__=" + json.dumps(bundle, ensure_ascii=False, separators=(",", ":")) + ";",
        encoding="utf-8",
    )


def bump_index_bundle_version() -> None:
    if not INDEX_PATH.exists():
        return
    stamp = datetime.now().strftime("%Y%m%d%H%M%S")
    text = INDEX_PATH.read_text(encoding="utf-8")
    if "./data/dashboard-data.js" not in text:
        text = text.replace('<script src="./script.js', f'<script src="./data/dashboard-data.js?v={stamp}"></script>\n    <script src="./script.js')
    else:
        text = re.sub(r'(\./data/dashboard-data\.js(?:\?v=)?)[^"]*', f"./data/dashboard-data.js?v={stamp}", text, count=1)
    INDEX_PATH.write_text(text, encoding="utf-8")


def index_loads_static_bundle() -> bool:
    return INDEX_PATH.exists() and "./data/dashboard-data.js" in INDEX_PATH.read_text(encoding="utf-8")


def print_statuses(statuses: list[Status]) -> None:
    updated_count = sum(1 for status in statuses if status.changed)
    unchanged_count = len(statuses) - updated_count
    print("탭별 데이터 확인 결과:")
    print(f"- 전체 항목: {len(statuses)}개")
    print(f"- 신규 데이터 반영: {updated_count}개")
    print(f"- 업데이트 불필요: {unchanged_count}개")
    current_tab = ""
    for status in statuses:
        if status.tab != current_tab:
            current_tab = status.tab
            print(f"\n[{current_tab}]")
        state = "신규 데이터 반영" if status.changed else "업데이트 불필요"
        print(f"- {status.label}: {state} / 최신 시점: {status.latest or '-'}")


def main() -> int:
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    try:
        records = parse_url_records()
        if not records:
            raise ValueError("md 파일에서 OPENAPI URL을 찾지 못했습니다.")
        sme_changed, sme_statuses = update_sme_profile(records)
        regular_changed, regular_statuses = update_regular_datasets(records)
        sheet_changed, sheet_statuses = update_google_sheet_datasets()
        statuses = sme_statuses + regular_statuses + sheet_statuses
        print_statuses(statuses)
        changed = sme_changed or regular_changed or sheet_changed
        needs_bundle = changed or not STATIC_BUNDLE_PATH.exists()
        needs_index_update = changed or not index_loads_static_bundle()
        if needs_bundle:
            write_static_bundle()
        if needs_index_update:
            bump_index_bundle_version()
        print("\n반영 결과:")
        print("- data/dashboard-data.js", "생성/갱신" if needs_bundle else "변경 없음")
        print("- index.html", "반영 완료" if needs_index_update else "변경 없음")
        print("- 종합 판정:", "신규 데이터가 있어 반영했습니다." if changed else "모든 항목이 기존 최신 데이터와 동일합니다.")
        return 0
    except (HTTPError, URLError) as error:
        print(f"API 호출 실패: {error}", file=sys.stderr)
        print("네트워크 응답 시간이 초과되었거나, 회사/보안망에서 KOSIS 또는 Google Sheet 접속이 차단되었을 수 있습니다.", file=sys.stderr)
        print("인터넷/VPN 상태를 확인한 뒤 다시 실행해주세요. 이미 받아둔 기존 data 파일은 삭제되지 않습니다.", file=sys.stderr)
        return 1
    except Exception as error:
        print(f"실행 실패: {error}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
