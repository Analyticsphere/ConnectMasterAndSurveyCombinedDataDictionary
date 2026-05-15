#!/usr/bin/env python3
from __future__ import annotations

import argparse
import datetime as dt
import difflib
import hashlib
import json
import os
import re
import sys
import urllib.request
import zipfile
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET


ROOT = Path(__file__).resolve().parent
DATA_DIR = ROOT / "data"
LIBRARY_DIR = ROOT / "library"

WORKBOOK_URL = (
    "https://raw.githubusercontent.com/"
    "Analyticsphere/ConnectMasterAndSurveyCombinedDataDictionary/"
    "main/MasterSurveyComb_20220317.xlsx"
)
WORKBOOK_PAGE_URL = (
    "https://github.com/Analyticsphere/"
    "ConnectMasterAndSurveyCombinedDataDictionary/blob/main/"
    "MasterSurveyComb_20220317.xlsx"
)

WORKBOOK_PATH = DATA_DIR / "MasterSurveyComb_latest.xlsx"
INDEX_PATH = DATA_DIR / "connect_dictionary_index.json"
METADATA_PATH = DATA_DIR / "workbook_metadata.json"
LIBRARY_PATH = LIBRARY_DIR / "connect_reference_library.json"
INDEX_SCHEMA_VERSION = 3

STATUS_FIELD = "Deprecated, New, or Revised"
STATUS_DATE_FIELD = "Date Deprecated, New, or Revised Variable Pushed to Prod"

NS = {
    "m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
}

MASTER_HEADER_OVERRIDES = {
    2: "conceptId_Primary Source",
    4: "conceptId_Secondary Source",
    6: "conceptId_Source Question",
    13: "conceptId_Current Question Text",
    22: "conceptId_Current Format/Value",
}

CONTEXT_FIELDS = {
    "Primary Source",
    "Secondary Source",
    "conceptId_Primary Source",
    "conceptId_Secondary Source",
    "conceptId_Source Question",
    "Current Source Question",
    "V2 Source Question",
    "V1 Source Question",
    "GridID/Source Question Name",
    "conceptId_Current Question Text",
    "Current Question Text",
    "V2 Question Text",
    "V1 Question Text",
    "Variable Label",
    "Variable Name",
    "Variable Type",
    "Variable Length",
    "Required",
    "PII",
    "Notes",
    "Derivation Notes",
    "Dictionary",
    "State Attribute",
    "Default Variable",
    "GCP Document/Table",
    "Question Type",
}

IMPORTANT_FIELDS = [
    "Current Question Text",
    "Question Text",
    "Variable Label",
    "Variable Name",
    "Current Format/Value",
    "Format/Value",
    "GCP Document/Table",
    "Notes",
    "Primary Source",
    "Secondary Source",
]

STOPWORDS = {
    "a",
    "an",
    "and",
    "are",
    "for",
    "in",
    "is",
    "of",
    "or",
    "the",
    "to",
}


def now_iso() -> str:
    return dt.datetime.now().astimezone().isoformat(timespec="seconds")


def ensure_dirs() -> None:
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    LIBRARY_DIR.mkdir(parents=True, exist_ok=True)


def read_json(path: Path, default: Any) -> Any:
    if not path.exists():
        return default
    with path.open("r", encoding="utf-8") as handle:
        return json.load(handle)


def write_json(path: Path, data: Any) -> None:
    ensure_dirs()
    tmp = path.with_suffix(path.suffix + ".tmp")
    with tmp.open("w", encoding="utf-8") as handle:
        json.dump(data, handle, indent=2, ensure_ascii=False, sort_keys=True)
        handle.write("\n")
    os.replace(tmp, path)


def file_sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def download_workbook() -> dict[str, Any]:
    ensure_dirs()
    tmp = WORKBOOK_PATH.with_suffix(".xlsx.tmp")
    request = urllib.request.Request(
        WORKBOOK_URL,
        headers={"User-Agent": "connect-dictionary-search/1.0"},
    )
    with urllib.request.urlopen(request, timeout=60) as response:
        payload = response.read()
    with tmp.open("wb") as handle:
        handle.write(payload)
    os.replace(tmp, WORKBOOK_PATH)

    metadata = {
        "downloaded_at": now_iso(),
        "source_url": WORKBOOK_URL,
        "source_page": WORKBOOK_PAGE_URL,
        "workbook_path": str(WORKBOOK_PATH),
        "sha256": hashlib.sha256(payload).hexdigest(),
        "size_bytes": len(payload),
    }
    write_json(METADATA_PATH, metadata)
    return metadata


def refresh_or_fail(allow_stale: bool = False) -> tuple[dict[str, Any], bool]:
    previous = read_json(METADATA_PATH, {})
    try:
        metadata = download_workbook()
    except Exception as exc:
        if allow_stale and WORKBOOK_PATH.exists():
            metadata = previous or {
                "downloaded_at": "unknown",
                "source_url": WORKBOOK_URL,
                "sha256": file_sha256(WORKBOOK_PATH),
                "warning": f"Refresh failed; using stale workbook: {exc}",
            }
            metadata["warning"] = f"Refresh failed; using stale workbook: {exc}"
            return metadata, False
        raise RuntimeError(
            "Could not refresh the workbook from GitHub. "
            "Use --allow-stale only if you intentionally want the local cache."
        ) from exc

    changed = metadata.get("sha256") != previous.get("sha256")
    return metadata, changed


def column_index(cell_ref: str) -> int:
    letters = "".join(ch for ch in cell_ref if ch.isalpha())
    number = 0
    for letter in letters.upper():
        number = number * 26 + ord(letter) - 64
    return number - 1


def column_letter(index: int) -> str:
    index += 1
    letters = ""
    while index:
        index, remainder = divmod(index - 1, 26)
        letters = chr(65 + remainder) + letters
    return letters


def cell_value(cell: ET.Element, shared_strings: list[str]) -> str:
    cell_type = cell.attrib.get("t")
    value = cell.find("m:v", NS)
    if cell_type == "s" and value is not None and value.text is not None:
        return shared_strings[int(value.text)]
    if cell_type == "inlineStr":
        return "".join(node.text or "" for node in cell.findall(".//m:t", NS))
    if value is None or value.text is None:
        return ""
    return value.text


def load_shared_strings(zipped: zipfile.ZipFile) -> list[str]:
    if "xl/sharedStrings.xml" not in zipped.namelist():
        return []
    root = ET.fromstring(zipped.read("xl/sharedStrings.xml"))
    return [
        "".join(node.text or "" for node in item.findall(".//m:t", NS))
        for item in root.findall("m:si", NS)
    ]


def workbook_sheets(zipped: zipfile.ZipFile) -> list[tuple[str, str]]:
    workbook = ET.fromstring(zipped.read("xl/workbook.xml"))
    rels = ET.fromstring(zipped.read("xl/_rels/workbook.xml.rels"))
    targets = {
        rel.attrib["Id"]: rel.attrib["Target"]
        for rel in rels.findall("rel:Relationship", NS)
    }

    sheets = []
    for sheet in workbook.findall("m:sheets/m:sheet", NS):
        rel_id = sheet.attrib[f"{{{NS['r']}}}id"]
        target = targets[rel_id]
        path = target.lstrip("/") if target.startswith("/") else f"xl/{target}"
        sheets.append((sheet.attrib["name"], path))
    return sheets


def parse_rows(
    zipped: zipfile.ZipFile,
    sheet_path: str,
    shared_strings: list[str],
) -> list[tuple[int, list[str]]]:
    root = ET.fromstring(zipped.read(sheet_path))
    parsed = []
    for row in root.findall("m:sheetData/m:row", NS):
        cells: dict[int, str] = {}
        max_index = -1
        for cell in row.findall("m:c", NS):
            ref = cell.attrib.get("r", "")
            if not ref:
                continue
            index = column_index(ref)
            max_index = max(max_index, index)
            cells[index] = cell_value(cell, shared_strings).strip()
        if max_index < 0:
            continue
        values = [""] * (max_index + 1)
        for index, value in cells.items():
            values[index] = value
        parsed.append((int(row.attrib.get("r", len(parsed) + 1)), values))
    return parsed


def header_for(sheet_name: str, values: list[str]) -> list[str]:
    headers: list[str] = []
    seen: dict[str, int] = {}
    for index, value in enumerate(values):
        header = value.strip() or f"Column {column_letter(index)}"
        if sheet_name.startswith("MasterSurveyComb_"):
            header = MASTER_HEADER_OVERRIDES.get(index, header)
        count = seen.get(header, 0)
        seen[header] = count + 1
        if count:
            header = f"{header} [{column_letter(index)}]"
        headers.append(header)
    return headers


def row_to_record(
    sheet_name: str,
    row_number: int,
    headers: list[str],
    values: list[str],
    context: dict[str, str],
) -> dict[str, Any]:
    cells: dict[str, str] = {}
    cells_by_col: dict[str, str] = {}
    max_len = max(len(headers), len(values))
    for index in range(max_len):
        header = headers[index] if index < len(headers) else f"Column {column_letter(index)}"
        value = values[index] if index < len(values) else ""
        if value:
            cells[header] = value
            cells_by_col[column_letter(index)] = value

    record_context = {
        key: value
        for key, value in context.items()
        if value and key not in cells
    }
    search_parts = list(cells.values()) + list(record_context.values())
    return {
        "sheet": sheet_name,
        "row": row_number,
        "cells": cells,
        "cells_by_col": cells_by_col,
        "context": record_context,
        "search_text": " ".join(search_parts),
    }


def build_index(metadata: dict[str, Any]) -> dict[str, Any]:
    if not WORKBOOK_PATH.exists():
        raise FileNotFoundError(f"Workbook cache not found: {WORKBOOK_PATH}")

    records: list[dict[str, Any]] = []
    sheet_summaries = []
    with zipfile.ZipFile(WORKBOOK_PATH) as zipped:
        shared_strings = load_shared_strings(zipped)
        for sheet_name, sheet_path in workbook_sheets(zipped):
            rows = parse_rows(zipped, sheet_path, shared_strings)
            if not rows:
                continue
            _, header_values = rows[0]
            headers = header_for(sheet_name, header_values)
            context: dict[str, str] = {}
            sheet_count = 0
            for row_number, values in rows[1:]:
                record = row_to_record(sheet_name, row_number, headers, values, context)
                if record["cells"]:
                    records.append(record)
                    sheet_count += 1
                if sheet_name.startswith("MasterSurveyComb_"):
                    for field in CONTEXT_FIELDS:
                        value = record["cells"].get(field, "")
                        if value:
                            context[field] = value
            sheet_summaries.append(
                {
                    "name": sheet_name,
                    "path": sheet_path,
                    "rows_indexed": sheet_count,
                }
            )

    index = {
        "schema_version": INDEX_SCHEMA_VERSION,
        "built_at": now_iso(),
        "metadata": metadata,
        "record_count": len(records),
        "sheets": sheet_summaries,
        "records": records,
    }
    write_json(INDEX_PATH, index)
    return index


def ensure_index(refresh: bool, allow_stale: bool = False) -> dict[str, Any]:
    if refresh:
        metadata, changed = refresh_or_fail(allow_stale=allow_stale)
    else:
        metadata = read_json(METADATA_PATH, {})
        changed = False

    index = read_json(INDEX_PATH, {})
    indexed_sha = index.get("metadata", {}).get("sha256")
    current_sha = metadata.get("sha256") or (file_sha256(WORKBOOK_PATH) if WORKBOOK_PATH.exists() else None)
    if changed or not index or indexed_sha != current_sha or index.get("schema_version") != INDEX_SCHEMA_VERSION:
        return build_index(metadata)
    if refresh and metadata:
        index["metadata"] = metadata
        write_json(INDEX_PATH, index)
    return index


def normalize(value: Any) -> str:
    text = str(value).lower()
    text = re.sub(r"[^a-z0-9_]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def query_tokens(query: str) -> list[str]:
    return [token for token in normalize(query).split() if token not in STOPWORDS]


def get_field(record: dict[str, Any], name: str) -> str:
    return record.get("cells", {}).get(name) or record.get("context", {}).get(name, "")


def get_cell_field(record: dict[str, Any], name: str) -> str:
    return record.get("cells", {}).get(name, "")


def state_attribute(record: dict[str, Any]) -> str:
    combined = " ".join(
        list(record.get("cells", {}).values()) + list(record.get("context", {}).values())
    )
    match = re.search(r"\bstate_d_\d+\b", combined)
    if match:
        return match.group(0)

    state_flag = get_field(record, "State Attribute").lower()
    concept_id = get_field(record, "conceptId_Current Question Text")
    if state_flag == "yes" and concept_id.isdigit():
        return f"state_d_{concept_id}"
    return ""


def concept_ids(record: dict[str, Any]) -> dict[str, str]:
    state_d = state_attribute(record)
    question_concept_id = get_field(record, "conceptId_Current Question Text")
    format_concept_id = get_field(record, "conceptId_Current Format/Value")
    source_concept_id = get_field(record, "conceptId_Source Question")
    preferred = state_d or question_concept_id or format_concept_id or source_concept_id
    return {
        "preferred": preferred,
        "state_d": state_d,
        "question_concept_id": question_concept_id,
        "format_value_concept_id": format_concept_id,
        "source_question_concept_id": source_concept_id,
    }


def record_label(record: dict[str, Any]) -> str:
    for field in [
        "Current Question Text",
        "Question Text",
        "Variable Label",
        "Current Format/Value",
        "Variable Name",
        "Primary Source",
    ]:
        value = get_field(record, field)
        if value:
            return value
    cells = record.get("cells", {})
    return next(iter(cells.values()), "(blank)")


def score_record(record: dict[str, Any], query: str) -> float:
    query_norm = normalize(query)
    tokens = query_tokens(query)
    if not query_norm:
        return 0

    search_text = normalize(record.get("search_text", ""))
    score = 0.0
    if query_norm in search_text:
        score += 120
    for token in tokens:
        if token in search_text:
            score += 8

    for field in IMPORTANT_FIELDS:
        for source, weight in (("cells", 1.0), ("context", 0.45)):
            field_text = record.get(source, {}).get(field, "")
            if not field_text:
                continue
            field_norm = normalize(field_text)
            if query_norm in field_norm:
                score += 220 * weight
            matched_tokens = sum(1 for token in tokens if token in field_norm)
            score += matched_tokens * 28 * weight
            ratio = difflib.SequenceMatcher(None, query_norm, field_norm).ratio()
            score = max(score, ratio * 100 * weight)

    if record.get("cells", {}).get("Current Question Text"):
        score += 35
    if record.get("cells", {}).get("Variable Label"):
        score += 15

    ids = concept_ids(record)
    if query_norm and query_norm in normalize(" ".join(ids.values())):
        score += 300
    return score


def search_records(index: dict[str, Any], query: str, limit: int = 10) -> list[tuple[float, dict[str, Any]]]:
    scored = [
        (score_record(record, query), record)
        for record in index.get("records", [])
    ]
    matches = [(score, record) for score, record in scored if score > 0]
    matches.sort(key=lambda item: (-item[0], item[1]["sheet"], item[1]["row"]))
    return matches[:limit]


def source_line(record: dict[str, Any]) -> str:
    parts = [
        get_field(record, "Primary Source"),
        get_field(record, "Secondary Source"),
    ]
    return " > ".join(part for part in parts if part)


def print_record(record: dict[str, Any], score: float | None = None, rank: int | None = None) -> None:
    ids = concept_ids(record)
    prefix = f"{rank}. " if rank is not None else ""
    print(f"{prefix}{record_label(record)}")
    if ids["preferred"]:
        print(f"   CID: {ids['preferred']}")
    if ids["question_concept_id"] and ids["question_concept_id"] != ids["preferred"]:
        print(f"   Question conceptId: {ids['question_concept_id']}")
    variable = get_field(record, "Variable Name")
    if variable:
        print(f"   Variable: {variable}")
    src = source_line(record)
    if src:
        print(f"   Source: {src}")
    pii = get_field(record, "PII")
    if pii:
        print(f"   PII: {pii}")
    status = get_cell_field(record, STATUS_FIELD)
    if status:
        print(f"   {STATUS_FIELD}: {status}")
    status_date = get_cell_field(record, STATUS_DATE_FIELD)
    if status_date:
        print(f"   {STATUS_DATE_FIELD}: {status_date}")
    gcp_document_table = get_field(record, "GCP Document/Table")
    if gcp_document_table:
        print(f"   GCP Document/Table: {gcp_document_table}")
    location = f"{record['sheet']} row {record['row']}"
    if score is not None:
        location = f"{location}; score {score:.1f}"
    print(f"   Location: {location}")


def print_refresh_status(index: dict[str, Any]) -> None:
    metadata = index.get("metadata", {})
    print(f"Downloaded: {metadata.get('downloaded_at', 'unknown')}")
    print(f"Workbook SHA-256: {metadata.get('sha256', 'unknown')}")
    print(f"Indexed rows: {index.get('record_count', 0)}")
    for sheet in index.get("sheets", []):
        print(f"- {sheet['name']}: {sheet['rows_indexed']} rows")
    warning = metadata.get("warning")
    if warning:
        print(f"Warning: {warning}")


def load_library() -> dict[str, Any]:
    data = read_json(LIBRARY_PATH, {"entries": []})
    data.setdefault("entries", [])
    return data


def save_library(data: dict[str, Any]) -> None:
    data["entries"] = sorted(data.get("entries", []), key=lambda item: normalize(item.get("name", "")))
    write_json(LIBRARY_PATH, data)


def library_entry_from_record(
    name: str,
    query: str,
    record: dict[str, Any],
    metadata: dict[str, Any],
    note: str = "",
) -> dict[str, Any]:
    ids = concept_ids(record)
    return {
        "name": name,
        "query": query,
        "cid": ids["preferred"],
        "state_d": ids["state_d"],
        "question_concept_id": ids["question_concept_id"],
        "format_value_concept_id": ids["format_value_concept_id"],
        "label": record_label(record),
        "variable_name": get_field(record, "Variable Name"),
        "source": source_line(record),
        "gcp_document_table": get_field(record, "GCP Document/Table"),
        "status": get_cell_field(record, STATUS_FIELD),
        "status_date": get_cell_field(record, STATUS_DATE_FIELD),
        "pii": get_field(record, "PII"),
        "sheet": record["sheet"],
        "row": record["row"],
        "note": note,
        "updated_at": now_iso(),
        "workbook_sha256": metadata.get("sha256", ""),
        "workbook_downloaded_at": metadata.get("downloaded_at", ""),
        "source_page": WORKBOOK_PAGE_URL,
    }


def find_library_entry(library: dict[str, Any], query: str) -> dict[str, Any] | None:
    query_norm = normalize(query)
    entries = library.get("entries", [])
    for entry in entries:
        if normalize(entry.get("name", "")) == query_norm:
            return entry
    scored = [
        (
            max(
                difflib.SequenceMatcher(None, query_norm, normalize(entry.get("name", ""))).ratio(),
                difflib.SequenceMatcher(None, query_norm, normalize(entry.get("query", ""))).ratio(),
            ),
            entry,
        )
        for entry in entries
    ]
    scored.sort(key=lambda item: item[0], reverse=True)
    if scored and scored[0][0] >= 0.72:
        return scored[0][1]
    return None


def print_library_entry(entry: dict[str, Any]) -> None:
    print(entry.get("name", "(unnamed)"))
    if entry.get("cid"):
        print(f"   CID: {entry['cid']}")
    if entry.get("label"):
        print(f"   Label: {entry['label']}")
    if entry.get("variable_name"):
        print(f"   Variable: {entry['variable_name']}")
    if entry.get("source"):
        print(f"   Source: {entry['source']}")
    if entry.get("gcp_document_table"):
        print(f"   GCP Document/Table: {entry['gcp_document_table']}")
    if entry.get("status"):
        print(f"   {STATUS_FIELD}: {entry['status']}")
    if entry.get("status_date"):
        print(f"   {STATUS_DATE_FIELD}: {entry['status_date']}")
    if entry.get("pii"):
        print(f"   PII: {entry['pii']}")
    if entry.get("note"):
        print(f"   Note: {entry['note']}")
    if entry.get("sheet") and entry.get("row"):
        print(f"   Location: {entry['sheet']} row {entry['row']}")
    if entry.get("workbook_downloaded_at"):
        print(f"   Workbook refreshed: {entry['workbook_downloaded_at']}")


def command_refresh(args: argparse.Namespace) -> int:
    index = ensure_index(refresh=True, allow_stale=args.allow_stale)
    print_refresh_status(index)
    return 0


def command_search(args: argparse.Namespace) -> int:
    index = ensure_index(refresh=not args.no_refresh, allow_stale=args.allow_stale)
    matches = search_records(index, args.query, limit=args.limit)
    if not matches:
        print(f"No matches found for: {args.query}")
        return 1
    for rank, (score, record) in enumerate(matches, start=1):
        print_record(record, score=score, rank=rank)
    return 0


def command_cid(args: argparse.Namespace) -> int:
    index = ensure_index(refresh=not args.no_refresh, allow_stale=args.allow_stale)
    matches = search_records(index, args.query, limit=max(args.limit, 1))
    if not matches:
        print(f"No CID match found for: {args.query}")
        return 1

    best_score, best_record = matches[0]
    ids = concept_ids(best_record)
    if ids["preferred"]:
        print(ids["preferred"])
    else:
        print("No CID found on the best matching row.")
    print_record(best_record, score=best_score)

    if args.limit > 1 and len(matches) > 1:
        print("\nOther close matches:")
        for rank, (score, record) in enumerate(matches[1:args.limit], start=2):
            print_record(record, score=score, rank=rank)
    return 0


def command_remember(args: argparse.Namespace) -> int:
    index = ensure_index(refresh=not args.no_refresh, allow_stale=args.allow_stale)
    matches = search_records(index, args.query, limit=1)
    if not matches:
        print(f"No match found to remember for: {args.query}")
        return 1

    _, record = matches[0]
    name = args.name or args.query
    library = load_library()
    entry = library_entry_from_record(
        name=name,
        query=args.query,
        record=record,
        metadata=index.get("metadata", {}),
        note=args.note or "",
    )

    entries = [
        existing
        for existing in library.get("entries", [])
        if normalize(existing.get("name", "")) != normalize(name)
    ]
    entries.append(entry)
    library["entries"] = entries
    save_library(library)

    print("Saved reference:")
    print_library_entry(entry)
    return 0


def command_library(args: argparse.Namespace) -> int:
    library = load_library()
    entries = library.get("entries", [])
    if not entries:
        print("Reference library is empty.")
        return 0
    for index, entry in enumerate(entries, start=1):
        cid = entry.get("cid", "")
        label = entry.get("label", "")
        print(f"{index}. {entry.get('name', '(unnamed)')} - {cid}")
        if label:
            print(f"   {label}")
    return 0


def command_show(args: argparse.Namespace) -> int:
    library = load_library()
    entry = find_library_entry(library, args.query)
    if not entry:
        print(f"No saved reference found for: {args.query}")
        return 1
    print_library_entry(entry)
    return 0


def command_forget(args: argparse.Namespace) -> int:
    library = load_library()
    entry = find_library_entry(library, args.query)
    if not entry:
        print(f"No saved reference found for: {args.query}")
        return 1
    library["entries"] = [
        existing
        for existing in library.get("entries", [])
        if existing is not entry
    ]
    save_library(library)
    print(f"Removed reference: {entry.get('name', args.query)}")
    return 0


def add_refresh_flags(parser: argparse.ArgumentParser) -> None:
    parser.add_argument(
        "--allow-stale",
        action="store_true",
        help="Use the local workbook cache if GitHub refresh fails.",
    )
    parser.add_argument(
        "--no-refresh",
        action="store_true",
        help="Skip the automatic GitHub refresh for this command.",
    )


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Refresh and search the Connect master/survey combined data dictionary.",
    )
    subparsers = parser.add_subparsers(dest="command", required=True)

    refresh = subparsers.add_parser("refresh", help="Download and index the latest workbook.")
    refresh.add_argument(
        "--allow-stale",
        action="store_true",
        help="Use the local workbook cache if GitHub refresh fails.",
    )
    refresh.set_defaults(func=command_refresh)

    search = subparsers.add_parser("search", help="Search the refreshed dictionary.")
    search.add_argument("query")
    search.add_argument("--limit", type=int, default=10)
    add_refresh_flags(search)
    search.set_defaults(func=command_search)

    cid = subparsers.add_parser("cid", help="Return the best CID/state attribute match.")
    cid.add_argument("query")
    cid.add_argument("--limit", type=int, default=1, help="Show additional close matches.")
    add_refresh_flags(cid)
    cid.set_defaults(func=command_cid)

    remember = subparsers.add_parser("remember", help="Save the best current match to the library.")
    remember.add_argument("query")
    remember.add_argument("--name", help="Reference name to save. Defaults to the query.")
    remember.add_argument("--note", help="Optional internal note.")
    add_refresh_flags(remember)
    remember.set_defaults(func=command_remember)

    library = subparsers.add_parser("library", help="List saved references.")
    library.set_defaults(func=command_library)

    show = subparsers.add_parser("show", help="Show one saved reference.")
    show.add_argument("query")
    show.set_defaults(func=command_show)

    forget = subparsers.add_parser("forget", help="Remove one saved reference.")
    forget.add_argument("query")
    forget.set_defaults(func=command_forget)
    return parser


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)
    try:
        return args.func(args)
    except RuntimeError as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 2
    except KeyboardInterrupt:
        print("Interrupted.", file=sys.stderr)
        return 130


if __name__ == "__main__":
    raise SystemExit(main())
