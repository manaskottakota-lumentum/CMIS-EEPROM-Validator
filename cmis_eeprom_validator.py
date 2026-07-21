from __future__ import annotations

import argparse
import csv
import math
import os
import re
import sys
import zipfile
import xml.etree.ElementTree as ET
from dataclasses import dataclass
from typing import Dict, Iterable, List, Optional, Tuple


IGNORED_WORKSHEET_KEYS = {"CMISMAP"}


@dataclass
class ExpectedParameter:
    parameter: str
    page: str
    start: int
    length: int
    data_type: str
    expected: str
    comparator: str = "exact"
    significance: str = ""
    mode: str = ""
    tolerance: str = ""
    minimum: str = ""
    maximum: str = ""
    mask: str = ""
    description: str = ""
    check: str = ""
    active_check: bool = True
    bank: str = ""
    address_text: str = ""
    cmis_type: str = ""
    simple_type: str = ""
    value: str = ""
    is_memory_map: bool = False
    source_sheet: str = ""
    source_row: int = 0
    bits: str = ""


@dataclass
class ValidationResult:
    parameter: ExpectedParameter
    status: str
    actual: str
    expected: str
    raw_hex: str
    page: str
    address: str
    message: str


class ValidationError(Exception):
    pass


def normalize_header(value: str) -> str:
    text = re.sub(r"[^a-z0-9]+", "_", str(value).strip().lower()).strip("_")
    aliases = {
        "parameter": "parameter",
        "parameter_name": "parameter",
        "name": "parameter",
        "field": "parameter",
        "field_name": "parameter",
        "byte": "byte",
        "hex": "hex",
        "bits": "bits",
        "check": "check",
        "checked": "check",
        "active": "check",
        "enabled": "check",
        "page": "page",
        "cmis_page": "page",
        "memory_page": "page",
        "bank": "bank",
        "start": "start",
        "start_address": "start",
        "address": "address",
        "offset": "start",
        "byte_offset": "start",
        "length": "length",
        "bytes": "length",
        "size": "length",
        "size_bytes": "length",
        "size_byte": "length",
        "data_type": "data_type",
        "type": "data_type",
        "format": "data_type",
        "cmis_type": "cmis_type",
        "cmis_rev_5_0_type": "cmis_type",
        "simple_type": "simple_type",
        "programming_type": "simple_type",
        "expected": "expected",
        "expected_value": "expected",
        "spec": "expected",
        "msft_spec": "expected",
        "spec_value": "expected",
        "target": "expected",
        "value": "value",
        "actual": "value",
        "actual_value": "value",
        "comparator": "comparator",
        "compare": "comparator",
        "rule": "comparator",
        "comparison": "comparator",
        "significance": "significance",
        "impact": "significance",
        "meaning": "significance",
        "failure_significance": "significance",
        "mode": "mode",
        "power_mode": "mode",
        "tolerance": "tolerance",
        "tol": "tolerance",
        "min": "minimum",
        "minimum": "minimum",
        "max": "maximum",
        "maximum": "maximum",
        "mask": "mask",
        "bit_mask": "mask",
        "description": "description",
    }
    return aliases.get(text, text)


def compact_text(value: object) -> str:
    if value is None:
        return ""
    text = str(value).replace("\r", " ").replace("\n", " ").strip()
    if text.endswith(".0"):
        whole = text[:-2]
        if re.fullmatch(r"-?\d+", whole):
            return whole
    return text


def worksheet_key(value: object) -> str:
    return re.sub(r"[^A-Z0-9]+", "", compact_text(value).upper())


def is_check_yes(value: object) -> bool:
    return compact_text(value).strip().lower() in {"yes", "y", "true", "1", "x", "checked"}


def parse_int(value: object) -> int:
    text = compact_text(value)
    if not text:
        raise ValidationError("Missing integer value.")
    if isinstance(value, int):
        return value
    if isinstance(value, float) and value.is_integer():
        return int(value)
    text = text.replace("_", "").replace(",", "")
    if text.lower().startswith("0x"):
        return int(text, 16)
    if text.lower().startswith("0b"):
        return int(text, 2)
    if text.lower().endswith("h"):
        return int(text[:-1], 16)
    if text.lower().endswith("b") and re.fullmatch(r"[01]+b", text.lower()):
        return int(text[:-1], 2)
    return int(float(text))


def parse_float(value: object) -> float:
    text = compact_text(value)
    if not text:
        raise ValidationError("Missing numeric value.")
    number = re.search(r"-?\d+(?:\.\d+)?(?:[eE][+-]?\d+)?", text.replace(",", ""))
    if not number:
        raise ValidationError(f"Missing numeric value: {text}")
    return float(number.group(0))


def parse_length(value: object, start: Optional[int] = None, end: Optional[int] = None) -> int:
    text = compact_text(value)
    byte_count = re.search(r"(\d+)\s*(?:bytes?|b)\b", text, re.IGNORECASE)
    if byte_count:
        return int(byte_count.group(1))
    if text:
        range_match = re.search(r"(?:0x)?([0-9a-fA-F]{1,2})h?\s*-\s*(?:0x)?([0-9a-fA-F]{1,2})h?", text)
        if range_match:
            return int(range_match.group(2), 16) - int(range_match.group(1), 16) + 1
        return parse_int(text)
    if start is not None and end is not None:
        return end - start + 1
    return 1


def parse_address_range(page_value: object, address_value: object, length_value: object) -> Tuple[str, int, int]:
    page_text = compact_text(page_value).lower()
    address_text = compact_text(address_value)

    upper_match = re.search(
        r"(?:page\s*)?(?:0x)?([0-9a-fA-F]{1,2})h?(?:\s*-\s*(?:0x)?([0-9a-fA-F]{1,2})h?)?\s*:\s*(?:0x)?([0-9a-fA-F]{1,2})h?"
        r"(?:\s*-\s*(?:0x)?([0-9a-fA-F]{1,2})h?)?",
        address_text,
    )
    if upper_match:
        page = f"{int(upper_match.group(1), 16):02X}"
        start = int(upper_match.group(3), 16)
        end = int(upper_match.group(4), 16) if upper_match.group(4) else None
        return page, start, parse_length(length_value, start, end)

    range_match = re.search(r"(?:0x)?([0-9a-fA-F]{1,2})h?(?:\s*-\s*(?:0x)?([0-9a-fA-F]{1,2})h?)?", address_text)
    if range_match:
        start = int(range_match.group(1), 16)
        end = int(range_match.group(2), 16) if range_match.group(2) else None
        if "upper" in page_text:
            raise ValidationError("Upper-memory rows need an address like 01h:80h-FFh so the upper page number is explicit.")
        return "lower", start, parse_length(length_value, start, end)

    raise ValidationError(f"Could not parse address range: {address_text}")


def normalize_page(value: object) -> str:
    text = compact_text(value).lower()
    if not text or text in {"lower", "lower memory", "lower page"}:
        return "lower"
    match = re.search(r"(?:page\s*)?(?:0x)?([0-9a-fA-F]{1,2})h?", text)
    if match:
        return f"{int(match.group(1), 16):02X}"
    if "upper" in text:
        return "upper"
    return text


def normalize_mode(value: object) -> str:
    text = compact_text(value).lower()
    if "low" in text:
        return "low"
    if "high" in text:
        return "high"
    return text


def hex_tokens_from_line(line: str) -> List[int]:
    if ":" in line:
        payload = line.split(":", 1)[1]
    else:
        payload = line
        if not re.fullmatch(r"\s*(?:0x)?[0-9A-Fa-f]{2}(?:h)?(?:[\s,;-]+(?:0x)?[0-9A-Fa-f]{2}(?:h)?)*\s*", payload):
            return []
    tokens = re.findall(r"(?<![0-9A-Fa-f])(?:0x)?([0-9A-Fa-f]{2})(?:h)?(?![0-9A-Fa-f])", payload)
    return [int(token, 16) for token in tokens]


def packed_hex_tokens_from_payload(payload: str) -> List[int]:
    tokens: List[int] = []
    for word in re.findall(r"(?<![0-9A-Fa-f])([0-9A-Fa-f]{4,64})(?![0-9A-Fa-f])", payload):
        if len(word) % 2:
            continue
        tokens.extend(int(word[index : index + 2], 16) for index in range(0, len(word), 2))
    return tokens


def parse_dump_page_header(line: str) -> Optional[str]:
    lower_line = line.lower()
    if "lower" in lower_line and re.search(r"\b(page|memory|register|idprom|eeprom|cmis)\b", lower_line):
        return "lower"

    if "upper" not in lower_line:
        return None

    page_match = re.search(r"\bpage\s+(?:0x)?([0-9a-fA-F]{1,2})h?\b", line, re.IGNORECASE)
    if not page_match:
        page_match = re.search(r"\bupper\s+(?:page\s*)?(?:0x)?([0-9a-fA-F]{1,2})h?\b", line, re.IGNORECASE)
    if page_match:
        return f"{int(page_match.group(1), 16):02X}"

    return None


def parse_dump_data_line(line: str) -> Tuple[Optional[int], List[int]]:
    offset_match = re.match(r"^\s*([0-9A-Fa-f]{4})\s*:?\s+(.*)$", line)
    if offset_match:
        payload = offset_match.group(2)
        tokens = hex_tokens_from_line(payload) or packed_hex_tokens_from_payload(payload)
        return int(offset_match.group(1), 16), tokens

    return None, hex_tokens_from_line(line)


def write_dump_tokens(page_data: bytearray, page: str, offset: Optional[int], tokens: List[int]) -> None:
    if not tokens:
        return
    if offset is None:
        page_data.extend(tokens)
        return
    if page != "lower" and offset < 128:
        offset += 128
    end = offset + len(tokens)
    if len(page_data) < end:
        page_data.extend(b"\x00" * (end - len(page_data)))
    page_data[offset:end] = bytes(tokens)


class CmisDump:
    def __init__(self) -> None:
        self.pages: Dict[str, Dict[str, bytearray]] = {}

    @classmethod
    def from_file(cls, path: str) -> "CmisDump":
        for encoding in ("utf-8-sig", "utf-16", "latin-1"):
            try:
                with open(path, "r", encoding=encoding) as handle:
                    return cls.from_text(handle.read())
            except UnicodeError:
                continue
        raise ValidationError(f"Could not read hex dump file: {path}")

    @classmethod
    def from_text(cls, text: str) -> "CmisDump":
        dump = cls()
        current_mode = "default"
        current_page = "lower"
        dump.pages.setdefault(current_mode, {}).setdefault(current_page, bytearray())

        for raw_line in text.splitlines():
            line = raw_line.strip()
            if not line:
                continue
            lower_line = line.lower()
            if "low power" in lower_line:
                current_mode = "low"
                dump.pages.setdefault(current_mode, {})
            elif "high power" in lower_line:
                current_mode = "high"
                dump.pages.setdefault(current_mode, {})

            page_header = parse_dump_page_header(line)
            if page_header is not None:
                current_page = page_header
                dump.pages.setdefault(current_mode, {}).setdefault(current_page, bytearray())
                continue

            if "lower page" in lower_line or "lower memory" in lower_line:
                current_page = "lower"
                dump.pages.setdefault(current_mode, {}).setdefault(current_page, bytearray())
                if not hex_tokens_from_line(line):
                    continue
            page_match = re.search(r"upper\s+page\s+(?:0x)?([0-9a-fA-F]{1,2})h?", line, re.IGNORECASE)
            if page_match:
                current_page = f"{int(page_match.group(1), 16):02X}"
                dump.pages.setdefault(current_mode, {}).setdefault(current_page, bytearray())
                if not hex_tokens_from_line(line):
                    continue
            offset, tokens = parse_dump_data_line(line)
            if tokens:
                page_data = dump.pages.setdefault(current_mode, {}).setdefault(current_page, bytearray())
                write_dump_tokens(page_data, current_page, offset, tokens)
        return dump

    def page_bytes(self, page: str, mode: str = "") -> bytes:
        normalized_page = normalize_page(page)
        normalized_mode = normalize_mode(mode)
        mode_order: List[str] = []
        if normalized_mode:
            mode_order.append(normalized_mode)
        mode_order.extend(["default", "low", "high"])
        mode_order.extend(self.pages.keys())

        seen = set()
        for candidate_mode in mode_order:
            if candidate_mode in seen:
                continue
            seen.add(candidate_mode)
            pages = self.pages.get(candidate_mode, {})
            if normalized_page in pages and pages[normalized_page]:
                return bytes(pages[normalized_page])
            if normalized_page != "lower" and "lower" in pages and len(pages["lower"]) >= 256:
                start = 128 if normalized_page == "00" else 0
                return bytes(pages["lower"][start : start + 128])
        raise ValidationError(f"Page {page} was not found in the EEPROM dump.")


def read_delimited_file(path: str) -> List[Dict[str, str]]:
    for encoding in ("utf-8-sig", "latin-1"):
        try:
            with open(path, "r", newline="", encoding=encoding) as handle:
                sample = handle.read(2048)
                handle.seek(0)
                try:
                    dialect = csv.Sniffer().sniff(sample, delimiters=",\t;")
                except csv.Error:
                    dialect = csv.excel_tab if path.lower().endswith(".tsv") else csv.excel
                return rows_to_dicts(list(csv.reader(handle, dialect)))
        except UnicodeError:
            continue
    raise ValidationError(f"Could not read expected-values file: {path}")


def read_xlsx_file(path: str) -> List[Dict[str, str]]:
    with zipfile.ZipFile(path) as archive:
        shared_strings = read_shared_strings(archive)
        sheet_refs = read_workbook_sheet_refs(archive)
        if not sheet_refs:
            raise ValidationError("The Excel file does not contain any worksheets.")
        records: List[Dict[str, str]] = []
        for sheet_name, sheet_path in sheet_refs:
            if worksheet_key(sheet_name) in IGNORED_WORKSHEET_KEYS:
                continue
            if sheet_path not in archive.namelist():
                continue
            rows = read_sheet_rows(archive, sheet_path, shared_strings)
            records.extend(rows_to_dicts(rows, sheet_name=sheet_name, raise_if_missing=False))
        if not records:
            raise ValidationError("The Excel file does not contain any CMIS map rows.")
        return records


def read_workbook_sheet_refs(archive: zipfile.ZipFile) -> List[Tuple[str, str]]:
    names = archive.namelist()
    if "xl/workbook.xml" not in names or "xl/_rels/workbook.xml.rels" not in names:
        return [
            (os.path.splitext(os.path.basename(name))[0], name)
            for name in sorted(names)
            if name.startswith("xl/worksheets/sheet") and name.endswith(".xml")
        ]
    workbook_root = ET.fromstring(archive.read("xl/workbook.xml"))
    rels_root = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
    workbook_ns = {
        "x": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
        "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    }
    rels_ns = {"rel": "http://schemas.openxmlformats.org/package/2006/relationships"}
    rels = {
        relationship.attrib.get("Id", ""): relationship.attrib.get("Target", "")
        for relationship in rels_root.findall("rel:Relationship", rels_ns)
    }
    refs: List[Tuple[str, str]] = []
    for sheet in workbook_root.findall("x:sheets/x:sheet", workbook_ns):
        sheet_name = sheet.attrib.get("name", "")
        rel_id = sheet.attrib.get(f"{{{workbook_ns['r']}}}id", "")
        target = rels.get(rel_id, "")
        if not target:
            continue
        sheet_path = target.lstrip("/") if target.startswith("/") else "xl/" + target.lstrip("/")
        refs.append((sheet_name, sheet_path.replace("\\", "/")))
    return refs


def read_shared_strings(archive: zipfile.ZipFile) -> List[str]:
    if "xl/sharedStrings.xml" not in archive.namelist():
        return []
    root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
    namespace = {"x": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
    values: List[str] = []
    for item in root.findall("x:si", namespace):
        values.append("".join(node.text or "" for node in item.findall(".//x:t", namespace)))
    return values


def read_sheet_rows(archive: zipfile.ZipFile, sheet_path: str, shared_strings: List[str]) -> List[List[str]]:
    root = ET.fromstring(archive.read(sheet_path))
    namespace = {"x": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
    rows: List[List[str]] = []
    for row in root.findall(".//x:sheetData/x:row", namespace):
        values: Dict[int, str] = {}
        max_index = -1
        for cell in row.findall("x:c", namespace):
            ref = cell.attrib.get("r", "")
            col_index = column_index_from_ref(ref)
            values[col_index] = read_cell_value(cell, namespace, shared_strings)
            max_index = max(max_index, col_index)
        if max_index >= 0:
            rows.append([values.get(index, "") for index in range(max_index + 1)])
    return rows


def column_index_from_ref(ref: str) -> int:
    letters = re.match(r"([A-Z]+)", ref.upper())
    if not letters:
        return 0
    index = 0
    for char in letters.group(1):
        index = index * 26 + (ord(char) - ord("A") + 1)
    return index - 1


def read_cell_value(cell: ET.Element, namespace: Dict[str, str], shared_strings: List[str]) -> str:
    cell_type = cell.attrib.get("t", "")
    if cell_type == "inlineStr":
        return "".join(node.text or "" for node in cell.findall(".//x:t", namespace))
    value_node = cell.find("x:v", namespace)
    if value_node is None or value_node.text is None:
        return ""
    value = value_node.text
    if cell_type == "s":
        try:
            return shared_strings[int(value)]
        except (IndexError, ValueError):
            return value
    if cell_type == "b":
        return "TRUE" if value == "1" else "FALSE"
    return value


def msft_page_from_sheet_name(sheet_name: str) -> str:
    text = compact_text(sheet_name).strip().lower()
    if text == "lower memory":
        return "lower"
    match = re.fullmatch(r"page\s+([0-9a-fA-F]{1,2})h", text)
    if match:
        return f"{int(match.group(1), 16):02X}"
    return ""


def parse_msft_byte_range(value: object, previous_start: Optional[int], previous_end: Optional[int]) -> Tuple[Optional[int], Optional[int]]:
    text = compact_text(value)
    if not text:
        return previous_start, previous_end
    match = re.fullmatch(r"(\d+)(?:\s*-\s*(\d+))?", text)
    if not match:
        raise ValidationError(f"Could not parse MSFT Byte value: {text}")
    start = int(match.group(1))
    end = int(match.group(2)) if match.group(2) else start
    return start, end


def msft_address(page: str, start: int, end: int) -> str:
    start_hex = f"{start:02X}h"
    end_hex = f"{end:02X}h"
    address = start_hex if start == end else f"{start_hex}-{end_hex}"
    return address if page == "lower" else f"{page}h:{address}"


def bit_mask_from_text(bits: object) -> Optional[int]:
    text = compact_text(bits).strip()
    if not text or text.lower() == "all":
        return None
    range_match = re.fullmatch(r"(\d+)\s*-\s*(\d+)", text)
    if range_match:
        high = int(range_match.group(1))
        low = int(range_match.group(2))
        if high == 7 and low == 0:
            return None
        return sum(1 << bit for bit in range(low, high + 1))
    if re.fullmatch(r"\d+", text):
        return 1 << int(text)
    return None


def lowest_set_bit(mask: int) -> int:
    return (mask & -mask).bit_length() - 1


def int_to_hex_bytes(value: int, length: int) -> str:
    length = max(1, length)
    modulus = 1 << (8 * length)
    value %= modulus
    hex_text = f"{value:0{length * 2}X}"
    return " ".join(hex_text[index : index + 2] for index in range(0, len(hex_text), 2))


def hex_bytes_from_msft_spec(spec: object, length: int) -> str:
    text = compact_text(spec).upper()
    match = re.search(r"0X([0-9A-F]+)", text) or re.search(r"([0-9A-F]+)H", text)
    if not match:
        return ""
    digits = match.group(1)
    if len(digits) % 2:
        digits = "0" + digits
    if length > 1 and len(digits) == 2:
        digits *= length
    return " ".join(digits[index : index + 2] for index in range(0, len(digits), 2))


def scaled_unit_hex_from_msft_spec(spec: object, field_name: object, length: int) -> str:
    text = compact_text(spec)
    field = compact_text(field_name).upper()
    number_match = re.search(r"-?\d+(?:\.\d+)?", text)
    if not number_match:
        return ""
    value = float(number_match.group(0))
    upper_text = text.upper()
    if "TEMP" in field and re.search(r"(?:\u02daC|\u00b0C|\u2103|DEGC|\bC\b)", upper_text):
        scaled = round(value * 256) if length > 1 else round(value)
        return int_to_hex_bytes(scaled, length)
    if ("VCC" in field or "VOLTAGE" in field) and "V" in upper_text:
        scaled = round(value / 0.0001)
        return int_to_hex_bytes(scaled, length)
    return ""


def normalize_msft_requirement(spec: object, length: int, bits: object, field_name: object) -> Tuple[str, str, str, str]:
    text = compact_text(spec)
    if not text or text == "*":
        return "", "read", "", "No"
    mask = bit_mask_from_text(bits)
    binary_match = re.fullmatch(r"([01]+)b", text.strip(), re.IGNORECASE)
    if binary_match:
        value = int(binary_match.group(1), 2)
        if mask is not None:
            return str(value << lowest_set_bit(mask)), "mask", str(mask), "YES"
        return int_to_hex_bytes(value, length), "exact", "", "YES"
    hex_expected = hex_bytes_from_msft_spec(text, length)
    if hex_expected:
        if mask is not None:
            return str(int(normalize_hex(hex_expected), 16)), "mask", str(mask), "YES"
        return hex_expected, "exact", "", "YES"
    unit_expected = scaled_unit_hex_from_msft_spec(text, field_name, length)
    if unit_expected:
        return unit_expected, "exact", "", "YES"
    if looks_like_hex_wildcard_bytes(text):
        return text, "exact", "", "YES"
    return "", "read", "", "No"


def msft_rows_to_dicts(rows: List[List[str]], headers: List[str], header_row_index: int, sheet_name: str) -> List[Dict[str, str]]:
    page = msft_page_from_sheet_name(sheet_name)
    if not page:
        return []
    records: List[Dict[str, str]] = []
    current_start: Optional[int] = None
    current_end: Optional[int] = None
    for row_index, row in enumerate(rows[header_row_index + 1 :], start=header_row_index + 2):
        if not any(compact_text(cell) for cell in row):
            continue
        raw_record: Dict[str, str] = {"source_row": str(row_index), "source_sheet": sheet_name}
        for col_index, header in enumerate(headers):
            if header:
                raw_record[header] = compact_text(row[col_index]) if col_index < len(row) else ""
        start, end = parse_msft_byte_range(raw_record.get("byte", ""), current_start, current_end)
        if start is None or end is None:
            continue
        if raw_record.get("byte", ""):
            current_start, current_end = start, end
        length = end - start + 1
        field_name = raw_record.get("parameter", "")
        expected, comparator, mask, check = normalize_msft_requirement(
            raw_record.get("expected", ""), length, raw_record.get("bits", "7-0"), field_name
        )
        description = raw_record.get("description", "")
        description = f"{field_name} - {description}" if description else field_name
        records.append(
            {
                "source_row": str(row_index),
                "source_sheet": sheet_name,
                "check": check,
                "parameter": field_name,
                "page": "lower" if page == "lower" else "Upper",
                "address": msft_address(page, start, end),
                "length": f"{length} byte{'s' if length != 1 else ''}",
                "bits": raw_record.get("bits", "7-0"),
                "data_type": raw_record.get("data_type", "") or "Hex",
                "expected": expected,
                "value": expected,
                "comparator": comparator,
                "mask": mask,
                "description": description,
                "significance": "MSFT requirement mismatch for this CMIS field.",
            }
        )
    return records


def rows_to_dicts(rows: List[List[str]], sheet_name: str = "", raise_if_missing: bool = True) -> List[Dict[str, str]]:
    header_row_index = None
    headers: List[str] = []
    for index, row in enumerate(rows):
        normalized = [normalize_header(cell) for cell in row]
        has_msft_requirements_shape = (
            bool(msft_page_from_sheet_name(sheet_name))
            and "byte" in normalized
            and "parameter" in normalized
            and "expected" in normalized
        )
        has_expected_shape = "parameter" in normalized and "expected" in normalized and (
            "address" in normalized or "start" in normalized or "page" in normalized
        )
        has_memory_map_shape = "page" in normalized and "address" in normalized and "length" in normalized
        if has_msft_requirements_shape:
            return msft_rows_to_dicts(rows, normalized, index, sheet_name)
        if has_expected_shape or has_memory_map_shape:
            header_row_index = index
            headers = normalized
            break
    if header_row_index is None:
        if not raise_if_missing:
            return []
        raise ValidationError(
            "Spreadsheet needs parameter/expected columns, CMIS map Page/Address/Length columns, or an MSFT Byte/MSFT spec memory-map sheet."
        )
    records: List[Dict[str, str]] = []
    for row_index, row in enumerate(rows[header_row_index + 1 :], start=header_row_index + 2):
        if not any(compact_text(cell) for cell in row):
            continue
        normalized_row = [normalize_header(cell) for cell in row]
        if "page" in normalized_row and "address" in normalized_row and "length" in normalized_row:
            continue
        record: Dict[str, str] = {"source_row": str(row_index), "source_sheet": sheet_name}
        for col_index, header in enumerate(headers):
            if header:
                record[header] = compact_text(row[col_index]) if col_index < len(row) else ""
        records.append(record)
    return records


def load_expected_parameters(path: str) -> List[ExpectedParameter]:
    ext = os.path.splitext(path)[1].lower()
    if ext == ".xlsx":
        records = read_xlsx_file(path)
    elif ext in {".csv", ".tsv", ".txt"}:
        records = read_delimited_file(path)
    else:
        raise ValidationError("Expected-values file must be .xlsx, .csv, .tsv, or .txt.")
    parameters: List[ExpectedParameter] = []
    for record in records:
        try:
            address_text = record.get("address", "").strip()
            if address_text:
                page, start, length = parse_address_range(record.get("page", "lower"), address_text, record.get("length", "1"))
            else:
                page = record.get("page", "lower").strip() or "lower"
                start = parse_int(record.get("start", ""))
                length = parse_length(record.get("length", "1"))
            cmis_type = record.get("cmis_type", "").strip()
            simple_type = record.get("simple_type", "").strip()
            data_type = record.get("data_type", "").strip() or cmis_type or simple_type or "Hex"
            description = record.get("description", "").strip()
            value = record.get("value", "").strip()
            expected = record.get("expected", "").strip() or value
            has_check_column = "check" in record
            check = record.get("check", "").strip()
            active_check = is_check_yes(check) if has_check_column else bool(expected)
            comparator = record.get("comparator", "").strip() or ("exact" if active_check and expected else "read")
            parameter = ExpectedParameter(
                parameter=record.get("parameter", "").strip() or description or address_text or f"Row {record.get('source_row', '?')}",
                page=page,
                start=start,
                length=length,
                data_type=data_type,
                expected=expected,
                comparator=comparator,
                significance=record.get("significance", "").strip(),
                mode=record.get("mode", "").strip(),
                tolerance=record.get("tolerance", "").strip(),
                minimum=record.get("minimum", "").strip(),
                maximum=record.get("maximum", "").strip(),
                mask=record.get("mask", "").strip(),
                description=description,
                check=check,
                active_check=active_check,
                bank=record.get("bank", "").strip() or "N/A",
                address_text=address_text,
                cmis_type=cmis_type or data_type,
                simple_type=simple_type,
                value=value,
                is_memory_map=bool(address_text),
                source_sheet=record.get("source_sheet", "").strip(),
                source_row=parse_int(record.get("source_row", "0")) if record.get("source_row", "") else 0,
                bits=record.get("bits", "").strip(),
            )
            parameters.append(parameter)
        except Exception as exc:
            raise ValidationError(f"Could not parse expected-values row {record.get('source_row', '?')}: {exc}") from exc
    return parameters


def read_parameter_bytes(dump: CmisDump, parameter: ExpectedParameter) -> bytes:
    data = dump.page_bytes(parameter.page, parameter.mode)
    if parameter.start < 0 or parameter.length < 1:
        raise ValidationError("Invalid address or length.")
    end = parameter.start + parameter.length
    if end > len(data):
        raise ValidationError(f"Dump does not contain bytes for {address_display(parameter)}.")
    return data[parameter.start:end]


def decode_value(raw: bytes, data_type: str) -> str:
    dtype = compact_text(data_type).lower()
    compact_dtype = re.sub(r"[^a-z0-9]+", "", dtype)
    if "ascii" in dtype or dtype in {"str", "string"}:
        return raw.decode("ascii", errors="replace").rstrip()
    if "bit" in dtype:
        return "0b" + "".join(f"{byte:08b}" for byte in raw)
    if compact_dtype in {"u8", "u16", "u32", "u64", "uint", "unsigned", "unsignedinteger", "int", "decimal"}:
        return str(int.from_bytes(raw, byteorder="big", signed=False))
    if compact_dtype in {"s8", "s16", "s32", "s64", "signed", "signedinteger"}:
        return str(int.from_bytes(raw, byteorder="big", signed=True))
    return raw_hex_value(raw)


def raw_hex_value(raw: bytes) -> str:
    return " ".join(f"{byte:02X}" for byte in raw)


def address_display(parameter: ExpectedParameter) -> str:
    page = normalize_page(parameter.page)
    start = parameter.start
    end = parameter.start + parameter.length - 1
    if page == "lower":
        return f"0x{start:02X}-0x{end:02X}" if end != start else f"0x{start:02X}"
    return f"{page}h / 0x{start:02X}-0x{end:02X}" if end != start else f"{page}h / 0x{start:02X}"


def normalize_hex(value: object) -> str:
    text = compact_text(value).upper()
    text = text.replace("0X", "").replace("H", "")
    return re.sub(r"[^0-9A-F]", "", text)


def looks_like_hex_bytes(value: object) -> bool:
    text = compact_text(value)
    return bool(
        re.fullmatch(
            r"\s*(?:0x)?[0-9A-Fa-f]{2}(?:h)?(?:[\s,;-]+(?:0x)?[0-9A-Fa-f]{2}(?:h)?)*\s*",
            text,
        )
    )


def looks_like_hex_wildcard_bytes(value: object) -> bool:
    text = compact_text(value)
    if "-" not in text:
        return False
    return bool(
        re.fullmatch(
            r"\s*(?:0x)?[0-9A-Fa-f-]{2}(?:h)?(?:[\s,;]+(?:0x)?[0-9A-Fa-f-]{2}(?:h)?)*\s*",
            text,
        )
    )


def normalize_hex_wildcard(value: object) -> str:
    text = compact_text(value).upper()
    text = text.replace("0X", "").replace("H", "")
    return re.sub(r"[\s,;]+", "", text)


def hex_wildcard_match(actual: object, expected: object) -> bool:
    actual_hex = normalize_hex(actual)
    expected_pattern = normalize_hex_wildcard(expected)
    if len(actual_hex) != len(expected_pattern):
        return False
    return all(expected_char == "-" or expected_char == actual_char for actual_char, expected_char in zip(actual_hex, expected_pattern))


def values_match(parameter: ExpectedParameter, actual: str, raw: bytes) -> Tuple[bool, str]:
    comparator = parameter.comparator.strip().lower() or "exact"
    dtype = parameter.data_type.strip().lower()
    compact_dtype = re.sub(r"[^a-z0-9]+", "", dtype)
    expected = parameter.expected
    if comparator == "range":
        if parameter.minimum or parameter.maximum:
            minimum = parse_float(parameter.minimum)
            maximum = parse_float(parameter.maximum)
        elif ".." in expected:
            left, right = expected.split("..", 1)
            minimum = parse_float(left)
            maximum = parse_float(right)
        else:
            raise ValidationError("Range comparisons need min/max columns or an expected value like 3.1..3.6.")
        actual_value = parse_float(actual)
        return minimum <= actual_value <= maximum, f"Expected {minimum:g} to {maximum:g}"
    if comparator == "tolerance":
        actual_value = parse_float(actual)
        expected_value = parse_float(expected)
        tolerance = parse_float(parameter.tolerance)
        return math.isclose(actual_value, expected_value, abs_tol=tolerance), f"Expected {expected_value:g} +/- {tolerance:g}"
    if comparator == "mask":
        mask = parse_int(parameter.mask)
        actual_value = int.from_bytes(raw, byteorder="big", signed=False) & mask
        expected_value = parse_int(expected) & mask
        return actual_value == expected_value, f"Expected masked value 0x{expected_value:X} with mask 0x{mask:X}"
    if looks_like_hex_bytes(actual) and looks_like_hex_wildcard_bytes(expected):
        return hex_wildcard_match(actual, expected), f"Expected {expected} (- is wildcard)"
    if dtype in {"hex", "byte", "bytes"}:
        return normalize_hex(actual) == normalize_hex(expected), f"Expected {expected}"
    if looks_like_hex_bytes(actual) and looks_like_hex_bytes(expected):
        return normalize_hex(actual) == normalize_hex(expected), f"Expected {expected}"
    if dtype in {"decimal", "unsigned", "unsigned integer", "signed", "signed integer", "int", "uint", "float", "float32"} or re.fullmatch(r"[us](8|16|32|64)", compact_dtype):
        try:
            return parse_float(actual) == parse_float(expected), f"Expected {expected}"
        except ValidationError:
            pass
    return actual.strip() == expected.strip(), f"Expected {expected}"


def compared_display_values(parameter: ExpectedParameter, actual: str, raw: bytes) -> Tuple[str, str]:
    if parameter.comparator.strip().lower() == "mask":
        mask = parse_int(parameter.mask)
        shift = lowest_set_bit(mask)
        actual_value = (int.from_bytes(raw, byteorder="big", signed=False) & mask) >> shift
        expected_value = (parse_int(parameter.expected) & mask) >> shift
        return str(actual_value), str(expected_value)
    return actual, parameter.expected


def validation_result_for_missing_reference(parameter: ExpectedParameter) -> ValidationResult:
    return ValidationResult(
        parameter=parameter,
        status="NOT CHECKED",
        actual="",
        expected=parameter.expected,
        raw_hex="",
        page=normalize_page(parameter.page),
        address=address_display(parameter),
        message="Check is not YES, so this row is visible as a reference row but was not automatically checked.",
    )


def validate_parameters(dump: CmisDump, parameters: Iterable[ExpectedParameter]) -> List[ValidationResult]:
    results: List[ValidationResult] = []
    for parameter in parameters:
        page = normalize_page(parameter.page)
        address = address_display(parameter)
        if not parameter.active_check:
            try:
                raw = read_parameter_bytes(dump, parameter)
                actual = raw_hex_value(raw)
                raw_hex = actual
            except Exception:
                results.append(validation_result_for_missing_reference(parameter))
                continue
            results.append(
                ValidationResult(
                    parameter=parameter,
                    status="NOT CHECKED",
                    actual=actual,
                    expected=parameter.expected,
                    raw_hex=raw_hex,
                    page=page,
                    address=address,
                    message="Check is not YES, so this row is visible as a reference row but was not automatically checked.",
                )
            )
            continue
        try:
            raw = read_parameter_bytes(dump, parameter)
            raw_hex = raw_hex_value(raw)
            actual = raw_hex if parameter.address_text or parameter.is_memory_map or parameter.comparator.lower() == "read" else decode_value(raw, parameter.data_type)
            if parameter.comparator.lower() == "read" or not parameter.expected:
                results.append(
                    ValidationResult(
                        parameter=parameter,
                        status="READ",
                        actual=actual,
                        expected=parameter.value or "",
                        raw_hex=raw_hex,
                        page=page,
                        address=address,
                        message=parameter.description or "Value populated from the dump using Page, Address, and Length.",
                    )
                )
                continue
            passed, expectation = values_match(parameter, actual, raw)
            display_actual, display_expected = compared_display_values(parameter, actual, raw)
            status = "PASS" if passed else "FAIL"
            if passed:
                message = parameter.significance or "Value matches the expected specification."
            else:
                significance = parameter.significance or "This mismatch means the module EEPROM content does not match the expected FAI specification."
                message = f"{significance} Actual value is {display_actual}; expected value is {display_expected}."
            results.append(
                ValidationResult(
                    parameter=parameter,
                    status=status,
                    actual=display_actual,
                    expected=display_expected if parameter.comparator.lower() != "range" else expectation,
                    raw_hex=raw_hex,
                    page=page,
                    address=address,
                    message=message,
                )
            )
        except Exception as exc:
            results.append(
                ValidationResult(
                    parameter=parameter,
                    status="ERROR",
                    actual="",
                    expected=parameter.expected,
                    raw_hex="",
                    page=page,
                    address=address,
                    message=str(exc),
                )
            )
    return results


def summarize_results(results: Iterable[ValidationResult]) -> Dict[str, int]:
    result_list = list(results)
    return {
        "total": len(result_list),
        "passed": sum(1 for result in result_list if result.status == "PASS"),
        "failed": sum(1 for result in result_list if result.status == "FAIL"),
        "read": sum(1 for result in result_list if result.status == "READ"),
        "notChecked": sum(1 for result in result_list if result.status == "NOT CHECKED"),
        "errors": sum(1 for result in result_list if result.status == "ERROR"),
    }


def run_self_test() -> int:
    base_dir = os.path.dirname(os.path.abspath(__file__))
    dump_path = os.path.join(base_dir, "sample_first_6_digits_eeprom_dump.hex")
    expected_path = os.path.join(base_dir, "samples", "sample_expected_values.csv")
    dump = CmisDump.from_file(dump_path)
    parameters = load_expected_parameters(expected_path)
    results = validate_parameters(dump, parameters)
    summary = summarize_results(results)
    print(summary)
    if summary["passed"] >= 1 and summary["failed"] == 0 and summary["errors"] == 0:
        print("Self-test passed.")
        return 0
    print("Self-test failed.")
    return 1


def main(argv: Optional[List[str]] = None) -> int:
    parser = argparse.ArgumentParser(description="Validate a CMIS EEPROM/register dump against expected values.")
    parser.add_argument("--self-test", action="store_true", help="Run the included sample validation.")
    parser.add_argument("--dump", help="Path to EEPROM/register hex dump.")
    parser.add_argument("--expected", help="Path to expected-values workbook/CSV.")
    args = parser.parse_args(argv)
    if args.self_test:
        return run_self_test()
    if args.dump and args.expected:
        dump = CmisDump.from_file(args.dump)
        parameters = load_expected_parameters(args.expected)
        results = validate_parameters(dump, parameters)
        print(summarize_results(results))
        for result in results:
            if result.status in {"FAIL", "ERROR"}:
                print(f"{result.status}: {result.parameter.parameter} {result.address}: {result.message}")
        return 0 if all(result.status not in {"FAIL", "ERROR"} for result in results) else 1
    parser.print_help()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
