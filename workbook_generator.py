from __future__ import annotations

import os
import math
from pathlib import Path
import re
import shutil
import struct
import tempfile
import zipfile
from copy import copy
from xml.sax.saxutils import escape

from openpyxl import load_workbook

from cmis_eeprom_validator import CmisDump


DUMP_VALUE_HEADER = "Dump Value"
CONVERTED_VALUE_HEADER = "Converted Value"
LOWER_HEADERS = ["Byte", "Hex", "Bits", "Field Name", "Type", DUMP_VALUE_HEADER, CONVERTED_VALUE_HEADER, "Description"]
UPPER_HEADERS = ["Byte", "Hex", "Size(bytes)", "Bits", "Field Name", "Type", DUMP_VALUE_HEADER, CONVERTED_VALUE_HEADER, "Description"]
DEFAULT_TEMPLATE_NAME = "1_6T_2xDR4_MemoryMap_MSFT_requirements_v1_2.xlsx"


def generate_workbook_from_dump(dump_path: str, output_path: str) -> int:
    template_path = find_default_template()
    if template_path:
        return generate_template_workbook_from_dump(dump_path, output_path, template_path)
    return generate_basic_workbook_from_dump(dump_path, output_path)


def find_default_template() -> Path | None:
    candidates = [
        Path(__file__).resolve().parent / "samples" / DEFAULT_TEMPLATE_NAME,
        Path(__file__).resolve().parent / DEFAULT_TEMPLATE_NAME,
        Path(os.environ.get("USERPROFILE", "")) / "OneDrive - Lumentum Operations LLC" / "Documents" / DEFAULT_TEMPLATE_NAME,
    ]
    for candidate in candidates:
        if candidate.exists() and is_readable_file(candidate):
            return candidate
    return None


def is_readable_file(path: Path) -> bool:
    try:
        with path.open("rb"):
            return True
    except OSError:
        return False


def generate_template_workbook_from_dump(dump_path: str, output_path: str, template_path: Path) -> int:
    dump = CmisDump.from_file(dump_path)
    sheets = dump.pages.get("default") or next(iter(dump.pages.values()), {})
    output = Path(output_path)
    output.parent.mkdir(parents=True, exist_ok=True)

    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as template_copy:
        temp_template = Path(template_copy.name)
    try:
        shutil.copyfile(template_path, temp_template)
        workbook = load_workbook(temp_template)
    finally:
        try:
            temp_template.unlink()
        except OSError:
            pass

    row_count = 0
    for worksheet in workbook.worksheets:
        page = page_from_template_sheet(worksheet.title)
        if not page:
            continue
        page_bytes = sheets.get(page)
        headers = header_columns(worksheet)
        byte_col = headers.get("byte")
        spec_col = headers.get("msft spec") or headers.get("spec") or headers.get("dump value") or headers.get("converted value")
        values_col = headers.get("values") or headers.get("value") or headers.get("converted value")
        bits_col = headers.get("bits")
        field_col = headers.get("field name") or headers.get("name")
        type_col = headers.get("type")
        description_col = headers.get("description")
        if not byte_col or not spec_col:
            continue
        worksheet.cell(1, spec_col).value = DUMP_VALUE_HEADER
        values_col = ensure_values_column(worksheet, spec_col, values_col)
        if description_col and description_col >= values_col:
            description_col += 1
        expand_template_reserved_custom_rows(
            worksheet,
            byte_col=byte_col,
            hex_col=headers.get("hex"),
            size_col=headers.get("size(bytes)") or headers.get("size"),
            field_col=field_col,
            bits_col=bits_col,
            type_col=type_col,
            description_col=description_col,
        )

        current_start: int | None = None
        current_end: int | None = None
        for row_index in range(2, worksheet.max_row + 1):
            start, end = parse_byte_range(worksheet.cell(row_index, byte_col).value, current_start, current_end)
            if start is None or end is None:
                continue
            current_start, current_end = start, end
            bits_value = worksheet.cell(row_index, bits_col).value if bits_col else None
            bits = str(bits_value) if bits_value is not None and str(bits_value).strip() != "" else "7-0"
            field_name = str(worksheet.cell(row_index, field_col).value or "") if field_col else ""
            type_text = str(worksheet.cell(row_index, type_col).value or "") if type_col else ""
            value = template_value_for_range(page_bytes, start, end, bits)
            description_text = str(worksheet.cell(row_index, description_col).value or "") if description_col else ""
            interpreted = interpreted_value_for_range(page_bytes, start, end, bits, field_name, type_text)
            worksheet.cell(row_index, spec_col).value = value
            worksheet.cell(row_index, values_col).value = value_column_text(interpreted, description_text, value, page_bytes, start, end, bits)
            row_count += 1

    row_count += ensure_defined_page_sheet(workbook, sheets, "03")
    for page in ("10", "11"):
        row_count += ensure_defined_page_sheet(workbook, sheets, page)
    row_count += ensure_defined_page_sheet(workbook, sheets, "1E")
    row_count += ensure_defined_page_sheet(workbook, sheets, "1F")

    if row_count == 0:
        raise ValueError("No matching MSFT memory-map rows were found in the template workbook.")
    rename_msft_spec_headers(workbook)
    workbook.save(output)
    return row_count


def rename_msft_spec_headers(workbook) -> None:
    for worksheet in workbook.worksheets:
        for row in worksheet.iter_rows():
            for cell in row:
                if str(cell.value or "").strip().lower() in {"msft spec", "spec"}:
                    cell.value = DUMP_VALUE_HEADER


def ensure_values_column(worksheet, spec_col: int, values_col: int | None) -> int:
    if values_col:
        worksheet.cell(1, values_col).value = CONVERTED_VALUE_HEADER
        return values_col

    values_col = spec_col + 1
    worksheet.insert_cols(values_col)
    worksheet.cell(1, values_col).value = CONVERTED_VALUE_HEADER
    copy_column_format(worksheet, spec_col, values_col)
    return values_col


def copy_column_format(worksheet, source_col: int, target_col: int) -> None:
    source_letter = worksheet.cell(1, source_col).column_letter
    target_letter = worksheet.cell(1, target_col).column_letter
    if source_letter in worksheet.column_dimensions:
        worksheet.column_dimensions[target_letter].width = worksheet.column_dimensions[source_letter].width
    for row_index in range(1, worksheet.max_row + 1):
        source = worksheet.cell(row_index, source_col)
        target = worksheet.cell(row_index, target_col)
        if source.has_style:
            target.font = copy(source.font)
            target.fill = copy(source.fill)
            target.border = copy(source.border)
            target.alignment = copy(source.alignment)
            target.number_format = source.number_format
            target.protection = copy(source.protection)


def expand_template_reserved_custom_rows(
    worksheet,
    byte_col: int,
    hex_col: int | None,
    size_col: int | None,
    field_col: int | None,
    bits_col: int | None,
    type_col: int | None,
    description_col: int | None,
) -> None:
    if not field_col:
        return
    for row_index in range(worksheet.max_row, 1, -1):
        start, end = parse_byte_range(worksheet.cell(row_index, byte_col).value, None, None)
        if start is None or end is None or end <= start:
            continue
        definition = {
            "start": start,
            "end": end,
            "bits": worksheet.cell(row_index, bits_col).value if bits_col else "7-0",
            "field": worksheet.cell(row_index, field_col).value or "",
            "type": worksheet.cell(row_index, type_col).value if type_col else "",
            "description": worksheet.cell(row_index, description_col).value if description_col else "",
        }
        if not is_reserved_or_custom_section(definition):
            continue
        source_values = [worksheet.cell(row_index, column).value for column in range(1, worksheet.max_column + 1)]
        worksheet.insert_rows(row_index + 1, end - start)
        for offset, address in enumerate(range(start, end + 1)):
            target_row = row_index + offset
            if offset:
                copy_row_format(worksheet, row_index, target_row)
                for column, value in enumerate(source_values, start=1):
                    worksheet.cell(target_row, column).value = value
            worksheet.cell(target_row, byte_col).value = address
            if hex_col:
                worksheet.cell(target_row, hex_col).value = f"{address:02X}"
            if size_col:
                worksheet.cell(target_row, size_col).value = 1
            if bits_col:
                worksheet.cell(target_row, bits_col).value = "7-0"
            worksheet.cell(target_row, field_col).value = byte_section_name(str(definition["field"]), address)


def copy_row_format(worksheet, source_row: int, target_row: int) -> None:
    worksheet.row_dimensions[target_row].height = worksheet.row_dimensions[source_row].height
    for column in range(1, worksheet.max_column + 1):
        source = worksheet.cell(source_row, column)
        target = worksheet.cell(target_row, column)
        if source.has_style:
            target.font = copy(source.font)
            target.fill = copy(source.fill)
            target.border = copy(source.border)
            target.alignment = copy(source.alignment)
            target.number_format = source.number_format
            target.protection = copy(source.protection)


def ensure_defined_page_sheet(workbook, sheets: dict[str, bytearray], page: str) -> int:
    page_bytes = sheets.get(page)
    definitions = cmis_page_definitions(page)
    if not definitions:
        return 0
    definitions = split_reserved_custom_ranges(definitions)

    sheet_name = _sheet_name(page)
    if sheet_name in workbook.sheetnames:
        worksheet = workbook[sheet_name]
    else:
        reference = workbook["Page 2Fh"] if "Page 2Fh" in workbook.sheetnames else workbook.worksheets[-1]
        worksheet = workbook.copy_worksheet(reference)
        worksheet.title = sheet_name

    style_template = capture_sheet_styles(worksheet)
    clear_sheet_values(worksheet)
    populate_defined_page_sheet(worksheet, page, page_bytes, style_template, definitions)
    return len(definitions)


def ensure_custom_page_sheet(workbook, sheets: dict[str, bytearray], page: str) -> int:
    page_bytes = sheets.get(page)
    if page_bytes is None:
        return 0

    sheet_name = _sheet_name(page)
    if sheet_name in workbook.sheetnames:
        worksheet = workbook[sheet_name]
    else:
        reference = workbook["Page 2Fh"] if "Page 2Fh" in workbook.sheetnames else workbook.worksheets[-1]
        worksheet = workbook.copy_worksheet(reference)
        worksheet.title = sheet_name

    style_template = capture_sheet_styles(worksheet)
    clear_sheet_values(worksheet)
    definitions = [
        field(
            address,
            address,
            "7-0",
            f"CustomPage{page.upper()}Byte{address:02X}h",
            "Custom",
            f"Vendor-specific custom byte {address:02X}h from {_page_label(page)}",
        )
        for address in range(128, min(len(page_bytes), 256))
    ]
    populate_defined_page_sheet(worksheet, page, page_bytes, style_template, definitions)
    return len(definitions)


def capture_sheet_styles(worksheet) -> dict[tuple[int, int], object]:
    styles: dict[tuple[int, int], object] = {}
    for row in range(1, min(worksheet.max_row, 3) + 1):
        for column in range(1, min(worksheet.max_column, len(UPPER_HEADERS)) + 1):
            cell = worksheet.cell(row, column)
            if cell.has_style:
                styles[(row, column)] = (
                    copy(cell.font),
                    copy(cell.fill),
                    copy(cell.border),
                    copy(cell.alignment),
                    cell.number_format,
                    copy(cell.protection),
                )
    return styles


def clear_sheet_values(worksheet) -> None:
    for merged_range in list(worksheet.merged_cells.ranges):
        worksheet.unmerge_cells(str(merged_range))
    if worksheet.max_row:
        worksheet.delete_rows(1, worksheet.max_row)


def populate_defined_page_sheet(
    worksheet,
    page: str,
    page_bytes: bytearray,
    style_template: dict[tuple[int, int], object],
    definitions: list[dict[str, object]],
) -> None:
    for column_index, header in enumerate(UPPER_HEADERS, start=1):
        cell = worksheet.cell(1, column_index)
        cell.value = header
        apply_captured_style(cell, style_template.get((1, column_index)))

    for row_index, definition in enumerate(definitions, start=2):
        start = int(definition["start"])
        end = int(definition.get("end", start))
        bits = str(definition.get("bits", "7-0"))
        value = template_value_for_range(page_bytes, start, end, bits)
        interpreted = interpreted_value_for_range(
            page_bytes,
            start,
            end,
            bits,
            str(definition["field"]),
            str(definition.get("type", "")),
        )
        interpreted = value_column_text(interpreted, str(definition.get("description", "")), value, page_bytes, start, end, bits)
        values = [
            byte_range_label(start, end),
            hex_range_label(start, end),
            end - start + 1,
            bits,
            definition["field"],
            definition.get("type", "Hex"),
            value,
            interpreted,
            definition.get("description", f"{definition['field']} from {_page_label(page)}"),
        ]
        for column_index, value in enumerate(values, start=1):
            cell = worksheet.cell(row_index, column_index)
            cell.value = value
            apply_captured_style(cell, style_template.get((min(row_index, 3), column_index)))

    worksheet.freeze_panes = "A2"
    worksheet.auto_filter.ref = f"A1:I{worksheet.max_row}"


def cmis_page_definitions(page: str) -> list[dict[str, object]]:
    if page == "03":
        return page_03h_definitions()
    if page == "10":
        return page_10h_definitions()
    if page == "11":
        return page_11h_definitions()
    if page == "1E":
        return page_1eh_definitions()
    if page == "1F":
        return page_1fh_definitions()
    return []


def page_03h_definitions() -> list[dict[str, object]]:
    return [
        field(
            address,
            address,
            "7-0",
            f"UserEepromByte{address:02X}h",
            "RO/RW",
            "CMIS Page 03h user EEPROM byte",
        )
        for address in range(128, 256)
    ]


def page_10h_definitions() -> list[dict[str, object]]:
    rows: list[dict[str, object]] = []
    rows += bit_lane_rows(128, "DPDeinitLane", "RW Rqd.", "Data Path initialization control")
    rows += bit_lane_rows(129, "InputPolarityFlipTx", "RW Adv.", "Tx input polarity flip")
    rows += bit_lane_rows(130, "OutputDisableTx", "RW Adv.", "Tx output disable")
    rows += bit_lane_rows(131, "AutoSquelchDisableTx", "RW Adv.", "Disable automatic Tx squelch")
    rows += bit_lane_rows(132, "OutputSquelchForceTx", "RW Adv.", "Force Tx output squelch")
    rows.append(field(133, 133, "7-0", "Reserved[1]", "RO", "Reserved byte"))
    rows += bit_lane_rows(134, "AdaptiveInputEqFreezeTx", "RW Adv.", "Freeze Tx input equalizer adaptation")
    rows += two_bit_lane_rows(135, "AdaptiveInputEqStoreTx", 1, 4, "WO Adv.", "Store Tx input equalizer adaptation")
    rows += two_bit_lane_rows(136, "AdaptiveInputEqStoreTx", 5, 8, "WO Adv.", "Store Tx input equalizer adaptation")
    rows += bit_lane_rows(137, "OutputPolarityFlipRx", "RW Adv.", "Rx output polarity flip")
    rows += bit_lane_rows(138, "OutputDisableRx", "RW Adv.", "Rx output disable")
    rows += bit_lane_rows(139, "AutoSquelchDisableRx", "RW Adv.", "Disable automatic Rx squelch")
    rows.append(field(140, 142, "7-0", "Reserved[3]", "RO", "Reserved bytes"))
    rows += bit_lane_rows(143, "SCS0ApplyDPInitLane", "WO Rqd.", "Staged Control Set 0 ApplyDPInit trigger")
    rows += bit_lane_rows(144, "SCS0ApplyImmediateLane", "WO Cnd.", "Staged Control Set 0 ApplyImmediate trigger")
    rows += dp_config_rows(145, "SCS0", "RW Rqd.")
    rows += bit_lane_rows(153, "SCS0AdaptiveInputEqEnableTx", "RW Adv.", "SCS0 adaptive Tx input equalizer enable")
    rows += two_bit_lane_rows(154, "SCS0AdaptiveInputEqRecallTx", 1, 4, "RW Adv.", "SCS0 recall Tx input equalizer adaptation")
    rows += two_bit_lane_rows(155, "SCS0AdaptiveInputEqRecallTx", 5, 8, "RW Adv.", "SCS0 recall Tx input equalizer adaptation")
    rows += nibble_lane_rows(156, "SCS0FixedInputEqTargetTx", 1, 2, "RW Adv.", "SCS0 fixed Tx input equalizer target")
    rows += nibble_lane_rows(157, "SCS0FixedInputEqTargetTx", 3, 4, "RW Adv.", "SCS0 fixed Tx input equalizer target")
    rows += nibble_lane_rows(158, "SCS0FixedInputEqTargetTx", 5, 6, "RW Adv.", "SCS0 fixed Tx input equalizer target")
    rows += nibble_lane_rows(159, "SCS0FixedInputEqTargetTx", 7, 8, "RW Adv.", "SCS0 fixed Tx input equalizer target")
    rows += bit_lane_rows(160, "SCS0CDREnableTx", "RW Adv.", "SCS0 Tx CDR enable")
    rows += bit_lane_rows(161, "SCS0CDREnableRx", "RW Adv.", "SCS0 Rx CDR enable")
    rows += rx_nibble_target_rows(162, "SCS0OutputEqPreCursorTargetRx", "RW Adv.", "SCS0 Rx output pre-cursor target")
    rows += rx_nibble_target_rows(166, "SCS0OutputEqPostCursorTargetRx", "RW Adv.", "SCS0 Rx output post-cursor target")
    rows += rx_nibble_target_rows(170, "SCS0OutputAmplitudeTargetRx", "RW Adv.", "SCS0 Rx output amplitude target")
    rows.append(field(174, 177, "7-0", "Reserved[4]", "RO", "Reserved bytes"))
    rows += bit_lane_rows(178, "SCS1ApplyDPInitLane", "WO Adv.", "Staged Control Set 1 ApplyDPInit trigger")
    rows += bit_lane_rows(179, "SCS1ApplyImmediateLane", "WO Cnd.", "Staged Control Set 1 ApplyImmediate trigger")
    rows += dp_config_rows(180, "SCS1", "RW Adv.")
    rows += bit_lane_rows(188, "SCS1AdaptiveInputEqEnableTx", "RW Adv.", "SCS1 adaptive Tx input equalizer enable")
    rows += two_bit_lane_rows(189, "SCS1AdaptiveInputEqRecallTx", 1, 4, "RW Adv.", "SCS1 recall Tx input equalizer adaptation")
    rows += two_bit_lane_rows(190, "SCS1AdaptiveInputEqRecallTx", 5, 8, "RW Adv.", "SCS1 recall Tx input equalizer adaptation")
    rows += nibble_lane_rows(191, "SCS1FixedInputEqTargetTx", 1, 2, "RW Adv.", "SCS1 fixed Tx input equalizer target")
    rows += nibble_lane_rows(192, "SCS1FixedInputEqTargetTx", 3, 4, "RW Adv.", "SCS1 fixed Tx input equalizer target")
    rows += nibble_lane_rows(193, "SCS1FixedInputEqTargetTx", 5, 6, "RW Adv.", "SCS1 fixed Tx input equalizer target")
    rows += nibble_lane_rows(194, "SCS1FixedInputEqTargetTx", 7, 8, "RW Adv.", "SCS1 fixed Tx input equalizer target")
    rows += bit_lane_rows(195, "SCS1CDREnableTx", "RW Adv.", "SCS1 Tx CDR enable")
    rows += bit_lane_rows(196, "SCS1CDREnableRx", "RW Adv.", "SCS1 Rx CDR enable")
    rows += rx_nibble_target_rows(197, "SCS1OutputEqPreCursorTargetRx", "RW Adv.", "SCS1 Rx output pre-cursor target")
    rows += rx_nibble_target_rows(201, "SCS1OutputEqPostCursorTargetRx", "RW Adv.", "SCS1 Rx output post-cursor target")
    rows += rx_nibble_target_rows(205, "SCS1OutputAmplitudeTargetRx", "RW Adv.", "SCS1 Rx output amplitude target")
    rows.append(field(209, 212, "7-0", "Reserved[4]", "RO", "Reserved bytes"))
    for address, name, type_text in [
        (213, "DPStateChangedMask", "RW Rqd."),
        (214, "FailureMaskTx", "RW Adv."),
        (215, "LOSMaskTx", "RW Adv."),
        (216, "CDRLOLMaskTx", "RW Adv."),
        (217, "AdaptiveInputEqFailMaskTx", "RW Adv."),
        (218, "OpticalPowerHighAlarmMaskTx", "RW Adv."),
        (219, "OpticalPowerLowAlarmMaskTx", "RW Adv."),
        (220, "OpticalPowerHighWarningMaskTx", "RW Adv."),
        (221, "OpticalPowerLowWarningMaskTx", "RW Adv."),
        (222, "LaserBiasHighAlarmMaskTx", "RW Adv."),
        (223, "LaserBiasLowAlarmMaskTx", "RW Adv."),
        (224, "LaserBiasHighWarningMaskTx", "RW Adv."),
        (225, "LaserBiasLowWarningMaskTx", "RW Adv."),
        (226, "LOSMaskRx", "RW Adv."),
        (227, "CDRLOLMaskRx", "RW Adv."),
        (228, "OpticalPowerHighAlarmMaskRx", "RW Adv."),
        (229, "OpticalPowerLowAlarmMaskRx", "RW Adv."),
        (230, "OpticalPowerHighWarningMaskRx", "RW Adv."),
        (231, "OpticalPowerLowWarningMaskRx", "RW Adv."),
        (232, "OutputStatusChangedMaskRx", "RW Rqd."),
    ]:
        rows += bit_lane_rows(address, name, type_text, f"Mask bits for {name.replace('Mask', 'Flag')}")
    rows.append(field(233, 239, "7-0", "Reserved[7]", "RO", "Reserved bytes"))
    rows.append(field(240, 255, "7-0", "Custom[16]", "Custom", "Custom bytes"))
    return sorted(rows, key=lambda row: (int(row["start"]), bit_sort_key(str(row.get("bits", "7-0")))))


def page_11h_definitions() -> list[dict[str, object]]:
    rows: list[dict[str, object]] = []
    for address, first_lane in zip(range(128, 132), range(1, 8, 2)):
        rows += nibble_lane_rows(address, "DPStateHostLane", first_lane, first_lane + 1, "RO Rqd.", "Data Path state")
    rows += bit_lane_rows(132, "OutputStatusRx", "RO Rqd.", "Rx output signal status")
    rows += bit_lane_rows(133, "OutputStatusTx", "RO Rqd.", "Tx output signal status")
    for address, name, type_text in [
        (134, "DPStateChangedFlag", "RO/COR Rqd."),
        (135, "FailureFlagTx", "RO/COR Adv."),
        (136, "LOSFlagTx", "RO/COR Adv."),
        (137, "CDRLOLFlagTx", "RO/COR Adv."),
        (138, "AdaptiveInputEqFailFlagTx", "RO/COR Adv."),
        (139, "OpticalPowerHighAlarmFlagTx", "RO/COR Adv."),
        (140, "OpticalPowerLowAlarmFlagTx", "RO/COR Adv."),
        (141, "OpticalPowerHighWarningFlagTx", "RO/COR Adv."),
        (142, "OpticalPowerLowWarningFlagTx", "RO/COR Adv."),
        (143, "LaserBiasHighAlarmFlagTx", "RO/COR Adv."),
        (144, "LaserBiasLowAlarmFlagTx", "RO/COR Adv."),
        (145, "LaserBiasHighWarningFlagTx", "RO/COR Adv."),
        (146, "LaserBiasLowWarningFlagTx", "RO/COR Adv."),
        (147, "LOSFlagRx", "RO/COR Adv."),
        (148, "CDRLOLFlagRx", "RO/COR Adv."),
        (149, "OpticalPowerHighAlarmFlagRx", "RO/COR Adv."),
        (150, "OpticalPowerLowAlarmFlagRx", "RO/COR Adv."),
        (151, "OpticalPowerHighWarningFlagRx", "RO/COR Adv."),
        (152, "OpticalPowerLowWarningFlagRx", "RO/COR Adv."),
        (153, "OutputStatusChangedFlagRx", "RO/COR Rqd."),
    ]:
        rows += bit_lane_rows(address, name, type_text, f"Lane-specific {name}")
    rows += u16_lane_rows(154, "OpticalPowerTx", "RO Adv.", "Measured Tx output optical power")
    rows += u16_lane_rows(170, "LaserBiasTx", "RO Adv.", "Measured Tx laser bias current")
    rows += u16_lane_rows(186, "OpticalPowerRx", "RO Adv.", "Measured Rx input optical power")
    for address, first_lane in zip(range(202, 206), range(1, 8, 2)):
        rows += nibble_lane_rows(address, "ConfigStatusLane", first_lane, first_lane + 1, "RO Rqd.", "Configuration command status")
    rows += dp_config_rows(206, "ACS", "RO Rqd.")
    rows += bit_lane_rows(214, "ACSAdaptiveInputEqEnableTx", "RO Adv.", "Active Tx adaptive equalizer enable")
    rows += two_bit_lane_rows(215, "ACSAdaptiveInputEqRecalledTx", 1, 4, "RO Adv.", "Active Tx equalizer recall status")
    rows += two_bit_lane_rows(216, "ACSAdaptiveInputEqRecalledTx", 5, 8, "RO Adv.", "Active Tx equalizer recall status")
    rows += nibble_lane_rows(217, "ACSFixedInputEqTargetTx", 1, 2, "RO Adv.", "Active fixed Tx input equalizer target")
    rows += nibble_lane_rows(218, "ACSFixedInputEqTargetTx", 3, 4, "RO Adv.", "Active fixed Tx input equalizer target")
    rows += nibble_lane_rows(219, "ACSFixedInputEqTargetTx", 5, 6, "RO Adv.", "Active fixed Tx input equalizer target")
    rows += nibble_lane_rows(220, "ACSFixedInputEqTargetTx", 7, 8, "RO Adv.", "Active fixed Tx input equalizer target")
    rows += bit_lane_rows(221, "ACSCDREnableTx", "RO Adv.", "Active Tx CDR enable")
    rows += bit_lane_rows(222, "ACSCDREnableRx", "RO Adv.", "Active Rx CDR enable")
    rows += rx_nibble_target_rows(223, "ACSOutputEqPreCursorTargetRx", "RO Adv.", "Active Rx output pre-cursor target")
    rows += rx_nibble_target_rows(227, "ACSOutputEqPostCursorTargetRx", "RO Adv.", "Active Rx output post-cursor target")
    rows += rx_nibble_target_rows(231, "ACSOutputAmplitudeTargetRx", "RO Adv.", "Active Rx output amplitude target")
    rows += bit_lane_rows(235, "DPInitPendingLane", "RO Rqd.", "Data Path initialization pending")
    rows.append(field(236, 239, "7-0", "Reserved[4]", "RO", "Reserved bytes"))
    for lane, address in enumerate(range(240, 248), start=1):
        rows.append(field(address, address, "7-4", f"MediaLaneToWavelengthMappingTx{lane}", "RO Cnd.", "Tx media lane wavelength mapping"))
        rows.append(field(address, address, "3-0", f"MediaLaneToFiberMappingTx{lane}", "RO Cnd.", "Tx media lane fiber mapping"))
    for lane, address in enumerate(range(248, 256), start=1):
        rows.append(field(address, address, "7-4", f"MediaLaneToWavelengthMappingRx{lane}", "RO Cnd.", "Rx media lane wavelength mapping"))
        rows.append(field(address, address, "3-0", f"MediaLaneToFiberMappingRx{lane}", "RO Cnd.", "Rx media lane fiber mapping"))
    return sorted(rows, key=lambda row: (int(row["start"]), bit_sort_key(str(row.get("bits", "7-0")))))


def page_1eh_definitions() -> list[dict[str, object]]:
    rows: list[dict[str, object]] = []
    rows.append(field(128, 131, "7-0", "GoogleCustomReserved[4]", "RO", "Google custom area reserved for future/custom use"))
    threshold_names = [
        "Custom2DSP2TempHighAlarmThreshold",
        "Custom2DSP2TempLowAlarmThreshold",
        "Custom2DSP2TempHighWarningThreshold",
        "Custom2DSP2TempLowWarningThreshold",
        "Custom1DSP1TempHighAlarmThreshold",
        "Custom1DSP1TempLowAlarmThreshold",
        "Custom1DSP1TempHighWarningThreshold",
        "Custom1DSP1TempLowWarningThreshold",
        "Custom4TIAPD2TempHighAlarmThreshold",
        "Custom4TIAPD2TempLowAlarmThreshold",
        "Custom4TIAPD2TempHighWarningThreshold",
        "Custom4TIAPD2TempLowWarningThreshold",
        "Custom3TIAPD1TempHighAlarmThreshold",
        "Custom3TIAPD1TempLowAlarmThreshold",
        "Custom3TIAPD1TempHighWarningThreshold",
        "Custom3TIAPD1TempLowWarningThreshold",
        "Custom6PCB2TempHighAlarmThreshold",
        "Custom6PCB2TempLowAlarmThreshold",
        "Custom6PCB2TempHighWarningThreshold",
        "Custom6PCB2TempLowWarningThreshold",
        "Custom5PCB1TempHighAlarmThreshold",
        "Custom5PCB1TempLowAlarmThreshold",
        "Custom5PCB1TempHighWarningThreshold",
        "Custom5PCB1TempLowWarningThreshold",
        "Custom8PIC2TempHighAlarmThreshold",
        "Custom8PIC2TempLowAlarmThreshold",
        "Custom8PIC2TempHighWarningThreshold",
        "Custom8PIC2TempLowWarningThreshold",
        "Custom7PIC1TempHighAlarmThreshold",
        "Custom7PIC1TempLowAlarmThreshold",
        "Custom7PIC1TempHighWarningThreshold",
        "Custom7PIC1TempLowWarningThreshold",
        "Custom10Driver2TempHighAlarmThreshold",
        "Custom10Driver2TempLowAlarmThreshold",
        "Custom10Driver2TempHighWarningThreshold",
        "Custom10Driver2TempLowWarningThreshold",
        "Custom9Driver1TempHighAlarmThreshold",
        "Custom9Driver1TempLowAlarmThreshold",
        "Custom9Driver1TempHighWarningThreshold",
        "Custom9Driver1TempLowWarningThreshold",
        "Custom12UncooledLaser1TempHighAlarmThreshold",
        "Custom12UncooledLaser1TempLowAlarmThreshold",
        "Custom12UncooledLaser1TempHighWarningThreshold",
        "Custom12UncooledLaser1TempLowWarningThreshold",
        "Custom13UncooledLaser2TempHighAlarmThreshold",
        "Custom13UncooledLaser2TempLowAlarmThreshold",
        "Custom13UncooledLaser2TempHighWarningThreshold",
        "Custom13UncooledLaser2TempLowWarningThreshold",
        "Custom14UncooledLaser3TempHighAlarmThreshold",
        "Custom14UncooledLaser3TempLowAlarmThreshold",
        "Custom14UncooledLaser3TempHighWarningThreshold",
        "Custom14UncooledLaser3TempLowWarningThreshold",
        "Custom15UncooledLaser4TempHighAlarmThreshold",
        "Custom15UncooledLaser4TempLowAlarmThreshold",
        "Custom15UncooledLaser4TempHighWarningThreshold",
        "Custom15UncooledLaser4TempLowWarningThreshold",
    ]
    for address, name in zip(range(132, 188), threshold_names):
        rows.append(field(address, address, "7-0", name, "RO", "Google custom non-volatile component temperature threshold"))
    rows.append(field(188, 223, "7-0", "Reserved[36]", "RO", "Google custom reserved bytes"))
    for lane, address in enumerate(range(224, 240, 2), start=1):
        rows.append(field(address, address + 1, "7-0", f"TxLaserLaneBiasBoL{lane}", "RO", "Saved Tx laser bias current at beginning of life"))
    for lane, address in enumerate(range(240, 256, 2), start=1):
        rows.append(field(address, address + 1, "7-0", f"TxLanePowerBoL{lane}", "RO", "Saved Tx optical lane power at beginning of life"))
    return rows


def page_1fh_definitions() -> list[dict[str, object]]:
    rows: list[dict[str, object]] = []
    for lane, address in enumerate(range(128, 136), start=1):
        rows.append(
            field(
                address,
                address,
                "7-0",
                f"FailureCauseMediaTxLane{lane}",
                "RO/COR",
                "Media Tx lane latched failure cause",
            )
        )
    for lane, address in enumerate(range(136, 144), start=1):
        rows.append(
            field(
                address,
                address,
                "7-0",
                f"FailureCauseMediaRxLane{lane}",
                "RO/COR",
                "Media Rx lane latched failure cause",
            )
        )
    custom_flag_rows = [
        (144, "7-4", "Custom2Flags"),
        (144, "3-0", "Custom1Flags"),
        (145, "7-4", "Custom4Flags"),
        (145, "3-0", "Custom3Flags"),
        (146, "7-4", "Custom6Flags"),
        (146, "3-0", "Custom5Flags"),
        (147, "7-4", "Custom8Flags"),
        (147, "3-0", "Custom7Flags"),
        (148, "7-4", "Custom10Flags"),
        (148, "3-0", "Custom9Flags"),
        (149, "7-4", "Custom12Flags"),
        (149, "3-0", "Custom13Flags"),
        (150, "7-4", "Custom14Flags"),
        (151, "3-0", "Custom15Flags"),
    ]
    for address, bits, name in custom_flag_rows:
        rows.append(field(address, address, bits, name, "RO/COR", "Google custom latched flag group; contributes to 00h:3.0"))
    rows.append(field(152, 152, "7-0", "Reserved[1]", "RO", "Reserved for future custom flags"))

    temp_names = [
        "Custom1DSP1Temp",
        "Custom2DSP2Temp",
        "Custom3TIAPD1Temp",
        "Custom4TIAPD2Temp",
        "Custom5PCBTemp1",
        "Custom6PCBTemp2",
        "Custom7PIC1Temp",
        "Custom8PIC2Temp",
        "Custom9Driver1Temp",
        "Custom10Driver2Temp",
        "Custom12UncolledLaser1Temp",
        "Custom13UncolledLaser2Temp",
        "Custom14UncolledLaser3Temp",
        "Custom15UncolledLaser4Temp",
    ]
    for address, name in zip(range(153, 167), temp_names):
        rows.append(field(address, address, "7-0", name, "RO", "Google custom component temperature, units of 1 C"))
    rows.append(field(167, 170, "7-0", "Reserved[4]", "RO", "Reserved for future custom parameters"))
    rows.append(field(171, 255, "7-0", "ReservedForFutureUse[85]", "RO", "For future use"))
    return split_reserved_custom_ranges(rows)


def field(start: int, end: int, bits: str, name: str, type_text: str, description: str) -> dict[str, object]:
    return {"start": start, "end": end, "bits": bits, "field": name, "type": type_text, "description": description}


def split_reserved_custom_ranges(definitions: list[dict[str, object]]) -> list[dict[str, object]]:
    rows: list[dict[str, object]] = []
    for definition in definitions:
        start = int(definition["start"])
        end = int(definition.get("end", start))
        if end <= start or not is_reserved_or_custom_section(definition):
            rows.append(definition)
            continue
        for address in range(start, end + 1):
            split_row = dict(definition)
            split_row["start"] = address
            split_row["end"] = address
            split_row["field"] = byte_section_name(str(definition["field"]), address)
            rows.append(split_row)
    return rows


def is_reserved_or_custom_section(definition: dict[str, object]) -> bool:
    field_name = str(definition.get("field", "")).strip().upper()
    type_text = str(definition.get("type", "")).strip().upper()
    description = str(definition.get("description", "")).strip().upper()
    if field_name.startswith("RESERVED") or type_text == "RESERVED":
        return True
    if field_name in {"CUSTOM"} or re.fullmatch(r"CUSTOM\[\d+\]", field_name):
        return True
    if field_name.startswith("GOOGLECUSTOMRESERVED"):
        return True
    return "RESERVED BYTES" in description or "CUSTOM BYTES" in description


def byte_section_name(field_name: str, address: int) -> str:
    prefix = re.sub(r"\[\d+\]", "", field_name).strip() or "Byte"
    if prefix.upper() == "RESERVED":
        return f"ReservedByte{address:02X}h"
    if prefix.upper() == "CUSTOM":
        return f"CustomByte{address:02X}h"
    return f"{prefix}Byte{address:02X}h"


def bit_lane_rows(address: int, base_name: str, type_text: str, description: str) -> list[dict[str, object]]:
    return [field(address, address, str(bit), f"{base_name}{lane}", type_text, description) for bit, lane in zip(range(7, -1, -1), range(8, 0, -1))]


def two_bit_lane_rows(address: int, base_name: str, first_lane: int, last_lane: int, type_text: str, description: str) -> list[dict[str, object]]:
    ranges = ["1-0", "3-2", "5-4", "7-6"]
    return [field(address, address, bits, f"{base_name}{lane}", type_text, description) for bits, lane in zip(ranges, range(first_lane, last_lane + 1))]


def nibble_lane_rows(address: int, base_name: str, first_lane: int, last_lane: int, type_text: str, description: str) -> list[dict[str, object]]:
    return [
        field(address, address, "3-0", f"{base_name}{first_lane}", type_text, description),
        field(address, address, "7-4", f"{base_name}{last_lane}", type_text, description),
    ]


def rx_nibble_target_rows(start_address: int, base_name: str, type_text: str, description: str) -> list[dict[str, object]]:
    rows: list[dict[str, object]] = []
    for offset, first_lane in enumerate(range(1, 8, 2)):
        rows += nibble_lane_rows(start_address + offset, base_name, first_lane, first_lane + 1, type_text, description)
    return rows


def dp_config_rows(start_address: int, prefix: str, type_text: str) -> list[dict[str, object]]:
    rows: list[dict[str, object]] = []
    for lane, address in enumerate(range(start_address, start_address + 8), start=1):
        rows.append(field(address, address, "7-4", f"{prefix}AppSelCodeLane{lane}", type_text, "Application select code"))
        rows.append(field(address, address, "3-1", f"{prefix}DataPathIDLane{lane}", type_text, "Data Path ID"))
        rows.append(field(address, address, "0", f"{prefix}ExplicitControlLane{lane}", type_text, "Explicit signal-integrity control"))
    return rows


def u16_lane_rows(start_address: int, base_name: str, type_text: str, description: str) -> list[dict[str, object]]:
    return [field(start_address + (lane - 1) * 2, start_address + (lane - 1) * 2 + 1, "7-0", f"{base_name}{lane}", type_text, description) for lane in range(1, 9)]


def bit_sort_key(bits: str) -> int:
    match = re.match(r"(\d+)", bits)
    return -int(match.group(1)) if match else -7


def byte_range_label(start: int, end: int) -> str:
    return str(start) if start == end else f"{start}-{end}"


def hex_range_label(start: int, end: int) -> str:
    return f"{start:02X}" if start == end else f"{start:02X}-{end:02X}"


def apply_captured_style(target, style: object) -> None:
    if not style:
        return
    font, fill, border, alignment, number_format, protection = style
    target.font = copy(font)
    target.fill = copy(fill)
    target.border = copy(border)
    target.alignment = copy(alignment)
    target.number_format = number_format
    target.protection = copy(protection)


def generate_basic_workbook_from_dump(dump_path: str, output_path: str) -> int:
    dump = CmisDump.from_file(dump_path)
    sheets = dump.pages.get("default") or next(iter(dump.pages.values()), {})
    workbook_rows: dict[str, list[list[str]]] = {
        "Change history": [
            ["#", "Version", "Date", "Descriptions"],
            ["1", "generated", "", f"Generated from {Path(dump_path).name}. Values are populated from switch EEPROM dump bytes."],
        ]
    }

    for page, data in sorted(sheets.items(), key=lambda item: _page_sort_key(item[0])):
        sheet_name = _sheet_name(page)
        rows = [LOWER_HEADERS if page == "lower" else UPPER_HEADERS]
        for address, byte_value in enumerate(bytes(data)):
            if page != "lower" and address < 128:
                continue
            if page != "lower" and address >= 256:
                break
            if page == "lower":
                spec_value = f"{byte_value:02X}h"
                rows.append(
                    [
                        str(address),
                        f"{address:02X}",
                        "7-0",
                        f"RawByte{address:02X}h",
                        "Hex",
                        spec_value,
                        interpreted_value_for_range(data, address, address, "7-0", f"RawByte{address:02X}h", "Hex"),
                        f"EEPROM dump byte {address:02X}h from {_page_label(page)}",
                    ]
                )
            else:
                spec_value = f"{byte_value:02X}h"
                rows.append(
                    [
                        str(address),
                        f"{address:02X}",
                        "1",
                        "7-0",
                        f"RawByte{address:02X}h",
                        "Hex",
                        spec_value,
                        interpreted_value_for_range(data, address, address, "7-0", f"RawByte{address:02X}h", "Hex"),
                        f"EEPROM dump byte {address:02X}h from {_page_label(page)}",
                    ]
                )
        workbook_rows[sheet_name] = rows

    if len(workbook_rows) == 1:
        raise ValueError("No EEPROM page data found in the selected dump file.")

    write_xlsx(output_path, workbook_rows)
    return sum(max(len(rows) - 1, 0) for name, rows in workbook_rows.items() if name != "Change history")


def page_from_template_sheet(sheet_name: str) -> str:
    text = sheet_name.strip().lower()
    if text == "lower memory":
        return "lower"
    match = re.fullmatch(r"page\s+([0-9a-f]{1,2})h", text)
    if match:
        return f"{int(match.group(1), 16):02X}"
    return ""


def header_columns(worksheet) -> dict[str, int]:
    columns: dict[str, int] = {}
    for cell in worksheet[1]:
        text = str(cell.value or "").strip().lower()
        if text:
            columns[text] = cell.column
    return columns


def parse_byte_range(value: object, previous_start: int | None, previous_end: int | None) -> tuple[int | None, int | None]:
    text = "" if value is None else str(value).strip()
    if not text:
        return previous_start, previous_end
    match = re.fullmatch(r"(\d+)(?:\s*-\s*(\d+))?", text)
    if not match:
        return None, None
    start = int(match.group(1))
    end = int(match.group(2)) if match.group(2) else start
    return start, end


def template_value_for_range(page_bytes: object, start: int, end: int, bits: str) -> str:
    if page_bytes is None:
        return "*"
    data = bytes(page_bytes)
    if start >= len(data) or end >= len(data):
        return "*"
    raw = data[start : end + 1]
    bit_range = parse_bits(bits)
    if bit_range and len(raw) == 1:
        high, low = bit_range
        mask = sum(1 << bit for bit in range(low, high + 1))
        width = high - low + 1
        return f"{((raw[0] & mask) >> low):0{width}b}b"
    return " ".join(f"{value:02X}h" for value in raw)


def value_column_text(
    interpreted: str,
    description: str,
    spec_value: str,
    page_bytes: object,
    start: int,
    end: int,
    bits: str,
) -> str:
    if not interpreted:
        return interpreted
    if normalize_compare_text(interpreted) != normalize_compare_text(description):
        return interpreted
    raw_code = raw_code_value(page_bytes, start, end, bits)
    return raw_code or spec_value


def normalize_compare_text(value: object) -> str:
    return re.sub(r"\s+", " ", str(value or "").strip()).casefold()


def raw_code_value(page_bytes: object, start: int, end: int, bits: str) -> str:
    if page_bytes is None:
        return "*"
    data = bytes(page_bytes)
    if start >= len(data) or end >= len(data):
        return "*"
    raw = data[start : end + 1]
    bit_range = parse_bits(bits)
    if bit_range and len(raw) == 1:
        high, low = bit_range
        mask = sum(1 << bit for bit in range(low, high + 1))
        return str((raw[0] & mask) >> low)
    if len(raw) == 1:
        return f"{raw[0]:02X}h ({raw[0]})"
    return " ".join(f"{value:02X}h" for value in raw)


def interpreted_value_for_range(page_bytes: object, start: int, end: int, bits: str, field_name: str, type_text: str) -> str:
    if page_bytes is None:
        return "*"
    data = bytes(page_bytes)
    if start >= len(data) or end >= len(data):
        return "*"
    raw = data[start : end + 1]
    field = str(field_name or "")
    type_label = str(type_text or "")
    upper_field = field.upper()
    upper_type = type_label.upper()

    bit_range = parse_bits(bits)
    if bit_range and len(raw) == 1:
        high, low = bit_range
        mask = sum(1 << bit for bit in range(low, high + 1))
        bit_value = (raw[0] & mask) >> low
        if upper_field == "MODULESTATE":
            return MODULE_STATES.get(bit_value, f"Reserved/unknown state ({bit_value})")
        if upper_field == "MEMORYMODEL":
            return MEMORY_MODELS.get(bit_value, f"Reserved/unknown memory model ({bit_value})")
        if upper_field == "MODULEPOWERCLASS":
            return f"Power Class {bit_value + 1}"
        if upper_field == "OPTICALDETECTORTYPE":
            return OPTICAL_DETECTOR_TYPES.get(bit_value, f"Unknown detector type ({bit_value})")
        if upper_field.startswith("MAXDURATION"):
            return f"code {bit_value} (CMIS duration code)"
        if upper_field.endswith("CONFIGONLY") or upper_field.startswith(("STEPPED", "LOWPWR", "SOFTWARE", "SQUELCH")):
            return bool_value(bit_value)
        return str(bit_value)

    if is_ascii_field(field, type_label, len(raw)):
        return decode_ascii(raw)

    if upper_field == "SFF8024IDENTIFIER" and len(raw) == 1:
        return SFF8024_IDENTIFIERS.get(raw[0], f"Unknown ({raw[0]})")

    if upper_field == "CMISREVISION" and len(raw) == 1:
        return f"{raw[0] >> 4}.{raw[0] & 0x0F}"

    if upper_field == "MODULESTATE" and len(raw) == 1:
        state = (raw[0] >> 1) & 0x07
        return MODULE_STATES.get(state, f"Unknown ({state})")

    if upper_field == "MODULEPOWERCLASS" and len(raw) == 1:
        return f"Power Class {(raw[0] >> 5) + 1}"

    if upper_field == "MAXPOWER" and len(raw) == 1:
        return f"{raw[0] * 0.25:g} W"

    if upper_field == "CONNECTORTYPE" and len(raw) == 1:
        return CONNECTOR_TYPES.get(raw[0], f"Unknown connector type ({raw[0]:02X}h)")

    if upper_field == "MEDIAINTERFACETECHNOLOGY" and len(raw) == 1:
        return MEDIA_INTERFACE_TECHNOLOGIES.get(raw[0], f"Unknown media interface technology ({raw[0]:02X}h)")

    if upper_field == "NOMINALWAVELENGTH" and len(raw) == 2:
        return f"{uint(raw) * 0.05:g} nm"

    if upper_field == "WAVELENGTHTOLERANCE" and len(raw) == 2:
        return f"+/-{uint(raw) * 0.005:g} nm"

    if upper_field in {"MODULETEMPMAX", "MODULETEMPMIN"} and len(raw) == 1:
        return f"{int.from_bytes(raw, 'big', signed=True):.1f} C"

    if upper_field in {"OPERATINGVOLTAGEMAX", "OPERATINGVOLTAGEMIN"} and len(raw) == 1:
        return f"{raw[0] * 0.02:.1f} V"

    if "TEMP" in upper_field and "THRESHOLD" in upper_field and len(raw) == 1:
        return f"{raw[0]:.1f} C"

    if "TEMP" in upper_field and "CUSTOM" in upper_field and len(raw) == 1:
        return f"{raw[0]:.1f} C"

    if ("TEMPMONVALUE" in upper_field or upper_field.startswith("TEMPMON")) and len(raw) == 2:
        return f"{int.from_bytes(raw, 'big', signed=True) / 256:.1f} C"

    if ("VCCMONVOLTAGE" in upper_field or upper_field.startswith("VCCMON")) and len(raw) == 2:
        return f"{int.from_bytes(raw, 'big', signed=False) * 0.0001:.1f} V"

    if ("OPTICALPOWER" in upper_field or "TXLANEPOWER" in upper_field or "POWERBOL" in upper_field) and len(raw) == 2:
        return optical_power_dbm(raw)

    if ("LASERBIAS" in upper_field or "LASERLANEBIAS" in upper_field) and len(raw) == 2:
        return f"{int.from_bytes(raw, 'big', signed=False) * 0.002:g} mA"

    if "AUX" in upper_field and "MONVALUE" in upper_field and len(raw) == 2:
        return percent_monitor_value(raw)

    if ("AUX" in upper_field or "CUSTOMMON" in upper_field or "CUSTOMMONITOR" in upper_field) and len(raw) == 2:
        return percent_monitor_value(raw)

    if "FLOAT32" in upper_type and len(raw) == 4:
        try:
            return f"{struct.unpack('>f', raw)[0]:g}"
        except struct.error:
            return hex_value(raw)

    signed = bool(re.search(r"\bS(?:8|16|32|64)\b|SIGNED", upper_type))
    unsigned = bool(re.search(r"\bU(?:8|16|32|64)\b|UNSIGNED|DECIMAL", upper_type))
    if signed or unsigned:
        return str(int.from_bytes(raw, "big", signed=signed))

    if len(raw) == 1 and is_small_numeric_field(field, type_label):
        return str(raw[0])

    return hex_value(raw)


def uint(raw: bytes) -> int:
    return int.from_bytes(raw, "big", signed=False)


def bool_value(value: int) -> str:
    return "Enabled/true" if value else "Disabled/false"


def percent_monitor_value(raw: bytes) -> str:
    value = int.from_bytes(raw, "big", signed=True)
    return f"{value * 100 / 32767:.1f}%"


def optical_power_dbm(raw: bytes) -> str:
    mw = int.from_bytes(raw, "big", signed=False) * 0.0001
    if mw <= 0:
        return "-inf dBm"
    return f"{10 * math.log10(mw):.1f} dBm"


def is_ascii_field(field_name: str, type_text: str, length: int) -> bool:
    upper_field = str(field_name or "").upper()
    upper_type = str(type_text or "").upper()
    if "ASCII" in upper_type:
        return True
    ascii_markers = [
        "VENDORNAME",
        "VENDORPN",
        "VENDORREV",
        "VENDORSN",
        "DATECODE",
        "CLEICODE",
        "CUSTOMERPN",
        "FWVERSION",
        "FIRMWAREVERSION",
    ]
    return length > 1 and any(marker in upper_field for marker in ascii_markers)


def decode_ascii(raw: bytes) -> str:
    text = raw.decode("ascii", errors="replace").replace("\x00", "").strip()
    return text if text else "(blank ASCII)"


def is_small_numeric_field(field_name: str, type_text: str) -> bool:
    upper_field = str(field_name or "").upper()
    upper_type = str(type_text or "").upper()
    if "HEX" in upper_type:
        return False
    numeric_markers = [
        "REVISION",
        "MODEL",
        "CLASS",
        "CODES",
        "CODE",
        "COUNT",
        "STATUS",
        "STATE",
        "SPEED",
        "VERSION",
        "CONTROL",
        "CONFIG",
        "MAPPING",
        "TARGET",
        "VALUE",
    ]
    return any(marker in upper_field for marker in numeric_markers)


def hex_value(raw: bytes) -> str:
    return " ".join(f"{value:02X}" for value in raw)


SFF8024_IDENTIFIERS = {
    0x18: "QSFP-DD Double Density 8X Pluggable Transceiver",
    0x19: "OSFP 8X Pluggable Transceiver",
    0x1E: "QSFP+ or later with CMIS",
}


MODULE_STATES = {
    0b001: "ModuleLowPwr",
    0b010: "ModulePwrUp",
    0b011: "ModuleReady",
    0b100: "ModulePwrDn",
    0b101: "ModuleFault",
}


MEMORY_MODELS = {
    0: "Paged memory",
    1: "Flat memory",
}


OPTICAL_DETECTOR_TYPES = {
    0: "PIN detector",
    1: "APD detector",
}


CONNECTOR_TYPES = {
    0x00: "Unknown or unspecified",
    0x01: "SC",
    0x02: "Fibre Channel Style 1 copper",
    0x03: "Fibre Channel Style 2 copper",
    0x04: "BNC/TNC",
    0x05: "Fibre Channel coaxial",
    0x06: "FiberJack",
    0x07: "LC",
    0x08: "MT-RJ",
    0x09: "MU",
    0x0A: "SG",
    0x0B: "Optical pigtail",
    0x0C: "MPO 1x12",
    0x0D: "MPO 2x16",
    0x20: "HSSDC II",
    0x21: "Copper pigtail",
    0x22: "RJ45",
    0x23: "No separable connector",
    0x24: "MXC 2x16",
    0x25: "CS optical connector",
    0x26: "SN optical connector",
    0x27: "MPO 2x12",
    0x28: "MPO 1x16",
}


MEDIA_INTERFACE_TECHNOLOGIES = {
    0x00: "850 nm VCSEL",
    0x01: "1310 nm VCSEL",
    0x02: "1550 nm VCSEL",
    0x03: "1310 nm FP laser",
    0x04: "1310 nm DFB laser",
    0x05: "1550 nm DFB laser",
    0x06: "1310 nm EML",
    0x07: "1550 nm EML",
    0x08: "Copper cable, unequalized",
    0x09: "Copper cable, passive equalized",
    0x0A: "Copper cable, near/far-end limiting active equalizers",
    0x0B: "Copper cable, far-end limiting active equalizers",
    0x0C: "Copper cable, near-end limiting active equalizers",
    0x0D: "Copper cable, linear active equalizers",
}


def parse_bits(bits: str) -> tuple[int, int] | None:
    text = str(bits or "").strip().lower()
    if text in {"", "all", "7-0"}:
        return None
    range_match = re.fullmatch(r"(\d+)\s*-\s*(\d+)", text)
    if range_match:
        high = int(range_match.group(1))
        low = int(range_match.group(2))
        if high > 7 or low < 0:
            return None
        return (high, low) if high >= low else (low, high)
    if re.fullmatch(r"\d+", text):
        bit = int(text)
        if bit > 7:
            return None
        return bit, bit
    return None


def write_xlsx(path: str, sheets: dict[str, list[list[str]]]) -> None:
    output = Path(path)
    output.parent.mkdir(parents=True, exist_ok=True)
    sheet_items = list(sheets.items())

    with zipfile.ZipFile(output, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        archive.writestr("[Content_Types].xml", _content_types_xml(len(sheet_items)))
        archive.writestr("_rels/.rels", _root_rels_xml())
        archive.writestr("xl/workbook.xml", _workbook_xml(sheet_items))
        archive.writestr("xl/_rels/workbook.xml.rels", _workbook_rels_xml(len(sheet_items)))
        archive.writestr("xl/styles.xml", _styles_xml())
        for index, (_name, rows) in enumerate(sheet_items, start=1):
            archive.writestr(f"xl/worksheets/sheet{index}.xml", _worksheet_xml(rows))


def _worksheet_xml(rows: list[list[str]]) -> str:
    row_xml = []
    for row_index, row in enumerate(rows, start=1):
        cells = []
        for col_index, value in enumerate(row, start=1):
            cell_ref = f"{_column_name(col_index)}{row_index}"
            text = escape(str(value))
            style = ' s="1"' if row_index == 1 else ""
            cells.append(f'<c r="{cell_ref}"{style} t="inlineStr"><is><t>{text}</t></is></c>')
        row_xml.append(f'<row r="{row_index}">{"".join(cells)}</row>')
    last_column = _column_name(max(len(row) for row in rows))
    last_row = len(rows)
    widths = _column_widths(max(len(row) for row in rows))
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        '<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>'
        + widths
        + '<sheetData>'
        + "".join(row_xml)
        + f'</sheetData><autoFilter ref="A1:{last_column}{last_row}"/></worksheet>'
    )


def _workbook_xml(sheet_items: list[tuple[str, list[list[str]]]]) -> str:
    sheets_xml = []
    for index, (name, _rows) in enumerate(sheet_items, start=1):
        safe_name = escape(name[:31])
        sheets_xml.append(f'<sheet name="{safe_name}" sheetId="{index}" r:id="rId{index}"/>')
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        "<sheets>"
        + "".join(sheets_xml)
        + "</sheets></workbook>"
    )


def _workbook_rels_xml(sheet_count: int) -> str:
    rels = []
    for index in range(1, sheet_count + 1):
        rels.append(
            f'<Relationship Id="rId{index}" '
            'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" '
            f'Target="worksheets/sheet{index}.xml"/>'
        )
    rels.append(
        f'<Relationship Id="rId{sheet_count + 1}" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" '
        'Target="styles.xml"/>'
    )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        + "".join(rels)
        + "</Relationships>"
    )


def _content_types_xml(sheet_count: int) -> str:
    overrides = [
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>',
        '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>',
    ]
    for index in range(1, sheet_count + 1):
        overrides.append(
            f'<Override PartName="/xl/worksheets/sheet{index}.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        + "".join(overrides)
        + "</Types>"
    )


def _root_rels_xml() -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
        'Target="xl/workbook.xml"/>'
        "</Relationships>"
    )


def _styles_xml() -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        '<fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><name val="Calibri"/></font></fonts>'
        '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FFD9EAF7"/><bgColor indexed="64"/></patternFill></fill></fills>'
        '<borders count="2"><border/><border><left style="thin"/><right style="thin"/><top style="thin"/><bottom style="thin"/></border></borders>'
        '<cellStyleXfs count="1"><xf/></cellStyleXfs>'
        '<cellXfs count="2"><xf xfId="0"/><xf xfId="0" fontId="1" fillId="1" borderId="1" applyFont="1" applyFill="1" applyBorder="1"/></cellXfs>'
        "</styleSheet>"
    )


def _column_widths(column_count: int) -> str:
    defaults = [12, 10, 12, 10, 30, 14, 16, 60]
    cols = []
    for index in range(1, column_count + 1):
        width = defaults[index - 1] if index <= len(defaults) else 18
        cols.append(f'<col min="{index}" max="{index}" width="{width}" customWidth="1"/>')
    return "<cols>" + "".join(cols) + "</cols>"


def _column_name(index: int) -> str:
    name = ""
    while index:
        index, remainder = divmod(index - 1, 26)
        name = chr(65 + remainder) + name
    return name


def _page_sort_key(page: str) -> tuple[int, int]:
    if page == "lower":
        return (0, 0)
    try:
        return (1, int(page, 16))
    except ValueError:
        return (2, 0)


def _sheet_name(page: str) -> str:
    return "Lower Memory" if page == "lower" else f"Page {page.upper()}h"


def _page_label(page: str) -> str:
    return "Lower Memory" if page == "lower" else f"Upper Page {page.upper()}h"
