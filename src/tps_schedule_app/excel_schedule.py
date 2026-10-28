from __future__ import annotations

from collections import OrderedDict
from dataclasses import dataclass
from pathlib import Path
import re
import zipfile
from xml.etree import ElementTree as ET

from .models import DaySchedule, PersonSchedule, ScheduleEvent, WorkbookSchedule


XML_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS = {"a": XML_NS, "r": REL_NS}
CELL_REF_RE = re.compile(r"([A-Z]+)(\d+)")
TIME_RE = re.compile(r"(\d{1,2})(?::(\d{2}))?\s*(am|pm)\b", re.IGNORECASE)


@dataclass(frozen=True)
class CellValue:
    column_index: int
    column_name: str
    value: str
    style_id: int
    fill_id: int


@dataclass(frozen=True)
class SheetSource:
    name: str
    xml_path: str


def load_workbook_schedule(workbook_path: str | Path) -> WorkbookSchedule:
    workbook = Path(workbook_path)
    with zipfile.ZipFile(workbook) as archive:
        shared_strings = _load_shared_strings(archive)
        fill_ids = _load_fill_ids(archive)
        sheet_sources = _load_sheet_sources(archive)

        people_by_name: OrderedDict[str, list[DaySchedule]] = OrderedDict()
        for sheet in sheet_sources:
            day_schedules = _load_sheet_people(
                archive=archive,
                sheet=sheet,
                shared_strings=shared_strings,
                fill_ids=fill_ids,
            )
            for name, events in day_schedules.items():
                people_by_name.setdefault(name, []).append(
                    DaySchedule(sheet_name=sheet.name, events=events)
                )

    people = [
        PersonSchedule(name=name, days=days)
        for name, days in people_by_name.items()
    ]
    return WorkbookSchedule(
        sheet_names=[sheet.name for sheet in sheet_sources],
        people=people,
    )


def _load_sheet_people(
    archive: zipfile.ZipFile,
    sheet: SheetSource,
    shared_strings: list[str],
    fill_ids: list[int],
) -> OrderedDict[str, list[ScheduleEvent]]:
    root = ET.fromstring(archive.read(sheet.xml_path))
    sheet_data = root.find("a:sheetData", NS)
    rows = list(sheet_data) if sheet_data is not None else []
    if not rows:
        return OrderedDict()

    header_times = _load_header_times(rows[0], shared_strings)
    header_columns = sorted(header_times)
    schedules: OrderedDict[str, list[ScheduleEvent]] = OrderedDict()

    for row in rows[1:]:
        row_cells = _load_row_cells(row, shared_strings, fill_ids)
        full_name = _build_full_name(row_cells)
        if not full_name:
            continue

        events = _extract_events(
            row_cells=row_cells,
            header_times=header_times,
            header_columns=header_columns,
            sheet_name=sheet.name,
        )
        schedules[full_name] = events

    return schedules


def _load_shared_strings(archive: zipfile.ZipFile) -> list[str]:
    try:
        root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
    except KeyError:
        return []
    return [_read_inline_text(item) for item in root]


def _load_fill_ids(archive: zipfile.ZipFile) -> list[int]:
    styles_root = ET.fromstring(archive.read("xl/styles.xml"))
    cell_xfs = styles_root.find("a:cellXfs", NS)
    if cell_xfs is None:
        return [0]

    fill_ids = []
    for xf in cell_xfs:
        fill_ids.append(int(xf.attrib.get("fillId", "0")))
    return fill_ids or [0]


def _load_sheet_sources(archive: zipfile.ZipFile) -> list[SheetSource]:
    workbook_root = ET.fromstring(archive.read("xl/workbook.xml"))
    rel_root = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))

    rel_map = {
        relationship.attrib["Id"]: relationship.attrib["Target"]
        for relationship in rel_root
    }

    sheet_sources = []
    sheets = workbook_root.find("a:sheets", NS)
    if sheets is None:
        return sheet_sources

    for sheet in sheets:
        rel_id = sheet.attrib[f"{{{REL_NS}}}id"]
        target = rel_map[rel_id].replace("\\", "/")
        if not target.startswith("worksheets/"):
            target = f"worksheets/{target.split('/')[-1]}"
        sheet_sources.append(
            SheetSource(
                name=sheet.attrib["name"],
                xml_path=f"xl/{target}",
            )
        )

    return sheet_sources


def _load_header_times(row: ET.Element, shared_strings: list[str]) -> dict[int, int]:
    header_times: dict[int, int] = {}
    for cell in row:
        column_name, _ = _split_cell_reference(cell.attrib["r"])
        if column_name == "A":
            continue
        value = _cell_text(cell, shared_strings)
        if not value:
            continue
        header_times[_column_to_index(column_name)] = _parse_time_value(value)
    return header_times


def _load_row_cells(
    row: ET.Element,
    shared_strings: list[str],
    fill_ids: list[int],
) -> dict[int, CellValue]:
    row_cells: dict[int, CellValue] = {}
    for cell in row:
        column_name, _ = _split_cell_reference(cell.attrib["r"])
        column_index = _column_to_index(column_name)
        style_id = int(cell.attrib.get("s", "0"))
        value = _cell_text(cell, shared_strings)
        fill_id = fill_ids[style_id] if style_id < len(fill_ids) else 0
        row_cells[column_index] = CellValue(
            column_index=column_index,
            column_name=column_name,
            value=value,
            style_id=style_id,
            fill_id=fill_id,
        )
    return row_cells


def _build_full_name(row_cells: dict[int, CellValue]) -> str:
    first_name = row_cells.get(1)
    last_name = row_cells.get(2)
    if first_name is None or last_name is None:
        return ""
    if not first_name.value or not last_name.value:
        return ""
    return f"{first_name.value} {last_name.value}".strip()


def _extract_events(
    row_cells: dict[int, CellValue],
    header_times: dict[int, int],
    header_columns: list[int],
    sheet_name: str,
) -> list[ScheduleEvent]:
    events = []
    for cell in sorted(row_cells.values(), key=lambda item: item.column_index):
        if cell.column_index <= 2 or not cell.value or cell.fill_id <= 0:
            continue

        block_end = cell.column_index
        candidate = cell.column_index + 1
        while candidate in header_times:
            next_cell = row_cells.get(candidate)
            if next_cell is None or next_cell.fill_id != cell.fill_id or next_cell.value:
                break
            block_end = candidate
            candidate += 1

        end_minutes = None
        if block_end > cell.column_index:
            next_header_column = _next_header_column(header_columns, block_end)
            if next_header_column is not None:
                end_minutes = header_times[next_header_column]
            else:
                end_minutes = _extrapolate_end_time(header_columns, header_times, block_end)

        events.append(
            ScheduleEvent(
                description=cell.value,
                start_minutes=header_times[cell.column_index],
                end_minutes=end_minutes,
                sheet_name=sheet_name,
                start_column=cell.column_name,
                fill_id=cell.fill_id,
            )
        )

    return events


def _next_header_column(header_columns: list[int], current_column: int) -> int | None:
    for column in header_columns:
        if column > current_column:
            return column
    return None


def _extrapolate_end_time(
    header_columns: list[int],
    header_times: dict[int, int],
    current_column: int,
) -> int:
    if len(header_columns) < 2:
        return header_times[current_column]

    current_position = header_columns.index(current_column)
    if current_position == 0:
        step = 15
    else:
        previous_column = header_columns[current_position - 1]
        step = (header_times[current_column] - header_times[previous_column]) % (24 * 60)
        if step <= 0:
            step = 15

    return (header_times[current_column] + step) % (24 * 60)


def _cell_text(cell: ET.Element, shared_strings: list[str]) -> str:
    cell_type = cell.attrib.get("t")
    if cell_type == "inlineStr":
        return _normalize_text(_read_inline_text(cell))

    value_node = cell.find("a:v", NS)
    if value_node is None or value_node.text is None:
        return ""

    if cell_type == "s":
        index = int(value_node.text)
        return _normalize_text(shared_strings[index])

    return _normalize_text(value_node.text)


def _read_inline_text(node: ET.Element) -> str:
    return "".join(
        text_node.text or ""
        for text_node in node.iter(f"{{{XML_NS}}}t")
    )


def _normalize_text(value: str) -> str:
    return re.sub(r"\s+", " ", value).strip()


def _split_cell_reference(reference: str) -> tuple[str, int]:
    match = CELL_REF_RE.fullmatch(reference)
    if not match:
        raise ValueError(f"Unsupported cell reference: {reference}")
    return match.group(1), int(match.group(2))


def _column_to_index(column_name: str) -> int:
    result = 0
    for character in column_name:
        result = result * 26 + ord(character) - 64
    return result


def _parse_time_value(value: str) -> int:
    match = TIME_RE.search(value)
    if match is None:
        raise ValueError(f"Unable to parse time from header value: {value!r}")

    hour = int(match.group(1))
    minute = int(match.group(2) or "0")
    suffix = match.group(3).lower()

    if hour == 12:
        hour = 0
    if suffix == "pm":
        hour += 12

    return hour * 60 + minute
