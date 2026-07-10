from dataclasses import dataclass
from html import escape
from io import BytesIO
from pathlib import Path
from threading import Event
import posixpath
import re
import zipfile
from xml.etree import ElementTree

from excel_common import ExcelReadError, iter_excel_files


MAIN_NAMESPACE = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
OFFICE_RELATIONSHIP_NAMESPACE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
WORKBOOK_PART_PATH = "xl/workbook.xml"
WORKBOOK_RELS_PART_PATH = "xl/_rels/workbook.xml.rels"
SHARED_STRINGS_PART_PATH = "xl/sharedStrings.xml"
CELL_REFERENCE_PATTERN = re.compile(r"^([A-Za-z]+)(\d+)$")
SHARED_VALUE_PATTERN = re.compile(rb"<v>\s*(\d+)\s*</v>")
ROW_TAG = f"{{{MAIN_NAMESPACE}}}row"
CELL_TAG = f"{{{MAIN_NAMESPACE}}}c"
TEXT_TAG = f"{{{MAIN_NAMESPACE}}}t"
SHARED_STRING_ITEM_TAG = f"{{{MAIN_NAMESPACE}}}si"
BOOLEAN_SEARCH_KEYWORDS = {"true", "false"}
ROW_FILTER_MODE_NONE = "无"
ROW_FILTER_MODE_DISABLED = "禁用"
ROW_FILTER_MODE_RANGE = "区间"
ROW_FILTER_MODES = (ROW_FILTER_MODE_NONE, ROW_FILTER_MODE_DISABLED, ROW_FILTER_MODE_RANGE)


@dataclass(frozen=True)
class ExcelCellMatch:
    path: Path
    workbook_name: str
    sheet_name: str
    row_index: int
    column_index: int
    value: str


@dataclass(frozen=True)
class DataFilterConfig:
    enable_header_filter: bool = False
    header_row_count: int = 1
    enable_column_filter: bool = False
    column_marker_row_index: int = 1
    disabled_column_marker: str = "$"
    enable_disabled_row_filter: bool = False
    disabled_row_marker_column_index: int = 1
    disabled_row_contains: str = ""
    enable_range_row_filter: bool = False
    range_row_marker_column_index: int = 1
    range_start_text: str = ""
    range_end_text: str = ""


@dataclass(frozen=True)
class DataFilterState:
    header_row_count: int = 0
    disabled_columns: frozenset[int] = frozenset()
    disabled_rows: frozenset[int] = frozenset()
    enabled_rows: frozenset[int] | None = None


DEFAULT_DATA_FILTER_CONFIG = DataFilterConfig()


def search_excel_cells(
    folder_path: str | Path,
    keyword: str,
    disabled_sheet_marker: str | None = None,
    data_filter_config: DataFilterConfig | None = None,
    cancel_event: Event | None = None,
) -> list[ExcelCellMatch]:
    normalized_keyword = keyword.strip().lower()
    if not normalized_keyword:
        raise ExcelReadError("Search keyword cannot be empty")

    matches: list[ExcelCellMatch] = []
    for path in iter_excel_files(folder_path):
        if _is_cancelled(cancel_event):
            break
        matches.extend(
            _search_excel_file_cells(
                path,
                normalized_keyword,
                disabled_sheet_marker,
                data_filter_config,
                cancel_event,
            )
        )

    return matches


def _search_excel_file_cells(
    path: Path,
    keyword: str,
    disabled_sheet_marker: str | None,
    data_filter_config: DataFilterConfig | None,
    cancel_event: Event | None,
) -> list[ExcelCellMatch]:
    suffix = path.suffix.lower()
    if suffix in {".xlsx", ".xlsm", ".xltx", ".xltm"}:
        return _search_xlsx_xml_cells(path, keyword, disabled_sheet_marker, data_filter_config, cancel_event)
    if suffix == ".xls":
        return _search_xlrd_cells(path, keyword, disabled_sheet_marker, data_filter_config, cancel_event)
    if suffix == ".xlsb":
        return _search_xlsb_cells(path, keyword, disabled_sheet_marker, data_filter_config, cancel_event)
    return []


def _search_xlsx_xml_cells(
    path: Path,
    keyword: str,
    disabled_sheet_marker: str | None,
    data_filter_config: DataFilterConfig | None,
    cancel_event: Event | None,
) -> list[ExcelCellMatch]:
    try:
        with zipfile.ZipFile(path) as archive:
            if _is_cancelled(cancel_event):
                return []

            shared_strings = _read_shared_strings(archive, cancel_event) if data_filter_config is not None else None
            matching_shared_strings = (
                _filter_matching_shared_strings(shared_strings, keyword, cancel_event)
                if shared_strings is not None
                else _read_matching_shared_strings(archive, keyword, cancel_event)
            )
            if _is_cancelled(cancel_event):
                return []

            matching_shared_string_ids = {
                str(shared_string_index).encode("ascii")
                for shared_string_index in matching_shared_strings
            }
            sheets = _read_workbook_sheets(archive)
            matches: list[ExcelCellMatch] = []
            for sheet_name, sheet_xml_path in sheets:
                if _is_cancelled(cancel_event):
                    break
                if not _should_search_sheet(sheet_name, disabled_sheet_marker):
                    continue
                matches.extend(
                    _search_sheet_xml_cells(
                        archive,
                        path,
                        sheet_name,
                        sheet_xml_path,
                        matching_shared_strings,
                        matching_shared_string_ids,
                        keyword,
                        shared_strings,
                        data_filter_config,
                        cancel_event,
                    )
                )
            return matches
    except (OSError, KeyError, zipfile.BadZipFile, ElementTree.ParseError) as exc:
        raise ExcelReadError(f"Could not open Excel file: {path}") from exc


def _read_workbook_sheets(archive: zipfile.ZipFile) -> list[tuple[str, str]]:
    workbook_root = ElementTree.fromstring(archive.read(WORKBOOK_PART_PATH))
    relationships = _read_workbook_relationships(archive)
    sheets: list[tuple[str, str]] = []

    for sheet_element in workbook_root.iter(_xml_tag("sheet")):
        sheet_name = sheet_element.attrib.get("name", "")
        relationship_id = sheet_element.attrib.get(f"{{{OFFICE_RELATIONSHIP_NAMESPACE}}}id", "")
        target = relationships.get(relationship_id)
        if sheet_name and target:
            sheets.append((sheet_name, target))

    return sheets


def _read_workbook_relationships(archive: zipfile.ZipFile) -> dict[str, str]:
    relationships_root = ElementTree.fromstring(archive.read(WORKBOOK_RELS_PART_PATH))
    relationships: dict[str, str] = {}
    for relationship in relationships_root:
        if _local_name(relationship.tag) != "Relationship":
            continue
        if relationship.attrib.get("TargetMode") == "External":
            continue

        relationship_id = relationship.attrib.get("Id", "")
        target = relationship.attrib.get("Target", "")
        if relationship_id and target:
            relationships[relationship_id] = _resolve_part_path(WORKBOOK_PART_PATH, target)
    return relationships


def _read_shared_strings(archive: zipfile.ZipFile, cancel_event: Event | None = None) -> list[str]:
    try:
        shared_strings_xml = archive.read(SHARED_STRINGS_PART_PATH)
    except KeyError:
        return []

    shared_strings: list[str] = []
    with BytesIO(shared_strings_xml) as stream:
        for _event, element in ElementTree.iterparse(stream, events=("end",)):
            if _is_cancelled(cancel_event):
                break
            if element.tag != SHARED_STRING_ITEM_TAG and _local_name(element.tag) != "si":
                continue
            shared_strings.append(_rich_text(element))
            element.clear()
    return shared_strings


def _filter_matching_shared_strings(
    shared_strings: list[str],
    keyword: str,
    cancel_event: Event | None = None,
) -> dict[int, str]:
    matches: dict[int, str] = {}
    for shared_string_index, text in enumerate(shared_strings):
        if _is_cancelled(cancel_event):
            break
        if _get_matched_text(text, keyword) is not None:
            matches[shared_string_index] = text
    return matches


def _read_matching_shared_strings(
    archive: zipfile.ZipFile,
    keyword: str,
    cancel_event: Event | None = None,
) -> dict[int, str]:
    try:
        shared_strings_xml = archive.read(SHARED_STRINGS_PART_PATH)
    except KeyError:
        return {}

    if _is_cancelled(cancel_event):
        return {}

    normalized_shared_strings_xml = shared_strings_xml.lower()
    if not _has_direct_keyword_hint(normalized_shared_strings_xml, keyword) and b"<r>" not in normalized_shared_strings_xml:
        return {}

    matching_shared_strings: dict[int, str] = {}
    shared_string_index = 0
    with BytesIO(shared_strings_xml) as stream:
        for _event, element in ElementTree.iterparse(stream, events=("end",)):
            if _is_cancelled(cancel_event):
                break
            if element.tag != SHARED_STRING_ITEM_TAG and _local_name(element.tag) != "si":
                continue

            text = _rich_text(element)
            if _get_matched_text(text, keyword) is not None:
                matching_shared_strings[shared_string_index] = text
            shared_string_index += 1
            element.clear()
    return matching_shared_strings


def _search_sheet_xml_cells(
    archive: zipfile.ZipFile,
    path: Path,
    sheet_name: str,
    sheet_xml_path: str,
    matching_shared_strings: dict[int, str],
    matching_shared_string_ids: set[bytes],
    keyword: str,
    shared_strings: list[str] | None,
    data_filter_config: DataFilterConfig | None,
    cancel_event: Event | None,
) -> list[ExcelCellMatch]:
    matches: list[ExcelCellMatch] = []
    current_row_index = 0
    current_column_index = 0
    sheet_xml = _read_sheet_xml_if_search_hint(
        archive,
        sheet_xml_path,
        matching_shared_string_ids,
        keyword,
        cancel_event,
    )
    if sheet_xml is None:
        return matches

    if _is_cancelled(cancel_event):
        return matches

    data_filter_state = _build_xlsx_data_filter_state(sheet_xml, shared_strings or [], data_filter_config, cancel_event)
    with BytesIO(sheet_xml) as stream:
        for event, element in ElementTree.iterparse(stream, events=("start", "end")):
            if _is_cancelled(cancel_event):
                break
            if event == "start" and (element.tag == ROW_TAG or _local_name(element.tag) == "row"):
                current_row_index = _parse_int(element.attrib.get("r"), current_row_index + 1)
                current_column_index = 0
                continue
            if event != "end" or (element.tag != CELL_TAG and _local_name(element.tag) != "c"):
                continue

            current_column_index += 1
            cell_reference = element.attrib.get("r", "")
            row_index, column_index = _get_cell_position(
                cell_reference,
                current_row_index,
                current_column_index,
            )
            if not _is_enabled_cell(row_index, column_index, data_filter_state):
                element.clear()
                continue
            matched_value = _get_xml_cell_match(element, matching_shared_strings, keyword)
            if matched_value is not None:
                matches.append(
                    ExcelCellMatch(
                        path=path,
                        workbook_name=path.name,
                        sheet_name=sheet_name,
                        row_index=row_index,
                        column_index=column_index,
                        value=matched_value,
                    )
                )
            element.clear()
    return matches


def _read_sheet_xml_if_search_hint(
    archive: zipfile.ZipFile,
    sheet_xml_path: str,
    matching_shared_string_ids: set[bytes],
    keyword: str,
    cancel_event: Event | None = None,
) -> bytes | None:
    if _is_cancelled(cancel_event):
        return None

    sheet_xml = archive.read(sheet_xml_path)
    if _is_cancelled(cancel_event):
        return None

    normalized_sheet_xml = sheet_xml.lower()

    if _sheet_xml_has_search_hint(normalized_sheet_xml, matching_shared_string_ids, keyword):
        return sheet_xml
    return None


def _sheet_xml_has_search_hint(
    normalized_sheet_xml: bytes,
    matching_shared_string_ids: set[bytes],
    keyword: str,
) -> bool:
    if matching_shared_string_ids and _has_matching_shared_string_reference(
        normalized_sheet_xml,
        matching_shared_string_ids,
    ):
        return True

    if _has_direct_keyword_hint(normalized_sheet_xml, keyword):
        return True

    if keyword in BOOLEAN_SEARCH_KEYWORDS and b't="b"' in normalized_sheet_xml:
        return True

    return b't="inlinestr"' in normalized_sheet_xml


def _build_xlsx_data_filter_state(
    sheet_xml: bytes,
    shared_strings: list[str],
    data_filter_config: DataFilterConfig | None,
    cancel_event: Event | None = None,
) -> DataFilterState | None:
    if data_filter_config is None:
        return None

    current_row_index = 0
    current_column_index = 0
    disabled_row_marker_values: dict[int, str] = {}
    range_row_marker_values: dict[int, str] = {}
    marker_row_values: dict[int, str] = {}
    available_row_indices: set[int] = set()

    with BytesIO(sheet_xml) as stream:
        for event, element in ElementTree.iterparse(stream, events=("start", "end")):
            if _is_cancelled(cancel_event):
                return None
            if event == "start" and (element.tag == ROW_TAG or _local_name(element.tag) == "row"):
                current_row_index = _parse_int(element.attrib.get("r"), current_row_index + 1)
                current_column_index = 0
                continue
            if event != "end" or (element.tag != CELL_TAG and _local_name(element.tag) != "c"):
                continue

            current_column_index += 1
            row_index, column_index = _get_cell_position(
                element.attrib.get("r", ""),
                current_row_index,
                current_column_index,
            )
            available_row_indices.add(row_index)
            text = _clean_filter_text(_get_xml_cell_text(element, shared_strings))
            if (
                data_filter_config.enable_disabled_row_filter
                and column_index == data_filter_config.disabled_row_marker_column_index
            ):
                disabled_row_marker_values[row_index] = text
            if (
                data_filter_config.enable_range_row_filter
                and column_index == data_filter_config.range_row_marker_column_index
            ):
                range_row_marker_values[row_index] = text
            if data_filter_config.enable_column_filter and row_index == data_filter_config.column_marker_row_index:
                marker_row_values[column_index] = text
            element.clear()

    return _build_data_filter_state(
        disabled_row_marker_values,
        range_row_marker_values,
        marker_row_values,
        data_filter_config,
        available_row_indices,
    )


def _build_data_filter_state(
    disabled_row_marker_values: dict[int, str],
    range_row_marker_values: dict[int, str],
    marker_row_values: dict[int, str],
    data_filter_config: DataFilterConfig | None,
    available_row_indices: set[int] | None = None,
) -> DataFilterState | None:
    if data_filter_config is None:
        return None

    header_row_count = data_filter_config.header_row_count if data_filter_config.enable_header_filter else 0
    marker = data_filter_config.disabled_column_marker.strip()
    disabled_columns = frozenset(
        column_index
        for column_index, text in marker_row_values.items()
        if data_filter_config.enable_column_filter and marker and text.startswith(marker)
    )

    disabled_rows: frozenset[int] = frozenset()
    enabled_rows: frozenset[int] | None = None

    if data_filter_config.enable_disabled_row_filter:
        disabled_text = data_filter_config.disabled_row_contains.strip()
        disabled_rows = frozenset(
            row_index
            for row_index, text in disabled_row_marker_values.items()
            if disabled_text and disabled_text in text
        )
    if data_filter_config.enable_range_row_filter:
        enabled_rows = _get_enabled_range_rows(
            range_row_marker_values,
            data_filter_config.range_start_text.strip(),
            data_filter_config.range_end_text.strip(),
            available_row_indices or set(range_row_marker_values),
        )

    return DataFilterState(header_row_count, disabled_columns, disabled_rows, enabled_rows)


def _get_enabled_range_rows(
    row_marker_values: dict[int, str],
    range_start_text: str,
    range_end_text: str,
    available_row_indices: set[int],
) -> frozenset[int]:
    if not range_start_text or not range_end_text:
        return frozenset()

    enabled_rows: set[int] = set()
    range_start_row: int | None = None
    sorted_available_rows = sorted(available_row_indices)
    for row_index, text in sorted(row_marker_values.items()):
        if range_start_text in text:
            range_start_row = row_index
        if range_start_row is not None and range_end_text in text:
            enabled_rows.update(
                available_row
                for available_row in sorted_available_rows
                if range_start_row <= available_row <= row_index
            )
            range_start_row = None

    if range_start_row is not None:
        enabled_rows.update(
            available_row
            for available_row in sorted_available_rows
            if available_row >= range_start_row
        )
    return frozenset(enabled_rows)


def _is_enabled_cell(row_index: int, column_index: int, data_filter_state: DataFilterState | None) -> bool:
    if data_filter_state is None:
        return True
    if data_filter_state.header_row_count > 0 and row_index <= data_filter_state.header_row_count:
        return False
    if column_index in data_filter_state.disabled_columns:
        return False
    if row_index in data_filter_state.disabled_rows:
        return False
    if data_filter_state.enabled_rows is not None and row_index not in data_filter_state.enabled_rows:
        return False
    return True


def _has_matching_shared_string_reference(
    normalized_sheet_xml: bytes,
    matching_shared_string_ids: set[bytes],
) -> bool:
    for match in SHARED_VALUE_PATTERN.finditer(normalized_sheet_xml):
        if match.group(1) in matching_shared_string_ids:
            return True
    return False


def _has_direct_keyword_hint(normalized_sheet_xml: bytes, keyword: str) -> bool:
    for hint in _direct_keyword_hints(keyword):
        if hint and hint in normalized_sheet_xml:
            return True
    return False


def _direct_keyword_hints(keyword: str) -> set[bytes]:
    hints = {keyword.encode("utf-8")}
    escaped_keyword = escape(keyword, quote=False).lower()
    hints.add(escaped_keyword.encode("utf-8"))
    return hints


def _get_xml_cell_match(
    element: ElementTree.Element,
    matching_shared_strings: dict[int, str],
    keyword: str,
) -> str | None:
    cell_type = element.attrib.get("t", "")
    if cell_type == "s":
        raw_value = _find_child_text(element, "v")
        if raw_value is None:
            return None

        shared_string_index = _parse_int(raw_value, -1)
        return matching_shared_strings.get(shared_string_index)

    if cell_type == "inlineStr":
        return _get_matched_text(_rich_text(element), keyword)

    raw_value = _find_child_text(element, "v")
    if raw_value is None:
        return None

    if cell_type == "b":
        value = "TRUE" if raw_value == "1" else "FALSE" if raw_value == "0" else raw_value
        return _get_matched_text(value, keyword)

    return _get_matched_text(raw_value, keyword)


def _get_xml_cell_text(element: ElementTree.Element, shared_strings: list[str]) -> str:
    cell_type = element.attrib.get("t", "")
    if cell_type == "s":
        raw_value = _find_child_text(element, "v")
        if raw_value is None:
            return ""

        shared_string_index = _parse_int(raw_value, -1)
        if shared_string_index < 0 or shared_string_index >= len(shared_strings):
            return ""
        return shared_strings[shared_string_index]

    if cell_type == "inlineStr":
        return _rich_text(element)

    raw_value = _find_child_text(element, "v")
    if raw_value is None:
        return ""

    if cell_type == "b":
        return "TRUE" if raw_value == "1" else "FALSE" if raw_value == "0" else raw_value
    return raw_value


def _get_cell_position(
    cell_reference: str,
    fallback_row_index: int,
    fallback_column_index: int,
) -> tuple[int, int]:
    match = CELL_REFERENCE_PATTERN.match(cell_reference)
    if match is None:
        return fallback_row_index, fallback_column_index

    column_letters, row_text = match.groups()
    return _parse_int(row_text, fallback_row_index), _column_letters_to_index(column_letters)


def _column_letters_to_index(column_letters: str) -> int:
    column_index = 0
    for char in column_letters.upper():
        column_index = column_index * 26 + ord(char) - ord("A") + 1
    return column_index


def _resolve_part_path(base_part_path: str, target: str) -> str:
    if target.startswith("/"):
        return target.lstrip("/")
    return posixpath.normpath(posixpath.join(posixpath.dirname(base_part_path), target))


def _find_child_text(element: ElementTree.Element, child_name: str) -> str | None:
    for child in element:
        if _local_name(child.tag) == child_name:
            return child.text or ""
    return None


def _rich_text(element: ElementTree.Element) -> str:
    text_parts: list[str] = []
    for text_element in element.iter():
        if (text_element.tag == TEXT_TAG or _local_name(text_element.tag) == "t") and text_element.text:
            text_parts.append(text_element.text)
    return "".join(text_parts)


def _xml_tag(name: str) -> str:
    return f"{{{MAIN_NAMESPACE}}}{name}"


def _local_name(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def _parse_int(value: object, default: int) -> int:
    try:
        return int(str(value))
    except (TypeError, ValueError):
        return default


def _clean_filter_text(value: object) -> str:
    if value is None:
        return ""
    return str(value).strip()


def normalize_row_filter_mode(row_filter_mode: str) -> str:
    return row_filter_mode if row_filter_mode in ROW_FILTER_MODES else ROW_FILTER_MODE_NONE


def validate_data_filter_config(config: DataFilterConfig) -> None:
    if config.enable_header_filter and config.header_row_count < 1:
        raise ExcelReadError("表头长度必须大于等于 1")
    if config.enable_column_filter and config.column_marker_row_index < 1:
        raise ExcelReadError("列未启用判断行必须大于等于 1")
    if config.enable_column_filter and not config.disabled_column_marker.strip():
        raise ExcelReadError("启用列禁用时，必须填写列未启用标识")
    if config.enable_disabled_row_filter and config.disabled_row_marker_column_index < 1:
        raise ExcelReadError("行禁用判断列必须大于等于 1")
    if config.enable_disabled_row_filter and not config.disabled_row_contains.strip():
        raise ExcelReadError("行未启用方式为“禁用”时，必须填写行禁用包含数据")
    if config.enable_range_row_filter:
        if config.range_row_marker_column_index < 1:
            raise ExcelReadError("行区间判断列必须大于等于 1")
        if not config.range_start_text.strip():
            raise ExcelReadError("行未启用方式为“区间”时，必须填写区间起始数据")
        if not config.range_end_text.strip():
            raise ExcelReadError("行未启用方式为“区间”时，必须填写区间终止数据")


def _search_xlrd_cells(
    path: Path,
    keyword: str,
    disabled_sheet_marker: str | None,
    data_filter_config: DataFilterConfig | None,
    cancel_event: Event | None,
) -> list[ExcelCellMatch]:
    try:
        import xlrd
    except ImportError as exc:
        raise ExcelReadError("xlrd is required to read old .xls files") from exc

    try:
        workbook = xlrd.open_workbook(str(path), on_demand=True)
    except Exception as exc:
        raise ExcelReadError(f"Could not open Excel file: {path}") from exc

    matches: list[ExcelCellMatch] = []
    try:
        for sheet_name in workbook.sheet_names():
            if _is_cancelled(cancel_event):
                break
            if not _should_search_sheet(sheet_name, disabled_sheet_marker):
                continue
            worksheet = workbook.sheet_by_name(sheet_name)
            data_filter_state = _build_xlrd_data_filter_state(worksheet, data_filter_config, cancel_event)
            for row_offset in range(worksheet.nrows):
                if _is_cancelled(cancel_event):
                    break
                for column_offset in range(worksheet.ncols):
                    row_index = row_offset + 1
                    column_index = column_offset + 1
                    if not _is_enabled_cell(row_index, column_index, data_filter_state):
                        continue
                    matched_value = _get_matched_text(
                        worksheet.cell_value(row_offset, column_offset),
                        keyword,
                    )
                    if matched_value is None:
                        continue
                    matches.append(
                        ExcelCellMatch(
                            path=path,
                            workbook_name=path.name,
                            sheet_name=sheet_name,
                            row_index=row_index,
                            column_index=column_index,
                            value=matched_value,
                        )
                    )
        return matches
    finally:
        workbook.release_resources()


def _search_xlsb_cells(
    path: Path,
    keyword: str,
    disabled_sheet_marker: str | None,
    data_filter_config: DataFilterConfig | None,
    cancel_event: Event | None,
) -> list[ExcelCellMatch]:
    try:
        from pyxlsb import open_workbook
    except ImportError as exc:
        raise ExcelReadError("pyxlsb is required to read .xlsb files") from exc

    matches: list[ExcelCellMatch] = []
    try:
        with open_workbook(str(path)) as workbook:
            for sheet_name in workbook.sheets:
                if _is_cancelled(cancel_event):
                    break
                if not _should_search_sheet(sheet_name, disabled_sheet_marker):
                    continue
                with workbook.get_sheet(sheet_name) as worksheet:
                    rows: list[list[object]] = []
                    for row in worksheet.rows():
                        if _is_cancelled(cancel_event):
                            break
                        rows.append([cell.v for cell in row])
                    if _is_cancelled(cancel_event):
                        break
                    data_filter_state = _build_tabular_data_filter_state(rows, data_filter_config)
                    for row_index, row in enumerate(rows, start=1):
                        if _is_cancelled(cancel_event):
                            break
                        for column_index, value in enumerate(row, start=1):
                            if not _is_enabled_cell(row_index, column_index, data_filter_state):
                                continue
                            matched_value = _get_matched_text(value, keyword)
                            if matched_value is None:
                                continue
                            matches.append(
                                ExcelCellMatch(
                                    path=path,
                                    workbook_name=path.name,
                                    sheet_name=sheet_name,
                                    row_index=row_index,
                                    column_index=column_index,
                                    value=matched_value,
                                )
                            )
        return matches
    except Exception as exc:
        raise ExcelReadError(f"Could not open Excel file: {path}") from exc


def _build_xlrd_data_filter_state(
    worksheet: object,
    data_filter_config: DataFilterConfig | None,
    cancel_event: Event | None = None,
) -> DataFilterState | None:
    if data_filter_config is None:
        return None

    disabled_row_marker_values: dict[int, str] = {}
    range_row_marker_values: dict[int, str] = {}
    marker_row_values: dict[int, str] = {}
    marker_row_offset = data_filter_config.column_marker_row_index - 1
    disabled_row_marker_column_offset = data_filter_config.disabled_row_marker_column_index - 1
    range_row_marker_column_offset = data_filter_config.range_row_marker_column_index - 1

    if data_filter_config.enable_column_filter and 0 <= marker_row_offset < getattr(worksheet, "nrows", 0):
        for column_offset in range(getattr(worksheet, "ncols", 0)):
            if _is_cancelled(cancel_event):
                return None
            marker_row_values[column_offset + 1] = _clean_filter_text(
                worksheet.cell_value(marker_row_offset, column_offset)
            )

    for row_offset in range(getattr(worksheet, "nrows", 0)):
        if _is_cancelled(cancel_event):
            return None
        if (
            data_filter_config.enable_disabled_row_filter
            and 0 <= disabled_row_marker_column_offset < getattr(worksheet, "ncols", 0)
        ):
            disabled_row_marker_values[row_offset + 1] = _clean_filter_text(
                worksheet.cell_value(row_offset, disabled_row_marker_column_offset)
            )
        if (
            data_filter_config.enable_range_row_filter
            and 0 <= range_row_marker_column_offset < getattr(worksheet, "ncols", 0)
        ):
            range_row_marker_values[row_offset + 1] = _clean_filter_text(
                worksheet.cell_value(row_offset, range_row_marker_column_offset)
            )

    return _build_data_filter_state(
        disabled_row_marker_values,
        range_row_marker_values,
        marker_row_values,
        data_filter_config,
        set(range(1, getattr(worksheet, "nrows", 0) + 1)),
    )


def _build_tabular_data_filter_state(
    rows: list[list[object]],
    data_filter_config: DataFilterConfig | None,
) -> DataFilterState | None:
    if data_filter_config is None:
        return None

    disabled_row_marker_column_index = data_filter_config.disabled_row_marker_column_index
    range_row_marker_column_index = data_filter_config.range_row_marker_column_index
    disabled_row_marker_values = {
        row_index: _clean_filter_text(row[disabled_row_marker_column_index - 1])
        for row_index, row in enumerate(rows, start=1)
        if data_filter_config.enable_disabled_row_filter
        and disabled_row_marker_column_index >= 1
        and len(row) >= disabled_row_marker_column_index
    }
    range_row_marker_values = {
        row_index: _clean_filter_text(row[range_row_marker_column_index - 1])
        for row_index, row in enumerate(rows, start=1)
        if data_filter_config.enable_range_row_filter
        and range_row_marker_column_index >= 1
        and len(row) >= range_row_marker_column_index
    }

    marker_row_values: dict[int, str] = {}
    marker_row_index = data_filter_config.column_marker_row_index
    if data_filter_config.enable_column_filter and 1 <= marker_row_index <= len(rows):
        marker_row_values = {
            column_index: _clean_filter_text(value)
            for column_index, value in enumerate(rows[marker_row_index - 1], start=1)
        }

    return _build_data_filter_state(
        disabled_row_marker_values,
        range_row_marker_values,
        marker_row_values,
        data_filter_config,
        set(range(1, len(rows) + 1)),
    )


def _should_search_sheet(sheet_name: str, disabled_sheet_marker: str | None) -> bool:
    marker = (disabled_sheet_marker or "").strip()
    return not marker or not sheet_name.startswith(marker)


def _is_cancelled(cancel_event: Event | None) -> bool:
    return cancel_event is not None and cancel_event.is_set()


def _get_matched_text(value: object, keyword: str) -> str | None:
    if value is None:
        return None

    text = str(value)
    if text.startswith("="):
        return None

    if keyword in text.lower():
        return text
    return None
