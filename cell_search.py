"""Excel 单元格内容检索逻辑。

新版 .xlsx/.xlsm 等文件会先快速预筛 XML 字节，命中后再流式定位单元格；
大型 XML 的预筛也会分块执行，.xls 和 .xlsb 使用对应第三方库处理。
"""

from dataclasses import dataclass
from datetime import date, datetime, time
from html import escape, unescape
from pathlib import Path
from threading import Event
import posixpath
import re
import zipfile
from xml.etree import ElementTree

from openpyxl.styles.numbers import BUILTIN_FORMATS, is_date_format
from openpyxl.utils.datetime import CALENDAR_MAC_1904, CALENDAR_WINDOWS_1900, from_excel

from excel_common import EXCEL_MAX_COLUMNS, EXCEL_MAX_ROWS, ExcelReadError, iter_excel_files


MAIN_NAMESPACE = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
OFFICE_RELATIONSHIP_NAMESPACE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
WORKBOOK_PART_PATH = "xl/workbook.xml"
WORKBOOK_RELS_PART_PATH = "xl/_rels/workbook.xml.rels"
SHARED_STRINGS_PART_PATH = "xl/sharedStrings.xml"
STYLES_PART_PATH = "xl/styles.xml"
CELL_REFERENCE_PATTERN = re.compile(r"^([A-Za-z]+)(\d+)$")
SHARED_VALUE_PATTERN = re.compile(rb"<v>\s*(\d+)\s*</v>")
INLINE_RICH_TEXT_CELL_PATTERN = re.compile(
    rb"<c\b(?=[^>]*\bt=\"inlinestr\")[^>]*>.*?<is\b[^>]*>.*?<r\b",
    re.DOTALL,
)
SHARED_STRING_ITEM_PATTERN = re.compile(rb"<si\b[^>]*>(.*?)</si>", re.DOTALL)
SHARED_STRING_TEXT_PATTERN = re.compile(rb"<t(?:\s[^>]*)?>(.*?)</t>", re.DOTALL)
CELL_STYLE_PATTERN = re.compile(rb"<c\b[^>]*\bs=\"(\d+)\"[^>]*>")
ROW_TAG = f"{{{MAIN_NAMESPACE}}}row"
CELL_TAG = f"{{{MAIN_NAMESPACE}}}c"
TEXT_TAG = f"{{{MAIN_NAMESPACE}}}t"
DEFAULT_MAX_SEARCH_RESULTS = 20_000
FAST_PREFILTER_MEMORY_LIMIT = 8 * 1024 * 1024
PREFILTER_CHUNK_SIZE = 1024 * 1024
PREFILTER_OVERLAP_SIZE = 4096
ROW_FILTER_MODE_NONE = "无"
ROW_FILTER_MODE_DISABLED = "禁用"
ROW_FILTER_MODE_RANGE = "区间"
ROW_FILTER_MODES = (ROW_FILTER_MODE_NONE, ROW_FILTER_MODE_DISABLED, ROW_FILTER_MODE_RANGE)


@dataclass(frozen=True)
class ExcelCellMatch:
    """单元格检索命中结果，包含打开文件并定位单元格所需的信息。"""

    path: Path
    workbook_name: str
    sheet_name: str
    row_index: int
    column_index: int
    value: str


@dataclass(frozen=True)
class ExcelCellSearchIssue:
    """单个工作簿搜索失败时记录路径和错误，不中断其他文件。"""

    path: Path
    message: str


@dataclass(frozen=True)
class WorkbookFormatInfo:
    """工作簿日期系统和各单元格样式对应的数字格式。"""

    epoch: datetime = CALENDAR_WINDOWS_1900
    number_formats: tuple[str, ...] = ()


@dataclass(frozen=True)
class DataFilterConfig:
    """启用数据过滤配置。

    可跳过表头、排除空白表头列、排除禁用列、排除禁用行，
    或只保留某个起止标记之间的行。
    """

    enable_header_filter: bool = False
    header_row_count: int = 1
    enable_blank_header_column_filter: bool = False
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
    """把过滤配置转换成搜索时可快速判断的集合状态。"""

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
    issues: list[ExcelCellSearchIssue] | None = None,
    max_results: int | None = DEFAULT_MAX_SEARCH_RESULTS,
    exact_match: bool = False,
) -> list[ExcelCellMatch]:
    """在目录内所有 Excel 文件中搜索单元格缓存值。"""

    normalized_keyword = keyword.strip().lower()
    if not normalized_keyword:
        raise ExcelReadError("Search keyword cannot be empty")

    matches: list[ExcelCellMatch] = []
    for path in iter_excel_files(folder_path, cancel_event=cancel_event):
        if _is_cancelled(cancel_event):
            break
        remaining_results = None if max_results is None else max(max_results - len(matches), 0)
        if remaining_results == 0:
            break
        try:
            file_matches = _search_excel_file_cells(
                path,
                normalized_keyword,
                disabled_sheet_marker,
                data_filter_config,
                cancel_event,
                remaining_results,
                exact_match,
            )
        except ExcelReadError as exc:
            if issues is not None:
                issues.append(ExcelCellSearchIssue(path=path, message=str(exc)))
            continue
        matches.extend(file_matches)

    return matches


def _search_excel_file_cells(
    path: Path,
    keyword: str,
    disabled_sheet_marker: str | None,
    data_filter_config: DataFilterConfig | None,
    cancel_event: Event | None,
    max_results: int | None,
    exact_match: bool,
) -> list[ExcelCellMatch]:
    """根据 Excel 文件格式分发到对应搜索实现。"""

    suffix = path.suffix.lower()
    if suffix in {".xlsx", ".xlsm", ".xltx", ".xltm"}:
        return _search_xlsx_xml_cells(
            path,
            keyword,
            disabled_sheet_marker,
            data_filter_config,
            cancel_event,
            max_results,
            exact_match,
        )
    if suffix == ".xls":
        return _search_xlrd_cells(
            path,
            keyword,
            disabled_sheet_marker,
            data_filter_config,
            cancel_event,
            max_results,
            exact_match,
        )
    if suffix == ".xlsb":
        return _search_xlsb_cells(
            path,
            keyword,
            disabled_sheet_marker,
            data_filter_config,
            cancel_event,
            max_results,
            exact_match,
        )
    return []


def _search_xlsx_xml_cells(
    path: Path,
    keyword: str,
    disabled_sheet_marker: str | None,
    data_filter_config: DataFilterConfig | None,
    cancel_event: Event | None,
    max_results: int | None,
    exact_match: bool,
) -> list[ExcelCellMatch]:
    """直接解析 xlsx 压缩包 XML 来搜索单元格，避免完整加载工作簿。"""

    try:
        with zipfile.ZipFile(path) as archive:
            if _is_cancelled(cancel_event):
                return []

            # 启用数据过滤时后面还要读取列/行标记，所以先完整拿到共享字符串表。
            shared_strings = _read_shared_strings(archive, cancel_event) if data_filter_config is not None else None
            matching_shared_strings = (
                _filter_matching_shared_strings(shared_strings, keyword, cancel_event, exact_match)
                if shared_strings is not None
                else _read_matching_shared_strings(archive, keyword, cancel_event, exact_match)
            )
            if _is_cancelled(cancel_event):
                return []

            matching_shared_string_ids = {
                str(shared_string_id).encode("ascii")
                for shared_string_id in matching_shared_strings
            }
            sheets = _read_workbook_sheets(archive)
            format_info = (
                _read_workbook_format_info(archive)
                if _keyword_may_target_formatted_number(keyword)
                else WorkbookFormatInfo()
            )
            relevant_style_ids = _get_relevant_formatted_style_ids(format_info, keyword)
            direct_hints = _direct_keyword_hints(keyword)
            matches: list[ExcelCellMatch] = []
            for sheet_name, sheet_xml_path in sheets:
                if _is_cancelled(cancel_event):
                    break
                remaining_results = None if max_results is None else max(max_results - len(matches), 0)
                if remaining_results == 0:
                    break
                if not _should_search_sheet(sheet_name, disabled_sheet_marker):
                    continue
                if not _sheet_xml_has_search_hint(
                    archive,
                    sheet_xml_path,
                    matching_shared_string_ids,
                    keyword,
                    direct_hints,
                    relevant_style_ids,
                    cancel_event,
                ):
                    continue
                matches.extend(
                    _search_sheet_xml_cells(
                        archive,
                        path,
                        sheet_name,
                        sheet_xml_path,
                        matching_shared_strings,
                        keyword,
                        shared_strings,
                        data_filter_config,
                        format_info,
                        cancel_event,
                        remaining_results,
                        exact_match,
                    )
                )
            return matches
    except (OSError, KeyError, zipfile.BadZipFile, ElementTree.ParseError) as exc:
        raise ExcelReadError(f"Could not open Excel file: {path}") from exc


def _read_workbook_sheets(archive: zipfile.ZipFile) -> list[tuple[str, str]]:
    """从 workbook.xml 和关系文件中解析工作表名称与 XML 路径。"""

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
    """读取 workbook.xml.rels，把关系 ID 映射到实际工作表 XML 路径。"""

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


def _read_workbook_format_info(archive: zipfile.ZipFile) -> WorkbookFormatInfo:
    """读取日期系统和数字格式，用于按 Excel 中可见的值进行搜索。"""

    workbook_root = ElementTree.fromstring(archive.read(WORKBOOK_PART_PATH))
    epoch = CALENDAR_WINDOWS_1900
    for element in workbook_root.iter():
        if _local_name(element.tag) != "workbookPr":
            continue
        if str(element.attrib.get("date1904", "")).lower() in {"1", "true"}:
            epoch = CALENDAR_MAC_1904
        break

    try:
        styles_root = ElementTree.fromstring(archive.read(STYLES_PART_PATH))
    except KeyError:
        return WorkbookFormatInfo(epoch=epoch)

    custom_formats: dict[int, str] = {}
    cell_formats: list[str] = []
    for element in styles_root.iter():
        local_name = _local_name(element.tag)
        if local_name == "numFmt":
            num_fmt_id = _parse_int(element.attrib.get("numFmtId"), -1)
            format_code = element.attrib.get("formatCode", "")
            if num_fmt_id >= 0 and format_code:
                custom_formats[num_fmt_id] = format_code
        if local_name != "cellXfs":
            continue
        for cell_format in element:
            if _local_name(cell_format.tag) != "xf":
                continue
            num_fmt_id = _parse_int(cell_format.attrib.get("numFmtId"), 0)
            cell_formats.append(custom_formats.get(num_fmt_id, BUILTIN_FORMATS.get(num_fmt_id, "General")))

    return WorkbookFormatInfo(epoch=epoch, number_formats=tuple(cell_formats))


def _get_relevant_formatted_style_ids(format_info: WorkbookFormatInfo, keyword: str) -> set[bytes]:
    """找出可能让底层数值格式化后命中关键词的单元格样式。"""

    return {
        str(style_index).encode("ascii")
        for style_index, number_format in enumerate(format_info.number_formats)
        if _number_format_may_match_keyword(number_format, keyword)
    }


def _keyword_may_target_formatted_number(keyword: str) -> bool:
    """普通文字无需读取 styles.xml；仅疑似格式化数字搜索才加载样式。"""

    return any(char.isdigit() for char in keyword) and any(
        marker in keyword for marker in ("-", "/", ":", ".", ",", "%", "年", "月", "日", "¥", "￥", "$", "€", "£")
    )


def _number_format_may_match_keyword(number_format: str, keyword: str) -> bool:
    """判断搜索词是否可能命中某种日期、百分比、货币或小数显示格式。"""

    format_section = number_format.split(";", 1)[0]
    if is_date_format(number_format):
        return any(char.isdigit() for char in keyword) or any(
            marker in keyword for marker in ("-", "/", ":", "年", "月", "日")
        )
    if "%" in format_section and "%" in keyword:
        return True
    if any(symbol in format_section and symbol in keyword for symbol in ("¥", "￥", "$", "€", "£")):
        return True
    if "," in format_section and "," in keyword:
        return True
    return "." in keyword and re.search(r"\.[0#]+", format_section) is not None


def _sheet_xml_has_search_hint(
    archive: zipfile.ZipFile,
    sheet_xml_path: str,
    matching_shared_string_ids: set[bytes],
    keyword: str,
    direct_hints: set[bytes],
    relevant_style_ids: set[bytes],
    cancel_event: Event | None = None,
) -> bool:
    """先快速扫描 XML 字节；小文件整块检查，大文件分块检查以限制内存。"""

    if _is_cancelled(cancel_event):
        return False

    sheet_info = archive.getinfo(sheet_xml_path)
    if sheet_info.file_size <= FAST_PREFILTER_MEMORY_LIMIT:
        normalized_xml = archive.read(sheet_xml_path).lower()
        return _prefilter_buffer_has_hint(
            normalized_xml,
            matching_shared_string_ids,
            direct_hints,
            keyword,
            relevant_style_ids,
        )

    overlap = b""
    with archive.open(sheet_xml_path) as stream:
        while not _is_cancelled(cancel_event):
            chunk = stream.read(PREFILTER_CHUNK_SIZE)
            if not chunk:
                return False
            normalized_chunk = (overlap + chunk).lower()
            if _prefilter_buffer_has_hint(
                normalized_chunk,
                matching_shared_string_ids,
                direct_hints,
                keyword,
                relevant_style_ids,
            ):
                return True
            overlap = normalized_chunk[-PREFILTER_OVERLAP_SIZE:]
    return False


def _prefilter_buffer_has_hint(
    normalized_xml: bytes,
    matching_shared_string_ids: set[bytes],
    direct_hints: set[bytes],
    keyword: str,
    relevant_style_ids: set[bytes],
) -> bool:
    """检查一段小写 XML 字节中是否存在需要进一步解析的迹象。"""

    if matching_shared_string_ids and _has_typed_shared_string_reference(
        normalized_xml,
        matching_shared_string_ids,
    ):
        return True
    if any(hint and hint in normalized_xml for hint in direct_hints):
        return True
    # 普通 inlineStr 和公式缓存字符串已经由直接字节命中覆盖；只对可能跨多个
    # <r><t>...</t></r> 片段的富文本保留完整解析兜底。
    if b"<r>" in normalized_xml and b't="inlinestr"' in normalized_xml:
        if INLINE_RICH_TEXT_CELL_PATTERN.search(normalized_xml) is not None:
            return True
    if keyword in {"true", "false"} and b't="b"' in normalized_xml:
        return True
    if relevant_style_ids:
        for match in CELL_STYLE_PATTERN.finditer(normalized_xml):
            if match.group(1) in relevant_style_ids:
                return True
    return False


def _has_typed_shared_string_reference(
    normalized_xml: bytes,
    matching_shared_string_ids: set[bytes],
) -> bool:
    """先线性定位候选编号，仅在编号命中时确认它属于 t="s" 单元格。"""

    for shared_string_id in matching_shared_string_ids:
        value_tag = b"<v>" + shared_string_id + b"</v>"
        search_start = 0
        while True:
            value_start = normalized_xml.find(value_tag, search_start)
            if value_start < 0:
                break
            cell_start = normalized_xml.rfind(b"<c", max(0, value_start - 1024), value_start)
            if cell_start >= 0:
                cell_tag_end = normalized_xml.find(b">", cell_start, value_start)
                if cell_tag_end >= 0 and b't="s"' in normalized_xml[cell_start : cell_tag_end + 1]:
                    return True
            search_start = value_start + len(value_tag)
    return False


def _direct_keyword_hints(keyword: str) -> set[bytes]:
    """生成 XML 字节预筛需要匹配的原文和转义文本。"""

    return {
        keyword.encode("utf-8"),
        escape(keyword, quote=False).lower().encode("utf-8"),
    }


def _read_shared_strings(archive: zipfile.ZipFile, cancel_event: Event | None = None) -> list[str]:
    """读取 xlsx 共享字符串表，索引位置对应单元格里的 shared string id。"""

    try:
        shared_strings_xml = archive.read(SHARED_STRINGS_PART_PATH)
    except KeyError:
        return []

    shared_strings: list[str] = []
    for item_match in SHARED_STRING_ITEM_PATTERN.finditer(shared_strings_xml):
        if _is_cancelled(cancel_event):
            break
        shared_strings.append(_decode_shared_string_item(item_match.group(1)))
    return shared_strings


def _filter_matching_shared_strings(
    shared_strings: list[str],
    keyword: str,
    cancel_event: Event | None = None,
    exact_match: bool = False,
) -> dict[int, str]:
    """在已读取的共享字符串表中筛出包含关键词的条目。"""

    matches: dict[int, str] = {}
    for shared_string_index, text in enumerate(shared_strings):
        if _is_cancelled(cancel_event):
            break
        if _get_matched_text(text, keyword, exact_match) is not None:
            matches[shared_string_index] = text
    return matches


def _read_matching_shared_strings(
    archive: zipfile.ZipFile,
    keyword: str,
    cancel_event: Event | None = None,
    exact_match: bool = False,
) -> dict[int, str]:
    """只读取命中关键词的共享字符串，用于不需要数据过滤的快速路径。"""

    try:
        shared_strings_xml = archive.read(SHARED_STRINGS_PART_PATH)
    except KeyError:
        return {}

    if _is_cancelled(cancel_event):
        return {}

    normalized_xml = shared_strings_xml.lower()
    direct_hints = _direct_keyword_hints(keyword)
    if not any(hint and hint in normalized_xml for hint in direct_hints) and b"<r>" not in normalized_xml:
        return {}

    matching_shared_strings: dict[int, str] = {}
    for shared_string_index, item_match in enumerate(SHARED_STRING_ITEM_PATTERN.finditer(shared_strings_xml)):
        if _is_cancelled(cancel_event):
            break
        item_xml = item_match.group(1)
        normalized_item_xml = item_xml.lower()
        if not any(hint and hint in normalized_item_xml for hint in direct_hints) and b"<r>" not in normalized_item_xml:
            continue
        text = _decode_shared_string_item(item_xml)
        if _get_matched_text(text, keyword, exact_match) is not None:
            matching_shared_strings[shared_string_index] = text
    return matching_shared_strings


def _decode_shared_string_item(item_xml: bytes) -> str:
    """直接拼接一个共享字符串条目中的文本段，避免构建 ElementTree 节点。"""

    text_parts: list[str] = []
    for text_match in SHARED_STRING_TEXT_PATTERN.finditer(item_xml):
        encoded_text = text_match.group(1)
        try:
            text = encoded_text.decode("utf-8")
        except UnicodeDecodeError:
            text = encoded_text.decode("utf-8", errors="replace")
        text_parts.append(unescape(text))
    return "".join(text_parts)


def _search_sheet_xml_cells(
    archive: zipfile.ZipFile,
    path: Path,
    sheet_name: str,
    sheet_xml_path: str,
    matching_shared_strings: dict[int, str],
    keyword: str,
    shared_strings: list[str] | None,
    data_filter_config: DataFilterConfig | None,
    format_info: WorkbookFormatInfo,
    cancel_event: Event | None,
    max_results: int | None,
    exact_match: bool,
) -> list[ExcelCellMatch]:
    """解析单个工作表 XML，定位真正命中的行、列和值。"""

    matches: list[ExcelCellMatch] = []
    current_row_index = 0
    current_column_index = 0
    sheet_data_element: ElementTree.Element | None = None
    if _is_cancelled(cancel_event):
        return matches

    data_filter_state = None
    if data_filter_config is not None:
        with archive.open(sheet_xml_path) as filter_stream:
            data_filter_state = _build_xlsx_data_filter_state(
                filter_stream,
                shared_strings or [],
                data_filter_config,
                cancel_event,
            )
    if _is_cancelled(cancel_event):
        return matches

    with archive.open(sheet_xml_path) as stream:
        for event, element in ElementTree.iterparse(stream, events=("start", "end")):
            if _is_cancelled(cancel_event):
                break
            if event == "start" and _local_name(element.tag) == "sheetData":
                sheet_data_element = element
                continue
            if event == "start" and (element.tag == ROW_TAG or _local_name(element.tag) == "row"):
                current_row_index = _parse_int(element.attrib.get("r"), current_row_index + 1)
                current_column_index = 0
                continue
            if event == "end" and (element.tag == ROW_TAG or _local_name(element.tag) == "row"):
                element.clear()
                if sheet_data_element is not None:
                    sheet_data_element.clear()
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
            matched_value = _get_xml_cell_match(
                element,
                matching_shared_strings,
                keyword,
                format_info,
                exact_match,
            )
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
                if max_results is not None and len(matches) >= max_results:
                    element.clear()
                    break
            element.clear()
    return matches


def _build_xlsx_data_filter_state(
    sheet_stream: object,
    shared_strings: list[str],
    data_filter_config: DataFilterConfig | None,
    cancel_event: Event | None = None,
) -> DataFilterState | None:
    """从 xlsx 工作表 XML 中提取启用数据过滤所需的行列标记。"""

    if data_filter_config is None:
        return None

    current_row_index = 0
    current_column_index = 0
    disabled_row_marker_values: dict[int, str] = {}
    range_row_marker_values: dict[int, str] = {}
    marker_row_values: dict[int, str] = {}
    available_row_indices: set[int] = set()
    available_column_indices: set[int] = set()
    header_column_values: dict[int, list[str]] = {}
    sheet_data_element: ElementTree.Element | None = None

    for event, element in ElementTree.iterparse(sheet_stream, events=("start", "end")):
        if _is_cancelled(cancel_event):
            return None
        if event == "start" and _local_name(element.tag) == "sheetData":
            sheet_data_element = element
            continue
        if event == "start" and (element.tag == ROW_TAG or _local_name(element.tag) == "row"):
            current_row_index = _parse_int(element.attrib.get("r"), current_row_index + 1)
            current_column_index = 0
            continue
        if event == "end" and (element.tag == ROW_TAG or _local_name(element.tag) == "row"):
            element.clear()
            if sheet_data_element is not None:
                sheet_data_element.clear()
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
        available_column_indices.add(column_index)
        text = _clean_filter_text(_get_xml_cell_text(element, shared_strings))
        if (
            data_filter_config.enable_blank_header_column_filter
            and row_index <= data_filter_config.header_row_count
        ):
            header_column_values.setdefault(column_index, []).append(text)
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
        available_column_indices,
        header_column_values,
    )


def _build_data_filter_state(
    disabled_row_marker_values: dict[int, str],
    range_row_marker_values: dict[int, str],
    marker_row_values: dict[int, str],
    data_filter_config: DataFilterConfig | None,
    available_row_indices: set[int] | None = None,
    available_column_indices: set[int] | None = None,
    header_column_values: dict[int, list[str]] | None = None,
) -> DataFilterState | None:
    """把原始标记值转换成可直接判断的禁用列、禁用行和启用行集合。"""

    if data_filter_config is None:
        return None

    header_row_count = data_filter_config.header_row_count if data_filter_config.enable_header_filter else 0
    marker = data_filter_config.disabled_column_marker.strip()
    disabled_columns = frozenset(
        column_index
        for column_index, text in marker_row_values.items()
        if data_filter_config.enable_column_filter and marker and text.startswith(marker)
    )
    blank_header_columns = _get_blank_header_columns(
        header_column_values or {},
        available_column_indices or set(),
        data_filter_config,
    )
    disabled_columns = frozenset((*disabled_columns, *blank_header_columns))

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


def _get_blank_header_columns(
    header_column_values: dict[int, list[str]],
    available_column_indices: set[int],
    data_filter_config: DataFilterConfig,
) -> frozenset[int]:
    """找出表头范围内没有任何非空值的列。"""

    if not data_filter_config.enable_blank_header_column_filter:
        return frozenset()

    header_non_empty_columns = {
        column_index
        for column_index, values in header_column_values.items()
        if any(value.strip() for value in values)
    }
    return frozenset(
        column_index
        for column_index in available_column_indices
        if column_index not in header_non_empty_columns
    )


def _get_enabled_range_rows(
    row_marker_values: dict[int, str],
    range_start_text: str,
    range_end_text: str,
    available_row_indices: set[int],
) -> frozenset[int]:
    """根据起始/结束标记计算允许检索的行区间。"""

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
    """判断某个单元格是否通过所有启用数据过滤规则。"""

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


def _get_xml_cell_match(
    element: ElementTree.Element,
    matching_shared_strings: dict[int, str],
    keyword: str,
    format_info: WorkbookFormatInfo,
    exact_match: bool = False,
) -> str | None:
    """读取 XML 单元格显示值，并判断是否包含关键词。"""

    cell_type = element.attrib.get("t", "")
    if cell_type == "s":
        raw_value = _find_child_text(element, "v")
        if raw_value is None:
            return None

        shared_string_index = _parse_int(raw_value, -1)
        return matching_shared_strings.get(shared_string_index)

    if cell_type == "inlineStr":
        return _get_matched_text(_rich_text(element), keyword, exact_match)

    raw_value = _find_child_text(element, "v")
    if raw_value is None:
        return None

    if cell_type == "b":
        value = "TRUE" if raw_value == "1" else "FALSE" if raw_value == "0" else raw_value
        return _get_matched_text(value, keyword, exact_match)

    style_index = _parse_int(element.attrib.get("s"), -1)
    if 0 <= style_index < len(format_info.number_formats):
        formatted_values = _format_numeric_search_values(
            raw_value,
            format_info.number_formats[style_index],
            format_info.epoch,
        )
        formatted_match = _get_first_matched_text(formatted_values, keyword, exact_match)
        if formatted_match is not None:
            return formatted_match

    return _get_matched_text(raw_value, keyword, exact_match)


def _format_numeric_search_values(raw_value: str, number_format: str, epoch: datetime) -> tuple[str, ...]:
    """生成日期、百分比和常见数字格式的可搜索显示值。"""

    try:
        numeric_value = float(raw_value)
    except (TypeError, ValueError, OverflowError):
        return ()

    if is_date_format(number_format):
        try:
            converted_value = from_excel(numeric_value, epoch=epoch)
        except (TypeError, ValueError, OverflowError):
            return ()
        return _format_date_search_values(converted_value)

    format_section = number_format.split(";", 1)[0]
    decimal_match = re.search(r"\.([0#]+)", format_section)
    decimal_places = len(decimal_match.group(1)) if decimal_match else 0
    values: list[str] = []
    if "%" in format_section:
        values.append(f"{numeric_value * 100:.{decimal_places}f}%")
    elif "," in format_section or decimal_match:
        use_grouping = "," in format_section
        format_spec = f",.{decimal_places}f" if use_grouping else f".{decimal_places}f"
        number_text = format(numeric_value, format_spec)
        values.append(number_text)
        for symbol in ("¥", "￥", "$", "€", "£"):
            if symbol in format_section:
                values.extend((f"{symbol}{number_text}", f"{number_text}{symbol}"))
    return tuple(dict.fromkeys(values))


def _format_date_search_values(value: object) -> tuple[str, ...]:
    """生成常见日期/时间写法，使搜索值与 Excel 可见内容更接近。"""

    values: list[str] = []
    if isinstance(value, datetime):
        values.extend((value.isoformat(sep=" "), value.isoformat(sep=" ", timespec="seconds")))
        date_value: date | None = value.date()
    elif isinstance(value, date):
        date_value = value
    else:
        date_value = None

    if date_value is not None:
        year, month, day = date_value.year, date_value.month, date_value.day
        values.extend(
            (
                date_value.isoformat(),
                f"{year}/{month:02d}/{day:02d}",
                f"{year}/{month}/{day}",
                f"{year}-{month}-{day}",
                f"{year}年{month}月{day}日",
                f"{year}年{month:02d}月{day:02d}日",
            )
        )
    if isinstance(value, time):
        values.extend((value.isoformat(), value.isoformat(timespec="seconds")))
    return tuple(dict.fromkeys(values))


def _get_first_matched_text(
    values: tuple[str, ...],
    keyword: str,
    exact_match: bool = False,
) -> str | None:
    """从多个显示候选值中返回第一个命中项。"""

    for value in values:
        matched_value = _get_matched_text(value, keyword, exact_match)
        if matched_value is not None:
            return matched_value
    return None


def _get_xml_cell_text(element: ElementTree.Element, shared_strings: list[str]) -> str:
    """读取 XML 单元格文本值，用于构建启用数据过滤状态。"""

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
    """把 A1 这类引用转换成行列号；缺少引用时使用遍历位置兜底。"""

    match = CELL_REFERENCE_PATTERN.match(cell_reference)
    if match is None:
        return fallback_row_index, fallback_column_index

    column_letters, row_text = match.groups()
    return _parse_int(row_text, fallback_row_index), _column_letters_to_index(column_letters)


def _column_letters_to_index(column_letters: str) -> int:
    """把 Excel 列字母转换成 1 起始列号，例如 A=1、AA=27。"""

    column_index = 0
    for char in column_letters.upper():
        column_index = column_index * 26 + ord(char) - ord("A") + 1
    return column_index


def _resolve_part_path(base_part_path: str, target: str) -> str:
    """把关系文件中的相对路径解析成压缩包内的规范路径。"""

    if target.startswith("/"):
        return target.lstrip("/")
    return posixpath.normpath(posixpath.join(posixpath.dirname(base_part_path), target))


def _find_child_text(element: ElementTree.Element, child_name: str) -> str | None:
    """在 XML 元素的直接子节点中查找指定名称的文本。"""

    for child in element:
        if _local_name(child.tag) == child_name:
            return child.text or ""
    return None


def _rich_text(element: ElementTree.Element) -> str:
    """拼接富文本单元格中的所有文本段。"""

    text_parts: list[str] = []
    for text_element in element.iter():
        if (text_element.tag == TEXT_TAG or _local_name(text_element.tag) == "t") and text_element.text:
            text_parts.append(text_element.text)
    return "".join(text_parts)


def _xml_tag(name: str) -> str:
    """生成带主命名空间的 Excel XML 标签名。"""

    return f"{{{MAIN_NAMESPACE}}}{name}"


def _local_name(tag: str) -> str:
    """去掉 XML 命名空间，只保留标签本名。"""

    return tag.rsplit("}", 1)[-1]


def _parse_int(value: object, default: int) -> int:
    """安全转换整数，失败时返回默认值。"""

    try:
        return int(str(value))
    except (TypeError, ValueError):
        return default


def _clean_filter_text(value: object) -> str:
    """清理过滤配置比较时使用的文本。"""

    if value is None:
        return ""
    return str(value).strip()


def normalize_row_filter_mode(row_filter_mode: str) -> str:
    """兼容旧缓存中的行过滤模式值。"""

    return row_filter_mode if row_filter_mode in ROW_FILTER_MODES else ROW_FILTER_MODE_NONE


def validate_data_filter_config(config: DataFilterConfig) -> None:
    """校验启用数据过滤配置，避免搜索时出现无意义的行列参数。"""

    if config.enable_header_filter and config.header_row_count < 1:
        raise ExcelReadError("表头长度必须大于等于 1")
    if config.header_row_count > EXCEL_MAX_ROWS:
        raise ExcelReadError(f"表头长度不能大于 {EXCEL_MAX_ROWS}")
    if config.enable_blank_header_column_filter and config.header_row_count < 1:
        raise ExcelReadError("屏蔽空白表头时，表头长度必须大于等于 1")
    if config.enable_column_filter and config.column_marker_row_index < 1:
        raise ExcelReadError("列未启用判断行必须大于等于 1")
    if config.column_marker_row_index > EXCEL_MAX_ROWS:
        raise ExcelReadError(f"列未启用判断行不能大于 {EXCEL_MAX_ROWS}")
    if config.enable_column_filter and not config.disabled_column_marker.strip():
        raise ExcelReadError("启用列禁用时，必须填写列未启用标识")
    if config.enable_disabled_row_filter and config.disabled_row_marker_column_index < 1:
        raise ExcelReadError("行禁用判断列必须大于等于 1")
    if config.disabled_row_marker_column_index > EXCEL_MAX_COLUMNS:
        raise ExcelReadError(f"行禁用判断列不能大于 {EXCEL_MAX_COLUMNS}")
    if config.enable_disabled_row_filter and not config.disabled_row_contains.strip():
        raise ExcelReadError("行未启用方式为“禁用”时，必须填写行禁用包含数据")
    if config.enable_range_row_filter:
        if config.range_row_marker_column_index < 1:
            raise ExcelReadError("行区间判断列必须大于等于 1")
        if config.range_row_marker_column_index > EXCEL_MAX_COLUMNS:
            raise ExcelReadError(f"行区间判断列不能大于 {EXCEL_MAX_COLUMNS}")
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
    max_results: int | None,
    exact_match: bool,
) -> list[ExcelCellMatch]:
    """用 xlrd 搜索老格式 .xls 文件。"""

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
            if max_results is not None and len(matches) >= max_results:
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
                    cell = worksheet.cell(row_offset, column_offset)
                    matched_value = None
                    if cell.ctype == xlrd.XL_CELL_DATE:
                        try:
                            date_value = xlrd.xldate_as_datetime(cell.value, workbook.datemode)
                        except (TypeError, ValueError, OverflowError, xlrd.XLDateError):
                            date_value = None
                        if date_value is not None:
                            matched_value = _get_first_matched_text(
                                _format_date_search_values(date_value),
                                keyword,
                                exact_match,
                            )
                    if matched_value is None:
                        matched_value = _get_matched_text(cell.value, keyword, exact_match)
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
                    if max_results is not None and len(matches) >= max_results:
                        break
                if max_results is not None and len(matches) >= max_results:
                    break
        return matches
    finally:
        workbook.release_resources()


def _search_xlsb_cells(
    path: Path,
    keyword: str,
    disabled_sheet_marker: str | None,
    data_filter_config: DataFilterConfig | None,
    cancel_event: Event | None,
    max_results: int | None,
    exact_match: bool,
) -> list[ExcelCellMatch]:
    """用 pyxlsb 搜索 .xlsb 文件。"""

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
                if max_results is not None and len(matches) >= max_results:
                    break
                if not _should_search_sheet(sheet_name, disabled_sheet_marker):
                    continue
                with workbook.get_sheet(sheet_name) as worksheet:
                    data_filter_state = _build_xlsb_data_filter_state(
                        worksheet,
                        data_filter_config,
                        cancel_event,
                    )
                    if _is_cancelled(cancel_event):
                        break
                    for row_index, row in enumerate(worksheet.rows(), start=1):
                        if _is_cancelled(cancel_event):
                            break
                        for column_index, cell in enumerate(row, start=1):
                            if not _is_enabled_cell(row_index, column_index, data_filter_state):
                                continue
                            matched_value = _get_matched_text(cell.v, keyword, exact_match)
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
                            if max_results is not None and len(matches) >= max_results:
                                break
                        if max_results is not None and len(matches) >= max_results:
                            break
        return matches
    except Exception as exc:
        raise ExcelReadError(f"Could not open Excel file: {path}") from exc


def _build_xlrd_data_filter_state(
    worksheet: object,
    data_filter_config: DataFilterConfig | None,
    cancel_event: Event | None = None,
) -> DataFilterState | None:
    """从 xlrd 工作表中提取启用数据过滤状态。"""

    if data_filter_config is None:
        return None

    disabled_row_marker_values: dict[int, str] = {}
    range_row_marker_values: dict[int, str] = {}
    marker_row_values: dict[int, str] = {}
    available_column_indices: set[int] = set(range(1, getattr(worksheet, "ncols", 0) + 1))
    header_column_values: dict[int, list[str]] = {}
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

    if data_filter_config.enable_blank_header_column_filter:
        header_rows = min(data_filter_config.header_row_count, getattr(worksheet, "nrows", 0))
        for row_offset in range(header_rows):
            if _is_cancelled(cancel_event):
                return None
            for column_offset in range(getattr(worksheet, "ncols", 0)):
                header_column_values.setdefault(column_offset + 1, []).append(
                    _clean_filter_text(worksheet.cell_value(row_offset, column_offset))
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
        available_column_indices,
        header_column_values,
    )


def _build_xlsb_data_filter_state(
    worksheet: object,
    data_filter_config: DataFilterConfig | None,
    cancel_event: Event | None = None,
) -> DataFilterState | None:
    """流式读取 xlsb 行来构建过滤状态，避免把整张工作表放进内存。"""

    if data_filter_config is None:
        return None

    disabled_row_marker_values: dict[int, str] = {}
    range_row_marker_values: dict[int, str] = {}
    marker_row_values: dict[int, str] = {}
    available_row_indices: set[int] = set()
    available_column_indices: set[int] = set()
    header_column_values: dict[int, list[str]] = {}
    disabled_row_marker_column_index = data_filter_config.disabled_row_marker_column_index
    range_row_marker_column_index = data_filter_config.range_row_marker_column_index
    marker_row_index = data_filter_config.column_marker_row_index

    for row_index, row in enumerate(worksheet.rows(), start=1):
        if _is_cancelled(cancel_event):
            return None
        values = [cell.v for cell in row]
        available_row_indices.add(row_index)
        available_column_indices.update(range(1, len(values) + 1))
        if data_filter_config.enable_blank_header_column_filter and row_index <= data_filter_config.header_row_count:
            for column_index, value in enumerate(values, start=1):
                header_column_values.setdefault(column_index, []).append(_clean_filter_text(value))
        if data_filter_config.enable_column_filter and row_index == marker_row_index:
            marker_row_values = {
                column_index: _clean_filter_text(value)
                for column_index, value in enumerate(values, start=1)
            }
        if (
            data_filter_config.enable_disabled_row_filter
            and 1 <= disabled_row_marker_column_index <= len(values)
        ):
            disabled_row_marker_values[row_index] = _clean_filter_text(
                values[disabled_row_marker_column_index - 1]
            )
        if data_filter_config.enable_range_row_filter and 1 <= range_row_marker_column_index <= len(values):
            range_row_marker_values[row_index] = _clean_filter_text(values[range_row_marker_column_index - 1])

    return _build_data_filter_state(
        disabled_row_marker_values,
        range_row_marker_values,
        marker_row_values,
        data_filter_config,
        available_row_indices,
        available_column_indices,
        header_column_values,
    )


def _should_search_sheet(sheet_name: str, disabled_sheet_marker: str | None) -> bool:
    """根据工作表名称前缀判断是否跳过未启用工作表。"""

    marker = (disabled_sheet_marker or "").strip()
    return not marker or not sheet_name.startswith(marker)


def _is_cancelled(cancel_event: Event | None) -> bool:
    """判断后台单元格搜索是否已取消。"""

    return cancel_event is not None and cancel_event.is_set()


def _get_matched_text(value: object, keyword: str, exact_match: bool = False) -> str | None:
    """按包含或完整相等方式匹配最终缓存值；公式文本本身不参与匹配。"""

    if value is None:
        return None

    text = str(value)
    # 这里跳过公式表达式本身，搜索的是 Excel 保存后的缓存结果。
    if text.startswith("="):
        return None

    normalized_text = text.lower()
    is_match = normalized_text == keyword if exact_match else keyword in normalized_text
    if is_match:
        return text
    return None
