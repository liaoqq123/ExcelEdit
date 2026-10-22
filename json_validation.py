"""Excel 工作表中 JSON 列的批量格式检查。"""

from dataclasses import dataclass, field
import json
from pathlib import Path
from threading import Event
import zipfile
from xml.etree import ElementTree

from cell_search import (
    CELL_TAG,
    ROW_TAG,
    _get_cell_position,
    _get_xml_cell_text,
    _local_name,
    _read_shared_strings,
    _read_workbook_sheets,
)
from excel_common import ExcelReadError, clean_cell_text, iter_excel_files
from search_status import SearchTracker, describe_error


JSON_SAMPLE_ROWS = (4, 6)
JSON_DATA_START_ROW = min(JSON_SAMPLE_ROWS)
DEFAULT_MAX_JSON_VALIDATION_ERRORS = 20_000


@dataclass(frozen=True)
class JsonValidationError:
    """单个 JSON 格式错误的位置、原始内容和原因。"""

    path: Path
    workbook_name: str
    sheet_name: str
    row_index: int
    column_index: int
    value: str
    message: str


@dataclass(frozen=True)
class JsonValidationFileIssue:
    """无法读取的工作簿信息。"""

    path: Path
    message: str


@dataclass
class JsonValidationReport:
    """一次检查的汇总数据。"""

    errors: list[JsonValidationError] = field(default_factory=list)
    file_issues: list[JsonValidationFileIssue] = field(default_factory=list)
    workbook_count: int = 0
    sheet_count: int = 0
    json_column_count: int = 0
    checked_cell_count: int = 0
    limit_reached: bool = False


def validate_excel_json_columns(
    folder_path: str | Path,
    cancel_event: Event | None = None,
    max_errors: int | None = DEFAULT_MAX_JSON_VALIDATION_ERRORS,
    tracker: SearchTracker | None = None,
) -> JsonValidationReport:
    """检查所有非 `$` 工作表中由第 4、6 行识别出的 JSON 列。

    当同一列的第 4 行和第 6 行都能被正确解析为 JSON 时，该列被识别为 JSON
    列。随后检查从第 4 行开始、排除表头后的全部非空单元格；空白单元格不视为
    格式错误。
    """

    report = JsonValidationReport()
    tracker = tracker or SearchTracker()
    for path in iter_excel_files(folder_path, cancel_event=cancel_event, tracker=tracker):
        if _is_cancelled(cancel_event) or _has_reached_error_limit(report, max_errors):
            break
        tracker.scan(path)
        try:
            _validate_workbook(path, report, cancel_event, max_errors)
        except ExcelReadError as exc:
            report.file_issues.append(JsonValidationFileIssue(path, describe_error(exc)))
            tracker.issue(path, exc)
        tracker.result_count = len(report.errors)
        if _has_reached_error_limit(report, max_errors):
            report.limit_reached = True
            break

    tracker.finish("cancelled" if _is_cancelled(cancel_event) else "limit" if report.limit_reached else "completed")

    return report


def _validate_workbook(
    path: Path,
    report: JsonValidationReport,
    cancel_event: Event | None,
    max_errors: int | None,
) -> None:
    """根据 Excel 文件类型分发检查逻辑。"""

    suffix = path.suffix.lower()
    if suffix in {".xlsx", ".xlsm", ".xltx", ".xltm"}:
        _validate_openxml_workbook(path, report, cancel_event, max_errors)
        return
    if suffix == ".xls":
        _validate_xlrd_workbook(path, report, cancel_event, max_errors)
        return
    if suffix == ".xlsb":
        _validate_xlsb_workbook(path, report, cancel_event, max_errors)
        return
    raise ExcelReadError(f"不支持的 Excel 文件类型：{path.suffix}")


def _validate_openxml_workbook(
    path: Path,
    report: JsonValidationReport,
    cancel_event: Event | None,
    max_errors: int | None,
) -> None:
    """直接流式解析 Open XML，避免 openpyxl 随机读取时反复扫描工作表。"""

    try:
        with zipfile.ZipFile(path) as archive:
            shared_strings = _read_shared_strings(archive, cancel_event)
            if _is_cancelled(cancel_event):
                return

            report.workbook_count += 1
            for sheet_name, sheet_xml_path in _read_workbook_sheets(archive):
                if _is_cancelled(cancel_event) or _has_reached_error_limit(report, max_errors):
                    break
                if sheet_name.startswith("$"):
                    continue
                report.sheet_count += 1
                with archive.open(sheet_xml_path) as sheet_stream:
                    _validate_openxml_sheet(
                        path,
                        sheet_name,
                        sheet_stream,
                        shared_strings,
                        report,
                        cancel_event,
                        max_errors,
                    )
    except (OSError, KeyError, zipfile.BadZipFile, ElementTree.ParseError) as exc:
        raise ExcelReadError(f"无法打开 Excel 文件：{path}") from exc


def _validate_openxml_sheet(
    path: Path,
    sheet_name: str,
    sheet_stream: object,
    shared_strings: list[str],
    report: JsonValidationReport,
    cancel_event: Event | None,
    max_errors: int | None,
) -> None:
    """只处理 XML 中实际存在的数据单元格，跳过格式化造成的空白区域。"""

    current_row_index = 0
    current_column_index = 0
    sample_values: dict[int, dict[int, str]] = {
        row_index: {}
        for row_index in range(JSON_DATA_START_ROW, max(JSON_SAMPLE_ROWS) + 1)
    }
    candidate_columns: set[int] | None = None
    sheet_data_element: ElementTree.Element | None = None

    for event, element in ElementTree.iterparse(sheet_stream, events=("start", "end")):
        if _is_cancelled(cancel_event) or _has_reached_error_limit(report, max_errors):
            return
        if event == "start" and _local_name(element.tag) == "sheetData":
            sheet_data_element = element
            continue
        if event == "start" and (element.tag == ROW_TAG or _local_name(element.tag) == "row"):
            current_row_index = _parse_int(element.attrib.get("r"), current_row_index + 1)
            current_column_index = 0
            continue
        if event == "end" and (element.tag == CELL_TAG or _local_name(element.tag) == "c"):
            current_column_index += 1
            row_index, column_index = _get_cell_position(
                element.attrib.get("r", ""),
                current_row_index,
                current_column_index,
            )
            value = _get_xml_cell_text(element, shared_strings)
            if row_index in sample_values:
                sample_values[row_index][column_index] = value
            elif candidate_columns is not None and row_index >= JSON_DATA_START_ROW and column_index in candidate_columns:
                _validate_value(
                    path,
                    sheet_name,
                    row_index,
                    column_index,
                    value,
                    report,
                    max_errors,
                )
            element.clear()
            continue
        if event != "end" or (element.tag != ROW_TAG and _local_name(element.tag) != "row"):
            continue

        if current_row_index == max(JSON_SAMPLE_ROWS):
            max_column = max(
                (max(row_values, default=0) for row_values in sample_values.values()),
                default=0,
            )
            candidate_columns = _find_json_columns(
                max_column,
                lambda column_index: [
                    sample_values.get(sample_row, {}).get(column_index)
                    for sample_row in JSON_SAMPLE_ROWS
                ],
            )
            report.json_column_count += len(candidate_columns)
            for sample_row in range(JSON_DATA_START_ROW, current_row_index + 1):
                for column_index in candidate_columns:
                    _validate_value(
                        path,
                        sheet_name,
                        sample_row,
                        column_index,
                        sample_values.get(sample_row, {}).get(column_index),
                        report,
                        max_errors,
                    )

        element.clear()
        if sheet_data_element is not None:
            sheet_data_element.clear()


def _validate_xlrd_workbook(
    path: Path,
    report: JsonValidationReport,
    cancel_event: Event | None,
    max_errors: int | None,
) -> None:
    try:
        import xlrd
    except ImportError as exc:
        raise ExcelReadError("读取 .xls 文件需要 xlrd") from exc

    try:
        workbook = xlrd.open_workbook(str(path), on_demand=True)
    except Exception as exc:
        raise ExcelReadError(f"无法打开 Excel 文件：{path}") from exc

    try:
        report.workbook_count += 1
        for sheet_name in workbook.sheet_names():
            if _is_cancelled(cancel_event) or _has_reached_error_limit(report, max_errors):
                break
            if sheet_name.startswith("$"):
                continue
            worksheet = workbook.sheet_by_name(sheet_name)
            report.sheet_count += 1
            _validate_sheet(
                path,
                sheet_name,
                worksheet.nrows,
                worksheet.ncols,
                lambda row_index, column_index: worksheet.cell_value(row_index - 1, column_index - 1),
                report,
                cancel_event,
                max_errors,
            )
    finally:
        workbook.release_resources()


def _validate_xlsb_workbook(
    path: Path,
    report: JsonValidationReport,
    cancel_event: Event | None,
    max_errors: int | None,
) -> None:
    try:
        from pyxlsb import open_workbook
    except ImportError as exc:
        raise ExcelReadError("读取 .xlsb 文件需要 pyxlsb") from exc

    try:
        with open_workbook(str(path)) as workbook:
            report.workbook_count += 1
            for sheet_name in workbook.sheets:
                if _is_cancelled(cancel_event) or _has_reached_error_limit(report, max_errors):
                    break
                if sheet_name.startswith("$"):
                    continue
                report.sheet_count += 1
                _validate_xlsb_sheet(path, workbook, sheet_name, report, cancel_event, max_errors)
    except ExcelReadError:
        raise
    except Exception as exc:
        raise ExcelReadError(f"无法打开 Excel 文件：{path}") from exc


def _validate_xlsb_sheet(
    path: Path,
    workbook: object,
    sheet_name: str,
    report: JsonValidationReport,
    cancel_event: Event | None,
    max_errors: int | None,
) -> None:
    """按流式行读取 .xlsb，先收集第 4、6 行再校验整列。"""

    with workbook.get_sheet(sheet_name) as worksheet:
        _validate_streamed_rows(
            path,
            sheet_name,
            (
                (row_index, [cell.v for cell in row])
                for row_index, row in enumerate(worksheet.rows(), start=1)
            ),
            report,
            cancel_event,
            max_errors,
        )


def _validate_streamed_rows(
    path: Path,
    sheet_name: str,
    rows: object,
    report: JsonValidationReport,
    cancel_event: Event | None,
    max_errors: int | None,
) -> None:
    """顺序读取一次工作表，在第 6 行确认 JSON 列后校验整列数据。"""

    if not hasattr(rows, "__iter__"):
        return

    sample_row_limit = max(JSON_SAMPLE_ROWS)
    first_rows: dict[int, list[object]] = {}
    candidate_columns: set[int] | None = None
    for row_index, values in rows:
        if _is_cancelled(cancel_event) or _has_reached_error_limit(report, max_errors):
            return

        if row_index <= sample_row_limit:
            first_rows[row_index] = values
        if row_index == sample_row_limit:
            max_column = max((len(row) for row in first_rows.values()), default=0)
            candidate_columns = _find_json_columns(
                max_column,
                lambda column_index: [
                    _get_row_value(first_rows.get(sample_row, []), column_index)
                    for sample_row in JSON_SAMPLE_ROWS
                ],
            )
            report.json_column_count += len(candidate_columns)
            for buffered_row_index in range(JSON_DATA_START_ROW, row_index + 1):
                _validate_row_values(
                    path,
                    sheet_name,
                    buffered_row_index,
                    first_rows.get(buffered_row_index, []),
                    candidate_columns,
                    report,
                    max_errors,
                )
            continue

        if candidate_columns is not None and row_index >= JSON_DATA_START_ROW:
            _validate_row_values(
                path,
                sheet_name,
                row_index,
                values,
                candidate_columns,
                report,
                max_errors,
            )


def _validate_sheet(
    path: Path,
    sheet_name: str,
    max_row: int,
    max_column: int,
    get_value: object,
    report: JsonValidationReport,
    cancel_event: Event | None,
    max_errors: int | None,
) -> None:
    """校验普通二维工作表中的候选 JSON 列。"""

    if max_row < max(JSON_SAMPLE_ROWS) or max_column < 1:
        return
    value_getter = get_value
    if not callable(value_getter):
        return
    candidate_columns = _find_json_columns(
        max_column,
        lambda column_index: [value_getter(sample_row, column_index) for sample_row in JSON_SAMPLE_ROWS],
    )
    report.json_column_count += len(candidate_columns)
    for column_index in candidate_columns:
        for row_index in range(JSON_DATA_START_ROW, max_row + 1):
            if _is_cancelled(cancel_event) or _has_reached_error_limit(report, max_errors):
                return
            _validate_value(
                path,
                sheet_name,
                row_index,
                column_index,
                value_getter(row_index, column_index),
                report,
                max_errors,
            )


def _find_json_columns(max_column: int, get_samples: object) -> set[int]:
    """依据第 4、6 行能否正确解析为 JSON 来定位 JSON 列。"""

    if not callable(get_samples):
        return set()
    columns: set[int] = set()
    for column_index in range(1, max_column + 1):
        samples = [clean_cell_text(value) for value in get_samples(column_index)]
        if samples and all(_is_valid_json(sample) for sample in samples):
            columns.add(column_index)
    return columns


def _validate_row_values(
    path: Path,
    sheet_name: str,
    row_index: int,
    values: list[object],
    candidate_columns: set[int],
    report: JsonValidationReport,
    max_errors: int | None,
) -> None:
    for column_index in candidate_columns:
        if _has_reached_error_limit(report, max_errors):
            return
        _validate_value(
            path,
            sheet_name,
            row_index,
            column_index,
            _get_row_value(values, column_index),
            report,
            max_errors,
        )


def _validate_value(
    path: Path,
    sheet_name: str,
    row_index: int,
    column_index: int,
    value: object,
    report: JsonValidationReport,
    max_errors: int | None,
) -> None:
    text = clean_cell_text(value)
    if not text or _has_reached_error_limit(report, max_errors):
        return
    report.checked_cell_count += 1
    try:
        _strict_json_loads(text)
    except json.JSONDecodeError as exc:
        report.errors.append(
            JsonValidationError(
                path=path,
                workbook_name=path.name,
                sheet_name=sheet_name,
                row_index=row_index,
                column_index=column_index,
                value=text,
                message=f"{exc.msg}（JSON 字符位置 {exc.pos + 1}）",
            )
        )
        if _has_reached_error_limit(report, max_errors):
            report.limit_reached = True
    except ValueError as exc:
        report.errors.append(
            JsonValidationError(
                path=path,
                workbook_name=path.name,
                sheet_name=sheet_name,
                row_index=row_index,
                column_index=column_index,
                value=text,
                message=str(exc),
            )
        )
        if _has_reached_error_limit(report, max_errors):
            report.limit_reached = True


def _is_valid_json(value: str) -> bool:
    """判断单元格是否为符合标准的 JSON 数据。"""

    if not value:
        return False
    try:
        _strict_json_loads(value)
    except (json.JSONDecodeError, ValueError):
        return False
    return True


def _strict_json_loads(value: str) -> object:
    """按 JSON 标准解析，拒绝 Python 默认会接纳的 NaN 和 Infinity。"""

    return json.loads(value, parse_constant=_reject_non_standard_json_constant)


def _reject_non_standard_json_constant(value: str) -> object:
    """拒绝不是 JSON 标准字面量的 NaN、Infinity 等特殊常量。"""

    raise ValueError(f"不支持的 JSON 常量：{value}")


def _get_row_value(values: list[object], column_index: int) -> object:
    offset = column_index - 1
    return values[offset] if 0 <= offset < len(values) else None


def _parse_int(value: object, default: int) -> int:
    """安全读取 XML 中的行号，缺失或异常时使用顺序行号兜底。"""

    try:
        return int(str(value))
    except (TypeError, ValueError):
        return default


def _has_reached_error_limit(report: JsonValidationReport, max_errors: int | None) -> bool:
    return max_errors is not None and len(report.errors) >= max_errors


def _is_cancelled(cancel_event: Event | None) -> bool:
    return cancel_event is not None and cancel_event.is_set()
