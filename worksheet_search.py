import re
from dataclasses import dataclass
from pathlib import Path

from openpyxl import load_workbook

from excel_common import (
    ExcelReadError,
    clean_cell_text,
    get_sheet_names,
    iter_excel_files,
)


@dataclass(frozen=True)
class ExcelWorkbookInfo:
    path: Path
    workbook_name: str
    sheet_names: list[str]
    error: str | None = None


@dataclass(frozen=True)
class ReferenceLookupConfig:
    sample_text: str = "FOREIGN:BBB.Id"
    reference_row_index: int = 6
    table_name: str = "BBB"
    field_name: str = "Id"
    field_row_index: int = 1


@dataclass(frozen=True)
class ExcelSheetReference:
    sheet_name: str
    reference_field_name: str
    source_field_name: str = ""


DEFAULT_REFERENCE_LOOKUP_CONFIG = ReferenceLookupConfig()


def scan_excel_workbooks(folder_path: str | Path, keyword: str = "") -> list[ExcelWorkbookInfo]:
    normalized_keyword = keyword.strip().lower()
    results: list[ExcelWorkbookInfo] = []

    for path in iter_excel_files(folder_path):
        try:
            sheet_names = get_sheet_names(path)
            error = None
        except ExcelReadError as exc:
            sheet_names = []
            error = str(exc)

        workbook = ExcelWorkbookInfo(
            path=path,
            workbook_name=path.name,
            sheet_names=sheet_names,
            error=error,
        )
        if _matches_keyword(workbook, normalized_keyword):
            results.append(workbook)

    return results


def find_sheet_references(
    file_path: str | Path,
    sheet_name: str,
    row_index: int = 6,
) -> list[str]:
    config = ReferenceLookupConfig(reference_row_index=row_index)
    matches = find_sheet_reference_matches(file_path, sheet_name, config)
    references: list[str] = []
    for match in matches:
        if match.sheet_name and match.sheet_name not in references:
            references.append(match.sheet_name)
    return references


def find_sheet_reference_matches(
    file_path: str | Path,
    sheet_name: str,
    config: ReferenceLookupConfig | None = None,
) -> list[ExcelSheetReference]:
    lookup_config = config or DEFAULT_REFERENCE_LOOKUP_CONFIG
    validate_reference_lookup_config(lookup_config)
    path = Path(file_path)
    if not path.exists():
        raise ExcelReadError(f"File does not exist: {path}")

    suffix = path.suffix.lower()
    if suffix in {".xlsx", ".xlsm", ".xltx", ".xltm"}:
        return _find_openpyxl_sheet_reference_matches(path, sheet_name, lookup_config)
    if suffix == ".xls":
        return _find_xlrd_sheet_reference_matches(path, sheet_name, lookup_config)
    if suffix == ".xlsb":
        return _find_xlsb_sheet_reference_matches(path, sheet_name, lookup_config)

    raise ExcelReadError(f"Unsupported Excel file type: {path.suffix}")


def validate_reference_lookup_config(config: ReferenceLookupConfig) -> None:
    _build_reference_pattern(config)


def _find_openpyxl_sheet_reference_matches(
    path: Path,
    sheet_name: str,
    config: ReferenceLookupConfig,
) -> list[ExcelSheetReference]:
    try:
        workbook = load_workbook(path, read_only=True, data_only=True)
    except Exception as exc:
        raise ExcelReadError(f"Could not open Excel file: {path}") from exc

    try:
        if sheet_name not in workbook.sheetnames:
            raise ExcelReadError(f"Sheet does not exist: {sheet_name}")

        worksheet = workbook[sheet_name]
        pattern = _build_reference_pattern(config)
        source_fields = _get_openpyxl_row_values(worksheet, config.field_row_index)
        references: list[ExcelSheetReference] = []
        for row in worksheet.iter_rows(
            min_row=config.reference_row_index,
            max_row=config.reference_row_index,
        ):
            for column_index, cell in enumerate(row, start=1):
                _add_unique_reference_matches(
                    references,
                    cell.value,
                    source_fields.get(column_index, ""),
                    pattern,
                )
        return references
    finally:
        workbook.close()


def _find_xlrd_sheet_reference_matches(
    path: Path,
    sheet_name: str,
    config: ReferenceLookupConfig,
) -> list[ExcelSheetReference]:
    try:
        import xlrd
    except ImportError as exc:
        raise ExcelReadError("xlrd is required to read old .xls files") from exc

    try:
        workbook = xlrd.open_workbook(str(path), on_demand=True)
    except Exception as exc:
        raise ExcelReadError(f"Could not open Excel file: {path}") from exc

    try:
        try:
            worksheet = workbook.sheet_by_name(sheet_name)
        except xlrd.XLRDError as exc:
            raise ExcelReadError(f"Sheet does not exist: {sheet_name}") from exc

        row_offset = config.reference_row_index - 1
        references: list[ExcelSheetReference] = []
        if worksheet.nrows <= row_offset:
            return references

        pattern = _build_reference_pattern(config)
        source_fields = _get_xlrd_row_values(worksheet, config.field_row_index)
        for column_offset in range(worksheet.ncols):
            _add_unique_reference_matches(
                references,
                worksheet.cell_value(row_offset, column_offset),
                source_fields.get(column_offset + 1, ""),
                pattern,
            )
        return references
    finally:
        workbook.release_resources()


def _find_xlsb_sheet_reference_matches(
    path: Path,
    sheet_name: str,
    config: ReferenceLookupConfig,
) -> list[ExcelSheetReference]:
    try:
        from pyxlsb import open_workbook
    except ImportError as exc:
        raise ExcelReadError("pyxlsb is required to read .xlsb files") from exc

    try:
        with open_workbook(str(path)) as workbook:
            if sheet_name not in workbook.sheets:
                raise ExcelReadError(f"Sheet does not exist: {sheet_name}")

            pattern = _build_reference_pattern(config)
            references: list[ExcelSheetReference] = []
            source_fields: dict[int, str] = {}
            reference_values: list[tuple[int, object]] = []
            last_needed_row = max(config.reference_row_index, config.field_row_index)
            with workbook.get_sheet(sheet_name) as worksheet:
                for current_row_index, row in enumerate(worksheet.rows(), start=1):
                    if current_row_index == config.field_row_index:
                        for column_index, cell in enumerate(row, start=1):
                            source_fields[column_index] = clean_cell_text(cell.v)
                    if current_row_index == config.reference_row_index:
                        reference_values = [
                            (column_index, cell.v)
                            for column_index, cell in enumerate(row, start=1)
                        ]
                    if current_row_index >= last_needed_row:
                        break

            for column_index, value in reference_values:
                _add_unique_reference_matches(
                    references,
                    value,
                    source_fields.get(column_index, ""),
                    pattern,
                )
            return references
    except ExcelReadError:
        raise
    except Exception as exc:
        raise ExcelReadError(f"Could not open Excel file: {path}") from exc


def _matches_keyword(workbook: ExcelWorkbookInfo, keyword: str) -> bool:
    if not keyword:
        return True

    searchable_parts = [
        workbook.workbook_name,
        workbook.path.stem,
        str(workbook.path),
        *workbook.sheet_names,
    ]
    return any(keyword in part.lower() for part in searchable_parts)


def _get_openpyxl_row_values(worksheet: object, row_index: int) -> dict[int, str]:
    values: dict[int, str] = {}
    for row in worksheet.iter_rows(min_row=row_index, max_row=row_index):
        for column_index, cell in enumerate(row, start=1):
            values[column_index] = clean_cell_text(cell.value)
    return values


def _get_xlrd_row_values(worksheet: object, row_index: int) -> dict[int, str]:
    row_offset = row_index - 1
    if worksheet.nrows <= row_offset:
        return {}

    return {
        column_offset + 1: clean_cell_text(worksheet.cell_value(row_offset, column_offset))
        for column_offset in range(worksheet.ncols)
    }


def _add_unique_reference_matches(
    references: list[ExcelSheetReference],
    value: object,
    source_field_name: str,
    pattern: re.Pattern[str],
) -> None:
    if value is None:
        return

    for match in pattern.finditer(str(value)):
        sheet_name = match.group("table").strip()
        reference_field_name = match.group("field").strip()
        if not sheet_name:
            continue

        reference = ExcelSheetReference(
            sheet_name=sheet_name,
            reference_field_name=reference_field_name,
            source_field_name=source_field_name.strip(),
        )
        if reference not in references:
            references.append(reference)


def _build_reference_pattern(config: ReferenceLookupConfig) -> re.Pattern[str]:
    sample_text = config.sample_text.strip()
    table_name = config.table_name.strip()
    field_name = config.field_name.strip()

    if not sample_text:
        raise ExcelReadError("引用样例不能为空")
    if not table_name:
        raise ExcelReadError("引用表格名称不能为空")
    if not field_name:
        raise ExcelReadError("引用字段不能为空")
    if config.reference_row_index < 1:
        raise ExcelReadError("引用行必须大于等于 1")
    if config.field_row_index < 1:
        raise ExcelReadError("引用字段行必须大于等于 1")

    table_start = sample_text.find(table_name)
    field_start = sample_text.find(field_name)
    if table_start < 0:
        raise ExcelReadError("引用样例中找不到填写的引用表格名称")
    if field_start < 0:
        raise ExcelReadError("引用样例中找不到填写的引用字段")

    spans = [
        ("table", table_start, len(table_name)),
        ("field", field_start, len(field_name)),
    ]
    spans.sort(key=lambda item: item[1])
    first_end = spans[0][1] + spans[0][2]
    if first_end > spans[1][1]:
        raise ExcelReadError("引用表格名称和引用字段在样例中不能重叠")

    pattern_parts: list[str] = []
    cursor = 0
    for index, (group_name, start, length) in enumerate(spans):
        pattern_parts.append(re.escape(sample_text[cursor:start]))
        next_start = spans[index + 1][1] if index + 1 < len(spans) else len(sample_text)
        next_literal = sample_text[start + length:next_start]
        pattern_parts.append(f"(?P<{group_name}>.+?)" if next_literal else f"(?P<{group_name}>[^\\s,;，；]+)")
        cursor = start + length
    pattern_parts.append(re.escape(sample_text[cursor:]))

    return re.compile("".join(pattern_parts), re.IGNORECASE)
