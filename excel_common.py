from pathlib import Path
from typing import Iterable, Sequence

from openpyxl import load_workbook
from xlsxwriter import Workbook


EXCEL_EXTENSIONS = {".xlsx", ".xlsm", ".xltx", ".xltm", ".xls", ".xlsb"}


class ExcelReadError(RuntimeError):
    pass


def get_sheet_names(file_path: str | Path) -> list[str]:
    path = Path(file_path)
    if not path.exists():
        raise ExcelReadError(f"File does not exist: {path}")

    suffix = path.suffix.lower()
    if suffix in {".xlsx", ".xlsm", ".xltx", ".xltm"}:
        return _get_openpyxl_sheet_names(path)
    if suffix == ".xls":
        return _get_xlrd_sheet_names(path)
    if suffix == ".xlsb":
        return _get_xlsb_sheet_names(path)

    raise ExcelReadError(f"Unsupported Excel file type: {path.suffix}")


def iter_excel_files(folder_path: str | Path) -> list[Path]:
    folder = validate_excel_folder(folder_path)
    return [
        path
        for path in sorted(folder.rglob("*"), key=lambda item: str(item).lower())
        if is_excel_file(path)
    ]


def validate_excel_folder(folder_path: str | Path) -> Path:
    folder = Path(folder_path)
    if not folder.exists():
        raise ExcelReadError(f"Folder does not exist: {folder}")
    if not folder.is_dir():
        raise ExcelReadError(f"Not a folder: {folder}")
    return folder


def is_excel_file(path: Path) -> bool:
    if not path.is_file():
        return False
    if path.name.startswith("~$"):
        return False
    return path.suffix.lower() in EXCEL_EXTENSIONS


def clean_cell_text(value: object) -> str:
    if value is None:
        return ""
    return str(value).strip()


def read_excel_preview(
    file_path: str | Path,
    sheet_name: str | None = None,
    max_rows: int = 20,
) -> list[list[object]]:
    path = Path(file_path)
    if not path.exists():
        raise ExcelReadError(f"File does not exist: {path}")

    try:
        workbook = load_workbook(path, read_only=True, data_only=True)
    except Exception as exc:
        raise ExcelReadError(f"Could not open Excel file: {path}") from exc

    try:
        worksheet = workbook[sheet_name] if sheet_name else workbook.active
        rows: list[list[object]] = []
        for row in worksheet.iter_rows(max_row=max_rows, values_only=True):
            rows.append(list(row))
        return rows
    finally:
        workbook.close()


def write_rows_to_excel(
    file_path: str | Path,
    rows: Iterable[Sequence[object]],
    sheet_name: str = "Sheet1",
) -> None:
    path = Path(file_path)
    path.parent.mkdir(parents=True, exist_ok=True)

    workbook = Workbook(str(path))
    worksheet = workbook.add_worksheet(sheet_name[:31] or "Sheet1")

    for row_index, row in enumerate(rows):
        for column_index, value in enumerate(row):
            worksheet.write(row_index, column_index, value)

    workbook.close()


def _get_openpyxl_sheet_names(path: Path) -> list[str]:
    try:
        workbook = load_workbook(path, read_only=True, data_only=True)
    except Exception as exc:
        raise ExcelReadError(f"Could not open Excel file: {path}") from exc

    try:
        return list(workbook.sheetnames)
    finally:
        workbook.close()


def _get_xlrd_sheet_names(path: Path) -> list[str]:
    try:
        import xlrd
    except ImportError as exc:
        raise ExcelReadError("xlrd is required to read old .xls files") from exc

    try:
        workbook = xlrd.open_workbook(str(path), on_demand=True)
    except Exception as exc:
        raise ExcelReadError(f"Could not open Excel file: {path}") from exc

    try:
        return list(workbook.sheet_names())
    finally:
        workbook.release_resources()


def _get_xlsb_sheet_names(path: Path) -> list[str]:
    try:
        from pyxlsb import open_workbook
    except ImportError as exc:
        raise ExcelReadError("pyxlsb is required to read .xlsb files") from exc

    try:
        with open_workbook(str(path)) as workbook:
            return list(workbook.sheets)
    except Exception as exc:
        raise ExcelReadError(f"Could not open Excel file: {path}") from exc
