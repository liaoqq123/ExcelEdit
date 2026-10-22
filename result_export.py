"""将检索结果以纯文本单元格导出，完成写入后再替换目标文件。"""

import os
import tempfile
from collections.abc import Iterable, Sequence
from contextlib import suppress
from pathlib import Path

from openpyxl import Workbook
from openpyxl.cell import WriteOnlyCell
from openpyxl.cell.cell import Cell
from openpyxl.worksheet._write_only import WriteOnlyWorksheet


EXCEL_MAX_ROWS = 1_048_576
EXCEL_MAX_COLUMNS = 16_384
EXCEL_MAX_TEXT_LENGTH = 32_767


def export_result_rows(
    path: Path,
    headers: Sequence[str],
    rows: Iterable[Sequence[object]],
) -> int:
    """流式导出 XLSX，返回数据行数；失败时保留已有文件并抛出异常。"""

    destination = Path(path)
    temporary_path: Path | None = None
    workbook = Workbook(write_only=True)
    worksheet = workbook.create_sheet("检索结果")
    worksheet.freeze_panes = "A2"
    try:
        # 同目录临时文件保证替换发生在同一文件系统，Windows 下先关闭临时句柄。
        with tempfile.NamedTemporaryFile(
            dir=destination.parent,
            prefix=f".{destination.name}.",
            suffix=".tmp",
            delete=False,
        ) as temporary_file:
            temporary_path = Path(temporary_file.name)

        worksheet.append(_text_cells(worksheet, headers))
        count = 0
        for row in rows:
            if count >= EXCEL_MAX_ROWS - 1:
                raise ValueError("导出结果超过 Excel 单个工作表的行数上限。")
            worksheet.append(_text_cells(worksheet, row))
            count += 1

        with temporary_path.open("wb") as output:
            workbook.save(output)
            output.flush()
            os.fsync(output.fileno())
        os.replace(temporary_path, destination)
        return count
    finally:
        # write_only 会使用内部 XML 临时文件，迭代器或保存失败时也要关闭并清理。
        with suppress(Exception):
            if not worksheet.closed:
                worksheet.close()
        writer = getattr(worksheet, "_writer", None)
        if writer is not None and Path(writer.out).exists():
            with suppress(OSError, ValueError):
                writer.cleanup()
        with suppress(Exception):
            workbook.close()
        if temporary_path is not None:
            temporary_path.unlink(missing_ok=True)


def _text_cells(worksheet: WriteOnlyWorksheet, values: Sequence[object]) -> list[Cell]:
    if len(values) > EXCEL_MAX_COLUMNS:
        raise ValueError("导出结果超过 Excel 单个工作表的列数上限。")
    cells = []
    for value in values:
        text = "" if value is None else str(value)
        if len(text) > EXCEL_MAX_TEXT_LENGTH:
            raise ValueError("单元格内容超过 Excel 的 32767 字符上限，无法完整导出。")
        cell = WriteOnlyCell(worksheet, value=text)
        # 显式指定字符串类型；包括以 =、+、-、@ 开头的原始内容。
        cell.data_type = "s"
        cells.append(cell)
    return cells
