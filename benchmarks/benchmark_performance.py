"""同机性能复测：python benchmarks/benchmark_performance.py。不会读取或修改用户配置。"""

import argparse
from pathlib import Path
import sys
import tempfile
from time import perf_counter
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import xlsxwriter

from cell_search import DataFilterConfig, _has_typed_shared_string_reference, search_excel_cells
import cell_search


def main() -> None:
    parser = argparse.ArgumentParser(description="默认只运行不创建窗口的检索基准。")
    parser.add_argument("--gui", action="store_true", help="显式启用会创建窗口的界面基准")
    args = parser.parse_args()
    xml = b'<c r="A1" t="s"><v>0</v></c>' * 40000
    identifiers = {str(index).encode() for index in range(1, 5001)}
    started = perf_counter()
    assert not _has_typed_shared_string_reference(xml, identifiers)
    print(f"共享字符串预筛：{perf_counter() - started:.4f} 秒")

    with tempfile.TemporaryDirectory() as directory:
        workbook = xlsxwriter.Workbook(Path(directory) / "search.xlsx")
        misses = workbook.add_worksheet("NoMatch")
        hits = workbook.add_worksheet("Hits")
        for row in range(40000):
            misses.write(row, 0, f"其他中文配置{row}")
        misses.write_formula("B1", '="说明"&"其他"', None, "普通说明")
        number_format = workbook.add_format({"num_format": "0.00"})
        for row in range(1000):
            misses.write_number(row, 2, row / 100, number_format)
        for row in range(10000):
            hits.write(row, 0, f"needle-{row}")
        workbook.close()
        started = perf_counter()
        results = search_excel_cells(directory, "needle")
        assert len(results) == 10000
        print(f"50,000 单元格工作簿完整检索：{perf_counter() - started:.4f} 秒")
        for keyword in ("不存在的技能展示", "missing1001001", "864208642"):
            started = perf_counter()
            with patch("cell_search._search_sheet_xml_cells", wraps=cell_search._search_sheet_xml_cells) as parse:
                results = search_excel_cells(directory, keyword, data_filter_config=DataFilterConfig(enable_header_filter=True, header_row_count=6))
            assert not results and parse.call_count == 0
            print(f"中文/公式/数据过滤场景未命中 {keyword!r}：{perf_counter() - started:.4f} 秒；完整解析 {parse.call_count} 张工作表")

    if not args.gui:
        print("界面基准已跳过；如需测量界面绘制，请显式添加 --gui。")
        return

    from file_search import FileSearchResult
    from gui import ExcelEditApp

    with patch("gui.load_cache_data", return_value={}):
        app = ExcelEditApp()
    app.withdraw()
    app._save_cache_from_inputs = lambda: None
    try:
        for count in (500, 20000):
            results = [FileSearchResult(Path(f"E:/benchmark/{index}.txt"), f"{index}.txt", "") for index in range(count)]
            started = perf_counter()
            app._render_file_results(results)
            app.update_idletasks()
            print(f"{count:,} 条文件结果首屏：{perf_counter() - started:.4f} 秒；绘制 {len(app.file_thumbnail_cards)} 张卡片")
            started = perf_counter()
            app.file_thumbnail_frame.canvas.yview_moveto(1)
            app.update_idletasks()
            assert results[-1] in app.file_thumbnail_cards
            print(f"滚动到底部：{perf_counter() - started:.4f} 秒")
    finally:
        app.close_app()


if __name__ == "__main__":
    main()
