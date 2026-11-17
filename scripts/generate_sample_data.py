"""Regenerate the three intentionally dirty demonstration files."""

from __future__ import annotations

from pathlib import Path

import pandas as pd
import xlwt


PROJECT_ROOT = Path(__file__).resolve().parents[1]
SAMPLE_DIR = PROJECT_ROOT / "sample_data"
HEADERS = ["订单号", "日期", "客户", "地区", "金额", "备注"]


def _write_xls(path: Path, rows: list[list[object]]) -> None:
    """Write a legacy Excel 97-2003 workbook using xlwt."""

    workbook = xlwt.Workbook(encoding="utf-8")
    sheet = workbook.add_sheet("销售数据")
    for column, header in enumerate(HEADERS):
        sheet.write(0, column, header)
    for row_index, row in enumerate(rows, start=1):
        for column, value in enumerate(row):
            sheet.write(row_index, column, value)
    workbook.save(str(path))


def main() -> None:
    """Create dirty XLSX, XLS, and GBK CSV files in ``sample_data``."""

    SAMPLE_DIR.mkdir(parents=True, exist_ok=True)

    east_rows = [
        ["A001", "2024/1/5", " 华东科技 ", "华东", "¥ 1,200.50", "首单"],
        ["A001", "2024/1/5", "华东科技", "华东", "¥ 1,200.50", "首单"],
        ["A002", "2024-01-16", "星河商贸", "华东", "￥880", ""],
        ["A003", "2024年2月3日", "远山文创", "华东", "", "缺金额"],
        ["A004", "20240315", " 青禾供应链", "华东 ", "2,350", None],
    ]
    with pd.ExcelWriter(SAMPLE_DIR / "region_east.xlsx", engine="openpyxl") as writer:
        pd.DataFrame(east_rows, columns=HEADERS).to_excel(
            writer, sheet_name="销售数据", index=False
        )
        pd.DataFrame(
            {"说明": ["这是故意制作的脏数据，用于演示日期、金额、空值和空格清洗。"]}
        ).to_excel(writer, sheet_name="说明", index=False)

    west_rows = [
        ["W001", "2024.05.01", "云海商贸", "西部", "¥ 430.00", "重复检查"],
        ["W002", "2024/6/12", " 北岸贸易 ", " 西部 ", "1,005", ""],
        ["W003", "2024-13-01", " 云海商贸", "西部", "待确认", "异常日期和金额"],
        ["W004", "07/08/2024", "木棉工作室", "西部", "（999）", "括号金额"],
        ["W005", "2024/8/20", "北岸贸易", "西部", "3,200 ", None],
    ]
    _write_xls(SAMPLE_DIR / "region_west.xls", west_rows)

    south_frame = pd.DataFrame(
        [
            ["C001", "2024年7月1日", "青禾公司", "华南", "￥2,800.90", "  紧急 "],
            ["C002", "2024-07-03", "青禾公司", "华南", "￥1,300", ""],
            ["C003", "", " 南岭文创 ", "华南 ", "700.00", None],
            ["C004", "2024/7/15", "青禾公司", "华南", "1,150", "加急"],
        ],
        columns=HEADERS,
    )
    south_frame.to_csv(
        SAMPLE_DIR / "region_south_gbk.csv",
        index=False,
        encoding="gbk",
    )

    print(f"已生成 {len(list(SAMPLE_DIR.iterdir()))} 个样例文件：{SAMPLE_DIR}")


if __name__ == "__main__":
    main()
