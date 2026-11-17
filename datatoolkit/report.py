"""Excel summary reports built from cleaned table data."""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Iterable

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill

from datatoolkit.core import CleaningOptions, clean_amount, clean_dataframe, read_table


@dataclass(frozen=True)
class ReportResult:
    """Counts returned after a report workbook has been written."""

    output_path: Path
    original_rows: int
    cleaned_rows: int
    grouped_rows: int
    anomaly_count: int


def _is_blank(value: object) -> bool:
    """Return whether a cell should be treated as blank."""

    return pd.isna(value) or (isinstance(value, str) and not value.strip())


def _is_valid_date(value: object) -> bool:
    """Check one date value against the cleaning formats."""

    if _is_blank(value):
        return True
    if isinstance(value, (pd.Timestamp, datetime, date)):
        return True
    text = str(value).strip()
    formats = (
        "%Y-%m-%d",
        "%Y/%m/%d",
        "%Y.%m.%d",
        "%Y年%m月%d日",
        "%m/%d/%Y",
        "%d-%m-%Y",
        "%Y%m%d",
    )
    if any(not pd.isna(pd.to_datetime(text, format=fmt, errors="coerce")) for fmt in formats):
        return True
    return not pd.isna(pd.to_datetime(text, errors="coerce"))


def _duplicate_mask(
    frame: pd.DataFrame,
    dedupe_columns: Iterable[str],
) -> pd.Series:
    """Mark duplicate rows using the configured comparison columns."""

    compare_frame = frame.copy()
    for column in compare_frame.columns:
        compare_frame[column] = compare_frame[column].map(
            lambda value: value.strip() if isinstance(value, str) else value
        )
    subset = list(dedupe_columns) or None
    return compare_frame.duplicated(subset=subset, keep="first")


def collect_anomalies(
    frame: pd.DataFrame,
    options: CleaningOptions,
) -> pd.DataFrame:
    """List blank, invalid-date, invalid-amount, and duplicate rows.

    Row numbers use Excel conventions: the first data row is row 2.
    """

    anomalies: list[dict[str, object]] = []
    duplicate_rows = (
        _duplicate_mask(frame, options.dedupe_columns)
        if options.drop_duplicates
        else pd.Series(False, index=frame.index)
    )

    for index, row in frame.iterrows():
        excel_row = int(index) + 2
        for column in frame.columns:
            value = row[column]
            if _is_blank(value):
                anomalies.append(
                    {
                        "行号": excel_row,
                        "字段": column,
                        "原始值": "",
                        "问题": "空值",
                    }
                )
            elif column in options.date_columns and not _is_valid_date(value):
                anomalies.append(
                    {
                        "行号": excel_row,
                        "字段": column,
                        "原始值": str(value),
                        "问题": "无效日期",
                    }
                )
            elif column in options.amount_columns and pd.isna(clean_amount(value)):
                anomalies.append(
                    {
                        "行号": excel_row,
                        "字段": column,
                        "原始值": str(value),
                        "问题": "无效金额",
                    }
                )

        if bool(duplicate_rows.loc[index]):
            anomalies.append(
                {
                    "行号": excel_row,
                    "字段": "整行",
                    "原始值": " | ".join(str(value) for value in row.tolist()),
                    "问题": "重复行",
                }
            )

    return pd.DataFrame(
        anomalies,
        columns=["行号", "字段", "原始值", "问题"],
    )


def _group_summary(
    frame: pd.DataFrame,
    group_column: str,
    sum_column: str,
) -> pd.DataFrame:
    """Calculate record counts and a numeric sum for each group."""

    working = frame.copy()
    working["__求和值"] = pd.to_numeric(working[sum_column], errors="coerce").fillna(0)
    grouped = (
        working.groupby(group_column, dropna=False)
        .agg(
            记录数=(group_column, "size"),
            金额_求和=("__求和值", "sum"),
        )
        .reset_index()
    )
    return grouped


def _style_workbook(path: Path) -> None:
    """Apply compact portfolio-friendly formatting to all report sheets."""

    workbook = load_workbook(path)
    header_fill = PatternFill("solid", fgColor="1F4E78")
    header_font = Font(color="FFFFFF", bold=True)
    for worksheet in workbook.worksheets:
        worksheet.freeze_panes = "A2"
        worksheet.auto_filter.ref = worksheet.dimensions
        for cell in worksheet[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center")
        for letter, cells in worksheet.column_dimensions.items():
            values = [str(cell.value or "") for cell in worksheet[letter]]
            width = min(max(len(value) for value in values) + 4, 45)
            cells.width = max(width, 12)
    workbook.save(path)


def build_report(
    input_path: str | Path,
    output_path: str | Path,
    options: CleaningOptions,
    group_column: str,
    sum_column: str,
    sheet: str | int | None = None,
) -> ReportResult:
    """Clean a table and write summary, grouped totals, and anomalies to Excel."""

    raw = read_table(input_path, sheet=sheet)
    from datatoolkit.core import _validate_columns

    _validate_columns(raw, (group_column, sum_column), "报表")
    anomalies = collect_anomalies(raw, options)
    cleaned = clean_dataframe(raw, options)
    grouped = _group_summary(cleaned, group_column, sum_column)

    summary = pd.DataFrame(
        {
            "指标": [
                "原始行数",
                "清洗后行数",
                "删除行数",
                "列数",
                "分组列",
                "求和列",
                "异常记录数",
            ],
            "值": [
                len(raw),
                len(cleaned),
                len(raw) - len(cleaned),
                len(raw.columns),
                group_column,
                sum_column,
                len(anomalies),
            ],
        }
    )

    destination = Path(output_path).expanduser()
    destination.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(destination, engine="openpyxl") as writer:
        summary.to_excel(writer, sheet_name="汇总", index=False)
        grouped.to_excel(writer, sheet_name="分组汇总", index=False)
        anomalies.to_excel(writer, sheet_name="异常值", index=False)
    _style_workbook(destination)

    return ReportResult(
        output_path=destination,
        original_rows=len(raw),
        cleaned_rows=len(cleaned),
        grouped_rows=len(grouped),
        anomaly_count=len(anomalies),
    )
