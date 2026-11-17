from __future__ import annotations

from pathlib import Path

import pandas as pd

from datatoolkit.core import CleaningOptions
from datatoolkit.report import build_report


def test_report_has_summary_grouped_values_and_anomalies(tmp_path: Path) -> None:
    input_path = tmp_path / "dirty.csv"
    pd.DataFrame(
        {
            "订单号": ["A1", "A1", "A2", "A3"],
            "日期": ["2024/1/5", "2024/1/5", "2024-01-06", "2024-13-01"],
            "地区": ["华东", "华东", "华南", None],
            "金额": ["¥ 100", "¥ 100", "待确认", "200"],
        }
    ).to_csv(input_path, index=False, encoding="utf-8-sig")
    output_path = tmp_path / "report.xlsx"
    options = CleaningOptions(
        trim_whitespace=True,
        drop_duplicates=True,
        dedupe_columns=("订单号",),
        null_strategy="fill",
        fill_value="未知",
        date_columns=("日期",),
        amount_columns=("金额",),
    )

    result = build_report(
        input_path,
        output_path,
        options,
        group_column="地区",
        sum_column="金额",
    )

    summary = pd.read_excel(output_path, sheet_name="汇总")
    grouped = pd.read_excel(output_path, sheet_name="分组汇总")
    anomalies = pd.read_excel(output_path, sheet_name="异常值")
    summary_values = dict(zip(summary["指标"], summary["值"]))

    assert output_path.exists()
    assert result.original_rows == 4
    assert result.cleaned_rows == 3
    assert int(summary_values["原始行数"]) == 4
    assert int(summary_values["清洗后行数"]) == 3
    assert grouped.loc[grouped["地区"] == "华东", "金额_求和"].iloc[0] == 100
    assert grouped.loc[grouped["地区"] == "华东", "记录数"].iloc[0] == 1
    assert set(anomalies["问题"]) >= {"重复行", "无效日期", "无效金额", "空值"}
