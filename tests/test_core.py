from __future__ import annotations

from pathlib import Path

import pandas as pd
import pytest

from datatoolkit.core import (
    CleaningOptions,
    clean_dataframe,
    merge_directory,
    read_table,
    write_table,
)


def test_read_xlsx_supports_selected_sheet(tmp_path: Path) -> None:
    path = tmp_path / "book.xlsx"
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        pd.DataFrame({"值": ["错误页"]}).to_excel(writer, sheet_name="概览", index=False)
        pd.DataFrame({"值": ["目标页"]}).to_excel(writer, sheet_name="明细", index=False)

    result = read_table(path, sheet="明细")

    assert result["值"].tolist() == ["目标页"]


def test_read_csv_detects_gbk_encoding(tmp_path: Path) -> None:
    path = tmp_path / "gbk.csv"
    path.write_bytes("客户,金额\n云海,100\n".encode("gbk"))

    result = read_table(path)

    assert result.loc[0, "客户"] == "云海"


def test_drop_duplicates_can_use_selected_columns() -> None:
    frame = pd.DataFrame(
        {
            "订单号": ["A1", "A1", "A2"],
            "日期": ["2024-01-01", "2024-01-02", "2024-01-03"],
            "金额": ["10", "20", "30"],
        }
    )

    result = clean_dataframe(
        frame,
        CleaningOptions(drop_duplicates=True, dedupe_columns=("订单号",)),
    )

    assert result["订单号"].tolist() == ["A1", "A2"]
    assert result["日期"].tolist() == ["2024-01-01", "2024-01-03"]


def test_null_strategy_drop_removes_rows_with_missing_values() -> None:
    frame = pd.DataFrame({"客户": ["甲", None, "丙"], "金额": [10, 20, None]})

    result = clean_dataframe(frame, CleaningOptions(null_strategy="drop"))

    assert result["客户"].tolist() == ["甲"]


def test_null_strategy_fill_uses_configured_value() -> None:
    frame = pd.DataFrame({"客户": ["甲", None], "金额": [10, 20]})

    result = clean_dataframe(
        frame,
        CleaningOptions(null_strategy="fill", fill_value="未知"),
    )

    assert result["客户"].tolist() == ["甲", "未知"]


def test_null_strategy_ffill_uses_previous_row() -> None:
    frame = pd.DataFrame({"客户": ["甲", None, None], "金额": [10, 20, 30]})

    result = clean_dataframe(frame, CleaningOptions(null_strategy="ffill"))

    assert result["客户"].tolist() == ["甲", "甲", "甲"]


def test_dates_and_amounts_are_normalized() -> None:
    frame = pd.DataFrame(
        {
            "日期": ["2024/1/5", "2024年2月3日", "20240315"],
            "金额": ["¥ 1,200.50", "￥880", "2,350"],
        }
    )

    result = clean_dataframe(
        frame,
        CleaningOptions(date_columns=("日期",), amount_columns=("金额",)),
    )

    assert result["日期"].tolist() == ["2024-01-05", "2024-02-03", "2024-03-15"]
    assert result["金额"].tolist() == [1200.5, 880.0, 2350.0]
    assert pd.api.types.is_float_dtype(result["金额"])


def test_string_columns_are_trimmed() -> None:
    frame = pd.DataFrame({"客户": [" 云海 ", " 星河  "], "备注": [None, " 加急 "]})

    result = clean_dataframe(frame, CleaningOptions(trim_whitespace=True))

    assert result["客户"].tolist() == ["云海", "星河"]
    assert result.loc[1, "备注"] == "加急"


def test_merge_directory_reads_same_structure_and_adds_source_filename(
    tmp_path: Path,
) -> None:
    write_table(pd.DataFrame({"客户": ["甲"], "金额": [10]}), tmp_path / "a.csv")
    write_table(pd.DataFrame({"客户": ["乙"], "金额": [20]}), tmp_path / "b.csv")

    result = merge_directory(tmp_path)

    assert result.columns.tolist() == ["来源文件名", "客户", "金额"]
    assert result["来源文件名"].tolist() == ["a.csv", "b.csv"]
    assert result["客户"].tolist() == ["甲", "乙"]


def test_merge_directory_rejects_mismatched_columns(tmp_path: Path) -> None:
    write_table(pd.DataFrame({"客户": ["甲"]}), tmp_path / "a.csv")
    write_table(pd.DataFrame({"姓名": ["乙"]}), tmp_path / "b.csv")

    with pytest.raises(ValueError, match="结构"):
        merge_directory(tmp_path)


def test_invalid_date_and_amount_become_missing_values() -> None:
    frame = pd.DataFrame({"日期": ["不是日期"], "金额": ["待确认"]})

    result = clean_dataframe(
        frame,
        CleaningOptions(date_columns=("日期",), amount_columns=("金额",)),
    )

    assert pd.isna(result.loc[0, "日期"])
    assert pd.isna(result.loc[0, "金额"])
