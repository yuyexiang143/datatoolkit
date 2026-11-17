from __future__ import annotations

import subprocess
import sys
from pathlib import Path

import pandas as pd


PROJECT_ROOT = Path(__file__).resolve().parents[1]


def run_module(*args: str) -> subprocess.CompletedProcess[str]:
    return subprocess.run(
        [sys.executable, "-m", "datatoolkit", *args],
        cwd=PROJECT_ROOT,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        check=False,
    )


def test_cli_help_lists_three_commands() -> None:
    result = run_module("--help")

    assert result.returncode == 0
    assert {"clean", "merge", "report"} <= set(result.stdout.split())


def test_cli_clean_writes_cleaned_workbook(tmp_path: Path) -> None:
    source = tmp_path / "source.csv"
    output = tmp_path / "clean.xlsx"
    pd.DataFrame(
        {"客户": [" 云海 ", " 云海 "], "日期": ["2024/1/5", "2024/1/5"], "金额": ["¥ 10", "¥ 10"]}
    ).to_csv(source, index=False, encoding="utf-8-sig")

    result = run_module(
        "clean",
        str(source),
        "-o",
        str(output),
        "--drop-duplicates",
        "--date-columns",
        "日期",
        "--amount-columns",
        "金额",
    )

    assert result.returncode == 0, result.stderr
    assert output.exists()
    cleaned = pd.read_excel(output)
    assert cleaned.to_dict("records") == [{"客户": "云海", "日期": "2024-01-05", "金额": 10.0}]
