"""Table input, cleaning, directory merging, and output helpers."""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Literal, Sequence

import numpy as np
import pandas as pd


SUPPORTED_INPUT_SUFFIXES = {".xlsx", ".xls", ".csv"}
SOURCE_COLUMN = "来源文件名"
NullStrategy = Literal["keep", "drop", "fill", "ffill"]


class DataToolkitError(ValueError):
    """Raised when a table cannot be read, cleaned, merged, or written."""


@dataclass(frozen=True)
class CleaningOptions:
    """Options shared by the command line, GUI, and report layers.

    Args:
        trim_whitespace: Remove leading and trailing spaces from string cells.
        drop_duplicates: Remove duplicate rows after other cleaning steps.
        dedupe_columns: Columns used for duplicate detection. Empty means all.
        null_strategy: Keep, delete, fill, or forward-fill missing values.
        fill_value: Replacement used when ``null_strategy`` is ``"fill"``.
        date_columns: Columns to normalize to ``YYYY-MM-DD``.
        amount_columns: Columns to parse as numeric currency amounts.
    """

    trim_whitespace: bool = True
    drop_duplicates: bool = False
    dedupe_columns: tuple[str, ...] = ()
    null_strategy: NullStrategy = "keep"
    fill_value: object = ""
    date_columns: tuple[str, ...] = ()
    amount_columns: tuple[str, ...] = ()


def _validate_columns(
    frame: pd.DataFrame,
    columns: Sequence[str],
    role: str,
) -> None:
    """Raise a descriptive error if any requested column is absent."""

    missing = [column for column in columns if column not in frame.columns]
    if missing:
        available = "、".join(str(column) for column in frame.columns)
        raise DataToolkitError(
            f"{role}列不存在：{'、'.join(missing)}；当前可用列：{available}"
        )


def _normalize_date(value: object) -> object:
    """Convert one supported date value to ISO text, or return a missing value."""

    if pd.isna(value):
        return pd.NA
    if isinstance(value, pd.Timestamp):
        return value.strftime("%Y-%m-%d")
    if isinstance(value, (datetime, date)):
        return value.strftime("%Y-%m-%d")

    text = str(value).strip()
    if not text:
        return pd.NA

    formats = (
        "%Y-%m-%d",
        "%Y/%m/%d",
        "%Y.%m.%d",
        "%Y年%m月%d日",
        "%m/%d/%Y",
        "%d-%m-%Y",
        "%Y%m%d",
    )
    for date_format in formats:
        parsed = pd.to_datetime(text, format=date_format, errors="coerce")
        if not pd.isna(parsed):
            return parsed.strftime("%Y-%m-%d")

    parsed = pd.to_datetime(text, errors="coerce")
    if pd.isna(parsed):
        return pd.NA
    return parsed.strftime("%Y-%m-%d")


def clean_amount(value: object) -> float:
    """Remove common currency decorations and convert a value to ``float``.

    Blank or unparseable values become ``numpy.nan`` so the normal null-value
    strategy can process them and the report can list them separately.
    """

    if pd.isna(value):
        return float("nan")

    text = str(value).strip()
    if not text:
        return float("nan")

    negative = text.startswith("(") and text.endswith(")")
    for token in ("¥", "￥", ",", "，", " "):
        text = text.replace(token, "")
    if negative:
        text = text.strip("()")
    if text in {"", "-", "--"}:
        return float("nan")

    try:
        amount = float(text)
    except ValueError:
        return float("nan")
    return -amount if negative else amount


def _trim_string_cells(frame: pd.DataFrame) -> pd.DataFrame:
    """Return a copy with leading and trailing spaces removed from strings."""

    result = frame.copy()
    for column in result.columns:
        result[column] = result[column].map(
            lambda value: value.strip() if isinstance(value, str) else value
        )
    return result


def clean_dataframe(
    frame: pd.DataFrame,
    options: CleaningOptions | None = None,
) -> pd.DataFrame:
    """Apply the configured cleaning rules without modifying the input frame.

    The order is intentionally stable: whitespace, amount, date, null handling,
    then deduplication. This makes outputs reproducible across the CLI and GUI.
    """

    active = options or CleaningOptions()
    date_columns = tuple(active.date_columns)
    amount_columns = tuple(active.amount_columns)
    dedupe_columns = tuple(active.dedupe_columns)
    _validate_columns(frame, date_columns, "日期")
    _validate_columns(frame, amount_columns, "金额")
    _validate_columns(frame, dedupe_columns, "去重")

    result = frame.copy()
    if active.trim_whitespace:
        result = _trim_string_cells(result)

    for column in amount_columns:
        result[column] = result[column].map(clean_amount).astype(float)

    for column in date_columns:
        result[column] = result[column].map(_normalize_date)
        result[column] = result[column].astype("object")

    if active.null_strategy == "drop":
        result = result.dropna(how="any")
    elif active.null_strategy == "fill":
        result = result.fillna(active.fill_value)
    elif active.null_strategy == "ffill":
        result = result.ffill()
    elif active.null_strategy != "keep":
        raise DataToolkitError(f"不支持的空值策略：{active.null_strategy}")

    if active.drop_duplicates:
        subset = list(dedupe_columns) or None
        result = result.drop_duplicates(subset=subset, keep="first")

    return result.reset_index(drop=True)


def _read_csv(path: Path) -> pd.DataFrame:
    """Read a CSV using UTF-8 (with BOM) before falling back to GBK."""

    errors: list[str] = []
    for encoding in ("utf-8-sig", "gbk"):
        try:
            return pd.read_csv(path, encoding=encoding)
        except UnicodeDecodeError:
            errors.append(encoding)
        except pd.errors.EmptyDataError as exc:
            raise DataToolkitError(f"CSV 文件为空：{path}") from exc
        except pd.errors.ParserError as exc:
            raise DataToolkitError(f"CSV 内容无法解析：{path}") from exc
    raise DataToolkitError(
        f"CSV 编码识别失败，已尝试 {'、'.join(errors)}：{path}"
    )


def read_table(path: str | Path, sheet: str | int | None = None) -> pd.DataFrame:
    """Read ``.xlsx``, ``.xls``, or ``.csv`` into a DataFrame.

    Args:
        path: Input table path.
        sheet: Optional Excel worksheet name or zero-based index.

    Raises:
        FileNotFoundError: If the input path does not exist.
        DataToolkitError: If the format is unsupported or the file cannot parse.
    """

    source = Path(path).expanduser()
    if not source.exists():
        raise FileNotFoundError(f"文件不存在：{source}")
    if not source.is_file():
        raise DataToolkitError(f"输入路径不是文件：{source}")

    suffix = source.suffix.lower()
    if suffix not in SUPPORTED_INPUT_SUFFIXES:
        supported = "、".join(sorted(SUPPORTED_INPUT_SUFFIXES))
        raise DataToolkitError(f"不支持的文件格式 {suffix}，仅支持：{supported}")

    try:
        if suffix == ".csv":
            if sheet is not None:
                raise DataToolkitError("CSV 文件不能指定工作表")
            return _read_csv(source)

        return pd.read_excel(source, sheet_name=sheet if sheet is not None else 0)
    except ImportError as exc:
        raise DataToolkitError(
            "读取该 Excel 格式需要额外依赖，请先安装 requirements.txt"
        ) from exc
    except ValueError as exc:
        if isinstance(exc, DataToolkitError):
            raise
        raise DataToolkitError(f"无法读取表格 {source}：{exc}") from exc


def write_table(
    frame: pd.DataFrame,
    path: str | Path,
) -> Path:
    """Write a DataFrame to ``.xlsx`` or UTF-8 BOM CSV and return the path."""

    destination = Path(path).expanduser()
    destination.parent.mkdir(parents=True, exist_ok=True)
    suffix = destination.suffix.lower()

    if suffix == ".xlsx":
        frame.to_excel(destination, index=False, engine="openpyxl")
    elif suffix == ".csv":
        frame.to_csv(destination, index=False, encoding="utf-8-sig")
    else:
        raise DataToolkitError("输出仅支持 .xlsx 或 .csv")
    return destination


def merge_directory(
    directory: str | Path,
    pattern: str = "*",
    sheet: str | int | None = None,
) -> pd.DataFrame:
    """Merge matching tables with identical schemas and add source filenames.

    Files are sorted by name for deterministic output. Temporary files beginning
    with ``~$`` and unsupported suffixes are ignored.
    """

    source_dir = Path(directory).expanduser()
    if not source_dir.exists():
        raise FileNotFoundError(f"目录不存在：{source_dir}")
    if not source_dir.is_dir():
        raise DataToolkitError(f"输入路径不是目录：{source_dir}")

    paths = sorted(
        path
        for path in source_dir.glob(pattern)
        if path.is_file()
        and path.suffix.lower() in SUPPORTED_INPUT_SUFFIXES
        and not path.name.startswith("~$")
    )
    if not paths:
        raise DataToolkitError(f"目录中没有可合并的表格文件：{source_dir}")

    frames: list[pd.DataFrame] = []
    expected_columns: list[object] | None = None
    expected_source: Path | None = None
    for source in paths:
        frame = read_table(source, sheet=sheet)
        if SOURCE_COLUMN in frame.columns:
            raise DataToolkitError(
                f"源文件已包含“{SOURCE_COLUMN}”列，无法自动添加来源：{source.name}"
            )
        if expected_columns is None:
            expected_columns = list(frame.columns)
            expected_source = source
        elif list(frame.columns) != expected_columns:
            raise DataToolkitError(
                "文件结构不一致："
                f"{expected_source.name}={expected_columns}，"
                f"{source.name}={list(frame.columns)}"
            )

        marked = frame.copy()
        marked.insert(0, SOURCE_COLUMN, source.name)
        frames.append(marked)

    return pd.concat(frames, ignore_index=True)
