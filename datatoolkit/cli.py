"""Argparse command line interface for datatoolkit."""

from __future__ import annotations

import argparse
import sys
from pathlib import Path
from typing import Sequence

from datatoolkit.core import (
    CleaningOptions,
    DataToolkitError,
    clean_dataframe,
    merge_directory,
    read_table,
    write_table,
)
from datatoolkit.report import build_report


NULL_LABELS = {
    "drop": "删除含空值行",
    "fill": "填充默认值",
    "ffill": "用上一行填充",
    "keep": "保留空值",
}


def _parser() -> argparse.ArgumentParser:
    """Create the complete command line parser."""

    parser = argparse.ArgumentParser(
        prog="python -m datatoolkit",
        description="Excel / CSV 数据清洗、目录合并和汇总报表工具",
    )
    subparsers = parser.add_subparsers(dest="command", required=True)

    clean = subparsers.add_parser("clean", help="清洗一个表格文件")
    clean.add_argument("input", type=Path, help="输入 .xlsx/.xls/.csv")
    clean.add_argument("-o", "--output", type=Path, required=True, help="输出 .xlsx/.csv")
    clean.add_argument("--sheet", help="Excel 工作表名称或序号")
    _add_cleaning_arguments(clean)
    clean.set_defaults(handler=_run_clean)

    merge = subparsers.add_parser("merge", help="合并目录中的同结构文件")
    merge.add_argument("directory", type=Path, help="包含多个表格的目录")
    merge.add_argument("-o", "--output", type=Path, required=True, help="合并后 .xlsx/.csv")
    merge.add_argument("--pattern", default="*", help="文件名匹配模式，默认 *")
    merge.add_argument("--sheet", help="每个 Excel 文件读取的工作表名称或序号")
    merge.set_defaults(handler=_run_merge)

    report = subparsers.add_parser("report", help="清洗并生成多 Sheet 汇总报表")
    report.add_argument("input", type=Path, help="输入 .xlsx/.xls/.csv")
    report.add_argument("-o", "--output", type=Path, required=True, help="输出 .xlsx")
    report.add_argument("--sheet", help="Excel 工作表名称或序号")
    report.add_argument("--group-column", required=True, help="分组列")
    report.add_argument("--sum-column", required=True, help="求和的金额列")
    _add_cleaning_arguments(report)
    report.set_defaults(handler=_run_report)
    return parser


def _add_cleaning_arguments(parser: argparse.ArgumentParser) -> None:
    """Register shared cleaning flags on a subcommand."""

    parser.add_argument(
        "--no-trim",
        action="store_true",
        help="关闭字符串首尾空格清理",
    )
    parser.add_argument(
        "--drop-duplicates",
        action="store_true",
        help="删除重复行",
    )
    parser.add_argument(
        "--dedupe-columns",
        nargs="+",
        default=[],
        help="按这些列判断重复；不填表示比较整行",
    )
    parser.add_argument(
        "--null-strategy",
        choices=("drop", "fill", "ffill", "keep"),
        default="keep",
        help="空值处理：drop 删除、fill 填充、ffill 上一行、keep 保留",
    )
    parser.add_argument(
        "--fill-value",
        default="",
        help="--null-strategy fill 时使用的默认值",
    )
    parser.add_argument(
        "--date-columns",
        nargs="+",
        default=[],
        help="需要统一日期格式的列",
    )
    parser.add_argument(
        "--amount-columns",
        nargs="+",
        default=[],
        help="需要清洗为金额数值的列",
    )


def _split_columns(values: Sequence[str]) -> tuple[str, ...]:
    """Accept both space-separated args and comma-separated column lists."""

    result: list[str] = []
    for value in values:
        result.extend(item.strip() for item in value.split(",") if item.strip())
    return tuple(result)


def _options_from_args(args: argparse.Namespace) -> CleaningOptions:
    """Translate argparse values into the core options dataclass."""

    return CleaningOptions(
        trim_whitespace=not args.no_trim,
        drop_duplicates=args.drop_duplicates,
        dedupe_columns=_split_columns(args.dedupe_columns),
        null_strategy=args.null_strategy,
        fill_value=args.fill_value,
        date_columns=_split_columns(args.date_columns),
        amount_columns=_split_columns(args.amount_columns),
    )


def _run_clean(args: argparse.Namespace) -> int:
    """Handle the ``clean`` subcommand."""

    raw = read_table(args.input, sheet=args.sheet)
    options = _options_from_args(args)
    cleaned = clean_dataframe(raw, options)
    output = write_table(cleaned, args.output)

    dedupe_text = "、".join(options.dedupe_columns) or "整行"
    print(f"输入文件：{args.input}")
    print(f"工作表：{args.sheet or '默认'}")
    print(f"清洗前：{len(raw)} 行，{len(raw.columns)} 列")
    print(f"字符串去空格：{'已执行' if options.trim_whitespace else '未执行'}")
    print(f"空值策略:{NULL_LABELS[options.null_strategy]}")
    print(
        "去重："
        + (
            f"已按 {dedupe_text} 去重"
            if options.drop_duplicates
            else "未执行"
        )
    )
    print(f"日期列：{'、'.join(options.date_columns) or '无'}")
    print(f"金额列：{'、'.join(options.amount_columns) or '无'}")
    print(f"清洗后：{len(cleaned)} 行，{len(cleaned.columns)} 列")
    print(f"输出文件：{output}")
    return 0


def _run_merge(args: argparse.Namespace) -> int:
    """Handle the ``merge`` subcommand."""

    merged = merge_directory(args.directory, pattern=args.pattern, sheet=args.sheet)
    output = write_table(merged, args.output)
    source_count = merged["来源文件名"].nunique()

    print(f"输入目录：{args.directory}")
    print(f"匹配文件：{source_count} 个")
    print(f"合并结果：{len(merged)} 行，{len(merged.columns)} 列")
    print(f"来源列：来源文件名")
    print(f"输出文件：{output}")
    return 0


def _run_report(args: argparse.Namespace) -> int:
    """Handle the ``report`` subcommand."""

    options = _options_from_args(args)
    result = build_report(
        args.input,
        args.output,
        options,
        group_column=args.group_column,
        sum_column=args.sum_column,
        sheet=args.sheet,
    )

    print(f"输入文件：{args.input}")
    print(f"原始行数：{result.original_rows}")
    print(f"清洗后行数：{result.cleaned_rows}")
    print(f"删除行数：{result.original_rows - result.cleaned_rows}")
    print(f"分组列：{args.group_column}")
    print(f"求和列：{args.sum_column}")
    print(f"分组数：{result.grouped_rows}")
    print(f"异常清单：{result.anomaly_count} 条")
    print(f"输出文件：{result.output_path}")
    return 0


def main(argv: Sequence[str] | None = None) -> int:
    """Run the CLI and return a process exit code."""

    parser = _parser()
    args = parser.parse_args(argv)
    try:
        return int(args.handler(args))
    except (DataToolkitError, FileNotFoundError, PermissionError) as exc:
        print(f"错误：{exc}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
