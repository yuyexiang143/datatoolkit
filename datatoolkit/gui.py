"""Simple tkinter front end for the data cleaning workflow."""

from __future__ import annotations

import tkinter as tk
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from typing import Mapping

from datatoolkit.core import CleaningOptions, clean_dataframe, read_table, write_table


NULL_STRATEGY_MAP: Mapping[str, str] = {
    "删除含空值行": "drop",
    "填充默认值": "fill",
    "用上一行填充": "ffill",
    "保留空值": "keep",
}


class DataToolkitApp:
    """Main tkinter application with file, cleaning, output, and log controls."""

    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("Excel / CSV 数据清洗工具")
        self.root.geometry("760x620")
        self.root.minsize(680, 560)

        self.input_var = tk.StringVar()
        self.output_var = tk.StringVar()
        self.sheet_var = tk.StringVar()
        self.null_var = tk.StringVar(value="删除含空值行")
        self.fill_var = tk.StringVar(value="")
        self.dedupe_columns_var = tk.StringVar()
        self.date_columns_var = tk.StringVar()
        self.amount_columns_var = tk.StringVar()
        self.drop_duplicates_var = tk.BooleanVar(value=False)
        self.trim_var = tk.BooleanVar(value=True)

        self._build()

    def _build(self) -> None:
        """Create and place all widgets."""

        container = ttk.Frame(self.root, padding=16)
        container.pack(fill="both", expand=True)
        container.columnconfigure(1, weight=1)

        ttk.Label(container, text="输入文件").grid(row=0, column=0, sticky="w", pady=5)
        ttk.Entry(container, textvariable=self.input_var).grid(
            row=0, column=1, sticky="ew", padx=8, pady=5
        )
        ttk.Button(container, text="选择文件", command=self._choose_input).grid(
            row=0, column=2, pady=5
        )

        ttk.Label(container, text="工作表").grid(row=1, column=0, sticky="w", pady=5)
        ttk.Entry(container, textvariable=self.sheet_var).grid(
            row=1, column=1, sticky="ew", padx=8, pady=5
        )
        ttk.Label(container, text="CSV 可留空").grid(
            row=1, column=2, sticky="w", pady=5
        )

        ttk.Label(container, text="输出文件").grid(row=2, column=0, sticky="w", pady=5)
        ttk.Entry(container, textvariable=self.output_var).grid(
            row=2, column=1, sticky="ew", padx=8, pady=5
        )
        ttk.Button(container, text="选择位置", command=self._choose_output).grid(
            row=2, column=2, pady=5
        )

        options = ttk.LabelFrame(container, text="清洗选项", padding=10)
        options.grid(row=3, column=0, columnspan=3, sticky="ew", pady=(12, 8))
        options.columnconfigure(1, weight=1)

        ttk.Checkbutton(
            options,
            text="去除字符串首尾空格",
            variable=self.trim_var,
        ).grid(row=0, column=0, sticky="w", pady=4)
        ttk.Checkbutton(
            options,
            text="去除重复行",
            variable=self.drop_duplicates_var,
        ).grid(row=0, column=1, sticky="w", pady=4)

        ttk.Label(options, text="重复判断列").grid(row=1, column=0, sticky="w", pady=4)
        ttk.Entry(options, textvariable=self.dedupe_columns_var).grid(
            row=1, column=1, sticky="ew", padx=8, pady=4
        )
        ttk.Label(options, text="多个列用逗号分隔").grid(
            row=1, column=2, sticky="w", pady=4
        )

        ttk.Label(options, text="空值处理").grid(row=2, column=0, sticky="w", pady=4)
        ttk.Combobox(
            options,
            textvariable=self.null_var,
            values=tuple(NULL_STRATEGY_MAP),
            state="readonly",
        ).grid(row=2, column=1, sticky="w", padx=8, pady=4)

        ttk.Label(options, text="默认填充值").grid(row=3, column=0, sticky="w", pady=4)
        ttk.Entry(options, textvariable=self.fill_var).grid(
            row=3, column=1, sticky="ew", padx=8, pady=4
        )

        ttk.Label(options, text="日期列").grid(row=4, column=0, sticky="w", pady=4)
        ttk.Entry(options, textvariable=self.date_columns_var).grid(
            row=4, column=1, sticky="ew", padx=8, pady=4
        )
        ttk.Label(options, text="例如：日期,下单日期").grid(
            row=4, column=2, sticky="w", pady=4
        )

        ttk.Label(options, text="金额列").grid(row=5, column=0, sticky="w", pady=4)
        ttk.Entry(options, textvariable=self.amount_columns_var).grid(
            row=5, column=1, sticky="ew", padx=8, pady=4
        )
        ttk.Label(options, text="例如：金额,销售额").grid(
            row=5, column=2, sticky="w", pady=4
        )

        ttk.Button(
            container,
            text="开始清洗",
            command=self._execute,
        ).grid(row=4, column=0, columnspan=3, sticky="ew", pady=8)

        log_frame = ttk.LabelFrame(container, text="运行日志", padding=8)
        log_frame.grid(row=5, column=0, columnspan=3, sticky="nsew")
        container.rowconfigure(5, weight=1)
        log_frame.rowconfigure(0, weight=1)
        log_frame.columnconfigure(0, weight=1)

        self.log = tk.Text(log_frame, height=12, wrap="word", state="disabled")
        self.log.grid(row=0, column=0, sticky="nsew")
        scrollbar = ttk.Scrollbar(log_frame, command=self.log.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.log.configure(yscrollcommand=scrollbar.set)

    @staticmethod
    def _split_columns(value: str) -> tuple[str, ...]:
        """Split a comma-separated field into column names."""

        comma = "，"
        normalized = value.replace(comma, ",")
        return tuple(item.strip() for item in normalized.split(",") if item.strip())

    def _log(self, message: str) -> None:
        """Append a timestamped message to the visual log."""

        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log.configure(state="normal")
        self.log.insert("end", f"[{timestamp}] {message}\n")
        self.log.see("end")
        self.log.configure(state="disabled")

    def _choose_input(self) -> None:
        """Ask for an input file and suggest a matching output name."""

        selected = filedialog.askopenfilename(
            title="选择表格文件",
            filetypes=[
                ("表格文件", "*.xlsx *.xls *.csv"),
                ("Excel 文件", "*.xlsx *.xls"),
                ("CSV 文件", "*.csv"),
            ],
        )
        if selected:
            self.input_var.set(selected)
            suggested = Path(selected).with_name(f"{Path(selected).stem}_cleaned.xlsx")
            self.output_var.set(str(suggested))
            self._log(f"已选择输入：{selected}")

    def _choose_output(self) -> None:
        """Ask for an output workbook path."""

        selected = filedialog.asksaveasfilename(
            title="保存清洗结果",
            defaultextension=".xlsx",
            filetypes=[("Excel 文件", "*.xlsx"), ("CSV 文件", "*.csv")],
        )
        if selected:
            self.output_var.set(selected)

    def _execute(self) -> None:
        """Run cleaning with the current form values and report errors in the UI."""

        input_path = self.input_var.get().strip()
        output_path = self.output_var.get().strip()
        if not input_path or not output_path:
            messagebox.showwarning("缺少信息", "请选择输入文件和输出文件。")
            return

        self._log("开始读取文件")
        try:
            sheet = self.sheet_var.get().strip() or None
            raw = read_table(input_path, sheet=sheet)
            options = CleaningOptions(
                trim_whitespace=self.trim_var.get(),
                drop_duplicates=self.drop_duplicates_var.get(),
                dedupe_columns=self._split_columns(self.dedupe_columns_var.get()),
                null_strategy=NULL_STRATEGY_MAP[self.null_var.get()],  # type: ignore[arg-type]
                fill_value=self.fill_var.get(),
                date_columns=self._split_columns(self.date_columns_var.get()),
                amount_columns=self._split_columns(self.amount_columns_var.get()),
            )
            cleaned = clean_dataframe(raw, options)
            output = write_table(cleaned, output_path)
            self._log(
                f"读取完成：{len(raw)} 行；清洗后：{len(cleaned)} 行；结果已保存到 {output}"
            )
            messagebox.showinfo("完成", f"清洗完成，共 {len(cleaned)} 行。")
        except Exception as exc:  # GUI must keep running after user-input errors.
            self._log(f"失败：{exc}")
            messagebox.showerror("执行失败", str(exc))


def main() -> None:
    """Start the tkinter application."""

    root = tk.Tk()
    DataToolkitApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
