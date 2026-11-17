# datatoolkit — Excel / CSV 数据清洗与报表自动化工具

把"每个月手工整理这几张表"变成一条命令。

中小企业和个体户最常见的数据活：多个人各自维护的表格，日期格式不统一（`2024/1/5` 和 `2024-01-16` 混着）、金额带着 `¥` 和千分位逗号、有重复行、有空值、还有一堆分散在不同文件里的分区数据。手工做一次两小时，每月都要重做。

这个工具做三件事：**清洗、合并、出汇总报表**。原始文件全程只读，不覆盖。

---

## 特性

- 读取 `.xlsx` / `.xls` / `.csv`，CSV 自动识别 `utf-8` / `GBK` 编码（中文 CSV 的常见坑）
- 去重（可指定按哪几列判断，默认整行）
- 空值策略：删除 / 填充默认值 / 用上一行填充 / 保留
- 日期归一化：多种输入格式统一为 `YYYY-MM-DD`
- 金额清洗：去掉 `¥`、`￥`、千分位逗号、空格，转成数值
- 字符串首尾空格清理
- 多文件合并，自动追加「来源文件名」列
- 生成汇总报表：行数对比、分组求和/计数、异常值清单（三个 Sheet）
- 命令行 + 图形界面两种用法
- 输入输出完全分离，原始文件不会被修改

---

## 安装

需要 Python 3.10+。

```bash
pip install -r requirements.txt
```

开发环境（跑测试）：

```bash
pip install -r requirements-dev.txt
```

---

## 快速开始

仓库自带 `sample_data/`，里面是三个故意的「脏」文件：

- `region_east.xlsx` — 有重复行、空值、`¥ 1,200.50` 这种金额、`2024/1/5` 这种日期
- `region_south_gbk.csv` — **GBK 编码**的中文 CSV
- `region_west.xls` — 老版 Excel 格式

### 1. 清洗单个文件

```bash
python -m datatoolkit clean sample_data/region_east.xlsx -o output/cleaned_east.xlsx \
    --drop-duplicates --dedupe-columns 订单号 \
    --date-columns 日期 --amount-columns 金额 \
    --null-strategy fill --fill-value 无
```

真实输出：

```
输入文件：sample_data\region_east.xlsx
工作表：默认
清洗前：5 行，6 列
字符串去空格：已执行
空值策略:填充默认值
去重：已按 订单号 去重
日期列：日期
金额列：金额
清洗后：4 行，6 列
输出文件：output\cleaned_east.xlsx
```

清洗前后对比：

| 清洗前 | 清洗后 |
|---|---|
| `2024/1/5` | `2024-01-05` |
| `￥880` | `880` |
| `¥ 1,200.50` | `1200.5` |
| `A001` 出现两次 | 只保留一条 |
| 空值 | `无` |

### 2. GBK 编码的 CSV 不用管

```bash
python -m datatoolkit clean sample_data/region_south_gbk.csv -o output/cleaned_south.xlsx \
    --drop-duplicates --date-columns 日期 --amount-columns 金额
```

```
输入文件：sample_data\region_south_gbk.csv
工作表：默认
清洗前：4 行，6 列
字符串去空格：已执行
空值策略:保留空值
去重：已按 整行 去重
日期列：日期
金额列：金额
清洗后：4 行，6 列
输出文件：output\cleaned_south.xlsx
```

编码会自动识别，不需要手动指定。

### 3. 合并一个目录下的所有文件

```bash
python -m datatoolkit merge sample_data -o output/merged.xlsx
```

```
输入目录：sample_data
匹配文件：3 个
合并结果：14 行，7 列
来源列：来源文件名
输出文件：output\merged.xlsx
```

7 列 = 原有 6 列 + 自动追加的「来源文件名」，方便追溯每一行来自哪个分区的表。

### 4. 生成汇总报表

```bash
python -m datatoolkit report output/cleaned_east.xlsx -o output/report.xlsx \
    --group-column 地区 --sum-column 金额
```

```
输入文件：output/cleaned_east.xlsx
原始行数：4
清洗后行数：4
删除行数：0
分组列：地区
求和列：金额
分组数：1
异常清单：0 条
输出文件：output\report.xlsx
```

报表包含三个 Sheet：

**汇总**

| 指标 | 值 |
|---|---|
| 原始行数 | 4 |
| 清洗后行数 | 4 |
| 删除行数 | 0 |
| 列数 | 6 |
| 分组列 | 地区 |
| 求和列 | 金额 |
| 异常记录数 | 0 |

**分组汇总**

| 地区 | 记录数 | 金额_求和 |
|---|---|---|
| 华东 | 4 | 4430.5 |

**异常值** —— 空表，因为样例数据里没有清洗失败的记录。

---

## 图形界面

不想敲命令：

```bash
python -m datatoolkit.gui
```

先是选文件、勾选项、点执行、看日志。界面很朴素，功能上和命令行等价。

---

## 目录结构

```
datatoolkit/
  __init__.py
  __main__.py     # python -m datatoolkit 入口
  core.py         # 清洗逻辑（去重/空值/日期/金额/合并），不依赖 CLI
  cli.py          # argparse 命令行
  report.py       # 汇总报表生成
  gui.py          # tkinter 界面
sample_data/      # 三个脏数据样例
tests/            # pytest 测试
output/           # 输出目录（不纳入版本管理）
```

`core.py` 不依赖 `cli.py`，可以单独 import 用在别的脚本里：

```python
from datatoolkit.core import clean_frame

df = clean_frame(df, date_columns=["日期"], amount_columns=["金额"])
```

---

## 测试

```bash
python -m pytest -q
```

真实结果：

```
14 passed in 12.03s
```

14 个用例覆盖：去重（整行 / 指定列）、空值四种策略、日期多格式归一、金额清洗、GBK 读取、多文件合并、报表三个 Sheet 的生成、异常清单。

---

## 已知限制

写在这里，免得你踩到：

1. **空值填充会污染数值列。** 如果对「金额」列用 `--fill-value 无`，该列会变成字符串类型，后续无法直接求和。数值列建议用 `--null-strategy drop` 或保留空值。
2. **异常清单只覆盖清洗失败的记录**（比如日期无法解析、金额无法转换），不检测业务逻辑层面的问题（比如「金额缺失」本身算不算异常）。业务规则需要另外判断。
3. **`.xls` 老格式依赖 `xlrd`**，复杂的老文件（多个 sheet、有合并单元格）可能读不干净。遇到读不了的文件会明确报错，不会静默产出错误结果。
4. **不支持合并单元格。** 带合并单元格的表需要先在 Excel 里拆开。
5. **不加 `--drop-duplicates` 就不会去重**，默认行为是只做格式清理。这一点容易误会。
6. **大文件（10 万行以上）走的是整表读入内存**，没有做流式处理，内存占用会比较高。

---

## 为什么原始文件不会被改

所有写操作都指向 `-o` 指定的输出路径。输入文件只读打开，不做任何写入。这样即使参数填错，原始数据也还在。
