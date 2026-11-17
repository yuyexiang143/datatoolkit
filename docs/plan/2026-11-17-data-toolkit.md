# Data Toolkit Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build and verify a runnable Excel/CSV cleaning, merging, and reporting portfolio project.

**Architecture:** A dependency-light Python package separates table I/O and cleaning (`core`), workbook reporting (`report`), argparse entry points (`cli`), and tkinter UI (`gui`). Tests exercise the real pandas/openpyxl behavior with temporary files.

**Tech Stack:** Python 3, pandas, openpyxl, xlrd, tkinter, pytest

**Spec:** `docs/superpowers/specs/2026-11-17-data-toolkit-design.md`

## Global Constraints

- Support `.xlsx`, `.xls`, and `.csv`; CSV encoding is UTF-8 or GBK.
- Support a selected Excel sheet.
- Cleaning must cover selected-column and full-row deduplication, drop/fill/forward-fill null strategies, date normalization, amount normalization, and whitespace trimming.
- Directory merge must add a source filename column.
- Report must include before/after row counts, grouped sum/count, and a separate anomalies sheet.
- CLI is built with argparse and exposes `clean`, `merge`, and `report`.
- GUI is a runnable tkinter window.
- All production functions use docstrings and type annotations.

---

### Task 1: Core reading and cleaning

**Files:**
- Create: `datatoolkit/core.py`
- Create: `tests/test_core.py`

**Interfaces:**
- Produces: `read_table(path, sheet=None) -> pd.DataFrame`
- Produces: `clean_dataframe(frame, options) -> pd.DataFrame`
- Produces: `merge_directory(directory, pattern="*") -> pd.DataFrame`
- Produces: `write_table(frame, path) -> None`

- [ ] Write failing tests for XLSX sheet selection, GBK CSV, XLS, deduplication, null strategies, dates, amounts, trimming, and merging.
- [ ] Run `pytest tests/test_core.py -q` and confirm imports fail because the package does not exist.
- [ ] Implement the core functions with explicit column validation and descriptive exceptions.
- [ ] Run `pytest tests/test_core.py -q` and confirm all core tests pass.

### Task 2: Report generation

**Files:**
- Create: `datatoolkit/report.py`
- Create: `tests/test_report.py`

**Interfaces:**
- Consumes: `CleaningOptions`, `clean_dataframe`, `read_table`
- Produces: `build_report(input_path, output_path, options, group_column, sum_column) -> ReportResult`

- [ ] Write a failing integration test that reads the generated workbook and checks summary, grouped values, and anomaly rows.
- [ ] Run `pytest tests/test_report.py -q` and confirm the report module is missing.
- [ ] Implement the three-sheet report and return row/group/anomaly counts.
- [ ] Run `pytest tests/test_report.py -q` and confirm it passes.

### Task 3: CLI and sample data

**Files:**
- Create: `datatoolkit/__init__.py`
- Create: `datatoolkit/__main__.py`
- Create: `datatoolkit/cli.py`
- Create: `datatoolkit/gui.py`
- Create: `scripts/generate_sample_data.py`
- Create: `requirements.txt`
- Create: `tests/test_cli.py`

**Interfaces:**
- Consumes: all core and report public functions
- Produces: `python -m datatoolkit clean|merge|report`

- [ ] Write a failing subprocess test for `--help` and one real clean command.
- [ ] Run the CLI test and confirm failure before creating the entry point.
- [ ] Implement argparse commands and the tkinter form.
- [ ] Generate three dirty sample files with reproducible script.
- [ ] Run clean, merge, and report for real.

### Task 4: Portfolio documentation and verification

**Files:**
- Create: `README.md`
- Create: `.gitignore`
- Create: `pyproject.toml`

**Interfaces:**
- Consumes: verified terminal output from Task 3

- [ ] Paste real CLI output into Chinese README usage examples.
- [ ] Document install, structure, tests, and limitations.
- [ ] Run the full pytest suite.
- [ ] Launch the tkinter UI as a background process, confirm it remains running, then stop the smoke-test process.
