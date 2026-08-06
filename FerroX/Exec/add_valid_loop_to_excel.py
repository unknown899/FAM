#!/usr/bin/env python3
"""Calculate P-V loop validity and update an Excel workbook.

Use ``python add_valid_loop_to_excel.py --help`` for the complete command
workflow and copyable examples.

For each case directory whose name starts with ``--prefix``, read
``figs/MFIS_PV_curve.csv`` and define (CSV header excluded):

    P_start = -P_mean at data row 1
    P_r-    = -P_mean at data row 10
    P_r+    = -P_mean at data row 28
    P_end   = -P_mean at data row 37

The loop is valid when all of the following are true:

    abs(P_r+ - P_r-) > open_threshold
    P_r+ > 0 and P_r- < 0
    abs(P_start - P_end) < close_threshold

Only worksheet rows whose ``folder`` matches a selected directory are
changed. Other rows are left untouched. If a matching case has no CSV, its
``valid_loop`` value is set to ``no_data`` and processing continues.

With ``--dry-run``, proposed values are written only to a temporary workbook.
``build_dash.py`` then uses that temporary workbook to generate an HTML
preview, after which the temporary workbook is deleted. Without ``--dry-run``,
the workbook selected by ``--excel`` is updated.
"""

from __future__ import annotations

import argparse
import math
import os
import tempfile
from copy import copy
from pathlib import Path
from typing import Any

from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet

import build_dash
from utils import loop_metrics


DEFAULT_EXCEL_NAME = "MFIS_dataset.xlsx"
DEFAULT_SHEET_NAME = "experiments"
DEFAULT_FILE_COLUMN = "folder"
DEFAULT_VALID_COLUMN = "valid_loop"
DEFAULT_PREVIEW_HTML = "index.html"
MISSING_DATA_VALUE = "no_data"
RELATIVE_CSV_PATH = Path("figs") / "MFIS_PV_curve.csv"

HELP_EPILOG = r"""
使用流程:
  1. 先用 --dry-run 產生 HTML，確認 valid_loop 與對應圖片。
  2. 視需要調整 threshold，重新執行 dry-run。
  3. 確認後移除 --dry-run，才會寫入指定 Excel。

結果值:
  1       三項 loop 條件皆通過
  0       至少一項 loop 條件未通過
  no_data 找不到該 case 的 figs/MFIS_PV_curve.csv

Excel 欄位:
  完全沿用 add_EPR 的新增欄位邏輯：新欄繼承 gamma 欄格式，並延伸
  AutoFilter 與 Excel Table；欄位已存在時也會補齊表頭、格式與篩選範圍。

範例 1：預覽 MFIS_t_5_ cases
  python add_valid_loop_to_excel.py \
      --prefix MFIS_t_5_ \
      --open-threshold 0.1 \
      --close-threshold 0.02 \
      --dry-run

  預設讀取 MFIS_dataset.xlsx，並產生 index.html；正式 Excel 不會改變。

範例 2：指定 Excel 與預覽 HTML
  python add_valid_loop_to_excel.py \
      --prefix MFIS_t_5_ \
      --open-threshold 0.1 \
      --close-threshold 0.02 \
      --excel other_dataset.xlsx \
      --html valid_preview.html \
      --dry-run

範例 3：確認後寫入 Excel
  python add_valid_loop_to_excel.py \
      --prefix MFIS_t_5_ \
      --open-threshold 0.1 \
      --close-threshold 0.02 \
      --excel other_dataset.xlsx

提示:
  --dry_run 等同 --dry-run
  --excel-name 等同 --excel
  --preview-html 等同 --html
"""


def nonnegative_float(text: str) -> float:
    """argparse converter for a finite, non-negative threshold."""

    try:
        value = float(text)
    except ValueError as exc:
        raise argparse.ArgumentTypeError(f"not a number: {text!r}") from exc

    if not math.isfinite(value) or value < 0:
        raise argparse.ArgumentTypeError(
            f"threshold must be finite and >= 0, got {text!r}"
        )
    return value


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Set valid_loop for case folders in the execution directory "
            "whose names begin with a specified prefix."
        ),
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=HELP_EPILOG,
    )
    parser.add_argument(
        "--prefix",
        required=True,
        help="Case-directory prefix, for example MFIS_t_5_nomi_3.5_gvar10_4_",
    )
    parser.add_argument(
        "--open-threshold",
        required=True,
        type=nonnegative_float,
        help="Required lower bound for abs(P_r_plus - P_r_minus).",
    )
    parser.add_argument(
        "--close-threshold",
        required=True,
        type=nonnegative_float,
        help="Required upper bound for abs(P_start - P_end).",
    )
    parser.add_argument(
        "--root",
        type=Path,
        default=Path.cwd(),
        help="Execution directory containing the Excel file and case folders (default: cwd).",
    )
    parser.add_argument(
        "--excel",
        "--excel-name",
        dest="excel",
        type=Path,
        default=Path(DEFAULT_EXCEL_NAME),
        help=(
            "Excel filename/path, relative to --root unless absolute "
            f"(default: {DEFAULT_EXCEL_NAME})."
        ),
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=None,
        help="Optional output workbook. By default, update --excel in place.",
    )
    parser.add_argument(
        "--sheet",
        default=DEFAULT_SHEET_NAME,
        help=f"Worksheet name (default: {DEFAULT_SHEET_NAME}).",
    )
    parser.add_argument(
        "--file-column",
        default=DEFAULT_FILE_COLUMN,
        help=f"Column containing case-directory names (default: {DEFAULT_FILE_COLUMN}).",
    )
    parser.add_argument(
        "--valid-column",
        default=DEFAULT_VALID_COLUMN,
        help=f"Column to create/update (default: {DEFAULT_VALID_COLUMN}).",
    )
    parser.add_argument(
        "--p-column",
        default=loop_metrics.DEFAULT_P_COLUMN,
        help=(
            "Polarization column in each CSV "
            f"(default: {loop_metrics.DEFAULT_P_COLUMN})."
        ),
    )
    parser.add_argument(
        "--dry-run",
        "--dry_run",
        dest="dry_run",
        action="store_true",
        help=(
            "Write proposed values to a temporary workbook and generate "
            "a preview dashboard without changing the real workbook."
        ),
    )
    parser.add_argument(
        "--preview-html",
        "--html",
        dest="preview_html",
        type=Path,
        default=Path(DEFAULT_PREVIEW_HTML),
        help=(
            "HTML generated during --dry-run, relative to --root unless "
            f"absolute (default: {DEFAULT_PREVIEW_HTML})."
        ),
    )
    return parser.parse_args()


def resolve_under_root(path: Path, root: Path) -> Path:
    return path if path.is_absolute() else root / path


def _normalize_header(value: Any) -> str:
    if value is None:
        return ""
    return "".join(str(value).strip().casefold().split())


def _find_header_column(
    worksheet: Any,
    header_row: int,
    aliases: set[str],
) -> int:
    normalized_aliases = {_normalize_header(alias) for alias in aliases}
    matches = [
        column
        for column in range(1, worksheet.max_column + 1)
        if _normalize_header(worksheet.cell(header_row, column).value)
        in normalized_aliases
    ]
    if not matches:
        raise KeyError(
            f"Could not find any of {sorted(aliases)!r} in header row {header_row}"
        )
    if len(matches) > 1:
        raise KeyError(f"Multiple matching columns for {sorted(aliases)!r}: {matches}")
    return matches[0]


def _ensure_output_column(
    worksheet: Any,
    header_row: int,
    name: str,
    style_source_column: int,
) -> tuple[int, bool]:
    normalized_name = _normalize_header(name)
    for column in range(1, worksheet.max_column + 1):
        if _normalize_header(worksheet.cell(header_row, column).value) == normalized_name:
            return column, False

    column = worksheet.max_column + 1
    source = worksheet.cell(header_row, style_source_column)
    target = worksheet.cell(header_row, column, name)
    if source.has_style:
        target._style = copy(source._style)
    target.alignment = copy(source.alignment)
    target.protection = copy(source.protection)
    return column, True


def _copy_style(source: Any, target: Any) -> None:
    if source.has_style:
        target._style = copy(source._style)
    target.alignment = copy(source.alignment)
    target.protection = copy(source.protection)


def _extend_filter_and_tables(
    worksheet: Any,
    *,
    header_row: int,
    old_max_column: int,
    new_max_column: int,
) -> None:
    if new_max_column <= old_max_column:
        return

    from openpyxl.utils import get_column_letter, range_boundaries
    from openpyxl.worksheet.table import TableColumn

    def extend_ref(reference: str) -> str:
        min_col, min_row, max_col, max_row = range_boundaries(reference)
        if min_row == header_row and max_col == old_max_column:
            max_col = new_max_column
        return (
            f"{get_column_letter(min_col)}{min_row}:"
            f"{get_column_letter(max_col)}{max_row}"
        )

    if worksheet.auto_filter.ref:
        worksheet.auto_filter.ref = extend_ref(worksheet.auto_filter.ref)

    for table in worksheet.tables.values():
        _, min_row, max_col, _ = range_boundaries(table.ref)
        if min_row != header_row or max_col != old_max_column:
            continue

        next_id = max((column.id for column in table.tableColumns), default=0) + 1
        for column in range(old_max_column + 1, new_max_column + 1):
            table.tableColumns.append(
                TableColumn(
                    id=next_id,
                    name=str(worksheet.cell(header_row, column).value),
                )
            )
            next_id += 1
        table.ref = extend_ref(table.ref)
        if table.autoFilter is not None and table.autoFilter.ref:
            table.autoFilter.ref = extend_ref(table.autoFilter.ref)


def _repair_existing_output_column(
    worksheet: Any,
    *,
    header_row: int,
    style_source_column: int,
    output_column: int,
    width: float,
) -> None:
    """Repair a column created by an older run but left outside the filter."""

    from openpyxl.utils import get_column_letter, range_boundaries
    from openpyxl.worksheet.table import TableColumn

    _copy_style(
        worksheet.cell(header_row, style_source_column),
        worksheet.cell(header_row, output_column),
    )
    worksheet.column_dimensions[get_column_letter(output_column)].width = width

    def extend_ref(reference: str) -> str:
        min_col, min_row, max_col, max_row = range_boundaries(reference)
        if min_row == header_row and max_col < output_column:
            max_col = output_column
        return (
            f"{get_column_letter(min_col)}{min_row}:"
            f"{get_column_letter(max_col)}{max_row}"
        )

    if worksheet.auto_filter.ref:
        worksheet.auto_filter.ref = extend_ref(worksheet.auto_filter.ref)

    for table in worksheet.tables.values():
        _, min_row, max_col, _ = range_boundaries(table.ref)
        if min_row != header_row or max_col >= output_column:
            continue

        next_id = max((column.id for column in table.tableColumns), default=0) + 1
        for column in range(max_col + 1, output_column + 1):
            table.tableColumns.append(
                TableColumn(
                    id=next_id,
                    name=str(worksheet.cell(header_row, column).value),
                )
            )
            next_id += 1
        table.ref = extend_ref(table.ref)
        if table.autoFilter is not None and table.autoFilter.ref:
            table.autoFilter.ref = extend_ref(table.autoFilter.ref)


def discover_case_directories(root: Path, prefix: str) -> list[Path]:
    if not prefix:
        raise ValueError("--prefix must not be empty.")
    return sorted(
        (
            path
            for path in root.iterdir()
            if path.is_dir() and path.name.startswith(prefix)
        ),
        key=lambda path: path.name,
    )


def worksheet_rows_by_case(
    worksheet: Worksheet,
    file_column: int,
) -> dict[str, list[int]]:
    rows: dict[str, list[int]] = {}
    for row_number in range(2, worksheet.max_row + 1):
        value = worksheet.cell(row=row_number, column=file_column).value
        if value is None:
            continue
        case_name = str(value).strip()
        if case_name:
            rows.setdefault(case_name, []).append(row_number)
    return rows


def save_workbook_atomically(workbook: object, output_path: Path) -> None:
    """Write beside the destination, then replace it only after save succeeds."""

    output_path.parent.mkdir(parents=True, exist_ok=True)
    descriptor, temporary_name = tempfile.mkstemp(
        prefix=f".{output_path.stem}.",
        suffix=output_path.suffix,
        dir=output_path.parent,
    )
    os.close(descriptor)
    temporary_path = Path(temporary_name)

    try:
        workbook.save(temporary_path)
        os.replace(temporary_path, output_path)
    finally:
        temporary_path.unlink(missing_ok=True)


def format_result(case_name: str, metrics: loop_metrics.LoopMetrics) -> str:
    return (
        f"{case_name}: valid_loop={metrics.valid_loop} | "
        f"P_r-={metrics.p_r_minus:.8g}, P_r+={metrics.p_r_plus:.8g}, "
        f"open_gap={metrics.open_gap:.8g} "
        f"({'PASS' if metrics.open_ok else 'FAIL'}) | "
        f"sign={'PASS' if metrics.sign_ok else 'FAIL'} | "
        f"P_start={metrics.p_start:.8g}, P_end={metrics.p_end:.8g}, "
        f"close_gap={metrics.close_gap:.8g} "
        f"({'PASS' if metrics.close_ok else 'FAIL'})"
    )


def run(args: argparse.Namespace) -> int:
    root = args.root.expanduser().resolve()
    if not root.is_dir():
        raise NotADirectoryError(f"Execution directory does not exist: {root}")

    excel_path = resolve_under_root(args.excel.expanduser(), root).resolve()
    output_path = (
        excel_path
        if args.output is None
        else resolve_under_root(args.output.expanduser(), root).resolve()
    )
    preview_html_path = resolve_under_root(
        args.preview_html.expanduser(),
        root,
    ).resolve()
    if not excel_path.is_file():
        raise FileNotFoundError(f"Workbook does not exist: {excel_path}")

    case_directories = discover_case_directories(root, args.prefix)
    if not case_directories:
        raise FileNotFoundError(
            f"No directories beginning with {args.prefix!r} were found under {root}."
        )

    workbook = load_workbook(excel_path)
    if args.sheet not in workbook.sheetnames:
        raise KeyError(
            f"Worksheet {args.sheet!r} was not found. "
            f"Available worksheets: {workbook.sheetnames}"
        )
    worksheet = workbook[args.sheet]

    file_column = _find_header_column(worksheet, 1, {args.file_column})
    gamma_column = _find_header_column(worksheet, 1, {"gamma", "γ"})

    old_max_column = worksheet.max_column
    valid_column, _ = _ensure_output_column(
        worksheet,
        1,
        args.valid_column,
        gamma_column,
    )

    _extend_filter_and_tables(
        worksheet,
        header_row=1,
        old_max_column=old_max_column,
        new_max_column=worksheet.max_column,
    )
    # The column may already exist from an older run but still be outside the
    # AutoFilter and have an unformatted header.  Reapplying the EPR-style
    # formatting makes reruns repair that partial state as well as add new data.
    _repair_existing_output_column(
        worksheet,
        header_row=1,
        style_source_column=gamma_column,
        output_column=valid_column,
        width=14.0,
    )

    excel_rows = worksheet_rows_by_case(worksheet, file_column)
    for row_numbers in excel_rows.values():
        for row_number in row_numbers:
            _copy_style(
                worksheet.cell(row_number, gamma_column),
                worksheet.cell(row_number, valid_column),
            )
    analyzer = loop_metrics.LoopAnalyzer(
        open_threshold=args.open_threshold,
        close_threshold=args.close_threshold,
        p_column=args.p_column,
    )

    pending_updates: list[tuple[str, list[int], loop_metrics.LoopMetrics]] = []
    missing_csv_updates: list[tuple[str, list[int], Path]] = []
    unmatched_directories: list[str] = []
    errors: list[str] = []

    for case_directory in case_directories:
        case_name = case_directory.name
        matching_rows = excel_rows.get(case_name)
        if not matching_rows:
            unmatched_directories.append(case_name)
            continue

        csv_path = case_directory / RELATIVE_CSV_PATH
        if not csv_path.is_file():
            missing_csv_updates.append((case_name, matching_rows, csv_path))
            continue

        try:
            metrics = analyzer.analyze_csv(csv_path)
        except (OSError, KeyError, ValueError) as exc:
            errors.append(f"{case_name}: {csv_path}: {exc}")
            continue

        pending_updates.append((case_name, matching_rows, metrics))

    if errors:
        details = "\n".join(f"  - {error}" for error in errors)
        raise RuntimeError(
            "No workbook changes were saved because one or more selected "
            f"cases could not be evaluated:\n{details}"
        )
    if not pending_updates and not missing_csv_updates:
        raise RuntimeError(
            "No selected case directory matched a value in the worksheet "
            f"column {args.file_column!r}; no changes were saved."
        )

    valid_count = 0
    invalid_count = 0
    updated_row_count = 0

    for case_name, row_numbers, metrics in pending_updates:
        for row_number in row_numbers:
            target_cell = worksheet.cell(
                row=row_number,
                column=valid_column,
            )
            target_cell.value = metrics.valid_loop
            target_cell.number_format = "General"
            updated_row_count += 1

        if metrics.valid_loop:
            valid_count += 1
        else:
            invalid_count += 1
        print(format_result(case_name, metrics))

    for case_name, row_numbers, csv_path in missing_csv_updates:
        for row_number in row_numbers:
            target_cell = worksheet.cell(
                row=row_number,
                column=valid_column,
            )
            target_cell.value = MISSING_DATA_VALUE
            target_cell.number_format = "General"
            updated_row_count += 1
        print(
            f"[NO DATA] {case_name}: CSV not found: {csv_path}; "
            f"valid_loop={MISSING_DATA_VALUE}"
        )

    for case_name in unmatched_directories:
        print(
            f"[WARNING] {case_name}: directory matched the prefix but was not "
            f"found in worksheet column {args.file_column!r}; skipped."
        )

    if args.dry_run:
        descriptor, temporary_name = tempfile.mkstemp(
            prefix=f".{excel_path.stem}.valid_loop_preview.",
            suffix=".xlsx",
            dir=root,
        )
        os.close(descriptor)
        temporary_excel_path = Path(temporary_name)

        try:
            workbook.save(temporary_excel_path)
            build_dash.build_dashboard(
                temporary_excel_path,
                data_root=root,
                output_path=preview_html_path,
                sheet_name=args.sheet,
            )
        finally:
            temporary_excel_path.unlink(missing_ok=True)

        print(
            f"[DRY RUN] Would update {updated_row_count} worksheet row(s): "
            f"{valid_count} valid case(s), {invalid_count} invalid case(s), "
            f"{len(missing_csv_updates)} case(s) without CSV."
        )
        print(f"Preview dashboard: {preview_html_path}")
    else:
        save_workbook_atomically(workbook, output_path)
        print(
            f"Saved {output_path} | updated {updated_row_count} worksheet "
            f"row(s): {valid_count} valid case(s), "
            f"{invalid_count} invalid case(s), "
            f"{len(missing_csv_updates)} case(s) without CSV."
        )

    return 0


def main() -> int:
    args = parse_args()
    try:
        return run(args)
    except (FileNotFoundError, NotADirectoryError, KeyError, ValueError, RuntimeError) as exc:
        raise SystemExit(f"ERROR: {exc}") from exc


if __name__ == "__main__":
    raise SystemExit(main())
