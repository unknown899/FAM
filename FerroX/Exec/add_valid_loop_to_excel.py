#!/usr/bin/env python3
"""Add/update ``valid_loop`` in MFIS_dataset.xlsx.

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
"""

from __future__ import annotations

import argparse
import math
import os
import tempfile
from pathlib import Path

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
        )
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


def normalized_name(value: object) -> str:
    return "" if value is None else str(value).strip().casefold()


def find_header_column(
    worksheet: Worksheet,
    column_name: str,
    *,
    create: bool,
) -> int:
    """Find a row-1 header case-insensitively, optionally appending it."""

    wanted = normalized_name(column_name)
    matches = [
        cell.column
        for cell in worksheet[1]
        if normalized_name(cell.value) == wanted
    ]

    if len(matches) > 1:
        raise ValueError(
            f"Worksheet {worksheet.title!r} contains duplicate {column_name!r} headers."
        )
    if matches:
        return matches[0]
    if not create:
        available = [
            str(cell.value).strip()
            for cell in worksheet[1]
            if cell.value is not None and str(cell.value).strip()
        ]
        raise KeyError(
            f"Column {column_name!r} was not found in worksheet "
            f"{worksheet.title!r}. Available headers: {available}"
        )

    last_header_column = max(
        (
            cell.column
            for cell in worksheet[1]
            if cell.value is not None and str(cell.value).strip()
        ),
        default=0,
    )
    new_column = last_header_column + 1
    worksheet.cell(row=1, column=new_column, value=column_name)
    return new_column


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
        prefix=f".{output_path.name}.",
        suffix=".tmp",
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

    file_column = find_header_column(
        worksheet,
        args.file_column,
        create=False,
    )
    valid_column = find_header_column(
        worksheet,
        args.valid_column,
        create=True,
    )
    excel_rows = worksheet_rows_by_case(worksheet, file_column)
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
            worksheet.cell(
                row=row_number,
                column=valid_column,
                value=metrics.valid_loop,
            )
            updated_row_count += 1

        if metrics.valid_loop:
            valid_count += 1
        else:
            invalid_count += 1
        print(format_result(case_name, metrics))

    for case_name, row_numbers, csv_path in missing_csv_updates:
        for row_number in row_numbers:
            worksheet.cell(
                row=row_number,
                column=valid_column,
                value=MISSING_DATA_VALUE,
            )
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
