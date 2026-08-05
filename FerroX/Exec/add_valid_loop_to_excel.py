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

Only worksheet rows whose ``file_name`` matches a selected directory are
changed. Other rows are left untouched.
"""

from __future__ import annotations

import argparse
import csv
import math
import os
import tempfile
from dataclasses import dataclass
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet


DEFAULT_EXCEL_NAME = "MFIS_dataset.xlsx"
DEFAULT_SHEET_NAME = "experiments"
DEFAULT_FILE_COLUMN = "file_name"
DEFAULT_VALID_COLUMN = "valid_loop"
DEFAULT_P_COLUMN = "P_mean"
RELATIVE_CSV_PATH = Path("figs") / "MFIS_PV_curve.csv"


@dataclass(frozen=True)
class LoopMetrics:
    """Polarization values and the resulting validity checks for one case."""

    p_start: float
    p_r_minus: float
    p_r_plus: float
    p_end: float
    open_gap: float
    close_gap: float
    open_ok: bool
    sign_ok: bool
    close_ok: bool

    @property
    def valid_loop(self) -> int:
        return int(self.open_ok and self.sign_ok and self.close_ok)


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
        type=Path,
        default=Path(DEFAULT_EXCEL_NAME),
        help=f"Input workbook, relative to --root unless absolute (default: {DEFAULT_EXCEL_NAME}).",
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
        default=DEFAULT_P_COLUMN,
        help=f"Polarization column in each CSV (default: {DEFAULT_P_COLUMN}).",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Calculate and print results without writing the workbook.",
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


def read_p_mean_rows(csv_path: Path, p_column: str) -> list[float]:
    """Read all nonblank data rows from the requested P_mean column."""

    with csv_path.open("r", encoding="utf-8-sig", newline="") as handle:
        reader = csv.DictReader(handle)
        if reader.fieldnames is None:
            raise ValueError("CSV has no header row.")

        header_lookup = {
            normalized_name(header): header
            for header in reader.fieldnames
            if header is not None
        }
        actual_p_column = header_lookup.get(normalized_name(p_column))
        if actual_p_column is None:
            raise KeyError(
                f"Column {p_column!r} was not found. "
                f"Available columns: {reader.fieldnames}"
            )

        values: list[float] = []
        for csv_row_number, row in enumerate(reader, start=2):
            if not any(
                value is not None and str(value).strip()
                for value in row.values()
            ):
                continue

            raw_value = row.get(actual_p_column)
            try:
                value = float(str(raw_value).strip())
            except (TypeError, ValueError) as exc:
                raise ValueError(
                    f"Invalid {actual_p_column} value at CSV row "
                    f"{csv_row_number}: {raw_value!r}"
                ) from exc
            if not math.isfinite(value):
                raise ValueError(
                    f"Non-finite {actual_p_column} value at CSV row "
                    f"{csv_row_number}: {raw_value!r}"
                )
            values.append(value)

    if len(values) < 37:
        raise ValueError(
            f"At least 37 data rows are required, but only {len(values)} were found."
        )
    return values


def calculate_loop_metrics(
    p_mean_values: list[float],
    open_threshold: float,
    close_threshold: float,
) -> LoopMetrics:
    # User-facing row numbers are 1-based; list indices are 0-based.
    p_start = -p_mean_values[0]
    p_r_minus = -p_mean_values[9]
    p_r_plus = -p_mean_values[27]
    p_end = -p_mean_values[36]

    open_gap = abs(p_r_plus - p_r_minus)
    close_gap = abs(p_start - p_end)

    return LoopMetrics(
        p_start=p_start,
        p_r_minus=p_r_minus,
        p_r_plus=p_r_plus,
        p_end=p_end,
        open_gap=open_gap,
        close_gap=close_gap,
        open_ok=open_gap > open_threshold,
        sign_ok=p_r_plus > 0 and p_r_minus < 0,
        close_ok=close_gap < close_threshold,
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


def format_result(case_name: str, metrics: LoopMetrics) -> str:
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

    pending_updates: list[tuple[str, list[int], LoopMetrics]] = []
    unmatched_directories: list[str] = []
    errors: list[str] = []
    missing_csv_row_count = 0
    missing_csv_case_count = 0

    for case_directory in case_directories:
        case_name = case_directory.name
        matching_rows = excel_rows.get(case_name)
        if not matching_rows:
            unmatched_directories.append(case_name)
            continue

        csv_path = case_directory / RELATIVE_CSV_PATH
        if not csv_path.is_file():
            for row_number in matching_rows:
                worksheet.cell(
                    row=row_number,
                    column=valid_column,
                    value="no_data",
                )
                missing_csv_row_count += 1

            missing_csv_case_count += 1
            print(f"[NO DATA] {case_name}: CSV not found: {csv_path}")
            continue
        try:
            p_mean_values = read_p_mean_rows(csv_path, args.p_column)
            metrics = calculate_loop_metrics(
                p_mean_values,
                args.open_threshold,
                args.close_threshold,
            )
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
    if not pending_updates and missing_csv_row_count == 0:
        raise RuntimeError(
            "No selected case directory matched a value in the worksheet "
            f"column {args.file_column!r}; no changes were saved."
        )

    valid_count = 0
    invalid_count = 0
    updated_row_count = missing_csv_row_count

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

    for case_name in unmatched_directories:
        print(
            f"[WARNING] {case_name}: directory matched the prefix but was not "
            f"found in worksheet column {args.file_column!r}; skipped."
        )

    if args.dry_run:
        print(
            f"[DRY RUN] Would update {updated_row_count} worksheet row(s): "
            f"{valid_count} valid case(s), {invalid_count} invalid case(s), "
            f"{missing_csv_case_count} case(s) without CSV."
        )
    else:
        save_workbook_atomically(workbook, output_path)
        print(
            f"Saved {output_path} | updated {updated_row_count} worksheet "
            f"row(s): {valid_count} valid case(s), "
            f"{invalid_count} invalid case(s),"
            f"{missing_csv_case_count} case(s) without CSV."
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
