#!/usr/bin/env python3
"""Add intrinsic Landau EPR columns to MFIS_dataset.xlsx.

Expected source columns in the ``experiments`` worksheet are ``alpha``,
``beta`` and ``gamma``. The script adds or updates these columns:

    Ec, Pr, Pc0, rp, landau_consistency_pass

The operation is idempotent: rerunning the script updates existing result
columns instead of appending duplicates.
"""

from __future__ import annotations

import argparse
import math
import os
import sys
import tempfile
from copy import copy
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from utils.landau_EPR_transformer import check_landau_EPR, landau_to_EPR


OUTPUT_HEADERS = (
    "Ec",
    "Pr",
    "Pc0",
    "rp",
    "landau_consistency_pass",
)


@dataclass(frozen=True, slots=True)
class UpdateSummary:
    output_path: Path
    calculated_rows: int
    passed_rows: int
    failed_rows: int
    skipped_blank_rows: int
    row_errors: tuple[str, ...]


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


def _to_float(value: Any, name: str) -> float:
    if value is None or (isinstance(value, str) and not value.strip()):
        raise ValueError(f"missing {name}")
    if isinstance(value, bool):
        raise ValueError(f"{name} is Boolean, not numeric")
    if isinstance(value, str) and value.lstrip().startswith("="):
        raise ValueError(f"{name} formula has no cached numeric value")
    try:
        result = float(value)
    except (TypeError, ValueError) as error:
        raise ValueError(f"invalid {name}: {value!r}") from error
    if not math.isfinite(result):
        raise ValueError(f"{name} must be finite; got {result!r}")
    return result


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


def add_EPR_to_workbook(
    input_path: str | Path = "MFIS_dataset.xlsx",
    output_path: str | Path | None = None,
    *,
    sheet_name: str = "experiments",
    header_row: int = 1,
    rtol: float = 1.0e-9,
    atol: float = 0.0,
    overwrite: bool = False,
) -> UpdateSummary:
    """Calculate and write EPR quantities for every valid experiments row."""

    try:
        from openpyxl import load_workbook
        from openpyxl.utils import get_column_letter
    except ImportError as error:
        raise RuntimeError(
            "openpyxl is required: python -m pip install openpyxl"
        ) from error

    source = Path(input_path).expanduser().resolve()
    if not source.is_file():
        raise FileNotFoundError(source)
    if source.suffix.casefold() not in {".xlsx", ".xlsm"}:
        raise ValueError("Only .xlsx and .xlsm workbooks are supported")
    if header_row < 1:
        raise ValueError("header_row must be at least 1")

    if output_path is None:
        destination = source.with_name(
            f"{source.stem}_with_EPR{source.suffix}"
        )
    else:
        destination = Path(output_path).expanduser().resolve()

    same_file = destination == source
    if destination.exists() and not (overwrite or same_file):
        raise FileExistsError(
            f"Output already exists: {destination}; use --overwrite"
        )

    keep_vba = source.suffix.casefold() == ".xlsm"
    workbook = load_workbook(
        source,
        data_only=False,
        keep_vba=keep_vba,
        keep_links=True,
    )
    value_workbook = load_workbook(
        source,
        data_only=True,
        read_only=True,
        keep_vba=keep_vba,
        keep_links=True,
    )

    try:
        if sheet_name not in workbook.sheetnames:
            raise KeyError(
                f"Worksheet {sheet_name!r} not found; available: {workbook.sheetnames}"
            )

        worksheet = workbook[sheet_name]
        value_worksheet = value_workbook[sheet_name]
        alpha_col = _find_header_column(worksheet, header_row, {"alpha", "α"})
        beta_col = _find_header_column(worksheet, header_row, {"beta", "β"})
        gamma_col = _find_header_column(worksheet, header_row, {"gamma", "γ"})

        old_max_column = worksheet.max_column
        output_columns: dict[str, int] = {}
        new_columns: set[int] = set()
        for header in OUTPUT_HEADERS:
            column, is_new = _ensure_output_column(
                worksheet,
                header_row,
                header,
                gamma_col,
            )
            output_columns[header] = column
            if is_new:
                new_columns.add(column)

        _extend_filter_and_tables(
            worksheet,
            header_row=header_row,
            old_max_column=old_max_column,
            new_max_column=worksheet.max_column,
        )

        widths = {
            "Ec": 18.0,
            "Pr": 18.0,
            "Pc0": 18.0,
            "rp": 13.0,
            "landau_consistency_pass": 26.0,
        }
        for header, column in output_columns.items():
            if column in new_columns:
                worksheet.column_dimensions[get_column_letter(column)].width = widths[header]

        calculated = passed = failed = skipped = 0
        errors: list[str] = []

        for row in range(header_row + 1, worksheet.max_row + 1):
            raw = (
                value_worksheet.cell(row, alpha_col).value,
                value_worksheet.cell(row, beta_col).value,
                value_worksheet.cell(row, gamma_col).value,
            )
            cells = {
                header: worksheet.cell(row, column)
                for header, column in output_columns.items()
            }

            if all(
                value is None or (isinstance(value, str) and not value.strip())
                for value in raw
            ):
                for cell in cells.values():
                    cell.value = None
                skipped += 1
                continue

            for column in new_columns:
                _copy_style(
                    worksheet.cell(row, gamma_col),
                    worksheet.cell(row, column),
                )

            try:
                alpha = _to_float(raw[0], "alpha")
                beta = _to_float(raw[1], "beta")
                gamma = _to_float(raw[2], "gamma")
                epr = landau_to_EPR(alpha, beta, gamma)
                check = check_landau_EPR(
                    alpha,
                    beta,
                    gamma,
                    epr.Ec,
                    epr.Pr,
                    epr.Pc0,
                    epr.rp,
                    rtol=rtol,
                    atol=atol,
                )

                cells["Ec"].value = epr.Ec
                cells["Pr"].value = epr.Pr
                cells["Pc0"].value = epr.Pc0
                cells["rp"].value = epr.rp
                cells["landau_consistency_pass"].value = check.passed

                for header in ("Ec", "Pr", "Pc0"):
                    cells[header].number_format = "0.000000E+00"
                cells["rp"].number_format = "0.000000"
                cells["landau_consistency_pass"].number_format = "General"

                calculated += 1
                if check.passed:
                    passed += 1
                else:
                    failed += 1
                    errors.append(
                        f"row {row}: {check.transition_order} consistency check failed"
                    )
            except (ArithmeticError, OverflowError, ValueError) as error:
                for header in ("Ec", "Pr", "Pc0", "rp"):
                    cells[header].value = None
                cells["landau_consistency_pass"].value = False
                failed += 1
                errors.append(f"row {row}: {error}")

        destination.parent.mkdir(parents=True, exist_ok=True)
        descriptor, temporary_name = tempfile.mkstemp(
            prefix=f".{destination.stem}.",
            suffix=destination.suffix,
            dir=destination.parent,
        )
        os.close(descriptor)
        temporary = Path(temporary_name)
        try:
            workbook.save(temporary)
            os.replace(temporary, destination)
        finally:
            temporary.unlink(missing_ok=True)

        return UpdateSummary(
            output_path=destination,
            calculated_rows=calculated,
            passed_rows=passed,
            failed_rows=failed,
            skipped_blank_rows=skipped,
            row_errors=tuple(errors),
        )
    finally:
        value_workbook.close()
        workbook.close()


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Add Ec, Pr, Pc0, rp and a Landau consistency flag to Excel."
    )
    parser.add_argument("--input", default="MFIS_dataset.xlsx")
    parser.add_argument(
        "--output",
        help="Default: <input_stem>_with_EPR.xlsx",
    )
    parser.add_argument(
        "--in-place",
        action="store_true",
        help="Atomically replace the input workbook.",
    )
    parser.add_argument("--overwrite", action="store_true")
    parser.add_argument("--sheet", default="experiments")
    parser.add_argument("--header-row", type=int, default=1)
    parser.add_argument("--rtol", type=float, default=1.0e-9)
    parser.add_argument("--atol", type=float, default=0.0)
    return parser


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)
    if args.in_place and args.output:
        parser.error("--in-place and --output cannot be used together")

    output = args.input if args.in_place else args.output
    try:
        summary = add_EPR_to_workbook(
            args.input,
            output,
            sheet_name=args.sheet,
            header_row=args.header_row,
            rtol=args.rtol,
            atol=args.atol,
            overwrite=args.overwrite,
        )
    except Exception as error:
        print(f"ERROR: {error}", file=sys.stderr)
        return 1

    print(f"Saved: {summary.output_path}")
    print(
        f"Rows: calculated={summary.calculated_rows}, "
        f"passed={summary.passed_rows}, failed={summary.failed_rows}, "
        f"blank/skipped={summary.skipped_blank_rows}"
    )
    for message in summary.row_errors[:20]:
        print(f"WARNING: {message}", file=sys.stderr)
    if len(summary.row_errors) > 20:
        print(
            f"WARNING: {len(summary.row_errors) - 20} additional errors omitted",
            file=sys.stderr,
        )
    return 0 if summary.failed_rows == 0 else 2


if __name__ == "__main__":
    raise SystemExit(main())
