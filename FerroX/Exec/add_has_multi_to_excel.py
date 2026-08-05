#!/usr/bin/env python3
"""Add/update ``has_multi`` in MFIS_dataset.xlsx from extracted Pz data.

The default paths assume this script is executed from ``FerroX/Exec``::

    python add_has_multi_to_excel.py --var-threshold 0.1

For each ``extracted_pz/<case>/Pz_Phi_FE_all_voltage.npz``, the script loads
``Pz_stack`` and inspects ``Pz_stack[voltage_index, :, z_index]``.  A case is
marked as ``has_multi = 1`` as soon as one voltage point satisfies both:

1. ``abs(P_max - P_min) > var_threshold``
2. ``P_max > 0`` and ``P_min < 0``

If no voltage point satisfies the conditions, the value is 0.  Excel cases
without a corresponding NPZ file are marked with ``沒資料``.
"""

from __future__ import annotations

import argparse
import math
import os
import shutil
import tempfile
from dataclasses import dataclass
from pathlib import Path
from typing import Final

import numpy as np
from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet


DEFAULT_EXCEL_PATH: Final = Path("MFIS_dataset.xlsx")
DEFAULT_EXTRACTED_PZ_DIR: Final = Path("extracted_pz")
DEFAULT_SHEET_NAME: Final = "experiments"
DEFAULT_NPZ_NAME: Final = "Pz_Phi_FE_all_voltage.npz"
DEFAULT_Z_INDEX: Final = 5
DEFAULT_MISSING_LABEL: Final = "沒資料"

# Checked in this order when --case-column is not supplied.
CASE_COLUMN_CANDIDATES: Final = (
    "file_name",
    "folder",
    "folder_name",
    "case_key",
    "case",
)


@dataclass(frozen=True)
class MultiResult:
    has_multi: int
    voltage_index: int | None = None
    p_min: float | None = None
    p_max: float | None = None


def normalized_header(value: object) -> str:
    """Normalize a worksheet header for case-insensitive matching."""
    if value is None:
        return ""
    return str(value).strip().casefold()


def normalized_case_name(value: object) -> str:
    """Convert an Excel case-name cell to the folder-name lookup key."""
    if value is None:
        return ""
    return str(value).strip()


def analyze_pz_stack(
    pz_stack: np.ndarray,
    var_threshold: float,
    z_index: int = DEFAULT_Z_INDEX,
) -> MultiResult:
    """Return whether any voltage slice contains positive/negative domains."""
    pz_stack = np.asarray(pz_stack)

    if pz_stack.ndim != 3:
        raise ValueError(
            f"Pz_stack must be 3-D, but its shape is {pz_stack.shape}."
        )
    if not 0 <= z_index < pz_stack.shape[2]:
        raise IndexError(
            f"z_index={z_index} is outside the third dimension "
            f"of Pz_stack with shape {pz_stack.shape}."
        )

    found_finite_value = False

    for voltage_index in range(pz_stack.shape[0]):
        try:
            line = np.asarray(
                pz_stack[voltage_index, :, z_index],
                dtype=float,
            )
        except (TypeError, ValueError) as exc:
            raise ValueError(
                "Pz_stack cannot be converted to numeric values."
            ) from exc

        finite_line = line[np.isfinite(line)]
        if finite_line.size == 0:
            continue

        found_finite_value = True
        p_min = float(np.min(finite_line))
        p_max = float(np.max(finite_line))

        if (
            abs(p_max - p_min) > var_threshold
            and p_max > 0.0
            and p_min < 0.0
        ):
            return MultiResult(
                has_multi=1,
                voltage_index=voltage_index,
                p_min=p_min,
                p_max=p_max,
            )

    if not found_finite_value:
        raise ValueError(
            "No finite value exists in any selected Pz_stack slice."
        )

    return MultiResult(has_multi=0)


def analyze_npz(
    npz_path: Path,
    var_threshold: float,
    z_index: int,
) -> MultiResult:
    """Load and analyze one NPZ file."""
    try:
        with np.load(npz_path, allow_pickle=True) as data:
            if "Pz_stack" not in data.files:
                raise KeyError("NPZ does not contain the key 'Pz_stack'.")
            pz_stack = data["Pz_stack"]
    except Exception as exc:
        raise RuntimeError(f"Failed to load {npz_path}: {exc}") from exc

    try:
        return analyze_pz_stack(
            pz_stack=pz_stack,
            var_threshold=var_threshold,
            z_index=z_index,
        )
    except Exception as exc:
        raise RuntimeError(f"Failed to analyze {npz_path}: {exc}") from exc


def discover_npz_files(
    extracted_pz_dir: Path,
    npz_name: str,
) -> dict[str, Path]:
    """Find every target NPZ and map its parent folder name to its path."""
    if not extracted_pz_dir.is_dir():
        raise FileNotFoundError(
            f"Extracted-Pz directory does not exist: {extracted_pz_dir}"
        )

    result: dict[str, Path] = {}
    for npz_path in sorted(extracted_pz_dir.rglob(npz_name)):
        case_name = npz_path.parent.name.strip()
        if not case_name:
            raise ValueError(f"Cannot determine the case name for {npz_path}.")
        if case_name in result:
            raise ValueError(
                f"Duplicate NPZ files found for case {case_name!r}:\n"
                f"  {result[case_name]}\n"
                f"  {npz_path}"
            )
        result[case_name] = npz_path

    return result


def find_column(
    worksheet: Worksheet,
    header_row: int,
    requested_name: str | None,
) -> tuple[int, str]:
    """Find the experiment case-name column."""
    header_to_columns: dict[str, list[int]] = {}
    for column_index in range(1, worksheet.max_column + 1):
        header = normalized_header(
            worksheet.cell(row=header_row, column=column_index).value
        )
        if header:
            header_to_columns.setdefault(header, []).append(column_index)

    names_to_try = (
        (normalized_header(requested_name),)
        if requested_name is not None
        else CASE_COLUMN_CANDIDATES
    )

    for name in names_to_try:
        matches = header_to_columns.get(name, [])
        if len(matches) == 1:
            original = worksheet.cell(
                row=header_row,
                column=matches[0],
            ).value
            return matches[0], str(original)
        if len(matches) > 1:
            raise ValueError(
                f"Worksheet contains more than one column named {name!r}."
            )

    available = [
        str(worksheet.cell(row=header_row, column=i).value)
        for i in range(1, worksheet.max_column + 1)
        if worksheet.cell(row=header_row, column=i).value is not None
    ]
    if requested_name is not None:
        wanted = repr(requested_name)
    else:
        wanted = "one of " + ", ".join(repr(x) for x in CASE_COLUMN_CANDIDATES)
    raise ValueError(
        f"Could not find case-name column {wanted} in row {header_row}. "
        f"Available columns: {available}"
    )


def find_or_create_has_multi_column(
    worksheet: Worksheet,
    header_row: int,
) -> int:
    """Return the existing has_multi column or append a new one."""
    matches = [
        column_index
        for column_index in range(1, worksheet.max_column + 1)
        if normalized_header(
            worksheet.cell(row=header_row, column=column_index).value
        )
        == "has_multi"
    ]
    if len(matches) > 1:
        raise ValueError("Worksheet contains more than one 'has_multi' column.")
    if matches:
        return matches[0]

    new_column = worksheet.max_column + 1
    worksheet.cell(row=header_row, column=new_column, value="has_multi")
    return new_column


def save_workbook_atomically(workbook, excel_path: Path) -> None:
    """Save beside the source file and atomically replace it on success."""
    temp_path: Path | None = None
    try:
        with tempfile.NamedTemporaryFile(
            prefix=f".{excel_path.stem}_",
            suffix=excel_path.suffix,
            dir=excel_path.parent,
            delete=False,
        ) as temp_file:
            temp_path = Path(temp_file.name)

        workbook.save(temp_path)
        shutil.copymode(excel_path, temp_path)
        os.replace(temp_path, excel_path)
    finally:
        if temp_path is not None and temp_path.exists():
            temp_path.unlink()


def update_excel(
    excel_path: Path,
    sheet_name: str,
    header_row: int,
    case_column_name: str | None,
    results: dict[str, MultiResult],
    missing_label: str,
    dry_run: bool,
) -> tuple[int, int, set[str]]:
    """Update all non-empty experiment rows and return summary counts."""
    if not excel_path.is_file():
        raise FileNotFoundError(f"Excel file does not exist: {excel_path}")

    workbook = load_workbook(excel_path)
    try:
        if sheet_name not in workbook.sheetnames:
            raise KeyError(
                f"Worksheet {sheet_name!r} does not exist. "
                f"Available worksheets: {workbook.sheetnames}"
            )

        worksheet = workbook[sheet_name]
        case_column, detected_name = find_column(
            worksheet=worksheet,
            header_row=header_row,
            requested_name=case_column_name,
        )
        has_multi_column = find_or_create_has_multi_column(
            worksheet=worksheet,
            header_row=header_row,
        )

        matched_cases: set[str] = set()
        updated_count = 0
        missing_count = 0

        for row_index in range(header_row + 1, worksheet.max_row + 1):
            case_name = normalized_case_name(
                worksheet.cell(row=row_index, column=case_column).value
            )
            if not case_name:
                continue

            result = results.get(case_name)
            if result is None:
                value: int | str = missing_label
                missing_count += 1
            else:
                value = result.has_multi
                matched_cases.add(case_name)

            worksheet.cell(
                row=row_index,
                column=has_multi_column,
                value=value,
            )
            updated_count += 1

        if not dry_run:
            save_workbook_atomically(workbook, excel_path)

        print(f"Case-name column: {detected_name}")
        print(f"Updated experiment rows: {updated_count}")
        print(f"Rows marked {missing_label!r}: {missing_count}")
        return updated_count, missing_count, matched_cases
    finally:
        workbook.close()


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Set experiments.has_multi using Pz_stack data under extracted_pz."
        )
    )
    parser.add_argument(
        "--var-threshold",
        type=float,
        required=True,
        help="Required minimum abs(P_max - P_min); comparison is strict (>).",
    )
    parser.add_argument(
        "--excel-path",
        type=Path,
        default=DEFAULT_EXCEL_PATH,
        help=f"Excel workbook path (default: {DEFAULT_EXCEL_PATH}).",
    )
    parser.add_argument(
        "--extracted-pz-dir",
        type=Path,
        default=DEFAULT_EXTRACTED_PZ_DIR,
        help=(
            "Directory containing one subdirectory per case "
            f"(default: {DEFAULT_EXTRACTED_PZ_DIR})."
        ),
    )
    parser.add_argument(
        "--sheet-name",
        default=DEFAULT_SHEET_NAME,
        help=f"Worksheet name (default: {DEFAULT_SHEET_NAME}).",
    )
    parser.add_argument(
        "--header-row",
        type=int,
        default=1,
        help="One-based worksheet header row (default: 1).",
    )
    parser.add_argument(
        "--case-column",
        default=None,
        help=(
            "Experiment case-name column. If omitted, detect file_name, "
            "folder, folder_name, case_key, or case automatically."
        ),
    )
    parser.add_argument(
        "--npz-name",
        default=DEFAULT_NPZ_NAME,
        help=f"NPZ basename to discover recursively (default: {DEFAULT_NPZ_NAME}).",
    )
    parser.add_argument(
        "--z-index",
        type=int,
        default=DEFAULT_Z_INDEX,
        help="Index selected from the third Pz_stack dimension (default: 5).",
    )
    parser.add_argument(
        "--missing-label",
        default=DEFAULT_MISSING_LABEL,
        help=f"Value for Excel cases without NPZ data (default: {DEFAULT_MISSING_LABEL}).",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Analyze and report results without saving the workbook.",
    )
    args = parser.parse_args()

    if not math.isfinite(args.var_threshold) or args.var_threshold < 0.0:
        parser.error("--var-threshold must be a finite number >= 0.")
    if args.header_row < 1:
        parser.error("--header-row must be >= 1.")
    if args.z_index < 0:
        parser.error("--z-index must be >= 0.")
    if not args.npz_name.strip():
        parser.error("--npz-name cannot be empty.")

    return args


def main() -> None:
    args = parse_args()

    npz_paths = discover_npz_files(
        extracted_pz_dir=args.extracted_pz_dir,
        npz_name=args.npz_name,
    )
    print(f"Discovered NPZ files: {len(npz_paths)}")

    # Analyze every discovered NPZ before opening/writing Excel.  Thus a bad
    # input file cannot leave the workbook only partially updated.
    results: dict[str, MultiResult] = {}
    for case_name, npz_path in npz_paths.items():
        result = analyze_npz(
            npz_path=npz_path,
            var_threshold=args.var_threshold,
            z_index=args.z_index,
        )
        results[case_name] = result

        if result.has_multi:
            assert result.voltage_index is not None
            assert result.p_min is not None
            assert result.p_max is not None
            variation = abs(result.p_max - result.p_min)
            print(
                f"[1] {case_name}: voltage_index={result.voltage_index}, "
                f"P_min={result.p_min:.8g}, P_max={result.p_max:.8g}, "
                f"variation={variation:.8g}"
            )
        else:
            print(f"[0] {case_name}")

    _, _, matched_cases = update_excel(
        excel_path=args.excel_path,
        sheet_name=args.sheet_name,
        header_row=args.header_row,
        case_column_name=args.case_column,
        results=results,
        missing_label=args.missing_label,
        dry_run=args.dry_run,
    )

    unmatched_npz_cases = sorted(set(results) - matched_cases)
    if unmatched_npz_cases:
        print(
            "Warning: the following NPZ cases were analyzed but do not have "
            "a matching Excel row:"
        )
        for case_name in unmatched_npz_cases:
            print(f"  - {case_name}")

    if args.dry_run:
        print("Dry run complete; Excel was not modified.")
    else:
        print(f"Saved: {args.excel_path}")


if __name__ == "__main__":
    main()
