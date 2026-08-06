"""Reusable P-V loop metric calculation.

``LoopAnalyzer`` keeps the thresholds together and can analyze either an
already-loaded P_mean sequence or an ``MFIS_PV_curve.csv`` file.
"""

from __future__ import annotations

import csv
import math
from collections.abc import Sequence
from dataclasses import dataclass
from pathlib import Path


DEFAULT_P_COLUMN = "P_mean"

# Data-row positions in MFIS_PV_curve.csv (header excluded).
_P_START_INDEX = 0
_P_R_MINUS_INDEX = 9
_P_R_PLUS_INDEX = 27
_P_END_INDEX = 36
_REQUIRED_DATA_ROWS = _P_END_INDEX + 1


@dataclass(frozen=True)
class LoopMetrics:
    """Polarization values and validity checks for one P-V loop."""

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


@dataclass(frozen=True)
class LoopAnalyzer:
    """Read and evaluate P-V loops using one reusable set of criteria."""

    open_threshold: float
    close_threshold: float
    p_column: str = DEFAULT_P_COLUMN

    def __post_init__(self) -> None:
        for name in ("open_threshold", "close_threshold"):
            raw_value = getattr(self, name)
            try:
                value = float(raw_value)
            except (TypeError, ValueError) as exc:
                raise ValueError(f"{name} must be a number, got {raw_value!r}") from exc
            if not math.isfinite(value) or value < 0:
                raise ValueError(f"{name} must be finite and >= 0, got {raw_value!r}")
            object.__setattr__(self, name, value)

        if not isinstance(self.p_column, str) or not self.p_column.strip():
            raise ValueError("p_column must be a non-empty string")

    def read_p_mean_rows(self, csv_path: str | Path) -> list[float]:
        """Read all nonblank values from the selected polarization column."""

        path = Path(csv_path)
        with path.open("r", encoding="utf-8-sig", newline="") as handle:
            reader = csv.DictReader(handle)
            if reader.fieldnames is None:
                raise ValueError("CSV has no header row.")

            header_lookup = {
                _normalized_name(header): header
                for header in reader.fieldnames
                if header is not None
            }
            actual_p_column = header_lookup.get(_normalized_name(self.p_column))
            if actual_p_column is None:
                raise KeyError(
                    f"Column {self.p_column!r} was not found. "
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

        return values

    def calculate_loop_metrics(
        self,
        p_mean_values: Sequence[float],
    ) -> LoopMetrics:
        """Calculate metrics from P_mean values in sweep order."""

        if len(p_mean_values) < _REQUIRED_DATA_ROWS:
            raise ValueError(
                f"At least {_REQUIRED_DATA_ROWS} data rows are required, "
                f"but only {len(p_mean_values)} were found."
            )

        def flipped_value(index: int, label: str) -> float:
            raw_value = p_mean_values[index]
            try:
                value = -float(raw_value)
            except (TypeError, ValueError) as exc:
                raise ValueError(f"Invalid {label} value: {raw_value!r}") from exc
            if not math.isfinite(value):
                raise ValueError(f"Non-finite {label} value: {raw_value!r}")
            return value

        p_start = flipped_value(_P_START_INDEX, "P_start")
        p_r_minus = flipped_value(_P_R_MINUS_INDEX, "P_r_minus")
        p_r_plus = flipped_value(_P_R_PLUS_INDEX, "P_r_plus")
        p_end = flipped_value(_P_END_INDEX, "P_end")

        open_gap = abs(p_r_plus - p_r_minus)
        close_gap = abs(p_start - p_end)

        return LoopMetrics(
            p_start=p_start,
            p_r_minus=p_r_minus,
            p_r_plus=p_r_plus,
            p_end=p_end,
            open_gap=open_gap,
            close_gap=close_gap,
            open_ok=open_gap > self.open_threshold,
            sign_ok=p_r_plus > 0 and p_r_minus < 0,
            close_ok=close_gap < self.close_threshold,
        )

    def analyze_csv(self, csv_path: str | Path) -> LoopMetrics:
        """Read one CSV and return its calculated loop metrics."""

        return self.calculate_loop_metrics(self.read_p_mean_rows(csv_path))


def _normalized_name(value: object) -> str:
    return "" if value is None else str(value).strip().casefold()
