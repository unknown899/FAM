"""Reusable detection of mixed positive/negative polarization domains.

``MultiAnalyzer`` keeps the threshold and z index together and can analyze
either an already-loaded ``Pz_stack`` array or one NPZ file.
"""

from __future__ import annotations

import math
import operator
from dataclasses import dataclass
from pathlib import Path

import numpy as np


DEFAULT_NPZ_KEY = "Pz_stack"
DEFAULT_Z_INDEX: int | None = None


@dataclass(frozen=True)
class MultiResult:
    has_multi: int
    voltage_index: int | None = None
    z_index: int | None = None
    p_min: float | None = None
    p_max: float | None = None

    @property
    def variation(self) -> float | None:
        if self.p_min is None or self.p_max is None:
            return None
        return abs(self.p_max - self.p_min)


@dataclass(frozen=True)
class MultiAnalyzer:
    """Detect mixed domains using one reusable set of criteria."""

    var_threshold: float
    z_index: int | None = DEFAULT_Z_INDEX
    npz_key: str = DEFAULT_NPZ_KEY

    def __post_init__(self) -> None:
        try:
            threshold = float(self.var_threshold)
        except (TypeError, ValueError) as exc:
            raise ValueError(
                f"var_threshold must be a number, got {self.var_threshold!r}"
            ) from exc
        if not math.isfinite(threshold) or threshold < 0:
            raise ValueError(
                "var_threshold must be finite and >= 0, "
                f"got {self.var_threshold!r}"
            )
        object.__setattr__(self, "var_threshold", threshold)

        if self.z_index is not None:
            try:
                z_index = operator.index(self.z_index)
            except TypeError as exc:
                raise ValueError(
                    f"z_index must be an integer or None, got {self.z_index!r}"
                ) from exc

            if z_index < 0:
                raise ValueError(f"z_index must be >= 0, got {z_index}")

            object.__setattr__(self, "z_index", z_index)

        if not isinstance(self.npz_key, str) or not self.npz_key.strip():
            raise ValueError("npz_key must be a non-empty string")

    def analyze_pz_stack(self, pz_stack: np.ndarray) -> MultiResult:
        """Count voltage points having at least one mixed-domain z slice."""

        stack = np.asarray(pz_stack)

        if stack.ndim != 3:
            raise ValueError(
                f"Pz_stack must be 3-D, but its shape is {stack.shape}."
            )

        if self.z_index is None:
            z_indices = range(stack.shape[2])
        else:
            if self.z_index >= stack.shape[2]:
                raise IndexError(
                    f"z_index={self.z_index} is outside the third dimension "
                    f"of Pz_stack with shape {stack.shape}."
                )
            z_indices = (self.z_index,)

        found_finite_value = False
        multi_voltage_count = 0

        # 保留第一個通過的 slice，供輸出詳細資訊使用
        first_match: tuple[int, int, float, float] | None = None

        for voltage_index in range(stack.shape[0]):
            for z_index in z_indices:
                try:
                    line = np.asarray(
                        stack[voltage_index, :, z_index],
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
                    abs(p_max - p_min) > self.var_threshold
                    and p_max > 0
                    and p_min < 0
                ):
                    multi_voltage_count += 1

                    if first_match is None:
                        first_match = (
                            voltage_index,
                            z_index,
                            p_min,
                            p_max,
                        )

                    # 此 voltage 已確定有 MD，不再檢查其他 z
                    break

        if not found_finite_value:
            raise ValueError(
                "No finite value exists in any selected Pz_stack slice."
            )

        if first_match is None:
            return MultiResult(has_multi=0)

        voltage_index, z_index, p_min, p_max = first_match

        return MultiResult(
            has_multi=multi_voltage_count,
            voltage_index=voltage_index,
            z_index=z_index,
            p_min=p_min,
            p_max=p_max,
        )

    def analyze_npz(self, npz_path: str | Path) -> MultiResult:
        """Load ``npz_key`` from one NPZ file and analyze its Pz stack."""

        path = Path(npz_path)
        try:
            with np.load(path, allow_pickle=True) as data:
                if self.npz_key not in data.files:
                    raise KeyError(
                        f"NPZ does not contain the key {self.npz_key!r}."
                    )
                pz_stack = data[self.npz_key]
        except (OSError, ValueError, KeyError) as exc:
            raise RuntimeError(f"Failed to load {path}: {exc}") from exc

        try:
            return self.analyze_pz_stack(pz_stack)
        except (IndexError, TypeError, ValueError) as exc:
            raise RuntimeError(f"Failed to analyze {path}: {exc}") from exc
