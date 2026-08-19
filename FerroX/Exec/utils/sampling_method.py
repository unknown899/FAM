"""Reusable EPR-space sampling utilities for FerroX batch generation.

This module deliberately knows nothing about FerroX, Landau conversion, Excel,
or folder creation.  It only generates points in the intrinsic parameter space
``(Ec, Pr, rp)`` and therefore can also be imported from notebooks.

Supported global designs
------------------------
``uniform``
    Independent random draws in each normalized coordinate.
``lhs``
    Latin-hypercube sampling.
``sobol``
    Scrambled Sobol low-discrepancy sampling (requires SciPy).
``maximin``
    Generate a large LHS pool and greedily select points that maximize their
    nearest-neighbour distance from supplied reference points and from points
    already selected in the same batch.
``screening``
    Deterministic coarse range screening for expensive FerroX runs.  The first
    15 accepted points prioritize the box center, 8 corners, and 6 face centers.
    If more points are requested, the design is augmented from a 5-level lattice
    using maximin distance in normalized EPR coordinates.

Supported local / variation designs
-----------------------------------
``variation`` is represented by :class:`VariationSpec` and can use:
``oat``
    One-at-a-time perturbations around an anchor.  Good for local sensitivity
    and finite-difference/Jacobian studies.
``uniform``, ``lhs``, ``sobol``, ``maximin``
    Jointly vary Ec, Pr and rp inside a local box around an anchor.  These are
    better than OAT when the goal is to add training data that captures
    parameter interactions.

Ec and Pr can be sampled on a linear or logarithmic coordinate.  rp is kept
linear because its physically allowed interval is narrow and bounded.
"""

from __future__ import annotations

import math
from dataclasses import dataclass
from typing import Callable, Iterable, Sequence
from pathlib import Path
import pandas as pd

import numpy as np


EPR_PARAMETER_NAMES = ("ec", "pr", "rp")
GLOBAL_DESIGNS = ("uniform", "lhs", "sobol", "maximin", "screening","f_thickness")
VARIATION_MODES = ("oat", "uniform", "lhs", "sobol", "maximin")


@dataclass(frozen=True)
class EPRPoint:
    """One intrinsic EPR point and a human-readable design label."""

    Ec: float
    Pr: float
    rp: float
    design: str

    def as_tuple(self) -> tuple[float, float, float]:
        return (self.Ec, self.Pr, self.rp)

    def as_dict(self) -> dict[str, float]:
        return {"ec": self.Ec, "pr": self.Pr, "rp": self.rp}


@dataclass(frozen=True)
class EPRBounds:
    """Sampling bounds for intrinsic Ec, Pr and rp."""

    ec: tuple[float, float]
    pr: tuple[float, float]
    rp: tuple[float, float]

    def validate_positive(self) -> None:
        for name, bounds in (
            ("Ec", self.ec),
            ("Pr", self.pr),
            ("rp", self.rp),
        ):
            low, high = bounds
            if not (math.isfinite(low) and math.isfinite(high)):
                raise ValueError(f"{name} bounds must be finite; got {bounds}")
            if low <= 0.0 or high <= low:
                raise ValueError(
                    f"{name} bounds must satisfy 0 < LOW < HIGH; got {bounds}"
                )

    def contains(self, point: EPRPoint | Sequence[float], *, atol: float = 0.0) -> bool:
        if isinstance(point, EPRPoint):
            values = point.as_tuple()
        else:
            values = tuple(float(value) for value in point)
        return all(
            low - atol <= value <= high + atol
            for value, (low, high) in zip(values, (self.ec, self.pr, self.rp), strict=True)
        )


@dataclass(frozen=True)
class VariationSpec:
    """Definition of a local EPR variation around one anchor.

    Parameters
    ----------
    anchor:
        ``(Ec, Pr, rp)`` center of the local study.
    mode:
        ``oat`` for one-at-a-time perturbations, or one of the joint sampling
        modes ``uniform/lhs/sobol/maximin``.
    ec_fraction, pr_fraction:
        Fractional half-width around the anchor.  For example 0.10 means
        ``Ec0*(1-0.10)`` to ``Ec0*(1+0.10)``.
    rp_delta:
        Absolute half-width in rp.  For example 0.02 means ``rp0 +/- 0.02``.
    parameters:
        Parameters varied by OAT. Joint modes sample all three coordinates
        inside the configured local box.
    levels:
        OAT perturbation levels.  With fraction 0.10 and levels (0.5, 1.0), Ec
        uses +/-5% and +/-10% perturbations.
    include_center:
        Include the anchor itself in an OAT design.
    """

    anchor: tuple[float, float, float]
    mode: str = "oat"
    ec_fraction: float = 0.10
    pr_fraction: float = 0.10
    rp_delta: float = 0.02
    parameters: tuple[str, ...] = EPR_PARAMETER_NAMES
    levels: tuple[float, ...] = (1.0,)
    include_center: bool = False

    def validate(self) -> None:
        if self.mode not in VARIATION_MODES:
            raise ValueError(
                f"Unknown variation mode {self.mode!r}; choose from {VARIATION_MODES}"
            )
        Ec, Pr, rp = self.anchor
        if not all(math.isfinite(v) and v > 0.0 for v in (Ec, Pr, rp)):
            raise ValueError(f"Variation anchor must contain positive finite values; got {self.anchor}")
        if self.ec_fraction < 0.0 or not math.isfinite(self.ec_fraction):
            raise ValueError("ec_fraction must be finite and >= 0")
        if self.pr_fraction < 0.0 or not math.isfinite(self.pr_fraction):
            raise ValueError("pr_fraction must be finite and >= 0")
        if self.rp_delta < 0.0 or not math.isfinite(self.rp_delta):
            raise ValueError("rp_delta must be finite and >= 0")
        if not self.parameters:
            raise ValueError("At least one variation parameter is required")
        bad_parameters = [p for p in self.parameters if p not in EPR_PARAMETER_NAMES]
        if bad_parameters:
            raise ValueError(
                f"Unknown variation parameters {bad_parameters}; choose from {EPR_PARAMETER_NAMES}"
            )
        if not self.levels:
            raise ValueError("At least one OAT variation level is required")
        if any((not math.isfinite(level) or level <= 0.0) for level in self.levels):
            raise ValueError("All OAT variation levels must be positive finite numbers")


def lhs_unit(n_samples: int, dimensions: int, rng: np.random.Generator) -> np.ndarray:
    """Dependency-free Latin-hypercube points in ``[0, 1)``."""
    if n_samples <= 0:
        raise ValueError("n_samples must be positive")
    if dimensions <= 0:
        raise ValueError("dimensions must be positive")
    points = np.empty((n_samples, dimensions), dtype=float)
    for dimension in range(dimensions):
        values = (np.arange(n_samples) + rng.random(n_samples)) / n_samples
        rng.shuffle(values)
        points[:, dimension] = values
    return points


def sobol_unit(n_samples: int, dimensions: int, seed: int) -> np.ndarray:
    """Return scrambled Sobol points in ``[0, 1)``.

    SciPy is imported lazily so users who only need uniform/LHS/maximin do not
    acquire an additional hard dependency.
    """
    if n_samples <= 0:
        raise ValueError("n_samples must be positive")
    if dimensions <= 0:
        raise ValueError("dimensions must be positive")
    try:
        from scipy.stats import qmc
    except ImportError as exc:  # pragma: no cover - depends on user environment
        raise ImportError(
            "Sobol sampling requires SciPy. Install it with `conda install scipy` "
            "or choose uniform/lhs/maximin."
        ) from exc

    exponent = int(math.ceil(math.log2(max(n_samples, 1))))
    sampler = qmc.Sobol(d=dimensions, scramble=True, seed=seed)
    points = sampler.random_base2(m=exponent)
    return np.asarray(points[:n_samples], dtype=float)


def scale_unit_value(
    unit_value: float,
    bounds: tuple[float, float],
    scale: str,
) -> float:
    """Map one value in ``[0, 1]`` to a linear or logarithmic interval."""
    low, high = bounds
    if scale == "linear":
        return low + unit_value * (high - low)
    if scale == "log":
        if low <= 0.0:
            raise ValueError(f"Log sampling requires positive lower bound; got {bounds}")
        return math.exp(math.log(low) + unit_value * (math.log(high) - math.log(low)))
    raise ValueError(f"Unknown sampling scale: {scale!r}")


def normalize_value(
    value: float,
    bounds: tuple[float, float],
    scale: str,
) -> float:
    """Map one physical value to its normalized sampling coordinate."""
    low, high = bounds
    if high <= low:
        # A zero-width local variation means the parameter is intentionally fixed.
        return 0.5
    if scale == "linear":
        return (value - low) / (high - low)
    if scale == "log":
        if low <= 0.0 or value <= 0.0:
            raise ValueError("Log normalization requires positive values")
        return (math.log(value) - math.log(low)) / (math.log(high) - math.log(low))
    raise ValueError(f"Unknown sampling scale: {scale!r}")


def unit_point_to_epr(
    point: Sequence[float],
    bounds: EPRBounds,
    ec_scale: str,
    pr_scale: str,
    *,
    design: str,
) -> EPRPoint:
    return EPRPoint(
        Ec=scale_unit_value(float(point[0]), bounds.ec, ec_scale),
        Pr=scale_unit_value(float(point[1]), bounds.pr, pr_scale),
        rp=scale_unit_value(float(point[2]), bounds.rp, "linear"),
        design=design,
    )


def normalized_epr(
    point: EPRPoint | Sequence[float],
    bounds: EPRBounds,
    ec_scale: str,
    pr_scale: str,
) -> np.ndarray:
    if isinstance(point, EPRPoint):
        Ec, Pr, rp = point.as_tuple()
    else:
        Ec, Pr, rp = (float(value) for value in point)
    return np.asarray(
        [
            normalize_value(Ec, bounds.ec, ec_scale),
            normalize_value(Pr, bounds.pr, pr_scale),
            normalize_value(rp, bounds.rp, "linear"),
        ],
        dtype=float,
    )


def minimum_squared_distances(points: np.ndarray, references: np.ndarray) -> np.ndarray:
    """Nearest-reference squared distance, evaluated in memory-safe chunks."""
    if len(references) == 0:
        return np.full(len(points), np.inf, dtype=float)

    result = np.full(len(points), np.inf, dtype=float)
    point_chunk_size = 2048
    reference_chunk_size = 128

    for point_start in range(0, len(points), point_chunk_size):
        point_stop = min(point_start + point_chunk_size, len(points))
        point_chunk = points[point_start:point_stop]
        local_minimum = np.full(len(point_chunk), np.inf, dtype=float)

        for reference_start in range(0, len(references), reference_chunk_size):
            reference_stop = min(reference_start + reference_chunk_size, len(references))
            differences = (
                point_chunk[:, np.newaxis, :]
                - references[np.newaxis, reference_start:reference_stop, :]
            )
            squared = np.einsum("ijk,ijk->ij", differences, differences)
            local_minimum = np.minimum(local_minimum, np.min(squared, axis=1))

        result[point_start:point_stop] = local_minimum

    return result


def variation_bounds(
    global_bounds: EPRBounds,
    spec: VariationSpec,
) -> EPRBounds:
    """Build a local box around ``spec.anchor``, clipped by global bounds."""
    spec.validate()
    global_bounds.validate_positive()
    Ec0, Pr0, rp0 = spec.anchor

    raw_ec = (Ec0 * (1.0 - spec.ec_fraction), Ec0 * (1.0 + spec.ec_fraction))
    raw_pr = (Pr0 * (1.0 - spec.pr_fraction), Pr0 * (1.0 + spec.pr_fraction))
    raw_rp = (rp0 - spec.rp_delta, rp0 + spec.rp_delta)

    def intersect(
        local: tuple[float, float],
        global_: tuple[float, float],
        name: str,
    ) -> tuple[float, float]:
        low = max(local[0], global_[0])
        high = min(local[1], global_[1])
        if high < low:
            raise ValueError(
                f"Local {name} variation {local} does not overlap global bounds {global_}"
            )
        return (low, high)

    return EPRBounds(
        ec=intersect(raw_ec, global_bounds.ec, "Ec"),
        pr=intersect(raw_pr, global_bounds.pr, "Pr"),
        rp=intersect(raw_rp, global_bounds.rp, "rp"),
    )


def oat_variation_points(
    global_bounds: EPRBounds,
    spec: VariationSpec,
) -> list[EPRPoint]:
    """Generate deterministic one-at-a-time perturbations around the anchor."""
    spec.validate()
    if spec.mode != "oat":
        raise ValueError("oat_variation_points requires spec.mode='oat'")
    global_bounds.validate_positive()

    Ec0, Pr0, rp0 = spec.anchor
    anchor_values = {"ec": Ec0, "pr": Pr0, "rp": rp0}
    results: list[EPRPoint] = []

    if spec.include_center:
        center = EPRPoint(Ec0, Pr0, rp0, "variation-oat-center")
        if not global_bounds.contains(center):
            raise ValueError(
                f"Variation anchor {center.as_tuple()} lies outside global bounds {global_bounds}"
            )
        results.append(center)

    for level in spec.levels:
        for parameter in spec.parameters:
            for sign in (-1.0, +1.0):
                values = dict(anchor_values)
                if parameter == "ec":
                    values[parameter] = Ec0 * (1.0 + sign * level * spec.ec_fraction)
                elif parameter == "pr":
                    values[parameter] = Pr0 * (1.0 + sign * level * spec.pr_fraction)
                else:
                    values[parameter] = rp0 + sign * level * spec.rp_delta

                point = EPRPoint(
                    Ec=values["ec"],
                    Pr=values["pr"],
                    rp=values["rp"],
                    design=f"variation-oat-{parameter}-{'plus' if sign > 0 else 'minus'}-L{level:g}",
                )
                if not global_bounds.contains(point, atol=1.0e-15):
                    raise ValueError(
                        "OAT variation point falls outside the global sampling bounds: "
                        f"{point.as_tuple()} from parameter={parameter}, level={level:g}. "
                        "Reduce the variation span or widen the corresponding --*-range."
                    )
                results.append(point)

    # Preserve deterministic order while removing exact duplicates, e.g. zero spans.
    unique: list[EPRPoint] = []
    seen: set[tuple[float, float, float]] = set()
    for point in results:
        key = point.as_tuple()
        if key in seen:
            continue
        seen.add(key)
        unique.append(point)
    return unique


def screening_unit_candidates() -> np.ndarray:
    """Return deterministic unit-cube candidates for coarse range screening.

    Priority of the first 27 points:

    1. box center (1 point)
    2. all low/high corners (8 points)
    3. face centers (6 points)
    4. edge centers (12 points)

    The remaining candidates come from a 5-level lattice with normalized levels
    ``0, 0.25, 0.5, 0.75, 1``.  Exact duplicates are removed while preserving
    priority.  This makes ``count=15`` a useful default for fast FerroX range
    checks, while larger counts progressively fill the interior.
    """
    candidates: list[tuple[float, float, float]] = []

    # 1) Center.
    candidates.append((0.5, 0.5, 0.5))

    # 2) 8 corners.
    for x in (0.0, 1.0):
        for y in (0.0, 1.0):
            for z in (0.0, 1.0):
                candidates.append((x, y, z))

    # 3) 6 face centers: one coordinate at a boundary, the other two centered.
    for axis in range(3):
        for boundary in (0.0, 1.0):
            point = [0.5, 0.5, 0.5]
            point[axis] = boundary
            candidates.append(tuple(point))

    # 4) 12 edge centers: two coordinates on boundaries, one centered.
    for center_axis in range(3):
        boundary_axes = [axis for axis in range(3) if axis != center_axis]
        for a in (0.0, 1.0):
            for b in (0.0, 1.0):
                point = [0.5, 0.5, 0.5]
                point[boundary_axes[0]] = a
                point[boundary_axes[1]] = b
                candidates.append(tuple(point))

    # 5) Finer 5-level lattice for count > 27 or to replace rejected core points.
    fine_levels = (0.0, 0.25, 0.5, 0.75, 1.0)
    for x in fine_levels:
        for y in fine_levels:
            for z in fine_levels:
                candidates.append((x, y, z))

    unique: list[tuple[float, float, float]] = []
    seen: set[tuple[float, float, float]] = set()
    for point in candidates:
        if point in seen:
            continue
        seen.add(point)
        unique.append(point)

    return np.asarray(unique, dtype=float)


def screening_epr_points(
    *,
    count: int,
    bounds: EPRBounds,
    ec_scale: str,
    pr_scale: str,
    reference_points: Iterable[Sequence[float]] = (),
    accept: Callable[[EPRPoint], bool] | None = None,
) -> list[EPRPoint]:
    """Generate deterministic coarse screening points for FerroX.

    ``count=15`` gives the recommended core design: center + corners + face
    centers, subject to ``accept``.  If a core point is rejected by a physical
    constraint, or if more than 15 points are requested, additional accepted
    candidates are selected from the 5-level lattice by maximin distance.

    Notes
    -----
    The screening design intentionally includes exact lower/upper bounds.  The
    caller therefore should not use mathematically singular bounds such as
    ``Pr=0`` when the downstream EPR-to-Landau transform requires ``Pr > 0``.
    """
    if count <= 0:
        raise ValueError("screening requires a positive count")

    unit_candidates = screening_unit_candidates()
    physical_candidates: list[EPRPoint] = []
    normalized_candidates: list[np.ndarray] = []

    for unit_point in unit_candidates:
        point = unit_point_to_epr(
            unit_point,
            bounds,
            ec_scale,
            pr_scale,
            design="screening",
        )
        if accept is not None and not accept(point):
            continue
        physical_candidates.append(point)
        normalized_candidates.append(np.asarray(unit_point, dtype=float))

    if len(physical_candidates) < count:
        raise RuntimeError(
            f"Only {len(physical_candidates)} accepted screening points are available "
            f"for count={count}. Reduce --count or relax the acceptance constraints."
        )

    # Keep the accepted high-priority core in its deterministic order.  The first
    # 15 raw candidates are center + 8 corners + 6 face centers.
    core_raw = screening_unit_candidates()[:15]
    core_keys = {tuple(float(v) for v in row) for row in core_raw}

    selected: list[EPRPoint] = []
    selected_norm: list[np.ndarray] = []
    used_indices: set[int] = set()

    for index, (point, norm) in enumerate(zip(physical_candidates, normalized_candidates, strict=True)):
        if tuple(float(v) for v in norm) not in core_keys:
            continue
        selected.append(point)
        selected_norm.append(norm)
        used_indices.add(index)
        if len(selected) >= count:
            return selected

    # Normalize supplied references and use them only to guide the augmentation.
    references_norm: list[np.ndarray] = []
    for ref in reference_points:
        ref_tuple = tuple(float(v) for v in ref)
        if bounds.contains(ref_tuple, atol=1.0e-12):
            references_norm.append(normalized_epr(ref_tuple, bounds, ec_scale, pr_scale))

    if selected_norm:
        references_norm.extend(selected_norm)

    all_norm = np.asarray(normalized_candidates, dtype=float)
    if references_norm:
        reference_array = np.asarray(references_norm, dtype=float).reshape(-1, 3)
        nearest_distance = minimum_squared_distances(all_norm, reference_array)
    else:
        # This can occur only if all 15 core points were rejected.  Favor points
        # farthest from the unit-cube center for the first replacement.
        center = np.asarray([[0.5, 0.5, 0.5]], dtype=float)
        nearest_distance = np.sum((all_norm - center) ** 2, axis=1)

    available = np.ones(len(physical_candidates), dtype=bool)
    for index in used_indices:
        available[index] = False

    while len(selected) < count and np.any(available):
        scores = np.where(available, nearest_distance, -np.inf)
        selected_index = int(np.argmax(scores))
        available[selected_index] = False
        selected.append(physical_candidates[selected_index])

        selected_point = all_norm[selected_index]
        squared_to_selected = np.sum((all_norm - selected_point) ** 2, axis=1)
        nearest_distance = np.minimum(nearest_distance, squared_to_selected)

    if len(selected) != count:
        raise RuntimeError(
            f"Screening selected only {len(selected)} points out of requested {count}"
        )
    return selected



def load_epr_points_from_thickness(
    *,
    dataset_path: str | Path = "MFM_dataset.xlsx",
    source_t_fe: float = 8e-9,
    sheet_name: str = "experiments",
    design: str = "thickness_copy",
    count: int,
    accept: Callable[[EPRPoint], bool] | None = None,
) -> list[EPRPoint]:
    """
    從 dataset 中讀取指定厚度的 Ec, Pr, rp，
    經 accept 篩選後最多回傳 count 個 EPRPoint。

    若符合條件的點數少於 count，
    則回傳所有剩餘點。
    """

    if count <= 0:
        raise ValueError("count must be positive")

    experiments = pd.read_excel(
        dataset_path,
        sheet_name=sheet_name,
    )

    mask = np.isclose(
        experiments["T_FE"].astype(float),
        source_t_fe,
        rtol=1e-6,
        atol=1e-15,
    )

    selected = (
        experiments.loc[
            mask,
            ["Ec", "Pr", "rp"],
        ]
        .dropna()
        .drop_duplicates()
    )

    points: list[EPRPoint] = []
    d_cnt = 0

    for row in selected.itertuples(index=False):
        point = EPRPoint(
            Ec=float(row.Ec),
            Pr=float(row.Pr),
            rp=float(row.rp),
            design=design,
        )

        if accept is not None and not accept(point):
            d_cnt += 1
            continue

        points.append(point)

        # 已經取得需要的數量，可以直接停止
        if len(points) >= count:
            break

    if d_cnt > 0:
        print(f"Skipped {d_cnt} points that did not meet the acceptance criteria.")

    return points

def _unit_design(
    design: str,
    n_samples: int,
    dimensions: int,
    *,
    seed: int,
    rng: np.random.Generator,
) -> np.ndarray:
    if design == "uniform":
        return rng.random((n_samples, dimensions))
    if design == "lhs":
        return lhs_unit(n_samples, dimensions, rng)
    if design == "sobol":
        return sobol_unit(n_samples, dimensions, seed)
    raise ValueError(f"Unit design does not support {design!r}")


def sample_epr_points(
    *,
    count: int | None,
    design: str,
    bounds: EPRBounds,
    ec_scale: str = "linear",
    pr_scale: str = "linear",
    seed: int = 0,
    reference_points: Iterable[Sequence[float]] = (),
    maximin_pool_size: int | None = None,
    variation: VariationSpec | None = None,
    accept: Callable[[EPRPoint], bool] | None = None,
    source_t_fe: float = 8e-9,
    excel_path: str | Path = "MFM_dataset.xlsx",
    
) -> list[EPRPoint]:
    """Generate EPR points for either a global or local/variation design.

    ``accept`` is an optional domain-specific predicate.  It is useful when the
    caller needs to reject points after converting them to another parameter
    system (for example a ``|alpha|`` constraint) while keeping this module
    independent of that physics.
    """
    bounds.validate_positive()
    if ec_scale not in {"linear", "log"}:
        raise ValueError("ec_scale must be 'linear' or 'log'")
    if pr_scale not in {"linear", "log"}:
        raise ValueError("pr_scale must be 'linear' or 'log'")

    if design == "variation":
        if variation is None:
            raise ValueError("design='variation' requires a VariationSpec")
        variation.validate()
        if variation.mode == "oat":
            proposed = oat_variation_points(bounds, variation)
            accepted = [point for point in proposed if accept is None or accept(point)]
            if count is not None and count != len(accepted):
                raise ValueError(
                    "For variation/oat, --count is optional. If supplied it must equal "
                    f"the number of accepted deterministic points ({len(accepted)}), "
                    f"but got {count}. Usually omit --count for OAT."
                )
            if not accepted:
                raise RuntimeError("No OAT variation points survived the acceptance rules")
            return accepted
        active_design = variation.mode
        active_bounds = variation_bounds(bounds, variation)
        label_prefix = "variation-"
    else:
        if design not in GLOBAL_DESIGNS:
            raise ValueError(
                f"Unknown design {design!r}; choose from {GLOBAL_DESIGNS} or 'variation'"
            )
        active_design = design
        active_bounds = bounds
        label_prefix = ""

    if count is None or count <= 0:
        raise ValueError(f"design={design!r} requires a positive count")

    rng = np.random.default_rng(seed)
    references = [tuple(float(v) for v in ref) for ref in reference_points]

    if active_design == "screening":
        if design == "variation":
            raise ValueError("screening is a global design and cannot be used as a variation mode")
        return screening_epr_points(
            count=count,
            bounds=active_bounds,
            ec_scale=ec_scale,
            pr_scale=pr_scale,
            reference_points=references,
            accept=accept,
        )
        
    if active_design == "f_thickness":
        if design == "variation":
            raise ValueError("f_thickness is a global design and cannot be used as a variation mode")
        return load_epr_points_from_thickness(
            dataset_path=excel_path,
            source_t_fe=source_t_fe,
            sheet_name="experiments",
            design="f_thickness",
            count=count,
            accept=accept,
        )

    if active_design in {"uniform", "lhs", "sobol"}:
        accepted_points: list[EPRPoint] = []
        seen: set[tuple[float, float, float]] = set()
        attempts = 0
        max_attempts = max(10_000, count * 2_000)
        batch_number = 0

        while len(accepted_points) < count and attempts < max_attempts:
            remaining = count - len(accepted_points)
            batch_size = max(remaining * 4, 32)
            batch_seed = seed + batch_number
            unit_points = _unit_design(
                active_design,
                batch_size,
                3,
                seed=batch_seed,
                rng=rng,
            )
            batch_number += 1

            for unit_point in unit_points:
                attempts += 1
                point = unit_point_to_epr(
                    unit_point,
                    active_bounds,
                    ec_scale,
                    pr_scale,
                    design=f"{label_prefix}{active_design}",
                )
                key = point.as_tuple()
                if key in seen:
                    continue
                if accept is not None and not accept(point):
                    if attempts >= max_attempts:
                        break
                    continue
                seen.add(key)
                accepted_points.append(point)
                if len(accepted_points) >= count or attempts >= max_attempts:
                    break

        if len(accepted_points) != count:
            raise RuntimeError(
                f"Could generate only {len(accepted_points)} accepted points out of {count}. "
                "Change the seed, relax constraints, or expand the sampling bounds."
            )
        return accepted_points

    # Maximin augmentation: build a large accepted pool first, then greedily
    # maximize nearest-neighbour distance in normalized EPR coordinates.
    pool_size = max(4096, count * 50) if maximin_pool_size is None else maximin_pool_size
    if pool_size < count:
        raise ValueError(
            f"maximin_pool_size must be at least count; got {pool_size} < {count}"
        )

    # Oversample because caller-side physical constraints may reject many points.
    raw_pool_size = max(pool_size, count)
    unit_pool = lhs_unit(raw_pool_size, 3, rng)
    pool: list[EPRPoint] = []
    pool_norm: list[np.ndarray] = []
    seen_pool: set[tuple[float, float, float]] = set()

    for unit_point in unit_pool:
        point = unit_point_to_epr(
            unit_point,
            active_bounds,
            ec_scale,
            pr_scale,
            design=f"{label_prefix}maximin",
        )
        key = point.as_tuple()
        if key in seen_pool:
            continue
        if accept is not None and not accept(point):
            continue
        seen_pool.add(key)
        pool.append(point)
        pool_norm.append(normalized_epr(point, active_bounds, ec_scale, pr_scale))

    if len(pool) < count:
        raise RuntimeError(
            f"Only {len(pool)} accepted maximin pool points were generated for count={count}. "
            "Increase --maximin-pool-size, relax constraints, or expand the bounds."
        )

    reference_norm: list[np.ndarray] = []
    for ref in references:
        if not active_bounds.contains(ref, atol=1.0e-12):
            continue
        reference_norm.append(normalized_epr(ref, active_bounds, ec_scale, pr_scale))

    pool_array = np.asarray(pool_norm, dtype=float)
    reference_array = np.asarray(reference_norm, dtype=float).reshape(-1, 3)
    nearest_distance = minimum_squared_distances(pool_array, reference_array)
    available = np.ones(len(pool), dtype=bool)
    selected: list[EPRPoint] = []

    while len(selected) < count and np.any(available):
        scores = np.where(available, nearest_distance, -np.inf)
        selected_index = int(np.argmax(scores))
        available[selected_index] = False
        selected.append(pool[selected_index])

        selected_point = pool_array[selected_index]
        squared_to_selected = np.sum((pool_array - selected_point) ** 2, axis=1)
        nearest_distance = np.minimum(nearest_distance, squared_to_selected)

    if len(selected) != count:
        raise RuntimeError(f"Maximin selected only {len(selected)} points out of requested {count}")
    return selected
