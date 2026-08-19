"""Bidirectional sixth-order Landau <-> intrinsic EPR conversion.

The model is

    E(P) = alpha * P + beta * P**3 + gamma * P**5

and this module works on the branch ``alpha < 0`` and ``gamma > 0``. On this
branch, the sign of beta determines the transition class and the valid rp range:

* beta < 0: first-order,  sqrt(5/3) < rp < 5**(1/4)
* beta = 0: tricritical,  rp = 5**(1/4)
* beta > 0: second-order, 5**(1/4) < rp < sqrt(3)

Here ``Ec`` is the positive magnitude of the intrinsic coercive field, ``Pr``
is the positive zero-field polarization, ``Pc0`` is the positive polarization
at the negative-field spinodal, and ``rp = Pr / Pc0``.
"""

from __future__ import annotations

import math
import sys
from dataclasses import dataclass
from typing import Iterable


FIRST_ORDER_RP_BOUNDS = (math.sqrt(5.0 / 3.0), 5.0**0.25)
SECOND_ORDER_RP_BOUNDS = (5.0**0.25, math.sqrt(3.0))
TRICRITICAL_RP = 5.0**0.25


@dataclass(frozen=True, slots=True)
class LandauParameters:
    alpha: float
    beta: float
    gamma: float


@dataclass(frozen=True, slots=True)
class EPRParameters:
    Ec: float
    Pr: float
    Pc0: float
    rp: float
    transition_order: str
    rp_lower_bound: float
    rp_upper_bound: float


@dataclass(frozen=True, slots=True)
class TransitionInfo:
    order: str
    rp_lower_bound: float
    rp_upper_bound: float


@dataclass(frozen=True, slots=True)
class ConsistencyResult:
    passed: bool
    transition_order: str
    rp_bounds_passed: bool
    coefficient_round_trip_passed: bool
    remanence_condition_passed: bool
    spinodal_condition_passed: bool
    coercive_field_condition_passed: bool
    polarization_relation_passed: bool
    max_coefficient_relative_error: float
    normalized_E_at_Pr: float
    normalized_dE_dP_at_Pc0: float
    normalized_coercive_field_error: float


def _finite(name: str, value: float) -> float:
    value = float(value)
    if not math.isfinite(value):
        raise ValueError(f"{name} must be finite; got {value!r}")
    return value


def _relative_error(actual: float, expected: float) -> float:
    scale = max(abs(actual), abs(expected), sys.float_info.min)
    return abs(actual - expected) / scale


def _normalized_residual(residual: float, terms: Iterable[float]) -> float:
    scale = max(sum(abs(term) for term in terms), sys.float_info.min)
    return abs(residual) / scale


def get_transition_info(beta: float) -> TransitionInfo:
    """Classify transition order from beta and return the corresponding rp bounds."""

    beta = _finite("beta", beta)
    if beta < 0.0:
        return TransitionInfo("first_order", *FIRST_ORDER_RP_BOUNDS)
    if beta > 0.0:
        return TransitionInfo("second_order", *SECOND_ORDER_RP_BOUNDS)
    return TransitionInfo("tricritical", TRICRITICAL_RP, TRICRITICAL_RP)


def _rp_is_in_expected_range(
    rp: float,
    transition: TransitionInfo,
    *,
    tolerance: float,
) -> bool:
    if transition.order == "tricritical":
        return math.isclose(rp, TRICRITICAL_RP, rel_tol=tolerance, abs_tol=tolerance)
    return (
        transition.rp_lower_bound - tolerance
        <= rp
        <= transition.rp_upper_bound + tolerance
    )


def landau_to_EPR(
    alpha: float,
    beta: float,
    gamma: float,
    *,
    validate_rp_bounds: bool = True,
) -> EPRParameters:
    """Convert ``(alpha, beta, gamma)`` to ``(Ec, Pr, Pc0, rp)``.

    The square-root formulas are selected according to the sign of beta to
    avoid cancellation near the first/second-order limits.
    """

    alpha = _finite("alpha", alpha)
    beta = _finite("beta", beta)
    gamma = _finite("gamma", gamma)

    if alpha >= 0.0:
        raise ValueError(
            "This rp parameterization and its requested bounds require alpha < 0"
        )
    if gamma <= 0.0:
        raise ValueError("A stable sixth-order model requires gamma > 0")

    transition = get_transition_info(beta)
    cross_scale = math.sqrt(-alpha) * math.sqrt(gamma)
    delta_r = math.hypot(beta, 2.0 * cross_scale)
    delta_c = math.hypot(3.0 * beta, math.sqrt(20.0) * cross_scale)

    if beta < 0.0:
        # These forms add positive terms when beta < 0.
        pr_sq = (-beta + delta_r) / (2.0 * gamma)
        pc0_sq = (-3.0 * beta + delta_c) / (10.0 * gamma)
    else:
        # Rationalized forms avoid subtracting nearly equal terms.
        pr_sq = (-2.0 * alpha) / (beta + delta_r)
        pc0_sq = (-2.0 * alpha) / (3.0 * beta + delta_c)

    if not (pr_sq > 0.0 and pc0_sq > 0.0):
        raise ArithmeticError(f"Expected positive Pr^2 and Pc0^2; got {pr_sq=}, {pc0_sq=}")

    Pr = math.sqrt(pr_sq)
    Pc0 = math.sqrt(pc0_sq)
    rp = Pr / Pc0
    Ec = 2.0 * Pc0**3 * (beta + 2.0 * gamma * pc0_sq)

    bound_tolerance = 64.0 * sys.float_info.epsilon
    if validate_rp_bounds and not _rp_is_in_expected_range(
        rp,
        transition,
        tolerance=bound_tolerance,
    ):
        raise ArithmeticError(
            f"rp={rp:.12g} is outside the expected {transition.order} range "
            f"[{transition.rp_lower_bound:.12g}, {transition.rp_upper_bound:.12g}]"
        )

    return EPRParameters(
        Ec=Ec,
        Pr=Pr,
        Pc0=Pc0,
        rp=rp,
        transition_order=transition.order,
        rp_lower_bound=transition.rp_lower_bound,
        rp_upper_bound=transition.rp_upper_bound,
    )


def EPR_to_landau(
    Ec: float,
    Pr: float,
    *,
    Pc0: float | None = None,
    rp: float | None = None,
    relation_rtol: float = 1.0e-9,
    relation_atol: float = 0.0,
    validate_rp_bounds: bool = True,
) -> LandauParameters:
    """Convert intrinsic EPR quantities back to ``(alpha, beta, gamma)``.

    Supply at least one of ``Pc0`` and ``rp``. If both are supplied, the
    function checks ``Pr == rp * Pc0`` before conversion.
    """

    Ec = _finite("Ec", Ec)
    Pr = _finite("Pr", Pr)
    if Ec <= 0.0 or Pr <= 0.0:
        raise ValueError("Ec and Pr must be positive")
    if Pc0 is None and rp is None:
        raise ValueError("Supply Pc0, rp, or both")

    if Pc0 is not None:
        Pc0 = _finite("Pc0", Pc0)
        if Pc0 <= 0.0:
            raise ValueError("Pc0 must be positive")
    if rp is not None:
        rp = _finite("rp", rp)
        if rp <= 1.0:
            raise ValueError("rp must be greater than 1")

    if Pc0 is None:
        assert rp is not None
        Pc0 = Pr / rp
    elif rp is None:
        rp = Pr / Pc0
    elif not math.isclose(
        Pr,
        rp * Pc0,
        rel_tol=relation_rtol,
        abs_tol=relation_atol,
    ):
        raise ValueError(
            "Inconsistent inputs: expected Pr == rp * Pc0; "
            f"got Pr={Pr:.12g}, rp*Pc0={rp * Pc0:.12g}"
        )

    r_sq = rp * rp
    denominator = 2.0 * (r_sq - 1.0) ** 2
    alpha = -Ec / Pc0 * (r_sq * (3.0 * r_sq - 5.0) / denominator)
    beta = Ec / Pc0**3 * ((r_sq**2 - 5.0) / denominator)
    gamma = -Ec / Pc0**5 * ((r_sq - 3.0) / denominator)

    result = LandauParameters(alpha=alpha, beta=beta, gamma=gamma)
    if validate_rp_bounds:
        if alpha >= 0.0 or gamma <= 0.0:
            raise ValueError(
                "The supplied rp is outside the alpha < 0, gamma > 0 branch"
            )
        transition = get_transition_info(beta)
        tolerance = max(10.0 * relation_rtol, 64.0 * sys.float_info.epsilon)
        if not _rp_is_in_expected_range(rp, transition, tolerance=tolerance):
            raise ValueError(
                f"rp={rp:.12g} is inconsistent with beta={beta:.12g} "
                f"({transition.order})"
            )

    return result


def check_landau_EPR(
    alpha: float,
    beta: float,
    gamma: float,
    Ec: float,
    Pr: float,
    Pc0: float,
    rp: float,
    *,
    rtol: float = 1.0e-9,
    atol: float = 0.0,
) -> ConsistencyResult:
    """Check round-trip coefficients, Landau conditions and beta-dependent rp bounds."""

    if rtol <= 0.0 or atol < 0.0:
        raise ValueError("rtol must be positive and atol must be non-negative")

    alpha = _finite("alpha", alpha)
    beta = _finite("beta", beta)
    gamma = _finite("gamma", gamma)
    Ec = _finite("Ec", Ec)
    Pr = _finite("Pr", Pr)
    Pc0 = _finite("Pc0", Pc0)
    rp = _finite("rp", rp)

    if alpha >= 0.0 or gamma <= 0.0:
        raise ValueError("Consistency check requires alpha < 0 and gamma > 0")
    if min(Ec, Pr, Pc0, rp) <= 0.0:
        raise ValueError("Ec, Pr, Pc0 and rp must be positive")

    transition = get_transition_info(beta)
    rebuilt = EPR_to_landau(
        Ec,
        Pr,
        Pc0=Pc0,
        rp=rp,
        relation_rtol=rtol,
        relation_atol=atol,
        validate_rp_bounds=False,
    )

    originals = (alpha, beta, gamma)
    reconstructions = (rebuilt.alpha, rebuilt.beta, rebuilt.gamma)
    coefficient_errors = tuple(
        _relative_error(actual, expected)
        for actual, expected in zip(reconstructions, originals)
    )

    # Near an rp boundary, recovering the coefficient that approaches zero is
    # ill-conditioned. Relax only that round-trip comparison by the predicted
    # floating-point amplification; all defining-equation checks retain rtol.
    r_sq = rp * rp
    epsilon = sys.float_info.epsilon
    conditions = (
        max(1.0, abs(3.0 * r_sq) / max(abs(3.0 * r_sq - 5.0), epsilon)),
        max(1.0, abs(r_sq**2) / max(abs(r_sq**2 - 5.0), epsilon)),
        max(1.0, abs(r_sq) / max(abs(r_sq - 3.0), epsilon)),
    )
    coefficient_tolerances = tuple(
        max(rtol, 32.0 * epsilon * condition) for condition in conditions
    )
    coefficient_round_trip_passed = all(
        abs(actual - expected) <= atol or error <= tolerance
        for actual, expected, error, tolerance in zip(
            reconstructions,
            originals,
            coefficient_errors,
            coefficient_tolerances,
        )
    )

    e_pr_terms = (alpha * Pr, beta * Pr**3, gamma * Pr**5)
    normalized_E_at_Pr = _normalized_residual(sum(e_pr_terms), e_pr_terms)

    slope_terms = (alpha, 3.0 * beta * Pc0**2, 5.0 * gamma * Pc0**4)
    normalized_dE_dP_at_Pc0 = _normalized_residual(sum(slope_terms), slope_terms)

    E_at_Pc0 = alpha * Pc0 + beta * Pc0**3 + gamma * Pc0**5
    normalized_coercive_field_error = _normalized_residual(
        E_at_Pc0 + Ec,
        (E_at_Pc0, Ec),
    )

    polarization_relation_passed = math.isclose(
        Pr,
        rp * Pc0,
        rel_tol=rtol,
        abs_tol=atol,
    )
    rp_bounds_passed = _rp_is_in_expected_range(
        rp,
        transition,
        tolerance=max(10.0 * rtol, 64.0 * epsilon),
    )
    remanence_condition_passed = normalized_E_at_Pr <= rtol
    spinodal_condition_passed = normalized_dE_dP_at_Pc0 <= rtol
    coercive_field_condition_passed = normalized_coercive_field_error <= rtol

    checks = (
        rp_bounds_passed,
        coefficient_round_trip_passed,
        remanence_condition_passed,
        spinodal_condition_passed,
        coercive_field_condition_passed,
        polarization_relation_passed,
    )
    return ConsistencyResult(
        passed=all(checks),
        transition_order=transition.order,
        rp_bounds_passed=rp_bounds_passed,
        coefficient_round_trip_passed=coefficient_round_trip_passed,
        remanence_condition_passed=remanence_condition_passed,
        spinodal_condition_passed=spinodal_condition_passed,
        coercive_field_condition_passed=coercive_field_condition_passed,
        polarization_relation_passed=polarization_relation_passed,
        max_coefficient_relative_error=max(coefficient_errors),
        normalized_E_at_Pr=normalized_E_at_Pr,
        normalized_dE_dP_at_Pc0=normalized_dE_dP_at_Pc0,
        normalized_coercive_field_error=normalized_coercive_field_error,
    )


__all__ = [
    "ConsistencyResult",
    "EPRParameters",
    "EPR_to_landau",
    "FIRST_ORDER_RP_BOUNDS",
    "LandauParameters",
    "SECOND_ORDER_RP_BOUNDS",
    "TRICRITICAL_RP",
    "TransitionInfo",
    "check_landau_EPR",
    "get_transition_info",
    "landau_to_EPR",
]


# =============================================================================
# Command-line interface (safe for imports)
# =============================================================================
# This block runs only when this file is executed directly. Importing any of
# the functions above from another Python file will not execute this block.
if __name__ == "__main__":
    import argparse
    import re

    parser = argparse.ArgumentParser(
        description=(
            "Convert sixth-order Landau parameters and intrinsic EPR "
            "parameters, then run the consistency test."
        ),
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  Landau -> EPR:
    python landau_EPR_transformer.py --mode landau_to_EPR --alpha -7.4e9 --beta 1.4e12 --gamma 5.0e12

  EPR -> Landau, using rp:
    python landau_EPR_transformer.py --mode EPR_to_landau --Ec 2.064e8 --Pr 7.204e-2 --rp 1.725

  EPR -> Landau, using Pc0:
    python landau_EPR_transformer.py --mode EPR_to_landau --Ec 2.064e8 --Pr 7.204e-2 --Pc0 4.176e-2

  EPR -> Landau, supplying both Pc0 and rp for relation checking:
    python landau_EPR_transformer.py --mode EPR_to_landau --Ec 2.064e8 --Pr 7.204e-2 --Pc0 4.176231884057971e-2 --rp 1.725
""",
    )
    # argparse's default negative-number matcher does not recognize scientific
    # notation such as -7.4e9 when it follows an option as a separate token.
    parser._negative_number_matcher = re.compile(
        r"^-?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?$"
    )
    parser.add_argument(
        "--mode",
        required=True,
        choices=("landau_to_EPR", "EPR_to_landau"),
        help="Select the conversion direction.",
    )

    landau_group = parser.add_argument_group("Landau input")
    landau_group.add_argument(
        "--alpha",
        type=float,
        help="Landau alpha; required for landau_to_EPR.",
    )
    landau_group.add_argument(
        "--beta",
        type=float,
        help="Landau beta; required for landau_to_EPR.",
    )
    landau_group.add_argument(
        "--gamma",
        type=float,
        help="Landau gamma; required for landau_to_EPR.",
    )

    epr_group = parser.add_argument_group("EPR input")
    epr_group.add_argument(
        "--Ec",
        "--ec",
        dest="Ec",
        type=float,
        help="Positive coercive-field magnitude; required for EPR_to_landau.",
    )
    epr_group.add_argument(
        "--Pr",
        "--pr",
        dest="Pr",
        type=float,
        help="Positive remanent polarization; required for EPR_to_landau.",
    )
    epr_group.add_argument(
        "--Pc0",
        "--pc0",
        dest="Pc0",
        type=float,
        help="Positive spinodal polarization; supply Pc0, rp, or both.",
    )
    epr_group.add_argument(
        "--rp",
        type=float,
        help="Polarization ratio Pr/Pc0; supply Pc0, rp, or both.",
    )

    test_group = parser.add_argument_group("Consistency-test options")
    test_group.add_argument(
        "--rtol",
        type=float,
        default=1.0e-9,
        help="Relative tolerance (default: %(default).1e).",
    )
    test_group.add_argument(
        "--atol",
        type=float,
        default=0.0,
        help="Absolute tolerance (default: %(default)g).",
    )

    args = parser.parse_args()

    def _require_arguments(*names: str) -> None:
        missing = [f"--{name}" for name in names if getattr(args, name) is None]
        if missing:
            parser.error(
                f"mode {args.mode!r} requires: {', '.join(missing)}"
            )

    def _print_dataclass(title: str, result: object) -> None:
        print(f"\n{title}")
        print("-" * len(title))
        for field_name in result.__dataclass_fields__:
            value = getattr(result, field_name)
            if isinstance(value, float):
                print(f"{field_name}: {value:.12e}")
            else:
                print(f"{field_name}: {value}")

    try:
        if args.mode == "landau_to_EPR":
            _require_arguments("alpha", "beta", "gamma")
            input_parameters = {
                "alpha": args.alpha,
                "beta": args.beta,
                "gamma": args.gamma,
            }
            transformed = landau_to_EPR(**input_parameters)
            consistency = check_landau_EPR(
                args.alpha,
                args.beta,
                args.gamma,
                transformed.Ec,
                transformed.Pr,
                transformed.Pc0,
                transformed.rp,
                rtol=args.rtol,
                atol=args.atol,
            )
            result_title = "EPR result"

        else:
            _require_arguments("Ec", "Pr")
            if args.Pc0 is None and args.rp is None:
                parser.error(
                    "mode 'EPR_to_landau' requires --Pc0, --rp, or both"
                )

            input_parameters = {"Ec": args.Ec, "Pr": args.Pr}
            if args.Pc0 is not None:
                input_parameters["Pc0"] = args.Pc0
            if args.rp is not None:
                input_parameters["rp"] = args.rp

            transformed = EPR_to_landau(
                **input_parameters,
                relation_rtol=args.rtol,
                relation_atol=args.atol,
            )

            Ec = args.Ec
            Pr = args.Pr
            if args.Pc0 is None:
                rp = args.rp
                Pc0 = Pr / rp
            elif args.rp is None:
                Pc0 = args.Pc0
                rp = Pr / Pc0
            else:
                Pc0 = args.Pc0
                rp = args.rp

            consistency = check_landau_EPR(
                transformed.alpha,
                transformed.beta,
                transformed.gamma,
                Ec,
                Pr,
                Pc0,
                rp,
                rtol=args.rtol,
                atol=args.atol,
            )
            result_title = "Landau result"

    except (ArithmeticError, ValueError) as error:
        parser.error(str(error))

    print(f"Selected mode: {args.mode}")
    print(f"Input parameters: {input_parameters}")
    _print_dataclass(result_title, transformed)
    _print_dataclass("Consistency test", consistency)
    print(f"\nOverall consistency: {'PASS' if consistency.passed else 'FAIL'}")

    
