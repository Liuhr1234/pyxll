"""Splice distribution support for Drisk.

This implementation follows the splice construction used by the team's
reference prototype:

- Left component contributes probability mass on ``x <= splice_point``
- Right component contributes probability mass on ``x > splice_point``
- Scaling constants are chosen so the two component heights match at the
  splice point and the full distribution integrates/sums to 1.

Nested left/right distributions are passed in as Drisk formula strings, just
like Compound.
"""

from __future__ import annotations

import math
from typing import Any, Dict, Iterable, List, Optional, Tuple

import numpy as np
from scipy.integrate import quad
from scipy.optimize import minimize_scalar

from constants import get_distribution_info
from distribution_base import DistributionBase

_EPS = 1e-12
_TAIL_PROB = 1e-10
_CORE_CACHE: Dict[Tuple[str, str, float], Dict[str, Any]] = {}


def _normalize_formula_text(formula_text: Any) -> str:
    text = str(formula_text).strip()
    if not text:
        raise ValueError("Splice requires non-empty distribution formulas")
    if not text.startswith("="):
        text = "=" + text
    return text


def _normalize_splice_point(value: Any) -> float:
    splice_point = float(value)
    if not math.isfinite(splice_point):
        raise ValueError("Splice point must be finite")
    return splice_point


def _normalize_params(
    left_formula: Any,
    right_formula: Any,
    splice_point: Any,
) -> Tuple[str, str, float]:
    return (
        _normalize_formula_text(left_formula),
        _normalize_formula_text(right_formula),
        _normalize_splice_point(splice_point),
    )


def _parse_distribution(formula_text: str):
    from statistical_functions_theo import _parse_distribution_from_formula_string

    dist = _parse_distribution_from_formula_string(formula_text)
    if dist is None:
        raise ValueError(f"Unable to parse nested distribution: {formula_text}")
    return dist


def _is_discrete_distribution(dist: Any) -> bool:
    func_name = getattr(dist, "func_name", None)
    if func_name:
        info = get_distribution_info(str(func_name))
        if info is not None:
            return bool(info.get("is_discrete", False))

    class_name = type(dist).__name__.lower()
    discrete_tokens = (
        "bernoulli",
        "binomial",
        "poisson",
        "negbin",
        "geomet",
        "hypergeo",
        "intuniform",
        "duniform",
        "discrete",
    )
    return any(token in class_name for token in discrete_tokens)


def _safe_min_value(dist: Any) -> float:
    try:
        value = float(dist.min_val())
        return value
    except Exception:
        return float("-inf")


def _safe_max_value(dist: Any) -> float:
    try:
        value = float(dist.max_val())
        return value
    except Exception:
        return float("inf")


def _safe_ppf_scalar(dist: Any, q: float) -> float:
    q = min(max(float(q), _EPS), 1.0 - _EPS)
    value = float(dist.ppf(q))
    if not math.isfinite(value):
        raise ValueError(f"Nested distribution returned non-finite ppf for q={q}")
    return value


def _finite_bound_from_tail(dist: Any, upper: bool) -> float:
    if upper:
        finite = _safe_max_value(dist)
        if math.isfinite(finite):
            return finite
        for q in (1.0 - _TAIL_PROB, 1.0 - 1e-8, 1.0 - 1e-6, 1.0 - 1e-4):
            value = _safe_ppf_scalar(dist, q)
            if math.isfinite(value):
                return value
        raise ValueError("Unable to determine finite upper bound for splice component")

    finite = _safe_min_value(dist)
    if math.isfinite(finite):
        return finite
    for q in (_TAIL_PROB, 1e-8, 1e-6, 1e-4):
        value = _safe_ppf_scalar(dist, q)
        if math.isfinite(value):
            return value
    raise ValueError("Unable to determine finite lower bound for splice component")


def _mass_or_density(dist: Any, is_discrete: bool, x: float) -> float:
    if is_discrete:
        if hasattr(dist, "pmf"):
            return float(dist.pmf(x))
        return float(dist.pdf(x))
    return float(dist.pdf(x))


def _explicit_support_points(dist: Any) -> Optional[np.ndarray]:
    for attr_name in ("x_vals_list", "x_vals"):
        values = getattr(dist, attr_name, None)
        if values is None:
            continue
        try:
            shift = float(getattr(dist, "shift_amount", 0.0))
            pts = sorted({float(x) + shift for x in values})
            if pts:
                return np.asarray(pts, dtype=float)
        except Exception:
            continue
    return None


def _discrete_support_points(dist: Any, upper_hint: Optional[float] = None) -> np.ndarray:
    explicit = _explicit_support_points(dist)
    if explicit is not None:
        return explicit

    start = _finite_bound_from_tail(dist, upper=False)
    end = upper_hint if upper_hint is not None else _finite_bound_from_tail(dist, upper=True)
    if not math.isfinite(end):
        end = _finite_bound_from_tail(dist, upper=True)

    points = []
    current = start
    max_steps = 100000
    steps = 0
    while current <= end + 1e-9 and steps < max_steps:
        points.append(float(current))
        current += 1.0
        steps += 1
    return np.asarray(points, dtype=float)


def _fast_vectorized_ppf(dist: Any, q_values: np.ndarray) -> Optional[np.ndarray]:
    scipy_dist = getattr(dist, "_dist", None)
    if scipy_dist is None or not hasattr(scipy_dist, "ppf") or not hasattr(scipy_dist, "cdf"):
        return None

    q = np.asarray(q_values, dtype=float)
    shift_amount = float(getattr(dist, "shift_amount", 0.0))

    if not dist.is_truncated():
        try:
            return np.asarray(scipy_dist.ppf(q), dtype=float) + shift_amount
        except Exception:
            return None

    lower, upper = dist.get_truncated_bounds()
    truncate_type = getattr(dist, "truncate_type", None)

    try:
        if truncate_type in ["value", "percentile"]:
            lower_p = 0.0 if lower is None else float(dist._original_cdf(lower))
            upper_p = 1.0 if upper is None else float(dist._original_cdf(upper))
            original_q = lower_p + q * (upper_p - lower_p)
            result = np.asarray(scipy_dist.ppf(original_q), dtype=float)
            if lower is not None:
                result = np.maximum(result, lower)
            if upper is not None:
                result = np.minimum(result, upper)
            return result + shift_amount

        if truncate_type in ["value2", "percentile2"]:
            lower_orig = None if lower is None else lower - shift_amount
            upper_orig = None if upper is None else upper - shift_amount
            lower_p = 0.0 if lower_orig is None else float(dist._original_cdf(lower_orig))
            upper_p = 1.0 if upper_orig is None else float(dist._original_cdf(upper_orig))
            original_q = lower_p + q * (upper_p - lower_p)
            result = np.asarray(scipy_dist.ppf(original_q), dtype=float)
            shifted_result = result + shift_amount
            if lower is not None:
                result = np.where(shifted_result < lower, lower - shift_amount, result)
            if upper is not None:
                result = np.where(shifted_result > upper, upper - shift_amount, result)
            return result + shift_amount
    except Exception:
        return None

    return None


def _ppf_many(dist: Any, q_values: np.ndarray) -> np.ndarray:
    fast = _fast_vectorized_ppf(dist, q_values)
    if fast is not None:
        return np.asarray(fast, dtype=float)
    return np.asarray([float(dist.ppf(float(q))) for q in q_values], dtype=float)


def _build_core(left_formula: Any, right_formula: Any, splice_point: Any) -> Dict[str, Any]:
    left_text, right_text, splice_value = _normalize_params(left_formula, right_formula, splice_point)
    cache_key = (left_text, right_text, splice_value)
    cached = _CORE_CACHE.get(cache_key)
    if cached is not None:
        return cached

    left_dist = _parse_distribution(left_text)
    right_dist = _parse_distribution(right_text)
    left_is_discrete = _is_discrete_distribution(left_dist)
    right_is_discrete = _is_discrete_distribution(right_dist)

    left_support_low = _safe_min_value(left_dist)
    right_support_high = _safe_max_value(right_dist)

    f1 = _mass_or_density(left_dist, left_is_discrete, splice_value)
    f2 = _mass_or_density(right_dist, right_is_discrete, splice_value)
    p1 = float(left_dist.cdf(splice_value))
    p2 = float(max(0.0, 1.0 - right_dist.cdf(splice_value)))

    if not math.isfinite(f1) or f1 <= 0.0:
        raise ValueError("Left splice density/mass must be positive at the splice point")
    if not math.isfinite(f2) or f2 <= 0.0:
        raise ValueError("Right splice density/mass must be positive at the splice point")
    if p1 <= 0.0 or p1 >= 1.0:
        raise ValueError("Left splice cumulative probability must lie strictly between 0 and 1")
    if p2 <= 0.0 or p2 >= 1.0:
        raise ValueError("Right tail probability beyond the splice point must lie strictly between 0 and 1")

    denominator = f1 * p2 + f2 * p1
    if not math.isfinite(denominator) or denominator <= 0.0:
        raise ValueError("Unable to normalize splice distribution")

    k1 = f2 / denominator
    k2 = f1 / denominator
    w1 = k1 * p1
    total_area = w1 + k2 * p2
    if not math.isfinite(total_area) or abs(total_area - 1.0) > 1e-8:
        raise ValueError("Splice normalization failed")

    data = {
        "left_formula": left_text,
        "right_formula": right_text,
        "splice_point": splice_value,
        "left_dist": left_dist,
        "right_dist": right_dist,
        "left_is_discrete": left_is_discrete,
        "right_is_discrete": right_is_discrete,
        "left_support_low": left_support_low,
        "right_support_high": right_support_high,
        "f1": float(f1),
        "f2": float(f2),
        "p1": float(p1),
        "p2": float(p2),
        "k1": float(k1),
        "k2": float(k2),
        "w1": float(w1),
        "right_cdf_splice": float(right_dist.cdf(splice_value)),
    }
    _CORE_CACHE[cache_key] = data
    return data


def splice_generator_single(rng: np.random.Generator, params: List[Any]) -> float:
    left_formula, right_formula, splice_point = params[0], params[1], params[2]
    return float(splice_generator_vectorized(rng, [left_formula, right_formula, splice_point], 1)[0])


def splice_generator_vectorized(
    rng: np.random.Generator,
    params: List[Any],
    n_samples: int,
) -> np.ndarray:
    core = _build_core(params[0], params[1], params[2])
    q = rng.uniform(_EPS, 1.0 - _EPS, size=n_samples)
    samples = np.empty(n_samples, dtype=float)

    left_mask = q <= core["w1"]
    if np.any(left_mask):
        q_left = np.clip(q[left_mask] / core["k1"], _EPS, 1.0 - _EPS)
        samples[left_mask] = _ppf_many(core["left_dist"], q_left)

    if np.any(~left_mask):
        lower_q = math.nextafter(core["right_cdf_splice"], 1.0)
        q_right = (q[~left_mask] - core["w1"]) / core["k2"] + core["right_cdf_splice"]
        q_right = np.clip(q_right, lower_q, 1.0 - _EPS)
        samples[~left_mask] = _ppf_many(core["right_dist"], q_right)

    return samples


def splice_cdf(x: float, left_formula: Any, right_formula: Any, splice_point: Any) -> float:
    core = _build_core(left_formula, right_formula, splice_point)
    x = float(x)
    if x <= core["splice_point"]:
        return max(0.0, min(1.0, core["k1"] * float(core["left_dist"].cdf(x))))
    right_part = core["k2"] * (float(core["right_dist"].cdf(x)) - core["right_cdf_splice"])
    return max(0.0, min(1.0, core["w1"] + right_part))


def splice_ppf(q: float, left_formula: Any, right_formula: Any, splice_point: Any) -> float:
    core = _build_core(left_formula, right_formula, splice_point)
    q = float(q)
    if q <= 0.0:
        return float(core["left_support_low"])
    if q >= 1.0:
        return float(core["right_support_high"])
    if q <= core["w1"]:
        q_left = min(max(q / core["k1"], _EPS), 1.0 - _EPS)
        return float(core["left_dist"].ppf(q_left))
    lower_q = math.nextafter(core["right_cdf_splice"], 1.0)
    q_right = (q - core["w1"]) / core["k2"] + core["right_cdf_splice"]
    q_right = min(max(q_right, lower_q), 1.0 - _EPS)
    return float(core["right_dist"].ppf(q_right))


def splice_raw_mean(left_formula: Any, right_formula: Any, splice_point: Any) -> float:
    dist = SpliceDistribution([left_formula, right_formula, splice_point], {}, "DriskSplice")
    return float(dist.mean())


class SpliceDistribution(DistributionBase):
    """Splice distribution with exact CDF/PPF and numerically stable moments."""

    def __init__(self, params: List[Any], markers: Dict[str, Any] = None, func_name: str = None):
        markers = markers or {}
        left_formula = params[0] if len(params) >= 1 else ""
        right_formula = params[1] if len(params) >= 2 else ""
        splice_point = params[2] if len(params) >= 3 else 0.0

        core = _build_core(left_formula, right_formula, splice_point)
        self.left_formula = core["left_formula"]
        self.right_formula = core["right_formula"]
        self.splice_point = core["splice_point"]
        self.left_dist = core["left_dist"]
        self.right_dist = core["right_dist"]
        self.left_is_discrete = core["left_is_discrete"]
        self.right_is_discrete = core["right_is_discrete"]
        self.k1 = core["k1"]
        self.k2 = core["k2"]
        self.w1 = core["w1"]
        self._right_cdf_splice = core["right_cdf_splice"]
        self._raw_mode = float("nan")
        self._raw_mean = float("nan")
        self._raw_var = float("nan")
        self._raw_skew = float("nan")
        self._raw_kurt = float("nan")
        self._truncated_stats: Optional[Dict[str, float]] = None

        super().__init__([self.left_formula, self.right_formula, self.splice_point], markers, func_name)

        self.support_low = core["left_support_low"]
        self.support_high = core["right_support_high"]

        if self.truncate_type in ["percentile", "percentile2"]:
            if self.truncate_lower_pct is not None:
                lower_value = self._original_ppf(self.truncate_lower_pct)
                if self.truncate_type == "percentile2":
                    lower_value += self.shift_amount
                self.truncate_lower = lower_value
            if self.truncate_upper_pct is not None:
                upper_value = self._original_ppf(self.truncate_upper_pct)
                if self.truncate_type == "percentile2":
                    upper_value += self.shift_amount
                self.truncate_upper = upper_value

        self._finalize_truncation()
        self._compute_raw_stats()
        self._compute_truncated_stats()

    def _original_pdf(self, x: float) -> float:
        x = float(x)
        if x <= self.splice_point:
            return self.k1 * _mass_or_density(self.left_dist, self.left_is_discrete, x)
        return self.k2 * _mass_or_density(self.right_dist, self.right_is_discrete, x)

    def _original_cdf(self, x: float) -> float:
        return splice_cdf(x, self.left_formula, self.right_formula, self.splice_point)

    def _original_ppf(self, q_prob: float) -> float:
        return splice_ppf(q_prob, self.left_formula, self.right_formula, self.splice_point)

    def _left_component_points(self, upper: float) -> np.ndarray:
        points = _discrete_support_points(self.left_dist, upper_hint=upper)
        return points[points <= upper + 1e-9]

    def _right_component_points(self, lower: float, upper: Optional[float]) -> np.ndarray:
        hint = upper if upper is not None and math.isfinite(upper) else None
        points = _discrete_support_points(self.right_dist, upper_hint=hint)
        mask = points > lower + 1e-9
        if upper is not None:
            mask &= points <= upper + 1e-9
        return points[mask]

    def _continuous_integral(self, dist: Any, lower: float, upper: float, power: int) -> float:
        if upper <= lower:
            return 0.0

        lb = lower
        ub = upper
        if not math.isfinite(lb):
            lb = _finite_bound_from_tail(dist, upper=False)
        if not math.isfinite(ub):
            ub = _finite_bound_from_tail(dist, upper=True)
        if ub <= lb:
            return 0.0

        def integrand(x: float) -> float:
            return (x ** power) * float(dist.pdf(x))

        result, _ = quad(integrand, lb, ub, limit=400)
        return float(result)

    def _component_moment(
        self,
        dist: Any,
        is_discrete: bool,
        weight: float,
        lower_exclusive: Optional[float],
        upper_inclusive: Optional[float],
        is_left: bool,
        power: int,
    ) -> float:
        if is_left:
            upper = self.splice_point if upper_inclusive is None else min(upper_inclusive, self.splice_point)
            if lower_exclusive is not None and upper <= lower_exclusive:
                return 0.0
            if is_discrete:
                points = self._left_component_points(upper)
                if lower_exclusive is not None:
                    points = points[points > lower_exclusive + 1e-9]
                if points.size == 0:
                    return 0.0
                masses = np.asarray([_mass_or_density(dist, True, float(x)) for x in points], dtype=float)
                return float(weight * np.sum((points ** power) * masses))
            lower = _safe_min_value(dist) if lower_exclusive is None else lower_exclusive
            return float(weight * self._continuous_integral(dist, lower, upper, power))

        lower = self.splice_point if lower_exclusive is None else max(lower_exclusive, self.splice_point)
        upper = upper_inclusive
        if upper is not None and upper <= lower:
            return 0.0
        if is_discrete:
            points = self._right_component_points(lower, upper)
            if points.size == 0:
                return 0.0
            masses = np.asarray([_mass_or_density(dist, True, float(x)) for x in points], dtype=float)
            return float(weight * np.sum((points ** power) * masses))
        upper_bound = _safe_max_value(dist) if upper is None else upper
        return float(weight * self._continuous_integral(dist, lower, upper_bound, power))

    def _moment_between(
        self, power: int, lower_exclusive: Optional[float] = None, upper_inclusive: Optional[float] = None
    ) -> float:
        left = self._component_moment(
            self.left_dist,
            self.left_is_discrete,
            self.k1,
            lower_exclusive,
            upper_inclusive,
            True,
            power,
        )
        right = self._component_moment(
            self.right_dist,
            self.right_is_discrete,
            self.k2,
            lower_exclusive,
            upper_inclusive,
            False,
            power,
        )
        return float(left + right)

    def _finite_mode_bounds(self, dist: Any, lower: float, upper: float) -> Tuple[float, float]:
        lb = lower if math.isfinite(lower) else _finite_bound_from_tail(dist, upper=False)
        ub = upper if math.isfinite(upper) else _finite_bound_from_tail(dist, upper=True)
        return lb, ub

    def _mode_on_interval(self, lower_exclusive: Optional[float], upper_inclusive: Optional[float]) -> float:
        best_x = float("nan")
        best_y = -float("inf")

        def consider(x: float) -> None:
            nonlocal best_x, best_y
            y = self._original_pdf(x)
            if math.isfinite(y) and y > best_y:
                best_x = float(x)
                best_y = float(y)

        left_upper = self.splice_point if upper_inclusive is None else min(upper_inclusive, self.splice_point)
        if lower_exclusive is None or left_upper > lower_exclusive:
            if self.left_is_discrete:
                for x in self._left_component_points(left_upper):
                    if lower_exclusive is None or x > lower_exclusive + 1e-9:
                        consider(float(x))
            else:
                left_lower = _safe_min_value(self.left_dist) if lower_exclusive is None else max(
                    _safe_min_value(self.left_dist), lower_exclusive
                )
                lb, ub = self._finite_mode_bounds(self.left_dist, left_lower, left_upper)
                if ub > lb:
                    consider(lb)
                    consider(ub)
                    try:
                        candidate = float(self.left_dist.mode())
                        if lb <= candidate <= ub:
                            consider(candidate)
                    except Exception:
                        pass
                    try:
                        result = minimize_scalar(
                            lambda x: -self.k1 * float(self.left_dist.pdf(x)),
                            bounds=(lb, ub),
                            method="bounded",
                        )
                        consider(float(result.x))
                    except Exception:
                        pass

        right_lower = self.splice_point if lower_exclusive is None else max(lower_exclusive, self.splice_point)
        if upper_inclusive is None or upper_inclusive > right_lower:
            if self.right_is_discrete:
                for x in self._right_component_points(right_lower, upper_inclusive):
                    consider(float(x))
            else:
                right_upper = _safe_max_value(self.right_dist) if upper_inclusive is None else min(
                    _safe_max_value(self.right_dist), upper_inclusive
                )
                lb, ub = self._finite_mode_bounds(self.right_dist, right_lower, right_upper)
                if ub > lb:
                    consider(lb)
                    consider(ub)
                    try:
                        candidate = float(self.right_dist.mode())
                        if lb <= candidate <= ub:
                            consider(candidate)
                    except Exception:
                        pass
                    try:
                        result = minimize_scalar(
                            lambda x: -self.k2 * float(self.right_dist.pdf(x)),
                            bounds=(lb, ub),
                            method="bounded",
                        )
                        consider(float(result.x))
                    except Exception:
                        pass

        if not math.isfinite(best_x):
            return float(self.splice_point)
        return best_x

    def _compute_raw_stats(self) -> None:
        m1 = self._moment_between(1)
        m2 = self._moment_between(2)
        m3 = self._moment_between(3)
        m4 = self._moment_between(4)

        variance = max(0.0, m2 - m1 * m1)
        if variance > _EPS:
            mu3 = m3 - 3.0 * m1 * m2 + 2.0 * (m1 ** 3)
            mu4 = m4 - 4.0 * m1 * m3 + 6.0 * (m1 ** 2) * m2 - 3.0 * (m1 ** 4)
            skewness = mu3 / (variance ** 1.5)
            kurtosis = mu4 / (variance ** 2)
        else:
            skewness = 0.0
            kurtosis = 3.0

        self._raw_mean = float(m1)
        self._raw_var = float(variance)
        self._raw_skew = float(skewness)
        self._raw_kurt = float(kurtosis)
        self._raw_mode = float(self._mode_on_interval(None, None))

    def _get_original_truncation_bounds(self) -> Tuple[Optional[float], Optional[float]]:
        low, high = self.get_truncated_bounds()
        if low is None and high is None:
            return None, None
        if self.truncate_type in ["value2", "percentile2"]:
            low = low - self.shift_amount if low is not None else None
            high = high - self.shift_amount if high is not None else None
        return low, high

    def _compute_truncated_stats(self) -> None:
        if not self.is_truncated():
            self._truncated_stats = None
            return

        low_orig, high_orig = self._get_original_truncation_bounds()
        if high_orig is not None and low_orig is not None and high_orig <= low_orig:
            self._truncated_stats = None
            return

        z = max(0.0, self._original_cdf(float("inf") if high_orig is None else high_orig) - self._original_cdf(
            float("-inf") if low_orig is None else low_orig
        ))
        if z <= _EPS:
            self._truncated_stats = None
            return

        m1 = self._moment_between(1, low_orig, high_orig) / z
        m2 = self._moment_between(2, low_orig, high_orig) / z
        m3 = self._moment_between(3, low_orig, high_orig) / z
        m4 = self._moment_between(4, low_orig, high_orig) / z
        variance = max(0.0, m2 - m1 * m1)

        if variance > _EPS:
            mu3 = m3 - 3.0 * m1 * m2 + 2.0 * (m1 ** 3)
            mu4 = m4 - 4.0 * m1 * m3 + 6.0 * (m1 ** 2) * m2 - 3.0 * (m1 ** 4)
            skewness = mu3 / (variance ** 1.5)
            kurtosis = mu4 / (variance ** 2)
        else:
            skewness = 0.0
            kurtosis = 3.0

        self._truncated_stats = {
            "mean": float(m1),
            "variance": float(variance),
            "skewness": float(skewness),
            "kurtosis": float(kurtosis),
            "mode": float(self._mode_on_interval(low_orig, high_orig)),
        }

    def mean(self) -> float:
        if self._truncate_invalid:
            return float("nan")
        if self._truncated_stats is not None:
            return float(self._truncated_stats["mean"] + self.shift_amount)
        return float(self._raw_mean + self.shift_amount)

    def variance(self) -> float:
        if self._truncate_invalid:
            return float("nan")
        if self._truncated_stats is not None:
            return float(self._truncated_stats["variance"])
        return float(self._raw_var)

    def skewness(self) -> float:
        if self._truncate_invalid:
            return float("nan")
        if self._truncated_stats is not None:
            return float(self._truncated_stats["skewness"])
        return float(self._raw_skew)

    def kurtosis(self) -> float:
        if self._truncate_invalid:
            return float("nan")
        if self._truncated_stats is not None:
            return float(self._truncated_stats["kurtosis"])
        return float(self._raw_kurt)

    def mode(self) -> float:
        if self._truncate_invalid:
            return float("nan")
        if self._truncated_stats is not None:
            return float(self._truncated_stats["mode"] + self.shift_amount)
        return float(self._raw_mode + self.shift_amount)

    def min_val(self) -> float:
        if self._truncate_invalid:
            return float("nan")
        lower, _ = self.get_effective_bounds()
        if lower is not None:
            return float(lower)
        return float(self.support_low + self.shift_amount)

    def max_val(self) -> float:
        if self._truncate_invalid:
            return float("nan")
        _, upper = self.get_effective_bounds()
        if upper is not None:
            return float(upper)
        if math.isfinite(self.support_high):
            return float(self.support_high + self.shift_amount)
        return float("inf")
