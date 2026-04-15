"""
General distribution support for Drisk.

Definition used here:
    X ~ General(min_val, max_val, x_table, p_table)

This is a bounded continuous distribution on [min_val, max_val].
The internal x_table/p_table define a piecewise-linear PDF profile.
Boundary points min_val and max_val are added with zero density and the
resulting PDF is normalized to integrate to 1.
"""

import math
from typing import List, Optional, Tuple, Union

import numpy as np

from distribution_base import DistributionBase

_EPS = 1e-12


def _extract_numbers_from_input(val) -> List[float]:
    numbers: List[float] = []

    def extract(item) -> None:
        if isinstance(item, (tuple, list, np.ndarray)):
            for sub in item:
                extract(sub)
        elif isinstance(item, str):
            clean = item.strip()
            if clean.startswith("{") and clean.endswith("}"):
                clean = clean[1:-1].strip()
            parts = [p.strip() for p in clean.split(",") if p.strip()]
            for part in parts:
                try:
                    numbers.append(float(part))
                except Exception:
                    pass
        else:
            try:
                if item is not None:
                    numbers.append(float(item))
            except Exception:
                pass

    extract(val)
    return numbers


def _parse_general_arrays(
    min_val: float, max_val: float, x_input, p_input
) -> Tuple[List[float], List[float]]:
    x_vals = _extract_numbers_from_input(x_input)
    p_vals = _extract_numbers_from_input(p_input)
    _validate_general_params(min_val, max_val, x_vals, p_vals)
    return x_vals, p_vals


def _validate_general_params(
    min_val: float, max_val: float, x_vals: List[float], p_vals: List[float]
) -> None:
    if max_val <= min_val:
        raise ValueError("General requires min_val < max_val")
    if len(x_vals) != len(p_vals):
        raise ValueError("X-Table and P-Table must have the same length")
    if len(x_vals) < 1:
        raise ValueError("General requires at least one internal X point")
    prev = None
    for x in x_vals:
        if not (min_val < x < max_val):
            raise ValueError("Internal X points must satisfy min_val < x < max_val")
        if prev is not None and x <= prev:
            raise ValueError("X-Table must be strictly increasing")
        prev = x
    if any(p < 0.0 for p in p_vals):
        raise ValueError("P-Table values must be nonnegative")
    if max(p_vals) <= 0.0:
        raise ValueError("P-Table must contain at least one positive value")


def _build_full_points(
    min_val: float, max_val: float, x_vals: List[float], p_vals: List[float]
) -> Tuple[np.ndarray, np.ndarray]:
    x_points = np.array([float(min_val)] + [float(x) for x in x_vals] + [float(max_val)], dtype=float)
    p_points = np.array([0.0] + [float(p) for p in p_vals] + [0.0], dtype=float)
    total_area = 0.0
    for i in range(len(x_points) - 1):
        total_area += 0.5 * (p_points[i] + p_points[i + 1]) * (x_points[i + 1] - x_points[i])
    if total_area <= _EPS:
        raise ValueError("General PDF area must be positive")
    p_points = p_points / total_area
    return x_points, p_points


def _segment_pdf_coeffs(x1: float, x2: float, p1: float, p2: float) -> Tuple[float, float]:
    if abs(x2 - x1) <= _EPS:
        return 0.0, p1
    slope = (p2 - p1) / (x2 - x1)
    intercept = p1 - slope * x1
    return slope, intercept


def _integrate_power_linear(
    power: int, lower: float, upper: float, slope: float, intercept: float
) -> float:
    if upper <= lower:
        return 0.0
    term1 = 0.0
    if abs(slope) > _EPS:
        term1 = slope * (upper ** (power + 2) - lower ** (power + 2)) / (power + 2)
    term2 = intercept * (upper ** (power + 1) - lower ** (power + 1)) / (power + 1)
    return float(term1 + term2)


class _GeneralCore:
    def __init__(self, min_val: float, max_val: float, x_vals: List[float], p_vals: List[float]):
        _validate_general_params(min_val, max_val, x_vals, p_vals)
        self.min_val = float(min_val)
        self.max_val = float(max_val)
        self.x_vals = [float(x) for x in x_vals]
        self.p_vals = [float(p) for p in p_vals]
        self.x_points, self.p_points = _build_full_points(self.min_val, self.max_val, self.x_vals, self.p_vals)
        self.n_segments = len(self.x_points) - 1
        self.cdf_points = np.zeros(len(self.x_points), dtype=float)
        self.slopes: List[float] = []
        self.intercepts: List[float] = []
        for i in range(self.n_segments):
            x1 = self.x_points[i]
            x2 = self.x_points[i + 1]
            p1 = self.p_points[i]
            p2 = self.p_points[i + 1]
            slope, intercept = _segment_pdf_coeffs(x1, x2, p1, p2)
            self.slopes.append(slope)
            self.intercepts.append(intercept)
            area = 0.5 * (p1 + p2) * (x2 - x1)
            self.cdf_points[i + 1] = self.cdf_points[i] + area
        self.cdf_points[-1] = 1.0

    def pdf(self, x: float) -> float:
        x = float(x)
        if x < self.min_val - _EPS or x > self.max_val + _EPS:
            return 0.0
        if x <= self.min_val + _EPS:
            return float(self.p_points[0])
        if x >= self.max_val - _EPS:
            return float(self.p_points[-1])
        idx = int(np.searchsorted(self.x_points, x, side="right") - 1)
        idx = max(0, min(idx, self.n_segments - 1))
        return float(self.slopes[idx] * x + self.intercepts[idx])

    def cdf(self, x: float) -> float:
        x = float(x)
        if x <= self.min_val:
            return 0.0
        if x >= self.max_val:
            return 1.0
        idx = int(np.searchsorted(self.x_points, x, side="right") - 1)
        idx = max(0, min(idx, self.n_segments - 1))
        x1 = self.x_points[idx]
        p1 = self.p_points[idx]
        p2 = self.p_points[idx + 1]
        width = self.x_points[idx + 1] - x1
        if width <= _EPS:
            return float(self.cdf_points[idx])
        t = (x - x1) / width
        area = width * (p1 * t + 0.5 * (p2 - p1) * t * t)
        return float(self.cdf_points[idx] + area)

    def ppf(self, q: float) -> float:
        q = float(q)
        if q <= 0.0:
            return self.min_val
        if q >= 1.0:
            return self.max_val
        idx = int(np.searchsorted(self.cdf_points, q, side="right") - 1)
        idx = max(0, min(idx, self.n_segments - 1))
        x1 = self.x_points[idx]
        x2 = self.x_points[idx + 1]
        p1 = self.p_points[idx]
        p2 = self.p_points[idx + 1]
        target = q - self.cdf_points[idx]
        width = x2 - x1
        if width <= _EPS:
            return float(x1)

        a = 0.5 * width * (p2 - p1)
        b = width * p1
        if abs(a) <= _EPS:
            if abs(b) <= _EPS:
                t = 0.0
            else:
                t = target / b
        else:
            c = -target
            disc = b * b - 4.0 * a * c
            disc = max(0.0, disc)
            sqrt_disc = math.sqrt(disc)
            roots = [(-b + sqrt_disc) / (2.0 * a), (-b - sqrt_disc) / (2.0 * a)]
            valid_roots = [r for r in roots if -_EPS <= r <= 1.0 + _EPS]
            if valid_roots:
                t = min(valid_roots, key=lambda r: abs(max(0.0, min(1.0, r)) - r))
            else:
                t = 0.0 if target <= 0.0 else 1.0
        t = max(0.0, min(1.0, t))
        return float(x1 + t * width)

    def moment_integrals(self, lower: float, upper: float, max_power: int = 4) -> List[float]:
        lower = max(float(lower), self.min_val)
        upper = min(float(upper), self.max_val)
        if upper < lower:
            return [0.0] * (max_power + 1)
        results = [0.0] * (max_power + 1)
        for i in range(self.n_segments):
            seg_low = max(lower, self.x_points[i])
            seg_high = min(upper, self.x_points[i + 1])
            if seg_high <= seg_low:
                continue
            slope = self.slopes[i]
            intercept = self.intercepts[i]
            for power in range(max_power + 1):
                results[power] += _integrate_power_linear(power, seg_low, seg_high, slope, intercept)
        return results

    def mode_on_interval(self, lower: float, upper: float) -> float:
        lower = max(float(lower), self.min_val)
        upper = min(float(upper), self.max_val)
        candidates = [lower, upper]
        for x in self.x_points:
            if lower - _EPS <= x <= upper + _EPS:
                candidates.append(float(x))
        best_x = lower
        best_pdf = -1.0
        for x in candidates:
            pdf_val = self.pdf(x)
            if pdf_val > best_pdf + _EPS or (abs(pdf_val - best_pdf) <= _EPS and x < best_x):
                best_pdf = pdf_val
                best_x = x
        return float(best_x)


def general_generator_single(
    rng: np.random.Generator,
    params: List[float],
    x_vals: List[float],
    p_vals: List[float],
) -> float:
    min_val = float(params[0])
    max_val = float(params[1])
    core = _GeneralCore(min_val, max_val, x_vals, p_vals)
    return core.ppf(float(rng.uniform(_EPS, 1.0 - _EPS)))


def general_generator_vectorized(
    rng: np.random.Generator,
    params: List[float],
    n_samples: int,
    x_vals: List[float],
    p_vals: List[float],
) -> np.ndarray:
    min_val = float(params[0])
    max_val = float(params[1])
    core = _GeneralCore(min_val, max_val, x_vals, p_vals)
    u = rng.uniform(_EPS, 1.0 - _EPS, size=n_samples)
    return np.array([core.ppf(float(q)) for q in u], dtype=float)


def general_generator(
    rng: np.random.Generator,
    params: List[float],
    n_samples: Optional[int] = None,
    x_vals: List[float] = None,
    p_vals: List[float] = None,
) -> Union[float, np.ndarray]:
    if x_vals is None or p_vals is None:
        raise ValueError("general_generator requires x_vals and p_vals")
    if n_samples is None:
        return general_generator_single(rng, params, x_vals, p_vals)
    return general_generator_vectorized(rng, params, n_samples, x_vals, p_vals)


def general_pdf(
    x: float, min_val: float, max_val: float, x_vals: List[float], p_vals: List[float]
) -> float:
    return _GeneralCore(min_val, max_val, x_vals, p_vals).pdf(x)


def general_cdf(
    x: float, min_val: float, max_val: float, x_vals: List[float], p_vals: List[float]
) -> float:
    return _GeneralCore(min_val, max_val, x_vals, p_vals).cdf(x)


def general_ppf(
    q: float, min_val: float, max_val: float, x_vals: List[float], p_vals: List[float]
) -> float:
    return _GeneralCore(min_val, max_val, x_vals, p_vals).ppf(q)


def general_raw_mean(min_val: float, max_val: float, x_vals: List[float], p_vals: List[float]) -> float:
    core = _GeneralCore(min_val, max_val, x_vals, p_vals)
    return float(core.moment_integrals(core.min_val, core.max_val, 1)[1])


def general_raw_var(min_val: float, max_val: float, x_vals: List[float], p_vals: List[float]) -> float:
    core = _GeneralCore(min_val, max_val, x_vals, p_vals)
    integrals = core.moment_integrals(core.min_val, core.max_val, 2)
    mean_val = integrals[1]
    variance = max(0.0, integrals[2] - mean_val * mean_val)
    return float(variance)


def general_raw_skewness(min_val: float, max_val: float, x_vals: List[float], p_vals: List[float]) -> float:
    core = _GeneralCore(min_val, max_val, x_vals, p_vals)
    integrals = core.moment_integrals(core.min_val, core.max_val, 3)
    mean_val = integrals[1]
    variance = max(0.0, integrals[2] - mean_val * mean_val)
    if variance <= _EPS:
        return 0.0
    mu3 = integrals[3] - 3.0 * mean_val * integrals[2] + 2.0 * mean_val ** 3
    return float(mu3 / (variance ** 1.5))


def general_raw_kurtosis(min_val: float, max_val: float, x_vals: List[float], p_vals: List[float]) -> float:
    core = _GeneralCore(min_val, max_val, x_vals, p_vals)
    integrals = core.moment_integrals(core.min_val, core.max_val, 4)
    mean_val = integrals[1]
    variance = max(0.0, integrals[2] - mean_val * mean_val)
    if variance <= _EPS:
        return 3.0
    mu4 = (
        integrals[4]
        - 4.0 * mean_val * integrals[3]
        + 6.0 * mean_val * mean_val * integrals[2]
        - 3.0 * mean_val ** 4
    )
    return float(mu4 / (variance ** 2))


def general_raw_mode(min_val: float, max_val: float, x_vals: List[float], p_vals: List[float]) -> float:
    core = _GeneralCore(min_val, max_val, x_vals, p_vals)
    return core.mode_on_interval(core.min_val, core.max_val)


class GeneralDistribution(DistributionBase):
    """General theoretical distribution with truncation support."""

    def __init__(self, params: List[float], markers: dict = None, func_name: str = None):
        markers = markers or {}
        if len(params) >= 4:
            min_val = float(params[0])
            max_val = float(params[1])
            x_source = markers.get("x_vals", params[2])
            p_source = markers.get("p_vals", params[3])
        else:
            min_val, max_val = -1.0, 1.0
            x_source = markers.get("x_vals", "-0.5,0,0.5")
            p_source = markers.get("p_vals", "2,3,2")

        x_vals, p_vals = _parse_general_arrays(min_val, max_val, x_source, p_source)
        self.min_support = float(min_val)
        self.max_support = float(max_val)
        self.x_vals = x_vals
        self.p_vals = p_vals
        self._core = _GeneralCore(min_val, max_val, x_vals, p_vals)

        self.support_low = self.min_support
        self.support_high = self.max_support
        self._raw_mean = general_raw_mean(min_val, max_val, x_vals, p_vals)
        self._raw_var = general_raw_var(min_val, max_val, x_vals, p_vals)
        self._raw_skew = general_raw_skewness(min_val, max_val, x_vals, p_vals)
        self._raw_kurt = general_raw_kurtosis(min_val, max_val, x_vals, p_vals)
        self._raw_mode = general_raw_mode(min_val, max_val, x_vals, p_vals)
        self._truncated_stats = None

        super().__init__([min_val, max_val, params[2] if len(params) >= 3 else x_source, params[3] if len(params) >= 4 else p_source], markers, func_name)

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
        self._compute_truncated_stats()

    def _original_cdf(self, x: float) -> float:
        return self._core.cdf(x)

    def _original_ppf(self, q_prob: float) -> float:
        return self._core.ppf(q_prob)

    def _original_pdf(self, x: float) -> float:
        return self._core.pdf(x)

    def _get_original_truncation_bounds(self) -> Tuple[Optional[float], Optional[float]]:
        low, high = self.get_truncated_bounds()
        if low is None and high is None:
            return None, None
        if self.truncate_type in ["value2", "percentile2"]:
            low_orig = low - self.shift_amount if low is not None else None
            high_orig = high - self.shift_amount if high is not None else None
        else:
            low_orig = low
            high_orig = high
        return low_orig, high_orig

    def _compute_truncated_stats(self) -> None:
        if not self.is_truncated():
            self._truncated_stats = None
            return

        low_orig, high_orig = self._get_original_truncation_bounds()
        lb = self.support_low if low_orig is None else max(self.support_low, low_orig)
        ub = self.support_high if high_orig is None else min(self.support_high, high_orig)
        if ub < lb:
            self._truncated_stats = None
            return

        integrals = self._core.moment_integrals(lb, ub, 4)
        mass = integrals[0]
        if mass <= _EPS:
            self._truncated_stats = None
            return

        m1 = integrals[1] / mass
        m2 = integrals[2] / mass
        m3_raw = integrals[3] / mass
        m4_raw = integrals[4] / mass
        variance = max(0.0, m2 - m1 * m1)
        if variance > _EPS:
            mu3 = m3_raw - 3.0 * m1 * m2 + 2.0 * m1 ** 3
            mu4 = m4_raw - 4.0 * m1 * m3_raw + 6.0 * m1 * m1 * m2 - 3.0 * m1 ** 4
            skewness = mu3 / (variance ** 1.5)
            kurtosis = mu4 / (variance ** 2)
        else:
            skewness = 0.0
            kurtosis = 3.0

        mode_raw = self._core.mode_on_interval(lb, ub)
        self._truncated_stats = {
            "mean": float(m1),
            "variance": float(variance),
            "skewness": float(skewness),
            "kurtosis": float(kurtosis),
            "mode": float(mode_raw),
        }

    def mean(self) -> float:
        if self._truncate_invalid:
            return float("nan")
        if self._truncated_stats is not None:
            return self._truncated_stats["mean"] + self.shift_amount
        return self._raw_mean + self.shift_amount

    def variance(self) -> float:
        if self._truncate_invalid:
            return float("nan")
        if self._truncated_stats is not None:
            return self._truncated_stats["variance"]
        return self._raw_var

    def skewness(self) -> float:
        if self._truncate_invalid:
            return float("nan")
        if self._truncated_stats is not None:
            return self._truncated_stats["skewness"]
        return self._raw_skew

    def kurtosis(self) -> float:
        if self._truncate_invalid:
            return float("nan")
        if self._truncated_stats is not None:
            return self._truncated_stats["kurtosis"]
        return self._raw_kurt

    def mode(self) -> float:
        if self._truncate_invalid:
            return float("nan")
        if self._truncated_stats is not None:
            return self._truncated_stats["mode"] + self.shift_amount
        return self._raw_mode + self.shift_amount

    def min_val(self) -> float:
        if self._truncate_invalid:
            return float("nan")
        lower, _ = self.get_effective_bounds()
        if lower is not None:
            return lower
        return self.apply_shift(self.support_low)

    def max_val(self) -> float:
        if self._truncate_invalid:
            return float("nan")
        _, upper = self.get_effective_bounds()
        if upper is not None:
            return upper
        return self.apply_shift(self.support_high)
