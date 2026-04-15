"""Levy distribution support for Drisk."""

import math
from typing import List, Optional, Tuple

import numpy as np
from scipy.integrate import quad
from scipy.stats import levy as scipy_levy

from distribution_base import DistributionBase

_EPS = 1e-12


def _validate_params(a: float, c: float) -> None:
    if c <= 0.0:
        raise ValueError("Levy requires c > 0")


def _create_dist(a: float, c: float):
    _validate_params(a, c)
    return scipy_levy(loc=float(a), scale=float(c))


def levy_pdf(x: float, a: float, c: float) -> float:
    if float(x) <= float(a):
        return 0.0
    dist = _create_dist(a, c)
    return float(dist.pdf(x))


def levy_cdf(x: float, a: float, c: float) -> float:
    if float(x) <= float(a):
        return 0.0
    dist = _create_dist(a, c)
    return float(dist.cdf(x))


def levy_ppf(q: float, a: float, c: float) -> float:
    q = float(q)
    if q <= 0.0:
        return float(a)
    if q >= 1.0:
        return float("inf")
    dist = _create_dist(a, c)
    return float(dist.ppf(q))


def levy_raw_mean(a: float, c: float) -> float:
    _validate_params(a, c)
    return float("nan")


def levy_generator_single(rng: np.random.Generator, params: List[float]) -> float:
    a = float(params[0])
    c = float(params[1])
    q = float(rng.uniform(_EPS, 1.0 - _EPS))
    return levy_ppf(q, a, c)


def levy_generator_vectorized(
    rng: np.random.Generator,
    params: List[float],
    n_samples: int,
) -> np.ndarray:
    a = float(params[0])
    c = float(params[1])
    dist = _create_dist(a, c)
    q = rng.uniform(_EPS, 1.0 - _EPS, size=n_samples)
    return np.asarray(dist.ppf(q), dtype=float)


class LevyDistribution(DistributionBase):
    """Levy theoretical distribution with truncation support."""

    def __init__(self, params: List[float], markers: dict = None, func_name: str = None):
        markers = markers or {}
        a = float(params[0]) if len(params) >= 1 else 0.0
        c = float(params[1]) if len(params) >= 2 else 1.0

        self.a = a
        self.c = c
        self._dist = _create_dist(a, c)
        self.support_low = float(a)
        self.support_high = float("inf")
        self._raw_mean = float("nan")
        self._raw_var = float("nan")
        self._raw_skew = float("nan")
        self._raw_kurt = float("nan")
        self._raw_mode = float(a + c / 3.0)
        self._truncated_stats = None

        super().__init__([a, c], markers, func_name)

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

    def _original_pdf(self, x: float) -> float:
        return levy_pdf(x, self.a, self.c)

    def _original_cdf(self, x: float) -> float:
        return levy_cdf(x, self.a, self.c)

    def _original_ppf(self, q_prob: float) -> float:
        return levy_ppf(q_prob, self.a, self.c)

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

    def _moment_integral(self, lower: float, upper: float, power: int) -> float:
        def integrand(x: float) -> float:
            return (x ** power) * self._original_pdf(x)

        result, _ = quad(integrand, lower, upper, limit=300)
        return float(result)

    def _compute_truncated_stats(self) -> None:
        if not self.is_truncated():
            self._truncated_stats = None
            return

        low_orig, high_orig = self._get_original_truncation_bounds()
        lb = self.a if low_orig is None else max(self.a, low_orig)
        ub = float("inf") if high_orig is None else high_orig
        if ub < lb:
            self._truncated_stats = None
            return

        if math.isinf(ub):
            self._truncated_stats = None
            return

        if lb <= self.a:
            lb = math.nextafter(self.a, float("inf"))

        mass = max(0.0, self._original_cdf(ub) - self._original_cdf(lb))
        if mass <= _EPS:
            self._truncated_stats = None
            return

        m1 = self._moment_integral(lb, ub, 1) / mass
        m2 = self._moment_integral(lb, ub, 2) / mass
        m3_raw = self._moment_integral(lb, ub, 3) / mass
        m4_raw = self._moment_integral(lb, ub, 4) / mass
        variance = max(0.0, m2 - m1 * m1)

        if variance > _EPS:
            mu3 = m3_raw - 3.0 * m1 * m2 + 2.0 * m1 ** 3
            mu4 = m4_raw - 4.0 * m1 * m3_raw + 6.0 * m1 * m1 * m2 - 3.0 * m1 ** 4
            skewness = mu3 / (variance ** 1.5)
            kurtosis = mu4 / (variance ** 2)
        else:
            skewness = 0.0
            kurtosis = 3.0

        mode_raw = min(max(self._raw_mode, lb), ub)
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
        return self._raw_mean

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
            mode_value = self._truncated_stats["mode"] + self.shift_amount
        else:
            mode_value = self._raw_mode + self.shift_amount

        if self.is_truncated():
            lower, upper = self.get_effective_bounds()
            if lower is not None:
                mode_value = max(mode_value, lower)
            if upper is not None:
                mode_value = min(mode_value, upper)

        return float(mode_value)

    def min_val(self) -> float:
        if self._truncate_invalid:
            return float("nan")
        lower, _ = self.get_effective_bounds()
        if lower is not None:
            return lower
        return float(self.a + self.shift_amount)

    def max_val(self) -> float:
        if self._truncate_invalid:
            return float("nan")
        _, upper = self.get_effective_bounds()
        if upper is not None:
            return upper
        return float("inf")
