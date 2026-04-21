"""Reciprocal distribution support for Drisk."""

import math
from typing import List, Optional, Tuple, Union

import numpy as np
from scipy.stats import reciprocal as scipy_reciprocal

from distribution_base import DistributionBase

_EPS = 1e-12


def _validate_params(min_val: float, max_val: float) -> None:
    if float(min_val) <= 0.0:
        raise ValueError("Reciprocal requires Min > 0")
    if float(max_val) <= float(min_val):
        raise ValueError("Reciprocal requires Max > Min")


def _create_dist(min_val: float, max_val: float):
    _validate_params(min_val, max_val)
    return scipy_reciprocal(a=float(min_val), b=float(max_val))


def _log_ratio(min_val: float, max_val: float) -> float:
    _validate_params(min_val, max_val)
    return float(math.log(float(max_val) / float(min_val)))


def reciprocal_pdf(x: float, min_val: float, max_val: float) -> float:
    x = float(x)
    if x < float(min_val) or x > float(max_val):
        return 0.0
    return float(1.0 / (x * _log_ratio(min_val, max_val)))


def reciprocal_cdf(x: float, min_val: float, max_val: float) -> float:
    x = float(x)
    if x <= float(min_val):
        return 0.0
    if x >= float(max_val):
        return 1.0
    log_ratio = _log_ratio(min_val, max_val)
    return float(math.log(x / float(min_val)) / log_ratio)


def reciprocal_ppf(q_prob: float, min_val: float, max_val: float) -> float:
    q_prob = float(q_prob)
    _validate_params(min_val, max_val)
    if q_prob <= 0.0:
        return float(min_val)
    if q_prob >= 1.0:
        return float(max_val)
    return float(float(min_val) * math.exp(q_prob * _log_ratio(min_val, max_val)))


def reciprocal_raw_mean(min_val: float, max_val: float) -> float:
    log_ratio = _log_ratio(min_val, max_val)
    return float((float(max_val) - float(min_val)) / log_ratio)


def reciprocal_generator_single(rng: np.random.Generator, params: List[float]) -> float:
    min_val = float(params[0])
    max_val = float(params[1])
    u = float(rng.uniform(_EPS, 1.0 - _EPS))
    return reciprocal_ppf(u, min_val, max_val)


def reciprocal_generator_vectorized(
    rng: np.random.Generator,
    params: List[Union[float, np.ndarray]],
    n_samples: int,
) -> np.ndarray:
    min_val = params[0]
    max_val = params[1]

    if not isinstance(min_val, np.ndarray):
        min_val = np.full(n_samples, float(min_val))
    if not isinstance(max_val, np.ndarray):
        max_val = np.full(n_samples, float(max_val))

    if np.any((min_val <= 0) | (max_val <= min_val)):
        raise ValueError("Reciprocal requires 0 < min < max")

    u = rng.uniform(_EPS, 1.0 - _EPS, size=n_samples)
    log_ratio = np.log(max_val / min_val)
    samples = min_val * np.exp(u * log_ratio)
    return samples


class ReciprocalDistribution(DistributionBase):
    """Reciprocal theoretical distribution with truncation support."""

    def __init__(self, params: List[float], markers: dict = None, func_name: str = None):
        markers = markers or {}
        min_val = float(params[0]) if len(params) >= 1 else 1.0
        max_val = float(params[1]) if len(params) >= 2 else 2.0

        _validate_params(min_val, max_val)

        self.min_point = min_val
        self.max_point = max_val
        self._dist = _create_dist(min_val, max_val)
        self.support_low = float(min_val)
        self.support_high = float(max_val)

        log_ratio = _log_ratio(min_val, max_val)
        q1 = (max_val - min_val) / log_ratio
        q2 = (max_val ** 2 - min_val ** 2) / (2.0 * log_ratio)
        q3 = (max_val ** 3 - min_val ** 3) / (3.0 * log_ratio)
        q4 = (max_val ** 4 - min_val ** 4) / (4.0 * log_ratio)

        self._raw_mean = float(q1)
        self._raw_var = float(q2 - q1 ** 2)

        if self._raw_var > _EPS:
            numerator = q3 - 3.0 * q1 * q2 + 2.0 * q1 ** 3
            self._raw_skew = float(numerator / (self._raw_var ** 1.5))
            numerator_kurt = q4 - 4.0 * q1 * q3 + 6.0 * (q1 ** 2) * q2 - 3.0 * q1 ** 4
            self._raw_kurt = float(numerator_kurt / (self._raw_var ** 2))
        else:
            self._raw_skew = 0.0
            self._raw_kurt = 3.0

        self._raw_mode = float(min_val)
        self._truncated_stats = None

        super().__init__([min_val, max_val], markers, func_name)

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
        return reciprocal_pdf(x, self.min_point, self.max_point)

    def _original_cdf(self, x: float) -> float:
        return reciprocal_cdf(x, self.min_point, self.max_point)

    def _original_ppf(self, q_prob: float) -> float:
        return reciprocal_ppf(q_prob, self.min_point, self.max_point)

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
        lb = self.min_point if low_orig is None else max(self.min_point, low_orig)
        ub = self.max_point if high_orig is None else min(self.max_point, high_orig)
        if ub < lb:
            self._truncated_stats = None
            return

        mass = max(0.0, self._original_cdf(ub) - self._original_cdf(lb))
        if mass <= _EPS:
            self._truncated_stats = None
            return

        m1 = float(self._dist.expect(lambda x: x, lb=lb, ub=ub, conditional=True))
        m2 = float(self._dist.expect(lambda x: x * x, lb=lb, ub=ub, conditional=True))
        m3_raw = float(self._dist.expect(lambda x: x ** 3, lb=lb, ub=ub, conditional=True))
        m4_raw = float(self._dist.expect(lambda x: x ** 4, lb=lb, ub=ub, conditional=True))

        variance = max(0.0, m2 - m1 * m1)
        if variance > _EPS:
            mu3 = m3_raw - 3.0 * m1 * m2 + 2.0 * m1 ** 3
            mu4 = m4_raw - 4.0 * m1 * m3_raw + 6.0 * m1 * m1 * m2 - 3.0 * m1 ** 4
            skewness = mu3 / (variance ** 1.5)
            kurtosis = mu4 / (variance ** 2)
        elif variance == 0.0:
            skewness = 0.0
            kurtosis = 3.0
        else:
            skewness = float("nan")
            kurtosis = float("nan")

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
        return float(self.min_point + self.shift_amount)

    def max_val(self) -> float:
        if self._truncate_invalid:
            return float("nan")
        _, upper = self.get_effective_bounds()
        if upper is not None:
            return upper
        return float(self.max_point + self.shift_amount)
