"""
DoubleTriang distribution support for Drisk.

Definition used here:
    X ~ DoubleTriang(min_val, m_likely, max_val, lower_p)
    support = [min_val, max_val]

Piecewise density:
    For min_val <= x <= m_likely:
        PDF = 2 * lower_p * (x - min_val) / (m_likely - min_val)^2

    For m_likely < x <= max_val:
        PDF = 2 * (1 - lower_p) * (max_val - x) / (max_val - m_likely)^2

Parameters:
    min_val: minimum value
    m_likely: most likely value
    max_val: maximum value
    lower_p: probability weight of the lower triangle, must satisfy 0 <= p <= 1
"""

import math
from typing import List, Optional, Union
import numpy as np
from scipy.integrate import quad
from distribution_base import DistributionBase
_EPS = 1e-12
def _validate_params(min_val: float, m_likely: float, max_val: float, lower_p: float) -> None:
    if not (min_val < m_likely < max_val):
        raise ValueError(
            f"DoubleTriang requires min_val < m_likely < max_val, got "
            f"min_val={min_val}, m_likely={m_likely}, max_val={max_val}"
        )
    if not (0.0 <= lower_p <= 1.0):
        raise ValueError(f"lower_p must satisfy 0 <= p <= 1, got {lower_p}")


def doubletriang_generator_single(rng: np.random.Generator, params: List[float]) -> float:
    min_val = float(params[0])
    m_likely = float(params[1])
    max_val = float(params[2])
    lower_p = float(params[3])
    _validate_params(min_val, m_likely, max_val, lower_p)
    return doubletriang_ppf(float(rng.random()), min_val, m_likely, max_val, lower_p)


def doubletriang_generator_vectorized(
    rng: np.random.Generator, params: List[float], n_samples: int
) -> np.ndarray:
    min_val = float(params[0])
    m_likely = float(params[1])
    max_val = float(params[2])
    lower_p = float(params[3])
    _validate_params(min_val, m_likely, max_val, lower_p)
    u = rng.random(size=n_samples)
    return np.array(
        [doubletriang_ppf(float(q), min_val, m_likely, max_val, lower_p) for q in u],
        dtype=float,
    )


def doubletriang_generator(
    rng: np.random.Generator,
    params: List[float],
    n_samples: Optional[int] = None,
) -> Union[float, np.ndarray]:
    if n_samples is None:
        return doubletriang_generator_single(rng, params)
    return doubletriang_generator_vectorized(rng, params, n_samples)


def doubletriang_pdf(x: float, min_val: float, m_likely: float, max_val: float, lower_p: float) -> float:
    _validate_params(min_val, m_likely, max_val, lower_p)
    if x < min_val or x > max_val:
        return 0.0

    left_range = m_likely - min_val
    right_range = max_val - m_likely
    if x <= m_likely:
        return (2.0 * lower_p * (x - min_val)) / (left_range ** 2)
    return (2.0 * (1.0 - lower_p) * (max_val - x)) / (right_range ** 2)


def doubletriang_cdf(x: float, min_val: float, m_likely: float, max_val: float, lower_p: float) -> float:
    _validate_params(min_val, m_likely, max_val, lower_p)
    if x <= min_val:
        return 0.0
    if x >= max_val:
        return 1.0

    left_range = m_likely - min_val
    right_range = max_val - m_likely
    if x <= m_likely:
        return lower_p * ((x - min_val) ** 2) / (left_range ** 2)
    return 1.0 - (1.0 - lower_p) * ((max_val - x) ** 2) / (right_range ** 2)


def doubletriang_ppf(
    q_prob: float, min_val: float, m_likely: float, max_val: float, lower_p: float
) -> float:
    _validate_params(min_val, m_likely, max_val, lower_p)
    if q_prob <= 0.0:
        return min_val
    if q_prob >= 1.0:
        return max_val

    q = float(q_prob)
    if q <= lower_p and lower_p > _EPS:
        return min_val + (m_likely - min_val) * math.sqrt(q / lower_p)
    if lower_p >= 1.0 - _EPS:
        return m_likely
    return max_val - (max_val - m_likely) * math.sqrt((1.0 - q) / (1.0 - lower_p))


def doubletriang_raw_mean(min_val: float, m_likely: float, max_val: float, lower_p: float) -> float:
    _validate_params(min_val, m_likely, max_val, lower_p)
    return (lower_p * min_val + 2.0 * m_likely + (1.0 - lower_p) * max_val) / 3.0


class DoubleTriangDistribution(DistributionBase):
    """DoubleTriang theoretical distribution with truncation support."""

    def __init__(self, params: List[float], markers: dict = None, func_name: str = None):
        if len(params) >= 4:
            min_val = float(params[0])
            m_likely = float(params[1])
            max_val = float(params[2])
            lower_p = float(params[3])
        else:
            min_val, m_likely, max_val, lower_p = -10.0, 0.0, 10.0, 0.4
        _validate_params(min_val, m_likely, max_val, lower_p)

        self.min_param = min_val
        self.m_likely = m_likely
        self.max_param = max_val
        self.lower_p = lower_p
        self.support_low = min_val
        self.support_high = max_val

        self._raw_mean = doubletriang_raw_mean(min_val, m_likely, max_val, lower_p)
        self._raw_var = float("nan")
        self._raw_skew = float("nan")
        self._raw_kurt = float("nan")
        self._raw_mode = m_likely
        self._truncated_stats = None

        super().__init__([min_val, m_likely, max_val, lower_p], markers, func_name)

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
        self._compute_stats()

    def _original_cdf(self, x: float) -> float:
        return doubletriang_cdf(x, self.min_param, self.m_likely, self.max_param, self.lower_p)

    def _original_ppf(self, q_prob: float) -> float:
        return doubletriang_ppf(q_prob, self.min_param, self.m_likely, self.max_param, self.lower_p)

    def _original_pdf(self, x: float) -> float:
        return doubletriang_pdf(x, self.min_param, self.m_likely, self.max_param, self.lower_p)

    def _get_original_truncation_bounds(self):
        low, high = self.get_truncated_bounds()
        if low is None and high is None:
            return self.min_param, self.max_param
        if self.truncate_type in ["value2", "percentile2"]:
            low_orig = low - self.shift_amount if low is not None else self.min_param
            high_orig = high - self.shift_amount if high is not None else self.max_param
        else:
            low_orig = low if low is not None else self.min_param
            high_orig = high if high is not None else self.max_param
        return low_orig, high_orig

    def _compute_stats(self):
        low_orig, high_orig = self._get_original_truncation_bounds()
        lower = max(self.min_param, low_orig)
        upper = min(self.max_param, high_orig)

        if upper < lower:
            self._truncate_invalid = True
            self._truncated_stats = None
            return

        norm = self._original_cdf(upper) - self._original_cdf(lower)
        if norm <= _EPS:
            point = lower
            self._truncated_stats = {
                "mean": point,
                "variance": 0.0,
                "skewness": 0.0,
                "kurtosis": 3.0,
                "mode": min(max(self.m_likely, lower), upper),
            }
            return

        pdf_func = lambda x: self._original_pdf(x)

        mean_raw = quad(lambda x: x * pdf_func(x), lower, upper, limit=200)[0] / norm
        variance = quad(lambda x: ((x - mean_raw) ** 2) * pdf_func(x), lower, upper, limit=200)[0] / norm
        variance = max(0.0, variance)

        if variance > _EPS:
            mu3 = quad(lambda x: ((x - mean_raw) ** 3) * pdf_func(x), lower, upper, limit=200)[0] / norm
            mu4 = quad(lambda x: ((x - mean_raw) ** 4) * pdf_func(x), lower, upper, limit=200)[0] / norm
            skewness = mu3 / (variance ** 1.5)
            kurtosis = mu4 / (variance ** 2)
        else:
            skewness = 0.0
            kurtosis = 3.0

        mode_raw = min(max(self.m_likely, lower), upper)

        self._truncated_stats = {
            "mean": mean_raw,
            "variance": variance,
            "skewness": skewness,
            "kurtosis": kurtosis,
            "mode": mode_raw,
        }

    def mean(self) -> float:
        if self._truncate_invalid or self._truncated_stats is None:
            return float("nan")
        return self._truncated_stats["mean"] + self.shift_amount

    def variance(self) -> float:
        if self._truncate_invalid or self._truncated_stats is None:
            return float("nan")
        return self._truncated_stats["variance"]

    def skewness(self) -> float:
        if self._truncate_invalid or self._truncated_stats is None:
            return float("nan")
        return self._truncated_stats["skewness"]

    def kurtosis(self) -> float:
        if self._truncate_invalid or self._truncated_stats is None:
            return float("nan")
        return self._truncated_stats["kurtosis"]

    def mode(self) -> float:
        if self._truncate_invalid or self._truncated_stats is None:
            return float("nan")
        return self._truncated_stats["mode"] + self.shift_amount

    def min_val(self) -> float:
        if self._truncate_invalid:
            return float("nan")
        if self.is_truncated():
            lower, _ = self.get_effective_bounds()
            if lower is not None:
                return lower
        return self.min_param + self.shift_amount

    def max_val(self) -> float:
        if self._truncate_invalid:
            return float("nan")
        if self.is_truncated():
            _, upper = self.get_effective_bounds()
            if upper is not None:
                return upper
        return self.max_param + self.shift_amount
