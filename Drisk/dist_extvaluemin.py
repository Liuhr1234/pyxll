"""
Extreme value type I (minimum) / Gumbel distribution support for Drisk.

Definition used here:
    X ~ ExtvalueMin(A, B)
    PDF = (1 / B) * exp(z - exp(z))
    CDF = 1 - exp(-exp(z))
    PPF = A + B * ln(-ln(1 - q))
    where z = (x - A) / B and B > 0.
"""

import math
from typing import List, Optional, Union

import numpy as np
import scipy.stats as sps

from distribution_base import DistributionBase

_EULER_GAMMA = 0.5772156649015329
_GUMBEL_SKEW = -1.1395470994046488
_GUMBEL_KURT = 5.4
_EPS = 1e-12


def _validate_params(a: float, b: float) -> None:
    if b <= 0.0:
        raise ValueError(f"B must be > 0, got {b}")


def extvaluemin_generator_single(rng: np.random.Generator, params: List[float]) -> float:
    a = float(params[0])
    b = float(params[1])
    _validate_params(a, b)
    return float(rng.gumbel(loc=-a, scale=b) * -1.0)


def extvaluemin_generator_vectorized(
    rng: np.random.Generator, params: List[float], n_samples: int
) -> np.ndarray:
    a = float(params[0])
    b = float(params[1])
    _validate_params(a, b)
    return (-rng.gumbel(loc=-a, scale=b, size=n_samples)).astype(float)


def extvaluemin_generator(
    rng: np.random.Generator,
    params: List[float],
    n_samples: Optional[int] = None,
) -> Union[float, np.ndarray]:
    if n_samples is None:
        return extvaluemin_generator_single(rng, params)
    return extvaluemin_generator_vectorized(rng, params, n_samples)


def extvaluemin_pdf(x: float, a: float, b: float) -> float:
    _validate_params(a, b)
    return float(sps.gumbel_l.pdf(x, loc=a, scale=b))


def extvaluemin_cdf(x: float, a: float, b: float) -> float:
    _validate_params(a, b)
    return float(sps.gumbel_l.cdf(x, loc=a, scale=b))


def extvaluemin_ppf(q_prob: float, a: float, b: float) -> float:
    _validate_params(a, b)
    if q_prob <= 0.0:
        return float("-inf")
    if q_prob >= 1.0:
        return float("inf")
    q = max(_EPS, min(1.0 - _EPS, float(q_prob)))
    return float(sps.gumbel_l.ppf(q, loc=a, scale=b))


def extvaluemin_raw_mean(a: float, b: float) -> float:
    _validate_params(a, b)
    return float(a - b * _EULER_GAMMA)


class ExtvalueMinDistribution(DistributionBase):
    """Extreme value type I (minimum) distribution with truncation support."""

    def __init__(self, params: List[float], markers: dict = None, func_name: str = None):
        if len(params) >= 2:
            a = float(params[0])
            b = float(params[1])
        else:
            a, b = 0.0, 10.0
        _validate_params(a, b)

        self.a = a
        self.b = b
        self._dist = sps.gumbel_l(loc=a, scale=b)

        self.support_low = float("-inf")
        self.support_high = float("inf")
        self._raw_mean = a - b * _EULER_GAMMA
        self._raw_var = (math.pi ** 2 / 6.0) * (b ** 2)
        self._raw_skew = _GUMBEL_SKEW
        self._raw_kurt = _GUMBEL_KURT
        self._raw_mode = a
        self._truncated_stats = None

        super().__init__([a, b], markers, func_name)

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
        return extvaluemin_cdf(x, self.a, self.b)

    def _original_ppf(self, q_prob: float) -> float:
        return extvaluemin_ppf(q_prob, self.a, self.b)

    def _original_pdf(self, x: float) -> float:
        return extvaluemin_pdf(x, self.a, self.b)

    def _get_original_truncation_bounds(self):
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

    def _compute_truncated_stats(self):
        if not self.is_truncated():
            self._truncated_stats = None
            return

        low_orig, high_orig = self._get_original_truncation_bounds()
        lb = -float("inf") if low_orig is None else low_orig
        ub = float("inf") if high_orig is None else high_orig

        try:
            m1 = self._dist.expect(lambda x: x, lb=lb, ub=ub, conditional=True)
            m2 = self._dist.expect(lambda x: x**2, lb=lb, ub=ub, conditional=True)
            m3 = self._dist.expect(lambda x: x**3, lb=lb, ub=ub, conditional=True)
            m4 = self._dist.expect(lambda x: x**4, lb=lb, ub=ub, conditional=True)
        except Exception:
            self._truncated_stats = None
            return

        variance = max(0.0, m2 - m1**2)
        if variance > _EPS:
            skewness = (m3 - 3 * m1 * variance - m1**3) / (variance ** 1.5)
            kurtosis = (m4 - 4 * m1 * m3 + 6 * m1 * m1 * m2 - 3 * (m1 ** 4)) / (variance ** 2)
        else:
            skewness = 0.0
            kurtosis = 3.0

        mode_raw = self.a
        if low_orig is not None and mode_raw < low_orig:
            mode_raw = low_orig
        if high_orig is not None and mode_raw > high_orig:
            mode_raw = high_orig

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
        return self.a + self.shift_amount
