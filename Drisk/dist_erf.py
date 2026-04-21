"""
Erf distribution support for Drisk.

Definition used here:
    X ~ Erf(h)
    PDF = h / sqrt(pi) * exp(-(h * x)^2)
    CDF = 0.5 * (1 + erf(h * x))
    PPF = erfinv(2q - 1) / h

Parameters:
    h: inverse scale parameter, must be > 0
"""

import math
from typing import List, Optional, Union

import numpy as np
from scipy.integrate import quad
from scipy.special import erfinv

from distribution_base import DistributionBase

_EPS = 1e-12
_SQRT_PI = math.sqrt(math.pi)
_SQRT2 = math.sqrt(2.0)


def _validate_params(h: float) -> None:
    if h <= 0.0:
        raise ValueError(f"h must be > 0, got {h}")


def erf_generator_single(rng: np.random.Generator, params: List[float]) -> float:
    h = float(params[0])
    _validate_params(h)
    sigma = 1.0 / (_SQRT2 * h)
    return float(rng.normal(0.0, sigma))


def erf_generator_vectorized(
    rng: np.random.Generator, params: List[Union[float, np.ndarray]], n_samples: int
) -> np.ndarray:
    h = params[0]
    if not isinstance(h, np.ndarray):
        h = np.full(n_samples, float(h))
    if np.any(h <= 0):
        raise ValueError("Erf requires h > 0")
    sigma = 1.0 / (np.sqrt(2.0) * h)
    return rng.normal(0.0, sigma, size=n_samples)


def erf_generator(
    rng: np.random.Generator,
    params: List[float],
    n_samples: Optional[int] = None,
) -> Union[float, np.ndarray]:
    if n_samples is None:
        return erf_generator_single(rng, params)
    return erf_generator_vectorized(rng, params, n_samples)


def erf_pdf(x: float, h: float) -> float:
    _validate_params(h)
    return (h / _SQRT_PI) * math.exp(-((h * x) ** 2))


def erf_cdf(x: float, h: float) -> float:
    _validate_params(h)
    return 0.5 * (1.0 + math.erf(h * x))


def erf_ppf(q_prob: float, h: float) -> float:
    _validate_params(h)
    if q_prob <= 0.0:
        return float("-inf")
    if q_prob >= 1.0:
        return float("inf")
    q = max(_EPS, min(1.0 - _EPS, float(q_prob)))
    return float(erfinv(2.0 * q - 1.0) / h)


def erf_raw_mean(h: float) -> float:
    _validate_params(h)
    return 0.0


class ErfDistribution(DistributionBase):
    """Erf theoretical distribution with truncation support."""

    def __init__(self, params: List[float], markers: dict = None, func_name: str = None):
        if len(params) >= 1:
            h = float(params[0])
        else:
            h = 8.5
        _validate_params(h)

        self.h = h
        self.support_low = float("-inf")
        self.support_high = float("inf")

        self._raw_mean = 0.0
        self._raw_var = 1.0 / (2.0 * h * h)
        self._raw_skew = 0.0
        self._raw_kurt = 3.0
        self._raw_mode = 0.0
        self._truncated_stats = None

        super().__init__([h], markers, func_name)

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
        return erf_cdf(x, self.h)

    def _original_ppf(self, q_prob: float) -> float:
        return erf_ppf(q_prob, self.h)

    def _original_pdf(self, x: float) -> float:
        return erf_pdf(x, self.h)

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

        norm = self._original_cdf(ub) - self._original_cdf(lb)
        if norm <= _EPS:
            self._truncate_invalid = True
            self._truncated_stats = None
            return

        pdf_func = lambda x: self._original_pdf(x)

        mean_raw = quad(lambda x: x * pdf_func(x), lb, ub, limit=200)[0] / norm
        variance = quad(lambda x: ((x - mean_raw) ** 2) * pdf_func(x), lb, ub, limit=200)[0] / norm
        variance = max(0.0, variance)

        if variance > _EPS:
            mu3 = quad(lambda x: ((x - mean_raw) ** 3) * pdf_func(x), lb, ub, limit=200)[0] / norm
            mu4 = quad(lambda x: ((x - mean_raw) ** 4) * pdf_func(x), lb, ub, limit=200)[0] / norm
            skewness = mu3 / (variance ** 1.5)
            kurtosis = mu4 / (variance ** 2)
        else:
            skewness = 0.0
            kurtosis = 3.0

        if low_orig is not None and low_orig > 0.0:
            mode_raw = low_orig
        elif high_orig is not None and high_orig < 0.0:
            mode_raw = high_orig
        else:
            mode_raw = 0.0

        self._truncated_stats = {
            "mean": mean_raw,
            "variance": variance,
            "skewness": skewness,
            "kurtosis": kurtosis,
            "mode": mode_raw,
        }

    def mean(self) -> float:
        if self._truncate_invalid:
            return float("nan")
        if self._truncated_stats is not None:
            return self._truncated_stats["mean"] + self.shift_amount
        return self.shift_amount

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
        return 0.0

    def kurtosis(self) -> float:
        if self._truncate_invalid:
            return float("nan")
        if self._truncated_stats is not None:
            return self._truncated_stats["kurtosis"]
        return 3.0

    def mode(self) -> float:
        if self._truncate_invalid:
            return float("nan")
        if self._truncated_stats is not None:
            return self._truncated_stats["mode"] + self.shift_amount
        return self.shift_amount
