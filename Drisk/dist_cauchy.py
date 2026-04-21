"""
Cauchy distribution support for Drisk.

Definition used here:
    X ~ Cauchy(gamma, beta)
    PDF = 1 / (pi * beta * (1 + ((x - gamma) / beta)^2))
    CDF = 0.5 + arctan((x - gamma) / beta) / pi
    PPF = gamma + beta * tan(pi * (q - 0.5))

Parameters:
    gamma: location parameter
    beta: scale parameter, must be > 0
"""

import math
from typing import List, Optional, Union

import numpy as np
from scipy.integrate import quad

from distribution_base import DistributionBase

_EPS = 1e-12


def _validate_params(gamma: float, beta: float) -> None:
    if beta <= 0.0:
        raise ValueError(f"beta must be > 0, got {beta}")


def cauchy_generator_single(rng: np.random.Generator, params: List[float]) -> float:
    gamma = float(params[0])
    beta = float(params[1])
    _validate_params(gamma, beta)
    return float(gamma + beta * rng.standard_cauchy())


def cauchy_generator_vectorized(
    rng: np.random.Generator, params: List[Union[float, np.ndarray]], n_samples: int
) -> np.ndarray:
    gamma = params[0]
    beta = params[1]

    if not isinstance(gamma, np.ndarray):
        gamma = np.full(n_samples, float(gamma))
    if not isinstance(beta, np.ndarray):
        beta = np.full(n_samples, float(beta))

    if np.any(beta <= 0):
        raise ValueError("Cauchy requires beta > 0")

    return gamma + beta * rng.standard_cauchy(size=n_samples)


def cauchy_generator(
    rng: np.random.Generator,
    params: List[float],
    n_samples: Optional[int] = None,
) -> Union[float, np.ndarray]:
    if n_samples is None:
        return cauchy_generator_single(rng, params)
    return cauchy_generator_vectorized(rng, params, n_samples)


def cauchy_pdf(x: float, gamma: float, beta: float) -> float:
    _validate_params(gamma, beta)
    z = (x - gamma) / beta
    return 1.0 / (math.pi * beta * (1.0 + z * z))


def cauchy_cdf(x: float, gamma: float, beta: float) -> float:
    _validate_params(gamma, beta)
    z = (x - gamma) / beta
    return 0.5 + math.atan(z) / math.pi


def cauchy_ppf(q_prob: float, gamma: float, beta: float) -> float:
    _validate_params(gamma, beta)
    if q_prob <= 0.0:
        return float("-inf")
    if q_prob >= 1.0:
        return float("inf")
    return gamma + beta * math.tan(math.pi * (q_prob - 0.5))


class CauchyDistribution(DistributionBase):
    """Cauchy theoretical distribution with truncation support."""

    def __init__(self, params: List[float], markers: dict = None, func_name: str = None):
        if len(params) >= 2:
            gamma = float(params[0])
            beta = float(params[1])
        else:
            gamma, beta = 0.0, 10.0
        _validate_params(gamma, beta)

        self.gamma = gamma
        self.beta = beta
        self.support_low = float("-inf")
        self.support_high = float("inf")

        self._raw_mean = float("nan")
        self._raw_var = float("nan")
        self._raw_skew = float("nan")
        self._raw_kurt = float("nan")
        self._raw_mode = gamma
        self._truncated_stats = None

        super().__init__([gamma, beta], markers, func_name)

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
        return cauchy_cdf(x, self.gamma, self.beta)

    def _original_ppf(self, q_prob: float) -> float:
        return cauchy_ppf(q_prob, self.gamma, self.beta)

    def _original_pdf(self, x: float) -> float:
        return cauchy_pdf(x, self.gamma, self.beta)

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
        if (
            low_orig is None or high_orig is None
            or math.isinf(low_orig) or math.isinf(high_orig)
            or high_orig <= low_orig
        ):
            self._truncated_stats = None
            return

        norm = self._original_cdf(high_orig) - self._original_cdf(low_orig)
        if norm <= _EPS:
            self._truncate_invalid = True
            self._truncated_stats = None
            return

        def pdf_func(x: float) -> float:
            return self._original_pdf(x)

        mean_raw = quad(lambda x: x * pdf_func(x), low_orig, high_orig, limit=200)[0] / norm
        variance = quad(lambda x: ((x - mean_raw) ** 2) * pdf_func(x), low_orig, high_orig, limit=200)[0] / norm
        variance = max(0.0, variance)

        if variance > _EPS:
            mu3 = quad(lambda x: ((x - mean_raw) ** 3) * pdf_func(x), low_orig, high_orig, limit=200)[0] / norm
            mu4 = quad(lambda x: ((x - mean_raw) ** 4) * pdf_func(x), low_orig, high_orig, limit=200)[0] / norm
            skewness = mu3 / (variance ** 1.5)
            kurtosis = mu4 / (variance ** 2)
        else:
            skewness = 0.0
            kurtosis = 3.0

        if self.gamma < low_orig:
            mode_raw = low_orig
        elif self.gamma > high_orig:
            mode_raw = high_orig
        else:
            mode_raw = self.gamma

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
        if self._truncated_stats is None:
            return float("nan")
        return self._truncated_stats["mean"] + self.shift_amount

    def variance(self) -> float:
        if self._truncate_invalid:
            return float("nan")
        if self._truncated_stats is None:
            return float("nan")
        return self._truncated_stats["variance"]

    def skewness(self) -> float:
        if self._truncate_invalid:
            return float("nan")
        if self._truncated_stats is None:
            return float("nan")
        return self._truncated_stats["skewness"]

    def kurtosis(self) -> float:
        if self._truncate_invalid:
            return float("nan")
        if self._truncated_stats is None:
            return float("nan")
        return self._truncated_stats["kurtosis"]

    def mode(self) -> float:
        if self._truncate_invalid:
            return float("nan")
        if self._truncated_stats is not None:
            return self._truncated_stats["mode"] + self.shift_amount
        return self.gamma + self.shift_amount
