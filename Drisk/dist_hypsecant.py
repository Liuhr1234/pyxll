"""
Hyperbolic secant distribution support for Drisk.
"""

import math
from typing import List, Optional, Tuple, Union

import numpy as np
from scipy.integrate import quad

from distribution_base import DistributionBase

_EPS = 1e-12
_PI = math.pi
_PI_OVER_2 = _PI / 2.0
_TWO_OVER_PI = 2.0 / _PI


def _sech(x: float) -> float:
    ax = abs(float(x))
    if ax > 40.0:
        # Avoid cosh overflow in tail integrals; sech(x) ~ 2 * exp(-|x|).
        return 2.0 * math.exp(-ax)
    return 1.0 / math.cosh(x)


def hypsecant_pdf(x: float, gamma: float, beta: float) -> float:
    z = (float(x) - float(gamma)) / float(beta)
    return (1.0 / (2.0 * float(beta))) * _sech(_PI_OVER_2 * z)


def hypsecant_cdf(x: float, gamma: float, beta: float) -> float:
    z = (float(x) - float(gamma)) / float(beta)
    return _TWO_OVER_PI * math.atan(math.exp(_PI_OVER_2 * z))


def hypsecant_ppf(q: float, gamma: float, beta: float) -> float:
    q = float(q)
    if q <= 0.0:
        return float("-inf")
    if q >= 1.0:
        return float("inf")
    if abs(q - 0.5) <= _EPS:
        return float(gamma)
    tan_val = math.tan(_PI_OVER_2 * q)
    if tan_val <= 0.0:
        return float("-inf") if q < 0.5 else float("inf")
    z = _TWO_OVER_PI * math.log(tan_val)
    return float(gamma + beta * z)


def hypsecant_raw_mean(gamma: float, beta: float) -> float:
    return float(gamma)


def hypsecant_raw_var(gamma: float, beta: float) -> float:
    return float(beta) ** 2


def hypsecant_raw_skewness(gamma: float, beta: float) -> float:
    return 0.0


def hypsecant_raw_kurtosis(gamma: float, beta: float) -> float:
    return 5.0


def hypsecant_raw_mode(gamma: float, beta: float) -> float:
    return float(gamma)


def hypsecant_generator_single(
    rng: np.random.Generator,
    params: List[float],
) -> float:
    gamma = float(params[0])
    beta = float(params[1])
    q = float(rng.uniform(_EPS, 1.0 - _EPS))
    return hypsecant_ppf(q, gamma, beta)


def hypsecant_generator_vectorized(
    rng: np.random.Generator, params: List[Union[float, np.ndarray]], n_samples: int
) -> np.ndarray:
    gamma = params[0]
    beta = params[1]

    if not isinstance(gamma, np.ndarray):
        gamma = np.full(n_samples, float(gamma))
    if not isinstance(beta, np.ndarray):
        beta = np.full(n_samples, float(beta))

    if np.any(beta <= 0):
        raise ValueError("HypSecant requires beta > 0")

    q = rng.uniform(_EPS, 1.0 - _EPS, size=n_samples)
    # 使用向量化计算
    tan_val = np.tan(np.pi * (q - 0.5))
    # 处理 tan_val <= 0 的情况，使用 np.where
    z = np.where(tan_val > 0, (2.0 / np.pi) * np.log(tan_val), -np.inf)
    samples = gamma + beta * z
    # 边界处理
    samples = np.where(q <= 0.0, -np.inf, samples)
    samples = np.where(q >= 1.0, np.inf, samples)
    return samples


class HypSecantDistribution(DistributionBase):
    """HypSecant theoretical distribution with truncation support."""

    def __init__(self, params: List[float], markers: dict = None, func_name: str = None):
        markers = markers or {}
        gamma = float(params[0]) if len(params) >= 1 else 0.0
        beta = float(params[1]) if len(params) >= 2 else 10.0
        if beta <= 0:
            raise ValueError("HypSecant requires beta > 0")

        self.gamma = gamma
        self.beta = beta
        self.support_low = float("-inf")
        self.support_high = float("inf")
        self._raw_mean = hypsecant_raw_mean(gamma, beta)
        self._raw_var = hypsecant_raw_var(gamma, beta)
        self._raw_skew = hypsecant_raw_skewness(gamma, beta)
        self._raw_kurt = hypsecant_raw_kurtosis(gamma, beta)
        self._raw_mode = hypsecant_raw_mode(gamma, beta)
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

    def _original_pdf(self, x: float) -> float:
        return hypsecant_pdf(x, self.gamma, self.beta)

    def _original_cdf(self, x: float) -> float:
        return hypsecant_cdf(x, self.gamma, self.beta)

    def _original_ppf(self, q_prob: float) -> float:
        return hypsecant_ppf(q_prob, self.gamma, self.beta)

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
        if upper <= lower:
            return 0.0

        def integrand(x: float) -> float:
            return (x ** power) * hypsecant_pdf(x, self.gamma, self.beta)

        result, _ = quad(integrand, lower, upper, limit=200)
        return float(result)

    def _compute_truncated_stats(self) -> None:
        if not self.is_truncated():
            self._truncated_stats = None
            return

        low_orig, high_orig = self._get_original_truncation_bounds()
        lb = float("-inf") if low_orig is None else low_orig
        ub = float("inf") if high_orig is None else high_orig
        if ub < lb:
            self._truncated_stats = None
            return

        mass = max(0.0, hypsecant_cdf(ub, self.gamma, self.beta) - hypsecant_cdf(lb, self.gamma, self.beta))
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

        if self.gamma < lb:
            mode_raw = lb
        elif self.gamma > ub:
            mode_raw = ub
        else:
            mode_raw = self.gamma

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
        return float("-inf")

    def max_val(self) -> float:
        if self._truncate_invalid:
            return float("nan")
        _, upper = self.get_effective_bounds()
        if upper is not None:
            return upper
        return float("inf")
