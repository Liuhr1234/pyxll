"""Burr12 distribution support for Drisk."""

import math
from typing import List, Optional, Tuple, Union

import numpy as np
import scipy.special as sp
from scipy.integrate import quad
from scipy.stats import burr12 as scipy_burr12

from distribution_base import DistributionBase

_EPS = 1e-12


def _validate_params(gamma: float, beta: float, alpha1: float, alpha2: float) -> None:
    if beta <= 0.0:
        raise ValueError("Burr12 requires beta > 0")
    if alpha1 <= 0.0:
        raise ValueError("Burr12 requires alpha1 > 0")
    if alpha2 <= 0.0:
        raise ValueError("Burr12 requires alpha2 > 0")


def _log1p_exp(value: float) -> float:
    if value > 50.0:
        return value
    if value < -50.0:
        return math.exp(value)
    return math.log1p(math.exp(value))


def _burr12_q(alpha1: float, alpha2: float, n: int) -> float:
    if alpha1 * alpha2 <= float(n):
        return float("inf")
    return float(alpha2 * sp.beta(1.0 + n / alpha1, alpha2 - n / alpha1))


def _create_dist(gamma: float, beta: float, alpha1: float, alpha2: float):
    _validate_params(gamma, beta, alpha1, alpha2)
    # SciPy burr12(c, d, loc, scale) uses the same loc/scale convention:
    # z = (x - gamma) / beta, x >= gamma.
    return scipy_burr12(c=float(alpha1), d=float(alpha2), loc=float(gamma), scale=float(beta))


def burr12_pdf(x: float, gamma: float, beta: float, alpha1: float, alpha2: float) -> float:
    _validate_params(gamma, beta, alpha1, alpha2)
    x = float(x)
    if x < gamma:
        return 0.0
    if x == gamma:
        if alpha1 < 1.0:
            return float("inf")
        if abs(alpha1 - 1.0) <= 1e-12:
            return float(alpha2 / beta)
        return 0.0

    z = (x - gamma) / beta
    log_z = math.log(z)
    log_denom = (alpha2 + 1.0) * _log1p_exp(alpha1 * log_z)
    log_pdf = (
        math.log(alpha1)
        + math.log(alpha2)
        - math.log(beta)
        + (alpha1 - 1.0) * log_z
        - log_denom
    )
    return math.exp(log_pdf)


def burr12_cdf(x: float, gamma: float, beta: float, alpha1: float, alpha2: float) -> float:
    _validate_params(gamma, beta, alpha1, alpha2)
    x = float(x)
    if x <= gamma:
        return 0.0

    z = (x - gamma) / beta
    log_survival = -alpha2 * _log1p_exp(alpha1 * math.log(z))
    if log_survival < -745.0:
        return 1.0
    return max(0.0, min(1.0, -math.expm1(log_survival)))


def burr12_ppf(q_prob: float, gamma: float, beta: float, alpha1: float, alpha2: float) -> float:
    _validate_params(gamma, beta, alpha1, alpha2)
    q_prob = float(q_prob)
    if q_prob <= 0.0:
        return float(gamma)
    if q_prob >= 1.0:
        return float("inf")

    q = max(_EPS, min(1.0 - _EPS, q_prob))
    inner = math.expm1(-math.log1p(-q) / alpha2)
    if inner <= 0.0:
        return float(gamma)
    return float(gamma + beta * (inner ** (1.0 / alpha1)))


def burr12_raw_mean(gamma: float, beta: float, alpha1: float, alpha2: float) -> float:
    _validate_params(gamma, beta, alpha1, alpha2)
    q1 = _burr12_q(alpha1, alpha2, 1)
    if not math.isfinite(q1):
        return float("inf")
    return float(gamma + beta * q1)


def burr12_raw_var(gamma: float, beta: float, alpha1: float, alpha2: float) -> float:
    _validate_params(gamma, beta, alpha1, alpha2)
    q1 = _burr12_q(alpha1, alpha2, 1)
    q2 = _burr12_q(alpha1, alpha2, 2)
    if not math.isfinite(q1) or not math.isfinite(q2):
        return float("inf")
    variance = (beta ** 2) * (q2 - q1 * q1)
    return max(0.0, variance)


def burr12_raw_skew(gamma: float, beta: float, alpha1: float, alpha2: float) -> float:
    _validate_params(gamma, beta, alpha1, alpha2)
    q1 = _burr12_q(alpha1, alpha2, 1)
    q2 = _burr12_q(alpha1, alpha2, 2)
    q3 = _burr12_q(alpha1, alpha2, 3)
    if not math.isfinite(q1) or not math.isfinite(q2) or not math.isfinite(q3):
        return float("nan")

    variance_base = q2 - q1 * q1
    if variance_base <= _EPS:
        return 0.0

    numerator = q3 - 3.0 * q1 * q2 + 2.0 * (q1 ** 3)
    return float(numerator / (variance_base ** 1.5))


def burr12_raw_kurt(gamma: float, beta: float, alpha1: float, alpha2: float) -> float:
    _validate_params(gamma, beta, alpha1, alpha2)
    q1 = _burr12_q(alpha1, alpha2, 1)
    q2 = _burr12_q(alpha1, alpha2, 2)
    q3 = _burr12_q(alpha1, alpha2, 3)
    q4 = _burr12_q(alpha1, alpha2, 4)
    if not math.isfinite(q1) or not math.isfinite(q2) or not math.isfinite(q3) or not math.isfinite(q4):
        return float("nan")

    variance_base = q2 - q1 * q1
    if variance_base <= _EPS:
        return 3.0

    fourth_central = q4 - 4.0 * q1 * q3 + 6.0 * (q1 ** 2) * q2 - 3.0 * (q1 ** 4)
    return float(fourth_central / (variance_base ** 2))


def burr12_raw_mode(gamma: float, beta: float, alpha1: float, alpha2: float) -> float:
    _validate_params(gamma, beta, alpha1, alpha2)
    if alpha1 <= 1.0:
        return float(gamma)
    return float(gamma + beta * (((alpha1 - 1.0) / (alpha1 * alpha2 + 1.0)) ** (1.0 / alpha1)))


def burr12_generator_single(rng: np.random.Generator, params: List[float]) -> float:
    gamma = float(params[0])
    beta = float(params[1])
    alpha1 = float(params[2])
    alpha2 = float(params[3])
    q = float(rng.uniform(_EPS, 1.0 - _EPS))
    return burr12_ppf(q, gamma, beta, alpha1, alpha2)


def burr12_generator_vectorized(
    rng: np.random.Generator,
    params: List[Union[float, np.ndarray]],
    n_samples: int,
) -> np.ndarray:
    gamma = params[0]
    beta = params[1]
    alpha1 = params[2]
    alpha2 = params[3]

    # 广播
    if not isinstance(gamma, np.ndarray):
        gamma = np.full(n_samples, float(gamma))
    if not isinstance(beta, np.ndarray):
        beta = np.full(n_samples, float(beta))
    if not isinstance(alpha1, np.ndarray):
        alpha1 = np.full(n_samples, float(alpha1))
    if not isinstance(alpha2, np.ndarray):
        alpha2 = np.full(n_samples, float(alpha2))

    # 验证
    valid = (beta > 0) & (alpha1 > 0) & (alpha2 > 0)
    if not np.all(valid):
        raise ValueError("Burr12 parameters invalid")

    q = rng.uniform(_EPS, 1.0 - _EPS, size=n_samples)
    inner = np.expm1(-np.log1p(-q) / alpha2)
    inner = np.clip(inner, 0.0, np.inf)
    samples = gamma + beta * np.power(inner, 1.0 / alpha1)
    return samples.astype(float)

class Burr12Distribution(DistributionBase):
    """Burr12 theoretical distribution with truncation support."""

    def __init__(self, params: List[float], markers: dict = None, func_name: str = None):
        markers = markers or {}
        gamma = float(params[0]) if len(params) >= 1 else 0.0
        beta = float(params[1]) if len(params) >= 2 else 1.0
        alpha1 = float(params[2]) if len(params) >= 3 else 2.0
        alpha2 = float(params[3]) if len(params) >= 4 else 2.0

        _validate_params(gamma, beta, alpha1, alpha2)

        self.gamma = gamma
        self.beta = beta
        self.alpha1 = alpha1
        self.alpha2 = alpha2
        self._dist = _create_dist(gamma, beta, alpha1, alpha2)
        self.support_low = gamma
        self.support_high = float("inf")
        self._raw_mean = burr12_raw_mean(gamma, beta, alpha1, alpha2)
        self._raw_var = burr12_raw_var(gamma, beta, alpha1, alpha2)
        self._raw_skew = burr12_raw_skew(gamma, beta, alpha1, alpha2)
        self._raw_kurt = burr12_raw_kurt(gamma, beta, alpha1, alpha2)
        self._raw_mode = burr12_raw_mode(gamma, beta, alpha1, alpha2)
        self._truncated_stats = None

        super().__init__([gamma, beta, alpha1, alpha2], markers, func_name)

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
        return burr12_pdf(x, self.gamma, self.beta, self.alpha1, self.alpha2)

    def _original_cdf(self, x: float) -> float:
        return burr12_cdf(x, self.gamma, self.beta, self.alpha1, self.alpha2)

    def _original_ppf(self, q_prob: float) -> float:
        return burr12_ppf(q_prob, self.gamma, self.beta, self.alpha1, self.alpha2)

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
        lb = self.gamma if low_orig is None else max(self.gamma, low_orig)
        ub = float("inf") if high_orig is None else high_orig
        if ub < lb:
            self._truncated_stats = None
            return

        mass = max(0.0, self._original_cdf(ub) - self._original_cdf(lb))
        if mass <= _EPS:
            self._truncated_stats = None
            return

        integral_lb = lb
        if integral_lb <= self.gamma:
            integral_lb = math.nextafter(self.gamma, float("inf"))

        try:
            m1 = float(self._dist.expect(lambda x: x, lb=lb, ub=ub, conditional=True))
            m2 = float(self._dist.expect(lambda x: x * x, lb=lb, ub=ub, conditional=True))
            m3_raw = float(self._dist.expect(lambda x: x ** 3, lb=lb, ub=ub, conditional=True))
            m4_raw = float(self._dist.expect(lambda x: x ** 4, lb=lb, ub=ub, conditional=True))
        except Exception:
            finite_ub = ub
            if math.isinf(finite_ub):
                finite_ub = self._original_ppf(1.0 - 1e-8)

            m1 = self._moment_integral(integral_lb, finite_ub, 1) / mass
            m2 = self._moment_integral(integral_lb, finite_ub, 2) / mass
            m3_raw = self._moment_integral(integral_lb, finite_ub, 3) / mass
            m4_raw = self._moment_integral(integral_lb, finite_ub, 4) / mass

        variance = max(0.0, m2 - m1 * m1) if math.isfinite(m2) and math.isfinite(m1) else float("inf")

        if math.isfinite(variance) and variance > _EPS:
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

        mode_raw = self._raw_mode
        if math.isfinite(mode_raw):
            if mode_raw < lb:
                mode_raw = lb
            if math.isfinite(ub) and mode_raw > ub:
                mode_raw = ub

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
        return float(self.gamma + self.shift_amount)

    def max_val(self) -> float:
        if self._truncate_invalid:
            return float("nan")
        _, upper = self.get_effective_bounds()
        if upper is not None:
            return upper
        return float("inf")
