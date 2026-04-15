"""
Dagum distribution support for Drisk.

Definition used here:
    X ~ Dagum(gamma, beta, alpha1, alpha2)
    support = [gamma, +inf)

    For x > gamma and z = (x - gamma) / beta:
        PDF = (alpha1 * alpha2 / beta) * z^(alpha1 * alpha2 - 1) / (1 + z^alpha1)^(alpha2 + 1)
        CDF = (1 + z^(-alpha1))^(-alpha2)
        PPF = gamma + beta * (q^(-1/alpha2) - 1)^(-1/alpha1)

Parameters:
    gamma: location parameter
    beta: scale parameter, must be > 0
    alpha1: shape parameter, must be > 0
    alpha2: shape parameter, must be > 0
"""

import math
from typing import List, Optional, Union

import numpy as np
import scipy.special as sp
from scipy.integrate import quad

from distribution_base import DistributionBase

_EPS = 1e-12


def _validate_params(gamma: float, beta: float, alpha1: float, alpha2: float) -> None:
    if beta <= 0.0:
        raise ValueError(f"beta must be > 0, got {beta}")
    if alpha1 <= 0.0:
        raise ValueError(f"alpha1 must be > 0, got {alpha1}")
    if alpha2 <= 0.0:
        raise ValueError(f"alpha2 must be > 0, got {alpha2}")


def _log1p_exp(value: float) -> float:
    if value > 50.0:
        return value
    if value < -50.0:
        return math.exp(value)
    return math.log1p(math.exp(value))


def _dagum_q(alpha1: float, alpha2: float, n: int) -> float:
    if alpha1 <= float(n):
        return float("nan")
    return float(alpha2 * sp.beta(alpha2 + n / alpha1, 1.0 - n / alpha1))


def dagum_generator_single(rng: np.random.Generator, params: List[float]) -> float:
    gamma = float(params[0])
    beta = float(params[1])
    alpha1 = float(params[2])
    alpha2 = float(params[3])
    _validate_params(gamma, beta, alpha1, alpha2)
    u = float(rng.random())
    return dagum_ppf(u, gamma, beta, alpha1, alpha2)


def dagum_generator_vectorized(
    rng: np.random.Generator, params: List[float], n_samples: int
) -> np.ndarray:
    gamma = float(params[0])
    beta = float(params[1])
    alpha1 = float(params[2])
    alpha2 = float(params[3])
    _validate_params(gamma, beta, alpha1, alpha2)
    u = rng.random(size=n_samples)
    return np.array(
        [dagum_ppf(float(q), gamma, beta, alpha1, alpha2) for q in u],
        dtype=float,
    )


def dagum_generator(
    rng: np.random.Generator,
    params: List[float],
    n_samples: Optional[int] = None,
) -> Union[float, np.ndarray]:
    if n_samples is None:
        return dagum_generator_single(rng, params)
    return dagum_generator_vectorized(rng, params, n_samples)


def dagum_pdf(x: float, gamma: float, beta: float, alpha1: float, alpha2: float) -> float:
    _validate_params(gamma, beta, alpha1, alpha2)
    if x <= gamma:
        return 0.0

    z = (x - gamma) / beta
    log_z = math.log(z)
    log_denom = (alpha2 + 1.0) * _log1p_exp(alpha1 * log_z)
    log_pdf = (
        math.log(alpha1)
        + math.log(alpha2)
        - math.log(beta)
        + (alpha1 * alpha2 - 1.0) * log_z
        - log_denom
    )
    return math.exp(log_pdf)


def dagum_cdf(x: float, gamma: float, beta: float, alpha1: float, alpha2: float) -> float:
    _validate_params(gamma, beta, alpha1, alpha2)
    if x <= gamma:
        return 0.0

    z = (x - gamma) / beta
    log_term = -alpha1 * math.log(z)
    log_cdf = -alpha2 * _log1p_exp(log_term)
    if log_cdf < -745.0:
        return 0.0
    return max(0.0, min(1.0, math.exp(log_cdf)))


def dagum_ppf(q_prob: float, gamma: float, beta: float, alpha1: float, alpha2: float) -> float:
    _validate_params(gamma, beta, alpha1, alpha2)
    if q_prob <= 0.0:
        return gamma
    if q_prob >= 1.0:
        return float("inf")

    q = max(_EPS, min(1.0 - _EPS, float(q_prob)))
    term = math.expm1(-math.log(q) / alpha2)
    if term <= 0.0:
        return gamma
    return gamma + beta * (term ** (-1.0 / alpha1))


def dagum_raw_mean(gamma: float, beta: float, alpha1: float, alpha2: float) -> float:
    _validate_params(gamma, beta, alpha1, alpha2)
    if alpha1 <= 1.0:
        return float("nan")
    return gamma + beta * _dagum_q(alpha1, alpha2, 1)


def dagum_raw_var(gamma: float, beta: float, alpha1: float, alpha2: float) -> float:
    _validate_params(gamma, beta, alpha1, alpha2)
    if alpha1 <= 2.0:
        return float("nan")
    q1 = _dagum_q(alpha1, alpha2, 1)
    q2 = _dagum_q(alpha1, alpha2, 2)
    variance = (beta ** 2) * (q2 - q1 * q1)
    return max(0.0, variance)


def dagum_raw_skew(gamma: float, beta: float, alpha1: float, alpha2: float) -> float:
    _validate_params(gamma, beta, alpha1, alpha2)
    if alpha1 <= 3.0:
        return float("nan")
    q1 = _dagum_q(alpha1, alpha2, 1)
    q2 = _dagum_q(alpha1, alpha2, 2)
    q3 = _dagum_q(alpha1, alpha2, 3)
    variance_base = q2 - q1 * q1
    if variance_base <= _EPS:
        return 0.0
    numerator = q3 - 3.0 * q1 * q2 + 2.0 * (q1 ** 3)
    return numerator / (variance_base ** 1.5)


def dagum_raw_kurt(gamma: float, beta: float, alpha1: float, alpha2: float) -> float:
    _validate_params(gamma, beta, alpha1, alpha2)
    if alpha1 <= 4.0:
        return float("nan")
    q1 = _dagum_q(alpha1, alpha2, 1)
    q2 = _dagum_q(alpha1, alpha2, 2)
    q3 = _dagum_q(alpha1, alpha2, 3)
    q4 = _dagum_q(alpha1, alpha2, 4)
    variance_base = q2 - q1 * q1
    if variance_base <= _EPS:
        return 3.0
    fourth_central = q4 - 4.0 * q1 * q3 + 6.0 * (q1 ** 2) * q2 - 3.0 * (q1 ** 4)
    return fourth_central / (variance_base ** 2)


def dagum_raw_mode(gamma: float, beta: float, alpha1: float, alpha2: float) -> float:
    _validate_params(gamma, beta, alpha1, alpha2)
    if alpha1 * alpha2 <= 1.0:
        return gamma
    return gamma + beta * (((alpha1 * alpha2 - 1.0) / (alpha1 + 1.0)) ** (1.0 / alpha1))


class DagumDistribution(DistributionBase):
    """Dagum theoretical distribution with truncation support."""

    def __init__(self, params: List[float], markers: dict = None, func_name: str = None):
        if len(params) >= 4:
            gamma = float(params[0])
            beta = float(params[1])
            alpha1 = float(params[2])
            alpha2 = float(params[3])
        else:
            gamma, beta, alpha1, alpha2 = 100.0, 10000.0, 1.0, 1.0
        _validate_params(gamma, beta, alpha1, alpha2)

        self.gamma = gamma
        self.beta = beta
        self.alpha1 = alpha1
        self.alpha2 = alpha2
        self.support_low = gamma
        self.support_high = float("inf")

        self._raw_mean = dagum_raw_mean(gamma, beta, alpha1, alpha2)
        self._raw_var = dagum_raw_var(gamma, beta, alpha1, alpha2)
        self._raw_skew = dagum_raw_skew(gamma, beta, alpha1, alpha2)
        self._raw_kurt = dagum_raw_kurt(gamma, beta, alpha1, alpha2)
        self._raw_mode = dagum_raw_mode(gamma, beta, alpha1, alpha2)
        self._truncated_stats = None
        self._untruncated_numeric_stats = None

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

    def _original_cdf(self, x: float) -> float:
        return dagum_cdf(x, self.gamma, self.beta, self.alpha1, self.alpha2)

    def _original_ppf(self, q_prob: float) -> float:
        return dagum_ppf(q_prob, self.gamma, self.beta, self.alpha1, self.alpha2)

    def _original_pdf(self, x: float) -> float:
        return dagum_pdf(x, self.gamma, self.beta, self.alpha1, self.alpha2)

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

    def _integrate(self, func, lower: float, upper: float) -> float:
        return quad(func, lower, upper, limit=200)[0]

    def _compute_truncated_stats(self):
        if not self.is_truncated():
            self._truncated_stats = None
            return

        low_orig, high_orig = self._get_original_truncation_bounds()
        lower = self.support_low if low_orig is None else max(self.support_low, low_orig)
        upper = self.support_high if high_orig is None else high_orig

        if upper <= lower:
            self._truncate_invalid = True
            self._truncated_stats = None
            return

        norm = self._original_cdf(upper) - self._original_cdf(lower)
        if norm <= _EPS:
            self._truncate_invalid = True
            self._truncated_stats = None
            return

        pdf_func = lambda x: self._original_pdf(x)
        mean_raw = self._integrate(lambda x: x * pdf_func(x), lower, upper) / norm
        if not math.isfinite(mean_raw):
            mean_raw = float("nan")

        variance = self._integrate(
            lambda x: ((x - mean_raw) ** 2) * pdf_func(x), lower, upper
        ) / norm if math.isfinite(mean_raw) else float("nan")
        if math.isfinite(variance):
            variance = max(0.0, variance)
        else:
            variance = float("nan")

        if math.isfinite(variance) and variance > _EPS:
            mu3 = self._integrate(lambda x: ((x - mean_raw) ** 3) * pdf_func(x), lower, upper) / norm
            mu4 = self._integrate(lambda x: ((x - mean_raw) ** 4) * pdf_func(x), lower, upper) / norm
            skewness = mu3 / (variance ** 1.5)
            kurtosis = mu4 / (variance ** 2)
        elif math.isfinite(variance):
            skewness = 0.0
            kurtosis = 3.0
        else:
            skewness = float("nan")
            kurtosis = float("nan")

        mode_raw = self._raw_mode
        if math.isfinite(mode_raw):
            if mode_raw < lower:
                mode_raw = lower
            if mode_raw > upper:
                mode_raw = upper

        self._truncated_stats = {
            "mean": mean_raw,
            "variance": variance,
            "skewness": skewness,
            "kurtosis": kurtosis,
            "mode": mode_raw,
        }

    def _compute_untruncated_numeric_stats(self):
        if self._untruncated_numeric_stats is not None:
            return self._untruncated_numeric_stats

        lower = self.gamma
        upper = self.gamma + 100.0 * self.beta
        pdf_func = lambda x: self._original_pdf(x)

        mean_raw = self._integrate(lambda x: x * pdf_func(x), lower, upper)
        variance = self._integrate(lambda x: ((x - mean_raw) ** 2) * pdf_func(x), lower, upper)
        variance = max(0.0, variance) if math.isfinite(variance) else float("nan")

        if math.isfinite(variance) and variance > _EPS:
            mu3 = self._integrate(lambda x: ((x - mean_raw) ** 3) * pdf_func(x), lower, upper)
            mu4 = self._integrate(lambda x: ((x - mean_raw) ** 4) * pdf_func(x), lower, upper)
            skewness = mu3 / (variance ** 1.5)
            kurtosis = mu4 / (variance ** 2)
        elif math.isfinite(variance):
            skewness = 0.0
            kurtosis = 3.0
        else:
            skewness = float("nan")
            kurtosis = float("nan")

        self._untruncated_numeric_stats = {
            "mean": mean_raw,
            "variance": variance,
            "skewness": skewness,
            "kurtosis": kurtosis,
        }
        return self._untruncated_numeric_stats

    def mean(self) -> float:
        if self._truncate_invalid:
            return float("nan")
        if self._truncated_stats is not None:
            return self._truncated_stats["mean"] + self.shift_amount
        if math.isfinite(self._raw_mean):
            return self._raw_mean + self.shift_amount
        return self._compute_untruncated_numeric_stats()["mean"] + self.shift_amount

    def variance(self) -> float:
        if self._truncate_invalid:
            return float("nan")
        if self._truncated_stats is not None:
            return self._truncated_stats["variance"]
        if math.isfinite(self._raw_var):
            return self._raw_var
        return self._compute_untruncated_numeric_stats()["variance"]

    def skewness(self) -> float:
        if self._truncate_invalid:
            return float("nan")
        if self._truncated_stats is not None:
            return self._truncated_stats["skewness"]
        if math.isfinite(self._raw_skew):
            return self._raw_skew
        return self._compute_untruncated_numeric_stats()["skewness"]

    def kurtosis(self) -> float:
        if self._truncate_invalid:
            return float("nan")
        if self._truncated_stats is not None:
            return self._truncated_stats["kurtosis"]
        if math.isfinite(self._raw_kurt):
            return self._raw_kurt
        return self._compute_untruncated_numeric_stats()["kurtosis"]

    def mode(self) -> float:
        if self._truncate_invalid:
            return float("nan")
        if self._truncated_stats is not None:
            return self._truncated_stats["mode"] + self.shift_amount
        return self._raw_mode + self.shift_amount

    def min_val(self) -> float:
        if self._truncate_invalid:
            return float("nan")
        if self.is_truncated():
            lower, _ = self.get_effective_bounds()
            if lower is not None:
                return lower
        return self.gamma + self.shift_amount

    def max_val(self) -> float:
        if self._truncate_invalid:
            return float("nan")
        if self.is_truncated():
            _, upper = self.get_effective_bounds()
            if upper is not None:
                return upper
        return float("inf")
