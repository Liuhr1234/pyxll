"""Pearson5 distribution support for Drisk."""

import math
from typing import List, Optional, Tuple, Union

import numpy as np
from scipy.integrate import quad
from scipy.special import gammaincc
from scipy.stats import invgamma as scipy_invgamma

from distribution_base import DistributionBase

_EPS = 1e-12


def _validate_params(alpha: float, beta: float) -> None:
    if alpha <= 0.0:
        raise ValueError("Pearson5 requires alpha > 0")
    if beta <= 0.0:
        raise ValueError("Pearson5 requires beta > 0")


def _create_dist(alpha: float, beta: float):
    _validate_params(alpha, beta)
    return scipy_invgamma(a=float(alpha), loc=0.0, scale=float(beta))


def pearson5_pdf(x: float, alpha: float, beta: float) -> float:
    if float(x) <= 0.0:
        return 0.0
    dist = _create_dist(alpha, beta)
    return float(dist.pdf(x))


def pearson5_cdf(x: float, alpha: float, beta: float) -> float:
    if float(x) <= 0.0:
        return 0.0
    _validate_params(alpha, beta)
    return float(gammaincc(float(alpha), float(beta) / float(x)))


def pearson5_ppf(q_prob: float, alpha: float, beta: float) -> float:
    q_prob = float(q_prob)
    if q_prob <= 0.0:
        return 0.0
    if q_prob >= 1.0:
        return float("inf")
    dist = _create_dist(alpha, beta)
    return float(dist.ppf(q_prob))


def pearson5_raw_mean(alpha: float, beta: float) -> float:
    _validate_params(alpha, beta)
    if alpha <= 1.0:
        return float("inf")
    return float(beta / (alpha - 1.0))


def pearson5_generator_single(rng: np.random.Generator, params: List[float]) -> float:
    alpha = float(params[0])
    beta = float(params[1])
    u = float(rng.uniform(_EPS, 1.0 - _EPS))
    return pearson5_ppf(u, alpha, beta)


def pearson5_generator_vectorized(
    rng: np.random.Generator,
    params: List[Union[float, np.ndarray]],
    n_samples: int,
) -> np.ndarray:
    alpha = params[0]
    beta = params[1]

    if not isinstance(alpha, np.ndarray):
        alpha = np.full(n_samples, float(alpha))
    if not isinstance(beta, np.ndarray):
        beta = np.full(n_samples, float(beta))

    if np.any((alpha <= 0) | (beta <= 0)):
        raise ValueError("Pearson5 requires alpha>0, beta>0")

    from scipy.stats import invgamma
    return invgamma.rvs(a=alpha, scale=beta, size=n_samples, random_state=rng)

class Pearson5Distribution(DistributionBase):
    """Pearson5 theoretical distribution with truncation support."""

    def __init__(self, params: List[float], markers: dict = None, func_name: str = None):
        markers = markers or {}
        alpha = float(params[0]) if len(params) >= 1 else 3.0
        beta = float(params[1]) if len(params) >= 2 else 10.0

        _validate_params(alpha, beta)

        self.alpha = alpha
        self.beta = beta
        self._dist = _create_dist(alpha, beta)
        self.support_low = 0.0
        self.support_high = float("inf")
        self._raw_mean = pearson5_raw_mean(alpha, beta)
        if alpha > 2.0:
            self._raw_var = float((beta * beta) / (((alpha - 1.0) ** 2) * (alpha - 2.0)))
        else:
            self._raw_var = float("inf")
        if alpha > 3.0:
            self._raw_skew = float((4.0 * math.sqrt(alpha - 2.0)) / (alpha - 3.0))
        else:
            self._raw_skew = float("nan")
        if alpha > 4.0:
            kurt_excess = 6.0 * (5.0 * alpha - 11.0) / ((alpha - 3.0) * (alpha - 4.0))
            self._raw_kurt = float(kurt_excess + 3.0)
        else:
            self._raw_kurt = float("nan")
        self._raw_mode = float(beta / (alpha + 1.0))
        self._truncated_stats = None

        super().__init__([alpha, beta], markers, func_name)

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
        return pearson5_pdf(x, self.alpha, self.beta)

    def _original_cdf(self, x: float) -> float:
        return pearson5_cdf(x, self.alpha, self.beta)

    def _original_ppf(self, q_prob: float) -> float:
        return pearson5_ppf(q_prob, self.alpha, self.beta)

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
        lb = 0.0 if low_orig is None else max(0.0, low_orig)
        ub = float("inf") if high_orig is None else high_orig
        if ub < lb:
            self._truncated_stats = None
            return

        mass = max(0.0, self._original_cdf(ub) - self._original_cdf(lb))
        if mass <= _EPS:
            self._truncated_stats = None
            return

        finite_ub = ub
        if math.isinf(finite_ub):
            finite_ub = self._original_ppf(1.0 - 1e-8)

        m1 = self._moment_integral(lb, finite_ub, 1) / mass
        m2 = self._moment_integral(lb, finite_ub, 2) / mass
        m3_raw = self._moment_integral(lb, finite_ub, 3) / mass
        m4_raw = self._moment_integral(lb, finite_ub, 4) / mass
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
        return float(self.shift_amount)

    def max_val(self) -> float:
        if self._truncate_invalid:
            return float("nan")
        _, upper = self.get_effective_bounds()
        if upper is not None:
            return upper
        return float("inf")
