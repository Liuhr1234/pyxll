"""Loglogistic distribution support for Drisk."""

import math
from typing import List, Optional, Tuple, Union

import numpy as np
from scipy.integrate import quad
from scipy.stats import fisk as scipy_fisk

from distribution_base import DistributionBase

_EPS = 1e-12


def _validate_params(gamma: float, beta: float, alpha: float) -> None:
    if beta <= 0.0:
        raise ValueError("Loglogistic requires beta > 0")
    if alpha <= 0.0:
        raise ValueError("Loglogistic requires alpha > 0")


def _create_dist(gamma: float, beta: float, alpha: float):
    _validate_params(gamma, beta, alpha)
    return scipy_fisk(float(alpha), loc=float(gamma), scale=float(beta))


def _raw_mode_value(gamma: float, beta: float, alpha: float) -> float:
    if alpha <= 1.0:
        return float(gamma)
    ratio = (alpha - 1.0) / (alpha + 1.0)
    return float(gamma + beta * (ratio ** (1.0 / alpha)))


def loglogistic_pdf(x: float, gamma: float, beta: float, alpha: float) -> float:
    if float(x) <= float(gamma):
        return 0.0
    dist = _create_dist(gamma, beta, alpha)
    return float(dist.pdf(x))


def loglogistic_cdf(x: float, gamma: float, beta: float, alpha: float) -> float:
    if float(x) <= float(gamma):
        return 0.0
    dist = _create_dist(gamma, beta, alpha)
    return float(dist.cdf(x))


def loglogistic_ppf(q: float, gamma: float, beta: float, alpha: float) -> float:
    q = float(q)
    if q <= 0.0:
        return float(gamma)
    if q >= 1.0:
        return float("inf")
    dist = _create_dist(gamma, beta, alpha)
    return float(dist.ppf(q))


def loglogistic_raw_mean(gamma: float, beta: float, alpha: float) -> float:
    _validate_params(gamma, beta, alpha)
    if alpha <= 1.0:
        return float("nan")
    theta = math.pi / alpha
    return float(gamma + beta * theta / math.sin(theta))


def loglogistic_generator_single(rng: np.random.Generator, params: List[float]) -> float:
    gamma = float(params[0])
    beta = float(params[1])
    alpha = float(params[2])
    q = float(rng.uniform(_EPS, 1.0 - _EPS))
    return loglogistic_ppf(q, gamma, beta, alpha)


def loglogistic_generator_vectorized(
    rng: np.random.Generator,
    params: List[Union[float, np.ndarray]],
    n_samples: int,
) -> np.ndarray:
    gamma = params[0]
    beta = params[1]
    alpha = params[2]

    for i, arr in enumerate([gamma, beta, alpha]):
        if not isinstance(arr, np.ndarray):
            arr = np.full(n_samples, float(arr))
        else:
            arr = arr.astype(float)
        if i == 0:
            gamma_arr = arr
        elif i == 1:
            beta_arr = arr
        else:
            alpha_arr = arr

    if np.any((beta_arr <= 0) | (alpha_arr <= 0)):
        raise ValueError("Loglogistic requires beta>0, alpha>0")

    from scipy.stats import fisk
    return fisk.rvs(c=alpha_arr, loc=gamma_arr, scale=beta_arr, size=n_samples, random_state=rng)


class LoglogisticDistribution(DistributionBase):
    """Loglogistic theoretical distribution with truncation support."""

    def __init__(self, params: List[float], markers: dict = None, func_name: str = None):
        markers = markers or {}
        gamma = float(params[0]) if len(params) >= 1 else 0.0
        beta = float(params[1]) if len(params) >= 2 else 1.0
        alpha = float(params[2]) if len(params) >= 3 else 1.0

        self.gamma = gamma
        self.beta = beta
        self.alpha = alpha
        self._dist = _create_dist(gamma, beta, alpha)
        self.support_low = float(gamma)
        self.support_high = float("inf")
        self._raw_mean = loglogistic_raw_mean(gamma, beta, alpha)
        self._raw_var = float("nan")
        self._raw_skew = float("nan")
        self._raw_kurt = float("nan")
        self._raw_mode = _raw_mode_value(gamma, beta, alpha)
        self._truncated_stats = None

        theta = math.pi / alpha
        if alpha > 2.0:
            self._raw_var = float(
                beta * beta * theta * (2.0 / math.sin(2.0 * theta) - theta / (math.sin(theta) ** 2))
            )
        if alpha > 3.0:
            sin_theta = math.sin(theta)
            sin_2theta = math.sin(2.0 * theta)
            sin_3theta = math.sin(3.0 * theta)
            numerator = (
                3.0 / sin_3theta
                - 6.0 * theta / (sin_2theta * sin_theta)
                + 2.0 * theta * theta / (sin_theta ** 3)
            )
            denominator = math.sqrt(theta) * (
                2.0 / sin_2theta - theta / (sin_theta ** 2)
            ) ** 1.5
            self._raw_skew = float(numerator / denominator)
        if alpha > 4.0:
            sin_theta = math.sin(theta)
            sin_2theta = math.sin(2.0 * theta)
            sin_3theta = math.sin(3.0 * theta)
            sin_4theta = math.sin(4.0 * theta)
            numerator = (
                4.0 / sin_4theta
                - 12.0 * theta / (sin_3theta * sin_theta)
                + 12.0 * theta * theta / (sin_2theta * (sin_theta ** 2))
                - 3.0 * (theta ** 3) / (sin_theta ** 4)
            )
            denominator = theta * (
                2.0 / sin_2theta - theta / (sin_theta ** 2)
            ) ** 2
            self._raw_kurt = float(numerator / denominator)

        super().__init__([gamma, beta, alpha], markers, func_name)

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
        return loglogistic_pdf(x, self.gamma, self.beta, self.alpha)

    def _original_cdf(self, x: float) -> float:
        return loglogistic_cdf(x, self.gamma, self.beta, self.alpha)

    def _original_ppf(self, q_prob: float) -> float:
        return loglogistic_ppf(q_prob, self.gamma, self.beta, self.alpha)

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

        if lb <= self.gamma:
            lb = math.nextafter(self.gamma, float("inf"))

        mass = max(0.0, self._original_cdf(ub) - self._original_cdf(lb))
        if mass <= _EPS:
            self._truncated_stats = None
            return

        stats = {
            "mean": float("nan"),
            "variance": float("nan"),
            "skewness": float("nan"),
            "kurtosis": float("nan"),
            "mode": float(min(max(self._raw_mode, lb), ub)),
        }

        finite_upper = not math.isinf(ub)

        if finite_upper or self.alpha > 1.0:
            m1 = self._moment_integral(lb, ub, 1) / mass
            stats["mean"] = float(m1)
        else:
            self._truncated_stats = stats
            return

        if finite_upper or self.alpha > 2.0:
            m2 = self._moment_integral(lb, ub, 2) / mass
            variance = max(0.0, m2 - stats["mean"] * stats["mean"])
            stats["variance"] = float(variance)
        else:
            self._truncated_stats = stats
            return

        variance = stats["variance"]
        if variance <= _EPS:
            stats["skewness"] = 0.0
            stats["kurtosis"] = 3.0
            self._truncated_stats = stats
            return

        if finite_upper or self.alpha > 3.0:
            m3_raw = self._moment_integral(lb, ub, 3) / mass
            mu3 = m3_raw - 3.0 * stats["mean"] * m2 + 2.0 * (stats["mean"] ** 3)
            stats["skewness"] = float(mu3 / (variance ** 1.5))

        if finite_upper or self.alpha > 4.0:
            m4_raw = self._moment_integral(lb, ub, 4) / mass
            mu4 = (
                m4_raw
                - 4.0 * stats["mean"] * (self._moment_integral(lb, ub, 3) / mass)
                + 6.0 * (stats["mean"] ** 2) * m2
                - 3.0 * (stats["mean"] ** 4)
            )
            stats["kurtosis"] = float(mu4 / (variance ** 2))

        self._truncated_stats = stats

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
