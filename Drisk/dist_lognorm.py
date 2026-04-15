"""Lognorm distribution support for Drisk."""

import math
from typing import List, Optional, Tuple

import numpy as np
from scipy.integrate import quad
from scipy.stats import lognorm as scipy_lognorm

from distribution_base import DistributionBase

_EPS = 1e-12


def _validate_params(mu: float, sigma: float) -> None:
    if mu <= 0.0:
        raise ValueError("Lognorm requires mu > 0")
    if sigma <= 0.0:
        raise ValueError("Lognorm requires sigma > 0")


def _to_internal_params(mu: float, sigma: float) -> Tuple[float, float]:
    _validate_params(mu, sigma)
    sigma_prime = math.sqrt(math.log(1.0 + (sigma / mu) ** 2))
    mu_prime = math.log((mu * mu) / math.sqrt(sigma * sigma + mu * mu))
    return mu_prime, sigma_prime


def _create_dist(mu: float, sigma: float):
    mu_prime, sigma_prime = _to_internal_params(mu, sigma)
    return scipy_lognorm(s=sigma_prime, scale=math.exp(mu_prime), loc=0.0)


def lognorm_pdf(x: float, mu: float, sigma: float) -> float:
    if float(x) <= 0.0:
        return 0.0
    dist = _create_dist(mu, sigma)
    return float(dist.pdf(x))


def lognorm_cdf(x: float, mu: float, sigma: float) -> float:
    if float(x) <= 0.0:
        return 0.0
    dist = _create_dist(mu, sigma)
    return float(dist.cdf(x))


def lognorm_ppf(q: float, mu: float, sigma: float) -> float:
    q = float(q)
    if q <= 0.0:
        return 0.0
    if q >= 1.0:
        return float("inf")
    dist = _create_dist(mu, sigma)
    return float(dist.ppf(q))


def lognorm_raw_mean(mu: float, sigma: float) -> float:
    _validate_params(mu, sigma)
    return float(mu)


def lognorm_generator_single(rng: np.random.Generator, params: List[float]) -> float:
    mu = float(params[0])
    sigma = float(params[1])
    q = float(rng.uniform(_EPS, 1.0 - _EPS))
    return lognorm_ppf(q, mu, sigma)


def lognorm_generator_vectorized(
    rng: np.random.Generator,
    params: List[float],
    n_samples: int,
) -> np.ndarray:
    mu = float(params[0])
    sigma = float(params[1])
    dist = _create_dist(mu, sigma)
    q = rng.uniform(_EPS, 1.0 - _EPS, size=n_samples)
    return np.asarray(dist.ppf(q), dtype=float)


class LognormDistribution(DistributionBase):
    """Lognorm theoretical distribution with truncation support."""

    def __init__(self, params: List[float], markers: dict = None, func_name: str = None):
        markers = markers or {}
        mu = float(params[0]) if len(params) >= 1 else 10.0
        sigma = float(params[1]) if len(params) >= 2 else 10.0

        self.mu = mu
        self.sigma = sigma
        self.mu_prime, self.sigma_prime = _to_internal_params(mu, sigma)
        self._dist = _create_dist(mu, sigma)
        self.support_low = 0.0
        self.support_high = float("inf")
        self._raw_mean = float(mu)
        self._raw_var = float(sigma * sigma)
        sigma_over_mu = sigma / mu
        omega = 1.0 + sigma_over_mu * sigma_over_mu
        self._raw_skew = float(sigma_over_mu ** 3 + 3.0 * sigma_over_mu)
        self._raw_kurt = float(omega ** 4 + 2.0 * omega ** 3 + 3.0 * omega ** 2 - 3.0)
        self._raw_mode = float(math.exp(self.mu_prime - self.sigma_prime ** 2))
        self._truncated_stats = None

        super().__init__([mu, sigma], markers, func_name)

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
        return lognorm_pdf(x, self.mu, self.sigma)

    def _original_cdf(self, x: float) -> float:
        return lognorm_cdf(x, self.mu, self.sigma)

    def _original_ppf(self, q_prob: float) -> float:
        return lognorm_ppf(q_prob, self.mu, self.sigma)

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

        if lb <= 0.0:
            lb = math.nextafter(0.0, float("inf"))

        mass = max(0.0, self._original_cdf(ub) - self._original_cdf(lb))
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
