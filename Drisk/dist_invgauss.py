"""
Inverse Gaussian (Wald) distribution support for Drisk.
Provides sampling, CDF/PPF/PDF, and theoretical statistics with truncation support.
"""

import math
from typing import List, Optional, Union

import numpy as np

from distribution_base import DistributionBase


def _validate_params(mu: float, lam: float) -> None:
    if mu <= 0:
        raise ValueError(f"mu must be > 0, got {mu}")
    if lam <= 0:
        raise ValueError(f"lambda must be > 0, got {lam}")


def invgauss_generator_single(rng: np.random.Generator, params: List[float]) -> float:
    mu, lam = float(params[0]), float(params[1])
    _validate_params(mu, lam)

    y = rng.chisquare(1.0)
    term = (mu * mu * y) / (2.0 * lam)
    sqrt_term = math.sqrt(4.0 * mu * lam * y + (mu * y) ** 2)
    x1 = mu + term - (mu / (2.0 * lam)) * sqrt_term
    x2 = mu * mu / x1
    return x1 if rng.random() <= mu / (mu + x1) else x2


def invgauss_generator_vectorized(
    rng: np.random.Generator, params: List[float], n_samples: int
) -> np.ndarray:
    mu, lam = float(params[0]), float(params[1])
    _validate_params(mu, lam)

    y = rng.chisquare(1.0, n_samples)
    term = (mu * mu * y) / (2.0 * lam)
    sqrt_term = np.sqrt(4.0 * mu * lam * y + (mu * y) ** 2)
    x1 = mu + term - (mu / (2.0 * lam)) * sqrt_term
    x2 = mu * mu / x1
    keep_x1 = rng.random(n_samples) <= (mu / (mu + x1))
    return np.where(keep_x1, x1, x2)


def invgauss_generator(
    rng: np.random.Generator,
    params: List[float],
    n_samples: Optional[int] = None,
) -> Union[float, np.ndarray]:
    if n_samples is None:
        return invgauss_generator_single(rng, params)
    return invgauss_generator_vectorized(rng, params, n_samples)


def _std_normal_cdf(z: float) -> float:
    return 0.5 * (1.0 + math.erf(z / math.sqrt(2.0)))


def _safe_scaled_prob(exp_arg: float, prob: float) -> float:
    if prob <= 0.0:
        return 0.0
    log_value = exp_arg + math.log(prob)
    if log_value < -745:
        return 0.0
    if log_value > 709:
        return float("inf")
    return math.exp(log_value)


def invgauss_cdf(x: float, mu: float, lam: float) -> float:
    if x <= 0:
        return 0.0
    _validate_params(mu, lam)

    sqrt_term = math.sqrt(lam / x)
    term1 = sqrt_term * (x / mu - 1.0)
    term2 = sqrt_term * (x / mu + 1.0)

    phi1 = _std_normal_cdf(term1)
    phi2 = _std_normal_cdf(-term2)
    second = _safe_scaled_prob(2.0 * lam / mu, phi2)
    return max(0.0, min(1.0, phi1 + second))


def invgauss_ppf(
    q: float,
    mu: float,
    lam: float,
    tol: float = 1e-12,
    max_iter: int = 1000,
) -> float:
    if q <= 0.0:
        return 0.0
    if q >= 1.0:
        return float("inf")

    _validate_params(mu, lam)

    std_dev = math.sqrt(mu ** 3 / lam)
    lower = 0.0
    upper = max(mu + 10.0 * std_dev, mu * 2.0)
    while invgauss_cdf(upper, mu, lam) < q:
        upper *= 2.0

    for _ in range(max_iter):
        mid = (lower + upper) / 2.0
        cdf_mid = invgauss_cdf(mid, mu, lam)
        if abs(cdf_mid - q) < tol:
            return mid
        if cdf_mid < q:
            lower = mid
        else:
            upper = mid
    return (lower + upper) / 2.0


def invgauss_pdf(x: float, mu: float, lam: float) -> float:
    if x <= 0:
        return 0.0
    _validate_params(mu, lam)
    part1 = math.sqrt(lam / (2.0 * math.pi * x ** 3))
    exponent = -(lam * (x - mu) ** 2) / (2.0 * mu ** 2 * x)
    return part1 * math.exp(exponent)


def invgauss_raw_mean(mu: float, lam: float) -> float:
    return mu


def invgauss_raw_var(mu: float, lam: float) -> float:
    return mu ** 3 / lam


def invgauss_raw_skew(mu: float, lam: float) -> float:
    return 3.0 * math.sqrt(mu / lam)


def invgauss_raw_kurt(mu: float, lam: float) -> float:
    return 3.0 + 15.0 * (mu / lam)


def invgauss_raw_mode(mu: float, lam: float) -> float:
    return mu * (
        math.sqrt(1.0 + (9.0 * mu ** 2) / (4.0 * lam ** 2)) - (3.0 * mu) / (2.0 * lam)
    )


class InvgaussDistribution(DistributionBase):
    """Inverse Gaussian theoretical distribution with truncated moments."""

    _INF_CDF_TOL = 1e-12

    def __init__(self, params: List[float], markers: dict = None, func_name: str = None):
        if len(params) < 2:
            mu, lam = 1.0, 1.0
        else:
            mu, lam = float(params[0]), float(params[1])
        _validate_params(mu, lam)

        self.mu = mu
        self.lam = lam
        self.support_low = 0.0
        self.support_high = float("inf")

        self._raw_mean = invgauss_raw_mean(mu, lam)
        self._raw_var = invgauss_raw_var(mu, lam)
        self._raw_skew = invgauss_raw_skew(mu, lam)
        self._raw_kurt = invgauss_raw_kurt(mu, lam)
        self._raw_mode = invgauss_raw_mode(mu, lam)

        super().__init__(params, markers, func_name)

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
        self._truncated_stats = None
        self._compute_truncated_stats()

    def _adaptive_simpsons(self, f, a, b, eps=1e-9, max_depth=20):
        def simpson_rule(x0, x1):
            mid = (x0 + x1) / 2.0
            return (x1 - x0) * (f(x0) + 4.0 * f(mid) + f(x1)) / 6.0

        def recurse(x0, x1, fa, fm, fb, whole, depth):
            mid = (x0 + x1) / 2.0
            left_mid = (x0 + mid) / 2.0
            right_mid = (mid + x1) / 2.0

            f_left_mid = f(left_mid)
            f_right_mid = f(right_mid)

            left = (mid - x0) * (fa + 4.0 * f_left_mid + fm) / 6.0
            right = (x1 - mid) * (fm + 4.0 * f_right_mid + fb) / 6.0

            if depth >= max_depth or abs(left + right - whole) <= 15.0 * eps:
                return left + right + (left + right - whole) / 15.0

            return recurse(x0, mid, fa, f_left_mid, fm, left, depth + 1) + recurse(
                mid, x1, fm, f_right_mid, fb, right, depth + 1
            )

        if a >= b:
            return 0.0

        fa = f(a)
        fb = f(b)
        mid = (a + b) / 2.0
        fm = f(mid)
        whole = simpson_rule(a, b)
        return recurse(a, b, fa, fm, fb, whole, 0)

    def _finite_upper_bound(self) -> float:
        target = 1.0 - self._INF_CDF_TOL
        bound = max(self.mu + 12.0 * math.sqrt(self.mu ** 3 / self.lam), self.mu * 2.0)
        for _ in range(64):
            if invgauss_cdf(bound, self.mu, self.lam) >= target:
                return bound
            bound *= 2.0
        return bound

    def _finite_integration_bounds(
        self, low_orig: Optional[float], high_orig: Optional[float]
    ) -> tuple[float, float]:
        lower = self.support_low if low_orig is None else max(self.support_low, low_orig)
        upper = self._finite_upper_bound() if high_orig is None or math.isinf(high_orig) else high_orig
        return lower, upper

    def _compute_truncated_stats(self):
        if not self.is_truncated():
            return

        low, high = self.get_truncated_bounds()
        if low is None and high is None:
            return

        if self.truncate_type in ["value2", "percentile2"]:
            low_orig = low - self.shift_amount if low is not None else None
            high_orig = high - self.shift_amount if high is not None else None
        else:
            low_orig = low
            high_orig = high

        lower, upper = self._finite_integration_bounds(low_orig, high_orig)
        if lower >= upper:
            self._truncated_stats = None
            return

        z = invgauss_cdf(upper, self.mu, self.lam) - invgauss_cdf(lower, self.mu, self.lam)
        if z <= 0:
            self._truncated_stats = None
            return

        def make_integrand(power: int):
            return lambda x: (x ** power) * invgauss_pdf(x, self.mu, self.lam)

        m1 = self._adaptive_simpsons(make_integrand(1), lower, upper) / z
        m2 = self._adaptive_simpsons(make_integrand(2), lower, upper) / z
        m3 = self._adaptive_simpsons(make_integrand(3), lower, upper) / z
        m4 = self._adaptive_simpsons(make_integrand(4), lower, upper) / z

        var_trunc = max(0.0, m2 - m1 ** 2)
        if var_trunc > 0:
            mu3 = m3 - 3.0 * m1 * m2 + 2.0 * m1 ** 3
            mu4 = m4 - 4.0 * m1 * m3 + 6.0 * (m1 ** 2) * m2 - 3.0 * m1 ** 4
            skew_trunc = mu3 / (var_trunc ** 1.5)
            kurt_trunc = mu4 / (var_trunc ** 2)
        else:
            skew_trunc = 0.0
            kurt_trunc = 3.0

        if low_orig is not None and self._raw_mode < low_orig:
            mode_trunc = low_orig
        elif high_orig is not None and self._raw_mode > high_orig:
            mode_trunc = high_orig
        else:
            mode_trunc = self._raw_mode

        self._truncated_stats = {
            "mean": m1,
            "variance": var_trunc,
            "skewness": skew_trunc,
            "kurtosis": kurt_trunc,
            "mode": mode_trunc,
        }

    def _original_cdf(self, x: float) -> float:
        return invgauss_cdf(x, self.mu, self.lam)

    def _original_ppf(self, q: float) -> float:
        return invgauss_ppf(q, self.mu, self.lam)

    def _original_pdf(self, x: float) -> float:
        return invgauss_pdf(x, self.mu, self.lam)

    def mean(self) -> float:
        if self._truncate_invalid:
            return float("nan")
        if self._truncated_stats:
            return self._truncated_stats["mean"] + self.shift_amount
        return self._raw_mean + self.shift_amount

    def variance(self) -> float:
        if self._truncate_invalid:
            return float("nan")
        if self._truncated_stats:
            return self._truncated_stats["variance"]
        return self._raw_var

    def skewness(self) -> float:
        if self._truncate_invalid:
            return float("nan")
        if self._truncated_stats:
            return self._truncated_stats["skewness"]
        return self._raw_skew

    def kurtosis(self) -> float:
        if self._truncate_invalid:
            return float("nan")
        if self._truncated_stats:
            return self._truncated_stats["kurtosis"]
        return self._raw_kurt

    def mode(self) -> float:
        if self._truncate_invalid:
            return float("nan")
        if self._truncated_stats:
            return self._truncated_stats["mode"] + self.shift_amount
        raw_mode = self._raw_mode
        if self.is_truncated():
            low, high = self.get_effective_bounds()
            if low is not None and raw_mode + self.shift_amount < low:
                return low
            if high is not None and raw_mode + self.shift_amount > high:
                return high
        return raw_mode + self.shift_amount

    def min_val(self) -> float:
        if self._truncate_invalid:
            return float("nan")
        lower, _ = self.get_effective_bounds()
        if lower is not None:
            return lower
        return self.apply_shift(self.support_low)

    def max_val(self) -> float:
        if self._truncate_invalid:
            return float("nan")
        _, upper = self.get_effective_bounds()
        if upper is not None:
            return upper
        return self.apply_shift(self.support_high)
