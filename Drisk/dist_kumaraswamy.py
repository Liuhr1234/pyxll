"""Kumaraswamy distribution support for Drisk."""

import math
from typing import List, Optional, Tuple

import numpy as np
from scipy.integrate import quad
from scipy.special import beta as beta_func

from distribution_base import DistributionBase

_EPS = 1e-12


def _validate_params(alpha1: float, alpha2: float, min_val: float, max_val: float) -> None:
    if alpha1 <= 0:
        raise ValueError("Kumaraswamy requires alpha1 > 0")
    if alpha2 <= 0:
        raise ValueError("Kumaraswamy requires alpha2 > 0")
    if max_val <= min_val:
        raise ValueError("Kumaraswamy requires max_val > min_val")


def _range(min_val: float, max_val: float) -> float:
    return float(max_val) - float(min_val)


def _moment_q(alpha1: float, alpha2: float, n: int) -> float:
    return float(alpha2 * beta_func(1.0 + float(n) / float(alpha1), float(alpha2)))


def _z_from_x(x: float, min_val: float, max_val: float) -> float:
    rng = _range(min_val, max_val)
    if rng <= 0:
        return 0.0
    z = (float(x) - float(min_val)) / rng
    return float(min(max(z, 0.0), 1.0))


def _x_from_z(z: float, min_val: float, max_val: float) -> float:
    return float(min_val) + _range(min_val, max_val) * float(z)


def kumaraswamy_pdf(x: float, alpha1: float, alpha2: float, min_val: float, max_val: float) -> float:
    _validate_params(alpha1, alpha2, min_val, max_val)
    x = float(x)
    if x < min_val or x > max_val:
        return 0.0

    rng = _range(min_val, max_val)
    z = _z_from_x(x, min_val, max_val)

    if z <= 0.0:
        if alpha1 < 1.0:
            return float("inf")
        if abs(alpha1 - 1.0) <= _EPS:
            return float(alpha1 * alpha2 / rng)
        return 0.0

    if z >= 1.0:
        if alpha2 < 1.0:
            return float("inf")
        if abs(alpha2 - 1.0) <= _EPS:
            return float(alpha1 * alpha2 / rng)
        return 0.0

    z_pow = z ** alpha1
    return float((alpha1 * alpha2 / rng) * (z ** (alpha1 - 1.0)) * ((1.0 - z_pow) ** (alpha2 - 1.0)))


def kumaraswamy_cdf(x: float, alpha1: float, alpha2: float, min_val: float, max_val: float) -> float:
    _validate_params(alpha1, alpha2, min_val, max_val)
    x = float(x)
    if x <= min_val:
        return 0.0
    if x >= max_val:
        return 1.0

    z = _z_from_x(x, min_val, max_val)
    z_pow = z ** alpha1
    return float(1.0 - ((1.0 - z_pow) ** alpha2))


def kumaraswamy_ppf(q: float, alpha1: float, alpha2: float, min_val: float, max_val: float) -> float:
    _validate_params(alpha1, alpha2, min_val, max_val)
    q = float(q)
    if q <= 0.0:
        return float(min_val)
    if q >= 1.0:
        return float(max_val)

    # Stable form of 1 - (1 - q)^(1/alpha2)
    term = -math.expm1(math.log1p(-q) / float(alpha2))
    if term <= 0.0:
        return float(min_val)
    z = term ** (1.0 / float(alpha1))
    return float(_x_from_z(z, min_val, max_val))


def kumaraswamy_raw_mean(alpha1: float, alpha2: float, min_val: float, max_val: float) -> float:
    _validate_params(alpha1, alpha2, min_val, max_val)
    return float(min_val + _range(min_val, max_val) * _moment_q(alpha1, alpha2, 1))


def kumaraswamy_generator_single(rng: np.random.Generator, params: List[float]) -> float:
    alpha1 = float(params[0])
    alpha2 = float(params[1])
    min_val = float(params[2])
    max_val = float(params[3])
    q = float(rng.uniform(_EPS, 1.0 - _EPS))
    return kumaraswamy_ppf(q, alpha1, alpha2, min_val, max_val)


def kumaraswamy_generator_vectorized(
    rng: np.random.Generator,
    params: List[float],
    n_samples: int,
) -> np.ndarray:
    alpha1 = float(params[0])
    alpha2 = float(params[1])
    min_val = float(params[2])
    max_val = float(params[3])
    _validate_params(alpha1, alpha2, min_val, max_val)
    q = rng.uniform(_EPS, 1.0 - _EPS, size=n_samples)
    term = -np.expm1(np.log1p(-q) / alpha2)
    z = np.power(term, 1.0 / alpha1)
    return np.asarray(min_val + (max_val - min_val) * z, dtype=float)


class KumaraswamyDistribution(DistributionBase):
    """Kumaraswamy theoretical distribution with truncation support."""

    def __init__(self, params: List[float], markers: dict = None, func_name: str = None):
        markers = markers or {}
        alpha1 = float(params[0]) if len(params) >= 1 else 2.0
        alpha2 = float(params[1]) if len(params) >= 2 else 2.0
        min_val = float(params[2]) if len(params) >= 3 else -10.0
        max_val = float(params[3]) if len(params) >= 4 else 10.0

        _validate_params(alpha1, alpha2, min_val, max_val)

        self.alpha1 = alpha1
        self.alpha2 = alpha2
        self.min_param = min_val
        self.max_param = max_val
        self.range_param = _range(min_val, max_val)
        self.support_low = min_val
        self.support_high = max_val

        self.q1 = _moment_q(alpha1, alpha2, 1)
        self.q2 = _moment_q(alpha1, alpha2, 2)
        self.q3 = _moment_q(alpha1, alpha2, 3)
        self.q4 = _moment_q(alpha1, alpha2, 4)

        self._raw_mean = float(min_val + self.range_param * self.q1)
        self._raw_var = float((self.range_param ** 2) * (self.q2 - self.q1 ** 2))
        self._raw_skew = self._calculate_raw_skewness()
        self._raw_kurt = self._calculate_raw_kurtosis()
        self._raw_mode = self._calculate_raw_mode()
        self._truncated_stats = None

        super().__init__([alpha1, alpha2, min_val, max_val], markers, func_name)

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

    def _calculate_raw_skewness(self) -> float:
        variance_component = self.q2 - self.q1 ** 2
        if variance_component <= 0.0:
            return 0.0
        numerator = self.q3 - 3.0 * self.q1 * self.q2 + 2.0 * self.q1 ** 3
        return float(numerator / (variance_component ** 1.5))

    def _calculate_raw_kurtosis(self) -> float:
        variance_component = self.q2 - self.q1 ** 2
        if variance_component <= 0.0:
            return 3.0
        numerator = self.q4 - 4.0 * self.q1 * self.q3 + 6.0 * (self.q1 ** 2) * self.q2 - 3.0 * (self.q1 ** 4)
        return float(numerator / (variance_component ** 2))

    def _calculate_raw_mode(self) -> float:
        if self.alpha1 > 1.0 and self.alpha2 > 1.0:
            z_mode = ((self.alpha1 - 1.0) / (self.alpha1 * self.alpha2 - 1.0)) ** (1.0 / self.alpha1)
            return float(_x_from_z(z_mode, self.min_param, self.max_param))
        if self.alpha1 <= 1.0 and self.alpha2 > 1.0:
            return float(self.min_param)
        if self.alpha1 > 1.0 and self.alpha2 <= 1.0:
            return float(self.max_param)
        return float(self._raw_mean)

    def _original_pdf(self, x: float) -> float:
        return kumaraswamy_pdf(x, self.alpha1, self.alpha2, self.min_param, self.max_param)

    def _original_cdf(self, x: float) -> float:
        return kumaraswamy_cdf(x, self.alpha1, self.alpha2, self.min_param, self.max_param)

    def _original_ppf(self, q_prob: float) -> float:
        return kumaraswamy_ppf(q_prob, self.alpha1, self.alpha2, self.min_param, self.max_param)

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
        z_lower = _z_from_x(lower, self.min_param, self.max_param)
        z_upper = _z_from_x(upper, self.min_param, self.max_param)
        if z_upper <= z_lower:
            return 0.0

        def integrand(z: float) -> float:
            x = _x_from_z(z, self.min_param, self.max_param)
            return (x ** power) * self.alpha1 * self.alpha2 * (z ** (self.alpha1 - 1.0)) * ((1.0 - z ** self.alpha1) ** (self.alpha2 - 1.0))

        result, _ = quad(integrand, z_lower, z_upper, limit=300)
        return float(result)

    def _compute_truncated_stats(self) -> None:
        if not self.is_truncated():
            self._truncated_stats = None
            return

        low_orig, high_orig = self._get_original_truncation_bounds()
        lb = self.min_param if low_orig is None else max(self.min_param, low_orig)
        ub = self.max_param if high_orig is None else min(self.max_param, high_orig)
        if ub < lb:
            self._truncated_stats = None
            return

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
            return self._truncated_stats["mode"] + self.shift_amount
        return self._raw_mode + self.shift_amount

    def min_val(self) -> float:
        if self._truncate_invalid:
            return float("nan")
        lower, _ = self.get_effective_bounds()
        if lower is not None:
            return lower
        return float(self.min_param + self.shift_amount)

    def max_val(self) -> float:
        if self._truncate_invalid:
            return float("nan")
        _, upper = self.get_effective_bounds()
        if upper is not None:
            return upper
        return float(self.max_param + self.shift_amount)
