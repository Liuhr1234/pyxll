"""BetaGeneral distribution support for Drisk."""

import math
from typing import List, Optional, Tuple, Union

import numpy as np
from scipy.integrate import quad
from scipy.special import betainc, betaincinv
from scipy.stats import beta as scipy_beta

from distribution_base import DistributionBase

_EPS = 1e-12


def _validate_params(alpha1: float, alpha2: float, min_val: float, max_val: float) -> None:
    if float(alpha1) <= 0.0:
        raise ValueError("BetaGeneral requires alpha1 > 0")
    if float(alpha2) <= 0.0:
        raise ValueError("BetaGeneral requires alpha2 > 0")
    if float(max_val) <= float(min_val):
        raise ValueError("BetaGeneral requires Max > Min")


def _span(min_val: float, max_val: float) -> float:
    return float(max_val) - float(min_val)


def _create_dist(alpha1: float, alpha2: float, min_val: float, max_val: float):
    _validate_params(alpha1, alpha2, min_val, max_val)
    return scipy_beta(a=float(alpha1), b=float(alpha2), loc=float(min_val), scale=_span(min_val, max_val))


def betageneral_pdf(x: float, alpha1: float, alpha2: float, min_val: float, max_val: float) -> float:
    x = float(x)
    if x < float(min_val) or x > float(max_val):
        return 0.0
    dist = _create_dist(alpha1, alpha2, min_val, max_val)
    return float(dist.pdf(x))


def betageneral_cdf(x: float, alpha1: float, alpha2: float, min_val: float, max_val: float) -> float:
    x = float(x)
    if x <= float(min_val):
        return 0.0
    if x >= float(max_val):
        return 1.0
    _validate_params(alpha1, alpha2, min_val, max_val)
    z = (x - float(min_val)) / _span(min_val, max_val)
    return float(betainc(float(alpha1), float(alpha2), z))


def betageneral_ppf(q_prob: float, alpha1: float, alpha2: float, min_val: float, max_val: float) -> float:
    q_prob = float(q_prob)
    if q_prob <= 0.0:
        return float(min_val)
    if q_prob >= 1.0:
        return float(max_val)
    _validate_params(alpha1, alpha2, min_val, max_val)
    z = float(betaincinv(float(alpha1), float(alpha2), q_prob))
    return float(float(min_val) + _span(min_val, max_val) * z)


def betageneral_raw_mean(alpha1: float, alpha2: float, min_val: float, max_val: float) -> float:
    _validate_params(alpha1, alpha2, min_val, max_val)
    return float(float(min_val) + _span(min_val, max_val) * float(alpha1) / (float(alpha1) + float(alpha2)))


def betageneral_generator_single(rng: np.random.Generator, params: List[float]) -> float:
    alpha1 = float(params[0])
    alpha2 = float(params[1])
    min_val = float(params[2])
    max_val = float(params[3])
    _validate_params(alpha1, alpha2, min_val, max_val)
    return float(min_val + _span(min_val, max_val) * rng.beta(alpha1, alpha2))


def betageneral_generator_vectorized(
    rng: np.random.Generator,
    params: List[Union[float, np.ndarray]],
    n_samples: int,
) -> np.ndarray:
    alpha1 = params[0]
    alpha2 = params[1]
    min_val = params[2]
    max_val = params[3]

    # 广播为数组
    if not isinstance(alpha1, np.ndarray):
        alpha1 = np.full(n_samples, float(alpha1))
    if not isinstance(alpha2, np.ndarray):
        alpha2 = np.full(n_samples, float(alpha2))
    if not isinstance(min_val, np.ndarray):
        min_val = np.full(n_samples, float(min_val))
    if not isinstance(max_val, np.ndarray):
        max_val = np.full(n_samples, float(max_val))

    # 验证（逐个元素，可简化为向量化检查，这里保持简单）
    valid = (alpha1 > 0) & (alpha2 > 0) & (max_val > min_val)
    if not np.all(valid):
        raise ValueError("BetaGeneral parameters invalid")

    span = max_val - min_val
    # rng.beta 支持数组参数 a, b
    samples = min_val + span * rng.beta(alpha1, alpha2, size=n_samples)
    return samples.astype(float)


class BetaGeneralDistribution(DistributionBase):
    """BetaGeneral theoretical distribution with truncation support."""

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
        self.range_param = _span(min_val, max_val)
        self._dist = _create_dist(alpha1, alpha2, min_val, max_val)
        self.support_low = min_val
        self.support_high = max_val

        alpha_sum = alpha1 + alpha2
        self._raw_mean = float(min_val + self.range_param * alpha1 / alpha_sum)
        self._raw_var = float((self.range_param ** 2) * alpha1 * alpha2 / ((alpha_sum ** 2) * (alpha_sum + 1.0)))

        if alpha1 > 0.0 and alpha2 > 0.0:
            self._raw_skew = float(
                2.0 * (alpha2 - alpha1) / (alpha_sum + 2.0) * math.sqrt((alpha_sum + 1.0) / (alpha1 * alpha2))
            )
            self._raw_kurt = float(
                3.0
                * ((alpha_sum + 1.0) * (2.0 * (alpha_sum ** 2) + alpha1 * alpha2 * (alpha_sum - 6.0)))
                / (alpha1 * alpha2 * (alpha_sum + 2.0) * (alpha_sum + 3.0))
            )
        else:
            self._raw_skew = float("nan")
            self._raw_kurt = float("nan")

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

    def _calculate_raw_mode(self) -> float:
        if self.alpha1 > 1.0 and self.alpha2 > 1.0:
            z_mode = (self.alpha1 - 1.0) / (self.alpha1 + self.alpha2 - 2.0)
            return float(self.min_param + self.range_param * z_mode)
        if self.alpha1 < 1.0 and self.alpha2 < 1.0:
            return float((self.min_param + self.max_param) / 2.0)
        if self.alpha1 <= 1.0 and self.alpha2 <= 1.0:
            return float((self.min_param + self.max_param) / 2.0)
        if self.alpha1 < 1.0:
            return float(self.min_param)
        return float(self.max_param)

    def _original_pdf(self, x: float) -> float:
        return betageneral_pdf(x, self.alpha1, self.alpha2, self.min_param, self.max_param)

    def _original_cdf(self, x: float) -> float:
        return betageneral_cdf(x, self.alpha1, self.alpha2, self.min_param, self.max_param)

    def _original_ppf(self, q_prob: float) -> float:
        return betageneral_ppf(q_prob, self.alpha1, self.alpha2, self.min_param, self.max_param)

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
        lb = self.min_param if low_orig is None else max(self.min_param, low_orig)
        ub = self.max_param if high_orig is None else min(self.max_param, high_orig)
        if ub < lb:
            self._truncated_stats = None
            return

        mass = max(0.0, self._original_cdf(ub) - self._original_cdf(lb))
        if mass <= _EPS:
            self._truncated_stats = None
            return

        try:
            m1 = float(self._dist.expect(lambda x: x, lb=lb, ub=ub, conditional=True))
            m2 = float(self._dist.expect(lambda x: x * x, lb=lb, ub=ub, conditional=True))
            m3_raw = float(self._dist.expect(lambda x: x ** 3, lb=lb, ub=ub, conditional=True))
            m4_raw = float(self._dist.expect(lambda x: x ** 4, lb=lb, ub=ub, conditional=True))
        except Exception:
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
        elif variance == 0.0:
            skewness = 0.0
            kurtosis = 3.0
        else:
            skewness = float("nan")
            kurtosis = float("nan")

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
        return float(self.min_param + self.shift_amount)

    def max_val(self) -> float:
        if self._truncate_invalid:
            return float("nan")
        _, upper = self.get_effective_bounds()
        if upper is not None:
            return upper
        return float(self.max_param + self.shift_amount)
