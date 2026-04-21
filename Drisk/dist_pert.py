"""Pert distribution support for Drisk."""

import math
from typing import List, Optional, Tuple, Union

import numpy as np
from scipy.special import betainc, betaincinv
from scipy.stats import beta as scipy_beta

from distribution_base import DistributionBase

_EPS = 1e-12


def _validate_params(min_val: float, m_likely: float, max_val: float) -> None:
    if float(max_val) <= float(min_val):
        raise ValueError("Pert requires Min < Max")
    if float(m_likely) < float(min_val) or float(m_likely) > float(max_val):
        raise ValueError("Pert requires Min <= M.likely <= Max")


def _shape_params(min_val: float, m_likely: float, max_val: float) -> Tuple[float, float, float]:
    _validate_params(min_val, m_likely, max_val)
    mean_val = (float(min_val) + 4.0 * float(m_likely) + float(max_val)) / 6.0
    span = float(max_val) - float(min_val)
    alpha1 = 6.0 * (mean_val - float(min_val)) / span
    alpha2 = 6.0 * (float(max_val) - mean_val) / span
    return mean_val, alpha1, alpha2


def _create_dist(min_val: float, m_likely: float, max_val: float):
    _, alpha1, alpha2 = _shape_params(min_val, m_likely, max_val)
    return scipy_beta(
        a=float(alpha1),
        b=float(alpha2),
        loc=float(min_val),
        scale=float(max_val) - float(min_val),
    )


def pert_pdf(x: float, min_val: float, m_likely: float, max_val: float) -> float:
    x = float(x)
    if x < float(min_val) or x > float(max_val):
        return 0.0
    dist = _create_dist(min_val, m_likely, max_val)
    return float(dist.pdf(x))


def pert_cdf(x: float, min_val: float, m_likely: float, max_val: float) -> float:
    x = float(x)
    if x <= float(min_val):
        return 0.0
    if x >= float(max_val):
        return 1.0
    _, alpha1, alpha2 = _shape_params(min_val, m_likely, max_val)
    z = (x - float(min_val)) / (float(max_val) - float(min_val))
    return float(betainc(float(alpha1), float(alpha2), z))


def pert_ppf(q_prob: float, min_val: float, m_likely: float, max_val: float) -> float:
    q_prob = float(q_prob)
    if q_prob <= 0.0:
        return float(min_val)
    if q_prob >= 1.0:
        return float(max_val)
    _, alpha1, alpha2 = _shape_params(min_val, m_likely, max_val)
    z = float(betaincinv(float(alpha1), float(alpha2), q_prob))
    return float(float(min_val) + (float(max_val) - float(min_val)) * z)


def pert_raw_mean(min_val: float, m_likely: float, max_val: float) -> float:
    mean_val, _, _ = _shape_params(min_val, m_likely, max_val)
    return float(mean_val)


def pert_generator_single(rng: np.random.Generator, params: List[float]) -> float:
    min_val = float(params[0])
    m_likely = float(params[1])
    max_val = float(params[2])
    u = float(rng.uniform(_EPS, 1.0 - _EPS))
    return pert_ppf(u, min_val, m_likely, max_val)


def pert_generator_vectorized(
    rng: np.random.Generator,
    params: List[Union[float, np.ndarray]],
    n_samples: int,
) -> np.ndarray:
    """
    向量化生成 Pert 分布样本，支持参数为标量或数组（每个迭代不同参数）。

    Args:
        rng: NumPy 随机生成器
        params: [min, likely, max]，每个可以是标量或长度为 n_samples 的数组
        n_samples: 样本数量

    Returns:
        长度为 n_samples 的样本数组
    """
    min_val = params[0]
    m_likely = params[1]
    max_val = params[2]

    # 将标量广播为数组
    if not isinstance(min_val, np.ndarray):
        min_val = np.full(n_samples, float(min_val))
    if not isinstance(m_likely, np.ndarray):
        m_likely = np.full(n_samples, float(m_likely))
    if not isinstance(max_val, np.ndarray):
        max_val = np.full(n_samples, float(max_val))

    # 向量化计算 α1, α2
    mean_val = (min_val + 4.0 * m_likely + max_val) / 6.0
    span = max_val - min_val
    alpha1 = 6.0 * (mean_val - min_val) / span
    alpha2 = 6.0 * (max_val - mean_val) / span

    # 生成均匀随机数
    u = rng.uniform(_EPS, 1.0 - _EPS, size=n_samples)

    # 使用 scipy.stats.beta.ppf 支持数组参数
    z = scipy_beta.ppf(u, alpha1, alpha2)

    # 计算最终样本
    samples = min_val + span * z
    return samples.astype(float)


class PertDistribution(DistributionBase):
    """Pert theoretical distribution with truncation support."""

    def __init__(self, params: List[float], markers: dict = None, func_name: str = None):
        markers = markers or {}
        min_point = float(params[0]) if len(params) >= 1 else -10.0
        mode_point = float(params[1]) if len(params) >= 2 else 0.0
        max_point = float(params[2]) if len(params) >= 3 else 10.0

        mean_val, alpha1, alpha2 = _shape_params(min_point, mode_point, max_point)

        self.min_point = min_point
        self.mode_point = mode_point
        self.max_point = max_point
        self.alpha1 = float(alpha1)
        self.alpha2 = float(alpha2)
        self._dist = _create_dist(min_point, mode_point, max_point)
        self.support_low = float(min_point)
        self.support_high = float(max_point)
        self._raw_mean = float(mean_val)
        self._raw_var = float((mean_val - min_point) * (max_point - mean_val) / 7.0)

        if self._raw_var > 0.0:
            numerator = (min_point + max_point - 2.0 * mean_val) / 4.0
            denominator = math.sqrt((mean_val - min_point) * (max_point - mean_val) / 7.0)
            self._raw_skew = float(numerator / denominator)
        else:
            self._raw_skew = 0.0

        self._raw_kurt = float(21.0 / (self.alpha1 * self.alpha2))
        self._raw_mode = float(mode_point)
        self._truncated_stats = None

        super().__init__([min_point, mode_point, max_point], markers, func_name)

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
        return pert_pdf(x, self.min_point, self.mode_point, self.max_point)

    def _original_cdf(self, x: float) -> float:
        return pert_cdf(x, self.min_point, self.mode_point, self.max_point)

    def _original_ppf(self, q_prob: float) -> float:
        return pert_ppf(q_prob, self.min_point, self.mode_point, self.max_point)

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

    def _compute_truncated_stats(self) -> None:
        if not self.is_truncated():
            self._truncated_stats = None
            return

        low_orig, high_orig = self._get_original_truncation_bounds()
        lb = self.min_point if low_orig is None else max(self.min_point, low_orig)
        ub = self.max_point if high_orig is None else min(self.max_point, high_orig)
        if ub < lb:
            self._truncated_stats = None
            return

        mass = max(0.0, self._original_cdf(ub) - self._original_cdf(lb))
        if mass <= _EPS:
            self._truncated_stats = None
            return

        m1 = float(self._dist.expect(lambda x: x, lb=lb, ub=ub, conditional=True))
        m2 = float(self._dist.expect(lambda x: x * x, lb=lb, ub=ub, conditional=True))
        m3_raw = float(self._dist.expect(lambda x: x ** 3, lb=lb, ub=ub, conditional=True))
        m4_raw = float(self._dist.expect(lambda x: x ** 4, lb=lb, ub=ub, conditional=True))

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
        return float(self.min_point + self.shift_amount)

    def max_val(self) -> float:
        if self._truncate_invalid:
            return float("nan")
        _, upper = self.get_effective_bounds()
        if upper is not None:
            return upper
        return float(self.max_point + self.shift_amount)