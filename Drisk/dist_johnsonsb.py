"""Johnson SB distribution support for Drisk."""

import math
from typing import List, Optional, Tuple, Union

import numpy as np
from scipy.integrate import quad
from scipy.optimize import minimize_scalar
from scipy.stats import johnsonsb as scipy_johnsonsb

from distribution_base import DistributionBase

_EPS = 1e-12


def _create_dist(alpha1: float, alpha2: float, a: float, b: float):
    if alpha2 <= 0:
        raise ValueError("JohnsonSB requires alpha2 > 0")
    if b <= a:
        raise ValueError("JohnsonSB requires b > a")
    # Match @RISK's JohnsonSB(alpha1, alpha2, a, b) parameter order.
    return scipy_johnsonsb(alpha1, alpha2, loc=a, scale=b - a)


def johnsonsb_pdf(x: float, alpha1: float, alpha2: float, a: float, b: float) -> float:
    dist = _create_dist(alpha1, alpha2, a, b)
    return float(dist.pdf(x))


def johnsonsb_cdf(x: float, alpha1: float, alpha2: float, a: float, b: float) -> float:
    dist = _create_dist(alpha1, alpha2, a, b)
    return float(dist.cdf(x))


def johnsonsb_ppf(q: float, alpha1: float, alpha2: float, a: float, b: float) -> float:
    q = float(q)
    if q <= 0.0:
        return float(a)
    if q >= 1.0:
        return float(b)
    dist = _create_dist(alpha1, alpha2, a, b)
    return float(dist.ppf(q))


def johnsonsb_raw_mean(alpha1: float, alpha2: float, a: float, b: float) -> float:
    dist = _create_dist(alpha1, alpha2, a, b)
    return float(dist.mean())


def johnsonsb_generator_single(rng: np.random.Generator, params: List[float]) -> float:
    alpha1 = float(params[0])
    alpha2 = float(params[1])
    a = float(params[2])
    b = float(params[3])
    q = float(rng.uniform(_EPS, 1.0 - _EPS))
    return johnsonsb_ppf(q, alpha1, alpha2, a, b)


def johnsonsb_generator_vectorized(
    rng: np.random.Generator,
    params: List[Union[float, np.ndarray]],
    n_samples: int,
) -> np.ndarray:
    alpha1 = params[0]
    alpha2 = params[1]
    a = params[2]
    b = params[3]

    # 广播
    for i, arr in enumerate([alpha1, alpha2, a, b]):
        if not isinstance(arr, np.ndarray):
            arr = np.full(n_samples, float(arr))
        else:
            arr = arr.astype(float)
        if i == 0:
            alpha1_arr = arr
        elif i == 1:
            alpha2_arr = arr
        elif i == 2:
            a_arr = arr
        else:
            b_arr = arr

    if np.any((alpha2_arr <= 0) | (b_arr <= a_arr)):
        raise ValueError("JohnsonSB requires alpha2>0, b>a")

    q = rng.uniform(_EPS, 1.0 - _EPS, size=n_samples)
    # 使用 scipy.stats.johnsonsb.ppf 支持数组参数
    from scipy.stats import johnsonsb
    samples = johnsonsb.ppf(q, alpha1_arr, alpha2_arr, loc=a_arr, scale=b_arr - a_arr)
    return samples


class JohnsonSBDistribution(DistributionBase):
    """Johnson SB theoretical distribution with truncation support."""

    def __init__(self, params: List[float], markers: dict = None, func_name: str = None):
        markers = markers or {}
        alpha1 = float(params[0]) if len(params) >= 1 else 1.0
        alpha2 = float(params[1]) if len(params) >= 2 else 1.0
        a = float(params[2]) if len(params) >= 3 else -10.0
        b = float(params[3]) if len(params) >= 4 else 100.0

        self.alpha1 = alpha1
        self.alpha2 = alpha2
        self.a = a
        self.b = b
        self._dist = _create_dist(alpha1, alpha2, a, b)
        self.support_low = a
        self.support_high = b
        self._raw_mean = float(self._dist.mean())
        self._raw_var = float(self._dist.var())
        self._raw_skew = float(self._dist.stats(moments="s"))
        self._raw_kurt = float(self._dist.stats(moments="k")) + 3.0
        self._raw_mode = self._find_mode()
        self._truncated_stats = None

        super().__init__([alpha1, alpha2, a, b], markers, func_name)

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

    def _find_mode(self) -> float:
        eps = max((self.b - self.a) * 1e-8, 1e-10)
        left = self.a + eps
        right = self.b - eps
        if right <= left:
            return float((self.a + self.b) / 2.0)

        def neg_pdf(x: float) -> float:
            return -float(self._dist.pdf(x))

        try:
            result = minimize_scalar(neg_pdf, bounds=(left, right), method="bounded")
            if result.success:
                return float(result.x)
        except Exception:
            pass

        grid = np.linspace(left, right, 4000)
        pdf_vals = self._dist.pdf(grid)
        return float(grid[int(np.argmax(pdf_vals))])

    def _original_pdf(self, x: float) -> float:
        return float(self._dist.pdf(x))

    def _original_cdf(self, x: float) -> float:
        return float(self._dist.cdf(x))

    def _original_ppf(self, q_prob: float) -> float:
        return float(self._dist.ppf(q_prob))

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
            return (x ** power) * self._original_pdf(x)

        result, _ = quad(integrand, lower, upper, limit=200)
        return float(result)

    def _compute_truncated_stats(self) -> None:
        if not self.is_truncated():
            self._truncated_stats = None
            return

        low_orig, high_orig = self._get_original_truncation_bounds()
        lb = self.a if low_orig is None else max(self.a, low_orig)
        ub = self.b if high_orig is None else min(self.b, high_orig)
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
        return self.a + self.shift_amount

    def max_val(self) -> float:
        if self._truncate_invalid:
            return float("nan")
        _, upper = self.get_effective_bounds()
        if upper is not None:
            return upper
        return self.b + self.shift_amount
