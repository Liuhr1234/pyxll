"""
FatigueLife (Birnbaum-Saunders) distribution support for Drisk.

Definition used here:
    X ~ FatigueLife(Y, beta, alpha)
    support = [Y, +inf)
    scale beta > 0
    shape alpha > 0
"""

import math
from typing import List, Optional, Union

import numpy as np
import scipy.stats as sps
from scipy.optimize import minimize_scalar

from distribution_base import DistributionBase

_EPS = 1e-12


def _validate_params(y: float, beta: float, alpha: float) -> None:
    if beta <= 0.0:
        raise ValueError(f"beta must be > 0, got {beta}")
    if alpha <= 0.0:
        raise ValueError(f"alpha must be > 0, got {alpha}")


def fatiguelife_generator_single(rng: np.random.Generator, params: List[float]) -> float:
    y = float(params[0])
    beta = float(params[1])
    alpha = float(params[2])
    _validate_params(y, beta, alpha)

    z = float(rng.normal())
    t = 0.5 * alpha * z
    return float(y + beta * (t + math.sqrt(t * t + 1.0)) ** 2)


def fatiguelife_generator_vectorized(
    rng: np.random.Generator, params: List[float], n_samples: int
) -> np.ndarray:
    y = float(params[0])
    beta = float(params[1])
    alpha = float(params[2])
    _validate_params(y, beta, alpha)

    z = rng.normal(size=n_samples)
    t = 0.5 * alpha * z
    return (y + beta * (t + np.sqrt(t * t + 1.0)) ** 2).astype(float)


def fatiguelife_generator(
    rng: np.random.Generator,
    params: List[float],
    n_samples: Optional[int] = None,
) -> Union[float, np.ndarray]:
    if n_samples is None:
        return fatiguelife_generator_single(rng, params)
    return fatiguelife_generator_vectorized(rng, params, n_samples)


def fatiguelife_pdf(x: float, y: float, beta: float, alpha: float) -> float:
    _validate_params(y, beta, alpha)
    return float(sps.fatiguelife.pdf(x, alpha, loc=y, scale=beta))


def fatiguelife_cdf(x: float, y: float, beta: float, alpha: float) -> float:
    _validate_params(y, beta, alpha)
    return float(sps.fatiguelife.cdf(x, alpha, loc=y, scale=beta))


def fatiguelife_ppf(q_prob: float, y: float, beta: float, alpha: float) -> float:
    _validate_params(y, beta, alpha)
    if q_prob <= 0.0:
        return float(y)
    if q_prob >= 1.0:
        return float("inf")
    q = max(_EPS, min(1.0 - _EPS, float(q_prob)))
    return float(sps.fatiguelife.ppf(q, alpha, loc=y, scale=beta))


def fatiguelife_raw_mean(y: float, beta: float, alpha: float) -> float:
    _validate_params(y, beta, alpha)
    return float(y + beta * (1.0 + (alpha * alpha) / 2.0))


def fatiguelife_raw_var(beta: float, alpha: float) -> float:
    return float((alpha * beta) ** 2 * (1.0 + 1.25 * alpha * alpha))


def fatiguelife_raw_mode(y: float, beta: float, alpha: float) -> float:
    _validate_params(y, beta, alpha)
    dist = sps.fatiguelife(alpha, loc=y, scale=beta)
    lower = float(y)
    upper = float(dist.ppf(1.0 - 1e-9))
    if not math.isfinite(upper) or upper <= lower:
        upper = y + beta * max(10.0, 20.0 * alpha * alpha)

    def objective(x: float) -> float:
        return -float(dist.pdf(x))

    try:
        result = minimize_scalar(objective, bounds=(lower + 1e-12, upper), method="bounded")
        if result.success and math.isfinite(result.x):
            return float(result.x)
    except Exception:
        pass
    return float(y)


class FatigueLifeDistribution(DistributionBase):
    """FatigueLife theoretical distribution with truncation support."""

    def __init__(self, params: List[float], markers: dict = None, func_name: str = None):
        if len(params) >= 3:
            y = float(params[0])
            beta = float(params[1])
            alpha = float(params[2])
        else:
            y, beta, alpha = 0.0, 10.0, 2.0
        _validate_params(y, beta, alpha)

        self.y = y
        self.beta = beta
        self.alpha = alpha
        self._dist = sps.fatiguelife(alpha, loc=y, scale=beta)

        self.support_low = float(y)
        self.support_high = float("inf")
        mean_val, var_val, skew_val, kurt_excess = self._dist.stats(moments="mvsk")
        self._raw_mean = float(mean_val)
        self._raw_var = float(var_val)
        self._raw_skew = float(skew_val)
        self._raw_kurt = float(kurt_excess) + 3.0
        self._raw_mode = fatiguelife_raw_mode(y, beta, alpha)
        self._truncated_stats = None

        super().__init__([y, beta, alpha], markers, func_name)

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
        return fatiguelife_cdf(x, self.y, self.beta, self.alpha)

    def _original_ppf(self, q_prob: float) -> float:
        return fatiguelife_ppf(q_prob, self.y, self.beta, self.alpha)

    def _original_pdf(self, x: float) -> float:
        return fatiguelife_pdf(x, self.y, self.beta, self.alpha)

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

    def _compute_truncated_stats(self):
        if not self.is_truncated():
            self._truncated_stats = None
            return

        low_orig, high_orig = self._get_original_truncation_bounds()
        lb = self.support_low if low_orig is None else max(self.support_low, low_orig)
        ub = float("inf") if high_orig is None else high_orig

        try:
            m1 = self._dist.expect(lambda x: x, lb=lb, ub=ub, conditional=True)
            m2 = self._dist.expect(lambda x: x**2, lb=lb, ub=ub, conditional=True)
            m3 = self._dist.expect(lambda x: x**3, lb=lb, ub=ub, conditional=True)
            m4 = self._dist.expect(lambda x: x**4, lb=lb, ub=ub, conditional=True)
        except Exception:
            self._truncated_stats = None
            return

        variance = max(0.0, m2 - m1**2)
        if variance > _EPS:
            skewness = (m3 - 3 * m1 * variance - m1**3) / (variance ** 1.5)
            kurtosis = (m4 - 4 * m1 * m3 + 6 * m1 * m1 * m2 - 3 * (m1 ** 4)) / (variance ** 2)
        else:
            skewness = 0.0
            kurtosis = 3.0

        mode_raw = self._raw_mode
        if low_orig is not None and mode_raw < low_orig:
            mode_raw = low_orig
        if high_orig is not None and mode_raw > high_orig:
            mode_raw = high_orig

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
        return self.apply_shift(self.support_low)

    def max_val(self) -> float:
        if self._truncate_invalid:
            return float("nan")
        _, upper = self.get_effective_bounds()
        if upper is not None:
            return upper
        return self.apply_shift(self.support_high)
