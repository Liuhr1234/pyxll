"""BetaSubj distribution support for Drisk."""

import math
from typing import List, Optional, Tuple

import numpy as np
from scipy.integrate import quad
from scipy.optimize import least_squares
from scipy.special import betainc, betaincinv
from scipy.stats import beta as scipy_beta

from distribution_base import DistributionBase

_EPS = 1e-12


def _span(min_val: float, max_val: float) -> float:
    return float(max_val) - float(min_val)


def _validate_input_params(min_val: float, m_likely: float, mean_val: float, max_val: float) -> None:
    if float(max_val) <= float(min_val):
        raise ValueError("BetaSubj requires Max > Min")
    if not (float(min_val) < float(mean_val) < float(max_val)):
        raise ValueError("BetaSubj requires Min < Mean < Max")
    if not (float(min_val) < float(m_likely) < float(max_val)):
        raise ValueError("BetaSubj requires Min < M.likely < Max")


def _validate_alpha(alpha1: float, alpha2: float) -> None:
    if float(alpha1) <= 0.0:
        raise ValueError("BetaSubj requires alpha1 > 0")
    if float(alpha2) <= 0.0:
        raise ValueError("BetaSubj requires alpha2 > 0")


def _standardize(x: float, min_val: float, max_val: float) -> float:
    return (float(x) - float(min_val)) / _span(min_val, max_val)


def _solve_alpha_params(min_val: float, m_likely: float, mean_val: float, max_val: float) -> Tuple[float, float]:
    _validate_input_params(min_val, m_likely, mean_val, max_val)

    p_mean = _standardize(mean_val, min_val, max_val)
    p_mode = _standardize(m_likely, min_val, max_val)

    if abs(p_mode - p_mean) <= 1e-12:
        if abs(p_mean - 0.5) > 1e-12:
            raise ValueError("BetaSubj requires Mean=(Min+Max)/2 when M.likely=Mean")
        return 2.0, 2.0

    alpha1 = float("nan")
    alpha2 = float("nan")

    try:
        s = (2.0 * p_mode - 1.0) / (p_mode - p_mean)
        alpha1 = p_mean * s
        alpha2 = (1.0 - p_mean) * s
    except Exception:
        pass

    if math.isfinite(alpha1) and math.isfinite(alpha2) and alpha1 > 1.0 and alpha2 > 1.0:
        implied_mean = alpha1 / (alpha1 + alpha2)
        implied_mode = (alpha1 - 1.0) / (alpha1 + alpha2 - 2.0)
        if abs(implied_mean - p_mean) <= 1e-10 and abs(implied_mode - p_mode) <= 1e-10:
            return float(alpha1), float(alpha2)

    def residuals(params: np.ndarray) -> np.ndarray:
        a1, a2 = float(params[0]), float(params[1])
        mean_res = a1 / (a1 + a2) - p_mean
        if a1 > 1.0 and a2 > 1.0:
            mode_res = (a1 - 1.0) / (a1 + a2 - 2.0) - p_mode
        else:
            mode_res = 10.0
        return np.asarray([mean_res, mode_res], dtype=float)

    guess = np.asarray([max(1.1, p_mean * 4.0), max(1.1, (1.0 - p_mean) * 4.0)], dtype=float)
    result = least_squares(
        residuals,
        guess,
        bounds=([1.0 + 1e-9, 1.0 + 1e-9], [1e6, 1e6]),
        xtol=1e-14,
        ftol=1e-14,
        gtol=1e-14,
    )

    alpha1 = float(result.x[0])
    alpha2 = float(result.x[1])
    _validate_alpha(alpha1, alpha2)

    implied_mean = alpha1 / (alpha1 + alpha2)
    implied_mode = (alpha1 - 1.0) / (alpha1 + alpha2 - 2.0)
    if abs(implied_mean - p_mean) > 1e-8 or abs(implied_mode - p_mode) > 1e-8:
        raise ValueError("BetaSubj parameters are inconsistent")

    return alpha1, alpha2


def _create_dist(min_val: float, max_val: float, alpha1: float, alpha2: float):
    _validate_input_params(min_val, (min_val + max_val) / 2.0, (min_val + max_val) / 2.0, max_val)
    _validate_alpha(alpha1, alpha2)
    return scipy_beta(a=float(alpha1), b=float(alpha2), loc=float(min_val), scale=_span(min_val, max_val))


def betasubj_pdf(x: float, min_val: float, m_likely: float, mean_val: float, max_val: float) -> float:
    alpha1, alpha2 = _solve_alpha_params(min_val, m_likely, mean_val, max_val)
    x = float(x)
    if x < float(min_val) or x > float(max_val):
        return 0.0
    dist = scipy_beta(a=float(alpha1), b=float(alpha2), loc=float(min_val), scale=_span(min_val, max_val))
    return float(dist.pdf(x))


def betasubj_cdf(x: float, min_val: float, m_likely: float, mean_val: float, max_val: float) -> float:
    x = float(x)
    if x <= float(min_val):
        return 0.0
    if x >= float(max_val):
        return 1.0
    alpha1, alpha2 = _solve_alpha_params(min_val, m_likely, mean_val, max_val)
    z = _standardize(x, min_val, max_val)
    return float(betainc(alpha1, alpha2, z))


def betasubj_ppf(q_prob: float, min_val: float, m_likely: float, mean_val: float, max_val: float) -> float:
    q_prob = float(q_prob)
    if q_prob <= 0.0:
        return float(min_val)
    if q_prob >= 1.0:
        return float(max_val)
    alpha1, alpha2 = _solve_alpha_params(min_val, m_likely, mean_val, max_val)
    z = float(betaincinv(alpha1, alpha2, q_prob))
    return float(float(min_val) + _span(min_val, max_val) * z)


def betasubj_raw_mean(min_val: float, m_likely: float, mean_val: float, max_val: float) -> float:
    _solve_alpha_params(min_val, m_likely, mean_val, max_val)
    return float(mean_val)


def betasubj_generator_single(rng: np.random.Generator, params: List[float]) -> float:
    min_val = float(params[0])
    m_likely = float(params[1])
    mean_val = float(params[2])
    max_val = float(params[3])
    alpha1, alpha2 = _solve_alpha_params(min_val, m_likely, mean_val, max_val)
    return float(min_val + _span(min_val, max_val) * rng.beta(alpha1, alpha2))


def betasubj_generator_vectorized(
    rng: np.random.Generator,
    params: List[float],
    n_samples: int,
) -> np.ndarray:
    min_val = float(params[0])
    m_likely = float(params[1])
    mean_val = float(params[2])
    max_val = float(params[3])
    alpha1, alpha2 = _solve_alpha_params(min_val, m_likely, mean_val, max_val)
    return np.asarray(min_val + _span(min_val, max_val) * rng.beta(alpha1, alpha2, size=n_samples), dtype=float)


class BetaSubjDistribution(DistributionBase):
    """BetaSubj theoretical distribution with truncation support."""

    def __init__(self, params: List[float], markers: dict = None, func_name: str = None):
        markers = markers or {}
        min_val = float(params[0]) if len(params) >= 1 else -10.0
        m_likely = float(params[1]) if len(params) >= 2 else 6.0
        mean_val = float(params[2]) if len(params) >= 3 else 2.0
        max_val = float(params[3]) if len(params) >= 4 else 10.0

        alpha1, alpha2 = _solve_alpha_params(min_val, m_likely, mean_val, max_val)

        self.min_param = min_val
        self.m_likely = m_likely
        self.mean_input = mean_val
        self.max_param = max_val
        self.alpha1 = float(alpha1)
        self.alpha2 = float(alpha2)
        self.range_param = _span(min_val, max_val)
        self._dist = scipy_beta(a=self.alpha1, b=self.alpha2, loc=min_val, scale=self.range_param)
        self.support_low = min_val
        self.support_high = max_val

        alpha_sum = self.alpha1 + self.alpha2
        self._raw_mean = float(min_val + self.range_param * self.alpha1 / alpha_sum)
        self._raw_var = float(
            (self.range_param ** 2) * self.alpha1 * self.alpha2 / ((alpha_sum ** 2) * (alpha_sum + 1.0))
        )
        self._raw_skew = float(
            2.0 * (self.alpha2 - self.alpha1) / (alpha_sum + 2.0) * math.sqrt((alpha_sum + 1.0) / (self.alpha1 * self.alpha2))
        )
        self._raw_kurt = float(
            3.0
            * ((alpha_sum + 1.0) * (2.0 * (alpha_sum ** 2) + self.alpha1 * self.alpha2 * (alpha_sum - 6.0)))
            / (self.alpha1 * self.alpha2 * (alpha_sum + 2.0) * (alpha_sum + 3.0))
        )
        self._raw_mode = float(m_likely)
        self._truncated_stats = None

        super().__init__([min_val, m_likely, mean_val, max_val], markers, func_name)

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
        x = float(x)
        if x < self.min_param or x > self.max_param:
            return 0.0
        return float(self._dist.pdf(x))

    def _original_cdf(self, x: float) -> float:
        x = float(x)
        if x <= self.min_param:
            return 0.0
        if x >= self.max_param:
            return 1.0
        return float(self._dist.cdf(x))

    def _original_ppf(self, q_prob: float) -> float:
        q_prob = float(q_prob)
        if q_prob <= 0.0:
            return float(self.min_param)
        if q_prob >= 1.0:
            return float(self.max_param)
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
