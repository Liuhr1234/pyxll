"""
Geometric distribution support for Drisk.
Definition used here:
    X = number of failures before the first success
    support = {0, 1, 2, ...}
    PMF = p * (1 - p)^x
"""
import math
from typing import List, Optional, Union
import numpy as np
from distribution_base import DistributionBase
_EPS = 1e-12
def _validate_params(p: float) -> None:
    if not (0.0 < p <= 1.0):
        raise ValueError(f"Geometric parameter p must satisfy 0 < p <= 1, got {p}")
def geomet_generator_single(rng: np.random.Generator, params: List[float]) -> float:
    p = float(params[0])
    _validate_params(p)
    return float(rng.geometric(p) - 1)
def geomet_generator_vectorized(
    rng: np.random.Generator, params: List[float], n_samples: int
) -> np.ndarray:
    p = params[0]
    if np.isscalar(p):
        p_arr = np.full(n_samples, float(p), dtype=float)
    else:
        p_arr = np.asarray(p, dtype=float)
        if p_arr.shape != (n_samples,):
            p_arr = np.broadcast_to(p_arr, (n_samples,))
    if np.any((p_arr <= 0.0) | (p_arr > 1.0)):
        raise ValueError("Geometric parameter p must satisfy 0 < p <= 1")
    return (rng.geometric(p_arr, size=n_samples) - 1).astype(float)
def geomet_generator(
    rng: np.random.Generator,
    params: List[float],
    n_samples: Optional[int] = None,
) -> Union[float, np.ndarray]:
    if n_samples is None:
        return geomet_generator_single(rng, params)
    return geomet_generator_vectorized(rng, params, n_samples)
def geomet_pmf(x: float, p: float) -> float:
    _validate_params(p)
    if x < 0 or abs(x - round(x)) > _EPS:
        return 0.0
    k = int(round(x))
    q = 1.0 - p
    return p * (q ** k)
def geomet_cdf(x: float, p: float) -> float:
    _validate_params(p)
    if x < 0:
        return 0.0
    if p >= 1.0:
        return 1.0
    k = int(math.floor(x))
    q = 1.0 - p
    return 1.0 - (q ** (k + 1))
def geomet_ppf(q_prob: float, p: float) -> float:
    _validate_params(p)
    if q_prob <= 0.0:
        return 0.0
    if q_prob >= 1.0:
        return float("inf")
    if p >= 1.0:
        return 0.0
    q = 1.0 - p
    result = math.ceil(math.log1p(-q_prob) / math.log(q)) - 1
    return float(max(0, result))
def geomet_raw_mean(p: float) -> float:
    return (1.0 / p) - 1.0
def geomet_raw_var(p: float) -> float:
    return (1.0 - p) / (p ** 2)
def geomet_raw_skew(p: float) -> float:
    q = 1.0 - p
    if q <= 0.0:
        return float("nan")
    return (2.0 - p) / math.sqrt(q)
def geomet_raw_kurt(p: float) -> float:
    q = 1.0 - p
    if q <= 0.0:
        return float("nan")
    return 9.0 + (p ** 2) / q
def geomet_raw_mode(p: float) -> float:
    _validate_params(p)
    return 0.0
def _tail_moment_sum(order: int, start: int, p: float) -> float:
    if start < 0:
        start = 0
    q = 1.0 - p
    q_to_start = q ** start
    if order == 0:
        return q_to_start
    if order == 1:
        return q_to_start * (start + q / p)
    if order == 2:
        return q_to_start * (
            start ** 2
            + (2.0 * start * q) / p
            + (q * (1.0 + q)) / (p ** 2)
        )
    if order == 3:
        return q_to_start * (
            start ** 3
            + (3.0 * (start ** 2) * q) / p
            + (3.0 * start * q * (1.0 + q)) / (p ** 2)
            + (q * (1.0 + 4.0 * q + q ** 2)) / (p ** 3)
        )
    if order == 4:
        return q_to_start * (
            start ** 4
            + (4.0 * (start ** 3) * q) / p
            + (6.0 * (start ** 2) * q * (1.0 + q)) / (p ** 2)
            + (4.0 * start * q * (1.0 + 4.0 * q + q ** 2)) / (p ** 3)
            + (q * (1.0 + 11.0 * q + 11.0 * q ** 2 + q ** 3)) / (p ** 4)
        )
    raise ValueError(f"Unsupported moment order: {order}")
def _interval_moment_sum(order: int, start: int, end: Optional[int], p: float) -> float:
    if end is not None and end < start:
        return 0.0
    tail_start = _tail_moment_sum(order, start, p)
    if end is None:
        return tail_start
    return tail_start - _tail_moment_sum(order, end + 1, p)
class GeometDistribution(DistributionBase):
    """Geometric distribution with exact discrete truncation handling."""
    def __init__(self, params: List[float], markers: dict = None, func_name: str = None):
        p = float(params[0]) if params else 0.5
        _validate_params(p)
        self.p = p
        self.q = 1.0 - p
        self.support_low = 0.0
        self.support_high = float("inf")
        self._raw_mean = geomet_raw_mean(p)
        self._raw_var = geomet_raw_var(p)
        self._raw_skew = geomet_raw_skew(p)
        self._raw_kurt = geomet_raw_kurt(p)
        self._raw_mode = geomet_raw_mode(p)
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
        self._window_cache = None
        self._truncated_stats = None
        if self.is_truncated():
            window = self._get_probability_window()
            if window is None or window["norm"] <= _EPS:
                self._truncate_invalid = True
            else:
                self._compute_truncated_stats()
    def _original_cdf(self, x: float) -> float:
        return geomet_cdf(x, self.p)
    def _original_ppf(self, q_prob: float) -> float:
        return geomet_ppf(q_prob, self.p)
    def _original_pdf(self, x: float) -> float:
        return geomet_pmf(x, self.p)
    def _raw_included_mass_up_to(
        self, k: int, lower_prob: float, upper_prob: float
    ) -> float:
        if k < 0:
            return 0.0
        cdf_k = self._original_cdf(float(k))
        return max(0.0, min(cdf_k, upper_prob) - lower_prob)
    def _raw_included_point_mass(
        self, k: int, lower_prob: float, upper_prob: float
    ) -> float:
        return self._raw_included_mass_up_to(
            k, lower_prob, upper_prob
        ) - self._raw_included_mass_up_to(k - 1, lower_prob, upper_prob)
    def _get_probability_window(self):
        if self._window_cache is not None:
            return self._window_cache
        if not self.is_truncated():
            self._window_cache = {
                "lower_prob": 0.0,
                "upper_prob": 1.0,
                "norm": 1.0,
                "start": 0,
                "end": None,
                "boundary_start": 0,
                "boundary_end": None,
            }
            return self._window_cache
        if self.truncate_type in ["value", "value2"]:
            low, high = self.get_truncated_bounds()
            if self.truncate_type == "value2":
                low_orig = low - self.shift_amount if low is not None else None
                high_orig = high - self.shift_amount if high is not None else None
            else:
                low_orig = low
                high_orig = high
            boundary_start = max(0, int(math.ceil(low_orig - _EPS))) if low_orig is not None else 0
            boundary_end = None
            if high_orig is not None and not math.isinf(high_orig):
                boundary_end = int(math.floor(high_orig + _EPS))
                if boundary_end < boundary_start:
                    self._window_cache = None
                    return None
            lower_prob = 0.0 if boundary_start <= 0 else self._original_cdf(boundary_start - 1)
            upper_prob = 1.0 if boundary_end is None else self._original_cdf(boundary_end)
        else:
            lower_prob = 0.0 if self.truncate_lower_pct is None else float(self.truncate_lower_pct)
            upper_prob = 1.0 if self.truncate_upper_pct is None else float(self.truncate_upper_pct)
            boundary_start = 0 if lower_prob <= 0.0 else int(self._original_ppf(lower_prob))
            boundary_end = None if upper_prob >= 1.0 else int(self._original_ppf(upper_prob))
        norm = upper_prob - lower_prob
        if norm <= _EPS:
            self._window_cache = None
            return None
        start = boundary_start
        while self._raw_included_point_mass(start, lower_prob, upper_prob) <= _EPS:
            start += 1
            if boundary_end is not None and start > boundary_end:
                self._window_cache = None
                return None
        end = boundary_end
        if end is not None:
            while end >= start and self._raw_included_point_mass(end, lower_prob, upper_prob) <= _EPS:
                end -= 1
            if end < start:
                self._window_cache = None
                return None
        self._window_cache = {
            "lower_prob": lower_prob,
            "upper_prob": upper_prob,
            "norm": norm,
            "start": start,
            "end": end,
            "boundary_start": boundary_start,
            "boundary_end": boundary_end,
        }
        return self._window_cache
    def _compute_truncated_stats(self):
        window = self._get_probability_window()
        if not window:
            self._truncated_stats = None
            return
        lower_prob = window["lower_prob"]
        upper_prob = window["upper_prob"]
        norm = window["norm"]
        boundary_start = window["boundary_start"]
        boundary_end = window["boundary_end"]
        moments = []
        for order in range(5):
            total = _interval_moment_sum(order, boundary_start, boundary_end, self.p)
            if lower_prob > 0.0:
                prev_cdf = self._original_cdf(boundary_start - 1)
                removed_lower = max(0.0, lower_prob - prev_cdf)
                total -= removed_lower * (boundary_start ** order)
            if boundary_end is not None:
                prev_cdf = self._original_cdf(boundary_end - 1)
                kept_upper = max(0.0, upper_prob - prev_cdf)
                removed_upper = max(0.0, self._original_pdf(boundary_end) - kept_upper)
                total -= removed_upper * (boundary_end ** order)
            moments.append(total / norm)
        m1, m2, m3, m4 = moments[1], moments[2], moments[3], moments[4]
        var_trunc = max(0.0, m2 - m1 ** 2)
        if var_trunc > _EPS:
            mu3 = m3 - 3.0 * m1 * m2 + 2.0 * (m1 ** 3)
            mu4 = m4 - 4.0 * m1 * m3 + 6.0 * (m1 ** 2) * m2 - 3.0 * (m1 ** 4)
            skew_trunc = mu3 / (var_trunc ** 1.5)
            kurt_trunc = mu4 / (var_trunc ** 2)
        else:
            skew_trunc = 0.0
            kurt_trunc = 3.0
        self._truncated_stats = {
            "mean": m1,
            "variance": var_trunc,
            "skewness": skew_trunc,
            "kurtosis": kurt_trunc,
            "mode": float(window["start"]),
        }
    def ppf(self, q_prob: float) -> float:
        if q_prob < 0.0 or q_prob > 1.0:
            return float("nan")
        if self._truncate_invalid:
            return float("nan")
        if not self.is_truncated():
            return self.apply_shift(self._original_ppf(q_prob))
        window = self._get_probability_window()
        if not window:
            return float("nan")
        if q_prob <= 0.0:
            return self.apply_shift(float(window["start"]))
        if q_prob >= 1.0:
            if window["end"] is None:
                return float("inf")
            return self.apply_shift(float(window["end"]))
        target_prob = window["lower_prob"] + q_prob * window["norm"]
        return self.apply_shift(self._original_ppf(target_prob))
    def cdf(self, x: float) -> float:
        if self._truncate_invalid:
            return float("nan")
        x_orig = self.apply_unshift(x)
        if not self.is_truncated():
            return self._original_cdf(x_orig)
        window = self._get_probability_window()
        if not window:
            return float("nan")
        raw_mass = self._raw_included_mass_up_to(
            int(math.floor(x_orig)), window["lower_prob"], window["upper_prob"]
        )
        return max(0.0, min(1.0, raw_mass / window["norm"]))
    def pdf(self, x: float) -> float:
        if self._truncate_invalid:
            return float("nan")
        x_orig = self.apply_unshift(x)
        if abs(x_orig - round(x_orig)) > _EPS:
            return 0.0
        if not self.is_truncated():
            return self._original_pdf(x_orig)
        window = self._get_probability_window()
        if not window:
            return float("nan")
        k = int(round(x_orig))
        raw_mass = self._raw_included_point_mass(
            k, window["lower_prob"], window["upper_prob"]
        )
        return raw_mass / window["norm"]
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
        return self._raw_mode + self.shift_amount
    def min_val(self) -> float:
        if self._truncate_invalid:
            return float("nan")
        lower, _ = self.get_effective_bounds()
        if lower is not None:
            return lower
        return self.apply_shift(0.0)
    def max_val(self) -> float:
        if self._truncate_invalid:
            return float("nan")
        _, upper = self.get_effective_bounds()
        if upper is not None:
            return upper
        return float("inf")
