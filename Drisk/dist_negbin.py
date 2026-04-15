"""
Negative binomial distribution support for Drisk.

Definition used here:
    X = number of failures before the s-th success
    support = {0, 1, 2, ...}
    PMF = C(s + x - 1, x) * p^s * (1 - p)^x

Parameters:
    s: required number of successes, must be a positive integer
    p: success probability per trial, must satisfy 0 < p <= 1
"""

import math
from typing import List, Optional, Union

import numpy as np
import scipy.stats as sps

from distribution_base import DistributionBase

_EPS = 1e-12


def _to_int(name: str, value: float) -> int:
    num = float(value)
    if abs(num - round(num)) > 1e-9:
        raise ValueError(f"{name} must be an integer, got {value}")
    return int(round(num))


def _validate_params(s: int, p: float) -> None:
    if s <= 0:
        raise ValueError(f"s must be a positive integer, got {s}")
    if not (0.0 < p <= 1.0):
        raise ValueError(f"p must satisfy 0 < p <= 1, got {p}")


def negbin_generator_single(rng: np.random.Generator, params: List[float]) -> float:
    s = _to_int("s", params[0])
    p = float(params[1])
    _validate_params(s, p)
    return float(rng.negative_binomial(s, p))


def negbin_generator_vectorized(
    rng: np.random.Generator, params: List[float], n_samples: int
) -> np.ndarray:
    s = params[0]
    p = params[1]

    if np.isscalar(s):
        s_arr = np.full(n_samples, _to_int("s", s), dtype=int)
    else:
        s_raw = np.asarray(s, dtype=float)
        if s_raw.shape != (n_samples,):
            s_raw = np.broadcast_to(s_raw, (n_samples,))
        if np.any(np.abs(s_raw - np.round(s_raw)) > 1e-9):
            raise ValueError(f"s must be an integer, got {s}")
        s_arr = np.round(s_raw).astype(int)

    if np.isscalar(p):
        p_arr = np.full(n_samples, float(p), dtype=float)
    else:
        p_arr = np.asarray(p, dtype=float)
        if p_arr.shape != (n_samples,):
            p_arr = np.broadcast_to(p_arr, (n_samples,))

    if np.any(s_arr <= 0):
        raise ValueError("s must be a positive integer")
    if np.any((p_arr <= 0.0) | (p_arr > 1.0)):
        raise ValueError("p must satisfy 0 < p <= 1")

    return rng.negative_binomial(s_arr, p_arr, size=n_samples).astype(float)


def negbin_generator(
    rng: np.random.Generator,
    params: List[float],
    n_samples: Optional[int] = None,
) -> Union[float, np.ndarray]:
    if n_samples is None:
        return negbin_generator_single(rng, params)
    return negbin_generator_vectorized(rng, params, n_samples)


def negbin_pmf(x: float, s: float, p: float) -> float:
    s_i = _to_int("s", s)
    p_f = float(p)
    _validate_params(s_i, p_f)
    if x < 0 or abs(x - round(x)) > _EPS:
        return 0.0
    k = int(round(x))
    if p_f >= 1.0 - _EPS:
        return 1.0 if k == 0 else 0.0
    return float(sps.nbinom.pmf(k, s_i, p_f))


def negbin_cdf(x: float, s: float, p: float) -> float:
    s_i = _to_int("s", s)
    p_f = float(p)
    _validate_params(s_i, p_f)
    if x < 0:
        return 0.0
    if p_f >= 1.0 - _EPS:
        return 1.0
    return float(sps.nbinom.cdf(math.floor(x), s_i, p_f))


def negbin_ppf(q_prob: float, s: float, p: float) -> float:
    s_i = _to_int("s", s)
    p_f = float(p)
    _validate_params(s_i, p_f)
    if q_prob <= 0.0:
        return 0.0
    if p_f >= 1.0 - _EPS:
        return 0.0
    if q_prob >= 1.0:
        return float("inf")
    return float(sps.nbinom.ppf(q_prob, s_i, p_f))


def negbin_raw_mean(s: float, p: float) -> float:
    s_i = _to_int("s", s)
    p_f = float(p)
    _validate_params(s_i, p_f)
    if p_f >= 1.0 - _EPS:
        return 0.0
    return s_i * (1.0 - p_f) / p_f


def negbin_raw_var(s: float, p: float) -> float:
    s_i = _to_int("s", s)
    p_f = float(p)
    _validate_params(s_i, p_f)
    if p_f >= 1.0 - _EPS:
        return 0.0
    return s_i * (1.0 - p_f) / (p_f ** 2)


def negbin_raw_skew(s: float, p: float) -> float:
    s_i = _to_int("s", s)
    p_f = float(p)
    _validate_params(s_i, p_f)
    if p_f >= 1.0 - _EPS:
        return 0.0
    return (2.0 - p_f) / math.sqrt(s_i * (1.0 - p_f))


def negbin_raw_kurt(s: float, p: float) -> float:
    s_i = _to_int("s", s)
    p_f = float(p)
    _validate_params(s_i, p_f)
    if p_f >= 1.0 - _EPS:
        return 3.0
    return 3.0 + 6.0 / s_i + (p_f ** 2) / (s_i * (1.0 - p_f))


def negbin_raw_mode(s: float, p: float) -> float:
    s_i = _to_int("s", s)
    p_f = float(p)
    _validate_params(s_i, p_f)
    if p_f >= 1.0 - _EPS or s_i <= 1:
        return 0.0
    z = (s_i - 1.0) * (1.0 - p_f) / p_f
    if z < 0.0:
        return 0.0
    if abs(z - round(z)) <= _EPS:
        return float("nan")
    return float(math.floor(z))


class NegbinDistribution(DistributionBase):
    """Negative binomial distribution with exact discrete truncation handling."""

    def __init__(self, params: List[float], markers: dict = None, func_name: str = None):
        if len(params) >= 2:
            s = _to_int("s", params[0])
            p = float(params[1])
        else:
            s, p = 1, 0.5
        _validate_params(s, p)

        self.s = s
        self.p = p
        self.support_low = 0.0
        self.support_high = 0.0 if p >= 1.0 - _EPS else float("inf")

        self._raw_mean = negbin_raw_mean(s, p)
        self._raw_var = negbin_raw_var(s, p)
        self._raw_skew = negbin_raw_skew(s, p)
        self._raw_kurt = negbin_raw_kurt(s, p)
        self._raw_mode = negbin_raw_mode(s, p)

        super().__init__([float(s), float(p)], markers, func_name)

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
        self._truncated_point_probs = None
        self._compute_truncated_stats()

    def _original_cdf(self, x: float) -> float:
        return negbin_cdf(x, self.s, self.p)

    def _original_ppf(self, q_prob: float) -> float:
        return negbin_ppf(q_prob, self.s, self.p)

    def _original_pdf(self, x: float) -> float:
        return negbin_pmf(x, self.s, self.p)

    def _find_boundary_info(self, pct: float) -> tuple[int, float]:
        pct = max(0.0, min(1.0, float(pct)))
        if pct <= 0.0:
            return 0, 0.0
        if self.p >= 1.0 - _EPS:
            return 0, 1.0
        idx = int(self._original_ppf(pct))
        point_prob = self._original_pdf(float(idx))
        cdf_before = self._original_cdf(float(idx - 1)) if idx > 0 else 0.0
        if point_prob <= _EPS:
            return idx, 0.0
        fraction = (pct - cdf_before) / point_prob
        return idx, max(0.0, min(1.0, fraction))

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

    def _get_point_keep_fraction(
        self,
        k: int,
        lower_idx: Optional[int],
        lower_fraction: float,
        upper_idx: Optional[int],
        upper_fraction: float,
    ) -> float:
        if lower_idx is not None and upper_idx is not None and lower_idx == upper_idx == k:
            return max(0.0, upper_fraction - lower_fraction)

        keep_fraction = 1.0
        if lower_idx is not None:
            if k < lower_idx:
                return 0.0
            if k == lower_idx:
                keep_fraction *= (1.0 - lower_fraction)

        if upper_idx is not None:
            if k > upper_idx:
                return 0.0
            if k == upper_idx:
                keep_fraction *= upper_fraction

        return max(0.0, min(1.0, keep_fraction))

    def _enumeration_upper_limit(self) -> int:
        if self.p >= 1.0 - _EPS:
            return 0
        approx = self._original_ppf(1.0 - 1e-12)
        if math.isinf(approx):
            approx = self._original_ppf(1.0 - 1e-10)
        return max(0, int(approx))

    def _build_truncated_point_probs(self):
        if self._truncate_invalid:
            return None
        if not self.is_truncated():
            return None
        if self._truncated_point_probs is not None:
            return self._truncated_point_probs

        low_orig, high_orig = self._get_original_truncation_bounds()
        start_k = 0 if low_orig is None else max(0, int(math.ceil(low_orig - _EPS)))

        lower_idx = None
        upper_idx = None
        lower_fraction = 0.0
        upper_fraction = 1.0

        if self.truncate_type in ["percentile", "percentile2"]:
            if self.truncate_lower_pct is not None and self.truncate_lower_pct > 0.0:
                lower_idx, lower_fraction = self._find_boundary_info(self.truncate_lower_pct)
            if self.truncate_upper_pct is not None and self.truncate_upper_pct < 1.0:
                upper_idx, upper_fraction = self._find_boundary_info(self.truncate_upper_pct)

        end_k = self._enumeration_upper_limit()
        if high_orig is not None and not math.isinf(high_orig):
            end_k = min(end_k, int(math.floor(high_orig + _EPS)))
        if upper_idx is not None:
            end_k = min(end_k, upper_idx)
        if end_k < start_k:
            self._truncate_invalid = True
            self._truncated_point_probs = []
            return None

        point_probs = []
        total_prob = 0.0
        for k in range(start_k, end_k + 1):
            x = float(k)
            if low_orig is not None and x < low_orig - _EPS:
                continue
            if high_orig is not None and not math.isinf(high_orig) and x > high_orig + _EPS:
                continue

            prob = self._original_pdf(x)
            if self.truncate_type in ["percentile", "percentile2"]:
                prob *= self._get_point_keep_fraction(
                    k, lower_idx, lower_fraction, upper_idx, upper_fraction
                )

            if prob > _EPS:
                point_probs.append((x, prob))
                total_prob += prob

        if total_prob <= _EPS:
            self._truncate_invalid = True
            self._truncated_point_probs = []
            return None

        self._truncated_point_probs = [(x, prob / total_prob) for x, prob in point_probs]
        return self._truncated_point_probs

    def _compute_truncated_stats(self):
        if not self.is_truncated():
            return

        point_probs = self._build_truncated_point_probs()
        if not point_probs:
            self._truncated_stats = None
            return

        values = [x for x, _ in point_probs]
        probs = [p for _, p in point_probs]

        mean_trunc = sum(v * p for v, p in zip(values, probs))
        var_trunc = sum(((v - mean_trunc) ** 2) * p for v, p in zip(values, probs))

        if var_trunc > _EPS:
            skew_trunc = sum(((v - mean_trunc) ** 3) * p for v, p in zip(values, probs)) / (var_trunc ** 1.5)
            kurt_trunc = sum(((v - mean_trunc) ** 4) * p for v, p in zip(values, probs)) / (var_trunc ** 2)
        else:
            skew_trunc = 0.0
            kurt_trunc = 3.0

        max_prob = max(probs)
        mode_values = [v for v, prob in zip(values, probs) if abs(prob - max_prob) <= _EPS]
        if len(mode_values) == 1:
            mode_trunc = mode_values[0]
        else:
            mode_trunc = float("nan")

        self._truncated_stats = {
            "mean": mean_trunc,
            "variance": var_trunc,
            "skewness": skew_trunc,
            "kurtosis": kurt_trunc,
            "mode": mode_trunc,
        }

    def ppf(self, q: float) -> float:
        if q < 0.0 or q > 1.0:
            return float("nan")
        if self._truncate_invalid:
            return float("nan")
        if not self.is_truncated():
            return self.apply_shift(self._original_ppf(q))

        point_probs = self._build_truncated_point_probs()
        if not point_probs:
            return float("nan")
        if q <= 0.0:
            return self.apply_shift(point_probs[0][0])
        if q >= 1.0:
            return self.apply_shift(point_probs[-1][0])

        cumulative = 0.0
        for value, prob in point_probs:
            cumulative += prob
            if cumulative >= q - _EPS:
                return self.apply_shift(value)
        return self.apply_shift(point_probs[-1][0])

    def cdf(self, x: float) -> float:
        if self._truncate_invalid:
            return float("nan")
        x_orig = self.apply_unshift(x)
        if not self.is_truncated():
            return self._original_cdf(x_orig)

        point_probs = self._build_truncated_point_probs()
        if not point_probs:
            return float("nan")
        if x_orig < point_probs[0][0] - _EPS:
            return 0.0
        if x_orig >= point_probs[-1][0] - _EPS:
            return 1.0

        total = 0.0
        for value, prob in point_probs:
            if value <= x_orig + _EPS:
                total += prob
        return min(1.0, total)

    def pdf(self, x: float) -> float:
        if self._truncate_invalid:
            return float("nan")
        x_orig = self.apply_unshift(x)
        if abs(x_orig - round(x_orig)) > _EPS:
            return 0.0
        if not self.is_truncated():
            return self._original_pdf(x_orig)

        point_probs = self._build_truncated_point_probs()
        if not point_probs:
            return float("nan")

        x_round = float(round(x_orig))
        for value, prob in point_probs:
            if abs(value - x_round) <= _EPS:
                return prob
        return 0.0

    def mean(self) -> float:
        if self._truncate_invalid:
            return float("nan")
        if self.is_truncated() and self._truncated_stats is None:
            self._compute_truncated_stats()
        if self._truncated_stats:
            return self._truncated_stats["mean"] + self.shift_amount
        return self._raw_mean + self.shift_amount

    def variance(self) -> float:
        if self._truncate_invalid:
            return float("nan")
        if self.is_truncated() and self._truncated_stats is None:
            self._compute_truncated_stats()
        if self._truncated_stats:
            return self._truncated_stats["variance"]
        return self._raw_var

    def skewness(self) -> float:
        if self._truncate_invalid:
            return float("nan")
        if self.is_truncated() and self._truncated_stats is None:
            self._compute_truncated_stats()
        if self._truncated_stats:
            return self._truncated_stats["skewness"]
        return self._raw_skew

    def kurtosis(self) -> float:
        if self._truncate_invalid:
            return float("nan")
        if self.is_truncated() and self._truncated_stats is None:
            self._compute_truncated_stats()
        if self._truncated_stats:
            return self._truncated_stats["kurtosis"]
        return self._raw_kurt

    def mode(self) -> float:
        if self._truncate_invalid:
            return float("nan")
        if self.is_truncated() and self._truncated_stats is None:
            self._compute_truncated_stats()
        if self._truncated_stats:
            mode_val = self._truncated_stats["mode"]
            return mode_val + self.shift_amount if not math.isnan(mode_val) else float("nan")
        return self._raw_mode + self.shift_amount if not math.isnan(self._raw_mode) else float("nan")

    def min_val(self) -> float:
        if self._truncate_invalid:
            return float("nan")
        if self.is_truncated():
            point_probs = self._build_truncated_point_probs()
            if point_probs:
                return self.apply_shift(point_probs[0][0])
        lower, _ = self.get_effective_bounds()
        if lower is not None:
            return lower
        return self.apply_shift(self.support_low)

    def max_val(self) -> float:
        if self._truncate_invalid:
            return float("nan")
        if self.is_truncated():
            _, upper = self.get_effective_bounds()
            if upper is None and self.support_high == float("inf"):
                return float("inf")
            point_probs = self._build_truncated_point_probs()
            if point_probs:
                return self.apply_shift(point_probs[-1][0])
        if self.support_high != float("inf"):
            return self.apply_shift(self.support_high)
        _, upper = self.get_effective_bounds()
        if upper is not None:
            return upper
        return float("inf")
