"""
Hypergeometric distribution support for Drisk.

Definition used here:
    X = number of target items drawn in n draws without replacement
    from a population of size M containing D target items.

Parameters:
    n: sample size
    D: number of target items in the population
    M: population size
"""

import math
from functools import lru_cache
from typing import List, Optional, Union

import numpy as np

from distribution_base import DistributionBase

_EPS = 1e-12


def _to_int(name: str, value: float) -> int:
    num = float(value)
    if not num.is_integer():
        raise ValueError(f"{name} must be an integer, got {value}")
    return int(num)


def _validate_params(n: int, D: int, M: int) -> None:
    if n <= 0:
        raise ValueError(f"n must be a positive integer, got {n}")
    if D <= 0:
        raise ValueError(f"D must be a positive integer, got {D}")
    if M <= 0:
        raise ValueError(f"M must be a positive integer, got {M}")
    if n > M:
        raise ValueError(f"n must satisfy n <= M, got n={n}, M={M}")
    if D > M:
        raise ValueError(f"D must satisfy D <= M, got D={D}, M={M}")


def _support_bounds(n: int, D: int, M: int) -> tuple[int, int]:
    return max(0, n + D - M), min(n, D)


@lru_cache(maxsize=256)
def _build_support_and_probs(n: int, D: int, M: int) -> tuple[tuple[int, ...], tuple[float, ...]]:
    _validate_params(n, D, M)
    low, high = _support_bounds(n, D, M)
    xs = tuple(range(low, high + 1))
    denom = math.comb(M, n)
    probs = []
    for x in xs:
        prob = (
            math.comb(D, x)
            * math.comb(M - D, n - x)
            / denom
        )
        probs.append(float(prob))
    total = sum(probs)
    if total <= 0.0:
        raise ValueError("Invalid hypergeometric probability table")
    probs = [p / total for p in probs]
    return xs, tuple(probs)


def _mode_value(n: int, D: int, M: int, xs: tuple[int, ...], probs: tuple[float, ...]) -> float:
    x_m = ((n + 1) * (D + 1)) / (M + 2)
    if abs(x_m - round(x_m)) < 1e-12:
        upper_mode = int(round(x_m))
        lower_mode = upper_mode - 1
        valid = [x for x in (lower_mode, upper_mode) if x in xs]
        if valid:
            return float(sum(valid) / len(valid))
    floor_mode = math.floor(x_m)
    if floor_mode in xs:
        return float(floor_mode)
    max_prob = max(probs)
    modes = [x for x, p in zip(xs, probs) if abs(p - max_prob) <= 1e-12]
    return float(sum(modes) / len(modes))


@lru_cache(maxsize=256)
def _raw_stats(n: int, D: int, M: int) -> dict:
    xs, probs = _build_support_and_probs(n, D, M)
    mean = n * D / M
    if M > 1:
        variance = n * D * (M - D) * (M - n) / (M * M * (M - 1))
    else:
        variance = 0.0

    if variance > 0.0:
        mu3 = sum(((x - mean) ** 3) * p for x, p in zip(xs, probs))
        mu4 = sum(((x - mean) ** 4) * p for x, p in zip(xs, probs))
        skewness = mu3 / (variance ** 1.5)
        kurtosis = mu4 / (variance ** 2)
    else:
        skewness = 0.0
        kurtosis = 3.0

    return {
        "mean": float(mean),
        "variance": float(variance),
        "skewness": float(skewness),
        "kurtosis": float(kurtosis),
        "mode": _mode_value(n, D, M, xs, probs),
    }


def hypergeo_generator_single(rng: np.random.Generator, params: List[float]) -> float:
    n = _to_int("n", params[0])
    D = _to_int("D", params[1])
    M = _to_int("M", params[2])
    _validate_params(n, D, M)
    return float(rng.hypergeometric(D, M - D, n))


def hypergeo_generator_vectorized(
    rng: np.random.Generator, params: List[float], n_samples: int
) -> np.ndarray:
    n = params[0]
    D = params[1]
    M = params[2]

    if np.isscalar(n):
        n_arr = np.full(n_samples, _to_int("n", n), dtype=int)
    else:
        n_raw = np.asarray(n, dtype=float)
        if n_raw.shape != (n_samples,):
            n_raw = np.broadcast_to(n_raw, (n_samples,))
        if np.any(np.abs(n_raw - np.round(n_raw)) > 1e-9):
            raise ValueError(f"n must be an integer, got {n}")
        n_arr = np.round(n_raw).astype(int)

    if np.isscalar(D):
        D_arr = np.full(n_samples, _to_int("D", D), dtype=int)
    else:
        D_raw = np.asarray(D, dtype=float)
        if D_raw.shape != (n_samples,):
            D_raw = np.broadcast_to(D_raw, (n_samples,))
        if np.any(np.abs(D_raw - np.round(D_raw)) > 1e-9):
            raise ValueError(f"D must be an integer, got {D}")
        D_arr = np.round(D_raw).astype(int)

    if np.isscalar(M):
        M_arr = np.full(n_samples, _to_int("M", M), dtype=int)
    else:
        M_raw = np.asarray(M, dtype=float)
        if M_raw.shape != (n_samples,):
            M_raw = np.broadcast_to(M_raw, (n_samples,))
        if np.any(np.abs(M_raw - np.round(M_raw)) > 1e-9):
            raise ValueError(f"M must be an integer, got {M}")
        M_arr = np.round(M_raw).astype(int)

    if np.any(n_arr <= 0):
        raise ValueError("n must be a positive integer")
    if np.any(D_arr <= 0):
        raise ValueError("D must be a positive integer")
    if np.any(M_arr <= 0):
        raise ValueError("M must be a positive integer")
    if np.any(n_arr > M_arr):
        raise ValueError("n must satisfy n <= M")
    if np.any(D_arr > M_arr):
        raise ValueError("D must satisfy D <= M")

    return rng.hypergeometric(D_arr, M_arr - D_arr, n_arr, size=n_samples).astype(float)


def hypergeo_generator(
    rng: np.random.Generator,
    params: List[float],
    n_samples: Optional[int] = None,
) -> Union[float, np.ndarray]:
    if n_samples is None:
        return hypergeo_generator_single(rng, params)
    return hypergeo_generator_vectorized(rng, params, n_samples)


def hypergeo_pmf(x: float, n: float, D: float, M: float) -> float:
    n_i = _to_int("n", n)
    D_i = _to_int("D", D)
    M_i = _to_int("M", M)
    _validate_params(n_i, D_i, M_i)
    if abs(x - round(x)) > _EPS:
        return 0.0
    k = int(round(x))
    xs, probs = _build_support_and_probs(n_i, D_i, M_i)
    if k < xs[0] or k > xs[-1]:
        return 0.0
    return float(probs[k - xs[0]])


def hypergeo_cdf(x: float, n: float, D: float, M: float) -> float:
    n_i = _to_int("n", n)
    D_i = _to_int("D", D)
    M_i = _to_int("M", M)
    _validate_params(n_i, D_i, M_i)
    xs, probs = _build_support_and_probs(n_i, D_i, M_i)
    if x < xs[0]:
        return 0.0
    if x >= xs[-1]:
        return 1.0
    limit = int(math.floor(x))
    total = 0.0
    for xi, pi in zip(xs, probs):
        if xi <= limit:
            total += pi
        else:
            break
    return float(total)


def hypergeo_ppf(q_prob: float, n: float, D: float, M: float) -> float:
    n_i = _to_int("n", n)
    D_i = _to_int("D", D)
    M_i = _to_int("M", M)
    _validate_params(n_i, D_i, M_i)
    xs, probs = _build_support_and_probs(n_i, D_i, M_i)
    if q_prob <= 0.0:
        return float(xs[0])
    if q_prob >= 1.0:
        return float(xs[-1])
    cumulative = 0.0
    for xi, pi in zip(xs, probs):
        cumulative += pi
        if cumulative >= q_prob - _EPS:
            return float(xi)
    return float(xs[-1])


def hypergeo_raw_mean(n: float, D: float, M: float) -> float:
    n_i = _to_int("n", n)
    D_i = _to_int("D", D)
    M_i = _to_int("M", M)
    _validate_params(n_i, D_i, M_i)
    return _raw_stats(n_i, D_i, M_i)["mean"]


class HypergeoDistribution(DistributionBase):
    """Hypergeometric distribution with exact discrete truncation handling."""

    def __init__(self, params: List[float], markers: dict = None, func_name: str = None):
        if len(params) >= 3:
            n = _to_int("n", params[0])
            D = _to_int("D", params[1])
            M = _to_int("M", params[2])
        else:
            n, D, M = 10, 20, 30
        _validate_params(n, D, M)

        self.n = n
        self.D = D
        self.M = M
        self.support_low, self.support_high = _support_bounds(n, D, M)
        self.x_vals, self.p_vals = _build_support_and_probs(n, D, M)

        stats = _raw_stats(n, D, M)
        self._raw_mean = stats["mean"]
        self._raw_var = stats["variance"]
        self._raw_skew = stats["skewness"]
        self._raw_kurt = stats["kurtosis"]
        self._raw_mode = stats["mode"]

        super().__init__([float(n), float(D), float(M)], markers, func_name)

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
        return hypergeo_cdf(x, self.n, self.D, self.M)

    def _original_ppf(self, q_prob: float) -> float:
        return hypergeo_ppf(q_prob, self.n, self.D, self.M)

    def _original_pdf(self, x: float) -> float:
        return hypergeo_pmf(x, self.n, self.D, self.M)

    def _raw_included_mass_up_to(
        self, k: int, lower_prob: float, upper_prob: float
    ) -> float:
        if k < self.support_low:
            return 0.0
        cdf_k = self._original_cdf(float(k))
        return max(0.0, min(cdf_k, upper_prob) - lower_prob)

    def _raw_included_point_mass(
        self, k: int, lower_prob: float, upper_prob: float
    ) -> float:
        return (
            self._raw_included_mass_up_to(k, lower_prob, upper_prob)
            - self._raw_included_mass_up_to(k - 1, lower_prob, upper_prob)
        )

    def _get_probability_window(self):
        if self._window_cache is not None:
            return self._window_cache

        if not self.is_truncated():
            self._window_cache = {
                "lower_prob": 0.0,
                "upper_prob": 1.0,
                "norm": 1.0,
                "start": self.support_low,
                "end": self.support_high,
                "boundary_start": self.support_low,
                "boundary_end": self.support_high,
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

            boundary_start = self.support_low
            if low_orig is not None:
                boundary_start = max(self.support_low, int(math.ceil(low_orig - _EPS)))

            boundary_end = self.support_high
            if high_orig is not None:
                boundary_end = min(self.support_high, int(math.floor(high_orig + _EPS)))
                if boundary_end < boundary_start:
                    self._window_cache = None
                    return None

            lower_prob = (
                0.0 if boundary_start <= self.support_low
                else self._original_cdf(boundary_start - 1)
            )
            upper_prob = (
                1.0 if boundary_end >= self.support_high
                else self._original_cdf(boundary_end)
            )
        else:
            lower_prob = 0.0 if self.truncate_lower_pct is None else float(self.truncate_lower_pct)
            upper_prob = 1.0 if self.truncate_upper_pct is None else float(self.truncate_upper_pct)
            boundary_start = (
                self.support_low
                if lower_prob <= 0.0
                else int(self._original_ppf(lower_prob))
            )
            boundary_end = (
                self.support_high
                if upper_prob >= 1.0
                else int(self._original_ppf(upper_prob))
            )

        boundary_start = int(boundary_start)
        boundary_end = int(boundary_end)

        norm = upper_prob - lower_prob
        if norm <= _EPS:
            self._window_cache = None
            return None

        start = boundary_start
        while (
            start <= self.support_high
            and self._raw_included_point_mass(start, lower_prob, upper_prob) <= _EPS
        ):
            start += 1
            if start > boundary_end:
                self._window_cache = None
                return None

        end = boundary_end
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

        values = []
        probs = []
        for k in range(window["start"], window["end"] + 1):
            prob = self._raw_included_point_mass(k, lower_prob, upper_prob)
            if prob > _EPS:
                values.append(float(k))
                probs.append(prob)

        if not values:
            self._truncated_stats = None
            return

        norm_probs = [p / norm for p in probs]
        mean_trunc = sum(v * p for v, p in zip(values, norm_probs))
        var_trunc = max(0.0, sum(((v - mean_trunc) ** 2) * p for v, p in zip(values, norm_probs)))

        if var_trunc > _EPS:
            mu3 = sum(((v - mean_trunc) ** 3) * p for v, p in zip(values, norm_probs))
            mu4 = sum(((v - mean_trunc) ** 4) * p for v, p in zip(values, norm_probs))
            skew_trunc = mu3 / (var_trunc ** 1.5)
            kurt_trunc = mu4 / (var_trunc ** 2)
        else:
            skew_trunc = 0.0
            kurt_trunc = 3.0

        max_prob = max(norm_probs)
        modes = [v for v, p in zip(values, norm_probs) if abs(p - max_prob) <= _EPS]

        self._truncated_stats = {
            "mean": mean_trunc,
            "variance": var_trunc,
            "skewness": skew_trunc,
            "kurtosis": kurt_trunc,
            "mode": float(sum(modes) / len(modes)),
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
        if self.is_truncated():
            window = self._get_probability_window()
            if window:
                return self.apply_shift(float(window["start"]))
        return self.apply_shift(float(self.support_low))

    def max_val(self) -> float:
        if self._truncate_invalid:
            return float("nan")
        if self.is_truncated():
            window = self._get_probability_window()
            if window:
                return self.apply_shift(float(window["end"]))
        return self.apply_shift(float(self.support_high))
