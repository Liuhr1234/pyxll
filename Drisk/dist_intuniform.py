"""
Integer uniform distribution support for Drisk.

Definition used here:
    X is uniformly distributed over the integers [min_val, max_val].
    PMF = 1 / (max_val - min_val + 1) for integer x in the support.
"""

import math
from typing import List, Optional, Union

import numpy as np

from dist_duniform import DUniformDistribution

_EPS = 1e-12


def _to_int(name: str, value: float) -> int:
    num = float(value)
    rounded = round(num)
    if abs(num - rounded) <= 1e-9:
        return int(rounded)
    if not num.is_integer():
        raise ValueError(f"{name} must be an integer, got {value}")
    return int(num)


def _validate_params(min_val: int, max_val: int) -> None:
    if min_val >= max_val:
        raise ValueError(
            f"Intuniform requires min_val < max_val, got min_val={min_val}, max_val={max_val}"
        )


def _build_x_values(min_val: int, max_val: int) -> list[float]:
    _validate_params(min_val, max_val)
    return [float(x) for x in range(min_val, max_val + 1)]


def _has_invalid_percentile_marker(markers: dict) -> bool:
    if not markers:
        return False
    for key in ("truncate_p", "truncatep", "truncate_p2", "truncatep2"):
        raw = markers.get(key)
        if raw is None:
            continue
        raw_text = str(raw).strip()
        if raw_text.startswith("(") and raw_text.endswith(")"):
            raw_text = raw_text[1:-1]
        parts = [p.strip() for p in raw_text.split(",")]
        for part in parts:
            if not part:
                continue
            try:
                value = float(part)
            except Exception:
                return True
            if value < 0.0 or value > 1.0:
                return True
    return False


def intuniform_generator_single(rng: np.random.Generator, params: List[float]) -> float:
    min_val = _to_int("min_val", params[0])
    max_val = _to_int("max_val", params[1])
    _validate_params(min_val, max_val)
    return float(rng.integers(min_val, max_val + 1))


def intuniform_generator_vectorized(
    rng: np.random.Generator, params: List[float], n_samples: int
) -> np.ndarray:
    min_val = params[0]
    max_val = params[1]

    if np.isscalar(min_val):
        min_arr = np.full(n_samples, _to_int("min_val", min_val), dtype=int)
    else:
        min_raw = np.asarray(min_val, dtype=float)
        if min_raw.shape != (n_samples,):
            min_raw = np.broadcast_to(min_raw, (n_samples,))
        if np.any(np.abs(min_raw - np.round(min_raw)) > 1e-9):
            raise ValueError(f"min_val must be an integer, got {min_val}")
        min_arr = np.round(min_raw).astype(int)

    if np.isscalar(max_val):
        max_arr = np.full(n_samples, _to_int("max_val", max_val), dtype=int)
    else:
        max_raw = np.asarray(max_val, dtype=float)
        if max_raw.shape != (n_samples,):
            max_raw = np.broadcast_to(max_raw, (n_samples,))
        if np.any(np.abs(max_raw - np.round(max_raw)) > 1e-9):
            raise ValueError(f"max_val must be an integer, got {max_val}")
        max_arr = np.round(max_raw).astype(int)

    if np.any(min_arr >= max_arr):
        raise ValueError("Intuniform requires min_val < max_val")

    return rng.integers(min_arr, max_arr + 1, size=n_samples).astype(float)


def intuniform_generator(
    rng: np.random.Generator,
    params: List[float],
    n_samples: Optional[int] = None,
) -> Union[float, np.ndarray]:
    if n_samples is None:
        return intuniform_generator_single(rng, params)
    return intuniform_generator_vectorized(rng, params, n_samples)


def intuniform_pmf(x: float, min_val: float, max_val: float) -> float:
    min_i = _to_int("min_val", min_val)
    max_i = _to_int("max_val", max_val)
    _validate_params(min_i, max_i)
    if abs(x - round(x)) > _EPS:
        return 0.0
    k = int(round(x))
    if k < min_i or k > max_i:
        return 0.0
    return 1.0 / (max_i - min_i + 1)


def intuniform_cdf(x: float, min_val: float, max_val: float) -> float:
    min_i = _to_int("min_val", min_val)
    max_i = _to_int("max_val", max_val)
    _validate_params(min_i, max_i)
    if x < min_i:
        return 0.0
    if x >= max_i:
        return 1.0
    count = int(math.floor(x)) - min_i + 1
    count = max(0, min(count, max_i - min_i + 1))
    return count / (max_i - min_i + 1)


def intuniform_ppf(q_prob: float, min_val: float, max_val: float) -> float:
    min_i = _to_int("min_val", min_val)
    max_i = _to_int("max_val", max_val)
    _validate_params(min_i, max_i)
    n = max_i - min_i + 1
    if q_prob <= 0.0:
        return float(min_i)
    if q_prob >= 1.0:
        return float(max_i)
    idx = min(int(math.floor(q_prob * n)), n - 1)
    return float(min_i + idx)


class IntuniformDistribution(DUniformDistribution):
    """Integer uniform distribution built on the discrete-uniform engine."""

    def __init__(self, params: List[float], markers: dict = None, func_name: str = None):
        markers = dict(markers or {})
        if len(params) >= 2:
            min_val = _to_int("min_val", params[0])
            max_val = _to_int("max_val", params[1])
        else:
            min_val, max_val = 0, 10
        _validate_params(min_val, max_val)

        self.min_param = min_val
        self.max_param = max_val
        self.min_int = min_val
        self.max_int = max_val
        x_vals = _build_x_values(min_val, max_val)
        x_str = ",".join(str(int(x)) for x in x_vals)

        super().__init__([x_str], markers, func_name)

        # Keep the public params aligned with the exposed function signature.
        self.params = [float(min_val), float(max_val)]
        self.support_low = float(min_val)
        self.support_high = float(max_val)

        # Re-validate truncation against the true integer support.
        if _has_invalid_percentile_marker(markers):
            self._truncate_invalid = True
            self._truncated_stats = None
        else:
            self._truncate_invalid = False
            self._finalize_truncation()
            self._truncated_stats = None
            self._compute_truncated_stats()

    def _compute_truncated_stats(self):
        if not self.is_truncated():
            self._truncated_stats = None
            return

        low, high = self.get_truncated_bounds()
        if self.truncate_type in ['value2', 'percentile2']:
            low_orig = low - self.shift_amount if low is not None else None
            high_orig = high - self.shift_amount if high is not None else None
        else:
            low_orig = low
            high_orig = high

        if low_orig is None:
            low_orig = self.min_int
        if high_orig is None:
            high_orig = self.max_int

        # Intuniform 的百分位截断按整数点子集处理，不做边界概率拆分。
        values = [x for x in self.x_vals if low_orig - _EPS <= x <= high_orig + _EPS]

        if not values:
            self._truncated_stats = None
            return

        count = len(values)
        mean_trunc = sum(values) / count
        var_trunc = sum((v - mean_trunc) ** 2 for v in values) / count
        if var_trunc > 0:
            skew_trunc = sum((v - mean_trunc) ** 3 for v in values) / count / (var_trunc ** 1.5)
            kurt_trunc = sum((v - mean_trunc) ** 4 for v in values) / count / (var_trunc ** 2)
        else:
            skew_trunc = 0.0
            kurt_trunc = 3.0

        self._truncated_stats = {
            'mean': mean_trunc,
            'variance': var_trunc,
            'skewness': skew_trunc,
            'kurtosis': kurt_trunc,
            'mode': float('nan'),
        }
