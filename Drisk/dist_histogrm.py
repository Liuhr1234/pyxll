"""
Histogrm distribution support for Drisk.

Definition used here:
    X ~ Histogrm(min_val, max_val, p_table)

This is a bounded continuous distribution on [min_val, max_val].
The interval [min_val, max_val] is split into N equal-width bins where
N = len(p_table). Each p_table entry is the unnormalized constant density
height of one bin. The heights are normalized to integrate to 1.
"""

import math
from typing import List, Optional, Tuple, Union

import numpy as np

from distribution_base import DistributionBase

_EPS = 1e-12


def _extract_numbers_from_input(val) -> List[float]:
    numbers: List[float] = []

    def extract(item) -> None:
        if isinstance(item, (tuple, list, np.ndarray)):
            for sub in item:
                extract(sub)
        elif isinstance(item, str):
            clean = item.strip()
            if clean.startswith("{") and clean.endswith("}"):
                clean = clean[1:-1].strip()
            parts = [p.strip() for p in clean.split(",") if p.strip()]
            for part in parts:
                try:
                    numbers.append(float(part))
                except Exception:
                    pass
        else:
            try:
                if item is not None:
                    numbers.append(float(item))
            except Exception:
                pass

    extract(val)
    return numbers


def _parse_histogrm_p_table(p_input) -> List[float]:
    p_vals = _extract_numbers_from_input(p_input)
    _validate_histogrm_params(0.0, 1.0, p_vals, validate_bounds=False)
    return p_vals


def _validate_histogrm_params(
    min_val: float,
    max_val: float,
    p_vals: List[float],
    validate_bounds: bool = True,
) -> None:
    if validate_bounds and max_val <= min_val:
        raise ValueError("Histogrm requires min_val < max_val")
    if len(p_vals) < 1:
        raise ValueError("Histogrm requires at least one P-Table value")
    if any(p <= 0.0 for p in p_vals):
        raise ValueError("P-Table values must all be positive")


class _HistogrmCore:
    def __init__(self, min_val: float, max_val: float, p_vals: List[float]):
        _validate_histogrm_params(min_val, max_val, p_vals)
        self.min_val = float(min_val)
        self.max_val = float(max_val)
        self.p_vals_raw = np.array([float(p) for p in p_vals], dtype=float)
        self.n_bins = len(self.p_vals_raw)
        self.bin_width = (self.max_val - self.min_val) / self.n_bins
        self.x_points = np.linspace(self.min_val, self.max_val, self.n_bins + 1, dtype=float)
        total_area = float(np.sum(self.p_vals_raw) * self.bin_width)
        if total_area <= _EPS:
            raise ValueError("Histogrm PDF area must be positive")
        self.p_vals = self.p_vals_raw / total_area
        self.bin_masses = self.p_vals * self.bin_width
        self.cdf_points = np.zeros(self.n_bins + 1, dtype=float)
        self.cdf_points[1:] = np.cumsum(self.bin_masses)
        self.cdf_points[-1] = 1.0

    def _find_bin(self, x: float) -> int:
        if x < self.min_val:
            return -1
        if x > self.max_val:
            return self.n_bins
        if x >= self.max_val:
            return self.n_bins - 1
        idx = int((x - self.min_val) / self.bin_width)
        return max(0, min(idx, self.n_bins - 1))

    def pdf(self, x: float) -> float:
        x = float(x)
        idx = self._find_bin(x)
        if idx < 0 or idx >= self.n_bins:
            return 0.0
        left = self.x_points[idx]
        right = self.x_points[idx + 1]
        if idx == self.n_bins - 1:
            if x < left or x > right:
                return 0.0
        else:
            if x < left or x >= right:
                return 0.0
        return float(self.p_vals[idx])

    def cdf(self, x: float) -> float:
        x = float(x)
        if x <= self.min_val:
            return 0.0
        if x >= self.max_val:
            return 1.0
        idx = self._find_bin(x)
        if idx < 0:
            return 0.0
        left = self.x_points[idx]
        return float(self.cdf_points[idx] + self.p_vals[idx] * (x - left))

    def ppf(self, q: float) -> float:
        q = float(q)
        if q <= 0.0:
            return self.min_val
        if q >= 1.0:
            return self.max_val
        idx = int(np.searchsorted(self.cdf_points, q, side="right") - 1)
        idx = max(0, min(idx, self.n_bins - 1))
        while idx < self.n_bins - 1 and self.bin_masses[idx] <= _EPS:
            idx += 1
        left = self.x_points[idx]
        if self.p_vals[idx] <= _EPS:
            return float(left)
        offset = (q - self.cdf_points[idx]) / self.p_vals[idx]
        return float(min(self.x_points[idx + 1], max(left, left + offset)))

    def moment_integrals(self, lower: float, upper: float, max_power: int = 4) -> List[float]:
        lower = max(float(lower), self.min_val)
        upper = min(float(upper), self.max_val)
        if upper < lower:
            return [0.0] * (max_power + 1)
        results = [0.0] * (max_power + 1)
        for i in range(self.n_bins):
            seg_low = max(lower, self.x_points[i])
            seg_high = min(upper, self.x_points[i + 1])
            if seg_high <= seg_low:
                continue
            density = self.p_vals[i]
            for power in range(max_power + 1):
                results[power] += density * (
                    seg_high ** (power + 1) - seg_low ** (power + 1)
                ) / (power + 1)
        return results

    def mode_on_interval(self, lower: float, upper: float) -> float:
        lower = max(float(lower), self.min_val)
        upper = min(float(upper), self.max_val)
        best_pdf = -1.0
        best_mid = lower
        for i in range(self.n_bins):
            left = max(lower, self.x_points[i])
            right = min(upper, self.x_points[i + 1])
            if right <= left:
                continue
            density = float(self.p_vals[i])
            mid = 0.5 * (left + right)
            if density > best_pdf + _EPS or (abs(density - best_pdf) <= _EPS and mid < best_mid):
                best_pdf = density
                best_mid = mid
        return float(best_mid)


def histogrm_generator_single(
    rng: np.random.Generator,
    params: List[float],
    p_vals: List[float],
) -> float:
    core = _HistogrmCore(float(params[0]), float(params[1]), p_vals)
    return core.ppf(float(rng.uniform(_EPS, 1.0 - _EPS)))


def histogrm_generator_vectorized(
    rng: np.random.Generator,
    params: List[float],
    n_samples: int,
    p_vals: List[float],
) -> np.ndarray:
    core = _HistogrmCore(float(params[0]), float(params[1]), p_vals)
    u = rng.uniform(_EPS, 1.0 - _EPS, size=n_samples)
    return np.array([core.ppf(float(q)) for q in u], dtype=float)


def histogrm_generator(
    rng: np.random.Generator,
    params: List[float],
    n_samples: Optional[int] = None,
    p_vals: List[float] = None,
) -> Union[float, np.ndarray]:
    if p_vals is None:
        raise ValueError("histogrm_generator requires p_vals")
    if n_samples is None:
        return histogrm_generator_single(rng, params, p_vals)
    return histogrm_generator_vectorized(rng, params, n_samples, p_vals)


def histogrm_pdf(x: float, min_val: float, max_val: float, p_vals: List[float]) -> float:
    return _HistogrmCore(min_val, max_val, p_vals).pdf(x)


def histogrm_cdf(x: float, min_val: float, max_val: float, p_vals: List[float]) -> float:
    return _HistogrmCore(min_val, max_val, p_vals).cdf(x)


def histogrm_ppf(q: float, min_val: float, max_val: float, p_vals: List[float]) -> float:
    return _HistogrmCore(min_val, max_val, p_vals).ppf(q)


def histogrm_raw_mean(min_val: float, max_val: float, p_vals: List[float]) -> float:
    core = _HistogrmCore(min_val, max_val, p_vals)
    return float(core.moment_integrals(core.min_val, core.max_val, 1)[1])


def histogrm_raw_var(min_val: float, max_val: float, p_vals: List[float]) -> float:
    core = _HistogrmCore(min_val, max_val, p_vals)
    integrals = core.moment_integrals(core.min_val, core.max_val, 2)
    mean_val = integrals[1]
    return float(max(0.0, integrals[2] - mean_val * mean_val))


def histogrm_raw_skewness(min_val: float, max_val: float, p_vals: List[float]) -> float:
    core = _HistogrmCore(min_val, max_val, p_vals)
    integrals = core.moment_integrals(core.min_val, core.max_val, 3)
    mean_val = integrals[1]
    variance = max(0.0, integrals[2] - mean_val * mean_val)
    if variance <= _EPS:
        return 0.0
    mu3 = integrals[3] - 3.0 * mean_val * integrals[2] + 2.0 * mean_val ** 3
    return float(mu3 / (variance ** 1.5))


def histogrm_raw_kurtosis(min_val: float, max_val: float, p_vals: List[float]) -> float:
    core = _HistogrmCore(min_val, max_val, p_vals)
    integrals = core.moment_integrals(core.min_val, core.max_val, 4)
    mean_val = integrals[1]
    variance = max(0.0, integrals[2] - mean_val * mean_val)
    if variance <= _EPS:
        return 3.0
    mu4 = (
        integrals[4]
        - 4.0 * mean_val * integrals[3]
        + 6.0 * mean_val * mean_val * integrals[2]
        - 3.0 * mean_val ** 4
    )
    return float(mu4 / (variance ** 2))


def histogrm_raw_mode(min_val: float, max_val: float, p_vals: List[float]) -> float:
    core = _HistogrmCore(min_val, max_val, p_vals)
    return core.mode_on_interval(core.min_val, core.max_val)


class HistogrmDistribution(DistributionBase):
    """Histogrm theoretical distribution with truncation support."""

    def __init__(self, params: List[float], markers: dict = None, func_name: str = None):
        markers = markers or {}
        if len(params) >= 3:
            min_val = float(params[0])
            max_val = float(params[1])
            p_source = markers.get("p_vals", params[2])
        else:
            min_val, max_val = -1.0, 1.0
            p_source = markers.get("p_vals", "{0.1,0.2,0.4,0.2,0.1}")

        p_vals = _parse_histogrm_p_table(p_source)
        self.min_support = float(min_val)
        self.max_support = float(max_val)
        self.p_vals = p_vals
        self._core = _HistogrmCore(min_val, max_val, p_vals)

        self.support_low = self.min_support
        self.support_high = self.max_support
        self._raw_mean = histogrm_raw_mean(min_val, max_val, p_vals)
        self._raw_var = histogrm_raw_var(min_val, max_val, p_vals)
        self._raw_skew = histogrm_raw_skewness(min_val, max_val, p_vals)
        self._raw_kurt = histogrm_raw_kurtosis(min_val, max_val, p_vals)
        self._raw_mode = histogrm_raw_mode(min_val, max_val, p_vals)
        self._truncated_stats = None

        super().__init__([min_val, max_val, params[2] if len(params) >= 3 else p_source], markers, func_name)

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
        return self._core.cdf(x)

    def _original_ppf(self, q_prob: float) -> float:
        return self._core.ppf(q_prob)

    def _original_pdf(self, x: float) -> float:
        return self._core.pdf(x)

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
        lb = self.support_low if low_orig is None else max(self.support_low, low_orig)
        ub = self.support_high if high_orig is None else min(self.support_high, high_orig)
        if ub < lb:
            self._truncated_stats = None
            return

        integrals = self._core.moment_integrals(lb, ub, 4)
        mass = integrals[0]
        if mass <= _EPS:
            self._truncated_stats = None
            return

        m1 = integrals[1] / mass
        m2 = integrals[2] / mass
        m3_raw = integrals[3] / mass
        m4_raw = integrals[4] / mass
        variance = max(0.0, m2 - m1 * m1)
        if variance > _EPS:
            mu3 = m3_raw - 3.0 * m1 * m2 + 2.0 * m1 ** 3
            mu4 = m4_raw - 4.0 * m1 * m3_raw + 6.0 * m1 * m1 * m2 - 3.0 * m1 ** 4
            skewness = mu3 / (variance ** 1.5)
            kurtosis = mu4 / (variance ** 2)
        else:
            skewness = 0.0
            kurtosis = 3.0

        mode_raw = self._core.mode_on_interval(lb, ub)
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

