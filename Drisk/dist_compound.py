"""Compound distribution support for Drisk."""

import math
from typing import Any, Dict, List, Optional, Tuple

import numpy as np

from distribution_base import DistributionBase

_EPS = 1e-12
_EMPIRICAL_SAMPLE_SIZE = 120000
_EMPIRICAL_SEED = 20260408
_EMPIRICAL_CACHE: Dict[Tuple[str, str, float, float], Dict[str, Any]] = {}


def _normalize_formula_text(formula_text: Any) -> str:
    text = str(formula_text).strip()
    if not text:
        raise ValueError("Compound requires non-empty distribution formulas")
    if not text.startswith("="):
        text = "=" + text
    return text


def _normalize_deductible(value: Any) -> float:
    if value is None:
        return 0.0
    if isinstance(value, str) and not value.strip():
        return 0.0
    deductible = float(value)
    if not math.isfinite(deductible):
        raise ValueError("Compound deductible must be finite")
    if deductible < 0.0:
        raise ValueError("Compound deductible must be >= 0")
    return deductible


def _normalize_limit(value: Any) -> float:
    if value is None:
        return float("inf")
    if isinstance(value, str) and not value.strip():
        return float("inf")
    limit = float(value)
    if math.isnan(limit):
        return float("inf")
    if limit < 0.0:
        raise ValueError("Compound limit must be >= 0")
    return limit


def _normalize_params(
    frequency_formula: Any,
    severity_formula: Any,
    deductible: Any = 0.0,
    limit: Any = float("inf"),
) -> Tuple[str, str, float, float]:
    frequency_text = _normalize_formula_text(frequency_formula)
    severity_text = _normalize_formula_text(severity_formula)
    deductible_value = _normalize_deductible(deductible)
    limit_value = _normalize_limit(limit)
    return frequency_text, severity_text, deductible_value, limit_value


def _parse_distribution(formula_text: str):
    from statistical_functions_theo import _parse_distribution_from_formula_string

    dist = _parse_distribution_from_formula_string(formula_text)
    if dist is None:
        raise ValueError(f"Unable to parse nested distribution: {formula_text}")
    return dist


def _fast_vectorized_ppf(dist, q_values: np.ndarray) -> Optional[np.ndarray]:
    """Use the nested scipy distribution directly when available.

    This preserves the existing deterministic sampling path while avoiding
    Python-level loops for each nested `ppf` call.
    """
    scipy_dist = getattr(dist, "_dist", None)
    if scipy_dist is None or not hasattr(scipy_dist, "ppf") or not hasattr(scipy_dist, "cdf"):
        return None

    q = np.asarray(q_values, dtype=float)
    shift_amount = float(getattr(dist, "shift_amount", 0.0))

    if not dist.is_truncated():
        try:
            result = np.asarray(scipy_dist.ppf(q), dtype=float)
            return result + shift_amount
        except Exception:
            return None

    lower, upper = dist.get_truncated_bounds()
    truncate_type = getattr(dist, "truncate_type", None)

    try:
        if truncate_type in ["value", "percentile"]:
            lower_p = 0.0 if lower is None else float(dist._original_cdf(lower))
            upper_p = 1.0 if upper is None else float(dist._original_cdf(upper))
            original_q = lower_p + q * (upper_p - lower_p)
            result = np.asarray(scipy_dist.ppf(original_q), dtype=float)
            if lower is not None:
                result = np.maximum(result, lower)
            if upper is not None:
                result = np.minimum(result, upper)
            return result + shift_amount

        if truncate_type in ["value2", "percentile2"]:
            lower_orig = None if lower is None else lower - shift_amount
            upper_orig = None if upper is None else upper - shift_amount
            lower_p = 0.0 if lower_orig is None else float(dist._original_cdf(lower_orig))
            upper_p = 1.0 if upper_orig is None else float(dist._original_cdf(upper_orig))
            original_q = lower_p + q * (upper_p - lower_p)
            result = np.asarray(scipy_dist.ppf(original_q), dtype=float)
            shifted_result = result + shift_amount
            if lower is not None:
                result = np.where(shifted_result < lower, lower - shift_amount, result)
            if upper is not None:
                result = np.where(shifted_result > upper, upper - shift_amount, result)
            return result + shift_amount
    except Exception:
        return None

    return None


def _draw_samples_from_distribution(dist, rng: np.random.Generator, n_samples: int) -> np.ndarray:
    if n_samples <= 0:
        return np.empty(0, dtype=float)
    q = rng.uniform(_EPS, 1.0 - _EPS, size=n_samples)
    fast_samples = _fast_vectorized_ppf(dist, q)
    if fast_samples is not None:
        return np.asarray(fast_samples, dtype=float)
    return np.asarray([float(dist.ppf(float(p))) for p in q], dtype=float)


def _generate_compound_samples_with_dist_objects(
    rng: np.random.Generator,
    frequency_dist,
    severity_dist,
    deductible: float,
    limit: float,
    n_samples: int,
) -> np.ndarray:
    if n_samples <= 0:
        return np.empty(0, dtype=float)

    frequency_raw = _draw_samples_from_distribution(frequency_dist, rng, n_samples)
    if not np.all(np.isfinite(frequency_raw)):
        raise ValueError("Compound frequency distribution generated non-finite values")

    counts = np.floor(frequency_raw).astype(np.int64)
    counts[counts < 0] = 0

    totals = np.zeros(n_samples, dtype=float)
    total_claims = int(counts.sum())
    if total_claims <= 0:
        return totals

    severity_raw = _draw_samples_from_distribution(severity_dist, rng, total_claims)
    if not np.all(np.isfinite(severity_raw)):
        raise ValueError("Compound severity distribution generated non-finite values")

    adjusted = severity_raw - deductible
    adjusted = np.maximum(adjusted, 0.0)
    if math.isfinite(limit):
        adjusted = np.minimum(adjusted, limit)

    positive_mask = counts > 0
    positive_counts = counts[positive_mask]
    starts = np.cumsum(np.r_[0, positive_counts[:-1]], dtype=np.int64)
    totals[positive_mask] = np.add.reduceat(adjusted, starts)
    return totals


def _generate_compound_samples(
    rng: np.random.Generator,
    frequency_formula: Any,
    severity_formula: Any,
    deductible: Any = 0.0,
    limit: Any = float("inf"),
    n_samples: int = 1,
) -> np.ndarray:
    frequency_text, severity_text, deductible_value, limit_value = _normalize_params(
        frequency_formula, severity_formula, deductible, limit
    )
    frequency_dist = _parse_distribution(frequency_text)
    severity_dist = _parse_distribution(severity_text)
    return _generate_compound_samples_with_dist_objects(
        rng,
        frequency_dist,
        severity_dist,
        deductible_value,
        limit_value,
        n_samples,
    )


def _estimate_mode(samples: np.ndarray) -> float:
    if samples.size == 0:
        return 0.0

    zero_mass = float(np.mean(np.isclose(samples, 0.0, atol=1e-12)))
    positive = samples[samples > 0.0]
    if positive.size == 0:
        return 0.0

    hist, edges = np.histogram(positive, bins="auto")
    if hist.size == 0:
        return 0.0

    max_bin_mass = float(np.max(hist)) / float(samples.size)
    if zero_mass >= max_bin_mass:
        return 0.0

    mode_index = int(np.argmax(hist))
    return float((edges[mode_index] + edges[mode_index + 1]) / 2.0)


def _bandwidth(samples: np.ndarray) -> float:
    if samples.size <= 1:
        return 1.0

    q75, q25 = np.percentile(samples, [75.0, 25.0])
    iqr = float(q75 - q25)
    if iqr > 0.0:
        bw = 2.0 * iqr / (samples.size ** (1.0 / 3.0))
        if math.isfinite(bw) and bw > 0.0:
            return bw

    std = float(np.std(samples, ddof=1))
    if math.isfinite(std) and std > 0.0:
        bw = 1.06 * std * (samples.size ** (-1.0 / 5.0))
        if math.isfinite(bw) and bw > 0.0:
            return bw

    scale = float(np.max(samples) - np.min(samples))
    if not math.isfinite(scale) or scale <= 0.0:
        scale = max(float(np.mean(np.abs(samples))), 1.0)
    return max(scale * 1e-3, 1e-6)


def _get_empirical_data(
    frequency_formula: Any,
    severity_formula: Any,
    deductible: Any = 0.0,
    limit: Any = float("inf"),
) -> Dict[str, Any]:
    frequency_text, severity_text, deductible_value, limit_value = _normalize_params(
        frequency_formula, severity_formula, deductible, limit
    )
    cache_key = (frequency_text, severity_text, deductible_value, limit_value)
    cached = _EMPIRICAL_CACHE.get(cache_key)
    if cached is not None:
        return cached

    frequency_dist = _parse_distribution(frequency_text)
    severity_dist = _parse_distribution(severity_text)
    rng = np.random.default_rng(_EMPIRICAL_SEED)
    samples = _generate_compound_samples_with_dist_objects(
        rng,
        frequency_dist,
        severity_dist,
        deductible_value,
        limit_value,
        _EMPIRICAL_SAMPLE_SIZE,
    )
    sorted_samples = np.sort(samples.astype(float))
    positive_samples = sorted_samples[sorted_samples > 0.0]

    if sorted_samples.size > 1:
        variance = float(np.var(sorted_samples, ddof=1))
    else:
        variance = 0.0

    if sorted_samples.size > 2 and variance > _EPS:
        centered = sorted_samples - float(np.mean(sorted_samples))
        mu2 = float(np.mean(centered ** 2))
        mu3 = float(np.mean(centered ** 3))
        mu4 = float(np.mean(centered ** 4))
        skewness = mu3 / (mu2 ** 1.5) if mu2 > _EPS else 0.0
        kurtosis = mu4 / (mu2 ** 2) if mu2 > _EPS else 3.0
    else:
        skewness = 0.0
        kurtosis = 3.0

    data = {
        "samples": sorted_samples,
        "mean": float(np.mean(sorted_samples)) if sorted_samples.size else 0.0,
        "variance": max(0.0, variance),
        "skewness": float(skewness),
        "kurtosis": float(kurtosis),
        "mode": _estimate_mode(sorted_samples),
        "pdf_bandwidth": _bandwidth(positive_samples if positive_samples.size else sorted_samples),
    }
    _EMPIRICAL_CACHE[cache_key] = data
    return data


def compound_generator_single(rng: np.random.Generator, params: List[Any]) -> float:
    frequency_formula = params[0]
    severity_formula = params[1]
    deductible = params[2] if len(params) >= 3 else 0.0
    limit = params[3] if len(params) >= 4 else float("inf")
    return float(
        _generate_compound_samples(rng, frequency_formula, severity_formula, deductible, limit, 1)[0]
    )


def compound_generator_vectorized(
    rng: np.random.Generator,
    params: List[Any],
    n_samples: int,
) -> np.ndarray:
    frequency_formula = params[0]
    severity_formula = params[1]
    deductible = params[2] if len(params) >= 3 else 0.0
    limit = params[3] if len(params) >= 4 else float("inf")
    return _generate_compound_samples(rng, frequency_formula, severity_formula, deductible, limit, n_samples)


def compound_cdf(
    x: float,
    frequency_formula: Any,
    severity_formula: Any,
    deductible: Any = 0.0,
    limit: Any = float("inf"),
) -> float:
    data = _get_empirical_data(frequency_formula, severity_formula, deductible, limit)
    samples = data["samples"]
    if x < 0.0:
        return 0.0
    return float(np.searchsorted(samples, x, side="right") / samples.size)


def compound_ppf(
    q: float,
    frequency_formula: Any,
    severity_formula: Any,
    deductible: Any = 0.0,
    limit: Any = float("inf"),
) -> float:
    if q <= 0.0:
        return 0.0
    if q >= 1.0:
        return float("inf")
    data = _get_empirical_data(frequency_formula, severity_formula, deductible, limit)
    return float(np.quantile(data["samples"], q))


def compound_raw_mean(
    frequency_formula: Any,
    severity_formula: Any,
    deductible: Any = 0.0,
    limit: Any = float("inf"),
) -> float:
    data = _get_empirical_data(frequency_formula, severity_formula, deductible, limit)
    return float(data["mean"])


class CompoundDistribution(DistributionBase):
    """Compound distribution backed by a deterministic empirical approximation."""

    def __init__(self, params: List[Any], markers: Dict[str, Any] = None, func_name: str = None):
        markers = markers or {}
        frequency_formula = params[0] if len(params) >= 1 else ""
        severity_formula = params[1] if len(params) >= 2 else ""
        deductible = params[2] if len(params) >= 3 else 0.0
        limit = params[3] if len(params) >= 4 else float("inf")

        (
            self.frequency_formula,
            self.severity_formula,
            self.deductible,
            self.limit,
        ) = _normalize_params(frequency_formula, severity_formula, deductible, limit)
        self._empirical = _get_empirical_data(
            self.frequency_formula,
            self.severity_formula,
            self.deductible,
            self.limit,
        )
        self._samples = self._empirical["samples"]
        self._raw_mean = self._empirical["mean"]
        self._raw_variance = self._empirical["variance"]
        self._raw_skewness = self._empirical["skewness"]
        self._raw_kurtosis = self._empirical["kurtosis"]
        self._raw_mode = self._empirical["mode"]
        self._pdf_bandwidth = self._empirical["pdf_bandwidth"]
        self._truncated_stats: Optional[Dict[str, float]] = None

        super().__init__(
            [self.frequency_formula, self.severity_formula, self.deductible, self.limit],
            markers,
            func_name,
        )

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

    def _get_original_truncation_bounds(self) -> Tuple[Optional[float], Optional[float]]:
        lower, upper = self.get_truncated_bounds()
        if lower is None and upper is None:
            return None, None
        if self.truncate_type in ["value2", "percentile2"]:
            lower = lower - self.shift_amount if lower is not None else None
            upper = upper - self.shift_amount if upper is not None else None
        return lower, upper

    def _compute_truncated_stats(self) -> None:
        if not self.is_truncated():
            self._truncated_stats = None
            return

        lower, upper = self._get_original_truncation_bounds()
        mask = np.ones(self._samples.size, dtype=bool)
        if lower is not None:
            mask &= self._samples >= lower
        if upper is not None:
            mask &= self._samples <= upper

        truncated = self._samples[mask]
        if truncated.size == 0:
            self._truncated_stats = None
            return

        if truncated.size > 1:
            variance = float(np.var(truncated, ddof=1))
        else:
            variance = 0.0

        if truncated.size > 2 and variance > _EPS:
            centered = truncated - float(np.mean(truncated))
            mu2 = float(np.mean(centered ** 2))
            mu3 = float(np.mean(centered ** 3))
            mu4 = float(np.mean(centered ** 4))
            skewness = mu3 / (mu2 ** 1.5) if mu2 > _EPS else 0.0
            kurtosis = mu4 / (mu2 ** 2) if mu2 > _EPS else 3.0
        else:
            skewness = 0.0
            kurtosis = 3.0

        mode_value = _estimate_mode(truncated)
        self._truncated_stats = {
            "mean": float(np.mean(truncated)),
            "variance": max(0.0, variance),
            "skewness": float(skewness),
            "kurtosis": float(kurtosis),
            "mode": float(mode_value),
        }

    def _original_ppf(self, q: float) -> float:
        return compound_ppf(q, self.frequency_formula, self.severity_formula, self.deductible, self.limit)

    def _original_cdf(self, x: float) -> float:
        return compound_cdf(x, self.frequency_formula, self.severity_formula, self.deductible, self.limit)

    def _original_pdf(self, x: float) -> float:
        if x < 0.0:
            return 0.0
        h = max(self._pdf_bandwidth, 1e-6)
        lower = max(0.0, x - h)
        upper = x + h
        if upper <= lower:
            return 0.0
        cdf_upper = self._original_cdf(upper)
        cdf_lower = self._original_cdf(lower)
        return max(0.0, float((cdf_upper - cdf_lower) / (upper - lower)))

    def mean(self) -> float:
        if self._truncate_invalid:
            return float("nan")
        if self._truncated_stats is not None:
            return float(self._truncated_stats["mean"] + self.shift_amount)
        return float(self._raw_mean + self.shift_amount)

    def variance(self) -> float:
        if self._truncate_invalid:
            return float("nan")
        if self._truncated_stats is not None:
            return float(self._truncated_stats["variance"])
        return float(self._raw_variance)

    def skewness(self) -> float:
        if self._truncate_invalid:
            return float("nan")
        if self._truncated_stats is not None:
            return float(self._truncated_stats["skewness"])
        return float(self._raw_skewness)

    def kurtosis(self) -> float:
        if self._truncate_invalid:
            return float("nan")
        if self._truncated_stats is not None:
            return float(self._truncated_stats["kurtosis"])
        return float(self._raw_kurtosis)

    def mode(self) -> float:
        if self._truncate_invalid:
            return float("nan")
        if self._truncated_stats is not None:
            mode_value = float(self._truncated_stats["mode"] + self.shift_amount)
        else:
            mode_value = float(self._raw_mode + self.shift_amount)

        if self.is_truncated():
            lower, upper = self.get_effective_bounds()
            if lower is not None:
                mode_value = max(mode_value, lower)
            if upper is not None:
                mode_value = min(mode_value, upper)
        return mode_value

    def min_val(self) -> float:
        if self._truncate_invalid:
            return float("nan")
        lower, _ = self.get_effective_bounds()
        if lower is not None:
            return float(lower)
        return float(self.shift_amount)

    def max_val(self) -> float:
        if self._truncate_invalid:
            return float("nan")
        _, upper = self.get_effective_bounds()
        if upper is not None:
            return float(upper)
        return float("inf")
