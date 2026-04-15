from dist_discrete import (
    _extract_numbers_from_input,
    discrete_generator_single,
    discrete_generator_vectorized,
    discrete_pmf,
    discrete_cdf,
    discrete_ppf,
    DiscreteDistribution,
)
import math
import numpy as np


def _normalize_x_vals(val):
    xs = sorted(float(x) for x in _extract_numbers_from_input(val))
    if not xs:
        raise ValueError('DUniform requires at least one X-table value')
    for i in range(1, len(xs)):
        if abs(xs[i] - xs[i - 1]) < 1e-12:
            raise ValueError('DUniform X-table values must be unique')
    return xs


def _resolve_x_vals(params=None, x_vals=None, markers=None):
    if x_vals is not None:
        return _normalize_x_vals(x_vals)
    if markers and markers.get('x_vals') is not None:
        return _normalize_x_vals(markers['x_vals'])
    if params and len(params) >= 1:
        return _normalize_x_vals(params[0])
    raise ValueError('Unable to resolve DUniform X-table values')


def _equal_probs(x_vals):
    p = 1.0 / len(x_vals)
    return [p] * len(x_vals)


def duniform_generator_single(rng, params, x_vals=None):
    xs = _resolve_x_vals(params, x_vals=x_vals)
    return discrete_generator_single(rng, params, xs, _equal_probs(xs))


def duniform_generator_vectorized(rng, params, n_samples, x_vals=None):
    xs = _resolve_x_vals(params, x_vals=x_vals)
    return discrete_generator_vectorized(rng, params, n_samples, xs, _equal_probs(xs))


def duniform_generator(rng, params, n_samples=None, x_vals=None):
    if n_samples is None:
        return duniform_generator_single(rng, params, x_vals=x_vals)
    return duniform_generator_vectorized(rng, params, n_samples, x_vals=x_vals)


def duniform_pmf(x, x_vals):
    xs = _resolve_x_vals(x_vals=x_vals)
    return discrete_pmf(x, xs, _equal_probs(xs))


def duniform_cdf(x, x_vals):
    xs = _resolve_x_vals(x_vals=x_vals)
    return discrete_cdf(x, xs, _equal_probs(xs))


def duniform_ppf(q, x_vals):
    xs = _resolve_x_vals(x_vals=x_vals)
    return discrete_ppf(q, xs, _equal_probs(xs))


class DUniformDistribution(DiscreteDistribution):
    def __init__(self, params, markers=None, func_name=None):
        markers = dict(markers or {})
        xs = _resolve_x_vals(params, markers=markers)
        p = 1.0 / len(xs)
        x_str = ','.join(str(x) for x in xs)
        p_str = ','.join(str(p) for _ in xs)
        markers['x_vals'] = x_str
        markers['p_vals'] = p_str
        super().__init__([x_str, p_str], markers, func_name)
        self._raw_mode = float(xs[0]) if len(xs) == 1 else float('nan')
        if self.truncate_type == 'percentile2':
            if self.truncate_lower is not None:
                self.truncate_lower += self.shift_amount
            if self.truncate_upper is not None:
                self.truncate_upper += self.shift_amount
            self._truncate_invalid = False
            self.adjusted_truncate_lower = self.truncate_lower
            self.adjusted_truncate_upper = self.truncate_upper
            self._finalize_truncation()
            self._truncated_stats = None
            self._compute_truncated_stats()

    def _compute_truncated_stats(self):
        if not self.is_truncated():
            return

        low, high = self.get_truncated_bounds()
        if low is None and high is None:
            return

        if self.truncate_type in ['value2', 'percentile2']:
            low_orig = low - self.shift_amount if low is not None else None
            high_orig = high - self.shift_amount if high is not None else None
        else:
            low_orig = low
            high_orig = high

        if low_orig is None:
            low_orig = min(self.x_vals)
        if high_orig is None:
            high_orig = max(self.x_vals)

        lower_fraction = 1.0
        upper_fraction = 1.0
        lower_idx = None
        upper_idx = None
        if self.truncate_type in ['percentile', 'percentile2']:
            if self.truncate_lower_pct is not None:
                lower_idx, lower_fraction = self._find_boundary_info(self.truncate_lower_pct)
            if self.truncate_upper_pct is not None:
                upper_idx, upper_fraction = self._find_boundary_info(self.truncate_upper_pct)

        values = []
        probs = []
        for i, (x, p) in enumerate(zip(self.x_vals, self.p_vals)):
            if x < low_orig - 1e-12 or x > high_orig + 1e-12:
                continue
            prob = p
            if lower_idx is not None and i == lower_idx:
                prob *= lower_fraction
            if upper_idx is not None and i == upper_idx:
                prob *= upper_fraction
            if prob > 0:
                values.append(x)
                probs.append(prob)

        if not values:
            self._truncated_stats = None
            return

        total_prob = sum(probs)
        if total_prob <= 0:
            self._truncated_stats = None
            return

        norm_probs = [p / total_prob for p in probs]
        mean_trunc = sum(v * p for v, p in zip(values, norm_probs))
        var_trunc = sum((v - mean_trunc) ** 2 * p for v, p in zip(values, norm_probs))
        if var_trunc > 0:
            skew_trunc = sum((v - mean_trunc) ** 3 * p for v, p in zip(values, norm_probs)) / (var_trunc ** 1.5)
            kurt_trunc = sum((v - mean_trunc) ** 4 * p for v, p in zip(values, norm_probs)) / (var_trunc ** 2)
        else:
            skew_trunc = 0.0
            kurt_trunc = 3.0

        max_prob = max(norm_probs)
        mode_indices = [i for i, p in enumerate(norm_probs) if abs(p - max_prob) < 1e-12]
        mode_trunc = values[mode_indices[0]] if len(mode_indices) == 1 else float('nan')

        self._truncated_stats = {
            'mean': mean_trunc,
            'variance': var_trunc,
            'skewness': skew_trunc,
            'kurtosis': kurt_trunc,
            'mode': mode_trunc,
        }
