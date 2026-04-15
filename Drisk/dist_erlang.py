"""
Erlang distribution support for Drisk.

Definition used here:
    X ~ Erlang(m, beta)
    where m is a positive integer shape parameter and beta > 0 is the scale.

This is the integer-shape special case of the Gamma distribution:
    PDF = x^(m-1) * exp(-x / beta) / (beta^m * Gamma(m)), x >= 0
    CDF = GammaCDF(x; shape=m, scale=beta)
    PPF = GammaPPF(q; shape=m, scale=beta)
"""

import math
from typing import List, Optional, Union

import numpy as np
import scipy.stats as sps

from distribution_base import GammaDistribution

_EPS = 1e-12


def _to_int(name: str, value: float) -> int:
    num = float(value)
    if abs(num - round(num)) > 1e-9:
        raise ValueError(f"{name} must be an integer, got {value}")
    return int(round(num))


def _validate_params(m: int, beta: float) -> None:
    if m <= 0:
        raise ValueError(f"m must be a positive integer, got {m}")
    if beta <= 0.0:
        raise ValueError(f"beta must be > 0, got {beta}")


def erlang_generator_single(rng: np.random.Generator, params: List[float]) -> float:
    m = _to_int("m", params[0])
    beta = float(params[1])
    _validate_params(m, beta)
    return float(rng.gamma(shape=m, scale=beta))


def erlang_generator_vectorized(
    rng: np.random.Generator, params: List[float], n_samples: int
) -> np.ndarray:
    m = _to_int("m", params[0])
    beta = float(params[1])
    _validate_params(m, beta)
    return rng.gamma(shape=m, scale=beta, size=n_samples).astype(float)


def erlang_generator(
    rng: np.random.Generator,
    params: List[float],
    n_samples: Optional[int] = None,
) -> Union[float, np.ndarray]:
    if n_samples is None:
        return erlang_generator_single(rng, params)
    return erlang_generator_vectorized(rng, params, n_samples)


def erlang_pdf(x: float, m: float, beta: float) -> float:
    m_i = _to_int("m", m)
    beta_f = float(beta)
    _validate_params(m_i, beta_f)
    if x < 0.0:
        return 0.0
    return float(sps.gamma.pdf(x, a=m_i, scale=beta_f))


def erlang_cdf(x: float, m: float, beta: float) -> float:
    m_i = _to_int("m", m)
    beta_f = float(beta)
    _validate_params(m_i, beta_f)
    if x < 0.0:
        return 0.0
    return float(sps.gamma.cdf(x, a=m_i, scale=beta_f))


def erlang_ppf(q_prob: float, m: float, beta: float) -> float:
    m_i = _to_int("m", m)
    beta_f = float(beta)
    _validate_params(m_i, beta_f)
    if q_prob <= 0.0:
        return 0.0
    if q_prob >= 1.0:
        return float("inf")
    return float(sps.gamma.ppf(q_prob, a=m_i, scale=beta_f))


def erlang_raw_mean(m: float, beta: float) -> float:
    m_i = _to_int("m", m)
    beta_f = float(beta)
    _validate_params(m_i, beta_f)
    return float(m_i * beta_f)


class ErlangDistribution(GammaDistribution):
    """Erlang distribution implemented as the integer-shape Gamma special case."""

    def __init__(self, params: List[float], markers: dict = None, func_name: str = None):
        if len(params) >= 2:
            m = _to_int("m", params[0])
            beta = float(params[1])
        else:
            m, beta = 2, 10.0
        _validate_params(m, beta)

        self.m = m
        self.beta = beta
        super().__init__([float(m), beta], markers, func_name)
        self.shape = float(m)
        self.scale = beta

    def mode(self) -> float:
        if self._truncate_invalid:
            return float("nan")
        mode_val = self.beta * (self.m - 1) if self.m >= 1 else 0.0
        return mode_val + self.shift_amount

    def skewness(self) -> float:
        if self._truncate_invalid:
            return float("nan")
        if self._truncated_stats:
            return self._truncated_stats["skewness"]
        return 2.0 / math.sqrt(self.m)

    def kurtosis(self) -> float:
        if self._truncate_invalid:
            return float("nan")
        if self._truncated_stats:
            return self._truncated_stats["kurtosis"]
        return 3.0 + 6.0 / self.m
