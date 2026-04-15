# dist_binomial.py
"""
二项分布（Binomial）专属模块
用于 Drisk 系统集成，提供随机数生成、CDF/PPF/PMF、理论统计类（含离散求和截断矩）。
完全独立，不依赖 scipy。
包含向量化批量生成函数，适用于索引模拟。
"""

import math
import numpy as np
from typing import List, Optional, Union
from scipy.special import gammaln  # 仅用于组合数对数，可选；若不想依赖scipy，可自行实现
from distribution_base import DistributionBase

# -------------------- 参数验证 --------------------
def _validate_params(n: int, p: float) -> None:
    if n <= 0 or not float(n).is_integer():
        raise ValueError(f"试验次数 n 必须为正整数，收到 {n}")
    if not (0 < p < 1):
        raise ValueError(f"成功概率 p 必须满足 0 < p < 1，收到 {p}")

# -------------------- 单个随机数生成器 --------------------
def binomial_generator_single(rng: np.random.Generator, params: List[float]) -> float:
    n = int(params[0])
    p = params[1]
    _validate_params(n, p)
    return float(rng.binomial(n, p))

# -------------------- 向量化批量生成器 --------------------
def binomial_generator_vectorized(rng, params, n_samples):
    """
    向量化二项分布生成器
    params: [n, p] 其中 n 和 p 可以是标量或长度等于 n_samples 的数组
    """
    n = params[0]
    p = params[1]
    # 将 n 转为整数数组，并广播到 n_samples
    if np.isscalar(n):
        n_arr = np.full(n_samples, int(n), dtype=int)
    else:
        n_arr = np.asarray(n, dtype=int)
        if n_arr.shape != (n_samples,):
            n_arr = np.broadcast_to(n_arr, (n_samples,))
    # 将 p 转为浮点数组，并广播到 n_samples
    if np.isscalar(p):
        p_arr = np.full(n_samples, float(p), dtype=float)
    else:
        p_arr = np.asarray(p, dtype=float)
        if p_arr.shape != (n_samples,):
            p_arr = np.broadcast_to(p_arr, (n_samples,))
    # 生成样本（NumPy 的 binomial 支持数组参数）
    samples = rng.binomial(n_arr, p_arr, size=n_samples)
    return samples.astype(float)

def binomial_generator(rng: np.random.Generator, params: List[float], n_samples: Optional[int] = None) -> Union[float, np.ndarray]:
    if n_samples is None:
        return binomial_generator_single(rng, params)
    else:
        return binomial_generator_vectorized(rng, params, n_samples)

# -------------------- 原始分布函数 --------------------
# 为计算精确，可预先计算所有可能的 PMF 和 CDF，但对于大 n 可能内存过大。
# 此处采用直接计算（使用对数以避免溢出）
def _log_comb(n: int, k: int) -> float:
    """组合数的对数"""
    return gammaln(n + 1) - gammaln(k + 1) - gammaln(n - k + 1)

def binomial_pmf(k: int, n: int, p: float) -> float:
    if k < 0 or k > n:
        return 0.0
    log_pmf = _log_comb(n, k) + k * math.log(p) + (n - k) * math.log(1 - p)
    return math.exp(log_pmf)

def binomial_cdf(x: float, n: int, p: float) -> float:
    if x < 0:
        return 0.0
    if x >= n:
        return 1.0
    k_max = int(math.floor(x))
    # 累加 pmf 直到 k_max
    total = 0.0
    for k in range(0, k_max + 1):
        total += binomial_pmf(k, n, p)
    return total

def binomial_ppf(q: float, n: int, p: float) -> float:
    if q <= 0:
        return 0.0
    if q >= 1:
        return float(n)
    cum = 0.0
    for k in range(0, n + 1):
        cum += binomial_pmf(k, n, p)
        if cum >= q - 1e-12:
            return float(k)
    return float(n)

# -------------------- 原始矩 --------------------
def binomial_raw_mean(n: int, p: float) -> float:
    return n * p

def binomial_raw_var(n: int, p: float) -> float:
    return n * p * (1 - p)

def binomial_raw_skew(n: int, p: float) -> float:
    std = math.sqrt(binomial_raw_var(n, p))
    if std == 0:
        return 0.0
    return (1 - 2 * p) / std

def binomial_raw_kurt(n: int, p: float) -> float:
    var = n * p * (1 - p)
    if var == 0:
        return 3.0
    return (1 - 6 * p * (1 - p)) / (n * p * (1 - p)) + 3.0

def binomial_raw_mode(n: int, p: float) -> float:
    # 众数为 floor((n+1)p)
    mode_val = math.floor((n + 1) * p)
    if mode_val > n:
        mode_val = n
    if mode_val < 0:
        mode_val = 0
    return float(mode_val)

# -------------------- 理论统计类 --------------------
class BinomialDistribution(DistributionBase):
    """
    二项分布理论统计类。
    继承 DistributionBase，通过离散点求和计算截断矩。
    """

    def __init__(self, params: List[float], markers: dict = None, func_name: str = None):
        if len(params) < 2:
            n, p = 10, 0.5
        else:
            n = int(params[0])
            p = float(params[1])
        _validate_params(n, p)

        self.n = n
        self.p = p

        self._raw_mean = binomial_raw_mean(n, p)
        self._raw_var = binomial_raw_var(n, p)
        self._raw_skew = binomial_raw_skew(n, p)
        self._raw_kurt = binomial_raw_kurt(n, p)
        self._raw_mode = binomial_raw_mode(n, p)

        super().__init__(params, markers, func_name)

        if self.truncate_type in ['percentile', 'percentile2']:
            if self.truncate_lower_pct is not None:
                self.truncate_lower = self._original_ppf(self.truncate_lower_pct)
            if self.truncate_upper_pct is not None:
                self.truncate_upper = self._original_ppf(self.truncate_upper_pct)

        # 完成截断设置后调用
        self._finalize_truncation()

        self._truncated_stats = None
        self._compute_truncated_stats()

    def _find_boundary_info(self, pct: float):
        """用于百分位数截断：找到边界点索引和保留比例"""
        cum = 0.0
        for k in range(0, self.n + 1):
            pmf_k = binomial_pmf(k, self.n, self.p)
            if cum <= pct < cum + pmf_k:
                fraction = (pct - cum) / pmf_k if pmf_k > 0 else 0.0
                return k, fraction
            cum += pmf_k
        return self.n, 1.0

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
            low_orig = 0
        if high_orig is None:
            high_orig = self.n

        # 处理百分位数截断的边界比例
        lower_fraction = 1.0
        upper_fraction = 1.0
        lower_idx = None
        upper_idx = None
        if self.truncate_type in ['percentile', 'percentile2']:
            if self.truncate_lower_pct is not None:
                lower_idx, lower_fraction = self._find_boundary_info(self.truncate_lower_pct)
            if self.truncate_upper_pct is not None:
                upper_idx, upper_fraction = self._find_boundary_info(self.truncate_upper_pct)

        # 收集截断区间内的点及其概率
        values = []
        probs = []
        for k in range(0, self.n + 1):
            x = float(k)
            if x < low_orig - 1e-12 or x > high_orig + 1e-12:
                continue
            prob = binomial_pmf(k, self.n, self.p)
            # 应用边界分数
            if lower_idx is not None and k == lower_idx:
                prob *= lower_fraction
            if upper_idx is not None and k == upper_idx:
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

        # 计算矩
        mean_trunc = sum(v * p for v, p in zip(values, norm_probs))
        var_trunc = sum((v - mean_trunc)**2 * p for v, p in zip(values, norm_probs))
        if var_trunc > 0:
            skew_trunc = sum((v - mean_trunc)**3 * p for v, p in zip(values, norm_probs)) / (var_trunc ** 1.5)
            kurt_trunc = sum((v - mean_trunc)**4 * p for v, p in zip(values, norm_probs)) / (var_trunc ** 2)
        else:
            skew_trunc = 0.0
            kurt_trunc = 3.0

        # 众数：取概率最大的值，若多个则平均
        max_prob = max(norm_probs)
        mode_indices = [i for i, p in enumerate(norm_probs) if abs(p - max_prob) < 1e-12]
        if len(mode_indices) == 1:
            mode_trunc = values[mode_indices[0]]
        else:
            mode_trunc = sum(values[i] for i in mode_indices) / len(mode_indices)

        self._truncated_stats = {
            'mean': mean_trunc,
            'variance': var_trunc,
            'skewness': skew_trunc,
            'kurtosis': kurt_trunc,
            'mode': mode_trunc
        }

    def _original_cdf(self, x: float) -> float:
        return binomial_cdf(x, self.n, self.p)

    def _original_ppf(self, q: float) -> float:
        return binomial_ppf(q, self.n, self.p)

    def _original_pdf(self, x: float) -> float:
        # 对于离散分布，pdf 实为 pmf
        return binomial_pmf(int(round(x)), self.n, self.p) if abs(x - round(x)) < 1e-12 else 0.0

    def mean(self) -> float:
        if self._truncate_invalid:
            return float('nan')
        if self._truncated_stats:
            return self._truncated_stats['mean'] + self.shift_amount
        return self._raw_mean + self.shift_amount

    def variance(self) -> float:
        if self._truncate_invalid:
            return float('nan')
        if self._truncated_stats:
            return self._truncated_stats['variance']
        return self._raw_var

    def skewness(self) -> float:
        if self._truncate_invalid:
            return float('nan')
        if self._truncated_stats:
            return self._truncated_stats['skewness']
        return self._raw_skew

    def kurtosis(self) -> float:
        if self._truncate_invalid:
            return float('nan')
        if self._truncated_stats:
            return self._truncated_stats['kurtosis']
        return self._raw_kurt

    def mode(self) -> float:
        if self._truncate_invalid:
            return float('nan')
        if self._truncated_stats:
            return self._truncated_stats['mode'] + self.shift_amount
        raw_mode = self._raw_mode
        if self.is_truncated():
            low, high = self.get_effective_bounds()
            if low is not None and raw_mode + self.shift_amount < low:
                return low
            if high is not None and raw_mode + self.shift_amount > high:
                return high
        return raw_mode + self.shift_amount

    # ========== 修正 min_val 和 max_val，确保应用平移 ==========
    def min_val(self) -> float:
        if self._truncate_invalid:
            return float('nan')
        lower, upper = self.get_effective_bounds()
        if lower is not None:
            return lower
        return self.apply_shift(self.support_low)

    def max_val(self) -> float:
        if self._truncate_invalid:
            return float('nan')
        lower, upper = self.get_effective_bounds()
        if upper is not None:
            return upper
        return self.apply_shift(self.support_high)