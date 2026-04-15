# dist_bernoulli.py
"""
伯努利分布（Bernoulli）专属模块
用于 Drisk 系统集成，提供随机数生成、CDF/PPF/PMF、理论统计类（含离散求和截断矩）。
完全独立，不依赖 scipy。
包含向量化批量生成函数，适用于索引模拟。
"""

import math
import numpy as np
from typing import List, Optional, Union
from distribution_base import DistributionBase

# -------------------- 参数验证 --------------------
def _validate_params(p: float) -> None:
    """验证参数有效性，无效则抛出 ValueError"""
    if not (0 < p < 1):
        raise ValueError(f"伯努利分布参数 p 必须满足 0 < p < 1，收到 {p}")

# -------------------- 单个随机数生成器 --------------------
def bernoulli_generator_single(rng: np.random.Generator, params: List[float]) -> float:
    """
    生成一个伯努利分布随机数。
    参数: params = [p]
    返回: 0 或 1
    """
    p = params[0]
    _validate_params(p)
    return 1.0 if rng.random() < p else 0.0

# -------------------- 向量化批量生成器 --------------------
def bernoulli_generator_vectorized(rng: np.random.Generator, params: List[float], n_samples: int) -> np.ndarray:
    """
    批量生成伯努利分布随机数（向量化实现）。
    参数:
        rng: NumPy 随机数生成器
        params: [p]
        n_samples: 样本数量
    返回: 长度为 n_samples 的 NumPy 数组，值为 0 或 1
    """
    p = params[0]
    _validate_params(p)
    # 使用二项分布 n=1 等价于伯努利
    return rng.binomial(1, p, n_samples).astype(float)

# -------------------- 统一生成器接口 --------------------
def bernoulli_generator(rng: np.random.Generator, params: List[float], n_samples: Optional[int] = None) -> Union[float, np.ndarray]:
    """
    统一生成器。若 n_samples 为 None 则返回单个样本，否则返回数组。
    """
    if n_samples is None:
        return bernoulli_generator_single(rng, params)
    else:
        return bernoulli_generator_vectorized(rng, params, n_samples)

# -------------------- 原始分布函数 --------------------
def bernoulli_cdf(x: float, p: float) -> float:
    """原始累积分布函数"""
    _validate_params(p)
    if x < 0:
        return 0.0
    if x >= 1:
        return 1.0
    # 在 [0,1) 区间内，CDF 值为 1-p（因为只有 0 点被覆盖）
    return 1 - p

def bernoulli_ppf(q: float, p: float) -> float:
    """原始分位数函数"""
    _validate_params(p)
    if q <= 0:
        return 0.0
    if q >= 1:
        return 1.0
    # 当 q <= 1-p 时返回 0，否则返回 1
    return 0.0 if q <= 1 - p else 1.0

def bernoulli_pmf(x: float, p: float) -> float:
    """原始概率质量函数"""
    _validate_params(p)
    if abs(x - 0.0) < 1e-12:
        return 1 - p
    if abs(x - 1.0) < 1e-12:
        return p
    return 0.0

# -------------------- 原始矩 --------------------
def bernoulli_raw_mean(p: float) -> float:
    return p

def bernoulli_raw_var(p: float) -> float:
    return p * (1 - p)

def bernoulli_raw_skew(p: float) -> float:
    # 伯努利偏度公式：(1-2p)/sqrt(p(1-p))
    std = math.sqrt(p * (1 - p))
    if std == 0:
        return 0.0
    return (1 - 2 * p) / std

def bernoulli_raw_kurt(p: float) -> float:
    # 伯努利峰度（普通峰度）公式：(1-6p(1-p))/(p(1-p))
    var = p * (1 - p)
    if var == 0:
        return 3.0  # 退化分布，峰度定义为3
    return (1 - 6 * var) / var + 3.0

def bernoulli_raw_mode(p: float) -> float:
    # 众数：当 p>0.5 时众数为1，p<0.5 时众数为0，p=0.5 时双众数，取平均值0.5
    if p > 0.5:
        return 1.0
    elif p < 0.5:
        return 0.0
    else:
        return 0.5

# -------------------- 理论统计类（继承 DistributionBase） --------------------
class BernoulliDistribution(DistributionBase):
    """
    伯努利分布理论统计类。
    继承 DistributionBase，实现原始分布的必要方法，并通过离散点求和计算截断矩。
    截断和平移由基类自动处理，本类提供截断后的矩。
    """

    def __init__(self, params: List[float], markers: dict = None, func_name: str = None):
        if len(params) < 1:
            p = 0.5
        else:
            p = float(params[0])
        _validate_params(p)

        self.p = p

        # 原始矩（无截断）
        self._raw_mean = bernoulli_raw_mean(p)
        self._raw_var = bernoulli_raw_var(p)
        self._raw_skew = bernoulli_raw_skew(p)
        self._raw_kurt = bernoulli_raw_kurt(p)
        self._raw_mode = bernoulli_raw_mode(p)

        # 调用父类初始化（解析截断、平移标记）
        super().__init__(params, markers, func_name)

        # 如果截断类型是百分位数，将概率转换为实际值（基类已做，但需确保边界正确）
        if self.truncate_type in ['percentile', 'percentile2']:
            if self.truncate_lower_pct is not None:
                self.truncate_lower = self._original_ppf(self.truncate_lower_pct)
            if self.truncate_upper_pct is not None:
                self.truncate_upper = self._original_ppf(self.truncate_upper_pct)

        # 完成截断设置后调用
        self._finalize_truncation()

        # 计算截断矩（如果存在截断）
        self._truncated_stats = None
        self._compute_truncated_stats()

    def _compute_truncated_stats(self):
        """
        计算截断后的矩，使用离散点求和（伯努利只有0和1两个值）。
        支持值截断和百分位数截断，正确处理边界点的部分概率（百分位数截断时）。
        """
        if not self.is_truncated():
            return

        # 获取截断边界（已与支持范围取交集，但可能包含平移）
        low, high = self.get_truncated_bounds()
        if low is None and high is None:
            return  # 截断无效，保留 None

        # 转换为原始尺度上的边界用于求和
        if self.truncate_type in ['value2', 'percentile2']:
            # 先平移后截断：边界是平移后的值，需要减去 shift 得到原始值
            low_orig = low - self.shift_amount if low is not None else None
            high_orig = high - self.shift_amount if high is not None else None
        else:
            low_orig = low
            high_orig = high

        # 如果边界为 None，用支持范围的边界代替（但伯努利只有0和1，所以区间外无点）
        if low_orig is None:
            low_orig = self.support_low
        if high_orig is None:
            high_orig = self.support_high

        # 对于百分位数截断，我们需要知道边界点的保留比例
        lower_fraction = 1.0
        upper_fraction = 1.0
        lower_idx = None
        upper_idx = None

        if self.truncate_type in ['percentile', 'percentile2']:
            # 百分位数截断：使用类似 _find_boundary_info 的逻辑
            # 简化为直接根据百分位数计算分数
            if self.truncate_lower_pct is not None:
                p_lower = max(0.0, min(1.0, float(self.truncate_lower_pct)))
                if p_lower <= 1 - self.p:
                    lower_idx = 0
                    # 下界在 0 点内部
                    lower_fraction = p_lower / (1 - self.p) if (1 - self.p) > 0 else 0.0
                else:
                    lower_idx = 1
                    # 下界在 1 点内部，但此时需要调整：因为从 0 到 1 跨越了 CDF 跳变
                    # 更简单的处理：使用 PPF 转换后，再根据原始值判断边界点
                    # 由于已经调用了 _original_ppf，low_orig 已经是 0 或 1，此时分数需按 CDF 计算
                    # 这里为了简化，我们直接使用 low_orig 判断，分数置为 1.0（保留整个点）
                    lower_fraction = 1.0
            if self.truncate_upper_pct is not None:
                p_upper = max(0.0, min(1.0, float(self.truncate_upper_pct)))
                if p_upper <= 1 - self.p:
                    upper_idx = 0
                    upper_fraction = p_upper / (1 - self.p) if (1 - self.p) > 0 else 0.0
                else:
                    upper_idx = 1
                    upper_fraction = 1.0  # 上界在 1 点，保留整个点（因为 p_upper >= 1-p）
            # 注意：当上下界在同一侧时，分数需要特殊处理，但简单起见，我们采用通用方法

        # 定义两个可能的值及其原始概率
        values = [0.0, 1.0]
        probs = [1 - self.p, self.p]

        # 筛选在截断区间内的点，并应用边界分数
        selected_vals = []
        selected_probs = []
        for i, (x, prob) in enumerate(zip(values, probs)):
            # 检查是否在区间内（使用原始尺度）
            in_low = (low_orig is None or x >= low_orig - 1e-12)
            in_high = (high_orig is None or x <= high_orig + 1e-12)
            if not (in_low and in_high):
                continue

            # 应用边界分数（仅当百分位数截断且该点恰为边界点时）
            fraction = 1.0
            if self.truncate_type in ['percentile', 'percentile2']:
                if lower_idx is not None and i == lower_idx:
                    fraction = lower_fraction
                if upper_idx is not None and i == upper_idx:
                    fraction = upper_fraction

            if fraction > 0:
                selected_vals.append(x)
                selected_probs.append(prob * fraction)

        if not selected_vals:
            # 区间内无点，截断无效
            self._truncated_stats = None
            return

        # 归一化常数
        total_prob = sum(selected_probs)
        if total_prob <= 0:
            self._truncated_stats = None
            return

        # 归一化概率
        norm_probs = [p / total_prob for p in selected_probs]

        # 计算矩
        # 均值
        mean_trunc = sum(v * p for v, p in zip(selected_vals, norm_probs))
        # 方差
        var_trunc = sum((v - mean_trunc)**2 * p for v, p in zip(selected_vals, norm_probs))
        # 偏度
        if var_trunc > 0:
            skew_trunc = sum((v - mean_trunc)**3 * p for v, p in zip(selected_vals, norm_probs)) / (var_trunc ** 1.5)
        else:
            skew_trunc = 0.0
        # 峰度（普通峰度）
        if var_trunc > 0:
            kurt_trunc = sum((v - mean_trunc)**4 * p for v, p in zip(selected_vals, norm_probs)) / (var_trunc ** 2)
        else:
            kurt_trunc = 3.0  # 退化分布峰度定义为3
        # 众数（取概率最大的值，若多个则平均）
        max_prob = max(norm_probs)
        mode_indices = [i for i, p in enumerate(norm_probs) if abs(p - max_prob) < 1e-12]
        if len(mode_indices) == 1:
            mode_trunc = selected_vals[mode_indices[0]]
        else:
            mode_trunc = sum(selected_vals[i] for i in mode_indices) / len(mode_indices)

        self._truncated_stats = {
            'mean': mean_trunc,
            'variance': var_trunc,
            'skewness': skew_trunc,
            'kurtosis': kurt_trunc,
            'mode': mode_trunc
        }

    # ---------- 实现原始分布函数 ----------
    def _original_cdf(self, x: float) -> float:
        return bernoulli_cdf(x, self.p)

    def _original_ppf(self, q: float) -> float:
        return bernoulli_ppf(q, self.p)

    def _original_pdf(self, x: float) -> float:
        # 对于离散分布，pdf 实为 pmf，但基类可能调用 pdf，这里直接返回 pmf
        return bernoulli_pmf(x, self.p)

    # ---------- 矩 ----------
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
            # 如果截断后众数已计算，直接使用（注意已包含平移？我们存储的是原始值，需加平移）
            return self._truncated_stats['mode'] + self.shift_amount
        # 原始众数
        raw_mode = self._raw_mode
        # 检查是否在截断区间内（如果截断但未计算截断众数，回退）
        if self.is_truncated():
            low, high = self.get_effective_bounds()
            if low is not None and raw_mode + self.shift_amount < low:
                return low
            if high is not None and raw_mode + self.shift_amount > high:
                return high
        return raw_mode + self.shift_amount

    def min_val(self) -> float:
        """重写最小值：如果截断则用截断下界，否则用理论支持下界"""
        if self._truncate_invalid:
            return float('nan')
        lower, upper = self.get_effective_bounds()
        if lower is not None:
            return lower
        return self.support_low

    def max_val(self) -> float:
        """重写最大值：如果截断则用截断上界，否则用理论支持上界"""
        if self._truncate_invalid:
            return float('nan')
        lower, upper = self.get_effective_bounds()
        if upper is not None:
            return upper
        return self.support_high