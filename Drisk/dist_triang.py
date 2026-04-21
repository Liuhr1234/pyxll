# dist_triang.py
"""
三角分布（Triangular）专属模块
用于 Drisk 系统集成，提供随机数生成、CDF/PPF/PDF、理论统计类（含数值积分截断矩）。
完全独立，不依赖 scipy。
包含向量化批量生成函数，适用于索引模拟。
"""

import math
import numpy as np
from typing import List, Optional, Union
from distribution_base import DistributionBase

# -------------------- 参数验证 --------------------
def _validate_params(a: float, c: float, b: float) -> None:
    """验证参数有效性，无效则抛出 ValueError"""
    if not (a < c < b):
        raise ValueError(f"三角分布参数必须满足 a < c < b，收到 a={a}, c={c}, b={b}")

# -------------------- 单个随机数生成器（带参数保护） --------------------
def triang_generator_single(rng: np.random.Generator, params: List[float]) -> float:
    """
    生成一个三角分布随机数。
    参数: params = [a, c, b]
    返回: 一个样本值
    """
    a, c, b = params[0], params[1], params[2]
    # 如果参数顺序错误，尝试交换 c 和 b（防御性）
    if c > b:
        print(f"警告: 三角分布参数顺序可能错误，在生成器中交换 c={c} 与 b={b}")
        c, b = b, c
    _validate_params(a, c, b)

    # 使用逆变换法
    fc = (c - a) / (b - a)
    u = rng.random()
    if u <= fc:
        return a + math.sqrt(u * (b - a) * (c - a))
    else:
        return b - math.sqrt((1 - u) * (b - a) * (b - c))

# -------------------- 向量化批量生成器（带参数保护） --------------------
def triang_generator_vectorized(rng: np.random.Generator, params: List[Union[float, np.ndarray]], n_samples: int) -> np.ndarray:
    a = params[0]
    c = params[1]
    b = params[2]

    # 广播
    if not isinstance(a, np.ndarray):
        a = np.full(n_samples, float(a))
    if not isinstance(c, np.ndarray):
        c = np.full(n_samples, float(c))
    if not isinstance(b, np.ndarray):
        b = np.full(n_samples, float(b))

    # 修正顺序（如果 c > b）
    swap = c > b
    if np.any(swap):
        c_swapped = np.where(swap, b, c)
        b_swapped = np.where(swap, c, b)
        c = c_swapped
        b = b_swapped

    if np.any((a >= c) | (c >= b)):
        raise ValueError("Triangular distribution requires a < c < b")

    fc = (c - a) / (b - a)
    u = rng.random(n_samples)
    left_mask = u <= fc
    samples = np.empty(n_samples, dtype=float)
    samples[left_mask] = a[left_mask] + np.sqrt(u[left_mask] * (b[left_mask] - a[left_mask]) * (c[left_mask] - a[left_mask]))
    right_mask = ~left_mask
    samples[right_mask] = b[right_mask] - np.sqrt((1 - u[right_mask]) * (b[right_mask] - a[right_mask]) * (b[right_mask] - c[right_mask]))
    return samples

# 统一生成器接口
def triang_generator(rng: np.random.Generator, params: List[float], n_samples: Optional[int] = None) -> Union[float, np.ndarray]:
    if n_samples is None:
        return triang_generator_single(rng, params)
    else:
        return triang_generator_vectorized(rng, params, n_samples)

# -------------------- 原始分布函数 --------------------
def triang_cdf(x: float, a: float, c: float, b: float) -> float:
    if x <= a:
        return 0.0
    if x >= b:
        return 1.0
    if x <= c:
        return (x - a) ** 2 / ((b - a) * (c - a))
    else:
        return 1 - (b - x) ** 2 / ((b - a) * (b - c))

def triang_ppf(q: float, a: float, c: float, b: float) -> float:
    if q <= 0:
        return a
    if q >= 1:
        return b
    fc = (c - a) / (b - a)
    if q <= fc:
        return a + math.sqrt(q * (b - a) * (c - a))
    else:
        return b - math.sqrt((1 - q) * (b - a) * (b - c))

def triang_pdf(x, a: float, c: float, b: float):
    """
    三角分布的概率密度函数。
    参数:
        x: 输入值，可以是标量或 numpy 数组
        a: 最小值
        c: 众数
        b: 最大值
    返回: 概率密度值，与输入类型相同
    """
    x_arr = np.asarray(x)
    result = np.zeros_like(x_arr, dtype=float)

    left_mask = (x_arr >= a) & (x_arr <= c)
    if np.any(left_mask):
        result[left_mask] = 2 * (x_arr[left_mask] - a) / ((b - a) * (c - a))

    right_mask = (x_arr > c) & (x_arr <= b)
    if np.any(right_mask):
        result[right_mask] = 2 * (b - x_arr[right_mask]) / ((b - a) * (b - c))

    if np.isscalar(x):
        return float(result.item())
    return result

# -------------------- 原始矩 --------------------
def triang_raw_mean(a: float, c: float, b: float) -> float:
    return (a + b + c) / 3

def triang_raw_var(a: float, c: float, b: float) -> float:
    return (a**2 + b**2 + c**2 - a*b - a*c - b*c) / 18

def triang_raw_skew(a: float, c: float, b: float) -> float:
    numerator = (a + b - 2*c) * (2*c - a - b)
    denom = (b - a) * math.sqrt((b - a)**2 + 2*(c - a)*(b - c))
    if denom == 0:
        return 0.0
    return (math.sqrt(2) / 5) * numerator / denom

def triang_raw_kurt(a: float, c: float, b: float) -> float:
    return 2.4

def triang_raw_mode(a: float, c: float, b: float) -> float:
    return c

# -------------------- 理论统计类 --------------------
class TriangDistribution(DistributionBase):
    """
    三角分布理论统计类。
    继承 DistributionBase，通过数值积分计算截断矩。
    """

    def __init__(self, params: List[float], markers: dict = None, func_name: str = None):
        if len(params) < 3:
            a, c, b = 0.0, 0.5, 1.0
        else:
            a, c, b = float(params[0]), float(params[1]), float(params[2])
        # 如果参数顺序错误，尝试修复
        if c > b:
            print(f"警告: 三角分布参数顺序可能错误，在类初始化中交换 c={c} 与 b={b}")
            c, b = b, c
        _validate_params(a, c, b)

        self.a = a
        self.c = c
        self.b = b

        # ===== 显式设置支持范围（必须在调用父类之前设置） =====
        self.support_low = a
        self.support_high = b
        # ====================================================

        # 原始矩
        self._raw_mean = triang_raw_mean(a, c, b)
        self._raw_var = triang_raw_var(a, c, b)
        self._raw_skew = triang_raw_skew(a, c, b)
        self._raw_kurt = triang_raw_kurt(a, c, b)
        self._raw_mode = triang_raw_mode(a, c, b)

        super().__init__(params, markers, func_name)

        # 处理百分位数截断
        if self.truncate_type in ['percentile', 'percentile2']:
            if self.truncate_lower_pct is not None:
                self.truncate_lower = self._original_ppf(self.truncate_lower_pct)
            if self.truncate_upper_pct is not None:
                self.truncate_upper = self._original_ppf(self.truncate_upper_pct)

        # 完成截断设置后调用
        self._finalize_truncation()

        self._truncated_stats = None
        self._compute_truncated_stats()

    # ---------- 数值积分辅助 ----------
    def _adaptive_simpsons(self, f, a, b, eps=1e-10, max_depth=20):
        """自适应 Simpson 积分（内部使用）"""
        def simpson_rule(a, b):
            c = (a + b) / 2
            h = (b - a) / 6
            return h * (f(a) + 4*f(c) + f(b))

        def _recursive(a, b, fa, fb, fc, I, depth):
            c = (a + b) / 2
            d = (a + c) / 2
            e = (c + b) / 2
            fd = f(d)
            fe = f(e)
            I_left = (c - a) / 6 * (fa + 4*fd + fc)
            I_right = (b - c) / 6 * (fc + 4*fe + fb)
            if depth >= max_depth or abs(I_left + I_right - I) < 15 * eps:
                return I_left + I_right
            return _recursive(a, c, fa, fc, fd, I_left, depth+1) + \
                   _recursive(c, b, fc, fb, fe, I_right, depth+1)

        if a >= b:
            return 0.0
        fa = f(a)
        fb = f(b)
        fc = f((a + b) / 2)
        I = simpson_rule(a, b)
        return _recursive(a, b, fa, fb, fc, I, 0)

    def _compute_truncated_stats(self):
        """计算截断后的矩（数值积分）"""
        if not self.is_truncated():
            return

        low, high = self.get_truncated_bounds()
        if low is None and high is None:
            return

        # 转换为原始尺度上的边界用于积分
        if self.truncate_type in ['value2', 'percentile2']:
            low_orig = low - self.shift_amount if low is not None else None
            high_orig = high - self.shift_amount if high is not None else None
        else:
            low_orig = low
            high_orig = high

        if low_orig is None:
            low_orig = self.support_low
        if high_orig is None:
            high_orig = self.support_high

        if low_orig >= high_orig:
            self._truncated_stats = {
                'mean': low_orig + self.shift_amount,
                'variance': 0.0,
                'skewness': 0.0,
                'kurtosis': 3.0,
                'mode': low_orig + self.shift_amount
            }
            return

        # 归一化常数 Z
        Z = triang_cdf(high_orig, self.a, self.c, self.b) - triang_cdf(low_orig, self.a, self.c, self.b)
        if Z <= 0:
            self._truncated_stats = None
            return

        # 定义被积函数 x^k * pdf(x)
        def make_integrand(k):
            return lambda x: (x ** k) * triang_pdf(x, self.a, self.c, self.b)

        try:
            m1 = self._adaptive_simpsons(make_integrand(1), low_orig, high_orig) / Z
            m2 = self._adaptive_simpsons(make_integrand(2), low_orig, high_orig) / Z
            m3 = self._adaptive_simpsons(make_integrand(3), low_orig, high_orig) / Z
            m4 = self._adaptive_simpsons(make_integrand(4), low_orig, high_orig) / Z

            var = m2 - m1**2
            if var <= 0:
                skew = 0.0
                kurt = 3.0
            else:
                mu3 = m3 - 3*m1*m2 + 2*m1**3
                mu4 = m4 - 4*m1*m3 + 6*m1**2*m2 - 3*m1**4
                skew = mu3 / (var ** 1.5)
                kurt = mu4 / (var ** 2)

            # ========== 修正：计算截断后的众数（根据原始众数位置） ==========
            c = self.c  # 原始众数
            if low_orig <= c <= high_orig:
                mode = c
            elif c < low_orig:
                mode = low_orig
            else:  # c > high_orig
                mode = high_orig
            # ============================================================

            self._truncated_stats = {
                'mean': m1,
                'variance': var,
                'skewness': skew,
                'kurtosis': kurt,
                'mode': mode
            }
        except Exception:
            self._truncated_stats = None

    # ---------- 实现原始分布函数 ----------
    def _original_cdf(self, x: float) -> float:
        return triang_cdf(x, self.a, self.c, self.b)

    def _original_ppf(self, q: float) -> float:
        return triang_ppf(q, self.a, self.c, self.b)

    def _original_pdf(self, x: float) -> float:
        return triang_pdf(x, self.a, self.c, self.b)

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
            return self._truncated_stats['mode'] + self.shift_amount
        raw_mode = self._raw_mode
        if self.is_truncated():
            low, high = self.get_effective_bounds()
            if low is not None and raw_mode + self.shift_amount < low:
                return low
            if high is not None and raw_mode + self.shift_amount > high:
                return high
        return raw_mode + self.shift_amount

    def min_val(self) -> float:
        lower, upper = self.get_effective_bounds()
        if lower is not None:
            return lower
        return self.apply_shift(self.support_low)

    def max_val(self) -> float:
        lower, upper = self.get_effective_bounds()
        if upper is not None:
            return upper
        return self.apply_shift(self.support_high)


def triang_truncated_mean(a: float, c: float, b: float, lower: Optional[float], upper: Optional[float]) -> float:
    """
    计算三角分布截断均值。
    lower, upper: 原始尺度上的截断边界（不含平移），可为 None 表示无边界。
    """
    low = a if lower is None else max(lower, a)
    high = b if upper is None else min(upper, b)
    if low >= high:
        return low

    Z = triang_cdf(high, a, c, b) - triang_cdf(low, a, c, b)
    if Z <= 0:
        return (low + high) / 2

    def integrand(x):
        return x * triang_pdf(x, a, c, b)

    n = 1000
    x = np.linspace(low, high, n)
    dx = (high - low) / (n - 1)
    y = integrand(x)
    integral = np.trapz(y, dx=dx)
    return integral / Z