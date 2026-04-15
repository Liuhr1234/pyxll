# dist_cumul.py
"""
累积分布（Cumul）专属模块
由用户提供 X 和 P 数组（严格递增，P 从 0 到 1），构建分段线性 CDF。
用于 Drisk 系统集成，提供随机数生成、CDF/PPF/PDF、理论统计类（含数值积分截断矩）。
完全独立，不依赖 scipy。
包含向量化批量生成函数，适用于索引模拟。
"""

import math
import numpy as np
import traceback  # 新增：用于打印异常堆栈
from typing import List, Optional, Union, Tuple
from distribution_base import DistributionBase

# 允许的浮点误差
EPS = 1e-12

# -------------------- 辅助函数：从各种输入中提取数值 --------------------
def _extract_numbers_from_input(val):
    """
    从各种输入（字符串、元组、列表）中提取数值列表。
    返回浮点数列表。
    """
    numbers = []

    def extract(item):
        if isinstance(item, (tuple, list, np.ndarray)):
            for sub in item:
                extract(sub)
        elif isinstance(item, str):
            clean = item.strip()
            if clean.startswith('{') and clean.endswith('}'):
                inner = clean[1:-1].strip()
                parts = [p.strip() for p in inner.split(',') if p.strip()]
                for p in parts:
                    try:
                        numbers.append(float(p))
                    except:
                        pass
            else:
                if ',' in clean:
                    parts = [p.strip() for p in clean.split(',') if p.strip()]
                    for p in parts:
                        try:
                            numbers.append(float(p))
                        except:
                            pass
                else:
                    try:
                        numbers.append(float(clean))
                    except:
                        pass
        else:
            try:
                if item is not None:
                    numbers.append(float(item))
            except:
                pass

    extract(val)
    return numbers

def _parse_arrays(x_input, p_input) -> Tuple[List[float], List[float]]:
    """
    将输入解析为浮点数列表，并进行验证。
    x_input 和 p_input 可以是字符串、元组、列表等。
    要求 x 严格递增（允许 EPS 误差），p 严格递增，且 p[0]==0, p[-1]==1。
    若无效则抛出 ValueError。
    """
    x_vals = _extract_numbers_from_input(x_input)
    p_vals = _extract_numbers_from_input(p_input)

    if len(x_vals) != len(p_vals):
        raise ValueError("X 和 P 数组长度必须相等")
    if len(x_vals) < 2:
        raise ValueError("至少需要两个点")
    # 检查 X 严格递增（允许误差）
    for i in range(len(x_vals)-1):
        if x_vals[i] >= x_vals[i+1] - EPS:
            raise ValueError(f"X 值必须严格递增，发现 {x_vals[i]} >= {x_vals[i+1]}")
    # 检查 P 严格递增（允许误差）
    for i in range(len(p_vals)-1):
        if p_vals[i] >= p_vals[i+1] - EPS:
            raise ValueError(f"P 值必须严格递增，发现 {p_vals[i]} >= {p_vals[i+1]}")
    if not all(0 - EPS <= p <= 1 + EPS for p in p_vals):
        raise ValueError("P 值必须在 [0,1] 范围内")
    if abs(p_vals[0] - 0.0) > EPS:
        raise ValueError("第一个 P 值必须为 0")
    if abs(p_vals[-1] - 1.0) > EPS:
        raise ValueError("最后一个 P 值必须为 1")
    return x_vals, p_vals

def _parse_internal_arrays(x_input, p_input) -> Tuple[List[float], List[float]]:
    """
    解析内部点（不强制首尾为0和1），仅验证递增和范围 [0,1]。
    若无效则抛出 ValueError。
    """
    x_vals = _extract_numbers_from_input(x_input)
    p_vals = _extract_numbers_from_input(p_input)

    if len(x_vals) != len(p_vals):
        raise ValueError("X 和 P 数组长度必须相等")
    if len(x_vals) < 1:
        raise ValueError("至少需要一个内部点")
    for i in range(len(x_vals)-1):
        if x_vals[i] >= x_vals[i+1] - EPS:
            raise ValueError(f"内部 X 值必须严格递增，发现 {x_vals[i]} >= {x_vals[i+1]}")
    for i in range(len(p_vals)-1):
        if p_vals[i] >= p_vals[i+1] - EPS:
            raise ValueError(f"内部 P 值必须严格递增，发现 {p_vals[i]} >= {p_vals[i+1]}")
    if not all(0 - EPS <= p <= 1 + EPS for p in p_vals):
        raise ValueError("内部 P 值必须在 [0,1] 范围内")
    return x_vals, p_vals

# -------------------- 内部函数：使用浮点数列表操作 --------------------
def _find_segment(x: float, x_vals: List[float]) -> int:
    """返回 x 所在区间索引 i 使得 x_vals[i] <= x <= x_vals[i+1]；若 x < x_vals[0] 返回 -1，若 x > x_vals[-1] 返回 len-1"""
    if x < x_vals[0] - EPS:
        return -1
    if x > x_vals[-1] + EPS:
        return len(x_vals) - 1
    # 二分查找
    lo, hi = 0, len(x_vals) - 1
    while lo <= hi:
        mid = (lo + hi) // 2
        if x < x_vals[mid] - EPS:
            hi = mid - 1
        elif x > x_vals[mid+1] + EPS:
            lo = mid + 1
        else:
            return mid
    return -1  # 不应发生

def _interpolate(x: float, i: int, x_vals: List[float], p_vals: List[float]) -> float:
    """在区间 i 内线性插值 CDF"""
    if i == -1:
        return 0.0
    if i == len(x_vals) - 1:
        return 1.0
    x1, x2 = x_vals[i], x_vals[i+1]
    p1, p2 = p_vals[i], p_vals[i+1]
    return p1 + (p2 - p1) * (x - x1) / (x2 - x1)

def _inverse_interpolate(q: float, x_vals: List[float], p_vals: List[float]) -> float:
    """给定概率 q，返回对应分位数"""
    if q <= 0 + EPS:
        return x_vals[0]
    if q >= 1 - EPS:
        return x_vals[-1]
    # 找到 q 所在的区间
    lo, hi = 0, len(p_vals) - 1
    while lo < hi:
        mid = (lo + hi) // 2
        if q < p_vals[mid] - EPS:
            hi = mid
        elif q > p_vals[mid+1] + EPS:
            lo = mid + 1
        else:
            lo = mid
            break
    i = lo
    if i == len(p_vals) - 1:
        i -= 1
    x1, x2 = x_vals[i], x_vals[i+1]
    p1, p2 = p_vals[i], p_vals[i+1]
    return x1 + (x2 - x1) * (q - p1) / (p2 - p1)

def _pdf(x: float, x_vals: List[float], p_vals: List[float]) -> float:
    """返回 x 处的 PDF 值（分段常数）"""
    i = _find_segment(x, x_vals)
    if i == -1 or i == len(x_vals) - 1:
        return 0.0
    x1, x2 = x_vals[i], x_vals[i+1]
    p1, p2 = p_vals[i], p_vals[i+1]
    if x2 <= x1 + EPS:
        return 0.0
    return (p2 - p1) / (x2 - x1)

# -------------------- 公共函数（接受列表参数）--------------------
def cumul_cdf(x: float, x_vals: List[float], p_vals: List[float]) -> float:
    """累积分布函数（接受列表参数）"""
    i = _find_segment(x, x_vals)
    if i == -1:
        return 0.0
    if i == len(x_vals) - 1:
        return 1.0
    return _interpolate(x, i, x_vals, p_vals)

def cumul_ppf(q: float, x_vals: List[float], p_vals: List[float]) -> float:
    """分位数函数（接受列表参数）"""
    return _inverse_interpolate(q, x_vals, p_vals)

def cumul_pdf(x: float, x_vals: List[float], p_vals: List[float]) -> float:
    """概率密度函数（接受列表参数）"""
    return _pdf(x, x_vals, p_vals)

# -------------------- 随机数生成器（接受列表参数） --------------------
def cumul_generator_single(rng: np.random.Generator, params: List[float], x_vals: List[float], p_vals: List[float]) -> float:
    """生成单个随机数"""
    u = rng.random()
    return _inverse_interpolate(u, x_vals, p_vals)

def cumul_generator_vectorized(rng: np.random.Generator, params: List[float], n_samples: int, x_vals: List[float], p_vals: List[float]) -> np.ndarray:
    """批量生成随机数（向量化）"""
    u = rng.random(n_samples)
    samples = np.empty(n_samples, dtype=float)
    for i in range(n_samples):
        samples[i] = _inverse_interpolate(u[i], x_vals, p_vals)
    return samples

def cumul_generator(rng: np.random.Generator, params: List[float], n_samples: Optional[int] = None, x_vals: List[float] = None, p_vals: List[float] = None) -> Union[float, np.ndarray]:
    """统一生成器接口，接受列表参数"""
    if x_vals is None or p_vals is None:
        raise ValueError("cumul_generator 需要 x_vals 和 p_vals")
    if n_samples is None:
        return cumul_generator_single(rng, params, x_vals, p_vals)
    else:
        return cumul_generator_vectorized(rng, params, n_samples, x_vals, p_vals)

# -------------------- 截断均值计算函数（供索引模拟使用）--------------------
def cumul_truncated_mean(x_vals: List[float], p_vals: List[float], lower: Optional[float], upper: Optional[float]) -> float:
    """
    计算累积分布截断均值。
    lower, upper: 原始尺度上的截断边界，可为 None。
    """
    total = 0.0
    norm = 0.0
    for i in range(len(x_vals) - 1):
        x1, x2 = x_vals[i], x_vals[i+1]
        p1, p2 = p_vals[i], p_vals[i+1]
        slope = (p2 - p1) / (x2 - x1)  # 常数 pdf

        a = max(x1, lower) if lower is not None else x1
        b = min(x2, upper) if upper is not None else x2
        if a >= b - EPS:
            continue

        # 积分 x * slope dx = slope * (b^2 - a^2)/2
        total += slope * (b**2 - a**2) / 2
        norm += slope * (b - a)

    if norm <= 0:
        # 截断区间内无质量，退化为单点（取中点）
        low = lower if lower is not None else x_vals[0]
        high = upper if upper is not None else x_vals[-1]
        return (low + high) / 2
    return total / norm

# -------------------- 理论统计类（继承 DistributionBase）--------------------
class CumulDistribution(DistributionBase):
    """
    累积分布理论统计类。
    继承 DistributionBase，通过数值积分计算截断矩。
    参数：params 应包含四个元素：[min, max, x_str, p_str]
    markers 可包含其他属性，但 x_vals 和 p_vals 优先从 markers 中获取（若存在）。
    """

    def __init__(self, params: List[float], markers: dict = None, func_name: str = None):
        self._invalid = False
        try:
            # ===== 优先使用 markers 中已构建好的完整数组 =====
            if markers and 'x_vals' in markers and 'p_vals' in markers:
                # 从 markers 获取完整数组（可能为字符串或列表）
                x_vals = markers['x_vals']
                p_vals = markers['p_vals']
                # 统一解析为数值列表
                if isinstance(x_vals, str):
                    x_vals = _extract_numbers_from_input(x_vals)
                if isinstance(p_vals, str):
                    p_vals = _extract_numbers_from_input(p_vals)
                # 验证完整数组（严格递增、端点正确）
                x_vals, p_vals = _parse_arrays(x_vals, p_vals)
                self.x_vals = x_vals
                self.p_vals = p_vals
                self.x_str = ','.join(str(x) for x in x_vals)
                self.p_str = ','.join(str(p) for p in p_vals)
                # 获取 min_val, max_val（用于后续检查，但已不再需要构造）
                min_val = x_vals[0]
                max_val = x_vals[-1]
            else:
                # 原有逻辑：从 params 解析内部点并构造完整数组
                if len(params) < 4:
                    min_val = 0.0
                    max_val = 1.0
                    x_str = "0,0.5,1"
                    p_str = "0,0.5,1"
                else:
                    min_val = float(params[0])
                    max_val = float(params[1])
                    x_str = str(params[2])
                    p_str = str(params[3])

                if min_val >= max_val - EPS:
                    raise ValueError("最小值必须小于最大值")

                # 解析内部点（不强制0和1）
                try:
                    x_inner, p_inner = _parse_internal_arrays(x_str, p_str)
                except ValueError as e:
                    raise ValueError(f"内部点解析失败: {e}")

                # 验证内部点范围
                for x in x_inner:
                    if not (min_val - EPS <= x <= max_val + EPS):
                        raise ValueError(f"内部 X 值 {x} 不在 [{min_val}, {max_val}] 范围内")

                # 构造完整的数组：包含边界点，并自动过滤掉等于边界的内部点（避免重复）
                x_full = [min_val]
                p_full = [0.0]
                for xi, pi in zip(x_inner, p_inner):
                    if abs(xi - min_val) > EPS and abs(xi - max_val) > EPS:
                        x_full.append(xi)
                        p_full.append(pi)
                    else:
                        # 若内部点等于边界，则忽略该点（仅给出警告）
                        pass
                x_full.append(max_val)
                p_full.append(1.0)

                # 用 _parse_arrays 验证完整数组（检查严格递增、端点）
                x_vals, p_vals = _parse_arrays(x_full, p_full)

                self.x_vals = x_vals
                self.p_vals = p_vals
                self.x_str = ','.join(str(x) for x in x_vals)
                self.p_str = ','.join(str(p) for p in p_vals)

            # 原始矩（数值积分）
            self._raw_mean = self._integrate_moment(1)
            self._raw_var = self._integrate_central_moment(2)
            if self._raw_var > 0:
                self._raw_skew = self._integrate_central_moment(3) / (self._raw_var ** 1.5)
                self._raw_kurt = self._integrate_central_moment(4) / (self._raw_var ** 2)
            else:
                self._raw_skew = 0.0
                self._raw_kurt = 3.0

            # 众数：PDF 最大值对应的 x，取最大斜率区间中点
            slopes = [(self.p_vals[i+1] - self.p_vals[i]) / (self.x_vals[i+1] - self.x_vals[i]) for i in range(len(self.x_vals)-1)]
            max_slope_idx = np.argmax(slopes)
            self._raw_mode = (self.x_vals[max_slope_idx] + self.x_vals[max_slope_idx+1]) / 2

        except Exception as e:
            # 打印详细错误信息以帮助定位问题
            traceback.print_exc()
            self._invalid = True
            self.x_vals = []
            self.p_vals = []
            self._raw_mean = self._raw_var = self._raw_skew = self._raw_kurt = self._raw_mode = float('nan')
            # 仍然调用父类初始化，但标记已失效
            super().__init__(params, markers, func_name)
            self._truncate_invalid = True  # 强制无效
            return

        super().__init__(params, markers, func_name)

        # 标记截断是否无效
        self._truncate_invalid = False

        if self.truncate_type in ['percentile', 'percentile2']:
            if self.truncate_lower_pct is not None:
                self.truncate_lower = self._original_ppf(self.truncate_lower_pct)
            if self.truncate_upper_pct is not None:
                self.truncate_upper = self._original_ppf(self.truncate_upper_pct)

        self._finalize_truncation()

        self._truncated_stats = None
        self._compute_truncated_stats()

    def _integrate_moment(self, k: int, low: float = None, high: float = None) -> float:
        """计算 ∫ x^k * pdf(x) dx 从 low 到 high"""
        if self._invalid:
            return float('nan')
        if low is None:
            low = self.x_vals[0]
        if high is None:
            high = self.x_vals[-1]
        total = 0.0
        for i in range(len(self.x_vals) - 1):
            x1, x2 = self.x_vals[i], self.x_vals[i+1]
            p1, p2 = self.p_vals[i], self.p_vals[i+1]
            slope = (p2 - p1) / (x2 - x1)
            a = max(x1, low)
            b = min(x2, high)
            if a >= b - EPS:
                continue
            # ∫ x^k * slope dx = slope * (b^(k+1) - a^(k+1)) / (k+1)
            total += slope * (b**(k+1) - a**(k+1)) / (k+1)
        return total

    def _integrate_central_moment(self, k: int, mean: float = None, low: float = None, high: float = None) -> float:
        """计算 ∫ (x - mean)^k * pdf(x) dx"""
        if self._invalid:
            return float('nan')
        if mean is None:
            mean = self._integrate_moment(1, low, high)
        total = 0.0
        for i in range(len(self.x_vals) - 1):
            x1, x2 = self.x_vals[i], self.x_vals[i+1]
            p1, p2 = self.p_vals[i], self.p_vals[i+1]
            slope = (p2 - p1) / (x2 - x1)
            a = max(x1, low) if low is not None else x1
            b = min(x2, high) if high is not None else x2
            if a >= b - EPS:
                continue
            # ∫ (x-mean)^k dx = [ (x-mean)^(k+1) / (k+1) ]_a^b
            total += slope * ((b - mean)**(k+1) - (a - mean)**(k+1)) / (k+1)
        return total

    def _compute_truncated_stats(self):
        """计算截断后的矩，并检查截断有效性"""
        if self._invalid:
            self._truncated_stats = None
            return
        if not self.is_truncated():
            return

        low, high = self.get_truncated_bounds()
        if low is None and high is None:
            return

        # 根据平移模式调整边界到原始尺度
        if self.truncate_type in ['value2', 'percentile2']:
            low_orig = low - self.shift_amount if low is not None else None
            high_orig = high - self.shift_amount if high is not None else None
        else:
            low_orig = low
            high_orig = high

        if low_orig is None:
            low_orig = self.x_vals[0]
        if high_orig is None:
            high_orig = self.x_vals[-1]

        # 检查截断范围是否与支撑域有交集
        if low_orig > self.x_vals[-1] + EPS or high_orig < self.x_vals[0] - EPS:
            self._truncate_invalid = True
            self._truncated_stats = None
            return

        # 如果 low_orig >= high_orig，退化为单点
        if low_orig >= high_orig - EPS:
            self._truncated_stats = {
                'mean': low_orig,
                'variance': 0.0,
                'skewness': 0.0,
                'kurtosis': 3.0,
                'mode': low_orig
            }
            return

        # 计算归一化常数 Z
        Z = self._original_cdf(high_orig) - self._original_cdf(low_orig)
        if Z <= EPS:
            self._truncate_invalid = True
            self._truncated_stats = None
            return

        mean_trunc = self._integrate_moment(1, low_orig, high_orig) / Z
        var_trunc = self._integrate_central_moment(2, mean_trunc, low_orig, high_orig) / Z
        if var_trunc > EPS:
            skew_trunc = self._integrate_central_moment(3, mean_trunc, low_orig, high_orig) / Z / (var_trunc ** 1.5)
            kurt_trunc = self._integrate_central_moment(4, mean_trunc, low_orig, high_orig) / Z / (var_trunc ** 2)
        else:
            skew_trunc = 0.0
            kurt_trunc = 3.0
        # 众数：取截断区间内 PDF 最大值点，这里简化为区间中点
        mode_trunc = (low_orig + high_orig) / 2

        self._truncated_stats = {
            'mean': mean_trunc,
            'variance': var_trunc,
            'skewness': skew_trunc,
            'kurtosis': kurt_trunc,
            'mode': mode_trunc
        }

    def _original_cdf(self, x: float) -> float:
        if self._invalid:
            return float('nan')
        return cumul_cdf(x, self.x_vals, self.p_vals)

    def _original_ppf(self, q: float) -> float:
        if self._invalid:
            return float('nan')
        return cumul_ppf(q, self.x_vals, self.p_vals)

    def _original_pdf(self, x: float) -> float:
        if self._invalid:
            return float('nan')
        return cumul_pdf(x, self.x_vals, self.p_vals)

    def mean(self) -> float:
        if self._invalid or self._truncate_invalid:
            return float('nan')
        if self._truncated_stats:
            return self._truncated_stats['mean'] + self.shift_amount
        return self._raw_mean + self.shift_amount

    def variance(self) -> float:
        if self._invalid or self._truncate_invalid:
            return float('nan')
        if self._truncated_stats:
            return self._truncated_stats['variance']
        return self._raw_var

    def skewness(self) -> float:
        if self._invalid or self._truncate_invalid:
            return float('nan')
        if self._truncated_stats:
            return self._truncated_stats['skewness']
        return self._raw_skew

    def kurtosis(self) -> float:
        if self._invalid or self._truncate_invalid:
            return float('nan')
        if self._truncated_stats:
            return self._truncated_stats['kurtosis']
        return self._raw_kurt

    def mode(self) -> float:
        if self._invalid or self._truncate_invalid:
            return float('nan')
        if self._truncated_stats:
            return self._truncated_stats['mode'] + self.shift_amount
        raw_mode = self._raw_mode
        if self.is_truncated():
            low, high = self.get_effective_bounds()
            if low is not None and raw_mode + self.shift_amount < low - EPS:
                return low
            if high is not None and raw_mode + self.shift_amount > high + EPS:
                return high
        return raw_mode + self.shift_amount

    def min_val(self) -> float:
        if self._invalid:
            return float('nan')
        lower, upper = self.get_effective_bounds()
        if lower is not None:
            return lower
        return self.x_vals[0] + self.shift_amount

    def max_val(self) -> float:
        if self._invalid:
            return float('nan')
        lower, upper = self.get_effective_bounds()
        if upper is not None:
            return upper
        return self.x_vals[-1] + self.shift_amount