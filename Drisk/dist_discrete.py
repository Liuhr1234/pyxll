# dist_discrete.py
"""
离散分布（Discrete）专属模块
由用户提供 X 和 P 数组（X 值，P 概率，自动归一化）。
用于 Drisk 系统集成，提供随机数生成、CDF/PPF/PMF、理论统计类（含离散求和截断矩）。
完全独立，不依赖 scipy。
包含向量化批量生成函数，适用于索引模拟。
"""

import math
import numpy as np
from typing import List, Optional, Union
from distribution_base import DistributionBase

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

def _validate_discrete_params(x_vals: List[float], p_vals: List[float]) -> None:
    if len(x_vals) != len(p_vals):
        raise ValueError("X 和 P 数组长度必须相等")
    if len(x_vals) == 0:
        raise ValueError("至少需要一个点")
    # 检查概率非负
    if any(p < 0 for p in p_vals):
        raise ValueError("概率值不能为负")
    # 归一化
    total = sum(p_vals)
    if total <= 0:
        raise ValueError("概率和必须大于 0")

def _normalize(p_vals: List[float]) -> List[float]:
    total = sum(p_vals)
    return [p / total for p in p_vals]

# -------------------- 随机数生成器（接受列表参数） --------------------
def discrete_generator_single(rng: np.random.Generator, params: List[float], x_vals: List[float], p_vals: List[float]) -> float:
    # 使用轮盘赌选择，确保 p_vals 是浮点数
    u = rng.random()
    cum = 0.0
    for x, p in zip(x_vals, p_vals):
        cum += p
        if u <= cum:
            return float(x)
    return float(x_vals[-1])

def discrete_generator_vectorized(rng: np.random.Generator, params: List[float], n_samples: int, x_vals: List[float], p_vals: List[float]) -> np.ndarray:
    # 使用 numpy 的 choice 函数
    return rng.choice(x_vals, size=n_samples, p=p_vals).astype(float)

def discrete_generator(rng: np.random.Generator, params: List[float], n_samples: Optional[int] = None, x_vals: List[float] = None, p_vals: List[float] = None) -> Union[float, np.ndarray]:
    """
    统一生成器接口，接受列表参数。
    """
    if x_vals is None or p_vals is None:
        raise ValueError("discrete_generator 需要 x_vals 和 p_vals")
    if n_samples is None:
        return discrete_generator_single(rng, params, x_vals, p_vals)
    else:
        return discrete_generator_vectorized(rng, params, n_samples, x_vals, p_vals)

# -------------------- 原始分布函数（接受列表参数） --------------------
def discrete_pmf(x: float, x_vals: List[float], p_vals: List[float]) -> float:
    for xi, pi in zip(x_vals, p_vals):
        if abs(x - xi) < 1e-12:
            return pi
    return 0.0

def discrete_cdf(x: float, x_vals: List[float], p_vals: List[float]) -> float:
    cum = 0.0
    for xi, pi in zip(x_vals, p_vals):
        if x >= xi - 1e-12:
            cum += pi
        else:
            break
    return cum

def discrete_ppf(q: float, x_vals: List[float], p_vals: List[float]) -> float:
    if q <= 0:
        return x_vals[0]
    if q >= 1:
        return x_vals[-1]
    cum = 0.0
    for xi, pi in zip(x_vals, p_vals):
        cum += pi
        if cum >= q - 1e-12:
            return float(xi)
    return x_vals[-1]

# -------------------- 原始矩（接受列表参数） --------------------
def discrete_raw_mean(x_vals: List[float], p_vals: List[float]) -> float:
    return sum(x * p for x, p in zip(x_vals, p_vals))

def discrete_raw_var(x_vals: List[float], p_vals: List[float]) -> float:
    mean = discrete_raw_mean(x_vals, p_vals)
    return sum((x - mean)**2 * p for x, p in zip(x_vals, p_vals))

def discrete_raw_skew(x_vals: List[float], p_vals: List[float]) -> float:
    mean = discrete_raw_mean(x_vals, p_vals)
    var = discrete_raw_var(x_vals, p_vals)
    if var == 0:
        return 0.0
    m3 = sum((x - mean)**3 * p for x, p in zip(x_vals, p_vals))
    return m3 / (var ** 1.5)

def discrete_raw_kurt(x_vals: List[float], p_vals: List[float]) -> float:
    mean = discrete_raw_mean(x_vals, p_vals)
    var = discrete_raw_var(x_vals, p_vals)
    if var == 0:
        return 3.0
    m4 = sum((x - mean)**4 * p for x, p in zip(x_vals, p_vals))
    return m4 / (var ** 2)

def discrete_raw_mode(x_vals: List[float], p_vals: List[float]) -> float:
    max_prob = max(p_vals)
    indices = [i for i, p in enumerate(p_vals) if abs(p - max_prob) < 1e-12]
    if len(indices) == 1:
        return x_vals[indices[0]]
    else:
        # 多个众数，返回平均值
        return sum(x_vals[i] for i in indices) / len(indices)

# -------------------- 理论统计类 --------------------
class DiscreteDistribution(DistributionBase):
    """
    离散分布理论统计类。
    继承 DistributionBase，通过离散点求和计算截断矩。
    参数：x_vals 和 p_vals 应为列表（由基类通过 markers 传入）
    """

    def __init__(self, params: List[float], markers: dict = None, func_name: str = None):
        if markers is None:
            markers = {}
        # 从 markers 获取 x_vals 和 p_vals（已解析为列表）
        x_vals = markers.get('x_vals')
        p_vals = markers.get('p_vals')
        
        # 修改处：UI 层（ui_modeler.py）提取到的单元格引用数据（如 E5:J5 展平后的字符串）是存放在 params 列表中传给后端的。
        # 由于 markers 字典里根本没有这两项数据，触发了 if x_vals is None or p_vals is None: 条件，直接退化到了写死的默认示例 [1, 2, 3] 和 [0.2, 0.3, 0.5]
        # 增加对 params 的读取兜底逻辑
        if x_vals is None or p_vals is None:
            if params and len(params) >= 2:
                x_vals = params[0]
                p_vals = params[1]
            else:
                # 默认示例
                x_vals = [1, 2, 3]
                p_vals = [0.2, 0.3, 0.5]

        # 统一进行字符串解析（如果从 params 读出来的是逗号分隔的字符串）
        if isinstance(x_vals, str):
            x_vals = _extract_numbers_from_input(x_vals)
        if isinstance(p_vals, str):
            p_vals = _extract_numbers_from_input(p_vals)

        # 确保 x_vals 和 p_vals 是浮点数列表
        x_vals = [float(x) for x in x_vals]
        p_vals = [float(p) for p in p_vals]
        _validate_discrete_params(x_vals, p_vals)
        # 归一化概率
        total = sum(p_vals)
        p_vals = [p / total for p in p_vals]

        self.x_vals = x_vals
        self.p_vals = p_vals
        # 存储字符串以备后用（例如用于输出）
        self.x_vals_str = ','.join(str(x) for x in x_vals)
        self.p_vals_str = ','.join(str(p) for p in p_vals)

        self._raw_mean = discrete_raw_mean(x_vals, p_vals)
        self._raw_var = discrete_raw_var(x_vals, p_vals)
        self._raw_skew = discrete_raw_skew(x_vals, p_vals)
        self._raw_kurt = discrete_raw_kurt(x_vals, p_vals)
        self._raw_mode = discrete_raw_mode(x_vals, p_vals)

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
        for i, (x, p) in enumerate(zip(self.x_vals, self.p_vals)):
            if cum <= pct < cum + p:
                fraction = (pct - cum) / p if p > 0 else 0.0
                return i, fraction
            cum += p
        return len(self.x_vals)-1, 1.0

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

        # 计算矩
        mean_trunc = sum(v * p for v, p in zip(values, norm_probs))
        var_trunc = sum((v - mean_trunc)**2 * p for v, p in zip(values, norm_probs))
        if var_trunc > 0:
            skew_trunc = sum((v - mean_trunc)**3 * p for v, p in zip(values, norm_probs)) / (var_trunc ** 1.5)
            kurt_trunc = sum((v - mean_trunc)**4 * p for v, p in zip(values, norm_probs)) / (var_trunc ** 2)
        else:
            skew_trunc = 0.0
            kurt_trunc = 3.0

        # 众数
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
        return discrete_cdf(x, self.x_vals, self.p_vals)

    def _original_ppf(self, q: float) -> float:
        return discrete_ppf(q, self.x_vals, self.p_vals)

    def _original_pdf(self, x: float) -> float:
        # 离散分布返回 PMF
        return discrete_pmf(x, self.x_vals, self.p_vals)

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
        if self._truncate_invalid:
            return float('nan')
        lower, upper = self.get_effective_bounds()
        if lower is not None:
            return lower
        return min(self.x_vals) + self.shift_amount

    def max_val(self) -> float:
        if self._truncate_invalid:
            return float('nan')
        lower, upper = self.get_effective_bounds()
        if upper is not None:
            return upper
        return max(self.x_vals) + self.shift_amount

# 在 dist_discrete.py 末尾添加
def discrete_truncated_mean(x_vals: List[float], p_vals: List[float], lower: Optional[float], upper: Optional[float]) -> float:
    """
    计算离散分布截断均值。
    x_vals: 离散点值
    p_vals: 对应概率（已归一化）
    lower, upper: 原始尺度上的截断边界
    """
    total = 0.0
    norm = 0.0
    for x, p in zip(x_vals, p_vals):
        in_lower = lower is None or x >= lower - 1e-12
        in_upper = upper is None or x <= upper + 1e-12
        if in_lower and in_upper:
            total += x * p
            norm += p
    if norm <= 0:
        # 无点，返回边界中点
        low = lower if lower is not None else min(x_vals)
        high = upper if upper is not None else max(x_vals)
        return (low + high) / 2
    return total / norm