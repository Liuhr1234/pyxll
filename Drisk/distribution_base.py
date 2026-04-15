# distribution_base.py
"""基础分布类模块 - 包含分布基类和所有内置分布（正态、均匀、伽马等）"""

import math
import numpy as np
from typing import Dict, List, Tuple, Optional, Any
import scipy.stats as sps
from constants import get_distribution_support

# -------------------- 辅助函数 --------------------
def _safe_norm_ppf(p: float) -> float:
    """安全的正态分布逆累积分布函数（备用）"""
    if p <= 0.0:
        return float('-inf')
    if p >= 1.0:
        return float('inf')
    try:
        return sps.norm.ppf(p)
    except:
        # 降级到近似算法（很少用到）
        q = p - 0.5
        if abs(q) <= 0.425:
            r = q * q
            num = (((-25.44106049637 * r + 41.39119773534) * r - 18.61500062529) * r + 2.50662823884) * q
            den = (((3.13082909833 * r + -21.06224101826) * r + 23.08336743743) * r + -8.47351093090) * r + 1.0
            return num / den
        else:
            if q < 0:
                r = p
            else:
                r = 1 - p
            r = math.sqrt(-math.log(r))
            if r <= 5.0:
                r = r - 1.6
                num = (((0.000776 * r + 0.0261) * r + 0.3617) * r + 1.7815) * r + 1.8213
                den = ((0.0034 * r + 0.158) * r + 1.089) * r + 1.0
            else:
                r = r - 5.0
                num = (((0.000026 * r + 0.0011) * r + 0.054) * r + 0.706) * r + 2.307
                den = ((0.00041 * r + 0.036) * r + 0.406) * r + 1.0
            result = num / den
            if q < 0:
                return -result
            else:
                return result

# ==================== 基础分布类（增加支持范围处理） ====================
class DistributionBase:
    """分布基类 - 提供统一的接口（修复版，加入支持范围检查）"""

    def __init__(self, params: List[float], markers: Dict[str, Any] = None, func_name: str = None):
        """
        初始化分布

        Args:
            params: 分布参数列表
            markers: 属性标记字典
            func_name: 分布函数名（用于获取支持范围）
        """
        self.params = params
        self.markers = markers or {}
        self.func_name = func_name

        # 获取分布的理论支持范围
        if func_name:
            self.support_low, self.support_high = get_distribution_support(func_name, params)
        else:
            self.support_low, self.support_high = float('-inf'), float('inf')

        # 平移参数
        self.shift_amount = self._get_shift_amount()
        self.shift_mode = self._get_shift_mode()

        # 截断参数
        self.truncate_type = None  # 'value', 'percentile', 'value2', 'percentile2'
        self.truncate_lower = None
        self.truncate_upper = None
        self.truncate_lower_pct = None  # 百分位数下界
        self.truncate_upper_pct = None  # 百分位数上界

        # 截断有效性标记
        self._truncate_invalid = False

        self._parse_truncate_params()

    def _get_shift_amount(self) -> float:
        """获取平移量"""
        shift = self.markers.get('shift')
        if shift is not None:
            try:
                return float(shift)
            except:
                return 0.0
        return 0.0

    def _get_shift_mode(self) -> str:
        """获取平移模式：before_truncation 或 after_truncation"""
        # 检查是否有truncate2或truncateP2标记
        if self.markers.get('truncate2'):
            return 'before_truncation'  # 先平移后截断
        else:
            return 'after_truncation'  # 先截断后平移

    def _parse_truncate_params(self):
        """解析截断参数 - 修复版：正确处理字符串和数值类型，支持百分位数截断，增强单边截断支持"""
        # 解析truncate/truncateP（先截断后平移）
        truncate_val = self.markers.get('truncate')
        truncate_p_val = self.markers.get('truncate_p') or self.markers.get('truncatep')

        # 解析truncate2/truncateP2（先平移后截断）
        truncate2_val = self.markers.get('truncate2')
        truncate_p2_val = self.markers.get('truncate_p2') or self.markers.get('truncatep2')

        # 处理truncate（按值截断）
        if truncate_val is not None:
            self.truncate_type = 'value'
            self._parse_value_truncate(truncate_val)

        # 处理truncate_p（按百分位数截断）
        elif truncate_p_val is not None:
            self.truncate_type = 'percentile'
            self._parse_percentile_truncate(truncate_p_val)

        # 处理truncate2（按值截断，先平移后截断）
        elif truncate2_val is not None:
            self.truncate_type = 'value2'
            self._parse_value_truncate(truncate2_val)

        # 处理truncate_p2（按百分位数截断，先平移后截断）
        elif truncate_p2_val is not None:
            self.truncate_type = 'percentile2'
            self._parse_percentile_truncate(truncate_p2_val)

    def _parse_value_truncate(self, truncate_val):
        """
        解析按值截断参数 - 增强版：支持单边截断，如 "(,0.8)" 或 "(0.2,)"
        """
        # 确保是字符串
        if not isinstance(truncate_val, str):
            truncate_val = str(truncate_val)

        # 清理字符串：移除空格，移除可能的括号
        truncate_val = truncate_val.strip()
        if truncate_val.startswith('(') and truncate_val.endswith(')'):
            truncate_val = truncate_val[1:-1]

        # 分割参数，保留空字符串（表示缺失）
        parts = [p.strip() for p in truncate_val.split(',')]

        # 初始化变量
        lower = None
        upper = None

        # 处理第一个参数（如果存在且非空）
        if len(parts) >= 1 and parts[0]:
            try:
                lower = float(parts[0])
            except Exception:
                pass

        # 处理第二个参数（如果存在且非空）
        if len(parts) >= 2 and parts[1]:
            try:
                upper = float(parts[1])
            except Exception:
                pass

        # 如果两个参数都为空，则忽略截断
        if lower is None and upper is None:
            return

        # 设置截断边界
        self.truncate_lower = lower
        self.truncate_upper = upper

    def _parse_percentile_truncate(self, truncate_p_val):
        """
        解析按百分位数截断参数 - 增强版：支持单边截断（例如 (0.2,) 或 (,0.8)）
        """
        # 确保是字符串
        if not isinstance(truncate_p_val, str):
            truncate_p_val = str(truncate_p_val)

        truncate_p_val = truncate_p_val.strip()
        if truncate_p_val.startswith('(') and truncate_p_val.endswith(')'):
            truncate_p_val = truncate_p_val[1:-1]

        # 分割参数，保留空字符串（表示缺失）
        parts = [p.strip() for p in truncate_p_val.split(',')]

        # 初始化变量
        lower_pct = None
        upper_pct = None
        valid = True

        # 处理第一个参数（如果存在且非空）
        if len(parts) >= 1 and parts[0]:
            try:
                val = float(parts[0])
                # 检查是否在0-1范围内，否则可能为百分比（如5）
                if 0 <= val <= 1:
                    lower_pct = val
                elif 0 <= val <= 100:
                    lower_pct = val / 100.0
                else:
                    valid = False
            except Exception:
                valid = False

        # 处理第二个参数（如果存在且非空）
        if len(parts) >= 2 and parts[1]:
            try:
                val = float(parts[1])
                if 0 <= val <= 1:
                    upper_pct = val
                elif 0 <= val <= 100:
                    upper_pct = val / 100.0
                else:
                    valid = False
            except Exception:
                valid = False

        # 如果只有一个非空参数且没有第二个逗号（即 parts 长度为1且参数非空）
        if len(parts) == 1 and lower_pct is not None:
            # 单边截断：根据值判断是下界还是上界
            if lower_pct < 0.5:
                upper_pct = None
            else:
                upper_pct = lower_pct
                lower_pct = None

        if not valid:
            self._truncate_invalid = True
            return

        # 根据解析结果设置截断边界
        if lower_pct is not None:
            self.truncate_lower_pct = lower_pct
        if upper_pct is not None:
            self.truncate_upper_pct = upper_pct

    def _finalize_truncation(self):
        """
        完成截断设置后调用，检查截断是否有效（与支撑域有交集）。
        如果无效，设置 _truncate_invalid = True。
        """
        if self.truncate_type is None:
            return
        lower, upper = self.get_truncated_bounds()
        if lower is None and upper is None:
            self._truncate_invalid = True

    def apply_shift(self, x: float) -> float:
        """应用平移"""
        return x + self.shift_amount

    def apply_unshift(self, x: float) -> float:
        """移除平移"""
        return x - self.shift_amount

    def _intersect_with_support(self, lower: Optional[float], upper: Optional[float], is_shifted: bool = False) -> Tuple[Optional[float], Optional[float]]:
        """
        与分布支持范围取交集，返回调整后的边界（保持原尺度）

        Args:
            lower: 截断下界（可为 None）
            upper: 截断上界（可为 None）
            is_shifted: 边界是否已经包含平移（即是否是平移后的值）

        Returns:
            调整后的 (lower, upper)，若交集为空则返回 (None, None)
        """
        if lower is None and upper is None:
            return lower, upper

        # 如果是平移后的值，先转换到原始尺度
        if is_shifted:
            orig_lower = lower - self.shift_amount if lower is not None else None
            orig_upper = upper - self.shift_amount if upper is not None else None
        else:
            orig_lower, orig_upper = lower, upper

        # 取交集
        low_support, high_support = self.support_low, self.support_high

        if orig_lower is not None:
            if low_support == float('-inf'):
                new_lower = orig_lower
            else:
                new_lower = max(orig_lower, low_support)
        else:
            new_lower = None

        if orig_upper is not None:
            if high_support == float('inf'):
                new_upper = orig_upper
            else:
                new_upper = min(orig_upper, high_support)
        else:
            new_upper = None

        # 检查交集是否非空（双边截断）
        if new_lower is not None and new_upper is not None and new_lower > new_upper:
            return None, None  # 交集为空，截断无效

        # ===== 新增：检查单边截断是否完全超出支持范围 =====
        # 如果下界大于支持上界（有限值）
        if new_lower is not None and self.support_high != float('inf') and new_lower > self.support_high:
            return None, None
        # 如果上界小于支持下界（有限值）
        if new_upper is not None and self.support_low != float('-inf') and new_upper < self.support_low:
            return None, None
        # =================================================

        # 如果是平移后的值，再加回 shift
        if is_shifted:
            result_lower = new_lower + self.shift_amount if new_lower is not None else None
            result_upper = new_upper + self.shift_amount if new_upper is not None else None
            return result_lower, result_upper
        else:
            return new_lower, new_upper

    def get_truncated_bounds(self) -> Tuple[Optional[float], Optional[float]]:
        """
        获取截断边界（已考虑平移模式，并与支持范围取交集）
        """
        if not self.truncate_type:
            return None, None

        if self.truncate_type in ['value', 'value2']:
            # 值截断，需要与支持范围取交集
            lower = self.truncate_lower
            upper = self.truncate_upper
            is_shifted = (self.truncate_type == 'value2')
            return self._intersect_with_support(lower, upper, is_shifted)

        elif self.truncate_type in ['percentile', 'percentile2']:
            # 百分位数截断，边界已经通过分位数转换得到值（已在 __init__ 中计算）
            # 理论上分位数一定在支持范围内，但为防止意外，仍可与支持范围取交集（但无需转换平移）
            lower = self.truncate_lower
            upper = self.truncate_upper
            # 由于分位数已在 support 内，直接返回即可，但为安全可调用 _intersect_with_support（不改变值）
            # 注意：对于 percentile2，边界是平移后的值，需要 is_shifted=True
            if self.truncate_type == 'percentile2':
                # percentile2 的边界已经是平移后的值，与支持取交集时需要先减 shift
                return self._intersect_with_support(lower, upper, is_shifted=True)
            else:
                return self._intersect_with_support(lower, upper, is_shifted=False)

        return None, None

    def get_effective_bounds(self) -> Tuple[Optional[float], Optional[float]]:
        """获取有效的截断边界（考虑平移）"""
        if not self.truncate_type:
            return None, None

        if self.truncate_type in ['value', 'percentile']:
            # 先截断后平移：获取原始截断边界，然后应用平移
            lower, upper = self.get_truncated_bounds()
            if lower is not None:
                lower = self.apply_shift(lower)
            if upper is not None:
                upper = self.apply_shift(upper)
            return lower, upper
        else:
            # 先平移后截断：截断边界已经考虑了平移
            if self.truncate_type in ['value2', 'percentile2']:
                return self.get_truncated_bounds()

        return None, None

    def is_truncated(self) -> bool:
        """检查是否有截断"""
        return self.truncate_type is not None

    def is_valid(self) -> bool:
        """检查分布是否有效（参数有效且截断有效）"""
        return not self._truncate_invalid

    def _original_ppf(self, p: float) -> float:
        """原始分布的分位数函数（在子类中实现）"""
        raise NotImplementedError

    def _original_cdf(self, x: float) -> float:
        """原始分布的累积分布函数（在子类中实现）"""
        raise NotImplementedError

    def _original_pdf(self, x: float) -> float:
        """原始分布的概率密度/质量函数（在子类中实现）"""
        raise NotImplementedError

    # 抽象方法
    def mean(self) -> float:
        """均值"""
        raise NotImplementedError

    def variance(self) -> float:
        """方差"""
        raise NotImplementedError

    def std_dev(self) -> float:
        """标准差"""
        var = self.variance()
        return math.sqrt(var) if var >= 0 else 0.0

    def skewness(self) -> float:
        """偏度"""
        raise NotImplementedError

    def kurtosis(self) -> float:
        """峰度（普通峰度，正态分布为3）"""
        raise NotImplementedError

    def mode(self) -> float:
        """众数"""
        raise NotImplementedError

    def min_val(self) -> float:
        """最小值"""
        lower, upper = self.get_effective_bounds()
        if lower is not None:
            return lower
        return float('-inf')

    def max_val(self) -> float:
        """最大值"""
        lower, upper = self.get_effective_bounds()
        if upper is not None:
            return upper
        return float('inf')

    def range_val(self) -> float:
        """极差"""
        min_val = self.min_val()
        max_val = self.max_val()
        if min_val == float('-inf') or max_val == float('inf'):
            return float('inf')
        return max_val - min_val

    # ========== 通用分位数函数实现 ==========
    def ppf(self, p: float) -> float:
        """分位数函数（通用实现，依赖子类的 _original_ppf 和 _original_cdf）"""
        if p < 0 or p > 1:
            return float('nan')

        if not self.is_truncated():
            result = self._original_ppf(p)
        else:
            lower, upper = self.get_truncated_bounds()
            # 根据平移模式确定用于CDF的原始边界
            if self.truncate_type in ['value', 'percentile']:
                lower_p = 0.0 if lower is None else self._original_cdf(lower)
                upper_p = 1.0 if upper is None else self._original_cdf(upper)
            else:  # value2, percentile2
                if lower is not None:
                    lower_orig = lower - self.shift_amount
                else:
                    lower_orig = None
                if upper is not None:
                    upper_orig = upper - self.shift_amount
                else:
                    upper_orig = None
                lower_p = 0.0 if lower_orig is None else self._original_cdf(lower_orig)
                upper_p = 1.0 if upper_orig is None else self._original_cdf(upper_orig)

            original_p = lower_p + p * (upper_p - lower_p)
            original_p = max(0.0, min(1.0, original_p))
            result = self._original_ppf(original_p)

            # 确保结果在截断范围内
            if self.truncate_type in ['value', 'percentile']:
                if lower is not None and result < lower:
                    result = lower
                if upper is not None and result > upper:
                    result = upper
            else:  # value2, percentile2
                shifted_result = result + self.shift_amount
                if lower is not None and shifted_result < lower:
                    result = lower - self.shift_amount
                if upper is not None and shifted_result > upper:
                    result = upper - self.shift_amount

        return self.apply_shift(result)

    # ========== 新增：通用累积分布函数实现 ==========
    def cdf(self, x: float) -> float:
        """累积分布函数（通用实现，依赖子类的 _original_cdf）"""
        x_original = self.apply_unshift(x)
        if not self.is_truncated():
            return self._original_cdf(x_original)

        lower, upper = self.get_truncated_bounds()
        # 根据平移模式确定用于CDF的原始边界
        if self.truncate_type in ['value', 'percentile']:
            lower_p = 0.0 if lower is None else self._original_cdf(lower)
            upper_p = 1.0 if upper is None else self._original_cdf(upper)
            # 检查 x 是否在截断范围内（原始尺度）
            if lower is not None and x_original < lower:
                return 0.0
            if upper is not None and x_original > upper:
                return 1.0
        else:  # value2, percentile2
            if lower is not None:
                lower_orig = lower - self.shift_amount
            else:
                lower_orig = None
            if upper is not None:
                upper_orig = upper - self.shift_amount
            else:
                upper_orig = None
            lower_p = 0.0 if lower_orig is None else self._original_cdf(lower_orig)
            upper_p = 1.0 if upper_orig is None else self._original_cdf(upper_orig)
            # 检查 x 是否在截断范围内（x 是平移后的值）
            if lower is not None and x < lower:
                return 0.0
            if upper is not None and x > upper:
                return 1.0

        orig_cdf = self._original_cdf(x_original)
        return (orig_cdf - lower_p) / (upper_p - lower_p)

    # ========== 新增：通用概率密度/质量函数实现 ==========
    def pdf(self, x: float) -> float:
        """概率密度函数（通用实现，依赖子类的 _original_pdf）"""
        x_original = self.apply_unshift(x)
        if not self.is_truncated():
            return self._original_pdf(x_original)

        lower, upper = self.get_truncated_bounds()
        # 根据平移模式确定用于CDF的原始边界（用于归一化）
        if self.truncate_type in ['value', 'percentile']:
            lower_p = 0.0 if lower is None else self._original_cdf(lower)
            upper_p = 1.0 if upper is None else self._original_cdf(upper)
            # 检查 x 是否在截断范围内（原始尺度）
            if lower is not None and x_original < lower:
                return 0.0
            if upper is not None and x_original > upper:
                return 0.0
        else:  # value2, percentile2
            if lower is not None:
                lower_orig = lower - self.shift_amount
            else:
                lower_orig = None
            if upper is not None:
                upper_orig = upper - self.shift_amount
            else:
                upper_orig = None
            lower_p = 0.0 if lower_orig is None else self._original_cdf(lower_orig)
            upper_p = 1.0 if upper_orig is None else self._original_cdf(upper_orig)
            # 检查 x 是否在截断范围内（x 是平移后的值）
            if lower is not None and x < lower:
                return 0.0
            if upper is not None and x > upper:
                return 0.0

        # 对于连续分布，_original_pdf 返回 PDF；对于离散分布，返回 PMF，这里统一处理
        return self._original_pdf(x_original) / (upper_p - lower_p)

    def pmf(self, x: float) -> float:
        """概率质量函数（默认与 pdf 相同，离散分布可覆盖）"""
        return self.pdf(x)


# ==================== 正态分布（使用 scipy 精确计算） ====================
class NormalDistribution(DistributionBase):
    """正态分布（精确版）"""

    def __init__(self, params: List[float], markers: Dict[str, Any] = None, func_name: str = None):
        # 先处理参数，再调用父类（父类需要 params 和 markers）
        if len(params) < 2:
            self.mu = 0.0
            self.sigma = 1.0
        else:
            self.mu = float(params[0])
            self.sigma = float(params[1])

        if self.sigma <= 0:
            self.sigma = 1.0

        # 创建 scipy 分布对象（用于后续的分位数等）
        self._dist = sps.norm(loc=self.mu, scale=self.sigma)

        # 调用父类初始化（此时父类会解析截断标记，并设置支持范围）
        super().__init__(params, markers, func_name)

        # 处理百分位数截断：将概率转换为实际值
        if self.truncate_type in ['percentile', 'percentile2']:
            if self.truncate_lower_pct is not None:
                self.truncate_lower = self._original_ppf(self.truncate_lower_pct)
            if self.truncate_upper_pct is not None:
                self.truncate_upper = self._original_ppf(self.truncate_upper_pct)

        # 完成截断设置后调用
        self._finalize_truncation()

        # 缓存截断后的统计量
        self._truncated_stats = None
        self._update_truncated_stats()

    def _update_truncated_stats(self):
        """更新截断后的统计量（使用 scipy 精确计算）"""
        if not self.is_truncated():
            self._truncated_stats = None
            return

        # 获取截断边界（已与支持范围取交集）
        lower, upper = self.get_truncated_bounds()
        if lower is None and upper is None:
            # 交集为空，截断无效，使用原始分布
            self._truncated_stats = None
            return

        # 根据平移模式调整边界到原始尺度
        if self.truncate_type in ['value', 'percentile']:
            # 先截断后平移：边界已经是原始值
            pass
        else:
            # 先平移后截断：需要将边界减去 shift 得到原始值
            if lower is not None:
                lower = lower - self.shift_amount
            if upper is not None:
                upper = upper - self.shift_amount

        lb = -float('inf') if lower is None else lower
        ub = float('inf') if upper is None else upper

        if lb == -float('inf') and ub == float('inf'):
            # 无截断（实际上不会进入，因为 is_truncated 为 True）
            self._truncated_stats = {
                'mean': self.mu,
                'variance': self.sigma ** 2,
                'skewness': 0.0,
                'kurtosis': 3.0
            }
        else:
            # 使用 scipy 的截断正态分布精确计算矩
            a = (lb - self.mu) / self.sigma if lb != -float('inf') else -float('inf')
            b = (ub - self.mu) / self.sigma if ub != float('inf') else float('inf')
            trunc_dist = sps.truncnorm(a, b, loc=self.mu, scale=self.sigma)
            mean_trunc = trunc_dist.mean()
            var_trunc = trunc_dist.var()
            # 偏度和峰度使用 scipy 的 stats 方法，注意峰度返回超额峰度
            skew_trunc = trunc_dist.stats(moments='s')
            kurt_trunc = trunc_dist.stats(moments='k') + 3.0  # 转换为普通峰度

            self._truncated_stats = {
                'mean': mean_trunc,
                'variance': var_trunc,
                'skewness': skew_trunc,
                'kurtosis': kurt_trunc
            }

    def mean(self) -> float:
        if self._truncate_invalid:
            return float('nan')
        if self._truncated_stats:
            return self._truncated_stats['mean'] + self.shift_amount
        return self.mu + self.shift_amount

    def variance(self) -> float:
        if self._truncate_invalid:
            return float('nan')
        if self._truncated_stats:
            return self._truncated_stats['variance']
        return self.sigma ** 2

    def skewness(self) -> float:
        if self._truncate_invalid:
            return float('nan')
        if self._truncated_stats:
            return self._truncated_stats['skewness']
        return 0.0

    def kurtosis(self) -> float:
        if self._truncate_invalid:
            return float('nan')
        if self._truncated_stats:
            return self._truncated_stats['kurtosis']
        return 3.0

    def mode(self) -> float:
        if self._truncate_invalid:
            return float('nan')
        # 对称分布，众数 = 均值
        return self.mean()

    def min_val(self) -> float:
        """
        获取分布的有效最小值。

        返回分布的有效下界，若下界未设置则返回负无穷。

        Returns:
            float: 分布的有效最小值，若下界未定义则返回负无穷。
        """
        lower, upper = self.get_effective_bounds()
        if lower is not None:
            return lower
        return float('-inf')

    def max_val(self) -> float:
        lower, upper = self.get_effective_bounds()
        if upper is not None:
            return upper
        return float('inf')

    def _original_ppf(self, p: float) -> float:
        return self._dist.ppf(p)

    def ppf(self, p: float) -> float:
        if p < 0 or p > 1:
            return float('nan')
        if self._truncate_invalid:
            return float('nan')

        if not self.is_truncated():
            result = self._original_ppf(p)
        else:
            lower, upper = self.get_truncated_bounds()
            if self.truncate_type in ['value', 'percentile']:
                lower_p = 0.0 if lower is None else self._original_cdf(lower)
                upper_p = 1.0 if upper is None else self._original_cdf(upper)
            else:
                if lower is not None:
                    lower_orig = lower - self.shift_amount
                else:
                    lower_orig = None
                if upper is not None:
                    upper_orig = upper - self.shift_amount
                else:
                    upper_orig = None
                lower_p = 0.0 if lower_orig is None else self._original_cdf(lower_orig)
                upper_p = 1.0 if upper_orig is None else self._original_cdf(upper_orig)

            original_p = lower_p + p * (upper_p - lower_p)
            original_p = max(0.0, min(1.0, original_p))
            result = self._original_ppf(original_p)

            if self.truncate_type in ['value', 'percentile']:
                if lower is not None and result < lower:
                    result = lower
                if upper is not None and result > upper:
                    result = upper
            else:
                shifted_result = result + self.shift_amount
                if lower is not None and shifted_result < lower:
                    result = lower - self.shift_amount
                if upper is not None and shifted_result > upper:
                    result = upper - self.shift_amount

        return self.apply_shift(result)

    def _original_cdf(self, x: float) -> float:
        return self._dist.cdf(x)

    def cdf(self, x: float) -> float:
        if self._truncate_invalid:
            return float('nan')
        x_original = self.apply_unshift(x)

        if not self.is_truncated():
            return self._original_cdf(x_original)
        else:
            lower, upper = self.get_truncated_bounds()
            if self.truncate_type in ['value', 'percentile']:
                lower_p = 0.0 if lower is None else self._original_cdf(lower)
                upper_p = 1.0 if upper is None else self._original_cdf(upper)
            else:
                if lower is not None:
                    lower_orig = lower - self.shift_amount
                else:
                    lower_orig = None
                if upper is not None:
                    upper_orig = upper - self.shift_amount
                else:
                    upper_orig = None
                lower_p = 0.0 if lower_orig is None else self._original_cdf(lower_orig)
                upper_p = 1.0 if upper_orig is None else self._original_cdf(upper_orig)

            # 检查 x 是否在截断范围内
            if self.truncate_type in ['value', 'percentile']:
                if lower is not None and x_original < lower:
                    return 0.0
                if upper is not None and x_original > upper:
                    return 1.0
            else:
                shifted_x = x
                if lower is not None and shifted_x < lower:
                    return 0.0
                if upper is not None and shifted_x > upper:
                    return 1.0

            orig_cdf = self._original_cdf(x_original)
            return (orig_cdf - lower_p) / (upper_p - lower_p)

    def _original_pdf(self, x: float) -> float:
        return self._dist.pdf(x)

    def pdf(self, x: float) -> float:
        if self._truncate_invalid:
            return float('nan')
        x_original = self.apply_unshift(x)

        if not self.is_truncated():
            return self._original_pdf(x_original)
        else:
            lower, upper = self.get_truncated_bounds()
            if self.truncate_type in ['value', 'percentile']:
                lower_p = 0.0 if lower is None else self._original_cdf(lower)
                upper_p = 1.0 if upper is None else self._original_cdf(upper)
            else:
                if lower is not None:
                    lower_orig = lower - self.shift_amount
                else:
                    lower_orig = None
                if upper is not None:
                    upper_orig = upper - self.shift_amount
                else:
                    upper_orig = None
                lower_p = 0.0 if lower_orig is None else self._original_cdf(lower_orig)
                upper_p = 1.0 if upper_orig is None else self._original_cdf(upper_orig)

            # 检查 x 是否在截断范围内
            if self.truncate_type in ['value', 'percentile']:
                if lower is not None and x_original < lower:
                    return 0.0
                if upper is not None and x_original > upper:
                    return 0.0
            else:
                if lower is not None and x < lower:
                    return 0.0
                if upper is not None and x > upper:
                    return 0.0

            return self._original_pdf(x_original) / (upper_p - lower_p)


# ==================== 均匀分布（使用 scipy，保持精确） ====================
class UniformDistribution(DistributionBase):
    """均匀分布（精确版）"""

    def __init__(self, params: List[float], markers: Dict[str, Any] = None, func_name: str = None):
        if len(params) < 2:
            self.a = 0.0
            self.b = 1.0
        else:
            self.a = float(params[0])
            self.b = float(params[1])

        if self.b <= self.a:
            self.b = self.a + 1.0

        self._dist = sps.uniform(loc=self.a, scale=self.b - self.a)

        super().__init__(params, markers, func_name)

        if self.truncate_type in ['percentile', 'percentile2']:
            if self.truncate_lower_pct is not None:
                self.truncate_lower = self._original_ppf(self.truncate_lower_pct)
            if self.truncate_upper_pct is not None:
                self.truncate_upper = self._original_ppf(self.truncate_upper_pct)

        self._finalize_truncation()
        self._update_effective_bounds()
        self._truncated_stats = None
        self._compute_truncated_stats()

    def _update_effective_bounds(self):
        """更新有效边界（用于快速计算）"""
        lower, upper = self.get_truncated_bounds()
        if self.truncate_type in ['value', 'percentile']:
            # 先截断后平移：边界是原始值
            pass
        else:
            # 先平移后截断：需要将边界转换为原始值
            if lower is not None:
                lower = lower - self.shift_amount
            if upper is not None:
                upper = upper - self.shift_amount

        if lower is None:
            self._a_effective = self.a
        else:
            self._a_effective = max(self.a, lower)

        if upper is None:
            self._b_effective = self.b
        else:
            self._b_effective = min(self.b, upper)

        if self._a_effective >= self._b_effective:
            self._a_effective = self.a
            self._b_effective = self.b

    def _compute_truncated_stats(self):
        """计算截断后的矩（均匀分布有闭式解）"""
        if not self.is_truncated():
            return

        # 获取有效边界（原始尺度）
        lower, upper = self.get_truncated_bounds()
        if lower is None and upper is None:
            # 交集为空，截断无效
            self._truncated_stats = None
            return

        if self.truncate_type in ['value', 'percentile']:
            L = lower if lower is not None else self.a
            R = upper if upper is not None else self.b
        else:
            # 对于 value2/percentile2，边界是平移后的，需要转换到原始尺度
            if lower is not None:
                L = lower - self.shift_amount
            else:
                L = self.a
            if upper is not None:
                R = upper - self.shift_amount
            else:
                R = self.b
            L = max(L, self.a)
            R = min(R, self.b)

        width = R - L

        if width <= 0:
            # 退化为单点
            self._truncated_stats = {
                'mean': L,
                'variance': 0.0,
                'skewness': 0.0,
                'kurtosis': 3.0
            }
        else:
            mean_trunc = (L + R) / 2
            var_trunc = width**2 / 12
            skew_trunc = 0.0
            kurt_trunc = 1.8

            self._truncated_stats = {
                'mean': mean_trunc,
                'variance': var_trunc,
                'skewness': skew_trunc,
                'kurtosis': kurt_trunc
            }

    def mean(self) -> float:
        if self._truncate_invalid:
            return float('nan')
        if self._truncated_stats:
            return self._truncated_stats['mean'] + self.shift_amount
        return (self.a + self.b) / 2 + self.shift_amount

    def variance(self) -> float:
        if self._truncate_invalid:
            return float('nan')
        if self._truncated_stats:
            return self._truncated_stats['variance']
        return (self.b - self.a) ** 2 / 12

    def skewness(self) -> float:
        if self._truncate_invalid:
            return float('nan')
        if self._truncated_stats:
            return self._truncated_stats['skewness']
        return 0.0

    def kurtosis(self) -> float:
        if self._truncate_invalid:
            return float('nan')
        if self._truncated_stats:
            return self._truncated_stats['kurtosis']
        return 1.8

    def mode(self) -> float:
        if self._truncate_invalid:
            return float('nan')
        return self.mean()

    def min_val(self) -> float:
        lower, upper = self.get_effective_bounds()
        if lower is not None:
            return lower
        return self.apply_shift(self.a)

    def max_val(self) -> float:
        lower, upper = self.get_effective_bounds()
        if upper is not None:
            return upper
        return self.apply_shift(self.b)

    def _original_ppf(self, p: float) -> float:
        return self._dist.ppf(p)

    def ppf(self, p: float) -> float:
        if p < 0 or p > 1:
            return float('nan')
        if self._truncate_invalid:
            return float('nan')

        if not self.is_truncated():
            result = self._original_ppf(p)
        else:
            if p <= 0:
                return self.min_val()
            if p >= 1:
                return self.max_val()
            # 使用有效边界直接计算
            lower_eff, upper_eff = self.get_effective_bounds()
            if lower_eff is None or upper_eff is None:
                # 单边截断，无法简单线性映射，回退到通用方法
                # 获取原始尺度上的有效边界
                lower_raw, upper_raw = self.get_truncated_bounds()
                if self.truncate_type in ['value', 'percentile']:
                    L = lower_raw if lower_raw is not None else self.a
                    R = upper_raw if upper_raw is not None else self.b
                else:
                    L = (lower_raw - self.shift_amount) if lower_raw is not None else self.a
                    R = (upper_raw - self.shift_amount) if upper_raw is not None else self.b
                    L = max(L, self.a)
                    R = min(R, self.b)
                # 线性映射到截断区间
                result = L + p * (R - L)
            else:
                # 双边有效边界，直接线性映射
                result = lower_eff + p * (upper_eff - lower_eff)
        return self.apply_shift(result) if not self.truncate_type in ['value2', 'percentile2'] else result

    def _original_cdf(self, x: float) -> float:
        return self._dist.cdf(x)

    def cdf(self, x: float) -> float:
        if self._truncate_invalid:
            return float('nan')
        x_original = self.apply_unshift(x)
        lower_eff, upper_eff = self.get_effective_bounds()
        if lower_eff is not None and x < lower_eff:
            return 0.0
        if upper_eff is not None and x > upper_eff:
            return 1.0

        # 将 x 映射到原始尺度，然后计算线性比例
        if self.truncate_type in ['value', 'percentile']:
            # 先截断后平移：x_original 是原始值，有效边界也是原始值
            L = lower_eff - self.shift_amount if lower_eff is not None else self.a
            R = upper_eff - self.shift_amount if upper_eff is not None else self.b
        else:
            # 先平移后截断：有效边界是平移后的值，x 是平移后的值，需要映射回原始尺度比较
            L = lower_eff - self.shift_amount if lower_eff is not None else self.a
            R = upper_eff - self.shift_amount if upper_eff is not None else self.b
            # 此时 x_original 已经正确

        return (x_original - L) / (R - L)

    def _original_pdf(self, x: float) -> float:
        return self._dist.pdf(x)

    def pdf(self, x: float) -> float:
        if self._truncate_invalid:
            return float('nan')
        lower_eff, upper_eff = self.get_effective_bounds()
        if lower_eff is not None and x < lower_eff:
            return 0.0
        if upper_eff is not None and x > upper_eff:
            return 0.0
        width = (upper_eff - lower_eff) if (lower_eff is not None and upper_eff is not None) else (self.b - self.a)
        return 1.0 / width


# ==================== 伽马分布（使用 scipy） ====================
class GammaDistribution(DistributionBase):
    """伽马分布（精确版）"""

    def __init__(self, params: List[float], markers: Dict[str, Any] = None, func_name: str = None):
        if len(params) < 2:
            self.shape = 1.0
            self.scale = 1.0
        else:
            self.shape = float(params[0])
            self.scale = float(params[1])

        if self.shape <= 0:
            self.shape = 1.0
        if self.scale <= 0:
            self.scale = 1.0

        self._dist = sps.gamma(a=self.shape, scale=self.scale)

        super().__init__(params, markers, func_name)

        if self.truncate_type in ['percentile', 'percentile2']:
            if self.truncate_lower_pct is not None:
                self.truncate_lower = self._original_ppf(self.truncate_lower_pct)
            if self.truncate_upper_pct is not None:
                self.truncate_upper = self._original_ppf(self.truncate_upper_pct)

        self._finalize_truncation()
        self._truncated_stats = None
        self._update_truncated_stats()

    def _update_truncated_stats(self):
        """计算截断后的矩（使用 scipy.expect）"""
        if not self.is_truncated():
            return

        lower, upper = self.get_truncated_bounds()
        if lower is None and upper is None:
            self._truncated_stats = None
            return

        if self.truncate_type in ['value', 'percentile']:
            lb = -float('inf') if lower is None else lower
            ub = float('inf') if upper is None else upper
        else:
            # value2/percentile2: 边界是平移后的值，需转换为原始值
            lb = -float('inf') if lower is None else (lower - self.shift_amount)
            ub = float('inf') if upper is None else (upper - self.shift_amount)

        if lb == -float('inf') and ub == float('inf'):
            moments = self._dist.stats(moments='mvsk')
            mean_val = moments[0]
            var_val = moments[1]
            skew_val = moments[2]
            kurt_val = moments[3] + 3.0
        else:
            try:
                # 使用 expect 计算条件期望
                m1 = self._dist.expect(lambda x: x, lb=lb, ub=ub, conditional=True)
                m2 = self._dist.expect(lambda x: x**2, lb=lb, ub=ub, conditional=True)
                m3 = self._dist.expect(lambda x: x**3, lb=lb, ub=ub, conditional=True)
                m4 = self._dist.expect(lambda x: x**4, lb=lb, ub=ub, conditional=True)

                mean_val = m1
                var_val = m2 - m1**2
                if var_val > 0:
                    skew_val = (m3 - 3*m1*var_val - m1**3) / (var_val**1.5)
                    kurt_val = (m4 - 4*m1*m3 + 6*m1**2*m2 - 3*m1**4) / (var_val**2)
                else:
                    skew_val = 0.0
                    kurt_val = 3.0
            except Exception as e:
                moments = self._dist.stats(moments='mvsk')
                mean_val = moments[0]
                var_val = moments[1]
                skew_val = moments[2]
                kurt_val = moments[3] + 3.0

        self._truncated_stats = {
            'mean': mean_val,
            'variance': var_val,
            'skewness': skew_val,
            'kurtosis': kurt_val
        }

    def mean(self) -> float:
        if self._truncate_invalid:
            return float('nan')
        if self._truncated_stats:
            return self._truncated_stats['mean'] + self.shift_amount
        return self._dist.mean() + self.shift_amount

    def variance(self) -> float:
        if self._truncate_invalid:
            return float('nan')
        if self._truncated_stats:
            return self._truncated_stats['variance']
        return self._dist.var()

    def skewness(self) -> float:
        if self._truncate_invalid:
            return float('nan')
        if self._truncated_stats:
            return self._truncated_stats['skewness']
        return self._dist.stats(moments='s')

    def kurtosis(self) -> float:
        if self._truncate_invalid:
            return float('nan')
        if self._truncated_stats:
            return self._truncated_stats['kurtosis']
        return self._dist.stats(moments='k') + 3.0

    def mode(self) -> float:
        if self._truncate_invalid:
            return float('nan')
        if self.shape >= 1:
            mode_val = (self.shape - 1) * self.scale
        else:
            mode_val = 0.0
        return mode_val + self.shift_amount

    def min_val(self) -> float:
        lower, upper = self.get_effective_bounds()
        if lower is not None:
            return lower
        return self.apply_shift(0.0)

    def max_val(self) -> float:
        lower, upper = self.get_effective_bounds()
        if upper is not None:
            return upper
        return float('inf')

    def _original_ppf(self, p: float) -> float:
        return self._dist.ppf(p)

    def _original_cdf(self, x: float) -> float:
        return self._dist.cdf(x)

    def cdf(self, x: float) -> float:
        if self._truncate_invalid:
            return float('nan')
        x_original = self.apply_unshift(x)

        if not self.is_truncated():
            return self._original_cdf(x_original)
        else:
            lower, upper = self.get_truncated_bounds()
            if self.truncate_type in ['value', 'percentile']:
                lower_p = 0.0 if lower is None else self._original_cdf(lower)
                upper_p = 1.0 if upper is None else self._original_cdf(upper)
            else:
                if lower is not None:
                    lower_orig = lower - self.shift_amount
                else:
                    lower_orig = None
                if upper is not None:
                    upper_orig = upper - self.shift_amount
                else:
                    upper_orig = None
                lower_p = 0.0 if lower_orig is None else self._original_cdf(lower_orig)
                upper_p = 1.0 if upper_orig is None else self._original_cdf(upper_orig)

            # 检查 x 是否在截断范围内
            if self.truncate_type in ['value', 'percentile']:
                if lower is not None and x_original < lower:
                    return 0.0
                if upper is not None and x_original > upper:
                    return 1.0
            else:
                if lower is not None and x < lower:
                    return 0.0
                if upper is not None and x > upper:
                    return 1.0

            orig_cdf = self._original_cdf(x_original)
            return (orig_cdf - lower_p) / (upper_p - lower_p)

    def _original_pdf(self, x: float) -> float:
        return self._dist.pdf(x)

    def pdf(self, x: float) -> float:
        if self._truncate_invalid:
            return float('nan')
        x_original = self.apply_unshift(x)

        if not self.is_truncated():
            return self._original_pdf(x_original)
        else:
            lower, upper = self.get_truncated_bounds()
            if self.truncate_type in ['value', 'percentile']:
                lower_p = 0.0 if lower is None else self._original_cdf(lower)
                upper_p = 1.0 if upper is None else self._original_cdf(upper)
            else:
                if lower is not None:
                    lower_orig = lower - self.shift_amount
                else:
                    lower_orig = None
                if upper is not None:
                    upper_orig = upper - self.shift_amount
                else:
                    upper_orig = None
                lower_p = 0.0 if lower_orig is None else self._original_cdf(lower_orig)
                upper_p = 1.0 if upper_orig is None else self._original_cdf(upper_orig)

            # 检查 x 是否在截断范围内
            if self.truncate_type in ['value', 'percentile']:
                if lower is not None and x_original < lower:
                    return 0.0
                if upper is not None and x_original > upper:
                    return 0.0
            else:
                if lower is not None and x < lower:
                    return 0.0
                if upper is not None and x > upper:
                    return 0.0

            return self._original_pdf(x_original) / (upper_p - lower_p)


# ==================== 贝塔分布（使用 scipy） ====================
class BetaDistribution(DistributionBase):
    """贝塔分布（精确版）"""

    def __init__(self, params: List[float], markers: Dict[str, Any] = None, func_name: str = None):
        if len(params) < 2:
            self.alpha = 1.0
            self.beta = 1.0
        else:
            self.alpha = float(params[0])
            self.beta = float(params[1])

        if self.alpha <= 0:
            self.alpha = 1.0
        if self.beta <= 0:
            self.beta = 1.0

        self._dist = sps.beta(a=self.alpha, b=self.beta)

        super().__init__(params, markers, func_name)

        if self.truncate_type in ['percentile', 'percentile2']:
            if self.truncate_lower_pct is not None:
                self.truncate_lower = self._original_ppf(self.truncate_lower_pct)
            if self.truncate_upper_pct is not None:
                self.truncate_upper = self._original_ppf(self.truncate_upper_pct)

        self._finalize_truncation()
        self._truncated_stats = None
        self._update_truncated_stats()

    def _update_truncated_stats(self):
        """计算截断后的矩"""
        if not self.is_truncated():
            return

        lower, upper = self.get_truncated_bounds()
        if lower is None and upper is None:
            self._truncated_stats = None
            return

        if self.truncate_type in ['value', 'percentile']:
            lb = -float('inf') if lower is None else lower
            ub = float('inf') if upper is None else upper
        else:
            lb = -float('inf') if lower is None else (lower - self.shift_amount)
            ub = float('inf') if upper is None else (upper - self.shift_amount)

        if lb == -float('inf') and ub == float('inf'):
            moments = self._dist.stats(moments='mvsk')
            mean_val = moments[0]
            var_val = moments[1]
            skew_val = moments[2]
            kurt_val = moments[3] + 3.0
        else:
            try:
                m1 = self._dist.expect(lambda x: x, lb=lb, ub=ub, conditional=True)
                m2 = self._dist.expect(lambda x: x**2, lb=lb, ub=ub, conditional=True)
                m3 = self._dist.expect(lambda x: x**3, lb=lb, ub=ub, conditional=True)
                m4 = self._dist.expect(lambda x: x**4, lb=lb, ub=ub, conditional=True)

                mean_val = m1
                var_val = m2 - m1**2
                if var_val > 0:
                    skew_val = (m3 - 3*m1*var_val - m1**3) / (var_val**1.5)
                    kurt_val = (m4 - 4*m1*m3 + 6*m1**2*m2 - 3*m1**4) / (var_val**2)
                else:
                    skew_val = 0.0
                    kurt_val = 3.0
            except Exception as e:
                moments = self._dist.stats(moments='mvsk')
                mean_val = moments[0]
                var_val = moments[1]
                skew_val = moments[2]
                kurt_val = moments[3] + 3.0

        self._truncated_stats = {
            'mean': mean_val,
            'variance': var_val,
            'skewness': skew_val,
            'kurtosis': kurt_val
        }

    def mean(self) -> float:
        if self._truncate_invalid:
            return float('nan')
        if self._truncated_stats:
            return self._truncated_stats['mean'] + self.shift_amount
        return self._dist.mean() + self.shift_amount

    def variance(self) -> float:
        if self._truncate_invalid:
            return float('nan')
        if self._truncated_stats:
            return self._truncated_stats['variance']
        return self._dist.var()

    def skewness(self) -> float:
        if self._truncate_invalid:
            return float('nan')
        if self._truncated_stats:
            return self._truncated_stats['skewness']
        return self._dist.stats(moments='s')

    def kurtosis(self) -> float:
        if self._truncate_invalid:
            return float('nan')
        if self._truncated_stats:
            return self._truncated_stats['kurtosis']
        return self._dist.stats(moments='k') + 3.0

    def mode(self) -> float:
        if self._truncate_invalid:
            return float('nan')
        if self.alpha > 1 and self.beta > 1:
            mode_val = (self.alpha - 1) / (self.alpha + self.beta - 2)
        elif self.alpha < 1 and self.beta < 1:
            mode_val = 0.5
        elif self.alpha <= 1 and self.beta > 1:
            mode_val = 0.0
        elif self.alpha > 1 and self.beta <= 1:
            mode_val = 1.0
        else:
            mode_val = self.mean() - self.shift_amount
        return mode_val + self.shift_amount

    def min_val(self) -> float:
        lower, upper = self.get_effective_bounds()
        if lower is not None:
            return lower
        return self.apply_shift(0.0)

    def max_val(self) -> float:
        lower, upper = self.get_effective_bounds()
        if upper is not None:
            return upper
        return self.apply_shift(1.0)

    def _original_ppf(self, p: float) -> float:
        return self._dist.ppf(p)

    def _original_cdf(self, x: float) -> float:
        return self._dist.cdf(x)

    def cdf(self, x: float) -> float:
        if self._truncate_invalid:
            return float('nan')
        x_original = self.apply_unshift(x)

        if not self.is_truncated():
            return self._original_cdf(x_original)
        else:
            lower, upper = self.get_truncated_bounds()
            if self.truncate_type in ['value', 'percentile']:
                lower_p = 0.0 if lower is None else self._original_cdf(lower)
                upper_p = 1.0 if upper is None else self._original_cdf(upper)
            else:
                if lower is not None:
                    lower_orig = lower - self.shift_amount
                else:
                    lower_orig = None
                if upper is not None:
                    upper_orig = upper - self.shift_amount
                else:
                    upper_orig = None
                lower_p = 0.0 if lower_orig is None else self._original_cdf(lower_orig)
                upper_p = 1.0 if upper_orig is None else self._original_cdf(upper_orig)

            if self.truncate_type in ['value', 'percentile']:
                if lower is not None and x_original < lower:
                    return 0.0
                if upper is not None and x_original > upper:
                    return 1.0
            else:
                if lower is not None and x < lower:
                    return 0.0
                if upper is not None and x > upper:
                    return 1.0

            orig_cdf = self._original_cdf(x_original)
            return (orig_cdf - lower_p) / (upper_p - lower_p)

    def _original_pdf(self, x: float) -> float:
        return self._dist.pdf(x)

    def pdf(self, x: float) -> float:
        if self._truncate_invalid:
            return float('nan')
        x_original = self.apply_unshift(x)

        if not self.is_truncated():
            return self._original_pdf(x_original)
        else:
            lower, upper = self.get_truncated_bounds()
            if self.truncate_type in ['value', 'percentile']:
                lower_p = 0.0 if lower is None else self._original_cdf(lower)
                upper_p = 1.0 if upper is None else self._original_cdf(upper)
            else:
                if lower is not None:
                    lower_orig = lower - self.shift_amount
                else:
                    lower_orig = None
                if upper is not None:
                    upper_orig = upper - self.shift_amount
                else:
                    upper_orig = None
                lower_p = 0.0 if lower_orig is None else self._original_cdf(lower_orig)
                upper_p = 1.0 if upper_orig is None else self._original_cdf(upper_orig)

            if self.truncate_type in ['value', 'percentile']:
                if lower is not None and x_original < lower:
                    return 0.0
                if upper is not None and x_original > upper:
                    return 0.0
            else:
                if lower is not None and x < lower:
                    return 0.0
                if upper is not None and x > upper:
                    return 0.0

            return self._original_pdf(x_original) / (upper_p - lower_p)


# ==================== 泊松分布（离散，使用 scipy） ====================
class PoissonDistribution(DistributionBase):
    """泊松分布（精确版，离散）"""

    def __init__(self, params: List[float], markers: Dict[str, Any] = None, func_name: str = None):
        if len(params) < 1:
            self.mu = 1.0
        else:
            self.mu = float(params[0])

        if self.mu <= 0:
            self.mu = 1.0

        self._dist = sps.poisson(mu=self.mu)

        super().__init__(params, markers, func_name)

        if self.truncate_type in ['percentile', 'percentile2']:
            if self.truncate_lower_pct is not None:
                self.truncate_lower = self._original_ppf(self.truncate_lower_pct)
            if self.truncate_upper_pct is not None:
                self.truncate_upper = self._original_ppf(self.truncate_upper_pct)

        self._finalize_truncation()
        self._truncated_stats = None
        self._update_truncated_stats()

    def _update_truncated_stats(self):
        """计算截断后的矩（使用 scipy.expect 求和）"""
        if not self.is_truncated():
            return

        lower, upper = self.get_truncated_bounds()
        if lower is None and upper is None:
            self._truncated_stats = None
            return

        if self.truncate_type in ['value', 'percentile']:
            lb = -float('inf') if lower is None else lower
            ub = float('inf') if upper is None else upper
        else:
            lb = -float('inf') if lower is None else (lower - self.shift_amount)
            ub = float('inf') if upper is None else (upper - self.shift_amount)

        if lb == -float('inf') and ub == float('inf'):
            moments = self._dist.stats(moments='mvsk')
            mean_val = moments[0]
            var_val = moments[1]
            skew_val = moments[2]
            kurt_val = moments[3] + 3.0
        else:
            try:
                # 对于泊松，expect 自动处理为求和
                m1 = self._dist.expect(lambda x: x, lb=lb, ub=ub, conditional=True)
                m2 = self._dist.expect(lambda x: x**2, lb=lb, ub=ub, conditional=True)
                m3 = self._dist.expect(lambda x: x**3, lb=lb, ub=ub, conditional=True)
                m4 = self._dist.expect(lambda x: x**4, lb=lb, ub=ub, conditional=True)

                mean_val = m1
                var_val = m2 - m1**2
                if var_val > 0:
                    skew_val = (m3 - 3*m1*var_val - m1**3) / (var_val**1.5)
                    kurt_val = (m4 - 4*m1*m3 + 6*m1**2*m2 - 3*m1**4) / (var_val**2)
                else:
                    skew_val = 0.0
                    kurt_val = 3.0
            except Exception as e:
                moments = self._dist.stats(moments='mvsk')
                mean_val = moments[0]
                var_val = moments[1]
                skew_val = moments[2]
                kurt_val = moments[3] + 3.0

        self._truncated_stats = {
            'mean': mean_val,
            'variance': var_val,
            'skewness': skew_val,
            'kurtosis': kurt_val
        }

    def mean(self) -> float:
        if self._truncate_invalid:
            return float('nan')
        if self._truncated_stats:
            return self._truncated_stats['mean'] + self.shift_amount
        return self._dist.mean() + self.shift_amount

    def variance(self) -> float:
        if self._truncate_invalid:
            return float('nan')
        if self._truncated_stats:
            return self._truncated_stats['variance']
        return self._dist.var()

    def skewness(self) -> float:
        if self._truncate_invalid:
            return float('nan')
        if self._truncated_stats:
            return self._truncated_stats['skewness']
        return self._dist.stats(moments='s')

    def kurtosis(self) -> float:
        if self._truncate_invalid:
            return float('nan')
        if self._truncated_stats:
            return self._truncated_stats['kurtosis']
        return self._dist.stats(moments='k') + 3.0

    def mode(self) -> float:
        if self._truncate_invalid:
            return float('nan')
        mode_val = math.floor(self.mu)
        return mode_val + self.shift_amount

    def min_val(self) -> float:
        lower, upper = self.get_effective_bounds()
        if lower is not None:
            return lower
        return self.apply_shift(0.0)

    def max_val(self) -> float:
        lower, upper = self.get_effective_bounds()
        if upper is not None:
            return upper
        return float('inf')

    def _original_ppf(self, p: float) -> float:
        return self._dist.ppf(p)

    def _original_cdf(self, x: float) -> float:
        return self._dist.cdf(x)

    def cdf(self, x: float) -> float:
        if self._truncate_invalid:
            return float('nan')
        x_original = self.apply_unshift(x)

        if not self.is_truncated():
            return self._original_cdf(x_original)
        else:
            lower, upper = self.get_truncated_bounds()
            if self.truncate_type in ['value', 'percentile']:
                lower_p = 0.0 if lower is None else self._original_cdf(lower)
                upper_p = 1.0 if upper is None else self._original_cdf(upper)
            else:
                if lower is not None:
                    lower_orig = lower - self.shift_amount
                else:
                    lower_orig = None
                if upper is not None:
                    upper_orig = upper - self.shift_amount
                else:
                    upper_orig = None
                lower_p = 0.0 if lower_orig is None else self._original_cdf(lower_orig)
                upper_p = 1.0 if upper_orig is None else self._original_cdf(upper_orig)

            if self.truncate_type in ['value', 'percentile']:
                if lower is not None and x_original < lower:
                    return 0.0
                if upper is not None and x_original > upper:
                    return 1.0
            else:
                if lower is not None and x < lower:
                    return 0.0
                if upper is not None and x > upper:
                    return 1.0

            orig_cdf = self._original_cdf(x_original)
            return (orig_cdf - lower_p) / (upper_p - lower_p)

    def _original_pdf(self, x: float) -> float:
        return self._dist.pmf(x)

    def pdf(self, x: float) -> float:
        # 对于离散分布，返回 pmf
        return self.pmf(x)

    def pmf(self, x: float) -> float:
        if self._truncate_invalid:
            return float('nan')
        x_original = self.apply_unshift(x)

        if not self.is_truncated():
            return self._original_pdf(x_original)
        else:
            lower, upper = self.get_truncated_bounds()
            if self.truncate_type in ['value', 'percentile']:
                lower_p = 0.0 if lower is None else self._original_cdf(lower)
                upper_p = 1.0 if upper is None else self._original_cdf(upper)
            else:
                if lower is not None:
                    lower_orig = lower - self.shift_amount
                else:
                    lower_orig = None
                if upper is not None:
                    upper_orig = upper - self.shift_amount
                else:
                    upper_orig = None
                lower_p = 0.0 if lower_orig is None else self._original_cdf(lower_orig)
                upper_p = 1.0 if upper_orig is None else self._original_cdf(upper_orig)

            if self.truncate_type in ['value', 'percentile']:
                if lower is not None and x_original < lower:
                    return 0.0
                if upper is not None and x_original > upper:
                    return 0.0
            else:
                if lower is not None and x < lower:
                    return 0.0
                if upper is not None and x > upper:
                    return 0.0

            return self._original_pdf(x_original) / (upper_p - lower_p)


# ==================== 卡方分布（使用 scipy） ====================
class ChiSquaredDistribution(DistributionBase):
    """卡方分布（精确版）"""

    def __init__(self, params: List[float], markers: Dict[str, Any] = None, func_name: str = None):
        if len(params) < 1:
            self.df = 1.0
        else:
            self.df = float(params[0])

        if self.df <= 0:
            self.df = 1.0

        self._dist = sps.chi2(df=self.df)

        super().__init__(params, markers, func_name)

        if self.truncate_type in ['percentile', 'percentile2']:
            if self.truncate_lower_pct is not None:
                self.truncate_lower = self._original_ppf(self.truncate_lower_pct)
            if self.truncate_upper_pct is not None:
                self.truncate_upper = self._original_ppf(self.truncate_upper_pct)

        self._finalize_truncation()
        self._truncated_stats = None
        self._update_truncated_stats()

    def _update_truncated_stats(self):
        if not self.is_truncated():
            return

        lower, upper = self.get_truncated_bounds()
        if lower is None and upper is None:
            self._truncated_stats = None
            return

        if self.truncate_type in ['value', 'percentile']:
            lb = -float('inf') if lower is None else lower
            ub = float('inf') if upper is None else upper
        else:
            lb = -float('inf') if lower is None else (lower - self.shift_amount)
            ub = float('inf') if upper is None else (upper - self.shift_amount)

        if lb == -float('inf') and ub == float('inf'):
            moments = self._dist.stats(moments='mvsk')
            mean_val = moments[0]
            var_val = moments[1]
            skew_val = moments[2]
            kurt_val = moments[3] + 3.0
        else:
            try:
                m1 = self._dist.expect(lambda x: x, lb=lb, ub=ub, conditional=True)
                m2 = self._dist.expect(lambda x: x**2, lb=lb, ub=ub, conditional=True)
                m3 = self._dist.expect(lambda x: x**3, lb=lb, ub=ub, conditional=True)
                m4 = self._dist.expect(lambda x: x**4, lb=lb, ub=ub, conditional=True)

                mean_val = m1
                var_val = m2 - m1**2
                if var_val > 0:
                    skew_val = (m3 - 3*m1*var_val - m1**3) / (var_val**1.5)
                    kurt_val = (m4 - 4*m1*m3 + 6*m1**2*m2 - 3*m1**4) / (var_val**2)
                else:
                    skew_val = 0.0
                    kurt_val = 3.0
            except Exception as e:
                moments = self._dist.stats(moments='mvsk')
                mean_val = moments[0]
                var_val = moments[1]
                skew_val = moments[2]
                kurt_val = moments[3] + 3.0

        self._truncated_stats = {
            'mean': mean_val,
            'variance': var_val,
            'skewness': skew_val,
            'kurtosis': kurt_val
        }

    def mean(self) -> float:
        if self._truncate_invalid:
            return float('nan')
        if self._truncated_stats:
            return self._truncated_stats['mean'] + self.shift_amount
        return self._dist.mean() + self.shift_amount

    def variance(self) -> float:
        if self._truncate_invalid:
            return float('nan')
        if self._truncated_stats:
            return self._truncated_stats['variance']
        return self._dist.var()

    def skewness(self) -> float:
        if self._truncate_invalid:
            return float('nan')
        if self._truncated_stats:
            return self._truncated_stats['skewness']
        return self._dist.stats(moments='s')

    def kurtosis(self) -> float:
        if self._truncate_invalid:
            return float('nan')
        if self._truncated_stats:
            return self._truncated_stats['kurtosis']
        return self._dist.stats(moments='k') + 3.0

    def mode(self) -> float:
        if self._truncate_invalid:
            return float('nan')
        if self.df >= 2:
            mode_val = self.df - 2
        else:
            mode_val = 0.0
        return mode_val + self.shift_amount

    def min_val(self) -> float:
        lower, upper = self.get_effective_bounds()
        if lower is not None:
            return lower
        return self.apply_shift(0.0)

    def max_val(self) -> float:
        lower, upper = self.get_effective_bounds()
        if upper is not None:
            return upper
        return float('inf')

    def _original_ppf(self, p: float) -> float:
        return self._dist.ppf(p)

    def _original_cdf(self, x: float) -> float:
        return self._dist.cdf(x)

    def cdf(self, x: float) -> float:
        if self._truncate_invalid:
            return float('nan')
        x_original = self.apply_unshift(x)

        if not self.is_truncated():
            return self._original_cdf(x_original)
        else:
            lower, upper = self.get_truncated_bounds()
            if self.truncate_type in ['value', 'percentile']:
                lower_p = 0.0 if lower is None else self._original_cdf(lower)
                upper_p = 1.0 if upper is None else self._original_cdf(upper)
            else:
                if lower is not None:
                    lower_orig = lower - self.shift_amount
                else:
                    lower_orig = None
                if upper is not None:
                    upper_orig = upper - self.shift_amount
                else:
                    upper_orig = None
                lower_p = 0.0 if lower_orig is None else self._original_cdf(lower_orig)
                upper_p = 1.0 if upper_orig is None else self._original_cdf(upper_orig)

            if self.truncate_type in ['value', 'percentile']:
                if lower is not None and x_original < lower:
                    return 0.0
                if upper is not None and x_original > upper:
                    return 1.0
            else:
                if lower is not None and x < lower:
                    return 0.0
                if upper is not None and x > upper:
                    return 1.0

            orig_cdf = self._original_cdf(x_original)
            return (orig_cdf - lower_p) / (upper_p - lower_p)

    def _original_pdf(self, x: float) -> float:
        return self._dist.pdf(x)

    def pdf(self, x: float) -> float:
        if self._truncate_invalid:
            return float('nan')
        x_original = self.apply_unshift(x)

        if not self.is_truncated():
            return self._original_pdf(x_original)
        else:
            lower, upper = self.get_truncated_bounds()
            if self.truncate_type in ['value', 'percentile']:
                lower_p = 0.0 if lower is None else self._original_cdf(lower)
                upper_p = 1.0 if upper is None else self._original_cdf(upper)
            else:
                if lower is not None:
                    lower_orig = lower - self.shift_amount
                else:
                    lower_orig = None
                if upper is not None:
                    upper_orig = upper - self.shift_amount
                else:
                    upper_orig = None
                lower_p = 0.0 if lower_orig is None else self._original_cdf(lower_orig)
                upper_p = 1.0 if upper_orig is None else self._original_cdf(upper_orig)

            if self.truncate_type in ['value', 'percentile']:
                if lower is not None and x_original < lower:
                    return 0.0
                if upper is not None and x_original > upper:
                    return 0.0
            else:
                if lower is not None and x < lower:
                    return 0.0
                if upper is not None and x > upper:
                    return 0.0

            return self._original_pdf(x_original) / (upper_p - lower_p)


# ==================== F 分布（使用 scipy） ====================
class FDistribution(DistributionBase):
    """F 分布（精确版）"""

    def __init__(self, params: List[float], markers: Dict[str, Any] = None, func_name: str = None):
        if len(params) < 2:
            self.dfn = 1.0
            self.dfd = 1.0
        else:
            self.dfn = float(params[0])
            self.dfd = float(params[1])

        if self.dfn <= 0:
            self.dfn = 1.0
        if self.dfd <= 0:
            self.dfd = 1.0

        self._dist = sps.f(dfn=self.dfn, dfd=self.dfd)

        super().__init__(params, markers, func_name)

        if self.truncate_type in ['percentile', 'percentile2']:
            if self.truncate_lower_pct is not None:
                self.truncate_lower = self._original_ppf(self.truncate_lower_pct)
            if self.truncate_upper_pct is not None:
                self.truncate_upper = self._original_ppf(self.truncate_upper_pct)

        self._finalize_truncation()
        self._truncated_stats = None
        self._update_truncated_stats()

    def _update_truncated_stats(self):
        if not self.is_truncated():
            return

        lower, upper = self.get_truncated_bounds()
        if lower is None and upper is None:
            self._truncated_stats = None
            return

        if self.truncate_type in ['value', 'percentile']:
            lb = -float('inf') if lower is None else lower
            ub = float('inf') if upper is None else upper
        else:
            lb = -float('inf') if lower is None else (lower - self.shift_amount)
            ub = float('inf') if upper is None else (upper - self.shift_amount)

        if lb == -float('inf') and ub == float('inf'):
            moments = self._dist.stats(moments='mvsk')
            mean_val = moments[0]
            var_val = moments[1]
            skew_val = moments[2]
            kurt_val = moments[3] + 3.0
        else:
            try:
                m1 = self._dist.expect(lambda x: x, lb=lb, ub=ub, conditional=True)
                m2 = self._dist.expect(lambda x: x**2, lb=lb, ub=ub, conditional=True)
                m3 = self._dist.expect(lambda x: x**3, lb=lb, ub=ub, conditional=True)
                m4 = self._dist.expect(lambda x: x**4, lb=lb, ub=ub, conditional=True)

                mean_val = m1
                var_val = m2 - m1**2
                if var_val > 0:
                    skew_val = (m3 - 3*m1*var_val - m1**3) / (var_val**1.5)
                    kurt_val = (m4 - 4*m1*m3 + 6*m1**2*m2 - 3*m1**4) / (var_val**2)
                else:
                    skew_val = 0.0
                    kurt_val = 3.0
            except Exception as e:
                moments = self._dist.stats(moments='mvsk')
                mean_val = moments[0]
                var_val = moments[1]
                skew_val = moments[2]
                kurt_val = moments[3] + 3.0

        self._truncated_stats = {
            'mean': mean_val,
            'variance': var_val,
            'skewness': skew_val,
            'kurtosis': kurt_val
        }

    def mean(self) -> float:
        if self._truncate_invalid:
            return float('nan')
        if self._truncated_stats:
            return self._truncated_stats['mean'] + self.shift_amount
        return self._dist.mean() + self.shift_amount

    def variance(self) -> float:
        if self._truncate_invalid:
            return float('nan')
        if self._truncated_stats:
            return self._truncated_stats['variance']
        return self._dist.var()

    def skewness(self) -> float:
        if self._truncate_invalid:
            return float('nan')
        if self._truncated_stats:
            return self._truncated_stats['skewness']
        return self._dist.stats(moments='s')

    def kurtosis(self) -> float:
        if self._truncate_invalid:
            return float('nan')
        if self._truncated_stats:
            return self._truncated_stats['kurtosis']
        return self._dist.stats(moments='k') + 3.0

    def mode(self) -> float:
        if self._truncate_invalid:
            return float('nan')
        if self.dfn > 2:
            mode_val = (self.dfn - 2) / self.dfn * self.dfd / (self.dfd + 2)
        else:
            mode_val = 0.0
        return mode_val + self.shift_amount

    def min_val(self) -> float:
        lower, upper = self.get_effective_bounds()
        if lower is not None:
            return lower
        return self.apply_shift(0.0)

    def max_val(self) -> float:
        lower, upper = self.get_effective_bounds()
        if upper is not None:
            return upper
        return float('inf')

    def _original_ppf(self, p: float) -> float:
        return self._dist.ppf(p)

    def _original_cdf(self, x: float) -> float:
        return self._dist.cdf(x)

    def cdf(self, x: float) -> float:
        if self._truncate_invalid:
            return float('nan')
        x_original = self.apply_unshift(x)

        if not self.is_truncated():
            return self._original_cdf(x_original)
        else:
            lower, upper = self.get_truncated_bounds()
            if self.truncate_type in ['value', 'percentile']:
                lower_p = 0.0 if lower is None else self._original_cdf(lower)
                upper_p = 1.0 if upper is None else self._original_cdf(upper)
            else:
                if lower is not None:
                    lower_orig = lower - self.shift_amount
                else:
                    lower_orig = None
                if upper is not None:
                    upper_orig = upper - self.shift_amount
                else:
                    upper_orig = None
                lower_p = 0.0 if lower_orig is None else self._original_cdf(lower_orig)
                upper_p = 1.0 if upper_orig is None else self._original_cdf(upper_orig)

            if self.truncate_type in ['value', 'percentile']:
                if lower is not None and x_original < lower:
                    return 0.0
                if upper is not None and x_original > upper:
                    return 1.0
            else:
                if lower is not None and x < lower:
                    return 0.0
                if upper is not None and x > upper:
                    return 1.0

            orig_cdf = self._original_cdf(x_original)
            return (orig_cdf - lower_p) / (upper_p - lower_p)

    def _original_pdf(self, x: float) -> float:
        return self._dist.pdf(x)

    def pdf(self, x: float) -> float:
        if self._truncate_invalid:
            return float('nan')
        x_original = self.apply_unshift(x)

        if not self.is_truncated():
            return self._original_pdf(x_original)
        else:
            lower, upper = self.get_truncated_bounds()
            if self.truncate_type in ['value', 'percentile']:
                lower_p = 0.0 if lower is None else self._original_cdf(lower)
                upper_p = 1.0 if upper is None else self._original_cdf(upper)
            else:
                if lower is not None:
                    lower_orig = lower - self.shift_amount
                else:
                    lower_orig = None
                if upper is not None:
                    upper_orig = upper - self.shift_amount
                else:
                    upper_orig = None
                lower_p = 0.0 if lower_orig is None else self._original_cdf(lower_orig)
                upper_p = 1.0 if upper_orig is None else self._original_cdf(upper_orig)

            if self.truncate_type in ['value', 'percentile']:
                if lower is not None and x_original < lower:
                    return 0.0
                if upper is not None and x_original > upper:
                    return 0.0
            else:
                if lower is not None and x < lower:
                    return 0.0
                if upper is not None and x > upper:
                    return 0.0

            return self._original_pdf(x_original) / (upper_p - lower_p)


# ==================== t 分布（使用 scipy） ====================
class TDistribution(DistributionBase):
    """t 分布（精确版）"""

    def __init__(self, params: List[float], markers: Dict[str, Any] = None, func_name: str = None):
        if len(params) < 1:
            self.df = 1.0
        else:
            self.df = float(params[0])

        if self.df <= 0:
            self.df = 1.0

        self._dist = sps.t(df=self.df)

        super().__init__(params, markers, func_name)

        if self.truncate_type in ['percentile', 'percentile2']:
            if self.truncate_lower_pct is not None:
                self.truncate_lower = self._original_ppf(self.truncate_lower_pct)
            if self.truncate_upper_pct is not None:
                self.truncate_upper = self._original_ppf(self.truncate_upper_pct)

        self._finalize_truncation()
        self._truncated_stats = None
        self._update_truncated_stats()

    def _update_truncated_stats(self):
        if not self.is_truncated():
            return

        lower, upper = self.get_truncated_bounds()
        if lower is None and upper is None:
            self._truncated_stats = None
            return

        if self.truncate_type in ['value', 'percentile']:
            lb = -float('inf') if lower is None else lower
            ub = float('inf') if upper is None else upper
        else:
            lb = -float('inf') if lower is None else (lower - self.shift_amount)
            ub = float('inf') if upper is None else (upper - self.shift_amount)

        if lb == -float('inf') and ub == float('inf'):
            moments = self._dist.stats(moments='mvsk')
            mean_val = moments[0]
            var_val = moments[1]
            skew_val = moments[2]
            kurt_val = moments[3] + 3.0
        else:
            try:
                m1 = self._dist.expect(lambda x: x, lb=lb, ub=ub, conditional=True)
                m2 = self._dist.expect(lambda x: x**2, lb=lb, ub=ub, conditional=True)
                m3 = self._dist.expect(lambda x: x**3, lb=lb, ub=ub, conditional=True)
                m4 = self._dist.expect(lambda x: x**4, lb=lb, ub=ub, conditional=True)

                mean_val = m1
                var_val = m2 - m1**2
                if var_val > 0:
                    skew_val = (m3 - 3*m1*var_val - m1**3) / (var_val**1.5)
                    kurt_val = (m4 - 4*m1*m3 + 6*m1**2*m2 - 3*m1**4) / (var_val**2)
                else:
                    skew_val = 0.0
                    kurt_val = 3.0
            except Exception as e:
                moments = self._dist.stats(moments='mvsk')
                mean_val = moments[0]
                var_val = moments[1]
                skew_val = moments[2]
                kurt_val = moments[3] + 3.0

        self._truncated_stats = {
            'mean': mean_val,
            'variance': var_val,
            'skewness': skew_val,
            'kurtosis': kurt_val
        }

    def mean(self) -> float:
        if self._truncate_invalid:
            return float('nan')
        if self._truncated_stats:
            return self._truncated_stats['mean'] + self.shift_amount
        return self._dist.mean() + self.shift_amount

    def variance(self) -> float:
        if self._truncate_invalid:
            return float('nan')
        if self._truncated_stats:
            return self._truncated_stats['variance']
        return self._dist.var()

    def skewness(self) -> float:
        if self._truncate_invalid:
            return float('nan')
        if self._truncated_stats:
            return self._truncated_stats['skewness']
        return self._dist.stats(moments='s')

    def kurtosis(self) -> float:
        if self._truncate_invalid:
            return float('nan')
        if self._truncated_stats:
            return self._truncated_stats['kurtosis']
        return self._dist.stats(moments='k') + 3.0

    def mode(self) -> float:
        if self._truncate_invalid:
            return float('nan')
        return self.shift_amount

    def min_val(self) -> float:
        lower, upper = self.get_effective_bounds()
        if lower is not None:
            return lower
        return float('-inf')

    def max_val(self) -> float:
        lower, upper = self.get_effective_bounds()
        if upper is not None:
            return upper
        return float('inf')

    def _original_ppf(self, p: float) -> float:
        return self._dist.ppf(p)

    def _original_cdf(self, x: float) -> float:
        return self._dist.cdf(x)

    def cdf(self, x: float) -> float:
        if self._truncate_invalid:
            return float('nan')
        x_original = self.apply_unshift(x)

        if not self.is_truncated():
            return self._original_cdf(x_original)
        else:
            lower, upper = self.get_truncated_bounds()
            if self.truncate_type in ['value', 'percentile']:
                lower_p = 0.0 if lower is None else self._original_cdf(lower)
                upper_p = 1.0 if upper is None else self._original_cdf(upper)
            else:
                if lower is not None:
                    lower_orig = lower - self.shift_amount
                else:
                    lower_orig = None
                if upper is not None:
                    upper_orig = upper - self.shift_amount
                else:
                    upper_orig = None
                lower_p = 0.0 if lower_orig is None else self._original_cdf(lower_orig)
                upper_p = 1.0 if upper_orig is None else self._original_cdf(upper_orig)

            if self.truncate_type in ['value', 'percentile']:
                if lower is not None and x_original < lower:
                    return 0.0
                if upper is not None and x_original > upper:
                    return 1.0
            else:
                if lower is not None and x < lower:
                    return 0.0
                if upper is not None and x > upper:
                    return 1.0

            orig_cdf = self._original_cdf(x_original)
            return (orig_cdf - lower_p) / (upper_p - lower_p)

    def _original_pdf(self, x: float) -> float:
        return self._dist.pdf(x)

    def pdf(self, x: float) -> float:
        if self._truncate_invalid:
            return float('nan')
        x_original = self.apply_unshift(x)

        if not self.is_truncated():
            return self._original_pdf(x_original)
        else:
            lower, upper = self.get_truncated_bounds()
            if self.truncate_type in ['value', 'percentile']:
                lower_p = 0.0 if lower is None else self._original_cdf(lower)
                upper_p = 1.0 if upper is None else self._original_cdf(upper)
            else:
                if lower is not None:
                    lower_orig = lower - self.shift_amount
                else:
                    lower_orig = None
                if upper is not None:
                    upper_orig = upper - self.shift_amount
                else:
                    upper_orig = None
                lower_p = 0.0 if lower_orig is None else self._original_cdf(lower_orig)
                upper_p = 1.0 if upper_orig is None else self._original_cdf(upper_orig)

            if self.truncate_type in ['value', 'percentile']:
                if lower is not None and x_original < lower:
                    return 0.0
                if upper is not None and x_original > upper:
                    return 0.0
            else:
                if lower is not None and x < lower:
                    return 0.0
                if upper is not None and x > upper:
                    return 0.0

            return self._original_pdf(x_original) / (upper_p - lower_p)


# ==================== 指数分布（使用 scipy） ====================
class ExponentialDistribution(DistributionBase):
    """指数分布（精确版）"""

    def __init__(self, params: List[float], markers: Dict[str, Any] = None, func_name: str = None):
        if len(params) < 1:
            scale = 1.0  # 期望值
        else:
            scale = float(params[0])

        if scale <= 0:
            scale = 1.0

        self.scale = scale
        self.rate = 1.0 / scale

        self._dist = sps.expon(scale=self.scale)

        super().__init__(params, markers, func_name)

        self.scale = 1.0 / self.rate  # 修复：确保scale是1/rate，而不是rate

        if self.truncate_type in ['percentile', 'percentile2']:
            if self.truncate_lower_pct is not None:
                self.truncate_lower = self._original_ppf(self.truncate_lower_pct)
            if self.truncate_upper_pct is not None:
                self.truncate_upper = self._original_ppf(self.truncate_upper_pct)

        self._finalize_truncation()
        self._truncated_stats = None
        self._update_truncated_stats()

    def _update_truncated_stats(self):
        if not self.is_truncated():
            return

        lower, upper = self.get_truncated_bounds()
        if lower is None and upper is None:
            self._truncated_stats = None
            return

        if self.truncate_type in ['value', 'percentile']:
            lb = -float('inf') if lower is None else lower
            ub = float('inf') if upper is None else upper
        else:
            lb = -float('inf') if lower is None else (lower - self.shift_amount)
            ub = float('inf') if upper is None else (upper - self.shift_amount)

        if lb == -float('inf') and ub == float('inf'):
            moments = self._dist.stats(moments='mvsk')
            mean_val = moments[0]
            var_val = moments[1]
            skew_val = moments[2]
            kurt_val = moments[3] + 3.0
        else:
            try:
                m1 = self._dist.expect(lambda x: x, lb=lb, ub=ub, conditional=True)
                m2 = self._dist.expect(lambda x: x**2, lb=lb, ub=ub, conditional=True)
                m3 = self._dist.expect(lambda x: x**3, lb=lb, ub=ub, conditional=True)
                m4 = self._dist.expect(lambda x: x**4, lb=lb, ub=ub, conditional=True)

                mean_val = m1
                var_val = m2 - m1**2
                if var_val > 0:
                    skew_val = (m3 - 3*m1*var_val - m1**3) / (var_val**1.5)
                    kurt_val = (m4 - 4*m1*m3 + 6*m1**2*m2 - 3*m1**4) / (var_val**2)
                else:
                    skew_val = 0.0
                    kurt_val = 3.0
            except Exception as e:
                moments = self._dist.stats(moments='mvsk')
                mean_val = moments[0]
                var_val = moments[1]
                skew_val = moments[2]
                kurt_val = moments[3] + 3.0

        self._truncated_stats = {
            'mean': mean_val,
            'variance': var_val,
            'skewness': skew_val,
            'kurtosis': kurt_val
        }

    def mean(self) -> float:
        if self._truncate_invalid:
            return float('nan')
        if self._truncated_stats:
            return self._truncated_stats['mean'] + self.shift_amount
        return self._dist.mean() + self.shift_amount

    def variance(self) -> float:
        if self._truncate_invalid:
            return float('nan')
        if self._truncated_stats:
            return self._truncated_stats['variance']
        return self._dist.var()

    def skewness(self) -> float:
        if self._truncate_invalid:
            return float('nan')
        if self._truncated_stats:
            return self._truncated_stats['skewness']
        return self._dist.stats(moments='s')

    def kurtosis(self) -> float:
        if self._truncate_invalid:
            return float('nan')
        if self._truncated_stats:
            return self._truncated_stats['kurtosis']
        return self._dist.stats(moments='k') + 3.0

    def mode(self) -> float:
        if self._truncate_invalid:
            return float('nan')
        return self.shift_amount

    def min_val(self) -> float:
        lower, upper = self.get_effective_bounds()
        if lower is not None:
            return lower
        return self.apply_shift(0.0)

    def max_val(self) -> float:
        lower, upper = self.get_effective_bounds()
        if upper is not None:
            return upper
        return float('inf')

    def _original_ppf(self, p: float) -> float:
        return self._dist.ppf(p)

    def _original_cdf(self, x: float) -> float:
        return self._dist.cdf(x)

    def cdf(self, x: float) -> float:
        if self._truncate_invalid:
            return float('nan')
        x_original = self.apply_unshift(x)

        if not self.is_truncated():
            return self._original_cdf(x_original)
        else:
            lower, upper = self.get_truncated_bounds()
            if self.truncate_type in ['value', 'percentile']:
                lower_p = 0.0 if lower is None else self._original_cdf(lower)
                upper_p = 1.0 if upper is None else self._original_cdf(upper)
            else:
                if lower is not None:
                    lower_orig = lower - self.shift_amount
                else:
                    lower_orig = None
                if upper is not None:
                    upper_orig = upper - self.shift_amount
                else:
                    upper_orig = None
                lower_p = 0.0 if lower_orig is None else self._original_cdf(lower_orig)
                upper_p = 1.0 if upper_orig is None else self._original_cdf(upper_orig)

            if self.truncate_type in ['value', 'percentile']:
                if lower is not None and x_original < lower:
                    return 0.0
                if upper is not None and x_original > upper:
                    return 1.0
            else:
                if lower is not None and x < lower:
                    return 0.0
                if upper is not None and x > upper:
                    return 1.0

            orig_cdf = self._original_cdf(x_original)
            return (orig_cdf - lower_p) / (upper_p - lower_p)

    def _original_pdf(self, x: float) -> float:
        return self._dist.pdf(x)

    def pdf(self, x: float) -> float:
        if self._truncate_invalid:
            return float('nan')
        x_original = self.apply_unshift(x)

        if not self.is_truncated():
            return self._original_pdf(x_original)
        else:
            lower, upper = self.get_truncated_bounds()
            if self.truncate_type in ['value', 'percentile']:
                lower_p = 0.0 if lower is None else self._original_cdf(lower)
                upper_p = 1.0 if upper is None else self._original_cdf(upper)
            else:
                if lower is not None:
                    lower_orig = lower - self.shift_amount
                else:
                    lower_orig = None
                if upper is not None:
                    upper_orig = upper - self.shift_amount
                else:
                    upper_orig = None
                lower_p = 0.0 if lower_orig is None else self._original_cdf(lower_orig)
                upper_p = 1.0 if upper_orig is None else self._original_cdf(upper_orig)

            if self.truncate_type in ['value', 'percentile']:
                if lower is not None and x_original < lower:
                    return 0.0
                if upper is not None and x_original > upper:
                    return 0.0
            else:
                if lower is not None and x < lower:
                    return 0.0
                if upper is not None and x > upper:
                    return 0.0

            return self._original_pdf(x_original) / (upper_p - lower_p)