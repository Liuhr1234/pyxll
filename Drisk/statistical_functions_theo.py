# statistical_functions_theo.py
"""理论统计函数模块 - 精确计算分布的理论统计量，支持截断和平移（完全精确版）"""

import re
import math
import numpy as np
from typing import Dict, List, Tuple, Optional, Union, Any
from pyxll import xl_func, xl_app
from attribute_functions import ERROR_MARKER
from formula_parser import extract_all_distribution_functions
import warnings

# ==================== 导入 scipy 进行精确计算 ====================
try:
    import scipy.stats as sps
except ImportError:
    raise ImportError("scipy is required for exact statistical calculations. Please install scipy.")

# 导入分布基类和内置分布
from distribution_base import (
    DistributionBase,
    NormalDistribution,
    UniformDistribution,
    GammaDistribution,
    BetaDistribution,
    PoissonDistribution,
    ChiSquaredDistribution,
    FDistribution,
    TDistribution,
    ExponentialDistribution
)

# 导入分布支持信息
from constants import get_distribution_support, DISTRIBUTION_FUNCTION_NAMES

# ==================== 调试配置 ====================
DEBUG_MODE = False  # 设置为 True 以启用调试输出
USE_STATIC_FOR_NESTED = False  # 嵌套函数参数使用模式：True-用均值，False-用随机样本（默认改为True）

_NESTED_LOCK_ROUND_INT_FUNCS = {
    'driskbernoulli',
    'driskbinomial',
    'driskgeomet',
    'driskhypergeo',
    'driskintuniform',
    'drisknegbin',
}


def _has_invalid_strict_percentile_marker(markers: Dict[str, Any]) -> bool:
    if not markers:
        return False
    for key in ("truncate_p", "truncatep", "truncate_p2", "truncatep2"):
        raw = markers.get(key)
        if raw is None:
            continue
        raw_text = str(raw).strip()
        if raw_text.startswith("(") and raw_text.endswith(")"):
            raw_text = raw_text[1:-1]
        parts = [p.strip() for p in raw_text.split(",")]
        for part in parts:
            if not part:
                continue
            try:
                value = float(part)
            except Exception:
                return True
            if value < 0.0 or value > 1.0:
                return True
    return False


def _formula_has_invalid_erlang_percentile(formula_text: str) -> bool:
    if not formula_text:
        return False
    text = re.sub(r"\s+", "", str(formula_text)).lower()
    if "driskerlang(" not in text:
        return False
    matches = re.findall(r"drisktruncatep2?\(([^)]*)\)", text)
    for match in matches:
        for part in match.split(","):
            part = part.strip()
            if not part:
                continue
            try:
                value = float(part)
            except Exception:
                return True
            if value < 0.0 or value > 1.0:
                return True
    return False

def debug_print(*args, **kwargs):
    """调试打印函数"""
    if DEBUG_MODE:
        print(*args, **kwargs)

# ==================== 新增：增强的分割参数函数（支持花括号）====================
def _split_outer_arguments_enhanced(args_string: str) -> List[str]:
    """
    将参数字符串分割成顶层参数列表，忽略括号和花括号内的逗号
    支持括号 () 和花括号 {} 的嵌套
    Args:
        args_string: 如 "DriskTriang(0,0.5,1), 0.9" 或 "{1,2,3}, {0,0.5,1}"
    Returns:
        分割后的参数列表
    """
    if not args_string:
        return []

    result = []
    current = []
    paren_count = 0          # 普通括号计数
    brace_count = 0          # 花括号计数
    in_quotes = False
    quote_char = None

    for ch in args_string:
        if ch in ('"', "'") and not in_quotes:
            in_quotes = True
            quote_char = ch
            current.append(ch)
        elif ch == quote_char and in_quotes:
            in_quotes = False
            quote_char = None
            current.append(ch)
        elif ch == '(' and not in_quotes:
            paren_count += 1
            current.append(ch)
        elif ch == ')' and not in_quotes:
            paren_count -= 1
            current.append(ch)
        elif ch == '{' and not in_quotes:
            brace_count += 1
            current.append(ch)
        elif ch == '}' and not in_quotes:
            brace_count -= 1
            current.append(ch)
        elif ch == ',' and paren_count == 0 and brace_count == 0 and not in_quotes:
            result.append(''.join(current).strip())
            current = []
        else:
            current.append(ch)

    if current:
        result.append(''.join(current).strip())

    return result

# ==================== 新增：从 Excel 区域获取数值列表 ====================
def _get_range_values(range_str: str) -> Optional[str]:
    """
    从 Excel 区域引用（如 "B5:B7"）中提取数值，返回逗号分隔的字符串。
    如果区域无效或没有数值，返回 None。
    """
    try:
        app = xl_app()
        # 尝试解析区域，可能包含工作表名
        if '!' in range_str:
            sheet_name, addr = range_str.split('!')
            sheet = app.ActiveWorkbook.Worksheets(sheet_name)
            range_obj = sheet.Range(addr)
        else:
            range_obj = app.ActiveSheet.Range(range_str)

        values = range_obj.Value2
        if values is None:
            return None

        # 将 values 展平为一维数值列表
        flat_numbers = []
        if isinstance(values, tuple):
            # 二维区域
            for row in values:
                if isinstance(row, tuple):
                    for cell in row:
                        if cell is not None and isinstance(cell, (int, float)):
                            flat_numbers.append(float(cell))
                else:
                    if row is not None and isinstance(row, (int, float)):
                        flat_numbers.append(float(row))
        else:
            # 单个单元格
            if values is not None and isinstance(values, (int, float)):
                flat_numbers.append(float(values))

        if flat_numbers:
            return ','.join(str(x) for x in flat_numbers)
        return None
    except Exception as e:
        debug_print(f"获取区域值失败 {range_str}: {e}")
        return None

# ==================== 分布工厂（包含新增的伯努利、三角、二项、Trigen、Cumul、Discrete） ====================
class DistributionFactory:
    """分布工厂 - 根据函数名创建对应的分布对象（精确匹配）"""
    # 支持的分布函数名列表（小写，用于精确匹配）
    SUPPORTED_DISTRIBUTIONS = {
        'drisknormal': 'normal',
        'driskuniform': 'uniform',
        'driskerf': 'erf',
        'driskextvalue': 'extvalue',
        'driskextvaluemin': 'extvaluemin',
        'driskfatiguelife': 'fatiguelife',
        'driskfrechet': 'frechet',
        'driskgeneral': 'general',
        'driskhistogrm': 'histogrm',
        'driskhypsecant': 'hypsecant',
        'driskjohnsonsb': 'johnsonsb',
        'driskjohnsonsu': 'johnsonsu',
        'driskkumaraswamy': 'kumaraswamy',
        'drisklaplace': 'laplace',
        'drisklogistic': 'logistic',
        'driskloglogistic': 'loglogistic',
        'drisklognorm': 'lognorm',
        'drisklognorm2': 'lognorm2',
        'driskbetageneral': 'betageneral',
        'driskbetasubj': 'betasubj',
        'driskburr12': 'burr12',
        'driskcompound': 'compound',
        'drisksplice': 'splice',
        'driskpert': 'pert',
        'driskreciprocal': 'reciprocal',
        'driskrayleigh': 'rayleigh',
        'driskweibull': 'weibull',
        'driskpearson5': 'pearson5',
        'driskpearson6': 'pearson6',
        'driskpareto2': 'pareto2',
        'driskpareto': 'pareto',
        'drisklevy': 'levy',
        'driskcauchy': 'cauchy',
        'driskdagum': 'dagum',
        'driskdoubletriang': 'doubletriang',
        'driskgamma': 'gamma',
        'driskerlang': 'erlang',
        'driskbeta': 'beta',
        'driskpoisson': 'poisson',
        'driskchisq': 'chisq',
        'driskf': 'f',
        'driskstudent': 'student',
        'driskexpon': 'expon',
        'driskbernoulli': 'bernoulli',
        'drisktriang': 'triang',
        'driskbinomial': 'binomial',
        'drisknegbin': 'negbin',
        'driskinvgauss': 'invgauss',
        'driskduniform': 'duniform',
        'driskgeomet': 'geomet',
        'driskhypergeo': 'hypergeo',
        'driskintuniform': 'intuniform',
        'drisktrigen': 'trigen',
        'driskcumul': 'cumul',
        'driskdiscrete': 'discrete'
    }

    @staticmethod
    def create_distribution(func_name: str, params: List[float],
                           markers: Dict[str, Any] = None) -> DistributionBase:
        """
        创建分布对象（精确匹配函数名，不区分大小写）

        Args:
            func_name: 函数名（如 "DriskNormal"）
            params: 参数列表
            markers: 属性标记

        Returns:
            DistributionBase: 分布对象
        """
        func_name_lower = func_name.lower()
        if func_name_lower in ['driskintuniform', 'intuniform', 'driskerlang', 'erlang'] and _has_invalid_strict_percentile_marker(markers or {}):
            raise ValueError(f"{func_name} 的百分位截断参数必须介于 0 和 1 之间")

        # 精确匹配
        if func_name_lower in ['drisknormal', 'normal']:
            return NormalDistribution(params, markers, func_name)
        elif func_name_lower in ['driskuniform', 'uniform']:
            return UniformDistribution(params, markers, func_name)
        elif func_name_lower in ['driskerf', 'erf']:
            from dist_erf import ErfDistribution
            return ErfDistribution(params, markers, func_name)
        elif func_name_lower in ['driskextvalue', 'extvalue']:
            from dist_extvalue import ExtvalueDistribution
            return ExtvalueDistribution(params, markers, func_name)
        elif func_name_lower in ['driskextvaluemin', 'extvaluemin']:
            from dist_extvaluemin import ExtvalueMinDistribution
            return ExtvalueMinDistribution(params, markers, func_name)
        elif func_name_lower in ['driskfatiguelife', 'fatiguelife']:
            from dist_fatiguelife import FatigueLifeDistribution
            return FatigueLifeDistribution(params, markers, func_name)
        elif func_name_lower in ['driskfrechet', 'frechet']:
            from dist_frechet import FrechetDistribution
            return FrechetDistribution(params, markers, func_name)
        elif func_name_lower in ['driskgeneral', 'general']:
            from dist_general import GeneralDistribution
            return GeneralDistribution(params, markers, func_name)
        elif func_name_lower in ['driskhistogrm', 'histogrm']:
            from dist_histogrm import HistogrmDistribution
            return HistogrmDistribution(params, markers, func_name)
        elif func_name_lower in ['driskhypsecant', 'hypsecant']:
            from dist_hypsecant import HypSecantDistribution
            return HypSecantDistribution(params, markers, func_name)
        elif func_name_lower in ['driskjohnsonsb', 'johnsonsb']:
            from dist_johnsonsb import JohnsonSBDistribution
            return JohnsonSBDistribution(params, markers, func_name)
        elif func_name_lower in ['driskjohnsonsu', 'johnsonsu']:
            from dist_johnsonsu import JohnsonSUDistribution
            return JohnsonSUDistribution(params, markers, func_name)
        elif func_name_lower in ['driskkumaraswamy', 'kumaraswamy']:
            from dist_kumaraswamy import KumaraswamyDistribution
            return KumaraswamyDistribution(params, markers, func_name)
        elif func_name_lower in ['drisklaplace', 'laplace']:
            from dist_laplace import LaplaceDistribution
            return LaplaceDistribution(params, markers, func_name)
        elif func_name_lower in ['drisklogistic', 'logistic']:
            from dist_logistic import LogisticDistribution
            return LogisticDistribution(params, markers, func_name)
        elif func_name_lower in ['driskloglogistic', 'loglogistic']:
            from dist_loglogistic import LoglogisticDistribution
            return LoglogisticDistribution(params, markers, func_name)
        elif func_name_lower in ['drisklognorm', 'lognorm']:
            from dist_lognorm import LognormDistribution
            return LognormDistribution(params, markers, func_name)
        elif func_name_lower in ['drisklognorm2', 'lognorm2']:
            from dist_lognorm2 import Lognorm2Distribution
            return Lognorm2Distribution(params, markers, func_name)
        elif func_name_lower in ['driskbetageneral', 'betageneral']:
            from dist_betageneral import BetaGeneralDistribution
            return BetaGeneralDistribution(params, markers, func_name)
        elif func_name_lower in ['driskbetasubj', 'betasubj']:
            from dist_betasubj import BetaSubjDistribution
            return BetaSubjDistribution(params, markers, func_name)
        elif func_name_lower in ['driskburr12', 'burr12']:
            from dist_burr12 import Burr12Distribution
            return Burr12Distribution(params, markers, func_name)
        elif func_name_lower in ['driskcompound', 'compound']:
            from dist_compound import CompoundDistribution
            return CompoundDistribution(params, markers, func_name)
        elif func_name_lower in ['drisksplice', 'splice']:
            from dist_splice import SpliceDistribution
            return SpliceDistribution(params, markers, func_name)
        elif func_name_lower in ['driskpert', 'pert']:
            from dist_pert import PertDistribution
            return PertDistribution(params, markers, func_name)
        elif func_name_lower in ['driskreciprocal', 'reciprocal']:
            from dist_reciprocal import ReciprocalDistribution
            return ReciprocalDistribution(params, markers, func_name)
        elif func_name_lower in ['driskrayleigh', 'rayleigh']:
            from dist_rayleigh import RayleighDistribution
            return RayleighDistribution(params, markers, func_name)
        elif func_name_lower in ['driskweibull', 'weibull']:
            from dist_weibull import WeibullDistribution
            return WeibullDistribution(params, markers, func_name)
        elif func_name_lower in ['driskpearson5', 'pearson5']:
            from dist_pearson5 import Pearson5Distribution
            return Pearson5Distribution(params, markers, func_name)
        elif func_name_lower in ['driskpearson6', 'pearson6']:
            from dist_pearson6 import Pearson6Distribution
            return Pearson6Distribution(params, markers, func_name)
        elif func_name_lower in ['driskpareto2', 'pareto2']:
            from dist_pareto2 import Pareto2Distribution
            return Pareto2Distribution(params, markers, func_name)
        elif func_name_lower in ['driskpareto', 'pareto']:
            from dist_pareto import ParetoDistribution
            return ParetoDistribution(params, markers, func_name)
        elif func_name_lower in ['drisklevy', 'levy']:
            from dist_levy import LevyDistribution
            return LevyDistribution(params, markers, func_name)
        elif func_name_lower in ['driskcauchy', 'cauchy']:
            from dist_cauchy import CauchyDistribution
            return CauchyDistribution(params, markers, func_name)
        elif func_name_lower in ['driskdagum', 'dagum']:
            from dist_dagum import DagumDistribution
            return DagumDistribution(params, markers, func_name)
        elif func_name_lower in ['driskdoubletriang', 'doubletriang']:
            from dist_doubletriang import DoubleTriangDistribution
            return DoubleTriangDistribution(params, markers, func_name)
        elif func_name_lower in ['driskgamma', 'gamma']:
            return GammaDistribution(params, markers, func_name)
        elif func_name_lower in ['driskerlang', 'erlang']:
            from dist_erlang import ErlangDistribution
            return ErlangDistribution(params, markers, func_name)
        elif func_name_lower in ['driskbeta', 'beta']:
            return BetaDistribution(params, markers, func_name)
        elif func_name_lower in ['driskpoisson', 'poisson']:
            return PoissonDistribution(params, markers, func_name)
        elif func_name_lower in ['driskchisq', 'driskchi', 'chisq', 'chi2']:
            return ChiSquaredDistribution(params, markers, func_name)
        elif func_name_lower in ['driskf', 'f']:
            return FDistribution(params, markers, func_name)
        elif func_name_lower in ['driskstudent', 'student']:
            return TDistribution(params, markers, func_name)
        elif func_name_lower in ['driskexpon', 'expon', 'exponential']:
            return ExponentialDistribution(params, markers, func_name)
        elif func_name_lower in ['driskbernoulli', 'bernoulli']:
            from dist_bernoulli import BernoulliDistribution
            return BernoulliDistribution(params, markers, func_name)
        elif func_name_lower in ['drisktriang', 'triang']:
            from dist_triang import TriangDistribution
            return TriangDistribution(params, markers, func_name)
        elif func_name_lower in ['driskbinomial', 'binomial']:
            from dist_binomial import BinomialDistribution
            return BinomialDistribution(params, markers, func_name)
        elif func_name_lower in ['drisknegbin', 'negbin']:
            from dist_negbin import NegbinDistribution
            return NegbinDistribution(params, markers, func_name)
        elif func_name_lower in ['driskinvgauss', 'invgauss']:
            from dist_invgauss import InvgaussDistribution
            return InvgaussDistribution(params, markers, func_name)
        elif func_name_lower in ['driskgeomet', 'geomet']:
            from dist_geomet import GeometDistribution
            return GeometDistribution(params, markers, func_name)
        elif func_name_lower in ['driskhypergeo', 'hypergeo']:
            from dist_hypergeo import HypergeoDistribution
            return HypergeoDistribution(params, markers, func_name)
        elif func_name_lower in ['driskintuniform', 'intuniform']:
            from dist_intuniform import IntuniformDistribution
            return IntuniformDistribution(params, markers, func_name)
        elif func_name_lower in ['driskduniform', 'duniform']:
            from dist_duniform import DUniformDistribution
            return DUniformDistribution(params, markers, func_name)
        elif func_name_lower in ['drisktrigen', 'trigen']:
            from dist_trigen import TrigenDistribution
            return TrigenDistribution(params, markers, func_name)
        elif func_name_lower in ['driskcumul', 'cumul']:
            from dist_cumul import CumulDistribution
            return CumulDistribution(params, markers, func_name)
        elif func_name_lower in ['driskdiscrete', 'discrete']:
            from dist_discrete import DiscreteDistribution
            return DiscreteDistribution(params, markers, func_name)
        else:
            # 默认使用正态分布
            print(f"警告：未知分布函数名 '{func_name}'，使用正态分布")
            return NormalDistribution(params, markers, func_name)

# ==================== 以下为原有的辅助函数和 Excel 函数（保持原有逻辑，但已修复嵌套解析） ====================

def _get_formula_from_caller():
    """从调用者获取公式"""
    try:
        app = xl_app()
        caller = app.Caller
        if hasattr(caller, 'Formula'):
            formula = caller.Formula
            if formula and isinstance(formula, str):
                return formula
    except Exception:
        pass
    return None

def _extract_complete_args_text_from_formula(formula: str, func_name: str) -> str:
    """
    从公式中提取完整的参数文本
    修复 extract_all_distribution_functions 返回不完整参数的问题

    Args:
        formula: Excel公式
        func_name: 函数名

    Returns:
        完整的参数文本
    """
    if not formula or not isinstance(formula, str):
        return ""

    # 去掉等号
    if formula.startswith('='):
        formula = formula[1:]

    # 查找函数名位置（不区分大小写）
    pattern = re.compile(re.escape(func_name), re.IGNORECASE)
    match = pattern.search(formula)

    if not match:
        return ""

    # 找到函数名后的第一个左括号
    start_pos = match.end()
    while start_pos < len(formula) and formula[start_pos].isspace():
        start_pos += 1

    if start_pos >= len(formula) or formula[start_pos] != '(':
        return ""

    # 查找匹配的右括号
    paren_count = 1
    pos = start_pos + 1
    while pos < len(formula) and paren_count > 0:
        char = formula[pos]
        if char == '(':
            paren_count += 1
        elif char == ')':
            paren_count -= 1
        pos += 1

    if paren_count > 0:
        return ""

    # 提取参数文本（不包括外层的括号）
    args_text = formula[start_pos + 1:pos - 1]
    return args_text

def _parse_marker_string(marker_str: str) -> Tuple[str, Any]:
    """
    解析标记字符串，提取标记类型和值

    Args:
        marker_str: 标记字符串，如 "DriskTruncate(-1,1)"

    Returns:
        (标记类型, 标记值)
    """
    # 清理空白字符
    marker_str = marker_str.strip()

    # 如果没有括号，可能是简单标记
    if '(' not in marker_str:
        if marker_str.startswith('Drisk'):
            # 提取标记类型
            marker_type = marker_str[5:].lower()  # 去掉"Drisk"
            return marker_type, True
        else:
            return None, None

    # 提取函数名和参数
    match = re.match(r'Drisk(\w+)\s*\(\s*(.*?)\s*\)', marker_str, re.IGNORECASE)
    if not match:
        # 尝试匹配不完整的括号
        match = re.match(r'Drisk(\w+)\s*\(\s*(.*)', marker_str, re.IGNORECASE)
        if match:
            pass
        else:
            return None, None

    marker_type = match.group(1).lower()
    params_str = match.group(2).strip()

    # 检查参数字符串是否包含不完整的括号
    if '(' in params_str and ')' not in params_str:
        # 尝试修复：添加缺失的右括号
        params_str = params_str + ')'

    # 根据标记类型处理参数
    if marker_type in ['truncate', 'truncate2']:
        # 截断标记：参数为两个数值，如 "-1,1"
        if params_str:
            # 清理参数：移除可能的额外括号
            params_str = params_str.replace('(', '').replace(')', '')
            parts = [p.strip() for p in params_str.split(',')]
            if len(parts) >= 2:
                try:
                    lower = "" if not parts[0] else float(_evaluate_parameter(parts[0]))
                    upper = "" if not parts[1] else float(_evaluate_parameter(parts[1]))
                    result = f"{lower},{upper}"
                    return marker_type, result
                except Exception:
                    # 尝试更激进的清理
                    try:
                        cleaned = re.sub(r'[^\d\.,\-]', '', params_str)
                        parts = cleaned.split(',')
                        if len(parts) >= 2:
                            lower = float(parts[0])
                            upper = float(parts[1])
                            result = f"{lower},{upper}"
                            return marker_type, result
                    except Exception:
                        return marker_type, params_str
            else:
                return marker_type, params_str
        else:
            return marker_type, ""

    elif marker_type in ['truncatep', 'truncate_p', 'truncatep2', 'truncate_p2']:
        # 百分比截断标记：参数为两个百分比，如 "5,95"
        # 统一标记类型名称
        if marker_type == 'truncatep':
            marker_type = 'truncate_p'
        elif marker_type == 'truncatep2':
            marker_type = 'truncate_p2'

        # 检查百分比是否在有效范围内
        try:
            parts = [p.strip() for p in params_str.split(',')]
            lower_pct = None
            upper_pct = None
            if len(parts) >= 1 and parts[0]:
                val = float(_evaluate_parameter(parts[0]))
                if 0 <= val <= 1:
                    lower_pct = val
                elif 0 <= val <= 100:
                    lower_pct = val / 100.0
                else:
                    # 无效，返回 None 标记截断无效
                    return None, None
            if len(parts) >= 2 and parts[1]:
                val = float(_evaluate_parameter(parts[1]))
                if 0 <= val <= 1:
                    upper_pct = val
                elif 0 <= val <= 100:
                    upper_pct = val / 100.0
                else:
                    return None, None
            # 重构字符串
            result = f"{lower_pct if lower_pct is not None else ''},{upper_pct if upper_pct is not None else ''}"
            return marker_type, result
        except:
            return marker_type, params_str

    elif marker_type == 'shift':
        # 平移标记：参数为单个数值
        try:
            value = float(params_str)
            return marker_type, value
        except Exception:
            # 尝试清理非数字字符
            try:
                cleaned = re.sub(r'[^\d\.\-]', '', params_str)
                if cleaned:
                    value = float(cleaned)
                    return marker_type, value
            except Exception:
                pass
            return marker_type, params_str

    elif marker_type == 'static':
        # 静态值标记：参数为单个数值
        try:
            value = float(params_str)
            return marker_type, value
        except Exception:
            return marker_type, params_str

    elif marker_type in ['name', 'units', 'category']:
        # 字符串标记：参数为字符串，可能带引号
        # 去除引号
        if params_str and (params_str.startswith('"') or params_str.startswith("'")):
            params_str = params_str[1:-1]
        return marker_type, params_str

    elif marker_type == 'seed':
        # 种子标记：参数为类型和种子值，如 "1,42"
        return marker_type, params_str

    elif marker_type == 'loc':
        # 位置标记：无参数
        return marker_type, True

    elif marker_type in ['isdate', 'is_date']:
        # 日期标记：参数为布尔值
        if params_str.lower() in ['true', '1', 'yes']:
            return 'is_date', True
        else:
            return 'is_date', False

    elif marker_type in ['isdiscrete', 'is_discrete']:
        # 离散标记：参数为布尔值
        if params_str.lower() in ['true', '1', 'yes']:
            return 'is_discrete', True
        else:
            return 'is_discrete', False

    else:
        # 其他标记：返回原始字符串
        return marker_type, params_str

def _fix_unmatched_parentheses(args_string: str) -> str:
    """
    修复不匹配的括号

    Args:
        args_string: 参数字符串

    Returns:
        修复后的字符串
    """
    if not args_string:
        return args_string

    # 统计括号
    open_paren = 0
    in_quotes = False
    quote_char = None

    for char in args_string:
        if char in ['"', "'"] and not in_quotes:
            in_quotes = True
            quote_char = char
        elif char == quote_char and in_quotes:
            in_quotes = False
            quote_char = None
        elif not in_quotes:
            if char == '(':
                open_paren += 1
            elif char == ')':
                open_paren -= 1

    # 如果缺少右括号，添加它们
    if open_paren > 0:
        args_string += ')' * open_paren

    return args_string

def _extract_markers_from_args_string(args_string: str) -> Dict[str, Any]:
    """
    从参数字符串中提取属性标记 - 修复版：正确处理嵌套括号和不完整字符串

    Args:
        args_string: 参数字符串，如 "0, 1, DriskTruncate(-1,1), DriskShift(5)"

    Returns:
        标记字典
    """
    if not args_string:
        return {}

    # 修复不匹配的括号
    args_string = _fix_unmatched_parentheses(args_string)

    markers = {}

    # 分割参数，但要小心处理嵌套括号
    args_list = []
    current_arg = ""
    paren_count = 0
    in_quotes = False
    quote_char = None

    for i, char in enumerate(args_string):
        if char in ['"', "'"] and not in_quotes:
            in_quotes = True
            quote_char = char
            current_arg += char
        elif char == quote_char and in_quotes:
            in_quotes = False
            quote_char = None
            current_arg += char
        elif char == '(' and not in_quotes:
            paren_count += 1
            current_arg += char
        elif char == ')' and not in_quotes:
            paren_count -= 1
            current_arg += char
        elif char == ',' and paren_count == 0 and not in_quotes:
            args_list.append(current_arg.strip())
            current_arg = ""
        else:
            current_arg += char

    if current_arg.strip():
        args_list.append(current_arg.strip())

    # 提取标记
    for arg in args_list:
        # 检查是否以Drisk开头
        arg_lower = arg.strip().lower()
        if arg_lower.startswith('drisk'):
            # 检查是否是分布函数名（如DriskNormal, DriskUniform等）
            is_distribution = False
            for dist_name in DistributionFactory.SUPPORTED_DISTRIBUTIONS.keys():
                # 检查是否以分布函数名开头
                if arg_lower.startswith(dist_name):
                    # 进一步检查：分布函数名后应该是括号，而不是其他字符
                    # 例如：DriskNormal(5,2) 是分布函数
                    #       DriskTruncate(5,2) 是标记
                    if len(arg_lower) > len(dist_name):
                        next_char = arg_lower[len(dist_name)]
                        if next_char == '(':
                            is_distribution = True
                            break
                    else:
                        # 如果长度相等，可能是DriskNormal没有参数，但这种情况不应该发生
                        is_distribution = True
                        break

            # 如果是分布函数，不要将其作为标记处理
            if is_distribution:
                debug_print(f"跳过分布函数参数: {arg}")
                continue

            marker_type, marker_value = _parse_marker_string(arg)
            if marker_type:
                markers[marker_type] = marker_value

    return markers

def _has_invalid_strict_intuniform_percentile_args(args_string: str) -> bool:
    """
    检查 Intuniform 的百分位截断参数是否严格位于 [0, 1]。
    这里保留老代码对其他分布的兼容口径，只对 Intuniform 做严格限制。
    """
    if not args_string:
        return False

    try:
        args_list = _split_outer_arguments_enhanced(args_string)
    except Exception:
        return False

    for arg in args_list:
        text = str(arg).strip()
        lower_text = text.lower()
        if not (
            lower_text.startswith('drisktruncatep(')
            or lower_text.startswith('drisktruncatep2(')
        ):
            continue

        start = text.find('(')
        end = text.rfind(')')
        if start < 0 or end <= start:
            continue

        inner = text[start + 1:end]
        parts = [p.strip() for p in inner.split(',')]
        for part in parts[:2]:
            if not part:
                continue
            try:
                value = float(part)
            except Exception:
                return True
            if value < 0.0 or value > 1.0:
                return True

    return False

def _evaluate_parameter(param_str: str, formula_parser_func=None) -> float:
    """
    评估参数值 - 修复版：支持嵌套分布函数，并根据 USE_STATIC_FOR_NESTED 决定返回均值还是随机样本
    修改：如果标记中存在 'loc' 或 'lock'，则强制返回均值（无论 USE_STATIC_FOR_NESTED）
    """
    debug_print(f"\n===== 开始评估参数 =====")
    debug_print(f"参数字符串: {param_str}")

    if not param_str or not isinstance(param_str, str):
        debug_print("参数为空或不是字符串，返回0.0")
        return 0.0

    param_str = param_str.strip()

    # 1. 尝试作为数值直接转换
    try:
        result = float(param_str)
        debug_print(f"作为数值转换成功: {result}")
        return result
    except:
        debug_print("无法作为数值转换")
        pass

    # 2. 尝试作为单元格引用
    try:
        app = xl_app()
        cell = app.ActiveSheet.Range(param_str)
        val = cell.Value
        if isinstance(val, (int, float, np.number)):
            result = float(val)
            debug_print(f"作为单元格引用获取数值成功: {result}")
            return result
        elif isinstance(val, str):
            # 如果是字符串，尝试转换为数值
            try:
                result = float(val)
                debug_print(f"单元格字符串转换为数值成功: {result}")
                return result
            except:
                # 可能是公式，尝试获取计算结果
                if cell.HasFormula:
                    # 获取计算后的值
                    result = cell.Value
                    if isinstance(result, (int, float, np.number)):
                        result = float(result)
                        debug_print(f"单元格公式计算结果: {result}")
                        return result
    except Exception as e:
        debug_print(f"作为单元格引用处理失败: {e}")
        pass

    # 3. 检查是否是分布函数（包含Drisk且包含括号）
    # 更精确的检查：检查是否包含Drisk且包含括号
    if 'Drisk' in param_str and '(' in param_str and ')' in param_str:
        debug_print("参数包含Drisk且包含括号，尝试作为分布函数解析")

        # 确保是完整的公式
        formula = param_str
        if not formula.startswith('='):
            formula = f'={formula}'

        debug_print(f"处理公式: {formula}")

        # 使用公式解析器提取分布函数
        if formula_parser_func:
            dist_funcs = formula_parser_func(formula)
        else:
            dist_funcs = extract_all_distribution_functions(formula)

        debug_print(f"公式解析函数得到的分布函数数量: {len(dist_funcs)}")

        if dist_funcs:
            # 取第一个分布函数（在参数中，我们期望只有一个分布函数）
            first_func = dist_funcs[0]
            func_name = first_func.get('func_name', '')

            # 提取完整参数文本
            complete_args_text = _extract_complete_args_text_from_formula(formula, func_name)

            if not complete_args_text:
                args_text = first_func.get('args_text', '')
                complete_args_text = args_text

            debug_print(f"找到函数: {func_name}")
            debug_print(f"完整参数文本: {complete_args_text}")

            if func_name.lower() in ['driskintuniform', 'intuniform'] and _has_invalid_strict_intuniform_percentile_args(complete_args_text):
                debug_print("Intuniform 的百分位截断参数超出 [0,1]，返回 NaN")
                return float('nan')

            # 提取标记
            markers = _extract_markers_from_args_string(complete_args_text)
            debug_print(f"提取到的标记: {markers}")

            # 提取分布参数
            dist_params = []

            # 先分割参数
            args_list = _split_outer_arguments_enhanced(complete_args_text)
            debug_print(f"分割后的参数列表: {args_list}")

            # 处理每个参数（递归评估）
            for i, arg in enumerate(args_list):
                debug_print(f"处理参数 {i+1}: {arg}")

                # 检查是否是标记
                arg_lower = arg.strip().lower()
                is_marker = False

                # 检查是否以Drisk开头但不是分布函数
                if arg_lower.startswith('drisk'):
                    # 检查是否是分布函数
                    is_distribution = False
                    for dist_name in DistributionFactory.SUPPORTED_DISTRIBUTIONS.keys():
                        if arg_lower.startswith(dist_name) and len(arg_lower) > len(dist_name) and arg_lower[len(dist_name)] == '(':
                            is_distribution = True
                            break

                    # 如果不是分布函数，则可能是标记
                    if not is_distribution:
                        is_marker = True
                        debug_print(f"  跳过标记参数")

                if not is_marker:
                    # 递归评估参数，支持嵌套分布
                    val = _evaluate_parameter(arg, formula_parser_func)
                    dist_params.append(val)
                    debug_print(f"  递归评估结果: {val}")

            debug_print(f"分布参数: {dist_params}")

            # 创建分布对象（传入 func_name 以支持支持范围检查）
            dist = DistributionFactory.create_distribution(func_name, dist_params, markers)

            # 如果标记中存在 'loc' 或 'lock'，则强制返回理论均值
            if 'loc' in markers or 'lock' in markers:
                result = dist.mean()
                if func_name.lower() in _NESTED_LOCK_ROUND_INT_FUNCS:
                    result = float(round(result))
                debug_print(f"检测到 loc/lock 标记，使用均值模式，结果: {result}")
            else:
                # 否则根据全局变量决定返回均值还是随机样本
                if USE_STATIC_FOR_NESTED:
                    result = dist.mean()
                    debug_print(f"使用均值模式，结果: {result}")
                else:
                    # 生成随机样本（逆变换法）
                    result = dist.ppf(np.random.random())
                    debug_print(f"使用随机模式，结果: {result}")

            debug_print(f"===== 结束评估参数 =====\n")
            return result
        else:
            debug_print("未提取到分布函数")
    else:
        debug_print(f"参数不包含Drisk或括号，尝试其他解析方式")

    # 4. 尝试作为数学表达式
    try:
        # 替换Excel运算符
        expr = param_str.replace('^', '**')
        debug_print(f"尝试作为数学表达式: {expr}")

        # 安全地计算表达式
        val = eval(expr, {"__builtins__": {}}, {
            'exp': math.exp, 'log': math.log, 'sqrt': math.sqrt,
            'sin': math.sin, 'cos': math.cos, 'tan': math.tan,
            'pi': math.pi, 'e': math.e
        })
        result = float(val)
        debug_print(f"数学表达式计算成功: {result}")
        debug_print(f"===== 结束评估参数 =====\n")
        return result
    except Exception as e:
        debug_print(f"数学表达式计算失败: {e}")
        pass

    # 所有尝试都失败，返回0.0
    debug_print("所有尝试都失败，返回0.0")
    debug_print(f"===== 结束评估参数 =====\n")
    return 0.0

def _extract_outermost_distribution(formula: str, formula_parser_func) -> List[Dict]:
    """
    从公式中提取最外层分布函数（忽略非分布函数如 DriskOutput）

    Args:
        formula: Excel公式
        formula_parser_func: 公式解析函数（应为 extract_all_distribution_functions）

    Returns:
        分布函数列表（只包含最外层分布）
    """
    debug_print(f"\n===== 开始提取最外层分布 =====")
    debug_print(f"原始公式: {formula}")

    if not formula or not isinstance(formula, str):
        debug_print("公式为空或不是字符串")
        return []

    # 确保公式以等号开头
    if not formula.startswith('='):
        formula = f'={formula}'

    # 使用公式解析函数提取所有分布函数
    all_dist_funcs = formula_parser_func(formula)
    debug_print(f"提取到的所有分布函数数量: {len(all_dist_funcs)}")

    if not all_dist_funcs:
        debug_print("未提取到任何分布函数")
        return []

    # 按起始位置排序
    all_dist_funcs.sort(key=lambda x: x.get('start_pos', 0))

    # 找出最外层的分布函数（不被任何其他函数包含）
    outermost_funcs = []
    for i, func in enumerate(all_dist_funcs):
        func_start = func.get('start_pos', 0)
        func_end = func.get('end_pos', 0)
        is_contained = False
        for j, other in enumerate(all_dist_funcs):
            if i != j:
                other_start = other.get('start_pos', 0)
                other_end = other.get('end_pos', 0)
                if other_start <= func_start and other_end >= func_end:
                    is_contained = True
                    break
        if not is_contained:
            outermost_funcs.append(func)

    # 按结束位置排序，取最后一个（最右侧的）
    if outermost_funcs:
        outermost_funcs.sort(key=lambda x: x.get('end_pos', 0))
        result = outermost_funcs[-1]
        debug_print(f"选择的最外层分布函数: {result.get('func_name', 'N/A')}")
        debug_print(f"===== 结束提取最外层分布 =====\n")
        return [result]

    # 后备：如果没有找到最外层函数，返回最后一个分布函数
    debug_print("未找到最外层函数，返回最后一个分布函数")
    result = all_dist_funcs[-1]
    debug_print(f"选择的函数: {result.get('func_name', 'N/A')}")
    debug_print(f"===== 结束提取最外层分布 =====\n")
    return [all_dist_funcs[-1]]

# ==================== 修改：从数据源提取分布信息的简化版本（已修复对花括号的支持） ====================
def _extract_distribution_from_data_source_simple(data_source) -> Optional[DistributionBase]:
    """
    从数据源提取分布信息的简化版本（已修复嵌套函数支持，并增强对花括号数组的处理）
    如果 data_source 是数值，则尝试从调用者公式中恢复表达式。
    如果失败，返回 None 并记录警告。
    """
    debug_print(f"\n===== 开始提取分布信息 =====")
    debug_print(f"数据源类型: {type(data_source)}")
    debug_print(f"数据源值: {data_source}")

    try:
        # 情况1: 如果data_source是数值（Excel传递的计算结果）
        if isinstance(data_source, (int, float, np.number)):
            debug_print("数据源是数值类型")

            # 获取调用者公式，直接尝试提取
            formula = _get_formula_from_caller()
            debug_print(f"从调用者获取的公式: {formula}")

            if not formula or not isinstance(formula, str):
                debug_print("未获取到有效的公式")
                return None

            # 提取外层函数（应该是 DriskTheoMean 等）的参数文本
            match = re.search(r'^=([A-Za-z_][A-Za-z0-9_]*)\(', formula, re.IGNORECASE)
            if match:
                outer_func = match.group(1)
                param_expr = _extract_complete_args_text_from_formula(formula, outer_func)
                debug_print(f"提取的外层函数参数表达式: {param_expr}")

                # 使用增强的分割函数分割参数，取第一个作为分布表达式
                args_list = _split_outer_arguments_enhanced(param_expr)
                if args_list:
                    nested_expr = args_list[0].strip()
                    # 检查是否为分布函数表达式（包含 Drisk 和括号）或者包含花括号数组
                    if ('Drisk' in nested_expr and '(' in nested_expr and ')' in nested_expr) or \
                       (nested_expr.startswith('{') and nested_expr.endswith('}')):
                        nested_formula = f"={nested_expr}"
                        return _parse_distribution_from_formula_string(nested_formula)
                    else:
                        # 否则按原有逻辑处理单元格引用（可能参数是单元格引用，如 "E43"）
                        cell_ref = nested_expr
                        debug_print(f"单元格引用: {cell_ref}")

                        try:
                            app = xl_app()
                            cell = app.ActiveSheet.Range(cell_ref)
                            formula = cell.Formula
                            debug_print(f"单元格公式: {formula}")

                            if formula and isinstance(formula, str):
                                return _parse_distribution_from_formula_string(formula)
                        except Exception as e:
                            debug_print(f"处理单元格引用时出错: {e}")
                            return None
                else:
                    debug_print("参数表达式为空，无法提取分布")
                    return None
            else:
                debug_print("无法从公式中提取外层函数名")
                return None

        # 情况2: 如果data_source是字符串
        elif isinstance(data_source, str):
            debug_print("数据源是字符串类型")

            # 检查是否是直接传入的分布函数字符串或花括号数组
            if ('Drisk' in data_source and '(' in data_source) or \
               (data_source.startswith('{') and data_source.endswith('}')):
                debug_print("数据源是直接的分布函数字符串或花括号数组")
                return _parse_distribution_from_formula_string(data_source)
            else:
                debug_print("数据源不是直接的分布函数字符串，可能是单元格引用")
                # 可能是单元格引用
                try:
                    app = xl_app()
                    cell = app.ActiveSheet.Range(data_source)
                    formula = cell.Formula
                    debug_print(f"单元格公式: {formula}")

                    if formula and isinstance(formula, str):
                        return _parse_distribution_from_formula_string(formula)
                except Exception as e:
                    debug_print(f"处理字符串单元格引用时出错: {e}")
                    return None

        debug_print("数据源类型不支持")
        return None

    except Exception as e:
        debug_print(f"提取分布信息失败: {e}")
        import traceback
        debug_print(f"错误详情: {traceback.format_exc()}")
        return None

# ==================== 修改：从公式字符串解析分布对象（支持特殊分布和花括号，增强区域引用支持） ====================
def _parse_distribution_from_formula_string(formula_str: str) -> Optional[DistributionBase]:
    """
    从公式字符串解析分布对象（内部函数），增强对花括号数组和单元格区域引用的支持

    Args:
        formula_str: 公式字符串，必须以等号开头，例如 "=DriskNormal(0,1)" 或 "={1,2,3}"

    Returns:
        分布对象，如果解析失败则返回 None
    """
    debug_print(f"\n===== 开始从公式字符串解析分布 =====")
    debug_print(f"公式字符串: {formula_str}")

    if not formula_str or not isinstance(formula_str, str):
        return None

    # 确保以等号开头
    if not formula_str.startswith('='):
        formula_str = f'={formula_str}'
    if _formula_has_invalid_erlang_percentile(formula_str):
        debug_print("Erlang 的百分位截断参数必须介于 0 和 1 之间")
        return None

    # 提取分布函数信息
    dist_funcs = _extract_outermost_distribution(formula_str, extract_all_distribution_functions)
    if not dist_funcs:
        debug_print("未提取到分布函数")
        return None

    last_func = dist_funcs[-1]
    func_name = last_func.get('func_name', '')

    debug_print(f"找到的函数名: {func_name}")

    # 提取完整参数文本
    complete_args_text = _extract_complete_args_text_from_formula(formula_str, func_name)
    if not complete_args_text:
        args_text = last_func.get('args_text', '')
        complete_args_text = args_text

    debug_print(f"完整的参数文本: {complete_args_text}")

    if func_name.lower() in ['driskintuniform', 'intuniform'] and _has_invalid_strict_intuniform_percentile_args(complete_args_text):
        debug_print("Intuniform 的百分位截断参数超出 [0,1]，理论分布判定为无效")
        return None

    # 提取标记
    markers = _extract_markers_from_args_string(complete_args_text)
    debug_print(f"提取到的标记: {markers}")

    # 分割参数（使用增强的分割函数）
    args_list = _split_outer_arguments_enhanced(complete_args_text)
    debug_print(f"分割后的参数列表: {args_list}")

    # 特殊处理 DriskCumul 和 DriskDiscrete：不评估参数，而是将数组参数转换为数值列表字符串存入 markers
    if func_name.lower() == 'driskcumul':
        # DriskCumul 需要四个参数：min, max, x_vals, p_vals
        if len(args_list) >= 4:
            # 辅助函数：将参数转换为数值列表字符串
            def param_to_string(param):
                param = param.strip()
                # 如果是单元格区域引用（如 "B5:B7" 或 "Sheet1!B5:B7"）
                if re.match(r'^([A-Za-z0-9_]+!)?[A-Z]+[0-9]+:[A-Z]+[0-9]+$', param, re.IGNORECASE):
                    # 从区域获取数值
                    range_str = _get_range_values(param)
                    if range_str:
                        return range_str
                    else:
                        # 获取失败，返回原字符串（可能导致后续错误）
                        return param
                # 如果是单个单元格引用（如 "B5"），获取该单元格的值并尝试解析
                elif re.match(r'^([A-Za-z0-9_]+!)?[A-Z]+[0-9]+$', param, re.IGNORECASE):
                    # 获取单元格的值
                    try:
                        app = xl_app()
                        if '!' in param:
                            sheet_name, addr = param.split('!')
                            sheet = app.ActiveWorkbook.Worksheets(sheet_name)
                            cell = sheet.Range(addr)
                        else:
                            cell = app.ActiveSheet.Range(param)
                        val = cell.Value2
                        if val is None:
                            return param
                        # 如果值是数组（如数组公式返回多单元格），需要处理
                        if isinstance(val, tuple):
                            # 展平
                            flat = []
                            for row in val:
                                if isinstance(row, tuple):
                                    for v in row:
                                        if v is not None and isinstance(v, (int, float)):
                                            flat.append(float(v))
                                else:
                                    if row is not None and isinstance(row, (int, float)):
                                        flat.append(float(row))
                            if flat:
                                return ','.join(str(x) for x in flat)
                        elif isinstance(val, (int, float)):
                            return str(float(val))
                        elif isinstance(val, str):
                            # 如果值是字符串，可能本身就是逗号分隔或花括号
                            return val
                    except Exception as e:
                        debug_print(f"读取单元格 {param} 失败: {e}")
                        return param
                # 其他情况（花括号或逗号分隔字符串），直接返回
                return param

            min_raw = args_list[0].strip().strip('"\'')
            max_raw = args_list[1].strip().strip('"\'')
            x_raw = args_list[2].strip().strip('"\'')
            p_raw = args_list[3].strip().strip('"\'')

            # 解析 min 和 max 为数值
            try:
                min_val = float(param_to_string(min_raw) if param_to_string(min_raw) != min_raw else min_raw)
                max_val = float(param_to_string(max_raw) if param_to_string(max_raw) != max_raw else max_raw)
            except:
                debug_print(f"解析 min/max 失败")
                return None

            x_vals_str = param_to_string(x_raw)
            p_vals_str = param_to_string(p_raw)

            # 解析内部点，验证非空
            def parse_numbers(s):
                s = s.strip()
                if s.startswith('{') and s.endswith('}'):
                    s = s[1:-1]
                parts = [p.strip() for p in s.split(',') if p.strip()]
                return [float(p) for p in parts]

            x_inner = parse_numbers(x_vals_str) if x_vals_str else []
            p_inner = parse_numbers(p_vals_str) if p_vals_str else []

            if len(x_inner) == 0 or len(p_inner) == 0:
                debug_print(f"错误：内部点不能为空")
                return None
            if len(x_inner) != len(p_inner):
                debug_print(f"错误：X 和 P 数组长度不相等")
                return None

            # 构造完整数组
            x_full = [min_val] + x_inner + [max_val]
            p_full = [0.0] + p_inner + [1.0]

            x_full_str = ','.join(str(x) for x in x_full)
            p_full_str = ','.join(str(p) for p in p_full)

            markers['x_vals'] = x_full_str
            markers['p_vals'] = p_full_str
            # 剩余参数（如果有）作为普通参数，但通常不会有
            dist_params = [min_val, max_val, x_full_str, p_full_str]
            debug_print(f"特殊处理 DriskCumul，x_full={x_full_str}, p_full={p_full_str}")
        else:
            debug_print(f"错误：DriskCumul 需要至少四个参数")
            return None
    elif func_name.lower() == 'driskgeneral':
        if len(args_list) >= 4:
            def param_to_string(param):
                param = param.strip()
                if re.match(r'^([A-Za-z0-9_]+!)?[A-Z]+[0-9]+:[A-Z]+[0-9]+$', param, re.IGNORECASE):
                    range_str = _get_range_values(param)
                    return range_str if range_str else param
                elif re.match(r'^([A-Za-z0-9_]+!)?[A-Z]+[0-9]+$', param, re.IGNORECASE):
                    try:
                        app = xl_app()
                        if '!' in param:
                            sheet_name, addr = param.split('!')
                            sheet = app.ActiveWorkbook.Worksheets(sheet_name)
                            cell = sheet.Range(addr)
                        else:
                            cell = app.ActiveSheet.Range(param)
                        val = cell.Value2
                        if val is None:
                            return param
                        if isinstance(val, tuple):
                            flat = []
                            for row in val:
                                if isinstance(row, tuple):
                                    for v in row:
                                        if v is not None and isinstance(v, (int, float)):
                                            flat.append(float(v))
                                else:
                                    if row is not None and isinstance(row, (int, float)):
                                        flat.append(float(row))
                            if flat:
                                return ','.join(str(x) for x in flat)
                        elif isinstance(val, (int, float)):
                            return str(float(val))
                        elif isinstance(val, str):
                            return val
                    except Exception as e:
                        debug_print(f"读取单元格 {param} 失败: {e}")
                        return param
                return param

            min_raw = args_list[0].strip().strip('"\'')
            max_raw = args_list[1].strip().strip('"\'')
            x_raw = args_list[2].strip().strip('"\'')
            p_raw = args_list[3].strip().strip('"\'')

            try:
                min_val = float(param_to_string(min_raw) if param_to_string(min_raw) != min_raw else min_raw)
                max_val = float(param_to_string(max_raw) if param_to_string(max_raw) != max_raw else max_raw)
            except Exception:
                debug_print("解析 General 的 min/max 失败")
                return None

            x_vals_str = param_to_string(x_raw)
            p_vals_str = param_to_string(p_raw)
            try:
                from dist_general import _parse_general_arrays as general_parse_arrays
                general_parse_arrays(min_val, max_val, x_vals_str, p_vals_str)
            except Exception as e:
                debug_print(f"General 参数验证失败: {e}")
                return None

            markers['x_vals'] = x_vals_str
            markers['p_vals'] = p_vals_str
            dist_params = [min_val, max_val, x_vals_str, p_vals_str]
            debug_print(f"特殊处理 DriskGeneral，x_vals={x_vals_str}, p_vals={p_vals_str}")
        else:
            debug_print("错误：DriskGeneral 需要至少四个参数")
            return None
    elif func_name.lower() == 'driskhistogrm':
        if len(args_list) >= 3:
            def param_to_string(param):
                param = param.strip()
                if re.match(r'^([A-Za-z0-9_]+!)?[A-Z]+[0-9]+:[A-Z]+[0-9]+$', param, re.IGNORECASE):
                    range_str = _get_range_values(param)
                    return range_str if range_str else param
                elif re.match(r'^([A-Za-z0-9_]+!)?[A-Z]+[0-9]+$', param, re.IGNORECASE):
                    try:
                        app = xl_app()
                        if '!' in param:
                            sheet_name, addr = param.split('!')
                            sheet = app.ActiveWorkbook.Worksheets(sheet_name)
                            cell = sheet.Range(addr)
                        else:
                            cell = app.ActiveSheet.Range(param)
                        val = cell.Value2
                        if val is None:
                            return param
                        if isinstance(val, tuple):
                            flat = []
                            for row in val:
                                if isinstance(row, tuple):
                                    for v in row:
                                        if v is not None and isinstance(v, (int, float)):
                                            flat.append(float(v))
                                else:
                                    if row is not None and isinstance(row, (int, float)):
                                        flat.append(float(row))
                            if flat:
                                return ','.join(str(x) for x in flat)
                        elif isinstance(val, (int, float)):
                            return str(float(val))
                        elif isinstance(val, str):
                            return val
                    except Exception as e:
                        debug_print(f"读取单元格 {param} 失败: {e}")
                        return param
                return param

            min_raw = args_list[0].strip().strip('"\'')
            max_raw = args_list[1].strip().strip('"\'')
            p_raw = args_list[2].strip().strip('"\'')

            try:
                min_val = float(param_to_string(min_raw) if param_to_string(min_raw) != min_raw else min_raw)
                max_val = float(param_to_string(max_raw) if param_to_string(max_raw) != max_raw else max_raw)
            except Exception:
                debug_print("解析 Histogrm 的 min/max 失败")
                return None

            p_vals_str = param_to_string(p_raw)
            try:
                from dist_histogrm import _parse_histogrm_p_table as histogrm_parse_p_table
                histogrm_parse_p_table(p_vals_str)
            except Exception as e:
                debug_print(f"Histogrm 参数验证失败: {e}")
                return None

            markers['p_vals'] = p_vals_str
            dist_params = [min_val, max_val, p_vals_str]
            debug_print(f"特殊处理 DriskHistogrm，p_vals={p_vals_str}")
        else:
            debug_print("错误：DriskHistogrm 需要至少三个参数")
            return None
    elif func_name.lower() == 'driskduniform':
        if len(args_list) >= 1:
            def param_to_string(param):
                param = param.strip()
                if re.match(r'^([A-Za-z0-9_]+!)?[A-Z]+[0-9]+:[A-Z]+[0-9]+$', param, re.IGNORECASE):
                    range_str = _get_range_values(param)
                    return range_str if range_str else param
                elif re.match(r'^([A-Za-z0-9_]+!)?[A-Z]+[0-9]+$', param, re.IGNORECASE):
                    try:
                        app = xl_app()
                        if '!' in param:
                            sheet_name, addr = param.split('!')
                            sheet = app.ActiveWorkbook.Worksheets(sheet_name)
                            cell = sheet.Range(addr)
                        else:
                            cell = app.ActiveSheet.Range(param)
                        val = cell.Value2
                        if val is None:
                            return param
                        if isinstance(val, tuple):
                            flat = []
                            for row in val:
                                if isinstance(row, tuple):
                                    for v in row:
                                        if v is not None and isinstance(v, (int, float)):
                                            flat.append(float(v))
                                else:
                                    if row is not None and isinstance(row, (int, float)):
                                        flat.append(float(row))
                            if flat:
                                return ','.join(str(x) for x in flat)
                        elif isinstance(val, (int, float)):
                            return str(float(val))
                        elif isinstance(val, str):
                            return val
                    except Exception as e:
                        debug_print(f"????????{param} ???: {e}")
                        return param
                return param

            x_raw = args_list[0].strip().strip("'\"")
            x_vals_str = param_to_string(x_raw)

            def parse_numbers(s):
                s = s.strip()
                if s.startswith('{') and s.endswith('}'):
                    s = s[1:-1]
                parts = [p.strip() for p in s.split(',') if p.strip()]
                return [float(p) for p in parts]

            x_vals = parse_numbers(x_vals_str) if x_vals_str else []
            if len(x_vals) == 0:
                debug_print(f"?????-Table ??????")
                return None
            if len(x_vals) != len(set(x_vals)):
                debug_print(f"?????-Table ???????????")
                return None

            p = 1.0 / len(x_vals)
            p_vals = [p] * len(x_vals)
            p_vals_str = ','.join(str(v) for v in p_vals)

            markers['x_vals'] = x_vals_str
            markers['p_vals'] = p_vals_str
            dist_params = [x_vals_str]
            debug_print(f"?????? DriskDUniform??_vals={x_vals_str}, p_vals={p_vals_str}")
        else:
            debug_print(f"?????riskDUniform ?????????????")
            return None
    elif func_name.lower() == 'driskdiscrete':
        # DriskDiscrete 需要两个参数
        if len(args_list) >= 2:
            # 辅助函数：将参数转换为数值列表字符串
            def param_to_string(param):
                param = param.strip()
                # 如果是单元格区域引用（如 "B5:B7" 或 "Sheet1!B5:B7"）
                if re.match(r'^([A-Za-z0-9_]+!)?[A-Z]+[0-9]+:[A-Z]+[0-9]+$', param, re.IGNORECASE):
                    # 从区域获取数值
                    range_str = _get_range_values(param)
                    if range_str:
                        return range_str
                    else:
                        return param
                # 如果是单个单元格引用（如 "B5"），获取该单元格的值并尝试解析
                elif re.match(r'^([A-Za-z0-9_]+!)?[A-Z]+[0-9]+$', param, re.IGNORECASE):
                    try:
                        app = xl_app()
                        if '!' in param:
                            sheet_name, addr = param.split('!')
                            sheet = app.ActiveWorkbook.Worksheets(sheet_name)
                            cell = sheet.Range(addr)
                        else:
                            cell = app.ActiveSheet.Range(param)
                        val = cell.Value2
                        if val is None:
                            return param
                        if isinstance(val, tuple):
                            flat = []
                            for row in val:
                                if isinstance(row, tuple):
                                    for v in row:
                                        if v is not None and isinstance(v, (int, float)):
                                            flat.append(float(v))
                                else:
                                    if row is not None and isinstance(row, (int, float)):
                                        flat.append(float(row))
                            if flat:
                                return ','.join(str(x) for x in flat)
                        elif isinstance(val, (int, float)):
                            return str(float(val))
                        elif isinstance(val, str):
                            return val
                    except Exception as e:
                        debug_print(f"读取单元格 {param} 失败: {e}")
                        return param
                return param

            x_raw = args_list[0].strip().strip('"\'')
            p_raw = args_list[1].strip().strip('"\'')
            x_vals_str = param_to_string(x_raw)
            p_vals_str = param_to_string(p_raw)

            # 验证非空
            def parse_numbers(s):
                s = s.strip()
                if s.startswith('{') and s.endswith('}'):
                    s = s[1:-1]
                parts = [p.strip() for p in s.split(',') if p.strip()]
                return [float(p) for p in parts]

            x_vals = parse_numbers(x_vals_str) if x_vals_str else []
            p_vals = parse_numbers(p_vals_str) if p_vals_str else []

            if len(x_vals) == 0 or len(p_vals) == 0:
                debug_print(f"错误：X-Table 或 P-Table 不能为空")
                return None
            if len(x_vals) != len(p_vals):
                debug_print(f"错误：X 和 P 数组长度不相等")
                return None

            # 对于离散分布，概率需要归一化
            total = sum(p_vals)
            if total > 0:
                p_vals = [p / total for p in p_vals]
            p_vals_str = ','.join(str(p) for p in p_vals)

            markers['x_vals'] = x_vals_str
            markers['p_vals'] = p_vals_str
            dist_params = [x_vals_str, p_vals_str]
            debug_print(f"特殊处理 DriskDiscrete，x_vals={x_vals_str}, p_vals={p_vals_str}")
        else:
            debug_print(f"错误：DriskDiscrete 需要至少两个参数")
            return None
    else:
        # 普通分布：递归评估每个参数
        dist_params = []
        if func_name.lower() == 'driskcompound':
            if len(args_list) < 2:
                debug_print(f"错误：DriskCompound 需要至少两个参数")
                return None
            frequency_formula = args_list[0].strip().strip('"\'')
            severity_formula = args_list[1].strip().strip('"\'')
            deductible_raw = args_list[2].strip().strip('"\'') if len(args_list) >= 3 else ''
            limit_raw = args_list[3].strip().strip('"\'') if len(args_list) >= 4 else ''
            try:
                deductible = 0.0 if deductible_raw == '' else float(deductible_raw)
                limit = float('inf') if limit_raw == '' else float(limit_raw)
            except Exception:
                debug_print("Compound 参数解析失败")
                return None
            dist_params = [frequency_formula, severity_formula, deductible, limit]
        elif func_name.lower() == 'drisksplice':
            if len(args_list) < 3:
                debug_print("错误：DriskSplice 需要三个参数")
                return None
            left_formula = args_list[0].strip().strip('"\'')
            right_formula = args_list[1].strip().strip('"\'')
            splice_raw = args_list[2].strip().strip('"\'')
            try:
                splice_point = float(splice_raw)
            except Exception:
                debug_print("Splice 参数解析失败")
                return None
            dist_params = [left_formula, right_formula, splice_point]
        for i, arg in enumerate(args_list):
            # 检查是否是标记
            arg_lower = arg.strip().lower()
            is_marker = False
            if arg_lower.startswith('drisk'):
                # 检查是否是分布函数
                is_distribution = False
                for dist_name in DistributionFactory.SUPPORTED_DISTRIBUTIONS.keys():
                    if arg_lower.startswith(dist_name) and len(arg_lower) > len(dist_name) and arg_lower[len(dist_name)] == '(':
                        is_distribution = True
                        break
                if not is_distribution:
                    is_marker = True
                    debug_print(f"  跳过标记参数")

            if not is_marker:
                val = _evaluate_parameter(arg, extract_all_distribution_functions)
                dist_params.append(val)
                debug_print(f"  参数值: {val}")

    debug_print(f"最终的分布参数: {dist_params}")

    if func_name.lower() == 'driskcompound':
        dist_params = [frequency_formula, severity_formula, deductible, limit]
    elif func_name.lower() == 'drisksplice':
        dist_params = [left_formula, right_formula, splice_point]

    try:
        dist = DistributionFactory.create_distribution(func_name, dist_params, markers)
        # 检查分布是否有效，若无效则返回 None
        if not dist.is_valid():
            debug_print("分布无效（截断无效或参数错误），返回 None")
            return None
    except Exception as e:
        debug_print(f"创建分布对象失败: {e}")
        return None
    debug_print(f"创建分布对象: {type(dist).__name__}")
    debug_print(f"===== 结束从公式字符串解析分布 =====\n")
    return dist

# ==================== Excel函数 ====================
# 所有函数均返回 #ERROR! 当分布无效时

@xl_func("var data_source: var", category="Drisk Theoretical", volatile=True)
def DriskTheoMean(data_source):
    """返回理论分布的均值"""
    dist = _extract_distribution_from_data_source_simple(data_source)
    if dist is None:
        return ERROR_MARKER

    try:
        result = dist.mean()
        if result == float('inf') or result == float('-inf') or np.isnan(result):
            return ERROR_MARKER
        return result
    except Exception as e:
        warnings.warn(f"计算理论均值失败: {e}")
        return ERROR_MARKER

@xl_func("var data_source: var", category="Drisk Theoretical", volatile=True)
def DriskTheoStdDev(data_source):
    """返回理论分布的标准差"""
    dist = _extract_distribution_from_data_source_simple(data_source)
    if dist is None:
        return ERROR_MARKER

    try:
        result = dist.std_dev()
        if result == float('inf') or result == float('-inf') or np.isnan(result):
            return ERROR_MARKER
        return result
    except Exception as e:
        warnings.warn(f"计算理论标准差失败: {e}")
        return ERROR_MARKER

@xl_func("var data_source: var", category="Drisk Theoretical", volatile=True)
def DriskTheoVariance(data_source):
    """返回理论分布的方差"""
    dist = _extract_distribution_from_data_source_simple(data_source)
    if dist is None:
        return ERROR_MARKER

    try:
        result = dist.variance()
        if result == float('inf') or result == float('-inf') or np.isnan(result):
            return ERROR_MARKER
        return result
    except Exception as e:
        warnings.warn(f"计算理论方差失败: {e}")
        return ERROR_MARKER

@xl_func("var data_source: var", category="Drisk Theoretical", volatile=True)
def DriskTheoSkewness(data_source):
    """返回理论分布的偏度"""
    dist = _extract_distribution_from_data_source_simple(data_source)
    if dist is None:
        return ERROR_MARKER

    try:
        result = dist.skewness()
        if result == float('inf') or result == float('-inf') or np.isnan(result):
            return ERROR_MARKER
        return result
    except Exception as e:
        warnings.warn(f"计算理论偏度失败: {e}")
        return ERROR_MARKER

@xl_func("var data_source: var", category="Drisk Theoretical", volatile=True)
def DriskTheoKurtosis(data_source):
    """返回理论分布的峰度（普通峰度，正态分布为3）"""
    dist = _extract_distribution_from_data_source_simple(data_source)
    if dist is None:
        return ERROR_MARKER

    try:
        result = dist.kurtosis()
        if result == float('inf') or result == float('-inf') or np.isnan(result):
            return ERROR_MARKER
        return result
    except Exception as e:
        warnings.warn(f"计算理论峰度失败: {e}")
        return ERROR_MARKER

@xl_func("var data_source: var", category="Drisk Theoretical", volatile=True)
def DriskTheoMin(data_source):
    """返回理论分布的最小值"""
    dist = _extract_distribution_from_data_source_simple(data_source)
    if dist is None:
        return ERROR_MARKER

    try:
        result = dist.min_val()
        if result == float('-inf') or result == float('inf') or np.isnan(result):
            return ERROR_MARKER
        return result
    except Exception as e:
        warnings.warn(f"计算理论最小值失败: {e}")
        return ERROR_MARKER

@xl_func("var data_source: var", category="Drisk Theoretical", volatile=True)
def DriskTheoMax(data_source):
    """返回理论分布的最大值"""
    dist = _extract_distribution_from_data_source_simple(data_source)
    if dist is None:
        return ERROR_MARKER

    try:
        result = dist.max_val()
        if result == float('-inf') or result == float('inf') or np.isnan(result):
            return ERROR_MARKER
        return result
    except Exception as e:
        warnings.warn(f"计算理论最大值失败: {e}")
        return ERROR_MARKER

@xl_func("var data_source: var", category="Drisk Theoretical", volatile=True)
def DriskTheoRange(data_source):
    """返回理论分布的极差"""
    dist = _extract_distribution_from_data_source_simple(data_source)
    if dist is None:
        return ERROR_MARKER

    try:
        result = dist.range_val()
        if result == float('inf') or result == float('-inf') or np.isnan(result):
            return ERROR_MARKER
        return result
    except Exception as e:
        warnings.warn(f"计算理论极差失败: {e}")
        return ERROR_MARKER

@xl_func("var data_source: var", category="Drisk Theoretical", volatile=True)
def DriskTheoMode(data_source):
    """返回理论分布的众数"""
    dist = _extract_distribution_from_data_source_simple(data_source)
    if dist is None:
        return ERROR_MARKER

    try:
        result = dist.mode()
        if result == float('inf') or result == float('-inf') or np.isnan(result):
            return ERROR_MARKER
        return result
    except Exception as e:
        warnings.warn(f"计算理论众数失败: {e}")
        return ERROR_MARKER

@xl_func("var data_source, float p_value: var", category="Drisk Theoretical", volatile=True)
def DriskTheoPtoX(data_source, p_value):
    """返回理论分布在指定概率p下对应的x值"""
    if p_value < 0 or p_value > 1:
        return ERROR_MARKER

    dist = _extract_distribution_from_data_source_simple(data_source)
    if dist is None:
        return ERROR_MARKER

    try:
        result = dist.ppf(p_value)
        if result == float('inf') or result == float('-inf') or np.isnan(result):
            return ERROR_MARKER
        return result
    except Exception as e:
        warnings.warn(f"计算理论分位数失败: {e}")
        return ERROR_MARKER

@xl_func("var data_source, float x_value: var", category="Drisk Theoretical", volatile=True)
def DriskTheoXtoP(data_source, x_value):
    """返回理论分布在指定x值处的累积概率"""
    dist = _extract_distribution_from_data_source_simple(data_source)
    if dist is None:
        return ERROR_MARKER

    try:
        result = dist.cdf(x_value)
        if np.isnan(result):
            return ERROR_MARKER
        return result
    except Exception as e:
        warnings.warn(f"计算理论累积概率失败: {e}")
        return ERROR_MARKER

@xl_func("var data_source, float x_value: var", category="Drisk Theoretical", volatile=True)
def DriskTheoXtoY(data_source, x_value):
    """返回理论分布在指定x值处的y值（概率密度/质量函数值）"""
    dist = _extract_distribution_from_data_source_simple(data_source)
    if dist is None:
        return ERROR_MARKER

    try:
        result = dist.pdf(x_value)
        if result == float('inf') or result == float('-inf') or np.isnan(result):
            return ERROR_MARKER
        return result
    except Exception as e:
        warnings.warn(f"计算理论密度函数失败: {e}")
        return ERROR_MARKER
