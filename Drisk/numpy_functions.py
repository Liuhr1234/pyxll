# numpy_functions.py v5.4 - 修复 SUBTOTAL 嵌套导致的重复计算问题
"""
纯向量化蒙特卡洛模拟宏（DriskNumpyMC）
"""

import os
import re
import math
import time
import traceback
import zlib
import datetime
from typing import Dict, List, Tuple, Optional, Any, Set, Union
import numpy as np

# PyXLL 相关
from pyxll import xl_macro, xl_app, xlcAlert

# 现有系统模块
from constants import DISTRIBUTION_FUNCTION_NAMES, get_distribution_info
from attribute_functions import extract_markers_from_args, is_marker_string, get_static_mode, set_static_mode
from corrmat_functions import _get_named_range, _is_positive_semidefinite, _nearest_psd, _read_matrix_from_range
from simulation_manager import clear_simulations, create_simulation, get_simulation
from sampling_functions import generate_latin_hypercube_1d
from formula_parser import (
    extract_nested_distributions_advanced,
    parse_formula_references,
    is_distribution_function,
    is_simtable_function,
    is_makeinput_function,
    is_output_cell,
    extract_all_distribution_functions_with_index,
    extract_simtable_functions,
    extract_makeinput_functions,
    extract_output_info,
    parse_args_with_nested_functions,
    extract_makeinput_attributes,
    ATTRIBUTE_FUNCTIONS  # 导入属性函数列表
)
from dependency_tracker import (
    find_all_simulation_cells_in_workbook,
    EnhancedDependencyAnalyzer,
    get_all_dependents_direct,
    is_formula_cell
)

# 从 index_functions 导入可重用的辅助类、函数和管理器获取函数
from index_functions import (
    TruncatableDistributionGenerator,
    StatusBarProgress,
    _parse_truncate_from_func,
    normalize_cell_address,
    _extract_cross_sheet_references,
    _count_distributions_in_cell,
    get_index_simulation_manager,
    bernoulli_generator_vectorized,
    triang_generator_vectorized,
    binomial_generator_vectorized,
    invgauss_generator_vectorized,
    duniform_generator_vectorized,
    trigen_generator_vectorized,
    cumul_generator_vectorized,
    discrete_generator_vectorized
)

from dist_bernoulli import bernoulli_generator_vectorized
from dist_triang import triang_generator_vectorized
from dist_binomial import binomial_generator_vectorized
from dist_erf import erf_cdf, erf_cdf, erf_generator_vectorized, erf_ppf
from dist_extvalue import extvalue_generator_vectorized, extvalue_ppf, extvalue_cdf, extvalue_raw_mean
from dist_extvaluemin import extvaluemin_generator_vectorized, extvaluemin_ppf, extvaluemin_cdf, extvaluemin_raw_mean
from dist_fatiguelife import fatiguelife_generator_vectorized, fatiguelife_ppf, fatiguelife_cdf, fatiguelife_raw_mean
from dist_frechet import frechet_generator_vectorized, frechet_ppf, frechet_cdf, frechet_raw_mean
from dist_general import general_generator_vectorized, general_ppf, general_cdf, general_raw_mean, _parse_general_arrays as general_parse_arrays
from dist_histogrm import histogrm_generator_vectorized, histogrm_ppf, histogrm_cdf, histogrm_raw_mean, _parse_histogrm_p_table as histogrm_parse_p_table
from dist_hypsecant import hypsecant_generator_vectorized, hypsecant_ppf, hypsecant_cdf, hypsecant_raw_mean
from dist_johnsonsb import johnsonsb_generator_vectorized, johnsonsb_ppf, johnsonsb_cdf, johnsonsb_raw_mean
from dist_johnsonsu import johnsonsu_generator_vectorized, johnsonsu_ppf, johnsonsu_cdf, johnsonsu_raw_mean
from dist_kumaraswamy import kumaraswamy_generator_vectorized, kumaraswamy_ppf, kumaraswamy_cdf, kumaraswamy_raw_mean
from dist_laplace import laplace_generator_vectorized, laplace_ppf, laplace_cdf, laplace_raw_mean
from dist_logistic import logistic_generator_vectorized, logistic_ppf, logistic_cdf, logistic_raw_mean
from dist_loglogistic import loglogistic_generator_vectorized, loglogistic_ppf, loglogistic_cdf, loglogistic_raw_mean
from dist_lognorm import lognorm_generator_vectorized, lognorm_ppf, lognorm_cdf, lognorm_raw_mean
from dist_lognorm2 import lognorm2_generator_vectorized, lognorm2_ppf, lognorm2_cdf, lognorm2_raw_mean
from dist_betageneral import betageneral_generator_vectorized, betageneral_ppf, betageneral_cdf, betageneral_raw_mean
from dist_betasubj import betasubj_generator_vectorized, betasubj_ppf, betasubj_cdf, betasubj_raw_mean
from dist_burr12 import burr12_generator_vectorized, burr12_ppf, burr12_cdf, burr12_raw_mean
from dist_compound import compound_generator_vectorized, compound_ppf, compound_cdf, compound_raw_mean
from dist_splice import splice_generator_vectorized, splice_ppf, splice_cdf, splice_raw_mean
from dist_pert import pert_generator_vectorized, pert_ppf, pert_cdf, pert_raw_mean
from dist_reciprocal import reciprocal_generator_vectorized, reciprocal_ppf, reciprocal_cdf, reciprocal_raw_mean
from dist_rayleigh import rayleigh_generator_vectorized, rayleigh_ppf, rayleigh_cdf, rayleigh_raw_mean
from dist_weibull import weibull_generator_vectorized, weibull_ppf, weibull_cdf, weibull_raw_mean
from dist_pearson5 import pearson5_cdf, pearson5_generator_vectorized, pearson5_ppf, pearson5_raw_mean
from dist_pearson6 import pearson6_cdf, pearson6_generator_vectorized, pearson6_ppf, pearson6_raw_mean
from dist_pareto2 import pareto2_cdf, pareto2_cdf, pareto2_generator_vectorized, pareto2_ppf, pareto2_raw_mean
from dist_pareto import pareto_generator_vectorized, pareto_ppf, pareto_cdf, pareto_raw_mean
from dist_levy import levy_generator_vectorized, levy_ppf, levy_cdf, levy_raw_mean
from dist_erlang import erlang_cdf, erlang_generator_vectorized, erlang_ppf
from dist_cauchy import cauchy_cdf, cauchy_generator_vectorized, cauchy_ppf
from dist_dagum import dagum_generator_vectorized, dagum_ppf, dagum_cdf, dagum_raw_mean
from dist_doubletriang import doubletriang_cdf, doubletriang_cdf, doubletriang_generator_vectorized, doubletriang_ppf, doubletriang_raw_mean
from dist_negbin import negbin_cdf, negbin_generator_vectorized, negbin_ppf
from dist_invgauss import invgauss_generator_vectorized, invgauss_ppf
from dist_duniform import duniform_generator_vectorized, duniform_ppf
from dist_geomet import geomet_cdf, geomet_cdf, geomet_generator_vectorized, geomet_ppf
from dist_hypergeo import hypergeo_cdf, hypergeo_generator_vectorized, hypergeo_ppf
from dist_intuniform import intuniform_cdf, intuniform_generator_vectorized, intuniform_ppf
from dist_trigen import trigen_generator_vectorized, _convert_trigen_to_triang
from dist_cumul import cumul_generator_vectorized, _parse_arrays as cumul_parse_arrays
from dist_discrete import discrete_generator_vectorized
from dist_triang import triang_truncated_mean
from dist_cumul import cumul_truncated_mean
from dist_discrete import discrete_truncated_mean
# 导入 PPF 函数（用于后备截断处理）
from dist_bernoulli import bernoulli_ppf
from dist_triang import triang_ppf
from dist_binomial import binomial_ppf
from dist_invgauss import invgauss_ppf
from dist_duniform import duniform_ppf
from dist_trigen import trigen_ppf
from dist_cumul import cumul_ppf
from dist_discrete import discrete_ppf
# 导入所有分布类（用于统一截断均值计算）
from distribution_base import (
    GammaDistribution, BetaDistribution, ChiSquaredDistribution,
    FDistribution, TDistribution, PoissonDistribution,
    ExponentialDistribution, UniformDistribution, NormalDistribution
)
from dist_bernoulli import BernoulliDistribution
from dist_binomial import BinomialDistribution
from dist_erf import ErfDistribution
from dist_extvalue import ExtvalueDistribution
from dist_extvaluemin import ExtvalueMinDistribution
from dist_fatiguelife import FatigueLifeDistribution
from dist_frechet import FrechetDistribution
from dist_general import GeneralDistribution
from dist_histogrm import HistogrmDistribution
from dist_hypsecant import HypSecantDistribution
from dist_johnsonsb import JohnsonSBDistribution
from dist_johnsonsu import JohnsonSUDistribution
from dist_kumaraswamy import KumaraswamyDistribution
from dist_laplace import LaplaceDistribution
from dist_logistic import LogisticDistribution
from dist_loglogistic import LoglogisticDistribution
from dist_lognorm import LognormDistribution
from dist_lognorm2 import Lognorm2Distribution
from dist_betageneral import BetaGeneralDistribution
from dist_betasubj import BetaSubjDistribution
from dist_burr12 import Burr12Distribution
from dist_compound import CompoundDistribution
from dist_splice import SpliceDistribution
from dist_pert import PertDistribution
from dist_reciprocal import ReciprocalDistribution
from dist_rayleigh import RayleighDistribution
from dist_weibull import WeibullDistribution
from dist_pearson5 import Pearson5Distribution
from dist_pearson6 import Pearson6Distribution
from dist_pareto2 import Pareto2Distribution
from dist_pareto import ParetoDistribution
from dist_levy import LevyDistribution
from dist_erlang import ErlangDistribution
from dist_cauchy import CauchyDistribution
from dist_dagum import DagumDistribution
from dist_doubletriang import DoubleTriangDistribution
from dist_negbin import NegbinDistribution
from dist_invgauss import InvgaussDistribution
from dist_duniform import DUniformDistribution
from dist_geomet import GeometDistribution
from dist_hypergeo import HypergeoDistribution
from dist_intuniform import IntuniformDistribution
from dist_triang import TriangDistribution
from dist_trigen import TrigenDistribution
from dist_cumul import CumulDistribution
from dist_discrete import DiscreteDistribution

# 导入索引模拟函数（用于回退）
from index_functions import run_index_simulation

# 导入内置函数解析器
from resolve_functions import _get_xl_ns, _subtotal_core

# ==================== 全局常量与缓存 ====================
ERROR_MARKER = "#ERROR!"
_EXCEL_ERROR_STRINGS = {'#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#N/A', '#NULL!'}
DEBUG = True  # 设置为 True 可开启调试输出（正式版设为 False）

# 公式解析缓存
_formula_parse_cache: Dict[str, List[Tuple]] = {}

# ==================== 工作表名规范化（新增） ====================
# 定义中文标点正则（全角逗号、句号、问号、分号、冒号、引号、感叹号、省略号等）
CHINESE_PUNCTUATION = r'[\u3000-\u303F\uFF00-\uFFEF]'

def normalize_sheet_name(name: str, strict: bool = True) -> str:
    """
    规范化工作表名：
    - 去除首尾空格
    - 将双写单引号还原为单引号
    - 转换为大写（Excel 不区分大小写）
    - 如果 strict=True 且包含全角标点，则抛出 ValueError
    """
    if not isinstance(name, str):
        return name
    # 去除首尾空格
    name = name.strip()
    # 还原双写单引号为单引号
    name = name.replace("''", "'")
    # 检查中文标点
    if strict and re.search(CHINESE_PUNCTUATION, name):
        raise ValueError(f"工作表名包含中文标点，不允许：{name}")
    # 转换为大写
    return name.upper()

def normalize_cell_key(key: str) -> str:
    """
    规范化单元格键（格式：工作表名!单元格地址）
    返回规范化后的键，工作表名使用 normalize_sheet_name 处理。
    """
    if '!' in key:
        sheet_part, cell_part = key.split('!', 1)
        try:
            norm_sheet = normalize_sheet_name(sheet_part, strict=True)
        except ValueError as e:
            # 如果检测到中文标点，重新抛出以便上层捕获
            raise e
        # 单元格地址保持原样（但可以转为大写，以便后续匹配）
        cell_part_norm = cell_part.replace('$', '').upper()
        return f"{norm_sheet}!{cell_part_norm}"
    else:
        # 没有工作表名的键（一般不会出现），直接返回原样
        return key.upper()

# ==================== 动态构建分布函数正则表达式 ====================
# 使用 DISTRIBUTION_FUNCTION_NAMES 构建模式，支持所有已注册的 Drisk 分布函数
_DRISK_PATTERN = re.compile(
    r'(?:_xll\.)?(' + '|'.join(re.escape(name) for name in DISTRIBUTION_FUNCTION_NAMES) + r')\s*\(',
    re.I
)

# ==================== 增强的单元格引用正则（支持带引号工作表名，含转义单引号）====================
# 匹配两种形式：
#   1. 带引号的工作表名：'...'! 内部允许双写单引号
#   2. 不带引号的工作表名：字母开头，仅包含字母数字下划线
# 捕获组：
#   组1：引号内的内容（如果存在）
#   组2：无引号的工作表名（如果存在）
#   组3：列字母（可能带$）
#   组4：行号（可能带$）
_CELL_REF_PATTERN = re.compile(
    r"(?<![0-9.])(?:(?:'((?:[^']|'')*)'|([A-Za-z][A-Za-z0-9_]*))!)?\$?([A-Z]+)\$?(\d+)(?!\s*\(|[A-Za-z])",
    re.I
)

_RANGE_PAT = re.compile(
    r"^(?:(?:'[^']+'|[A-Za-z0-9_]+)!)?\$?([A-Z]+\d+):\$?([A-Z]+\d+)$", re.I)
_FUNC_PATTERN        = re.compile(r'([A-Za-z_][A-Za-z0-9_.]*)\s*\(')
_EMPTY_STRING_PATTERN = re.compile(r'^(""|\'\')$')
_OUTPUT_PATTERN      = re.compile(r'(?:_xll\.)?DriskOutput\s*\(', re.I)
_DRISK_NAME_RE       = re.compile(r'^Drisk(?!Output\b)', re.I)

# ==================== 辅助函数：获取工作簿 ====================
def _get_workbook_path() -> Optional[str]:
    """获取当前工作簿路径（用于兼容，实际不使用文件）"""
    try:
        import xlwings as xw
        path = xw.Book.caller().fullname
        if path and os.path.exists(path):
            return path
    except:
        pass
    return None

def _get_active_workbook_com():
    """获取当前活动工作簿的 COM 对象"""
    try:
        import win32com.client as win32
        app = win32.GetActiveObject("Excel.Application")
        wb = app.ActiveWorkbook
        return wb, wb.Name
    except Exception as e:
        print(f"  [COM] 获取工作簿失败: {e}")
        return None, None

def _is_excel_error(val) -> bool:
    """判断是否为 Excel 错误值"""
    if val is None:
        return False
    if isinstance(val, str) and val.strip().upper() in _EXCEL_ERROR_STRINGS:
        return True
    if isinstance(val, (int, float)):
        # COM 错误代码通常为负整数
        return int(val) in {-2146826281, -2146826246, -2146826259, -2146826288,
                            -2146826265, -2146826252, -2146826273}
    return False

def _n2col(n: int) -> str:
    """将列号转换为字母"""
    s = ''
    while n > 0:
        n -= 1
        s = chr(n % 26 + 65) + s
        n //= 26
    return s

def _strip_xll_prefix(formula: str) -> str:
    """移除公式中的 _xll. 前缀"""
    return re.sub(r'_xll\.', '', formula, flags=re.I)

def _normalize_key(sheet: str, coord: str) -> str:
    """标准化键名：Sheet!A1 转为大写"""
    return f"{sheet}!{coord}".upper()

def _col2n(c: str) -> int:
    """列字母转数字（忽略 $ 符号）"""
    c = c.replace('$', '')
    n = 0
    for ch in c.upper():
        n = n * 26 + ord(ch) - 64
    return n

def _convert_excel_operators(expr: str) -> str:
    """
    将 Excel 公式中的运算符转换为 Python 可求值的形式。
    处理：^ -> **, % -> /100, = -> ==, <> -> !=, & -> +
    同时保护字符串字面量不被替换。
    """
    # 1. 找出所有字符串字面量（双引号或单引号括起来的内容）
    string_pattern = r'("[^"\\]*(?:\\.[^"\\]*)*"|\'[^\'\\]*(?:\\.[^\'\\]*)*\')'
    strings = []
    placeholders = []

    def string_repl(m):
        s = m.group(0)
        placeholder = f"__STR_{len(strings)}__"
        strings.append(s)
        placeholders.append(placeholder)
        return placeholder

    # 临时替换字符串为占位符
    expr_no_strings = re.sub(string_pattern, string_repl, expr)

    # 2. 运算符替换（注意使用单词边界或前后断言避免误替换变量名）
    # 先处理 <> 为 !=
    expr_no_strings = re.sub(r'<>', '!=', expr_no_strings)

    # 将独立的 = 替换为 == (前后不能是 < > = !)
    expr_no_strings = re.sub(r'(?<![<>=!])=(?![=])', '==', expr_no_strings)

    # 将 ^ 替换为 **
    expr_no_strings = expr_no_strings.replace('^', '**')

    # 将 & 替换为 + (前后不是字母数字或下划线)
    expr_no_strings = expr_no_strings.replace('&', '+')

    # 处理百分比：匹配数字后跟 %，替换为 /100.0
    # 匹配整数、小数（如 5, 5.5, .5）
    expr_no_strings = re.sub(r'(\d+(?:\.\d*)?|\.\d+)%', r'(\1)/100.0', expr_no_strings)

    # 3. 恢复字符串字面量
    for placeholder, s in zip(placeholders, strings):
        expr_no_strings = expr_no_strings.replace(placeholder, s)

    return expr_no_strings

# ==================== 新增：安全求值数组元素 ====================
def _safe_eval_array_element(expr: str) -> float:
    """
    安全地求值数组常量中的简单数学表达式，仅支持数字、运算符和内置数学函数。
    如果表达式无法求值，抛出 ValueError。
    """
    # 允许的命名空间
    safe_ns = {
        'abs': abs,
        'sqrt': math.sqrt,
        'exp': math.exp,
        'log': math.log,
        'log10': math.log10,
        'sin': math.sin,
        'cos': math.cos,
        'tan': math.tan,
        'pi': math.pi,
        'e': math.e,
        '__builtins__': {}
    }
    # 先将 ^ 转换为 **（Excel 风格）
    expr = expr.replace('^', '**')
    try:
        result = eval(expr, {"__builtins__": {}}, safe_ns)
        return float(result)
    except Exception:
        raise ValueError(f"无法求值数组元素: {expr}")

# ==================== 辅助函数：解析数组参数（从 index_functions 复制并调整） ====================
def _split_cumul_discrete_args(args_text: str) -> List[str]:
    """自定义参数分割函数，正确处理花括号数组常量"""
    args = []
    current = []
    brace_depth = 0
    in_quotes = False
    quote_char = None
    for ch in args_text:
        if ch in ('"', "'"):
            if not in_quotes:
                in_quotes = True
                quote_char = ch
            elif ch == quote_char:
                in_quotes = False
            current.append(ch)
        elif not in_quotes:
            if ch == '{':
                brace_depth += 1
                current.append(ch)
            elif ch == '}':
                brace_depth -= 1
                current.append(ch)
            elif ch == ',' and brace_depth == 0:
                arg = ''.join(current).strip()
                if arg:
                    args.append(arg)
                current = []
            else:
                current.append(ch)
        else:
            current.append(ch)
    if current:
        arg = ''.join(current).strip()
        if arg:
            args.append(arg)
    return args

def _read_range_values(app, range_str: str) -> List[float]:
    """从 Excel 区域读取数值，返回浮点数列表"""
    try:
        if '!' in range_str:
            sheet_name, addr = range_str.split('!', 1)
            sheet = app.ActiveWorkbook.Worksheets(sheet_name)
        else:
            sheet = app.ActiveSheet
            addr = range_str
        addr = addr.strip()
        if ':' not in addr:
            cell = sheet.Range(addr)
            val = cell.Value2
            try:
                return [float(val)]
            except (TypeError, ValueError):
                raise ValueError(f"单元格 {range_str} 的值不是数字")
        else:
            range_obj = sheet.Range(addr)
            values_2d = range_obj.Value2
            result = []
            if isinstance(values_2d, tuple):
                for row in values_2d:
                    if isinstance(row, tuple):
                        for val in row:
                            try:
                                result.append(float(val))
                            except (TypeError, ValueError):
                                raise ValueError(f"区域 {range_str} 中包含非数字值")
                    else:
                        try:
                            result.append(float(row))
                        except (TypeError, ValueError):
                            raise ValueError(f"区域 {range_str} 中包含非数字值")
            else:
                try:
                    result.append(float(values_2d))
                except (TypeError, ValueError):
                    raise ValueError(f"区域 {range_str} 的值不是数字")
            return result
    except Exception as e:
        print(f"读取区域 {range_str} 失败: {str(e)}")
        raise

def _parse_cumul_4args(args_text: str, app) -> Tuple[Optional[str], Optional[str]]:
    """从 DriskCumul 的参数字符串中解析 x_vals 和 p_vals 字符串（已包含边界）"""
    if not args_text:
        return None, None
    args_list = _split_cumul_discrete_args(args_text)
    if len(args_list) < 4:
        return None, None
    min_raw = args_list[0].strip().strip('"\'')
    max_raw = args_list[1].strip().strip('"\'')
    x_raw = args_list[2].strip().strip('"\'')
    p_raw = args_list[3].strip().strip('"\'')
    # 解析 min 和 max
    try:
        min_val = float(_read_range_values(app, min_raw)[0] if re.match(r'[A-Z]+\d+', min_raw, re.I) else float(min_raw))
        max_val = float(_read_range_values(app, max_raw)[0] if re.match(r'[A-Z]+\d+', max_raw, re.I) else float(max_raw))
    except Exception as e:
        print(f"解析 min/max 失败: {e}")
        return None, None
    # 解析内部 X 数组
    if re.match(r'^[A-Z]{1,3}\d', x_raw, re.I) or ('!' in x_raw and re.match(r'.+![A-Z]{1,3}\d', x_raw, re.I)):
        try:
            x_inner = _read_range_values(app, x_raw)
        except Exception as e:
            print(f"读取 X 区域失败: {e}")
            return None, None
    else:
        if x_raw.startswith('{') and x_raw.endswith('}'):
            x_raw = x_raw[1:-1]
        parts = [p.strip() for p in x_raw.split(',') if p.strip()]
        x_inner = []
        for p in parts:
            try:
                x_inner.append(float(p))
            except ValueError:
                try:
                    x_inner.append(_safe_eval_array_element(p))
                except ValueError as e:
                    print(f"X 数组元素解析失败: {e}")
                    return None, None
    # 解析内部 P 数组
    if re.match(r'^[A-Z]{1,3}\d', p_raw, re.I) or ('!' in p_raw and re.match(r'.+![A-Z]{1,3}\d', p_raw, re.I)):
        try:
            p_inner = _read_range_values(app, p_raw)
        except Exception as e:
            print(f"读取 P 区域失败: {e}")
            return None, None
    else:
        if p_raw.startswith('{') and p_raw.endswith('}'):
            p_raw = p_raw[1:-1]
        parts = [p.strip() for p in p_raw.split(',') if p.strip()]
        p_inner = []
        for p in parts:
            try:
                p_inner.append(float(p))
            except ValueError:
                try:
                    p_inner.append(_safe_eval_array_element(p))
                except ValueError as e:
                    print(f"P 数组元素解析失败: {e}")
                    return None, None
    if len(x_inner) != len(p_inner):
        print(f"X 和 P 数组长度不相等: {len(x_inner)} vs {len(p_inner)}")
        return None, None
    # 验证内部点范围
    for x in x_inner:
        if not (min_val <= x <= max_val):
            print(f"X 值 {x} 超出范围 [{min_val}, {max_val}]")
            return None, None
    for p in p_inner:
        if not (0 <= p <= 1):
            print(f"P 值 {p} 超出 [0,1]")
            return None, None
    # 构造完整数组
    x_full = [min_val] + x_inner + [max_val]
    p_full = [0.0] + p_inner + [1.0]
    # 验证完整数组有效性（可选）
    try:
        from dist_cumul import _parse_arrays as cumul_parse_arrays
        cumul_parse_arrays(x_full, p_full)
    except Exception as e:
        print(f"Cumul 数组验证失败: {e}")
        return None, None
    x_full_str = ','.join(str(x) for x in x_full)
    p_full_str = ','.join(str(p) for p in p_full)
    return x_full_str, p_full_str

def _parse_cumul_discrete_args(args_text: str, app) -> Tuple[Optional[str], Optional[str]]:
    """从 DriskDiscrete 的参数字符串中解析 x_vals 和 p_vals 字符串"""
    if not args_text:
        return None, None
    args_list = _split_cumul_discrete_args(args_text)
    if len(args_list) < 2:
        return None, None
    x_raw = args_list[0].strip().strip('"\'')
    p_raw = args_list[1].strip().strip('"\'')
    # 解析 X 数组
    if re.match(r'^[A-Z]{1,3}\d', x_raw, re.I) or ('!' in x_raw and re.match(r'.+![A-Z]{1,3}\d', x_raw, re.I)):
        try:
            x_vals = _read_range_values(app, x_raw)
        except Exception as e:
            print(f"读取 X 区域失败: {e}")
            return None, None
    else:
        if x_raw.startswith('{') and x_raw.endswith('}'):
            x_raw = x_raw[1:-1]
        parts = [p.strip() for p in x_raw.split(',') if p.strip()]
        x_vals = []
        for p in parts:
            try:
                x_vals.append(float(p))
            except ValueError:
                try:
                    x_vals.append(_safe_eval_array_element(p))
                except ValueError as e:
                    print(f"X 数组元素解析失败: {e}")
                    return None, None
    # 解析 P 数组
    if re.match(r'^[A-Z]{1,3}\d', p_raw, re.I) or ('!' in p_raw and re.match(r'.+![A-Z]{1,3}\d', p_raw, re.I)):
        try:
            p_vals = _read_range_values(app, p_raw)
        except Exception as e:
            print(f"读取 P 区域失败: {e}")
            return None, None
    else:
        if p_raw.startswith('{') and p_raw.endswith('}'):
            p_raw = p_raw[1:-1]
        parts = [p.strip() for p in p_raw.split(',') if p.strip()]
        p_vals = []
        for p in parts:
            try:
                p_vals.append(float(p))
            except ValueError:
                try:
                    p_vals.append(_safe_eval_array_element(p))
                except ValueError as e:
                    print(f"P 数组元素解析失败: {e}")
                    return None, None
    if len(x_vals) != len(p_vals):
        print(f"X 和 P 数组长度不相等: {len(x_vals)} vs {len(p_vals)}")
        return None, None
    # 归一化概率
    total = sum(p_vals)
    if total > 0:
        p_vals = [p / total for p in p_vals]
    x_str = ','.join(str(x) for x in x_vals)
    p_str = ','.join(str(p) for p in p_vals)
    return x_str, p_str

def _parse_duniform_args(args_text: str, app) -> Tuple[Optional[str], Optional[str]]:
    if not args_text:
        print("DUniform: 参数文本为空")
        return None, None
    args_list = _split_cumul_discrete_args(args_text)
    if len(args_list) < 1:
        print(f"DUniform: 分割参数失败，args_text='{args_text}', args_list={args_list}")
        return None, None
    x_raw = args_list[0].strip().strip('"\'')
    print(f"DUniform: 原始 X 参数 = '{x_raw}'")
    
    # 解析 X 值列表
    if re.match(r'^[A-Z]{1,3}\d', x_raw, re.I) or ('!' in x_raw and re.match(r'.+![A-Z]{1,3}\d', x_raw, re.I)):
        try:
            x_vals = _read_range_values(app, x_raw)
            print(f"DUniform: 从区域读取 X 值成功: {x_vals}")
        except Exception as e:
            print(f"DUniform: 读取 X 区域失败: {e}")
            return None, None
    else:
        # 处理花括号或普通逗号分隔字符串
        if x_raw.startswith('{') and x_raw.endswith('}'):
            x_raw = x_raw[1:-1]
        parts = [p.strip() for p in x_raw.split(',') if p.strip()]
        if not parts:
            print(f"DUniform: 分割后无有效元素，x_raw='{x_raw}'")
            return None, None
        x_vals = []
        for p in parts:
            try:
                x_vals.append(float(p))
            except ValueError:
                try:
                    x_vals.append(_safe_eval_array_element(p))
                except ValueError as e:
                    print(f"DUniform: X 数组元素解析失败: {e}, 元素='{p}'")
                    return None, None
    if not x_vals:
        print("DUniform: X 值列表为空")
        return None, None
    if len(x_vals) != len(set(x_vals)):
        print(f"DUniform: X 值必须唯一，当前值: {x_vals}")
        return None, None
    
    # 生成等概率
    p = 1.0 / len(x_vals)
    p_vals = [p] * len(x_vals)
    x_str = ','.join(str(x) for x in x_vals)
    p_str = ','.join(str(p) for p in p_vals)
    print(f"DUniform: 解析成功，x_str='{x_str}', p_str='{p_str}'")
    return x_str, p_str

def _preprocess_distribution_cells(distribution_cells: Dict[str, List[Dict]], app):
    """预处理分布单元格，解析数组参数并存入 markers"""
    for cell_addr, funcs in distribution_cells.items():
        for func in funcs:
            func_name = func.get('func_name', '')
            args_text = func.get('args_text', '')
            if func_name == 'DriskCumul':
                x_str, p_str = _parse_cumul_4args(args_text, app)
                if x_str and p_str:
                    if 'markers' not in func:
                        func['markers'] = {}
                    func['markers']['x_vals'] = x_str
                    func['markers']['p_vals'] = p_str
                    if DEBUG:
                        print(f"为 {func.get('input_key', cell_addr)} 补充 Cumul 数组参数: x_vals={x_str}, p_vals={p_str}")
            elif func_name == 'DriskDiscrete':
                x_str, p_str = _parse_cumul_discrete_args(args_text, app)
                if x_str and p_str:
                    if 'markers' not in func:
                        func['markers'] = {}
                    func['markers']['x_vals'] = x_str
                    func['markers']['p_vals'] = p_str
                    if DEBUG:
                        print(f"为 {func.get('input_key', cell_addr)} 补充 Discrete 数组参数: x_vals={x_str}, p_vals={p_str}")
            elif func_name == 'DriskDUniform':
                x_str, p_str = _parse_duniform_args(args_text, app)
                if x_str and p_str:
                    if 'markers' not in func:
                        func['markers'] = {}
                    func['markers']['x_vals'] = x_str
                    func['markers']['p_vals'] = p_str
                    if DEBUG:
                        print(f"为 {func.get('input_key', cell_addr)} 补充 DUniform 数组参数: x_vals={x_str}, p_vals={p_str}")

# ==================== 新增：获取命名区域映射 ====================
def _get_names_map(wb_com) -> Dict[str, str]:
    """
    返回字典：{名称（大写）: 规范化的完整地址（如 'RISK FACTORS!F3'）}
    """
    names_map = {}
    try:
        for name_obj in wb_com.Names:
            name = name_obj.Name
            # 跳过内部名称（如 _xlfn. 开头）
            if name.startswith('_xlfn'):
                continue
            refers_to = name_obj.RefersTo
            # 解析引用字符串，格式如 '=Risk factors!$F$3' 或 '=Risk factors!$F$3:$G$10'
            if not isinstance(refers_to, str):
                continue
            # 去掉开头的 '='
            if refers_to.startswith('='):
                refers_to = refers_to[1:]
            # 解析工作表名和地址
            if '!' in refers_to:
                sheet_part, addr = refers_to.split('!', 1)
                # 去除可能的外层引号
                sheet_part = sheet_part.strip()
                if sheet_part.startswith("'") and sheet_part.endswith("'"):
                    sheet_part = sheet_part[1:-1].replace("''", "'")
                # 标准化工作表名
                try:
                    sheet_norm = normalize_sheet_name(sheet_part, strict=True)
                except ValueError:
                    continue
                # 将地址规范化（去除$，转为大写）
                addr = addr.replace('$', '').upper()
                full_addr = f"{sheet_norm}!{addr}"
                names_map[name.upper()] = full_addr
    except Exception as e:
        print(f"获取命名区域映射失败: {e}")
    return names_map

def _replace_names_in_formula(formula: str, names_map: Dict[str, str]) -> str:
    """
    将公式中的命名区域替换为实际地址（使用单词边界）。
    如果工作表名包含空格或特殊字符，自动添加单引号。
    """
    if not names_map:
        return formula
    new_formula = formula
    # 按名称长度降序排序，避免短名称替换时影响长名称
    for name in sorted(names_map.keys(), key=len, reverse=True):
        addr = names_map[name]
        # 解析工作表名和单元格地址
        if '!' in addr:
            sheet_part, cell_part = addr.split('!', 1)
            # 如果工作表名包含空格或特殊字符，需要加单引号
            if ' ' in sheet_part or any(c in sheet_part for c in "[]:?*"):
                safe_sheet = f"'{sheet_part}'"
            else:
                safe_sheet = sheet_part
            safe_addr = f"{safe_sheet}!{cell_part}"
        else:
            safe_addr = addr
        pattern = r'\b' + re.escape(name) + r'\b'
        def repl(match):
            return safe_addr
        new_formula = re.sub(pattern, repl, new_formula, flags=re.IGNORECASE)
    return new_formula

# ==================== COM 扫描函数 ====================
def _scan_workbook_com(wb_com) -> Tuple[Dict[str, str], Dict[str, Any], List[str]]:
    """
    通过 COM 对象直接从 Excel 内存读取公式和值，错误单元格标记为 #ERROR!。
    返回: (all_formulas, static_values, output_keys)
    """
    all_formulas: Dict[str, str] = {}
    static_values: Dict[str, Any] = {}
    output_keys: List[str] = []

    try:
        sheets = wb_com.Worksheets
        sheet_count = sheets.Count
    except Exception as e:
        print(f"  [COM] 获取工作表失败: {e}")
        return all_formulas, static_values, output_keys

    for si in range(1, sheet_count + 1):
        try:
            ws = sheets(si)
            sn = ws.Name
        except Exception:
            continue

        try:
            used = ws.UsedRange
            if used is None:
                continue
            row0 = used.Row
            col0 = used.Column
            rows = used.Rows.Count
            cols = used.Columns.Count
        except Exception:
            continue

        if rows == 0 or cols == 0:
            continue

        try:
            values_block = used.Value2
            formula_block = used.Formula
        except Exception as e:
            print(f"  [COM] 读取区域失败 sheet={sn}: {e}")
            continue

        if not isinstance(values_block, (list, tuple)):
            values_block = [[values_block]]
            formula_block = [[formula_block]]
        else:
            values_block = [list(r) for r in values_block]
            formula_block = [list(r) for r in formula_block]

        for ri, (vrow, frow) in enumerate(zip(values_block, formula_block)):
            for ci, (val, fml) in enumerate(zip(vrow, frow)):
                abs_row = row0 + ri
                abs_col = col0 + ci
                coord = f"{_n2col(abs_col)}{abs_row}"
                key = _normalize_key(sn, coord)   # 使用规范化键

                # 读取值
                if val is not None:
                    if _is_excel_error(val):
                        static_values[key] = ERROR_MARKER
                        static_values[coord.upper()] = ERROR_MARKER
                        if DEBUG:
                            print(f"  [WARN] {key} 含 Excel 错误值 ({val})，标记为 #ERROR!")
                    else:
                        try:
                            # 尝试转为浮点数，若失败则保留原值（如字符串）
                            fval = float(val)
                            static_values[key] = fval
                            static_values[coord.upper()] = fval
                        except (TypeError, ValueError):
                            # 非数字类型（如字符串、日期等），保留原值
                            static_values[key] = val
                            static_values[coord.upper()] = val

                # 读取公式
                if fml and isinstance(fml, str) and fml.startswith('='):
                    clean = _strip_xll_prefix(fml)
                    all_formulas[key] = clean
                    if is_output_cell(clean):
                        output_keys.append(key)

    return all_formulas, static_values, output_keys

# ==================== 向量化求值引擎 ====================
def _get_tokenized_expr(expr: str) -> List[Tuple]:
    """获取表达式的 token 列表，带缓存"""
    if expr not in _formula_parse_cache:
        _formula_parse_cache[expr] = _tokenize_top(expr)
    return _formula_parse_cache[expr]

def _tokenize_top(expr: str) -> List[Tuple]:
    """将表达式拆分为原子和函数调用"""
    tokens = []
    i = 0
    text = expr
    while i < len(text):
        m = _FUNC_PATTERN.search(text, i)
        if not m:
            tail = text[i:].strip()
            if tail: tokens.append(('atom', tail))
            break
        if m.start() > i:
            piece = text[i:m.start()].strip()
            if piece: tokens.append(('atom', piece))
        func_name = m.group(1)
        depth = 0
        j = m.end() - 1
        while j < len(text):
            if text[j] == '(':   depth += 1
            elif text[j] == ')': depth -= 1
            if depth == 0:
                inner = text[m.end():j]
                tokens.append(('call', func_name, inner))
                i = j + 1
                break
            j += 1
        else:
            tokens.append(('atom', text[m.start():]))
            i = len(text)
    return tokens

def _parse_drisk_args(inner: str) -> List[str]:
    """拆分函数参数，支持嵌套括号"""
    args, depth, cur = [], 0, []
    for ch in inner:
        if ch == '(':   depth += 1; cur.append(ch)
        elif ch == ')': depth -= 1; cur.append(ch)
        elif ch == ',' and depth == 0:
            args.append(''.join(cur).strip()); cur = []
        else:
            cur.append(ch)
    if cur:
        args.append(''.join(cur).strip())
    return args

# ---------- 单元格引用解析 ----------
def _extract_cell_refs_from_text(text: str, context_node: str) -> List[str]:
    """
    提取文本中的单元格引用，返回带工作表名的完整地址（大写，规范化）
    处理带引号的工作表名，还原内部双引号为单引号，并规范化为统一格式。
    """
    refs = []
    for m in _CELL_REF_PATTERN.finditer(text):
        quoted_sheet = m.group(1)
        unquoted_sheet = m.group(2)
        col = m.group(3)
        row = m.group(4)
        if quoted_sheet is not None:
            # 引号内的工作表名：还原双写单引号为单引号
            sheet = quoted_sheet.replace("''", "'")
            # 去除可能存在的首尾空格（Excel中工作表名不能以空格开头，但可能存在）
            sheet = sheet.strip()
            # 规范化工作表名（转为大写，检查非法字符）
            try:
                sheet_norm = normalize_sheet_name(sheet, strict=True)
            except ValueError as e:
                # 如果包含非法字符，跳过该引用（不记录依赖）
                continue
            # 构建引用字符串（不带外层引号）
            ref = f"{sheet_norm}!{col}{row}"
        elif unquoted_sheet is not None:
            sheet = unquoted_sheet.strip()
            try:
                sheet_norm = normalize_sheet_name(sheet, strict=True)
            except ValueError:
                continue
            ref = f"{sheet_norm}!{col}{row}"
        else:
            # 无工作表名，使用 context_node 确定工作表
            if '!' in context_node:
                sheet = context_node.split('!')[0]
                # 直接使用已规范化的工作表名（context_node 应已规范化）
                sheet_norm = sheet
            else:
                # 如果 context_node 没有工作表名，取当前默认工作表（但实际不应发生）
                sheet_norm = "Sheet1"
            ref = f"{sheet_norm}!{col}{row}"
        full_ref = ref.upper()
        if full_ref not in refs:
            refs.append(full_ref)
    return refs

def _resolve_cell_ref_live(
    ref: str, n: int,
    static_values: Dict[str, Any],
    live_arrays: Dict[str, np.ndarray],
    default_sheet: str,
    simtable_values: Dict[str, float] = None
) -> np.ndarray:
    """
    解析单元格引用，返回长度为 n 的数组。
    使用规范化键进行查找，确保与依赖图一致。
    """
    # 解析引用，得到规范化键
    # 注意：ref 可能是 "Sheet1!A1" 或 "'Sheet1'!A1" 等形式，需要先规范化
    # 通过 _extract_cell_refs_from_text 提取，但这里只需一个引用，所以直接处理
    if '!' in ref:
        sheet_part, cell_part = ref.split('!', 1)
        # 去除可能的外层引号
        sheet_part = sheet_part.strip()
        if sheet_part.startswith("'") and sheet_part.endswith("'"):
            sheet_part = sheet_part[1:-1].replace("''", "'")
        try:
            sheet_clean = normalize_sheet_name(sheet_part, strict=True)
        except ValueError:
            return np.full(n, ERROR_MARKER, dtype=object)
        cell_clean = cell_part.replace('$', '').upper()
        full_key = f"{sheet_clean}!{cell_clean}"
    else:
        # 无工作表名
        try:
            sheet_clean = normalize_sheet_name(default_sheet, strict=True)
        except ValueError:
            return np.full(n, ERROR_MARKER, dtype=object)
        cell_clean = ref.replace('$', '').upper()
        full_key = f"{sheet_clean}!{cell_clean}"

    # 1. simtable_values
    if simtable_values:
        if full_key in simtable_values:
            val = simtable_values[full_key]
            if val == ERROR_MARKER:
                return np.full(n, ERROR_MARKER, dtype=object)
            try:
                return np.full(n, float(val), dtype=float)
            except:
                return np.full(n, np.nan, dtype=float)
        if cell_clean in simtable_values:
            val = simtable_values[cell_clean]
            if val == ERROR_MARKER:
                return np.full(n, ERROR_MARKER, dtype=object)
            try:
                return np.full(n, float(val), dtype=float)
            except:
                return np.full(n, np.nan, dtype=float)

    # 2. live_arrays
    if full_key in live_arrays:
        if DEBUG:
            print(f"    [resolve] {ref} -> live_arrays[{full_key}] 存在")
        return live_arrays[full_key]

    # 3. static_values
    val = static_values.get(full_key)
    if val is None:
        val = static_values.get(cell_clean)
    if val == ERROR_MARKER:
        return np.full(n, ERROR_MARKER, dtype=object)
    if val is not None:
        if isinstance(val, str) and val == ERROR_MARKER:
            return np.full(n, ERROR_MARKER, dtype=object)
        try:
            fval = float(val)
            if np.isnan(fval):
                if DEBUG:
                    print(f"    [resolve] {ref} 静态值为 NaN，返回 NaN 数组")
                return np.full(n, np.nan, dtype=float)
            if DEBUG:
                print(f"    [resolve] {ref} 静态值为 {fval}，返回全等数组")
            return np.full(n, fval, dtype=float)
        except (TypeError, ValueError):
            return np.full(n, np.nan, dtype=float)
    else:
        # 4. 实时读取
        try:
            from pyxll import xl_app
            app = xl_app()
            if '!' in ref:
                # 重新解析工作表名和单元格地址，因为 ref 可能包含引号
                sheet_part, cell_addr = ref.split('!', 1)
                sheet_part = sheet_part.strip()
                if sheet_part.startswith("'") and sheet_part.endswith("'"):
                    sheet_part = sheet_part[1:-1].replace("''", "'")
                try:
                    sheet = app.ActiveWorkbook.Worksheets(sheet_part)
                except:
                    sheet = app.ActiveSheet
            else:
                sheet = app.ActiveSheet
                cell_addr = ref
            cell = sheet.Range(cell_addr)
            cell_val = cell.Value2
            if cell_val is None:
                return np.full(n, np.nan, dtype=float)
            if _is_excel_error(cell_val):
                return np.full(n, ERROR_MARKER, dtype=object)
            try:
                fval = float(cell_val)
                static_values[full_key] = fval
                static_values[cell_clean] = fval
                if DEBUG:
                    print(f"    [resolve] {ref} 从 Excel 实时读取到 {fval}")
                return np.full(n, fval, dtype=float)
            except:
                return np.full(n, np.nan, dtype=float)
        except Exception as e:
            if DEBUG:
                print(f"    [resolve] 从 Excel 实时读取 {ref} 失败: {e}")
            return np.full(n, np.nan, dtype=float)

def _resolve_atoms_live(
    text: str, n: int,
    static_values: Dict[str, Any],
    live_arrays: Dict[str, np.ndarray],
    default_sheet: str,
    simtable_values: Dict[str, float] = None
) -> Tuple[str, Dict[str, np.ndarray]]:
    """将原子表达式中的单元格引用替换为变量名"""
    extra_vars: Dict[str, np.ndarray] = {}
    idx_counter = [0]

    def repl(m):
        # m 是匹配对象，组1：引号内工作表名，组2：无引号工作表名，组3：列，组4：行
        quoted_sheet = m.group(1)
        unquoted_sheet = m.group(2)
        col = m.group(3)
        row = m.group(4)
        # 重建原始引用字符串（包括可能的工作表名和 !）
        if quoted_sheet is not None:
            # 引号内可能包含双写单引号，还原为单引号
            sheet = quoted_sheet.replace("''", "'")
            ref = f"'{sheet}'!{col}{row}"
        elif unquoted_sheet is not None:
            ref = f"{unquoted_sheet}!{col}{row}"
        else:
            ref = f"{col}{row}"
        # 注意：如果原始引用中带有 $，在匹配时已经忽略了，所以重建的引用没有 $
        arr = _resolve_cell_ref_live(ref, n, static_values, live_arrays, default_sheet, simtable_values)
        if isinstance(arr, np.ndarray):
            if arr.dtype == object and np.all(arr == ERROR_MARKER):
                return ERROR_MARKER
            first = arr.flat[0]
            is_error = isinstance(first, str) and first == ERROR_MARKER
            is_nan = False
            if isinstance(first, (int, float, np.number)):
                is_nan = np.isnan(first)
            if not np.all(arr == first) or is_error or is_nan:
                vname = f'__cell{idx_counter[0]}__'
                idx_counter[0] += 1
                extra_vars[vname] = arr
                if DEBUG:
                    print(f"    [atom] 引用 {ref} 生成变量 {vname}，数组形状 {arr.shape}，第一元素 {first}")
                return vname
            else:
                if DEBUG:
                    print(f"    [atom] 引用 {ref} 为常量 {first}")
                return str(first)
        else:
            vname = f'__cell{idx_counter[0]}__'
            idx_counter[0] += 1
            extra_vars[vname] = arr
            return vname

    # 使用新的正则表达式进行替换
    result = _CELL_REF_PATTERN.sub(repl, text)
    return result, extra_vars

def _eval_range_live(
    range_str: str, n: int,
    static_values: Dict[str, Any],
    live_arrays: Dict[str, np.ndarray],
    default_sheet: str,
    simtable_values: Dict[str, float] = None
) -> np.ndarray:
    """
    评估区域引用（如 A1:B2），返回一个二维数组，形状 (num_cells, n)，
    其中 num_cells 是区域内的单元格数，n 是迭代次数。
    """
    # 解析工作表名和区域
    if '!' in range_str:
        sheet_part, range_part = range_str.split('!', 1)
        sheet_part = sheet_part.strip()
        if sheet_part.startswith("'") and sheet_part.endswith("'"):
            sheet_part = sheet_part[1:-1].replace("''", "'")
        sheet = sheet_part
    else:
        sheet = default_sheet
        range_part = range_str

    match = re.match(r'([A-Z]+)(\d+):([A-Z]+)(\d+)', range_part, re.I)
    if not match:
        return np.full((1, n), np.nan, dtype=float)
    start_col, start_row, end_col, end_row = match.groups()
    start_row = int(start_row)
    end_row = int(end_row)
    start_col_num = _col2n(start_col)
    end_col_num = _col2n(end_col)

    values = []
    for row in range(start_row, end_row+1):
        for col_num in range(start_col_num, end_col_num+1):
            col_letter = _n2col(col_num)
            coord = f"{col_letter}{row}"
            full_key = normalize_cell_key(f"{sheet}!{coord}")

            if full_key in live_arrays:
                arr = live_arrays[full_key]
                if arr.ndim != 1:
                    arr = arr.flatten()
                if len(arr) != n:
                    if len(arr) > n:
                        arr = arr[:n]
                    else:
                        arr = np.pad(arr, (0, n - len(arr)), constant_values=np.nan)
                values.append(arr)
            else:
                val = static_values.get(full_key)
                if val is None:
                    val = np.nan
                if isinstance(val, str) and val == ERROR_MARKER:
                    arr = np.full(n, ERROR_MARKER, dtype=object)
                else:
                    try:
                        arr = np.full(n, float(val), dtype=float)
                    except:
                        arr = np.full(n, np.nan, dtype=float)
                values.append(arr)

    if values:
        return np.stack(values, axis=0)
    else:
        return np.full((0, n), np.nan, dtype=float)
    
def _eval_atom_vec_live(
    atom: str, n: int,
    static_values: Dict[str, Any],
    live_arrays: Dict[str, np.ndarray],
    default_sheet: str,
    simtable_values: Dict[str, float] = None
) -> np.ndarray:
    """评估原子表达式（不含函数调用）"""
    if ':' in atom and _RANGE_PAT.match(atom):
        return _eval_range_live(atom, n, static_values, live_arrays, default_sheet, simtable_values)

    # 处理百分号（如 "1%" -> 0.01）
    if isinstance(atom, str):
        atom_stripped = atom.strip()
        if atom_stripped.endswith('%'):
            try:
                num_str = atom_stripped.rstrip('%')
                val = float(num_str) / 100.0
                return np.full(n, val, dtype=float)
            except ValueError:
                pass

    try:
        val = float(atom)
        return np.full(n, val, dtype=float)
    except ValueError:
        pass

    if _EMPTY_STRING_PATTERN.match(atom):
        return np.zeros(n, dtype=float)

    resolved, extra_vars = _resolve_atoms_live(atom, n, static_values, live_arrays, default_sheet, simtable_values)
    resolved = _convert_excel_operators(resolved)
    eval_ns = {'np': np, '__builtins__': {}}
    eval_ns.update(extra_vars)
    try:
        result = eval(resolved, eval_ns)
        if isinstance(result, np.ndarray):
            if result.dtype.kind in 'fc':
                return result.astype(float)
            else:
                return result
        return np.full(n, float(result), dtype=float)
    except Exception:
        return np.full(n, np.nan, dtype=float)
    
def _range_to_coords(rng: str) -> List[str]:
    """将区域引用（如 A1:B2）展开为单元格坐标列表（大写字母，已去除 $）"""
    rng = rng.replace('$', '')
    m = re.match(r'([A-Z]+)(\d+):([A-Z]+)(\d+)', rng, re.I)
    if not m:
        return [rng.upper()]
    sc, sr, ec, er = m.groups()
    coords = []
    for row in range(int(sr), int(er) + 1):
        for cn in range(_col2n(sc), _col2n(ec) + 1):
            coords.append(f"{_n2col(cn)}{row}")
    return coords

# ==================== 增强的聚合函数（支持 ERROR_MARKER 和 NaN）====================
def _eval_aggregate_live(
    fn_upper: str, raw_args: List[str], n: int,
    static_values: Dict[str, Any],
    live_arrays: Dict[str, np.ndarray],
    xl_ns: dict, default_sheet: str,
    simtable_values: Dict[str, float] = None,
    _error_sink: Optional[List] = None
) -> np.ndarray:
    """
    评估聚合函数（SUM, AVERAGE, MAX, MIN, PRODUCT, COUNT, COUNTA, STDEV, STDEVP, VAR, VARP, MEDIAN 等）
    返回长度为 n 的向量。
    增强：将参数中的 ERROR_MARKER 字符串转换为 NaN，使用 nan-safe 函数计算结果。
    修复：区域引用中的单元格去重，避免重复计算。
    """
    vals = []
    seen_cells = set()

    for arg in raw_args:
        arg = arg.strip()
        if _RANGE_PAT.match(arg):
            # 解析工作表名和区域
            sheet = default_sheet
            rng_str = arg
            if '!' in arg:
                sheet_part, rng_str = arg.split('!', 1)
                sheet_part = sheet_part.strip()
                if sheet_part.startswith("'") and sheet_part.endswith("'"):
                    sheet_part = sheet_part[1:-1].replace("''", "'")
                sheet = sheet_part
            # 展开区域为单元格坐标列表
            coords = _range_to_coords(rng_str)
            for coord in coords:
                ref_key = _normalize_key(sheet, coord)
                if ref_key in seen_cells:
                    continue
                seen_cells.add(ref_key)

                # 获取单元格的数组值
                if ref_key in live_arrays:
                    v = live_arrays[ref_key]
                    if v.dtype.kind == 'O':
                        v_float = np.full(len(v), np.nan, dtype=float)
                        mask = (v == ERROR_MARKER)
                        valid_mask = ~mask
                        try:
                            v_float[valid_mask] = v[valid_mask].astype(float)
                        except:
                            pass
                        v = v_float
                    elif v.dtype.kind in 'fc':
                        v = v.astype(float)
                    if DEBUG:
                        print(f"聚合函数 {fn_upper}: 区域单元格 {ref_key} 形状 {v.shape}, 前5值 {v[:5]}")
                    vals.append(v)
                else:
                    sv = static_values.get(ref_key, static_values.get(coord.upper(), np.nan))
                    if DEBUG:
                        print(f"警告：区域 {rng_str} 中的 {coord} 未在 live_arrays 中找到，使用静态值 {sv}")
                    if sv == ERROR_MARKER:
                        vals.append(np.full(n, ERROR_MARKER, dtype=object))
                    else:
                        try:
                            v = np.full(n, float(sv)) if not np.isnan(float(sv)) else np.full(n, np.nan)
                            vals.append(v)
                        except:
                            vals.append(np.full(n, np.nan))
        else:
            # 非区域引用：评估表达式
            v = _eval_expr_vec_live(
                arg, n, static_values, live_arrays, xl_ns, default_sheet,
                simtable_values=simtable_values, _error_sink=_error_sink)
            # 将 ERROR_MARKER 转换为 NaN
            if isinstance(v, np.ndarray) and v.dtype.kind == 'O':
                v_float = np.full(len(v), np.nan, dtype=float)
                mask = (v == ERROR_MARKER)
                valid_mask = ~mask
                try:
                    v_float[valid_mask] = v[valid_mask].astype(float)
                except:
                    pass
                v = v_float
            elif isinstance(v, np.ndarray) and v.dtype.kind in 'fc':
                v = v.astype(float)
            vals.append(v)

    if not vals:
        return np.zeros(n, dtype=float)

    # 确保所有数组都是一维且长度一致
    for i, v in enumerate(vals):
        if v.ndim != 1:
            v = v.flatten()
        if len(v) != n:
            if len(v) > n:
                v = v[:n]
            else:
                v = np.pad(v, (0, n - len(v)), constant_values=np.nan)
        vals[i] = v

    # 堆叠成二维数组
    stacked = np.vstack(vals)

    # 根据聚合函数类型计算
    if fn_upper in ('SUM', 'SUMA'):
        result = np.nansum(stacked, axis=0)
    elif fn_upper in ('AVERAGE', 'AVERAGEA'):
        result = np.nanmean(stacked, axis=0)
    elif fn_upper in ('MAX', 'MAXA'):
        result = np.nanmax(stacked, axis=0)
    elif fn_upper in ('MIN', 'MINA'):
        result = np.nanmin(stacked, axis=0)
    elif fn_upper == 'PRODUCT':
        result = np.nanprod(stacked, axis=0)
    elif fn_upper == 'COUNT':
        mask = ~np.isnan(stacked)
        result = mask.sum(axis=0).astype(float)
    elif fn_upper == 'COUNTA':
        mask = ~np.isnan(stacked)
        result = mask.sum(axis=0).astype(float)
    elif fn_upper in ('STDEV', 'STDEVA', 'STDEVP', 'STDEVPA', 'VAR', 'VARA', 'VARP', 'VARPA'):
        ddof = 1 if fn_upper in ('STDEV', 'STDEVA', 'VAR', 'VARA') else 0
        if fn_upper.startswith('STDEV'):
            result = np.nanstd(stacked, axis=0, ddof=ddof)
        else:
            result = np.nanvar(stacked, axis=0, ddof=ddof)
    elif fn_upper == 'MEDIAN':
        result = np.nanmedian(stacked, axis=0)
    else:
        result = np.zeros(n, dtype=float)

    # 如果结果全为 NaN，可选记录错误
    if np.all(np.isnan(result)):
        msg = f"{default_sheet} 聚合函数 {fn_upper} 结果全为 NaN，参数可能全部无效"
        if _error_sink is not None:
            _error_sink.append(msg)
        if DEBUG:
            print(f"警告: {msg}")

    return result

def _npv_vectorized(rate, *cash_flows):
    """向量化 NPV 计算，rate 和每个现金流都是长度为 n 的数组"""
    rate_arr = np.asarray(rate)
    n = len(rate_arr)
    cf_list = []
    for cf in cash_flows:
        cf_arr = np.asarray(cf)
        if cf_arr.ndim == 0:
            cf_arr = np.full(n, cf_arr)
        elif len(cf_arr) != n:
            if len(cf_arr) == 1:
                cf_arr = np.full(n, cf_arr[0])
            else:
                raise ValueError(f"NPV 参数长度 {len(cf_arr)} 与贴现率长度 {n} 不匹配")
        cf_list.append(cf_arr)
    if not cf_list:
        return np.zeros(n, dtype=float)
    cf_stack = np.vstack(cf_list)
    t = cf_stack.shape[0]
    rate_expanded = rate_arr.reshape(1, -1)
    t_vals = np.arange(1, t+1).reshape(-1, 1)
    denominator = (1 + rate_expanded) ** t_vals
    discount_factors = 1.0 / denominator
    npv = np.sum(cf_stack * discount_factors, axis=0)
    return npv

# ==================== 分布类型辅助函数 ====================
def _get_dist_type_from_func_name(func_name: str) -> str:
    """
    根据分布函数名返回内部类型字符串（小写）。
    使用映射表提高效率和可维护性。
    """
    # 去除可能的 _xll. 前缀并转为小写
    clean_name = func_name.lower()
    if clean_name.startswith('_xll.'):
        clean_name = clean_name[5:]

    # 映射表：分布函数名（小写，不带 _xll. 前缀） -> 类型字符串
    DIST_TYPE_MAP = {
        # 连续分布
        'drisknormal': 'normal',
        'driskuniform': 'uniform',
        'driskgamma': 'gamma',
        'driskbeta': 'beta',
        'driskchisq': 'chisq',
        'driskf': 'f',
        'driskstudent': 'student',
        'driskexpon': 'expon',
        'driskbernoulli': 'bernoulli',
        'driskbinomial': 'binomial',
        'drisktriang': 'triang',
        'driskinvgauss': 'invgauss',
        'driskduniform': 'duniform', 
        'driskpoisson': 'poisson',
        'drisktrigen': 'trigen',
        'driskcumul': 'cumul',
        'driskdiscrete': 'discrete',
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
        'driskerlang': 'erlang',
        'driskcauchy': 'cauchy',
        'driskdagum': 'dagum',
        'driskdoubletriang': 'doubletriang',
        'drisknegbin': 'negbin',
        'driskgeomet': 'geomet',
        'driskhypergeo': 'hypergeo',
        'driskintuniform': 'intuniform',
        # 若还有其他别名，可继续添加
    }

    # 优先精确匹配（直接查找）
    if clean_name in DIST_TYPE_MAP:
        return DIST_TYPE_MAP[clean_name]

    # 后备：尝试子串匹配（处理可能带参数前缀的情况，但通常不应发生）
    # 注意：DUniform 必须优先于 Uniform，但由于精确匹配已处理，这里不会混淆
    for name, dtype in DIST_TYPE_MAP.items():
        if name in clean_name:
            return dtype

    # 默认返回 normal
    return 'normal'
# ==================== 新的向量化分布生成器 ====================
try:
    from distribution_base import (
        GammaDistribution, BetaDistribution, ChiSquaredDistribution,
        FDistribution, TDistribution, PoissonDistribution,
        ExponentialDistribution, UniformDistribution, NormalDistribution
    )
    from dist_bernoulli import BernoulliDistribution
    from dist_binomial import BinomialDistribution
    from dist_invgauss import InvgaussDistribution
    from dist_duniform import DUniformDistribution
    from dist_triang import TriangDistribution
    from dist_trigen import TrigenDistribution
    from dist_cumul import CumulDistribution
    from dist_discrete import DiscreteDistribution
    DIST_CLASSES_AVAILABLE = True
except ImportError:
    DIST_CLASSES_AVAILABLE = False

class VectorizedDistributionGenerator:
    """
    向量化分布生成器，支持参数为数组（长度 n），一次性生成所有样本。
    处理所有截断类型、平移、loc/lock、静态值，错误标记传播。
    增强版本：参数验证、截断边界调整、精确 loc 值计算。
    支持采样方法：蒙特卡洛 (MC)、拉丁超立方 (LHC)、Sobol 序列 (SOBOL)。
    """
    # 映射分布类型到对应的分布类（用于计算截断均值）
    DIST_CLASS_MAP = {
        'normal': NormalDistribution,
        'uniform': UniformDistribution,
        'gamma': GammaDistribution,
        'beta': BetaDistribution,
        'poisson': PoissonDistribution,
        'chisq': ChiSquaredDistribution,
        'f': FDistribution,
        'student': TDistribution,
        'expon': ExponentialDistribution,
        'bernoulli': BernoulliDistribution,
        'triang': TriangDistribution,
        'binomial': BinomialDistribution,
        'invgauss': InvgaussDistribution,
        'duniform': DUniformDistribution,
        'trigen': TrigenDistribution,
        'cumul': CumulDistribution,
        'discrete': DiscreteDistribution,
        # 新增分布类（如果存在对应类，否则回退到原始均值）
        'erf': ErfDistribution,
        'extvalue': ExtvalueDistribution,
        'extvaluemin': ExtvalueMinDistribution,
        'fatiguelife': FatigueLifeDistribution,
        'frechet': FrechetDistribution,
        'general': GeneralDistribution,
        'histogrm': HistogrmDistribution,
        'hypsecant': HypSecantDistribution,
        'johnsonsb': JohnsonSBDistribution,
        'johnsonsu': JohnsonSUDistribution,
        'kumaraswamy': KumaraswamyDistribution,
        'laplace': LaplaceDistribution,
        'logistic': LogisticDistribution,
        'loglogistic': LoglogisticDistribution,
        'lognorm': LognormDistribution,
        'lognorm2': Lognorm2Distribution,
        'betageneral': BetaGeneralDistribution,
        'betasubj': BetaSubjDistribution,
        'burr12': Burr12Distribution,
        'compound': CompoundDistribution,
        'splice': SpliceDistribution,
        'pert': PertDistribution,
        'reciprocal': ReciprocalDistribution,
        'rayleigh': RayleighDistribution,
        'weibull': WeibullDistribution,
        'pearson5': Pearson5Distribution,
        'pearson6': Pearson6Distribution,
        'pareto2': Pareto2Distribution,
        'pareto': ParetoDistribution,
        'levy': LevyDistribution,
        'erlang': ErlangDistribution,
        'cauchy': CauchyDistribution,
        'dagum': DagumDistribution,
        'doubletriang': DoubleTriangDistribution,
        'negbin': NegbinDistribution,
        'geomet': GeometDistribution,
        'hypergeo': HypergeoDistribution,
        'intuniform': IntuniformDistribution,
    }

    def __init__(self, func_name: str, dist_type: str, markers: Dict[str, Any], n: int, method: str = 'MC'):
        """
        初始化向量化分布生成器。

        Args:
            func_name: 分布函数名（如 'DriskNormal'）
            dist_type: 分布类型（小写）
            markers: 属性标记字典
            n: 样本数量（迭代次数）
            method: 采样方法，可选 'MC', 'LHC', 'SOBOL'
        """
        self.func_name = func_name
        self.dist_type = dist_type
        self.markers = markers
        self.n = n
        self.method = method

        self.shift_amount = 0.0
        self.truncate_type = None
        self.truncate_lower = None
        self.truncate_upper = None
        self.truncate_lower_pct = None
        self.truncate_upper_pct = None
        self._extract_markers()

        self.x_vals_list = None
        self.p_vals_list = None
        self._array_parse_failed = False

        # 处理需要数组参数的分布（Cumul, Discrete, DUniform, Histogrm, General）
        if self.dist_type in ['cumul', 'discrete', 'duniform', 'histogrm', 'general']:
            x_vals_str = markers.get('x_vals')
            p_vals_str = markers.get('p_vals')

            # Histogrm: 只有 p_vals
            if self.dist_type == 'histogrm' and p_vals_str and not x_vals_str:
                try:
                    def clean(s):
                        s = s.strip()
                        if s.startswith('"') and s.endswith('"'):
                            s = s[1:-1]
                        if s.startswith("'") and s.endswith("'"):
                            s = s[1:-1]
                        if s.startswith('{') and s.endswith('}'):
                            s = s[1:-1]
                        return s
                    p_vals_str_clean = clean(p_vals_str)
                    self.p_vals_list = [float(p) for p in p_vals_str_clean.split(',') if p.strip()]
                    if DEBUG:
                        print(f"向量化生成器解析 Histogrm P-Table 成功: p={self.p_vals_list}")
                except Exception as e:
                    print(f"向量化生成器解析 Histogrm 参数失败: {e}")
                    self._array_parse_failed = True

            # General: 需要 x_vals 和 p_vals
            elif self.dist_type == 'general' and x_vals_str and p_vals_str:
                try:
                    def clean(s):
                        s = s.strip()
                        if s.startswith('"') and s.endswith('"'):
                            s = s[1:-1]
                        if s.startswith("'") and s.endswith("'"):
                            s = s[1:-1]
                        if s.startswith('{') and s.endswith('}'):
                            s = s[1:-1]
                        return s
                    x_vals_str_clean = clean(x_vals_str)
                    p_vals_str_clean = clean(p_vals_str)
                    self.x_vals_list = [float(x) for x in x_vals_str_clean.split(',') if x.strip()]
                    self.p_vals_list = [float(p) for p in p_vals_str_clean.split(',') if p.strip()]
                    if len(self.x_vals_list) != len(self.p_vals_list):
                        raise ValueError("X 和 P 数组长度不相等")
                    if DEBUG:
                        print(f"向量化生成器解析 General 数组成功: x={self.x_vals_list}, p={self.p_vals_list}")
                except Exception as e:
                    print(f"向量化生成器解析 General 参数失败: {e}")
                    self._array_parse_failed = True

            # 其他数组分布（cumul, discrete, duniform）
            elif x_vals_str and p_vals_str:
                try:
                    def clean(s):
                        s = s.strip()
                        if s.startswith('"') and s.endswith('"'):
                            s = s[1:-1]
                        if s.startswith("'") and s.endswith("'"):
                            s = s[1:-1]
                        if s.startswith('{') and s.endswith('}'):
                            s = s[1:-1]
                        return s
                    x_vals_str_clean = clean(x_vals_str)
                    p_vals_str_clean = clean(p_vals_str)
                    def parse_numbers(s):
                        parts = [p.strip() for p in s.split(',') if p.strip()]
                        return [float(p) for p in parts]
                    self.x_vals_list = parse_numbers(x_vals_str_clean)
                    self.p_vals_list = parse_numbers(p_vals_str_clean)
                    # 归一化并修正最后一个概率
                    if self.dist_type in ('discrete', 'duniform'):
                        total = sum(self.p_vals_list)
                        if total > 0:
                            self.p_vals_list = [p / total for p in self.p_vals_list]
                        if len(self.p_vals_list) > 0:
                            self.p_vals_list[-1] = 1.0 - sum(self.p_vals_list[:-1])
                    if DEBUG:
                        print(f"向量化生成器解析数组成功: x={self.x_vals_list}, p={self.p_vals_list}")
                except Exception as e:
                    print(f"向量化生成器解析数组参数失败: {e}, x_vals='{x_vals_str}', p_vals='{p_vals_str}'")
                    self._array_parse_failed = True
                    self.x_vals_list = self.p_vals_list = None
            elif self.dist_type == 'duniform':
                x_vals_str = markers.get('x_vals')
                p_vals_str = markers.get('p_vals')
                if x_vals_str and p_vals_str:
                    try:
                        def clean(s):
                            s = s.strip()
                            if s.startswith('"') and s.endswith('"'):
                                s = s[1:-1]
                            if s.startswith("'") and s.endswith("'"):
                                s = s[1:-1]
                            if s.startswith('{') and s.endswith('}'):
                                s = s[1:-1]
                            return s
                        x_vals_str_clean = clean(x_vals_str)
                        p_vals_str_clean = clean(p_vals_str)
                        self.x_vals_list = [float(x) for x in x_vals_str_clean.split(',') if x.strip()]
                        self.p_vals_list = [float(p) for p in p_vals_str_clean.split(',') if p.strip()]
                        if len(self.x_vals_list) != len(self.p_vals_list):
                            raise ValueError("X 和 P 数组长度不相等")
                        # 归一化概率（确保和为1）
                        total = sum(self.p_vals_list)
                        if total > 0:
                            self.p_vals_list = [p / total for p in self.p_vals_list]
                        # 修正最后一个概率的舍入误差
                        if len(self.p_vals_list) > 0:
                            self.p_vals_list[-1] = 1.0 - sum(self.p_vals_list[:-1])
                        if DEBUG:
                            print(f"向量化生成器解析 DUniform 成功: x={self.x_vals_list}, p={self.p_vals_list}")
                    except Exception as e:
                        print(f"向量化生成器解析 DUniform 参数失败: {e}, x_vals='{x_vals_str}', p_vals='{p_vals_str}'")
                        self._array_parse_failed = True
                        self.x_vals_list = self.p_vals_list = None
                else:
                    # 如果没有 markers，尝试从参数文本解析（不应发生，但保留）
                    print("DUniform: markers 中缺少 x_vals 或 p_vals，无法解析")
                    self._array_parse_failed = True

        self.scipy_available = False
        try:
            import scipy.stats
            self.scipy_available = True
        except ImportError:
            pass

    def _extract_markers(self):
        shift_val = self.markers.get('shift')
        if shift_val is not None:
            try:
                self.shift_amount = float(shift_val)
            except:
                self.shift_amount = 0.0

        truncate_types = ['truncate', 'truncate2', 'truncatep', 'truncatep2']
        for tt in truncate_types:
            val = self.markers.get(tt)
            if val is not None:
                self.truncate_type = tt
                if isinstance(val, str):
                    val = val.strip()
                    if val.startswith('(') and val.endswith(')'):
                        val = val[1:-1]
                    parts = [p.strip() for p in val.split(',')]
                    lower = None
                    upper = None
                    if len(parts) >= 1 and parts[0]:
                        try:
                            lower = float(parts[0])
                        except:
                            pass
                    if len(parts) >= 2 and parts[1]:
                        try:
                            upper = float(parts[1])
                        except:
                            pass
                    if tt in ('truncatep', 'truncatep2'):
                        self.truncate_lower_pct = lower
                        self.truncate_upper_pct = upper
                    else:
                        self.truncate_lower = lower
                        self.truncate_upper = upper
                else:
                    try:
                        v = float(val)
                        if tt in ('truncatep', 'truncatep2'):
                            self.truncate_lower_pct = v
                        else:
                            self.truncate_lower = v
                    except:
                        pass
                break

    def _get_support_range(self, params_dict: Dict[str, np.ndarray]) -> Tuple[np.ndarray, np.ndarray]:
        n = self.n
        dist_type = self.dist_type
        if dist_type == 'normal':
            return np.full(n, -np.inf), np.full(n, np.inf)
        elif dist_type == 'uniform':
            return params_dict['min'], params_dict['max']
        elif dist_type == 'gamma':
            return np.full(n, 0.0), np.full(n, np.inf)
        elif dist_type == 'poisson':
            return np.full(n, 0.0), np.full(n, np.inf)
        elif dist_type == 'beta':
            return np.full(n, 0.0), np.full(n, 1.0)
        elif dist_type == 'chisq':
            return np.full(n, 0.0), np.full(n, np.inf)
        elif dist_type == 'f':
            return np.full(n, 0.0), np.full(n, np.inf)
        elif dist_type == 'student':
            return np.full(n, -np.inf), np.full(n, np.inf)
        elif dist_type == 'expon':
            return np.full(n, 0.0), np.full(n, np.inf)
        elif dist_type == 'bernoulli':
            return np.full(n, 0.0), np.full(n, 1.0)
        elif dist_type == 'triang':
            return params_dict['a'], params_dict['b']
        elif dist_type == 'binomial':
            return np.full(n, 0.0), params_dict['n']
        elif dist_type == 'invgauss':
            return np.full(n, 0.0), np.full(n, np.inf)
        elif dist_type == 'duniform':
            if self.x_vals_list:
                return np.full(n, min(self.x_vals_list)), np.full(n, max(self.x_vals_list))
            else:
                return np.full(n, -np.inf), np.full(n, np.inf)
        elif dist_type == 'trigen':
            return np.full(n, -np.inf), np.full(n, np.inf)
        elif dist_type == 'cumul':
            if self.x_vals_list:
                return np.full(n, min(self.x_vals_list)), np.full(n, max(self.x_vals_list))
            else:
                return np.full(n, -np.inf), np.full(n, np.inf)
        elif dist_type == 'discrete':
            if self.x_vals_list:
                return np.full(n, min(self.x_vals_list)), np.full(n, max(self.x_vals_list))
            else:
                return np.full(n, -np.inf), np.full(n, np.inf)
        # 新增分布
        elif dist_type == 'erf':
            return np.full(n, -np.inf), np.full(n, np.inf)
        elif dist_type in ('extvalue', 'extvaluemin'):
            return np.full(n, -np.inf), np.full(n, np.inf)
        elif dist_type == 'fatiguelife':
            return params_dict['y'], np.full(n, np.inf)
        elif dist_type == 'frechet':
            return params_dict['y'], np.full(n, np.inf)
        elif dist_type == 'general':
            if self.x_vals_list:
                return np.full(n, min(self.x_vals_list)), np.full(n, max(self.x_vals_list))
            else:
                return params_dict['min'], params_dict['max']
        elif dist_type == 'histogrm':
            return params_dict['min'], params_dict['max']
        elif dist_type == 'hypsecant':
            return np.full(n, -np.inf), np.full(n, np.inf)
        elif dist_type == 'johnsonsb':
            return params_dict['a'], params_dict['b']
        elif dist_type == 'johnsonsu':
            return np.full(n, -np.inf), np.full(n, np.inf)
        elif dist_type == 'kumaraswamy':
            return params_dict['min_val'], params_dict['max_val']
        elif dist_type == 'laplace':
            return np.full(n, -np.inf), np.full(n, np.inf)
        elif dist_type == 'logistic':
            return np.full(n, -np.inf), np.full(n, np.inf)
        elif dist_type == 'loglogistic':
            return params_dict['gamma'], np.full(n, np.inf)
        elif dist_type in ('lognorm', 'lognorm2'):
            return np.full(n, 0.0), np.full(n, np.inf)
        elif dist_type == 'betageneral':
            return params_dict['min_val'], params_dict['max_val']
        elif dist_type == 'betasubj':
            return params_dict['min_val'], params_dict['max_val']
        elif dist_type == 'burr12':
            return params_dict['gamma'], np.full(n, np.inf)
        elif dist_type == 'compound':
            return np.full(n, 0.0), np.full(n, np.inf)
        elif dist_type == 'splice':
            return np.full(n, -np.inf), np.full(n, np.inf)
        elif dist_type == 'pert':
            return params_dict['min_val'], params_dict['max_val']
        elif dist_type == 'reciprocal':
            return params_dict['min_val'], params_dict['max_val']
        elif dist_type == 'rayleigh':
            return np.full(n, 0.0), np.full(n, np.inf)
        elif dist_type == 'weibull':
            return np.full(n, 0.0), np.full(n, np.inf)
        elif dist_type == 'pearson5':
            return np.full(n, 0.0), np.full(n, np.inf)
        elif dist_type == 'pearson6':
            return np.full(n, 0.0), np.full(n, np.inf)
        elif dist_type == 'pareto2':
            return np.full(n, 0.0), np.full(n, np.inf)
        elif dist_type == 'pareto':
            return params_dict['alpha'], np.full(n, np.inf)
        elif dist_type == 'levy':
            return params_dict['a'], np.full(n, np.inf)
        elif dist_type == 'erlang':
            return np.full(n, 0.0), np.full(n, np.inf)
        elif dist_type == 'cauchy':
            return np.full(n, -np.inf), np.full(n, np.inf)
        elif dist_type == 'dagum':
            return params_dict['gamma'], np.full(n, np.inf)
        elif dist_type == 'doubletriang':
            return params_dict['min_val'], params_dict['max_val']
        elif dist_type == 'negbin':
            p = params_dict.get('p')
            if p is not None and np.all(p >= 1.0):
                return np.full(n, 0.0), np.full(n, 0.0)
            else:
                return np.full(n, 0.0), np.full(n, np.inf)
        elif dist_type == 'geomet':
            return np.full(n, 0.0), np.full(n, np.inf)
        elif dist_type == 'hypergeo':
            n_param = params_dict['n']
            D = params_dict['D']
            M = params_dict['M']
            low = np.maximum(0, n_param + D - M)
            high = np.minimum(n_param, D)
            return low, high
        elif dist_type == 'intuniform':
            return params_dict['min_val'], params_dict['max_val']
        else:
            return np.full(n, -np.inf), np.full(n, np.inf)

    def _validate_params(self, params_dict: Dict[str, np.ndarray]) -> np.ndarray:
        n = self.n
        valid = np.ones(n, dtype=bool)
        dist_type = self.dist_type

        # 基本数值验证
        if dist_type == 'normal':
            std = params_dict.get('std')
            if std is not None:
                valid &= (std > 0)
        elif dist_type == 'uniform':
            min_ = params_dict.get('min')
            max_ = params_dict.get('max')
            if min_ is not None and max_ is not None:
                valid &= (min_ < max_)
        elif dist_type == 'gamma':
            shape = params_dict.get('shape')
            scale = params_dict.get('scale')
            if shape is not None and scale is not None:
                valid &= (shape > 0) & (scale > 0)
        elif dist_type == 'poisson':
            lam = params_dict.get('lam')
            if lam is not None:
                valid &= (lam >= 0)
        elif dist_type == 'beta':
            a = params_dict.get('a')
            b = params_dict.get('b')
            if a is not None and b is not None:
                valid &= (a > 0) & (b > 0)
        elif dist_type == 'chisq':
            df = params_dict.get('df')
            if df is not None:
                valid &= (df > 0)
        elif dist_type == 'f':
            df1 = params_dict.get('df1')
            df2 = params_dict.get('df2')
            if df1 is not None and df2 is not None:
                valid &= (df1 > 0) & (df2 > 0)
        elif dist_type == 'student':
            df = params_dict.get('df')
            if df is not None:
                valid &= (df > 0)
        elif dist_type == 'expon':
            lam = params_dict.get('lam')
            if lam is not None:
                valid &= (lam > 0)
        elif dist_type == 'bernoulli':
            p = params_dict.get('p')
            if p is not None:
                valid &= (p >= 0) & (p <= 1)
        elif dist_type == 'triang':
            a = params_dict.get('a')
            c = params_dict.get('c')
            b = params_dict.get('b')
            if a is not None and c is not None and b is not None:
                valid &= (a <= c) & (c <= b)
        elif dist_type == 'binomial':
            n_trials = params_dict.get('n')
            p = params_dict.get('p')
            if n_trials is not None and p is not None:
                valid &= (n_trials > 0) & (p >= 0) & (p <= 1)
        elif dist_type == 'invgauss':
            mu = params_dict.get('mu')
            lam = params_dict.get('lam')
            if mu is not None and lam is not None:
                valid &= (mu > 0) & (lam > 0)
        elif dist_type in ('duniform', 'cumul', 'discrete'):
            if self._array_parse_failed or not self.x_vals_list or not self.p_vals_list:
                valid &= False
            elif dist_type == 'discrete':
                total_p = sum(self.p_vals_list)
                if abs(total_p - 1.0) > 1e-10:
                    valid &= False
        # 新增分布参数验证
        elif dist_type == 'erf':
            h = params_dict.get('h')
            if h is not None:
                valid &= (h > 0)
        elif dist_type in ('extvalue', 'extvaluemin'):
            B = params_dict.get('B')
            if B is not None:
                valid &= (B > 0)
        elif dist_type == 'fatiguelife':
            beta = params_dict.get('beta')
            alpha = params_dict.get('alpha')
            if beta is not None and alpha is not None:
                valid &= (beta > 0) & (alpha > 0)
        elif dist_type == 'frechet':
            beta = params_dict.get('beta')
            alpha = params_dict.get('alpha')
            if beta is not None and alpha is not None:
                valid &= (beta > 0) & (alpha > 0)
        elif dist_type == 'general':
            if self._array_parse_failed or not self.x_vals_list or not self.p_vals_list:
                valid &= False
        elif dist_type == 'histogrm':
            if self._array_parse_failed or not self.p_vals_list:
                valid &= False
            min_val = params_dict.get('min')
            max_val = params_dict.get('max')
            if min_val is not None and max_val is not None:
                valid &= (min_val < max_val)
        elif dist_type == 'hypsecant':
            beta = params_dict.get('beta')
            if beta is not None:
                valid &= (beta > 0)
        elif dist_type == 'johnsonsb':
            alpha2 = params_dict.get('alpha2')
            a = params_dict.get('a')
            b = params_dict.get('b')
            if alpha2 is not None and a is not None and b is not None:
                valid &= (alpha2 > 0) & (b > a)
        elif dist_type == 'johnsonsu':
            alpha2 = params_dict.get('alpha2')
            beta = params_dict.get('beta')
            if alpha2 is not None and beta is not None:
                valid &= (alpha2 > 0) & (beta > 0)
        elif dist_type == 'kumaraswamy':
            alpha1 = params_dict.get('alpha1')
            alpha2 = params_dict.get('alpha2')
            min_val = params_dict.get('min_val')
            max_val = params_dict.get('max_val')
            if alpha1 is not None and alpha2 is not None and min_val is not None and max_val is not None:
                valid &= (alpha1 > 0) & (alpha2 > 0) & (max_val > min_val)
        elif dist_type == 'laplace':
            sigma = params_dict.get('sigma')
            if sigma is not None:
                valid &= (sigma > 0)
        elif dist_type == 'logistic':
            beta = params_dict.get('beta')
            if beta is not None:
                valid &= (beta > 0)
        elif dist_type == 'loglogistic':
            beta = params_dict.get('beta')
            alpha = params_dict.get('alpha')
            if beta is not None and alpha is not None:
                valid &= (beta > 0) & (alpha > 0)
        elif dist_type == 'lognorm':
            mu = params_dict.get('mu')
            sigma = params_dict.get('sigma')
            if mu is not None and sigma is not None:
                valid &= (mu > 0) & (sigma > 0)
        elif dist_type == 'lognorm2':
            sigma = params_dict.get('sigma')
            if sigma is not None:
                valid &= (sigma > 0)
        elif dist_type == 'betageneral':
            alpha1 = params_dict.get('alpha1')
            alpha2 = params_dict.get('alpha2')
            min_val = params_dict.get('min_val')
            max_val = params_dict.get('max_val')
            if alpha1 is not None and alpha2 is not None and min_val is not None and max_val is not None:
                valid &= (alpha1 > 0) & (alpha2 > 0) & (max_val > min_val)
        elif dist_type == 'betasubj':
            min_val = params_dict.get('min_val')
            m_likely = params_dict.get('m_likely')
            mean_val = params_dict.get('mean_val')
            max_val = params_dict.get('max_val')
            if min_val is not None and m_likely is not None and mean_val is not None and max_val is not None:
                valid &= (min_val < m_likely < max_val) & (min_val < mean_val < max_val)
        elif dist_type == 'burr12':
            beta = params_dict.get('beta')
            alpha1 = params_dict.get('alpha1')
            alpha2 = params_dict.get('alpha2')
            if beta is not None and alpha1 is not None and alpha2 is not None:
                valid &= (beta > 0) & (alpha1 > 0) & (alpha2 > 0)
        elif dist_type == 'compound':
            # Compound 的验证比较复杂，默认通过
            pass
        elif dist_type == 'splice':
            # Splice 的验证通过参数格式判断，默认通过
            pass
        elif dist_type == 'pert':
            min_val = params_dict.get('min_val')
            m_likely = params_dict.get('m_likely')
            max_val = params_dict.get('max_val')
            if min_val is not None and m_likely is not None and max_val is not None:
                valid &= (min_val <= m_likely <= max_val) & (min_val < max_val)
        elif dist_type == 'reciprocal':
            min_val = params_dict.get('min_val')
            max_val = params_dict.get('max_val')
            if min_val is not None and max_val is not None:
                valid &= (min_val > 0) & (max_val > min_val)
        elif dist_type == 'rayleigh':
            b = params_dict.get('b')
            if b is not None:
                valid &= (b > 0)
        elif dist_type == 'weibull':
            alpha = params_dict.get('alpha')
            beta = params_dict.get('beta')
            if alpha is not None and beta is not None:
                valid &= (alpha > 0) & (beta > 0)
        elif dist_type == 'pearson5':
            alpha = params_dict.get('alpha')
            beta = params_dict.get('beta')
            if alpha is not None and beta is not None:
                valid &= (alpha > 0) & (beta > 0)
        elif dist_type == 'pearson6':
            alpha1 = params_dict.get('alpha1')
            alpha2 = params_dict.get('alpha2')
            beta = params_dict.get('beta')
            if alpha1 is not None and alpha2 is not None and beta is not None:
                valid &= (alpha1 > 0) & (alpha2 > 0) & (beta > 0)
        elif dist_type == 'pareto2':
            b = params_dict.get('b')
            q = params_dict.get('q')
            if b is not None and q is not None:
                valid &= (b > 0) & (q > 0)
        elif dist_type == 'pareto':
            theta = params_dict.get('theta')
            alpha = params_dict.get('alpha')
            if theta is not None and alpha is not None:
                valid &= (theta > 0) & (alpha > 0)
        elif dist_type == 'levy':
            c = params_dict.get('c')
            if c is not None:
                valid &= (c > 0)
        elif dist_type == 'erlang':
            m = params_dict.get('m')
            beta = params_dict.get('beta')
            if m is not None and beta is not None:
                valid &= (m > 0) & (beta > 0) & (np.abs(m - np.round(m)) < 1e-9)
        elif dist_type == 'cauchy':
            beta = params_dict.get('beta')
            if beta is not None:
                valid &= (beta > 0)
        elif dist_type == 'dagum':
            beta = params_dict.get('beta')
            alpha1 = params_dict.get('alpha1')
            alpha2 = params_dict.get('alpha2')
            if beta is not None and alpha1 is not None and alpha2 is not None:
                valid &= (beta > 0) & (alpha1 > 0) & (alpha2 > 0)
        elif dist_type == 'doubletriang':
            min_val = params_dict.get('min_val')
            m_likely = params_dict.get('m_likely')
            max_val = params_dict.get('max_val')
            lower_p = params_dict.get('lower_p')
            if min_val is not None and m_likely is not None and max_val is not None and lower_p is not None:
                valid &= (min_val < m_likely < max_val) & (0 <= lower_p <= 1)
        elif dist_type == 'negbin':
            s = params_dict.get('s')
            p = params_dict.get('p')
            if s is not None and p is not None:
                valid &= (s > 0) & (p > 0) & (p <= 1)
        elif dist_type == 'geomet':
            p = params_dict.get('p')
            if p is not None:
                valid &= (p > 0) & (p <= 1)
        elif dist_type == 'hypergeo':
            n_param = params_dict.get('n')
            D = params_dict.get('D')
            M = params_dict.get('M')
            if n_param is not None and D is not None and M is not None:
                valid &= (n_param > 0) & (D > 0) & (M > 0) & (n_param <= M) & (D <= M)
        elif dist_type == 'intuniform':
            min_val = params_dict.get('min_val')
            max_val = params_dict.get('max_val')
            if min_val is not None and max_val is not None:
                valid &= (min_val < max_val) & (np.abs(min_val - np.round(min_val)) < 1e-9) & (np.abs(max_val - np.round(max_val)) < 1e-9)
        elif dist_type == 'trigen':
            L = params_dict.get('L')
            M = params_dict.get('M')
            U = params_dict.get('U')
            alpha = params_dict.get('alpha')
            beta = params_dict.get('beta')
            if L is not None and M is not None and U is not None and alpha is not None and beta is not None:
                valid &= (L <= M <= U) & (0 <= alpha < beta <= 1)
        return valid

    def _adjust_truncation_boundaries(self, params_dict: Dict[str, np.ndarray]) -> Tuple[Optional[np.ndarray], Optional[np.ndarray], np.ndarray]:
        n = self.n
        if self.truncate_type is None:
            return None, None, np.ones(n, dtype=bool)

        if self.truncate_type in ('truncate', 'truncate2'):
            lower = np.full(n, self.truncate_lower) if self.truncate_lower is not None else None
            upper = np.full(n, self.truncate_upper) if self.truncate_upper is not None else None
        else:
            lower_p = np.full(n, self.truncate_lower_pct) if self.truncate_lower_pct is not None else None
            upper_p = np.full(n, self.truncate_upper_pct) if self.truncate_upper_pct is not None else None
            valid = np.ones(n, dtype=bool)
            if lower_p is not None:
                valid &= (lower_p >= 0) & (lower_p <= 1)
            if upper_p is not None:
                valid &= (upper_p >= 0) & (upper_p <= 1)
            return lower_p, upper_p, valid

        low_support, high_support = self._get_support_range(params_dict)

        if self.truncate_type == 'truncate2':
            orig_lower = lower - self.shift_amount if lower is not None else None
            orig_upper = upper - self.shift_amount if upper is not None else None
        else:
            orig_lower = lower
            orig_upper = upper

        if orig_lower is not None:
            new_lower = np.maximum(orig_lower, low_support)
        else:
            new_lower = None

        if orig_upper is not None:
            new_upper = np.minimum(orig_upper, high_support)
        else:
            new_upper = None

        valid = np.ones(n, dtype=bool)
        if new_lower is not None and new_upper is not None:
            valid &= (new_lower <= new_upper)
        elif new_lower is not None:
            valid &= ~( (high_support != np.inf) & (new_lower > high_support) )
        elif new_upper is not None:
            valid &= ~( (low_support != -np.inf) & (new_upper < low_support) )

        if self.truncate_type == 'truncate2':
            if new_lower is not None:
                new_lower = new_lower + self.shift_amount
            if new_upper is not None:
                new_upper = new_upper + self.shift_amount

        return new_lower, new_upper, valid

    def _handle_error_markers(self, param_arrays: List[np.ndarray]) -> Optional[np.ndarray]:
        if not param_arrays:
            return None
        valid = np.ones(self.n, dtype=bool)
        for arr in param_arrays:
            if arr.dtype.kind == 'O':
                err_mask = (arr == ERROR_MARKER)
                valid &= ~err_mask
            elif arr.dtype.kind in 'fc':
                err_mask = np.isnan(arr)
                valid &= ~err_mask
        if np.all(valid):
            return None
        return valid

    def _apply_shift(self, samples: np.ndarray) -> np.ndarray:
        if self.shift_amount != 0:
            return samples + self.shift_amount
        return samples

    def _get_param_names(self) -> List[str]:
        dist_type = self.dist_type
        if dist_type == 'normal':
            return ['mean', 'std']
        elif dist_type == 'uniform':
            return ['min', 'max']
        elif dist_type == 'gamma':
            return ['shape', 'scale']
        elif dist_type == 'poisson':
            return ['lam']
        elif dist_type == 'beta':
            return ['a', 'b']
        elif dist_type == 'chisq':
            return ['df']
        elif dist_type == 'f':
            return ['df1', 'df2']
        elif dist_type == 'student':
            return ['df']
        elif dist_type == 'expon':
            return ['lam']
        elif dist_type == 'bernoulli':
            return ['p']
        elif dist_type == 'triang':
            return ['a', 'c', 'b']
        elif dist_type == 'binomial':
            return ['n', 'p']
        elif dist_type == 'invgauss':
            return ['mu', 'lam']
        elif dist_type == 'duniform':
            return []
        elif dist_type == 'trigen':
            return ['L', 'M', 'U', 'alpha', 'beta']
        elif dist_type == 'cumul':
            return ['min', 'max']
        elif dist_type == 'discrete':
            return []
        elif dist_type == 'erf':
            return ['h']
        elif dist_type == 'extvalue':
            return ['A', 'B']
        elif dist_type == 'extvaluemin':
            return ['A', 'B']
        elif dist_type == 'fatiguelife':
            return ['y', 'beta', 'alpha']
        elif dist_type == 'frechet':
            return ['y', 'beta', 'alpha']
        elif dist_type == 'general':
            return ['min', 'max']
        elif dist_type == 'histogrm':
            return ['min', 'max']
        elif dist_type == 'hypsecant':
            return ['gamma', 'beta']
        elif dist_type == 'johnsonsb':
            return ['alpha1', 'alpha2', 'a', 'b']
        elif dist_type == 'johnsonsu':
            return ['alpha1', 'alpha2', 'gamma', 'beta']
        elif dist_type == 'kumaraswamy':
            return ['alpha1', 'alpha2', 'min_val', 'max_val']
        elif dist_type == 'laplace':
            return ['mu', 'sigma']
        elif dist_type == 'logistic':
            return ['alpha', 'beta']
        elif dist_type == 'loglogistic':
            return ['gamma', 'beta', 'alpha']
        elif dist_type == 'lognorm':
            return ['mu', 'sigma']
        elif dist_type == 'lognorm2':
            return ['mu', 'sigma']
        elif dist_type == 'betageneral':
            return ['alpha1', 'alpha2', 'min_val', 'max_val']
        elif dist_type == 'betasubj':
            return ['min_val', 'm_likely', 'mean_val', 'max_val']
        elif dist_type == 'burr12':
            return ['gamma', 'beta', 'alpha1', 'alpha2']
        elif dist_type == 'compound':
            return ['frequency', 'severity', 'deductible', 'limit']
        elif dist_type == 'splice':
            return ['left_dist', 'right_dist', 'splice_point']
        elif dist_type == 'pert':
            return ['min_val', 'm_likely', 'max_val']
        elif dist_type == 'reciprocal':
            return ['min_val', 'max_val']
        elif dist_type == 'rayleigh':
            return ['b']
        elif dist_type == 'weibull':
            return ['alpha', 'beta']
        elif dist_type == 'pearson5':
            return ['alpha', 'beta']
        elif dist_type == 'pearson6':
            return ['alpha1', 'alpha2', 'beta']
        elif dist_type == 'pareto2':
            return ['b', 'q']
        elif dist_type == 'pareto':
            return ['theta', 'alpha']
        elif dist_type == 'levy':
            return ['a', 'c']
        elif dist_type == 'erlang':
            return ['m', 'beta']
        elif dist_type == 'cauchy':
            return ['gamma', 'beta']
        elif dist_type == 'dagum':
            return ['gamma', 'beta', 'alpha1', 'alpha2']
        elif dist_type == 'doubletriang':
            return ['min_val', 'm_likely', 'max_val', 'lower_p']
        elif dist_type == 'negbin':
            return ['s', 'p']
        elif dist_type == 'geomet':
            return ['p']
        elif dist_type == 'hypergeo':
            return ['n', 'D', 'M']
        elif dist_type == 'intuniform':
            return ['min_val', 'max_val']
        elif dist_type == 'compound':
            return ['frequency', 'severity', 'deductible', 'limit']
        elif dist_type == 'splice':
            return ['left_dist', 'right_dist', 'splice_point']
        else:
            return []

    def _ppf(self, params_dict: Dict[str, np.ndarray], q: np.ndarray) -> np.ndarray:
        dist_type = self.dist_type
        n = self.n
        if dist_type == 'normal':
            from scipy.stats import norm
            return norm.ppf(q, loc=params_dict['mean'], scale=params_dict['std'])
        elif dist_type == 'uniform':
            a = params_dict['min']
            b = params_dict['max']
            return a + q * (b - a)
        elif dist_type == 'gamma':
            from scipy.stats import gamma
            return gamma.ppf(q, a=params_dict['shape'], scale=params_dict['scale'])
        elif dist_type == 'poisson':
            from scipy.stats import poisson
            return poisson.ppf(q, mu=params_dict['lam']).astype(float)
        elif dist_type == 'beta':
            from scipy.stats import beta
            return beta.ppf(q, a=params_dict['a'], b=params_dict['b'])
        elif dist_type == 'chisq':
            from scipy.stats import chi2
            return chi2.ppf(q, df=params_dict['df'])
        elif dist_type == 'f':
            from scipy.stats import f
            return f.ppf(q, dfn=params_dict['df1'], dfd=params_dict['df2'])
        elif dist_type == 'student':
            from scipy.stats import t
            return t.ppf(q, df=params_dict['df'])
        elif dist_type == 'expon':
            from scipy.stats import expon
            return expon.ppf(q, scale=params_dict['lam'])
        elif dist_type == 'bernoulli':
            p = params_dict['p']
            return np.where(q <= 1 - p, 0.0, 1.0)
        elif dist_type == 'triang':
            from index_functions import triang_ppf
            a, c, b = params_dict['a'], params_dict['c'], params_dict['b']
            return np.array([triang_ppf(qi, a[i], c[i], b[i]) for i, qi in enumerate(q)])
        elif dist_type == 'binomial':
            from index_functions import binomial_ppf
            n_param = params_dict['n'].astype(int)
            p = params_dict['p']
            return np.array([binomial_ppf(qi, n_param[i], p[i]) for i, qi in enumerate(q)])
        elif dist_type == 'invgauss':
            from index_functions import invgauss_ppf
            mu, lam = params_dict['mu'], params_dict['lam']
            return np.array([invgauss_ppf(qi, mu[i], lam[i]) for i, qi in enumerate(q)])
        elif dist_type == 'duniform':
            if self.x_vals_list is not None:
                from index_functions import duniform_ppf
                return np.array([duniform_ppf(qi, self.x_vals_list) for qi in q])
            else:
                return np.full_like(q, np.nan)
        elif dist_type == 'trigen':
            from index_functions import trigen_ppf
            L, M, U, alpha, beta = params_dict['L'], params_dict['M'], params_dict['U'], params_dict['alpha'], params_dict['beta']
            return np.array([trigen_ppf(qi, L[i], M[i], U[i], alpha[i], beta[i]) for i, qi in enumerate(q)])
        elif dist_type == 'cumul':
            if self.x_vals_list is not None and self.p_vals_list is not None:
                from index_functions import cumul_ppf
                return np.array([cumul_ppf(qi, self.x_vals_list, self.p_vals_list) for qi in q])
            else:
                return np.full_like(q, np.nan)
        elif dist_type == 'discrete':
            if self.x_vals_list is not None and self.p_vals_list is not None:
                from index_functions import discrete_ppf
                return np.array([discrete_ppf(qi, self.x_vals_list, self.p_vals_list) for qi in q])
            else:
                return np.full_like(q, np.nan)
        # 新增分布 PPF
        elif dist_type == 'erf':
            h = params_dict['h']
            return np.array([erf_ppf(qi, h[i]) for i, qi in enumerate(q)])
        elif dist_type == 'extvalue':
            A, B = params_dict['A'], params_dict['B']
            return np.array([extvalue_ppf(qi, A[i], B[i]) for i, qi in enumerate(q)])
        elif dist_type == 'extvaluemin':
            A, B = params_dict['A'], params_dict['B']
            return np.array([extvaluemin_ppf(qi, A[i], B[i]) for i, qi in enumerate(q)])
        elif dist_type == 'fatiguelife':
            y, beta, alpha = params_dict['y'], params_dict['beta'], params_dict['alpha']
            return np.array([fatiguelife_ppf(qi, y[i], beta[i], alpha[i]) for i, qi in enumerate(q)])
        elif dist_type == 'frechet':
            y, beta, alpha = params_dict['y'], params_dict['beta'], params_dict['alpha']
            return np.array([frechet_ppf(qi, y[i], beta[i], alpha[i]) for i, qi in enumerate(q)])
        elif dist_type == 'general':
            if self.x_vals_list is not None and self.p_vals_list is not None:
                min_val, max_val = params_dict['min'], params_dict['max']
                return np.array([general_ppf(qi, min_val[i], max_val[i], self.x_vals_list, self.p_vals_list) for i, qi in enumerate(q)])
            else:
                return np.full_like(q, np.nan)
        elif dist_type == 'histogrm':
            if self.p_vals_list is not None:
                min_val, max_val = params_dict['min'], params_dict['max']
                return np.array([histogrm_ppf(qi, min_val[i], max_val[i], self.p_vals_list) for i, qi in enumerate(q)])
            else:
                return np.full_like(q, np.nan)
        elif dist_type == 'hypsecant':
            gamma, beta = params_dict['gamma'], params_dict['beta']
            return np.array([hypsecant_ppf(qi, gamma[i], beta[i]) for i, qi in enumerate(q)])
        elif dist_type == 'johnsonsb':
            alpha1, alpha2, a, b = params_dict['alpha1'], params_dict['alpha2'], params_dict['a'], params_dict['b']
            return np.array([johnsonsb_ppf(qi, alpha1[i], alpha2[i], a[i], b[i]) for i, qi in enumerate(q)])
        elif dist_type == 'johnsonsu':
            alpha1, alpha2, gamma, beta = params_dict['alpha1'], params_dict['alpha2'], params_dict['gamma'], params_dict['beta']
            return np.array([johnsonsu_ppf(qi, alpha1[i], alpha2[i], gamma[i], beta[i]) for i, qi in enumerate(q)])
        elif dist_type == 'kumaraswamy':
            alpha1, alpha2, min_val, max_val = params_dict['alpha1'], params_dict['alpha2'], params_dict['min_val'], params_dict['max_val']
            return np.array([kumaraswamy_ppf(qi, alpha1[i], alpha2[i], min_val[i], max_val[i]) for i, qi in enumerate(q)])
        elif dist_type == 'laplace':
            mu, sigma = params_dict['mu'], params_dict['sigma']
            return np.array([laplace_ppf(qi, mu[i], sigma[i]) for i, qi in enumerate(q)])
        elif dist_type == 'logistic':
            alpha, beta = params_dict['alpha'], params_dict['beta']
            return np.array([logistic_ppf(qi, alpha[i], beta[i]) for i, qi in enumerate(q)])
        elif dist_type == 'loglogistic':
            gamma, beta, alpha = params_dict['gamma'], params_dict['beta'], params_dict['alpha']
            return np.array([loglogistic_ppf(qi, gamma[i], beta[i], alpha[i]) for i, qi in enumerate(q)])
        elif dist_type == 'lognorm':
            mu, sigma = params_dict['mu'], params_dict['sigma']
            return np.array([lognorm_ppf(qi, mu[i], sigma[i]) for i, qi in enumerate(q)])
        elif dist_type == 'lognorm2':
            mu, sigma = params_dict['mu'], params_dict['sigma']
            return np.array([lognorm2_ppf(qi, mu[i], sigma[i]) for i, qi in enumerate(q)])
        elif dist_type == 'betageneral':
            alpha1, alpha2, min_val, max_val = params_dict['alpha1'], params_dict['alpha2'], params_dict['min_val'], params_dict['max_val']
            return np.array([betageneral_ppf(qi, alpha1[i], alpha2[i], min_val[i], max_val[i]) for i, qi in enumerate(q)])
        elif dist_type == 'betasubj':
            min_val, m_likely, mean_val, max_val = params_dict['min_val'], params_dict['m_likely'], params_dict['mean_val'], params_dict['max_val']
            return np.array([betasubj_ppf(qi, min_val[i], m_likely[i], mean_val[i], max_val[i]) for i, qi in enumerate(q)])
        elif dist_type == 'burr12':
            gamma, beta, alpha1, alpha2 = params_dict['gamma'], params_dict['beta'], params_dict['alpha1'], params_dict['alpha2']
            return np.array([burr12_ppf(qi, gamma[i], beta[i], alpha1[i], alpha2[i]) for i, qi in enumerate(q)])
        elif dist_type == 'compound':
            # Compound 的 PPF 复杂，使用简单回退
            return np.full_like(q, np.nan)
        elif dist_type == 'splice':
            left_dist, right_dist, splice_point = params_dict['left_dist'], params_dict['right_dist'], params_dict['splice_point']
            # 简化：假设 splice_point 是标量，使用 splice_ppf
            point = splice_point[0] if isinstance(splice_point, np.ndarray) else splice_point
            left = left_dist[0] if isinstance(left_dist, np.ndarray) else left_dist
            right = right_dist[0] if isinstance(right_dist, np.ndarray) else right_dist
            return np.array([splice_ppf(qi, left, right, point) for qi in q])
        elif dist_type == 'pert':
            min_val, m_likely, max_val = params_dict['min_val'], params_dict['m_likely'], params_dict['max_val']
            return np.array([pert_ppf(qi, min_val[i], m_likely[i], max_val[i]) for i, qi in enumerate(q)])
        elif dist_type == 'reciprocal':
            min_val, max_val = params_dict['min_val'], params_dict['max_val']
            return np.array([reciprocal_ppf(qi, min_val[i], max_val[i]) for i, qi in enumerate(q)])
        elif dist_type == 'rayleigh':
            b = params_dict['b']
            return np.array([rayleigh_ppf(qi, b[i]) for i, qi in enumerate(q)])
        elif dist_type == 'weibull':
            alpha, beta = params_dict['alpha'], params_dict['beta']
            return np.array([weibull_ppf(qi, alpha[i], beta[i]) for i, qi in enumerate(q)])
        elif dist_type == 'pearson5':
            alpha, beta = params_dict['alpha'], params_dict['beta']
            return np.array([pearson5_ppf(qi, alpha[i], beta[i]) for i, qi in enumerate(q)])
        elif dist_type == 'pearson6':
            alpha1, alpha2, beta = params_dict['alpha1'], params_dict['alpha2'], params_dict['beta']
            return np.array([pearson6_ppf(qi, alpha1[i], alpha2[i], beta[i]) for i, qi in enumerate(q)])
        elif dist_type == 'pareto2':
            b, q_shape = params_dict['b'], params_dict['q']
            return np.array([pareto2_ppf(qi, b[i], q_shape[i]) for i, qi in enumerate(q)])
        elif dist_type == 'pareto':
            theta, alpha = params_dict['theta'], params_dict['alpha']
            return np.array([pareto_ppf(qi, theta[i], alpha[i]) for i, qi in enumerate(q)])
        elif dist_type == 'levy':
            a, c = params_dict['a'], params_dict['c']
            return np.array([levy_ppf(qi, a[i], c[i]) for i, qi in enumerate(q)])
        elif dist_type == 'erlang':
            m, beta = params_dict['m'], params_dict['beta']
            return np.array([erlang_ppf(qi, m[i], beta[i]) for i, qi in enumerate(q)])
        elif dist_type == 'cauchy':
            gamma, beta = params_dict['gamma'], params_dict['beta']
            return np.array([cauchy_ppf(qi, gamma[i], beta[i]) for i, qi in enumerate(q)])
        elif dist_type == 'dagum':
            gamma, beta, alpha1, alpha2 = params_dict['gamma'], params_dict['beta'], params_dict['alpha1'], params_dict['alpha2']
            return np.array([dagum_ppf(qi, gamma[i], beta[i], alpha1[i], alpha2[i]) for i, qi in enumerate(q)])
        elif dist_type == 'doubletriang':
            min_val, m_likely, max_val, lower_p = params_dict['min_val'], params_dict['m_likely'], params_dict['max_val'], params_dict['lower_p']
            return np.array([doubletriang_ppf(qi, min_val[i], m_likely[i], max_val[i], lower_p[i]) for i, qi in enumerate(q)])
        elif dist_type == 'negbin':
            s, p = params_dict['s'], params_dict['p']
            return np.array([negbin_ppf(qi, s[i], p[i]) for i, qi in enumerate(q)])
        elif dist_type == 'geomet':
            p = params_dict['p']
            return np.array([geomet_ppf(qi, p[i]) for i, qi in enumerate(q)])
        elif dist_type == 'hypergeo':
            n_param, D, M = params_dict['n'], params_dict['D'], params_dict['M']
            return np.array([hypergeo_ppf(qi, n_param[i], D[i], M[i]) for i, qi in enumerate(q)])
        elif dist_type == 'intuniform':
            min_val, max_val = params_dict['min_val'], params_dict['max_val']
            return np.array([intuniform_ppf(qi, min_val[i], max_val[i]) for i, qi in enumerate(q)])
        else:
            return np.full_like(q, np.nan)

    def _cdf(self, x: np.ndarray, params_dict: Dict[str, np.ndarray]) -> np.ndarray:
        dist_type = self.dist_type
        n = self.n
        if dist_type == 'normal':
            from scipy.stats import norm
            return norm.cdf(x, loc=params_dict['mean'], scale=params_dict['std'])
        elif dist_type == 'uniform':
            a = params_dict['min']
            b = params_dict['max']
            return np.clip((x - a) / (b - a), 0, 1)
        elif dist_type == 'gamma':
            from scipy.stats import gamma
            return gamma.cdf(x, a=params_dict['shape'], scale=params_dict['scale'])
        elif dist_type == 'poisson':
            from scipy.stats import poisson
            return poisson.cdf(x, mu=params_dict['lam'])
        elif dist_type == 'beta':
            from scipy.stats import beta
            return beta.cdf(x, a=params_dict['a'], b=params_dict['b'])
        elif dist_type == 'chisq':
            from scipy.stats import chi2
            return chi2.cdf(x, df=params_dict['df'])
        elif dist_type == 'f':
            from scipy.stats import f
            return f.cdf(x, dfn=params_dict['df1'], dfd=params_dict['df2'])
        elif dist_type == 'student':
            from scipy.stats import t
            return t.cdf(x, df=params_dict['df'])
        elif dist_type == 'expon':
            from scipy.stats import expon
            return expon.cdf(x, scale=params_dict['lam'])
        elif dist_type == 'bernoulli':
            p = params_dict['p']
            return np.where(x < 0, 0, np.where(x < 1, 1-p, 1))
        elif dist_type == 'triang':
            from index_functions import triang_cdf
            a, c, b = params_dict['a'], params_dict['c'], params_dict['b']
            return np.array([triang_cdf(x[i], a[i], c[i], b[i]) for i in range(n)])
        elif dist_type == 'binomial':
            from index_functions import binomial_cdf
            n_param = params_dict['n'].astype(int)
            p = params_dict['p']
            return np.array([binomial_cdf(x[i], n_param[i], p[i]) for i in range(n)])
        elif dist_type == 'invgauss':
            from index_functions import invgauss_cdf
            mu, lam = params_dict['mu'], params_dict['lam']
            return np.array([invgauss_cdf(x[i], mu[i], lam[i]) for i in range(n)])
        elif dist_type == 'duniform':
            if self.x_vals_list is not None:
                from index_functions import duniform_cdf
                return np.array([duniform_cdf(x[i], self.x_vals_list) for i in range(n)])
            else:
                return np.full(n, np.nan)
        elif dist_type == 'trigen':
            from index_functions import trigen_cdf
            L, M, U, alpha, beta = params_dict['L'], params_dict['M'], params_dict['U'], params_dict['alpha'], params_dict['beta']
            return np.array([trigen_cdf(x[i], L[i], M[i], U[i], alpha[i], beta[i]) for i in range(n)])
        elif dist_type == 'cumul':
            if self.x_vals_list is not None and self.p_vals_list is not None:
                from index_functions import cumul_cdf
                return np.array([cumul_cdf(x[i], self.x_vals_list, self.p_vals_list) for i in range(n)])
            else:
                return np.full(n, np.nan)
        elif dist_type == 'discrete':
            if self.x_vals_list is not None and self.p_vals_list is not None:
                from index_functions import discrete_cdf
                return np.array([discrete_cdf(x[i], self.x_vals_list, self.p_vals_list) for i in range(n)])
            else:
                return np.full(n, np.nan)
        # 新增分布 CDF
        elif dist_type == 'erf':
            h = params_dict['h']
            return np.array([erf_cdf(x[i], h[i]) for i in range(n)])
        elif dist_type == 'extvalue':
            A, B = params_dict['A'], params_dict['B']
            return np.array([extvalue_cdf(x[i], A[i], B[i]) for i in range(n)])
        elif dist_type == 'extvaluemin':
            A, B = params_dict['A'], params_dict['B']
            return np.array([extvaluemin_cdf(x[i], A[i], B[i]) for i in range(n)])
        elif dist_type == 'fatiguelife':
            y, beta, alpha = params_dict['y'], params_dict['beta'], params_dict['alpha']
            return np.array([fatiguelife_cdf(x[i], y[i], beta[i], alpha[i]) for i in range(n)])
        elif dist_type == 'frechet':
            y, beta, alpha = params_dict['y'], params_dict['beta'], params_dict['alpha']
            return np.array([frechet_cdf(x[i], y[i], beta[i], alpha[i]) for i in range(n)])
        elif dist_type == 'general':
            if self.x_vals_list is not None and self.p_vals_list is not None:
                min_val, max_val = params_dict['min'], params_dict['max']
                return np.array([general_cdf(x[i], min_val[i], max_val[i], self.x_vals_list, self.p_vals_list) for i in range(n)])
            else:
                return np.full(n, np.nan)
        elif dist_type == 'histogrm':
            if self.p_vals_list is not None:
                min_val, max_val = params_dict['min'], params_dict['max']
                return np.array([histogrm_cdf(x[i], min_val[i], max_val[i], self.p_vals_list) for i in range(n)])
            else:
                return np.full(n, np.nan)
        elif dist_type == 'hypsecant':
            gamma, beta = params_dict['gamma'], params_dict['beta']
            return np.array([hypsecant_cdf(x[i], gamma[i], beta[i]) for i in range(n)])
        elif dist_type == 'johnsonsb':
            alpha1, alpha2, a, b = params_dict['alpha1'], params_dict['alpha2'], params_dict['a'], params_dict['b']
            return np.array([johnsonsb_cdf(x[i], alpha1[i], alpha2[i], a[i], b[i]) for i in range(n)])
        elif dist_type == 'johnsonsu':
            alpha1, alpha2, gamma, beta = params_dict['alpha1'], params_dict['alpha2'], params_dict['gamma'], params_dict['beta']
            return np.array([johnsonsu_cdf(x[i], alpha1[i], alpha2[i], gamma[i], beta[i]) for i in range(n)])
        elif dist_type == 'kumaraswamy':
            alpha1, alpha2, min_val, max_val = params_dict['alpha1'], params_dict['alpha2'], params_dict['min_val'], params_dict['max_val']
            return np.array([kumaraswamy_cdf(x[i], alpha1[i], alpha2[i], min_val[i], max_val[i]) for i in range(n)])
        elif dist_type == 'laplace':
            mu, sigma = params_dict['mu'], params_dict['sigma']
            return np.array([laplace_cdf(x[i], mu[i], sigma[i]) for i in range(n)])
        elif dist_type == 'logistic':
            alpha, beta = params_dict['alpha'], params_dict['beta']
            return np.array([logistic_cdf(x[i], alpha[i], beta[i]) for i in range(n)])
        elif dist_type == 'loglogistic':
            gamma, beta, alpha = params_dict['gamma'], params_dict['beta'], params_dict['alpha']
            return np.array([loglogistic_cdf(x[i], gamma[i], beta[i], alpha[i]) for i in range(n)])
        elif dist_type == 'lognorm':
            mu, sigma = params_dict['mu'], params_dict['sigma']
            return np.array([lognorm_cdf(x[i], mu[i], sigma[i]) for i in range(n)])
        elif dist_type == 'lognorm2':
            mu, sigma = params_dict['mu'], params_dict['sigma']
            return np.array([lognorm2_cdf(x[i], mu[i], sigma[i]) for i in range(n)])
        elif dist_type == 'betageneral':
            alpha1, alpha2, min_val, max_val = params_dict['alpha1'], params_dict['alpha2'], params_dict['min_val'], params_dict['max_val']
            return np.array([betageneral_cdf(x[i], alpha1[i], alpha2[i], min_val[i], max_val[i]) for i in range(n)])
        elif dist_type == 'betasubj':
            min_val, m_likely, mean_val, max_val = params_dict['min_val'], params_dict['m_likely'], params_dict['mean_val'], params_dict['max_val']
            return np.array([betasubj_cdf(x[i], min_val[i], m_likely[i], mean_val[i], max_val[i]) for i in range(n)])
        elif dist_type == 'burr12':
            gamma, beta, alpha1, alpha2 = params_dict['gamma'], params_dict['beta'], params_dict['alpha1'], params_dict['alpha2']
            return np.array([burr12_cdf(x[i], gamma[i], beta[i], alpha1[i], alpha2[i]) for i in range(n)])
        elif dist_type == 'compound':
            return np.full(n, np.nan)
        elif dist_type == 'splice':
            left_dist, right_dist, splice_point = params_dict['left_dist'], params_dict['right_dist'], params_dict['splice_point']
            point = splice_point[0] if isinstance(splice_point, np.ndarray) else splice_point
            left = left_dist[0] if isinstance(left_dist, np.ndarray) else left_dist
            right = right_dist[0] if isinstance(right_dist, np.ndarray) else right_dist
            return np.array([splice_cdf(x[i], left, right, point) for i in range(n)])
        elif dist_type == 'pert':
            min_val, m_likely, max_val = params_dict['min_val'], params_dict['m_likely'], params_dict['max_val']
            return np.array([pert_cdf(x[i], min_val[i], m_likely[i], max_val[i]) for i in range(n)])
        elif dist_type == 'reciprocal':
            min_val, max_val = params_dict['min_val'], params_dict['max_val']
            return np.array([reciprocal_cdf(x[i], min_val[i], max_val[i]) for i in range(n)])
        elif dist_type == 'rayleigh':
            b = params_dict['b']
            return np.array([rayleigh_cdf(x[i], b[i]) for i in range(n)])
        elif dist_type == 'weibull':
            alpha, beta = params_dict['alpha'], params_dict['beta']
            return np.array([weibull_cdf(x[i], alpha[i], beta[i]) for i in range(n)])
        elif dist_type == 'pearson5':
            alpha, beta = params_dict['alpha'], params_dict['beta']
            return np.array([pearson5_cdf(x[i], alpha[i], beta[i]) for i in range(n)])
        elif dist_type == 'pearson6':
            alpha1, alpha2, beta = params_dict['alpha1'], params_dict['alpha2'], params_dict['beta']
            return np.array([pearson6_cdf(x[i], alpha1[i], alpha2[i], beta[i]) for i in range(n)])
        elif dist_type == 'pareto2':
            b, q = params_dict['b'], params_dict['q']
            return np.array([pareto2_cdf(x[i], b[i], q[i]) for i in range(n)])
        elif dist_type == 'pareto':
            theta, alpha = params_dict['theta'], params_dict['alpha']
            return np.array([pareto_cdf(x[i], theta[i], alpha[i]) for i in range(n)])
        elif dist_type == 'levy':
            a, c = params_dict['a'], params_dict['c']
            return np.array([levy_cdf(x[i], a[i], c[i]) for i in range(n)])
        elif dist_type == 'erlang':
            m, beta = params_dict['m'], params_dict['beta']
            return np.array([erlang_cdf(x[i], m[i], beta[i]) for i in range(n)])
        elif dist_type == 'cauchy':
            gamma, beta = params_dict['gamma'], params_dict['beta']
            return np.array([cauchy_cdf(x[i], gamma[i], beta[i]) for i in range(n)])
        elif dist_type == 'dagum':
            gamma, beta, alpha1, alpha2 = params_dict['gamma'], params_dict['beta'], params_dict['alpha1'], params_dict['alpha2']
            return np.array([dagum_cdf(x[i], gamma[i], beta[i], alpha1[i], alpha2[i]) for i in range(n)])
        elif dist_type == 'doubletriang':
            min_val, m_likely, max_val, lower_p = params_dict['min_val'], params_dict['m_likely'], params_dict['max_val'], params_dict['lower_p']
            return np.array([doubletriang_cdf(x[i], min_val[i], m_likely[i], max_val[i], lower_p[i]) for i in range(n)])
        elif dist_type == 'negbin':
            s, p = params_dict['s'], params_dict['p']
            return np.array([negbin_cdf(x[i], s[i], p[i]) for i in range(n)])
        elif dist_type == 'geomet':
            p = params_dict['p']
            return np.array([geomet_cdf(x[i], p[i]) for i in range(n)])
        elif dist_type == 'hypergeo':
            n_param, D, M = params_dict['n'], params_dict['D'], params_dict['M']
            return np.array([hypergeo_cdf(x[i], n_param[i], D[i], M[i]) for i in range(n)])
        elif dist_type == 'intuniform':
            min_val, max_val = params_dict['min_val'], params_dict['max_val']
            return np.array([intuniform_cdf(x[i], min_val[i], max_val[i]) for i in range(n)])
        else:
            return np.full(n, np.nan)

    def _generate_raw_samples(self, rng: np.random.Generator, params_dict: Dict[str, np.ndarray]) -> np.ndarray:
        n = self.n
        dist_type = self.dist_type
        try:
            if dist_type == 'normal':
                mean = params_dict['mean']
                std = params_dict['std']
                return rng.normal(mean, std, n)
            elif dist_type == 'uniform':
                low = params_dict['min']
                high = params_dict['max']
                return rng.uniform(low, high, n)
            elif dist_type == 'gamma':
                shape = params_dict['shape']
                scale = params_dict['scale']
                return rng.gamma(shape, scale, n)
            elif dist_type == 'poisson':
                lam = params_dict['lam']
                return rng.poisson(lam, n).astype(float)
            elif dist_type == 'beta':
                a = params_dict['a']
                b = params_dict['b']
                return rng.beta(a, b, n)
            elif dist_type == 'chisq':
                df = params_dict['df']
                return rng.chisquare(df, n)
            elif dist_type == 'f':
                df1 = params_dict['df1']
                df2 = params_dict['df2']
                return rng.f(df1, df2, n)
            elif dist_type == 'student':
                df = params_dict['df']
                return rng.standard_t(df, n)
            elif dist_type == 'expon':
                lam = params_dict['lam']
                return rng.exponential(lam, n)
            elif dist_type == 'bernoulli':
                p = params_dict['p']
                return bernoulli_generator_vectorized(rng, [p], n)
            elif dist_type == 'triang':
                a, c, b = params_dict['a'], params_dict['c'], params_dict['b']
                return triang_generator_vectorized(rng, [a, c, b], n)
            elif dist_type == 'binomial':
                n_param = params_dict['n'].astype(int)
                p = params_dict['p']
                return binomial_generator_vectorized(rng, [n_param, p], n)
            elif dist_type == 'invgauss':
                mu, lam = params_dict['mu'], params_dict['lam']
                return invgauss_generator_vectorized(rng, [mu, lam], n)
            elif dist_type == 'duniform':
                if self.x_vals_list is not None and self.p_vals_list is not None:
                    try:
                        # 确保概率和为 1
                        p_sum = sum(self.p_vals_list)
                        if abs(p_sum - 1.0) > 1e-10:
                            print(f"DUniform 概率和不为1: {p_sum}, 重新归一化")
                            self.p_vals_list = [p / p_sum for p in self.p_vals_list]
                        # 使用 numpy 的 choice 进行向量化采样
                        indices = rng.choice(len(self.x_vals_list), size=n, p=self.p_vals_list)
                        samples = np.array([self.x_vals_list[i] for i in indices], dtype=float)
                        print(f"DUniform 生成样本成功，前5个: {samples[:5]}")
                        return samples
                    except Exception as e:
                        print(f"DUniform 生成样本异常: {e}", exc_info=True)
                        return np.full(n, ERROR_MARKER, dtype=object)
                else:
                    print(f"DUniform x_vals_list 或 p_vals_list 为空: x_vals={self.x_vals_list}, p_vals={self.p_vals_list}")
                    return np.full(n, ERROR_MARKER, dtype=object)
            elif dist_type == 'trigen':
                L, M, U, alpha, beta = params_dict['L'], params_dict['M'], params_dict['U'], params_dict['alpha'], params_dict['beta']
                return trigen_generator_vectorized(rng, [L, M, U, alpha, beta], n)
            elif dist_type == 'cumul':
                if self.x_vals_list is not None and self.p_vals_list is not None:
                    min_val = params_dict.get('min', np.full(n, 0.0))
                    max_val = params_dict.get('max', np.full(n, 1.0))
                    return cumul_generator_vectorized(rng, [min_val, max_val], n, self.x_vals_list, self.p_vals_list)
                else:
                    return np.full(n, ERROR_MARKER, dtype=object)
            elif dist_type == 'discrete':
                if self.x_vals_list is not None and self.p_vals_list is not None:
                    return discrete_generator_vectorized(rng, [], n, self.x_vals_list, self.p_vals_list)
                else:
                    return np.full(n, ERROR_MARKER, dtype=object)
            # 新增分布生成器
            elif dist_type == 'erf':
                h = params_dict['h']
                return erf_generator_vectorized(rng, [h], n)
            elif dist_type == 'extvalue':
                A, B = params_dict['A'], params_dict['B']
                return extvalue_generator_vectorized(rng, [A, B], n)
            elif dist_type == 'extvaluemin':
                A, B = params_dict['A'], params_dict['B']
                return extvaluemin_generator_vectorized(rng, [A, B], n)
            elif dist_type == 'fatiguelife':
                y, beta, alpha = params_dict['y'], params_dict['beta'], params_dict['alpha']
                return fatiguelife_generator_vectorized(rng, [y, beta, alpha], n)
            elif dist_type == 'frechet':
                y, beta, alpha = params_dict['y'], params_dict['beta'], params_dict['alpha']
                return frechet_generator_vectorized(rng, [y, beta, alpha], n)
            elif dist_type == 'general':
                if self.x_vals_list is not None and self.p_vals_list is not None:
                    # 获取 min 和 max 数组
                    if 'min' in params_dict and 'max' in params_dict:
                        min_arr = params_dict['min']
                        max_arr = params_dict['max']
                    elif 'min_val' in params_dict and 'max_val' in params_dict:
                        min_arr = params_dict['min_val']
                        max_arr = params_dict['max_val']
                    else:
                        # 后备：从 markers 中读取（不应发生）
                        min_arr = np.full(n, float(self.markers.get('min', -1.0)))
                        max_arr = np.full(n, float(self.markers.get('max', 1.0)))
                    
                    # 检查是否为常量数组（所有迭代相同）
                    if isinstance(min_arr, np.ndarray) and len(min_arr) > 1:
                        if np.all(min_arr == min_arr[0]) and np.all(max_arr == max_arr[0]):
                            # 常量情况：转换为标量，使用向量化生成器
                            min_scalar = min_arr[0]
                            max_scalar = max_arr[0]
                            return general_generator_vectorized(rng, [min_scalar, max_scalar], n, self.x_vals_list, self.p_vals_list)
                        else:
                            # 非常量情况：回退到逐迭代循环（因为 general_generator_vectorized 不支持数组边界）
                            samples = np.zeros(n, dtype=float)
                            for i in range(n):
                                samples[i] = general_generator_vectorized(rng, [min_arr[i], max_arr[i]], 1, self.x_vals_list, self.p_vals_list)[0]
                            return samples
                    else:
                        # 已经是标量或长度1的数组
                        return general_generator_vectorized(rng, [min_arr, max_arr], n, self.x_vals_list, self.p_vals_list)
                else:
                    return np.full(n, ERROR_MARKER, dtype=object)
            elif dist_type == 'histogrm':
                if self.p_vals_list is None:
                    print("Histogrm: p_vals_list 为空，无法生成样本")
                    return np.full(self.n, ERROR_MARKER, dtype=object)
                
                # 检查 params_dict 中是否存在 'min' 和 'max'
                if 'min' not in params_dict or 'max' not in params_dict:
                    print(f"Histogrm 参数缺失: min={params_dict.get('min')}, max={params_dict.get('max')}")
                    return np.full(self.n, ERROR_MARKER, dtype=object)
                
                min_val = params_dict['min']
                max_val = params_dict['max']
                
                # 处理数组参数（非常量边界）
                if isinstance(min_val, np.ndarray) and len(min_val) > 1:
                    if np.all(min_val == min_val[0]) and np.all(max_val == max_val[0]):
                        # 常量情况：使用向量化生成器
                        min_scalar = min_val[0]
                        max_scalar = max_val[0]
                        return histogrm_generator_vectorized(rng, [min_scalar, max_scalar], self.n, self.p_vals_list)
                    else:
                        # 非常量情况：逐迭代循环
                        samples = np.zeros(self.n, dtype=float)
                        for i in range(self.n):
                            samples[i] = histogrm_generator_vectorized(rng, [min_val[i], max_val[i]], 1, self.p_vals_list)[0]
                        return samples
                else:
                    # 标量或长度1的数组
                    return histogrm_generator_vectorized(rng, [min_val, max_val], self.n, self.p_vals_list)
            elif dist_type == 'hypsecant':
                gamma, beta = params_dict['gamma'], params_dict['beta']
                return hypsecant_generator_vectorized(rng, [gamma, beta], n)
            elif dist_type == 'johnsonsb':
                alpha1, alpha2, a, b = params_dict['alpha1'], params_dict['alpha2'], params_dict['a'], params_dict['b']
                return johnsonsb_generator_vectorized(rng, [alpha1, alpha2, a, b], n)
            elif dist_type == 'johnsonsu':
                alpha1, alpha2, gamma, beta = params_dict['alpha1'], params_dict['alpha2'], params_dict['gamma'], params_dict['beta']
                return johnsonsu_generator_vectorized(rng, [alpha1, alpha2, gamma, beta], n)
            elif dist_type == 'kumaraswamy':
                alpha1, alpha2, min_val, max_val = params_dict['alpha1'], params_dict['alpha2'], params_dict['min_val'], params_dict['max_val']
                return kumaraswamy_generator_vectorized(rng, [alpha1, alpha2, min_val, max_val], n)
            elif dist_type == 'laplace':
                mu, sigma = params_dict['mu'], params_dict['sigma']
                return laplace_generator_vectorized(rng, [mu, sigma], n)
            elif dist_type == 'logistic':
                alpha, beta = params_dict['alpha'], params_dict['beta']
                return logistic_generator_vectorized(rng, [alpha, beta], n)
            elif dist_type == 'loglogistic':
                gamma, beta, alpha = params_dict['gamma'], params_dict['beta'], params_dict['alpha']
                return loglogistic_generator_vectorized(rng, [gamma, beta, alpha], n)
            elif dist_type == 'lognorm':
                mu, sigma = params_dict['mu'], params_dict['sigma']
                return lognorm_generator_vectorized(rng, [mu, sigma], n)
            elif dist_type == 'lognorm2':
                mu, sigma = params_dict['mu'], params_dict['sigma']
                return lognorm2_generator_vectorized(rng, [mu, sigma], n)
            elif dist_type == 'betageneral':
                alpha1, alpha2, min_val, max_val = params_dict['alpha1'], params_dict['alpha2'], params_dict['min_val'], params_dict['max_val']
                return betageneral_generator_vectorized(rng, [alpha1, alpha2, min_val, max_val], n)
            elif dist_type == 'betasubj':
                min_val, m_likely, mean_val, max_val = params_dict['min_val'], params_dict['m_likely'], params_dict['mean_val'], params_dict['max_val']
                return betasubj_generator_vectorized(rng, [min_val, m_likely, mean_val, max_val], n)
            elif dist_type == 'burr12':
                gamma, beta, alpha1, alpha2 = params_dict['gamma'], params_dict['beta'], params_dict['alpha1'], params_dict['alpha2']
                return burr12_generator_vectorized(rng, [gamma, beta, alpha1, alpha2], n)
            elif dist_type == 'compound':
                # 从 params_dict 中获取参数
                # 注意：params_dict 中的键根据 _get_param_names 应为 ['frequency', 'severity', 'deductible', 'limit']
                frequency_formula = params_dict.get('frequency')
                severity_formula = params_dict.get('severity')
                deductible = params_dict.get('deductible', 0.0)
                limit = params_dict.get('limit', float('inf'))
                # 参数可能是标量或数组，取第一个元素（因为 Compound 的公式字符串是常量）
                freq = frequency_formula[0] if isinstance(frequency_formula, np.ndarray) else frequency_formula
                sev = severity_formula[0] if isinstance(severity_formula, np.ndarray) else severity_formula
                ded = deductible[0] if isinstance(deductible, np.ndarray) else deductible
                lim = limit[0] if isinstance(limit, np.ndarray) else limit
                # 调用向量化生成器，需要传入 static_values, live_arrays 等，但这里没有这些上下文
                # 因此不能直接调用，需要重新设计。为了修复，我们改为逐迭代循环（性能可接受）
                # 注意：在向量化模拟中，Compound 的生成不应该在 _generate_raw_samples 中，
                # 而应该由更上层的 _sample_drisk_live 处理。此处保留原有标量方式，但标记为不推荐。
                # 实际上，我们已经在 _sample_drisk_live 中对 Compound 做了特殊处理，不会走到这里。
                # 所以这里保持原样，但为了安全，我们调用原有的标量生成器。
                from dist_compound import compound_generator_single
                samples = np.zeros(n, dtype=float)
                for i in range(n):
                    samples[i] = compound_generator_single(rng, [freq, sev, ded, lim])
                return samples
            elif dist_type == 'splice':
                left_dist = params_dict.get('left_dist')
                right_dist = params_dict.get('right_dist')
                splice_point = params_dict.get('splice_point')
                # 确保 left_dist 和 right_dist 是字符串（公式）
                if isinstance(left_dist, np.ndarray):
                    left_dist = left_dist[0] if left_dist.size > 0 else ""
                if isinstance(right_dist, np.ndarray):
                    right_dist = right_dist[0] if right_dist.size > 0 else ""
                if isinstance(splice_point, np.ndarray):
                    splice_point = splice_point[0] if splice_point.size > 0 else 0.0
                from dist_splice import splice_generator_vectorized
                return splice_generator_vectorized(rng, [left_dist, right_dist, splice_point], n)
            elif dist_type == 'pert':
                min_val, m_likely, max_val = params_dict['min_val'], params_dict['m_likely'], params_dict['max_val']
                return pert_generator_vectorized(rng, [min_val, m_likely, max_val], n)
            elif dist_type == 'reciprocal':
                min_val, max_val = params_dict['min_val'], params_dict['max_val']
                return reciprocal_generator_vectorized(rng, [min_val, max_val], n)
            elif dist_type == 'rayleigh':
                b = params_dict['b']
                return rayleigh_generator_vectorized(rng, [b], n)
            elif dist_type == 'weibull':
                alpha, beta = params_dict['alpha'], params_dict['beta']
                return weibull_generator_vectorized(rng, [alpha, beta], n)
            elif dist_type == 'pearson5':
                alpha, beta = params_dict['alpha'], params_dict['beta']
                return pearson5_generator_vectorized(rng, [alpha, beta], n)
            elif dist_type == 'pearson6':
                alpha1, alpha2, beta = params_dict['alpha1'], params_dict['alpha2'], params_dict['beta']
                return pearson6_generator_vectorized(rng, [alpha1, alpha2, beta], n)
            elif dist_type == 'pareto2':
                b, q = params_dict['b'], params_dict['q']
                return pareto2_generator_vectorized(rng, [b, q], n)
            elif dist_type == 'pareto':
                theta, alpha = params_dict['theta'], params_dict['alpha']
                return pareto_generator_vectorized(rng, [theta, alpha], n)
            elif dist_type == 'levy':
                a, c = params_dict['a'], params_dict['c']
                return levy_generator_vectorized(rng, [a, c], n)
            elif dist_type == 'erlang':
                m, beta = params_dict['m'], params_dict['beta']
                return erlang_generator_vectorized(rng, [m, beta], n)
            elif dist_type == 'cauchy':
                gamma, beta = params_dict['gamma'], params_dict['beta']
                return cauchy_generator_vectorized(rng, [gamma, beta], n)
            elif dist_type == 'dagum':
                gamma, beta, alpha1, alpha2 = params_dict['gamma'], params_dict['beta'], params_dict['alpha1'], params_dict['alpha2']
                return dagum_generator_vectorized(rng, [gamma, beta, alpha1, alpha2], n)
            elif dist_type == 'doubletriang':
                min_val, m_likely, max_val, lower_p = params_dict['min_val'], params_dict['m_likely'], params_dict['max_val'], params_dict['lower_p']
                return doubletriang_generator_vectorized(rng, [min_val, m_likely, max_val, lower_p], n)
            elif dist_type == 'negbin':
                s, p = params_dict['s'], params_dict['p']
                return negbin_generator_vectorized(rng, [s, p], n)
            elif dist_type == 'geomet':
                p = params_dict['p']
                return geomet_generator_vectorized(rng, [p], n)
            elif dist_type == 'hypergeo':
                n_param, D, M = params_dict['n'], params_dict['D'], params_dict['M']
                return hypergeo_generator_vectorized(rng, [n_param, D, M], n)
            elif dist_type == 'intuniform':
                min_val, max_val = params_dict['min_val'], params_dict['max_val']
                return intuniform_generator_vectorized(rng, [min_val, max_val], n)
            else:
                return np.full(n, ERROR_MARKER, dtype=object)
        except Exception as e:
            print(f"[_generate_raw_samples] 生成样本异常: {e}, dist_type={dist_type}")
            return np.full(n, ERROR_MARKER, dtype=object)

    def _generate_truncated_samples(self, rng: np.random.Generator, params_dict: Dict[str, np.ndarray],
                                    lower: Optional[np.ndarray], upper: Optional[np.ndarray]) -> np.ndarray:
        """
        生成截断分布样本（向量化）。

        参数：
            rng: NumPy 随机数生成器
            params_dict: 分布参数字典
            lower: 截断下界（已根据截断类型调整，可能是平移后的值或原始值）
            upper: 截断上界（同上）

        返回：
            样本数组，长度 n
        """
        n = self.n

        # 对于先平移后截断的两种类型，需要将边界还原到原始尺度再计算 CDF
        if self.truncate_type in ('truncate2', 'truncatep2'):
            # 边界是平移后的值 -> 减去 shift 得到原始尺度上的边界
            orig_lower = lower - self.shift_amount if lower is not None else None
            orig_upper = upper - self.shift_amount if upper is not None else None
        else:
            # 其他截断类型（truncate, truncatep）边界已经是原始尺度上的值
            orig_lower = lower
            orig_upper = upper

        # 计算原始分布下边界对应的 CDF 值
        low_cdf = np.zeros(n)
        high_cdf = np.ones(n)
        if orig_lower is not None:
            low_cdf = self._cdf(orig_lower, params_dict)
        if orig_upper is not None:
            high_cdf = self._cdf(orig_upper, params_dict)

        # 防止无效区间（理论上不会发生，但安全处理）
        invalid_mask = low_cdf >= high_cdf
        if np.any(invalid_mask):
            mid_cdf = (low_cdf + high_cdf) / 2
            low_cdf = np.where(invalid_mask, mid_cdf, low_cdf)
            high_cdf = np.where(invalid_mask, mid_cdf + 1e-12, high_cdf)

        # 在截断的 CDF 区间内均匀采样
        u = rng.uniform(0, 1, n)
        q = low_cdf + u * (high_cdf - low_cdf)

        # 通过 PPF 转换为原始尺度上的样本
        samples = self._ppf(params_dict, q)

        # 对于无效区间（截断区间退化为点），强制使用中位数
        if np.any(invalid_mask):
            q_mid = mid_cdf[invalid_mask]
            # 构建子字典，仅包含无效索引对应的参数
            sub_dict = {}
            for k, v in params_dict.items():
                if isinstance(v, np.ndarray) and len(v) == n:
                    sub_dict[k] = v[invalid_mask]
                else:
                    sub_dict[k] = v
            samples_mid = self._ppf(sub_dict, q_mid)
            samples[invalid_mask] = samples_mid

        # 最后统一加上平移量（所有截断类型均适用）
        return samples + self.shift_amount

    def _calculate_loc_value(self, params_dict: Dict[str, np.ndarray], valid_mask: np.ndarray) -> np.ndarray:
        n = self.n
        if not DIST_CLASSES_AVAILABLE:
            raw_mean = self._raw_mean(params_dict)
            result = raw_mean + self.shift_amount
            result[~valid_mask] = ERROR_MARKER
            return result.astype(object)

        dist_class = self.DIST_CLASS_MAP.get(self.dist_type)
        if dist_class is None:
            raw_mean = self._raw_mean(params_dict)
            result = raw_mean + self.shift_amount
            result[~valid_mask] = ERROR_MARKER
            return result.astype(object)

        result = np.full(n, ERROR_MARKER, dtype=object)
        valid_indices = np.where(valid_mask)[0]
        if len(valid_indices) == 0:
            return result

        param_names = self._get_param_names()
        param_arrays = [params_dict[name] for name in param_names if name in params_dict]

        for idx in valid_indices:
            try:
                params = [arr[idx] for arr in param_arrays]
                dist_obj = dist_class(params, self.markers, self.func_name)
                if not dist_obj.is_valid():
                    result[idx] = ERROR_MARKER
                else:
                    mean_val = dist_obj.mean()
                    result[idx] = float(mean_val)
            except Exception as e:
                if DEBUG:
                    print(f"计算 loc 值失败 at index {idx}: {e}")
                result[idx] = ERROR_MARKER
        return result

    def _raw_mean(self, params_dict: Dict[str, np.ndarray]) -> np.ndarray:
        dist_type = self.dist_type
        n = self.n
        if dist_type == 'normal':
            return params_dict['mean']
        elif dist_type == 'uniform':
            return (params_dict['min'] + params_dict['max']) / 2
        elif dist_type == 'gamma':
            return params_dict['shape'] * params_dict['scale']
        elif dist_type == 'poisson':
            return params_dict['lam']
        elif dist_type == 'beta':
            a, b = params_dict['a'], params_dict['b']
            return a / (a + b)
        elif dist_type == 'chisq':
            return params_dict['df']
        elif dist_type == 'f':
            df2 = params_dict['df2']
            return np.where(df2 > 2, df2 / (df2 - 2), np.inf)
        elif dist_type == 'student':
            return np.zeros(n)
        elif dist_type == 'expon':
            return params_dict['lam']
        elif dist_type == 'bernoulli':
            return params_dict['p']
        elif dist_type == 'triang':
            a, c, b = params_dict['a'], params_dict['c'], params_dict['b']
            return (a + c + b) / 3
        elif dist_type == 'binomial':
            return params_dict['n'] * params_dict['p']
        elif dist_type == 'invgauss':
            return params_dict['mu']
        elif dist_type == 'duniform':
            if self.x_vals_list is not None:
                return np.full(n, sum(self.x_vals_list) / len(self.x_vals_list))
            else:
                return np.full(n, 0.0)
        elif dist_type == 'trigen':
            from index_functions import _convert_trigen_to_triang
            L, M, U, alpha, beta = params_dict['L'], params_dict['M'], params_dict['U'], params_dict['alpha'], params_dict['beta']
            a_arr, c_arr, b_arr = _convert_trigen_to_triang(L, M, U, alpha, beta)
            return (a_arr + c_arr + b_arr) / 3
        elif dist_type == 'cumul':
            if self.x_vals_list is not None and self.p_vals_list is not None:
                return np.full(n, sum(x * p for x, p in zip(self.x_vals_list, self.p_vals_list)))
            else:
                return np.full(n, 0.0)
        elif dist_type == 'discrete':
            if self.x_vals_list is not None and self.p_vals_list is not None:
                return np.full(n, sum(x * p for x, p in zip(self.x_vals_list, self.p_vals_list)))
            else:
                return np.full(n, 0.0)
        # 新增分布均值
        elif dist_type == 'erf':
            return np.zeros(n)
        elif dist_type == 'extvalue':
            A, B = params_dict['A'], params_dict['B']
            return extvalue_raw_mean(A, B)
        elif dist_type == 'extvaluemin':
            A, B = params_dict['A'], params_dict['B']
            return extvaluemin_raw_mean(A, B)
        elif dist_type == 'fatiguelife':
            y, beta, alpha = params_dict['y'], params_dict['beta'], params_dict['alpha']
            return fatiguelife_raw_mean(y, beta, alpha)
        elif dist_type == 'frechet':
            y, beta, alpha = params_dict['y'], params_dict['beta'], params_dict['alpha']
            return frechet_raw_mean(y, beta, alpha)
        elif dist_type == 'general':
            if self.x_vals_list is not None and self.p_vals_list is not None:
                min_val, max_val = params_dict['min'], params_dict['max']
                return general_raw_mean(min_val, max_val, self.x_vals_list, self.p_vals_list)
            else:
                return np.full(n, 0.0)
        elif dist_type == 'histogrm':
            if self.p_vals_list is not None:
                min_val, max_val = params_dict['min'], params_dict['max']
                return histogrm_raw_mean(min_val, max_val, self.p_vals_list)
            else:
                return np.full(n, 0.0)
        elif dist_type == 'hypsecant':
            gamma, beta = params_dict['gamma'], params_dict['beta']
            return hypsecant_raw_mean(gamma, beta)
        elif dist_type == 'johnsonsb':
            alpha1, alpha2, a, b = params_dict['alpha1'], params_dict['alpha2'], params_dict['a'], params_dict['b']
            return johnsonsb_raw_mean(alpha1, alpha2, a, b)
        elif dist_type == 'johnsonsu':
            alpha1, alpha2, gamma, beta = params_dict['alpha1'], params_dict['alpha2'], params_dict['gamma'], params_dict['beta']
            return johnsonsu_raw_mean(alpha1, alpha2, gamma, beta)
        elif dist_type == 'kumaraswamy':
            alpha1, alpha2, min_val, max_val = params_dict['alpha1'], params_dict['alpha2'], params_dict['min_val'], params_dict['max_val']
            return kumaraswamy_raw_mean(alpha1, alpha2, min_val, max_val)
        elif dist_type == 'laplace':
            mu, sigma = params_dict['mu'], params_dict['sigma']
            return laplace_raw_mean(mu, sigma)
        elif dist_type == 'logistic':
            alpha, beta = params_dict['alpha'], params_dict['beta']
            return logistic_raw_mean(alpha, beta)
        elif dist_type == 'loglogistic':
            gamma, beta, alpha = params_dict['gamma'], params_dict['beta'], params_dict['alpha']
            return loglogistic_raw_mean(gamma, beta, alpha)
        elif dist_type == 'lognorm':
            mu, sigma = params_dict['mu'], params_dict['sigma']
            return lognorm_raw_mean(mu, sigma)
        elif dist_type == 'lognorm2':
            mu, sigma = params_dict['mu'], params_dict['sigma']
            return lognorm2_raw_mean(mu, sigma)
        elif dist_type == 'betageneral':
            alpha1, alpha2, min_val, max_val = params_dict['alpha1'], params_dict['alpha2'], params_dict['min_val'], params_dict['max_val']
            return betageneral_raw_mean(alpha1, alpha2, min_val, max_val)
        elif dist_type == 'betasubj':
            min_val, m_likely, mean_val, max_val = params_dict['min_val'], params_dict['m_likely'], params_dict['mean_val'], params_dict['max_val']
            return betasubj_raw_mean(min_val, m_likely, mean_val, max_val)
        elif dist_type == 'burr12':
            gamma, beta, alpha1, alpha2 = params_dict['gamma'], params_dict['beta'], params_dict['alpha1'], params_dict['alpha2']
            return burr12_raw_mean(gamma, beta, alpha1, alpha2)
        elif dist_type == 'compound':
            deductible = params_dict.get('deductible', 0.0)
            limit = params_dict.get('limit', float('inf'))
            frequency = params_dict['frequency']
            severity = params_dict['severity']
            # 简化处理：取第一个元素
            freq = frequency[0] if isinstance(frequency, np.ndarray) else frequency
            sev = severity[0] if isinstance(severity, np.ndarray) else severity
            ded = deductible[0] if isinstance(deductible, np.ndarray) else deductible
            lim = limit[0] if isinstance(limit, np.ndarray) else limit
            return np.full(n, compound_raw_mean(freq, sev, ded, lim))
        elif dist_type == 'splice':
            left_dist, right_dist, splice_point = params_dict['left_dist'], params_dict['right_dist'], params_dict['splice_point']
            left = left_dist[0] if isinstance(left_dist, np.ndarray) else left_dist
            right = right_dist[0] if isinstance(right_dist, np.ndarray) else right_dist
            point = splice_point[0] if isinstance(splice_point, np.ndarray) else splice_point
            return np.full(n, splice_raw_mean(left, right, point))
        elif dist_type == 'pert':
            min_val, m_likely, max_val = params_dict['min_val'], params_dict['m_likely'], params_dict['max_val']
            return pert_raw_mean(min_val, m_likely, max_val)
        elif dist_type == 'reciprocal':
            min_val, max_val = params_dict['min_val'], params_dict['max_val']
            return reciprocal_raw_mean(min_val, max_val)
        elif dist_type == 'rayleigh':
            b = params_dict['b']
            return rayleigh_raw_mean(b)
        elif dist_type == 'weibull':
            alpha, beta = params_dict['alpha'], params_dict['beta']
            return weibull_raw_mean(alpha, beta)
        elif dist_type == 'pearson5':
            alpha, beta = params_dict['alpha'], params_dict['beta']
            return pearson5_raw_mean(alpha, beta)
        elif dist_type == 'pearson6':
            alpha1, alpha2, beta = params_dict['alpha1'], params_dict['alpha2'], params_dict['beta']
            return pearson6_raw_mean(alpha1, alpha2, beta)
        elif dist_type == 'pareto2':
            b, q = params_dict['b'], params_dict['q']
            return pareto2_raw_mean(b, q)
        elif dist_type == 'pareto':
            theta, alpha = params_dict['theta'], params_dict['alpha']
            return pareto_raw_mean(theta, alpha)
        elif dist_type == 'levy':
            a, c = params_dict['a'], params_dict['c']
            return levy_raw_mean(a, c)
        elif dist_type == 'erlang':
            m, beta = params_dict['m'], params_dict['beta']
            # m 必须是整数
            return (np.round(m) * beta).astype(float)
        elif dist_type == 'cauchy':
            return np.full(n, np.nan)
        elif dist_type == 'dagum':
            gamma, beta, alpha1, alpha2 = params_dict['gamma'], params_dict['beta'], params_dict['alpha1'], params_dict['alpha2']
            return dagum_raw_mean(gamma, beta, alpha1, alpha2)
        elif dist_type == 'doubletriang':
            min_val, m_likely, max_val, lower_p = params_dict['min_val'], params_dict['m_likely'], params_dict['max_val'], params_dict['lower_p']
            return doubletriang_raw_mean(min_val, m_likely, max_val, lower_p)
        elif dist_type == 'negbin':
            s, p = params_dict['s'], params_dict['p']
            return s * (1.0 - p) / p
        elif dist_type == 'geomet':
            p = params_dict['p']
            return (1.0 / p) - 1.0
        elif dist_type == 'hypergeo':
            n_param, D, M = params_dict['n'], params_dict['D'], params_dict['M']
            return n_param * D / M
        elif dist_type == 'intuniform':
            min_val, max_val = params_dict['min_val'], params_dict['max_val']
            return (min_val + max_val) / 2.0
        else:
            return np.full(n, 0.0)

    def generate_samples(self, rng_seed: int, param_arrays: List[np.ndarray]) -> np.ndarray:
        """
        生成 n 个样本，支持蒙特卡洛、拉丁超立方、Sobol 序列。

        Args:
            rng_seed: 随机种子
            param_arrays: 参数数组列表（每个参数对应一个长度为 n 的数组）

        Returns:
            长度为 n 的样本数组（数值或错误标记）
        """
        n = self.n

        # 专门处理 Compound 和 Splice
        if self.dist_type in ('compound', 'splice'):
            # 需要 static_values, live_arrays 等上下文，但这些在生成器中不可用。
            # 因此，Compound/Splice 的向量化生成应该在 _sample_drisk_live 中完成，不会进入这里。
            # 为安全起见，如果进入这里，抛出异常或回退到循环。
            rng = np.random.default_rng(rng_seed)
            if self.dist_type == 'compound':
                if len(param_arrays) < 3:
                    raise ValueError("Compound 分布需要至少 3 个参数：频率、严重性、免赔额")
                freq_arr = param_arrays[0]
                sev_arr = param_arrays[1]
                ded_arr = param_arrays[2] if len(param_arrays) > 2 else np.full(n, 0.0)
                lim_arr = param_arrays[3] if len(param_arrays) > 3 else np.full(n, float('inf'))
                # 假设所有参数都是标量（字符串相同），取第一个
                freq = freq_arr[0] if freq_arr is not None else ""
                sev = sev_arr[0] if sev_arr is not None else ""
                ded = ded_arr[0] if ded_arr is not None else 0.0
                lim = lim_arr[0] if lim_arr is not None else float('inf')
                from dist_compound import compound_generator_single
                samples = np.zeros(n, dtype=float)
                for i in range(n):
                    samples[i] = compound_generator_single(rng, [freq, sev, ded, lim])
                return samples
            elif self.dist_type == 'splice':
                left_arr = param_arrays[0] if len(param_arrays) > 0 else None
                right_arr = param_arrays[1] if len(param_arrays) > 1 else None
                point_arr = param_arrays[2] if len(param_arrays) > 2 else None
                left = left_arr[0] if left_arr is not None else ""
                right = right_arr[0] if right_arr is not None else ""
                point = point_arr[0] if point_arr is not None else 0.0
                from dist_splice import splice_generator_single
                samples = np.zeros(n, dtype=float)
                for i in range(n):
                    samples[i] = splice_generator_single(rng, [left, right, point])
                return samples

        err_valid_mask = self._handle_error_markers(param_arrays)
        if err_valid_mask is not None and not np.any(err_valid_mask):
            return np.full(n, ERROR_MARKER, dtype=object)

        param_names = self._get_param_names()
        params_dict = {}
        for i, name in enumerate(param_names):
            if i < len(param_arrays):
                arr = param_arrays[i]
                if err_valid_mask is not None:
                    arr_float = np.where(err_valid_mask, arr.astype(float), 0.0)
                else:
                    arr_float = arr.astype(float)
                params_dict[name] = arr_float

        params_valid_mask = self._validate_params(params_dict)
        if err_valid_mask is not None:
            valid_mask = err_valid_mask & params_valid_mask
        else:
            valid_mask = params_valid_mask

        # 处理 loc/lock 标记
        if self.markers.get('loc') or self.markers.get('lock'):
            samples = self._calculate_loc_value(params_dict, valid_mask)
            return samples

        adjusted_lower, adjusted_upper, trunc_valid_mask = self._adjust_truncation_boundaries(params_dict)
        final_valid_mask = valid_mask & trunc_valid_mask

        # 全有效情况
        if np.all(final_valid_mask):
            if self.method == 'LHC':
                # 拉丁超立方采样
                from sampling_functions import generate_latin_hypercube_1d
                u_lhc = generate_latin_hypercube_1d(n, rng_seed)
                if adjusted_lower is None and adjusted_upper is None:
                    samples = self._ppf(params_dict, u_lhc)
                else:
                    low_cdf = np.zeros(n)
                    high_cdf = np.ones(n)
                    if adjusted_lower is not None:
                        low_cdf = self._cdf(adjusted_lower, params_dict)
                    if adjusted_upper is not None:
                        high_cdf = self._cdf(adjusted_upper, params_dict)
                    u_mapped = low_cdf + u_lhc * (high_cdf - low_cdf)
                    samples = self._ppf(params_dict, u_mapped)
                samples = samples + self.shift_amount
                return samples.astype(float) if samples.dtype.kind in 'fc' else samples
            elif self.method == 'SOBOL':
                # Sobol 序列采样
                from sampling_functions import generate_sobol_1d
                u_sobol = generate_sobol_1d(n, rng_seed)
                if adjusted_lower is None and adjusted_upper is None:
                    samples = self._ppf(params_dict, u_sobol)
                else:
                    low_cdf = np.zeros(n)
                    high_cdf = np.ones(n)
                    if adjusted_lower is not None:
                        low_cdf = self._cdf(adjusted_lower, params_dict)
                    if adjusted_upper is not None:
                        high_cdf = self._cdf(adjusted_upper, params_dict)
                    u_mapped = low_cdf + u_sobol * (high_cdf - low_cdf)
                    samples = self._ppf(params_dict, u_mapped)
                samples = samples + self.shift_amount
                return samples.astype(float) if samples.dtype.kind in 'fc' else samples
            else:
                # 蒙特卡洛采样（原有逻辑）
                rng = np.random.default_rng(rng_seed)
                if adjusted_lower is None and adjusted_upper is None:
                    samples = self._generate_raw_samples(rng, params_dict)
                else:
                    samples = self._generate_truncated_samples(rng, params_dict, adjusted_lower, adjusted_upper)
                return samples.astype(float) if samples.dtype.kind in 'fc' else samples
        else:
            # 部分有效情况，逐索引处理
            samples_obj = np.full(n, ERROR_MARKER, dtype=object)
            valid_indices = np.where(final_valid_mask)[0]
            if len(valid_indices) == 0:
                return samples_obj

            sub_dict = {}
            for name, arr in params_dict.items():
                sub_dict[name] = arr[valid_indices]
            sub_lower = adjusted_lower[valid_indices] if adjusted_lower is not None else None
            sub_upper = adjusted_upper[valid_indices] if adjusted_upper is not None else None

            if self.method == 'LHC':
                from sampling_functions import generate_latin_hypercube_1d
                sub_n = len(valid_indices)
                sub_u = generate_latin_hypercube_1d(sub_n, rng_seed)
                if sub_lower is None and sub_upper is None:
                    sub_samples = self._ppf(sub_dict, sub_u)
                else:
                    low_cdf_sub = np.zeros(sub_n)
                    high_cdf_sub = np.ones(sub_n)
                    if sub_lower is not None:
                        low_cdf_sub = self._cdf(sub_lower, sub_dict)
                    if sub_upper is not None:
                        high_cdf_sub = self._cdf(sub_upper, sub_dict)
                    sub_u_mapped = low_cdf_sub + sub_u * (high_cdf_sub - low_cdf_sub)
                    sub_samples = self._ppf(sub_dict, sub_u_mapped)
                sub_samples = sub_samples + self.shift_amount
            elif self.method == 'SOBOL':
                from sampling_functions import generate_sobol_1d
                sub_n = len(valid_indices)
                sub_u = generate_sobol_1d(sub_n, rng_seed)
                if sub_lower is None and sub_upper is None:
                    sub_samples = self._ppf(sub_dict, sub_u)
                else:
                    low_cdf_sub = np.zeros(sub_n)
                    high_cdf_sub = np.ones(sub_n)
                    if sub_lower is not None:
                        low_cdf_sub = self._cdf(sub_lower, sub_dict)
                    if sub_upper is not None:
                        high_cdf_sub = self._cdf(sub_upper, sub_dict)
                    sub_u_mapped = low_cdf_sub + sub_u * (high_cdf_sub - low_cdf_sub)
                    sub_samples = self._ppf(sub_dict, sub_u_mapped)
                sub_samples = sub_samples + self.shift_amount
            else:
                # 蒙特卡洛采样
                sub_rng = np.random.default_rng(rng_seed)
                if sub_lower is None and sub_upper is None:
                    sub_samples = self._generate_raw_samples(sub_rng, sub_dict)
                else:
                    sub_samples = self._generate_truncated_samples(sub_rng, sub_dict, sub_lower, sub_upper)

            samples_obj[valid_indices] = sub_samples.astype(float) if sub_samples.dtype.kind in 'fc' else sub_samples
            return samples_obj
        
    def generate_samples_dynamic(self, rng_seed: int,
                                 x_arr: Optional[np.ndarray] = None,
                                 p_arr: Optional[np.ndarray] = None,
                                 min_arr: Optional[np.ndarray] = None,
                                 max_arr: Optional[np.ndarray] = None) -> np.ndarray:
        """
        动态生成样本，支持每个迭代的 x_vals/p_vals 不同。
        x_arr: 形状 (m, n) 或 (n,) ，其中 m 为点数，n 为迭代次数
        p_arr: 形状 (m, n) 或 (n,)
        min_arr, max_arr: 形状 (n,) 用于 cumul/general/histogrm
        """
        n = self.n
        dist_type = self.dist_type
        rng = np.random.default_rng(rng_seed)

        # 处理 loc/lock 标记（静态均值）
        if self.markers.get('loc') or self.markers.get('lock'):
            # 对于动态分布，loc/lock 比较复杂，暂时回退到静态均值计算
            # 简单返回错误标记
            return np.full(n, ERROR_MARKER, dtype=object)

        # 根据分布类型选择生成方式
        if dist_type == 'discrete':
            if x_arr is None or p_arr is None:
                return np.full(n, ERROR_MARKER, dtype=object)
            # 确保 x_arr 和 p_arr 为二维 (m, n)
            if x_arr.ndim == 1:
                x_arr = x_arr.reshape(1, -1)
            if p_arr.ndim == 1:
                p_arr = p_arr.reshape(1, -1)
            m = x_arr.shape[0]
            if p_arr.shape[0] != m:
                return np.full(n, ERROR_MARKER, dtype=object)
            samples = np.empty(n, dtype=float)
            for i in range(n):
                x_i = x_arr[:, i]
                p_i = p_arr[:, i]
                # 归一化概率（防止浮点误差）
                total = np.sum(p_i)
                if total <= 0:
                    samples[i] = np.nan
                else:
                    p_norm = p_i / total
                    # 使用 numpy 的 choice 采样
                    idx = rng.choice(m, p=p_norm)
                    samples[i] = x_i[idx]
            return samples

        elif dist_type == 'duniform':
            if x_arr is None:
                return np.full(n, ERROR_MARKER, dtype=object)
            if x_arr.ndim == 1:
                x_arr = x_arr.reshape(1, -1)
            m = x_arr.shape[0]
            samples = np.empty(n, dtype=float)
            for i in range(n):
                x_i = x_arr[:, i]
                # 等概率
                idx = rng.integers(0, m)
                samples[i] = x_i[idx]
            return samples

        elif dist_type == 'cumul':
            if x_arr is None or p_arr is None or min_arr is None or max_arr is None:
                return np.full(n, ERROR_MARKER, dtype=object)
            if x_arr.ndim == 1:
                x_arr = x_arr.reshape(1, -1)
            if p_arr.ndim == 1:
                p_arr = p_arr.reshape(1, -1)
            m = x_arr.shape[0]
            # 确保 min_arr 和 max_arr 长度为 n
            min_arr = np.broadcast_to(min_arr, n)
            max_arr = np.broadcast_to(max_arr, n)
            samples = np.empty(n, dtype=float)
            for i in range(n):
                x_i = x_arr[:, i]
                p_i = p_arr[:, i]
                min_val = min_arr[i]
                max_val = max_arr[i]
                # 构建完整累积分布（包括边界）
                x_full = np.concatenate(([min_val], x_i, [max_val]))
                p_full = np.concatenate(([0.0], p_i, [1.0]))
                # 确保单调递增
                if not (np.all(np.diff(x_full) > 0) and np.all(np.diff(p_full) >= 0)):
                    samples[i] = np.nan
                    continue
                # 通过反函数采样
                u = rng.random()
                # 找到 u 对应的分位数
                idx = np.searchsorted(p_full, u) - 1
                if idx < 0:
                    idx = 0
                if idx >= len(x_full)-1:
                    samples[i] = x_full[-1]
                else:
                    # 线性插值
                    t = (u - p_full[idx]) / (p_full[idx+1] - p_full[idx])
                    samples[i] = x_full[idx] + t * (x_full[idx+1] - x_full[idx])
            return samples

        elif dist_type == 'general':
            # 与 cumul 类似，但 x_vals 和 p_vals 已经包含边界
            if x_arr is None or p_arr is None or min_arr is None or max_arr is None:
                return np.full(n, ERROR_MARKER, dtype=object)
            if x_arr.ndim == 1:
                x_arr = x_arr.reshape(1, -1)
            if p_arr.ndim == 1:
                p_arr = p_arr.reshape(1, -1)
            m = x_arr.shape[0]
            min_arr = np.broadcast_to(min_arr, n)
            max_arr = np.broadcast_to(max_arr, n)
            samples = np.empty(n, dtype=float)
            for i in range(n):
                x_i = x_arr[:, i]
                p_i = p_arr[:, i]
                min_val = min_arr[i]
                max_val = max_arr[i]
                # 验证 min <= x_i <= max
                if np.any(x_i < min_val) or np.any(x_i > max_val):
                    samples[i] = np.nan
                    continue
                # 确保累积概率单调
                if not np.all(np.diff(p_i) >= 0):
                    samples[i] = np.nan
                    continue
                u = rng.random()
                idx = np.searchsorted(p_i, u) - 1
                if idx < 0:
                    idx = 0
                if idx >= len(x_i)-1:
                    samples[i] = x_i[-1]
                else:
                    t = (u - p_i[idx]) / (p_i[idx+1] - p_i[idx])
                    samples[i] = x_i[idx] + t * (x_i[idx+1] - x_i[idx])
            return samples

        elif dist_type == 'histogrm':
            # Histogrm: 区间等分，给定各区间概率
            if p_arr is None or min_arr is None or max_arr is None:
                return np.full(n, ERROR_MARKER, dtype=object)
            if p_arr.ndim == 1:
                p_arr = p_arr.reshape(1, -1)
            m = p_arr.shape[0]  # 区间个数
            min_arr = np.broadcast_to(min_arr, n)
            max_arr = np.broadcast_to(max_arr, n)
            samples = np.empty(n, dtype=float)
            for i in range(n):
                p_i = p_arr[:, i]
                min_val = min_arr[i]
                max_val = max_arr[i]
                total = np.sum(p_i)
                if total <= 0:
                    samples[i] = np.nan
                    continue
                p_norm = p_i / total
                # 选择区间
                idx = rng.choice(m, p=p_norm)
                # 在区间内均匀采样
                width = (max_val - min_val) / m
                a = min_val + idx * width
                b = a + width
                samples[i] = rng.uniform(a, b)
            return samples

        else:
            # 其他分布回退到原有静态方法
            return self.generate_samples(rng_seed, [])
                        
def _sample_drisk_live(
    func_name: str,
    inner: str,
    n: int,
    static_values: Dict[str, Any],
    live_arrays: Dict[str, np.ndarray],
    xl_ns: dict,
    default_sheet: str,
    cell_coord: str = "",
    input_key: str = None,
    seed_base: int = 42,
    simtable_values: Dict[str, float] = None,
    _error_sink: Optional[List] = None,
    markers: Dict = None,
    method: str = 'MC',
    cell_formulas: Dict[str, str] = None   # 新增参数
) -> Tuple[np.ndarray, bool, str]:
    """
    生成分布样本（支持向量化参数和单元格引用公式解析）
    """
    dist_type = _get_dist_type_from_func_name(func_name)

    # 使用 parse_args_with_nested_functions 替代 _parse_drisk_args
    raw_args = parse_args_with_nested_functions(inner)

    if not raw_args:
        stripped = inner.strip()
        if stripped:
            raw_args = [stripped]
        else:
            msg = f"{default_sheet}!{cell_coord} {func_name}() 参数为空"
            if _error_sink is not None:
                _error_sink.append(msg)
            return np.full(n, ERROR_MARKER, dtype=object), True, msg

    if markers is None:
        normal_args = []
        for tok in raw_args:
            tok_val = tok.strip().strip('"').strip("'")
            try:
                normal_args.append(float(tok_val))
            except:
                normal_args.append(tok_val)
        from attribute_functions import extract_markers_from_args
        _, markers = extract_markers_from_args(tuple(normal_args))

    func_name_lower = func_name.lower()
    # ================== 处理 Compound 和 Splice ==================
    if func_name_lower in ('driskcompound', 'drisksplice'):
        # 确保参数数量至少为 3
        while len(raw_args) < 3:
            raw_args.append('')
        freq_formula = raw_args[0].strip()
        sev_formula = raw_args[1].strip()
        deductible_str = raw_args[2].strip() if len(raw_args) > 2 else '0'
        limit_str = raw_args[3].strip() if len(raw_args) > 3 else ''

        # 评估免赔额和限额（仅 Compound 需要）
        deductible_arr = _eval_expr_vec_live(
            deductible_str, n, static_values, live_arrays, xl_ns, default_sheet,
            cell_coord=cell_coord, simtable_values=simtable_values,
            _error_sink=_error_sink
        )
        if limit_str:
            limit_arr = _eval_expr_vec_live(
                limit_str, n, static_values, live_arrays, xl_ns, default_sheet,
                cell_coord=cell_coord, simtable_values=simtable_values,
                _error_sink=_error_sink
            )
        else:
            limit_arr = np.full(n, np.inf, dtype=float)

        # ----- 辅助函数：判断表达式是否为分布函数调用 -----
        def _is_distribution_call(expr: str) -> bool:
            return bool(_DRISK_PATTERN.search(expr))

        # ----- 辅助函数：解析单元格引用，返回其公式（若存在）-----
        def _resolve_cell_to_formula(cell_ref: str) -> str:
            if cell_formulas is None:
                return ''
            norm_ref = normalize_cell_key(cell_ref)
            formula = cell_formulas.get(norm_ref, '')
            if formula and formula.startswith('='):
                formula = _strip_xll_prefix(formula.lstrip('='))
            return formula

        # ----- 统一处理分布参数（频率/严重性 或 左分布/右分布）-----
        def _prepare_distribution_source(expr: str) -> Union[str, np.ndarray]:
            """返回分布源：字符串（分布函数）或数值数组（固定值）"""
            if _is_distribution_call(expr):
                return expr
            # 检查是否为单元格引用
            is_cell_ref = (re.match(r'^[A-Z]+\d+$', expr, re.I) or
                           ('!' in expr and re.match(r'.+![A-Z]+\d+', expr, re.I)))
            if is_cell_ref:
                ref_formula = _resolve_cell_to_formula(expr)
                if ref_formula and _is_distribution_call(ref_formula):
                    return ref_formula
            # 否则作为固定值求值
            arr = _eval_expr_vec_live(
                expr, n, static_values, live_arrays, xl_ns, default_sheet,
                cell_coord=cell_coord, simtable_values=simtable_values,
                _error_sink=_error_sink
            )
            return arr

        # 计算最终种子
        final_seed = seed_base
        if input_key and '_nested_' not in input_key:
            key_hash = zlib.crc32(input_key.encode()) & 0x7fffffff
            final_seed = (seed_base + key_hash) & 0x7fffffff

        if func_name_lower == 'driskcompound':
            freq_source = _prepare_distribution_source(freq_formula)
            sev_source = _prepare_distribution_source(sev_formula)
            # 频率如果是固定值，需要转换为整数数组
            if isinstance(freq_source, np.ndarray):
                freq_int = np.zeros(n, dtype=np.int64)
                for i, val in enumerate(freq_source):
                    try:
                        v = float(val)
                        if np.isnan(v) or v < 0:
                            freq_int[i] = 0
                        else:
                            freq_int[i] = int(np.floor(v))
                    except (TypeError, ValueError):
                        freq_int[i] = 0
                freq_source = freq_int
            # 严重性如果是固定值，直接作为浮点数组
            if isinstance(sev_source, np.ndarray):
                sev_float = np.zeros(n, dtype=float)
                for i, val in enumerate(sev_source):
                    try:
                        sev_float[i] = float(val)
                    except (TypeError, ValueError):
                        sev_float[i] = np.nan
                sev_source = sev_float

            samples = _generate_compound_samples_vectorized(
                freq_source, sev_source, deductible_arr, limit_arr,
                n, static_values, live_arrays, xl_ns, default_sheet,
                simtable_values, final_seed, method, cell_formulas
            )
        else:  # drisksplice
            left_source = _prepare_distribution_source(freq_formula)   # freq_formula 实际上是左分布
            right_source = _prepare_distribution_source(sev_formula)   # sev_formula 实际上是右分布
            # 拼接点参数（可以是数值或表达式）
            point_arr = _eval_expr_vec_live(
                deductible_str, n, static_values, live_arrays, xl_ns, default_sheet,
                cell_coord=cell_coord, simtable_values=simtable_values,
                _error_sink=_error_sink
            )
            # 拼接点通常为标量，取第一个值（若为数组则广播）
            if point_arr.size > 0:
                splice_point = point_arr[0] if point_arr.ndim == 1 else point_arr.flat[0]
            else:
                splice_point = 0.0

            samples = _generate_splice_samples_vectorized(
                left_source, right_source, splice_point,
                n, static_values, live_arrays, xl_ns, default_sheet,
                simtable_values, final_seed, method, cell_formulas
            )

        if np.all(samples == ERROR_MARKER):
            return samples, True, f"{default_sheet}!{cell_coord} Compound/Splice 所有样本均为错误标记"
        return samples, False, ""

    # ================== 数组分布动态参数处理 ==================
    array_dist_types = {'discrete', 'cumul', 'duniform', 'general', 'histogrm'}
    if dist_type in array_dist_types:
        # 收集所有非属性函数参数（属性函数如 DriskSeed, DriskLoc 等会被过滤）
        attr_func_names = ATTRIBUTE_FUNCTIONS
        param_arrays = []
        for idx, tok in enumerate(raw_args):
            tok_s = tok.strip()
            tok_val = tok_s.strip('"').strip("'")
            # 检查是否为属性函数
            m_attr = re.match(r'^([A-Za-z_][A-Za-z0-9_.]*)\(', tok_s, re.I)
            is_attr = False
            if m_attr:
                func_candidate = m_attr.group(1)
                if func_candidate in attr_func_names:
                    is_attr = True
            if is_attr or is_marker_string(tok_val):
                continue
            # 对参数求值（支持区域引用、单元格引用、分布函数等）
            arr = _eval_expr_vec_live(
                tok_s, n, static_values, live_arrays, xl_ns,
                default_sheet, cell_coord=cell_coord,
                simtable_values=simtable_values, _error_sink=_error_sink,
                full_match_to_key=None, cell_formulas=None
            )
            param_arrays.append(arr)

        # 计算最终种子
        final_seed = seed_base
        if input_key and '_nested_' not in input_key:
            key_hash = zlib.crc32(input_key.encode()) & 0x7fffffff
            final_seed = (seed_base + key_hash) & 0x7fffffff
        rng = np.random.default_rng(final_seed)

        # 根据分布类型，从 param_arrays 中提取参数并动态采样
        if dist_type == 'discrete':
            if len(param_arrays) < 2:
                msg = f"{default_sheet}!{cell_coord} Discrete 参数不足，需要至少两个参数"
                if _error_sink:
                    _error_sink.append(msg)
                return np.full(n, ERROR_MARKER, dtype=object), True, msg
            x_arr = param_arrays[0]   # 可能是 (m, n) 或 (n,)
            p_arr = param_arrays[1]   # 可能是 (m, n) 或 (n,)
            # 确保 x_arr 和 p_arr 是二维 (m, n)
            if x_arr.ndim == 1:
                x_arr = x_arr.reshape(1, -1)
            if p_arr.ndim == 1:
                p_arr = p_arr.reshape(1, -1)
            m = x_arr.shape[0]
            if p_arr.shape[0] != m:
                msg = f"{default_sheet}!{cell_coord} Discrete 参数点数不一致"
                if _error_sink:
                    _error_sink.append(msg)
                return np.full(n, ERROR_MARKER, dtype=object), True, msg
            samples = np.empty(n, dtype=float)
            for i in range(n):
                x_i = x_arr[:, i]
                p_i = p_arr[:, i]
                total = np.sum(p_i)
                if total <= 0:
                    samples[i] = np.nan
                    continue
                p_norm = p_i / total
                # 使用 numpy 的 choice 采样
                idx = rng.choice(m, p=p_norm)
                samples[i] = x_i[idx]
            return samples, False, ""

        elif dist_type == 'duniform':
            if len(param_arrays) < 1:
                msg = f"{default_sheet}!{cell_coord} DUniform 参数不足，需要至少一个参数"
                if _error_sink:
                    _error_sink.append(msg)
                return np.full(n, ERROR_MARKER, dtype=object), True, msg
            x_arr = param_arrays[0]
            if x_arr.ndim == 1:
                x_arr = x_arr.reshape(1, -1)
            m = x_arr.shape[0]
            samples = np.empty(n, dtype=float)
            for i in range(n):
                x_i = x_arr[:, i]
                idx = rng.integers(0, m)
                samples[i] = x_i[idx]
            return samples, False, ""

        elif dist_type == 'cumul':
            if len(param_arrays) < 4:
                msg = f"{default_sheet}!{cell_coord} Cumul 参数不足，需要至少四个参数"
                if _error_sink:
                    _error_sink.append(msg)
                return np.full(n, ERROR_MARKER, dtype=object), True, msg
            min_arr = param_arrays[0]
            max_arr = param_arrays[1]
            x_arr = param_arrays[2]
            p_arr = param_arrays[3]
            # 确保 x_arr 和 p_arr 为二维 (m, n)
            if x_arr.ndim == 1:
                x_arr = x_arr.reshape(1, -1)
            if p_arr.ndim == 1:
                p_arr = p_arr.reshape(1, -1)
            m = x_arr.shape[0]
            if p_arr.shape[0] != m:
                msg = f"{default_sheet}!{cell_coord} Cumul 参数点数不一致"
                if _error_sink:
                    _error_sink.append(msg)
                return np.full(n, ERROR_MARKER, dtype=object), True, msg
            # 确保 min 和 max 为长度 n 的一维数组
            min_arr = np.broadcast_to(min_arr, n)
            max_arr = np.broadcast_to(max_arr, n)
            samples = np.empty(n, dtype=float)
            for i in range(n):
                x_i = x_arr[:, i]
                p_i = p_arr[:, i]
                min_val = min_arr[i]
                max_val = max_arr[i]
                # 构建完整累积分布（包括边界）
                x_full = np.concatenate(([min_val], x_i, [max_val]))
                p_full = np.concatenate(([0.0], p_i, [1.0]))
                # 确保单调递增
                if not (np.all(np.diff(x_full) > 0) and np.all(np.diff(p_full) >= 0)):
                    samples[i] = np.nan
                    continue
                u = rng.random()
                idx = np.searchsorted(p_full, u) - 1
                if idx < 0:
                    idx = 0
                if idx >= len(x_full)-1:
                    samples[i] = x_full[-1]
                else:
                    t = (u - p_full[idx]) / (p_full[idx+1] - p_full[idx])
                    samples[i] = x_full[idx] + t * (x_full[idx+1] - x_full[idx])
            return samples, False, ""

        elif dist_type == 'general':
            if len(param_arrays) < 4:
                msg = f"{default_sheet}!{cell_coord} General 参数不足，需要至少四个参数"
                if _error_sink:
                    _error_sink.append(msg)
                return np.full(n, ERROR_MARKER, dtype=object), True, msg
            min_arr = param_arrays[0]
            max_arr = param_arrays[1]
            x_arr = param_arrays[2]
            p_arr = param_arrays[3]
            if x_arr.ndim == 1:
                x_arr = x_arr.reshape(1, -1)
            if p_arr.ndim == 1:
                p_arr = p_arr.reshape(1, -1)
            m = x_arr.shape[0]
            if p_arr.shape[0] != m:
                msg = f"{default_sheet}!{cell_coord} General 参数点数不一致"
                if _error_sink:
                    _error_sink.append(msg)
                return np.full(n, ERROR_MARKER, dtype=object), True, msg
            min_arr = np.broadcast_to(min_arr, n)
            max_arr = np.broadcast_to(max_arr, n)
            samples = np.empty(n, dtype=float)
            for i in range(n):
                x_i = x_arr[:, i]
                p_i = p_arr[:, i]
                min_val = min_arr[i]
                max_val = max_arr[i]
                # 验证 min <= x_i <= max
                if np.any(x_i < min_val) or np.any(x_i > max_val):
                    samples[i] = np.nan
                    continue
                if not np.all(np.diff(p_i) >= 0):
                    samples[i] = np.nan
                    continue
                u = rng.random()
                idx = np.searchsorted(p_i, u) - 1
                if idx < 0:
                    idx = 0
                if idx >= len(x_i)-1:
                    samples[i] = x_i[-1]
                else:
                    t = (u - p_i[idx]) / (p_i[idx+1] - p_i[idx])
                    samples[i] = x_i[idx] + t * (x_i[idx+1] - x_i[idx])
            return samples, False, ""

        elif dist_type == 'histogrm':
            if len(param_arrays) < 3:
                msg = f"{default_sheet}!{cell_coord} Histogrm 参数不足，需要至少三个参数"
                if _error_sink:
                    _error_sink.append(msg)
                return np.full(n, ERROR_MARKER, dtype=object), True, msg
            min_arr = param_arrays[0]
            max_arr = param_arrays[1]
            p_arr = param_arrays[2]   # 各区间概率，形状 (m, n)
            if p_arr.ndim == 1:
                p_arr = p_arr.reshape(1, -1)
            m = p_arr.shape[0]
            min_arr = np.broadcast_to(min_arr, n)
            max_arr = np.broadcast_to(max_arr, n)
            samples = np.empty(n, dtype=float)
            for i in range(n):
                p_i = p_arr[:, i]
                min_val = min_arr[i]
                max_val = max_arr[i]
                total = np.sum(p_i)
                if total <= 0:
                    samples[i] = np.nan
                    continue
                p_norm = p_i / total
                idx = rng.choice(m, p=p_norm)
                width = (max_val - min_val) / m
                a = min_val + idx * width
                b = a + width
                samples[i] = rng.uniform(a, b)
            return samples, False, ""

    # ========== 对于数组分布且 markers 中已包含数组，直接使用向量化生成器（旧逻辑，已被动态分支替代，但保留作为回退） ==========
    if dist_type in ['cumul', 'discrete', 'general', 'histogrm'] and markers and (
        ('x_vals' in markers and 'p_vals' in markers) 
    ):
        final_seed_vec = seed_base
        if input_key and '_nested_' not in input_key:
            key_hash = zlib.crc32(input_key.encode()) & 0x7fffffff
            final_seed_vec = (seed_base + key_hash) & 0x7fffffff
        vec_gen = VectorizedDistributionGenerator(func_name, dist_type, markers, n, method=method)
        samples = vec_gen.generate_samples(final_seed_vec, [])  # 传入空参数数组
        return samples, False, ""

    if dist_type == 'duniform' and markers and 'x_vals' in markers and 'p_vals' in markers:
        final_seed_vec = seed_base
        if input_key and '_nested_' not in input_key:
            key_hash = zlib.crc32(input_key.encode()) & 0x7fffffff
            final_seed_vec = (seed_base + key_hash) & 0x7fffffff
        vec_gen = VectorizedDistributionGenerator(func_name, dist_type, markers, n, method=method)
        samples = vec_gen.generate_samples(final_seed_vec, [])
        return samples, False, ""

    # 修改点：histogrm 不再列入“不收集普通参数”的列表，而是走 else 分支收集 min, max
    if dist_type in ['duniform', 'cumul', 'discrete', 'general']:
        non_marker_arrays = []
        all_scalar = True
    else:
        non_marker_arrays = []
        all_scalar = True
        attr_func_names = ATTRIBUTE_FUNCTIONS
        range_pattern = re.compile(r'^([A-Za-z_]+!)?[A-Z]+\d+:[A-Z]+\d+$', re.I)

        for idx, tok in enumerate(raw_args):
            tok_s = tok.strip()
            tok_val = tok_s.strip('"').strip("'")
            # 对于 histogrm，跳过第三个参数（p_table），因为它已通过 markers 处理
            if dist_type == 'histogrm' and idx == 2:
                if DEBUG:
                    print(f"Histogrm: 跳过第三个参数 (p_table): {tok_s}")
                continue
            is_attr = False
            m = re.match(r'^([A-Za-z_][A-Za-z0-9_.]*)\(', tok_s, re.I)
            if m:
                func_candidate = m.group(1)
                if func_candidate in attr_func_names:
                    is_attr = True
            if is_attr or is_marker_string(tok_val):
                continue

            arr = _eval_expr_vec_live(
                tok_s, n, static_values, live_arrays, xl_ns,
                default_sheet, cell_coord=cell_coord,
                simtable_values=simtable_values, _error_sink=_error_sink
            )
            non_marker_arrays.append(arr)
            # 注意：原来的循环内参数不足检查已移除，移至循环后统一检查
            if len(arr) > 0:
                first = arr[0]
                is_numeric = isinstance(first, (int, float, np.number))
                if not is_numeric or (is_numeric and np.isnan(first)) or not np.all(arr == first):
                    all_scalar = False

        # 循环结束后，检查 histogrm 参数数量（至少需要 min 和 max）
        if dist_type == 'histogrm' and len(non_marker_arrays) < 2:
            msg = f"{default_sheet}!{cell_coord} Histogrm 参数不足：期望至少2个参数(min,max)，实际得到 {len(non_marker_arrays)}"
            if _error_sink is not None:
                _error_sink.append(msg)
            return np.full(n, ERROR_MARKER, dtype=object), True, msg

    # 如果所有参数都是标量，且不是数组分布，使用标量生成器（保留原有逻辑）
    if all_scalar and dist_type not in ['duniform', 'cumul', 'discrete', 'general', 'histogrm']:
        if not non_marker_arrays:
            msg = f"{default_sheet}!{cell_coord} {func_name}() 没有有效的参数，可能公式解析失败"
            if _error_sink is not None:
                _error_sink.append(msg)
            return np.full(n, ERROR_MARKER, dtype=object), True, msg
        dist_params = []
        for arr in non_marker_arrays:
            if len(arr) == 0:
                continue
            val = arr[0]
            if val == ERROR_MARKER or (isinstance(val, str) and val.upper() == "#ERROR!"):
                msg = f"{default_sheet}!{cell_coord} {func_name}() 参数中包含错误标记"
                if _error_sink is not None:
                    _error_sink.append(msg)
                return np.full(n, ERROR_MARKER, dtype=object), True, msg
            try:
                dist_params.append(float(val))
            except:
                dist_params.append(np.nan)
            from constants import get_distribution_info
        dist_info = get_distribution_info(func_name)
        min_params = dist_info.get('min_params', 0) if dist_info else 0
        if len(dist_params) < min_params:
            msg = f"{default_sheet}!{cell_coord} {func_name}() 参数数量不足，需要至少 {min_params} 个参数"
            if _error_sink is not None:
                _error_sink.append(msg)
            return np.full(n, ERROR_MARKER, dtype=object), True, msg

        nan_indices = [i for i, p in enumerate(dist_params) if np.isnan(p)]
        if nan_indices:
            msg = (f"{default_sheet}!{cell_coord} {func_name}() "
                   f"第 {[i+1 for i in nan_indices]} 个参数为 NaN")
            if _error_sink is not None:
                _error_sink.append(msg)
            return np.full(n, ERROR_MARKER, dtype=object), True, msg

        final_seed = seed_base
        if input_key and '_nested_' not in input_key:
            key_hash = zlib.crc32(input_key.encode()) & 0x7fffffff
            final_seed = (seed_base + key_hash) & 0x7fffffff
            if DEBUG:
                print(f"    [sample] 顶层分布 {input_key} 使用最终种子 {final_seed} (base={seed_base}, hash={key_hash})")

        try:
            generator = TruncatableDistributionGenerator(
                func_name, dist_type, dist_params, markers,
                input_key=input_key or f"{default_sheet}!{cell_coord}",
                seed=final_seed
            )
            samples = generator.generate_samples(n)
            samples_obj = np.empty(n, dtype=object)
            for i, val in enumerate(samples):
                if val == ERROR_MARKER or (isinstance(val, str) and val.upper() == "#ERROR!"):
                    samples_obj[i] = ERROR_MARKER
                else:
                    try:
                        samples_obj[i] = float(val)
                    except:
                        samples_obj[i] = ERROR_MARKER
            if np.any(samples_obj == ERROR_MARKER):
                param_str = ', '.join([str(p) for p in dist_params])
                return samples_obj, True, f"{default_sheet}!{cell_coord} 生成样本失败，参数 [{param_str}] 无效"
            if DEBUG:
                display_key = input_key if input_key else f"{default_sheet}!{cell_coord}"
                print(f"    [sample] {display_key} 批量生成样本，前5值: {samples_obj[:5]}")
            return samples_obj, False, ""
        except Exception as e:
            msg = f"{default_sheet}!{cell_coord} 生成样本异常: {e}"
            if _error_sink is not None:
                _error_sink.append(msg)
            return np.full(n, ERROR_MARKER, dtype=object), True, msg
    else:
        # ========== 向量化分支（包括所有非标量参数 或 数组分布） ==========
        try:
            final_seed_vec = seed_base
            if input_key and '_nested_' not in input_key:
                key_hash = zlib.crc32(input_key.encode()) & 0x7fffffff
                final_seed_vec = (seed_base + key_hash) & 0x7fffffff
            vec_gen = VectorizedDistributionGenerator(func_name, dist_type, markers, n, method=method)
            samples = vec_gen.generate_samples(final_seed_vec, non_marker_arrays)
            if np.all(samples == ERROR_MARKER):
                return samples, True, f"{default_sheet}!{cell_coord} 所有样本均为错误标记"
            if DEBUG:
                print(f"    [vec_sample] {input_key or cell_coord} 向量化生成样本，前5值: {samples[:5]}")
            return samples, False, ""
        except Exception as e:
            msg = f"{default_sheet}!{cell_coord} 向量化生成样本异常: {e}"
            if _error_sink is not None:
                _error_sink.append(msg)
            return np.full(n, ERROR_MARKER, dtype=object), True, msg
                                                                                          
def _eval_call_vec_live(
    func_name: str, inner: str, n: int,
    static_values: Dict[str, Any],
    live_arrays: Dict[str, np.ndarray],
    xl_ns: dict, default_sheet: str,
    cell_coord: str = "",
    input_key: str = None,
    seed_base: int = 42,
    simtable_values: Dict[str, float] = None,
    _error_sink: Optional[List] = None,
    full_match_to_key: Dict = None,
    cell_formulas: Dict[str, str] = None   # 新增：单元格公式字典
) -> np.ndarray:
    fn_upper = func_name.upper()

    if fn_upper == 'DRISKOUTPUT':
        return np.zeros(n, dtype=float)

    if fn_upper == 'DRISKMAKEINPUT':
        raw_args = _parse_drisk_args(inner)
        if raw_args:
            first_arg = raw_args[0].strip()
            return _eval_expr_vec_live(first_arg, n, static_values, live_arrays, xl_ns, default_sheet,
                                       cell_coord=cell_coord, simtable_values=simtable_values,
                                       _error_sink=_error_sink, full_match_to_key=full_match_to_key,
                                       cell_formulas=cell_formulas)
        return np.zeros(n, dtype=float)
    
    if fn_upper == 'RAND':
        # 生成 n 个 [0,1) 均匀分布随机数
        rng = np.random.default_rng(seed_base)
        return rng.random(n)

    if fn_upper == 'RANDBETWEEN':
        raw_args = _parse_drisk_args(inner)
        if len(raw_args) < 2:
            return np.full(n, np.nan, dtype=float)
        # 求值下界和上界参数
        lower_arr = _eval_expr_vec_live(
            raw_args[0], n, static_values, live_arrays, xl_ns, default_sheet,
            cell_coord=cell_coord, simtable_values=simtable_values,
            _error_sink=_error_sink, full_match_to_key=full_match_to_key,
            cell_formulas=cell_formulas
        )
        upper_arr = _eval_expr_vec_live(
            raw_args[1], n, static_values, live_arrays, xl_ns, default_sheet,
            cell_coord=cell_coord, simtable_values=simtable_values,
            _error_sink=_error_sink, full_match_to_key=full_match_to_key,
            cell_formulas=cell_formulas
        )
        # 确保下界和上界为整数
        lower_int = np.floor(lower_arr).astype(int)
        upper_int = np.floor(upper_arr).astype(int)
        # 如果下界 > 上界，交换
        lower_int, upper_int = np.minimum(lower_int, upper_int), np.maximum(lower_int, upper_int)
        # 使用随机生成器
        rng = np.random.default_rng(seed_base)
        # 生成整数随机数，支持数组参数（NumPy 1.20+）
        try:
            result = rng.integers(lower_int, upper_int + 1, size=n)
        except TypeError:
            # 兼容旧版 NumPy，使用列表推导
            result = np.array([rng.integers(lower_int[i], upper_int[i] + 1) for i in range(n)])
        return result.astype(float)

    if fn_upper == 'DRISKSIMTABLE':
        if simtable_values:
            full_addr = f"{default_sheet}!{cell_coord}".upper() if cell_coord else None
            val = None
            # 1. 尝试完整地址（带工作表名）
            if full_addr and full_addr in simtable_values:
                val = simtable_values[full_addr]
            # 2. 尝试纯单元格地址（不带工作表名）
            elif cell_coord and cell_coord.upper() in simtable_values:
                val = simtable_values[cell_coord.upper()]
            # 3. 尝试用 default_sheet 和 cell_coord 构建另一种格式（可能大小写不同）
            elif full_addr and full_addr.upper() in simtable_values:
                val = simtable_values[full_addr.upper()]
            if val is not None:
                if val == ERROR_MARKER:
                    return np.full(n, ERROR_MARKER, dtype=object)
                try:
                    return np.full(n, float(val), dtype=float)
                except:
                    return np.full(n, np.nan, dtype=float)
        # 未找到时返回错误标记，避免静默使用 0
        return np.full(n, ERROR_MARKER, dtype=object)

    if func_name in ATTRIBUTE_FUNCTIONS:
        return np.zeros(n, dtype=float)

    if fn_upper == 'SUBTOTAL':
        raw_args = _parse_drisk_args(inner)   # 原始参数列表（字符串）
        # 调用核心函数，传递 n、static_values、live_arrays、default_sheet 和 app
        return _subtotal_core(
            function_num=raw_args[0],
            ref_strs=raw_args[1:],
            n=n,
            static_values=static_values,
            live_arrays=live_arrays,
            default_sheet=default_sheet,
            app=xl_app(),
            cell_formulas=cell_formulas   # 传递公式字典
        )
    
    if fn_upper == 'NPV':
        raw_args = _parse_drisk_args(inner)
        expanded_args = []
        for arg in raw_args:
            arg = arg.strip()
            if ':' in arg and re.match(r'^[A-Z]+\d+:[A-Z]+\d+$', arg, re.I):
                match = re.match(r'([A-Z]+)(\d+):([A-Z]+)(\d+)', arg, re.I)
                if match:
                    start_col, start_row, end_col, end_row = match.groups()
                    start_row = int(start_row)
                    end_row = int(end_row)
                    start_col_num = _col2n(start_col)
                    end_col_num = _col2n(end_col)
                    for row in range(start_row, end_row+1):
                        for col_num in range(start_col_num, end_col_num+1):
                            col_letter = _n2col(col_num)
                            cell_ref = f"{col_letter}{row}"
                            expanded_args.append(cell_ref)
                else:
                    expanded_args.append(arg)
            else:
                expanded_args.append(arg)
        param_arrays = [
            _eval_expr_vec_live(a, n, static_values, live_arrays, xl_ns, default_sheet,
                                cell_coord=cell_coord, simtable_values=simtable_values,
                                _error_sink=_error_sink, full_match_to_key=full_match_to_key,
                                cell_formulas=cell_formulas)
            for a in expanded_args
        ]
        if len(param_arrays) < 1:
            return np.zeros(n, dtype=float)
        rate_arr = param_arrays[0]
        cash_flows = param_arrays[1:]
        return _npv_vectorized(rate_arr, *cash_flows)

    if _DRISK_NAME_RE.match(func_name):
        full_call = f"{func_name}({inner})"
        full_call_no_space = re.sub(r'\s+', '', full_call)
        precomputed_key = None
        if full_match_to_key and cell_coord:
            # 使用 default_sheet 和 cell_coord 构建键
            key = f"{default_sheet}!{cell_coord}!{full_call}"
            if key in full_match_to_key:
                precomputed_key = full_match_to_key[key]
            else:
                key_no_space = f"{default_sheet}!{cell_coord}!{full_call_no_space}"
                if key_no_space in full_match_to_key:
                    precomputed_key = full_match_to_key[key_no_space]
        if precomputed_key and precomputed_key.upper() in live_arrays:
            if DEBUG:
                print(f"重用预计算分布 {full_call} -> {precomputed_key}")
            return live_arrays[precomputed_key.upper()]
        else:
            samples, err, msg = _sample_drisk_live(
                func_name, inner, n, static_values, live_arrays, xl_ns,
                default_sheet, cell_coord=cell_coord, input_key=input_key,
                seed_base=seed_base, simtable_values=simtable_values,
                _error_sink=_error_sink, markers=None, cell_formulas=cell_formulas)
            if err and _error_sink is not None:
                _error_sink.append(msg)
            return samples

    aggregate_fns = {'SUM', 'SUMA', 'AVERAGE', 'AVERAGEA', 'MAX', 'MAXA', 'MIN', 'MINA',
                     'PRODUCT', 'COUNT', 'COUNTA', 'STDEV', 'STDEVA', 'STDEVP', 'STDEVPA',
                     'VAR', 'VARA', 'VARP', 'VARPA', 'MEDIAN'}
    if fn_upper in aggregate_fns:
        raw_args = _parse_drisk_args(inner)
        return _eval_aggregate_live(
            fn_upper, raw_args, n, static_values, live_arrays, xl_ns, default_sheet,
            simtable_values=simtable_values, _error_sink=_error_sink)

    fn_impl = xl_ns.get(fn_upper)
    if fn_impl is not None:
        raw_args = _parse_drisk_args(inner)
        if fn_upper == 'PI' and not inner.strip():
            return np.full(n, math.pi, dtype=float)
        param_arrays = [
            _eval_expr_vec_live(a, n, static_values, live_arrays, xl_ns, default_sheet,
                                cell_coord=cell_coord, simtable_values=simtable_values,
                                _error_sink=_error_sink, full_match_to_key=full_match_to_key,
                                cell_formulas=cell_formulas)
            for a in raw_args
        ]
        try:
            result = fn_impl(*param_arrays)
            if isinstance(result, np.ndarray):
                if result.dtype.kind in 'fc':
                    return result.astype(float)
                else:
                    return result
            return np.full(n, float(result), dtype=float)
        except Exception as e:
            print(f"    [函数 {func_name} 失败] {e}")
            return np.zeros(n, dtype=float)

    print(f"    [未知函数] {func_name}(...)，将触发回退（预检应已检测）")
    return np.zeros(n, dtype=float)

def _eval_expr_vec_live(
    expr: str, n: int,
    static_values: Dict[str, Any],
    live_arrays: Dict[str, np.ndarray],
    xl_ns: dict,
    default_sheet: str,
    cell_coord: str = "",
    input_key: str = None,
    seed_base: int = 42,
    simtable_values: Dict[str, float] = None,
    _error_sink: Optional[List] = None,
    full_match_to_key: Dict = None,
    cell_formulas: Dict[str, str] = None   # 新增
) -> np.ndarray:
    expr = expr.strip()
    if not expr:
        return np.zeros(n, dtype=float)

    tokens = _get_tokenized_expr(expr)

    if len(tokens) == 1 and tokens[0][0] == 'atom':
        return _eval_atom_vec_live(tokens[0][1], n, static_values, live_arrays, default_sheet, simtable_values)

    local_ns: Dict[str, Any] = {}
    var_idx = 0
    parts = []

    for tok in tokens:
        if tok[0] == 'call':
            _, func_name, inner = tok
            arr = _eval_call_vec_live(
                func_name, inner, n, static_values, live_arrays, xl_ns, default_sheet,
                cell_coord=cell_coord, input_key=input_key, seed_base=seed_base,
                simtable_values=simtable_values, _error_sink=_error_sink,
                full_match_to_key=full_match_to_key,
                cell_formulas=cell_formulas)   # 传递
            vname = f'__v{var_idx}__'
            var_idx += 1
            local_ns[vname] = arr
            parts.append(vname)
        else:
            resolved, extra_vars = _resolve_atoms_live(
                tok[1], n, static_values, live_arrays, default_sheet, simtable_values)
            resolved = _convert_excel_operators(resolved)
            local_ns.update(extra_vars)
            parts.append(resolved)

    for vname, arr in local_ns.items():
        if isinstance(arr, np.ndarray) and arr.dtype.kind == 'O':
            if np.any(arr == ERROR_MARKER):
                return np.full(n, ERROR_MARKER, dtype=object)

    combined = ''.join(parts)
    eval_ns = {'np': np, '__builtins__': {}}
    eval_ns.update(local_ns)
    try:
        result = eval(combined, eval_ns)
        if isinstance(result, np.ndarray):
            if result.dtype.kind in 'fc':
                return result.astype(float)
            else:
                return result
        return np.full(n, float(result), dtype=float)
    except ZeroDivisionError:
        msg = f"{default_sheet}!{cell_coord} 表达式含除零: {expr!r}"
        if _error_sink is not None:
            _error_sink.append(msg)
        return np.full(n, np.nan, dtype=float)
    except Exception as e:
        print(f"    [eval_live 失败] expr={expr!r}  err={e}")
        for v in local_ns.values():
            if isinstance(v, np.ndarray):
                if v.dtype.kind in 'fc':
                    return v.astype(float)
                else:
                    return v
        return np.full(n, np.nan, dtype=float)

def _generate_compound_samples_vectorized(
    freq_source,           # 可以是 str（分布函数）或 np.ndarray（固定整数数组）
    sev_source,            # 可以是 str（分布函数）或 np.ndarray（固定浮点数组）
    deductible_arr: np.ndarray,
    limit_arr: np.ndarray,
    n: int,
    static_values: Dict[str, Any],
    live_arrays: Dict[str, np.ndarray],
    xl_ns: dict,
    default_sheet: str,
    simtable_values: Dict[str, float],
    seed_base: int,
    method: str,
    cell_formulas: Dict[str, str] = None   # 新增参数
) -> np.ndarray:
    """
    向量化生成复合分布样本。
    支持频率和严重性为分布函数（字符串）或固定值（数组）。
    """
    # 1. 处理频率
    if isinstance(freq_source, str):
        # 频率是分布函数，采样得到频率数组
        freq_match = re.match(r'([A-Za-z_][A-Za-z0-9_.]*)\s*\(', freq_source, re.I)
        if not freq_match:
            raise ValueError(f"无法解析频率分布公式: {freq_source}")
        freq_func = freq_match.group(1)
        freq_inner = freq_source[freq_match.end():].rstrip(')').strip()
        freq_samples, err, msg = _sample_drisk_live(
            freq_func, freq_inner, n, static_values, live_arrays, xl_ns,
            default_sheet, cell_coord="", input_key=None, seed_base=seed_base,
            simtable_values=simtable_values, _error_sink=None, markers=None,
            method=method, cell_formulas=cell_formulas   # 传递 cell_formulas
        )
        if err:
            raise ValueError(f"生成频率样本失败: {msg}")
        # 转换为整数
        freq_int = np.zeros(n, dtype=np.int64)
        for i, val in enumerate(freq_samples):
            try:
                v = float(val)
                if np.isnan(v) or v < 0:
                    freq_int[i] = 0
                else:
                    freq_int[i] = int(np.floor(v))
            except (TypeError, ValueError):
                freq_int[i] = 0
        freq = freq_int
    else:
        # 频率是固定数组，直接使用（已转换为整数）
        freq = freq_source.astype(np.int64)

    # 2. 处理严重性
    if isinstance(sev_source, str):
        # 严重性是分布函数，需要按每个迭代生成多个样本
        sev_match = re.match(r'([A-Za-z_][A-Za-z0-9_.]*)\s*\(', sev_source, re.I)
        if not sev_match:
            raise ValueError(f"无法解析严重性分布公式: {sev_source}")
        sev_func = sev_match.group(1)
        sev_inner = sev_source[sev_match.end():].rstrip(')').strip()

        max_freq = int(np.max(freq))
        if max_freq == 0:
            return np.zeros(n, dtype=float)

        # 预先为所有索赔事件生成样本矩阵 (max_freq, n)
        sev_matrix = np.zeros((max_freq, n), dtype=float)
        # 每个索赔事件独立同分布，使用不同的随机种子
        for i in range(max_freq):
            layer_seed = (seed_base + i) & 0x7fffffff
            sev_layer, err, msg = _sample_drisk_live(
                sev_func, sev_inner, n, static_values, live_arrays, xl_ns,
                default_sheet, cell_coord="", input_key=None, seed_base=layer_seed,
                simtable_values=simtable_values, _error_sink=None, markers=None,
                method=method, cell_formulas=cell_formulas   # 传递 cell_formulas
            )
            if err:
                raise ValueError(f"生成严重性样本失败: {msg}")
            # 转换为浮点数
            if sev_layer.dtype.kind == 'O':
                sev_float = np.full(n, np.nan, dtype=float)
                for j, val in enumerate(sev_layer):
                    try:
                        sev_float[j] = float(val)
                    except (TypeError, ValueError):
                        pass
                sev_layer = sev_float
            else:
                sev_layer = sev_layer.astype(float)
            sev_matrix[i, :] = sev_layer
    else:
        # 严重性是固定数组，每个迭代的索赔额固定
        sev_float = sev_source.astype(float)
        max_freq = 1   # 无需生成矩阵，直接使用固定值
        sev_matrix = sev_float.reshape(1, -1)  # 形状 (1, n)

    # 3. 应用免赔额和限额
    deductible_arr_float = deductible_arr.astype(float)
    limit_arr_float = limit_arr.astype(float)
    if sev_matrix.shape[0] > 1:
        # 严重性为分布的情况
        sev_matrix = sev_matrix - deductible_arr_float
        sev_matrix = np.maximum(sev_matrix, 0.0)
        limit_broad = limit_arr_float.reshape(1, -1)
        sev_matrix = np.minimum(sev_matrix, limit_broad)
    else:
        # 固定严重性，直接处理一次
        fixed_sev = sev_matrix[0, :] - deductible_arr_float
        fixed_sev = np.maximum(fixed_sev, 0.0)
        fixed_sev = np.minimum(fixed_sev, limit_arr_float)
        sev_matrix = fixed_sev.reshape(1, -1)

    # 4. 按频率累加索赔额
    totals = np.zeros(n, dtype=float)
    if max_freq == 1 and not isinstance(sev_source, str):
        # 严重性固定且频率可能变化：每个迭代索赔额 = 频率 * 固定索赔额
        totals = freq * sev_matrix[0, :]
    else:
        # 一般情况：累加每个迭代的前 freq[i] 个样本
        for i in range(n):
            cnt = freq[i]
            if cnt > 0:
                if cnt > sev_matrix.shape[0]:
                    cnt = sev_matrix.shape[0]   # 安全截断
                totals[i] = np.sum(sev_matrix[:cnt, i])
    return totals

def _generate_splice_samples_vectorized(
    left_source,           # 可以是 str（分布函数）或 np.ndarray（固定浮点数组）
    right_source,          # 可以是 str（分布函数）或 np.ndarray（固定浮点数组）
    splice_point: float,   # 拼接点标量
    n: int,
    static_values: Dict[str, Any],
    live_arrays: Dict[str, np.ndarray],
    xl_ns: dict,
    default_sheet: str,
    simtable_values: Dict[str, float],
    seed_base: int,
    method: str,
    cell_formulas: Dict[str, str] = None
) -> np.ndarray:
    """
    向量化生成 Splice 分布样本。
    支持左分布和右分布为分布函数（字符串）或固定值（数组）。
    拼接点固定为标量。
    """
    # 辅助函数：判断是否为分布公式字符串
    def is_distribution_string(s):
        return isinstance(s, str) and any(f in s for f in DISTRIBUTION_FUNCTION_NAMES)

    # 辅助函数：将可能为数组的源转换为字符串（取第一个元素）
    def to_formula_string(source):
        if isinstance(source, str):
            return source
        if isinstance(source, np.ndarray) and source.size > 0:
            val = source[0]
            if isinstance(val, str):
                return val
            # 如果第一个元素是数值，说明是固定值，返回其字符串表示
            return str(val)
        raise ValueError(f"无法将 source 转换为公式字符串: {source}")

    # 情况1：左右分布都是分布公式字符串 -> 向量化生成
    if is_distribution_string(left_source) and is_distribution_string(right_source):
        rng = np.random.default_rng(seed_base)
        # 直接调用 dist_splice 模块中的向量化生成器
        from dist_splice import splice_generator_vectorized
        return splice_generator_vectorized(rng, [left_source, right_source, splice_point], n)

    # 情况2：左右分布包含固定值（例如从单元格读取的常数数组）-> 降级为循环
    from dist_splice import splice_generator_single
    rng = np.random.default_rng(seed_base)
    samples = np.zeros(n, dtype=float)

    # 将 left_source, right_source 转换为标量列表（如果是数组，逐元素取；否则重复标量）
    if isinstance(left_source, np.ndarray):
        left_list = left_source.tolist()
    else:
        left_list = [left_source] * n

    if isinstance(right_source, np.ndarray):
        right_list = right_source.tolist()
    else:
        right_list = [right_source] * n

    for i in range(n):
        left = left_list[i]
        right = right_list[i]
        # 如果左右是数值，构造一个“退化的分布”字符串（例如 "DriskNormal(值,0)" 不合适，直接使用数值）
        # 但 splice_generator_single 期望分布公式字符串，因此将数值转为简单字符串形式
        if isinstance(left, (int, float)) and isinstance(right, (int, float)):
            # 两个都是常数：根据拼接点决定取值（简化，实际应使用理论权重，但常数情况极少见）
            samples[i] = left if left <= splice_point else right
        else:
            left_str = str(left) if isinstance(left, (int, float)) else left
            right_str = str(right) if isinstance(right, (int, float)) else right
            samples[i] = splice_generator_single(rng, [left_str, right_str, splice_point])

    return samples

def _eval_formula_with_live(
    formula: str, default_sheet: str, n: int,
    static_values: Dict[str, Any],
    live_arrays: Dict[str, np.ndarray],
    xl_ns: dict,
    cell_coord: str = "",
    input_key: str = None,
    seed_base: int = 42,
    simtable_values: Dict[str, float] = None,
    _error_sink: Optional[List] = None,
    full_match_to_key: Dict = None,
    cell_formulas: Dict[str, str] = None   # 新增
) -> np.ndarray:
    expr = _strip_xll_prefix(formula).lstrip('=').strip()
    if DEBUG:
        print(f"[eval] 单元格 {default_sheet}!{cell_coord} 表达式: {expr}")
    return _eval_expr_vec_live(
        expr, n, static_values, live_arrays, xl_ns, default_sheet,
        cell_coord=cell_coord, input_key=input_key, seed_base=seed_base,
        simtable_values=simtable_values, _error_sink=_error_sink,
        full_match_to_key=full_match_to_key,
        cell_formulas=cell_formulas)

# ==================== 预检函数 ====================
def _preflight_check_drisk_cells(
    all_formulas: Dict[str, str],
    static_values: Dict[str, Any],
) -> List[str]:
    errors: List[str] = []

    for key, formula in all_formulas.items():
        if not _DRISK_PATTERN.search(formula):
            continue
        sheet = key.split("!")[0] if "!" in key else "Sheet1"
        reported = False

        formula_upper = formula.upper()
        for err_str in _EXCEL_ERROR_STRINGS:
            if err_str in formula_upper:
                errors.append(f"单元格 {key} 的公式含 Excel 错误值 {err_str}\n  公式: {formula}")
                reported = True
                break
        if reported:
            continue

        for m in _CELL_REF_PATTERN.finditer(formula):
            # 获取引用的工作表名和单元格地址
            quoted_sheet = m.group(1)
            unquoted_sheet = m.group(2)
            col = m.group(3)
            row = m.group(4)
            if quoted_sheet is not None:
                sheet_name = quoted_sheet.replace("''", "'")
                ref = f"'{sheet_name}'!{col}{row}"
            elif unquoted_sheet is not None:
                sheet_name = unquoted_sheet
                ref = f"{sheet_name}!{col}{row}"
            else:
                ref = f"{col}{row}"
            full_ref = _normalize_key(sheet, ref)
            val = static_values.get(full_ref)
            if val is None:
                val = static_values.get(ref)
            if val is not None:
                try:
                    if np.isnan(float(val)):
                        errors.append(
                            f"单元格 {key} 的参数引用了 {ref}，该单元格在 Excel 中显示错误值（已标记为 NaN）\n"
                            f"  公式: {formula}")
                        reported = True
                        break
                except:
                    pass
    return errors

def _get_known_function_names() -> Set[str]:
    known = set()
    known.update(f.upper() for f in DISTRIBUTION_FUNCTION_NAMES)
    known.update(f.upper() for f in ATTRIBUTE_FUNCTIONS)
    aggregate_fns = {'SUM', 'SUMA', 'AVERAGE', 'AVERAGEA', 'MAX', 'MAXA', 'MIN', 'MINA',
                     'PRODUCT', 'COUNT', 'COUNTA', 'STDEV', 'STDEVA', 'STDEVP', 'STDEVPA',
                     'VAR', 'VARA', 'VARP', 'VARPA', 'MEDIAN'}
    known.update(aggregate_fns)
    xl_ns = _get_xl_ns()
    known.update(f.upper() for f in xl_ns.keys())
    known.update({'DRISKOUTPUT', 'DRISKSIMTABLE', 'DRISKMAKEINPUT'})

    simulation_stats = {
        'DRISKMEAN', 'DRISKSTD', 'DRISKVARIANCE', 'DRISKMIN', 'DRISKMAX',
        'DRISKMEDIAN', 'DRISKPTX', 'DRISKDATA', 'DRISKSKEW', 'DRISKKURT',
        'DRISKRANGE', 'DRISKMODE', 'DRISKMEANABSDEV', 'DRISKXTOP',
        'DRISKCIMEAN', 'DRISKCIPERCENTILE', 'DRISKCOEFFOFVARIATION',
        'DRISKSEMISTDDEV', 'DRISKSEMIVARIANCE', 'DRISKSTDERROFMEAN'
    }
    known.update(simulation_stats)

    theoretical_stats = {
        'DRISKTHEOMEAN', 'DRISKTHEOSTDDEV', 'DRISKTHEOVARIANCE',
        'DRISKTHEOSKEWNESS', 'DRISKTHEOKURTOSIS', 'DRISKTHEOMIN',
        'DRISKTHEOMAX', 'DRISKTHEORANGE', 'DRISKTHEOMODE',
        'DRISKTHEOPTOX', 'DRISKTHEOXTOP', 'DRISKTHEOXTOY'
    }
    known.update(theoretical_stats)

    # 添加 Excel 原生随机函数
    known.update({'RAND', 'RANDBETWEEN'})

    return known

def _check_unsupported_functions(formulas: Dict[str, str], known_functions: Set[str]) -> List[str]:
    unsupported_cells = []
    func_pattern = re.compile(r'([A-Za-z_][A-Za-z0-9_.]*)\s*\(', re.I)
    for key, formula in formulas.items():
        if not formula.startswith('='):
            continue
        matches = func_pattern.findall(formula)
        for func in matches:
            func_upper = func.upper()
            if func_upper not in known_functions:
                unsupported_cells.append(f"{key} (包含未支持函数 {func})")
                break
    return unsupported_cells

def _extract_markers_from_args_text(args_text: str) -> Dict[str, Any]:
    markers = {}
    attr_patterns = {
        'seed': r'DriskSeed\s*\(\s*([^)]*)\s*\)',
        'shift': r'DriskShift\s*\(\s*([^)]*)\s*\)',
        'truncate': r'DriskTruncate\s*\(\s*([^)]*)\s*\)',
        'truncatep': r'DriskTruncateP\s*\(\s*([^)]*)\s*\)',
        'truncate2': r'DriskTruncate2\s*\(\s*([^)]*)\s*\)',
        'truncatep2': r'DriskTruncateP2\s*\(\s*([^)]*)\s*\)',
        'loc': r'DriskLoc\s*\(\s*\)',
        'lock': r'DriskLock\s*\(\s*([^)]*)\s*\)',
        'static': r'DriskStatic\s*\(\s*([^)]*)\s*\)',
        'name': r'DriskName\s*\(\s*([^)]*)\s*\)',
        'units': r'DriskUnits\s*\(\s*([^)]*)\s*\)',
        'category': r'DriskCategory\s*\(\s*([^)]*)\s*\)',
        'corrmat': r'DriskCorrmat\s*\(\s*([^)]*)\s*\)',
        'convergence': r'DriskConvergence\s*\(\s*([^)]*)\s*\)',
        'copula': r'DriskCopula\s*\(\s*([^)]*)\s*\)',
    }
    for marker_type, pattern in attr_patterns.items():
        matches = re.findall(pattern, args_text, re.IGNORECASE)
        if matches:
            if marker_type in ('loc', 'lock'):
                markers[marker_type] = True
            elif marker_type == 'seed':
                param_str = matches[0].strip()
                if param_str:
                    parts = [p.strip() for p in param_str.split(',')]
                    if len(parts) >= 2:
                        try:
                            markers['rng_type'] = int(parts[0])
                            markers['seed'] = int(parts[1])
                        except:
                            markers['seed'] = 42
                    elif len(parts) == 1:
                        try:
                            markers['seed'] = int(parts[0])
                        except:
                            markers['seed'] = 42
                else:
                    markers['seed'] = 42
            elif marker_type == 'corrmat':
                markers['corrmat'] = matches[0].strip()
            elif marker_type == 'copula':
                markers['copula'] = matches[0].strip()
            elif marker_type in ('truncate', 'truncate2', 'truncatep', 'truncatep2'):
                markers[marker_type] = matches[0].strip()
            else:
                markers[marker_type] = matches[0].strip()
    return markers

def _expand_range_refs_in_formula(formula: str, default_sheet: str) -> List[str]:
    """
    提取公式中的区域引用（如 A1:B2 或 Sheet1!C3:D4），返回展开后的完整单元格地址列表（带工作表名，已规范化）。
    """
    refs = []
    range_pattern = re.compile(r'(?:([A-Za-z_][A-Za-z0-9_]*!)?)([A-Z]+\d+):([A-Z]+\d+)', re.I)
    for m in range_pattern.finditer(formula):
        sheet_prefix = m.group(1) or ''
        start = m.group(2)
        end = m.group(3)
        start_match = re.match(r'([A-Z]+)(\d+)', start)
        end_match = re.match(r'([A-Z]+)(\d+)', end)
        if start_match and end_match:
            start_col, start_row = start_match.groups()
            end_col, end_row = end_match.groups()
            start_row = int(start_row)
            end_row = int(end_row)
            start_col_num = _col2n(start_col)
            end_col_num = _col2n(end_col)

            # 确定工作表名
            if sheet_prefix:
                sheet = sheet_prefix.rstrip('!').strip().strip("'").replace("''", "'")
                try:
                    sheet = normalize_sheet_name(sheet, strict=True)
                except ValueError:
                    continue  # 跳过非法工作表名
            else:
                sheet = default_sheet

            # 展开区域内所有单元格，并生成规范键
            for r in range(start_row, end_row + 1):
                for c in range(start_col_num, end_col_num + 1):
                    col_letter = _n2col(c)
                    cell = f"{col_letter}{r}"
                    full_ref = normalize_cell_key(f"{sheet}!{cell}")
                    refs.append(full_ref)
    return refs

def _normalize_key(sheet: str, coord: str) -> str:
    """标准化键名：Sheet!A1 转为大写，并规范化工作表名"""
    return normalize_cell_key(f"{sheet}!{coord}")

def _get_sheet_from_key(key: str) -> str:
    """从单元格键（如 'Sheet1!A1'）中提取工作表名，若无则返回默认 Sheet1"""
    if '!' in key:
        return key.split('!')[0]
    return "Sheet1"

def build_numpy_dependency_graph(
    distribution_cells: Dict[str, List[Dict]],
    makeinput_cells: Dict[str, List[Dict]],
    simtable_cells: Dict[str, List[Dict]],
    all_related_cells: Dict[str, Dict],
    all_formulas: Dict[str, str],
    static_values: Dict[str, Any],
    output_cells: Dict[str, Any] = None
) -> Tuple[Dict[str, List[str]], List[str]]:
    """
    构建依赖图（节点为单元格地址或输入键），返回 (graph, topo_order)。
    修复：区域引用展开后统一使用规范键，确保节点匹配。
    """
    # 初始化节点集合
    nodes = set()
    nodes.update(k.upper() for k in distribution_cells.keys())
    nodes.update(k.upper() for k in makeinput_cells.keys())
    nodes.update(k.upper() for k in simtable_cells.keys())
    if output_cells:
        nodes.update(k.upper() for k in output_cells.keys())
    nodes.update(k.upper() for k in all_related_cells.keys())

    # 辅助函数：获取单元格公式（优先从 all_related_cells 获取，否则从 all_formulas）
    def get_formula(cell_key: str) -> str:
        cell_upper = cell_key.upper()
        if cell_upper in all_related_cells:
            return all_related_cells[cell_upper].get('formula', '')
        return all_formulas.get(cell_upper, '')

    # ---------- BFS 扩展所有可能涉及的单元格 ----------
    from collections import deque
    queue = deque(nodes)
    visited = set(nodes)

    while queue:
        node = queue.popleft()
        formula = get_formula(node)
        if formula and formula.startswith('='):
            sheet = _get_sheet_from_key(node)
            # 提取单个单元格引用（已带工作表名，规范键）
            refs = _extract_cell_refs_from_text(formula, node)
            # 提取区域引用并展开（返回规范键列表）
            range_refs = _expand_range_refs_in_formula(formula, sheet)
            all_refs = list(set(refs + range_refs))
            for ref in all_refs:
                if ref not in nodes:
                    nodes.add(ref)
                    # 如果该单元格也有公式，加入队列继续展开
                    if get_formula(ref):
                        if ref not in visited:
                            visited.add(ref)
                            queue.append(ref)

    # ---------- 构建分布函数映射 ----------
    input_key_to_func_info = {}
    cell_to_input_keys = {}
    for cell_addr, funcs in distribution_cells.items():
        cell_addr_upper = cell_addr.upper()
        cell_keys = []
        for func in funcs:
            input_key = func.get('input_key')
            if not input_key:
                continue
            nodes.add(input_key)
            input_key_to_func_info[input_key] = func
            func['cell_addr'] = cell_addr
            cell_keys.append(input_key)
        if cell_keys:
            cell_to_input_keys[cell_addr_upper] = cell_keys

    # 确保所有 MakeInput、Simtable、输出单元格都在节点中
    for cell_addr in makeinput_cells.keys():
        nodes.add(cell_addr.upper())
    for cell_addr in simtable_cells.keys():
        nodes.add(cell_addr.upper())
    if output_cells:
        for cell_addr in output_cells.keys():
            nodes.add(cell_addr.upper())

    # ---------- 构建依赖图 ----------
    graph = {node: [] for node in nodes}
    # 构建 full_match -> input_key 映射，用于识别函数调用依赖
    full_match_to_key = {}
    for input_key, func in input_key_to_func_info.items():
        fm = func.get('full_match', '')
        if fm:
            cell_addr = func.get('cell_addr')
            if cell_addr:
                # 原始版本
                key_with_cell = f"{cell_addr}!{fm}"
                full_match_to_key[key_with_cell] = input_key
                key_no_space = re.sub(r'\s+', '', key_with_cell)
                full_match_to_key[key_no_space] = input_key

                # 修复：处理 @ 符号
                if fm.startswith('@'):
                    fm_no_at = fm[1:]
                    key_with_cell_no_at = f"{cell_addr}!{fm_no_at}"
                    full_match_to_key[key_with_cell_no_at] = input_key
                    key_no_space_no_at = re.sub(r'\s+', '', key_with_cell_no_at)
                    full_match_to_key[key_no_space_no_at] = input_key
                else:
                    fm_with_at = '@' + fm
                    key_with_cell_with_at = f"{cell_addr}!{fm_with_at}"
                    full_match_to_key[key_with_cell_with_at] = input_key
                    key_no_space_with_at = re.sub(r'\s+', '', key_with_cell_with_at)
                    full_match_to_key[key_no_space_with_at] = input_key

    def add_deps(node, deps):
        for d in deps:
            if d != node and d in graph:
                if d not in graph[node]:
                    graph[node].append(d)

    # 遍历所有节点，解析依赖
    for node in nodes:
        # 分布函数节点（输入键）
        if node in input_key_to_func_info:
            func = input_key_to_func_info[node]
            args_text = func.get('args_text', '')
            # 依赖其他分布函数（通过 full_match 字符串）
            for other_key, other_func in input_key_to_func_info.items():
                if other_key == node:
                    continue
                other_fm = other_func.get('full_match', '')
                if other_fm and other_fm in args_text:
                    add_deps(node, [other_key])
                other_fm_clean = re.sub(r'\s+', '', other_fm) if other_fm else ''
                if other_fm_clean and other_fm_clean in args_text:
                    add_deps(node, [other_key])
            # 依赖单元格引用（包括区域）
            sheet = _get_sheet_from_key(node)
            cell_refs = _extract_cell_refs_from_text(args_text, node)
            range_refs = _expand_range_refs_in_formula(args_text, sheet)
            all_refs = list(set(cell_refs + range_refs))
            for ref in all_refs:
                # 若引用是分布函数单元格，则依赖其所有输入键
                if ref in cell_to_input_keys:
                    for key in cell_to_input_keys[ref]:
                        if key != node:
                            add_deps(node, [key])
                elif ref in nodes:
                    add_deps(node, [ref])
        else:
            # 普通单元格节点
            formula = get_formula(node)
            if formula and formula.startswith('='):
                sheet = _get_sheet_from_key(node)
                refs = _extract_cell_refs_from_text(formula, node)
                range_refs = _expand_range_refs_in_formula(formula, sheet)
                all_refs = list(set(refs + range_refs))
                for full_ref in all_refs:
                    if full_ref in cell_to_input_keys:
                        for key in cell_to_input_keys[full_ref]:
                            if key != node:
                                add_deps(node, [key])
                    if full_ref in nodes:
                        if full_ref != node:
                            add_deps(node, [full_ref])
                # 依赖其他分布函数（通过 full_match 字符串）
                for other_key, other_func in input_key_to_func_info.items():
                    other_fm = other_func.get('full_match', '')
                    if other_fm and other_fm in formula:
                        if other_key != node:
                            add_deps(node, [other_key])
                    other_fm_clean = re.sub(r'\s+', '', other_fm) if other_fm else ''
                    if other_fm_clean and other_fm_clean in formula:
                        if other_key != node:
                            add_deps(node, [other_key])

    # ---------- 拓扑排序 ----------
    rev_graph = {node: [] for node in nodes}
    for node, deps in graph.items():
        for dep in deps:
            if dep in rev_graph:
                rev_graph[dep].append(node)
    in_degree = {node: len(graph[node]) for node in nodes}
    queue = [node for node in nodes if in_degree[node] == 0]
    order = []
    while queue:
        node = queue.pop(0)
        order.append(node)
        for dependent in rev_graph[node]:
            in_degree[dependent] -= 1
            if in_degree[dependent] == 0:
                queue.append(dependent)
    remaining = [n for n in nodes if n not in order]
    order.extend(remaining)

    if DEBUG:
        print("\n========== 依赖图 (Dependency Graph) ==========")
        print(f"节点总数: {len(nodes)}")
        print("依赖关系 (节点 -> 依赖列表):")
        for node in sorted(nodes):
            deps = graph[node]
            if deps:
                print(f"  {node} -> {deps}")
            else:
                print(f"  {node} -> []")
        print("\n拓扑排序顺序 (Topological Order):")
        for i, node in enumerate(order):
            print(f"  {i+1}. {node}")
        print("=============================================\n")

    return graph, order

# ==================== 优化后的 StatusBarProgress（模拟迭代进度，更新间隔0.01秒）====================
class OptimizedStatusBarProgress:
    """优化版状态栏进度：基于节点处理进度模拟迭代进度，更新频率极高，分场景显示，使用滑动窗口即时速度"""
    def __init__(self, app, n_iterations_per_scenario, n_scenarios, total_nodes):
        self.app = app
        self.n_iterations = n_iterations_per_scenario
        self.n_scenarios = n_scenarios
        self.total_nodes = total_nodes
        self.processed_nodes = 0
        self.current_scenario = 0
        self.start_time = time.time()
        self.last_update_time = 0
        self.update_interval = 0.01
        self.cancelled = False
        self.esc_key_pressed = False
        self.last_esc_check_time = 0
        self.esc_check_interval = 0.1

        # 滑动窗口速度相关
        self.node_timestamps = []          # 存储 (时间戳, 累计节点数)
        self.speed_window = 1.0            # 速度计算窗口（秒）

    def start_new_scenario(self, scenario_idx):
        self.current_scenario = scenario_idx
        self.processed_nodes = 0
        self.start_time = time.time()
        # 重置时间戳列表，加入起始点
        self.node_timestamps = [(self.start_time, 0)]
        self.update_status_bar()

    def check_esc_key(self):
        try:
            import win32api
            import win32con
            current_time = time.time()
            if current_time - self.last_esc_check_time < self.esc_check_interval:
                return self.esc_key_pressed
            self.last_esc_check_time = current_time
            key_state = win32api.GetAsyncKeyState(win32con.VK_ESCAPE)
            if key_state & 0x8000:
                self.esc_key_pressed = True
                self.cancelled = True
                return True
        except:
            pass
        return False

    def _compute_instant_speed(self) -> float:
        """基于滑动窗口计算即时节点处理速度（节点/秒），然后转换为迭代/秒"""
        current_time = time.time()
        # 清理窗口外的记录
        self.node_timestamps = [(t, n) for t, n in self.node_timestamps if current_time - t <= self.speed_window]
        # 确保至少有两个点才能计算速度
        if len(self.node_timestamps) >= 2:
            t1, n1 = self.node_timestamps[0]
            t2, n2 = self.node_timestamps[-1]
            dt = t2 - t1
            dn = n2 - n1
            if dt > 0:
                node_speed = dn / dt
                # 每个节点对应的迭代数：总迭代数 / 总节点数
                iter_per_node = self.n_iterations / self.total_nodes if self.total_nodes else 0
                iter_speed = node_speed * iter_per_node
                return iter_speed
        # 降级为全局平均速度
        elapsed = current_time - self.start_time
        if elapsed > 0 and self.processed_nodes > 0:
            node_speed_avg = self.processed_nodes / elapsed
            iter_per_node = self.n_iterations / self.total_nodes if self.total_nodes else 0
            return node_speed_avg * iter_per_node
        return 0.0

    def update_node(self, processed_nodes):
        if self.check_esc_key():
            self.cancelled = True
            return
        self.processed_nodes = processed_nodes
        current_time = time.time()
        # 记录当前时间戳
        self.node_timestamps.append((current_time, self.processed_nodes))
        # 清理窗口外的记录（保持列表大小可控）
        self.node_timestamps = [(t, n) for t, n in self.node_timestamps if current_time - t <= self.speed_window]

        if current_time - self.last_update_time >= self.update_interval:
            self.update_status_bar()
            self.last_update_time = current_time

    def update_status_bar(self):
        try:
            node_percent = (self.processed_nodes / self.total_nodes) * 100 if self.total_nodes else 0
            simulated_iter = int((self.processed_nodes / self.total_nodes) * self.n_iterations) if self.total_nodes else 0
            simulated_iter = min(simulated_iter, self.n_iterations)

            elapsed = time.time() - self.start_time
            elapsed_h = int(elapsed // 3600)
            elapsed_m = int((elapsed % 3600) // 60)
            elapsed_s = int(elapsed % 60)

            # 计算即时速度（迭代/秒）
            iter_speed = self._compute_instant_speed()

            if self.n_scenarios > 1:
                msg = (f"DRISK向量化模拟: 场景 {self.current_scenario+1}/{self.n_scenarios}, "
                       f"迭代 {simulated_iter:,}/{self.n_iterations:,} ({node_percent:.1f}%), "
                       f"时间 {elapsed_h:02d}:{elapsed_m:02d}:{elapsed_s:02d}, "
                       f"速度 {iter_speed:.1f} 迭代/秒")
            else:
                msg = (f"DRISK向量化模拟: 迭代 {simulated_iter:,}/{self.n_iterations:,} ({node_percent:.1f}%), "
                       f"时间 {elapsed_h:02d}:{elapsed_m:02d}:{elapsed_s:02d}, "
                       f"速度 {iter_speed:.1f} 迭代/秒")
            if not self.cancelled:
                msg += " | 按ESC键停止"
            self.app.StatusBar = msg
        except:
            pass

    def clear(self):
        try:
            self.app.StatusBar = False
        except:
            pass

    def is_cancelled(self):
        return self.cancelled
    
# ==================== 辅助函数：构建输入键映射 ====================
def _build_input_key_maps(distribution_cells: Dict[str, List[Dict]]) -> Tuple[Dict[str, Dict], Dict[str, List[str]]]:
    input_key_to_func_info = {}
    cell_to_input_keys = {}
    for cell_addr, funcs in distribution_cells.items():
        keys = []
        for func in funcs:
            input_key = func.get('input_key')
            if input_key:
                input_key_to_func_info[input_key] = func
                keys.append(input_key)
        if keys:
            cell_to_input_keys[cell_addr.upper()] = keys
    return input_key_to_func_info, cell_to_input_keys

# ==================== 回退补丁（保持兼容） ====================
def _apply_fallback_patch():
    import index_functions
    if not hasattr(index_functions.StatusBarProgress, '_original_update'):
        index_functions.StatusBarProgress._original_update = index_functions.StatusBarProgress.update_status_bar
        def patched_update(self):
            self._original_update()
            try:
                self.app.StatusBar = "[回退索引模拟] " + self.app.StatusBar
            except:
                pass
        index_functions.StatusBarProgress.update_status_bar = patched_update

def _remove_fallback_patch():
    import index_functions
    if hasattr(index_functions.StatusBarProgress, '_original_update'):
        index_functions.StatusBarProgress.update_status_bar = index_functions.StatusBarProgress._original_update
        delattr(index_functions.StatusBarProgress, '_original_update')

# ==================== 主模拟函数 ====================
def run_numpy_simulation(n_iterations: int, scenario_count: int = 1, method: str = 'MC'):
    app = None
    progress = None
    try:
        try:
            from com_fixer import _safe_excel_app
            app = _safe_excel_app()
        except ImportError:
            import win32com.client
            app = win32com.client.Dispatch("Excel.Application")

        clear_simulations()
        app.StatusBar = "正在初始化……"

        wb_com, wb_name = _get_active_workbook_com()
        if wb_com is None:
            xlcAlert("无法获取当前工作簿，请确保 Excel 正在运行。")
            return

        # ==================== 获取命名区域映射 ====================
        names_map = _get_names_map(wb_com)
        if DEBUG:
            print(f"命名区域映射: {names_map}")
        # 保存原始计算模式
        original_calc = app.Calculation
        # 强制设为自动计算，并完全重算
        app.Calculation = -4105  # xlCalculationAutomatic
        app.CalculateFullRebuild()
        # 执行扫描
        all_formulas, static_values, _ = _scan_workbook_com(wb_com)
        # 恢复计算模式（可选，稍后优化性能时会再次设置）
        app.Calculation = original_calc
        all_formulas, static_values, _ = _scan_workbook_com(wb_com)

        # ==================== 替换公式中的命名区域 ====================
        if names_map:
            all_formulas_replaced = {}
            for key, formula in all_formulas.items():
                all_formulas_replaced[key] = _replace_names_in_formula(formula, names_map)
            all_formulas = all_formulas_replaced

        preflight_errors = _preflight_check_drisk_cells(all_formulas, static_values)
        if preflight_errors:
            err_text = "\n\n".join(f"• {e}" for e in preflight_errors)
            xlcAlert(
                f"发现 {len(preflight_errors)} 处参数错误，模拟无法启动：\n\n{err_text}\n\n请修正以上单元格后重新运行。"
            )
            return

        try:
            app.StatusBar = "正在查找模拟单元格…………"
            distribution_cells, simtable_cells, makeinput_cells, output_cells, _, all_related = \
                find_all_simulation_cells_in_workbook(app)
        except Exception as e:
            print(f"查找模拟单元格失败: {e}")
            xlcAlert(f"查找模拟单元格失败，将回退到索引模拟模式: {e}")
            _apply_fallback_patch()
            try:
                run_index_simulation(n_iterations, scenario_count)
            finally:
                _remove_fallback_patch()
            return

        # ========== 新增：规范化所有单元格键（工作表名） ==========
        try:
            def normalize_dict_keys(d):
                if not d:
                    return d
                new_d = {}
                for key, value in d.items():
                    try:
                        norm_key = normalize_cell_key(key)
                    except ValueError as e:
                        # 中文标点，重新抛出
                        raise e
                    new_d[norm_key] = value
                return new_d

            distribution_cells = normalize_dict_keys(distribution_cells)
            simtable_cells = normalize_dict_keys(simtable_cells)
            makeinput_cells = normalize_dict_keys(makeinput_cells)
            output_cells = normalize_dict_keys(output_cells)
            all_related = normalize_dict_keys(all_related)
        except ValueError as e:
            xlcAlert(f"模拟中止：工作表名包含非法字符（中文标点等）：{e}")
            return

        # ==================== 替换 all_related 中的公式命名区域 ====================
        if names_map:
            for key, cell_info in all_related.items():
                if 'formula' in cell_info and cell_info['formula']:
                    cell_info['formula'] = _replace_names_in_formula(cell_info['formula'], names_map)

        # 扫描原生随机函数
        native_random_cells = set()
        for key, formula in all_formulas.items():
            if formula and isinstance(formula, str):
                formula_upper = formula.upper()
                if 'RAND(' in formula_upper or 'RANDBETWEEN(' in formula_upper:
                    native_random_cells.add(key.upper())

        # 如果有原生随机函数，将它们加入 all_related，以便依赖图包含它们
        for cell in native_random_cells:
            if cell not in all_related:
                all_related[cell] = {'formula': all_formulas.get(cell, '')}

        # 检查是否有任何可模拟的随机源
        if not distribution_cells and not makeinput_cells and not native_random_cells:
            xlcAlert("未找到分布函数、MakeInput等随机函数。")
            return

        if not output_cells:
            xlcAlert("未找到输出单元格。")
            return

        print(f"找到分布单元格: {len(distribution_cells)} 个，输出单元格: {len(output_cells)} 个")

        app.StatusBar = "正在解析分布函数数组参数..."
        _preprocess_distribution_cells(distribution_cells, app)

        # ==================== 新增：标记纯分布单元格 ====================
        pure_cell_map = {}
        for cell_addr, funcs in distribution_cells.items():
            # 获取原始公式（已替换命名区域）
            original_formula = all_formulas.get(cell_addr.upper(), '')
            if not original_formula:
                # 如果 all_formulas 中没有，尝试从 all_related 获取
                cell_info = all_related.get(cell_addr.upper(), {})
                original_formula = cell_info.get('formula', '')
            # 清理公式：去除前导 = 、_xll. 和 @
            clean_formula = original_formula.lstrip('=').strip()
            clean_formula = re.sub(r'^_xll\.', '', clean_formula, flags=re.I)
            clean_formula = re.sub(r'^@', '', clean_formula).strip()
            # 如果该单元格只有一个分布函数，且提取的 full_match 等于清理后的公式，则为纯分布
            is_pure = False
            if len(funcs) == 1:
                func = funcs[0]
                full_match = func.get('full_match', '')
                if full_match:
                    # 清理 full_match（去除可能的 _xll. 前缀）
                    full_match_clean = re.sub(r'^_xll\.', '', full_match, flags=re.I)
                    full_match_clean = full_match_clean.strip()
                    # 比较时忽略大小写和空白
                    if full_match_clean.upper() == clean_formula.upper():
                        is_pure = True
            pure_cell_map[cell_addr.upper()] = is_pure
        # ================================================================

        for cell_addr, funcs in distribution_cells.items():
            for i, func in enumerate(funcs):
                if 'input_key' not in func:
                    func['input_key'] = f"{cell_addr}_{i+1}"
                if 'markers' not in func or not func['markers']:
                    args_text = func.get('args_text', '')
                    markers = _extract_markers_from_args_text(args_text)
                    if 'markers' in func and func['markers']:
                        markers.update(func['markers'])
                    func['markers'] = markers
                if 'parameters' not in func:
                    args_text = func.get('args_text', '')
                    params = []
                    for arg in parse_args_with_nested_functions(args_text):
                        if re.match(r'Drisk(?:Seed|Shift|Truncate|TruncateP|Truncate2|TruncateP2|Loc|Lock|Static|Name|Units|Category)\s*\(', arg, re.I):
                            continue
                        try:
                            params.append(float(arg))
                        except:
                            pass
                    func['parameters'] = params

        makeinput_key_to_info = {}
        for cell_addr, funcs in makeinput_cells.items():
            if funcs:
                mi_func = funcs[0]
                attrs = mi_func.get('attributes', {})
                if not attrs:
                    full_formula = all_formulas.get(cell_addr, '')
                    if full_formula:
                        attrs = extract_makeinput_attributes(full_formula)
                input_key = f"{cell_addr}_MAKEINPUT"
                attrs['is_makeinput'] = True
                makeinput_key_to_info[input_key] = attrs

        app.StatusBar = "正在进行拓扑排序……"
        all_related = {k.upper(): v for k, v in all_related.items()}
        all_formulas = {k.upper(): v for k, v in all_formulas.items()}
        graph, topo_order = build_numpy_dependency_graph(
            distribution_cells, makeinput_cells, simtable_cells, all_related, all_formulas, static_values,
            output_cells=output_cells
        )
        print(f"依赖图节点数: {len(topo_order)}")

        all_nodes = set(topo_order)
        if len(all_nodes) != len(topo_order):
            print("检测到循环依赖，向量化模拟无法处理，将回退到索引模拟")
            xlcAlert("检测到循环依赖，向量化模拟无法处理，将自动回退到索引模拟模式。")
            _apply_fallback_patch()
            try:
                run_index_simulation(n_iterations, scenario_count)
            finally:
                _remove_fallback_patch()
            return

        related_cells = set()
        input_key_to_func_info = {}
        for cell_addr, funcs in distribution_cells.items():
            for func in funcs:
                input_key = func.get('input_key')
                if input_key:
                    input_key_to_func_info[input_key] = func
        for node in topo_order:
            if node in input_key_to_func_info:
                cell = input_key_to_func_info[node].get('cell_addr')
                if cell:
                    related_cells.add(cell)
            elif '!' in node or re.match(r'^[A-Z]+\d+$', node, re.I):
                related_cells.add(node)

        for cell_addr in output_cells:
            related_cells.add(cell_addr)
        for cell_addr in makeinput_cells:
            related_cells.add(cell_addr)

        relevant_formulas = {}
        for addr in related_cells:
            if addr in all_formulas:
                relevant_formulas[addr] = all_formulas[addr]
            else:
                try:
                    if '!' in addr:
                        sheet_name, cell_addr = addr.split('!', 1)
                        sheet = app.ActiveWorkbook.Worksheets(sheet_name)
                    else:
                        sheet = app.ActiveSheet
                        cell_addr = addr
                    cell = sheet.Range(cell_addr)
                    formula = cell.Formula
                    if isinstance(formula, str) and formula.startswith('='):
                        relevant_formulas[addr] = _strip_xll_prefix(formula)
                except:
                    pass

        known_functions = _get_known_function_names()
        unsupported = _check_unsupported_functions(relevant_formulas, known_functions)

        if unsupported:
            err_text = "\n".join(unsupported)
            xlcAlert(
                f"以下与模拟相关的单元格包含向量化引擎不支持的函数，将自动回退到索引模拟模式：\n\n{err_text}\n\n正在启动索引模拟..."
            )
            _apply_fallback_patch()
            try:
                run_index_simulation(n_iterations, scenario_count)
            finally:
                _remove_fallback_patch()
            return

        full_match_to_key = {}
        input_key_to_func_info.clear()
        cell_to_input_keys = {}
        for cell_addr, funcs in distribution_cells.items():
            for func in funcs:
                input_key = func.get('input_key')
                if input_key:
                    input_key_to_func_info[input_key] = func
                    fm = func.get('full_match', '')
                    if fm:
                        fm_no_space = re.sub(r'\s+', '', fm)
                        key = f"{cell_addr}!{fm}"
                        full_match_to_key[key] = input_key
                        key_no_space = f"{cell_addr}!{fm_no_space}"
                        full_match_to_key[key_no_space] = input_key

                        # 修复：处理 @ 符号
                        if fm.startswith('@'):
                            fm_no_at = fm[1:]
                            key_with_cell_no_at = f"{cell_addr}!{fm_no_at}"
                            full_match_to_key[key_with_cell_no_at] = input_key
                            key_no_space_no_at = re.sub(r'\s+', '', key_with_cell_no_at)
                            full_match_to_key[key_no_space_no_at] = input_key
                        else:
                            # 添加带 @ 的版本以防万一
                            fm_with_at = '@' + fm
                            key_with_cell_with_at = f"{cell_addr}!{fm_with_at}"
                            full_match_to_key[key_with_cell_with_at] = input_key
                            key_no_space_with_at = re.sub(r'\s+', '', key_with_cell_with_at)
                            full_match_to_key[key_no_space_with_at] = input_key
                    cell_keys = cell_to_input_keys.setdefault(cell_addr.upper(), [])
                    cell_keys.append(input_key)

        # ========== 构建单元格公式字典（用于 SUBTOTAL 忽略自身） ==========
        cell_formulas = {}
        for addr, info in all_related.items():
            if 'formula' in info and info['formula']:
                cell_formulas[addr.upper()] = info['formula']
        for addr, formula in all_formulas.items():
            addr_upper = addr.upper()
            if addr_upper not in cell_formulas and formula:
                cell_formulas[addr_upper] = formula
        # ================================================================

        progress = OptimizedStatusBarProgress(app, n_iterations, scenario_count, len(topo_order))

        base_time_seed = int(time.time() * 1000) & 0x7fffffff

        all_scenario_results = []
        for scenario_idx in range(scenario_count):
            if progress.check_esc_key():
                print("用户取消模拟")
                break
            progress.start_new_scenario(scenario_idx)

            sim_id = create_simulation(n_iterations, "Numpy_MC")
            sim = get_simulation(sim_id)
            sim.set_scenario_info(scenario_count, scenario_idx)
            try:
                sim.workbook_name = app.ActiveWorkbook.Name
            except:
                pass

            simtable_values = {}
            for cell_addr, funcs in simtable_cells.items():
                if funcs:
                    values = funcs[0].get('values', [])
                    # 调试：打印解析出的值列表
                    print(f"Simtable 单元格 {cell_addr} 的值列表: {values}")
                    if scenario_idx < len(values):
                        val = values[scenario_idx]
                        try:
                            val_float = float(val)
                            norm_cell = normalize_cell_key(cell_addr) 
                            # 存储原始键
                            simtable_values[cell_addr] = val_float
                            # 同时存储只包含单元格地址的键（用于可能的工作表名不匹配）
                            if '!' in cell_addr:
                                _, coord = cell_addr.split('!', 1)
                                simtable_values[coord.upper()] = val_float
                            # 再存储标准化后的全大写键
                            simtable_values[cell_addr.upper()] = val_float
                            print(f"场景 {scenario_idx+1} -> {cell_addr} = {val_float}")
                        except Exception as e:
                            simtable_values[cell_addr] = ERROR_MARKER
                            print(f"警告：Simtable {cell_addr} 值 {val} 无法转换为浮点数")
                    else:
                        simtable_values[cell_addr] = ERROR_MARKER
                        print(f"警告：Simtable {cell_addr} 场景索引 {scenario_idx} 超出范围")

            app.StatusBar = f"场景 {scenario_idx+1}/{scenario_count}：正在生成样本（向量化）..."
            progress.update_node(0)

            xl_ns = _get_xl_ns()
            live_arrays: Dict[str, np.ndarray] = {}
            all_errors: List[str] = []

            seed_base_map = {}
            for node in topo_order:
                if node in input_key_to_func_info:
                    func_info = input_key_to_func_info[node]
                    markers = func_info.get('markers', {})
                    key_hash = zlib.crc32(node.encode()) & 0x7fffffff
                    seed_marker = markers.get('seed')
                    user_seed = None
                    if seed_marker is not None:
                        try:
                            if isinstance(seed_marker, str):
                                parts = seed_marker.split(',')
                                if len(parts) >= 2:
                                    user_seed = int(parts[1].strip())
                                elif len(parts) == 1:
                                    user_seed = int(parts[0].strip())
                            else:
                                user_seed = int(seed_marker)
                        except:
                            pass
                    if user_seed is not None:
                        final_base_seed = (user_seed + key_hash) & 0x7fffffff
                    else:
                        final_base_seed = (base_time_seed + key_hash) & 0x7fffffff
                    seed_base_map[node] = final_base_seed
                else:
                    seed_base_map[node] = None

            def _to_float(val):
                if val is None:
                    return 0.0
                if isinstance(val, str):
                    s = val.strip()
                    if s.endswith('%'):
                        try:
                            return float(s[:-1]) / 100.0
                        except:
                            return 0.0
                try:
                    return float(val)
                except (ValueError, TypeError):
                    return 0.0

            processed_nodes = 0
            for idx, node in enumerate(topo_order):
                if progress.is_cancelled():
                    break
                if node in live_arrays:
                    continue

                if node in input_key_to_func_info:
                    func_info = input_key_to_func_info[node]
                    cell_addr = func_info.get('cell_addr', node)
                    sheet = cell_addr.split('!')[0] if '!' in cell_addr else 'Sheet1'
                    coord = cell_addr.split('!')[-1] if '!' in cell_addr else cell_addr
                    full_match = func_info.get('full_match', '')
                    if not full_match:
                        continue

                    func_name_match = re.match(r'([A-Za-z_][A-Za-z0-9_.]*)\s*\(', full_match, re.I)
                    if not func_name_match:
                        continue
                    func_name = func_name_match.group(1)
                    inner_start = full_match.find('(') + 1
                    inner_end = full_match.rfind(')')
                    if inner_start <= 0 or inner_end <= inner_start:
                        continue
                    inner = full_match[inner_start:inner_end]

                    samples, err, msg = _sample_drisk_live(
                        func_name, inner, n_iterations, static_values, live_arrays, xl_ns,
                        sheet, cell_coord=coord, input_key=node,
                        seed_base=seed_base_map.get(node, base_time_seed),
                        simtable_values=simtable_values, _error_sink=all_errors,
                        markers=func_info.get('markers', {}),
                        method=method, cell_formulas=cell_formulas   # 传递采样方法
                    )
                    if err:
                        progress.clear()
                        xlcAlert(f"模拟中止：\n{msg}")
                        return
                    live_arrays[node.upper()] = samples
                    # ==================== 修复点：仅当单元格为纯分布时才映射 ====================
                    if pure_cell_map.get(cell_addr.upper(), False):
                        live_arrays[cell_addr.upper()] = samples
                    # ========================================================================

                else:
                    cell_addr = node
                    formula = None
                    if node in all_related:
                        formula = all_related[node].get('formula', '')
                    if not formula and node in all_formulas:
                        formula = all_formulas[node]

                    if formula:
                        sheet = cell_addr.split('!')[0] if '!' in cell_addr else 'Sheet1'
                        coord = cell_addr.split('!')[-1] if '!' in cell_addr else cell_addr
                        arr = _eval_formula_with_live(
                            formula, sheet, n_iterations, static_values, live_arrays, xl_ns,
                            cell_coord=coord, simtable_values=simtable_values, _error_sink=all_errors,
                            full_match_to_key=full_match_to_key,
                            cell_formulas=cell_formulas
                        )
                        live_arrays[cell_addr.upper()] = arr
                    else:
                        val = static_values.get(node)
                        if val is not None:
                            fval = _to_float(val)
                            arr = np.full(n_iterations, fval, dtype=float)
                        else:
                            try:
                                if '!' in node:
                                    sheet_name, cell_addr_only = node.split('!', 1)
                                    sheet = app.ActiveWorkbook.Worksheets(sheet_name)
                                else:
                                    sheet = app.ActiveSheet
                                    cell_addr_only = node
                                cell = sheet.Range(cell_addr_only)
                                cell_val = cell.Value2
                                fval = _to_float(cell_val)
                                arr = np.full(n_iterations, fval, dtype=float)
                                static_values[node] = fval
                            except Exception:
                                arr = np.full(n_iterations, np.nan, dtype=float)
                        live_arrays[cell_addr.upper()] = arr

                processed_nodes += 1
                progress.update_node(processed_nodes)

            if progress.is_cancelled():
                break

            # ---------- 应用相关性矩阵调整（秩相关） ----------
            if distribution_cells:
                try:
                    from corrmat_functions import (
                        _get_corrmat_info_from_cells,
                        _read_target_matrices,
                        _apply_rank_correlation,
                        _get_mapping_table,
                        _adjust_to_psd_with_weights,
                        _nearest_psd
                    )
                    # 提取所有分布函数中附加的 DriskCorrmat 信息
                    matrix_name_to_inputs = _get_corrmat_info_from_cells(distribution_cells,app)
                    if matrix_name_to_inputs:
                        matrix_names = set(matrix_name_to_inputs.keys())
                        target_matrices = _read_target_matrices(app, matrix_names)
                        
                        # 获取权重矩阵映射表
                        mapping = _get_mapping_table(app)
                        
                        # 检查目标矩阵是否正定，并尝试调整（支持权重矩阵）
                        for mat_name, C in list(target_matrices.items()):
                            # 查找是否有对应的权重矩阵
                            weight_name = None
                            for w_name, c_name in mapping.items():
                                if c_name == mat_name:
                                    weight_name = w_name
                                    break
                            
                            # 如果存在权重矩阵，尝试读取
                            weight_matrix = None
                            if weight_name:
                                weight_range = _get_named_range(app, weight_name)
                                if weight_range:
                                    weight_matrix = _read_matrix_from_range(weight_range)
                                    if weight_matrix is None:
                                        print(f"警告：无法读取权重矩阵 {weight_name}，将使用自动调整")
                                        weight_matrix = None
                                    else:
                                        # 检查权重矩阵大小
                                        if weight_matrix.shape != C.shape:
                                            print(f"警告：权重矩阵 {weight_name} 与相关性矩阵 {mat_name} 大小不一致，将使用自动调整")
                                            weight_matrix = None
                            
                            # 根据是否有权重矩阵选择调整方法
                            if weight_matrix is not None:
                                # 使用加权调整
                                print(f"使用权重矩阵 {weight_name} 调整矩阵 {mat_name}")
                                try:
                                    C_adj = _adjust_to_psd_with_weights(C, weight_matrix)
                                    # 确保对角线为1（adjust函数已处理）
                                    # 确保元素在[-1,1]内
                                    if np.any(C_adj > 1) or np.any(C_adj < -1):
                                        raise ValueError("调整后矩阵元素超出 [-1,1] 范围")
                                    target_matrices[mat_name] = C_adj
                                    print(f"矩阵 {mat_name} 加权调整成功")
                                except Exception as e:
                                    print(f"加权调整失败：{e}，回退到自动调整")
                                    # 回退到自动调整
                                    C_adj = _nearest_psd(C)
                                    np.fill_diagonal(C_adj, 1.0)
                                    target_matrices[mat_name] = C_adj
                            else:
                                # 自动调整
                                if not _is_positive_semidefinite(C):
                                    print(f"警告：矩阵 {mat_name} 不是半正定，将进行自动调整")
                                    C_adj = _nearest_psd(C)
                                    np.fill_diagonal(C_adj, 1.0)
                                    target_matrices[mat_name] = C_adj
                                else:
                                    # 即使半正定，也可能数值上不可逆，检查 Cholesky
                                    try:
                                        np.linalg.cholesky(C)
                                    except np.linalg.LinAlgError:
                                        print(f"警告：矩阵 {mat_name} 数值上不是正定，将进行扰动")
                                        C_adj = C + np.eye(C.shape[0]) * 1e-8
                                        target_matrices[mat_name] = C_adj
                                    else:
                                        target_matrices[mat_name] = C
                                # 确保元素在[-1,1]内
                                C_adj = target_matrices[mat_name]
                                if np.any(C_adj > 1) or np.any(C_adj < -1):
                                    print(f"警告：调整后矩阵元素超出 [-1,1] 范围，将裁剪")
                                    C_adj = np.clip(C_adj, -1.0, 1.0)
                                    target_matrices[mat_name] = C_adj
                        
                        # 执行秩相关调整（直接修改 live_arrays）
                        _apply_rank_correlation(
                            live_arrays,
                            matrix_name_to_inputs,
                            target_matrices,
                            n_iterations
                        )
                        # 验证调整效果：计算前两个输入样本的秩相关系数
                        if matrix_name_to_inputs:
                            first_mat = next(iter(matrix_name_to_inputs.values()))
                            if len(first_mat) >= 2:
                                key1 = first_mat[0][0]
                                key2 = first_mat[1][0]
                                if key1 in live_arrays and key2 in live_arrays:
                                    arr1 = live_arrays[key1].astype(float)
                                    arr2 = live_arrays[key2].astype(float)
                                    # 过滤有效值
                                    mask = ~(np.isnan(arr1) | np.isnan(arr2))
                                    if np.sum(mask) > 1:
                                        from scipy.stats import spearmanr
                                        rho, _ = spearmanr(arr1[mask], arr2[mask])
                                        print(f"调整后 {key1} 和 {key2} 的秩相关系数: {rho:.6f}")
                        print("已应用相关性矩阵调整。")
                except Exception as e:
                    print(f"相关性矩阵调整失败：{e}")
                    import traceback
                    traceback.print_exc()

            # ---------- 应用 Copula 调整 ----------
            if distribution_cells:
                try:
                    from copula_functions import _get_copula_info_from_cells, _apply_copula_rank_correlation
                    copula_info = _get_copula_info_from_cells(distribution_cells, app)
                    if copula_info:
                        # 注意：_apply_copula_rank_correlation 会直接修改 live_arrays
                        _apply_copula_rank_correlation(live_arrays, copula_info, n_iterations)
                        print("已应用 Copula 秩相关调整。")
                except Exception as e:
                    print(f"Copula 调整失败：{e}")
                    traceback.print_exc()
            
            if all_errors:
                error_msg = "\n".join(all_errors)
                print(f"场景 {scenario_idx+1} 生成样本时发生警告：{error_msg}")

            app.StatusBar = f"场景 {scenario_idx+1}/{scenario_count}：正在收集输出..."
            progress.update_node(processed_nodes)

            input_count = 0
            for input_key, func_info in input_key_to_func_info.items():
                cell_addr = func_info.get('cell_addr', input_key)
                if '!' in cell_addr:
                    sheet_name, pure_cell = cell_addr.split('!', 1)
                else:
                    sheet_name = app.ActiveSheet.Name
                    pure_cell = cell_addr
                key_upper = input_key.upper()
                if key_upper not in live_arrays:
                    print(f"警告：分布函数 {input_key} 的数据未在 live_arrays 中找到")
                    continue
                arr = live_arrays[key_upper]

                if not np.issubdtype(arr.dtype, np.number):
                    try:
                        arr = arr.astype(float)
                    except:
                        arr_obj = np.full(arr.shape, ERROR_MARKER, dtype=object)
                        arr_list = arr_obj.tolist()
                        if isinstance(arr_list, list) and len(arr_list) > 0 and isinstance(arr_list[0], list):
                            arr_list = [item for sublist in arr_list for item in sublist]
                        if len(arr_list) != n_iterations:
                            if len(arr_list) > n_iterations:
                                arr_list = arr_list[:n_iterations]
                            else:
                                arr_list.extend([ERROR_MARKER] * (n_iterations - len(arr_list)))
                        if sim.add_input_result(pure_cell + "_" + input_key.split('_')[-1], arr_list, sheet_name, func_info.get('markers', {})):
                            input_count += 1
                        continue

                arr_obj = arr.astype(object)
                nan_mask = np.isnan(arr)
                inf_mask = np.isinf(arr)
                mask = nan_mask | inf_mask
                arr_obj[mask] = ERROR_MARKER

                if arr_obj.ndim != 1:
                    arr_obj = arr_obj.flatten()
                arr_list = arr_obj.tolist()
                if len(arr_list) != n_iterations:
                    if len(arr_list) > n_iterations:
                        arr_list = arr_list[:n_iterations]
                    else:
                        arr_list.extend([ERROR_MARKER] * (n_iterations - len(arr_list)))

                suffix = input_key.split('_')[-1] if '_' in input_key else ''
                pure_name = f"{pure_cell}_{suffix}" if suffix else pure_cell
                if sim.add_input_result(pure_name, arr_list, sheet_name, func_info.get('markers', {})):
                    input_count += 1

            for makeinput_key, attrs in makeinput_key_to_info.items():
                cell_addr = makeinput_key.replace('_MAKEINPUT', '')
                if '!' in cell_addr:
                    sheet_name, pure_cell = cell_addr.split('!', 1)
                else:
                    sheet_name = app.ActiveSheet.Name
                    pure_cell = cell_addr
                key_upper = cell_addr.upper()
                if key_upper not in live_arrays:
                    print(f"警告：MakeInput 单元格 {cell_addr} 的数据未在 live_arrays 中找到")
                    continue
                arr = live_arrays[key_upper]

                if not np.issubdtype(arr.dtype, np.number):
                    try:
                        arr = arr.astype(float)
                    except:
                        arr_obj = np.full(arr.shape, ERROR_MARKER, dtype=object)
                        arr_list = arr_obj.tolist()
                        if isinstance(arr_list, list) and len(arr_list) > 0 and isinstance(arr_list[0], list):
                            arr_list = [item for sublist in arr_list for item in sublist]
                        if len(arr_list) != n_iterations:
                            if len(arr_list) > n_iterations:
                                arr_list = arr_list[:n_iterations]
                            else:
                                arr_list.extend([ERROR_MARKER] * (n_iterations - len(arr_list)))
                        if sim.add_input_result(f"{pure_cell}_MAKEINPUT", arr_list, sheet_name, attrs):
                            input_count += 1
                        continue

                arr_obj = arr.astype(object)
                nan_mask = np.isnan(arr)
                inf_mask = np.isinf(arr)
                mask = nan_mask | inf_mask
                arr_obj[mask] = ERROR_MARKER

                if arr_obj.ndim != 1:
                    arr_obj = arr_obj.flatten()
                arr_list = arr_obj.tolist()
                if len(arr_list) != n_iterations:
                    if len(arr_list) > n_iterations:
                        arr_list = arr_list[:n_iterations]
                    else:
                        arr_list.extend([ERROR_MARKER] * (n_iterations - len(arr_list)))

                if sim.add_input_result(f"{pure_cell}_MAKEINPUT", arr_list, sheet_name, attrs):
                    input_count += 1

            output_count = 0
            for cell_addr, out_info in output_cells.items():
                key = cell_addr.upper()
                if key in live_arrays:
                    arr = live_arrays[key]

                    if not np.issubdtype(arr.dtype, np.number):
                        try:
                            arr = arr.astype(float)
                        except:
                            arr_obj = np.full(arr.shape, ERROR_MARKER, dtype=object)
                            arr_list = arr_obj.tolist()
                            if isinstance(arr_list, list) and len(arr_list) > 0 and isinstance(arr_list[0], list):
                                arr_list = [item for sublist in arr_list for item in sublist]
                            if len(arr_list) != n_iterations:
                                if len(arr_list) > n_iterations:
                                    arr_list = arr_list[:n_iterations]
                                else:
                                    arr_list.extend([ERROR_MARKER] * (n_iterations - len(arr_list)))
                            if sim.add_output_result(cell_addr.split('!')[-1] if '!' in cell_addr else cell_addr, arr_list, cell_addr.split('!')[0] if '!' in cell_addr else app.ActiveSheet.Name, out_info):
                                output_count += 1
                            continue

                    arr_obj = arr.astype(object)
                    nan_mask = np.isnan(arr)
                    inf_mask = np.isinf(arr)
                    mask = nan_mask | inf_mask
                    arr_obj[mask] = ERROR_MARKER

                    if arr_obj.ndim != 1:
                        arr_obj = arr_obj.flatten()
                    arr_list = arr_obj.tolist()
                    if len(arr_list) != n_iterations:
                        if len(arr_list) > n_iterations:
                            arr_list = arr_list[:n_iterations]
                        else:
                            arr_list.extend([ERROR_MARKER] * (n_iterations - len(arr_list)))

                    if '!' in cell_addr:
                        sheet_name, cell_addr_only = cell_addr.split('!')
                    else:
                        sheet_name = app.ActiveSheet.Name
                        cell_addr_only = cell_addr

                    if sim.add_output_result(cell_addr_only, arr_list, sheet_name, out_info):
                        output_count += 1

            sim.compute_input_statistics()
            sim.compute_output_statistics()
            sim.set_end_time()

            all_scenario_results.append({
                'sim_id': sim_id,
                'scenario_index': scenario_idx,
                'output_count': output_count,
                'input_count': input_count,
                'skipped_outputs': len(output_cells) - output_count
            })

            print(f"场景 {scenario_idx+1} 完成，模拟ID={sim_id}，输入={input_count}，输出={output_count}")
            app.StatusBar = f"场景 {scenario_idx+1}/{scenario_count} 完成"
            progress.update_node(processed_nodes)

        try:
            app.StatusBar = False
        except:
            pass

        # 所有场景处理完毕后，强制重算工作簿
        try:
            app.Calculate()
        except Exception as e:
            print(f"重算工作簿失败: {e}")

        elapsed = time.time() - progress.start_time if progress else 0
        msg = f"向量化蒙特卡洛模拟完成！\n"
        msg += f"采样方法: {'拉丁超立方 (LHC)' if method == 'LHC' else 'Sobol 低差异序列' if method == 'SOBOL' else '蒙特卡洛 (MC)'}\n"
        msg += f"场景数: {scenario_count}\n每场景迭代: {n_iterations:,}\n总耗时: {elapsed:.2f}秒\n"
        for r in all_scenario_results:
            msg += f"场景 {r['scenario_index']+1}: 模拟ID={r['sim_id']} (Input={r['input_count']}, Output={r['output_count']})\n"
        xlcAlert(msg)

    except Exception as e:
        import traceback
        error_details = traceback.format_exc()
        print(f"向量化模拟发生未捕获异常: {e}\n{error_details}")
        try:
            if progress:
                progress.clear()
            if app:
                app.StatusBar = False
        except:
            pass
        xlcAlert(f"向量化模拟失败，出现未预期错误：{str(e)}\n\n详细信息已记录至控制台。")
        
# ==================== 宏入口 ====================
@xl_macro
def DriskNumpyMC():
    """向量化蒙特卡洛模拟主宏"""
    app = None
    try:
        from com_fixer import _safe_excel_app
        app = _safe_excel_app()
    except ImportError:
        import win32com.client
        app = win32com.client.Dispatch("Excel.Application")

    try:
        iter_in = app.InputBox(
            Prompt="请输入模拟次数（默认1000）：",
            Title="DriskNumpyMC - 迭代次数",
            Type=1,
            Default="1000"
        )
        if iter_in is False:
            return
        n_iterations = int(iter_in)
        if n_iterations < 1:
            xlcAlert("模拟次数必须是正整数")
            return
    except Exception as e:
        xlcAlert(f"无效的迭代次数: {str(e)}")
        return

    scenario_count = 1
    try:
        dist_cells, simtable_cells, _, _, _, _ = find_all_simulation_cells_in_workbook(app)
        max_len = 0
        for funcs in simtable_cells.values():
            for f in funcs:
                vals = f.get('values', [])
                if len(vals) > max_len:
                    max_len = len(vals)
        if max_len > 0:
            try:
                sc_in = app.InputBox(
                    Prompt=f"检测到 Simtable，最大数组长度为 {max_len}。请输入场景数（默认 {max_len}）：",
                    Title="DriskNumpyMC - 场景数",
                    Type=1,
                    Default=str(max_len)
                )
                if sc_in is False:
                    return
                scenario_count = int(sc_in)
                if scenario_count < 1 :
                    xlcAlert("场景数必须是正整数")
                    return
            except:
                scenario_count = max_len
    except Exception as e:
        print(f"获取场景数失败: {e}")

    run_numpy_simulation(n_iterations, scenario_count, method='MC')

@xl_macro
def DriskNumpyLHC():
    """拉丁超立方模拟主宏"""
    app = None
    try:
        from com_fixer import _safe_excel_app
        app = _safe_excel_app()
    except ImportError:
        import win32com.client
        app = win32com.client.Dispatch("Excel.Application")

    try:
        iter_in = app.InputBox(
            Prompt="请输入模拟次数（默认1000）：",
            Title="DriskNumpyLHC - 迭代次数",
            Type=1,
            Default="1000"
        )
        if iter_in is False:
            return
        n_iterations = int(iter_in)
        if n_iterations < 1:
            xlcAlert("模拟次数必须是正整数")
            return
    except Exception as e:
        xlcAlert(f"无效的迭代次数: {str(e)}")
        return

    scenario_count = 1
    try:
        dist_cells, simtable_cells, _, _, _, _ = find_all_simulation_cells_in_workbook(app)
        max_len = 0
        for funcs in simtable_cells.values():
            for f in funcs:
                vals = f.get('values', [])
                if len(vals) > max_len:
                    max_len = len(vals)
        if max_len > 0:
            try:
                sc_in = app.InputBox(
                    Prompt=f"检测到 Simtable，最大数组长度为 {max_len}。请输入场景数（默认 {max_len}）：",
                    Title="DriskNumpyLHC - 场景数",
                    Type=1,
                    Default=str(max_len)
                )
                if sc_in is False:
                    return
                scenario_count = int(sc_in)
                if scenario_count < 1:
                    xlcAlert("场景数必须是正整数")
                    return
            except:
                scenario_count = max_len
    except Exception as e:
        print(f"获取场景数失败: {e}")

    run_numpy_simulation(n_iterations, scenario_count, method='LHC')

@xl_macro
def DriskNumpySobol():
    """Sobol序列准蒙特卡洛模拟主宏（自动将迭代次数调整为2的幂）"""
    app = None
    try:
        from com_fixer import _safe_excel_app
        app = _safe_excel_app()
    except ImportError:
        import win32com.client
        app = win32com.client.Dispatch("Excel.Application")

    try:
        # 默认值设为 1024
        iter_in = app.InputBox(
            Prompt="请输入模拟次数（默认1024，推荐使用2的幂次如1024、2048等）：",
            Title="DriskNumpySobol - 迭代次数",
            Type=1,
            Default="1024"
        )
        if iter_in is False:
            return
        n_input = int(iter_in)
        if n_input < 1:
            xlcAlert("模拟次数必须是正整数")
            return
        # 计算大于等于 n_input 的最小 2 的幂
        n_iterations = 1
        while n_iterations < n_input:
            n_iterations <<= 1
        if n_iterations != n_input:
            # 告知用户调整后的迭代次数
            xlcAlert(f"为获得最佳 Sobol 序列均匀性，已将迭代次数从 {n_input} 调整为 {n_iterations}（2的幂次）。")
    except Exception as e:
        xlcAlert(f"无效的迭代次数: {str(e)}")
        return

    scenario_count = 1
    try:
        # 查找 Simtable 单元格以确定场景数
        dist_cells, simtable_cells, _, _, _, _ = find_all_simulation_cells_in_workbook(app)
        max_len = 0
        for funcs in simtable_cells.values():
            for f in funcs:
                vals = f.get('values', [])
                if len(vals) > max_len:
                    max_len = len(vals)
        if max_len > 0:
            try:
                sc_in = app.InputBox(
                    Prompt=f"检测到 Simtable，最大数组长度为 {max_len}。请输入场景数（默认 {max_len}）：",
                    Title="DriskNumpySobol - 场景数",
                    Type=1,
                    Default=str(max_len)
                )
                if sc_in is False:
                    return
                scenario_count = int(sc_in)
                if scenario_count < 1:
                    xlcAlert("场景数必须是正整数")
                    return
            except:
                scenario_count = max_len
    except Exception as e:
        print(f"获取场景数失败: {e}")

    run_numpy_simulation(n_iterations, scenario_count, method='SOBOL')   
