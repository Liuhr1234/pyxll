# distribution_functions.py - 关键修复版 v8.0
"""分布函数模块 - 修复嵌套函数和引用单元格值问题 - 增强嵌套函数支持 - 完善单边截断退化处理 - 修正属性处理顺序 - 增加截断参数有效性检查 - 增加支持范围检查 - 修复loc/lock静态均值计算 - 支持花括号数组参数 - 支持区域引用自动转换 - 修复Cumul/Discrete解析 - 修复三角分布参数顺序错误 - 统一使用分布类计算截断均值 (涵盖所有分布) - 修复截断无效时返回#ERROR! - 修复百分位数超出范围返回#ERROR! - 新增对未知属性函数（Excel错误值）的检测，静态模式下返回#ERROR!"""
import re
import numpy as np
import math
import threading
import time
from pyxll import xl_func as _pyxll_xl_func, xl_app
from typing import List, Tuple, Dict, Any, Callable, Optional
# 导入属性处理函数（包括静态模式设置）
from attribute_functions import extract_markers_from_args, get_static_mode, is_marker_string, set_static_mode, create_marker_string
# 导入分布注册表
from constants import (
    DISTRIBUTION_REGISTRY, get_distribution_info, get_distribution_type,
    validate_distribution_params, DISTRIBUTION_FUNCTION_NAMES,
    get_distribution_support
)


def _extract_arg_names_from_xl_signature(signature: Optional[str]) -> List[str]:
    """从 xl_func 签名提取参数名（忽略 var* 占位参数）。"""
    if not isinstance(signature, str) or not signature.strip():
        return []
    left = signature.split(":", 1)[0]
    names: List[str] = []
    for part in left.split(","):
        token = part.strip()
        if not token:
            continue
        pieces = token.split()
        if len(pieces) < 2:
            continue
        arg_name = pieces[-1].strip()
        if not arg_name or "*" in arg_name:
            continue
        names.append(arg_name)
    return names


def _build_auto_arg_descriptions(func_name: str, signature: Optional[str]) -> Dict[str, str]:
    """根据注册表文案构建 PyXLL arg_descriptions。"""
    info = get_distribution_info(func_name) or {}
    labels = list(info.get("ui_param_labels", []) or [])
    param_descs = list(info.get("param_descriptions", []) or [])
    arg_names = _extract_arg_names_from_xl_signature(signature)

    mapping: Dict[str, str] = {}
    for idx, arg_name in enumerate(arg_names):
        label = str(labels[idx]).strip() if idx < len(labels) and labels[idx] is not None else ""
        desc = str(param_descs[idx]).strip() if idx < len(param_descs) and param_descs[idx] is not None else ""
        # 函数向导已单独显示参数名，这里只给描述，避免出现 A [A] 这种重复展示。
        if desc:
            mapping[arg_name] = desc
        elif label and label.lower() != arg_name.lower():
            mapping[arg_name] = label
    return mapping

def _get_formula_from_caller() -> Optional[str]:
    try:
        app = xl_app()
        caller = app.Caller
        if hasattr(caller, "Formula"):
            formula = caller.Formula
            if formula and isinstance(formula, str):
                return formula
    except Exception:
        pass
    return None

def _extract_function_args_from_formula(formula: str, func_name: str) -> Optional[List[str]]:
    """
    从 Excel 公式字符串中提取指定函数的参数列表。
    支持函数调用前后有其他表达式（如 DriskOutput() + DriskCompound(...)）。
    """
    if not isinstance(formula, str) or not formula.strip():
        return None
    # 不区分大小写匹配函数名
    pattern = re.compile(rf'\b{re.escape(func_name)}\s*\(', re.IGNORECASE)
    match = pattern.search(formula)
    if not match:
        return None
    start = match.end() - 1  # '(' 的位置
    # 括号匹配
    balance = 1
    i = start + 1
    while i < len(formula) and balance > 0:
        if formula[i] == '(':
            balance += 1
        elif formula[i] == ')':
            balance -= 1
        i += 1
    if balance != 0:
        return None
    inner = formula[start+1:i-1]   # 括号内的原始内容
    # 按逗号分割参数，忽略括号内的逗号
    args = []
    current = []
    balance = 0
    for ch in inner:
        if ch == '(':
            balance += 1
            current.append(ch)
        elif ch == ')':
            balance -= 1
            current.append(ch)
        elif ch == ',' and balance == 0:
            args.append(''.join(current).strip())
            current = []
        else:
            current.append(ch)
    if current:
        args.append(''.join(current).strip())
    return args

def _get_cell_formula(cell_ref: str) -> Optional[str]:
    """从单元格地址获取其公式字符串（若单元格为空或非公式则返回 None）。"""
    try:
        from pyxll import xl_app
        app = xl_app()
        if '!' in cell_ref:
            sheet_name, addr = cell_ref.split('!', 1)
            rng = app.ActiveWorkbook.Worksheets(sheet_name).Range(addr)
        else:
            rng = app.ActiveSheet.Range(cell_ref)
        formula = rng.Formula
        if formula and isinstance(formula, str) and formula.startswith('='):
            return formula
    except Exception:
        pass
    return None

def _resolve_distribution_formula(arg_str: str) -> str:
    """
    将参数原始字符串解析为分布公式字符串。
    支持：
        - 已经是公式字符串（以 '=' 开头）
        - 单元格引用（如 C16, Sheet2!A1）
        - 数值常量（直接返回，但 Compound/Splice 通常需要公式）
    """
    arg = arg_str.strip()
    if not arg:
        return ""
    if arg.startswith('='):
        return arg
    # 单元格引用匹配
    cell_pattern = re.compile(r'^([A-Za-z]+!\$?[A-Z]+[$]?\d+|\$?[A-Z]+[$]?\d+)$', re.IGNORECASE)
    if cell_pattern.match(arg):
        formula = _get_cell_formula(arg)
        if formula:
            return formula
    # 数值常量或其他，直接返回原字符串
    return arg

def _resolve_distribution_formula(arg_str: str) -> str:
    arg = arg_str.strip()
    if not arg:
        return ""

    # 已经是完整公式（以 '=' 开头）
    if arg.startswith('='):
        return arg

    # 处理以 '@' 开头的动态数组公式（如 @DriskLognorm(120,52)）
    if arg.startswith('@'):
        # 去掉 @ 后，可能得到 =Drisk... 或直接 Drisk...
        rest = arg[1:]
        if rest.startswith('='):
            return rest
        else:
            return '=' + rest

    # 检查是否为单元格引用（如 C16, Sheet2!A1, $A$1 等）
    cell_pattern = re.compile(r'^([A-Za-z]+!\$?[A-Z]+[$]?\d+|\$?[A-Z]+[$]?\d+)$', re.IGNORECASE)
    if cell_pattern.match(arg):
        formula = _get_cell_formula(arg)
        if formula:
            return formula
        # 如果无法获取公式，尝试获取单元格的值作为常数
        try:
            from pyxll import xl_app
            app = xl_app()
            rng = app.Range(arg) if '!' not in arg else app.ActiveWorkbook.Worksheets(arg.split('!')[0]).Range(arg.split('!')[1])
            value = rng.Value2
            if value is not None:
                # 转换为字符串（数值或文本）
                return str(value)
        except:
            pass
        # 如果仍然失败，返回原始引用（后续解析会失败，但不会崩溃）
        return arg

    # 可能是数值常数（如 5, 3.14）
    try:
        float(arg)
        return arg  # 直接返回数值字符串
    except ValueError:
        pass

    # 其他情况（可能是分布函数字符串如 DriskPoisson(13.8)）补上 '=' 前缀
    if arg.startswith('Drisk'):
        return '=' + arg

    # 最终返回原字符串
    return arg

def _extract_compound_raw_args_from_formula(formula: str) -> Optional[List[str]]:
    raw_args = _extract_function_args_from_formula(formula, "DriskCompound")
    if not raw_args:
        return None
    resolved = []
    for a in raw_args:
        resolved.append(_resolve_distribution_formula(a))
    return resolved

def _extract_splice_raw_args_from_formula(formula: str) -> Optional[List[str]]:
    raw_args = _extract_function_args_from_formula(formula, "DriskSplice")
    if not raw_args:
        return None
    resolved = []
    for a in raw_args:
        resolved.append(_resolve_distribution_formula(a))
    return resolved

def xl_func(signature=None, **kwargs):
    """
    对 pyxll.xl_func 做最小包装：自动补充分布函数参数说明。
    仅影响函数面板注册文案，不改变分布计算逻辑。
    """
    # 兼容 @xl_func 直接修饰
    if callable(signature):
        return _pyxll_xl_func(signature)

    def _decorator(func):
        local_kwargs = dict(kwargs)
        func_name = getattr(func, "__name__", "")
        if func_name in DISTRIBUTION_FUNCTION_NAMES and "arg_descriptions" not in local_kwargs:
            auto_arg_desc = _build_auto_arg_descriptions(func_name, signature)
            if auto_arg_desc:
                local_kwargs["arg_descriptions"] = auto_arg_desc
        return _pyxll_xl_func(signature, **local_kwargs)(func)

    return _decorator
# 导入属性函数列表，用于识别已知属性函数（避免误判）
from formula_parser import ATTRIBUTE_FUNCTIONS, parse_complete_formula
# 导入已有分布生成器
from dist_bernoulli import bernoulli_generator_single as bernoulli_generator, bernoulli_cdf, bernoulli_ppf
from dist_triang import triang_generator_single as triang_generator, triang_cdf, triang_ppf
from dist_binomial import binomial_generator_single as binomial_generator, binomial_cdf, binomial_ppf
from dist_erf import erf_generator_single as erf_generator, erf_cdf, erf_ppf
from dist_erlang import erlang_generator_single as erlang_generator, erlang_cdf, erlang_ppf
from dist_cauchy import cauchy_generator_single as cauchy_generator, cauchy_cdf, cauchy_ppf
from dist_dagum import dagum_generator_single as dagum_generator, dagum_cdf, dagum_ppf, dagum_raw_mean
from dist_doubletriang import doubletriang_generator_single as doubletriang_generator, doubletriang_cdf, doubletriang_ppf, doubletriang_raw_mean
from dist_extvalue import extvalue_generator_single as extvalue_generator, extvalue_cdf, extvalue_ppf, extvalue_raw_mean
from dist_extvaluemin import extvaluemin_generator_single as extvaluemin_generator, extvaluemin_cdf, extvaluemin_ppf, extvaluemin_raw_mean
from dist_fatiguelife import fatiguelife_generator_single as fatiguelife_generator, fatiguelife_cdf, fatiguelife_ppf, fatiguelife_raw_mean
from dist_frechet import frechet_generator_single as frechet_generator, frechet_cdf, frechet_ppf, frechet_raw_mean
from dist_hypsecant import hypsecant_generator_single as hypsecant_generator, hypsecant_cdf, hypsecant_ppf, hypsecant_raw_mean
from dist_johnsonsb import johnsonsb_generator_single as johnsonsb_generator, johnsonsb_cdf, johnsonsb_ppf, johnsonsb_raw_mean
from dist_johnsonsu import johnsonsu_generator_single as johnsonsu_generator, johnsonsu_cdf, johnsonsu_ppf, johnsonsu_raw_mean
from dist_kumaraswamy import kumaraswamy_generator_single as kumaraswamy_generator, kumaraswamy_cdf, kumaraswamy_ppf, kumaraswamy_raw_mean
from dist_laplace import laplace_generator_single as laplace_generator, laplace_cdf, laplace_ppf, laplace_raw_mean
from dist_logistic import logistic_generator_single as logistic_generator, logistic_cdf, logistic_ppf, logistic_raw_mean
from dist_loglogistic import loglogistic_generator_single as loglogistic_generator, loglogistic_cdf, loglogistic_ppf, loglogistic_raw_mean
from dist_lognorm import lognorm_generator_single as lognorm_generator, lognorm_cdf, lognorm_ppf, lognorm_raw_mean
from dist_lognorm2 import lognorm2_generator_single as lognorm2_generator, lognorm2_cdf, lognorm2_ppf, lognorm2_raw_mean
from dist_betageneral import betageneral_generator_single as betageneral_generator, betageneral_cdf, betageneral_ppf, betageneral_raw_mean
from dist_betasubj import betasubj_generator_single as betasubj_generator, betasubj_cdf, betasubj_ppf, betasubj_raw_mean
from dist_burr12 import burr12_generator_single as burr12_generator, burr12_cdf, burr12_ppf, burr12_raw_mean
from dist_compound import compound_generator_single as compound_generator, compound_cdf, compound_ppf, compound_raw_mean
from dist_splice import splice_generator_single as splice_generator, splice_cdf, splice_ppf, splice_raw_mean
from dist_pert import pert_generator_single as pert_generator, pert_cdf, pert_ppf, pert_raw_mean
from dist_reciprocal import reciprocal_generator_single as reciprocal_generator, reciprocal_cdf, reciprocal_ppf, reciprocal_raw_mean
from dist_rayleigh import rayleigh_generator_single as rayleigh_generator, rayleigh_cdf, rayleigh_ppf, rayleigh_raw_mean
from dist_weibull import weibull_generator_single as weibull_generator, weibull_cdf, weibull_ppf, weibull_raw_mean
from dist_pearson5 import pearson5_generator_single as pearson5_generator, pearson5_cdf, pearson5_ppf, pearson5_raw_mean
from dist_pearson6 import pearson6_generator_single as pearson6_generator, pearson6_cdf, pearson6_ppf, pearson6_raw_mean
from dist_pareto2 import pareto2_generator_single as pareto2_generator, pareto2_cdf, pareto2_ppf, pareto2_raw_mean
from dist_pareto import pareto_generator_single as pareto_generator, pareto_cdf, pareto_ppf, pareto_raw_mean
from dist_levy import levy_generator_single as levy_generator, levy_cdf, levy_ppf, levy_raw_mean
from dist_general import general_generator_single as general_generator, general_cdf, general_ppf, general_raw_mean, _parse_general_arrays as general_parse_arrays
from dist_histogrm import histogrm_generator_single as histogrm_generator, histogrm_cdf, histogrm_ppf, histogrm_raw_mean, _parse_histogrm_p_table as histogrm_parse_p_table
from dist_negbin import negbin_generator_single as negbin_generator, negbin_cdf, negbin_ppf
from dist_invgauss import invgauss_generator_single as invgauss_generator, invgauss_cdf, invgauss_ppf
from dist_duniform import duniform_generator_single as duniform_generator, duniform_cdf, duniform_ppf
from dist_geomet import geomet_generator_single as geomet_generator, geomet_cdf, geomet_ppf
from dist_hypergeo import hypergeo_generator_single as hypergeo_generator, hypergeo_cdf, hypergeo_ppf
from dist_intuniform import intuniform_generator_single as intuniform_generator, intuniform_cdf, intuniform_ppf
from dist_trigen import trigen_generator_single as trigen_generator, trigen_cdf, trigen_ppf
from dist_cumul import cumul_generator_single as cumul_generator, cumul_cdf, cumul_ppf, _parse_arrays as cumul_parse_arrays
from dist_discrete import discrete_generator_single as discrete_generator, discrete_cdf, discrete_ppf
# 导入截断均值函数（用于 loc/lock 标记）
from dist_triang import triang_truncated_mean
from dist_cumul import cumul_truncated_mean
from dist_discrete import discrete_truncated_mean
from dist_trigen import _convert_trigen_to_triang
# ==================== 导入所有分布类（用于统一截断均值计算）====================
from distribution_base import (
    GammaDistribution, BetaDistribution, ChiSquaredDistribution,
    FDistribution, TDistribution, PoissonDistribution,
    ExponentialDistribution, UniformDistribution, NormalDistribution
)
from dist_bernoulli import BernoulliDistribution
from dist_binomial import BinomialDistribution
from dist_erf import ErfDistribution
from dist_erlang import ErlangDistribution
from dist_cauchy import CauchyDistribution
from dist_dagum import DagumDistribution
from dist_doubletriang import DoubleTriangDistribution
from dist_extvalue import ExtvalueDistribution
from dist_extvaluemin import ExtvalueMinDistribution
from dist_fatiguelife import FatigueLifeDistribution
from dist_frechet import FrechetDistribution
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
from dist_general import GeneralDistribution
from dist_histogrm import HistogrmDistribution
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
# ==================== 导入scipy以精确计算截断均值 ====================
try:
    import scipy.stats as sps
    SCIPY_AVAILABLE = True
except ImportError:
    SCIPY_AVAILABLE = False
# ==================== 模拟状态管理器（用于在UDF中记录随机数） ====================
class SimulationStateManager:
    """模拟状态管理器 - 用于在UDF中记录随机数 - 增强嵌套函数支持"""
    _instance = None
    _lock = threading.RLock()
    def __new__(cls):
        with cls._lock:
            if cls._instance is None:
                cls._instance = super().__new__(cls)
                cls._instance._init_state()
            return cls._instance
    def _init_state(self):
        """初始化状态"""
        self.current_iteration = 0  # 当前迭代次数
        self.total_iterations = 0   # 总迭代次数
        self.simulation_running = False  # 模拟是否正在进行
        self.distribution_cache = {}  # 分布函数缓存 {stable_key: np.array}
        self.distribution_generators = {}  # 分布生成器 {stable_key: DistributionGenerator}
        self.stable_key_to_func_info = {}  # 稳定键到函数信息的映射
        self.initialized_keys = set()  # 已初始化的稳定键
        self.func_signature_to_stable_key = {}  # 函数签名到稳定键的映射
        self.param_based_signatures = {}  # 基于参数值的签名映射
        self.cell_value_cache = {}  # 单元格值缓存 {cell_addr: np.array} - 解决引用值不一致问题
        self.cell_value_locks = {}  # 单元格值锁，确保线程安全
        self.nested_relationships = {}  # 嵌套函数关系映射 {parent_key: [child_key1, child_key2]}
        self.nested_children_cache = {}  # 嵌套子函数缓存
        self.nested_parent_cache = {}  # 嵌套父函数缓存
        self.registered_stable_keys = set()  # 已注册的稳定键
        self.debug_log = []  # 调试日志
        self.last_debug_print = 0  # 最后调试打印时间
        self.cell_unique_counter = {}  # 单元格唯一计数器，解决下拉重复问题
    def log_debug(self, message: str):
        """记录调试信息"""
        current_time = time.time()
        timestamp = time.strftime("%H:%M:%S", time.localtime(current_time))
        debug_msg = f"[{timestamp}] {message}"
        self.debug_log.append(debug_msg)
        # 控制输出频率，避免日志过多
        if current_time - self.last_debug_print > 2.0 or len(self.debug_log) % 20 == 0:
            print(debug_msg)
            self.last_debug_print = current_time
    def initialize_simulation(self, total_iterations: int):
        """初始化模拟"""
        with self._lock:
            self.current_iteration = 0
            self.total_iterations = total_iterations
            self.simulation_running = True
            self.distribution_cache = {}
            self.distribution_generators = {}
            self.stable_key_to_func_info = {}
            self.initialized_keys = set()
            self.func_signature_to_stable_key = {}
            self.param_based_signatures = {}
            self.cell_value_cache = {}  # 清空单元格值缓存
            self.cell_value_locks = {}
            self.nested_relationships = {}
            self.nested_children_cache = {}
            self.nested_parent_cache = {}
            self.registered_stable_keys = set()
            self.cell_unique_counter = {}  # 清空计数器
            self.debug_log = []
            self.last_debug_print = 0
            # 关键修复：模拟运行时强制关闭静态模式，确保 lock 正常生效
            set_static_mode(False)
            self.log_debug(f"模拟状态管理器初始化: 总迭代次数={total_iterations}, 静态模式已关闭")
    def set_current_iteration(self, iteration: int):
        """设置当前迭代次数"""
        with self._lock:
            self.current_iteration = iteration
            self.log_debug(f"设置当前迭代: {iteration}")
    def get_current_iteration(self) -> int:
        """获取当前迭代次数"""
        with self._lock:
            return self.current_iteration
    def is_simulation_running(self) -> bool:
        """检查模拟是否正在进行"""
        with self._lock:
            return self.simulation_running
    def register_distribution(self, stable_key: str, func_info: Dict[str, Any]):
        """注册分布函数 - 修复：确保基于参数的签名也能匹配，支持嵌套函数"""
        with self._lock:
            if stable_key in self.registered_stable_keys:
                self.log_debug(f"稳定键已注册，跳过: {stable_key}")
                return
            if stable_key not in self.initialized_keys:
                # 初始化缓存数组
                self.distribution_cache[stable_key] = np.full(self.total_iterations, np.nan, dtype=float)
                self.stable_key_to_func_info[stable_key] = func_info
                self.initialized_keys.add(stable_key)
                self.registered_stable_keys.add(stable_key)
                # 创建函数签名映射
                func_name = func_info.get('func_name', '')
                dist_params = func_info.get('dist_params', [])
                args_text_normalized = func_info.get('args_text_normalized', '')
                is_nested = func_info.get('is_nested', False)
                depth = func_info.get('depth', 0)
                parent_key = func_info.get('parent_key', None)
                self.log_debug(f"注册分布函数: 稳定键={stable_key}, 函数名={func_name}, 嵌套={is_nested}, 深度={depth}, 父键={parent_key}")
                self.log_debug(f"参数={dist_params}, 参数文本={args_text_normalized}")
                # 对于嵌套函数，添加额外的标识信息
                if is_nested:
                    func_name_with_nested = f"{func_name}_nested{depth}"
                else:
                    func_name_with_nested = func_name
                # 创建标准化的函数签名
                if dist_params:
                    func_signature = f"{func_name_with_nested}({','.join([str(p) for p in dist_params])})[{args_text_normalized}]"
                else:
                    func_signature = f"{func_name_with_nested}()[{args_text_normalized}]"
                self.func_signature_to_stable_key[func_signature] = stable_key
                # 创建基于参数值的签名映射
                param_strs = []
                for param in dist_params:
                    if isinstance(param, (int, float)):
                        param_strs.append(f"{param:.10f}".rstrip('0').rstrip('.'))
                    else:
                        param_strs.append(str(param))
                if param_strs:
                    param_based_signature = f"{func_name_with_nested}({','.join(param_strs)})"
                else:
                    param_based_signature = f"{func_name_with_nested}()"
                self.param_based_signatures[param_based_signature] = stable_key
                # 记录嵌套关系
                if parent_key:
                    if parent_key not in self.nested_relationships:
                        self.nested_relationships[parent_key] = []
                    if stable_key not in self.nested_relationships[parent_key]:
                        self.nested_relationships[parent_key].append(stable_key)
                        self.log_debug(f"记录嵌套关系: {parent_key} -> {stable_key}")
                    if parent_key not in self.nested_children_cache:
                        self.nested_children_cache[parent_key] = []
                    self.nested_children_cache[parent_key].append(stable_key)
                    self.nested_parent_cache[stable_key] = parent_key
                self.log_debug(f"注册分布函数成功: {stable_key}, 参数签名: {param_based_signature}, 缓存大小={self.total_iterations}")
    def _create_param_based_signature(self, func_name: str, dist_params: List[float], is_nested: bool = False, depth: int = 0, args_text: str = None) -> str:
        """创建基于参数值的签名 - 新增：考虑参数文本"""
        if is_nested:
            func_name_with_nested = f"{func_name}_nested{depth}"
        else:
            func_name_with_nested = func_name
        if args_text:
            return f"{func_name_with_nested}[{args_text}]"
        param_strs = []
        for param in dist_params:
            if isinstance(param, (int, float)):
                param_strs.append(f"{param:.10f}".rstrip('0').rstrip('.'))
            else:
                param_strs.append(str(param))
        if param_strs:
            return f"{func_name_with_nested}({','.join(param_strs)})"
        else:
            return f"{func_name_with_nested}()"
    def get_or_generate_value(self, stable_key: str) -> float:
        """获取或生成随机数值（简化版）- 增加调试信息"""
        with self._lock:
            if not self.simulation_running:
                func_info = self.stable_key_to_func_info.get(stable_key)
                if not func_info:
                    self.log_debug(f"非模拟模式，稳定键未注册: {stable_key}")
                    return self._generate_random_value('normal', [0, 1], {})
                return self._generate_random_value(
                    func_info.get('func_name', 'normal'),
                    func_info.get('dist_params', [0, 1]),
                    func_info.get('markers', {})
                )
            if stable_key not in self.distribution_cache:
                self.log_debug(f"警告: 稳定键未注册，动态创建缓存: {stable_key}")
                self.distribution_cache[stable_key] = np.full(self.total_iterations, np.nan, dtype=float)
            current_iter = self.current_iteration
            cache_array = self.distribution_cache[stable_key]
            if current_iter >= len(cache_array):
                new_size = max(len(cache_array) * 2, current_iter + 1)
                new_array = np.full(new_size, np.nan, dtype=float)
                new_array[:len(cache_array)] = cache_array
                self.distribution_cache[stable_key] = new_array
                cache_array = new_array
            if cache_array[current_iter] is not np.nan and not np.isnan(cache_array[current_iter]):
                value = cache_array[current_iter]
                if value == ERROR_MARKER:
                    return ERROR_MARKER
                try:
                    return float(value)
                except:
                    return value
            func_info = self.stable_key_to_func_info.get(stable_key)
            if not func_info:
                self.log_debug(f"错误: 稳定键未注册 {stable_key}")
                return ERROR_MARKER
            parent_key = self.nested_parent_cache.get(stable_key)
            if parent_key:
                parent_value = self.get_or_generate_value(parent_key)
                if parent_value == ERROR_MARKER:
                    self.log_debug(f"父函数 {parent_key} 生成失败，跳过子函数 {stable_key}")
            random_value = self._generate_random_value(
                func_info.get('func_name', 'normal'),
                func_info.get('dist_params', [0, 1]),
                func_info.get('markers', {})
            )
            try:
                cache_array[current_iter] = random_value
                self.log_debug(f"生成随机数: {stable_key}[{current_iter}] = {random_value:.6f}, 嵌套={func_info.get('is_nested', False)}, 深度={func_info.get('depth', 0)}")
                return random_value if random_value == ERROR_MARKER else float(random_value)
            except Exception as e:
                self.log_debug(f"记录随机数失败 {stable_key}[{current_iter}]: {str(e)}")
                return random_value if random_value == ERROR_MARKER else float(random_value)
    def _generate_random_value(self, func_name: str, params: List[float], markers: Dict[str, Any]) -> float:
        """生成随机数值"""
        try:
            dist_type = get_distribution_type(func_name)
            static_value = markers.get('static')
            if static_value is not None and get_static_mode():
                return float(static_value)
            if markers.get('loc') or markers.get('lock'):
                generator = self._get_or_create_generator(func_name, dist_type, params, markers)
                return float(generator._calculate_loc_value())
            generator = self._get_or_create_generator(func_name, dist_type, params, markers)
            seed = self._generate_seed(func_name, params, markers)
            return float(generator.generate_sample(seed))
        except Exception as e:
            self.log_debug(f"生成随机数值失败 {func_name}: {str(e)}")
            return ERROR_MARKER
    def _get_or_create_generator(self, func_name: str, dist_type: str, params: List[float], markers: Dict[str, Any]) -> 'DistributionGenerator':
        """获取或创建分布生成器"""
        generator_key = f"{func_name}_{dist_type}_{hash(tuple(params))}_{hash(tuple(markers.items()))}"
        if generator_key not in self.distribution_generators:
            self.distribution_generators[generator_key] = DistributionGenerator(
                func_name, dist_type, params, markers
            )
        return self.distribution_generators[generator_key]
    def _generate_seed(self, func_name: str, params: List[float], markers: Dict[str, Any]) -> int:
        """生成唯一种子"""
        seed_base = markers.get('seed', 42)
        current_iter = self.get_current_iteration()
        param_hash = hash(tuple(params)) % 10000
        seed = seed_base + current_iter * 1000 + param_hash
        return int(seed)
    def get_distribution_cache(self, stable_key: str) -> Optional[np.ndarray]:
        """获取分布函数缓存"""
        with self._lock:
            return self.distribution_cache.get(stable_key)
    def clear_simulation(self):
        """清除模拟状态"""
        with self._lock:
            self.current_iteration = 0
            self.total_iterations = 0
            self.simulation_running = False
            self.distribution_cache.clear()
            self.distribution_generators.clear()
            self.stable_key_to_func_info.clear()
            self.func_signature_to_stable_key.clear()
            self.param_based_signatures.clear()
            self.initialized_keys.clear()
            self.cell_value_cache.clear()
            self.cell_value_locks.clear()
            self.nested_relationships.clear()
            self.nested_children_cache.clear()
            self.nested_parent_cache.clear()
            self.registered_stable_keys.clear()
            self.cell_unique_counter.clear()
            self.debug_log.clear()
            self.log_debug("模拟状态管理器已清除")
    def find_stable_key_by_params(self, func_name: str, dist_params: List[float], is_nested: bool = False, depth: int = 0, parent_key: str = None, args_text: str = None) -> Optional[str]:
        """通过函数名和参数查找稳定键 - 支持嵌套函数和参数文本"""
        param_based_signature = self._create_param_based_signature(func_name, dist_params, is_nested, depth, args_text)
        if param_based_signature in self.param_based_signatures:
            stable_key = self.param_based_signatures[param_based_signature]
            self.log_debug(f"通过参数签名找到稳定键: {param_based_signature} -> {stable_key}")
            return stable_key
        for signature, stable_key in self.param_based_signatures.items():
            if is_nested:
                nested_func_name = f"{func_name}_nested{depth}"
                if signature.startswith(f"{nested_func_name}(") or (args_text and f"[{args_text}]" in signature):
                    if '(' in signature:
                        param_str = signature[len(nested_func_name)+1:-1].split('[')[0] if '[' in signature else signature[len(nested_func_name)+1:-1]
                        try:
                            params_from_signature = [float(p.strip()) for p in param_str.split(',') if p.strip()]
                            if len(params_from_signature) == len(dist_params):
                                match = True
                                for i in range(len(dist_params)):
                                    if abs(params_from_signature[i] - dist_params[i]) > 0.0001:
                                        match = False
                                        break
                                if match:
                                    self.log_debug(f"通过模糊匹配找到嵌套函数稳定键: {signature} -> {stable_key}")
                                    return stable_key
                        except:
                            continue
            else:
                if signature.startswith(f"{func_name}(") or (args_text and f"[{args_text}]" in signature):
                    if '(' in signature:
                        param_str = signature[len(func_name)+1:-1].split('[')[0] if '[' in signature else signature[len(func_name)+1:-1]
                        try:
                            params_from_signature = [float(p.strip()) for p in param_str.split(',') if p.strip()]
                            if len(params_from_signature) == len(dist_params):
                                match = True
                                for i in range(len(dist_params)):
                                    if abs(params_from_signature[i] - dist_params[i]) > 0.0001:
                                        match = False
                                        break
                                if match:
                                    self.log_debug(f"通过模糊匹配找到稳定键: {signature} -> {stable_key}")
                                    return stable_key
                        except:
                            continue
        self.log_debug(f"未找到稳定键: 函数名={func_name}, 嵌套={is_nested}, 深度={depth}, 参数={dist_params}, 参数文本={args_text}")
        return None
    def record_cell_value(self, cell_addr: str, value: float):
        """记录单元格的值"""
        with self._lock:
            if not self.simulation_running:
                return False
            current_iter = self.current_iteration
            if 0 <= current_iter < self.total_iterations:
                if cell_addr not in self.cell_value_cache:
                    self.cell_value_cache[cell_addr] = np.full(self.total_iterations, np.nan, dtype=float)
                    self.cell_value_locks[cell_addr] = threading.RLock()
                with self.cell_value_locks[cell_addr]:
                    self.cell_value_cache[cell_addr][current_iter] = value
                return True
            return False
    def get_cell_value(self, cell_addr: str) -> Optional[float]:
        """获取单元格当前迭代的值"""
        with self._lock:
            if cell_addr in self.cell_value_cache:
                cache_array = self.cell_value_cache[cell_addr]
                current_iter = self.current_iteration
                if 0 <= current_iter < len(cache_array):
                    value = cache_array[current_iter]
                    if not np.isnan(value):
                        return float(value)
            return None
    def add_nested_relationship(self, parent_key: str, child_key: str):
        """添加嵌套函数关系"""
        with self._lock:
            if parent_key not in self.nested_relationships:
                self.nested_relationships[parent_key] = []
            if child_key not in self.nested_relationships[parent_key]:
                self.nested_relationships[parent_key].append(child_key)
                self.nested_parent_cache[child_key] = parent_key
                self.log_debug(f"记录嵌套关系: {parent_key} -> {child_key}")
    def get_nested_children(self, parent_key: str) -> List[str]:
        """获取嵌套子函数"""
        with self._lock:
            return self.nested_children_cache.get(parent_key, []).copy()
    def get_nested_parent(self, child_key: str) -> Optional[str]:
        """获取嵌套父函数"""
        with self._lock:
            return self.nested_parent_cache.get(child_key)
    def is_registered(self, stable_key: str) -> bool:
        """检查稳定键是否已注册"""
        with self._lock:
            return stable_key in self.registered_stable_keys
    def get_unique_seed_for_cell(self, cell_addr: str) -> int:
        """为单元格生成唯一种子"""
        with self._lock:
            if cell_addr not in self.cell_unique_counter:
                self.cell_unique_counter[cell_addr] = 0
            self.cell_unique_counter[cell_addr] += 1
            addr_hash = hash(cell_addr) % 1000000
            counter = self.cell_unique_counter[cell_addr]
            time_part = int(time.time() * 1000) % 1000000
            return int(addr_hash + counter * 1000 + time_part)
# 全局模拟状态管理器
_simulation_state = SimulationStateManager()
# ==================== 统一的分布生成器类 ====================
class DistributionGenerator:
    """统一的分布生成器，支持shift和truncate，包含支持范围检查和错误传播"""
    # 映射分布类型到对应的分布类（用于计算截断均值）
    DIST_CLASS_MAP = {
        'normal': NormalDistribution,
        'uniform': UniformDistribution,
        'erf': ErfDistribution,
        'gamma': GammaDistribution,
        'erlang': ErlangDistribution,
        'beta': BetaDistribution,
        'poisson': PoissonDistribution,
        'chisq': ChiSquaredDistribution,
        'f': FDistribution,
        'student': TDistribution,
        'expon': ExponentialDistribution,
        'bernoulli': BernoulliDistribution,
        'triang': TriangDistribution,
        'binomial': BinomialDistribution,
        'cauchy': CauchyDistribution,
        'dagum': DagumDistribution,
        'doubletriang': DoubleTriangDistribution,
        'extvalue': ExtvalueDistribution,
        'extvaluemin': ExtvalueMinDistribution,
        'fatiguelife': FatigueLifeDistribution,
        'frechet': FrechetDistribution,
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
        'general': GeneralDistribution,
        'histogrm': HistogrmDistribution,
        'negbin': NegbinDistribution,
        'invgauss': InvgaussDistribution,
        'duniform': DUniformDistribution,
        'geomet': GeometDistribution,
        'hypergeo': HypergeoDistribution,
        'intuniform': IntuniformDistribution,
        'trigen': TrigenDistribution,
        'cumul': CumulDistribution,
        'discrete': DiscreteDistribution
    }
    def __init__(self, func_name: str, dist_type: str, dist_params: List[float], markers: Dict[str, Any]):
        self.func_name = func_name
        self.dist_type = dist_type
        self.dist_params = dist_params
        self.markers = markers
        self.shift_amount, self.truncate_type, self.truncate_lower, self.truncate_upper = self._extract_shift_and_truncate_params()
        self.support_low, self.support_high = get_distribution_support(func_name, dist_params)
        self.truncate_invalid = False
        self.adjusted_truncate_lower = self.truncate_lower
        self.adjusted_truncate_upper = self.truncate_upper
        self._check_and_adjust_truncate_boundaries()
        # 解析数组参数（用于 Cumul 和 Discrete）
        self.x_vals_list = None
        self.p_vals_list = None
        self.x_vals_str = markers.get('x_vals')
        self.p_vals_str = markers.get('p_vals')
        self._array_parse_failed = False
        if self.x_vals_str and self.dist_type == 'duniform' and not self.p_vals_str:
            try:
                def parse_numbers(s):
                    s = s.strip()
                    if s.startswith('{') and s.endswith('}'):
                        s = s[1:-1]
                    parts = [p.strip() for p in s.split(',') if p.strip()]
                    return [float(p) for p in parts]
                self.x_vals_list = parse_numbers(self.x_vals_str)
                if self.x_vals_list:
                    p = 1.0 / len(self.x_vals_list)
                    self.p_vals_list = [p] * len(self.x_vals_list)
                    self.p_vals_str = ','.join(str(v) for v in self.p_vals_list)
            except Exception as e:
                print(f"DUniform: {e}")
                self.x_vals_list = self.p_vals_list = None
                self._array_parse_failed = True
        if self.dist_type == 'histogrm' and self.p_vals_str and not self.x_vals_str:
            try:
                def parse_numbers(s):
                    s = s.strip()
                    if s.startswith('{') and s.endswith('}'):
                        s = s[1:-1]
                    parts = [p.strip() for p in s.split(',') if p.strip()]
                    return [float(p) for p in parts]
                self.p_vals_list = parse_numbers(self.p_vals_str)
                histogrm_parse_p_table(self.p_vals_list)
            except Exception as e:
                print(f"Histogrm: {e}")
                self.x_vals_list = self.p_vals_list = None
                self._array_parse_failed = True
        if self.x_vals_str and self.p_vals_str:
            try:
                # 改进解析：处理花括号和普通逗号分隔字符串
                def parse_numbers(s):
                    s = s.strip()
                    if s.startswith('{') and s.endswith('}'):
                        s = s[1:-1]
                    parts = [p.strip() for p in s.split(',') if p.strip()]
                    return [float(p) for p in parts]
                self.x_vals_list = parse_numbers(self.x_vals_str)
                self.p_vals_list = parse_numbers(self.p_vals_str)
                # 对于 Discrete，归一化概率
                if self.dist_type in ('discrete', 'duniform'):
                    total = sum(self.p_vals_list)
                    if total > 0:
                        self.p_vals_list = [p / total for p in self.p_vals_list]
                # 对于 Cumul，验证数组有效性
                if self.dist_type == 'cumul':
                    from dist_cumul import _parse_arrays as cumul_parse_arrays
                    cumul_parse_arrays(self.x_vals_list, self.p_vals_list)
                elif self.dist_type == 'general':
                    general_parse_arrays(float(self.dist_params[0]), float(self.dist_params[1]), self.x_vals_list, self.p_vals_list)
                elif self.dist_type == 'histogrm':
                    histogrm_parse_p_table(self.p_vals_list)
            except Exception as e:
                print(f"解析数组参数失败: {e}")
                self.x_vals_list = self.p_vals_list = None
                self._array_parse_failed = True
        # ===== 新增：为离散/累积分布设置正确的支持范围 =====
        # 必须在解析数组参数后、再次检查截断前进行
        if self.dist_type in ('cumul', 'discrete', 'duniform') and self.x_vals_list is not None and len(self.x_vals_list) > 0:
            self.support_low = min(self.x_vals_list)
            self.support_high = max(self.x_vals_list)
            # 由于支持范围已改变，需要重新检查截断有效性
            self.truncate_invalid = False
            self.adjusted_truncate_lower = self.truncate_lower
            self.adjusted_truncate_upper = self.truncate_upper
        # ===================================================
                                # ===== 为 Trigen 设置正确的支持范围 =====
        if self.dist_type == 'trigen' and len(self.dist_params) >= 5:
            try:
                from dist_trigen import _convert_trigen_to_triang
                a, c, b = _convert_trigen_to_triang(
                    self.dist_params[0], self.dist_params[1], 
                    self.dist_params[2], self.dist_params[3], self.dist_params[4]
                )
                self.support_low = a
                self.support_high = b
            except Exception as e:
                print(f"Trigen 转换失败，无法设置支持范围: {e}")
        # ---------- 新增：修复三角分布参数顺序 ----------
        if self.dist_type == 'triang' and len(self.dist_params) >= 3:
            a, c, b = self.dist_params[0], self.dist_params[1], self.dist_params[2]
            # 如果 c 大于 b，很可能顺序错了，交换 c 和 b
            if c > b:
                print(f"警告: 三角分布参数顺序可能错误，将 c={c} 与 b={b} 交换")
                self.dist_params[1], self.dist_params[2] = b, c
        # ---------- 结束修复 ----------
        self._check_and_adjust_truncate_boundaries()
        self.original_generator = self._create_original_generator()
        self.generator_with_transform = self._create_truncated_generator()
    def _get_support(self) -> Tuple[float, float]:
        return self.support_low, self.support_high
    def _check_and_adjust_truncate_boundaries(self):
        if self.truncate_lower is not None and self.truncate_upper is not None:
            if self.truncate_lower > self.truncate_upper:
                self.truncate_invalid = True
                print(f"警告: 截断参数无效，下限 {self.truncate_lower} 大于上限 {self.truncate_upper}")
                return
        if self.truncate_type not in ['value', 'value2']:
            return
        if self.truncate_type == 'value2':
            orig_lower = self.truncate_lower - self.shift_amount if self.truncate_lower is not None else None
            orig_upper = self.truncate_upper - self.shift_amount if self.truncate_upper is not None else None
        else:
            orig_lower = self.truncate_lower
            orig_upper = self.truncate_upper
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
        # ===== 新增：单边截断超出支持范围检查 =====
        # 如果下界大于支持上界（有限），无有效值
        if new_lower is not None and high_support != float('inf') and new_lower > high_support:
            self.truncate_invalid = True
            print(f"警告: 截断下界 {new_lower} 大于分布支持上界 {high_support}，无有效值")
            return
        # 如果上界小于支持下界（有限），无有效值
        if new_upper is not None and low_support != float('-inf') and new_upper < low_support:
            self.truncate_invalid = True
            print(f"警告: 截断上界 {new_upper} 小于分布支持下界 {low_support}，无有效值")
            return
        if new_lower is not None and new_upper is not None:
            if new_lower > new_upper:
                self.truncate_invalid = True
                print(f"警告: 截断范围与支持范围无交集，下限 {new_lower} 大于上限 {new_upper}，无有效值")
                return
        if self.truncate_type == 'value2':
            self.adjusted_truncate_lower = new_lower + self.shift_amount if new_lower is not None else None
            self.adjusted_truncate_upper = new_upper + self.shift_amount if new_upper is not None else None
        else:
            self.adjusted_truncate_lower = new_lower
            self.adjusted_truncate_upper = new_upper
    def _extract_shift_and_truncate_params(self) -> Tuple[float, str, Optional[float], Optional[float]]:
        markers = self.markers
        shift_amount = 0.0
        truncate_type = None
        truncate_lower = None
        truncate_upper = None
        shift_val = markers.get('shift')
        if shift_val is not None:
            try:
                shift_amount = float(shift_val)
            except:
                shift_amount = 0.0
        truncate_val = markers.get('truncate')
        truncate_p_val = markers.get('truncate_p') or markers.get('truncatep')
        truncate2_val = markers.get('truncate2')
        truncate_p2_val = markers.get('truncate_p2') or markers.get('truncatep2')
        if truncate_val is not None:
            truncate_type = 'value'
            if isinstance(truncate_val, str):
                truncate_val = truncate_val.strip()
                if truncate_val.startswith('(') and truncate_val.endswith(')'):
                    truncate_val = truncate_val[1:-1]
                parts = truncate_val.split(',')
                truncate_lower = None
                truncate_upper = None
                if len(parts) >= 1:
                    lower_str = parts[0].strip()
                    if lower_str:
                        try:
                            truncate_lower = float(lower_str)
                        except:
                            pass
                if len(parts) >= 2:
                    upper_str = parts[1].strip()
                    if upper_str:
                        try:
                            truncate_upper = float(upper_str)
                        except:
                            pass
                if truncate_lower is None and truncate_upper is None:
                    try:
                        cleaned = re.sub(r'[^\d\.,\-]', '', truncate_val)
                        parts = cleaned.split(',')
                        if len(parts) >= 1:
                            lower_str = parts[0].strip()
                            if lower_str:
                                truncate_lower = float(lower_str)
                        if len(parts) >= 2:
                            upper_str = parts[1].strip()
                            if upper_str:
                                truncate_upper = float(upper_str)
                    except:
                        pass
        elif truncate2_val is not None:
            truncate_type = 'value2'
            if isinstance(truncate2_val, str):
                truncate2_val = truncate2_val.strip()
                if truncate2_val.startswith('(') and truncate2_val.endswith(')'):
                    truncate2_val = truncate2_val[1:-1]
                parts = truncate2_val.split(',')
                truncate_lower = None
                truncate_upper = None
                if len(parts) >= 1:
                    lower_str = parts[0].strip()
                    if lower_str:
                        try:
                            truncate_lower = float(lower_str)
                        except:
                            pass
                if len(parts) >= 2:
                    upper_str = parts[1].strip()
                    if upper_str:
                        try:
                            truncate_upper = float(upper_str)
                        except:
                            pass
                if truncate_lower is None and truncate_upper is None:
                    try:
                        cleaned = re.sub(r'[^\d\.,\-]', '', truncate2_val)
                        parts = cleaned.split(',')
                        if len(parts) >= 1:
                            lower_str = parts[0].strip()
                            if lower_str:
                                truncate_lower = float(lower_str)
                        if len(parts) >= 2:
                            upper_str = parts[1].strip()
                            if upper_str:
                                truncate_upper = float(upper_str)
                    except:
                        pass
        elif truncate_p_val is not None:
            truncate_type = 'percentile'
            if isinstance(truncate_p_val, str):
                truncate_p_val = truncate_p_val.strip()
                if truncate_p_val.startswith('(') and truncate_p_val.endswith(')'):
                    truncate_p_val = truncate_p_val[1:-1]
                parts = truncate_p_val.split(',')
                truncate_lower = None
                truncate_upper = None
                if len(parts) >= 1:
                    lower_str = parts[0].strip()
                    if lower_str:
                        try:
                            lower_pct = float(lower_str)
                            # 检查有效性：必须在 [0,1] 或 [0,100]
                            if 0 <= lower_pct <= 1:
                                truncate_lower = lower_pct
                            elif 0 <= lower_pct <= 100:
                                truncate_lower = lower_pct / 100.0
                            else:
                                # 无效值，标记截断无效
                                self.truncate_invalid = True
                                print(f"警告: 百分位数截断下界 {lower_pct} 无效，必须介于0和100之间")
                        except:
                            pass
                if len(parts) >= 2:
                    upper_str = parts[1].strip()
                    if upper_str:
                        try:
                            upper_pct = float(upper_str)
                            if 0 <= upper_pct <= 1:
                                truncate_upper = upper_pct
                            elif 0 <= upper_pct <= 100:
                                truncate_upper = upper_pct / 100.0
                            else:
                                self.truncate_invalid = True
                                print(f"警告: 百分位数截断上界 {upper_pct} 无效，必须介于0和100之间")
                        except:
                            pass
                if truncate_lower is None and truncate_upper is None:
                    try:
                        cleaned = re.sub(r'[^\d\.,\-]', '', truncate_p_val)
                        parts = cleaned.split(',')
                        if len(parts) >= 1:
                            lower_str = parts[0].strip()
                            if lower_str:
                                lower_pct = float(lower_str)
                                if 0 <= lower_pct <= 1:
                                    truncate_lower = lower_pct
                                elif 0 <= lower_pct <= 100:
                                    truncate_lower = lower_pct / 100.0
                                else:
                                    self.truncate_invalid = True
                        if len(parts) >= 2:
                            upper_str = parts[1].strip()
                            if upper_str:
                                upper_pct = float(upper_str)
                                if 0 <= upper_pct <= 1:
                                    truncate_upper = upper_pct
                                elif 0 <= upper_pct <= 100:
                                    truncate_upper = upper_pct / 100.0
                                else:
                                    self.truncate_invalid = True
                    except:
                        pass
        elif truncate_p2_val is not None:
            truncate_type = 'percentile2'
            if isinstance(truncate_p2_val, str):
                truncate_p2_val = truncate_p2_val.strip()
                if truncate_p2_val.startswith('(') and truncate_p2_val.endswith(')'):
                    truncate_p2_val = truncate_p2_val[1:-1]
                parts = truncate_p2_val.split(',')
                truncate_lower = None
                truncate_upper = None
                if len(parts) >= 1:
                    lower_str = parts[0].strip()
                    if lower_str:
                        try:
                            lower_pct = float(lower_str)
                            if 0 <= lower_pct <= 1:
                                truncate_lower = lower_pct
                            elif 0 <= lower_pct <= 100:
                                truncate_lower = lower_pct / 100.0
                            else:
                                self.truncate_invalid = True
                        except:
                            pass
                if len(parts) >= 2:
                    upper_str = parts[1].strip()
                    if upper_str:
                        try:
                            upper_pct = float(upper_str)
                            if 0 <= upper_pct <= 1:
                                truncate_upper = upper_pct
                            elif 0 <= upper_pct <= 100:
                                truncate_upper = upper_pct / 100.0
                            else:
                                self.truncate_invalid = True
                        except:
                            pass
                if truncate_lower is None and truncate_upper is None:
                    try:
                        cleaned = re.sub(r'[^\d\.,\-]', '', truncate_p2_val)
                        parts = cleaned.split(',')
                        if len(parts) >= 1:
                            lower_str = parts[0].strip()
                            if lower_str:
                                lower_pct = float(lower_str)
                                if 0 <= lower_pct <= 1:
                                    truncate_lower = lower_pct
                                elif 0 <= lower_pct <= 100:
                                    truncate_lower = lower_pct / 100.0
                                else:
                                    self.truncate_invalid = True
                        if len(parts) >= 2:
                            upper_str = parts[1].strip()
                            if upper_str:
                                upper_pct = float(upper_str)
                                if 0 <= upper_pct <= 1:
                                    truncate_upper = upper_pct
                                elif 0 <= upper_pct <= 100:
                                    truncate_upper = upper_pct / 100.0
                                else:
                                    self.truncate_invalid = True
                    except:
                        pass
        return shift_amount, truncate_type, truncate_lower, truncate_upper
    def _create_original_generator(self) -> Callable:
        """创建原始分布生成器，使用解析后的列表（如果存在）"""
        dist_type = self.dist_type
        dist_params = self.dist_params
        if dist_type == 'normal':
            def normal_generator(rng, params):
                return float(rng.normal(loc=params[0], scale=params[1]))
            return normal_generator
        elif dist_type == 'uniform':
            def uniform_generator(rng, params):
                return float(rng.uniform(low=params[0], high=params[1]))
            return uniform_generator
        elif dist_type == 'erf':
            return lambda rng, params: erf_generator(rng, params)
        elif dist_type == 'erlang':
            return lambda rng, params: erlang_generator(rng, params)
        elif dist_type == 'gamma':
            def gamma_generator(rng, params):
                return float(rng.gamma(shape=params[0], scale=params[1]))
            return gamma_generator
        elif dist_type == 'poisson':
            def poisson_generator(rng, params):
                return float(rng.poisson(lam=params[0]))
            return poisson_generator
        elif dist_type == 'beta':
            def beta_generator(rng, params):
                return float(rng.beta(a=params[0], b=params[1]))
            return beta_generator
        elif dist_type == 'chisq':
            def chisquare_generator(rng, params):
                return float(rng.chisquare(df=params[0]))
            return chisquare_generator
        elif dist_type == 'f':
            def f_generator(rng, params):
                return float(rng.f(dfnum=params[0], dfden=params[1]))
            return f_generator
        elif dist_type == 'student':
            def t_generator(rng, params):
                return float(rng.standard_t(df=params[0]))
            return t_generator
        elif dist_type == 'expon':
            def exponential_generator(rng, params):
                return float(rng.exponential(scale=params[0]))
            return exponential_generator
        elif dist_type == 'bernoulli':
            return lambda rng, params: bernoulli_generator(rng, params)
        elif dist_type == 'triang':
            return lambda rng, params: triang_generator(rng, params)
        elif dist_type == 'binomial':
            return lambda rng, params: binomial_generator(rng, params)
        elif dist_type == 'cauchy':
            return lambda rng, params: cauchy_generator(rng, params)
        elif dist_type == 'dagum':
            return lambda rng, params: dagum_generator(rng, params)
        elif dist_type == 'doubletriang':
            return lambda rng, params: doubletriang_generator(rng, params)
        elif dist_type == 'extvalue':
            return lambda rng, params: extvalue_generator(rng, params)
        elif dist_type == 'extvaluemin':
            return lambda rng, params: extvaluemin_generator(rng, params)
        elif dist_type == 'fatiguelife':
            return lambda rng, params: fatiguelife_generator(rng, params)
        elif dist_type == 'frechet':
            return lambda rng, params: frechet_generator(rng, params)
        elif dist_type == 'hypsecant':
            return lambda rng, params: hypsecant_generator(rng, params)
        elif dist_type == 'johnsonsb':
            return lambda rng, params: johnsonsb_generator(rng, params)
        elif dist_type == 'johnsonsu':
            return lambda rng, params: johnsonsu_generator(rng, params)
        elif dist_type == 'kumaraswamy':
            return lambda rng, params: kumaraswamy_generator(rng, params)
        elif dist_type == 'laplace':
            return lambda rng, params: laplace_generator(rng, params)
        elif dist_type == 'logistic':
            return lambda rng, params: logistic_generator(rng, params)
        elif dist_type == 'loglogistic':
            return lambda rng, params: loglogistic_generator(rng, params)
        elif dist_type == 'lognorm':
            return lambda rng, params: lognorm_generator(rng, params)
        elif dist_type == 'lognorm2':
            return lambda rng, params: lognorm2_generator(rng, params)
        elif dist_type == 'betageneral':
            return lambda rng, params: betageneral_generator(rng, params)
        elif dist_type == 'betasubj':
            return lambda rng, params: betasubj_generator(rng, params)
        elif dist_type == 'burr12':
            return lambda rng, params: burr12_generator(rng, params)
        elif dist_type == 'compound':
            return lambda rng, params: compound_generator(rng, params)
        elif dist_type == 'splice':
            return lambda rng, params: splice_generator(rng, params)
        elif dist_type == 'pert':
            return lambda rng, params: pert_generator(rng, params)
        elif dist_type == 'reciprocal':
            return lambda rng, params: reciprocal_generator(rng, params)
        elif dist_type == 'rayleigh':
            return lambda rng, params: rayleigh_generator(rng, params)
        elif dist_type == 'weibull':
            return lambda rng, params: weibull_generator(rng, params)
        elif dist_type == 'pearson5':
            return lambda rng, params: pearson5_generator(rng, params)
        elif dist_type == 'pearson6':
            return lambda rng, params: pearson6_generator(rng, params)
        elif dist_type == 'pareto2':
            return lambda rng, params: pareto2_generator(rng, params)
        elif dist_type == 'pareto':
            return lambda rng, params: pareto_generator(rng, params)
        elif dist_type == 'levy':
            return lambda rng, params: levy_generator(rng, params)
        elif dist_type == 'general':
            if self.x_vals_list is not None and self.p_vals_list is not None:
                return lambda rng, params: general_generator(rng, params, x_vals=self.x_vals_list, p_vals=self.p_vals_list)
            else:
                return lambda rng, params: ERROR_MARKER
        elif dist_type == 'histogrm':
            if self.p_vals_list is not None:
                return lambda rng, params: histogrm_generator(rng, params, p_vals=self.p_vals_list)
            else:
                return lambda rng, params: ERROR_MARKER
        elif dist_type == 'negbin':
            return lambda rng, params: negbin_generator(rng, params)
        elif dist_type == 'invgauss':
            return lambda rng, params: invgauss_generator(rng, params)
        elif dist_type == 'geomet':
            return lambda rng, params: geomet_generator(rng, params)
        elif dist_type == 'hypergeo':
            return lambda rng, params: hypergeo_generator(rng, params)
        elif dist_type == 'intuniform':
            return lambda rng, params: intuniform_generator(rng, params)
        elif dist_type == 'duniform':
            if self.x_vals_list is not None:
                return lambda rng, params: duniform_generator(rng, params, x_vals=self.x_vals_list)
            else:
                return lambda rng, params: ERROR_MARKER
        elif dist_type == 'trigen':
            return lambda rng, params: trigen_generator(rng, params)
        elif dist_type == 'cumul':
            if self.x_vals_list is not None and self.p_vals_list is not None:
                # 使用解析后的列表调用 cumul_generator（该函数需要列表参数）
                return lambda rng, params: cumul_generator(rng, params, self.x_vals_list, self.p_vals_list)
            else:
                return lambda rng, params: ERROR_MARKER
        elif dist_type == 'discrete':
            if self.x_vals_list is not None and self.p_vals_list is not None:
                # 使用解析后的列表调用 discrete_generator
                return lambda rng, params: discrete_generator(rng, params, self.x_vals_list, self.p_vals_list)
            else:
                return lambda rng, params: ERROR_MARKER
        else:
            def default_generator(rng, params):
                return float(rng.normal(loc=params[0], scale=params[1]))
            return default_generator
    def _create_truncated_generator(self) -> Callable:
        dist_type = self.dist_type
        dist_params = self.dist_params
        markers = self.markers
        if not self.truncate_type:
            def generator_with_shift(rng, params):
                value = self.original_generator(rng, params)
                return value + self.shift_amount
            return generator_with_shift
        is_percentile = self.truncate_type in ['percentile', 'percentile2']
        ppf_func = None
        if is_percentile:
            ppf_func = self._create_ppf_function()
        if self.truncate_type in ['value', 'value2']:
            if self.truncate_type == 'value':
                def generator_value_truncate(rng, params):
                    max_attempts = 1000
                    for _ in range(max_attempts):
                        val = self.original_generator(rng, params)
                        if self.adjusted_truncate_lower is not None and val < self.adjusted_truncate_lower:
                            continue
                        if self.adjusted_truncate_upper is not None and val > self.adjusted_truncate_upper:
                            continue
                        return val + self.shift_amount
                    if self.adjusted_truncate_lower is not None and self.adjusted_truncate_upper is not None:
                        return float(self.adjusted_truncate_lower + rng.random() * (self.adjusted_truncate_upper - self.adjusted_truncate_lower)) + self.shift_amount
                    elif self.adjusted_truncate_lower is not None:
                        ppf = self._create_ppf_function()
                        if ppf is not None:
                            lower_cdf = self._estimate_cdf(self.adjusted_truncate_lower)
                            u = rng.random()
                            mapped_p = lower_cdf + u * (1.0 - lower_cdf)
                            val = ppf(mapped_p)
                            return val + self.shift_amount
                        else:
                            return self.adjusted_truncate_lower + self.shift_amount
                    elif self.adjusted_truncate_upper is not None:
                        ppf = self._create_ppf_function()
                        if ppf is not None:
                            upper_cdf = self._estimate_cdf(self.adjusted_truncate_upper)
                            u = rng.random()
                            mapped_p = u * upper_cdf
                            val = ppf(mapped_p)
                            return val + self.shift_amount
                        else:
                            return self.adjusted_truncate_upper + self.shift_amount
                    else:
                        return val + self.shift_amount
                return generator_value_truncate
            else:  # 'value2'
                def generator_value_truncate2(rng, params):
                    max_attempts = 1000
                    for _ in range(max_attempts):
                        orig = self.original_generator(rng, params)
                        val = orig + self.shift_amount
                        if self.adjusted_truncate_lower is not None and val < self.adjusted_truncate_lower:
                            continue
                        if self.adjusted_truncate_upper is not None and val > self.adjusted_truncate_upper:
                            continue
                        return val
                    if self.adjusted_truncate_lower is not None and self.adjusted_truncate_upper is not None:
                        return float(self.adjusted_truncate_lower + rng.random() * (self.adjusted_truncate_upper - self.adjusted_truncate_lower))
                    elif self.adjusted_truncate_lower is not None:
                        orig_lower = self.adjusted_truncate_lower - self.shift_amount
                        ppf = self._create_ppf_function()
                        if ppf is not None:
                            lower_cdf = self._estimate_cdf(orig_lower)
                            u = rng.random()
                            mapped_p = lower_cdf + u * (1.0 - lower_cdf)
                            orig_val = ppf(mapped_p)
                            return orig_val + self.shift_amount
                        else:
                            return self.adjusted_truncate_lower
                    elif self.adjusted_truncate_upper is not None:
                        orig_upper = self.adjusted_truncate_upper - self.shift_amount
                        ppf = self._create_ppf_function()
                        if ppf is not None:
                            upper_cdf = self._estimate_cdf(orig_upper)
                            u = rng.random()
                            mapped_p = u * upper_cdf
                            orig_val = ppf(mapped_p)
                            return orig_val + self.shift_amount
                        else:
                            return self.adjusted_truncate_upper
                    else:
                        return val
                return generator_value_truncate2
        else:  # 百分位数截断
            if self.truncate_type == 'percentile':
                def generator_percentile_truncate(rng, params):
                    u = rng.random()
                    if self.truncate_lower is None:
                        lower_p = 0.0
                    else:
                        lower_p = self.truncate_lower
                    if self.truncate_upper is None:
                        upper_p = 1.0
                    else:
                        upper_p = self.truncate_upper
                    mapped_p = lower_p + u * (upper_p - lower_p)
                    if ppf_func:
                        value = ppf_func(mapped_p)
                    else:
                        max_attempts = 100
                        for _ in range(max_attempts):
                            value = self.original_generator(rng, params)
                            p_value = self._estimate_cdf(value)
                            if (self.truncate_lower is None or p_value >= self.truncate_lower) and \
                               (self.truncate_upper is None or p_value <= self.truncate_upper):
                                break
                        else:
                            if self.truncate_lower is not None and self.truncate_upper is not None:
                                value = ppf_func((self.truncate_lower + self.truncate_upper) / 2) if ppf_func else 0.0
                            elif self.truncate_lower is not None:
                                value = ppf_func(self.truncate_lower) if ppf_func else 0.0
                            else:
                                value = ppf_func(self.truncate_upper) if ppf_func else 0.0
                    return value + self.shift_amount
                return generator_percentile_truncate
            else:  # 'percentile2'
                def generator_percentile_truncate2(rng, params):
                    value = self.original_generator(rng, params)
                    value = value + self.shift_amount
                    original_value = value - self.shift_amount
                    p_value = self._estimate_cdf(original_value)
                    if (self.truncate_lower is not None and p_value < self.truncate_lower
                        ) or (
                        self.truncate_upper is not None and p_value > self.truncate_upper):
                        u = rng.random()
                        if self.truncate_lower is None:
                            lower_p = 0.0
                        else:
                            lower_p = self.truncate_lower
                        if self.truncate_upper is None:
                            upper_p = 1.0
                        else:
                            upper_p = self.truncate_upper
                        mapped_p = lower_p + u * (upper_p - lower_p)
                        if ppf_func:
                            try:
                                original_value = ppf_func(mapped_p)
                                value = original_value + self.shift_amount
                            except Exception:
                                pass
                        else:
                            if self.truncate_lower is not None and self.truncate_upper is not None:
                                mid_p = (self.truncate_lower + self.truncate_upper) / 2
                                if ppf_func:
                                    original_value = ppf_func(mid_p)
                                    value = original_value + self.shift_amount
                    return value
                return generator_percentile_truncate2
    def _create_ppf_function(self) -> Optional[Callable]:
        dist_type = self.dist_type
        dist_params = self.dist_params
        try:
            if dist_type == 'normal':
                from scipy.stats import norm
                def ppf(p):
                    return norm.ppf(p, dist_params[0], dist_params[1])
                return ppf
            elif dist_type == 'uniform':
                from scipy.stats import uniform
                def ppf(p):
                    return uniform.ppf(p, dist_params[0], dist_params[1] - dist_params[0])
                return ppf
            elif dist_type == 'gamma':
                from scipy.stats import gamma
                def ppf(p):
                    return gamma.ppf(p, dist_params[0], scale=dist_params[1])
                return ppf
            elif dist_type == 'erlang':
                m, beta = int(round(float(dist_params[0]))), dist_params[1]
                return lambda q: erlang_ppf(q, m, beta)
            elif dist_type == 'beta':
                from scipy.stats import beta
                def ppf(p):
                    return beta.ppf(p, dist_params[0], dist_params[1])
                return ppf
            elif dist_type == 'chisq':
                from scipy.stats import chi2
                def ppf(p):
                    return chi2.ppf(p, dist_params[0])
                return ppf
            elif dist_type == 'f':
                from scipy.stats import f
                def ppf(p):
                    return f.ppf(p, dist_params[0], dist_params[1])
                return ppf
            elif dist_type == 'student':
                from scipy.stats import t
                def ppf(p):
                    return t.ppf(p, dist_params[0])
                return ppf
            elif dist_type == 'expon':
                from scipy.stats import expon
                def ppf(p):
                    return expon.ppf(p, scale=dist_params[0])
                return ppf
            elif dist_type == 'bernoulli':
                p = dist_params[0]
                return lambda q: bernoulli_ppf(q, p)
            elif dist_type == 'triang':
                a, c, b = dist_params[0], dist_params[1], dist_params[2]
                return lambda q: triang_ppf(q, a, c, b)
            elif dist_type == 'binomial':
                n, p = int(dist_params[0]), dist_params[1]
                return lambda q: binomial_ppf(q, n, p)
            elif dist_type == 'cauchy':
                gamma, beta = dist_params[0], dist_params[1]
                return lambda q: cauchy_ppf(q, gamma, beta)
            elif dist_type == 'erf':
                h = dist_params[0]
                return lambda q: erf_ppf(q, h)
            elif dist_type == 'dagum':
                gamma, beta, alpha1, alpha2 = dist_params[0], dist_params[1], dist_params[2], dist_params[3]
                return lambda q: dagum_ppf(q, gamma, beta, alpha1, alpha2)
            elif dist_type == 'doubletriang':
                min_val, m_likely, max_val, lower_p = dist_params[0], dist_params[1], dist_params[2], dist_params[3]
                return lambda q: doubletriang_ppf(q, min_val, m_likely, max_val, lower_p)
            elif dist_type == 'extvalue':
                a, b = dist_params[0], dist_params[1]
                return lambda q: extvalue_ppf(q, a, b)
            elif dist_type == 'extvaluemin':
                a, b = dist_params[0], dist_params[1]
                return lambda q: extvaluemin_ppf(q, a, b)
            elif dist_type == 'fatiguelife':
                y, beta, alpha = dist_params[0], dist_params[1], dist_params[2]
                return lambda q: fatiguelife_ppf(q, y, beta, alpha)
            elif dist_type == 'frechet':
                y, beta, alpha = dist_params[0], dist_params[1], dist_params[2]
                return lambda q: frechet_ppf(q, y, beta, alpha)
            elif dist_type == 'hypsecant':
                gamma, beta = dist_params[0], dist_params[1]
                return lambda q: hypsecant_ppf(q, gamma, beta)
            elif dist_type == 'johnsonsb':
                alpha1, alpha2, a, b = dist_params[0], dist_params[1], dist_params[2], dist_params[3]
                return lambda q: johnsonsb_ppf(q, alpha1, alpha2, a, b)
            elif dist_type == 'johnsonsu':
                alpha1, alpha2, gamma, beta = dist_params[0], dist_params[1], dist_params[2], dist_params[3]
                return lambda q: johnsonsu_ppf(q, alpha1, alpha2, gamma, beta)
            elif dist_type == 'kumaraswamy':
                alpha1, alpha2, min_val, max_val = dist_params[0], dist_params[1], dist_params[2], dist_params[3]
                return lambda q: kumaraswamy_ppf(q, alpha1, alpha2, min_val, max_val)
            elif dist_type == 'laplace':
                mu, sigma = dist_params[0], dist_params[1]
                return lambda q: laplace_ppf(q, mu, sigma)
            elif dist_type == 'logistic':
                alpha, beta = dist_params[0], dist_params[1]
                return lambda q: logistic_ppf(q, alpha, beta)
            elif dist_type == 'loglogistic':
                gamma, beta, alpha = dist_params[0], dist_params[1], dist_params[2]
                return lambda q: loglogistic_ppf(q, gamma, beta, alpha)
            elif dist_type == 'lognorm':
                mu, sigma = dist_params[0], dist_params[1]
                return lambda q: lognorm_ppf(q, mu, sigma)
            elif dist_type == 'lognorm2':
                mu, sigma = dist_params[0], dist_params[1]
                return lambda q: lognorm2_ppf(q, mu, sigma)
            elif dist_type == 'betageneral':
                alpha1, alpha2, min_val, max_val = dist_params[0], dist_params[1], dist_params[2], dist_params[3]
                return lambda q: betageneral_ppf(q, alpha1, alpha2, min_val, max_val)
            elif dist_type == 'betasubj':
                min_val, m_likely, mean_val, max_val = dist_params[0], dist_params[1], dist_params[2], dist_params[3]
                return lambda q: betasubj_ppf(q, min_val, m_likely, mean_val, max_val)
            elif dist_type == 'burr12':
                gamma, beta, alpha1, alpha2 = dist_params[0], dist_params[1], dist_params[2], dist_params[3]
                return lambda q: burr12_ppf(q, gamma, beta, alpha1, alpha2)
            elif dist_type == 'compound':
                frequency_formula = dist_params[0]
                severity_formula = dist_params[1]
                deductible = dist_params[2] if len(dist_params) >= 3 else 0.0
                limit = dist_params[3] if len(dist_params) >= 4 else float("inf")
                return lambda q: compound_ppf(q, frequency_formula, severity_formula, deductible, limit)
            elif dist_type == 'splice':
                left_formula = dist_params[0]
                right_formula = dist_params[1]
                splice_point = dist_params[2]
                return lambda q: splice_ppf(q, left_formula, right_formula, splice_point)
            elif dist_type == 'pert':
                min_val, m_likely, max_val = dist_params[0], dist_params[1], dist_params[2]
                return lambda q: pert_ppf(q, min_val, m_likely, max_val)
            elif dist_type == 'reciprocal':
                return lambda q: reciprocal_ppf(q, dist_params[0], dist_params[1])
            elif dist_type == 'rayleigh':
                return lambda q: rayleigh_ppf(q, dist_params[0])
            elif dist_type == 'weibull':
                alpha, beta = dist_params[0], dist_params[1]
                return lambda q: weibull_ppf(q, alpha, beta)
            elif dist_type == 'pearson5':
                alpha, beta = dist_params[0], dist_params[1]
                return lambda q: pearson5_ppf(q, alpha, beta)
            elif dist_type == 'pearson6':
                alpha1, alpha2, beta = dist_params[0], dist_params[1], dist_params[2]
                return lambda q: pearson6_ppf(q, alpha1, alpha2, beta)
            elif dist_type == 'pareto2':
                b, q_shape = dist_params[0], dist_params[1]
                return lambda q: pareto2_ppf(q, b, q_shape)
            elif dist_type == 'pareto':
                theta, alpha = dist_params[0], dist_params[1]
                return lambda q: pareto_ppf(q, theta, alpha)
            elif dist_type == 'levy':
                a, c = dist_params[0], dist_params[1]
                return lambda q: levy_ppf(q, a, c)
            elif dist_type == 'general':
                if self.x_vals_list is not None and self.p_vals_list is not None:
                    min_val, max_val = dist_params[0], dist_params[1]
                    return lambda q: general_ppf(q, min_val, max_val, self.x_vals_list, self.p_vals_list)
                else:
                    return None
            elif dist_type == 'histogrm':
                if self.p_vals_list is not None:
                    min_val, max_val = dist_params[0], dist_params[1]
                    return lambda q: histogrm_ppf(q, min_val, max_val, self.p_vals_list)
                else:
                    return None
            elif dist_type == 'negbin':
                s, p = int(dist_params[0]), dist_params[1]
                return lambda q: negbin_ppf(q, s, p)
            elif dist_type == 'invgauss':
                mu, lam = dist_params[0], dist_params[1]
                return lambda q: invgauss_ppf(q, mu, lam)
            elif dist_type == 'geomet':
                p = dist_params[0]
                return lambda q: geomet_ppf(q, p)
            elif dist_type == 'hypergeo':
                n, D, M = int(dist_params[0]), int(dist_params[1]), int(dist_params[2])
                return lambda q: hypergeo_ppf(q, n, D, M)
            elif dist_type == 'intuniform':
                min_val, max_val = int(dist_params[0]), int(dist_params[1])
                return lambda q: intuniform_ppf(q, min_val, max_val)
            elif dist_type == 'duniform':
                if self.x_vals_list is not None:
                    return lambda q: duniform_ppf(q, self.x_vals_list)
                else:
                    return None
            elif dist_type == 'trigen':
                L, M, U, alpha, beta = dist_params
                return lambda q: trigen_ppf(q, L, M, U, alpha, beta)
            elif dist_type == 'cumul':
                if self.x_vals_list is not None and self.p_vals_list is not None:
                    return lambda q: cumul_ppf(q, self.x_vals_list, self.p_vals_list)
                else:
                    return None
            elif dist_type == 'discrete':
                if self.x_vals_list is not None and self.p_vals_list is not None:
                    return lambda q: discrete_ppf(q, self.x_vals_list, self.p_vals_list)
                else:
                    return None
            else:
                return None
        except ImportError:
            return None
    def _estimate_cdf(self, value: float) -> float:
        dist_type = self.dist_type
        dist_params = self.dist_params
        try:
            if dist_type == 'normal':
                from scipy.stats import norm
                return norm.cdf(value, dist_params[0], dist_params[1])
            elif dist_type == 'uniform':
                a, b = dist_params[0], dist_params[1]
                if value < a:
                    return 0.0
                elif value > b:
                    return 1.0
                else:
                    return (value - a) / (b - a)
            elif dist_type == 'gamma':
                from scipy.stats import gamma
                return gamma.cdf(value, dist_params[0], scale=dist_params[1])
            elif dist_type == 'erlang':
                return erlang_cdf(value, dist_params[0], dist_params[1])
            elif dist_type == 'beta':
                from scipy.stats import beta
                return beta.cdf(value, dist_params[0], dist_params[1])
            elif dist_type == 'chisq':
                from scipy.stats import chi2
                return chi2.cdf(value, dist_params[0])
            elif dist_type == 'f':
                from scipy.stats import f
                return f.cdf(value, dist_params[0], dist_params[1])
            elif dist_type == 'student':
                from scipy.stats import t
                return t.cdf(value, dist_params[0])
            elif dist_type == 'expon':
                from scipy.stats import expon
                return expon.cdf(value, scale=dist_params[0])
            elif dist_type == 'bernoulli':
                p = dist_params[0]
                return bernoulli_cdf(value, p)
            elif dist_type == 'triang':
                a, c, b = dist_params[0], dist_params[1], dist_params[2]
                return triang_cdf(value, a, c, b)
            elif dist_type == 'binomial':
                n, p = int(dist_params[0]), dist_params[1]
                return binomial_cdf(value, n, p)
            elif dist_type == 'cauchy':
                return cauchy_cdf(value, dist_params[0], dist_params[1])
            elif dist_type == 'erf':
                return erf_cdf(value, dist_params[0])
            elif dist_type == 'dagum':
                return dagum_cdf(value, dist_params[0], dist_params[1], dist_params[2], dist_params[3])
            elif dist_type == 'doubletriang':
                return doubletriang_cdf(value, dist_params[0], dist_params[1], dist_params[2], dist_params[3])
            elif dist_type == 'extvalue':
                return extvalue_cdf(value, dist_params[0], dist_params[1])
            elif dist_type == 'extvaluemin':
                return extvaluemin_cdf(value, dist_params[0], dist_params[1])
            elif dist_type == 'fatiguelife':
                return fatiguelife_cdf(value, dist_params[0], dist_params[1], dist_params[2])
            elif dist_type == 'frechet':
                return frechet_cdf(value, dist_params[0], dist_params[1], dist_params[2])
            elif dist_type == 'hypsecant':
                return hypsecant_cdf(value, dist_params[0], dist_params[1])
            elif dist_type == 'johnsonsb':
                return johnsonsb_cdf(value, dist_params[0], dist_params[1], dist_params[2], dist_params[3])
            elif dist_type == 'johnsonsu':
                return johnsonsu_cdf(value, dist_params[0], dist_params[1], dist_params[2], dist_params[3])
            elif dist_type == 'kumaraswamy':
                return kumaraswamy_cdf(value, dist_params[0], dist_params[1], dist_params[2], dist_params[3])
            elif dist_type == 'laplace':
                return laplace_cdf(value, dist_params[0], dist_params[1])
            elif dist_type == 'logistic':
                return logistic_cdf(value, dist_params[0], dist_params[1])
            elif dist_type == 'loglogistic':
                return loglogistic_cdf(value, dist_params[0], dist_params[1], dist_params[2])
            elif dist_type == 'lognorm':
                return lognorm_cdf(value, dist_params[0], dist_params[1])
            elif dist_type == 'lognorm2':
                return lognorm2_cdf(value, dist_params[0], dist_params[1])
            elif dist_type == 'betageneral':
                return betageneral_cdf(value, dist_params[0], dist_params[1], dist_params[2], dist_params[3])
            elif dist_type == 'betasubj':
                return betasubj_cdf(value, dist_params[0], dist_params[1], dist_params[2], dist_params[3])
            elif dist_type == 'burr12':
                return burr12_cdf(value, dist_params[0], dist_params[1], dist_params[2], dist_params[3])
            elif dist_type == 'compound':
                deductible = dist_params[2] if len(dist_params) >= 3 else 0.0
                limit = dist_params[3] if len(dist_params) >= 4 else float("inf")
                return compound_cdf(value, dist_params[0], dist_params[1], deductible, limit)
            elif dist_type == 'splice':
                return splice_cdf(value, dist_params[0], dist_params[1], dist_params[2])
            elif dist_type == 'pert':
                return pert_cdf(value, dist_params[0], dist_params[1], dist_params[2])
            elif dist_type == 'reciprocal':
                return reciprocal_cdf(value, dist_params[0], dist_params[1])
            elif dist_type == 'rayleigh':
                return rayleigh_cdf(value, dist_params[0])
            elif dist_type == 'weibull':
                return weibull_cdf(value, dist_params[0], dist_params[1])
            elif dist_type == 'pearson5':
                return pearson5_cdf(value, dist_params[0], dist_params[1])
            elif dist_type == 'pearson6':
                return pearson6_cdf(value, dist_params[0], dist_params[1], dist_params[2])
            elif dist_type == 'pareto2':
                return pareto2_cdf(value, dist_params[0], dist_params[1])
            elif dist_type == 'pareto':
                return pareto_cdf(value, dist_params[0], dist_params[1])
            elif dist_type == 'levy':
                return levy_cdf(value, dist_params[0], dist_params[1])
            elif dist_type == 'general':
                if self.x_vals_list is not None and self.p_vals_list is not None:
                    return general_cdf(value, dist_params[0], dist_params[1], self.x_vals_list, self.p_vals_list)
                else:
                    return 0.0
            elif dist_type == 'histogrm':
                if self.p_vals_list is not None:
                    return histogrm_cdf(value, dist_params[0], dist_params[1], self.p_vals_list)
                else:
                    return 0.0
            elif dist_type == 'negbin':
                s, p = int(dist_params[0]), dist_params[1]
                return negbin_cdf(value, s, p)
            elif dist_type == 'invgauss':
                return invgauss_cdf(value, dist_params[0], dist_params[1])
            elif dist_type == 'geomet':
                return geomet_cdf(value, dist_params[0])
            elif dist_type == 'hypergeo':
                return hypergeo_cdf(value, dist_params[0], dist_params[1], dist_params[2])
            elif dist_type == 'intuniform':
                return intuniform_cdf(value, dist_params[0], dist_params[1])
            elif dist_type == 'duniform':
                if self.x_vals_list is not None:
                    return duniform_cdf(value, self.x_vals_list)
                else:
                    return 0.0
            elif dist_type == 'trigen':
                L, M, U, alpha, beta = dist_params
                return trigen_cdf(value, L, M, U, alpha, beta)
            elif dist_type == 'cumul':
                if self.x_vals_list is not None and self.p_vals_list is not None:
                    return cumul_cdf(value, self.x_vals_list, self.p_vals_list)
                else:
                    return 0.5
            elif dist_type == 'discrete':
                if self.x_vals_list is not None and self.p_vals_list is not None:
                    return discrete_cdf(value, self.x_vals_list, self.p_vals_list)
                else:
                    return 0.5
            else:
                return 0.5
        except ImportError:
            return 0.5
    def _get_scipy_dist(self):
        if not SCIPY_AVAILABLE:
            return None
        try:
            if self.dist_type == 'normal':
                return sps.norm(loc=self.dist_params[0], scale=self.dist_params[1])
            elif self.dist_type == 'uniform':
                return sps.uniform(loc=self.dist_params[0], scale=self.dist_params[1]-self.dist_params[0])
            elif self.dist_type == 'erf':
                sigma = 1.0 / (math.sqrt(2.0) * self.dist_params[0])
                return sps.norm(loc=0.0, scale=sigma)
            elif self.dist_type == 'erlang':
                return sps.gamma(a=int(round(float(self.dist_params[0]))), scale=self.dist_params[1])
            elif self.dist_type == 'gamma':
                return sps.gamma(a=self.dist_params[0], scale=self.dist_params[1])
            elif self.dist_type == 'poisson':
                return sps.poisson(mu=self.dist_params[0])
            elif self.dist_type == 'beta':
                return sps.beta(a=self.dist_params[0], b=self.dist_params[1])
            elif self.dist_type == 'chisq':
                return sps.chi2(df=self.dist_params[0])
            elif self.dist_type == 'f':
                return sps.f(dfn=self.dist_params[0], dfd=self.dist_params[1])
            elif self.dist_type == 'student':
                return sps.t(df=self.dist_params[0])
            elif self.dist_type == 'cauchy':
                return sps.cauchy(loc=self.dist_params[0], scale=self.dist_params[1])
            elif self.dist_type == 'extvalue':
                return sps.gumbel_r(loc=self.dist_params[0], scale=self.dist_params[1])
            elif self.dist_type == 'extvaluemin':
                return sps.gumbel_l(loc=self.dist_params[0], scale=self.dist_params[1])
            elif self.dist_type == 'fatiguelife':
                return sps.fatiguelife(self.dist_params[2], loc=self.dist_params[0], scale=self.dist_params[1])
            elif self.dist_type == 'frechet':
                return sps.invweibull(self.dist_params[2], loc=self.dist_params[0], scale=self.dist_params[1])
            elif self.dist_type == 'hypsecant':
                return sps.hypsecant(loc=self.dist_params[0], scale=self.dist_params[1] * 2.0 / np.pi)
            elif self.dist_type == 'johnsonsb':
                return sps.johnsonsb(self.dist_params[0], self.dist_params[1], loc=self.dist_params[2], scale=self.dist_params[3] - self.dist_params[2])
            elif self.dist_type == 'johnsonsu':
                return sps.johnsonsu(self.dist_params[0], self.dist_params[1], loc=self.dist_params[2], scale=self.dist_params[3])
            elif self.dist_type == 'kumaraswamy':
                return None
            elif self.dist_type == 'laplace':
                return sps.laplace(loc=self.dist_params[0], scale=self.dist_params[1] / math.sqrt(2.0))
            elif self.dist_type == 'logistic':
                return sps.logistic(loc=self.dist_params[0], scale=self.dist_params[1])
            elif self.dist_type == 'loglogistic':
                return sps.fisk(self.dist_params[2], loc=self.dist_params[0], scale=self.dist_params[1])
            elif self.dist_type == 'lognorm':
                mu = self.dist_params[0]
                sigma = self.dist_params[1]
                sigma_prime = math.sqrt(math.log(1.0 + (sigma / mu) ** 2))
                mu_prime = math.log((mu * mu) / math.sqrt(sigma * sigma + mu * mu))
                return sps.lognorm(s=sigma_prime, scale=math.exp(mu_prime), loc=0.0)
            elif self.dist_type == 'lognorm2':
                return sps.lognorm(s=self.dist_params[1], scale=math.exp(self.dist_params[0]), loc=0.0)
            elif self.dist_type == 'betageneral':
                return sps.beta(
                    a=float(self.dist_params[0]),
                    b=float(self.dist_params[1]),
                    loc=float(self.dist_params[2]),
                    scale=float(self.dist_params[3]) - float(self.dist_params[2]),
                )
            elif self.dist_type == 'betasubj':
                from dist_betasubj import _solve_alpha_params
                alpha1, alpha2 = _solve_alpha_params(
                    float(self.dist_params[0]),
                    float(self.dist_params[1]),
                    float(self.dist_params[2]),
                    float(self.dist_params[3]),
                )
                return sps.beta(
                    a=float(alpha1),
                    b=float(alpha2),
                    loc=float(self.dist_params[0]),
                    scale=float(self.dist_params[3]) - float(self.dist_params[0]),
                )
            elif self.dist_type == 'burr12':
                return sps.burr12(
                    c=float(self.dist_params[2]),
                    d=float(self.dist_params[3]),
                    loc=float(self.dist_params[0]),
                    scale=float(self.dist_params[1]),
                )
            elif self.dist_type == 'compound':
                return None
            elif self.dist_type == 'splice':
                return None
            elif self.dist_type == 'pert':
                min_val, m_likely, max_val = self.dist_params[0], self.dist_params[1], self.dist_params[2]
                mean_val = (float(min_val) + 4.0 * float(m_likely) + float(max_val)) / 6.0
                alpha1 = 6.0 * (mean_val - float(min_val)) / (float(max_val) - float(min_val))
                alpha2 = 6.0 * (float(max_val) - mean_val) / (float(max_val) - float(min_val))
                return sps.beta(
                    a=alpha1,
                    b=alpha2,
                    loc=float(min_val),
                    scale=float(max_val) - float(min_val),
                )
            elif self.dist_type == 'reciprocal':
                return sps.reciprocal(a=float(self.dist_params[0]), b=float(self.dist_params[1]))
            elif self.dist_type == 'rayleigh':
                return sps.rayleigh(loc=0.0, scale=float(self.dist_params[0]))
            elif self.dist_type == 'weibull':
                return sps.weibull_min(c=float(self.dist_params[0]), loc=0.0, scale=float(self.dist_params[1]))
            elif self.dist_type == 'pearson5':
                return sps.invgamma(a=self.dist_params[0], scale=self.dist_params[1], loc=0.0)
            elif self.dist_type == 'pearson6':
                return sps.betaprime(
                    a=self.dist_params[0],
                    b=self.dist_params[1],
                    scale=self.dist_params[2],
                    loc=0.0,
                )
            elif self.dist_type == 'pareto2':
                return sps.lomax(c=self.dist_params[1], loc=0.0, scale=self.dist_params[0])
            elif self.dist_type == 'pareto':
                return sps.pareto(b=self.dist_params[0], loc=0.0, scale=self.dist_params[1])
            elif self.dist_type == 'levy':
                return sps.levy(loc=self.dist_params[0], scale=self.dist_params[1])
            elif self.dist_type == 'expon':
                return sps.expon(scale=self.dist_params[0])
            elif self.dist_type == 'hypergeo':
                return sps.hypergeom(M=int(self.dist_params[2]), n=int(self.dist_params[1]), N=int(self.dist_params[0]))
            elif self.dist_type == 'negbin':
                return sps.nbinom(int(self.dist_params[0]), self.dist_params[1])
            elif self.dist_type == 'intuniform':
                min_val = int(self.dist_params[0])
                max_val = int(self.dist_params[1])
                return sps.randint(low=min_val, high=max_val + 1)
            else:
                return None
        except Exception:
            return None
    def generate_sample(self, rng_seed: int = 42) -> float:
        if self._array_parse_failed or self.truncate_invalid:
            print("数组解析失败或截断参数无效，返回错误标记")
            return ERROR_MARKER
        static_value = self.markers.get('static')
        from attribute_functions import get_static_mode
        if static_value is not None and get_static_mode():
            return float(static_value)
        if self.markers.get('loc') or self.markers.get('lock'):
            loc_value = self._calculate_loc_value()
            if loc_value == ERROR_MARKER:
                return ERROR_MARKER
            return float(loc_value)
        # 对于三角分布和 Trigen，额外验证参数有效性
        if self.dist_type == 'triang':
            if len(self.dist_params) >= 3:
                a, c, b = self.dist_params[0], self.dist_params[1], self.dist_params[2]
                if not (a <= c <= b):
                    print(f"三角分布参数无效: a={a}, c={c}, b={b}，返回错误标记")
                    return ERROR_MARKER
        elif self.dist_type == 'trigen':
            if len(self.dist_params) >= 5:
                L, M, U, alpha, beta = self.dist_params
                # 验证 L <= M <= U 和 0 <= alpha < beta <= 1
                if not (L <= M <= U):
                    print(f"Trigen 参数 L<=M<=U 不满足: L={L}, M={M}, U={U}")
                    return ERROR_MARKER
                if not (0 <= alpha < beta <= 1):
                    print(f"Trigen 参数 alpha<beta 不满足: alpha={alpha}, beta={beta}")
                    return ERROR_MARKER
        import numpy as np
        rng = np.random.Generator(np.random.MT19937(seed=rng_seed))
        return self.generator_with_transform(rng, self.dist_params)
    def _raw_mean(self) -> float:
        """返回原始分布的均值（不考虑截断和平移）"""
        dist_type = self.dist_type
        params = self.dist_params
        if dist_type == 'normal':
            return params[0]
        elif dist_type == 'uniform':
            return (params[0] + params[1]) / 2
        elif dist_type == 'gamma':
            return params[0] * params[1]
        elif dist_type == 'poisson':
            return params[0]
        elif dist_type == 'beta':
            return params[0] / (params[0] + params[1])
        elif dist_type == 'chisq':
            return params[0]
        elif dist_type == 'f':
            if len(params) > 1 and params[1] > 2:
                return params[1] / (params[1] - 2)
            else:
                return params[0] if len(params) > 0 else 0.0
        elif dist_type == 'student':
            return 0.0
        elif dist_type == 'expon':
            return 1 / params[0] if params[0] != 0 else 0.0
        elif dist_type == 'bernoulli':
            return params[0]
        elif dist_type == 'triang':
            return (params[0] + params[1] + params[2]) / 3
        elif dist_type == 'binomial':
            return params[0] * params[1]
        elif dist_type == 'cauchy':
            return float('nan')
        elif dist_type == 'erf':
            return 0.0
        elif dist_type == 'erlang':
            return float(int(round(float(params[0])))) * float(params[1])
        elif dist_type == 'dagum':
            return dagum_raw_mean(params[0], params[1], params[2], params[3])
        elif dist_type == 'doubletriang':
            return doubletriang_raw_mean(params[0], params[1], params[2], params[3])
        elif dist_type == 'extvalue':
            return extvalue_raw_mean(params[0], params[1])
        elif dist_type == 'extvaluemin':
            return extvaluemin_raw_mean(params[0], params[1])
        elif dist_type == 'fatiguelife':
            return fatiguelife_raw_mean(params[0], params[1], params[2])
        elif dist_type == 'frechet':
            return frechet_raw_mean(params[0], params[1], params[2])
        elif dist_type == 'hypsecant':
            return hypsecant_raw_mean(params[0], params[1])
        elif dist_type == 'johnsonsb':
            return johnsonsb_raw_mean(params[0], params[1], params[2], params[3])
        elif dist_type == 'johnsonsu':
            return johnsonsu_raw_mean(params[0], params[1], params[2], params[3])
        elif dist_type == 'kumaraswamy':
            return kumaraswamy_raw_mean(params[0], params[1], params[2], params[3])
        elif dist_type == 'laplace':
            return laplace_raw_mean(params[0], params[1])
        elif dist_type == 'logistic':
            return logistic_raw_mean(params[0], params[1])
        elif dist_type == 'loglogistic':
            return loglogistic_raw_mean(params[0], params[1], params[2])
        elif dist_type == 'lognorm':
            return lognorm_raw_mean(params[0], params[1])
        elif dist_type == 'lognorm2':
            return lognorm2_raw_mean(params[0], params[1])
        elif dist_type == 'betageneral':
            return betageneral_raw_mean(params[0], params[1], params[2], params[3])
        elif dist_type == 'betasubj':
            return betasubj_raw_mean(params[0], params[1], params[2], params[3])
        elif dist_type == 'burr12':
            return burr12_raw_mean(params[0], params[1], params[2], params[3])
        elif dist_type == 'compound':
            deductible = params[2] if len(params) >= 3 else 0.0
            limit = params[3] if len(params) >= 4 else float("inf")
            return compound_raw_mean(params[0], params[1], deductible, limit)
        elif dist_type == 'splice':
            return splice_raw_mean(params[0], params[1], params[2])
        elif dist_type == 'pert':
            return pert_raw_mean(params[0], params[1], params[2])
        elif dist_type == 'reciprocal':
            return reciprocal_raw_mean(params[0], params[1])
        elif dist_type == 'rayleigh':
            return rayleigh_raw_mean(params[0])
        elif dist_type == 'weibull':
            return weibull_raw_mean(params[0], params[1])
        elif dist_type == 'pearson5':
            return pearson5_raw_mean(params[0], params[1])
        elif dist_type == 'pearson6':
            return pearson6_raw_mean(params[0], params[1], params[2])
        elif dist_type == 'pareto2':
            return pareto2_raw_mean(params[0], params[1])
        elif dist_type == 'pareto':
            return pareto_raw_mean(params[0], params[1])
        elif dist_type == 'levy':
            return levy_raw_mean(params[0], params[1])
        elif dist_type == 'general':
            if self.x_vals_list is not None and self.p_vals_list is not None:
                return general_raw_mean(params[0], params[1], self.x_vals_list, self.p_vals_list)
            return 0.0
        elif dist_type == 'histogrm':
            if self.p_vals_list is not None:
                return histogrm_raw_mean(params[0], params[1], self.p_vals_list)
            return 0.0
        elif dist_type == 'negbin':
            s = params[0]
            p = params[1]
            return s * (1.0 - p) / p
        elif dist_type == 'invgauss':
            return params[0]
        elif dist_type == 'geomet':
            p = params[0]
            return (1.0 / p) - 1.0
        elif dist_type == 'hypergeo':
            return params[0] * params[1] / params[2]
        elif dist_type == 'intuniform':
            return (params[0] + params[1]) / 2.0
        elif dist_type == 'duniform':
            if self.x_vals_list is not None and len(self.x_vals_list) > 0:
                return sum(self.x_vals_list) / len(self.x_vals_list)
            else:
                return 0.0
        elif dist_type == 'trigen':
            try:
                a, c, b = _convert_trigen_to_triang(params[0], params[1], params[2], params[3], params[4])
                return (a + c + b) / 3
            except Exception as e:
                print(f"Trigen 均值计算失败: {e}")
                return (params[0] + params[1] + params[2]) / 3
        elif dist_type == 'cumul':
            if self.x_vals_list is not None and self.p_vals_list is not None:
                return sum(x * p for x, p in zip(self.x_vals_list, self.p_vals_list))
            else:
                return 0.0
        elif dist_type == 'discrete':
            if self.x_vals_list is not None and self.p_vals_list is not None:
                return sum(x * p for x, p in zip(self.x_vals_list, self.p_vals_list))
            else:
                return 0.0
        else:
            return 0.0
    def _truncated_normal_mean(self, mu: float, sigma: float,
                               lower: Optional[float], upper: Optional[float]) -> float:
        import math
        if lower is None and upper is None:
            return mu
        sqrt2 = math.sqrt(2.0)
        sqrt2pi = math.sqrt(2.0 * math.pi)
        def phi(x):
            return math.exp(-0.5 * x * x) / sqrt2pi
        def Phi(x):
            return 0.5 * (1.0 + math.erf(x / sqrt2))
        if lower is None:
            beta = (upper - mu) / sigma
            phi_beta = phi(beta)
            Phi_beta = Phi(beta)
            if Phi_beta <= 0:
                return mu
            return mu - sigma * phi_beta / Phi_beta
        if upper is None:
            alpha = (lower - mu) / sigma
            phi_alpha = phi(alpha)
            Phi_alpha = Phi(alpha)
            one_minus_Phi = 1.0 - Phi_alpha
            if one_minus_Phi <= 0:
                return mu
            return mu + sigma * phi_alpha / one_minus_Phi
        alpha = (lower - mu) / sigma
        beta  = (upper - mu) / sigma
        phi_alpha = phi(alpha)
        phi_beta  = phi(beta)
        Phi_alpha = Phi(alpha)
        Phi_beta  = Phi(beta)
        Z = Phi_beta - Phi_alpha
        if Z <= 0:
            return (lower + upper) / 2.0
        return mu - sigma * (phi_beta - phi_alpha) / Z
    def _calculate_loc_value(self) -> float:
        """计算理论均值（用于 loc/lock 标记）- 统一使用分布类实现"""
        if self._array_parse_failed or self.truncate_invalid:
            return ERROR_MARKER
        # 使用分布类计算截断均值（精确）
        dist_class = self.DIST_CLASS_MAP.get(self.dist_type)
        if dist_class is not None:
            try:
                # 创建分布对象，传入所有标记（包括截断和shift）
                # 注意：分布类的 mean() 已经包含了 shift，我们减去 shift 得到原始均值
                dist_obj = dist_class(self.dist_params, self.markers, self.func_name)
                if not dist_obj.is_valid():
                    return ERROR_MARKER
                raw_mean = dist_obj.mean() - dist_obj.shift_amount
                value = raw_mean + self.shift_amount
                if self.dist_type in _LOCK_ROUND_INT_DIST_TYPES:
                    value = float(round(value))
                return value
            except Exception as e:
                print(f"使用分布类 {dist_class.__name__} 计算截断均值失败: {e}，回退到原始均值")
                value = self._raw_mean() + self.shift_amount
                if self.dist_type in _LOCK_ROUND_INT_DIST_TYPES:
                    value = float(round(value))
                return value
        else:
            # 没有对应分布类，回退到原始均值（理论上不应发生）
            value = self._raw_mean() + self.shift_amount
            if self.dist_type in _LOCK_ROUND_INT_DIST_TYPES:
                value = float(round(value))
            return value
        def _validate_params(self) -> bool:
            """验证参数是否有效 - 支持边界退化情况"""
            if self._array_parse_failed:
                return False

            dist_type = self.dist_type

            # 二项分布：允许 n=0, p=0, p=1
            if dist_type == 'binomial':
                if len(self.dist_params) >= 2:
                    n = self.dist_params[0]
                    p = self.dist_params[1]
                    if n < 0 or not float(n).is_integer():
                        return False
                    if p < 0 or p > 1:
                        return False
                    return True
                return False

            # 泊松分布：允许 lambda = 0
            if dist_type == 'poisson':
                if len(self.dist_params) >= 1 and self.dist_params[0] >= 0:
                    return True
                return False

            # 负二项分布：允许 s=0, p=1（p=0 仍无效）
            if dist_type == 'negbin':
                if len(self.dist_params) >= 2:
                    s = self.dist_params[0]
                    p = self.dist_params[1]
                    if s < 0 or abs(s - round(s)) > 1e-9:
                        return False
                    if not (0 < p <= 1):
                        return False
                    return True
                return False

            # 几何分布：允许 p=1（p=0 无效）
            if dist_type == 'geomet':
                if len(self.dist_params) >= 1:
                    p = self.dist_params[0]
                    if 0 < p <= 1:
                        return True
                    return False
                return False

            # 均匀分布：允许 min == max
            if dist_type == 'uniform':
                if len(self.dist_params) >= 2 and self.dist_params[1] >= self.dist_params[0]:
                    return True
                return False

            # 直方图分布：允许 min == max
            if dist_type == 'histogrm':
                if len(self.dist_params) >= 3:
                    min_val = self.dist_params[0]
                    max_val = self.dist_params[1]
                    if min_val <= max_val and self.p_vals_list is not None:
                        return True
                    return False
                return False

            # 一般分布：允许 min == max，允许内部 X 点等于边界
            if dist_type == 'general':
                if len(self.dist_params) >= 4:
                    min_val = self.dist_params[0]
                    max_val = self.dist_params[1]
                    if min_val <= max_val and self.x_vals_list is not None and self.p_vals_list is not None:
                        return True
                    return False
                return False

            # 三角分布：允许 a <= c <= b（包括相等）
            if dist_type == 'triang':
                if len(self.dist_params) >= 3:
                    a, c, b = self.dist_params[0], self.dist_params[1], self.dist_params[2]
                    if a <= c <= b:
                        return True
                    return False
                return False

            # PERT 分布：允许 min <= likely <= max（包括相等）
            if dist_type == 'pert':
                if len(self.dist_params) >= 3:
                    min_val, m_likely, max_val = self.dist_params[0], self.dist_params[1], self.dist_params[2]
                    if min_val <= m_likely <= max_val:
                        return True
                    return False
                return False

            # ========== 以下为其他分布的原有验证逻辑（未作改动） ==========
            if dist_type == 'normal':
                if len(self.dist_params) >= 2 and self.dist_params[1] > 0:
                    return True
                return False

            if dist_type == 'gamma':
                if len(self.dist_params) >= 2 and self.dist_params[0] > 0 and self.dist_params[1] > 0:
                    return True
                return False

            if dist_type == 'erlang':
                if len(self.dist_params) >= 2:
                    m = float(self.dist_params[0])
                    beta = float(self.dist_params[1])
                    if m > 0 and abs(m - round(m)) <= 1e-9 and beta > 0:
                        return True
                    return False
                return False

            if dist_type == 'beta':
                if len(self.dist_params) >= 2 and self.dist_params[0] > 0 and self.dist_params[1] > 0:
                    return True
                return False

            if dist_type == 'chisq':
                if len(self.dist_params) >= 1 and self.dist_params[0] > 0 and self.dist_params[0].is_integer():
                    return True
                return False

            if dist_type == 'f':
                if len(self.dist_params) >= 2:
                    if self.dist_params[0] > 0 and self.dist_params[1] > 0:
                        if self.dist_params[0].is_integer() and self.dist_params[1].is_integer():
                            return True
                    return False
                return False

            if dist_type == 'student':
                if len(self.dist_params) >= 1 and self.dist_params[0] > 0 and self.dist_params[0].is_integer():
                    return True
                return False

            if dist_type == 'expon':
                if len(self.dist_params) >= 1 and self.dist_params[0] > 0:
                    return True
                return False

            if dist_type == 'bernoulli':
                if len(self.dist_params) >= 1 and 0 <= self.dist_params[0] <= 1:
                    return True
                return False

            if dist_type == 'cauchy':
                if len(self.dist_params) >= 2 and self.dist_params[1] > 0:
                    return True
                return False

            if dist_type == 'erf':
                if len(self.dist_params) >= 1 and self.dist_params[0] > 0:
                    return True
                return False

            if dist_type == 'dagum':
                if len(self.dist_params) >= 4:
                    if self.dist_params[1] > 0 and self.dist_params[2] > 0 and self.dist_params[3] > 0:
                        return True
                    return False
                return False

            if dist_type == 'doubletriang':
                if len(self.dist_params) >= 4:
                    if (self.dist_params[0] < self.dist_params[1] < self.dist_params[2]) and (0 <= self.dist_params[3] <= 1):
                        return True
                    return False
                return False

            if dist_type == 'extvalue':
                if len(self.dist_params) >= 2 and self.dist_params[1] > 0:
                    return True
                return False

            if dist_type == 'extvaluemin':
                if len(self.dist_params) >= 2 and self.dist_params[1] > 0:
                    return True
                return False

            if dist_type == 'fatiguelife':
                if len(self.dist_params) >= 3 and self.dist_params[1] > 0 and self.dist_params[2] > 0:
                    return True
                return False

            if dist_type == 'frechet':
                if len(self.dist_params) >= 3 and self.dist_params[1] > 0 and self.dist_params[2] > 0:
                    return True
                return False

            if dist_type == 'hypsecant':
                if len(self.dist_params) >= 2 and self.dist_params[1] > 0:
                    return True
                return False

            if dist_type == 'johnsonsb':
                if len(self.dist_params) >= 4 and self.dist_params[1] > 0 and self.dist_params[3] > self.dist_params[2]:
                    return True
                return False

            if dist_type == 'johnsonsu':
                if len(self.dist_params) >= 4 and self.dist_params[1] > 0 and self.dist_params[3] > 0:
                    return True
                return False

            if dist_type == 'kumaraswamy':
                if len(self.dist_params) >= 4:
                    if (self.dist_params[0] > 0 and self.dist_params[1] > 0 and self.dist_params[3] > self.dist_params[2]):
                        return True
                    return False
                return False

            if dist_type == 'laplace':
                if len(self.dist_params) >= 2 and self.dist_params[1] > 0:
                    return True
                return False

            if dist_type == 'logistic':
                if len(self.dist_params) >= 2 and self.dist_params[1] > 0:
                    return True
                return False

            if dist_type == 'loglogistic':
                if len(self.dist_params) >= 3 and self.dist_params[1] > 0 and self.dist_params[2] > 0:
                    return True
                return False

            if dist_type == 'lognorm':
                if len(self.dist_params) >= 2 and self.dist_params[0] > 0 and self.dist_params[1] > 0:
                    return True
                return False

            if dist_type == 'lognorm2':
                if len(self.dist_params) >= 2 and self.dist_params[1] > 0:
                    return True
                return False

            if dist_type == 'betageneral':
                if len(self.dist_params) >= 4:
                    if (self.dist_params[0] > 0 and self.dist_params[1] > 0 and self.dist_params[3] > self.dist_params[2]):
                        return True
                    return False
                return False

            if dist_type == 'betasubj':
                if len(self.dist_params) >= 4:
                    if (self.dist_params[3] > self.dist_params[0] and
                        self.dist_params[1] > self.dist_params[0] and
                        self.dist_params[1] < self.dist_params[3] and
                        self.dist_params[2] > self.dist_params[0] and
                        self.dist_params[2] < self.dist_params[3]):
                        return True
                    return False
                return False

            if dist_type == 'burr12':
                if len(self.dist_params) >= 4:
                    if (self.dist_params[1] > 0 and self.dist_params[2] > 0 and self.dist_params[3] > 0):
                        return True
                    return False
                return False

            if dist_type == 'compound':
                if len(self.dist_params) < 2:
                    return False
                if not isinstance(self.dist_params[0], str) or not str(self.dist_params[0]).strip():
                    return False
                if not isinstance(self.dist_params[1], str) or not str(self.dist_params[1]).strip():
                    return False
                try:
                    deductible = float(self.dist_params[2]) if len(self.dist_params) >= 3 else 0.0
                    limit = float(self.dist_params[3]) if len(self.dist_params) >= 4 else float("inf")
                except Exception:
                    return False
                if deductible < 0 or limit < 0:
                    return False
                return True

            if dist_type == 'splice':
                if len(self.dist_params) < 3:
                    return False
                if not isinstance(self.dist_params[0], str) or not str(self.dist_params[0]).strip():
                    return False
                if not isinstance(self.dist_params[1], str) or not str(self.dist_params[1]).strip():
                    return False
                try:
                    splice_point = float(self.dist_params[2])
                except Exception:
                    return False
                if not math.isfinite(splice_point):
                    return False
                return True

            if dist_type == 'reciprocal':
                if len(self.dist_params) >= 2 and self.dist_params[0] > 0 and self.dist_params[1] > self.dist_params[0]:
                    return True
                return False

            if dist_type == 'rayleigh':
                if len(self.dist_params) >= 1 and self.dist_params[0] > 0:
                    return True
                return False

            if dist_type == 'weibull':
                if len(self.dist_params) >= 2 and self.dist_params[0] > 0 and self.dist_params[1] > 0:
                    return True
                return False

            if dist_type == 'pearson5':
                if len(self.dist_params) >= 2 and self.dist_params[0] > 0 and self.dist_params[1] > 0:
                    return True
                return False

            if dist_type == 'pearson6':
                if len(self.dist_params) >= 3 and self.dist_params[0] > 0 and self.dist_params[1] > 0 and self.dist_params[2] > 0:
                    return True
                return False

            if dist_type == 'pareto2':
                if len(self.dist_params) >= 2 and self.dist_params[0] > 0 and self.dist_params[1] > 0:
                    return True
                return False

            if dist_type == 'pareto':
                if len(self.dist_params) >= 2 and self.dist_params[0] > 0 and self.dist_params[1] > 0:
                    return True
                return False

            if dist_type == 'levy':
                if len(self.dist_params) >= 2 and self.dist_params[1] > 0:
                    return True
                return False

            if dist_type == 'invgauss':
                if len(self.dist_params) >= 2 and self.dist_params[0] > 0 and self.dist_params[1] > 0:
                    return True
                return False

            if dist_type == 'duniform':
                if self.x_vals_list is not None and len(self.x_vals_list) > 0:
                    return True
                return False

            if dist_type == 'cumul':
                # 验证由 _parse_cumul_4args 完成，此处仅检查参数数量
                if len(self.dist_params) >= 4:
                    return True
                return False

            if dist_type == 'discrete':
                if len(self.dist_params) >= 2 and self.x_vals_list is not None and self.p_vals_list is not None:
                    return True
                return False

            if dist_type == 'trigen':
                if len(self.dist_params) >= 5:
                    L, M, U, alpha, beta = self.dist_params[0], self.dist_params[1], self.dist_params[2], self.dist_params[3], self.dist_params[4]
                    if L <= M <= U and 0 <= alpha < beta <= 1:
                        return True
                    return False
                return False

            # 默认情况（未知分布类型）视为有效
            return True
# ==================== 全局常量和辅助函数 ====================
ERROR_MARKER = "#ERROR!"
def _is_excel_error(val) -> bool:
    """
    检测参数是否来自Excel错误值（如 #NAME?、#VALUE! 等）。
    pyxll 可能将错误值表示为特殊对象或字符串。
    """
    if hasattr(val, '__class__') and val.__class__.__name__ == 'ExcelError':
        return True
    if isinstance(val, str) and val.startswith('#'):
        return True
    # 某些情况下错误值可能是特殊浮点数，但暂不处理
    return False

def _extract_function_args_from_formula_robust(formula: str, func_name: str) -> Optional[List[str]]:
    """
    从任意 Excel 公式中提取指定函数的参数列表。
    支持前缀表达式、嵌套函数、单元格引用等。
    返回参数原始字符串列表，失败返回 None。
    """
    if not isinstance(formula, str) or not formula.strip():
        return None

    # 不区分大小写匹配函数名及其左括号
    pattern = re.compile(rf'\b{re.escape(func_name)}\s*\(', re.IGNORECASE)
    match = pattern.search(formula)
    if not match:
        return None

    start = match.end() - 1  # '(' 的位置
    # 括号匹配
    balance = 1
    i = start + 1
    n = len(formula)
    while i < n and balance > 0:
        ch = formula[i]
        if ch == '(':
            balance += 1
        elif ch == ')':
            balance -= 1
        i += 1
    if balance != 0:
        return None

    inner = formula[start+1:i-1]  # 括号内的原始内容

    # 按逗号分割参数，忽略括号内的逗号
    args = []
    current = []
    balance = 0
    for ch in inner:
        if ch == '(':
            balance += 1
            current.append(ch)
        elif ch == ')':
            balance -= 1
            current.append(ch)
        elif ch == ',' and balance == 0:
            args.append(''.join(current).strip())
            current = []
        else:
            current.append(ch)
    if current:
        args.append(''.join(current).strip())
    return args

def _convert_attribute_call_string(arg_text: str):
    """将嵌套公式里的属性函数字符串转换为内部 marker。"""
    if not isinstance(arg_text, str):
        return arg_text
    text = arg_text.strip()

    match = re.match(r'^(Drisk[A-Za-z0-9_]+)\((.*)\)$', text)
    if not match:
        return arg_text

    func_name = match.group(1)
    inner = match.group(2)

    if func_name == 'DriskLock':
        return create_marker_string('lock', "True")
    if func_name == 'DriskLoc':
        return create_marker_string('loc', "True")
    if func_name == 'DriskShift':
        return create_marker_string('shift', inner.strip())
    if func_name == 'DriskStatic':
        return create_marker_string('static', inner.strip() if inner.strip() else "0.0")
    if func_name == 'DriskTruncate':
        return create_marker_string('truncate', inner)
    if func_name == 'DriskTruncateP':
        return create_marker_string('truncatep', inner)
    if func_name == 'DriskTruncate2':
        return create_marker_string('truncate2', inner)

    return arg_text

_NESTED_LOCK_ROUND_INT_FUNCS = {
    "DriskBernoulli",
    "DriskBinomial",
    "DriskGeomet",
    "DriskHypergeo",
    "DriskIntuniform",
    "DriskNegbin",
}
_LOCK_ROUND_INT_DIST_TYPES = {
    "bernoulli",
    "binomial",
    "geomet",
    "hypergeo",
    "intuniform",
    "negbin",
}


def _has_invalid_strict_percentile_marker(markers: Dict[str, Any]) -> bool:
    """严格检查百分位截断参数，要求必须在 [0, 1] 内。"""
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
def parse_parameters(*args) -> Tuple[List[float], Dict[str, Any]]:
    """
    解析参数：支持任意数量和类型的参数
    返回：分布参数列表、标记字典
    修复：正确处理标记字符串，防止参数列表越界
    新增：对花括号数组如 "{1,2,3}" 进行识别，并存入 markers（但保留为普通参数，由上层处理）
    新增：对元组/列表（来自区域引用）直接保留，后续特殊处理
    新增：检测Excel错误值，发现则抛出异常
    """
    # 首先检查是否有参数本身就是错误值
    for arg in args:
        if _is_excel_error(arg):
            raise ValueError(f"参数中包含Excel错误值: {arg}")
    normal_params, markers = extract_markers_from_args(args)
    dist_params = []
    for param in normal_params:
        if param is None:
            continue
        elif isinstance(param, (int, float)):
            dist_params.append(float(param))
        elif isinstance(param, (tuple, list)):
            # 对于区域引用，pyxll 可能传入元组，直接保留，后续特殊处理
            dist_params.append(param)
        elif isinstance(param, str):
            param_str = str(param).strip()
            # 检查是否是标记字符串（以 __DRISK_...__: 开头）
            if is_marker_string(param_str):
                continue
            # 检查是否是Excel错误值（可能作为字符串传入）
            if param_str.startswith('#'):
                raise ValueError(f"参数中包含Excel错误值: {param_str}")
            # 检查花括号数组 {1,2,3} - 保留为字符串，不拆分为数值
            if param_str.startswith('{') and param_str.endswith('}'):
                # 保留原样，作为字符串参数
                dist_params.append(param_str)
                continue

            # 支持嵌套分布函数参数，例如 DriskIntuniform(DriskIntuniform(...,DriskLock()), 100)
            if re.match(r'^Drisk[A-Za-z_][A-Za-z0-9_]*\s*\(.*\)$', param_str):
                nested_func_name, nested_args = parse_complete_formula('=' + param_str)
                if nested_func_name in DISTRIBUTION_FUNCTION_NAMES:
                    resolved_nested_args = [_convert_attribute_call_string(arg) for arg in nested_args]
                    nested_value = _generic_distribution_function_with_simulation(nested_func_name, *resolved_nested_args)
                    if nested_value == ERROR_MARKER:
                        raise ValueError(f"无法解析嵌套分布参数 '{param_str}'")
                    if (
                        nested_func_name in _NESTED_LOCK_ROUND_INT_FUNCS
                        and ("DriskLock(" in param_str or "DriskLoc(" in param_str)
                    ):
                        nested_value = float(round(float(nested_value)))
                    dist_params.append(float(nested_value))
                    continue
            
            if re.match(r'^[A-Z]{1,3}\d{1,7}(?::[A-Z]{1,3}\d{1,7})?$', param_str, re.IGNORECASE) or ('!' in param_str and re.match(r'.+![A-Z]{1,3}\d{1,7}(?::[A-Z]{1,3}\d{1,7})?$', param_str, re.IGNORECASE)):
                dist_params.append(param_str)
                continue
            if ',' in param_str:
                parts = [p.strip() for p in param_str.split(',')]
                try:
                    [float(p) for p in parts if p]
                    # 验证通过，保留原字符串
                    dist_params.append(param_str)
                    continue
                except ValueError:
                    pass
            # 尝试转换为数值
            try:
                num_val = float(param_str)
                dist_params.append(num_val)
            except ValueError:
                # 尝试数学表达式
                try:
                    safe_dict = {
                        '__builtins__': {},
                        'abs': abs,
                        'sqrt': np.sqrt,
                        'exp': np.exp,
                        'log': np.log,
                        'log10': np.log10,
                        'sin': np.sin,
                        'cos': np.cos,
                        'tan': np.tan,
                        'pi': math.pi,
                        'e': math.e
                    }
                    for name in dir(math):
                        if not name.startswith('_') and name not in ['dist', 'errstate']:
                            safe_dict[name] = getattr(math, name)
                    expr = param_str.replace('^', '**')
                    value = eval(expr, {"__builtins__": {}}, safe_dict)
                    dist_params.append(float(value))
                except:
                    # 尝试单元格引用
                    if re.match(r'^[A-Z]{1,3}\d{1,7}$', param_str, re.IGNORECASE):
                        try:
                            from pyxll import xl_app
                            app = xl_app()
                            cell = app.ActiveSheet.Range(param_str)
                            cell_value = cell.Value
                            if cell_value is None:
                                cell_value = cell.Value2
                            if cell_value is not None:
                                try:
                                    dist_params.append(float(cell_value))
                                except:
                                    dist_params.append(0.0)
                            else:
                                dist_params.append(0.0)
                        except Exception:
                            dist_params.append(0.0)
                    else:
                        # 尝试百分比
                        try:
                            if param_str.endswith('%'):
                                num = float(param_str[:-1]) / 100.0
                                dist_params.append(num)
                            else:
                                # 无法解析，抛出异常
                                raise ValueError(f"无法解析参数 '{param_str}'")
                        except:
                            raise ValueError(f"无法解析参数 '{param_str}'")
        else:
            try:
                val = float(param)
                dist_params.append(val)
            except:
                raise ValueError(f"无法解析参数类型: {type(param)}")
    return dist_params, markers
def create_distribution_generator(func_name: str, dist_type: str, dist_params: List[float], markers: Dict[str, Any]) -> DistributionGenerator:
    return DistributionGenerator(func_name, dist_type, dist_params, markers)


def _get_formula_from_caller() -> Optional[str]:
    try:
        app = xl_app()
        caller = app.Caller
        if hasattr(caller, "Formula"):
            formula = caller.Formula
            if formula and isinstance(formula, str):
                return formula
    except Exception:
        pass
    return None

def _extract_compound_raw_args_from_formula(formula: str) -> Optional[List[str]]:
    # 优先使用稳健提取器
    raw_args = _extract_function_args_from_formula_robust(formula, "DriskCompound")
    if raw_args is None:
        # 回退到旧方法（兼容性）
        raw_args = _extract_function_args_from_formula(formula, "DriskCompound")
    if raw_args is None:
        return None

    resolved = []
    for arg in raw_args:
        resolved.append(_resolve_distribution_formula(arg))
    return resolved

def _extract_splice_raw_args_from_formula(formula: str) -> Optional[List[str]]:
    raw_args = _extract_function_args_from_formula_robust(formula, "DriskSplice")
    if raw_args is None:
        raw_args = _extract_function_args_from_formula(formula, "DriskSplice")
    if raw_args is None:
        return None

    resolved = []
    for arg in raw_args:
        resolved.append(_resolve_distribution_formula(arg))
    return resolved

def _compound_distribution_function_with_simulation(*args):
    try:
        raw_args = None
        caller_formula = _get_formula_from_caller()
        if caller_formula:
            raw_args = _extract_compound_raw_args_from_formula(caller_formula)

        if not raw_args or len(raw_args) < 2:
            print("DriskCompound: 无法从公式中提取参数，请确保调用格式正确")
            return ERROR_MARKER

        frequency_formula = raw_args[0].strip()
        severity_formula = raw_args[1].strip()
        deductible_text = raw_args[2].strip() if len(raw_args) >= 3 else ""
        limit_text = raw_args[3].strip() if len(raw_args) >= 4 else ""

        extra_args = list(args[4:]) if len(args) > 4 else []
        dist_params = [
            frequency_formula,
            severity_formula,
            0.0 if deductible_text == "" else float(deductible_text),
            float("inf") if limit_text == "" else float(limit_text),
        ]
        markers = {}
        if extra_args:
            _, markers = parse_parameters(*extra_args)

        info = get_distribution_info("DriskCompound")
        if not info:
            return ERROR_MARKER
        if not validate_distribution_params("DriskCompound", dist_params):
            return ERROR_MARKER

        static_mode = get_static_mode()
        simulation_running = _simulation_state.is_simulation_running()
        args_text = ",".join(str(arg) for arg in raw_args)

        if simulation_running and not static_mode:
            stable_key = _simulation_state.find_stable_key_by_params(
                "DriskCompound", dist_params, False, 0, None, args_text
            )
            if stable_key:
                return _simulation_state.get_or_generate_value(stable_key)

            func_info = {
                "func_name": "DriskCompound",
                "dist_params": dist_params,
                "markers": markers,
                "args_text_normalized": args_text,
                "args_text": args_text,
                "is_nested": False,
                "depth": 0,
                "parent_key": None,
            }
            temp_stable_key = f"temp_DriskCompound_{hash(args_text)}_{int(time.time())}"
            _simulation_state.register_distribution(temp_stable_key, func_info)
            return _simulation_state.get_or_generate_value(temp_stable_key)

        if markers.get("loc") or markers.get("lock"):
            generator = DistributionGenerator("DriskCompound", info["type"], dist_params, markers)
            loc_value = generator._calculate_loc_value()
            if loc_value == ERROR_MARKER:
                return ERROR_MARKER
            return float(loc_value)

        generator = DistributionGenerator("DriskCompound", info["type"], dist_params, markers)
        seed = int(time.time() * 1000000) % 1000000
        sample = generator.generate_sample(seed)
        if sample == ERROR_MARKER:
            return ERROR_MARKER
        return float(sample)
    except Exception as e:
        print(f"DriskCompound error: {str(e)}")
        import traceback
        traceback.print_exc()
        return ERROR_MARKER
    
def _splice_distribution_function_with_simulation(*args):
    try:
        raw_args = None
        caller_formula = _get_formula_from_caller()
        if caller_formula:
            raw_args = _extract_splice_raw_args_from_formula(caller_formula)

        if not raw_args or len(raw_args) < 3:
            print("DriskSplice: 无法从公式中提取参数，请确保调用格式正确")
            return ERROR_MARKER

        left_formula = raw_args[0].strip()
        right_formula = raw_args[1].strip()
        try:
            splice_point = float(raw_args[2].strip())
        except:
            print("DriskSplice: 拼接点参数无效")
            return ERROR_MARKER

        extra_args = list(args[3:]) if len(args) > 3 else []
        dist_params = [left_formula, right_formula, splice_point]
        markers = {}
        if extra_args:
            _, markers = parse_parameters(*extra_args)

        info = get_distribution_info("DriskSplice")
        if not info:
            return ERROR_MARKER
        if not validate_distribution_params("DriskSplice", dist_params):
            return ERROR_MARKER

        static_mode = get_static_mode()
        simulation_running = _simulation_state.is_simulation_running()
        args_text = ",".join(str(arg) for arg in raw_args)

        if simulation_running and not static_mode:
            stable_key = _simulation_state.find_stable_key_by_params(
                "DriskSplice", dist_params, False, 0, None, args_text
            )
            if stable_key:
                return _simulation_state.get_or_generate_value(stable_key)

            func_info = {
                "func_name": "DriskSplice",
                "dist_params": dist_params,
                "markers": markers,
                "args_text_normalized": args_text,
                "args_text": args_text,
                "is_nested": False,
                "depth": 0,
                "parent_key": None,
            }
            temp_stable_key = f"temp_DriskSplice_{hash(args_text)}_{int(time.time())}"
            _simulation_state.register_distribution(temp_stable_key, func_info)
            return _simulation_state.get_or_generate_value(temp_stable_key)

        if markers.get("loc") or markers.get("lock"):
            generator = DistributionGenerator("DriskSplice", info["type"], dist_params, markers)
            loc_value = generator._calculate_loc_value()
            if loc_value == ERROR_MARKER:
                return ERROR_MARKER
            return float(loc_value)

        generator = DistributionGenerator("DriskSplice", info["type"], dist_params, markers)
        seed = int(time.time() * 1000000) % 1000000
        sample = generator.generate_sample(seed)
        if sample == ERROR_MARKER:
            return ERROR_MARKER
        return float(sample)
    except Exception as e:
        print(f"DriskSplice error: {str(e)}")
        import traceback
        traceback.print_exc()
        return ERROR_MARKER


def _generic_distribution_function_with_simulation(func_name: str, *args):
    try:
        dist_params, markers = parse_parameters(*args)
        from constants import get_distribution_info, validate_distribution_params
        info = get_distribution_info(func_name)
        if not info:
            print(f"错误: 未找到分布函数信息 {func_name}")
            return ERROR_MARKER
        if func_name in {"DriskIntuniform", "DriskErlang"} and _has_invalid_strict_percentile_marker(markers):
            print("错误: Intuniform 的百分位截断参数必须介于 0 和 1 之间")
            return ERROR_MARKER
        min_params = info.get('min_params', 0)
        max_params = info.get('max_params', float('inf'))
        # ===== 关键修复：对 DriskCumul 特殊处理（四个参数） =====
        if func_name == 'DriskCumul':
            if len(dist_params) >= 4:
                # 前两个参数是 min 和 max
                min_val = dist_params[0]
                max_val = dist_params[1]
                x_vals_param = dist_params[2]
                p_vals_param = dist_params[3]
                # 改进的 to_string 函数：从各种输入中提取数值并生成干净的逗号分隔字符串
                def to_string(val):
                    # 收集数值的列表
                    numbers = []
                    def extract_values(item):
                        if isinstance(item, (tuple, list, np.ndarray)):
                            # 如果是容器类型，递归提取
                            for sub_item in item:
                                extract_values(sub_item)
                        elif isinstance(item, str):
                            # 字符串类型，尝试转换为数值
                            try:
                                clean_item = item.strip()
                                # 处理可能的花括号（数组常量）
                                if clean_item.startswith('{') and clean_item.endswith('}'):
                                    # 提取花括号内的内容
                                    inner = clean_item[1:-1].strip()
                                    # 分割逗号
                                    parts = [p.strip() for p in inner.split(',') if p.strip()]
                                    for part in parts:
                                        try:
                                            val = float(part)
                                            numbers.append(val)
                                        except:
                                            pass
                                else:
                                    # 普通字符串，尝试转换
                                    val = float(clean_item)
                                    numbers.append(val)
                            except (ValueError, TypeError):
                                # 不能转换为数值，跳过
                                pass
                        else:
                            # 其他类型（通常是数值）
                            try:
                                if item is not None:
                                    val = float(item)
                                    numbers.append(val)
                            except (ValueError, TypeError):
                                # 不能转换为数值，跳过
                                pass
                    extract_values(val)
                    # 将数值转换为字符串并用逗号连接
                    return ','.join(str(num) for num in numbers)
                # 解析内部点
                x_inner_str = to_string(x_vals_param)
                p_inner_str = to_string(p_vals_param)
                # 将内部点字符串转换为数值列表，用于构建完整数组
                def parse_numbers(s):
                    s = s.strip()
                    if s.startswith('{') and s.endswith('}'):
                        s = s[1:-1]
                    parts = [p.strip() for p in s.split(',') if p.strip()]
                    return [float(p) for p in parts]
                x_inner = parse_numbers(x_inner_str)
                p_inner = parse_numbers(p_inner_str)
                # 验证内部点不能为空
                if len(x_inner) == 0 or len(p_inner) == 0:
                    print(f"错误: X-Table 或 P-Table 不能为空")
                    return ERROR_MARKER
                # 验证内部点
                if len(x_inner) != len(p_inner):
                    print(f"错误: X-Table 和 P-Table 长度必须相等")
                    return ERROR_MARKER
                for x in x_inner:
                    if not (min_val - 1e-12 <= x <= max_val + 1e-12):
                        print(f"错误: X 值 {x} 不在 [{min_val}, {max_val}] 范围内")
                        return ERROR_MARKER
                for p in p_inner:
                    if not (0 <= p <= 1):
                        print(f"错误: P 值 {p} 不在 [0,1] 范围内")
                        return ERROR_MARKER
                # 构造完整数组
                x_full = [float(min_val)]
                p_full = [0.0]
                for x, p in zip(x_inner, p_inner):
                    if abs(x - min_val) <= 1e-12 or abs(x - max_val) <= 1e-12:
                        continue
                    x_full.append(float(x))
                    p_full.append(float(p))
                x_full.append(float(max_val))
                p_full.append(1.0)
                # 转换为字符串存入 markers
                x_full_str = ','.join(str(x) for x in x_full)
                p_full_str = ','.join(str(p) for p in p_full)
                # 额外验证数组的有效性（递增等）
                try:
                    from dist_cumul import _parse_arrays as cumul_parse_arrays
                    cumul_parse_arrays(x_full, p_full)
                except Exception as e:
                    print(f"错误: Cumul 数组验证失败: {e}")
                    return ERROR_MARKER
                markers['x_vals'] = x_full_str
                markers['p_vals'] = p_full_str
                # 保留 min 和 max 作为普通参数（可选，CumulDistribution 需要四个参数）
                dist_params = [min_val, max_val, x_full_str, p_full_str]
            else:
                print(f"错误: {func_name} 需要 4 个参数")
                return ERROR_MARKER
        
        elif func_name == 'DriskGeneral':
            if len(dist_params) >= 4:
                min_val = dist_params[0]
                max_val = dist_params[1]
                x_vals_param = dist_params[2]
                p_vals_param = dist_params[3]
                def to_string(val):
                    numbers = []
                    def extract_values(item):
                        if isinstance(item, (tuple, list, np.ndarray)):
                            for sub_item in item:
                                extract_values(sub_item)
                        elif isinstance(item, str):
                            try:
                                clean_item = item.strip()
                                if clean_item.startswith('{') and clean_item.endswith('}'):
                                    inner = clean_item[1:-1].strip()
                                    parts = [p.strip() for p in inner.split(',') if p.strip()]
                                    for part in parts:
                                        try:
                                            numbers.append(float(part))
                                        except:
                                            pass
                                else:
                                    range_pattern = r'^[A-Z]{1,3}\d{1,7}(?::[A-Z]{1,3}\d{1,7})?$'
                                    if re.match(range_pattern, clean_item, re.IGNORECASE) or ('!' in clean_item and re.match(r'.+![A-Z]{1,3}\d{1,7}(?::[A-Z]{1,3}\d{1,7})?$', clean_item, re.IGNORECASE)):
                                        try:
                                            from pyxll import xl_app
                                            app = xl_app()
                                            if '!' in clean_item:
                                                sheet_name, addr = clean_item.split('!', 1)
                                                range_obj = app.ActiveWorkbook.Worksheets(sheet_name).Range(addr)
                                            else:
                                                range_obj = app.ActiveSheet.Range(clean_item)
                                            values = range_obj.Value2
                                            def flatten_range(v):
                                                if isinstance(v, (tuple, list)):
                                                    for sub in v:
                                                        flatten_range(sub)
                                                elif v is not None:
                                                    numbers.append(float(v))
                                            flatten_range(values)
                                        except Exception:
                                            pass
                                    else:
                                        numbers.append(float(clean_item))
                            except:
                                pass
                        else:
                            try:
                                if item is not None:
                                    numbers.append(float(item))
                            except:
                                pass
                    extract_values(val)
                    return ','.join(str(num) for num in numbers)
                x_vals_str = to_string(x_vals_param)
                p_vals_str = to_string(p_vals_param)
                if not x_vals_str or not p_vals_str:
                    print("错误: General 的 X-Table 和 P-Table 不能为空")
                    return ERROR_MARKER
                try:
                    general_parse_arrays(float(min_val), float(max_val), x_vals_str, p_vals_str)
                except Exception as e:
                    print(f"错误: General 参数无效: {e}")
                    return ERROR_MARKER
                markers['x_vals'] = x_vals_str
                markers['p_vals'] = p_vals_str
                dist_params = [min_val, max_val, x_vals_str, p_vals_str]
            else:
                print(f"错误: {func_name} 需要 4 个参数")
                return ERROR_MARKER

        elif func_name == 'DriskHistogrm':
            if len(dist_params) >= 3:
                min_val = dist_params[0]
                max_val = dist_params[1]
                p_vals_param = dist_params[2]

                def to_string(val):
                    numbers = []

                    def extract_values(item):
                        if isinstance(item, (tuple, list, np.ndarray)):
                            for sub_item in item:
                                extract_values(sub_item)
                        elif isinstance(item, str):
                            try:
                                clean_item = item.strip()
                                if clean_item.startswith('{') and clean_item.endswith('}'):
                                    inner = clean_item[1:-1].strip()
                                    parts = [p.strip() for p in inner.split(',') if p.strip()]
                                    for part in parts:
                                        try:
                                            numbers.append(float(part))
                                        except:
                                            pass
                                else:
                                    range_pattern = r'^[A-Z]{1,3}\d{1,7}(?::[A-Z]{1,3}\d{1,7})?$'
                                    if re.match(range_pattern, clean_item, re.IGNORECASE) or ('!' in clean_item and re.match(r'.+![A-Z]{1,3}\d{1,7}(?::[A-Z]{1,3}\d{1,7})?$', clean_item, re.IGNORECASE)):
                                        try:
                                            from pyxll import xl_app
                                            app = xl_app()
                                            if '!' in clean_item:
                                                sheet_name, addr = clean_item.split('!', 1)
                                                range_obj = app.ActiveWorkbook.Worksheets(sheet_name).Range(addr)
                                            else:
                                                range_obj = app.ActiveSheet.Range(clean_item)
                                            values = range_obj.Value2

                                            def flatten_range(v):
                                                if isinstance(v, (tuple, list)):
                                                    for sub in v:
                                                        flatten_range(sub)
                                                elif v is not None:
                                                    numbers.append(float(v))

                                            flatten_range(values)
                                        except Exception:
                                            pass
                                    else:
                                        numbers.append(float(clean_item))
                            except:
                                pass
                        else:
                            try:
                                if item is not None:
                                    numbers.append(float(item))
                            except:
                                pass

                    extract_values(val)
                    return ','.join(str(num) for num in numbers)

                p_vals_str = to_string(p_vals_param)
                if not p_vals_str:
                    print("错误: Histogrm 的 P-Table 不能为空")
                    return ERROR_MARKER
                try:
                    histogrm_parse_p_table(p_vals_str)
                except Exception as e:
                    print(f"错误: Histogrm 参数无效: {e}")
                    return ERROR_MARKER
                markers['p_vals'] = p_vals_str
                dist_params = [min_val, max_val, p_vals_str]
            else:
                print(f"错误: {func_name} 需要 3 个参数")
                return ERROR_MARKER

        elif func_name == 'DriskDUniform':
            if len(dist_params) >= 1:
                x_vals_param = dist_params[0]
                def to_string(val):
                    numbers = []
                    def extract_values(item):
                        if isinstance(item, (tuple, list, np.ndarray)):
                            for sub_item in item:
                                extract_values(sub_item)
                        elif isinstance(item, str):
                            try:
                                clean_item = item.strip()
                                if clean_item.startswith('{') and clean_item.endswith('}'):
                                    inner = clean_item[1:-1].strip()
                                    parts = [p.strip() for p in inner.split(',') if p.strip()]
                                    for part in parts:
                                        try:
                                            numbers.append(float(part))
                                        except:
                                            pass
                                else:
                                    range_pattern = r'^[A-Z]{1,3}\d{1,7}(?::[A-Z]{1,3}\d{1,7})?$'
                                    if re.match(range_pattern, clean_item, re.IGNORECASE) or ('!' in clean_item and re.match(r'.+![A-Z]{1,3}\d{1,7}(?::[A-Z]{1,3}\d{1,7})?$', clean_item, re.IGNORECASE)):
                                        try:
                                            from pyxll import xl_app
                                            app = xl_app()
                                            if '!' in clean_item:
                                                sheet_name, addr = clean_item.split('!', 1)
                                                range_obj = app.ActiveWorkbook.Worksheets(sheet_name).Range(addr)
                                            else:
                                                range_obj = app.ActiveSheet.Range(clean_item)
                                            values = range_obj.Value2
                                            def flatten_range(v):
                                                if isinstance(v, (tuple, list)):
                                                    for sub in v:
                                                        flatten_range(sub)
                                                elif v is not None:
                                                    numbers.append(float(v))
                                            flatten_range(values)
                                        except Exception:
                                            pass
                                    else:
                                        numbers.append(float(clean_item))
                            except:
                                pass
                        else:
                            try:
                                if item is not None:
                                    numbers.append(float(item))
                            except:
                                pass
                    extract_values(val)
                    return ','.join(str(num) for num in numbers)
                x_vals_str = to_string(x_vals_param)
                if not x_vals_str:
                    print("??: X-Table ????")
                    return ERROR_MARKER
                x_vals = [float(p.strip()) for p in x_vals_str.split(',') if p.strip()]
                if len(x_vals) == 0:
                    print("??: X-Table ????")
                    return ERROR_MARKER
                if len(x_vals) != len(set(x_vals)):
                    print("??: X-Table ?????")
                    return ERROR_MARKER
                p = 1.0 / len(x_vals)
                p_vals_str = ','.join(str(p) for _ in x_vals)
                markers['x_vals'] = x_vals_str
                markers['p_vals'] = p_vals_str
                dist_params = [x_vals_str]
            else:
                print(f"??: {func_name} ????????")
                return ERROR_MARKER
        # ===== 处理 DriskDiscrete（保持原样两个参数） =====
        elif func_name == 'DriskDiscrete':
            if len(dist_params) >= 2:
                x_vals_param = dist_params[0]
                p_vals_param = dist_params[1]
                def to_string(val):
                    numbers = []
                    def extract_values(item):
                        if isinstance(item, (tuple, list, np.ndarray)):
                            for sub_item in item:
                                extract_values(sub_item)
                        elif isinstance(item, str):
                            try:
                                clean_item = item.strip()
                                if clean_item.startswith('{') and clean_item.endswith('}'):
                                    inner = clean_item[1:-1].strip()
                                    parts = [p.strip() for p in inner.split(',') if p.strip()]
                                    for part in parts:
                                        try:
                                            val = float(part)
                                            numbers.append(val)
                                        except:
                                            pass
                                else:
                                    val = float(clean_item)
                                    numbers.append(val)
                            except:
                                pass
                        else:
                            try:
                                if item is not None:
                                    val = float(item)
                                    numbers.append(val)
                            except:
                                pass
                    extract_values(val)
                    return ','.join(str(num) for num in numbers)
                x_vals_str = to_string(x_vals_param)
                p_vals_str = to_string(p_vals_param)
                # 对于 discrete，空数组也是允许的？但分布需要至少一个点，这里假设至少一个点
                if not x_vals_str or not p_vals_str:
                    print(f"错误: X-Table 或 P-Table 不能为空")
                    return ERROR_MARKER
                markers['x_vals'] = x_vals_str
                markers['p_vals'] = p_vals_str
                dist_params = [x_vals_str, p_vals_str]
            else:
                print(f"错误: {func_name} 需要至少两个参数")
                return ERROR_MARKER
        # ========== 其他分布：检查参数数量 ==========
        else:
            if len(dist_params) < min_params or len(dist_params) > max_params:
                print(f"错误: 参数数量无效 {func_name}, 期望 {min_params}-{max_params}, 实际 {len(dist_params)}")
                return ERROR_MARKER
        if not validate_distribution_params(func_name, dist_params):
            print(f"错误: 参数验证失败 {func_name}, 参数 {dist_params}")
            return ERROR_MARKER
        from attribute_functions import get_static_mode
        static_mode = get_static_mode()
        simulation_running = _simulation_state.is_simulation_running()
        is_nested = False
        depth = 0
        nested_parent_key = None
        args_text = ','.join([str(arg) for arg in args])
        for arg in args:
            if isinstance(arg, str):
                if 'Drisk' in arg and any(fname in arg for fname in DISTRIBUTION_FUNCTION_NAMES):
                    is_nested = True
                    depth = 1
                    print(f"检测到嵌套函数调用: {arg}, 深度={depth}")
        if simulation_running and not static_mode:
            param_strs = []
            for param in dist_params:
                if isinstance(param, (int, float)):
                    param_strs.append(f"{param:.10f}".rstrip('0').rstrip('.'))
                else:
                    param_strs.append(str(param))
            args_text_normalized = ','.join(param_strs)
            stable_key = _simulation_state.find_stable_key_by_params(
                func_name, dist_params, is_nested, depth, nested_parent_key, args_text
            )
            if stable_key:
                value = _simulation_state.get_or_generate_value(stable_key)
                _simulation_state.log_debug(f"通过稳定键获取值: {func_name}, 稳定键={stable_key}, 值={value:.6f}, 嵌套={is_nested}, 深度={depth}, 参数文本={args_text}")
                return value
            else:
                func_info = {
                    'func_name': func_name,
                    'dist_params': dist_params,
                    'markers': markers,
                    'args_text_normalized': args_text_normalized,
                    'args_text': args_text,
                    'is_nested': is_nested,
                    'depth': depth,
                    'parent_key': nested_parent_key
                }
                temp_stable_key = f"temp_{func_name}_{hash(args_text)}_{int(time.time())}"
                _simulation_state.register_distribution(temp_stable_key, func_info)
                value = _simulation_state.get_or_generate_value(temp_stable_key)
                _simulation_state.log_debug(f"临时注册并生成值: {func_name}, 临时稳定键={temp_stable_key}, 值={value:.6f}, 嵌套={is_nested}, 深度={depth}, 参数文本={args_text}")
                return value
        else:
            static_value = markers.get('static')
            if static_value is not None and static_mode:
                return float(static_value)
            if markers.get('loc') or markers.get('lock'):
                generator = DistributionGenerator(func_name, info['type'], dist_params, markers)
                loc_value = generator._calculate_loc_value()
                if loc_value == ERROR_MARKER:
                    return ERROR_MARKER
                return float(loc_value)
            generator = DistributionGenerator(func_name, info['type'], dist_params, markers)
            try:
                from pyxll import xlfCaller
                caller = xlfCaller()
                if caller:
                    cell_addr = str(caller)
                    seed = _simulation_state.get_unique_seed_for_cell(cell_addr)
                else:
                    seed = int(time.time() * 1000000) % 1000000
            except:
                seed = int(time.time() * 1000000) % 1000000
            sample = generator.generate_sample(seed)
            if sample == ERROR_MARKER:
                return ERROR_MARKER
            return float(sample)
    except Exception as e:
        print(f"{func_name}错误: {str(e)}")
        import traceback
        traceback.print_exc()
        return ERROR_MARKER
# ==================== 分布函数定义（使用新架构） ====================
@xl_func("var Mean, var Std_Dev, var 属性1, var 属性2, var 属性3, var 属性4, var 属性5, var 属性6, var 属性7, var 属性8, var 属性9, var 属性10, var 属性11, var 属性12, var 属性13, var 属性14, var 属性15, var 属性16, var 属性17, var 属性18, var 属性19, var 属性20, var*: var", volatile=True)
def DriskNormal(Mean, Std_Dev, *args):
    """
    定义一个经典的“钟形曲线”正态分布，广泛适用于描述大量数据集的统计分布特征
    """
    # Mean：是分布的均值。Std._Dev.：是分布的标准差，必须大于0。
    return _generic_distribution_function_with_simulation("DriskNormal", Mean, Std_Dev, *args)
@xl_func("var Minimum, var Maximum, var 属性1, var 属性2, var 属性3, var 属性4, var 属性5, var 属性6, var 属性7, var 属性8, var 属性9, var 属性10, var 属性11, var 属性12, var 属性13, var 属性14, var 属性15, var 属性16, var 属性17, var 属性18, var 属性19, var 属性20, var*: var", volatile=True)
def DriskUniform(Minimum, Maximum, *args):
    """
   定义一个均匀概率分布。在均匀分布的取值范围内，每个值出现的可能性都是相等的
    """
    # Minimum：最小值。Maximum：最大值。
    return _generic_distribution_function_with_simulation("DriskUniform", Minimum, Maximum, *args)
@xl_func("var h, var 属性1, var 属性2, var 属性3, var 属性4, var 属性5, var 属性6, var 属性7, var 属性8, var 属性9, var 属性10, var 属性11, var 属性12, var 属性13, var 属性14, var 属性15, var 属性16, var 属性17, var 属性18, var 属性19, var 属性20, var*: var", volatile=True)
def DriskErf(h, *args):
    """定义一个具有方差参数H的Gauss Error函数，该分布派生自Normal分布"""
    return _generic_distribution_function_with_simulation("DriskErf", h, *args)
@xl_func("var a, var b, var 属性1, var 属性2, var 属性3, var 属性4, var 属性5, var 属性6, var 属性7, var 属性8, var 属性9, var 属性10, var 属性11, var 属性12, var 属性13, var 属性14, var 属性15, var 属性16, var 属性17, var 属性18, var 属性19, var 属性20, var*: var", volatile=True)
def DriskExtvalue(a, b, *args):
    """定义一个Extvalue分布"""
    return _generic_distribution_function_with_simulation("DriskExtvalue", a, b, *args)
@xl_func("var a, var b, var 属性1, var 属性2, var 属性3, var 属性4, var 属性5, var 属性6, var 属性7, var 属性8, var 属性9, var 属性10, var 属性11, var 属性12, var 属性13, var 属性14, var 属性15, var 属性16, var 属性17, var 属性18, var 属性19, var 属性20,	var*:	var", volatile=True)
def DriskExtvalueMin(a, b, *args):
    """定义一个ExtvalueMin分布"""
    return _generic_distribution_function_with_simulation("DriskExtvalueMin", a, b, *args)
@xl_func("var y, var beta, var alpha, var 属性1, var 属性2, var 属性3, var 属性4, var 属性5, var 属性6, var 属性7, var 属性8, var 属性9, var 属 性10, var 属 性11, var 属 性12, var 属 性13, var 属 性14, var 属 性15, var 属 性16, var 属 性17, var 属 性18, var 属 性19, var 属 性20,	var*:	var", volatile=True)
def DriskFatigueLife(y, beta, alpha, *args):
    """定义一个FatigueLife分布"""
    return _generic_distribution_function_with_simulation("DriskFatigueLife", y, beta, alpha, *args)
@xl_func("var y, var beta, var alpha, var 属性1, var 属性2, var 属性3, var 属性4, var 属性5, var 属性6, var 属性7, var 属性8, var 属性9, var 属性10, var 属性11, var 属性12, var 属性13, var 属性14, var 属性15, var 属性16, var 属性17, var 属性18, var 属性19, var 属性20,	var*:	var", volatile=True)
def DriskFrechet(y, beta, alpha, *args):
    """定义一个Frechet分布"""
    return _generic_distribution_function_with_simulation("DriskFrechet", y, beta, alpha, *args)
    
@xl_func("var gamma, var beta, var 属性1, var 属性2, var 属性3, var 属性4, var 属性5, var 属性6, var 属性7, var 属性8, var 属性9, var 属性10, var 属性11, var 属性12, var 属性13, var 属性14, var 属性15, var 属性16, var 属性17, var 属性18, var 属性19, var 属性20, var*: var", volatile=True)

def DriskHypSecant(gamma, beta, *args):
    """定义一个HypSecant分布"""
    return _generic_distribution_function_with_simulation("DriskHypSecant", gamma, beta, *args)
    # 双曲正割分布：无界对称，峰态更尖，常用于收益率等重峰场景建模。
   
@xl_func("var alpha1, var alpha2, var gamma, var beta, var 属性1, var 属性2, var 属性3, var 属性4, var 属性5, var 属性6, var 属性7, var 属性8, var 属性9, var 属性10, var 属性11, var 属性12, var 属性13, var 属性14, var 属性15, var 属性16, var 属性17, var 属性18, var 属性19, var 属性20, var*: var", volatile=True)
def DriskJohnsonSU(alpha1, alpha2, gamma, beta, *args):
    """定义一个JohnsonSU（系统无界）分布"""
    # JohnsonSU分布：连续、无界，适合专家意见建模和成本分析等场景。
    return _generic_distribution_function_with_simulation("DriskJohnsonSU", alpha1, alpha2, gamma, beta, *args)

@xl_func("var alpha1, var alpha2, var a, var b, var 属性1, var 属性2, var 属性3, var 属性4, var 属性5, var 属性6, var 属性7, var 属性8, var 属性9, var 属性10, var 属性11, var 属性12, var 属性13, var 属性14, var 属性15, var 属性16, var 属性17, var 属性18, var 属性19, var 属性20, var*: var", volatile=True)
def DriskJohnsonSB(alpha1, alpha2, a, b, *args):
    """定义一个JohnsonSB（系统有界）分布"""
    return _generic_distribution_function_with_simulation("DriskJohnsonSB", alpha1, alpha2, a, b, *args)
@xl_func("var alpha1, var alpha2, var min_val, var max_val, var 属性1, var 属性2, var 属性3, var 属性4, var 属性5, var 属性6, var 属性7, var 属性8, var 属性9, var 属性10, var 属性11, var 属性12, var 属性13, var 属性14, var 属性15, var 属性16, var 属性17, var 属性18, var 属性19, var 属性20, var*: var", volatile=True)
def DriskKumaraswamy(alpha1, alpha2, min_val, max_val, *args):
    """定义一个4参数Kumaraswamy分布"""
    return _generic_distribution_function_with_simulation("DriskKumaraswamy", alpha1, alpha2, min_val, max_val, *args)
@xl_func("var mu, var sigma, var 属性1, var 属性2, var 属性3, var 属性4, var 属性5, var 属性6, var 属性7, var 属性8, var 属性9, var 属性10, var 属性11, var 属性12, var 属性13, var 属性14, var 属性15, var 属性16, var 属性17, var 属性18, var 属性19, var 属性20, var*: var", volatile=True)
def DriskLaplace(mu, sigma, *args):
    """定义一个Laplace分布"""
    return _generic_distribution_function_with_simulation("DriskLaplace", mu, sigma, *args)

@xl_func("var a, var c, var 属性1, var 属性2, var 属性3, var 属性4, var 属性5, var 属性6, var 属性7, var 属性8, var 属性9, var 属性10, var 属性11, var 属性12, var 属性13, var 属性14, var 属性15, var 属性16, var 属性17, var 属性18, var 属性19, var 属性20, var*: var", volatile=True)
def DriskLevy(a, c, *args):
    """定义一个Levy分布"""
    # Levy分布，用于建模左端有界、右偏厚尾的连续变量。
    return _generic_distribution_function_with_simulation("DriskLevy", a, c, *args)

@xl_func("var alpha, var beta, var 属性1, var 属性2, var 属性3, var 属性4, var 属性5, var 属性6, var 属性7, var 属性8, var 属性9, var 属性10, var 属性11, var 属性12, var 属性13, var 属性14, var 属性15, var 属性16, var 属性17, var 属性18, var 属性19, var 属性20, var*: var", volatile=True)
def DriskLogistic(alpha, beta, *args):
    """定义一个Logistic分布"""
    return _generic_distribution_function_with_simulation("DriskLogistic", alpha, beta, *args)

@xl_func("var gamma, var beta, var alpha, var 属性1, var 属性2, var 属性3, var 属性4, var 属性5, var 属性6, var 属性7, var 属性8, var 属性9, var 属性10, var 属性11, var 属性12, var 属性13, var 属性14, var 属性15, var 属性16, var 属性17, var 属性18, var 属性19, var 属性20, var*: var", volatile=True)
def DriskLoglogistic(gamma, beta, alpha, *args):
    """定义一个Loglogistic分布"""
    return _generic_distribution_function_with_simulation("DriskLoglogistic", gamma, beta, alpha, *args)

@xl_func("var mu, var sigma, var 属性1, var 属性2, var 属性3, var 属性4, var 属性5, var 属性6, var 属性7, var 属性8, var 属性9, var 属性10, var 属性11, var 属性12, var 属性13, var 属性14, var 属性15, var 属性16, var 属性17, var 属性18, var 属性19, var 属性20, var*: var", volatile=True)
def DriskLognorm(mu, sigma, *args):
    """定义一种对数正态分布，这种形式的对数正态分布的参数为该概率分布的实际均值和标准差"""
    return _generic_distribution_function_with_simulation("DriskLognorm", mu, sigma, *args)
@xl_func("var mu, var sigma, var 属性1, var 属性2, var 属性3, var 属性4, var 属性5, var 属性6, var 属性7, var 属性8, var 属性9, var 属性10, var 属性11, var 属性12, var 属性13, var 属性14, var 属性15, var 属性16, var 属性17, var 属性18, var 属性19, var 属性20, var*: var", volatile=True)
def DriskLognorm2(mu, sigma, *args):
    """指定一种对数正态分布，其中输入的均值和标准差等于相应正态分布的均值和标准差"""
    return _generic_distribution_function_with_simulation("DriskLognorm2", mu, sigma, *args)
@xl_func("var theta, var alpha, var 属性1, var 属性2, var 属性3, var 属性4, var 属性5, var 属性6, var 属性7, var 属性8, var 属性9, var 属性10, var 属性11, var 属性12, var 属性13, var 属性14, var 属性15, var 属性16, var 属性17, var 属性18, var 属性19, var 属性20, var*: var", volatile=True)
def DriskPareto(theta, alpha, *args):
    """定义一个Pareto分布"""
    return _generic_distribution_function_with_simulation("DriskPareto", theta, alpha, *args)
@xl_func("var b, var q, var 属性1, var 属性2, var 属性3, var 属性4, var 属性5, var 属性6, var 属性7, var 属性8, var 属性9, var 属性10, var 属性11, var 属性12, var 属性13, var 属性14, var 属性15, var 属性16, var 属性17, var 属性18, var 属性19, var 属性20, var*: var",volatile=True)
def DriskPareto2(b, q, *args):
    """定义一个Pareto2分布"""
    return _generic_distribution_function_with_simulation("DriskPareto2", b, q, *args)
@xl_func(
    "var alpha, var beta, var 属性1, var 属性2, var 属性3, var 属性4, var 属性5, var 属性6, var 属性7, var 属性8, var 属性9, var 属性10, var 属性11, var 属性12, var 属性13, var 属性14, var 属性15, var 属性16, var 属性17, var 属性18, var 属性19, var 属性20, var*: var",volatile=True)
def DriskPearson5(alpha, beta, *args):
    """定义一个Pearson5分布"""
    return _generic_distribution_function_with_simulation("DriskPearson5", alpha, beta, *args)
@xl_func(
    "var alpha1, var alpha2, var beta, var 属性1, var 属性2, var 属性3, var 属性4, var 属性5, var 属性6, var 属性7, var 属性8, var 属性9, var 属性10, var 属性11, var 属性12, var 属性13, var 属性14, var 属性15, var 属性16, var 属性17, var 属性18, var 属性19, var 属性20, var*: var",volatile=True)
def DriskPearson6(alpha1, alpha2, beta, *args):
    """定义一个Pearson6分布"""
    return _generic_distribution_function_with_simulation("DriskPearson6", alpha1, alpha2, beta, *args)
@xl_func("var min_val, var m_likely, var max_val, var 属性1, var 属性2, var 属性3, var 属性4, var 属性5, var 属性6, var 属性7, var 属性8, var 属性9, var 属性10, var 属性11, var 属性12, var 属性13, var 属性14, var 属性15, var 属性16, var 属性17, var 属性18, var 属性19, var 属性20, var*: var",volatile=True)
def DriskPert(min_val, m_likely, max_val, *args):
    """定义一个Pert分布（也是Beta分布的一种特殊形式）"""
    return _generic_distribution_function_with_simulation("DriskPert", min_val, m_likely, max_val, *args)
@xl_func("var b, var 属性1, var 属性2, var 属性3, var 属性4, var 属性5, var 属性6, var 属性7, var 属性8, var 属性9, var 属性10, var 属性11, var 属性12, var 属性13, var 属性14, var 属性15, var 属性16, var 属性17, var 属性18, var 属性19, var 属性20, var*: var",volatile=True)
def DriskRayleigh(b, *args):
    """定义一个Rayleigh分布"""
    return _generic_distribution_function_with_simulation("DriskRayleigh", b, *args)
@xl_func("var min_val, var max_val, var 属性1, var 属性2, var 属性3, var 属性4, var 属性5, var 属性6, var 属性7, var 属性8, var 属性9, var 属性10, var 属性11, var 属性12, var 属性13, var 属性14, var 属性15, var 属性16, var 属性17, var 属性18, var 属性19, var 属性20, var*: var",volatile=True)
def DriskReciprocal(min_val, max_val, *args):
    """定义一个Reciprocal分布"""
    return _generic_distribution_function_with_simulation("DriskReciprocal", min_val, max_val, *args)
@xl_func("var alpha, var beta, var 属性1, var 属性2, var 属性3, var 属性4, var 属性5, var 属性6, var 属性7, var 属性8, var 属性9, var 属性10, var 属性11, var 属性12, var 属性13, var 属性14, var 属性15, var 属性16, var 属性17, var 属性18, var 属性19, var 属性20, var*: var",volatile=True)
def DriskWeibull(alpha, beta, *args):
    """定义Weibull分布，这是一种连续型概率分布，其形状与尺度特性会随参数取值发生显著变化"""
    return _generic_distribution_function_with_simulation("DriskWeibull", alpha, beta, *args)
@xl_func("var left_dist, var right_dist, var splice_point, var 属性1, var 属性2, var 属性3, var 属性4, var 属性5, var 属性6, var 属性7, var 属性8, var 属性9, var 属性10, var 属性11, var 属性12, var 属性13, var 属性14, var 属性15, var 属性16, var 属性17, var 属性18, var 属性19, var 属性20, var*: var", volatile=True)
def DriskSplice(left_dist, right_dist, splice_point, *args):
    """在 x 等于Splice点处将分布 #1 和分布 #2 拼接在一起"""
    return _splice_distribution_function_with_simulation(left_dist, right_dist, splice_point, *args)

@xl_func("var gamma, var beta, var 属性1, var 属性2, var 属性3, var 属性4, var 属性5, var 属性6, var 属性7, var 属性8, var 属性9, var 属性10, var 属性11, var 属性12, var 属性13, var 属性14, var 属性15, var 属性16, var 属性17, var 属性18, var 属性19, var 属性20, var*: var", volatile=True)
def DriskCauchy(gamma, beta, *args):
    """定义一个Cauchy分布"""
    return _generic_distribution_function_with_simulation("DriskCauchy", gamma, beta, *args)
@xl_func("var gamma, var beta, var alpha1, var alpha2, var 属性1, var 属性2, var 属性3, var 属性4, var 属性5, var 属性6, var 属性7, var 属性8, var 属性9, var 属性10, var 属性11, var 属性12, var 属性13, var 属性14, var 属性15, var 属性16, var 属性17, var 属性18, var 属性19, var 属性20, var*: var", volatile=True)
def DriskDagum(gamma, beta, alpha1, alpha2, *args):
    """定义一个Dagum分布"""
    return _generic_distribution_function_with_simulation("DriskDagum", gamma, beta, alpha1, alpha2, *args)
@xl_func("var min_val, var m_likely, var max_val, var lower_p, var 属性1, var 属性2, var 属性3, var 属性4, var 属性5, var 属性6, var 属性7, var 属性8, var 属于9, var 属于10, var 属于11, var 属于12, var 属于13, var 属于14, var 属于15, var 属于16, var 属于17,	var 属于18,	var 属于19,	var 属于20,	var*:	var", volatile=True)
def DriskDoubleTriang(min_val, m_likely, max_val, lower_p, *args):
    """定义一个由三个点及下三角分布概率权重构成的Double Triangular分布，其中最小值与最大值出现的概率为零"""
    return _generic_distribution_function_with_simulation("DriskDoubleTriang", min_val, m_likely, max_val, lower_p, *args)
@xl_func("var shape, var scale, var 属性1, var 属性2, var 属性3, var 属性4, var 属性5, var 属性6, var 属性7, var 属性8, var 属性9, var 属性10, var 属性11, var 属性12, var 属性13, var 属性14, var 属性15, var 属性16, var 属性17, var 属性18, var 属性19, var 属性20, var*: var", volatile=True)
def DriskGamma(shape, scale, *args):
    """定义一个Gamma分布。Gamma分布是一个连续分布。"""
    return _generic_distribution_function_with_simulation("DriskGamma", shape, scale, *args)
@xl_func("var m, var beta, var 属性1, var 属性2, var 属性3, var 属性4, var 属性5, var 属性6, var 属性7, var 属性8, var 属性9, var 属性10, var 属性11, var 属性12, var 属性13, var 属性14, var 属性15, var 属性16, var 属性17, var 属性18, var 属性19, var 属性20, var*: var", volatile=True)
def DriskErlang(m, beta, *args):
    """定义一个具有指定m和beta参数的m-erlang分布。"""
    return _generic_distribution_function_with_simulation("DriskErlang", m, beta, *args)
@xl_func("var lam, var 属性1, var 属性2, var 属性3, var 属性4, var 属性5, var 属性6, var 属性7, var 属性8, var 属性9, var 属性10, var 属性11, var 属性12, var 属性13, var 属性14, var 属性15, var 属性16, var 属性17, var 属性18, var 属性19, var 属性20, var*: var", volatile=True)
def DriskPoisson(lam, *args):
    """定义一个泊松分布，该离散型分布仅返回大于或等于零的整数值"""
    return _generic_distribution_function_with_simulation("DriskPoisson", lam, *args)
@xl_func("var a, var b, var 属性1, var 属性2, var 属性3, var 属性4, var 属性5, var 属性6, var 属性7, var 属性8, var 属性9, var 属性10, var 属性11, var 属性12, var 属性13, var 属性14, var 属性15, var 属性16, var 属性17, var 属性18, var 属性19, var 属性20, var*: var", volatile=True)
def DriskBeta(a, b, *args):
    """使用形状参数 alpha1 和 alpha2 来定义一个Beta分布"""
    return _generic_distribution_function_with_simulation("DriskBeta", a, b, *args)
@xl_func("var alpha1, var alpha2, var min_val, var max_val, var 属性1, var 属性2, var 属性3, var 属性4, var 属性5, var 属性6, var 属性7, var 属性8, var 属性9, var 属性10, var 属性11, var 属性12, var 属性13, var 属性14, var 属性15, var 属性16, var 属性17, var 属性18, var 属性19, var 属性20, var*: var",volatile=True)
def DriskBetaGeneral(alpha1, alpha2, min_val, max_val, *args):
    """\u5b9a\u4e49\u4e00\u4e2a\u5177\u6709\u81ea\u5b9a\u4e49\u6700\u5c0f\u503c\u548c\u6700\u5927\u503c\u7684Beta\u5206\u5e03"""
    return _generic_distribution_function_with_simulation("DriskBetaGeneral", alpha1, alpha2, min_val, max_val, *args)
@xl_func("var min_val, var m_likely, var mean_val, var max_val, var 属性1, var 属性2, var 属性3, var 属性4, var 属性5, var 属性6, var 属性7, var 属性8, var 属性9, var 属性10, var 属性11, var 属性12, var 属性13, var 属性14, var 属性15, var 属性16, var 属性17, var 属性18, var 属性19, var 属性20, var*: var",volatile=True)
def DriskBetaSubj(min_val, m_likely, mean_val, max_val, *args):
    """定义一个具有明确最小值和最大值的Beta分布，其形状参数通过设定的最可能值和均值计算得出"""
    return _generic_distribution_function_with_simulation("DriskBetaSubj", min_val, m_likely, mean_val, max_val, *args)
@xl_func("var gamma, var beta, var alpha1, var alpha2, var 属性1, var 属性2, var 属性3, var 属性4, var 属性5, var 属性6, var 属性7, var 属性8, var 属性9, var 属性10, var 属性11, var 属性12, var 属性13, var 属性14, var 属性15, var 属性16, var 属性17, var 属性18, var 属性19, var 属性20, var*: var", volatile=True)
def DriskBurr12(gamma, beta, alpha1, alpha2, *args):
    """定义一个Burr 12分布"""
    return _generic_distribution_function_with_simulation("DriskBurr12", gamma, beta, alpha1, alpha2, *args)
@xl_func("var frequency, var severity, var deductible, var limit, var 属性1, var 属性2, var 属性3, var 属性4, var 属性5, var 属性6, var 属性7, var 属性8, var 属性9, var 属性10, var 属性11, var 属性12, var 属性13, var 属性14, var 属性15, var 属性16, var 属性17, var 属性18, var 属性19, var 属性20, var*: var", volatile=True)
def DriskCompound(frequency, severity, deductible, limit, *args):
    """生成基于Severity分布的Frequency样本"""
    return _compound_distribution_function_with_simulation(frequency, severity, deductible, limit, *args)

@xl_func("var df, var 属性1, var 属性2, var 属性3, var 属性4, var 属性5, var 属性6, var 属性7, var 属性8, var 属性9, var 属性10, var 属性11, var 属性12, var 属性13, var 属性14, var 属性15, var 属性16, var 属性17, var 属性18, var 属性19, var 属性20, var*: var", volatile=True)
def DriskChiSq(df, *args):
    """定义一个自由度为V的卡方分布"""
    return _generic_distribution_function_with_simulation("DriskChiSq", df, *args)
@xl_func("var df1, var df2, var 属性1, var 属性2, var 属性3, var 属性4, var 属性5, var 属性6, var 属性7, var 属性8, var 属性9, var 属性10, var 属性11, var 属性12, var 属性13, var 属性14, var 属性15, var 属性16, var 属性17, var 属性18, var 属性19, var 属性20, var*: var", volatile=True)
def DriskF(df1, df2, *args):
    """定义一个F分布"""
    return _generic_distribution_function_with_simulation("DriskF", df1, df2, *args)
@xl_func("var df, var 属性1, var 属性2, var 属性3, var 属性4, var 属性5, var 属性6, var 属性7, var 属性8, var 属性9, var 属性10, var 属性11, var 属性12, var 属性13, var 属性14, var 属性15, var 属性16, var 属性17, var 属性18, var 属性19, var 属性20, var*: var", volatile=True)
def DriskStudent(df, *args):
    """定义一个Student分布"""
    return _generic_distribution_function_with_simulation("DriskStudent", df, *args)
@xl_func("var lam, var 属性1, var 属性2, var 属性3, var 属性4, var 属性5, var 属性6, var 属性7, var 属性8, var 属性9, var 属性10, var 属性11, var 属性12, var 属性13, var 属性14, var 属性15, var 属性16, var 属性17, var 属性18, var 属性19, var 属性20, var*: var", volatile=True)
def DriskExpon(lam, *args):
    """定义一个具有指定beta值的指数分布"""
    return _generic_distribution_function_with_simulation("DriskExpon", lam, *args)
@xl_func("var p, var 属性1, var 属性2, var 属性3, var 属性4, var 属性5, var 属性6, var 属性7, var 属性8, var 属性9, var 属性10, var 属性11, var 属性12, var 属性13, var 属性14, var 属性15, var 属性16, var 属性17, var 属性18, var 属性19, var 属性20, var*: var", volatile=True)
def DriskBernoulli(p, *args):
    """定义一个伯努利分布，其中每次试验的成功概率为p。该离散分布仅返回大于或等于零的整数值"""
    return _generic_distribution_function_with_simulation("DriskBernoulli", p, *args)
# ==================== 新增分布函数 ====================
@xl_func("var a, var c, var b, var 属性1, var 属性2, var 属性3, var 属性4, var 属性5, var 属性6, var 属性7, var 属性8, var 属性9, var 属性10, var 属性11, var 属性12, var 属性13, var 属性14, var 属性15, var 属性16, var 属性17, var 属性18, var 属性19, var 属性20, var*: var", volatile=True)
def DriskTriang(a, c, b, *args):
    """定义一个Triangular分布，其最小值与最大值处的发生概率为零"""
    return _generic_distribution_function_with_simulation("DriskTriang", a, c, b, *args)
@xl_func("var n, var p, var 属性1, var 属性2, var 属性3, var 属性4, var 属性5, var 属性6, var 属性7, var 属性8, var 属性9, var 属性10, var 属性11, var 属性12, var 属性13, var 属性14, var 属性15, var 属性16, var 属性17, var 属性18, var 属性19, var 属性20, var*: var", volatile=True)
def DriskBinomial(n, p, *args):
    """定义一个二项分布，其中试验次数为N，每次试验的成功概率为P。该离散分布仅返回大于或等于零的整数值"""
    return _generic_distribution_function_with_simulation("DriskBinomial", n, p, *args)
@xl_func("var s, var p, var 属性1, var 属性2, var 属性3, var 属性4, var 属性5, var 属性6, var 属性7, var 属性8, var 属性9, var 属性10, var 属性11, var 属性12, var 属性13, var 属性14, var 属性15, var 属性16, var 属性17, var 属性18, var 属性19, var 属性20, var*: var", volatile=True)
def DriskNegbin(s, p, *args):
    """定义一个负二项分布，该离散型分布仅返回大于或等于零的整数值"""
    return _generic_distribution_function_with_simulation("DriskNegbin", s, p, *args)
@xl_func("var L, var M, var U, var alpha, var beta, var 属性1, var 属性2, var 属性3, var 属性4, var 属性5, var 属性6, var 属性7, var 属性8, var 属性9, var 属性10, var 属性11, var 属性12, var 属性13, var 属性14, var 属性15, var 属性16, var 属性17, var 属性18, var 属性19, var 属性20, var*: var", volatile=True)
def DriskTrigen(L, M, U, alpha, beta, *args):
    """定义一个三角分布，该分布具有三个点，其中 一个位于最可能值，另外两个位于指定的底部和顶部百分位数处。"""
    return _generic_distribution_function_with_simulation("DriskTrigen", L, M, U, alpha, beta, *args)
@xl_func("var min_val, var max_val, var x_vals, var p_vals, var 属性1, var 属性2, var 属性3, var 属性4, var 属性5, var 属性6, var 属性7, var 属性8, var 属性9, var 属性10, var 属性11, var 属性12, var 属性13, var 属性14, var 属性15, var 属性16, var 属性17, var 属性18, var 属性19, var 属性20, var*: var", volatile=True)
def DriskCumul(min_val, max_val, x_vals, p_vals, *args):
    """定义一个由最小值和最大值确定范围、包含n个点的Cumul分布"""
    return _generic_distribution_function_with_simulation("DriskCumul", min_val, max_val, x_vals, p_vals, *args)
@xl_func("var min_val, var max_val, var x_vals, var p_vals, var 属性1, var 属性2, var 属性3, var 属性4, var 属性5, var 属性6, var 属性7, var 属性8, var 属性9, var 属性10, var 属性11, var 属性12, var 属性13, var 属性14, var 属性15, var 属性16, var 属性17, var 属性18, var 属性19, var 属性20, var*: var", volatile=True)
def DriskGeneral(min_val, max_val, x_vals, p_vals, *args):
    """基于指定的 (x, p) 数据对所构建的密度曲线，生成一个广义概率分布"""
    # 一般分布：由 Min/Max 与 X-Table/P-Table 定义的有界分段线性连续分布。
    return _generic_distribution_function_with_simulation("DriskGeneral", min_val, max_val, x_vals, p_vals, *args)
@xl_func("var min_val, var max_val, var p_vals, var 属性1, var 属性2, var 属性3, var 属性4, var 属性5, var 属性6, var 属性7, var 属性8, var 属性9, var 属性10, var 属性11, var 属性12, var 属性13, var 属性14, var 属性15, var 属性16, var 属性17, var 属性18, var 属性19, var 属性20, var*: var", volatile=True)
def DriskHistogrm(min_val, max_val, p_vals, *args):
    """定义一个用户自定义的直方图分布，该分布在最小值(minimum)和最大值(maximum)之间包含若干等宽区间，且每个区间具有相应的概率权重p"""
    return _generic_distribution_function_with_simulation("DriskHistogrm", min_val, max_val, p_vals, *args)

@xl_func("var x_vals, var p_vals, var 属性1, var 属性2, var 属性3, var 属性4, var 属性5, var 属性6, var 属性7, var 属性8, var 属性9, var 属性10, var 属性11, var 属性12, var 属性13, var 属性14, var 属性15, var 属性16, var 属性17, var 属性18, var 属性19, var 属性20, var*: var", volatile=True)
def DriskDiscrete(x_vals, p_vals, *args):
    """定义一个具有指定结果数量的离散分布，可输入任意数量的可能结果"""
    return _generic_distribution_function_with_simulation("DriskDiscrete", x_vals, p_vals, *args)
@xl_func("var mu, var lam, var 属性1, var 属性2, var 属性3, var 属性4, var 属性5, var 属性6, var 属性7, var 属性8, var 属性9, var 属性10, var 属性11, var 属性12, var 属于13, var 属于14, var 属于15, var 属于16, var 属于17, var 属于18, var 属于19, var 属于20,	var*:	var", volatile=True)
def DriskInvgauss(mu, lam, *args):
    """\u5b9a\u4e49\u4e00\u4e2aInvgauss\u5206\u5e03"""
    return _generic_distribution_function_with_simulation("DriskInvgauss", mu, lam, *args)
@xl_func("var p, var 属性1, var 属性2, var 属性3, var 属性4, var 属性5, var 属性6, var 属性7, var 属性8, var 属性9, var 属性10, var 属性11, var 属性12, var 属性13, var 属性14, var 属性15, var 属性16, var 属性17, var 属性18, var 属性19, var 属性20, var*: var", volatile=True)
def DriskGeomet(p, *args):
    """定义一个几何分布，其返回值表示在一系列独立试验中首次成功前所经历的失败次数"""
    return _generic_distribution_function_with_simulation("DriskGeomet", p, *args)
@xl_func("var x_vals, var 属性1, var 属性2, var 属性3, var 属性4, var 属性5, var 属性6, var 属性7, var 属性8, var 属性9, var 属性10, var 属性11, var 属性12, var 属性13, var 属性14, var 属性15, var 属性16, var 属性17, var 属性18, var 属性19, var 属性20, var*: var", volatile=True)
def DriskDUniform(x_vals, *args):
    """定义一个离散均匀分布，该分布具有任意数量的可能结果，且每个结果发生的概率相等"""
    return _generic_distribution_function_with_simulation("DriskDUniform", x_vals, *args)
# ==================== 模拟状态管理函数（供模拟引擎调用） ====================
def initialize_simulation_state(total_iterations: int):
    _simulation_state.initialize_simulation(total_iterations)
def set_simulation_iteration(iteration: int):
    _simulation_state.set_current_iteration(iteration)
def get_simulation_iteration() -> int:
    return _simulation_state.get_current_iteration()
def is_simulation_running() -> bool:
    return _simulation_state.is_simulation_running()
def get_distribution_cache(stable_key: str) -> Optional[np.ndarray]:
    return _simulation_state.get_distribution_cache(stable_key)
def clear_simulation_state():
    _simulation_state.clear_simulation()
def register_distribution_for_simulation(stable_key: str, func_info: Dict[str, Any]):
    _simulation_state.register_distribution(stable_key, func_info)
def find_stable_key_by_params(func_name: str, dist_params: List[float], is_nested: bool = False, depth: int = 0, parent_key: str = None, args_text: str = None) -> Optional[str]:
    return _simulation_state.find_stable_key_by_params(func_name, dist_params, is_nested, depth, parent_key, args_text)
def record_cell_value_for_simulation(cell_addr: str, value: float):
    return _simulation_state.record_cell_value(cell_addr, value)
def get_cell_value_from_simulation(cell_addr: str) -> Optional[float]:
    return _simulation_state.get_cell_value(cell_addr)
def add_nested_relationship(parent_key: str, child_key: str):
    _simulation_state.add_nested_relationship(parent_key, child_key)
def get_nested_children(parent_key: str) -> List[str]:
    return _simulation_state.get_nested_children(parent_key)
def get_nested_parent(child_key: str) -> Optional[str]:
    return _simulation_state.get_nested_parent(child_key)
def log_debug(message: str):
    _simulation_state.log_debug(message)


@xl_func("var n, var D, var M, var 属性1, var 属性2, var 属性3, var 属性4, var 属性5, var 属性6, var 属性7, var 属性8, var 属性9, var 属性10, var*: var", volatile=True)
def DriskHypergeo(n, D, M, *args):
    """定义一个超几何分布，该离散型分布仅返回非负整数值"""
    return _generic_distribution_function_with_simulation("DriskHypergeo", n, D, M, *args)

@xl_func("var min_val, var max_val, var 属性1, var 属性2, var 属性3, var 属性4, var 属性5, var 属性6, var 属性7, var 属性8, var 属性9, var 属性10, var*: var", volatile=True)
def DriskIntuniform(min_val, max_val, *args):
    """定义一个返回最小值和最大值范围内整数的等概率分布"""
    return _generic_distribution_function_with_simulation("DriskIntuniform", min_val, max_val, *args)
 
