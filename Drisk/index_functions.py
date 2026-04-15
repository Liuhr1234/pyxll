# index_functions.py 
"""索引函数模块 - 修复依赖解析与键映射问题，支持内嵌分布，属性完整，每次模拟前删除旧隐藏表"""
import numpy as np
import math
import time
import datetime
import threading
import re
import logging
import traceback
import sys
import os
import zlib  # 用于计算稳定的哈希值
from typing import Dict, List, Tuple, Optional, Any, Set, Union
from pyxll import xl_func, xl_app, xl_macro, xlcAlert
from constants import DISTRIBUTION_FUNCTION_NAMES, get_distribution_info, DIST_TYPE_TO_FUNC_NAME
# 导入向量化生成器（包括新分布）
from dist_bernoulli import bernoulli_generator_vectorized
from dist_triang import triang_generator_vectorized
from dist_binomial import binomial_generator_vectorized
from dist_erf import erf_generator_vectorized, erf_ppf
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
from dist_pearson5 import pearson5_generator_vectorized, pearson5_ppf, pearson5_raw_mean
from dist_pearson6 import pearson6_generator_vectorized, pearson6_ppf, pearson6_raw_mean
from dist_pareto2 import pareto2_generator_vectorized, pareto2_ppf, pareto2_raw_mean
from dist_pareto import pareto_generator_vectorized, pareto_ppf, pareto_cdf, pareto_raw_mean
from dist_levy import levy_generator_vectorized, levy_ppf, levy_cdf, levy_raw_mean
from dist_erlang import erlang_generator_vectorized, erlang_ppf
from dist_cauchy import cauchy_generator_vectorized, cauchy_ppf
from dist_dagum import dagum_generator_vectorized, dagum_ppf, dagum_cdf, dagum_raw_mean
from dist_doubletriang import doubletriang_generator_vectorized, doubletriang_ppf, doubletriang_raw_mean
from dist_negbin import negbin_generator_vectorized, negbin_ppf
from dist_invgauss import invgauss_generator_vectorized, invgauss_ppf
from dist_duniform import duniform_generator_vectorized, duniform_ppf
from dist_geomet import geomet_generator_vectorized, geomet_ppf
from dist_hypergeo import hypergeo_generator_vectorized, hypergeo_ppf
from dist_intuniform import intuniform_generator_vectorized, intuniform_ppf
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
# 添加当前目录到Python路径
current_dir = os.path.dirname(os.path.abspath(__file__))
if current_dir not in sys.path:
    sys.path.insert(0, current_dir)
# 设置日志
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s : %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
)
logger = logging.getLogger(__name__)
# 特殊错误标记
ERROR_MARKER = "#ERROR!"
_LOCK_ROUND_INT_DIST_TYPES = {
    "bernoulli",
    "binomial",
    "geomet",
    "hypergeo",
    "intuniform",
    "negbin",
}
# 尝试导入 constants 以使用 support 函数
try:
    from constants import get_distribution_support, DIST_TYPE_TO_FUNC_NAME
    CONSTANTS_AVAILABLE = True
except ImportError:
    CONSTANTS_AVAILABLE = False
    logger.warning("constants 模块导入失败，将使用硬编码的分布支持范围")
# ==================== 公式解析器导入 ====================
try:
    from formula_parser import (
        is_distribution_function, is_makeinput_function, is_simtable_function,
        is_output_cell, has_static_attribute,
        parse_formula_references, extract_all_distribution_functions_with_index,
        extract_simtable_functions, extract_makeinput_functions,
        extract_output_info, extract_input_attributes, extract_makeinput_attributes,
        remove_output_function_from_formula, remove_makeinput_function_from_formula,
        parse_complete_formula, extract_all_attributes_from_formula,
        extract_nested_distributions_advanced, extract_dist_params_and_markers,
        parse_marker_function,
        extract_distribution_params_from_formula,
        parse_args_with_nested_functions,
        DISTRIBUTION_FUNCTIONS, ATTRIBUTE_FUNCTIONS
    )
except ImportError as e:
    logger.error(f"导入formula_parser失败: {e}")
    DISTRIBUTION_FUNCTIONS = DISTRIBUTION_FUNCTION_NAMES
    ATTRIBUTE_FUNCTIONS = [
        "DriskName", "DriskLoc", "DriskCategory", "DriskCollect", "DriskConvergence",
        "DriskCopula", "DriskCorrmat", "DriskFit", "DriskIsDate", "DriskIsDiscrete",
        "DriskLock", "DriskSeed", "DriskShift", "DriskStatic", "DriskTruncate",
        "DriskTruncateP", "DriskTruncate2", "DriskTruncateP2", "DriskUnits", "DriskMakeInput"
    ]
# 导入依赖追踪模块
def _filter_compound_nested_functions(dist_funcs: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    if not dist_funcs:
        return dist_funcs

    compound_or_splice_indices = {
        func.get('index')
        for func in dist_funcs
        if str(func.get('func_name', '')).lower() in {'driskcompound', 'drisksplice'}
    }
    if not compound_or_splice_indices:
        return dist_funcs

    filtered = []
    for func in dist_funcs:
        parent_indices = func.get('parent_indices', []) or []
        if any(parent in compound_or_splice_indices for parent in parent_indices):
            continue
        filtered.append(func)
    return filtered

try:
    from dependency_tracker import (
        find_all_simulation_cells_in_workbook,
        EnhancedDependencyAnalyzer,
        get_all_dependents_direct,
        is_formula_cell
    )
except ImportError as e:
    logger.error(f"导入dependency_tracker失败: {e}")
# 导入模拟管理器
try:
    from simulation_manager import clear_simulations, create_simulation, get_simulation, get_current_sim_id
except ImportError as e:
    logger.error(f"导入simulation_manager失败: {e}")
# 导入分布生成函数（用于理论均值和截断逻辑参考，但不直接使用）
try:
    from distribution_functions import DistributionGenerator
except ImportError as e:
    logger.error(f"导入DistributionGenerator失败: {e}")
    DistributionGenerator = None
# ==================== 状态栏进度显示 ====================
class StatusBarProgress:
    """在Excel状态栏显示进度信息，支持ESC键停止"""
    def __init__(self, app, n_iterations_per_scenario, n_scenarios):
        self.app = app
        self.n_iterations_per_scenario = n_iterations_per_scenario
        self.n_scenarios = n_scenarios
        self.total_iterations = n_iterations_per_scenario * n_scenarios
        self.completed_iterations = 0
        self.current_scenario_iteration = 0
        self.current_scenario_index = 0
        self.start_time = time.time()
        self.scenario_start_time = self.start_time
        self.cancelled = False
        self.esc_key_pressed = False
        self.last_update_time = 0
        self.last_esc_check_time = 0
        logger.info(f"状态栏进度显示器初始化完成: {n_iterations_per_scenario}次迭代, {n_scenarios}个场景")
    def start_new_scenario(self, scenario_index):
        """开始新场景"""
        self.current_scenario_index = scenario_index
        self.current_scenario_iteration = 0
        self.scenario_start_time = time.time()
        self.cancelled = False
        self.update_status_bar()
        logger.info(f"开始场景 {scenario_index+1}/{self.n_scenarios}")
    def check_esc_key(self):
        """检查ESC键是否被按下"""
        try:
            import win32api
            import win32con
            current_time = time.time()
            if current_time - self.last_esc_check_time < 0.05:
                return self.esc_key_pressed
            self.last_esc_check_time = current_time
            # 检查ESC键状态
            key_state = win32api.GetAsyncKeyState(win32con.VK_ESCAPE)
            # 如果ESC键被按下（高位为1）
            if key_state & 0x8000:
                logger.info("检测到ESC键被按下")
                self.esc_key_pressed = True
                self.cancelled = True
                return True
        except Exception as e:
            logger.error(f"检查ESC键失败: {str(e)}")
        return False
    def update(self, current_iteration_in_scenario):
        """更新进度"""
        # 检查ESC键
        if not self.cancelled:
            self.check_esc_key()
        # 如果已经取消，不再更新进度
        if self.cancelled:
            return True
        self.current_scenario_iteration = current_iteration_in_scenario
        current_time = time.time()
        if current_iteration_in_scenario % 10 == 0 or (current_time - self.last_update_time) > 0.5:
            self.update_status_bar()
            self.last_update_time = current_time
        # 检查是否应该取消
        if self.esc_key_pressed:
            self.cancelled = True
            logger.info("模拟因ESC键被取消")
            return True
        return False
    def update_status_bar(self):
        """更新Excel状态栏显示"""
        try:
            # 计算进度百分比
            if self.total_iterations > 0:
                total_completed = self.current_scenario_index * self.n_iterations_per_scenario + self.current_scenario_iteration
                total_percent = min(100.0, (total_completed / self.total_iterations) * 100)
            else:
                total_percent = 0
            # 计算场景进度百分比
            if self.n_iterations_per_scenario > 0:
                scenario_percent = min(100, int((self.current_scenario_iteration / self.n_iterations_per_scenario) * 100))
            else:
                scenario_percent = 0
            # 计算运行时间
            current_time = time.time()
            total_elapsed_time = current_time - self.start_time
            elapsed_hours = int(total_elapsed_time // 3600)
            elapsed_minutes = int((total_elapsed_time % 3600) // 60)
            elapsed_seconds = int(total_elapsed_time % 60)
            # 计算速度
            if total_elapsed_time > 0 and total_completed > 0:
                speed = total_completed / total_elapsed_time
            else:
                speed = 0
            # 构建状态栏消息
            if self.n_scenarios > 1:
                status_msg = f"DRISK索引模拟: 总进度 {total_percent:.1f}%, 场景 {self.current_scenario_index+1}/{self.n_scenarios}, "
                status_msg += f"迭代 {self.current_scenario_iteration:,}/{self.n_iterations_per_scenario:,} ({scenario_percent}%), "
                status_msg += f"时间 {elapsed_hours:02d}:{elapsed_minutes:02d}:{elapsed_seconds:02d}, "
                status_msg += f"速度 {speed:.1f} 迭代/秒"
            else:
                status_msg = f"DRISK索引模拟: 迭代 {self.current_scenario_iteration:,}/{self.n_iterations_per_scenario:,} ({scenario_percent}%), "
                status_msg += f"时间 {elapsed_hours:02d}:{elapsed_minutes:02d}:{elapsed_seconds:02d}, "
                status_msg += f"速度 {speed:.1f} 迭代/秒"
            # 添加ESC提示
            if not self.cancelled:
                status_msg += " | 按ESC键停止"
            else:
                status_msg += " | 正在停止..."
            # 更新Excel状态栏
            self.app.StatusBar = status_msg
        except Exception as e:
            logger.error(f"更新状态栏失败: {str(e)}")
    def clear(self):
        """清除状态栏"""
        try:
            self.app.StatusBar = False
            logger.info("状态栏已清除")
        except Exception as e:
            logger.error(f"清除状态栏失败: {str(e)}")
    def is_cancelled(self):
        """检查是否取消"""
        return self.cancelled
    def was_cancelled_by_esc(self):
        """检查是否通过ESC键取消"""
        return self.esc_key_pressed
# ==================== 改进的分布生成器（支持正确截断和平移，并增加支持范围检查） ====================
try:
    from scipy import stats
    SCIPY_AVAILABLE = True
except ImportError:
    SCIPY_AVAILABLE = False
    logger.warning("scipy未安装，截断分布将使用拒绝采样，可能较慢。建议安装scipy以获得更好的性能。")
class TruncatableDistributionGenerator:
    """改进的分布生成器，支持正确的截断和平移，并处理loc/lock标记，增加参数验证和错误传播，支持范围检查"""
    # 映射分布类型到对应的分布类（用于计算截断均值）
    DIST_CLASS_MAP = {
        'normal': NormalDistribution,
        'uniform': UniformDistribution,
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
    # ========== 修改点1：移除 global_time_seed 参数，改为 seed，由调用者直接传入 ==========
    def __init__(self, func_name: str, dist_type: str, params: List[float], markers: Dict[str, Any], input_key: str, seed: int = None):
        self.func_name = func_name
        self.dist_type = dist_type.lower()
        self.params = params
        self.markers = markers or {}
        self.input_key = input_key
        self.seed = seed  # 直接使用传入的种子，不再计算
        # 提取shift和truncate参数
        self.shift_amount = 0.0
        self.truncate_type = None
        self.truncate_lower = None
        self.truncate_upper = None
        self.truncate_lower_pct = None
        self.truncate_upper_pct = None
        self._extract_markers()
        self.x_vals_list = None
        self.p_vals_list = None
        self._array_parse_failed = False  # 新增：标记数组解析是否失败
        if self.dist_type in ['cumul', 'discrete', 'duniform', 'general', 'histogrm']:
            x_vals_str = markers.get('x_vals')
            p_vals_str = markers.get('p_vals')
            if self.dist_type == 'histogrm' and p_vals_str and not x_vals_str:
                try:
                    def parse_numbers(s):
                        s = s.strip()
                        if s.startswith('{') and s.endswith('}'):
                            s = s[1:-1]
                        parts = [p.strip() for p in s.split(',') if p.strip()]
                        return [float(p) for p in parts]
                    self.p_vals_list = parse_numbers(p_vals_str)
                    histogrm_parse_p_table(self.p_vals_list)
                    logger.debug(f"成功解析 Histogrm P-Table: {self.p_vals_list}")
                except Exception as e:
                    logger.error(f"解析 Histogrm P-Table 失败: {e}, p_vals='{p_vals_str}'")
                    self.x_vals_list = self.p_vals_list = None
                    self._array_parse_failed = True
            elif x_vals_str and p_vals_str:
                try:
                    # 增强解析：先去除可能的引号，然后按逗号分割，去除空格
                    def parse_numbers(s):
                        s = s.strip()
                        if s.startswith('{') and s.endswith('}'):
                            s = s[1:-1]
                        parts = [p.strip() for p in s.split(',') if p.strip()]
                        return [float(p) for p in parts]
                    self.x_vals_list = parse_numbers(x_vals_str)
                    self.p_vals_list = parse_numbers(p_vals_str)
                    # 对于 discrete，确保概率归一化
                    if self.dist_type == 'discrete':
                        total = sum(self.p_vals_list)
                        if total > 0:
                            self.p_vals_list = [p / total for p in self.p_vals_list]
                    # 对于 cumul，验证数组有效性
                    if self.dist_type == 'cumul':
                        cumul_parse_arrays(self.x_vals_list, self.p_vals_list)
                    elif self.dist_type == 'general':
                        general_parse_arrays(float(self.params[0]), float(self.params[1]), self.x_vals_list, self.p_vals_list)
                    elif self.dist_type == 'histogrm':
                        histogrm_parse_p_table(self.p_vals_list)
                    logger.debug(f"成功解析数组参数: {self.x_vals_list}, {self.p_vals_list}")
                except Exception as e:
                    logger.error(f"解析数组参数失败: {e}, x_vals='{x_vals_str}', p_vals='{p_vals_str}'")
                    self.x_vals_list = self.p_vals_list = None
                    self._array_parse_failed = True
            else:
                logger.warning(f"缺少数组参数: x_vals={x_vals_str}, p_vals={p_vals_str}")
                self._array_parse_failed = True
        # 获取分布的理论支持范围（如果constants可用，优先使用）
        if CONSTANTS_AVAILABLE:
            try:
                # 根据dist_type获取func_name
                if func_name and func_name in DIST_TYPE_TO_FUNC_NAME.values():
                    # 已经有func_name
                    pass
                else:
                    # 根据dist_type反向查找
                    for fname, dtype in DIST_TYPE_TO_FUNC_NAME.items():
                        if dtype == dist_type:
                            func_name = fname
                            break
                if func_name:
                    self.support_low, self.support_high = get_distribution_support(func_name, params)
                else:
                    # 后备：使用硬编码
                    self.support_low, self.support_high = self._get_support_hardcoded()
            except:
                self.support_low, self.support_high = self._get_support_hardcoded()
        else:
            self.support_low, self.support_high = self._get_support_hardcoded()
        # 如果 scipy 不可用且是百分位数截断，尝试用自定义 ppf 转换为数值截断
        if not SCIPY_AVAILABLE and self.truncate_type in ['truncatep', 'truncatep2']:
            if self._has_custom_ppf():
                logger.debug(f"scipy不可用，使用自定义ppf将百分位数截断转换为数值截断: {self.input_key}")
                # 转换概率边界为数值边界
                new_lower = None
                new_upper = None
                if self.truncate_lower is not None:
                    lower_q = max(0.0, min(1.0, self.truncate_lower))
                    lower_val = self._custom_ppf(lower_q)
                    if not np.isnan(lower_val):
                        new_lower = lower_val
                if self.truncate_upper is not None:
                    upper_q = max(0.0, min(1.0, self.truncate_upper))
                    upper_val = self._custom_ppf(upper_q)
                    if not np.isnan(upper_val):
                        new_upper = upper_val
                # 根据原类型设置新边界
                if self.truncate_type == 'truncatep':
                    # 先截断后平移：边界为原始值
                    self.truncate_lower = new_lower
                    self.truncate_upper = new_upper
                    self.truncate_type = 'truncate'
                else:  # truncatep2
                    # 先平移后截断：边界应为平移后的值
                    if new_lower is not None:
                        new_lower += self.shift_amount
                    if new_upper is not None:
                        new_upper += self.shift_amount
                    self.truncate_lower = new_lower
                    self.truncate_upper = new_upper
                    self.truncate_type = 'truncate2'
                logger.debug(f"转换后截断类型: {self.truncate_type}, 下界={self.truncate_lower}, 上界={self.truncate_upper}")
        # 截断有效性标志及边界调整
        self.truncate_invalid = False
        self.adjusted_truncate_lower = self.truncate_lower
        self.adjusted_truncate_upper = self.truncate_upper
        self._check_and_adjust_truncate_boundaries()
        # 用于逆变换的分布对象
        self.dist_obj = None
        if SCIPY_AVAILABLE:
            self.dist_obj = self._get_scipy_dist()
        # 对于Cumul和Discrete，保存数组字符串以便生成样本
        self.x_vals_str = markers.get('x_vals')
        self.p_vals_str = markers.get('p_vals')
    def _get_support_hardcoded(self) -> Tuple[float, float]:
        """硬编码的分布支持范围（后备）"""
        dist_type = self.dist_type
        params = self.params
        if dist_type == 'normal':
            return -np.inf, np.inf
        elif dist_type == 'uniform':
            return params[0], params[1]
        elif dist_type == 'erf':
            return -np.inf, np.inf
        elif dist_type == 'extvaluemin':
            return -np.inf, np.inf
        elif dist_type == 'extvalue':
            return -np.inf, np.inf
        elif dist_type == 'erlang':
            return 0.0, np.inf
        elif dist_type == 'cauchy':
            return -np.inf, np.inf
        elif dist_type == 'dagum':
            return float(params[0]), np.inf
        elif dist_type == 'fatiguelife':
            return float(params[0]), np.inf
        elif dist_type == 'frechet':
            return float(params[0]), np.inf
        elif dist_type == 'johnsonsb':
            return float(params[2]), float(params[3])
        elif dist_type == 'johnsonsu':
            return -np.inf, np.inf
        elif dist_type == 'kumaraswamy':
            return float(params[2]), float(params[3])
        elif dist_type == 'laplace':
            return -np.inf, np.inf
        elif dist_type == 'logistic':
            return -np.inf, np.inf
        elif dist_type == 'loglogistic':
            return float(params[0]), np.inf
        elif dist_type == 'lognorm':
            return 0.0, np.inf
        elif dist_type == 'lognorm2':
            return 0.0, np.inf
        elif dist_type == 'betageneral':
            return float(params[2]), float(params[3])
        elif dist_type == 'betasubj':
            return float(params[0]), float(params[3])
        elif dist_type == 'burr12':
            return float(params[0]), np.inf
        elif dist_type == 'compound':
            return 0.0, np.inf
        elif dist_type == 'splice':
            return -np.inf, np.inf
        elif dist_type == 'pert':
            return float(params[0]), float(params[2])
        elif dist_type == 'reciprocal':
            return float(params[0]), float(params[1])
        elif dist_type == 'rayleigh':
            return 0.0, np.inf
        elif dist_type == 'weibull':
            return 0.0, np.inf
        elif dist_type == 'pearson5':
            return 0.0, np.inf
        elif dist_type == 'pearson6':
            return 0.0, np.inf
        elif dist_type == 'pareto2':
            return 0.0, np.inf
        elif dist_type == 'pareto':
            return float(params[1]), np.inf
        elif dist_type == 'levy':
            return float(params[0]), np.inf
        elif dist_type == 'doubletriang':
            return float(params[0]), float(params[2])
        elif dist_type == 'gamma':
            return 0.0, np.inf
        elif dist_type == 'poisson':
            return 0.0, np.inf
        elif dist_type == 'beta':
            return 0.0, 1.0
        elif dist_type == 'chisq':
            return 0.0, np.inf
        elif dist_type == 'f':
            return 0.0, np.inf
        elif dist_type == 'student':
            return -np.inf, np.inf
        elif dist_type == 'expon':
            return 0.0, np.inf
        elif dist_type == 'bernoulli':
            return 0.0, 1.0
        elif dist_type == 'triang':
            return params[0], params[2]
        elif dist_type == 'binomial':
            return 0.0, params[0]
        elif dist_type == 'negbin':
            return 0.0, 0.0 if float(params[1]) >= 1.0 else np.inf
        elif dist_type == 'geomet':
            return 0.0, np.inf
        elif dist_type == 'hypergeo':
            n = int(float(params[0]))
            D = int(float(params[1]))
            M = int(float(params[2]))
            return float(max(0, n + D - M)), float(min(n, D))
        elif dist_type == 'intuniform':
            return float(int(float(params[0]))), float(int(float(params[1])))
        elif dist_type == 'trigen':
            return -np.inf, np.inf
        elif dist_type == 'general':
            return float(params[0]), float(params[1])
        elif dist_type == 'histogrm':
            return float(params[0]), float(params[1])
        elif dist_type == 'cumul':
            return -np.inf, np.inf
        elif dist_type == 'discrete':
            return -np.inf, np.inf
        else:
            return -np.inf, np.inf
    def _extract_markers(self):
        """从markers中提取shift和truncate参数，支持单边截断，并确保数值类型"""
        shift_val = self.markers.get('shift')
        if shift_val is not None:
            try:
                self.shift_amount = float(shift_val)
            except:
                self.shift_amount = 0.0
        truncate_types = ['truncate', 'truncate2', 'truncatep', 'truncatep2']
        for trunc_type in truncate_types:
            trunc_val = self.markers.get(trunc_type)
            if trunc_val is not None:
                self.truncate_type = trunc_type
                if isinstance(trunc_val, str):
                    trunc_val = trunc_val.strip()
                    if trunc_val.startswith('(') and trunc_val.endswith(')'):
                        trunc_val = trunc_val[1:-1]
                    parts = [p.strip() for p in trunc_val.split(',')]
                    trunc_lower = None
                    trunc_upper = None
                    if len(parts) >= 1 and parts[0]:
                        try:
                            trunc_lower = float(parts[0])
                        except:
                            logger.warning(f"无法转换截断下限 '{parts[0]}' 为浮点数")
                    if len(parts) >= 2 and parts[1]:
                        try:
                            trunc_upper = float(parts[1])
                        except:
                            logger.warning(f"无法转换截断上限 '{parts[1]}' 为浮点数")
                    if trunc_lower is None and trunc_upper is None:
                        try:
                            import re
                            cleaned = re.sub(r'[^\d\.,\-]', '', trunc_val)
                            parts = [p.strip() for p in cleaned.split(',')]
                            if len(parts) >= 1 and parts[0]:
                                trunc_lower = float(parts[0])
                            if len(parts) >= 2 and parts[1]:
                                trunc_upper = float(parts[1])
                        except:
                            pass
                    if trunc_type in ['truncatep', 'truncatep2']:
                        if trunc_lower is not None and (trunc_lower < 0.0 or trunc_lower > 1.0):
                            self.truncate_invalid = True
                        if trunc_upper is not None and (trunc_upper < 0.0 or trunc_upper > 1.0):
                            self.truncate_invalid = True
                        self.truncate_lower_pct = trunc_lower
                        self.truncate_upper_pct = trunc_upper
                        self.truncate_lower = trunc_lower
                        self.truncate_upper = trunc_upper
                    else:
                        self.truncate_lower = trunc_lower
                        self.truncate_upper = trunc_upper
                break
    def _has_invalid_intuniform_percentile_marker(self) -> bool:
        if self.dist_type != 'intuniform':
            return False
        for key in ['truncatep', 'truncatep2']:
            raw = self.markers.get(key)
            if raw is None:
                continue
            text = str(raw).strip()
            if text.startswith('(') and text.endswith(')'):
                text = text[1:-1]
            for part in [p.strip() for p in text.split(',')]:
                if not part:
                    continue
                try:
                    value = float(part)
                except Exception:
                    return True
                if value < 0.0 or value > 1.0:
                    return True
        return False
    def _check_and_adjust_truncate_boundaries(self):
        """
        检查截断参数的有效性，并根据支持范围调整边界。
        如果调整后的区间非空，则标记有效并存储调整后的边界；
        如果为空，则标记无效。
        """
        # 确保截断参数为数值类型，否则标记无效
        if self.truncate_lower is not None and not isinstance(self.truncate_lower, (int, float)):
            logger.warning(f"截断下限不是数值类型: {self.truncate_lower}")
            self.truncate_invalid = True
            return
        if self.truncate_upper is not None and not isinstance(self.truncate_upper, (int, float)):
            logger.warning(f"截断上限不是数值类型: {self.truncate_upper}")
            self.truncate_invalid = True
            return
        # 基本检查：下限 > 上限
        if self.truncate_lower is not None and self.truncate_upper is not None:
            if self.truncate_lower > self.truncate_upper:
                self.truncate_invalid = True
                logger.warning(f"截断参数无效，下限 {self.truncate_lower} 大于上限 {self.truncate_upper}，将返回错误标记")
                return
        if self._has_invalid_intuniform_percentile_marker():
            self.truncate_invalid = True
            logger.warning("Intuniform 的百分位截断参数必须介于0和1之间")
            return
        if self.truncate_type in ['truncatep', 'truncatep2']:
            if self.truncate_lower is not None and (self.truncate_lower < 0.0 or self.truncate_lower > 1.0):
                self.truncate_invalid = True
                logger.warning(f"百分位截断下界 {self.truncate_lower} 无效，必须介于0和1之间")
                return
            if self.truncate_upper is not None and (self.truncate_upper < 0.0 or self.truncate_upper > 1.0):
                self.truncate_invalid = True
                logger.warning(f"百分位截断上界 {self.truncate_upper} 无效，必须介于0和1之间")
                return
        # 只对值截断（truncate/truncate2）进行支持范围调整，百分位数截断不调整
        if self.truncate_type not in ['truncate', 'truncate2']:
            return
        # 对于 truncate2 类型（先平移后截断），需要先还原到原始尺度再与 support 取交集
        if self.truncate_type == 'truncate2':
            # 原始尺度上的边界 = 用户指定边界 - shift
            orig_lower = self.truncate_lower - self.shift_amount if self.truncate_lower is not None else None
            orig_upper = self.truncate_upper - self.shift_amount if self.truncate_upper is not None else None
        else:
            orig_lower = self.truncate_lower
            orig_upper = self.truncate_upper
        # 计算与支持范围的交集
        low_support, high_support = self.support_low, self.support_high
        # 调整下界：取 max(用户下界, 支持下界)
        if orig_lower is not None:
            if low_support == -np.inf:
                new_lower = orig_lower
            else:
                new_lower = max(orig_lower, low_support)
        else:
            new_lower = None
        # 调整上界：取 min(用户上界, 支持上界)
        if orig_upper is not None:
            if high_support == np.inf:
                new_upper = orig_upper
            else:
                new_upper = min(orig_upper, high_support)
        else:
            new_upper = None
        # ===== 新增：单边截断超出支持范围检查 =====
        # 如果下界大于支持上界（有限），无有效值
        if new_lower is not None and high_support != np.inf and new_lower > high_support:
            self.truncate_invalid = True
            logger.warning(f"截断下界 {new_lower} 大于分布支持上界 {high_support}，无有效值")
            return
        # 如果上界小于支持下界（有限），无有效值
        if new_upper is not None and low_support != -np.inf and new_upper < low_support:
            self.truncate_invalid = True
            logger.warning(f"截断上界 {new_upper} 小于分布支持下界 {low_support}，无有效值")
            return
        # 检查调整后的区间是否非空
        if new_lower is not None and new_upper is not None:
            if new_lower > new_upper:
                self.truncate_invalid = True
                logger.warning(f"截断范围与支持范围无交集，下限 {new_lower} 大于上限 {new_upper}，将返回错误标记")
                return
        # 对于 truncate2，将调整后的原始边界转换回平移后的边界
        if self.truncate_type == 'truncate2':
            self.adjusted_truncate_lower = new_lower + self.shift_amount if new_lower is not None else None
            self.adjusted_truncate_upper = new_upper + self.shift_amount if new_upper is not None else None
        else:
            self.adjusted_truncate_lower = new_lower
            self.adjusted_truncate_upper = new_upper
    def _get_scipy_dist(self):
        """获取scipy分布对象"""
        if self.dist_type == 'normal':
            return stats.norm(loc=self.params[0], scale=self.params[1])
        elif self.dist_type == 'uniform':
            return stats.uniform(loc=self.params[0], scale=self.params[1]-self.params[0])
        elif self.dist_type == 'erf':
            sigma = 1.0 / (math.sqrt(2.0) * self.params[0])
            return stats.norm(loc=0.0, scale=sigma)
        elif self.dist_type == 'extvaluemin':
            return stats.gumbel_l(loc=self.params[0], scale=self.params[1])
        elif self.dist_type == 'extvalue':
            return stats.gumbel_r(loc=self.params[0], scale=self.params[1])
        elif self.dist_type == 'fatiguelife':
            return stats.fatiguelife(self.params[2], loc=self.params[0], scale=self.params[1])
        elif self.dist_type == 'frechet':
            return stats.invweibull(self.params[2], loc=self.params[0], scale=self.params[1])
        elif self.dist_type == 'hypsecant':
            # Align SciPy's scale with Drisk's beta parameterization.
            return stats.hypsecant(loc=self.params[0], scale=self.params[1] * 2.0 / np.pi)
        elif self.dist_type == 'johnsonsb':
            return stats.johnsonsb(self.params[0], self.params[1], loc=self.params[2], scale=self.params[3] - self.params[2])
        elif self.dist_type == 'johnsonsu':
            return stats.johnsonsu(self.params[0], self.params[1], loc=self.params[2], scale=self.params[3])
        elif self.dist_type == 'kumaraswamy':
            return None
        elif self.dist_type == 'laplace':
            return stats.laplace(loc=self.params[0], scale=self.params[1] / math.sqrt(2.0))
        elif self.dist_type == 'logistic':
            return stats.logistic(loc=self.params[0], scale=self.params[1])
        elif self.dist_type == 'loglogistic':
            return stats.fisk(self.params[2], loc=self.params[0], scale=self.params[1])
        elif self.dist_type == 'lognorm':
            mu = self.params[0]
            sigma = self.params[1]
            sigma_prime = math.sqrt(math.log(1.0 + (sigma / mu) ** 2))
            mu_prime = math.log((mu * mu) / math.sqrt(sigma * sigma + mu * mu))
            return stats.lognorm(s=sigma_prime, scale=math.exp(mu_prime), loc=0.0)
        elif self.dist_type == 'lognorm2':
            return stats.lognorm(s=self.params[1], scale=math.exp(self.params[0]), loc=0.0)
        elif self.dist_type == 'betageneral':
            return stats.beta(
                a=float(self.params[0]),
                b=float(self.params[1]),
                loc=float(self.params[2]),
                scale=float(self.params[3]) - float(self.params[2]),
            )
        elif self.dist_type == 'betasubj':
            from dist_betasubj import _solve_alpha_params
            alpha1, alpha2 = _solve_alpha_params(
                float(self.params[0]),
                float(self.params[1]),
                float(self.params[2]),
                float(self.params[3]),
            )
            return stats.beta(
                a=float(alpha1),
                b=float(alpha2),
                loc=float(self.params[0]),
                scale=float(self.params[3]) - float(self.params[0]),
            )
        elif self.dist_type == 'burr12':
            return stats.burr12(
                c=float(self.params[2]),
                d=float(self.params[3]),
                loc=float(self.params[0]),
                scale=float(self.params[1]),
            )
        elif self.dist_type == 'compound':
            return None
        elif self.dist_type == 'pert':
            min_val, m_likely, max_val = self.params[0], self.params[1], self.params[2]
            mean_val = (float(min_val) + 4.0 * float(m_likely) + float(max_val)) / 6.0
            alpha1 = 6.0 * (mean_val - float(min_val)) / (float(max_val) - float(min_val))
            alpha2 = 6.0 * (float(max_val) - mean_val) / (float(max_val) - float(min_val))
            return stats.beta(
                a=alpha1,
                b=alpha2,
                loc=float(min_val),
                scale=float(max_val) - float(min_val),
            )
        elif self.dist_type == 'reciprocal':
            return stats.reciprocal(a=float(self.params[0]), b=float(self.params[1]))
        elif self.dist_type == 'rayleigh':
            return stats.rayleigh(loc=0.0, scale=float(self.params[0]))
        elif self.dist_type == 'weibull':
            return stats.weibull_min(c=float(self.params[0]), loc=0.0, scale=float(self.params[1]))
        elif self.dist_type == 'pearson5':
            return stats.invgamma(a=self.params[0], scale=self.params[1], loc=0.0)
        elif self.dist_type == 'pearson6':
            return stats.betaprime(
                a=self.params[0],
                b=self.params[1],
                scale=self.params[2],
                loc=0.0,
            )
        elif self.dist_type == 'pareto2':
            return stats.lomax(c=self.params[1], loc=0.0, scale=self.params[0])
        elif self.dist_type == 'pareto':
            return stats.pareto(b=self.params[0], loc=0.0, scale=self.params[1])
        elif self.dist_type == 'levy':
            return stats.levy(loc=self.params[0], scale=self.params[1])
        elif self.dist_type == 'erlang':
            return stats.gamma(a=int(round(float(self.params[0]))), scale=self.params[1])
        elif self.dist_type == 'cauchy':
            return stats.cauchy(loc=self.params[0], scale=self.params[1])
        elif self.dist_type == 'dagum':
            return None
        elif self.dist_type == 'doubletriang':
            return None
        elif self.dist_type == 'gamma':
            return stats.gamma(a=self.params[0], scale=self.params[1])
        elif self.dist_type == 'poisson':
            return stats.poisson(mu=self.params[0])
        elif self.dist_type == 'beta':
            return stats.beta(a=self.params[0], b=self.params[1])
        elif self.dist_type == 'chisq':
            return stats.chi2(df=self.params[0])
        elif self.dist_type == 'f':
            return stats.f(dfn=self.params[0], dfd=self.params[1])
        elif self.dist_type == 'student':
            return stats.t(df=self.params[0])
        elif self.dist_type == 'expon':
            return stats.expon(scale=self.params[0])
        elif self.dist_type == 'bernoulli':
            return stats.bernoulli(p=self.params[0])
        elif self.dist_type == 'negbin':
            return stats.nbinom(int(self.params[0]), self.params[1])
        elif self.dist_type == 'hypergeo':
            return stats.hypergeom(M=int(self.params[2]), n=int(self.params[1]), N=int(self.params[0]))
        elif self.dist_type == 'intuniform':
            min_val = int(self.params[0])
            max_val = int(self.params[1])
            return stats.randint(low=min_val, high=max_val + 1)
        else:
            return None
    # 判断是否有自定义 PPF
    def _has_custom_ppf(self) -> bool:
        return self.dist_type in ['bernoulli', 'triang', 'binomial', 'erf', 'extvaluemin', 'extvalue', 'fatiguelife', 'frechet', 'general', 'histogrm', 'hypsecant', 'johnsonsb', 'johnsonsu', 'kumaraswamy', 'laplace', 'logistic', 'loglogistic', 'lognorm', 'lognorm2', 'betageneral', 'betasubj', 'burr12', 'compound', 'splice', 'pert', 'reciprocal', 'rayleigh', 'weibull', 'pearson5', 'pearson6', 'pareto2', 'pareto', 'levy', 'erlang', 'cauchy', 'dagum', 'doubletriang', 'negbin', 'invgauss', 'duniform', 'geomet', 'hypergeo', 'intuniform', 'trigen', 'cumul', 'discrete']
    def _custom_ppf(self, q: float) -> float:
        """自定义 PPF 函数"""
        if self.dist_type == 'bernoulli':
            return bernoulli_ppf(q, self.params[0])
        elif self.dist_type == 'triang':
            return triang_ppf(q, self.params[0], self.params[1], self.params[2])
        elif self.dist_type == 'binomial':
            n = int(self.params[0])
            return binomial_ppf(q, n, self.params[1])
        elif self.dist_type == 'erf':
            return erf_ppf(q, self.params[0])
        elif self.dist_type == 'extvaluemin':
            return extvaluemin_ppf(q, self.params[0], self.params[1])
        elif self.dist_type == 'extvalue':
            return extvalue_ppf(q, self.params[0], self.params[1])
        elif self.dist_type == 'fatiguelife':
            return fatiguelife_ppf(q, self.params[0], self.params[1], self.params[2])
        elif self.dist_type == 'frechet':
            return frechet_ppf(q, self.params[0], self.params[1], self.params[2])
        elif self.dist_type == 'hypsecant':
            return hypsecant_ppf(q, self.params[0], self.params[1])
        elif self.dist_type == 'johnsonsb':
            return johnsonsb_ppf(q, self.params[0], self.params[1], self.params[2], self.params[3])
        elif self.dist_type == 'johnsonsu':
            return johnsonsu_ppf(q, self.params[0], self.params[1], self.params[2], self.params[3])
        elif self.dist_type == 'kumaraswamy':
            return kumaraswamy_ppf(q, self.params[0], self.params[1], self.params[2], self.params[3])
        elif self.dist_type == 'laplace':
            return laplace_ppf(q, self.params[0], self.params[1])
        elif self.dist_type == 'logistic':
            return logistic_ppf(q, self.params[0], self.params[1])
        elif self.dist_type == 'loglogistic':
            return loglogistic_ppf(q, self.params[0], self.params[1], self.params[2])
        elif self.dist_type == 'lognorm':
            return lognorm_ppf(q, self.params[0], self.params[1])
        elif self.dist_type == 'lognorm2':
            return lognorm2_ppf(q, self.params[0], self.params[1])
        elif self.dist_type == 'betageneral':
            return betageneral_ppf(q, self.params[0], self.params[1], self.params[2], self.params[3])
        elif self.dist_type == 'betasubj':
            return betasubj_ppf(q, self.params[0], self.params[1], self.params[2], self.params[3])
        elif self.dist_type == 'burr12':
            return burr12_ppf(q, self.params[0], self.params[1], self.params[2], self.params[3])
        elif self.dist_type == 'compound':
            deductible = self.params[2] if len(self.params) >= 3 else 0.0
            limit = self.params[3] if len(self.params) >= 4 else float("inf")
            return compound_ppf(q, self.params[0], self.params[1], deductible, limit)
        elif self.dist_type == 'splice':
            return splice_ppf(q, self.params[0], self.params[1], self.params[2])
        elif self.dist_type == 'pert':
            return pert_ppf(q, self.params[0], self.params[1], self.params[2])
        elif self.dist_type == 'reciprocal':
            return reciprocal_ppf(q, self.params[0], self.params[1])
        elif self.dist_type == 'rayleigh':
            return rayleigh_ppf(q, self.params[0])
        elif self.dist_type == 'weibull':
            return weibull_ppf(q, self.params[0], self.params[1])
        elif self.dist_type == 'pearson5':
            return pearson5_ppf(q, self.params[0], self.params[1])
        elif self.dist_type == 'pearson6':
            return pearson6_ppf(q, self.params[0], self.params[1], self.params[2])
        elif self.dist_type == 'pareto2':
            return pareto2_ppf(q, self.params[0], self.params[1])
        elif self.dist_type == 'pareto':
            return pareto_ppf(q, self.params[0], self.params[1])
        elif self.dist_type == 'levy':
            return levy_ppf(q, self.params[0], self.params[1])
        elif self.dist_type == 'general':
            if self.x_vals_list is not None and self.p_vals_list is not None:
                return general_ppf(q, self.params[0], self.params[1], self.x_vals_list, self.p_vals_list)
            else:
                return np.nan
        elif self.dist_type == 'histogrm':
            if self.p_vals_list is not None:
                return histogrm_ppf(q, self.params[0], self.params[1], self.p_vals_list)
            else:
                return np.nan
        elif self.dist_type == 'erlang':
            return erlang_ppf(q, self.params[0], self.params[1])
        elif self.dist_type == 'cauchy':
            return cauchy_ppf(q, self.params[0], self.params[1])
        elif self.dist_type == 'dagum':
            return dagum_ppf(q, self.params[0], self.params[1], self.params[2], self.params[3])
        elif self.dist_type == 'doubletriang':
            return doubletriang_ppf(q, self.params[0], self.params[1], self.params[2], self.params[3])
        elif self.dist_type == 'negbin':
            return negbin_ppf(q, self.params[0], self.params[1])
        elif self.dist_type == 'invgauss':
            return invgauss_ppf(q, self.params[0], self.params[1])
        elif self.dist_type == 'geomet':
            return geomet_ppf(q, self.params[0])
        elif self.dist_type == 'hypergeo':
            return hypergeo_ppf(q, self.params[0], self.params[1], self.params[2])
        elif self.dist_type == 'intuniform':
            return intuniform_ppf(q, self.params[0], self.params[1])
        elif self.dist_type == 'duniform':
            if self.x_vals_list is not None:
                return duniform_ppf(q, self.x_vals_list)
            else:
                return np.nan
        elif self.dist_type == 'trigen':
            L, M, U, alpha, beta = self.params
            return trigen_ppf(q, L, M, U, alpha, beta)
        elif self.dist_type == 'cumul':
            if self.x_vals_list is not None and self.p_vals_list is not None:
                return cumul_ppf(q, self.x_vals_list, self.p_vals_list)
            else:
                return np.nan
        elif self.dist_type == 'discrete':
            if self.x_vals_list is not None and self.p_vals_list is not None:
                return discrete_ppf(q, self.x_vals_list, self.p_vals_list)
            else:
                return np.nan
        else:
            return np.nan
    def _generate_truncated_samples_custom_percentile(self, rng: np.random.Generator, n_samples: int) -> np.ndarray:
        """
        对带自定义 PPF 的分布，按概率区间直接做逆变换采样。
        这条路径和 distribution_functions 的百分位截断逻辑保持一致，
        避免把 truncatep / truncatep2 的概率边界误当成数值边界。
        """
        low_p = self.truncate_lower if self.truncate_lower is not None else 0.0
        high_p = self.truncate_upper if self.truncate_upper is not None else 1.0
        low_p = max(0.0, min(1.0, float(low_p)))
        high_p = max(0.0, min(1.0, float(high_p)))

        if low_p >= high_p:
            mid_p = (low_p + high_p) / 2.0
            value = self._custom_ppf(mid_p)
            return np.full(n_samples, value + self.shift_amount, dtype=float)

        probs = rng.uniform(low_p, high_p, n_samples)
        samples = np.array([self._custom_ppf(float(q)) for q in probs], dtype=float)
        return samples + self.shift_amount
    def _raw_mean(self) -> float:
        """返回原始分布的均值（不考虑截断和平移）"""
        dist_type = self.dist_type
        params = self.params
        if dist_type == 'normal':
            return params[0]
        elif dist_type == 'uniform':
            return (params[0] + params[1]) / 2
        elif dist_type == 'erf':
            return 0.0
        elif dist_type == 'extvaluemin':
            return extvaluemin_raw_mean(params[0], params[1])
        elif dist_type == 'extvalue':
            return extvalue_raw_mean(params[0], params[1])
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
            else:
                return 0.0
        elif dist_type == 'histogrm':
            if self.p_vals_list is not None:
                return histogrm_raw_mean(params[0], params[1], self.p_vals_list)
            else:
                return 0.0
        elif dist_type == 'erlang':
            return float(int(round(float(params[0])))) * float(params[1])
        elif dist_type == 'cauchy':
            return float('nan')
        elif dist_type == 'dagum':
            return dagum_raw_mean(params[0], params[1], params[2], params[3])
        elif dist_type == 'doubletriang':
            return doubletriang_raw_mean(params[0], params[1], params[2], params[3])
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
        elif dist_type == 'negbin':
            s = params[0]
            p = params[1]
            return s * (1.0 - p) / p
        elif dist_type == 'geomet':
            p = params[0]
            return (1.0 / p) - 1.0
        elif dist_type == 'hypergeo':
            return params[0] * params[1] / params[2]
        elif dist_type == 'intuniform':
            return (params[0] + params[1]) / 2.0
        elif dist_type == 'trigen':
            # 将 Trigen 参数转换为三角分布参数，再计算均值
            try:
                from dist_trigen import _convert_trigen_to_triang
                a, c, b = _convert_trigen_to_triang(params[0], params[1], params[2], params[3], params[4])
                return (a + c + b) / 3
            except Exception as e:
                logger.error(f"Trigen 均值计算失败: {e}")
                return (params[0] + params[1] + params[2]) / 3  # 回退近似
        elif dist_type == 'cumul':
            if self.x_vals_list is not None and self.p_vals_list is not None:
                # 加权平均
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
    def _truncated_normal_mean_manual(self, mu: float, sigma: float,
                                      lower: Optional[float], upper: Optional[float]) -> float:
        """手动计算截断正态均值（原始尺度）"""
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
    def _calculate_truncated_mean(self) -> float:
        """计算截断分布的理论均值（不考虑平移）"""
        if self.truncate_type is None:
            return self._raw_mean()
        # 确定原始尺度上的截断边界
        if self.truncate_type in ['truncate', 'truncatep']:
            # 先截断后平移：边界已经是原始值（已调整）
            orig_lower = self.adjusted_truncate_lower
            orig_upper = self.adjusted_truncate_upper
        else:  # 'truncate2', 'truncatep2'
            # 先平移后截断：边界需要减去 shift 得到原始值（使用调整后的边界）
            if self.truncate_type == 'truncate2':
                orig_lower = self.adjusted_truncate_lower - self.shift_amount if self.adjusted_truncate_lower is not None else None
                orig_upper = self.adjusted_truncate_upper - self.shift_amount if self.adjusted_truncate_upper is not None else None
            else:  # 'truncatep2' 边界已经是概率，不需要转换
                orig_lower = self.truncate_lower
                orig_upper = self.truncate_upper
        if SCIPY_AVAILABLE and self.dist_obj is not None:
            # 使用 scipy 的条件期望计算截断均值
            try:
                lb = -np.inf if orig_lower is None else orig_lower
                ub = np.inf if orig_upper is None else orig_upper
                # 对于百分位数截断，边界是概率值，需要转换为原始值
                if self.truncate_type in ['truncatep', 'truncatep2']:
                    # 如果边界是概率值，先用 ppf 转换为原始值
                    if orig_lower is not None:
                        lb = self.dist_obj.ppf(max(0.0, min(orig_lower, 1.0)))
                    else:
                        lb = -np.inf
                    if orig_upper is not None:
                        ub = self.dist_obj.ppf(max(0.0, min(orig_upper, 1.0)))
                    else:
                        ub = np.inf
                # 计算条件期望
                mean_trunc = self.dist_obj.expect(lambda x: x, lb=lb, ub=ub, conditional=True)
                return float(mean_trunc)
            except Exception as e:
                logger.debug(f"scipy expect 计算截断均值失败: {e}，回退到原始均值")
                return self._raw_mean()
        else:
            # scipy 不可用，只对正态分布有手动公式
            if self.dist_type == 'normal':
                mu = self.params[0]
                sigma = self.params[1]
                # 此时 orig_lower/orig_upper 已经是原始值
                return self._truncated_normal_mean_manual(mu, sigma, orig_lower, orig_upper)
            else:
                # 其他分布无法计算截断均值，返回原始均值
                return self._raw_mean()
    def _validate_params(self) -> bool:
        """验证参数是否有效（例如正态分布标准差>0，整数参数检查）"""
        if self.dist_type == 'normal':
            if len(self.params) >= 2 and self.params[1] <= 0:
                return False
        elif self.dist_type == 'uniform':
            if len(self.params) >= 2 and self.params[1] <= self.params[0]:
                return False
        elif self.dist_type == 'erf':
            if len(self.params) >= 1 and self.params[0] <= 0:
                return False
        elif self.dist_type == 'extvaluemin':
            if len(self.params) >= 2 and self.params[1] <= 0:
                return False
        elif self.dist_type == 'extvalue':
            if len(self.params) >= 2 and self.params[1] <= 0:
                return False
        elif self.dist_type == 'fatiguelife':
            if len(self.params) >= 3 and (self.params[1] <= 0 or self.params[2] <= 0):
                return False
        elif self.dist_type == 'frechet':
            if len(self.params) >= 3 and (self.params[1] <= 0 or self.params[2] <= 0):
                return False
        elif self.dist_type == 'hypsecant':
            if len(self.params) >= 2 and self.params[1] <= 0:
                return False
        elif self.dist_type == 'johnsonsb':
            if len(self.params) >= 4 and (self.params[1] <= 0 or self.params[3] <= self.params[2]):
                return False
        elif self.dist_type == 'johnsonsu':
            if len(self.params) >= 4 and (self.params[1] <= 0 or self.params[3] <= 0):
                return False
        elif self.dist_type == 'kumaraswamy':
            if len(self.params) >= 4 and (
                self.params[0] <= 0
                or self.params[1] <= 0
                or self.params[3] <= self.params[2]
            ):
                return False
        elif self.dist_type == 'laplace':
            if len(self.params) >= 2 and self.params[1] <= 0:
                return False
        elif self.dist_type == 'logistic':
            if len(self.params) >= 2 and self.params[1] <= 0:
                return False
        elif self.dist_type == 'loglogistic':
            if len(self.params) >= 3 and (self.params[1] <= 0 or self.params[2] <= 0):
                return False
        elif self.dist_type == 'lognorm':
            if len(self.params) >= 2 and (self.params[0] <= 0 or self.params[1] <= 0):
                return False
        elif self.dist_type == 'lognorm2':
            if len(self.params) >= 2 and self.params[1] <= 0:
                return False
        elif self.dist_type == 'betageneral':
            if len(self.params) >= 4:
                if self.params[0] <= 0 or self.params[1] <= 0 or self.params[3] <= self.params[2]:
                    return False
            else:
                return False
        elif self.dist_type == 'betasubj':
            if len(self.params) >= 4:
                if (
                    self.params[3] <= self.params[0]
                    or self.params[1] <= self.params[0]
                    or self.params[1] >= self.params[3]
                    or self.params[2] <= self.params[0]
                    or self.params[2] >= self.params[3]
                ):
                    return False
            else:
                return False
        elif self.dist_type == 'burr12':
            if len(self.params) >= 4:
                if self.params[1] <= 0 or self.params[2] <= 0 or self.params[3] <= 0:
                    return False
            else:
                return False
        elif self.dist_type == 'compound':
            if len(self.params) < 2:
                return False
            if not isinstance(self.params[0], str) or not str(self.params[0]).strip():
                return False
            if not isinstance(self.params[1], str) or not str(self.params[1]).strip():
                return False
            try:
                deductible = float(self.params[2]) if len(self.params) >= 3 else 0.0
                limit = float(self.params[3]) if len(self.params) >= 4 else float("inf")
            except Exception:
                return False
            if deductible < 0 or limit < 0:
                return False
        elif self.dist_type == 'splice':
            if len(self.params) < 3:
                return False
            if not isinstance(self.params[0], str) or not str(self.params[0]).strip():
                return False
            if not isinstance(self.params[1], str) or not str(self.params[1]).strip():
                return False
            try:
                splice_point = float(self.params[2])
            except Exception:
                return False
            if not np.isfinite(splice_point):
                return False
        elif self.dist_type == 'pert':
            if len(self.params) >= 3:
                min_val, m_likely, max_val = self.params[0], self.params[1], self.params[2]
                if max_val <= min_val or m_likely < min_val or m_likely > max_val:
                    return False
            else:
                return False
        elif self.dist_type == 'reciprocal':
            if len(self.params) >= 2:
                if self.params[0] <= 0 or self.params[1] <= self.params[0]:
                    return False
            else:
                return False
        elif self.dist_type == 'rayleigh':
            if len(self.params) >= 1:
                if self.params[0] <= 0:
                    return False
            else:
                return False
        elif self.dist_type == 'weibull':
            if len(self.params) >= 2:
                if self.params[0] <= 0 or self.params[1] <= 0:
                    return False
            else:
                return False
        elif self.dist_type == 'pearson5':
            if len(self.params) >= 2 and (self.params[0] <= 0 or self.params[1] <= 0):
                return False
        elif self.dist_type == 'pearson6':
            if len(self.params) >= 3 and (
                self.params[0] <= 0 or self.params[1] <= 0 or self.params[2] <= 0
            ):
                return False
        elif self.dist_type == 'pareto':
            if len(self.params) >= 2 and (self.params[0] <= 0 or self.params[1] <= 0):
                return False
        elif self.dist_type == 'pareto2':
            if len(self.params) >= 2 and (self.params[0] <= 0 or self.params[1] <= 0):
                return False
        elif self.dist_type == 'levy':
            if len(self.params) >= 2 and self.params[1] <= 0:
                return False
        elif self.dist_type == 'general':
            if len(self.params) < 4:
                return False
            try:
                if self.x_vals_list is None or self.p_vals_list is None:
                    return False
                general_parse_arrays(float(self.params[0]), float(self.params[1]), self.x_vals_list, self.p_vals_list)
            except Exception:
                return False
        elif self.dist_type == 'histogrm':
            if len(self.params) < 3:
                return False
            try:
                if float(self.params[1]) <= float(self.params[0]):
                    return False
                if self.p_vals_list is None:
                    return False
                histogrm_parse_p_table(self.p_vals_list)
            except Exception:
                return False
        elif self.dist_type == 'erlang':
            if len(self.params) >= 2:
                m = float(self.params[0])
                beta = float(self.params[1])
                if m <= 0 or abs(m - round(m)) > 1e-9 or beta <= 0:
                    return False
        elif self.dist_type == 'cauchy':
            if len(self.params) >= 2 and self.params[1] <= 0:
                return False
        elif self.dist_type == 'dagum':
            if len(self.params) >= 4 and (self.params[1] <= 0 or self.params[2] <= 0 or self.params[3] <= 0):
                return False
        elif self.dist_type == 'doubletriang':
            if len(self.params) >= 4:
                if not (self.params[0] < self.params[1] < self.params[2]):
                    return False
                if self.params[3] < 0 or self.params[3] > 1:
                    return False
        elif self.dist_type == 'gamma':
            if len(self.params) >= 2 and (self.params[0] <= 0 or self.params[1] <= 0):
                return False
        elif self.dist_type == 'poisson':
            if len(self.params) >= 1 and self.params[0] < 0:
                return False
        elif self.dist_type == 'beta':
            if len(self.params) >= 2 and (self.params[0] <= 0 or self.params[1] <= 0):
                return False
        elif self.dist_type == 'chisq':
            if len(self.params) >= 1 and (self.params[0] <= 0 or not self.params[0].is_integer()):
                return False
        elif self.dist_type == 'f':
            if len(self.params) >= 2 and (self.params[0] <= 0 or self.params[1] <= 0 or not self.params[0].is_integer() or not self.params[1].is_integer()):
                return False
        elif self.dist_type == 'student':
            if len(self.params) >= 1 and (self.params[0] <= 0 or not self.params[0].is_integer()):
                return False
        elif self.dist_type == 'expon':
            if len(self.params) >= 1 and self.params[0] <= 0:
                return False
        elif self.dist_type == 'invgauss':
            if len(self.params) >= 2 and (self.params[0] <= 0 or self.params[1] <= 0):
                return False
        elif self.dist_type == 'bernoulli':
            if len(self.params) >= 1 and (self.params[0] < 0 or self.params[0] > 1):
                return False
        elif self.dist_type == 'triang':
            if len(self.params) >= 3 and not (self.params[0] <= self.params[1] <= self.params[2]):
                return False
        elif self.dist_type == 'binomial':
            if len(self.params) >= 2:
                n = self.params[0]
                p = self.params[1]
                if n <= 0 or not float(n).is_integer() or p <= 0 or p >= 1:
                    return False
        elif self.dist_type == 'negbin':
            if len(self.params) >= 2:
                s = self.params[0]
                p = self.params[1]
                if s <= 0 or abs(float(s) - round(float(s))) > 1e-9 or p <= 0 or p > 1:
                    return False
        elif self.dist_type == 'geomet':
            if len(self.params) >= 1 and (self.params[0] <= 0 or self.params[0] > 1):
                return False
        elif self.dist_type == 'hypergeo':
            if len(self.params) >= 3:
                n = float(self.params[0])
                D = float(self.params[1])
                M = float(self.params[2])
                if (
                    n <= 0 or not n.is_integer()
                    or D <= 0 or not D.is_integer()
                    or M <= 0 or not M.is_integer()
                    or n > M or D > M
                ):
                    return False
        elif self.dist_type == 'intuniform':
            if len(self.params) >= 2:
                min_val = float(self.params[0])
                max_val = float(self.params[1])
                if (
                    abs(min_val - round(min_val)) > 1e-9
                    or abs(max_val - round(max_val)) > 1e-9
                    or min_val >= max_val
                ):
                    return False
        elif self.dist_type == 'trigen':
            if len(self.params) >= 5:
                a = self.params[0]
                b = self.params[1]
                c = self.params[2]
                alpha = self.params[3]
                beta = self.params[4]
                if not (a <= b <= c or a <= c <= b):
                    return False
        return True
    def generate_samples(self, n_samples: int) -> np.ndarray:
        """生成n个样本，正确处理截断和平移，并验证参数有效性"""
        # 检查截断有效性
        if self.truncate_invalid:
            logger.warning(f"截断参数无效（下限 > 上限 或 截断范围完全超出分布支持），返回全错误标记")
            return np.full(n_samples, ERROR_MARKER, dtype=object)
        if not self._validate_params():
            logger.debug(f"参数验证失败: {self.dist_type}, params={self.params}")
            return np.full(n_samples, ERROR_MARKER, dtype=object)
        # 数组解析失败则返回错误标记
        if self._array_parse_failed:
            logger.warning(f"数组参数解析失败，返回错误标记: {self.input_key}")
            return np.full(n_samples, ERROR_MARKER, dtype=object)
        # 处理 loc/lock 标记
        if self.markers.get('loc') or self.markers.get('lock'):
            mean_val = self._calculate_loc_value()
            if mean_val == ERROR_MARKER:
                return np.full(n_samples, ERROR_MARKER, dtype=object)
            return np.full(n_samples, mean_val, dtype=float)
        # 创建随机数生成器（使用最终种子 self.seed）
        rng = np.random.default_rng(self.seed)
        if self.truncate_type is None:
            samples = self._generate_raw_samples(rng, n_samples)
            if self.shift_amount != 0:
                samples = samples + self.shift_amount
            return samples
        if self.truncate_type in ['truncatep', 'truncatep2'] and self._has_custom_ppf():
            return self._generate_truncated_samples_custom_percentile(rng, n_samples)
        # 检查是否为离散分布
        is_discrete = False
        if self.func_name:
            info = get_distribution_info(self.func_name)
            if info:
                is_discrete = info.get('is_discrete', False)
            else:
                logger.warning(f"未找到分布信息: {self.func_name}")
        else:
            logger.warning("func_name 为空")
        if SCIPY_AVAILABLE and self.dist_obj is not None and not is_discrete:
            return self._generate_truncated_samples_scipy(rng, n_samples)
        else:
            return self._generate_truncated_samples_rejection(rng, n_samples)
    def _generate_raw_samples(self, rng: np.random.Generator, n_samples: int) -> np.ndarray:
        """生成原始分布样本（无截断平移）"""
        if self.dist_type == 'normal':
            return rng.normal(self.params[0], self.params[1], n_samples)
        elif self.dist_type == 'uniform':
            return rng.uniform(self.params[0], self.params[1], n_samples)
        elif self.dist_type == 'erf':
            return erf_generator_vectorized(rng, self.params, n_samples)
        elif self.dist_type == 'extvaluemin':
            return extvaluemin_generator_vectorized(rng, self.params, n_samples)
        elif self.dist_type == 'extvalue':
            return extvalue_generator_vectorized(rng, self.params, n_samples)
        elif self.dist_type == 'fatiguelife':
            return fatiguelife_generator_vectorized(rng, self.params, n_samples)
        elif self.dist_type == 'frechet':
            return frechet_generator_vectorized(rng, self.params, n_samples)
        elif self.dist_type == 'hypsecant':
            return hypsecant_generator_vectorized(rng, self.params, n_samples)
        elif self.dist_type == 'johnsonsb':
            return johnsonsb_generator_vectorized(rng, self.params, n_samples)
        elif self.dist_type == 'johnsonsu':
            return johnsonsu_generator_vectorized(rng, self.params, n_samples)
        elif self.dist_type == 'kumaraswamy':
            return kumaraswamy_generator_vectorized(rng, self.params, n_samples)
        elif self.dist_type == 'laplace':
            return laplace_generator_vectorized(rng, self.params, n_samples)
        elif self.dist_type == 'logistic':
            return logistic_generator_vectorized(rng, self.params, n_samples)
        elif self.dist_type == 'loglogistic':
            return loglogistic_generator_vectorized(rng, self.params, n_samples)
        elif self.dist_type == 'lognorm':
            return lognorm_generator_vectorized(rng, self.params, n_samples)
        elif self.dist_type == 'lognorm2':
            return lognorm2_generator_vectorized(rng, self.params, n_samples)
        elif self.dist_type == 'betageneral':
            return betageneral_generator_vectorized(rng, self.params, n_samples)
        elif self.dist_type == 'betasubj':
            return betasubj_generator_vectorized(rng, self.params, n_samples)
        elif self.dist_type == 'burr12':
            return burr12_generator_vectorized(rng, self.params, n_samples)
        elif self.dist_type == 'compound':
            return compound_generator_vectorized(rng, self.params, n_samples)
        elif self.dist_type == 'splice':
            return splice_generator_vectorized(rng, self.params, n_samples)
        elif self.dist_type == 'pert':
            return pert_generator_vectorized(rng, self.params, n_samples)
        elif self.dist_type == 'reciprocal':
            return reciprocal_generator_vectorized(rng, self.params, n_samples)
        elif self.dist_type == 'rayleigh':
            return rayleigh_generator_vectorized(rng, self.params, n_samples)
        elif self.dist_type == 'weibull':
            return weibull_generator_vectorized(rng, self.params, n_samples)
        elif self.dist_type == 'pearson5':
            return pearson5_generator_vectorized(rng, self.params, n_samples)
        elif self.dist_type == 'pearson6':
            return pearson6_generator_vectorized(rng, self.params, n_samples)
        elif self.dist_type == 'pareto2':
            return pareto2_generator_vectorized(rng, self.params, n_samples)
        elif self.dist_type == 'pareto':
            return pareto_generator_vectorized(rng, self.params, n_samples)
        elif self.dist_type == 'levy':
            return levy_generator_vectorized(rng, self.params, n_samples)
        elif self.dist_type == 'erlang':
            return erlang_generator_vectorized(rng, self.params, n_samples)
        elif self.dist_type == 'cauchy':
            return cauchy_generator_vectorized(rng, self.params, n_samples)
        elif self.dist_type == 'dagum':
            return dagum_generator_vectorized(rng, self.params, n_samples)
        elif self.dist_type == 'doubletriang':
            return doubletriang_generator_vectorized(rng, self.params, n_samples)
        elif self.dist_type == 'gamma':
            return rng.gamma(self.params[0], self.params[1], n_samples)
        elif self.dist_type == 'poisson':
            return rng.poisson(self.params[0], n_samples).astype(float)
        elif self.dist_type == 'beta':
            return rng.beta(self.params[0], self.params[1], n_samples)
        elif self.dist_type == 'chisq':
            return rng.chisquare(self.params[0], n_samples)
        elif self.dist_type == 'f':
            return rng.f(self.params[0], self.params[1], n_samples)
        elif self.dist_type == 'student':
            return rng.standard_t(self.params[0], n_samples)
        elif self.dist_type == 'expon':
            return rng.exponential(self.params[0], n_samples)
        elif self.dist_type == 'bernoulli':
            p = self.params[0]
            return bernoulli_generator_vectorized(rng, [p], n_samples)
        elif self.dist_type == 'triang':
            a, c, b = self.params[0], self.params[1], self.params[2]
            return triang_generator_vectorized(rng, [a, c, b], n_samples)
        elif self.dist_type == 'binomial':
            n = int(self.params[0])
            p = self.params[1]
            return binomial_generator_vectorized(rng, [n, p], n_samples)
        elif self.dist_type == 'negbin':
            return negbin_generator_vectorized(rng, self.params, n_samples)
        elif self.dist_type == 'invgauss':
            return invgauss_generator_vectorized(rng, self.params, n_samples)
        elif self.dist_type == 'geomet':
            return geomet_generator_vectorized(rng, self.params, n_samples)
        elif self.dist_type == 'hypergeo':
            return hypergeo_generator_vectorized(rng, self.params, n_samples)
        elif self.dist_type == 'intuniform':
            return intuniform_generator_vectorized(rng, self.params, n_samples)
        elif self.dist_type == 'duniform':
            if self.x_vals_str is not None and self.x_vals_list is not None:
                return duniform_generator_vectorized(rng, self.params, n_samples, self.x_vals_list)
            else:
                return np.full(n_samples, ERROR_MARKER, dtype=object)
        elif self.dist_type == 'trigen':
            L, M, U, alpha, beta = self.params
            return trigen_generator_vectorized(rng, [L, M, U, alpha, beta], n_samples)
        elif self.dist_type == 'general':
            if self.x_vals_str is not None and self.p_vals_str is not None and self.x_vals_list is not None:
                return general_generator_vectorized(rng, self.params, n_samples, self.x_vals_list, self.p_vals_list)
            else:
                return np.full(n_samples, ERROR_MARKER, dtype=object)
        elif self.dist_type == 'histogrm':
            if self.p_vals_str is not None and self.p_vals_list is not None:
                return histogrm_generator_vectorized(rng, self.params, n_samples, self.p_vals_list)
            else:
                return np.full(n_samples, ERROR_MARKER, dtype=object)
        elif self.dist_type == 'cumul':
            if self.x_vals_str is not None and self.p_vals_str is not None and self.x_vals_list is not None:
                return cumul_generator_vectorized(rng, self.params, n_samples, self.x_vals_list, self.p_vals_list)
            else:
                return np.full(n_samples, ERROR_MARKER, dtype=object)
        elif self.dist_type == 'discrete':
            if self.x_vals_str is not None and self.p_vals_str is not None and self.x_vals_list is not None:
                return discrete_generator_vectorized(rng, self.params, n_samples, self.x_vals_list, self.p_vals_list)
            else:
                return np.full(n_samples, ERROR_MARKER, dtype=object)
        else:
            return np.full(n_samples, ERROR_MARKER, dtype=object)
    def _generate_truncated_samples_scipy(self, rng: np.random.Generator, n_samples: int) -> np.ndarray:
        """使用scipy逆变换生成截断样本，支持单边截断，正确区分平移顺序，使用调整后的边界"""
        # 确定用于CDF的原始边界
        if self.truncate_type in ['truncate', 'truncatep']:
            # 先截断后平移：边界已经是原始值（已调整）
            orig_lower = self.adjusted_truncate_lower
            orig_upper = self.adjusted_truncate_upper
        else:  # 'truncate2', 'truncatep2'
            if self.truncate_type == 'truncate2':
                # 先平移后截断：边界需要减去 shift 得到原始值（使用调整后的边界）
                orig_lower = self.adjusted_truncate_lower - self.shift_amount if self.adjusted_truncate_lower is not None else None
                orig_upper = self.adjusted_truncate_upper - self.shift_amount if self.adjusted_truncate_upper is not None else None
            else:  # 'truncatep2' 边界是概率值，不需要转换
                orig_lower = self.truncate_lower
                orig_upper = self.truncate_upper
        # 计算对应的CDF区间
        if self.truncate_type in ['truncate', 'truncate2']:
            # 值截断：用原始值计算CDF
            low_cdf = self.dist_obj.cdf(orig_lower) if orig_lower is not None else 0.0
            high_cdf = self.dist_obj.cdf(orig_upper) if orig_upper is not None else 1.0
        else:  # 'truncatep', 'truncatep2'
            # 百分位数截断：直接使用概率值
            low_cdf = orig_lower if orig_lower is not None else 0.0
            high_cdf = orig_upper if orig_upper is not None else 1.0
        # 确保边界有效
        low_cdf = max(0.0, min(low_cdf, 1.0))
        high_cdf = max(0.0, min(high_cdf, 1.0))
        if low_cdf >= high_cdf:
            mid = (low_cdf + high_cdf) / 2
            val = self.dist_obj.ppf(mid)
            return np.full(n_samples, val, dtype=float) + self.shift_amount
        # 生成均匀随机数并映射到截断区间
        u = rng.uniform(low_cdf, high_cdf, n_samples)
        samples = self.dist_obj.ppf(u)
        # 最后加上平移量（对所有类型都适用，因为生成的样本是原始尺度）
        return samples + self.shift_amount
    def _generate_truncated_samples_rejection(self, rng: np.random.Generator, n_samples: int) -> np.ndarray:
        """使用拒绝采样生成截断样本（后备方案），使用调整后的边界"""
        max_attempts_per_sample = 1000  # 增加尝试次数
        samples = np.zeros(n_samples, dtype=object)
        collected = 0
        attempts = 0
        max_total_attempts = n_samples * max_attempts_per_sample
        def accept(val):
            if self.truncate_type in ['truncate', 'truncatep']:
                # 先截断后平移：边界是原始值（已调整）
                if self.adjusted_truncate_lower is not None and val < self.adjusted_truncate_lower:
                    return False
                if self.adjusted_truncate_upper is not None and val > self.adjusted_truncate_upper:
                    return False
                return True
            else:
                # 先平移后截断：检查平移后的值
                shifted = val + self.shift_amount
                if self.adjusted_truncate_lower is not None and shifted < self.adjusted_truncate_lower:
                    return False
                if self.adjusted_truncate_upper is not None and shifted > self.adjusted_truncate_upper:
                    return False
                return True
        while collected < n_samples and attempts < max_total_attempts:
            batch_size = min(n_samples - collected, max(100, n_samples // 10))
            raw_batch = self._generate_raw_samples(rng, batch_size)
            for val in raw_batch:
                if accept(val):
                    if self.truncate_type in ['truncate', 'truncatep']:
                        samples[collected] = val + self.shift_amount
                    else:
                        samples[collected] = val + self.shift_amount
                    collected += 1
                    if collected >= n_samples:
                        break
                attempts += 1
        if collected < n_samples:
            logger.warning(f"拒绝采样未能在{max_total_attempts}次尝试内收集足够样本，将填充错误标记。")
            samples[collected:] = ERROR_MARKER
        return samples
    # ========== 精确计算 loc 值（考虑截断） ==========
    def _calculate_loc_value(self) -> float:
        """计算理论均值（用于 loc/lock 标记）- 统一使用分布类实现"""
        if self.truncate_invalid or self._array_parse_failed:
            return ERROR_MARKER
        # 使用分布类计算截断均值（精确）
        dist_class = self.DIST_CLASS_MAP.get(self.dist_type)
        if dist_class is not None:
            try:
                # 创建分布对象，传入所有标记（包括截断和shift）
                # 注意：分布类的 mean() 已经包含了 shift，我们减去 shift 得到原始均值
                dist_obj = dist_class(self.params, self.markers, self.func_name)
                if not dist_obj.is_valid():
                    return ERROR_MARKER
                raw_mean = dist_obj.mean() - dist_obj.shift_amount
                value = raw_mean + self.shift_amount
                if self.dist_type in _LOCK_ROUND_INT_DIST_TYPES:
                    value = float(round(value))
                return value
            except Exception as e:
                logger.warning(f"使用分布类 {dist_class.__name__} 计算截断均值失败: {e}，回退到原始均值")
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
# ==================== 索引模拟管理器 ====================
class IndexSimulationManager:
    """索引模拟管理器 - 使用隐藏工作表批量生成随机数，INDEX函数查表，支持依赖顺序生成和错误传播"""
    _instance = None
    _lock = threading.RLock()
    def __new__(cls):
        with cls._lock:
            if cls._instance is None:
                cls._instance = super().__new__(cls)
                cls._instance._init_manager()
            return cls._instance
    def _init_manager(self):
        """初始化管理器"""
        self.simulation_id = int(time.time() % 1000000)
        self.total_iterations = 0
        self.scenario_count = 1
        self.current_scenario = 0
        self.simulation_start_time = time.time()
        self.simulation_end_time = None
        self.simulation_name = ""
        self.is_running = False
        # 输入数据存储
        self.input_data = {}
        self.input_stats = {}
        # 输出数据存储
        self.output_data = {}
        self.output_stats = {}
        # MakeInput数据存储
        self.makeinput_data = {}
        # 单元格信息
        self.distribution_cells = {}
        self.simtable_cells = {}
        self.makeinput_cells = {}
        self.output_cells = {}
        # 隐藏表信息
        self.hidden_sheet = None
        self.input_columns = {}          # {input_key: column_letter}
        self.input_key_to_column = {}    # 新增：输入键到列字母的映射
        # 原始公式备份
        self.original_formulas = {}
        # Simtable原始值备份
        self.original_simtable_values = {}
        # Excel性能设置
        self.original_settings = {}
        # 状态栏进度
        self.progress = None
        # 控制行单元格
        self.control_cell = None
        # 缓存上次计算的工作表
        self.calculated_sheets_cache = set()
        # 分布生成器缓存
        self.distribution_generators = {}
        # 修复：记录已处理的输入键，避免重复
        self.processed_input_keys = set()
        # 修复：存储输入属性信息
        self.input_attributes_info = {}
        self.output_attributes_info = {}
        # 改进：单元格内分布函数计数映射
        self.cell_dist_count = {}  # {cell_addr: count}
        # 改进：跨sheet引用跟踪
        self.cross_sheet_refs = {}  # {cell_addr: [ref1, ref2, ...]}
        # 新增：输入键与分布函数信息的映射，用于依赖排序
        self.input_key_to_func_info = {}  # {input_key: func_info}
        # 新增：单元格地址到主输入键的映射（仅当单元格只有一个分布函数时）
        self.cell_to_input_key = {}  # {cell_addr: input_key}
        # 缓存 index 公式替换结果
        self.index_formula_cache = {}  # {(cell_addr, sheet_name, original_formula): new_formula}
        # ---------- 新增：输出引用批量读取相关 ----------
        self.output_refs = []          # [(cell_addr, type)] type为'output'或'makeinput'
        self.output_refs_range = None  # 隐藏表中存储引用公式的区域
        self.output_refs_start_row = 3
        self.output_refs_start_col = 2  # 初始值，实际在setup_output_references中会动态计算
        # ---------- 新增：依赖图相关 ----------
        self.dependency_graph = {}      # {input_key: [依赖的input_key列表]}
        self.sorted_keys = []           # 拓扑排序后的输入键列表
        # ==================== 新增：Simtable 错误处理相关 ====================  # NEW
        self.simtable_error_cells = []  # 存储当前场景中因索引超出而设置为#ERROR!的Simtable单元格地址
        self.skipped_scenarios = []     # 记录跳过的场景索引
        # ========== 新增：基础时间种子（用于无用户种子的输入） ==========
        self.base_time_seed = 0
    # ---------- 辅助方法：解析Cumul/Discrete数组参数 ----------
    def _split_cumul_discrete_args(self, args_text: str) -> List[str]:
        """
        自定义参数分割函数，正确处理花括号数组常量。
        例如输入: "{1,2,3},{0.2,0.3,0.5}" 返回 ["{1,2,3}", "{0.2,0.3,0.5}"]
        """
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
                    # 参数分隔符
                    arg = ''.join(current).strip()
                    if arg:
                        args.append(arg)
                    current = []
                else:
                    current.append(ch)
            else:
                current.append(ch)
        # 最后一个参数
        if current:
            arg = ''.join(current).strip()
            if arg:
                args.append(arg)
        return args
    def _read_range_values(self, app, range_str: str) -> List[float]:
        """
        从Excel区域读取数值，返回浮点数列表。
        支持跨工作表引用，如 'Sheet2!A1:A10'
        """
        try:
            # 处理可能的工作表名
            if '!' in range_str:
                sheet_name, addr = range_str.split('!', 1)
                sheet = app.ActiveWorkbook.Worksheets(sheet_name)
            else:
                sheet = app.ActiveSheet
                addr = range_str
            addr = addr.strip()
            if ':' not in addr:
                # 单个单元格
                cell = sheet.Range(addr)
                val = cell.Value2
                try:
                    return [float(val)]
                except (TypeError, ValueError):
                    raise ValueError(f"单元格 {range_str} 的值不是数字")
            else:
                # 区域
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
            logger.error(f"读取区域 {range_str} 失败: {str(e)}")
            raise
    # ========== 新增：专门解析 DriskCumul 的四个参数 ==========
    def _parse_cumul_4args(self, args_text: str, app) -> Tuple[Optional[str], Optional[str]]:
        """
        从 DriskCumul 的参数字符串中解析出完整的 x_vals 和 p_vals 字符串（已包含边界）。
        参数格式：min_val, max_val, x_param, p_param
        返回 (x_full_str, p_full_str)
        """
        if not args_text:
            return None, None
        # 分割参数
        args_list = self._split_cumul_discrete_args(args_text)
        if len(args_list) < 4:
            logger.error(f"Cumul 参数不足，需要4个，实际 {len(args_list)}: {args_text}")
            return None, None
        min_raw = args_list[0].strip().strip('"\'')
        max_raw = args_list[1].strip().strip('"\'')
        x_raw = args_list[2].strip().strip('"\'')
        p_raw = args_list[3].strip().strip('"\'')
        # 解析 min 和 max 为数值
        try:
            min_val = float(self._read_range_values(app, min_raw)[0] if re.match(r'[A-Z]+\d+', min_raw, re.I) else float(min_raw))
            max_val = float(self._read_range_values(app, max_raw)[0] if re.match(r'[A-Z]+\d+', max_raw, re.I) else float(max_raw))
        except Exception as e:
            logger.error(f"解析 min/max 失败: {e}")
            return None, None
        # 解析内部 X 数组
        x_inner = []
        if re.match(r'^[A-Z]{1,3}\d', x_raw, re.I) or ('!' in x_raw and re.match(r'.+![A-Z]{1,3}\d', x_raw, re.I)):
            try:
                x_inner = self._read_range_values(app, x_raw)
            except Exception as e:
                logger.error(f"读取 X 区域失败: {e}")
                return None, None
        else:
            if x_raw.startswith('{') and x_raw.endswith('}'):
                x_raw = x_raw[1:-1]
            parts = [p.strip() for p in x_raw.split(',') if p.strip()]
            try:
                x_inner = [float(p) for p in parts]
            except ValueError:
                logger.error(f"X 值解析失败: {x_raw}")
                return None, None
        # 解析内部 P 数组
        p_inner = []
        if re.match(r'^[A-Z]{1,3}\d', p_raw, re.I) or ('!' in p_raw and re.match(r'.+![A-Z]{1,3}\d', p_raw, re.I)):
            try:
                p_inner = self._read_range_values(app, p_raw)
            except Exception as e:
                logger.error(f"读取 P 区域失败: {e}")
                return None, None
        else:
            if p_raw.startswith('{') and p_raw.endswith('}'):
                p_raw = p_raw[1:-1]
            parts = [p.strip() for p in p_raw.split(',') if p.strip()]
            try:
                p_inner = [float(p) for p in parts]
            except ValueError:
                logger.error(f"P 值解析失败: {p_raw}")
                return None, None
        if len(x_inner) != len(p_inner):
            logger.error(f"X 和 P 数组长度不相等: {len(x_inner)} vs {len(p_inner)}")
            return None, None
        # 验证内部点范围
        for x in x_inner:
            if not (min_val - 1e-12 <= x <= max_val + 1e-12):
                logger.error(f"X 值 {x} 超出范围 [{min_val}, {max_val}]")
                return None, None
        for p in p_inner:
            if not (0 <= p <= 1):
                logger.error(f"P 值 {p} 超出 [0,1]")
                return None, None
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
        # 验证完整数组的有效性
        try:
            from dist_cumul import _parse_arrays as cumul_parse_arrays
            cumul_parse_arrays(x_full, p_full)
        except Exception as e:
            logger.error(f"Cumul 数组验证失败: {e}")
            return None, None
        x_full_str = ','.join(str(x) for x in x_full)
        p_full_str = ','.join(str(p) for p in p_full)
        return x_full_str, p_full_str
    # 保留原有的 _parse_cumul_discrete_args 用于 Discrete
    def _parse_cumul_discrete_args(self, args_text: str, app) -> Tuple[Optional[str], Optional[str]]:
        """
        从 DriskCumul（旧版）或 DriskDiscrete 的参数文本中提取 x_vals 和 p_vals 字符串。
        支持花括号数组常量（如 {1,2,3}）和单元格区域引用（如 A1:A10）。
        返回 (x_str, p_str)，已去除花括号和引号，并用逗号分隔。
        若解析失败返回 (None, None)
        """
        if not args_text:
            return None, None
        # 尝试方法1：使用自定义分割
        try:
            args_list = self._split_cumul_discrete_args(args_text)
            if len(args_list) >= 2:
                x_raw = args_list[0].strip().strip('"\'')
                p_raw = args_list[1].strip().strip('"\'')
                x_str = None
                p_str = None
                # 处理 X 参数
                # 判断是否为区域引用
                range_pattern = r'^([A-Z]{1,3}\d{1,7})(?::([A-Z]{1,3}\d{1,7}))?$'
                if re.match(range_pattern, x_raw, re.IGNORECASE) or ('!' in x_raw and re.match(r'.+![A-Z]{1,3}\d{1,7}(?::[A-Z]{1,3}\d{1,7})?$', x_raw, re.IGNORECASE)):
                    try:
                        x_vals = self._read_range_values(app, x_raw)
                        x_str = ','.join(str(v) for v in x_vals)
                    except Exception as e:
                        logger.error(f"读取X值区域失败: {e}")
                        return None, None
                else:
                    # 处理花括号或普通字符串
                    if x_raw.startswith('{') and x_raw.endswith('}'):
                        x_raw = x_raw[1:-1]
                    x_parts = [part.strip() for part in x_raw.split(',') if part.strip()]
                    try:
                        [float(x) for x in x_parts]
                        x_str = ','.join(x_parts)
                    except ValueError:
                        logger.error(f"X值解析失败，不是有效的数字列表: {x_raw}")
                        return None, None
                # 处理 P 参数
                if re.match(range_pattern, p_raw, re.IGNORECASE) or ('!' in p_raw and re.match(r'.+![A-Z]{1,3}\d{1,7}(?::[A-Z]{1,3}\d{1,7})?$', p_raw, re.IGNORECASE)):
                    try:
                        p_vals = self._read_range_values(app, p_raw)
                        p_str = ','.join(str(v) for v in p_vals)
                    except Exception as e:
                        logger.error(f"读取P值区域失败: {e}")
                        return None, None
                else:
                    if p_raw.startswith('{') and p_raw.endswith('}'):
                        p_raw = p_raw[1:-1]
                    p_parts = [part.strip() for part in p_raw.split(',') if part.strip()]
                    try:
                        [float(p) for p in p_parts]
                        p_str = ','.join(p_parts)
                    except ValueError:
                        logger.error(f"P值解析失败，不是有效的数字列表: {p_raw}")
                        return None, None
                if x_str and p_str:
                    logger.debug(f"解析Discrete参数成功: x_vals='{x_str}', p_vals='{p_str}'")
                    return x_str, p_str
        except Exception as e:
            logger.debug(f"方法1解析失败: {e}，尝试方法2")
        # 尝试方法2：使用正则表达式直接匹配两个花括号数组或区域引用
        try:
            # 匹配花括号数组 {1,2,3}
            pattern_brace = r'\{\s*([^}]+)\s*\}\s*,\s*\{\s*([^}]+)\s*\}'
            # 匹配区域引用 A1:A10, B1:B5
            pattern_range = r'([A-Z]{1,3}\d{1,7}(?::[A-Z]{1,3}\d{1,7})?)\s*,\s*([A-Z]{1,3}\d{1,7}(?::[A-Z]{1,3}\d{1,7})?)'
            # 匹配混合情况，先尝试花括号
            match = re.search(pattern_brace, args_text, re.IGNORECASE)
            if match:
                x_raw = match.group(1).strip()
                p_raw = match.group(2).strip()
            else:
                match = re.search(pattern_range, args_text, re.IGNORECASE)
                if match:
                    x_raw = match.group(1).strip()
                    p_raw = match.group(2).strip()
                else:
                    return None, None
            # 处理 X
            if re.match(r'^[A-Z]{1,3}\d', x_raw, re.IGNORECASE):
                try:
                    x_vals = self._read_range_values(app, x_raw)
                    x_str = ','.join(str(v) for v in x_vals)
                except:
                    return None, None
            else:
                x_parts = [part.strip() for part in x_raw.split(',') if part.strip()]
                try:
                    [float(x) for x in x_parts]
                    x_str = ','.join(x_parts)
                except:
                    return None, None
            # 处理 P
            if re.match(r'^[A-Z]{1,3}\d', p_raw, re.IGNORECASE):
                try:
                    p_vals = self._read_range_values(app, p_raw)
                    p_str = ','.join(str(v) for v in p_vals)
                except:
                    return None, None
            else:
                p_parts = [part.strip() for part in p_raw.split(',') if part.strip()]
                try:
                    [float(p) for p in p_parts]
                    p_str = ','.join(p_parts)
                except:
                    return None, None
            if x_str and p_str:
                logger.debug(f"方法2解析Discrete参数成功: x_vals='{x_str}', p_vals='{p_str}'")
                return x_str, p_str
        except Exception as e:
            logger.debug(f"方法2解析失败: {e}")
        # 所有方法失败
        logger.error(f"解析Discrete参数最终失败: args_text='{args_text}'")
        return None, None
    def _parse_duniform_args(self, args_text, app):
        """DriskDUniform  x_vals  p_vals """
        try:
            args_list = self._split_cumul_discrete_args(args_text)
            if len(args_list) < 1:
                return None, None
            x_raw = args_list[0].strip().strip("'\"")
            range_pattern = r'^[A-Z]{1,3}\d{1,7}(?::[A-Z]{1,3}\d{1,7})?$'
            if re.match(range_pattern, x_raw, re.IGNORECASE) or ('!' in x_raw and re.match(r'.+![A-Z]{1,3}\d{1,7}(?::[A-Z]{1,3}\d{1,7})?$', x_raw, re.IGNORECASE)):
                try:
                    x_vals = self._read_range_values(app, x_raw)
                    x_str = ','.join(str(v) for v in x_vals)
                except Exception as e:
                    logger.error(f"DUniform X: {e}")
                    return None, None
            else:
                if x_raw.startswith('{') and x_raw.endswith('}'):
                    x_raw = x_raw[1:-1]
                x_parts = [part.strip() for part in x_raw.split(',') if part.strip()]
                try:
                    x_vals = [float(part) for part in x_parts]
                    x_str = ','.join(str(v) for v in x_vals)
                except ValueError:
                    logger.error(f"DUniform X {x_raw}")
                    return None, None
            if not x_vals:
                return None, None
            if len(x_vals) != len(set(x_vals)):
                logger.error('DUniform X-table values must be unique')
                return None, None
            p = 1.0 / len(x_vals)
            p_str = ','.join(str(p) for _ in x_vals)
            return x_str, p_str
        except Exception as e:
            logger.error(f" DriskDUniform : {e}")
            return None, None
    def _parse_general_4args(self, args_text: str, app) -> Tuple[Optional[str], Optional[str]]:
        """
        从 DriskGeneral 的参数字符串中解析 x_vals 和 p_vals 字符串。
        参数格式：min_val, max_val, x_param, p_param
        返回 (x_str, p_str)
        """
        if not args_text:
            return None, None
        args_list = self._split_cumul_discrete_args(args_text)
        if len(args_list) < 4:
            logger.error(f"General 参数不足，需要 4 个，实际 {len(args_list)}: {args_text}")
            return None, None
        min_raw = args_list[0].strip().strip('"\'')
        max_raw = args_list[1].strip().strip('"\'')
        x_raw = args_list[2].strip().strip('"\'')
        p_raw = args_list[3].strip().strip('"\'')
        try:
            min_val = float(self._read_range_values(app, min_raw)[0] if re.match(r'[A-Z]+\d+', min_raw, re.I) else float(min_raw))
            max_val = float(self._read_range_values(app, max_raw)[0] if re.match(r'[A-Z]+\d+', max_raw, re.I) else float(max_raw))
        except Exception as e:
            logger.error(f"解析 General 的 min/max 失败: {e}")
            return None, None
        x_str, p_str = self._parse_cumul_discrete_args(f"{x_raw},{p_raw}", app)
        if not x_str or not p_str:
            return None, None
        try:
            general_parse_arrays(min_val, max_val, x_str, p_str)
        except Exception as e:
            logger.error(f"General 数组验证失败: {e}")
            return None, None
        return x_str, p_str
    def _parse_histogrm_3args(self, args_text: str, app) -> Optional[str]:
        """
        从 DriskHistogrm 的参数字符串中解析 p_vals 字符串。
        参数格式：min_val, max_val, p_param
        返回 p_str
        """
        if not args_text:
            return None
        args_list = self._split_cumul_discrete_args(args_text)
        if len(args_list) < 3:
            logger.error(f"Histogrm 参数不足，需要 3 个，实际 {len(args_list)}: {args_text}")
            return None
        p_raw = args_list[2].strip().strip('"\'')
        range_pattern = r'^[A-Z]{1,3}\d{1,7}(?::[A-Z]{1,3}\d{1,7})?$'
        if re.match(range_pattern, p_raw, re.IGNORECASE) or ('!' in p_raw and re.match(r'.+![A-Z]{1,3}\d{1,7}(?::[A-Z]{1,3}\d{1,7})?$', p_raw, re.IGNORECASE)):
            try:
                p_vals = self._read_range_values(app, p_raw)
                p_str = ','.join(str(v) for v in p_vals)
            except Exception as e:
                logger.error(f"Histogrm P-Table: {e}")
                return None
        else:
            if p_raw.startswith('{') and p_raw.endswith('}'):
                p_raw = p_raw[1:-1]
            p_parts = [part.strip() for part in p_raw.split(',') if part.strip()]
            try:
                p_vals = [float(part) for part in p_parts]
                p_str = ','.join(str(v) for v in p_vals)
            except ValueError:
                logger.error(f"Histogrm P-Table {p_raw}")
                return None
        try:
            histogrm_parse_p_table(p_str)
        except Exception as e:
            logger.error(f"Histogrm 数组验证失败: {e}")
            return None
        return p_str
    def save_excel_settings(self, app):
        """保存Excel原始设置"""
        try:
            self.original_settings = {
                'screen_updating': app.ScreenUpdating,
                'calculation': app.Calculation,
                'enable_events': app.EnableEvents,
                'display_alerts': app.DisplayAlerts
            }
            return True
        except Exception as e:
            logger.error(f"保存Excel设置失败: {str(e)}")
            return False
    def optimize_excel_performance(self, app):
        """优化Excel性能"""
        try:
            app.ScreenUpdating = False
            app.Calculation = -4135  # xlCalculationManual
            app.EnableEvents = False
            app.DisplayAlerts = False
            return True
        except Exception as e:
            logger.error(f"优化Excel性能失败: {str(e)}")
            return False
    def restore_excel_settings(self, app):
        """恢复Excel原始设置"""
        try:
            if not self.original_settings:
                return False
            app.ScreenUpdating = self.original_settings.get('screen_updating', True)
            app.Calculation = self.original_settings.get('calculation', -4105)  # xlCalculationAutomatic
            app.EnableEvents = self.original_settings.get('enable_events', True)
            app.DisplayAlerts = self.original_settings.get('display_alerts', True)
            return True
        except Exception as e:
            logger.error(f"恢复Excel设置失败: {str(e)}")
            return False
    def _column_to_letter(self, n: int) -> str:
        """将列号转换为字母"""
        result = ""
        while n > 0:
            n, remainder = divmod(n - 1, 26)
            result = chr(65 + remainder) + result
        return result
    def _ensure_hidden_sheet(self, app):
        """确保隐藏表存在 - 每次模拟前删除旧的并新建"""
        try:
            workbook = app.ActiveWorkbook
            # 尝试删除已存在的隐藏表
            try:
                old_sheet = workbook.Worksheets("DriskIndexHidden")
                old_sheet.Visible = True  # 先取消隐藏以便删除
                old_sheet.Delete()
                logger.info("删除旧的DriskIndexHidden工作表")
            except:
                pass
            # 创建新的隐藏表
            self.hidden_sheet = workbook.Worksheets.Add()
            self.hidden_sheet.Name = "DriskIndexHidden"
            self.hidden_sheet.Visible = 0  # 隐藏工作表
            logger.info("创建新的DriskIndexHidden工作表")
            # 写入表头行
            self.hidden_sheet.Cells(1, 1).Value = "Iteration"
            # 设置控制单元格 - 在A2单元格
            self.control_cell = self.hidden_sheet.Range("A2")
            self.control_cell.Value = 1
            # 输出引用行的起始列将在setup_output_references中动态计算，此处不清空
            return True
        except Exception as e:
            logger.error(f"创建隐藏工作表失败: {str(e)}")
            self.hidden_sheet = None
            return False
    # ---------- 关键修改：register_input 方法（每个场景强制覆盖数据，并记录列映射）----------
    def register_input(self, input_key: str, values: np.ndarray):
        """注册输入数据到隐藏表（每个场景强制覆盖数据）"""
        with self._lock:
            if not self.hidden_sheet:
                logger.error(f"隐藏工作表不存在，无法注册输入: {input_key}")
                return None
            if input_key in self.input_columns:
                column_letter = self.input_columns[input_key]
                col_index = 0
                for ch in column_letter.upper():
                    col_index = col_index * 26 + (ord(ch) - ord('A') + 1)
                logger.debug(f"重用已有列 {column_letter}({col_index}) 更新数据: {input_key}")
            else:
                next_col = len(self.input_columns) + 2
                column_letter = self._column_to_letter(next_col)
                col_index = next_col
                self.input_columns[input_key] = column_letter
                self.hidden_sheet.Cells(1, next_col).Value = input_key
                logger.debug(f"分配新列 {column_letter}({col_index}) 给 {input_key}")
            # 记录列映射
            self.input_key_to_column[input_key] = column_letter
            try:
                if len(values) != self.total_iterations:
                    logger.warning(f"输入数据长度{len(values)}与总迭代次数{self.total_iterations}不匹配，进行截断/填充")
                    if len(values) > self.total_iterations:
                        values = values[:self.total_iterations]
                    else:
                        values = np.append(values, np.full(self.total_iterations - len(values), values[-1] if len(values) > 0 else 0))
                start_row = 2
                end_row = start_row + len(values) - 1
                if end_row >= start_row:
                    range_obj = self.hidden_sheet.Range(f"{column_letter}{start_row}:{column_letter}{end_row}")
                    values_list = []
                    for v in values:
                        if v == ERROR_MARKER or (isinstance(v, str) and v.upper() == "#ERROR!"):
                            values_list.append(["#ERROR!"])
                        else:
                            try:
                                values_list.append([float(v)])
                            except:
                                values_list.append([0.0])
                    range_obj.Value2 = values_list
                    logger.debug(f"批量写入数据到 {column_letter}{start_row}:{column_letter}{end_row}")
            except Exception as e:
                logger.error(f"批量写入失败: {str(e)}")
                try:
                    for i, val in enumerate(values):
                        row = 2 + i
                        if val == ERROR_MARKER or (isinstance(val, str) and val.upper() == "#ERROR!"):
                            self.hidden_sheet.Cells(row, col_index).Value = "#ERROR!"
                        else:
                            try:
                                self.hidden_sheet.Cells(row, col_index).Value = float(val)
                            except:
                                self.hidden_sheet.Cells(row, col_index).Value = 0.0
                except Exception as e2:
                    logger.error(f"备用写入也失败: {str(e2)}")
            self.input_data[input_key] = values.copy()
            logger.debug(f"注册输入成功: {input_key} -> 列 {column_letter}")
            return column_letter
    def backup_original_formulas_and_values(self, app):
        """备份原始公式和Simtable值"""
        try:
            self.original_formulas = {}
            self.original_simtable_values = {}
            for cell_addr, dist_funcs in self.distribution_cells.items():
                if '_nested_' in cell_addr or '_virtual_' in cell_addr:
                    continue
                try:
                    if '!' in cell_addr:
                        sheet_name, addr = cell_addr.split('!')
                        sheet = app.ActiveWorkbook.Worksheets(sheet_name)
                        cell = sheet.Range(addr)
                    else:
                        cell = app.ActiveSheet.Range(cell_addr)
                    formula = cell.Formula
                    if isinstance(formula, str) and formula.startswith('='):
                        self.original_formulas[cell_addr] = formula
                except Exception as e:
                    logger.error(f"备份分布单元格公式失败 {cell_addr}: {str(e)}")
            for cell_addr, makeinput_funcs in self.makeinput_cells.items():
                try:
                    if '!' in cell_addr:
                        sheet_name, addr = cell_addr.split('!')
                        sheet = app.ActiveWorkbook.Worksheets(sheet_name)
                        cell = sheet.Range(addr)
                    else:
                        cell = app.ActiveSheet.Range(cell_addr)
                    formula = cell.Formula
                    if isinstance(formula, str) and formula.startswith('='):
                        self.original_formulas[cell_addr] = formula
                except Exception as e:
                    logger.error(f"备份MakeInput单元格公式失败 {cell_addr}: {str(e)}")
            for cell_addr, simtable_funcs in self.simtable_cells.items():
                try:
                    if '!' in cell_addr:
                        sheet_name, addr = cell_addr.split('!')
                        sheet = app.ActiveWorkbook.Worksheets(sheet_name)
                        cell = sheet.Range(addr)
                    else:
                        cell = app.ActiveSheet.Range(cell_addr)
                    value = cell.Value
                    formula = cell.Formula
                    if isinstance(formula, str) and formula.startswith('='):
                        self.original_simtable_values[cell_addr] = formula
                    else:
                        self.original_simtable_values[cell_addr] = value
                    logger.debug(f"备份Simtable单元格 {cell_addr}: 公式={formula}, 值={value}")
                except Exception as e:
                    logger.error(f"备份Simtable单元格值失败 {cell_addr}: {str(e)}")
            logger.info(f"备份了 {len(self.original_formulas)} 个单元格的原始公式和 {len(self.original_simtable_values)} 个Simtable值")
            return True
        except Exception as e:
            logger.error(f"备份原始公式和值失败: {str(e)}")
            return False
    def restore_original_formulas_and_values(self, app):
        """恢复原始公式和Simtable值"""
        try:
            restored_count = 0
            for cell_addr, formula in self.original_formulas.items():
                try:
                    if '!' in cell_addr:
                        sheet_name, addr = cell_addr.split('!')
                        sheet = app.ActiveWorkbook.Worksheets(sheet_name)
                        cell = sheet.Range(addr)
                    else:
                        cell = app.ActiveSheet.Range(cell_addr)
                    cell.Formula = formula
                    restored_count += 1
                except Exception as e:
                    logger.error(f"恢复公式失败 {cell_addr}: {str(e)}")
            simtable_restored_count = 0
            for cell_addr, original_value in self.original_simtable_values.items():
                try:
                    if '!' in cell_addr:
                        sheet_name, addr = cell_addr.split('!')
                        sheet = app.ActiveWorkbook.Worksheets(sheet_name)
                        cell = sheet.Range(addr)
                    else:
                        cell = app.ActiveSheet.Range(cell_addr)
                    if isinstance(original_value, str) and original_value.startswith('='):
                        cell.Formula = original_value
                    else:
                        cell.Value = original_value
                    simtable_restored_count += 1
                except Exception as e:
                    logger.error(f"恢复Simtable单元格值失败 {cell_addr}: {str(e)}")
            try:
                app.Calculate()
            except:
                pass
            logger.info(f"恢复了 {restored_count} 个单元格的原始公式和 {simtable_restored_count} 个Simtable值")
            return True
        except Exception as e:
            logger.error(f"恢复原始公式和值失败: {str(e)}")
            return False
    # ========== 关键修改：setup_index_formulas（逐函数替换分布单元格内的所有分布） ==========
    def setup_index_formulas(self, app):
        """设置INDEX公式 - 包含分布单元格和内嵌在MakeInput中的分布函数"""
        try:
            set_count = 0
            # ---------- 处理普通分布单元格 ----------
            for cell_addr, dist_funcs in self.distribution_cells.items():
                if '_nested_' in cell_addr or '_virtual_' in cell_addr:
                    continue
                # 获取该单元格的原始公式
                original_formula = self.original_formulas.get(cell_addr, '')
                if not original_formula or not original_formula.startswith('='):
                    logger.error(f"分布单元格 {cell_addr} 无原始公式，跳过")
                    continue
                # 按分布函数在原公式中的位置逆序排序（按start_pos降序）
                sorted_funcs = sorted(dist_funcs, key=lambda x: x.get('start_pos', 0), reverse=True)
                new_formula = original_formula
                replaced_any = False
                for func in sorted_funcs:
                    input_key = func.get('input_key')
                    if not input_key:
                        logger.error(f"分布函数 {cell_addr} 缺少input_key，跳过INDEX公式设置")
                        continue
                    if input_key not in self.input_columns:
                        logger.error(f"输入键 {input_key} 不在input_columns中，可能未注册，跳过INDEX公式设置")
                        continue
                    full_match = func.get('full_match')
                    if not full_match:
                        logger.warning(f"分布函数缺少full_match，无法替换: {func}")
                        continue
                    column_letter = self.input_columns[input_key]
                    index_part = f'INDEX(\'DriskIndexHidden\'!${column_letter}$2:${column_letter}${self.total_iterations+1}, \'DriskIndexHidden\'!$A$2)'
                    # 替换原分布函数的完整文本
                    new_formula = new_formula.replace(full_match, index_part)
                    replaced_any = True
                    logger.debug(f"替换分布函数: {full_match} -> {index_part}")
                if replaced_any:
                    try:
                        if '!' in cell_addr:
                            sheet_name, addr = cell_addr.split('!')
                            sheet = app.ActiveWorkbook.Worksheets(sheet_name)
                            cell = sheet.Range(addr)
                        else:
                            cell = app.ActiveSheet.Range(cell_addr)
                        cell.Formula = new_formula
                        set_count += 1
                        print(f"设置INDEX公式: {cell_addr} -> {new_formula}")
                    except Exception as e:
                        logger.error(f"设置INDEX公式失败 {cell_addr}: {str(e)}")
                else:
                    logger.warning(f"分布单元格 {cell_addr} 无任何可替换的分布函数")
            # ---------- 处理DriskMakeInput单元格（替换其内部的分布函数） ----------
            for cell_addr, makeinput_funcs in self.makeinput_cells.items():
                if not makeinput_funcs:
                    continue
                # 取第一个MakeInput函数（一个单元格通常只有一个）
                mi_func = makeinput_funcs[0]
                # 关键修改：直接使用 self.distribution_cells 中已有的分布函数信息（它们已包含 input_key）
                if cell_addr not in self.distribution_cells:
                    # 理论上该单元格的分布函数应该已被记录，但如果没有则跳过
                    logger.warning(f"MakeInput单元格 {cell_addr} 在 distribution_cells 中无对应项，跳过替换")
                    continue
                nested_funcs = self.distribution_cells[cell_addr]  # 使用已记录的分布函数列表
                if not nested_funcs:
                    continue   # 无内嵌分布，无需修改
                original_formula = self.original_formulas.get(cell_addr, '')
                if not original_formula or not original_formula.startswith('='):
                    continue
                # 按分布函数在原公式中的位置逆序替换，避免破坏索引
                sorted_funcs = sorted(nested_funcs, key=lambda x: x.get('start_pos', 0), reverse=True)
                new_formula = original_formula
                for nf in sorted_funcs:
                    input_key = nf.get('input_key')
                    if not input_key or input_key not in self.input_columns:
                        continue
                    column_letter = self.input_columns[input_key]
                    # 构造INDEX公式（不带最外层等号，用于替换）
                    index_part = f'INDEX(\'DriskIndexHidden\'!${column_letter}$2:${column_letter}${self.total_iterations+1}, \'DriskIndexHidden\'!$A$2)'
                    # 替换原分布函数的完整文本
                    full_match = nf.get('full_match', '')
                    if full_match:
                        new_formula = new_formula.replace(full_match, index_part)
                # 将修改后的公式写回单元格
                try:
                    if '!' in cell_addr:
                        sheet_name, addr = cell_addr.split('!')
                        sheet = app.ActiveWorkbook.Worksheets(sheet_name)
                        cell = sheet.Range(addr)
                    else:
                        cell = app.ActiveSheet.Range(cell_addr)
                    cell.Formula = new_formula
                    set_count += 1
                    print(f"替换MakeInput内嵌分布: {cell_addr} -> {new_formula}")
                except Exception as e:
                    logger.error(f"设置MakeInput公式失败 {cell_addr}: {str(e)}")
            logger.info(f"共设置了 {set_count} 个INDEX公式（含内嵌分布）")
            return True
        except Exception as e:
            logger.error(f"设置INDEX公式失败: {str(e)}")
            traceback.print_exc()
            return False
    # ========== 关键修改：setup_output_references（将输出引用列移至输入列右侧） ==========
    def setup_output_references(self, app):
        """在隐藏表创建对输出单元格和MakeInput单元格的引用公式，用于批量读取值"""
        try:
            if not self.hidden_sheet:
                logger.error("隐藏表不存在，无法设置输出引用")
                return False
            # 计算起始列：第一列是迭代列，第二列开始是输入列，所以输出引用从第 2 + len(self.input_columns) 列开始
            start_col = 2 + len(self.input_columns)
            # 清空从 start_col 开始的第3行
            clear_range = self.hidden_sheet.Range(
                self.hidden_sheet.Cells(self.output_refs_start_row, start_col),
                self.hidden_sheet.Cells(self.output_refs_start_row, self.hidden_sheet.Columns.Count)
            )
            clear_range.ClearContents()
            self.output_refs = []
            current_col = start_col
            # 处理输出单元格
            for cell_addr in self.output_cells.keys():
                if '!' in cell_addr:
                    sheet_name, addr = cell_addr.split('!', 1)
                else:
                    sheet_name = app.ActiveSheet.Name
                    addr = cell_addr
                # 构造绝对引用公式：='SheetName'!$A$1
                if ' ' in sheet_name or any(c in sheet_name for c in "[]:?*"):
                    sheet_ref = f"'{sheet_name}'"
                else:
                    sheet_ref = sheet_name
                clean_addr = addr.replace('$', '')
                match = re.match(r'([A-Za-z]+)(\d+)$', clean_addr)
                if match:
                    col_letter, row_num = match.groups()
                    abs_addr = f"${col_letter.upper()}${row_num}"
                    formula = f"={sheet_ref}!{abs_addr}"
                else:
                    # 若解析失败（如范围），直接使用原地址，但仍加$符号
                    formula = f"={sheet_ref}!${clean_addr}$"
                cell = self.hidden_sheet.Cells(self.output_refs_start_row, current_col)
                cell.Formula = formula
                self.output_refs.append((cell_addr, 'output'))
                current_col += 1
            # 处理MakeInput单元格
            for cell_addr in self.makeinput_cells.keys():
                if '!' in cell_addr:
                    sheet_name, addr = cell_addr.split('!', 1)
                else:
                    sheet_name = app.ActiveSheet.Name
                    addr = cell_addr
                if ' ' in sheet_name or any(c in sheet_name for c in "[]:?*"):
                    sheet_ref = f"'{sheet_name}'"
                else:
                    sheet_ref = sheet_name
                clean_addr = addr.replace('$', '')
                match = re.match(r'([A-Za-z]+)(\d+)$', clean_addr)
                if match:
                    col_letter, row_num = match.groups()
                    abs_addr = f"${col_letter.upper()}${row_num}"
                    formula = f"={sheet_ref}!{abs_addr}"
                else:
                    formula = f"={sheet_ref}!${clean_addr}$"
                cell = self.hidden_sheet.Cells(self.output_refs_start_row, current_col)
                cell.Formula = formula
                self.output_refs.append((cell_addr, 'makeinput'))
                current_col += 1
            if self.output_refs:
                self.output_refs_range = self.hidden_sheet.Range(
                    self.hidden_sheet.Cells(self.output_refs_start_row, start_col),
                    self.hidden_sheet.Cells(self.output_refs_start_row, current_col - 1)
                )
                logger.info(f"设置了 {len(self.output_refs)} 个输出引用公式，起始列: {start_col}")
            else:
                self.output_refs_range = None
                logger.warning("没有输出或MakeInput单元格需要引用")
            # 更新起始列属性，供后续可能使用
            self.output_refs_start_col = start_col
            return True
        except Exception as e:
            logger.error(f"设置输出引用公式失败: {str(e)}")
            traceback.print_exc()
            return False
    # ========== 构建单元格到输入键的映射（返回列表，支持多分布函数） ==========
    def _build_cell_to_input_key_map(self) -> Dict[str, List[str]]:
        """构建单元格地址到该单元格内所有 input_key 的映射"""
        cell_to_keys = {}
        for cell_addr, dist_funcs in self.distribution_cells.items():
            if '_nested_' in cell_addr or '_virtual_' in cell_addr:
                continue
            keys = []
            for func in dist_funcs:
                input_key = func.get('input_key')
                if input_key:
                    keys.append(input_key)
            if keys:
                cell_to_keys[cell_addr] = keys
        return cell_to_keys
    # ========== 解析单个函数的依赖 ==========
    def _parse_function_dependencies(self, func_info: Dict, cell_to_input_keys: Dict[str, List[str]]) -> List[str]:
        """
        解析一个分布函数的依赖 input_key 列表。
        依赖来源：
        1. 参数文本中直接出现的其他分布函数调用（通过 full_match 字符串匹配）
        2. 参数文本中出现的单元格引用，且该单元格包含分布函数
        """
        deps = set()
        args_text = func_info.get('args_text', '')
        full_match = func_info.get('full_match', '')
        # 方法1：扫描所有其他 input_key 的 full_match 是否出现在 args_text 中
        for other_key, other_info in self.input_key_to_func_info.items():
            if other_key == func_info.get('input_key'):
                continue
            other_full = other_info.get('full_match', '')
            if other_full and other_full in args_text:
                deps.add(other_key)
        # 方法2：扫描单元格引用
        cell_pattern = r'[A-Z]{1,3}\d{1,7}'
        cell_matches = re.findall(cell_pattern, args_text, re.IGNORECASE)
        for cell in cell_matches:
            cell_upper = cell.upper()
            # 确定工作表名
            current_sheet = None
            if '!' in func_info.get('key', ''):
                current_sheet = func_info['key'].split('!')[0]
            full_cell_addr = f"{current_sheet}!{cell_upper}" if current_sheet else cell_upper
            # 如果该单元格包含分布函数，则依赖该单元格的所有 input_key
            if full_cell_addr in cell_to_input_keys:
                for key in cell_to_input_keys[full_cell_addr]:
                    deps.add(key)
        return list(deps)
    # ========== 构建依赖图 ==========
    def build_dependency_graph(self, distribution_cells: Dict) -> Tuple[Dict[str, List[str]], List[str]]:
        """
        构建依赖图，返回 (graph, sorted_keys)
        graph: {input_key: [依赖的input_key列表]}
        sorted_keys: 拓扑排序后的 input_key 列表（若无循环）
        若存在循环依赖，则抛出异常
        """
        logger.info("开始构建依赖图...")
        # 首先确保 input_key_to_func_info 已填充（应在分配 input_key 时填充）
        all_keys = list(self.input_key_to_func_info.keys())
        graph = {key: [] for key in all_keys}
        # 构建单元格到 input_key 的映射
        cell_to_keys = self._build_cell_to_input_key_map()
        # 解析每个函数的依赖
        for key, func_info in self.input_key_to_func_info.items():
            deps = self._parse_function_dependencies(func_info, cell_to_keys)
            graph[key] = deps
        # 检测循环依赖并拓扑排序
        try:
            sorted_keys = self._dfs_topological_sort(graph, all_keys)
            logger.info(f"依赖图构建完成，共 {len(all_keys)} 个节点，拓扑排序顺序：{sorted_keys}")
            return graph, sorted_keys
        except Exception as e:
            logger.error(f"依赖图存在循环或排序失败: {str(e)}")
            raise ValueError(f"输入函数之间存在循环依赖，无法进行模拟: {str(e)}")
    def _dfs_topological_sort(self, graph, all_keys):
        """使用DFS进行拓扑排序"""
        visited = set()
        temp = set()
        order = []
        def dfs(node):
            if node in temp:
                return False
            if node in visited:
                return True
            temp.add(node)
            for neighbor in graph.get(node, []):
                if not dfs(neighbor):
                    return False
            temp.remove(node)
            visited.add(node)
            order.append(node)
            return True
        for key in all_keys:
            if key not in visited:
                if not dfs(key):
                    logger.warning("依赖图存在循环，将按原顺序生成")
                    return all_keys
        return order
    # ========== 解析参数并生成样本的辅助方法（增强：过滤属性函数） ==========
    def _is_attribute_function(self, arg_str: str) -> bool:
        """判断一个字符串是否为属性函数"""
        # 属性函数列表（可从全局导入，这里直接定义部分常用）
        attr_funcs = ["DriskShift", "DriskTruncate", "DriskTruncate2", "DriskTruncateP", "DriskTruncateP2",
                      "DriskLoc", "DriskLock", "DriskSeed", "DriskStatic", "DriskName", "DriskUnits",
                      "DriskCategory", "DriskCollect", "DriskConvergence", "DriskCopula", "DriskCorrmat",
                      "DriskFit", "DriskIsDate", "DriskIsDiscrete"]
        for af in attr_funcs:
            if arg_str.strip().startswith(af) and '(' in arg_str:
                return True
        return False
    # ========== 新增：解析参数，支持常量单元格回退 ==========
    def _parse_args_for_generation(self, key: str, func_info: Dict, full_match_to_key: Dict[str, str], app) -> Tuple[List, List]:
        """
        解析分布函数的参数，确定每个参数是常量还是依赖键。
        返回 (param_constants, param_deps)，长度相同。
        param_constants[i] 如果是常量则为数值，否则为 None
        param_deps[i] 如果是依赖则为依赖键，否则为 None

        增强：若参数是单元格引用且不在 input_key_to_func_info 中，则视为常量单元格，直接读取其当前值。
        """
        args_text = func_info.get('args_text', '')
        if not args_text:
            return [], []
        # 使用 formula_parser 分割参数（处理嵌套引号等）
        from formula_parser import parse_args_with_nested_functions
        arg_strings = parse_args_with_nested_functions(args_text)
        param_constants = []
        param_deps = []
        func_name_lower = str(func_info.get('func_name', '')).lower()
        is_compound = func_name_lower == 'driskcompound'
        is_splice = func_name_lower == 'drisksplice'
        for idx, arg_str in enumerate(arg_strings):
            arg_str = arg_str.strip()
            # 跳过属性函数
            if self._is_attribute_function(arg_str):
                continue
            # 尝试解析为数值常量
            if is_compound and idx in (0, 1) and arg_str.startswith('Drisk') and '(' in arg_str:
                param_constants.append(arg_str)
                param_deps.append(None)
                continue
            if is_splice and idx in (0, 1) and arg_str.startswith('Drisk') and '(' in arg_str:
                param_constants.append(arg_str)
                param_deps.append(None)
                continue
            if is_compound and idx in (2, 3) and not arg_str:
                param_constants.append(0.0 if idx == 2 else float('inf'))
                param_deps.append(None)
                continue
            try:
                val = float(arg_str)
                param_constants.append(val)
                param_deps.append(None)
                continue
            except:
                pass
            # 允许 Excel 数组常量作为普通常量参数传递，例如 {1,2,3}
            if arg_str.startswith('{') and arg_str.endswith('}'):
                param_constants.append(arg_str)
                param_deps.append(None)
                continue
            # 检查是否是另一个分布函数的 full_match
            if arg_str in full_match_to_key:
                param_constants.append(None)
                param_deps.append(full_match_to_key[arg_str])
                continue
            # 检查是否是单元格引用（如 A1 或 Sheet1!A1）
            cell_pattern = r'^([A-Z]{1,3}\d{1,7})$'
            sheet_cell_pattern = r'^(.+![A-Z]{1,3}\d{1,7})$'
            if re.match(cell_pattern, arg_str, re.IGNORECASE) or re.match(sheet_cell_pattern, arg_str, re.IGNORECASE):
                # 规范化地址
                if '!' in arg_str:
                    # 已有工作表名
                    full_cell = arg_str.upper()
                else:
                    # 从 key 中获取当前工作表名
                    if '!' in key:
                        current_sheet = key.split('!')[0]
                    else:
                        current_sheet = app.ActiveSheet.Name
                    full_cell = f"{current_sheet}!{arg_str.upper()}"
                # 检查是否对应一个分布函数
                if full_cell in self.input_key_to_func_info:
                    param_constants.append(None)
                    param_deps.append(full_cell)
                    continue
                else:
                    # 常量单元格引用，直接读取值
                    try:
                        if '!' in arg_str:
                            sheet_name, cell_addr = arg_str.split('!', 1)
                            sheet = app.ActiveWorkbook.Worksheets(sheet_name)
                        else:
                            sheet = app.ActiveSheet
                            cell_addr = arg_str
                        cell = sheet.Range(cell_addr)
                        val = cell.Value2
                        if val is None:
                            val = cell.Value
                        # 处理单元格值，转换为数值或错误标记
                        processed = self._process_cell_value(val)
                        if processed == ERROR_MARKER:
                            # 无法获取有效值，返回错误标记
                            param_constants.append(ERROR_MARKER)
                            param_deps.append(None)
                        else:
                            try:
                                param_constants.append(float(processed))
                                param_deps.append(None)
                            except:
                                param_constants.append(ERROR_MARKER)
                                param_deps.append(None)
                        continue
                    except Exception as e:
                        logger.error(f"读取常量单元格 {arg_str} 失败: {e}")
                        param_constants.append(ERROR_MARKER)
                        param_deps.append(None)
                        continue
            # 无法解析，报错
            raise ValueError(f"无法解析参数 '{arg_str}' 在函数 {key} 中，可能不是有效的常量、函数调用或单元格引用")
        return param_constants, param_deps
    # ========== 按依赖顺序生成样本（动态生成，含错误标记检测） ==========
    def generate_samples_with_dependencies(self, n_iterations: int, sim, app, progress) -> Tuple[Dict[str, np.ndarray], bool, str]:
        """
        按依赖顺序生成所有输入样本，并立即注册到隐藏表和模拟对象。
        对于无依赖的键批量生成，对于有依赖的键逐迭代生成。
        如果遇到任何输入性错误（包括生成器返回 ERROR_MARKER），立即停止生成，返回已生成的部分和错误信息。
        返回: (samples_dict, error_occurred, error_message)
        """
        samples_dict = {}
        error_occurred = False
        error_message = ""
        # 构建 full_match -> input_key 映射，去除空格以提高匹配成功率
        full_match_to_key = {}
        for key, info in self.input_key_to_func_info.items():
            fm = info.get('full_match')
            if fm:
                fm_clean = re.sub(r'\s+', '', fm)
                full_match_to_key[fm_clean] = key
                full_match_to_key[fm] = key
        # 获取拓扑排序结果
        sorted_keys = getattr(self, 'sorted_keys', [])
        if not sorted_keys:
            sorted_keys = list(self.input_key_to_func_info.keys())
        # 先解析所有键的参数，确定依赖关系（避免重复解析）
        # 我们仍需要逐键生成，但可以预解析参数常量/依赖列表
        # 为简化，我们在循环内解析，因为常量单元格读取可能涉及 Excel 操作，无法完全预计算
        # 预先生成所有无依赖键的样本（批量）
        for key in sorted_keys:
            if progress and progress.is_cancelled():
                error_occurred = True
                error_message = "用户取消"
                break
            # 如果该键有依赖，跳过，稍后逐迭代生成
            if self.dependency_graph.get(key, []):
                continue
            func_info = self.input_key_to_func_info.get(key)
            if not func_info:
                logger.warning(f"跳过未找到 func_info 的 key: {key}")
                continue
            # ========== 计算基础种子 ==========
            markers = func_info.get('markers', {})
            user_seed = None
            seed_marker = markers.get('seed')
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
            key_hash = zlib.crc32(key.encode()) & 0x7fffffff
            if user_seed is not None:
                final_base_seed = (user_seed + key_hash) & 0x7fffffff
            else:
                final_base_seed = (self.base_time_seed + key_hash) & 0x7fffffff
            # 解析参数常量/依赖
            try:
                param_constants, param_deps = self._parse_args_for_generation(key, func_info, full_match_to_key, app)
            except Exception as e:
                logger.error(f"解析参数失败 {key}: {str(e)}")
                error_occurred = True
                error_message = f"解析参数失败：{key} - {str(e)}"
                break
            # 检查参数中是否有错误标记
            if any(p == ERROR_MARKER for p in param_constants):
                error_occurred = True
                error_message = f"参数中包含错误标记：{key}"
                break
            # 无依赖键：所有 param_deps 应为 None
            if any(d is not None for d in param_deps):
                # 本应是无依赖键，但解析出依赖，说明依赖图构建有问题
                logger.error(f"键 {key} 被标记为无依赖，但解析出依赖: {param_deps}")
                error_occurred = True
                error_message = f"依赖图不一致：{key}"
                break
            # 批量生成样本
            original_func_name = func_info.get('func_name', '')
            func_name_lower = original_func_name.lower()
            # 处理数组参数（Cumul/Discrete）
            base_markers = markers.copy()
            if original_func_name == 'DriskCumul':
                x_str, p_str = self._parse_cumul_4args(func_info.get('args_text', ''), app)
                if x_str and p_str:
                    base_markers['x_vals'] = x_str
                    base_markers['p_vals'] = p_str
                else:
                    logger.error(f"无法解析 {key} 的数组参数")
                    error_occurred = True
                    error_message = f"无法解析 {key} 的数组参数"
                    break
            elif original_func_name == 'DriskDiscrete':
                x_str, p_str = self._parse_cumul_discrete_args(func_info.get('args_text', ''), app)
                if x_str and p_str:
                    base_markers['x_vals'] = x_str
                    base_markers['p_vals'] = p_str
                else:
                    logger.error(f"无法解析 {key} 的数组参数")
                    error_occurred = True
                    error_message = f"无法解析 {key} 的数组参数"
                    break
            elif original_func_name == 'DriskDUniform':
                x_str, p_str = self._parse_duniform_args(func_info.get('args_text', ''), app)
                if x_str and p_str:
                    base_markers['x_vals'] = x_str
                    base_markers['p_vals'] = p_str
                else:
                    logger.error(f"无法解析 {key} 的 DUniform 数组参数")
                    error_occurred = True
                    error_message = f"无法解析 {key} 的 DUniform 数组参数"
                    break
            elif original_func_name == 'DriskGeneral':
                x_str, p_str = self._parse_general_4args(func_info.get('args_text', ''), app)
                if x_str and p_str:
                    base_markers['x_vals'] = x_str
                    base_markers['p_vals'] = p_str
                else:
                    logger.error(f"无法解析 {key} 的 General 数组参数")
                    error_occurred = True
                    error_message = f"无法解析 {key} 的 General 数组参数"
                    break
            elif original_func_name == 'DriskGeneral':
                x_str, p_str = self._parse_general_4args(func_info.get('args_text', ''), app)
                if x_str and p_str:
                    base_markers['x_vals'] = x_str
                    base_markers['p_vals'] = p_str
                else:
                    logger.error(f"无法解析 {key} 的 General 数组参数")
                    error_occurred = True
                    error_message = f"无法解析 {key} 的 General 数组参数"
                    break
            elif original_func_name == 'DriskHistogrm':
                p_str = self._parse_histogrm_3args(func_info.get('args_text', ''), app)
                if p_str:
                    base_markers['p_vals'] = p_str
                else:
                    logger.error(f"无法解析 {key} 的 Histogrm P-Table 参数")
                    error_occurred = True
                    error_message = f"无法解析 {key} 的 Histogrm P-Table 参数"
                    break
            trunc_info = _parse_truncate_from_func(func_info)
            if trunc_info:
                trunc_type, trunc_value = trunc_info
                if trunc_type not in base_markers:
                    base_markers[trunc_type] = trunc_value
            base_markers['is_nested'] = func_info.get('is_nested', False)
            if func_info.get('is_at_function'):
                base_markers['is_at_function'] = True
            # 确定分布类型
            dist_type = 'normal'
            if 'duniform' in func_name_lower:
                dist_type = 'duniform'
            elif 'johnsonsb' in func_name_lower:
                dist_type = 'johnsonsb'
            elif 'johnsonsu' in func_name_lower:
                dist_type = 'johnsonsu'
            elif 'kumaraswamy' in func_name_lower:
                dist_type = 'kumaraswamy'
            elif 'laplace' in func_name_lower:
                dist_type = 'laplace'
            elif 'loglogistic' in func_name_lower:
                dist_type = 'loglogistic'
            elif 'lognorm2' in func_name_lower:
                dist_type = 'lognorm2'
            elif 'betageneral' in func_name_lower:
                dist_type = 'betageneral'
            elif 'betasubj' in func_name_lower:
                dist_type = 'betasubj'
            elif 'burr12' in func_name_lower:
                dist_type = 'burr12'
            elif 'compound' in func_name_lower:
                dist_type = 'compound'
            elif 'splice' in func_name_lower:
                dist_type = 'splice'
            elif 'pert' in func_name_lower:
                dist_type = 'pert'
            elif 'reciprocal' in func_name_lower:
                dist_type = 'reciprocal'
            elif 'rayleigh' in func_name_lower:
                dist_type = 'rayleigh'
            elif 'weibull' in func_name_lower:
                dist_type = 'weibull'
            elif 'pearson5' in func_name_lower:
                dist_type = 'pearson5'
            elif 'pearson6' in func_name_lower:
                dist_type = 'pearson6'
            elif 'pareto2' in func_name_lower:
                dist_type = 'pareto2'
            elif 'pareto' in func_name_lower:
                dist_type = 'pareto'
            elif 'lognorm' in func_name_lower:
                dist_type = 'lognorm'
            elif 'logistic' in func_name_lower:
                dist_type = 'logistic'
            elif 'levy' in func_name_lower:
                dist_type = 'levy'
            elif 'frechet' in func_name_lower:
                dist_type = 'frechet'
            elif 'histogrm' in func_name_lower:
                dist_type = 'histogrm'
            elif 'hypsecant' in func_name_lower:
                dist_type = 'hypsecant'
            elif 'general' in func_name_lower:
                dist_type = 'general'
            elif 'fatiguelife' in func_name_lower:
                dist_type = 'fatiguelife'
            elif 'extvaluemin' in func_name_lower:
                dist_type = 'extvaluemin'
            elif 'extvalue' in func_name_lower:
                dist_type = 'extvalue'
            elif 'invgauss' in func_name_lower:
                dist_type = 'invgauss'
            elif 'intuniform' in func_name_lower:
                dist_type = 'intuniform'
            elif 'erlang' in func_name_lower:
                dist_type = 'erlang'
            elif 'erf' in func_name_lower:
                dist_type = 'erf'
            elif 'doubletriang' in func_name_lower:
                dist_type = 'doubletriang'
            elif 'cauchy' in func_name_lower:
                dist_type = 'cauchy'
            elif 'dagum' in func_name_lower:
                dist_type = 'dagum'
            elif 'uniform' in func_name_lower:
                dist_type = 'uniform'
            elif 'gamma' in func_name_lower:
                dist_type = 'gamma'
            elif 'poisson' in func_name_lower:
                dist_type = 'poisson'
            elif 'beta' in func_name_lower:
                dist_type = 'beta'
            elif 'chisq' in func_name_lower:
                dist_type = 'chisq'
            elif 'f' in func_name_lower:
                dist_type = 'f'
            elif 'student' in func_name_lower:
                dist_type = 'student'
            elif 'expon' in func_name_lower:
                dist_type = 'expon'
            elif 'bernoulli' in func_name_lower:
                dist_type = 'bernoulli'
            elif 'negbin' in func_name_lower:
                dist_type = 'negbin'
            elif 'geomet' in func_name_lower:
                dist_type = 'geomet'
            elif 'hypergeo' in func_name_lower:
                dist_type = 'hypergeo'
            elif 'triang' in func_name_lower:
                dist_type = 'triang'
            elif 'binomial' in func_name_lower:
                dist_type = 'binomial'
            elif 'trigen' in func_name_lower:
                dist_type = 'trigen'
            elif 'cumul' in func_name_lower:
                dist_type = 'cumul'
            elif 'discrete' in func_name_lower:
                dist_type = 'discrete'
            try:
                generator = TruncatableDistributionGenerator(original_func_name, dist_type, param_constants, base_markers, key, seed=final_base_seed)
                samples = generator.generate_samples(n_iterations)
                if np.any(samples == ERROR_MARKER):
                    error_occurred = True
                    error_message = f"生成样本失败：{key} - 参数无效导致部分样本为错误标记"
                    break
                samples_dict[key] = samples
                logger.debug(f"批量生成样本 (无依赖): {key}")
            except Exception as e:
                logger.error(f"生成样本失败 {key}: {str(e)}")
                error_occurred = True
                error_message = f"生成样本失败：{key} - {str(e)}"
                break
        if error_occurred:
            return samples_dict, error_occurred, error_message
        # 再生成有依赖的键（逐迭代）
        for key in sorted_keys:
            if progress and progress.is_cancelled():
                error_occurred = True
                error_message = "用户取消"
                break
            if not self.dependency_graph.get(key, []):
                continue
            func_info = self.input_key_to_func_info.get(key)
            if not func_info:
                logger.warning(f"跳过未找到 func_info 的 key: {key}")
                continue
            # ========== 计算基础种子 ==========
            markers = func_info.get('markers', {})
            user_seed = None
            seed_marker = markers.get('seed')
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
            key_hash = zlib.crc32(key.encode()) & 0x7fffffff
            if user_seed is not None:
                final_base_seed = (user_seed + key_hash) & 0x7fffffff
            else:
                final_base_seed = (self.base_time_seed + key_hash) & 0x7fffffff
            # 解析参数常量/依赖
            try:
                param_constants, param_deps = self._parse_args_for_generation(key, func_info, full_match_to_key, app)
            except Exception as e:
                logger.error(f"解析参数失败 {key}: {str(e)}")
                error_occurred = True
                error_message = f"解析参数失败：{key} - {str(e)}"
                break
            # 检查参数中是否有错误标记
            if any(p == ERROR_MARKER for p in param_constants):
                error_occurred = True
                error_message = f"参数中包含错误标记：{key}"
                break
            missing_deps = [d for d in param_deps if d is not None and d not in samples_dict]
            if missing_deps:
                logger.error(f"依赖键 {missing_deps} 尚未生成，无法生成 {key}")
                error_occurred = True
                error_message = f"依赖键缺失：{key} 依赖 {missing_deps}"
                break
            samples = np.zeros(n_iterations, dtype=object)
            original_func_name = func_info.get('func_name', '')
            func_name_lower = original_func_name.lower()
            base_markers = markers.copy()
            # 处理数组参数（Cumul/Discrete）
            if original_func_name == 'DriskCumul':
                x_str, p_str = self._parse_cumul_4args(func_info.get('args_text', ''), app)
                if x_str and p_str:
                    base_markers['x_vals'] = x_str
                    base_markers['p_vals'] = p_str
                else:
                    logger.error(f"无法解析 {key} 的数组参数")
                    error_occurred = True
                    error_message = f"无法解析 {key} 的数组参数"
                    break
            elif original_func_name == 'DriskDiscrete':
                x_str, p_str = self._parse_cumul_discrete_args(func_info.get('args_text', ''), app)
                if x_str and p_str:
                    base_markers['x_vals'] = x_str
                    base_markers['p_vals'] = p_str
                else:
                    logger.error(f"无法解析 {key} 的数组参数")
                    error_occurred = True
                    error_message = f"无法解析 {key} 的数组参数"
                    break
            elif original_func_name == 'DriskDUniform':
                x_str, p_str = self._parse_duniform_args(func_info.get('args_text', ''), app)
                if x_str and p_str:
                    base_markers['x_vals'] = x_str
                    base_markers['p_vals'] = p_str
                else:
                    logger.error(f"无法解析 {key} 的 DUniform 数组参数")
                    error_occurred = True
                    error_message = f"无法解析 {key} 的 DUniform 数组参数"
                    break
            trunc_info = _parse_truncate_from_func(func_info)
            if trunc_info:
                trunc_type, trunc_value = trunc_info
                if trunc_type not in base_markers:
                    base_markers[trunc_type] = trunc_value
            base_markers['is_nested'] = func_info.get('is_nested', False)
            if func_info.get('is_at_function'):
                base_markers['is_at_function'] = True
            # 确定分布类型
            dist_type = 'normal'
            if 'duniform' in func_name_lower:
                dist_type = 'duniform'
            elif 'frechet' in func_name_lower:
                dist_type = 'frechet'
            elif 'johnsonsb' in func_name_lower:
                dist_type = 'johnsonsb'
            elif 'johnsonsu' in func_name_lower:
                dist_type = 'johnsonsu'
            elif 'kumaraswamy' in func_name_lower:
                dist_type = 'kumaraswamy'
            elif 'laplace' in func_name_lower:
                dist_type = 'laplace'
            elif 'loglogistic' in func_name_lower:
                dist_type = 'loglogistic'
            elif 'lognorm2' in func_name_lower:
                dist_type = 'lognorm2'
            elif 'betageneral' in func_name_lower:
                dist_type = 'betageneral'
            elif 'betasubj' in func_name_lower:
                dist_type = 'betasubj'
            elif 'burr12' in func_name_lower:
                dist_type = 'burr12'
            elif 'compound' in func_name_lower:
                dist_type = 'compound'
            elif 'splice' in func_name_lower:
                dist_type = 'splice'
            elif 'pert' in func_name_lower:
                dist_type = 'pert'
            elif 'reciprocal' in func_name_lower:
                dist_type = 'reciprocal'
            elif 'rayleigh' in func_name_lower:
                dist_type = 'rayleigh'
            elif 'weibull' in func_name_lower:
                dist_type = 'weibull'
            elif 'pearson5' in func_name_lower:
                dist_type = 'pearson5'
            elif 'pearson6' in func_name_lower:
                dist_type = 'pearson6'
            elif 'pareto2' in func_name_lower:
                dist_type = 'pareto2'
            elif 'pareto' in func_name_lower:
                dist_type = 'pareto'
            elif 'lognorm' in func_name_lower:
                dist_type = 'lognorm'
            elif 'logistic' in func_name_lower:
                dist_type = 'logistic'
            elif 'levy' in func_name_lower:
                dist_type = 'levy'
            elif 'histogrm' in func_name_lower:
                dist_type = 'histogrm'
            elif 'hypsecant' in func_name_lower:
                dist_type = 'hypsecant'
            elif 'general' in func_name_lower:
                dist_type = 'general'
            elif 'fatiguelife' in func_name_lower:
                dist_type = 'fatiguelife'
            elif 'extvaluemin' in func_name_lower:
                dist_type = 'extvaluemin'
            elif 'extvalue' in func_name_lower:
                dist_type = 'extvalue'
            elif 'invgauss' in func_name_lower:
                dist_type = 'invgauss'
            elif 'intuniform' in func_name_lower:
                dist_type = 'intuniform'
            elif 'erlang' in func_name_lower:
                dist_type = 'erlang'
            elif 'erf' in func_name_lower:
                dist_type = 'erf'
            elif 'doubletriang' in func_name_lower:
                dist_type = 'doubletriang'
            elif 'cauchy' in func_name_lower:
                dist_type = 'cauchy'
            elif 'dagum' in func_name_lower:
                dist_type = 'dagum'
            elif 'uniform' in func_name_lower:
                dist_type = 'uniform'
            elif 'gamma' in func_name_lower:
                dist_type = 'gamma'
            elif 'poisson' in func_name_lower:
                dist_type = 'poisson'
            elif 'beta' in func_name_lower:
                dist_type = 'beta'
            elif 'chisq' in func_name_lower:
                dist_type = 'chisq'
            elif 'f' in func_name_lower:
                dist_type = 'f'
            elif 'student' in func_name_lower:
                dist_type = 'student'
            elif 'expon' in func_name_lower:
                dist_type = 'expon'
            elif 'bernoulli' in func_name_lower:
                dist_type = 'bernoulli'
            elif 'negbin' in func_name_lower:
                dist_type = 'negbin'
            elif 'geomet' in func_name_lower:
                dist_type = 'geomet'
            elif 'hypergeo' in func_name_lower:
                dist_type = 'hypergeo'
            elif 'triang' in func_name_lower:
                dist_type = 'triang'
            elif 'binomial' in func_name_lower:
                dist_type = 'binomial'
            elif 'trigen' in func_name_lower:
                dist_type = 'trigen'
            elif 'cumul' in func_name_lower:
                dist_type = 'cumul'
            elif 'discrete' in func_name_lower:
                dist_type = 'discrete'
            for i in range(n_iterations):
                if progress and progress.is_cancelled():
                    error_occurred = True
                    error_message = "用户取消"
                    break
                params = []
                try:
                    for j in range(len(param_constants)):
                        if param_constants[j] is not None:
                            params.append(param_constants[j])
                        else:
                            dep_key = param_deps[j]
                            dep_samples = samples_dict[dep_key]
                            if not isinstance(dep_samples, np.ndarray):
                                raise TypeError(f"依赖键 {dep_key} 的样本不是数组，实际类型: {type(dep_samples)}")
                            if i >= len(dep_samples):
                                raise IndexError(f"依赖键 {dep_key} 的样本长度不足: 需要索引 {i}，但长度只有 {len(dep_samples)}")
                            val = dep_samples[i]
                            if val == ERROR_MARKER or (isinstance(val, str) and val.upper() == "#ERROR!"):
                                params.append(ERROR_MARKER)
                            else:
                                try:
                                    val_f = float(val)
                                    if np.isnan(val_f) or np.isinf(val_f):
                                        raise ValueError("无效数值")
                                    params.append(val_f)
                                except:
                                    params.append(ERROR_MARKER)
                except (IndexError, KeyError, TypeError, ValueError) as e:
                    error_occurred = True
                    error_message = f"生成第{i}个样本时访问依赖值失败：{key} - {str(e)}，参数索引{j}，依赖键{dep_key}"
                    break
                if any(p == ERROR_MARKER for p in params):
                    samples[i] = ERROR_MARKER
                    error_occurred = True
                    error_message = f"生成第{i}个样本时参数中包含错误标记：{key} - 参数：{params}"
                    break
                try:
                    # 每个迭代使用不同的种子：final_base_seed + i
                    iter_seed = (final_base_seed + i) & 0x7fffffff
                    generator = TruncatableDistributionGenerator(original_func_name, dist_type, params, base_markers.copy(), key, seed=iter_seed)
                    one_sample = generator.generate_samples(1)[0]
                    if one_sample == ERROR_MARKER:
                        error_occurred = True
                        error_message = f"生成第{i}个样本失败（参数无效）：{key} - 参数：{params}"
                        break
                    samples[i] = one_sample
                except Exception as e:
                    logger.error(f"生成第{i}个样本失败 {key}: {str(e)}")
                    error_occurred = True
                    error_message = f"生成第{i}个样本失败：{key} - {str(e)}，参数：{params}"
                    break
            if error_occurred:
                break
            samples_dict[key] = samples
            logger.debug(f"逐迭代生成样本 (有依赖): {key}")
        if error_occurred:
            return samples_dict, error_occurred, error_message
        # 所有样本生成完毕，注册到隐藏表和模拟对象
        for key, samples in samples_dict.items():
            if '!' in key:
                sheet_name, cell_with_index = key.split('!', 1)
            else:
                sheet_name = app.ActiveSheet.Name
                cell_with_index = key
            # 注册到隐藏表
            try:
                column_letter = self.register_input(key, samples)
            except Exception as e:
                logger.error(f"注册输入到隐藏表失败 {key}: {str(e)}")
                error_occurred = True
                error_message = f"注册输入失败：{key} - {str(e)}"
                break
            pure_key = cell_with_index
            markers = self.input_key_to_func_info[key].get('markers', {})
            if not sim.add_input_result(pure_key, samples, sheet_name, markers):
                logger.warning(f"添加Input结果失败 {key}")
        return samples_dict, error_occurred, error_message
    # ========== 原有方法：calculate_outputs_and_makeinputs ==========
    def calculate_outputs_and_makeinputs(self, app):
        """计算输出值和MakeInput值（使用批量读取优化）"""
        try:
            logger.info(f"开始计算输出值和MakeInput值，共 {self.total_iterations} 次迭代...")
            # 初始化输出数据
            for cell_addr in self.output_cells:
                self.output_data[cell_addr] = np.full(self.total_iterations, np.nan, dtype=object)
            # 初始化MakeInput数据
            for cell_addr in self.makeinput_cells:
                self.makeinput_data[cell_addr] = np.full(self.total_iterations, np.nan, dtype=object)
            if not self.output_refs or not self.output_refs_range:
                logger.error("输出引用未设置，无法批量读取")
                return False
            for iteration in range(self.total_iterations):
                if self.progress:
                    if self.progress.check_esc_key():
                        logger.info(f"模拟在第 {iteration} 次迭代时被取消")
                        break
                    self.progress.update(iteration)
                if self.control_cell:
                    self.control_cell.Value = iteration + 1
                else:
                    self.hidden_sheet.Cells(2, 1).Value = iteration + 1
                try:
                    app.Calculate()
                except Exception as e:
                    logger.error(f"Excel计算失败: {str(e)}")
                try:
                    values = self.output_refs_range.Value2
                except Exception as e:
                    logger.error(f"读取输出引用区域失败: {str(e)}")
                    for idx, (cell_addr, val_type) in enumerate(self.output_refs):
                        if val_type == 'output':
                            self.output_data[cell_addr][iteration] = ERROR_MARKER
                        else:
                            self.makeinput_data[cell_addr][iteration] = ERROR_MARKER
                    continue
                if values is None:
                    val_list = [np.nan] * len(self.output_refs)
                elif isinstance(values, (list, tuple)) and len(values) > 0:
                    if isinstance(values[0], (list, tuple)):
                        val_list = values[0]
                    else:
                        val_list = values
                else:
                    val_list = [values]
                if len(val_list) != len(self.output_refs):
                    logger.warning(f"读取的值数量({len(val_list)})与引用数量({len(self.output_refs)})不匹配，进行填充/截断")
                    if len(val_list) > len(self.output_refs):
                        val_list = val_list[:len(self.output_refs)]
                    else:
                        val_list = val_list + [ERROR_MARKER] * (len(self.output_refs) - len(val_list))
                for idx, (cell_addr, val_type) in enumerate(self.output_refs):
                    raw_val = val_list[idx]
                    processed_val = self._process_cell_value(raw_val)
                    if val_type == 'output':
                        self.output_data[cell_addr][iteration] = processed_val
                    else:
                        self.makeinput_data[cell_addr][iteration] = processed_val
                if iteration % 500 == 0 and iteration > 0:
                    logger.info(f"  已完成 {iteration}/{self.total_iterations} 次迭代")
            completed_iterations = min(iteration + 1, self.total_iterations)
            logger.info(f"输出值和MakeInput值计算完成，共处理 {completed_iterations} 次迭代")
            return True
        except Exception as e:
            logger.error(f"计算输出值和MakeInput值失败: {str(e)}")
            traceback.print_exc()
            return False
    def _process_cell_value(self, value):
        """处理单元格值，检测Excel原生错误值"""
        if value is None:
            return np.nan
        elif isinstance(value, (int, float, np.number)):
            if isinstance(value, (int, float)) and -2146826300 <= value <= -2146826200:
                return ERROR_MARKER
            return float(value)
        elif isinstance(value, str):
            str_value = str(value).strip()
            excel_errors = ["#ERROR!", "#VALUE!", "#N/A", "#REF!", "#DIV/0!", "#NAME?", "#NULL!", "#NUM!"]
            if str_value.upper() in excel_errors:
                return ERROR_MARKER
            else:
                try:
                    return float(str_value)
                except ValueError:
                    return str_value
        elif isinstance(value, bool):
            return 1.0 if value else 0.0
        elif isinstance(value, (datetime.datetime, datetime.date)):
            try:
                if isinstance(value, datetime.datetime):
                    return float(value.timestamp())
                else:
                    return float(value.toordinal())
            except:
                return ERROR_MARKER
        else:
            try:
                return str(value)
            except:
                return ERROR_MARKER
    def complete_simulation(self, app):
        """完成模拟"""
        with self._lock:
            self.simulation_end_time = time.time()
            self.is_running = False
            logger.info("正在完成模拟...")
            if self.original_formulas or self.original_simtable_values:
                self.restore_original_formulas_and_values(app)
            self.restore_excel_settings(app)
            if self.progress:
                self.progress.clear()
            logger.info("模拟完成")
    def clear_simulation(self, app):
        """清除模拟数据和隐藏表"""
        try:
            logger.info("清除模拟数据...")
            try:
                hidden_sheet = app.ActiveWorkbook.Worksheets("DriskIndexHidden")
                hidden_sheet.Visible = True
                hidden_sheet.Delete()
                logger.info("删除DriskIndexHidden工作表")
            except Exception as e:
                logger.info("未找到DriskIndexHidden工作表或删除失败")
            self._init_manager()
            try:
                app.StatusBar = False
            except:
                pass
            return True
        except Exception as e:
            logger.error(f"清除模拟数据失败: {str(e)}")
            return False
# 全局模拟管理器实例
_sim_manager = IndexSimulationManager()
# ==================== 辅助函数 ====================
def _extract_cross_sheet_references(formula: str) -> List[str]:
    """提取公式中的跨sheet引用"""
    if not formula or not isinstance(formula, str) or not formula.startswith('='):
        return []
    pattern = r'(?:[\'\"]([^\'\"]+)[\'\"]!|[A-Za-z0-9_]+!)(?:[\$]?[A-Za-z]+[\$]?\d+(?::[\$]?[A-Za-z]+[\$]?\d+)?|\$?[A-Za-z]+\$?\d+)'
    matches = re.findall(pattern, formula, re.IGNORECASE)
    refs = []
    for match in matches:
        if match:
            refs.append(match)
    return refs
def _count_distributions_in_cell(formula: str) -> int:
    """统计单元格中分布函数的数量"""
    try:
        from formula_parser import extract_nested_distributions_advanced
        dist_funcs = _filter_compound_nested_functions(
            extract_nested_distributions_advanced(formula, "temp")
        )
        return len(dist_funcs)
    except:
        count = 0
        dist_patterns = ['DRISKNORMAL', 'DRISKUNIFORM', 'DRISKGAMMA', 'DRISKPOISSON']
        for pattern in dist_patterns:
            count += formula.upper().count(pattern)
        return count
def _parse_truncate_from_func(func: Dict) -> Optional[Tuple[str, str]]:
    """从函数字典中解析截断信息"""
    trunc_types = ['truncate', 'truncate2', 'truncatep', 'truncatep2']
    for ttype in trunc_types:
        if ttype in func.get('markers', {}):
            return (ttype, func['markers'][ttype])
    args_text = func.get('args_text', '')
    full_match = func.get('full_match', '')
    patterns = {
        'truncate': r'DriskTruncate\s*\(([^)]+)\)',
        'truncate2': r'DriskTruncate2\s*\(([^)]+)\)',
        'truncatep': r'DriskTruncateP\s*\(([^)]+)\)',
        'truncatep2': r'DriskTruncateP2\s*\(([^)]+)\)',
    }
    for ttype, pattern in patterns.items():
        match = re.search(pattern, args_text, re.IGNORECASE)
        if match:
            value_str = match.group(1).strip()
            return (ttype, value_str)
        match = re.search(pattern, full_match, re.IGNORECASE)
        if match:
            value_str = match.group(1).strip()
            return (ttype, value_str)
    return None
def normalize_cell_address(addr: str, app=None) -> str:
    """规范化单元格地址，统一工作表名为大写"""
    if not isinstance(addr, str):
        return addr
    if '!' in addr:
        parts = addr.split('!')
        if len(parts) > 2:
            sheet = parts[0]
            cell = parts[-1]
        else:
            sheet, cell = parts
        return f"{sheet.upper()}!{cell}"
    else:
        if app is None:
            return addr
        try:
            current_sheet = app.ActiveSheet.Name
            return f"{current_sheet.upper()}!{addr}"
        except:
            return addr
# ==================== 新增函数：检查 Simtable 单元格是否被任何分布函数引用 ====================  # NEW
def _check_simtable_input_dependency(simtable_cell_addr: str, distribution_cells: Dict[str, List[Dict]]) -> bool:
    """
    检查给定的 Simtable 单元格是否被任何分布函数单元格引用。
    如果被至少一个分布函数引用，返回 True；否则返回 False。
    """
    # 遍历所有分布函数单元格
    for cell_addr, dist_funcs in distribution_cells.items():
        for dist_func in dist_funcs:
            # 检查参数文本中是否包含 simtable_cell_addr
            args_text = dist_func.get('args_text', '')
            # 提取参数文本中的单元格引用（简单匹配）
            if simtable_cell_addr.upper() in args_text.upper():
                return True
            # 也可以使用正则提取完整引用
            refs = parse_formula_references(f"={dist_func.get('full_match', '')}")
            for ref in refs:
                if ref.upper() == simtable_cell_addr.upper():
                    return True
    return False
# ==================== 主模拟函数 ====================
def run_index_simulation(n_iterations: int, scenario_count: int = 1):
    """运行索引模拟 - 修复版，确保注册与查找一致，支持内嵌分布，每次模拟前删除旧隐藏表"""
    app = None
    try:
        clear_simulations()
        try:
            from com_fixer import _safe_excel_app
            app = _safe_excel_app()
        except ImportError:
            import win32com.client
            app = win32com.client.Dispatch("Excel.Application")
        _sim_manager._init_manager()
        # ========== 设置基础时间种子 ==========
        _sim_manager.base_time_seed = int(time.time() * 1000) & 0x7fffffff
        _sim_manager.simulation_id = int(time.time() % 1000000)
        _sim_manager.total_iterations = n_iterations
        _sim_manager.scenario_count = scenario_count
        _sim_manager.simulation_name = f"Index_MC_{n_iterations}_{_sim_manager.simulation_id}"
        _sim_manager.simulation_start_time = time.time()
        _sim_manager.is_running = True
        logger.info(f"开始索引模拟: {n_iterations}次迭代, {scenario_count}个场景")
        app.StatusBar = "开始索引蒙特卡洛模拟..."
        _sim_manager.save_excel_settings(app)
        _sim_manager.optimize_excel_performance(app)
        app.StatusBar = "查找工作表中的单元格..."
        logger.info("查找模拟单元格...")
        try:
            distribution_cells, simtable_cells, makeinput_cells, output_cells, all_input_keys, all_related_cells = find_all_simulation_cells_in_workbook(app)
        except Exception as e:
            logger.error(f"使用dependency_tracker查找单元格失败: {str(e)}")
            from formula_parser import extract_nested_distributions_advanced, extract_simtable_functions, extract_makeinput_functions, extract_output_info
            distribution_cells = {}
            simtable_cells = {}
            makeinput_cells = {}
            output_cells = {}
            all_input_keys = []
            workbook = app.ActiveWorkbook
            for sheet in workbook.Worksheets:
                # 只处理可见工作表
                if sheet.Visible != -1:
                    continue
                used_range = sheet.UsedRange
                for cell in used_range:
                    try:
                        formula = cell.Formula
                        if not isinstance(formula, str) or not formula.startswith('='):
                            continue
                        address = f"{sheet.Name}!{cell.Address.replace('$', '')}"
                        if is_output_cell(formula):
                            output_info = extract_output_info(formula)
                            output_cells[address] = output_info
                        if is_distribution_function(formula):
                            dist_funcs = _filter_compound_nested_functions(
                                extract_nested_distributions_advanced(formula, address)
                            )
                            if dist_funcs:
                                distribution_cells[address] = dist_funcs
                        if is_makeinput_function(formula):
                            makeinput_funcs = extract_makeinput_functions(formula)
                            if makeinput_funcs:
                                makeinput_cells[address] = makeinput_funcs
                        if is_simtable_function(formula):
                            simtable_funcs = extract_simtable_functions(formula)
                            if simtable_funcs:
                                simtable_cells[address] = simtable_funcs
                    except:
                        continue
        def normalize_and_merge(cell_dict, app):
            normalized = {}
            for addr, value in cell_dict.items():
                norm_addr = normalize_cell_address(addr, app)
                normalized[norm_addr] = value
            return normalized
        distribution_cells = normalize_and_merge(distribution_cells, app)
        distribution_cells = {
            addr: filtered_funcs
            for addr, funcs in distribution_cells.items()
            if (filtered_funcs := _filter_compound_nested_functions(funcs))
        }
        simtable_cells = normalize_and_merge(simtable_cells, app)
        makeinput_cells = normalize_and_merge(makeinput_cells, app)
        output_cells = normalize_and_merge(output_cells, app)
        if len(distribution_cells) == 0 and len(makeinput_cells) == 0:
            app.StatusBar = False
            _sim_manager.restore_excel_settings(app)
            xlcAlert("❌ 未找到分布函数或MakeInput函数单元格")
            return
        if len(output_cells) == 0:
            app.StatusBar = False
            _sim_manager.restore_excel_settings(app)
            xlcAlert("❌ 未找到输出单元格")
            return
        _sim_manager.distribution_cells = distribution_cells
        _sim_manager.simtable_cells = simtable_cells
        _sim_manager.makeinput_cells = makeinput_cells
        _sim_manager.output_cells = output_cells
        _sim_manager.cell_dist_count = {}
        for cell_addr, dist_funcs in distribution_cells.items():
            _sim_manager.cell_dist_count[cell_addr] = len(dist_funcs)
            logger.info(f"单元格 {cell_addr} 包含 {len(dist_funcs)} 个分布函数")
        _sim_manager.cross_sheet_refs = {}
        for cell_addr, dist_funcs in distribution_cells.items():
            if '!' in cell_addr:
                sheet_name, addr = cell_addr.split('!')
                try:
                    sheet = app.ActiveWorkbook.Worksheets(sheet_name)
                    cell = sheet.Range(addr)
                    formula = cell.Formula
                    refs = _extract_cross_sheet_references(formula)
                    if refs:
                        _sim_manager.cross_sheet_refs[cell_addr] = refs
                except:
                    pass
        total_distribution_functions = sum(len(v) for v in distribution_cells.values())
        total_makeinput_functions = sum(len(v) for v in makeinput_cells.values())
        nested_dist_count = 0
        for cell_addr, dist_funcs in distribution_cells.items():
            for func in dist_funcs:
                if func.get('is_nested', False):
                    nested_dist_count += 1
        actual_input_count = total_distribution_functions + total_makeinput_functions
        logger.info(f"找到 {len(distribution_cells)} 个分布单元格, {len(makeinput_cells)} 个MakeInput单元格, {len(output_cells)} 个输出单元格")
        logger.info(f"分布函数: {total_distribution_functions}, 内嵌分布: {nested_dist_count}, MakeInput: {total_makeinput_functions}, 实际Input数据: {actual_input_count}")
        app.StatusBar = "备份原始公式和Simtable值..."
        _sim_manager.backup_original_formulas_and_values(app)
        app.StatusBar = "创建隐藏工作表..."
        if not _sim_manager._ensure_hidden_sheet(app):
            app.StatusBar = False
            _sim_manager.restore_excel_settings(app)
            xlcAlert("❌ 创建隐藏工作表失败")
            return
        _sim_manager.progress = StatusBarProgress(app, n_iterations, scenario_count)
        all_scenario_results = []
        # ==================== 新增：构建 Simtable 单元格与分布函数的依赖关系 ====================  # NEW
        simtable_input_dependency = {}  # {simtable_addr: bool}
        for sim_addr in simtable_cells.keys():
            simtable_input_dependency[sim_addr] = _check_simtable_input_dependency(sim_addr, distribution_cells)
        for scenario_idx in range(scenario_count):
            _sim_manager.current_scenario = scenario_idx
            _sim_manager.simtable_error_cells = []  # 重置错误列表
            if _sim_manager.progress.check_esc_key():
                logger.info("模拟被ESC键取消")
                break
            _sim_manager.progress.start_new_scenario(scenario_idx)
            # ==================== 新增：检查 Simtable 单元格索引是否超出 ====================  # MODIFIED
            any_simtable_out_of_range = False
            for cell_addr, simtable_funcs in simtable_cells.items():
                if simtable_funcs:
                    simtable_func = simtable_funcs[0]
                    values = simtable_func.get('values', [])
                    if scenario_idx >= len(values):
                        # 场景索引超出范围
                        any_simtable_out_of_range = True
                        _sim_manager.simtable_error_cells.append(cell_addr)
                        logger.warning(f"场景 {scenario_idx+1}: Simtable单元格 {cell_addr} 超出范围（最大索引 {len(values)-1}），将设置为#ERROR!")
            # 如果存在超出范围的 Simtable 单元格，检查它们是否被任何分布函数引用
            if any_simtable_out_of_range:
                # 检查是否至少有一个错误 Simtable 被输入依赖
                skip_this_scenario = False
                for err_addr in _sim_manager.simtable_error_cells:
                    if simtable_input_dependency.get(err_addr, False):
                        skip_this_scenario = True
                        logger.warning(f"Simtable单元格 {err_addr} 被分布函数引用，将跳过场景 {scenario_idx+1}")
                        break
                if skip_this_scenario:
                    _sim_manager.skipped_scenarios.append(scenario_idx + 1)  # 记录跳过的场景（1-based）
                    # 跳过当前场景
                    continue
            sim_id = create_simulation(n_iterations, "Index_MC")
            sim = get_simulation(sim_id)
            if sim is None:
                logger.error(f"创建模拟对象失败（场景{scenario_idx+1}）")
                continue
            sim.set_scenario_info(scenario_count, scenario_idx)
            try:
                sim.workbook_name = app.ActiveWorkbook.Name
                logger.info(f"工作簿名称: {sim.workbook_name}")
            except Exception as e:
                logger.error(f"获取工作簿名称失败: {str(e)}")
            sim.distribution_cells = distribution_cells
            sim.simtable_cells = simtable_cells
            sim.makeinput_cells = makeinput_cells
            sim.output_cells = output_cells
            sim.all_input_keys = []
            # 设置Simtable值 - 如果超出范围则设为#ERROR!
            if simtable_cells:
                for cell_addr, simtable_funcs in simtable_cells.items():
                    if simtable_funcs:
                        simtable_func = simtable_funcs[0]
                        values = simtable_func.get('values', [])
                        if scenario_idx < len(values):
                            value = values[scenario_idx]
                        else:
                            # 超出范围，设为错误标记
                            value = ERROR_MARKER
                            # 记录到错误列表（已记录）
                        try:
                            if '!' in cell_addr:
                                sheet_name, addr = cell_addr.split('!')
                                sheet = app.ActiveWorkbook.Worksheets(sheet_name)
                                cell = sheet.Range(addr)
                            else:
                                cell = app.ActiveSheet.Range(cell_addr)
                            # 设置值为错误标记或对应值
                            if value == ERROR_MARKER:
                                cell.Value = ERROR_MARKER
                            else:
                                cell.Value = float(value) if isinstance(value, (int, float)) else value
                        except Exception as e:
                            logger.error(f"设置Simtable单元格 {cell_addr} 失败: {str(e)}")
            # ---------- 为所有分布函数分配 input_key ----------
            _sim_manager.processed_input_keys.clear()
            _sim_manager.input_key_to_func_info.clear()
            _sim_manager.cell_to_input_key = _sim_manager._build_cell_to_input_key_map()  # 更新映射
            for cell_addr, dist_funcs in distribution_cells.items():
                if '!' in cell_addr:
                    sheet_name, cell_addr_only = cell_addr.split('!', 1)
                else:
                    sheet_name = app.ActiveSheet.Name
                    cell_addr_only = cell_addr
                for i, func in enumerate(dist_funcs):
                    index_in_cell = func.get('index', i+1)
                    input_key = f"{cell_addr}_{index_in_cell}"
                    func['input_key'] = input_key
                    # 同时更新 func['key'] 以确保依赖解析时能通过单元格引用找到
                    func['key'] = input_key
                    # 填充 input_key_to_func_info
                    params = func.get('parameters', [])
                    func_name = func.get('func_name', '')
                    markers = func.get('markers', {}).copy()
                    # 补充截断信息
                    trunc_info = _parse_truncate_from_func(func)
                    if trunc_info:
                        trunc_type, trunc_value = trunc_info
                        if trunc_type not in markers:
                            markers[trunc_type] = trunc_value
                    _sim_manager.input_key_to_func_info[input_key] = {
                        'func_name': func_name,
                        'args_text': func.get('args_text', ''),
                        'full_match': func.get('full_match', ''),
                        'parameters': params,
                        'markers': markers,
                        'is_nested': func.get('is_nested', False),
                        'is_at_function': func.get('is_at_function', False),
                        'key': input_key,  # 使用统一键
                        'index': index_in_cell
                    }
            # ---------- 解析依赖关系 ----------
            app.StatusBar = "正在解析依赖关系..."
            try:
                graph, sorted_keys = _sim_manager.build_dependency_graph(distribution_cells)
                _sim_manager.dependency_graph = graph
                _sim_manager.sorted_keys = sorted_keys
                logger.info(f"依赖解析完成，拓扑顺序：{sorted_keys}")
            except Exception as e:
                logger.error(f"依赖解析失败: {str(e)}")
                _sim_manager.restore_excel_settings(app)
                _sim_manager.restore_original_formulas_and_values(app)
                xlcAlert(f"依赖解析失败：{str(e)}")
                return
            # ---------- 生成输入样本（按依赖顺序） ----------
            app.StatusBar = "生成输入随机数..."
            samples_dict, error_occurred, error_msg = _sim_manager.generate_samples_with_dependencies(
                n_iterations, sim, app, _sim_manager.progress
            )
            if error_occurred:
                logger.error(f"样本生成错误: {error_msg}")
                if not _sim_manager.progress.was_cancelled_by_esc():
                    xlcAlert(f"模拟中止：{error_msg}")
                # 跳出场景循环，不再计算输出
                break
            # ---------- 设置INDEX公式（包含分布和内嵌） ----------
            app.StatusBar = "设置INDEX公式..."
            if not _sim_manager.setup_index_formulas(app):
                logger.warning("警告: 设置INDEX公式时出现问题")
            # ---------- 设置输出引用公式（已移至输入列右侧） ----------
            app.StatusBar = "设置输出引用公式..."
            if not _sim_manager.setup_output_references(app):
                logger.warning("警告: 设置输出引用公式时出现问题")
            # ---------- 计算输出值和MakeInput值 ----------
            app.StatusBar = "计算输出值和MakeInput值..."
            if not _sim_manager.calculate_outputs_and_makeinputs(app):
                logger.warning("警告: 计算输出值和MakeInput值时出现问题")
            # 保存输出数据到模拟对象
            logger.info(f"保存场景 {scenario_idx+1} 的输出数据...")
            output_saved_count = 0
            for cell_addr, data in _sim_manager.output_data.items():
                if '!' in cell_addr:
                    sheet_name, cell_addr_only = cell_addr.split('!')
                else:
                    sheet_name = app.ActiveSheet.Name
                    cell_addr_only = cell_addr
                out_info = {}
                if cell_addr in output_cells:
                    out_info = output_cells[cell_addr].copy()
                if 'name' not in out_info:
                    out_info['name'] = cell_addr_only
                if 'cell_address' not in out_info:
                    out_info['cell_address'] = cell_addr
                if sim.add_output_result(cell_addr_only, data, sheet_name, out_info):
                    output_saved_count += 1
            # 保存MakeInput数据到模拟对象（更新之前占位的NaN）
            logger.info(f"保存场景 {scenario_idx+1} 的MakeInput数据...")
            makeinput_saved_count = 0
            for cell_addr, data in _sim_manager.makeinput_data.items():
                if '!' in cell_addr:
                    sheet_name, cell_addr_only = cell_addr.split('!')
                else:
                    sheet_name = app.ActiveSheet.Name
                    cell_addr_only = cell_addr
                input_key = f"{cell_addr_only}_MakeInput"
                # 获取之前注册的属性
                attrs = sim.input_attributes.get(f"{sheet_name}!{input_key}", {})
                if not attrs:
                    full_formula = _sim_manager.original_formulas.get(cell_addr, '')
                    if full_formula:
                        attrs = extract_makeinput_attributes(full_formula)
                    attrs['is_makeinput'] = True
                # 更新模拟对象中的值（替换占位的 NaN）
                if sim.add_input_result(input_key, data, sheet_name, attrs):
                    makeinput_saved_count += 1
            # 预计算统计量
            logger.info(f"计算场景 {scenario_idx+1} 统计量...")
            sim.compute_input_statistics()
            sim.compute_output_statistics()
            sim.set_end_time()
            all_scenario_results.append({
                'sim_id': sim_id,
                'scenario_index': scenario_idx,
                'output_saved_count': output_saved_count,
                'makeinput_saved_count': makeinput_saved_count
            })
            logger.info(f"场景 {scenario_idx+1} 模拟完成: Output={output_saved_count}, MakeInput={makeinput_saved_count}")
        _sim_manager.complete_simulation(app)
        total_elapsed = time.time() - _sim_manager.simulation_start_time
        completed_scenarios = len(all_scenario_results)
        total_dist_functions = sum(len(v) for v in _sim_manager.distribution_cells.values())
        total_makeinput = sum(len(v) for v in _sim_manager.makeinput_cells.values())
        nested_count = sum(1 for v in _sim_manager.distribution_cells.values() for func in v if func.get('is_nested', False))
        actual_input_count = total_dist_functions + total_makeinput
        # 构建结果消息，增加跳过的场景信息
        if _sim_manager.progress and _sim_manager.progress.was_cancelled_by_esc():
            alert_msg = (
                f"索引蒙特卡洛模拟已取消 (ESC键停止)!\n\n"
                f"已完成场景: {completed_scenarios}/{scenario_count}\n"
                f"每场景迭代次数: {n_iterations:,}\n"
                f"总已完成迭代: {completed_scenarios * n_iterations:,}\n"
                f"分布函数: {total_dist_functions}\n"
                f"内嵌分布: {nested_count}\n"
                f"MakeInput: {total_makeinput}\n"
                f"实际Input数据: {actual_input_count}\n"
                f"输出单元格: {len(_sim_manager.output_cells)}\n"
                f"抽样方法: 索引蒙特卡洛\n"
                f"总耗时: {total_elapsed:.2f}秒\n"
            )
        else:
            alert_msg = (
                f"索引蒙特卡洛模拟完成!\n\n"
                f"场景数: {scenario_count}\n"
                f"每场景迭代次数: {n_iterations:,}\n"
                f"总迭代次数: {scenario_count * n_iterations:,}\n"
                f"分布函数: {total_dist_functions}\n"
                f"内嵌分布: {nested_count}\n"
                f"MakeInput: {total_makeinput}\n"
                f"实际Input数据: {actual_input_count}\n"
                f"输出单元格: {len(_sim_manager.output_cells)}\n"
                f"抽样方法: 索引蒙特卡洛\n"
                f"总耗时: {total_elapsed:.2f}秒\n"
                f"平均速度: {scenario_count * n_iterations/total_elapsed:.1f} 迭代/秒\n"
            )
        # 如果有跳过的场景，在消息中显示
        if hasattr(_sim_manager, 'skipped_scenarios') and _sim_manager.skipped_scenarios:
            alert_msg += f"\n⚠️ 跳过的场景（因Simtable超出范围且被输入依赖）: {', '.join(map(str, _sim_manager.skipped_scenarios))}\n"
        if all_scenario_results:
            alert_msg += f"\n场景模拟ID:\n"
            for result in all_scenario_results:
                sim_obj = get_simulation(result['sim_id'])
                input_count = len(sim_obj.input_cache) if sim_obj else 0
                output_count = len(sim_obj.output_cache) if sim_obj else 0
                alert_msg += f"  场景{result['scenario_index']+1}: 模拟ID={result['sim_id']} (Input={input_count}, Output={output_count})\n"
        alert_msg += f"\n提示：使用DriskInfo()函数查看具体场景的模拟结果"
        xlcAlert(alert_msg)
    except Exception as e:
        error_msg = f"索引蒙特卡洛模拟失败: {str(e)}"
        logger.error(f"模拟失败: {traceback.format_exc()}")
        if app:
            try:
                app.StatusBar = False
                try:
                    _sim_manager.restore_excel_settings(app)
                except:
                    pass
                try:
                    _sim_manager.restore_original_formulas_and_values(app)
                except:
                    pass
            except:
                pass
        xlcAlert(error_msg)
# ==================== 宏函数 ====================
@xl_macro
def DriskIndexMC():
    """索引蒙特卡洛模拟主宏"""
    try:
        try:
            from com_fixer import _safe_excel_app
            app = _safe_excel_app()
        except ImportError:
            import win32com.client
            app = win32com.client.Dispatch("Excel.Application")
    except Exception as e:
        xlcAlert(f"无法获取 Excel 应用对象: {str(e)}")
        return
    try:
        iter_in = app.InputBox(
            Prompt="请输入模拟次数（默认1000，范围100-100000）：", 
            Title="DriskIndexMC - 迭代次数", 
            Type=1, 
            Default="1000"
        )
        if iter_in is False:
            return
        n_iterations = int(iter_in)
        if n_iterations < 100 or n_iterations > 100000:
            xlcAlert("模拟次数必须在100到100000之间")
            return
    except Exception as e:
        xlcAlert(f"无效的迭代次数: {str(e)}")
        return
    try:
        logger.info("查找模拟单元格...")
        distribution_cells, simtable_cells, makeinput_cells, output_cells, all_input_keys, _ = find_all_simulation_cells_in_workbook(app)
        scenario_count = 1
        if simtable_cells:
            max_simtable_length = 0
            for cell_addr, simtable_funcs in simtable_cells.items():
                for simtable_func in simtable_funcs:
                    values = simtable_func.get('values', [])
                    if len(values) > max_simtable_length:
                        max_simtable_length = len(values)
            if max_simtable_length > 0:
                try:
                    scenario_in = app.InputBox(
                        Prompt=f"检测到Simtable函数，最大数组长度为{max_simtable_length}。请输入场景数（默认{max_simtable_length}，范围1-100）：", 
                        Title="DriskIndexMC - 场景数", 
                        Type=1, 
                        Default=str(max_simtable_length)
                    )
                    if scenario_in is False:
                        return
                    scenario_count = int(scenario_in)
                    if scenario_count < 1 or scenario_count > 100:
                        xlcAlert("场景数必须在1到100之间")
                        return
                except Exception as e:
                    scenario_count = 1
                    logger.error(f"获取场景数失败，使用默认值1: {str(e)}")
    except Exception as e:
        logger.error(f"查找Simtable单元格失败: {str(e)}")
        scenario_count = 1
    run_index_simulation(n_iterations, scenario_count)
@xl_macro
def DriskClearIndex():
    """清除索引模拟数据"""
    try:
        try:
            from com_fixer import _safe_excel_app
            app = _safe_excel_app()
        except ImportError:
            import win32com.client
            app = win32com.client.Dispatch("Excel.Application")
        _sim_manager.clear_simulation(app)
        xlcAlert("索引模拟数据已清除")
    except Exception as e:
        xlcAlert(f"清除索引模拟数据失败: {str(e)}")
def get_index_simulation_manager():
    """获取索引模拟管理器实例"""
    return _sim_manager
def clear_index_simulation(app):
    """清除索引模拟"""
    return _sim_manager.clear_simulation(app)
