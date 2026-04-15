# distribution_fit.py
# -*- coding: utf-8 -*-
"""
分布拟合模块 - 为 Excel 提供 DriskFitDist、DriskFitInfo 和 DriskFitBatch 宏
依赖于同目录下的 distribution_functions.py 和 constants.py
支持：
- 连续分布拟合（样本数据）
- 离散分布拟合（样本数据）
- 连续密度拟合（未标准化 / 标准化）
- 连续累积拟合
- 批量拟合多列数据并生成相关性矩阵
"""

import sys
import os
import re
import math
import traceback
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
from typing import List, Tuple, Dict, Any, Optional, Callable
import numpy as np
import scipy.stats as sps
from scipy.special import gammaln
from scipy.optimize import minimize, least_squares, minimize_scalar
import warnings

# 抑制 scipy 拟合时的数值警告
warnings.filterwarnings("ignore", category=RuntimeWarning)
warnings.filterwarnings("ignore", category=UserWarning)

# 尝试导入 pyxll
try:
    from pyxll import xl_macro, xl_app
    PYXLL_AVAILABLE = True
except ImportError:
    PYXLL_AVAILABLE = False
    def xl_macro(func):
        return func

# 导入同目录下的 distribution_functions 和 constants 模块
current_dir = os.path.dirname(os.path.abspath(__file__))
if current_dir not in sys.path:
    sys.path.insert(0, current_dir)

try:
    import distribution_functions as df
    from distribution_functions import (
        DistributionGenerator, ERROR_MARKER,
        NormalDistribution, UniformDistribution, ErfDistribution,
        GammaDistribution, ErlangDistribution, BetaDistribution,
        ChiSquaredDistribution, FDistribution, TDistribution,
        ExponentialDistribution, TriangDistribution, CauchyDistribution,
        DagumDistribution, DoubleTriangDistribution, ExtvalueDistribution,
        ExtvalueMinDistribution, FatigueLifeDistribution, FrechetDistribution,
        HypSecantDistribution, JohnsonSBDistribution, JohnsonSUDistribution,
        KumaraswamyDistribution, LaplaceDistribution, LogisticDistribution,
        LoglogisticDistribution, LognormDistribution, Lognorm2Distribution,
        BetaGeneralDistribution, BetaSubjDistribution, Burr12Distribution,
        PertDistribution, ReciprocalDistribution, RayleighDistribution,
        WeibullDistribution, Pearson5Distribution, Pearson6Distribution,
        Pareto2Distribution, ParetoDistribution, LevyDistribution,
        GeneralDistribution, HistogrmDistribution,
        PoissonDistribution, BernoulliDistribution, BinomialDistribution,
        NegbinDistribution, InvgaussDistribution, DUniformDistribution,
        GeometDistribution, HypergeoDistribution, IntuniformDistribution,
        TrigenDistribution, CumulDistribution, DiscreteDistribution,
    )
    DIST_FUNCS_AVAILABLE = True
except ImportError as e:
    DIST_FUNCS_AVAILABLE = False
    print(f"警告: 无法导入 distribution_functions 模块: {e}")

# 导入 constants 模块以使用参数验证和支撑集
try:
    import constants
    CONSTANTS_AVAILABLE = True
except ImportError:
    CONSTANTS_AVAILABLE = False
    print("警告: 无法导入 constants 模块，参数验证将受限")

# 导入 corrmat_functions 模块以复用矩阵写入和命名区域管理
try:
    import corrmat_functions as cmf
    CORRMAT_AVAILABLE = True
except ImportError as e:
    CORRMAT_AVAILABLE = False
    print(f"警告: 无法导入 corrmat_functions 模块: {e}")

# 连续分布列表
CONTINUOUS_DIST_NAMES = [
    'normal', 'uniform', 'erf', 'gamma', 'erlang', 'beta', 'chisq', 'f', 'student',
    'expon', 'triang', 'cauchy', 'dagum', 'doubletriang', 'extvalue', 'extvaluemin',
    'fatiguelife', 'frechet', 'hypsecant', 'johnsonsb', 'johnsonsu', 'kumaraswamy',
    'laplace', 'logistic', 'loglogistic', 'lognorm', 'lognorm2', 'betageneral',
    'betasubj', 'burr12', 'pert', 'reciprocal', 'rayleigh', 'weibull', 'pearson5',
    'pearson6', 'pareto2', 'pareto', 'levy', 'trigen', 'invgauss'
]

DISCRETE_DISTRIBUTIONS = [
    'poisson', 'bernoulli', 'binomial', 'negbin', 'geomet', 'hypergeo', 'intuniform'
]

# 映射分布名称到 Drisk 函数名
DIST_TO_DRISK_NAME = {
    'normal': 'DriskNormal', 'uniform': 'DriskUniform', 'erf': 'DriskErf',
    'gamma': 'DriskGamma', 'erlang': 'DriskErlang', 'beta': 'DriskBeta',
    'poisson': 'DriskPoisson', 'chisq': 'DriskChisq', 'f': 'DriskF',
    'student': 'DriskStudent', 'expon': 'DriskExpon', 'bernoulli': 'DriskBernoulli',
    'triang': 'DriskTriang', 'binomial': 'DriskBinomial', 'cauchy': 'DriskCauchy',
    'dagum': 'DriskDagum', 'doubletriang': 'DriskDoubleTriang', 'extvalue': 'DriskExtvalue',
    'extvaluemin': 'DriskExtvalueMin', 'fatiguelife': 'DriskFatigueLife',
    'frechet': 'DriskFrechet', 'hypsecant': 'DriskHypSecant', 'johnsonsb': 'DriskJohnsonSB',
    'johnsonsu': 'DriskJohnsonSU', 'kumaraswamy': 'DriskKumaraswamy', 'laplace': 'DriskLaplace',
    'logistic': 'DriskLogistic', 'loglogistic': 'DriskLogLogistic', 'lognorm': 'DriskLognorm',
    'lognorm2': 'DriskLognorm2', 'betageneral': 'DriskBetaGeneral', 'betasubj': 'DriskBetaSubj',
    'burr12': 'DriskBurr12', 'pert': 'DriskPert', 'reciprocal': 'DriskReciprocal',
    'rayleigh': 'DriskRayleigh', 'weibull': 'DriskWeibull', 'pearson5': 'DriskPearson5',
    'pearson6': 'DriskPearson6', 'pareto2': 'DriskPareto2', 'pareto': 'DriskPareto',
    'levy': 'DriskLevy', 'general': 'DriskGeneral', 'histogrm': 'DriskHistogrm',
    'negbin': 'DriskNegbin', 'invgauss': 'DriskInvgauss', 'duniform': 'DriskDuniform',
    'geomet': 'DriskGeomet', 'hypergeo': 'DriskHypergeo', 'intuniform': 'DriskIntuniform',
    'trigen': 'DriskTrigen', 'cumul': 'DriskCumul', 'discrete': 'DriskDiscrete'
}

FIT_METRICS = ['AICc', 'BIC', 'KS', 'AD', 'AVlog', 'chisq']
_last_fit_result = None

# ================== 辅助函数 ==================
def get_excel_app():
    if PYXLL_AVAILABLE:
        return xl_app()
    return None

def set_status_bar(text: str):
    if PYXLL_AVAILABLE:
        app = xl_app()
        if app:
            try:
                app.StatusBar = text
            except:
                pass

def clear_status_bar():
    if PYXLL_AVAILABLE:
        app = xl_app()
        if app:
            try:
                app.StatusBar = False
            except:
                pass

def parse_excel_range(range_str: str):
    app = get_excel_app()
    if not app:
        return []
    try:
        if '!' in range_str:
            sheet_name, addr = range_str.split('!')
            sheet = app.Sheets(sheet_name)
        else:
            sheet = app.ActiveSheet
            addr = range_str
        rng = sheet.Range(addr)
        values = []
        for cell in rng:
            val = cell.Value
            if val is not None:
                try:
                    values.append(float(val))
                except:
                    continue
        return values
    except Exception as e:
        print(f"解析范围失败: {e}")
        return []

def col_num_to_letter(col: int) -> str:
    letters = ""
    while col > 0:
        col, rem = divmod(col - 1, 26)
        letters = chr(65 + rem) + letters
    return letters

def get_current_column_range_string() -> str:
    app = get_excel_app()
    if not app:
        return ""
    try:
        cell = app.ActiveCell
        col = cell.Column
        row = cell.Row
        sheet = app.ActiveSheet
        start_row = row
        while start_row > 1:
            val = sheet.Cells(start_row - 1, col).Value
            if val is None or not isinstance(val, (int, float)):
                break
            start_row -= 1
        end_row = row
        while True:
            val = sheet.Cells(end_row + 1, col).Value
            if val is None or not isinstance(val, (int, float)):
                break
            end_row += 1
        if end_row < start_row:
            return ""
        col_letter = col_num_to_letter(col)
        sheet_name = sheet.Name
        if " " in sheet_name or "'" in sheet_name:
            sheet_name = "'" + sheet_name.replace("'", "''") + "'"
        return f"{sheet_name}!{col_letter}{start_row}:{col_letter}{end_row}"
    except:
        return ""

def is_numeric_data(data):
    if not data:
        return False
    return all(isinstance(x, (int, float)) for x in data)

def auto_detect_data_type(data):
    if all(float(x).is_integer() for x in data):
        unique_vals = len(set(data))
        if unique_vals <= 0.2 * len(data):
            return "discrete"
    return "continuous"

def _get_excel_input(app, prompt, title, default="", input_type=2):
    try:
        value = app.InputBox(Prompt=prompt, Title=title, Type=input_type, Default=default)
        if value is False:
            return None
        return value
    except Exception:
        if input_type == 8:
            return None
        try:
            return simpledialog.askstring(title, prompt, initialvalue=default)
        except Exception:
            return None

def parse_int_input(value):
    """将各种可能的输入转换为整数，若失败返回None"""
    if value is None:
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, int):
        return value
    if isinstance(value, float):
        if value.is_integer():
            return int(value)
        return None
    try:
        s = str(value).strip()
        if not s:
            return None
        f = float(s)
        if f.is_integer():
            return int(f)
        return None
    except Exception:
        return None

def parse_float_input(value):
    """将各种可能的输入转换为浮点数，若失败返回None"""
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return float(value)
    try:
        s = str(value).strip()
        if not s:
            return None
        return float(s)
    except Exception:
        return None

def _flatten_excel_range_values(value):
    if isinstance(value, tuple):
        for item in value:
            yield from _flatten_excel_range_values(item)
    else:
        yield value

def parse_excel_selection(range_obj):
    values = []
    try:
        if hasattr(range_obj, 'Value'):
            raw = range_obj.Value
        elif hasattr(range_obj, 'Value2'):
            raw = range_obj.Value2
        else:
            raw = range_obj
        if raw is None:
            return []
        for item in _flatten_excel_range_values(raw):
            if item is None:
                continue
            try:
                values.append(float(item))
            except Exception:
                continue
        return values
    except Exception:
        return []

def parse_excel_selection_2cols(range_obj):
    """解析两列数据，返回 (x_list, y_list)"""
    x_vals = []
    y_vals = []
    try:
        if hasattr(range_obj, 'Value'):
            raw = range_obj.Value
        elif hasattr(range_obj, 'Value2'):
            raw = range_obj.Value2
        else:
            raw = range_obj
        if raw is None:
            return [], []
        # raw 可能是一个二维元组 (行, 列)
        if isinstance(raw, tuple) and len(raw) > 0 and isinstance(raw[0], tuple):
            # 多行多列
            for row in raw:
                if len(row) >= 2:
                    try:
                        x = float(row[0])
                        y = float(row[1])
                        x_vals.append(x)
                        y_vals.append(y)
                    except:
                        continue
        else:
            # 单行或单列？
            return [], []
        return x_vals, y_vals
    except Exception:
        return [], []

def is_integer_like(data: np.ndarray) -> bool:
    if len(data) == 0:
        return False
    return np.all(np.isfinite(data) & (np.abs(data - np.round(data)) < 1e-8))

def fit_binomial_mle(data):
    """改进的二项分布MLE拟合，扩大搜索范围并增加矩估计分支"""
    mean = np.mean(data)
    var = np.var(data, ddof=1)
    max_val = int(np.max(data))
    if mean <= 0:
        n_est = max_val if max_val > 0 else 1
        p_est = 1e-6
        return n_est, p_est
    if mean >= max_val:
        n_est = max_val
        p_est = 1.0 - 1e-6
        return n_est, p_est

    candidates = []
    if var < mean and mean > 0:
        n_mom = mean**2 / (mean - var)
        n_mom = int(round(n_mom))
        if n_mom >= max_val:
            candidates.append(n_mom)

    search_max = max(max_val * 3, max_val + 200)
    for n_candidate in range(max_val, search_max + 1):
        candidates.append(n_candidate)

    candidates = sorted(set(candidates))

    best_nll = np.inf
    best_n = max_val
    best_p = 0.5

    for n_candidate in candidates:
        if n_candidate <= 0:
            continue
        p_candidate = mean / n_candidate
        if p_candidate <= 0:
            p_candidate = 1e-6
        elif p_candidate >= 1:
            p_candidate = 1.0 - 1e-6
        if np.any(data > n_candidate) or np.any(data < 0):
            continue
        try:
            nll = -np.sum(sps.binom.logpmf(data, n_candidate, p_candidate))
            if np.isfinite(nll) and nll < best_nll:
                best_nll = nll
                best_n = n_candidate
                best_p = p_candidate
        except:
            continue

    if best_p <= 0:
        best_p = 1e-6
    if best_p >= 1:
        best_p = 1.0 - 1e-6
    if best_n < max_val:
        best_n = max_val
    return best_n, best_p

# ========== 改进的超几何分布拟合 ==========
def fit_hypergeometric_mle(data: np.ndarray):
    """
    超几何分布 MLE 拟合，返回参数顺序为 (n, D, M)
    其中 n = 样本容量, D = 成功数, M = 总体规模
    """
    data = np.asarray(data)
    data = data[~np.isnan(data)]
    if len(data) == 0:
        return None
    if not is_integer_like(data):
        return None
    data_int = np.round(data).astype(int)
    if np.any(data_int < 0):
        return None

    unique_vals, counts = np.unique(data_int, return_counts=True)
    min_val = int(np.min(data_int))
    max_val = int(np.max(data_int))
    mean_val = np.mean(data_int)
    var_val = np.var(data_int, ddof=0)

    if max_val == min_val:
        k = min_val
        n_est = k
        D_est = k
        M_est = max(k + 1, 2)
        return n_est, D_est, M_est

    def neg_log_likelihood(n, D, M):
        if n <= 0 or D <= 0 or M <= 0:
            return np.inf
        if D > M or n > M:
            return np.inf
        lower = max(0, n + D - M)
        upper = min(n, D)
        if min_val < lower or max_val > upper:
            return np.inf
        try:
            log_pmf = sps.hypergeom.logpmf(unique_vals, M, D, n)
            nll = -np.sum(counts * log_pmf)
            if np.isfinite(nll):
                return nll
            else:
                return np.inf
        except:
            return np.inf

    n_min = max_val
    n_upper_by_var = max(n_min, int(4 * var_val) + 10) if var_val > 0 else n_min * 3
    n_max = min(n_upper_by_var + 50, max_val * 10 + 200)
    n_max = max(n_max, n_min + 10)

    n_candidates = []
    for n_cand in range(n_min, min(n_min + 51, n_max + 1)):
        n_candidates.append(n_cand)
    if n_max > n_min + 50:
        step = max(1, (n_max - (n_min + 50)) // 30)
        for n_cand in range(n_min + 50, n_max + 1, step):
            n_candidates.append(n_cand)
    n_candidates = sorted(set(n_candidates))

    best_nll = np.inf
    best_params = None

    def search_best_for_n(n):
        nonlocal best_nll, best_params
        if n <= 0:
            return
        p_est = mean_val / n if n > 0 else 0.5
        if p_est <= 0 or p_est >= 1:
            return
        A = n * p_est * (1 - p_est)
        if A <= 0:
            return
        ratio = var_val / A
        if ratio <= 0:
            M_est = n
        elif ratio >= 1:
            M_est = n
        else:
            M_est = (n - ratio) / (1 - ratio)
        M_est = int(round(M_est))
        M_est = max(n, min(M_est, n * 10 + 100))
        D_est = int(round(p_est * M_est))
        D_est = max(1, min(M_est - 1, D_est))

        for radius in [5, 10, 20]:
            M_low = max(n, M_est - radius)
            M_high = min(M_est + radius, n * 20 + 200)
            D_low = max(1, D_est - radius)
            D_high = min(M_high - 1, D_est + radius)
            for M_cand in range(M_low, M_high + 1):
                p_target = mean_val / n
                D_center = int(round(p_target * M_cand))
                D_range = max(3, radius // 2)
                D_start = max(1, D_center - D_range)
                D_end = min(M_cand - 1, D_center + D_range)
                for D_cand in range(D_start, D_end + 1):
                    nll = neg_log_likelihood(n, D_cand, M_cand)
                    if nll < best_nll:
                        best_nll = nll
                        best_params = (n, D_cand, M_cand)
            if best_nll < 1e-6:
                break

    for n_cand in n_candidates:
        search_best_for_n(n_cand)

    if best_params is not None:
        n_best, D_best, M_best = best_params
        refine_radius = 5
        n_low = max(n_min, n_best - refine_radius)
        n_high = n_best + refine_radius
        M_low = max(n_low, M_best - refine_radius)
        M_high = M_best + refine_radius
        D_low = max(1, D_best - refine_radius)
        D_high = min(M_high - 1, D_best + refine_radius)
        for n_cand in range(n_low, n_high + 1):
            for M_cand in range(max(n_cand, M_low), M_high + 1):
                D_center = int(round(mean_val * M_cand / n_cand)) if n_cand > 0 else D_best
                D_start = max(1, D_center - 5)
                D_end = min(M_cand - 1, D_center + 5)
                for D_cand in range(D_start, D_end + 1):
                    nll = neg_log_likelihood(n_cand, D_cand, M_cand)
                    if nll < best_nll:
                        best_nll = nll
                        best_params = (n_cand, D_cand, M_cand)

    if best_params is None:
        n_est = max_val
        D_est = max_val
        M_est = max_val * 2 + 10
        lower = max(0, n_est + D_est - M_est)
        upper = min(n_est, D_est)
        if min_val >= lower and max_val <= upper:
            best_params = (n_est, D_est, M_est)
        else:
            M_est = max_val * 3 + 20
            lower = max(0, n_est + D_est - M_est)
            if min_val >= lower and max_val <= upper:
                best_params = (n_est, D_est, M_est)
            else:
                M_est = max_val * 5
                n_est = max_val
                D_est = max_val
                best_params = (n_est, D_est, M_est)

    n, D, M = best_params
    n = int(round(n))
    D = int(round(D))
    M = int(round(M))
    if n <= 0:
        n = max_val
    if D <= 0:
        D = 1
    if M <= 0:
        M = max_val * 2
    if D > M:
        D = M - 1
    if n > M:
        n = M
    if n < max_val:
        n = max_val
    lower = max(0, n + D - M)
    upper = min(n, D)
    if np.any(data_int < lower) or np.any(data_int > upper):
        for extra in [10, 50, 100, 200]:
            M2 = M + extra
            lower2 = max(0, n + D - M2)
            if np.all(data_int >= lower2) and np.all(data_int <= upper):
                M = M2
                break
        else:
            M = max_val * 5 + 100
            n = max_val
            D = max_val
    return n, D, M

fit_hypergeometric_fast = fit_hypergeometric_mle

def fit_negative_binomial_moments(data):
    mean = np.mean(data)
    var = np.var(data, ddof=1)
    if var <= mean or mean <= 0:
        p = 1.0 / (mean + 1.0)
        p = min(max(p, 1e-6), 1.0)
        r = 1
        return r, p
    r = mean**2 / (var - mean)
    p = mean / var
    if r <= 0:
        r = 1.0
    if p <= 0:
        p = 1e-6
    if p >= 1:
        p = 1.0 - 1e-6
    return r, p

def _vectorize_scalar_function(func: Callable[[float], float]) -> Callable[[np.ndarray], np.ndarray]:
    def vectorized(x):
        x = np.asarray(x, dtype=float)
        if x.ndim == 0:
            x = x.reshape((1,))
        return np.array([func(float(xx)) for xx in x], dtype=float)
    return vectorized

def _validate_distribution_params(dist_name: str, params: List[float], shift: float) -> bool:
    if dist_name == 'student':
        df = params[0]
        return df > 0 and abs(df - round(df)) < 1e-6
    if dist_name == 'chisq':
        df = params[0]
        return df > 0 and abs(df - round(df)) < 1e-6
    if dist_name == 'f':
        dfn, dfd = params[0], params[1]
        return dfn > 0 and dfd > 0 and abs(dfn - round(dfn)) < 1e-6 and abs(dfd - round(dfd)) < 1e-6
    if dist_name == 'binomial':
        n, p = params[0], params[1]
        return n > 0 and abs(n - round(n)) < 1e-6 and 1e-6 <= p <= 1-1e-6
    if dist_name == 'poisson':
        lam = params[0]
        return lam > 1e-6
    if dist_name == 'negbin':
        r, p = params[0], params[1]
        return r > 0 and abs(r - round(r)) < 1e-6 and 1e-6 <= p <= 1
    if dist_name == 'geomet':
        p = params[0]
        return 1e-6 <= p <= 1
    if dist_name == 'hypergeo':
        n, D, M = params[0], params[1], params[2]
        return (n > 0 and D > 0 and M > 0 and
                abs(n - round(n)) < 1e-6 and abs(D - round(D)) < 1e-6 and abs(M - round(M)) < 1e-6 and
                D <= M and n <= M)
    if dist_name == 'intuniform':
        a, b = params[0], params[1]
        return a < b and abs(a - round(a)) < 1e-6 and abs(b - round(b)) < 1e-6
    if dist_name == 'erlang':
        m, beta = params[0], params[1]
        return m > 0 and abs(m - round(m)) < 1e-6 and beta > 0
    return True

def fit_scipy_distribution(data: np.ndarray, dist_name: str):
    """拟合分布，返回 (params, shift, scipy_dist) 或 None"""
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        data = np.array(data, dtype=float)
        if len(data) == 0 or np.any(np.isnan(data)):
            return None
        try:
            # 连续分布
            if dist_name == 'normal':
                loc, scale = sps.norm.fit(data)
                params = [loc, scale]
                shift = 0.0
                scipy_dist = sps.norm(loc=loc, scale=scale)
            elif dist_name == 'uniform':
                loc, scale = sps.uniform.fit(data)
                params = [loc, loc + scale]
                shift = 0.0
                scipy_dist = sps.uniform(loc=loc, scale=scale)
            elif dist_name == 'gamma':
                a, loc, scale = sps.gamma.fit(data)
                params = [a, scale]
                shift = loc
                scipy_dist = sps.gamma(a=a, scale=scale)
            elif dist_name == 'erlang':
                a, loc, scale = sps.gamma.fit(data)
                a = max(1, int(round(a)))
                params = [a, scale]
                shift = loc
                scipy_dist = sps.gamma(a=a, scale=scale)
            elif dist_name == 'beta':
                if np.min(data) < 0 or np.max(data) > 1:
                    return None
                a, b, loc, scale = sps.beta.fit(data)
                params = [a, b]
                shift = loc
                scipy_dist = sps.beta(a=a, b=b)
            elif dist_name == 'chisq':
                df, loc, scale = sps.chi2.fit(data)
                df = max(1, int(round(df)))
                params = [df]
                shift = loc
                scipy_dist = sps.chi2(df=df, loc=0, scale=scale)
            elif dist_name == 'f':
                dfn, dfd, loc, scale = sps.f.fit(data)
                dfn = max(1, int(round(dfn)))
                dfd = max(1, int(round(dfd)))
                params = [dfn, dfd]
                shift = loc
                scipy_dist = sps.f(dfn=dfn, dfd=dfd, loc=0, scale=scale)
            elif dist_name == 'student':
                df, loc, scale = sps.t.fit(data)
                df = max(1, int(round(df)))
                params = [df]
                shift = loc
                scipy_dist = sps.t(df=df, loc=0, scale=scale)
            elif dist_name == 'expon':
                loc, scale = sps.expon.fit(data)
                params = [scale]
                shift = loc
                scipy_dist = sps.expon(loc=0, scale=scale)
            elif dist_name == 'triang':
                c, loc, scale = sps.triang.fit(data)
                params = [loc, loc + c * scale, loc + scale]
                shift = 0.0
                scipy_dist = sps.triang(c=c, loc=loc, scale=scale)
            elif dist_name == 'cauchy':
                loc, scale = sps.cauchy.fit(data)
                params = [loc, scale]
                shift = 0.0
                scipy_dist = sps.cauchy(loc=loc, scale=scale)
            elif dist_name == 'extvalue':
                loc, scale = sps.gumbel_r.fit(data)
                params = [loc, scale]
                shift = 0.0
                scipy_dist = sps.gumbel_r(loc=loc, scale=scale)
            elif dist_name == 'extvaluemin':
                loc, scale = sps.gumbel_l.fit(data)
                params = [loc, scale]
                shift = 0.0
                scipy_dist = sps.gumbel_l(loc=loc, scale=scale)
            elif dist_name == 'fatiguelife':
                c, loc, scale = sps.fatiguelife.fit(data)
                params = [loc, scale, c]
                shift = 0.0
                scipy_dist = sps.fatiguelife(c, loc=loc, scale=scale)
            elif dist_name == 'frechet':
                c, loc, scale = sps.invweibull.fit(data)
                params = [loc, scale, c]
                shift = 0.0
                scipy_dist = sps.invweibull(c, loc=loc, scale=scale)
            elif dist_name == 'hypsecant':
                loc, scale = sps.hypsecant.fit(data)
                params = [loc, scale]
                shift = 0.0
                scipy_dist = sps.hypsecant(loc=loc, scale=scale)
            elif dist_name == 'johnsonsb':
                a, b, loc, scale = sps.johnsonsb.fit(data)
                params = [a, b, loc, loc + scale]
                shift = 0.0
                scipy_dist = sps.johnsonsb(a, b, loc=loc, scale=scale)
            elif dist_name == 'johnsonsu':
                a, b, loc, scale = sps.johnsonsu.fit(data)
                params = [a, b, loc, scale]
                shift = 0.0
                scipy_dist = sps.johnsonsu(a, b, loc=loc, scale=scale)
            elif dist_name == 'kumaraswamy':
                min_val = float(np.min(data))
                max_val = float(np.max(data))
                if max_val <= min_val:
                    return None
                scaled = (data - min_val) / (max_val - min_val)
                a, b, loc, scale = sps.beta.fit(scaled, floc=0.0, fscale=1.0)
                if a <= 0 or b <= 0:
                    return None
                params = [a, b, min_val, max_val]
                shift = 0.0
                scipy_dist = None
            elif dist_name == 'erf':
                shift = float(np.mean(data))
                centered = data - shift
                ssq = float(np.sum(centered ** 2))
                if ssq <= 0:
                    return None
                h = math.sqrt(len(data) / (2.0 * ssq))
                if h <= 0:
                    return None
                params = [h]
                shift = shift
                scipy_dist = None
            elif dist_name == 'laplace':
                loc, scale = sps.laplace.fit(data)
                sigma = scale * math.sqrt(2.0)
                params = [loc, sigma]
                shift = 0.0
                scipy_dist = sps.laplace(loc=loc, scale=scale)
            elif dist_name == 'logistic':
                loc, scale = sps.logistic.fit(data)
                params = [loc, scale]
                shift = 0.0
                scipy_dist = sps.logistic(loc=loc, scale=scale)
            elif dist_name == 'loglogistic':
                c, loc, scale = sps.fisk.fit(data)
                params = [loc, scale, c]
                shift = 0.0
                scipy_dist = sps.fisk(c, loc=loc, scale=scale)
            elif dist_name == 'lognorm':
                s, loc, scale = sps.lognorm.fit(data)
                mu = scale * math.exp(0.5 * s**2)
                sigma = math.sqrt((math.exp(s**2) - 1.0) * scale**2 * math.exp(s**2))
                params = [mu, sigma]
                shift = loc
                scipy_dist = sps.lognorm(s, loc=0, scale=scale)
            elif dist_name == 'lognorm2':
                s, loc, scale = sps.lognorm.fit(data)
                mu = math.log(scale)
                params = [mu, s]
                shift = loc
                scipy_dist = sps.lognorm(s, loc=0, scale=scale)
            elif dist_name == 'betageneral':
                a, b, loc, scale = sps.beta.fit(data)
                params = [a, b, loc, loc + scale]
                shift = 0.0
                scipy_dist = sps.beta(a=a, b=b, loc=loc, scale=scale)
            elif dist_name == 'betasubj':
                a, b, loc, scale = sps.beta.fit(data)
                params = [a, b, loc, loc + scale]
                shift = 0.0
                scipy_dist = sps.beta(a=a, b=b, loc=loc, scale=scale)
            elif dist_name == 'burr12':
                c, d, loc, scale = sps.burr12.fit(data)
                params = [loc, scale, c, d]
                shift = 0.0
                scipy_dist = sps.burr12(c, d, loc=loc, scale=scale)
            elif dist_name == 'pert':
                if np.min(data) < 0:
                    return None
                min_val = np.min(data)
                max_val = np.max(data)
                mean_val = float(np.mean(data))
                mode = (6.0 * mean_val - min_val - max_val) / 4.0
                mode = float(min(max(mode, min_val), max_val))
                params = [min_val, mode, max_val]
                shift = 0.0
                scipy_dist = None
            elif dist_name == 'reciprocal':
                a, b, loc, scale = sps.reciprocal.fit(data)
                params = [loc, scale]
                shift = 0.0
                scipy_dist = sps.reciprocal(loc=loc, scale=scale)
            elif dist_name == 'rayleigh':
                loc, scale = sps.rayleigh.fit(data)
                params = [scale]
                shift = loc
                scipy_dist = sps.rayleigh(loc=loc, scale=scale)
            elif dist_name == 'weibull':
                c, loc, scale = sps.weibull_min.fit(data)
                params = [c, scale]
                shift = loc
                scipy_dist = sps.weibull_min(c, loc=0, scale=scale)
            elif dist_name == 'pearson5':
                a, loc, scale = sps.invgamma.fit(data)
                params = [a, scale]
                shift = loc
                scipy_dist = sps.invgamma(a, loc=0, scale=scale)
            elif dist_name == 'pearson6':
                a, b, loc, scale = sps.betaprime.fit(data)
                params = [a, b, scale]
                shift = loc
                scipy_dist = sps.betaprime(a, b, loc=0, scale=scale)
            elif dist_name == 'pareto2':
                c, loc, scale = sps.lomax.fit(data)
                params = [scale, c]
                shift = loc
                scipy_dist = sps.lomax(c, loc=0, scale=scale)
            elif dist_name == 'pareto':
                b, loc, scale = sps.pareto.fit(data)
                params = [b, scale]
                shift = loc
                scipy_dist = sps.pareto(b, loc=0, scale=scale)
            elif dist_name == 'levy':
                loc, scale = sps.levy.fit(data)
                params = [loc, scale]
                shift = 0.0
                scipy_dist = sps.levy(loc=loc, scale=scale)
            elif dist_name == 'invgauss':
                mu, loc, scale = sps.invgauss.fit(data)
                params = [mu, scale]
                shift = loc
                scipy_dist = sps.invgauss(mu, loc=0, scale=scale)
            # 离散分布
            elif dist_name == 'geomet':
                if not is_integer_like(data):
                    return None
                if np.any(data < 1):
                    return None
                p = 1.0 / (np.mean(data) + 1.0)
                p = min(max(p, 1e-6), 1.0)
                params = [p]
                shift = 0.0
                scipy_dist = sps.geom(p, loc=0)
            elif dist_name == 'poisson':
                if np.any(data < 0):
                    return None
                lam = np.mean(data)
                if lam <= 1e-6:
                    lam = 1e-6
                params = [lam]
                shift = 0.0
                scipy_dist = sps.poisson(mu=lam)
            elif dist_name == 'binomial':
                if not is_integer_like(data):
                    return None
                if np.any(data < 0):
                    return None
                n_binom, p = fit_binomial_mle(data)
                n_binom = int(round(n_binom))
                if n_binom <= 0:
                    n_binom = max(1, int(np.max(data)))
                if p <= 0:
                    p = 1e-6
                if p >= 1:
                    p = 1.0 - 1e-6
                if np.any(data > n_binom):
                    n_binom = max(n_binom, int(np.max(data)))
                if n_binom <= 0 or p <= 0 or p >= 1:
                    return None
                params = [n_binom, p]
                shift = 0.0
                scipy_dist = sps.binom(n_binom, p)
            elif dist_name == 'negbin':
                if not is_integer_like(data):
                    return None
                if np.any(data < 0):
                    return None
                result = fit_negative_binomial_moments(data)
                if result is None:
                    return None
                r, p = result
                r = max(1, int(round(r)))
                p = min(max(p, 1e-6), 1.0)
                params = [r, p]
                shift = 0.0
                scipy_dist = sps.nbinom(r, p)
            elif dist_name == 'hypergeo':
                if not is_integer_like(data):
                    return None
                if np.any(data < 0):
                    return None
                fit_result = fit_hypergeometric_fast(data)
                if fit_result is None:
                    return None
                n, D, M = fit_result
                n = int(round(n))
                D = int(round(D))
                M = int(round(M))
                if n <= 0 or D <= 0 or M <= 0:
                    return None
                if D > M:
                    D = M - 1
                if n > M:
                    n = M
                if n <= 0 or D <= 0:
                    return None
                lower = max(0, n + D - M)
                upper = min(n, D)
                data_int = np.round(data).astype(int)
                if np.any(data_int < lower) or np.any(data_int > upper):
                    M = max(M, int(np.max(data)) * 2 + 20)
                    lower2 = max(0, n + D - M)
                    if np.any(data_int < lower2) or np.any(data_int > upper):
                        return None
                params = [n, D, M]
                shift = 0.0
                scipy_dist = sps.hypergeom(M, D, n)
            elif dist_name == 'intuniform':
                if not is_integer_like(data):
                    return None
                a = int(np.min(data))
                b = int(np.max(data))
                if a >= b:
                    return None
                if np.any(data < a) or np.any(data > b):
                    return None
                params = [a, b]
                shift = 0.0
                scipy_dist = sps.randint(low=a, high=b + 1)
            else:
                return None

            if not _validate_distribution_params(dist_name, params, shift):
                return None

            return params, shift, scipy_dist
        except Exception:
            return None

def _get_custom_distribution_funcs(dist_name: str, params: List[float]):
    if dist_name == 'erf':
        h = float(params[0])
        return (
            _vectorize_scalar_function(lambda x: df.erf_pdf(x, h)),
            _vectorize_scalar_function(lambda x: df.erf_cdf(x, h))
        )
    if dist_name == 'kumaraswamy':
        alpha1, alpha2, min_val, max_val = [float(v) for v in params]
        return (
            _vectorize_scalar_function(lambda x: df.kumaraswamy_pdf(x, alpha1, alpha2, min_val, max_val)),
            _vectorize_scalar_function(lambda x: df.kumaraswamy_cdf(x, alpha1, alpha2, min_val, max_val))
        )
    if dist_name == 'pert':
        min_val, mode, max_val = [float(v) for v in params]
        return (
            _vectorize_scalar_function(lambda x: df.pert_pdf(x, min_val, mode, max_val)),
            _vectorize_scalar_function(lambda x: df.pert_cdf(x, min_val, mode, max_val))
        )
    return None, None

class DistributionWrapper:
    def __init__(self, dist_name: str, params: List[float], shift: float, scipy_dist, pdf_func=None, cdf_func=None):
        self.dist_name = dist_name
        self.params = list(params)
        self.shift = float(shift)
        self.scipy_dist = scipy_dist
        self.pdf_func = pdf_func
        self.cdf_func = cdf_func

    def is_valid(self):
        return self.scipy_dist is not None or self.cdf_func is not None

    def logpdf(self, x):
        x = np.asarray(x, dtype=float)
        if self.scipy_dist is None and self.pdf_func is None:
            return np.full_like(x, -np.inf, dtype=float)
        try:
            if self.dist_name in DISCRETE_DISTRIBUTIONS:
                if self.scipy_dist is not None:
                    pmf = self.scipy_dist.logpmf(np.round(x - self.shift))
                    return np.where(np.isfinite(pmf), pmf, -np.inf)
                return np.full_like(x, -np.inf, dtype=float)
            if self.pdf_func is not None:
                pdf = self.pdf_func(x - self.shift)
                result = np.full_like(pdf, -np.inf, dtype=float)
                valid = (pdf > 0) & np.isfinite(pdf)
                result[valid] = np.log(pdf[valid])
                return result
            return self.scipy_dist.logpdf(x - self.shift)
        except Exception:
            return np.full_like(x, -np.inf, dtype=float)

    def cdf(self, x):
        x = np.asarray(x, dtype=float)
        if self.scipy_dist is None and self.cdf_func is None:
            return np.zeros_like(x, dtype=float)
        try:
            if self.cdf_func is not None:
                return self.cdf_func(x - self.shift)
            return self.scipy_dist.cdf(x - self.shift)
        except Exception:
            return np.zeros_like(x, dtype=float)

    def ppf(self, q):
        q = np.asarray(q, dtype=float)
        if self.scipy_dist is None and self.cdf_func is None:
            return np.full_like(q, np.nan, dtype=float)
        try:
            if self.scipy_dist is not None:
                return self.scipy_dist.ppf(q) + self.shift
            return np.full_like(q, np.nan, dtype=float)
        except Exception:
            return np.full_like(q, np.nan, dtype=float)

    def pdf(self, x):
        x = np.asarray(x, dtype=float)
        if self.scipy_dist is None and self.pdf_func is None:
            return np.zeros_like(x, dtype=float)
        try:
            if self.pdf_func is not None:
                return self.pdf_func(x - self.shift)
            return self.scipy_dist.pdf(x - self.shift)
        except Exception:
            return np.zeros_like(x, dtype=float)

    def mean(self):
        if self.scipy_dist is not None:
            return self.scipy_dist.mean() + self.shift
        # 对于自定义分布，尝试数值积分
        try:
            # 粗略估计支撑集
            x_min = self.ppf(0.001)
            x_max = self.ppf(0.999)
            if np.isnan(x_min) or np.isnan(x_max):
                return np.nan
            x_vals = np.linspace(x_min, x_max, 1000)
            pdf_vals = self.pdf(x_vals)
            dx = x_vals[1] - x_vals[0]
            mean_val = np.sum(x_vals * pdf_vals) * dx
            return mean_val
        except:
            return np.nan

    def std(self):
        if self.scipy_dist is not None:
            return self.scipy_dist.std()
        try:
            mean_val = self.mean()
            if np.isnan(mean_val):
                return np.nan
            x_min = self.ppf(0.001)
            x_max = self.ppf(0.999)
            if np.isnan(x_min) or np.isnan(x_max):
                return np.nan
            x_vals = np.linspace(x_min, x_max, 1000)
            pdf_vals = self.pdf(x_vals)
            dx = x_vals[1] - x_vals[0]
            var_val = np.sum((x_vals - mean_val)**2 * pdf_vals) * dx
            return np.sqrt(var_val)
        except:
            return np.nan

    def median(self):
        return self.ppf(0.5)

    def mode(self):
        """数值计算众数（PDF最大值点）"""
        try:
            # 获取支撑集
            x_min = self.ppf(0.001)
            x_max = self.ppf(0.999)
            if np.isnan(x_min) or np.isnan(x_max):
                return np.nan
            # 使用优化求最大值
            def neg_pdf(x):
                return -self.pdf(np.array([x]))[0]
            res = minimize_scalar(neg_pdf, bounds=(x_min, x_max), method='bounded')
            if res.success:
                return res.x
            else:
                # 退化为网格搜索
                x_vals = np.linspace(x_min, x_max, 500)
                pdf_vals = self.pdf(x_vals)
                return x_vals[np.argmax(pdf_vals)]
        except:
            return np.nan

    def support(self):
        """返回支撑集下界和上界（可能为 -inf/inf）"""
        if self.scipy_dist is not None:
            a, b = self.scipy_dist.support()
            return a + self.shift, b + self.shift
        # 对于自定义分布，从分位数估计
        try:
            x_min = self.ppf(1e-10)
            x_max = self.ppf(1 - 1e-10)
            if np.isnan(x_min) or np.isnan(x_max):
                return -np.inf, np.inf
            return x_min, x_max
        except:
            return -np.inf, np.inf

# ================== 拟合指标计算 ==================
def compute_aic(nll, k, n):
    return 2 * k + 2 * nll

def compute_aicc(nll, k, n):
    aic = compute_aic(nll, k, n)
    if n - k - 1 > 0:
        return aic + (2 * k * (k + 1)) / (n - k - 1)
    return np.inf

def compute_bic(nll, k, n):
    return k * np.log(n) + 2 * nll

def compute_ks_statistic(data, dist: DistributionWrapper):
    try:
        sorted_data = np.sort(data)
        cdf_vals = dist.cdf(sorted_data)
        ecdf = np.arange(1, len(data) + 1) / len(data)
        return np.max(np.abs(ecdf - cdf_vals))
    except Exception:
        return np.inf

def compute_ad_statistic(data, dist: DistributionWrapper):
    try:
        sorted_data = np.sort(data)
        n = len(data)
        cdf_vals = dist.cdf(sorted_data)
        cdf_vals = np.clip(cdf_vals, 1e-10, 1 - 1e-10)
        return -n - np.sum((2 * np.arange(1, n + 1) - 1) * (np.log(cdf_vals) + np.log(1 - cdf_vals[::-1]))) / n
    except Exception:
        return np.inf

def compute_chisq_statistic(data, dist: DistributionWrapper, bins='auto'):
    try:
        n = len(data)
        if bins == 'auto':
            bins = int(np.sqrt(n))
        counts, bin_edges = np.histogram(data, bins=bins, density=False)
        expected = np.array([(dist.cdf(bin_edges[i + 1]) - dist.cdf(bin_edges[i])) * n for i in range(len(bin_edges) - 1)])
        valid = expected > 0
        if not np.any(valid):
            return np.inf
        return np.sum((counts[valid] - expected[valid])**2 / expected[valid])
    except Exception:
        return np.inf

def compute_avlog(data, dist: DistributionWrapper):
    try:
        logpdf_vals = dist.logpdf(data)
        return -np.mean(logpdf_vals)
    except Exception:
        return np.inf

# ================== 密度/累积拟合专用函数 ==================
def _normalize_density(x, y):
    """对密度数据(y)进行标准化，使得曲线下面积=1（使用梯形积分）"""
    if len(x) < 2:
        return y
    area = np.trapz(y, x)
    if area <= 0:
        return y
    return y / area

def _compute_rmse(theoretical, observed):
    return np.sqrt(np.mean((theoretical - observed)**2))

def _get_distribution_pdf_func(dist_name: str, params: List[float]):
    """返回给定分布和参数的PDF函数（接受numpy数组）"""
    # 尝试使用scipy分布
    scipy_dist = None
    try:
        if dist_name == 'normal':
            loc, scale = params[0], params[1]
            scipy_dist = sps.norm(loc=loc, scale=scale)
        elif dist_name == 'uniform':
            a, b = params[0], params[1]
            loc, scale = a, b - a
            scipy_dist = sps.uniform(loc=loc, scale=scale)
        elif dist_name == 'gamma':
            a, scale = params[0], params[1]
            scipy_dist = sps.gamma(a=a, scale=scale)
        elif dist_name == 'erlang':
            m, beta = params[0], params[1]
            scipy_dist = sps.gamma(a=m, scale=beta)
        elif dist_name == 'beta':
            a, b = params[0], params[1]
            scipy_dist = sps.beta(a=a, b=b)
        elif dist_name == 'chisq':
            df = params[0]
            scipy_dist = sps.chi2(df=df)
        elif dist_name == 'f':
            dfn, dfd = params[0], params[1]
            scipy_dist = sps.f(dfn=dfn, dfd=dfd)
        elif dist_name == 'student':
            df = params[0]
            scipy_dist = sps.t(df=df)
        elif dist_name == 'expon':
            scale = params[0]
            scipy_dist = sps.expon(scale=scale)
        elif dist_name == 'triang':
            a, b, c = params[0], params[1], params[2]
            loc = a
            scale = b - a
            c_scaled = (c - a) / (b - a)
            scipy_dist = sps.triang(c=c_scaled, loc=loc, scale=scale)
        elif dist_name == 'cauchy':
            loc, scale = params[0], params[1]
            scipy_dist = sps.cauchy(loc=loc, scale=scale)
        elif dist_name == 'extvalue':
            loc, scale = params[0], params[1]
            scipy_dist = sps.gumbel_r(loc=loc, scale=scale)
        elif dist_name == 'extvaluemin':
            loc, scale = params[0], params[1]
            scipy_dist = sps.gumbel_l(loc=loc, scale=scale)
        elif dist_name == 'fatiguelife':
            loc, scale, c = params[0], params[1], params[2]
            scipy_dist = sps.fatiguelife(c, loc=loc, scale=scale)
        elif dist_name == 'frechet':
            loc, scale, c = params[0], params[1], params[2]
            scipy_dist = sps.invweibull(c, loc=loc, scale=scale)
        elif dist_name == 'hypsecant':
            loc, scale = params[0], params[1]
            scipy_dist = sps.hypsecant(loc=loc, scale=scale)
        elif dist_name == 'johnsonsb':
            a, b, loc, scale = params[0], params[1], params[2], params[3]
            scipy_dist = sps.johnsonsb(a, b, loc=loc, scale=scale - loc)
        elif dist_name == 'johnsonsu':
            a, b, loc, scale = params[0], params[1], params[2], params[3]
            scipy_dist = sps.johnsonsu(a, b, loc=loc, scale=scale)
        elif dist_name == 'kumaraswamy':
            # 使用自定义函数
            alpha1, alpha2, min_val, max_val = params
            def pdf_func(x):
                x_scaled = (x - min_val) / (max_val - min_val)
                valid = (x >= min_val) & (x <= max_val)
                res = np.zeros_like(x)
                res[valid] = (alpha1 * alpha2 * x_scaled[valid]**(alpha1-1) * 
                              (1 - x_scaled[valid]**alpha1)**(alpha2-1)) / (max_val - min_val)
                return res
            return pdf_func
        elif dist_name == 'laplace':
            loc, sigma = params[0], params[1]
            scale = sigma / math.sqrt(2.0)
            scipy_dist = sps.laplace(loc=loc, scale=scale)
        elif dist_name == 'logistic':
            loc, scale = params[0], params[1]
            scipy_dist = sps.logistic(loc=loc, scale=scale)
        elif dist_name == 'loglogistic':
            loc, scale, c = params[0], params[1], params[2]
            scipy_dist = sps.fisk(c, loc=loc, scale=scale)
        elif dist_name == 'lognorm':
            mu, sigma = params[0], params[1]
            s = math.sqrt(math.log(1 + (sigma/mu)**2))
            scale = mu / math.exp(0.5 * s**2)
            scipy_dist = sps.lognorm(s, scale=scale)
        elif dist_name == 'lognorm2':
            mu, s = params[0], params[1]
            scale = math.exp(mu)
            scipy_dist = sps.lognorm(s, scale=scale)
        elif dist_name == 'betageneral':
            a, b, low, high = params[0], params[1], params[2], params[3]
            scale = high - low
            scipy_dist = sps.beta(a, b, loc=low, scale=scale)
        elif dist_name == 'betasubj':
            a, b, low, high = params[0], params[1], params[2], params[3]
            scale = high - low
            scipy_dist = sps.beta(a, b, loc=low, scale=scale)
        elif dist_name == 'burr12':
            loc, scale, c, d = params[0], params[1], params[2], params[3]
            scipy_dist = sps.burr12(c, d, loc=loc, scale=scale)
        elif dist_name == 'pert':
            min_val, mode, max_val = params
            # 自定义PERT PDF
            def pdf_func(x):
                alpha = 1 + (mode - min_val) / (max_val - min_val) * 4
                beta = 1 + (max_val - mode) / (max_val - min_val) * 4
                from scipy.special import beta as beta_func
                B = beta_func(alpha, beta)
                x_scaled = (x - min_val) / (max_val - min_val)
                valid = (x >= min_val) & (x <= max_val)
                res = np.zeros_like(x)
                res[valid] = (x_scaled[valid]**(alpha-1) * (1-x_scaled[valid])**(beta-1)) / (B * (max_val - min_val))
                return res
            return pdf_func
        elif dist_name == 'reciprocal':
            loc, scale = params[0], params[1]
            scipy_dist = sps.reciprocal(loc=loc, scale=scale)
        elif dist_name == 'rayleigh':
            scale = params[0]
            scipy_dist = sps.rayleigh(scale=scale)
        elif dist_name == 'weibull':
            c, scale = params[0], params[1]
            scipy_dist = sps.weibull_min(c, scale=scale)
        elif dist_name == 'pearson5':
            a, scale = params[0], params[1]
            scipy_dist = sps.invgamma(a, scale=scale)
        elif dist_name == 'pearson6':
            a, b, scale = params[0], params[1], params[2]
            scipy_dist = sps.betaprime(a, b, scale=scale)
        elif dist_name == 'pareto2':
            scale, c = params[0], params[1]
            scipy_dist = sps.lomax(c, scale=scale)
        elif dist_name == 'pareto':
            b, scale = params[0], params[1]
            scipy_dist = sps.pareto(b, scale=scale)
        elif dist_name == 'levy':
            loc, scale = params[0], params[1]
            scipy_dist = sps.levy(loc=loc, scale=scale)
        elif dist_name == 'invgauss':
            mu, scale = params[0], params[1]
            scipy_dist = sps.invgauss(mu, scale=scale)
        elif dist_name == 'erf':
            h = params[0]
            def pdf_func(x):
                return df.erf_pdf(x, h)
            return pdf_func
        else:
            return None
        if scipy_dist is not None:
            return scipy_dist.pdf
    except:
        return None
    return None

def _get_distribution_cdf_func(dist_name: str, params: List[float]):
    """返回给定分布和参数的CDF函数"""
    # 先尝试获取PDF对应的分布对象，再取cdf
    pdf_func = _get_distribution_pdf_func(dist_name, params)
    if pdf_func is None:
        return None
    # 对于自定义分布，需要单独处理cdf
    if dist_name in ['kumaraswamy', 'pert', 'erf']:
        if dist_name == 'kumaraswamy':
            alpha1, alpha2, min_val, max_val = params
            def cdf_func(x):
                x_scaled = (x - min_val) / (max_val - min_val)
                valid = (x >= min_val) & (x <= max_val)
                res = np.zeros_like(x)
                res[valid] = 1 - (1 - x_scaled[valid]**alpha1)**alpha2
                res[x > max_val] = 1.0
                return res
            return cdf_func
        elif dist_name == 'pert':
            min_val, mode, max_val = params
            def cdf_func(x):
                alpha = 1 + (mode - min_val) / (max_val - min_val) * 4
                beta = 1 + (max_val - mode) / (max_val - min_val) * 4
                from scipy.special import betainc
                x_scaled = (x - min_val) / (max_val - min_val)
                valid = (x >= min_val) & (x <= max_val)
                res = np.zeros_like(x)
                res[valid] = betainc(alpha, beta, x_scaled[valid])
                res[x > max_val] = 1.0
                return res
            return cdf_func
        elif dist_name == 'erf':
            h = params[0]
            def cdf_func(x):
                return df.erf_cdf(x, h)
            return cdf_func
    else:
        # 使用scipy分布
        try:
            if dist_name == 'normal':
                loc, scale = params[0], params[1]
                return sps.norm(loc=loc, scale=scale).cdf
            elif dist_name == 'uniform':
                a, b = params[0], params[1]
                return sps.uniform(loc=a, scale=b-a).cdf
            elif dist_name == 'gamma':
                a, scale = params[0], params[1]
                return sps.gamma(a=a, scale=scale).cdf
            elif dist_name == 'erlang':
                m, beta = params[0], params[1]
                return sps.gamma(a=m, scale=beta).cdf
            elif dist_name == 'beta':
                a, b = params[0], params[1]
                return sps.beta(a=a, b=b).cdf
            elif dist_name == 'chisq':
                df = params[0]
                return sps.chi2(df=df).cdf
            elif dist_name == 'f':
                dfn, dfd = params[0], params[1]
                return sps.f(dfn=dfn, dfd=dfd).cdf
            elif dist_name == 'student':
                df = params[0]
                return sps.t(df=df).cdf
            elif dist_name == 'expon':
                scale = params[0]
                return sps.expon(scale=scale).cdf
            elif dist_name == 'triang':
                a, b, c = params[0], params[1], params[2]
                loc = a
                scale = b - a
                c_scaled = (c - a) / (b - a)
                return sps.triang(c=c_scaled, loc=loc, scale=scale).cdf
            elif dist_name == 'cauchy':
                loc, scale = params[0], params[1]
                return sps.cauchy(loc=loc, scale=scale).cdf
            elif dist_name == 'extvalue':
                loc, scale = params[0], params[1]
                return sps.gumbel_r(loc=loc, scale=scale).cdf
            elif dist_name == 'extvaluemin':
                loc, scale = params[0], params[1]
                return sps.gumbel_l(loc=loc, scale=scale).cdf
            elif dist_name == 'fatiguelife':
                loc, scale, c = params[0], params[1], params[2]
                return sps.fatiguelife(c, loc=loc, scale=scale).cdf
            elif dist_name == 'frechet':
                loc, scale, c = params[0], params[1], params[2]
                return sps.invweibull(c, loc=loc, scale=scale).cdf
            elif dist_name == 'hypsecant':
                loc, scale = params[0], params[1]
                return sps.hypsecant(loc=loc, scale=scale).cdf
            elif dist_name == 'johnsonsb':
                a, b, loc, scale = params[0], params[1], params[2], params[3]
                return sps.johnsonsb(a, b, loc=loc, scale=scale - loc).cdf
            elif dist_name == 'johnsonsu':
                a, b, loc, scale = params[0], params[1], params[2], params[3]
                return sps.johnsonsu(a, b, loc=loc, scale=scale).cdf
            elif dist_name == 'laplace':
                loc, sigma = params[0], params[1]
                scale = sigma / math.sqrt(2.0)
                return sps.laplace(loc=loc, scale=scale).cdf
            elif dist_name == 'logistic':
                loc, scale = params[0], params[1]
                return sps.logistic(loc=loc, scale=scale).cdf
            elif dist_name == 'loglogistic':
                loc, scale, c = params[0], params[1], params[2]
                return sps.fisk(c, loc=loc, scale=scale).cdf
            elif dist_name == 'lognorm':
                mu, sigma = params[0], params[1]
                s = math.sqrt(math.log(1 + (sigma/mu)**2))
                scale = mu / math.exp(0.5 * s**2)
                return sps.lognorm(s, scale=scale).cdf
            elif dist_name == 'lognorm2':
                mu, s = params[0], params[1]
                scale = math.exp(mu)
                return sps.lognorm(s, scale=scale).cdf
            elif dist_name == 'betageneral':
                a, b, low, high = params[0], params[1], params[2], params[3]
                scale = high - low
                return sps.beta(a, b, loc=low, scale=scale).cdf
            elif dist_name == 'betasubj':
                a, b, low, high = params[0], params[1], params[2], params[3]
                scale = high - low
                return sps.beta(a, b, loc=low, scale=scale).cdf
            elif dist_name == 'burr12':
                loc, scale, c, d = params[0], params[1], params[2], params[3]
                return sps.burr12(c, d, loc=loc, scale=scale).cdf
            elif dist_name == 'reciprocal':
                loc, scale = params[0], params[1]
                return sps.reciprocal(loc=loc, scale=scale).cdf
            elif dist_name == 'rayleigh':
                scale = params[0]
                return sps.rayleigh(scale=scale).cdf
            elif dist_name == 'weibull':
                c, scale = params[0], params[1]
                return sps.weibull_min(c, scale=scale).cdf
            elif dist_name == 'pearson5':
                a, scale = params[0], params[1]
                return sps.invgamma(a, scale=scale).cdf
            elif dist_name == 'pearson6':
                a, b, scale = params[0], params[1], params[2]
                return sps.betaprime(a, b, scale=scale).cdf
            elif dist_name == 'pareto2':
                scale, c = params[0], params[1]
                return sps.lomax(c, scale=scale).cdf
            elif dist_name == 'pareto':
                b, scale = params[0], params[1]
                return sps.pareto(b, scale=scale).cdf
            elif dist_name == 'levy':
                loc, scale = params[0], params[1]
                return sps.levy(loc=loc, scale=scale).cdf
            elif dist_name == 'invgauss':
                mu, scale = params[0], params[1]
                return sps.invgauss(mu, scale=scale).cdf
        except:
            pass
    return None

def _fit_distribution_to_density_or_cumulative(x, y, dist_name, fit_type):
    """
    使用最小二乘法拟合分布到密度或累积数据
    fit_type: 'density_raw', 'density_norm', 'cumulative'
    返回 (params, rmse, pdf_func, cdf_func) 或 None
    """
    x = np.asarray(x, dtype=float)
    y = np.asarray(y, dtype=float)
    if len(x) < 3 or len(y) != len(x):
        return None
    # 排序 x
    idx = np.argsort(x)
    x = x[idx]
    y = y[idx]
    if fit_type == 'density_norm':
        y = _normalize_density(x, y)
    # 定义目标函数：对于密度拟合，理论值是PDF；对于累积拟合，理论值是CDF
    def objective(params):
        try:
            if fit_type in ['density_raw', 'density_norm']:
                pdf_func = _get_distribution_pdf_func(dist_name, params)
                if pdf_func is None:
                    return np.inf
                theoretical = pdf_func(x)
            else:  # cumulative
                cdf_func = _get_distribution_cdf_func(dist_name, params)
                if cdf_func is None:
                    return np.inf
                theoretical = cdf_func(x)
            # 处理NaN或inf
            theoretical = np.nan_to_num(theoretical, nan=0.0, posinf=1e10, neginf=-1e10)
            return np.sum((theoretical - y)**2)
        except:
            return np.inf

    # 获取参数初始估计
    initial_params = _get_initial_params_for_distribution(dist_name, x, y, fit_type)
    if initial_params is None:
        return None

    # 参数边界
    bounds = _get_parameter_bounds(dist_name)
    # 使用L-BFGS-B优化
    try:
        res = minimize(objective, initial_params, method='L-BFGS-B', bounds=bounds,
                       options={'maxiter': 1000, 'ftol': 1e-8})
        if res.success or res.fun < 1e-4:
            best_params = res.x.tolist()
            rmse = np.sqrt(res.fun / len(x))
            # 获取最终函数
            if fit_type in ['density_raw', 'density_norm']:
                pdf_func = _get_distribution_pdf_func(dist_name, best_params)
                cdf_func = _get_distribution_cdf_func(dist_name, best_params)
            else:
                pdf_func = _get_distribution_pdf_func(dist_name, best_params)
                cdf_func = _get_distribution_cdf_func(dist_name, best_params)
            return best_params, rmse, pdf_func, cdf_func
    except:
        pass
    return None

def _get_initial_params_for_distribution(dist_name, x, y, fit_type):
    """为分布提供合理的初始参数猜测"""
    n_params = get_base_param_count(dist_name)
    if n_params is None:
        return None
    # 默认初始值
    default_params = {
        'normal': [0.0, 1.0],
        'uniform': [0.0, 1.0],
        'gamma': [1.0, 1.0],
        'erlang': [1.0, 1.0],
        'beta': [2.0, 2.0],
        'chisq': [1.0],
        'f': [5.0, 5.0],
        'student': [3.0],
        'expon': [1.0],
        'triang': [0.0, 1.0, 0.5],
        'cauchy': [0.0, 1.0],
        'extvalue': [0.0, 1.0],
        'extvaluemin': [0.0, 1.0],
        'fatiguelife': [0.0, 1.0, 1.0],
        'frechet': [0.0, 1.0, 1.0],
        'hypsecant': [0.0, 1.0],
        'johnsonsb': [0.0, 1.0, 0.0, 1.0],
        'johnsonsu': [0.0, 1.0, 0.0, 1.0],
        'kumaraswamy': [2.0, 2.0, np.min(x), np.max(x)],
        'laplace': [0.0, 1.0],
        'logistic': [0.0, 1.0],
        'loglogistic': [0.0, 1.0, 1.0],
        'lognorm': [1.0, 1.0],
        'lognorm2': [0.0, 1.0],
        'betageneral': [2.0, 2.0, np.min(x), np.max(x)],
        'betasubj': [2.0, 2.0, np.min(x), np.max(x)],
        'burr12': [0.0, 1.0, 1.0, 1.0],
        'pert': [np.min(x), np.median(x), np.max(x)],
        'reciprocal': [np.min(x), np.max(x)-np.min(x)],
        'rayleigh': [1.0],
        'weibull': [1.0, 1.0],
        'pearson5': [2.0, 1.0],
        'pearson6': [2.0, 2.0, 1.0],
        'pareto2': [1.0, 1.0],
        'pareto': [2.0, 1.0],
        'levy': [0.0, 1.0],
        'invgauss': [1.0, 1.0],
        'erf': [1.0],
    }
    if dist_name in default_params:
        params = default_params[dist_name][:n_params]
        # 根据数据范围调整
        if dist_name in ['uniform', 'beta', 'betageneral', 'betasubj', 'pert']:
            if len(params) >= 2:
                params[0] = np.min(x)
                params[1] = np.max(x)
            if dist_name == 'pert' and len(params) == 3:
                params[1] = np.median(x)
        if dist_name in ['gamma', 'erlang', 'weibull'] and np.min(x) < 0:
            # 如果x有负值，添加位置参数？但这里不考虑shift，所以可能不适合
            pass
        return params
    return None

def _get_parameter_bounds(dist_name):
    """返回参数边界，用于优化"""
    bounds = {
        'normal': [(-np.inf, np.inf), (1e-6, np.inf)],
        'uniform': [(-np.inf, np.inf), (1e-6, np.inf)],
        'gamma': [(1e-6, np.inf), (1e-6, np.inf)],
        'erlang': [(1, 100), (1e-6, np.inf)],
        'beta': [(1e-6, np.inf), (1e-6, np.inf)],
        'chisq': [(0.5, 100)],
        'f': [(0.5, 100), (0.5, 100)],
        'student': [(0.5, 100)],
        'expon': [(1e-6, np.inf)],
        'triang': [(-np.inf, np.inf), (-np.inf, np.inf), (-np.inf, np.inf)],
        'cauchy': [(-np.inf, np.inf), (1e-6, np.inf)],
        'extvalue': [(-np.inf, np.inf), (1e-6, np.inf)],
        'extvaluemin': [(-np.inf, np.inf), (1e-6, np.inf)],
        'fatiguelife': [(-np.inf, np.inf), (1e-6, np.inf), (1e-6, np.inf)],
        'frechet': [(-np.inf, np.inf), (1e-6, np.inf), (1e-6, np.inf)],
        'hypsecant': [(-np.inf, np.inf), (1e-6, np.inf)],
        'johnsonsb': [(-np.inf, np.inf), (1e-6, np.inf), (-np.inf, np.inf), (-np.inf, np.inf)],
        'johnsonsu': [(-np.inf, np.inf), (1e-6, np.inf), (-np.inf, np.inf), (1e-6, np.inf)],
        'kumaraswamy': [(1e-6, np.inf), (1e-6, np.inf), (-np.inf, np.inf), (-np.inf, np.inf)],
        'laplace': [(-np.inf, np.inf), (1e-6, np.inf)],
        'logistic': [(-np.inf, np.inf), (1e-6, np.inf)],
        'loglogistic': [(-np.inf, np.inf), (1e-6, np.inf), (1e-6, np.inf)],
        'lognorm': [(1e-6, np.inf), (1e-6, np.inf)],
        'lognorm2': [(-np.inf, np.inf), (1e-6, np.inf)],
        'betageneral': [(1e-6, np.inf), (1e-6, np.inf), (-np.inf, np.inf), (-np.inf, np.inf)],
        'betasubj': [(1e-6, np.inf), (1e-6, np.inf), (-np.inf, np.inf), (-np.inf, np.inf)],
        'burr12': [(-np.inf, np.inf), (1e-6, np.inf), (1e-6, np.inf), (1e-6, np.inf)],
        'pert': [(-np.inf, np.inf), (-np.inf, np.inf), (-np.inf, np.inf)],
        'reciprocal': [(-np.inf, np.inf), (1e-6, np.inf)],
        'rayleigh': [(1e-6, np.inf)],
        'weibull': [(1e-6, np.inf), (1e-6, np.inf)],
        'pearson5': [(1e-6, np.inf), (1e-6, np.inf)],
        'pearson6': [(1e-6, np.inf), (1e-6, np.inf), (1e-6, np.inf)],
        'pareto2': [(1e-6, np.inf), (1e-6, np.inf)],
        'pareto': [(1e-6, np.inf), (1e-6, np.inf)],
        'levy': [(-np.inf, np.inf), (1e-6, np.inf)],
        'invgauss': [(1e-6, np.inf), (1e-6, np.inf)],
        'erf': [(1e-6, np.inf)],
    }
    return bounds.get(dist_name, None)

def fit_density_or_cumulative(x, y, fit_type, metric='RMSE'):
    """
    对密度或累积数据进行分布拟合
    fit_type: 'density_raw', 'density_norm', 'cumulative'
    返回结果列表，格式与fit_distributions类似
    """
    results = []
    dist_names = CONTINUOUS_DIST_NAMES  # 只使用连续分布
    total = len(dist_names)
    for idx, dist_name in enumerate(dist_names):
        set_status_bar(f"正在拟合分布 {dist_name}... ({idx+1}/{total})")
        fit_result = _fit_distribution_to_density_or_cumulative(x, y, dist_name, fit_type)
        if fit_result is None:
            results.append({
                'dist_name': dist_name,
                'metric_name': metric,
                'metric_value': np.inf,
                'params': [],
                'base_param_count': get_base_param_count(dist_name),
                'shift': 0.0,
                'ignore_shift': True,
                'nll': np.inf,
                'k': get_base_param_count(dist_name),
                'n': len(x),
                'dist_obj': None,
                'bootstrap': None,
                'rmse': np.inf
            })
            continue
        params, rmse, pdf_func, cdf_func = fit_result
        # 创建包装对象
        wrapper = DistributionWrapper(dist_name, params, 0.0, None, pdf_func=pdf_func, cdf_func=cdf_func)
        results.append({
            'dist_name': dist_name,
            'metric_name': metric,
            'metric_value': rmse,
            'params': params,
            'base_param_count': get_base_param_count(dist_name),
            'shift': 0.0,
            'ignore_shift': True,
            'nll': np.nan,
            'k': len(params),
            'n': len(x),
            'dist_obj': wrapper,
            'bootstrap': None,
            'rmse': rmse
        })
        print(f"{dist_name}: RMSE={rmse:.6g}, params={params}")
    results.sort(key=lambda x: x['metric_value'] if np.isfinite(x['metric_value']) else np.inf)
    return results

# ================== Bootstrap 模拟 ==================
def bootstrap_estimate(data: np.ndarray, dist_name: str, base_param_count: int, n_bootstrap: int, conf_level: float = 0.95, dist_display_name: str = ""):
    n = len(data)
    param_list = []
    ks_vals = []
    ad_vals = []
    chisq_vals = []
    for i in range(n_bootstrap):
        if dist_display_name:
            percent = (i + 1) / n_bootstrap * 100
            set_status_bar(f"Bootstrap [{dist_display_name}]: 迭代 {i+1}/{n_bootstrap} ({percent:.1f}%)")
        resampled = np.random.choice(data, size=n, replace=True)
        fit_result = fit_scipy_distribution(resampled, dist_name)
        if fit_result is not None:
            params, shift, scipy_dist = fit_result
            wrapper = DistributionWrapper(dist_name, params, shift, scipy_dist)
            combined = list(params)
            if abs(shift) > 1e-8:
                combined.append(shift)
            param_list.append(combined)
            ks_vals.append(compute_ks_statistic(data, wrapper))
            ad_vals.append(compute_ad_statistic(data, wrapper))
            chisq_vals.append(compute_chisq_statistic(data, wrapper))
        else:
            param_list.append([np.nan] * (len(param_list[0]) if param_list else (base_param_count + 1)))
            ks_vals.append(np.nan)
            ad_vals.append(np.nan)
            chisq_vals.append(np.nan)
    param_matrix = np.array(param_list, dtype=float)
    alpha = (1 - conf_level) / 2
    lower = np.nanpercentile(param_matrix, 100 * alpha, axis=0)
    upper = np.nanpercentile(param_matrix, 100 * (1 - alpha), axis=0)
    width = upper - lower
    return param_matrix, lower, upper, width, ks_vals, ad_vals, chisq_vals

# ================== 主拟合流程 ==================
def fit_distributions(data: np.ndarray, dist_type: str, metric: str,
                      do_bootstrap: bool, bootstrap_conf: float,
                      n_bootstrap: int = 100, bootstrap_dist_names: List[str] = None):
    results = []
    dist_names = CONTINUOUS_DIST_NAMES if dist_type == 'continuous' else DISCRETE_DISTRIBUTIONS
    total = len(dist_names)
    all_integer = is_integer_like(data) if dist_type == 'discrete' else True

    for idx, dist_name in enumerate(dist_names):
        percent = (idx + 1) / total * 100
        set_status_bar(f"正在拟合分布 {dist_name}... ({idx+1}/{total}, {percent:.1f}%)")
        if dist_type == 'discrete' and not all_integer:
            print(f"拟合 {dist_name} 失败：离散分布要求数据为整数")
            results.append({
                'dist_name': dist_name,
                'metric_name': metric,
                'metric_value': np.inf,
                'params': [],
                'base_param_count': get_base_param_count(dist_name),
                'shift': 0.0,
                'ignore_shift': True,
                'nll': np.inf,
                'k': 0,
                'n': len(data),
                'dist_obj': None,
                'bootstrap': None
            })
            continue

        fit_result = fit_scipy_distribution(data, dist_name)
        if fit_result is None:
            print(f"拟合 {dist_name} 失败")
            results.append({
                'dist_name': dist_name,
                'metric_name': metric,
                'metric_value': np.inf,
                'params': [],
                'base_param_count': get_base_param_count(dist_name),
                'shift': 0.0,
                'ignore_shift': True,
                'nll': np.inf,
                'k': 0,
                'n': len(data),
                'dist_obj': None,
                'bootstrap': None
            })
            continue

        params, shift, scipy_dist = fit_result
        pdf_func = None
        cdf_func = None
        if scipy_dist is None:
            pdf_func, cdf_func = _get_custom_distribution_funcs(dist_name, params)
        wrapper = DistributionWrapper(dist_name, params, shift, scipy_dist, pdf_func=pdf_func, cdf_func=cdf_func)
        if not wrapper.is_valid():
            print(f"拟合 {dist_name} 失败（无效）")
            results.append({
                'dist_name': dist_name,
                'metric_name': metric,
                'metric_value': np.inf,
                'params': [],
                'base_param_count': get_base_param_count(dist_name),
                'shift': 0.0,
                'ignore_shift': True,
                'nll': np.inf,
                'k': 0,
                'n': len(data),
                'dist_obj': None,
                'bootstrap': None
            })
            continue

        nll = -np.sum(wrapper.logpdf(data))
        if np.isnan(nll) or np.isinf(nll):
            print(f"拟合 {dist_name} 失败（似然无效）")
            results.append({
                'dist_name': dist_name,
                'metric_name': metric,
                'metric_value': np.inf,
                'params': [],
                'base_param_count': get_base_param_count(dist_name),
                'shift': 0.0,
                'ignore_shift': True,
                'nll': np.inf,
                'k': 0,
                'n': len(data),
                'dist_obj': None,
                'bootstrap': None
            })
            continue

        k = len(params) + (0 if abs(shift) < 1e-8 else 1)
        n = len(data)
        if metric == 'AICc':
            metric_val = compute_aicc(nll, k, n)
        elif metric == 'BIC':
            metric_val = compute_bic(nll, k, n)
        elif metric == 'KS':
            metric_val = compute_ks_statistic(data, wrapper)
        elif metric == 'AD':
            metric_val = compute_ad_statistic(data, wrapper)
        elif metric == 'AVlog':
            metric_val = compute_avlog(data, wrapper)
        elif metric == 'chisq':
            metric_val = compute_chisq_statistic(data, wrapper)
        else:
            metric_val = compute_aicc(nll, k, n)

        results.append({
            'dist_name': dist_name,
            'metric_name': metric,
            'metric_value': metric_val,
            'params': params,
            'base_param_count': get_base_param_count(dist_name),
            'shift': shift,
            'ignore_shift': abs(shift) < 1e-8,
            'nll': nll,
            'k': k,
            'n': n,
            'dist_obj': wrapper,
            'bootstrap': None
        })
        print(f"{dist_name}: {metric}={metric_val:.6g}, params={params}, shift={shift:.4g}")

    results.sort(key=lambda x: x['metric_value'] if np.isfinite(x['metric_value']) else np.inf)

    if do_bootstrap and results:
        if bootstrap_dist_names is None:
            dists_to_bootstrap = [res for res in results if np.isfinite(res['metric_value'])]
        else:
            dists_to_bootstrap = [res for res in results if res['dist_name'] in bootstrap_dist_names and np.isfinite(res['metric_value'])]
        for i, res in enumerate(dists_to_bootstrap):
            total_boot = len(dists_to_bootstrap)
            percent = (i + 1) / total_boot * 100
            set_status_bar(f"正在 Bootstrap 模拟 [{res['dist_name']}]... ({i+1}/{total_boot}, {percent:.1f}%)")
            try:
                param_matrix, lower, upper, width, ks_vals, ad_vals, chisq_vals = bootstrap_estimate(
                    data, res['dist_name'], res['base_param_count'], n_bootstrap, bootstrap_conf,
                    dist_display_name=res['dist_name']
                )
                res['bootstrap'] = {
                    'param_matrix': param_matrix,
                    'conf_lower': lower,
                    'conf_upper': upper,
                    'conf_width': width,
                    'ks_vals': ks_vals,
                    'ad_vals': ad_vals,
                    'chisq_vals': chisq_vals,
                    'conf_level': bootstrap_conf
                }
                print(f"Bootstrap 完成 {res['dist_name']}")
            except Exception as e:
                print(f"Bootstrap 失败 {res['dist_name']}: {e}")
                res['bootstrap'] = None
    return results

def get_base_param_count(dist_name: str) -> Optional[int]:
    param_counts = {
        'normal': 2, 'uniform': 2, 'erf': 1, 'gamma': 2, 'erlang': 2, 'beta': 2,
        'poisson': 1, 'chisq': 1, 'f': 2, 'student': 1, 'expon': 1, 'bernoulli': 1,
        'triang': 3, 'binomial': 2, 'cauchy': 2, 'dagum': 4, 'doubletriang': 4,
        'extvalue': 2, 'extvaluemin': 2, 'fatiguelife': 3, 'frechet': 3,
        'hypsecant': 2, 'johnsonsb': 4, 'johnsonsu': 4, 'kumaraswamy': 4,
        'laplace': 2, 'logistic': 2, 'loglogistic': 3, 'lognorm': 2, 'lognorm2': 2,
        'betageneral': 4, 'betasubj': 4, 'burr12': 4, 'pert': 3, 'reciprocal': 2,
        'rayleigh': 1, 'weibull': 2, 'pearson5': 2, 'pearson6': 3, 'pareto2': 2,
        'pareto': 2, 'levy': 2, 'negbin': 2, 'invgauss': 2, 'geomet': 1,
        'hypergeo': 3, 'intuniform': 2, 'trigen': 5
    }
    return param_counts.get(dist_name)

def format_drisk_formula(dist_name: str, params: List[float], base_param_count: int,
                         shift: float, ignore_shift: bool) -> str:
    if not params:
        return "N/A"
    func_name = DIST_TO_DRISK_NAME.get(dist_name, f'Drisk{dist_name.capitalize()}')
    base_params = params[:base_param_count]
    param_strs = []
    for p in base_params:
        if abs(p - round(p, 6)) < 1e-6:
            param_strs.append(f"{p:.6g}")
        else:
            param_strs.append(f"{p:.6f}")
    if not ignore_shift and abs(shift) > 1e-6:
        param_strs.append(f"DriskShift({shift:.6g})")
    return f"{func_name}({', '.join(param_strs)})"

# ================== GUI 选择对话框 ==================
def _show_distribution_selection_dialog(results, metric_name):
    try:
        root = tk.Tk()
        root.title("选择拟合分布")
        root.geometry("800x500")
        root.resizable(True, True)
        root.attributes('-topmost', True)
        root.update_idletasks()
        x = (root.winfo_screenwidth() // 2) - (800 // 2)
        y = (root.winfo_screenheight() // 2) - (500 // 2)
        root.geometry(f"+{x}+{y}")

        label = ttk.Label(root, text=f"按 {metric_name} 排序的拟合结果（双击选择分布）", padding=5)
        label.pack()

        frame = ttk.Frame(root)
        frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        columns = ('distribution', 'metric', 'formula')
        tree = ttk.Treeview(frame, columns=columns, show='headings', selectmode='browse')
        tree.heading('distribution', text='分布类型')
        tree.heading('metric', text=metric_name)
        tree.heading('formula', text='Drisk 公式')
        tree.column('distribution', width=150)
        tree.column('metric', width=100)
        tree.column('formula', width=500)

        for idx, res in enumerate(results):
            if np.isfinite(res['metric_value']):
                metric_display = f"{res['metric_value']:.6g}"
                formula = format_drisk_formula(
                    res['dist_name'], res['params'], res['base_param_count'],
                    res['shift'], res['ignore_shift']
                )
            else:
                metric_display = "N/A"
                formula = "N/A"
            tree.insert('', 'end', iid=idx, values=(res['dist_name'], metric_display, formula))

        vsb = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=vsb.set)
        tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)

        selected_idx = tk.IntVar()
        selected_idx.set(-1)

        def on_confirm():
            selection = tree.selection()
            if selection:
                idx = int(selection[0])
                if np.isfinite(results[idx]['metric_value']):
                    selected_idx.set(idx)
                    root.destroy()
                else:
                    messagebox.showwarning("无效分布", "该分布拟合失败，请选择其他分布")
            else:
                messagebox.showwarning("未选择", "请先选择一个分布")

        def on_double_click(event):
            selection = tree.selection()
            if selection:
                idx = int(selection[0])
                if np.isfinite(results[idx]['metric_value']):
                    selected_idx.set(idx)
                    root.destroy()
                else:
                    messagebox.showwarning("无效分布", "该分布拟合失败，请选择其他分布")

        tree.bind("<Double-1>", on_double_click)

        def on_cancel():
            selected_idx.set(-1)
            root.destroy()

        btn_frame = ttk.Frame(root)
        btn_frame.pack(pady=10)
        ttk.Button(btn_frame, text="确定", width=10, command=on_confirm).pack(side=tk.LEFT, padx=10)
        ttk.Button(btn_frame, text="取消", width=10, command=on_cancel).pack(side=tk.RIGHT, padx=10)

        root.mainloop()
        idx = selected_idx.get()
        if idx >= 0:
            return results[idx]
        else:
            return None
    except Exception as e:
        print(f"创建对话框失败: {e}")
        traceback.print_exc()
        return None

def _select_distributions_multiple(results):
    try:
        root = tk.Tk()
        root.title("选择要进行 Bootstrap 的分布")
        root.geometry("500x400")
        root.resizable(True, True)
        root.attributes('-topmost', True)
        root.update_idletasks()
        x = (root.winfo_screenwidth() // 2) - (500 // 2)
        y = (root.winfo_screenheight() // 2) - (400 // 2)
        root.geometry(f"+{x}+{y}")

        label = ttk.Label(root, text="请勾选要进行 Bootstrap 的分布（可多选）:", padding=5)
        label.pack()

        frame = ttk.Frame(root)
        frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        listbox = tk.Listbox(frame, selectmode=tk.MULTIPLE, exportselection=False)
        scrollbar = ttk.Scrollbar(frame, orient=tk.VERTICAL, command=listbox.yview)
        listbox.configure(yscrollcommand=scrollbar.set)
        listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # 过滤掉 hypergeo 分布，不提供 Bootstrap 选项
        valid_res = [res for res in results if np.isfinite(res['metric_value']) and res['dist_name'] != 'hypergeo']
        for res in valid_res:
            listbox.insert(tk.END, res['dist_name'])

        def on_confirm():
            selected_indices = listbox.curselection()
            selected_names = [listbox.get(i) for i in selected_indices]
            root.destroy()
            root.selected_names = selected_names

        def on_cancel():
            root.selected_names = None
            root.destroy()

        btn_frame = ttk.Frame(root)
        btn_frame.pack(pady=10)
        ttk.Button(btn_frame, text="确定", width=10, command=on_confirm).pack(side=tk.LEFT, padx=10)
        ttk.Button(btn_frame, text="取消", width=10, command=on_cancel).pack(side=tk.RIGHT, padx=10)

        root.mainloop()
        return getattr(root, 'selected_names', None)
    except Exception as e:
        print(f"创建多选对话框失败: {e}")
        traceback.print_exc()
        return None

# ================== DriskFitDist 宏 ==================
@xl_macro
def DriskFitDist():
    if not PYXLL_AVAILABLE:
        print("错误: pyxll 不可用")
        return
    app = xl_app()
    if not app:
        return

    global _last_fit_result

    # 选择数据区域
    suggested_range = get_current_column_range_string()
    range_prompt = "请选择数据区域。\n对于样本数据（类型1或2），请选择单列数值。\n对于密度/累积拟合（类型3-5），请选择两列：X 和 Y（密度或累积概率）。"
    if suggested_range:
        range_prompt += f"\n默认范围: {suggested_range}"
    try:
        range_obj = _get_excel_input(
            app, range_prompt, "DriskFitDist - 选择数据区域",
            default=suggested_range, input_type=8
        )
        if range_obj is None:
            range_input = _get_excel_input(
                app, "请选择或输入数据所在的 Excel 范围 (例如 A1:A100):",
                "DriskFitDist - 输入数据范围", default=suggested_range, input_type=1
            )
            if not range_input:
                return
            data_vals = parse_excel_range(range_input)
            # 判断是单列还是多列
            x_vals, y_vals = [], []
            if len(data_vals) == 0:
                return
            # 简单认为用户输入了单列
            is_two_col = False
        else:
            # 尝试解析两列
            x_vals, y_vals = parse_excel_selection_2cols(range_obj)
            if len(x_vals) > 0 and len(y_vals) > 0:
                is_two_col = True
            else:
                data_vals = parse_excel_selection(range_obj)
                is_two_col = False
    except Exception:
        range_input = _get_excel_input(
            app, "请输入数据所在的 Excel 范围 (例如 A1:A100):",
            "DriskFitDist - 输入数据范围", default=suggested_range, input_type=1
        )
        if not range_input:
            return
        data_vals = parse_excel_range(range_input)
        is_two_col = False

    # 数据类型选择
    data_type_choice = _get_excel_input(
        app,
        "请选择数据类型:\n1-连续分布（样本数据）\n2-离散分布（样本数据）\n3-连续密度（未标准化）\n4-连续密度（标准化）\n5-连续累积",
        "DriskFitDist - 数据类型", default="1", input_type=1
    )
    if data_type_choice is None:
        return
    choice = parse_int_input(data_type_choice)
    if choice is None:
        messagebox.showerror("错误", "无效选择")
        return

    # 对于类型3-5，要求两列数据
    if choice in [3,4,5]:
        if not is_two_col:
            messagebox.showerror("错误", "对于密度或累积拟合，请选择两列数据：第一列为X值，第二列为密度或累积概率。")
            return
        x_data = np.array(x_vals, dtype=float)
        y_data = np.array(y_vals, dtype=float)
        if len(x_data) < 3:
            messagebox.showerror("错误", "数据点太少，至少需要3个点。")
            return
        # 拟合密度/累积
        fit_type_map = {3: 'density_raw', 4: 'density_norm', 5: 'cumulative'}
        fit_type = fit_type_map[choice]
        metric = 'RMSE'  # 固定使用RMSE
        set_status_bar("正在拟合分布...")
        try:
            fit_results = fit_density_or_cumulative(x_data, y_data, fit_type, metric)
        except Exception as e:
            traceback.print_exc()
            messagebox.showerror("错误", f"拟合过程中发生错误: {e}")
            clear_status_bar()
            return
        clear_status_bar()
        if not fit_results:
            messagebox.showinfo("提示", "没有成功拟合任何分布。")
            return
        # 保存结果
        global _last_fit_result
        _last_fit_result = {
            'data': None,
            'dist_type': f'fit_type_{fit_type}',
            'metric': metric,
            'results': fit_results,
            'bootstrap_done': False,
            'x_data': x_data,
            'y_data': y_data,
            'fit_type': fit_type
        }
        # 选择分布
        selected_res = _show_distribution_selection_dialog(fit_results, metric)
        if selected_res is None:
            messagebox.showinfo("提示", "未选择任何分布。")
            return
        # 输出公式
        output_cell = _get_excel_input(
            app,
            "请选择要写入公式的单元格 (例如 Sheet1!A1):",
            "选择输出单元格",
            input_type=8
        )
        if output_cell:
            formula_str = format_drisk_formula(
                selected_res['dist_name'], selected_res['params'],
                selected_res['base_param_count'], selected_res['shift'],
                selected_res['ignore_shift']
            )
            try:
                output_cell.Formula = "=" + formula_str
                messagebox.showinfo("完成", f"公式已写入 {output_cell.Address}")
            except Exception as e:
                messagebox.showerror("错误", f"写入公式失败: {e}")
        return

    # 原有流程：样本数据拟合（类型1或2）
    if choice == 1:
        dist_type = 'continuous'
    elif choice == 2:
        dist_type = 'discrete'
    else:
        messagebox.showerror("错误", "暂未定义该选项")
        return

    # 对于样本数据，data_vals应该存在
    if is_two_col:
        # 如果用户误选了多列，尝试只用第一列
        data_vals = x_vals
    data = np.array(data_vals, dtype=float)
    if data.size == 0 or not is_numeric_data(data.tolist()):
        messagebox.showerror("错误", "输入的数据不是纯数字或为空。")
        return

    metric_options = "\n".join([f"{i+1}-{m}" for i, m in enumerate(FIT_METRICS)])
    metric_choice = _get_excel_input(
        app, f"请选择主要判断指标:\n{metric_options}",
        "DriskFitDist - 拟合指标", default="1", input_type=1
    )
    if metric_choice is None:
        return
    metric_idx = parse_int_input(metric_choice)
    if metric_idx is None or metric_idx < 1 or metric_idx > len(FIT_METRICS):
        messagebox.showerror("错误", "无效选择")
        return
    metric = FIT_METRICS[metric_idx - 1]

    bootstrap_choice = False
    bootstrap_input = _get_excel_input(
        app, "是否使用 Bootstrap 估计参数？\n1-是\n0-否",
        "DriskFitDist - Bootstrap 估计", default="0", input_type=1
    )
    if bootstrap_input is None:
        return
    bootstrap_choice_value = parse_int_input(bootstrap_input)
    if bootstrap_choice_value == 1:
        bootstrap_choice = True
    elif bootstrap_choice_value == 0:
        bootstrap_choice = False
    else:
        messagebox.showerror("错误", "无效选择")
        return

    bootstrap_conf = 0.95
    n_bootstrap = 100
    bootstrap_dist_names = None

    if bootstrap_choice:
        n_bootstrap_input = _get_excel_input(
            app, "请输入 Bootstrap 次数（正整数，推荐 100-1000）:", "DriskFitDist - Bootstrap 次数",
            default="100", input_type=1
        )
        if n_bootstrap_input is None:
            return
        n_bootstrap = parse_int_input(n_bootstrap_input)
        if n_bootstrap is None or n_bootstrap <= 0:
            messagebox.showerror("错误", "无效的 Bootstrap 次数，请输入正整数")
            return

        conf_input = _get_excel_input(
            app, "请输入置信水平 (0-1):", "DriskFitDist - 置信区间",
            default="0.95", input_type=1
        )
        if conf_input is None:
            return
        bootstrap_conf = parse_float_input(conf_input)
        if bootstrap_conf is None or not 0 < bootstrap_conf < 1:
            messagebox.showerror("错误", "无效的置信水平，请输入 0 到 1 之间的数字")
            return

        set_status_bar("正在初步拟合分布...")
        print("初步拟合分布（不含 Bootstrap）...")
        temp_results = fit_distributions(data, dist_type, metric, False, bootstrap_conf, n_bootstrap, None)
        if not temp_results or not any(np.isfinite(r['metric_value']) for r in temp_results):
            messagebox.showinfo("提示", "没有成功拟合任何分布，无法进行 Bootstrap。")
            return

        selected_names = _select_distributions_multiple(temp_results)
        if selected_names is None or len(selected_names) == 0:
            messagebox.showinfo("提示", "未选择任何分布，将不进行 Bootstrap。")
            bootstrap_choice = False
        else:
            bootstrap_dist_names = selected_names
            print(f"将对以下分布进行 Bootstrap: {bootstrap_dist_names}")

    try:
        set_status_bar("正在拟合分布...")
        print("开始拟合分布...")
        fit_results = fit_distributions(data, dist_type, metric, bootstrap_choice, bootstrap_conf,
                                        n_bootstrap, bootstrap_dist_names)
        _last_fit_result = {
            'data': data,
            'dist_type': dist_type,
            'metric': metric,
            'results': fit_results,
            'bootstrap_done': bootstrap_choice
        }
    except Exception as e:
        traceback.print_exc()
        messagebox.showerror("错误", f"拟合过程中发生错误: {e}")
        clear_status_bar()
        return
    finally:
        clear_status_bar()

    if not fit_results:
        messagebox.showinfo("提示", "没有成功拟合任何分布。")
        return

    selected_res = _show_distribution_selection_dialog(fit_results, metric)
    if selected_res is None:
        messagebox.showinfo("提示", "未选择任何分布。")
        return

    output_cell = _get_excel_input(
        app,
        "请选择要写入公式的单元格 (例如 Sheet1!A1):",
        "选择输出单元格",
        input_type=8
    )
    if output_cell:
        formula_str = format_drisk_formula(
            selected_res['dist_name'], selected_res['params'],
            selected_res['base_param_count'], selected_res['shift'],
            selected_res['ignore_shift']
        )
        try:
            output_cell.Formula = "=" + formula_str
            messagebox.showinfo("完成", f"公式已写入 {output_cell.Address}")
        except Exception as e:
            messagebox.showerror("错误", f"写入公式失败: {e}")

# ================== DriskFitInfo 宏 ==================
@xl_macro
def DriskFitInfo():
    global _last_fit_result
    if not _last_fit_result:
        messagebox.showinfo("提示", "没有上一次拟合的结果。")
        return

    result = _last_fit_result
    # 判断是否为密度/累积拟合
    if 'fit_type' in result:
        # 密度/累积拟合信息显示
        results = result['results']
        metric = result['metric']
        fit_type = result['fit_type']
        x_data = result.get('x_data', [])
        y_data = result.get('y_data', [])

        root = tk.Tk()
        root.title("拟合信息 - 密度/累积拟合")
        root.geometry("900x600")
        root.attributes('-topmost', True)

        notebook = ttk.Notebook(root)
        notebook.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # 拟合结果表格
        frame1 = ttk.Frame(notebook)
        notebook.add(frame1, text="拟合结果")
        tree1 = ttk.Treeview(frame1, columns=('dist', 'metric', 'params'), show='headings')
        tree1.heading('dist', text='分布类型')
        tree1.heading('metric', text=metric)
        tree1.heading('params', text='参数')
        tree1.column('dist', width=150)
        tree1.column('metric', width=100)
        tree1.column('params', width=600)
        for res in results:
            if np.isfinite(res['metric_value']):
                metric_display = f"{res['metric_value']:.6g}"
                param_str = ', '.join([f"{p:.6g}" for p in res['params']])
            else:
                metric_display = "N/A"
                param_str = "N/A"
            tree1.insert('', tk.END, values=(res['dist_name'], metric_display, param_str))
        scroll1 = ttk.Scrollbar(frame1, orient=tk.VERTICAL, command=tree1.yview)
        tree1.configure(yscrollcommand=scroll1.set)
        tree1.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scroll1.pack(side=tk.RIGHT, fill=tk.Y)

        # 绘制拟合曲线（可选，简单显示）
        frame_plot = ttk.Frame(notebook)
        notebook.add(frame_plot, text="拟合曲线预览")
        try:
            import matplotlib
            matplotlib.use('TkAgg')
            import matplotlib.pyplot as plt
            from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
            fig, ax = plt.subplots(figsize=(6,4))
            ax.plot(x_data, y_data, 'o', label='实际数据', markersize=4)
            # 显示前3个最佳拟合
            top3 = [res for res in results if np.isfinite(res['metric_value'])][:3]
            for res in top3:
                dist_obj = res['dist_obj']
                if dist_obj:
                    x_smooth = np.linspace(min(x_data), max(x_data), 200)
                    if fit_type in ['density_raw', 'density_norm']:
                        y_smooth = dist_obj.pdf(x_smooth)
                    else:
                        y_smooth = dist_obj.cdf(x_smooth)
                    ax.plot(x_smooth, y_smooth, label=res['dist_name'])
            ax.legend()
            ax.set_xlabel('X')
            if fit_type.startswith('density'):
                ax.set_ylabel('概率密度')
            else:
                ax.set_ylabel('累积概率')
            canvas = FigureCanvasTkAgg(fig, master=frame_plot)
            canvas.draw()
            canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
        except Exception as e:
            ttk.Label(frame_plot, text=f"无法绘制图表: {e}").pack()

        root.mainloop()
        return

    # 原有样本数据拟合信息
    data = result['data']
    results = result['results']
    metric = result['metric']
    bootstrap_done = result['bootstrap_done']

    root = tk.Tk()
    root.title("拟合信息")
    root.geometry("1000x700")
    root.attributes('-topmost', True)

    notebook = ttk.Notebook(root)
    notebook.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

    frame1 = ttk.Frame(notebook)
    notebook.add(frame1, text="拟合结果")
    tree1 = ttk.Treeview(frame1, columns=('dist', 'metric', 'formula'), show='headings')
    tree1.heading('dist', text='分布类型')
    tree1.heading('metric', text=metric)
    tree1.heading('formula', text='参数公式')
    tree1.column('dist', width=150)
    tree1.column('metric', width=100)
    tree1.column('formula', width=600)
    for res in results:
        if np.isfinite(res['metric_value']):
            metric_display = f"{res['metric_value']:.6g}"
            formula = format_drisk_formula(res['dist_name'], res['params'], res['base_param_count'],
                                           res['shift'], res['ignore_shift'])
        else:
            metric_display = "N/A"
            formula = "N/A"
        tree1.insert('', tk.END, values=(res['dist_name'], metric_display, formula))
    scroll1 = ttk.Scrollbar(frame1, orient=tk.VERTICAL, command=tree1.yview)
    tree1.configure(yscrollcommand=scroll1.set)
    tree1.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
    scroll1.pack(side=tk.RIGHT, fill=tk.Y)

    if bootstrap_done:
        frame2 = ttk.Frame(notebook)
        notebook.add(frame2, text="Bootstrap 参数估计")
        columns = ('dist', 'param_idx', 'fit_value', 'lower_5%', 'upper_5%', 'ci_length')
        tree2 = ttk.Treeview(frame2, columns=columns, show='headings')
        tree2.heading('dist', text='分布')
        tree2.heading('param_idx', text='参数位置')
        tree2.heading('fit_value', text='拟合参数值')
        tree2.heading('lower_5%', text='下尾5%值')
        tree2.heading('upper_5%', text='上尾5%值')
        tree2.heading('ci_length', text='95%区间长度')
        tree2.column('dist', width=150)
        tree2.column('param_idx', width=80)
        tree2.column('fit_value', width=120)
        tree2.column('lower_5%', width=120)
        tree2.column('upper_5%', width=120)
        tree2.column('ci_length', width=120)

        for res in results:
            if res.get('bootstrap') and np.isfinite(res['metric_value']):
                boot = res['bootstrap']
                fit_params = list(res['params'])
                if not res['ignore_shift'] and abs(res['shift']) > 1e-8:
                    fit_params.append(res['shift'])
                lower = boot['conf_lower']
                upper = boot['conf_upper']
                n_params = len(fit_params)
                if len(lower) != n_params or len(upper) != n_params:
                    n_params = min(len(fit_params), len(lower), len(upper))
                    fit_params = fit_params[:n_params]
                    lower = lower[:n_params]
                    upper = upper[:n_params]
                for i in range(n_params):
                    param_idx = i + 1
                    fit_val = fit_params[i]
                    low_val = lower[i]
                    up_val = upper[i]
                    length = up_val - low_val
                    tree2.insert('', tk.END, values=(
                        res['dist_name'], param_idx,
                        f"{fit_val:.6g}", f"{low_val:.6g}", f"{up_val:.6g}", f"{length:.6g}"
                    ))
        vsb2 = ttk.Scrollbar(frame2, orient=tk.VERTICAL, command=tree2.yview)
        tree2.configure(yscrollcommand=vsb2.set)
        tree2.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb2.pack(side=tk.RIGHT, fill=tk.Y)

        frame3 = ttk.Frame(notebook)
        notebook.add(frame3, text="Chi-Square")
        tree3 = ttk.Treeview(frame3, columns=('dist', 'value', 'pvalue'), show='headings')
        tree3.heading('dist', text='分布')
        tree3.heading('value', text='统计量')
        tree3.heading('pvalue', text='p值')
        for res in results:
            if res.get('bootstrap') and np.isfinite(res['metric_value']):
                boot = res['bootstrap']
                chisq_vals = boot['chisq_vals']
                mean_chisq = np.nanmean(chisq_vals)
                df = res['k'] - 1
                if df > 0:
                    p_val = 1 - sps.chi2.cdf(mean_chisq, df)
                else:
                    p_val = np.nan
                tree3.insert('', tk.END, values=(res['dist_name'], f"{mean_chisq:.4f}", f"{p_val:.4f}"))
        tree3.pack(fill=tk.BOTH, expand=True)

        frame4 = ttk.Frame(notebook)
        notebook.add(frame4, text="KS")
        tree4 = ttk.Treeview(frame4, columns=('dist', 'value', 'pvalue'), show='headings')
        tree4.heading('dist', text='分布')
        tree4.heading('value', text='KS统计量')
        tree4.heading('pvalue', text='p值')
        for res in results:
            if res.get('bootstrap') and np.isfinite(res['metric_value']):
                boot = res['bootstrap']
                ks_vals = boot['ks_vals']
                mean_ks = np.nanmean(ks_vals)
                n = res['n']
                p_val = sps.kstwobign.sf(mean_ks * np.sqrt(n))
                tree4.insert('', tk.END, values=(res['dist_name'], f"{mean_ks:.4f}", f"{p_val:.4f}"))
        tree4.pack(fill=tk.BOTH, expand=True)

        frame5 = ttk.Frame(notebook)
        notebook.add(frame5, text="AD")
        tree5 = ttk.Treeview(frame5, columns=('dist', 'value', 'pvalue'), show='headings')
        tree5.heading('dist', text='分布')
        tree5.heading('value', text='AD统计量')
        tree5.heading('pvalue', text='p值')
        for res in results:
            if res.get('bootstrap') and np.isfinite(res['metric_value']):
                boot = res['bootstrap']
                ad_vals = boot['ad_vals']
                mean_ad = np.nanmean(ad_vals)
                dist_obj = res['dist_obj']
                if dist_obj is not None:
                    original_ad = compute_ad_statistic(data, dist_obj)
                else:
                    original_ad = np.nan
                if np.isfinite(original_ad) and len(ad_vals) > 0:
                    p_val = (np.sum(ad_vals >= original_ad) + 1) / (len(ad_vals) + 1)
                else:
                    p_val = np.nan
                tree5.insert('', tk.END, values=(res['dist_name'], f"{mean_ad:.4f}", f"{p_val:.4f}"))
            else:
                tree5.insert('', tk.END, values=(res['dist_name'], "N/A", "N/A"))
        tree5.pack(fill=tk.BOTH, expand=True)

    root.mainloop()

# ================== 批量拟合宏 DriskFitBatch ==================
@xl_macro
def DriskFitBatch():
    """
    批量拟合多列数据，生成带相关矩阵的汇总工作表。
    """
    if not PYXLL_AVAILABLE:
        print("错误: pyxll 不可用")
        return
    app = xl_app()
    if not app:
        return

    # 1. 选择数据区域（多列）
    try:
        range_obj = app.InputBox(
            Prompt="请选择数据区域（多列，每列为一组样本数据，第一行为列名可选）：",
            Title="DriskFitBatch - 选择数据区域",
            Type=8
        )
        if range_obj is False:
            return
    except:
        xlcAlert("选择区域取消或失败")
        return

    # 解析区域，按列提取数据
    try:
        # 获取区域的 Value2 二维数组
        vals = range_obj.Value2
        if vals is None:
            xlcAlert("区域无数据")
            return
        # 转换为二维列表
        if not isinstance(vals, (list, tuple)):
            vals = [[vals]]
        # 获取行数和列数
        if isinstance(vals[0], (list, tuple)):
            n_rows = len(vals)
            n_cols = len(vals[0])
        else:
            n_rows = len(vals)
            n_cols = 1
            vals = [vals]  # 统一为二维

        # 提取每列数据（跳过第一行如果第一行是文本）
        col_data_list = []
        col_ranges = []  # 存储每列的地址字符串
        col_names = []   # 存储列名

        # 获取工作表对象
        sheet = range_obj.Worksheet
        start_row = range_obj.Row
        start_col = range_obj.Column
        end_row = start_row + n_rows - 1
        end_col = start_col + n_cols - 1

        for col_idx in range(n_cols):
            col_data = []
            # 遍历行，提取数值
            for row_idx in range(n_rows):
                val = vals[row_idx][col_idx] if n_cols > 1 else vals[row_idx]
                # 尝试转换为浮点数
                try:
                    num = float(val)
                    col_data.append(num)
                except (ValueError, TypeError):
                    # 非数值，跳过
                    pass
            if len(col_data) == 0:
                xlcAlert(f"第 {col_idx+1} 列没有有效数值数据，请检查。")
                return
            col_data_list.append(col_data)

            # 获取该列的 Excel 范围地址
            col_letter = col_num_to_letter(start_col + col_idx)
            first_row = start_row
            last_row = end_row
            # 如果第一行是文本且被跳过了，实际数据起始行可能需要调整
            # 简单起见，使用整个区域
            range_addr = f"{sheet.Name}!{col_letter}{first_row}:{col_letter}{last_row}"
            col_ranges.append(range_addr)

            # 获取列名：向上查找10格内的第一个非空字符串
            col_name = None
            # 查找当前列向上10行
            search_row = start_row - 1
            for _ in range(10):
                if search_row < 1:
                    break
                try:
                    cell_val = sheet.Cells(search_row, start_col + col_idx).Value
                    if cell_val is not None and isinstance(cell_val, str) and cell_val.strip():
                        col_name = cell_val.strip()
                        break
                except:
                    pass
                search_row -= 1
            if not col_name:
                # 使用第一个数据单元格的地址
                col_name = f"{col_letter}{start_row}"
            col_names.append(col_name)

    except Exception as e:
        xlcAlert(f"解析数据区域失败: {e}")
        traceback.print_exc()
        return

    n_cols = len(col_data_list)

    # 2. 选择拟合指标
    metric_options = "\n".join([f"{i+1}-{m}" for i, m in enumerate(FIT_METRICS)])
    try:
        metric_choice = app.InputBox(
            f"请选择拟合优度指标（用于选择最佳分布）:\n{metric_options}",
            "DriskFitBatch - 拟合指标",
            Default="1",
            Type=1
        )
        if metric_choice is False:
            return
        metric_idx = int(metric_choice) if metric_choice else 1
        if metric_idx < 1 or metric_idx > len(FIT_METRICS):
            metric_idx = 1
        metric = FIT_METRICS[metric_idx - 1]
    except:
        metric = "AICc"

    # 3. 对每列进行拟合
    set_status_bar("正在批量拟合分布...")
    best_results = []  # 每个元素为 (dist_name, params, base_param_count, shift, ignore_shift, dist_obj, metric_value)
    for idx, col_data in enumerate(col_data_list):
        set_status_bar(f"拟合第 {idx+1}/{n_cols} 列: {col_names[idx]} ...")
        data = np.array(col_data, dtype=float)
        # 只拟合连续分布（可根据需要修改）
        fit_results = fit_distributions(data, 'continuous', metric, False, 0.95, 100, None)
        if not fit_results or not np.isfinite(fit_results[0]['metric_value']):
            xlcAlert(f"第 {idx+1} 列（{col_names[idx]}）拟合失败，请检查数据。")
            clear_status_bar()
            return
        best = fit_results[0]
        best_results.append(best)
        print(f"列 {col_names[idx]} 最佳分布: {best['dist_name']}, {metric}={best['metric_value']:.6g}")
    clear_status_bar()

    # 4. 计算原始数据的相关系数矩阵
    # 构造原始数据矩阵（n_samples x n_cols）
    max_len = max(len(col) for col in col_data_list)
    # 对齐长度（截断到最短列，或者用NaN填充？通常使用完整数据，取共同行数）
    # 简单起见，取最小行数
    min_len = min(len(col) for col in col_data_list)
    data_matrix = np.zeros((min_len, n_cols), dtype=float)
    for i in range(n_cols):
        data_matrix[:, i] = col_data_list[i][:min_len]
    # 计算皮尔逊相关系数
    corr_matrix = np.corrcoef(data_matrix, rowvar=False)
    # 处理 NaN（如果某列标准差为0，相关系数为NaN，填充0）
    corr_matrix = np.nan_to_num(corr_matrix, nan=0.0)

    # 5. 创建工作表 "拟合结果"
    wb = app.ActiveWorkbook
    sheet_name = "拟合结果"
    suffix = 0
    original_name = sheet_name
    while True:
        try:
            ws = wb.Worksheets(sheet_name)
            suffix += 1
            sheet_name = f"{original_name}({suffix})"
        except:
            # 工作表不存在，创建
            ws = wb.Worksheets.Add()
            ws.Name = sheet_name
            break
    # 移动到最后
    ws.Move(After=wb.Worksheets(wb.Worksheets.Count))

    # 6. 写入表头和数据
    # 行1: 名称
    ws.Cells(1, 1).Value = "名称"
    for j, name in enumerate(col_names):
        ws.Cells(1, j+2).Value = name
    # 行2: 区域
    ws.Cells(2, 1).Value = "区域"
    for j, rng_addr in enumerate(col_ranges):
        ws.Cells(2, j+2).Value = rng_addr
    # 行3: 最佳拟合（不带 DriskCorrmat）
    ws.Cells(3, 1).Value = "最佳拟合"
    best_formulas_no_corr = []
    for j, res in enumerate(best_results):
        formula = format_drisk_formula(
            res['dist_name'], res['params'], res['base_param_count'],
            res['shift'], res['ignore_shift']
        )
        ws.Cells(3, j+2).Value = formula
        best_formulas_no_corr.append(formula)
    # 行4: 公式（带 DriskCorrmat，稍后填充）
    ws.Cells(4, 1).Value = "公式"
    # 行5: 指标值
    ws.Cells(5, 1).Value = metric
    for j, res in enumerate(best_results):
        ws.Cells(5, j+2).Value = res['metric_value']
    # 行6: 最小值
    ws.Cells(6, 1).Value = "最小值"
    # 行7: 最大值
    ws.Cells(7, 1).Value = "最大值"
    # 行8: 均值
    ws.Cells(8, 1).Value = "均值"
    # 行9: 众数
    ws.Cells(9, 1).Value = "众数"
    # 行10: 中位数
    ws.Cells(10, 1).Value = "中位数"
    # 行11: 标准差
    ws.Cells(11, 1).Value = "标准差"

    # 计算每个分布的理论统计量
    for j, res in enumerate(best_results):
        # 构造分布对象
        params = res['params']
        shift = res['shift']
        scipy_dist = None
        pdf_func, cdf_func = None, None
        if res['dist_obj'] is None:
            # 重新创建
            _, _, scipy_dist = fit_scipy_distribution(np.array([0]), res['dist_name'])  # dummy
            # 但对于自定义分布，scipy_dist可能为None
            if scipy_dist is None:
                pdf_func, cdf_func = _get_custom_distribution_funcs(res['dist_name'], params)
        else:
            scipy_dist = res['dist_obj'].scipy_dist
            pdf_func = res['dist_obj'].pdf_func
            cdf_func = res['dist_obj'].cdf_func
        wrapper = DistributionWrapper(res['dist_name'], params, shift, scipy_dist, pdf_func, cdf_func)

        # 最小值
        min_val, max_val = wrapper.support()
        ws.Cells(6, j+2).Value = "-∞" if np.isneginf(min_val) else min_val
        ws.Cells(7, j+2).Value = "+∞" if np.isposinf(max_val) else max_val
        # 均值
        mean_val = wrapper.mean()
        ws.Cells(8, j+2).Value = mean_val if not np.isnan(mean_val) else "N/A"
        # 众数
        mode_val = wrapper.mode()
        ws.Cells(9, j+2).Value = mode_val if not np.isnan(mode_val) else "N/A"
        # 中位数
        median_val = wrapper.median()
        ws.Cells(10, j+2).Value = median_val if not np.isnan(median_val) else "N/A"
        # 标准差
        std_val = wrapper.std()
        ws.Cells(11, j+2).Value = std_val if not np.isnan(std_val) else "N/A"

    # 7. 创建相关性矩阵（下三角格式）
    # 确定矩阵放置的起始行（在统计信息下面空两行）
    matrix_start_row = 14
    # 确保有足够空间
    matrix_size = n_cols
    # 矩阵区域： (matrix_size+1) 行 x (matrix_size+1) 列（包含标签行/列）
    # 或者使用 Copula 风格：3行表头 + n行数据
    # 为了与 DriskMakeCorr 兼容，使用下三角格式，第一行第一列为矩阵名称，第一行为变量名，第一列为变量名
    # 先确定矩阵区域的左上角单元格
    matrix_top_left = ws.Cells(matrix_start_row, 1)
    # 创建矩阵区域范围
    matrix_range = ws.Range(
        matrix_top_left,
        ws.Cells(matrix_start_row + matrix_size, 1 + matrix_size)
    )
    # 生成矩阵名称
    existing_names = [name.Name for name in wb.Names]
    max_num = 0
    for name in existing_names:
        match = re.match(r'^相关矩阵_(\d+)$', name)
        if match:
            max_num = max(max_num, int(match.group(1)))
    matrix_name = f"相关矩阵_{max_num + 1}"

    # 准备标签（列名）
    labels = col_names[:]  # 使用列名作为变量名

    # 写入矩阵区域（下三角）
    # 使用 corrmat_functions 中的 _write_lower_tri_matrix 函数
    if CORRMAT_AVAILABLE:
        cmf._write_lower_tri_matrix(matrix_range, corr_matrix, labels, matrix_name=matrix_name)
    else:
        # 简易写入
        # 清空区域
        matrix_range.ClearContents()
        # 写入矩阵名称
        matrix_range.Cells(1, 1).Value = matrix_name
        # 写入行标签和列标签
        for i, label in enumerate(labels):
            matrix_range.Cells(1, i+2).Value = label
            matrix_range.Cells(i+2, 1).Value = label
        # 写入下三角
        for i in range(matrix_size):
            for j in range(i+1):
                matrix_range.Cells(i+2, j+2).Value = corr_matrix[i, j]
    # 设置命名区域
    if CORRMAT_AVAILABLE:
        cmf._set_named_range(app, matrix_name, matrix_range)
    else:
        try:
            wb.Names.Add(Name=matrix_name, RefersTo=matrix_range)
        except:
            pass

    # 8. 更新第四行公式，添加 DriskCorrmat 参数
    for j in range(n_cols):
        base_formula = best_formulas_no_corr[j]
        # 移除开头的 "=" 如果存在
        if base_formula.startswith('='):
            base_formula = base_formula[1:]
        # 确定位置参数（从1开始）
        pos = j + 1
        # 构造 DriskCorrmat 参数
        # 使用命名区域名称（带引号，因为可能包含中文）
        # DriskCorrmat 解析支持带引号的名称
        corrmat_arg = f'DriskCorrmat("{matrix_name}", {pos})'
        # 插入到参数列表末尾
        # 解析原公式，找到函数名和参数
        # 简单处理：找到左括号和右括号
        func_name_end = base_formula.find('(')
        if func_name_end == -1:
            continue
        func_name = base_formula[:func_name_end]
        args_part = base_formula[func_name_end+1:-1]  # 去掉最外层括号
        if args_part.strip():
            new_args = f"{args_part}, {corrmat_arg}"
        else:
            new_args = corrmat_arg
        new_formula = f"={func_name}({new_args})"
        ws.Cells(4, j+2).Value = new_formula

    # 9. 可选：格式化工作表
    # 自动调整列宽
    ws.Columns.AutoFit()
    # 添加标题说明
    ws.Cells(matrix_start_row - 2, 1).Value = "相关性矩阵"
    ws.Cells(matrix_start_row - 2, 1).Font.Bold = True

    xlcAlert(f"批量拟合完成！\n结果已写入工作表“{sheet_name}”。\n相关性矩阵命名区域：{matrix_name}。")
    clear_status_bar()

# 如果 corrmat_functions 中的 xlcAlert 未定义，使用 messagebox
try:
    from pyxll import xlcAlert
except:
    def xlcAlert(msg):
        messagebox.showinfo("提示", msg)
