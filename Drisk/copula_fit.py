# copula_fit.py
"""
Copula 拟合模块
提供 DriskFitCopula 和 DriskAppendCopula 两个宏，用于从历史数据拟合最优 Copula，
并将拟合结果附加到 Drisk 分布函数中。生成的 Copula 矩阵与现有系统完全兼容，
支持 DriskCheckCopula、DriskRemoveCopula、DriskMuteCopula 等宏。

增强功能（2025-03）：
- 精确极大似然法（模式 1）支持参数化边缘分布拟合：对每一列自动选择最优连续分布，
  使用其 CDF 转换为均匀分位数，代替非参数经验分布，显著提高多维依赖结构的建模精度。
"""

import re
import traceback
import numpy as np
import win32com.client
from pyxll import xl_macro, xl_app, xlcAlert
from typing import List, Tuple, Optional, Dict, Any, Callable
from scipy.optimize import minimize, minimize_scalar
from scipy.stats import norm, t, multivariate_normal, multivariate_t
from scipy.special import gammaln

# 复用现有模块的辅助函数
from copula_functions import (
    _get_excel_app, _get_named_range, _set_named_range,
    _read_matrix_from_range, _write_lower_tri_matrix,
    _ensure_copula_in_formula, COPLUA_TYPES,
    _read_data_range, _nearest_psd,
    _sample_positive_stable
)
from constants import DISTRIBUTION_FUNCTION_NAMES
from formula_parser import is_distribution_function

# 导入 cumul 经验分布函数（回退方案）
from dist_cumul import cumul_cdf

# ==================== 新增：导入分布拟合模块（参数化边缘） ====================
try:
    from distribution_fit import fit_distributions, DistributionWrapper, CONTINUOUS_DIST_NAMES, FIT_METRICS
    DISTFIT_AVAILABLE = True
except ImportError as e:
    DISTFIT_AVAILABLE = False
    print(f"警告: 无法导入 distribution_fit 模块，参数化边缘拟合不可用: {e}")

# ==================== 辅助函数 ====================
def _get_workbook():
    """获取当前活动工作簿"""
    app = _get_excel_app()
    return app.ActiveWorkbook

def _check_ties(data: np.ndarray, app=None) -> bool:
    """检查数据每列中是否存在并列值（ties）。"""
    n, d = data.shape
    has_ties = False
    tie_cols = []
    for j in range(d):
        col = data[:, j]
        if len(np.unique(col)) < len(col):
            has_ties = True
            tie_cols.append(j + 1)
    if has_ties:
        msg = (f"警告：数据中存在并列值（ties）。列 {', '.join(map(str, tie_cols))} 含有重复值。\n"
               "拟合过程中将使用平均秩次处理，但这可能影响 Copula 参数估计的精度。")
        if app:
            try:
                xlcAlert(msg)
            except:
                print(msg)
        else:
            print(msg)
    return has_ties

def _show_choice_dialog(title, message, option1, option2):
    """显示两个按钮的对话框，返回 1 或 2"""
    try:
        import tkinter as tk
        from tkinter import ttk
        root = tk.Tk()
        root.title(title)
        root.geometry("350x180")
        root.resizable(False, False)
        root.attributes('-topmost', True)
        root.update_idletasks()
        x = (root.winfo_screenwidth() // 2) - (350 // 2)
        y = (root.winfo_screenheight() // 2) - (180 // 2)
        root.geometry(f"+{x}+{y}")

        msg_frame = ttk.Frame(root, padding=10)
        msg_frame.pack(fill=tk.BOTH, expand=True)
        label = ttk.Label(msg_frame, text=message, wraplength=320, justify=tk.CENTER)
        label.pack(expand=True)

        btn_frame = ttk.Frame(root, padding=10)
        btn_frame.pack(fill=tk.X)
        result = tk.IntVar()
        result.set(0)

        def on_click(val):
            result.set(val)
            root.destroy()

        ttk.Button(btn_frame, text=option1, width=12, command=lambda: on_click(1)).pack(side=tk.LEFT, padx=10, expand=True)
        ttk.Button(btn_frame, text=option2, width=12, command=lambda: on_click(2)).pack(side=tk.RIGHT, padx=10, expand=True)

        root.mainloop()
        return result.get()
    except Exception:
        app = _get_excel_app()
        try:
            choice = app.InputBox(
                Prompt=f"{message}\n\n输入 1 选择“{option1}”，输入 2 选择“{option2}”：",
                Title=title, Type=1, Default="1"
            )
            if choice is False:
                return 0
            return int(choice)
        except:
            return 0

def _show_copula_selection_dialog(title, headers, rows):
    """显示一个美观的表格选择对话框，用于选择 Copula。"""
    try:
        import tkinter as tk
        from tkinter import ttk
        from tkinter import messagebox
        root = tk.Tk()
        root.title(title)
        root.geometry("600x400")
        root.resizable(True, True)
        root.attributes('-topmost', True)
        root.update_idletasks()
        x = (root.winfo_screenwidth() // 2) - (600 // 2)
        y = (root.winfo_screenheight() // 2) - (400 // 2)
        root.geometry(f"+{x}+{y}")

        label = ttk.Label(root, text="请选择要使用的 Copula（双击或按“确定”）：", padding=5)
        label.pack()

        frame = ttk.Frame(root)
        frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        tree = ttk.Treeview(frame, columns=headers, show='headings', selectmode='browse')
        for col in headers:
            tree.heading(col, text=col)
            tree.column(col, width=200, anchor='center')

        for i, row in enumerate(rows):
            tree.insert('', 'end', iid=i, values=row)

        vsb = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=vsb.set)
        tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)

        result = tk.IntVar()
        result.set(-1)

        def on_confirm():
            selection = tree.selection()
            if selection:
                result.set(int(selection[0]))
                root.destroy()
            else:
                messagebox.showwarning("未选择", "请先选择一个 Copula")

        def on_double_click(event):
            selection = tree.selection()
            if selection:
                result.set(int(selection[0]))
                root.destroy()

        tree.bind("<Double-1>", on_double_click)

        def on_cancel():
            result.set(-1)
            root.destroy()

        btn_frame = ttk.Frame(root)
        btn_frame.pack(pady=10)
        ttk.Button(btn_frame, text="确定", width=10, command=on_confirm).pack(side=tk.LEFT, padx=10)
        ttk.Button(btn_frame, text="取消", width=10, command=on_cancel).pack(side=tk.RIGHT, padx=10)

        root.mainloop()
        return result.get()
    except Exception as e:
        print(f"表格对话框失败，使用回退方案: {e}")
        app = _get_excel_app()
        items = [row[0] for row in rows]
        prompt = "请选择要使用的 Copula（输入数字）:\n\n"
        for i, item in enumerate(items, 1):
            prompt += f"{i}. {item}\n"
        try:
            choice = app.InputBox(prompt, title, Type=1, Default="1")
            if choice is False:
                return -1
            idx = int(choice) - 1
            if 0 <= idx < len(items):
                return idx
            return -1
        except:
            return -1

# ==================== 旋转 Copula 支持 ====================
ROTATION_MAP = {
    "FrankRX": ("Frank", 1),      # 垂直翻转 (Y轴) -> 第二列变换
    "ClaytonRX": ("Clayton", 0),  # 水平翻转 (X轴) -> 第一列变换
    "ClaytonRY": ("Clayton", 1),  # 垂直翻转 (Y轴) -> 第二列变换
    "GumbelRX": ("Gumbel", 0),    # 水平翻转 (X轴)
    "GumbelRY": ("Gumbel", 1),    # 垂直翻转 (Y轴)
}

def _apply_rotation(data: np.ndarray, rot_type: str) -> np.ndarray:
    """根据旋转类型对数据进行列变换。"""
    if rot_type not in ROTATION_MAP:
        return data.copy()
    base_type, col_idx = ROTATION_MAP[rot_type]
    rotated = data.copy()
    rotated[:, col_idx] = 1.0 - rotated[:, col_idx]
    rotated = np.clip(rotated, 1e-12, 1.0 - 1e-12)
    return rotated

def _base_type_from_rotated(rot_type: str) -> str:
    """从旋转类型获取基础 Copula 类型名称"""
    if rot_type in ROTATION_MAP:
        return ROTATION_MAP[rot_type][0]
    return rot_type

# ==================== 参数化边缘分布拟合（新增） ====================
def _fit_marginal_distribution_parametric(data_col: np.ndarray, metric: str = 'AICc'):
    """
    对单列数据拟合最优连续分布（使用 distribution_fit 模块）。
    返回 (cdf_func, dist_name, params, shift, wrapper) 或 None。
    """
    if not DISTFIT_AVAILABLE:
        return None
    try:
        # 只拟合连续分布
        results = fit_distributions(data_col, 'continuous', metric, False, 0.95, 100, None)
        if not results or not np.isfinite(results[0]['metric_value']):
            return None
        best = results[0]
        # 构造 DistributionWrapper 对象以获取 cdf
        if best['dist_obj'] is None:
            # 重新创建 wrapper（理论上 fit_distributions 已经创建了 dist_obj）
            return None
        wrapper = best['dist_obj']
        return (wrapper.cdf, best['dist_name'], best['params'], best['shift'], wrapper)
    except Exception as e:
        print(f"参数化边缘拟合失败: {e}")
        return None

def _transform_to_uniform_parametric(data_matrix: np.ndarray, metric: str = 'AICc'):
    """
    对数据矩阵的每一列进行参数化边缘拟合，并用拟合的 CDF 转换为均匀分位数。
    返回 (pseudo_obs, marginal_info_list)。
    marginal_info_list 每个元素为 (dist_name, params, shift) 或 None。
    """
    n, d = data_matrix.shape
    pseudo = np.zeros((n, d), dtype=float)
    marginal_info = []
    for j in range(d):
        col = data_matrix[:, j]
        fit_res = _fit_marginal_distribution_parametric(col, metric)
        if fit_res is None:
            # 拟合失败，回退到经验分布（使用 cumul）
            print(f"列 {j+1} 参数化拟合失败，使用经验分布。")
            # 使用现有的 _compute_pseudo_obs_cumul 对该列单独处理
            from dist_cumul import cumul_cdf
            unique_vals, counts = np.unique(col, return_counts=True)
            m = len(unique_vals)
            cum_counts = np.cumsum(counts)
            p_vals = (cum_counts - counts + 0.5 * counts) / n
            # 添加辅助点
            x_full = np.concatenate(([unique_vals[0] - 1e-12], unique_vals, [unique_vals[-1] + 1e-12]))
            p_full = np.concatenate(([0.0], p_vals, [1.0]))
            col_u = np.array([cumul_cdf(v, x_full.tolist(), p_full.tolist()) for v in col])
            col_u = np.clip(col_u, 1e-12, 1.0 - 1e-12)
            pseudo[:, j] = col_u
            marginal_info.append(None)
        else:
            cdf_func, dist_name, params, shift, _ = fit_res
            # 计算每个点的 CDF 值
            col_u = cdf_func(col)
            # 处理可能出现的 0 或 1
            col_u = np.clip(col_u, 1e-12, 1.0 - 1e-12)
            pseudo[:, j] = col_u
            marginal_info.append((dist_name, params, shift))
            print(f"列 {j+1} 边缘分布: {dist_name}, 参数: {params}, shift={shift}")
    return pseudo, marginal_info

# ==================== 使用 cumul 经验分布计算伪观测值（原有） ====================
def _compute_pseudo_obs_cumul(data: np.ndarray) -> np.ndarray:
    """基于经验累积分布（Cumul）计算伪观测值。"""
    n, d = data.shape
    pseudo = np.zeros((n, d), dtype=float)
    for j in range(d):
        col = data[:, j]
        unique_vals, counts = np.unique(col, return_counts=True)
        m = len(unique_vals)
        cum_counts = np.cumsum(counts)
        p_vals = (cum_counts - counts + 0.5 * counts) / n
        # 修正严格递增
        for k in range(1, m):
            if p_vals[k] <= p_vals[k-1]:
                p_vals[k] = p_vals[k-1] + 1e-12
        x_full = np.concatenate(([unique_vals[0] - 1e-12], unique_vals, [unique_vals[-1] + 1e-12]))
        p_full = np.concatenate(([0.0], p_vals, [1.0]))
        for i, val in enumerate(col):
            u = cumul_cdf(val, x_full.tolist(), p_full.tolist())
            pseudo[i, j] = np.clip(u, 1e-12, 1.0 - 1e-12)
    return pseudo

# ==================== Copula 参数估计（原有，略作调整） ====================
def _kendall_tau_matrix(data: np.ndarray) -> np.ndarray:
    """计算多维数据的 Kendall's tau 矩阵（精确计算，O(n^2)）"""
    n, d = data.shape
    tau_mat = np.zeros((d, d))
    for i in range(d):
        for j in range(d):
            if i == j:
                tau_mat[i, j] = 1.0
            else:
                x = data[:, i]
                y = data[:, j]
                concordant = 0
                discordant = 0
                for k in range(n):
                    for l in range(k+1, n):
                        sign = (x[k] - x[l]) * (y[k] - y[l])
                        if sign > 0:
                            concordant += 1
                        elif sign < 0:
                            discordant += 1
                total = concordant + discordant
                if total > 0:
                    tau = (concordant - discordant) / total
                else:
                    tau = 0.0
                tau_mat[i, j] = tau
    return tau_mat

def _tau_to_correlation(tau: np.ndarray) -> np.ndarray:
    R = np.sin(np.pi / 2 * tau)
    np.fill_diagonal(R, 1.0)
    return _nearest_psd(R)

def _estimate_clayton_theta_from_tau(tau: float) -> float:
    if tau <= 0:
        return 0.1
    theta = 2 * tau / (1 - tau)
    return max(theta, 0.01)

def _estimate_gumbel_theta_from_tau(tau: float) -> float:
    if tau <= 0:
        return 1.0
    theta = 1 / (1 - tau)
    return max(theta, 1.0)

def _estimate_frank_theta_from_tau(tau: float) -> float:
    if tau <= 0:
        return 0.1
    if tau >= 0.999:
        return 20.0
    from scipy.optimize import brentq
    from scipy.integrate import quad
    def func(theta):
        if theta == 0:
            return 0
        def integrand(t):
            return t / (np.exp(t) - 1)
        integral, _ = quad(integrand, 0, theta)
        tau_theta = 1 - 4/theta + 4/theta**2 * integral
        return tau_theta - tau
    try:
        theta = brentq(func, 0.01, 20.0)
        return theta
    except:
        return 4 * tau / (1 - tau**2)

def _estimate_t_dof_from_correlation(data_pseudo: np.ndarray, R: np.ndarray) -> float:
    """使用 profile likelihood + 整数网格搜索估计 t copula 的自由度"""
    n, d = data_pseudo.shape
    R_psd = _nearest_psd(R)
    np.fill_diagonal(R_psd, 1.0)
    nus = np.concatenate((
        np.arange(1, 31, dtype=float),
        np.array([35, 40, 45, 50, 60, 70, 80, 90, 100, 120, 150, 200, 250, 300])
    ))
    best_nu = 4.0
    best_ll = -np.inf
    for nu_try in nus:
        try:
            ll = _t_copula_loglik(data_pseudo, R_psd, nu_try)
            if np.isfinite(ll) and ll > best_ll:
                best_ll = ll
                best_nu = nu_try
        except:
            continue
    return best_nu

def _gaussian_copula_loglik(u: np.ndarray, R: np.ndarray) -> float:
    n, d = u.shape
    z = norm.ppf(u)
    R = _nearest_psd(R)
    np.fill_diagonal(R, 1.0)
    try:
        R_inv = np.linalg.inv(R)
    except:
        R = R + np.eye(d) * 1e-8
        R_inv = np.linalg.inv(R)
    logdet = np.linalg.slogdet(R)[1]
    loglik = 0.0
    for i in range(n):
        zi = z[i, :]
        quad = zi @ (R_inv - np.eye(d)) @ zi
        loglik += -0.5 * quad
    loglik -= 0.5 * n * logdet
    return loglik

def _t_copula_loglik(u: np.ndarray, R: np.ndarray, nu: float) -> float:
    n, d = u.shape
    z = t.ppf(u, df=nu)
    R = _nearest_psd(R)
    np.fill_diagonal(R, 1.0)
    try:
        invR = np.linalg.inv(R)
    except:
        R = R + np.eye(d) * 1e-8
        invR = np.linalg.inv(R)
    q = np.array([z[i] @ invR @ z[i] for i in range(n)])
    logdet = np.linalg.slogdet(R)[1]
    const_mv = gammaln((nu + d) / 2) - gammaln(nu / 2) - 0.5 * d * np.log(np.pi * nu) - 0.5 * logdet
    log_f_mv = const_mv - (nu + d) / 2 * np.log(1 + q / nu)
    const_marg = gammaln((nu + 1) / 2) - gammaln(nu / 2) - 0.5 * np.log(np.pi * nu)
    log_f_marg = const_marg - (nu + 1) / 2 * np.log(1 + z**2 / nu)
    loglik = np.sum(log_f_mv - np.sum(log_f_marg, axis=1))
    return loglik

def _bivariate_archimedean_loglik(copula_type: str, theta: float, u: np.ndarray) -> float:
    n, d = u.shape
    if d != 2:
        return 0.0
    log_lik = 0.0
    for k in range(n):
        u1, u2 = u[k, 0], u[k, 1]
        u1 = np.clip(u1, 1e-12, 1.0 - 1e-12)
        u2 = np.clip(u2, 1e-12, 1.0 - 1e-12)
        if copula_type == "Clayton":
            if theta <= 0:
                continue
            t1 = u1**(-theta) - 1
            t2 = u2**(-theta) - 1
            log_c = np.log(1+theta) - (1+theta)*(np.log(u1)+np.log(u2)) - (2+1/theta)*np.log(max(t1+t2, 1e-12))
            log_lik += log_c
        elif copula_type == "ClaytonR":
            u1, u2 = 1 - u1, 1 - u2
            u1 = np.clip(u1, 1e-12, 1.0 - 1e-12)
            u2 = np.clip(u2, 1e-12, 1.0 - 1e-12)
            t1 = u1**(-theta) - 1
            t2 = u2**(-theta) - 1
            log_c = np.log(1+theta) - (1+theta)*(np.log(u1)+np.log(u2)) - (2+1/theta)*np.log(max(t1+t2, 1e-12))
            log_lik += log_c
        elif copula_type == "Gumbel":
            if theta < 1:
                continue
            t1 = -np.log(u1)
            t2 = -np.log(u2)
            t1 = max(t1, 1e-12)
            t2 = max(t2, 1e-12)
            log_t1 = np.log(t1)
            log_t2 = np.log(t2)
            log_exp1 = theta * log_t1
            log_exp2 = theta * log_t2
            if log_exp1 > log_exp2:
                logS = log_exp1 + np.log1p(np.exp(log_exp2 - log_exp1))
            else:
                logS = log_exp2 + np.log1p(np.exp(log_exp1 - log_exp2))
            term1 = -np.exp(logS / theta)
            term2 = (theta - 1) * (log_t1 + log_t2)
            term3 = (2.0/theta - 2.0) * logS
            exp_neg_logS_over_theta = np.exp(-logS / theta)
            term4 = np.log1p((theta - 1) * exp_neg_logS_over_theta)
            term5 = -np.log(u1) - np.log(u2)
            log_c = term1 + term2 + term3 + term4 + term5
            log_lik += log_c
        elif copula_type == "GumbelR":
            u1, u2 = 1 - u1, 1 - u2
            u1 = np.clip(u1, 1e-12, 1.0 - 1e-12)
            u2 = np.clip(u2, 1e-12, 1.0 - 1e-12)
            t1 = -np.log(u1)
            t2 = -np.log(u2)
            t1 = max(t1, 1e-12)
            t2 = max(t2, 1e-12)
            log_t1 = np.log(t1)
            log_t2 = np.log(t2)
            log_exp1 = theta * log_t1
            log_exp2 = theta * log_t2
            if log_exp1 > log_exp2:
                logS = log_exp1 + np.log1p(np.exp(log_exp2 - log_exp1))
            else:
                logS = log_exp2 + np.log1p(np.exp(log_exp1 - log_exp2))
            term1 = -np.exp(logS / theta)
            term2 = (theta - 1) * (log_t1 + log_t2)
            term3 = (2.0/theta - 2.0) * logS
            exp_neg_logS_over_theta = np.exp(-logS / theta)
            term4 = np.log1p((theta - 1) * exp_neg_logS_over_theta)
            term5 = -np.log(u1) - np.log(u2)
            log_c = term1 + term2 + term3 + term4 + term5
            log_lik += log_c
        elif copula_type == "Frank":
            theta = min(theta, 20.0)
            if theta <= 0:
                continue
            exp_theta = np.exp(-theta)
            exp_theta_u1 = np.exp(-theta * u1)
            exp_theta_u2 = np.exp(-theta * u2)
            numerator = theta * (1 - exp_theta) * exp_theta_u1 * exp_theta_u2
            denominator = 1 - exp_theta - (1 - exp_theta_u1)*(1 - exp_theta_u2)
            denominator = max(abs(denominator), 1e-12)
            denominator = denominator if denominator > 0 else 1e-12
            if numerator <= 0:
                log_c = -1e10
            else:
                log_c = np.log(numerator) - 2 * np.log(denominator)
            log_lik += log_c
    return log_lik

def _pairwise_loglik_archimedean(copula_type: str, theta: float, u: np.ndarray) -> float:
    n, d = u.shape
    total_loglik = 0.0
    for i in range(d):
        for j in range(i+1, d):
            u_pair = u[:, [i, j]]
            loglik_pair = _bivariate_archimedean_loglik(copula_type, theta, u_pair)
            total_loglik += loglik_pair
    return total_loglik

def _clayton_multivariate_density(u: np.ndarray, theta: float) -> float:
    n, d = u.shape
    log_c = 0.0
    for i in range(n):
        u_i = u[i, :]
        if np.any(u_i <= 0) or np.any(u_i >= 1):
            return -1e10
        sum_term = np.sum(u_i**(-theta) - 1) + 1
        log_c += np.log(theta) * d + gammaln(1/theta + d) - gammaln(1/theta) \
                 - (theta+1) * np.sum(np.log(u_i)) - (1/theta + d) * np.log(sum_term)
    return log_c

def _claytonr_multivariate_density(u: np.ndarray, theta: float) -> float:
    u_rot = 1 - u
    return _clayton_multivariate_density(u_rot, theta)

def _fit_copula_mle_exact(copula_type: str, data_pseudo: np.ndarray) -> dict:
    """精确极大似然法估计 copula 参数"""
    n, d = data_pseudo.shape
    if copula_type in ROTATION_MAP:
        rotated_data = _apply_rotation(data_pseudo, copula_type)
        base_type = _base_type_from_rotated(copula_type)
        return _fit_copula_mle_approx(base_type, rotated_data)
    if copula_type == "Gaussian":
        z = norm.ppf(data_pseudo)
        R = np.corrcoef(z, rowvar=False)
        R = _nearest_psd(R)
        np.fill_diagonal(R, 1.0)
        return {'matrix': R}
    elif copula_type == "t":
        z = norm.ppf(data_pseudo)
        R0 = np.corrcoef(z, rowvar=False)
        R0 = _nearest_psd(R0)
        np.fill_diagonal(R0, 1.0)
        nu = _estimate_t_dof_from_correlation(data_pseudo, R0)
        return {'matrix': R0, 'param': nu}
    elif copula_type in ("Clayton", "ClaytonR"):
        def neg_log_lik(theta):
            if theta <= 0:
                return 1e10
            if copula_type == "Clayton":
                ll = _clayton_multivariate_density(data_pseudo, theta)
            else:
                ll = _claytonr_multivariate_density(data_pseudo, theta)
            if np.isnan(ll) or np.isinf(ll):
                return 1e10
            return -ll
        tau_mat = _kendall_tau_matrix(data_pseudo)
        tau_avg = np.mean(tau_mat[np.triu_indices_from(tau_mat, k=1)])
        initial = _estimate_clayton_theta_from_tau(tau_avg)
        res = minimize(neg_log_lik, initial, bounds=[(0.01, 50)], method='L-BFGS-B')
        theta = res.x[0]
        return {'param': theta}
    else:
        def neg_pairwise_loglik(theta):
            if copula_type in ("Gumbel", "GumbelR"):
                if theta < 1:
                    return 1e10
                if theta > 30:
                    return 1e10
            if copula_type == "Frank":
                if theta <= 0 or theta > 20:
                    return 1e10
            ll = _pairwise_loglik_archimedean(copula_type, theta, data_pseudo)
            if np.isnan(ll) or np.isinf(ll):
                return 1e10
            return -ll
        tau_mat = _kendall_tau_matrix(data_pseudo)
        tau_avg = np.mean(tau_mat[np.triu_indices_from(tau_mat, k=1)])
        if copula_type in ("Gumbel", "GumbelR"):
            initial = _estimate_gumbel_theta_from_tau(tau_avg)
            bounds = [(1.0, 30)]
        else:  # Frank
            initial = _estimate_frank_theta_from_tau(tau_avg)
            bounds = [(0.01, 20)]
        res = minimize(neg_pairwise_loglik, initial, bounds=bounds, method='L-BFGS-B')
        theta = res.x[0]
        return {'param': theta}

def _fit_copula_mle_approx(copula_type: str, data_pseudo: np.ndarray) -> dict:
    """近似极大似然法（pairwise 似然）"""
    n, d = data_pseudo.shape
    if copula_type in ROTATION_MAP:
        rotated_data = _apply_rotation(data_pseudo, copula_type)
        base_type = _base_type_from_rotated(copula_type)
        return _fit_copula_mle_approx(base_type, rotated_data)
    if copula_type == "Gaussian":
        tau = _kendall_tau_matrix(data_pseudo)
        R = _tau_to_correlation(tau)
        return {'matrix': R}
    elif copula_type == "t":
        tau = _kendall_tau_matrix(data_pseudo)
        R = _tau_to_correlation(tau)
        nu = _estimate_t_dof_from_correlation(data_pseudo, R)
        return {'matrix': R, 'param': nu}
    else:
        tau_mat = _kendall_tau_matrix(data_pseudo)
        tau_avg = np.mean(tau_mat[np.triu_indices_from(tau_mat, k=1)])
        if copula_type in ("Clayton", "ClaytonR"):
            initial = _estimate_clayton_theta_from_tau(tau_avg)
            bounds = [(0.01, 50)]
        elif copula_type in ("Gumbel", "GumbelR"):
            initial = _estimate_gumbel_theta_from_tau(tau_avg)
            bounds = [(1.0, 30)]
        else:
            initial = _estimate_frank_theta_from_tau(tau_avg)
            bounds = [(0.01, 20)]
        def neg_pairwise_loglik(theta):
            if copula_type in ("Gumbel", "GumbelR") and theta < 1:
                return 1e10
            if copula_type == "Frank" and theta <= 0:
                return 1e10
            ll = _pairwise_loglik_archimedean(copula_type, theta, data_pseudo)
            if np.isnan(ll) or np.isinf(ll):
                return 1e10
            return -ll
        res = minimize(neg_pairwise_loglik, initial, bounds=bounds, method='L-BFGS-B')
        theta = res.x[0]
        return {'param': theta}

def _fit_copula_kendall(copula_type: str, data_pseudo: np.ndarray) -> dict:
    """使用 Kendall's tau 估计参数"""
    n, d = data_pseudo.shape
    if copula_type in ROTATION_MAP:
        rotated_data = _apply_rotation(data_pseudo, copula_type)
        base_type = _base_type_from_rotated(copula_type)
        return _fit_copula_kendall(base_type, rotated_data)
    if copula_type == "Gaussian":
        tau = _kendall_tau_matrix(data_pseudo)
        R = _tau_to_correlation(tau)
        return {'matrix': R}
    elif copula_type == "t":
        tau = _kendall_tau_matrix(data_pseudo)
        R = _tau_to_correlation(tau)
        nu = _estimate_t_dof_from_correlation(data_pseudo, R)
        return {'matrix': R, 'param': nu}
    elif copula_type in ("Clayton", "ClaytonR"):
        tau = _kendall_tau_matrix(data_pseudo)
        tau_avg = np.mean(tau[np.triu_indices_from(tau, k=1)])
        theta = _estimate_clayton_theta_from_tau(tau_avg)
        return {'param': theta}
    elif copula_type in ("Gumbel", "GumbelR"):
        tau = _kendall_tau_matrix(data_pseudo)
        tau_avg = np.mean(tau[np.triu_indices_from(tau, k=1)])
        theta = _estimate_gumbel_theta_from_tau(tau_avg)
        return {'param': theta}
    elif copula_type == "Frank":
        tau = _kendall_tau_matrix(data_pseudo)
        tau_avg = np.mean(tau[np.triu_indices_from(tau, k=1)])
        theta = _estimate_frank_theta_from_tau(tau_avg)
        return {'param': theta}
    else:
        raise ValueError(f"不支持的 Copula 类型: {copula_type}")

def _compute_information_criteria(copula_type: str, params: dict, data_pseudo: np.ndarray, method: int = 1) -> Tuple[float, float, float]:
    n, d = data_pseudo.shape
    if copula_type in ROTATION_MAP:
        rotated_data = _apply_rotation(data_pseudo, copula_type)
        base_type = _base_type_from_rotated(copula_type)
        return _compute_information_criteria(base_type, params, rotated_data, method)
    if copula_type == "Gaussian":
        R = params['matrix']
        log_lik = _gaussian_copula_loglik(data_pseudo, R)
        k = d * (d - 1) // 2
    elif copula_type == "t":
        R = params['matrix']
        nu = params['param']
        log_lik = _t_copula_loglik(data_pseudo, R, nu)
        k = d * (d - 1) // 2 + 1
    elif copula_type in ("Clayton", "ClaytonR"):
        theta = params['param']
        if copula_type == "Clayton":
            log_lik = _clayton_multivariate_density(data_pseudo, theta)
        else:
            log_lik = _claytonr_multivariate_density(data_pseudo, theta)
        k = 1
    else:
        theta = params['param']
        if copula_type == "Frank":
            theta = min(theta, 20.0)
        log_lik = _pairwise_loglik_archimedean(copula_type, theta, data_pseudo)
        k = 1
    if np.isnan(log_lik) or np.isinf(log_lik):
        log_lik = -1e10
    aic = -2 * log_lik + 2 * k
    if n - k - 1 > 0:
        aicc = aic + (2 * k * (k + 1)) / (n - k - 1)
    else:
        aicc = aic
    bic = -2 * log_lik + k * np.log(n)
    avg_loglik = log_lik / n
    return aicc, bic, avg_loglik

def _create_copula_matrix(sheet, top_left_cell, copula_type: str, params: dict, n_vars: int, matrix_name: str):
    start_row = top_left_cell.Row
    start_col = top_left_cell.Column
    if copula_type in ("Gaussian", "t"):
        rows = n_vars + 3
        cols = n_vars + 1
        end_row = start_row + rows - 1
        end_col = start_col + cols - 1
        range_matrix = sheet.Range(sheet.Cells(start_row, start_col), sheet.Cells(end_row, end_col))
        labels = ["缺失值"] * n_vars
        data = [[None] * cols for _ in range(rows)]
        data[0][0] = matrix_name
        data[1][0] = "模式"
        data[1][1] = copula_type
        if copula_type == "t" and 'param' in params:
            data[1][2] = params['param']
        data[2][0] = "矩阵系数"
        for j, label in enumerate(labels):
            if 1 + j < cols:
                data[2][1 + j] = label
        for i, label in enumerate(labels):
            if 3 + i < rows:
                data[3 + i][0] = label
        if 'matrix' in params:
            mat = params['matrix']
        else:
            mat = np.eye(n_vars)
        for i in range(n_vars):
            for j in range(n_vars):
                if i >= j:
                    if 3 + i < rows and 1 + j < cols:
                        data[3 + i][1 + j] = mat[i, j]
        range_matrix.Value2 = data
    else:
        rows = 3
        cols = 3
        end_row = start_row + rows - 1
        end_col = start_col + cols - 1
        range_matrix = sheet.Range(sheet.Cells(start_row, start_col), sheet.Cells(end_row, end_col))
        data = [[None] * cols for _ in range(rows)]
        data[0][0] = matrix_name
        data[1][0] = "模式"
        data[1][1] = copula_type
        if 'param' in params:
            data[1][2] = params['param']
        data[2][0] = "维度"
        data[2][1] = n_vars
        range_matrix.Value2 = data
    _set_named_range(_get_excel_app(), matrix_name, range_matrix)

def _update_copula_labels(copula_name: str, new_labels: List[str]):
    app = _get_excel_app()
    range_obj = _get_named_range(app, copula_name)
    if range_obj is None:
        raise ValueError(f"找不到 Copula 矩阵 {copula_name}")
    vals = range_obj.Value2
    if vals is None:
        return
    rows = len(vals)
    cols = len(vals[0]) if vals and isinstance(vals[0], (list, tuple)) else 1
    n = rows - 3
    if rows >= 4 and cols >= 2 and n == len(new_labels):
        if isinstance(vals[2], (list, tuple)):
            for j, label in enumerate(new_labels):
                if 1 + j < cols:
                    vals[2][1 + j] = label
        for i, label in enumerate(new_labels):
            if 3 + i < rows:
                if isinstance(vals[3 + i], (list, tuple)):
                    vals[3 + i][0] = label
        range_obj.Value2 = vals

# ==================== 宏实现 ====================
@xl_macro
def DriskFitCopula():
    """从历史数据拟合最优 Copula，并生成 Copula 矩阵（与现有系统完全兼容）"""
    app = _get_excel_app()
    original_status = app.StatusBar
    try:
        app.StatusBar = "正在运行 copula 拟合……"

        # 1. 选择历史数据区域
        try:
            data_range_obj = app.InputBox(
                Prompt="请选择历史数据区域（纯数字，每列代表一个变量）",
                Title="DriskFitCopula - 选择数据区域",
                Type=8
            )
            if data_range_obj is False:
                return
        except:
            xlcAlert("选择区域操作取消或失败")
            return

        try:
            data_range_addr = data_range_obj.Address
            sheet_name = data_range_obj.Worksheet.Name
            if ' ' in sheet_name or any(c in sheet_name for c in "[]:?*"):
                safe_sheet = f"'{sheet_name}'"
            else:
                safe_sheet = sheet_name
            data_range = f"{safe_sheet}!{data_range_addr}"
            hist_data = _read_data_range(app, data_range)
        except Exception as e:
            xlcAlert(f"数据区域无效: {str(e)}")
            return

        n_vars = hist_data.shape[1]
        n_obs = hist_data.shape[0]
        if n_vars < 2:
            xlcAlert("至少需要两列数据才能拟合 Copula。")
            return

        has_ties = _check_ties(hist_data, app)

        # 2. 选择拟合方法
        method_choice = app.InputBox(
            "请选择拟合方法：\n1=精确极大似然法（MLE）\n2=近似极大似然法（基于pairwise）\n3=肯德尔等级相关（默认，稳健但保守）",
            "拟合方法", Type=1, Default="1"
        )
        if method_choice is False:
            return
        method = int(method_choice)
        if method not in (1, 2, 3):
            xlcAlert("无效的方法编号，请输入 1、2 或 3")
            return

        if method == 1 and DISTFIT_AVAILABLE:
            use_parametric_margins = True

        # 3. 生成伪观测值（均匀分位数）
        if use_parametric_margins and DISTFIT_AVAILABLE:
            # 参数化边缘拟合
            app.StatusBar = "正在进行参数化边缘分布拟合（每列自动选择最优分布）..."
            pseudo_obs, marginal_info = _transform_to_uniform_parametric(hist_data, metric='AICc')
            # 可选：显示边缘分布汇总
            margin_summary = "\n".join([f"列{i+1}: {info[0]} 参数={info[1]}" for i, info in enumerate(marginal_info) if info is not None])
            if margin_summary:
                print("边缘分布拟合结果:\n" + margin_summary)
            app.StatusBar = "边缘分布拟合完成，继续 Copula 参数估计..."
        else:
            # 传统经验分布方法（cumul）
            # 询问数据是否已经是均匀分布
            data_already_uniform = False
            try:
                import tkinter as tk
                from tkinter import ttk
                root = tk.Tk()
                root.title("数据预处理")
                root.geometry("350x150")
                root.attributes('-topmost', True)
                root.update_idletasks()
                x = (root.winfo_screenwidth() // 2) - (350 // 2)
                y = (root.winfo_screenheight() // 2) - (150 // 2)
                root.geometry(f"+{x}+{y}")
                label = ttk.Label(root, text="数据是否已经是均匀分布变量（已在[0,1]区间内）？", padding=10)
                label.pack()
                result = tk.BooleanVar()
                result.set(False)
                def on_yes():
                    result.set(True)
                    root.destroy()
                def on_no():
                    result.set(False)
                    root.destroy()
                btn_frame = ttk.Frame(root)
                btn_frame.pack(pady=10)
                ttk.Button(btn_frame, text="是", width=10, command=on_yes).pack(side=tk.LEFT, padx=10)
                ttk.Button(btn_frame, text="否", width=10, command=on_no).pack(side=tk.RIGHT, padx=10)
                root.mainloop()
                data_already_uniform = result.get()
            except Exception:
                choice = app.InputBox("数据是否已经是均匀分布变量（已在[0,1]区间内）？\n输入 1 表示是，0 表示否：", "数据预处理", Type=1, Default="0")
                if choice is False:
                    return
                data_already_uniform = (int(choice) == 1)

            if data_already_uniform:
                pseudo_obs = hist_data.copy()
                if np.any(pseudo_obs <= 0) or np.any(pseudo_obs >= 1):
                    xlcAlert("警告：您声称数据已经是均匀分布，但存在超出 (0,1) 范围的值。拟合可能不准确。")
                pseudo_obs = np.clip(pseudo_obs, 1e-12, 1.0 - 1e-12)
            else:
                pseudo_obs = _compute_pseudo_obs_cumul(hist_data)

        # 4. 选择检验指标（AICc/BIC/平均对数似然）
        criterion_choice = app.InputBox(
            "请选择排名指标：\n1=AICc（小样本修正赤池信息准则）\n2=BIC\n3=对数似然均值（Average log likelihood）",
            "检验指标", Type=1, Default="1"
        )
        if criterion_choice is False:
            return
        criterion = int(criterion_choice)
        if criterion not in (1, 2, 3):
            xlcAlert("无效的指标编号，请输入 1,2 或 3")
            return

        # 5. 对十二种 Copula 进行拟合
        copula_types = ["Gaussian", "t", "Clayton", "ClaytonR", "Gumbel", "GumbelR", "Frank",
                        "FrankRX", "ClaytonRX", "ClaytonRY", "GumbelRX", "GumbelRY"]
        results = []
        for ctype in copula_types:
            try:
                if method == 1:
                    params = _fit_copula_mle_exact(ctype, pseudo_obs)
                elif method == 2:
                    params = _fit_copula_mle_approx(ctype, pseudo_obs)
                else:
                    params = _fit_copula_kendall(ctype, pseudo_obs)
                aicc, bic, avg_loglik = _compute_information_criteria(ctype, params, pseudo_obs, method)
                results.append((ctype, params, aicc, bic, avg_loglik))
            except Exception as e:
                print(f"拟合 {ctype} 失败: {e}")
                continue

        if not results:
            xlcAlert("所有 Copula 拟合均失败，请检查数据。")
            return

        # 排序
        if criterion == 1:
            results.sort(key=lambda x: x[2])  # AICc
            criterion_name = "AICc"
            get_val = lambda r: r[2]
        elif criterion == 2:
            results.sort(key=lambda x: x[3])  # BIC
            criterion_name = "BIC"
            get_val = lambda r: r[3]
        else:
            results.sort(key=lambda x: -x[4])  # avg_loglik
            criterion_name = "Average log likelihood"
            get_val = lambda r: r[4]

        headers = ["Copula 类型", f"{criterion_name} 值"]
        rows = []
        for res in results:
            ctype = res[0]
            val = get_val(res)
            rows.append([ctype, f"{val:.6f}"])

        selected_idx = _show_copula_selection_dialog(
            f"选择 Copula - 按 {criterion_name} 排序（越小越好）" if criterion in (1,2) else f"选择 Copula - 按 {criterion_name} 排序（越大越好）",
            headers, rows
        )
        if selected_idx < 0:
            xlcAlert("未选择 Copula，操作取消。")
            return

        selected_ctype, selected_params, _, _, _ = results[selected_idx]

        # 6. 指定矩阵位置和名称
        try:
            matrix_top_left = app.InputBox(
                Prompt="请选择 Copula 矩阵的左上角单元格（矩阵将向右下扩展）",
                Title="DriskFitCopula - 矩阵位置",
                Type=8
            )
            if matrix_top_left is False:
                return
        except:
            xlcAlert("选择矩阵位置取消或失败")
            return

        wb = _get_workbook()
        existing_names = [name.Name for name in wb.Names]
        max_num = 0
        for name in existing_names:
            match = re.match(r'^Copula_(\d+)$', name)
            if match:
                max_num = max(max_num, int(match.group(1)))
        new_num = max_num + 1
        default_name = f"Copula_{new_num}"
        user_name = app.InputBox(f"Copula 矩阵名称（默认 {default_name}）:", "矩阵名称", Default=default_name)
        if user_name and isinstance(user_name, str):
            matrix_name = user_name.strip()
        else:
            matrix_name = default_name

        sheet = matrix_top_left.Worksheet
        _create_copula_matrix(sheet, matrix_top_left, selected_ctype, selected_params, n_vars, matrix_name)

        success_msg = (f"已成功创建 Copula 矩阵“{matrix_name}”，类型为 {selected_ctype}，参数：{selected_params}\n"
                       f"您可以使用 DriskAppendCopula 将此 Copula 附加到 Drisk 分布函数上。")
        if use_parametric_margins:
            success_msg += "\n\n本次拟合使用了参数化边缘分布（每列最优分布），建模精度更高。"
        if has_ties:
            success_msg += "\n\n注意：原始数据中存在并列值（ties），这可能影响 Copula 拟合的精度。建议检查数据或考虑使用 Kendall tau 估计方法。"
        xlcAlert(success_msg)

    except Exception as e:
        xlcAlert(f"执行 DriskFitCopula 时出错：{str(e)}")
        traceback.print_exc()
    finally:
        app.StatusBar = original_status

@xl_macro
def DriskAppendCopula():
    """将一个 Copula 矩阵附加到所选的 Drisk 分布函数上"""
    app = _get_excel_app()
    try:
        wb = _get_workbook()
        copula_names = []
        for name_obj in wb.Names:
            name = name_obj.Name
            if name.startswith('Copula_') or name.startswith('CopulaFit_') or name.startswith('copula'):
                copula_names.append(name)
        if not copula_names:
            xlcAlert("未找到任何 Copula 矩阵。请先使用 DriskMakeCopula 或 DriskFitCopula 创建 Copula。")
            return

        if len(copula_names) == 1:
            copula_name = copula_names[0]
        else:
            prompt_lines = ["请选择要附加的 Copula 矩阵（输入数字）:"]
            for i, name in enumerate(copula_names, 1):
                prompt_lines.append(f"{i}. {name}")
            prompt = "\n".join(prompt_lines)
            try:
                choice = app.InputBox(prompt, "DriskAppendCopula - 选择 Copula", Type=1, Default="1")
                if choice is False:
                    return
                idx = int(choice) - 1
                if 0 <= idx < len(copula_names):
                    copula_name = copula_names[idx]
                else:
                    xlcAlert("输入的数字无效。")
                    return
            except:
                xlcAlert("输入无效。")
                return

        range_obj = _get_named_range(app, copula_name)
        if range_obj is None:
            xlcAlert(f"无法找到命名区域 {copula_name}。")
            return
        vals = range_obj.Value2
        if vals is None or not isinstance(vals, (list, tuple)):
            xlcAlert("Copula 数据为空或格式不正确。")
            return
        try:
            copula_type = str(vals[1][1]).strip() if len(vals) > 1 and len(vals[1]) > 1 else ""
            if copula_type in ("Gaussian", "t"):
                n_vars = len(vals) - 3
                if n_vars < 2:
                    xlcAlert("无法确定 Copula 的变量数。")
                    return
            else:
                if len(vals) > 2 and len(vals[2]) > 1:
                    n_vars = int(vals[2][1])
                else:
                    xlcAlert("无法确定 Copula 的变量数。")
                    return
        except Exception as e:
            xlcAlert(f"解析 Copula 信息失败: {e}")
            return

        try:
            range_obj = app.InputBox(
                Prompt=f"请选择 {n_vars} 个包含 Drisk 分布函数的单元格（每个单元格应为有效的 Drisk 分布）",
                Title="DriskAppendCopula - 选择分布区域",
                Type=8
            )
            if range_obj is False:
                return
        except:
            xlcAlert("选择区域操作取消或失败")
            return

        cells_list = []
        for cell in range_obj:
            cells_list.append(cell)
        if len(cells_list) != n_vars:
            xlcAlert(f"选择的单元格数量 ({len(cells_list)}) 与 Copula 的变量数 ({n_vars}) 不一致。")
            return

        invalid_cells = []
        existing_copula_cells = []
        for cell in cells_list:
            try:
                formula = cell.Formula
                if not isinstance(formula, str) or not formula.startswith('='):
                    invalid_cells.append(cell.Address)
                elif not is_distribution_function(formula):
                    invalid_cells.append(cell.Address)
                elif re.search(r'DriskCopula\s*\(', formula, re.IGNORECASE):
                    existing_copula_cells.append(cell.Address)
            except:
                invalid_cells.append(cell.Address)
        if invalid_cells:
            xlcAlert(f"以下单元格不是有效的 Drisk 分布函数：{', '.join(invalid_cells)}\n请修正后重试。")
            return
        if existing_copula_cells:
            choice = _show_choice_dialog(
                "覆盖提示",
                f"以下单元格已经包含 DriskCopula 属性：{', '.join(existing_copula_cells)}\n是否覆盖？",
                "覆盖", "取消"
            )
            if choice != 1:
                return

        if copula_type in ("Gaussian", "t"):
            new_labels = []
            for cell in cells_list:
                try:
                    addr = cell.Address.replace('$', '')
                    new_labels.append(addr)
                except:
                    new_labels.append("缺失值")
            try:
                _update_copula_labels(copula_name, new_labels)
            except Exception as e:
                print(f"更新标签失败: {e}")

        pos = 1
        modified = []
        for cell in cells_list:
            formula = cell.Formula
            new_formula = _ensure_copula_in_formula(formula, copula_name, pos)
            if new_formula != formula:
                cell.Formula = new_formula
                modified.append(cell.Address)
            pos += 1

        if modified:
            xlcAlert(f"已为 {len(modified)} 个单元格添加/更新 DriskCopula 属性。\nCopula 名称：{copula_name}")
        else:
            xlcAlert("所有单元格已包含 DriskCopula，未作修改。")

    except Exception as e:
        xlcAlert(f"执行 DriskAppendCopula 时出错：{str(e)}")
        traceback.print_exc()