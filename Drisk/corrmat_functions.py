# corrmat_functions.py (完整版 - 修复版)
"""
相关性矩阵与秩相关模块
提供 DriskMakeCorr, DriskCheckCorr, DriskWeight 宏以及 DriskCorr, DriskCorrRank 函数。
实现基于 Iman‑Conover 方法的多变量秩相关控制，支持目标相关矩阵的正半定调整与权重。
"""

import re
import traceback
import numpy as np
import win32com.client
import win32api
import win32con
from pyxll import xl_arg, xl_macro, xl_func, xl_app, xlcAlert
from typing import List, Tuple, Optional, Dict, Any
from constants import DISTRIBUTION_FUNCTION_NAMES
from formula_parser import (
    parse_complete_formula,
    parse_args_with_nested_functions,
    extract_all_attributes_from_formula,
    is_distribution_function
)
from simulation_manager import get_simulation
from distribution_functions import ERROR_MARKER

# ---------- 映射表管理 ----------
MAPPING_NAME = "__DriskCorrMatMapping__"  # 隐藏命名区域名称

def _get_mapping_table(app):
    """
    读取映射表，返回字典 {weight_name: corr_name}
    映射表存储在工作簿的命名区域中，是一个两列的表格（第一列权重矩阵名称，第二列相关性矩阵名称）
    """
    try:
        range_obj = _get_named_range(app, MAPPING_NAME)
        if range_obj is None:
            return {}
        vals = range_obj.Value2
        if vals is None or not isinstance(vals, (list, tuple)):
            return {}
        # 转换为列表
        if not isinstance(vals[0], (list, tuple)):
            vals = [vals]
        mapping = {}
        for row in vals:
            if len(row) >= 2:
                w_name = str(row[0]).strip()
                c_name = str(row[1]).strip()
                if w_name and c_name:
                    mapping[w_name] = c_name
        return mapping
    except Exception as e:
        print(f"读取映射表失败: {e}")
        return {}

def _save_mapping_table(app, mapping: Dict[str, str]):
    """将映射表保存到工作簿的命名区域中"""
    if not mapping:
        return
    try:
        # 准备数据
        data = []
        for w_name, c_name in mapping.items():
            data.append([w_name, c_name])
        if not data:
            return
        # 查找或创建用于存储映射表的工作表
        wb = app.ActiveWorkbook
        sheet_name = "DriskHiddenMapping"
        try:
            sheet = wb.Worksheets(sheet_name)
        except:
            sheet = wb.Worksheets.Add()
            sheet.Name = sheet_name
            sheet.Visible = 0  # 隐藏
        # 清空旧内容
        used = sheet.UsedRange
        if used:
            used.ClearContents()
        # 写入数据
        start_cell = sheet.Cells(1, 1)
        end_cell = sheet.Cells(len(data), 2)
        range_obj = sheet.Range(start_cell, end_cell)
        range_obj.Value2 = data
        # 设置命名区域
        _set_named_range(app, MAPPING_NAME, range_obj)
        print(f"映射表已保存: {mapping}")
    except Exception as e:
        print(f"保存映射表失败: {e}")

def _update_corr_weight_mapping(app, weight_name: str, corr_name: str):
    """更新或添加映射关系"""
    mapping = _get_mapping_table(app)
    mapping[weight_name] = corr_name
    _save_mapping_table(app, mapping)

def _remove_mapping_entry(app, weight_name: str):
    """删除映射条目（当权重矩阵被删除时可选）"""
    mapping = _get_mapping_table(app)
    if weight_name in mapping:
        del mapping[weight_name]
        _save_mapping_table(app, mapping)

# ---------- 静音矩阵管理 ----------
_MUTED_NAMES = "__DriskMutedCorrMatrices__"

def _get_muted_corr_matrices(app):
    """返回当前工作簿中静音的相关性矩阵名称列表（去重）"""
    try:
        range_obj = _get_named_range(app, _MUTED_NAMES)
        if range_obj is None:
            return []
        vals = range_obj.Value2
        if vals is None:
            return []
        # 存储为一个逗号分隔的字符串（单行单列）
        if isinstance(vals, str):
            text = vals.strip()
        elif isinstance(vals, (list, tuple)):
            # 如果区域有多个单元格，合并
            text = ",".join(str(v).strip() for row in vals for v in (row if isinstance(row, (list, tuple)) else [row]) if v is not None)
        else:
            text = str(vals).strip()
        if not text:
            return []
        names = [n.strip() for n in text.split(',') if n.strip()]
        # 去重
        return list(dict.fromkeys(names))
    except Exception as e:
        print(f"获取静音矩阵列表失败: {e}")
        return []

def _save_muted_corr_matrices(app, muted_list):
    """保存静音矩阵列表（去重后）到命名区域"""
    if not muted_list:
        # 如果列表为空，删除命名区域（如果存在）
        try:
            app.ActiveWorkbook.Names(_MUTED_NAMES).Delete()
        except:
            pass
        return
    # 去重
    muted_list = list(dict.fromkeys(muted_list))
    text = ",".join(muted_list)
    # 存储到命名区域（单行单列）
    try:
        wb = app.ActiveWorkbook
        # 查找或创建用于存储映射表的工作表
        sheet_name = "DriskHiddenMapping"
        try:
            sheet = wb.Worksheets(sheet_name)
        except:
            sheet = wb.Worksheets.Add()
            sheet.Name = sheet_name
            sheet.Visible = 0  # 隐藏
        # 写入数据
        cell = sheet.Cells(1, 3)  # 使用第三列，避免与映射表冲突
        cell.Value = text
        # 设置命名区域
        _set_named_range(app, _MUTED_NAMES, cell)
    except Exception as e:
        print(f"保存静音矩阵列表失败: {e}")

def _add_muted_corr_matrix(app, matrix_name):
    """将矩阵添加到静音列表"""
    muted = _get_muted_corr_matrices(app)
    if matrix_name not in muted:
        muted.append(matrix_name)
        _save_muted_corr_matrices(app, muted)

def _remove_muted_corr_matrix(app, matrix_name):
    """从静音列表中移除矩阵"""
    muted = _get_muted_corr_matrices(app)
    if matrix_name in muted:
        muted.remove(matrix_name)
        _save_muted_corr_matrices(app, muted)

def _is_corr_matrix_muted(app, matrix_name) -> bool:
    """检查矩阵是否在静音列表中"""
    muted = _get_muted_corr_matrices(app)
    return matrix_name in muted

# ---------- 辅助函数 ----------
def _get_excel_app():
    """获取当前 Excel 应用对象（安全方式）"""
    try:
        from com_fixer import _safe_excel_app
        return _safe_excel_app()
    except ImportError:
        return xl_app()

def _get_named_range(app, name: str):
    """返回指定名称对应的区域对象（如果存在），否则返回 None"""
    try:
        return app.ActiveWorkbook.Names(name).RefersToRange
    except:
        return None

def _set_named_range(app, name: str, range_obj):
    """为区域设置名称"""
    try:
        # 如果名称已存在，先删除再创建
        try:
            app.ActiveWorkbook.Names(name).Delete()
        except:
            pass
        app.ActiveWorkbook.Names.Add(Name=name, RefersTo=range_obj)
    except Exception as e:
        raise ValueError(f"无法创建命名区域 {name}: {e}")

def _read_matrix_from_range(range_obj):
    """
    从 Excel 区域读取数值矩阵，兼容下三角格式，返回完整对称矩阵。
    如果读取失败，返回 None 并打印错误。
    改进版：优先提取数据部分（跳过标签行/列），避免字符串转换错误。
    """
    try:
        vals = range_obj.Value2
        if vals is None:
            print("矩阵区域无数据")
            return None

        # 转换为二维列表
        if not isinstance(vals, (list, tuple)):
            vals = [[vals]]
        rows = len(vals)
        cols = len(vals[0])

        # 如果区域太小，无法构成矩阵
        if rows < 2 or cols < 2:
            print("矩阵区域小于 2x2，无法读取")
            return None

        # Copula 自定义格式：第三行第一列为“矩阵系数”
        if rows >= 3 and cols >= 2:
            head = None
            try:
                head = vals[2][0]
            except Exception:
                head = None
            if isinstance(head, str) and "矩阵系数" in head:
                data_rows = rows - 3
                data_cols = cols - 1
                if data_rows < 1 or data_cols < 1:
                    return None
                n = min(data_rows, data_cols)
                lower_tri = np.zeros((n, n), dtype=float)
                for i in range(n):
                    for j in range(min(i + 1, data_cols)):
                        try:
                            val = vals[i + 3][j + 1]
                            lower_tri[i, j] = float(val) if val is not None else 0.0
                        except Exception:
                            lower_tri[i, j] = 0.0
                full = np.zeros((n, n), dtype=float)
                for i in range(n):
                    for j in range(n):
                        if i >= j:
                            full[i, j] = lower_tri[i, j]
                        else:
                            full[i, j] = lower_tri[j, i]
                return full

        # 方法1：假设区域是带标签的下三角矩阵（第一行和第一列为标签）
        # 提取数据部分（从第二行第二列开始）
        data_rows = rows - 1
        data_cols = cols - 1
        data = [[None] * data_cols for _ in range(data_rows)]
        valid_data = True
        for i in range(data_rows):
            for j in range(data_cols):
                try:
                    val = vals[i+1][j+1]
                    if val is None:
                        val = 0.0
                    data[i][j] = float(val)
                except (ValueError, TypeError) as e:
                    print(f"数据部分单元格 ({i+2},{j+2}) 无法转换为浮点数: {e}")
                    valid_data = False
                    break
            if not valid_data:
                break
        if valid_data:
            # 成功提取数据部分，构建对称矩阵
            lower_tri = np.array(data, dtype=float)
            # 确保是方阵
            if lower_tri.shape[0] != lower_tri.shape[1]:
                # 如果数据部分不是方阵，尝试判断是否为下三角形式
                # 实际上下三角矩阵在数据区域中只包含下三角部分（包括对角线）
                # 但我们的写入函数 _write_lower_tri_matrix 是写满整个 n x n 区域的（只填充下三角，上三角留空或0）
                # 因此这里假定数据部分已经是 n x n
                print(f"数据部分不是方阵 ({lower_tri.shape})，可能矩阵格式不正确")
                # 尝试强制转换为方阵（取前 min(rows,cols) 行和列）
                min_dim = min(lower_tri.shape)
                lower_tri = lower_tri[:min_dim, :min_dim]
            # 构建对称矩阵
            n = lower_tri.shape[0]
            full = np.zeros((n, n))
            for i in range(n):
                for j in range(n):
                    if i >= j:
                        full[i, j] = lower_tri[i, j]
                    else:
                        full[i, j] = lower_tri[j, i]
            return full

        # 方法2：尝试将整个区域作为纯数值方阵处理
        try:
            mat = np.array(vals, dtype=float)
            if mat.shape[0] == mat.shape[1]:
                return mat
            else:
                print(f"纯数值矩阵大小 {mat.shape} 不是方阵")
                return None
        except Exception as e:
            print(f"纯数值矩阵读取失败: {e}")
            return None

    except Exception as e:
        print(f"读取矩阵异常: {e}")
        traceback.print_exc()
        return None

def _write_lower_tri_matrix(range_obj, matrix, labels=None, matrix_name=None):
    n = matrix.shape[0]
    rows = range_obj.Rows.Count
    cols = range_obj.Columns.Count
    if rows == n + 1 and cols == n + 1:
        data = [[None] * (n + 1) for _ in range(n + 1)]
        if matrix_name:
            data[0][0] = matrix_name
        if labels:
            for j in range(n):
                data[0][j+1] = labels[j]
            for i in range(n):
                data[i+1][0] = labels[i]
        else:
            for j in range(n):
                data[0][j+1] = f"变量{j+1}"
            for i in range(n):
                data[i+1][0] = f"变量{i+1}"
        for i in range(n):
            for j in range(i+1):
                data[i+1][j+1] = matrix[i, j]
        range_obj.Value2 = data
        return
    if rows == n + 3 and cols == n + 1:
        vals = range_obj.Value2
        if vals is None or not isinstance(vals, (list, tuple)):
            vals = [[None] * cols for _ in range(rows)]
        data = [[None] * cols for _ in range(rows)]
        for r in range(min(2, len(vals))):
            for c in range(min(cols, len(vals[r]))):
                data[r][c] = vals[r][c]
        if matrix_name:
            data[0][0] = matrix_name
        data[2][0] = "矩阵系数"
        if labels:
            for j, label in enumerate(labels):
                if 1 + j < cols:
                    data[2][1 + j] = label
        else:
            if len(vals) > 2 and isinstance(vals[2], (list, tuple)):
                for j in range(1, min(cols, len(vals[2]))):
                    data[2][j] = vals[2][j]
        for i in range(n):
            if labels and i < len(labels) and labels[i] is not None:
                data[3 + i][0] = labels[i]
            elif len(vals) > 3 + i and isinstance(vals[3 + i], (list, tuple)) and vals[3 + i][0] is not None:
                data[3 + i][0] = vals[3 + i][0]
            else:
                data[3 + i][0] = f"变量{i+1}"
            for j in range(i+1):
                data[3 + i][1 + j] = matrix[i, j]
        range_obj.Value2 = data
        return
    raise ValueError(f"区域大小 ({rows}x{cols}) 与矩阵大小 ({n}x{n}) 不匹配，需要 ({n+1}x{n+1}) 或 ({n+3}x{n+1})")

def _is_positive_semidefinite(matrix, tol=1e-12):
    """
    检查矩阵是否为半正定（带容差）。
    如果所有特征值 >= -tol，则认为半正定。
    """
    try:
        eigvals = np.linalg.eigvalsh(matrix)
        return np.all(eigvals >= -tol)
    except Exception as e:
        print(f"半正定检查失败: {e}")
        return False

def _nearest_psd(matrix, max_iter=1000, tol=1e-12):
    """
    使用 Higham 算法找到最接近的半正定矩阵（投影到半正定锥），
    然后缩放对角线为1，以确保返回的是有效的相关矩阵。
    返回的矩阵满足：半正定，对角线为1，元素在[-1,1]内。
    """
    # 对称化
    sym = (matrix + matrix.T) / 2.0
    # 投影到正半定
    eigvals, eigvecs = np.linalg.eigh(sym)
    eigvals = np.maximum(eigvals, 0)   # 将所有负特征值置0
    psd = eigvecs @ np.diag(eigvals) @ eigvecs.T
    # 强制对称
    psd = (psd + psd.T) / 2.0

    # 缩放对角线为1，保持半正定
    diag = np.diag(psd)
    # 避免除零或负数
    diag = np.maximum(diag, 1e-12)
    inv_sqrt_diag = 1.0 / np.sqrt(diag)
    # 执行缩放： D * psd * D
    psd_scaled = psd * np.outer(inv_sqrt_diag, inv_sqrt_diag)
    # 强制对称（避免浮点误差）
    psd_scaled = (psd_scaled + psd_scaled.T) / 2.0
    # 确保对角线严格为1
    np.fill_diagonal(psd_scaled, 1.0)
    # 裁剪到 [-1,1]（安全）
    psd_scaled = np.clip(psd_scaled, -1.0, 1.0)

    # 再次检查半正定性（若因浮点误差导致微小负特征值，再次投影）
    if not _is_positive_semidefinite(psd_scaled):
        eigvals2, eigvecs2 = np.linalg.eigh(psd_scaled)
        eigvals2 = np.maximum(eigvals2, 0)
        psd_scaled = eigvecs2 @ np.diag(eigvals2) @ eigvecs2.T
        psd_scaled = (psd_scaled + psd_scaled.T) / 2.0
        np.fill_diagonal(psd_scaled, 1.0)
        psd_scaled = np.clip(psd_scaled, -1.0, 1.0)

    return psd_scaled

def _adjust_to_psd_with_weights(target_matrix, weight_matrix):
    """
    加权最小二乘调整：在保持对角线为1和对称性的约束下，
    寻找与 target_matrix 尽可能接近的半正定矩阵。
    权重矩阵 weight_matrix 中元素为 0-100，权重越高，该元素调整幅度越小。
    改进版：使用 scipy 优化并增加特征值惩罚项，确保矩阵变为半正定。
    增加边界检查：确保结果元素在 [-1,1] 内，权重矩阵元素在 [0,100] 内。
    最终强制返回半正定且对角线为1的矩阵。
    """
    n = target_matrix.shape[0]
    # 确保矩阵是对称的
    target = (target_matrix + target_matrix.T) / 2.0
    # 确保权重矩阵对称（只考虑下三角权重，上三角取对称值）
    weight_sym = (weight_matrix + weight_matrix.T) / 2.0

    # 检查权重矩阵元素是否在 [0,100] 范围内
    if np.any(weight_sym < 0) or np.any(weight_sym > 100):
        raise ValueError("权重矩阵元素必须介于 0 和 100 之间")

    # 如果已经是半正定且对角线为1，直接返回
    if _is_positive_semidefinite(target) and np.allclose(np.diag(target), 1.0):
        print("矩阵已是半正定且对角线为1，无需调整")
        return target

    # 尝试使用 scipy 优化（如果可用）
    try:
        from scipy.optimize import minimize

        # 只优化非对角线元素（下三角，不包括对角线）
        # 变量索引：对于 i>j，变量 idx = i*(i-1)//2 + j
        # 总变量数 M = n*(n-1)//2
        M = n * (n - 1) // 2

        # 初始值：目标矩阵的非对角线元素
        x0 = []
        for i in range(n):
            for j in range(i):
                x0.append(target[i, j])

        # 边界：所有元素在[-1,1]
        bounds = [(-1.0, 1.0)] * M

        # 定义目标函数：加权平方误差 + 特征值惩罚
        # 惩罚系数 λ 可以调整，越大越强制半正定
        lambda_penalty = 1000.0

        def objective(x):
            # 重构对称矩阵（对角线为1）
            mat = np.eye(n)
            idx = 0
            for i in range(n):
                for j in range(i):
                    mat[i, j] = x[idx]
                    mat[j, i] = x[idx]
                    idx += 1
            # 加权平方误差
            err = 0.0
            for i in range(n):
                for j in range(n):
                    w = weight_sym[i, j]
                    diff = mat[i, j] - target[i, j]
                    err += w * diff * diff
            # 特征值惩罚（对负特征值的平方和）
            eigvals = np.linalg.eigvalsh(mat)
            penalty = 0.0
            for e in eigvals:
                if e < 0:
                    penalty += e * e
            return err + lambda_penalty * penalty

        # 优化
        res = minimize(objective, x0, method='L-BFGS-B', bounds=bounds,
                       options={'maxiter': 1000, 'ftol': 1e-12, 'gtol': 1e-10})

        if res.success:
            # 重构调整后的矩阵
            mat_adj = np.eye(n)
            idx = 0
            for i in range(n):
                for j in range(i):
                    mat_adj[i, j] = res.x[idx]
                    mat_adj[j, i] = res.x[idx]
                    idx += 1

            # 确保对角线为1（已经保证）
            np.fill_diagonal(mat_adj, 1.0)

            # 裁剪到 [-1,1]（安全）
            mat_adj = np.clip(mat_adj, -1.0, 1.0)

            # 强制半正定缩放
            mat_adj = _nearest_psd(mat_adj)

            # 计算变化量
            diff_norm = np.linalg.norm(mat_adj - target)
            print(f"优化成功，使用 scipy 调整矩阵，变化量范数: {diff_norm:.6f}")
            if diff_norm < 1e-12:
                print("警告：优化后矩阵与原始矩阵几乎相同，可能无需调整或权重设置导致无变化。")
            return mat_adj
        else:
            print(f"优化失败: {res.message}，回退到交替投影法")
    except ImportError:
        print("scipy.optimize 不可用，使用交替投影法")
    except Exception as e:
        print(f"加权调整异常: {e}，回退到交替投影法")

    # ---------- 改进的交替投影法（带权重，增加步长因子，最终强制半正定缩放） ----------
    print("开始交替投影调整...")
    # 初始化 X = target
    X = target.copy()
    max_iter = 100
    tol = 1e-8
    # 步长因子，控制向目标靠近的速度（0<eta<1）
    eta = 0.5
    # 最大权重（用于归一化，避免步长过大）
    max_weight = np.max(weight_sym)
    if max_weight == 0:
        max_weight = 1.0  # 避免除零

    for iteration in range(max_iter):
        X_prev = X.copy()
        # 1. 投影到半正定锥
        X = _nearest_psd(X)
        # 2. 强制对角线为1
        np.fill_diagonal(X, 1.0)
        # 3. 根据权重更新 X，使其向 target 靠近（加权梯度步）
        # 更新量 = eta * (weight_sym / max_weight) * (target - X)
        update = eta * (weight_sym / max_weight) * (target - X)
        X = X + update
        X = (X + X.T) / 2.0
        # 检查收敛
        change = np.linalg.norm(X - X_prev)
        if change < tol:
            print(f"交替投影在第 {iteration+1} 次迭代收敛，变化量 {change:.2e}")
            break
    # 最终确保对角线为1
    np.fill_diagonal(X, 1.0)
    # 裁剪到 [-1,1]
    X = np.clip(X, -1.0, 1.0)
    # 强制半正定缩放
    X = _nearest_psd(X)
    print("交替投影调整完成")
    return X

def _apply_rank_correlation(samples_dict, matrix_name_to_inputs, matrix_name_to_target_matrix, n_iterations):
    """
    使用 Iman‑Conover 方法调整样本的秩相关矩阵。
    参数：
        samples_dict: dict {key: np.ndarray} 包含所有输入样本（每个数组长度为 n_iterations）
        matrix_name_to_inputs: dict {mat_name: [(input_key, position), ...]}
        matrix_name_to_target_matrix: dict {mat_name: np.ndarray (n x n)} 目标相关矩阵
        n_iterations: 迭代次数
    """
    from scipy.stats import norm, spearmanr
    for mat_name, inputs in matrix_name_to_inputs.items():
        if len(inputs) < 2:
            continue
        target = matrix_name_to_target_matrix.get(mat_name)
        if target is None:
            print(f"⚠️ 跳过 {mat_name}：未找到目标矩阵")
            continue
        inputs.sort(key=lambda x: x[1])
        keys = [key for key, _ in inputs]
        print(f"调整矩阵 {mat_name}，变量: {keys}")

        # 提取样本矩阵 (n_iterations x n_vars)
        try:
            samples = np.column_stack([samples_dict[key].astype(float) for key in keys])
        except KeyError as e:
            print(f"❌ 在 samples_dict 中找不到键 {e}，请检查输入键是否匹配")
            continue
        # 剔除含有 NaN 的行
        nan_mask = np.any(np.isnan(samples), axis=1)
        valid_samples = samples[~nan_mask]
        print(f"有效样本数: {valid_samples.shape[0]}/{n_iterations}")

        if valid_samples.shape[0] < 2:
            print(f"⚠️ 有效样本不足，跳过调整")
            continue

        # 1. 计算原始样本的秩并转换为正态分位数
        ranks = np.argsort(np.argsort(valid_samples, axis=0), axis=0) + 1
        u = ranks / (valid_samples.shape[0] + 1)
        Z = norm.ppf(u)

        # 2. 计算 Z 的样本相关矩阵
        I_mat = np.corrcoef(Z, rowvar=False)

        # 3. 确保目标矩阵可进行 Cholesky 分解（先强制半正定）
        C = target.copy()
        # 确保 C 是半正定且对角线为1
        if not _is_positive_semidefinite(C) or not np.allclose(np.diag(C), 1.0):
            C = _nearest_psd(C)
        try:
            np.linalg.cholesky(C)
        except np.linalg.LinAlgError:
            # 尝试对角线扰动
            eigvals = np.linalg.eigvalsh(C)
            min_eig = min(eigvals)
            if min_eig < 0:
                C = C + np.eye(C.shape[0]) * (-min_eig + 1e-8)
            else:
                C = C + np.eye(C.shape[0]) * 1e-8
            # 再次强制半正定
            C = _nearest_psd(C)
            try:
                np.linalg.cholesky(C)
                print(f"已对矩阵 {mat_name} 添加扰动使其正定")
            except:
                print(f"❌ 矩阵 {mat_name} 仍无法进行 Cholesky 分解，跳过调整")
                continue

        # 4. 计算变换矩阵 S
        P = np.linalg.cholesky(C)
        Q = np.linalg.cholesky(I_mat)
        S = P @ np.linalg.inv(Q)

        # 5. 应用变换得到新的正态分位数
        Z_new = Z @ S.T
        new_ranks = np.argsort(np.argsort(Z_new, axis=0), axis=0) + 1

        # 6. 将新秩映射回原始样本值
        for col, key in enumerate(keys):
            orig = samples_dict[key].astype(float)          # 原始样本（可能含 NaN）
            # 对原始样本排序（NaN 会排到最后）
            orig_sorted = np.sort(orig)
            # 初始化新数组，保持 NaN 位置
            new_vals = np.full_like(orig, np.nan, dtype=float)
            rank_idx = new_ranks[:, col] - 1
            new_vals[~nan_mask] = orig_sorted[rank_idx]
            # 写回字典
            samples_dict[key] = new_vals

        # 7. 打印调整后第一个变量与第二个变量的秩相关系数（可选）
        if len(keys) >= 2:
            key1 = keys[0]
            key2 = keys[1]
            arr1 = samples_dict[key1][~nan_mask]
            arr2 = samples_dict[key2][~nan_mask]
            if len(arr1) > 1:
                rho, _ = spearmanr(arr1, arr2)
                print(f"调整后 {key1} 与 {key2} 的秩相关系数: {rho:.6f}")
        print(f"✅ 矩阵 {mat_name} 调整完成")

def _get_corrmat_info_from_cells(distribution_cells, app=None):
    if app is None:
        app = xl_app()
    matrix_name_to_inputs = {}
    for cell_addr, dist_funcs in distribution_cells.items():
        for func in dist_funcs:
            input_key = func.get('input_key')
            if not input_key:
                continue
            markers = func.get('markers', {})
            corrmat_val = markers.get('corrmat')
            if corrmat_val is not None and isinstance(corrmat_val, str):
                parts = corrmat_val.split(',')
                if len(parts) >= 2:
                    mat_name = parts[0].strip()
                    # 去除可能的外层引号（单引号或双引号）
                    if (mat_name.startswith('"') and mat_name.endswith('"')) or \
                       (mat_name.startswith("'") and mat_name.endswith("'")):
                        mat_name = mat_name[1:-1]
                    try:
                        pos = int(parts[1].strip())
                    except:
                        pos = 0
                    # 将输入键转为大写，以便与 live_arrays 中的键匹配
                    input_key_upper = input_key.upper()

                    # ========== 新增：检查矩阵是否被静音 ==========
                    if _is_corr_matrix_muted(app, mat_name):
                        print(f"矩阵 {mat_name} 已被静音，忽略其相关分布")
                        continue
                    # ==========================================

                    if mat_name not in matrix_name_to_inputs:
                        matrix_name_to_inputs[mat_name] = []
                    matrix_name_to_inputs[mat_name].append((input_key_upper, pos))
    return matrix_name_to_inputs

def _read_target_matrices(app, matrix_names):
    matrices = {}
    for name in matrix_names:
        print(f"尝试读取矩阵: {name}")
        range_obj = _get_named_range(app, name)
        if range_obj is None:
            print(f"❌ 命名区域 {name} 不存在（或无法通过 _get_named_range 获取）")
            continue
        print(f"✅ 命名区域 {name} 存在，区域地址: {range_obj.Address}")
        mat = _read_matrix_from_range(range_obj)
        if mat is None:
            print(f"❌ 读取矩阵 {name} 失败（_read_matrix_from_range 返回 None）")
            # 可选：打印区域内容前几行
            try:
                vals = range_obj.Value2
                print(f"区域内容预览: {vals[:3] if isinstance(vals, list) else vals}")
            except:
                pass
            continue
        matrices[name] = np.array(mat, dtype=float)
        print(f"✅ 成功读取矩阵 {name}，形状 {mat.shape}")
    return matrices

def _ensure_corrmat_in_formula(formula, matrix_name, position):
    formula_clean = formula.lstrip('=').strip()
    for dist_name in DISTRIBUTION_FUNCTION_NAMES:
        pattern = re.compile(rf'\b{re.escape(dist_name)}\s*\(', re.IGNORECASE)
        match = pattern.search(formula_clean)
        if match:
            start = match.start()
            depth = 1
            end = match.end()
            while end < len(formula_clean):
                if formula_clean[end] == '(':
                    depth += 1
                elif formula_clean[end] == ')':
                    depth -= 1
                    if depth == 0:
                        break
                end += 1
            if depth == 0:
                func_call = formula_clean[start:end+1]
                args_str = func_call[func_call.index('(')+1:func_call.rindex(')')]
                args = parse_args_with_nested_functions(args_str)
                new_attr = f'DriskCorrmat("{matrix_name}", {position})'
                corrmat_present = False
                for idx, arg in enumerate(args):
                    if arg.strip().upper().startswith('DRISKCORRMAT'):
                        corrmat_present = True
                        args[idx] = new_attr
                        break
                if not corrmat_present:
                    if args:
                        args.append(new_attr)
                    else:
                        args = [new_attr]
                new_args_str = ','.join(args)
                new_func_call = f"{dist_name}({new_args_str})"
                new_formula_clean = formula_clean[:start] + new_func_call + formula_clean[end+1:]
                return '=' + new_formula_clean
    return formula

# ---------- 新增：从公式中移除 DriskCorrmat 的辅助函数 ----------
def _remove_corrmat_from_cells(app, matrix_name):
    """
    遍历工作簿中所有单元格，移除包含指定矩阵名的 DriskCorrmat 调用。
    直接修改单元格公式，并处理多余逗号。
    """
    wb = app.ActiveWorkbook
    for sheet in wb.Worksheets:
        try:
            used = sheet.UsedRange
            if used is None:
                continue
            for cell in used:
                try:
                    formula = cell.Formula
                    if not isinstance(formula, str) or not formula.startswith('='):
                        continue
                    # 检查公式中是否包含 DriskCorrmat 及该矩阵名
                    # 匹配模式： DriskCorrmat(...matrix_name...)
                    # 由于矩阵名可能带引号，我们使用正则匹配并移除
                    # 模式： DriskCorrmat\s*\(\s*["']?matrix_name["']?\s*,\s*\d+\s*\)
                    # 注意：矩阵名可能被引号包围，也可能没有引号
                    quoted_name = f'"{re.escape(matrix_name)}"'
                    unquoted_name = re.escape(matrix_name)
                    pattern = re.compile(
                        r'DriskCorrmat\s*\(\s*(?:"' + re.escape(matrix_name) + r'"|\'' + re.escape(matrix_name) + r'\'|' + re.escape(matrix_name) + r')\s*,\s*\d+\s*\)',
                        re.IGNORECASE
                    )
                    new_formula = pattern.sub('', formula)
                    if new_formula != formula:
                        # 清理多余逗号
                        # 删除前后可能残留的逗号
                        # 处理 "function(arg1, arg2, , arg3)" 之类
                        new_formula = re.sub(r',\s*,', ',', new_formula)
                        new_formula = re.sub(r'\(\s*,', '(', new_formula)
                        new_formula = re.sub(r',\s*\)', ')', new_formula)
                        # 如果公式以 "= " 开头等，可以忽略
                        cell.Formula = new_formula
                        print(f"已从 {sheet.Name}!{cell.Address} 移除 DriskCorrmat({matrix_name})")
                except:
                    pass
        except:
            pass

def _remove_all_corrmat_from_cells(app, matrix_names):
    """从所有单元格中移除指定列表中的所有矩阵对应的 DriskCorrmat"""
    for mat_name in matrix_names:
        _remove_corrmat_from_cells(app, mat_name)

# ---------- 美观对话框 ----------
def _show_choice_dialog(title, message, option1, option2):
    """显示带有两个自定义按钮的对话框，返回 1（第一个按钮）或 2（第二个按钮）"""
    try:
        import tkinter as tk
        from tkinter import ttk
        root = tk.Tk()
        root.title(title)
        root.geometry("350x180")
        root.resizable(False, False)
        root.attributes('-topmost', True)
        # 居中
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
    except Exception as e:
        # 回退到 InputBox
        try:
            app = _get_excel_app()
            choice = app.InputBox(
                Prompt=f"{message}\n\n输入 1 选择“{option1}”，输入 2 选择“{option2}”：",
                Title=title,
                Type=1,
                Default="1"
            )
            if choice is False:
                return 0
            return int(choice)
        except:
            return 0

def _show_list_choice_dialog(title, message, items):
    """显示列表选择对话框，返回所选项目的索引（0-based）或 -1 如果取消"""
    try:
        import tkinter as tk
        from tkinter import ttk
        root = tk.Tk()
        root.title(title)
        root.geometry("450x350")
        root.resizable(False, False)
        root.attributes('-topmost', True)
        root.update_idletasks()
        x = (root.winfo_screenwidth() // 2) - (450 // 2)
        y = (root.winfo_screenheight() // 2) - (350 // 2)
        root.geometry(f"+{x}+{y}")

        label = ttk.Label(root, text=message, wraplength=430, padding=5)
        label.pack()

        listbox = tk.Listbox(root, height=12, font=("TkDefaultFont", 9))
        listbox.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        for item in items:
            listbox.insert(tk.END, item)

        result = tk.IntVar()
        result.set(-1)

        def on_select():
            sel = listbox.curselection()
            if sel:
                result.set(sel[0])
                root.destroy()

        def on_cancel():
            result.set(-1)
            root.destroy()

        btn_frame = ttk.Frame(root)
        btn_frame.pack(pady=10)
        ttk.Button(btn_frame, text="确定", width=10, command=on_select).pack(side=tk.LEFT, padx=10)
        ttk.Button(btn_frame, text="取消", width=10, command=on_cancel).pack(side=tk.RIGHT, padx=10)

        root.mainloop()
        return result.get()
    except Exception:
        # 回退到 InputBox
        try:
            app = _get_excel_app()
            prompt = f"{message}\n\n"
            for i, item in enumerate(items, 1):
                prompt += f"{i}. {item}\n"
            prompt += "\n请输入数字选择："
            choice = app.InputBox(prompt, title, Type=1, Default="1")
            if choice is False:
                return -1
            idx = int(choice) - 1
            if 0 <= idx < len(items):
                return idx
            return -1
        except:
            return -1

# ---------- 宏实现 ----------
@xl_macro
def DriskMakeCorr():
    """创建相关性矩阵，并为所选分布添加 DriskCorrmat 属性（下三角格式）"""
    app = _get_excel_app()
    try:
        # 1. 选择分布区域
        try:
            range_obj = app.InputBox(Prompt="请选择包含分布函数的单元格区域（每个单元格应为 Drisk 分布）",
                                     Title="DriskMakeCorr - 选择分布区域",
                                     Type=8)
            if range_obj is False:
                return
        except:
            xlcAlert("选择区域操作取消或失败")
            return

        invalid_cells = []
        cells_list = []
        for cell in range_obj:
            cells_list.append(cell)
            try:
                formula = cell.Formula
                if not isinstance(formula, str) or not formula.startswith('='):
                    invalid_cells.append(cell.Address)
                elif not is_distribution_function(formula):
                    invalid_cells.append(cell.Address)
            except:
                invalid_cells.append(cell.Address)
        if invalid_cells:
            xlcAlert(f"以下单元格不是有效的 Drisk 分布函数：{', '.join(invalid_cells)}\n请修正后重试。")
            return

        # 检查是否已有 DriskCorrmat 属性
        corrmat_cells = []
        for cell in cells_list:
            try:
                formula = cell.Formula
                if formula and isinstance(formula, str) and re.search(r'DriskCorrmat\s*\(', formula, re.IGNORECASE):
                    corrmat_cells.append(cell.Address)
            except:
                pass

        if corrmat_cells:
            xlcAlert(
                f"以下单元格已经包含 DriskCorrmat 属性，请先删除或修改后再执行 DriskMakeCorr：\n"
                f"{', '.join(corrmat_cells)}\n\n"
                f"建议：手动删除公式中的 DriskCorrmat(...) 部分，或使用 Excel 的查找替换功能移除。"
            )
            return

        n_vars = len(cells_list)

        try:
            matrix_top_left = app.InputBox(Prompt="请选择矩阵表格的左上角单元格（矩阵将向右下扩展，占用 (n+1)×(n+1) 区域）",
                                           Title="DriskMakeCorr - 矩阵位置",
                                           Type=8)
            if matrix_top_left is False:
                return
        except:
            xlcAlert("选择矩阵位置取消或失败")
            return

        wb = app.ActiveWorkbook
        existing_names = [name.Name for name in wb.Names]
        max_num = 0
        for name in existing_names:
            match = re.match(r'^相关矩阵_(\d+)$', name)
            if match:
                max_num = max(max_num, int(match.group(1)))
        new_num = max_num + 1
        matrix_name = f"相关矩阵_{new_num}"
        user_name = app.InputBox(f"矩阵名称（默认 {matrix_name}）:", "矩阵名称", Default=matrix_name)
        if user_name and isinstance(user_name, str):
            matrix_name = user_name.strip()

        sheet = matrix_top_left.Worksheet
        start_row = matrix_top_left.Row
        start_col = matrix_top_left.Column
        end_row = start_row + n_vars
        end_col = start_col + n_vars
        range_matrix = sheet.Range(sheet.Cells(start_row, start_col), sheet.Cells(end_row, end_col))

        labels = []
        for cell in cells_list:
            try:
                addr = cell.Address
                addr = addr.replace('$', '')
                labels.append(addr)
            except:
                labels.append(f"变量{len(labels)+1}")

        matrix = np.eye(n_vars)
        _write_lower_tri_matrix(range_matrix, matrix, labels, matrix_name=matrix_name)
        _set_named_range(app, matrix_name, range_matrix)

        pos = 1
        modified = []
        for cell in cells_list:
            formula = cell.Formula
            new_formula = _ensure_corrmat_in_formula(formula, matrix_name, pos)
            if new_formula != formula:
                cell.Formula = new_formula
                modified.append(cell.Address)
            pos += 1

        if modified:
            xlcAlert(f"已为 {len(modified)} 个单元格添加/更新 DriskCorrmat 属性。\n矩阵名称：{matrix_name}\n现在可以手动修改下三角部分的相关系数（对角线上方留空）。")
        else:
            xlcAlert("所有单元格已包含 DriskCorrmat，未作修改。")

    except Exception as e:
        xlcAlert(f"执行 DriskMakeCorr 时出错：{str(e)}")
        traceback.print_exc()

@xl_macro
def DriskCheckCorr():
    """检查相关性矩阵的有效性，并提供调整选项（支持下三角格式）"""
    app = _get_excel_app()
    try:
        wb = app.ActiveWorkbook
        matrix_names = []
        for name_obj in wb.Names:
            name = name_obj.Name
            if name.startswith('相关矩阵_') or name.startswith('矩阵'):
                matrix_names.append(name)
        if not matrix_names:
            xlcAlert("未找到任何相关性矩阵命名区域。请先使用 DriskMakeCorr 创建矩阵。")
            return

        matrix_name = None
        if len(matrix_names) == 1:
            matrix_name = matrix_names[0]
        else:
            prompt_lines = ["请选择要检查的矩阵名称（输入数字）:"]
            for i, name in enumerate(matrix_names, 1):
                prompt_lines.append(f"{i}. {name}")
            prompt = "\n".join(prompt_lines)
            try:
                choice = app.InputBox(prompt, "选择矩阵", Type=1, Default="1")
                if choice is False:
                    return
                idx = int(choice) - 1
                if 0 <= idx < len(matrix_names):
                    matrix_name = matrix_names[idx]
                else:
                    xlcAlert("输入的数字无效。")
                    return
            except:
                xlcAlert("输入无效。")
                return

        range_obj = _get_named_range(app, matrix_name)
        if range_obj is None:
            xlcAlert(f"无法找到命名区域 {matrix_name}。")
            return
        full_mat = _read_matrix_from_range(range_obj)
        if full_mat is None:
            xlcAlert("矩阵数据为空或格式不正确。")
            return

        is_psd = _is_positive_semidefinite(full_mat)
        if is_psd:
            xlcAlert(f"矩阵 {matrix_name} 是半正定的，有效。")
            return
        else:
            # 使用自定义对话框选择处理方式
            choice = _show_choice_dialog(
                title="DriskCheckCorr - 矩阵调整",
                message=f"矩阵 {matrix_name} 不是半正定。请选择处理方式：",
                option1="自动调整",
                option2="手动调整"
            )
            if choice == 1:  # 自动调整
                # 使用改进的 nearest_psd 确保结果为半正定且对角线为1
                adj = _nearest_psd(full_mat)
                # 验证调整后的矩阵
                if not _is_positive_semidefinite(adj):
                    # 如果仍然不是半正定，再强制缩放一次
                    adj = _nearest_psd(adj)
                np.fill_diagonal(adj, 1.0)
                overwrite = app.InputBox("是否覆盖原矩阵？(Y/N)", "覆盖选项", Type=2, Default="Y")
                if overwrite and overwrite.upper() == 'Y':
                    vals = range_obj.Value2
                    if vals and len(vals) > 0 and len(vals[0]) > 0:
                        labels = []
                        if len(vals[0]) > 1:
                            for j in range(1, len(vals[0])):
                                labels.append(vals[0][j])
                        else:
                            labels = None
                        _write_lower_tri_matrix(range_obj, adj, labels, matrix_name=matrix_name)
                    else:
                        _write_lower_tri_matrix(range_obj, adj, matrix_name=matrix_name)
                    xlcAlert(f"矩阵 {matrix_name} 已调整为半正定。")
                else:
                    new_name = f"{matrix_name}_adj"
                    new_range = range_obj.Offset(range_obj.Rows.Count + 2, 0).Resize(range_obj.Rows.Count, range_obj.Columns.Count)
                    vals = range_obj.Value2
                    labels = None
                    if vals and len(vals) > 0 and len(vals[0]) > 1:
                        labels = [vals[0][j] for j in range(1, len(vals[0]))]
                    _write_lower_tri_matrix(new_range, adj, labels, matrix_name=new_name)
                    _set_named_range(app, new_name, new_range)
                    xlcAlert(f"已创建新矩阵 {new_name} 并调整为半正定。")
            elif choice == 2:  # 手动调整 -> 仅创建权重矩阵
                weight_top_left = app.InputBox("请选择权重矩阵的左上角单元格（矩阵大小应与相关性矩阵相同）", "权重矩阵位置", Type=8)
                if weight_top_left is False:
                    return
                n = full_mat.shape[0]
                weight_mat = np.zeros((n, n), dtype=float)
                sheet = weight_top_left.Worksheet
                start_row = weight_top_left.Row
                start_col = weight_top_left.Column
                end_row = start_row + n
                end_col = start_col + n
                weight_range = sheet.Range(sheet.Cells(start_row, start_col), sheet.Cells(end_row, end_col))
                vals = range_obj.Value2
                labels = None
                if vals and len(vals) > 0 and len(vals[0]) > 1:
                    labels = [vals[0][j] for j in range(1, len(vals[0]))]
                _write_lower_tri_matrix(weight_range, weight_mat, labels, matrix_name="权重矩阵")
                weight_name = f"WeightMat_{matrix_name}"
                _set_named_range(app, weight_name, weight_range)
                # 存储映射关系
                _update_corr_weight_mapping(app, weight_name, matrix_name)
                xlcAlert(f"已创建权重矩阵 {weight_name}，请手动修改权重值（0-100），然后运行 DriskWeight 进行调整。")
            else:
                xlcAlert("未选择任何操作。")

    except Exception as e:
        xlcAlert(f"执行 DriskCheckCorr 时出错：{str(e)}")
        traceback.print_exc()

@xl_macro
def DriskWeight():
    """根据权重矩阵调整相关性矩阵（支持下三角格式）"""
    app = _get_excel_app()
    try:
        wb = app.ActiveWorkbook
        weight_names = []
        for name_obj in wb.Names:
            name = name_obj.Name
            if name.startswith('WeightMat_'):
                weight_names.append(name)
        if not weight_names:
            xlcAlert("未找到任何权重矩阵。请先使用 DriskCheckCorr 创建权重矩阵。")
            return

        # 选择权重矩阵
        weight_name = None
        if len(weight_names) == 1:
            weight_name = weight_names[0]
        else:
            prompt_lines = ["请选择要使用的权重矩阵（输入数字）:"]
            for i, name in enumerate(weight_names, 1):
                prompt_lines.append(f"{i}. {name}")
            prompt = "\n".join(prompt_lines)
            try:
                choice = app.InputBox(prompt, "选择权重矩阵", Type=1, Default="1")
                if choice is False:
                    return
                idx = int(choice) - 1
                if 0 <= idx < len(weight_names):
                    weight_name = weight_names[idx]
                else:
                    xlcAlert("输入的数字无效。")
                    return
            except:
                xlcAlert("输入无效。")
                return

        print(f"选中的权重矩阵: {weight_name}")

        # 通过映射表查找对应的相关性矩阵
        mapping = _get_mapping_table(app)
        corr_name = mapping.get(weight_name)
        print(f"映射表查找结果: {corr_name}")

        # 如果映射表中没有，尝试自动推断
        if corr_name is None:
            candidate = weight_name.replace('WeightMat_', '')
            if _get_named_range(app, candidate) is not None:
                corr_name = candidate
                print(f"自动推断匹配到: {corr_name}")

        # 再尝试 Copula 映射表
        _get_copula_mapping_table = None
        if corr_name is None:
            try:
                from copula_functions import _get_mapping_table as _get_copula_mapping_table
            except Exception:
                _get_copula_mapping_table = None
        if corr_name is None and _get_copula_mapping_table is not None:
            try:
                copula_mapping = _get_copula_mapping_table(app)
                corr_name = copula_mapping.get(weight_name)
                if corr_name and _get_named_range(app, corr_name) is not None:
                    print(f"从 Copula 映射表匹配到: {corr_name}")
                else:
                    corr_name = None
            except Exception as e:
                print(f"检查 Copula 映射表失败: {e}")

        # 如果仍然没有，让用户手动选择
        if corr_name is None or _get_named_range(app, corr_name) is None:
            # 获取所有相关性矩阵名称
            corr_names = []
            for name_obj in wb.Names:
                name = name_obj.Name
                if name.startswith('相关矩阵_') or name.startswith('矩阵'):
                    corr_names.append(name)
            if _get_copula_mapping_table is not None:
                try:
                    copula_mapping = _get_copula_mapping_table(app)
                    for c_name in copula_mapping.values():
                        if c_name and c_name not in corr_names and _get_named_range(app, c_name) is not None:
                            corr_names.append(c_name)
                except Exception:
                    pass
            if not corr_names:
                # fallback: include all non-weight named ranges
                for name_obj in wb.Names:
                    name = name_obj.Name
                    if not name.startswith('WeightMat_') and not name.startswith('__'):
                        corr_names.append(name)
            if not corr_names:
                xlcAlert(f"找不到对应的相关性矩阵或 Copula 矩阵，请先创建目标矩阵。")
                return
            msg = f"权重矩阵 {weight_name} 无法自动匹配到相关性矩阵。\n请选择要调整的目标矩阵："
            idx = _show_list_choice_dialog("选择目标矩阵", msg, corr_names)
            if idx < 0:
                xlcAlert("未选择目标矩阵，操作取消。")
                return
            corr_name = corr_names[idx]
            # 更新映射表
            _update_corr_weight_mapping(app, weight_name, corr_name)
            print(f"用户手动选择: {corr_name}")

        # 验证矩阵存在
        corr_range = _get_named_range(app, corr_name)
        if corr_range is None:
            xlcAlert(f"无法找到相关性矩阵 {corr_name}，操作取消。")
            return

        weight_range = _get_named_range(app, weight_name)
        if weight_range is None:
            xlcAlert(f"找不到权重矩阵区域 {weight_name}。")
            return

        # 读取矩阵数据
        print("读取相关性矩阵...")
        c_np = _read_matrix_from_range(corr_range)
        print("读取权重矩阵...")
        w_np = _read_matrix_from_range(weight_range)
        if c_np is None:
            xlcAlert("相关性矩阵读取失败，请检查矩阵区域是否为正确的下三角格式。")
            return
        if w_np is None:
            xlcAlert("权重矩阵读取失败，请检查权重矩阵区域是否为正确的下三角格式。")
            return

        if w_np.shape != c_np.shape:
            xlcAlert(f"权重矩阵大小 {w_np.shape} 与相关性矩阵大小 {c_np.shape} 不一致，无法调整。")
            return

        # 执行调整
        print("开始调整矩阵...")
        try:
            adj = _adjust_to_psd_with_weights(c_np, w_np)
        except ValueError as e:
            xlcAlert(f"权重矩阵调整失败：{e}")
            return
        print("调整完成。")

        # 确保最终结果半正定（_adjust_to_psd_with_weights 内部已调用 _nearest_psd，但再次确认）
        if not _is_positive_semidefinite(adj):
            adj = _nearest_psd(adj)
        np.fill_diagonal(adj, 1.0)

        # 询问覆盖
        overwrite = app.InputBox("是否覆盖原相关性矩阵？(Y/N)", "覆盖选项", Type=2, Default="Y")
        vals = corr_range.Value2
        labels = None
        if vals and len(vals) > 0 and len(vals[0]) > 1:
            labels = [vals[0][j] for j in range(1, len(vals[0]))]
        if overwrite and overwrite.upper() == 'Y':
            # 写入调整后的矩阵
            _write_lower_tri_matrix(corr_range, adj, labels, matrix_name=corr_name)
            # 验证写入是否成功
            written = _read_matrix_from_range(corr_range)
            if written is None:
                xlcAlert("警告：写入后无法读取矩阵，请检查区域。")
            else:
                if np.allclose(written, adj, rtol=1e-6, atol=1e-8):
                    print("验证成功：写入后的矩阵与调整后的矩阵一致。")
                else:
                    print("验证失败：写入后的矩阵与调整后的矩阵不一致。")
                    diff = np.linalg.norm(written - adj)
                    print(f"差异范数: {diff}")
            xlcAlert(f"矩阵 {corr_name} 已根据权重矩阵 {weight_name} 调整。")
        else:
            # 创建新矩阵
            new_name = f"{corr_name}_adj"
            sheet = corr_range.Worksheet
            start_row = corr_range.Row + corr_range.Rows.Count + 2
            start_col = corr_range.Column
            end_row = start_row + corr_range.Rows.Count - 1
            end_col = start_col + corr_range.Columns.Count - 1
            new_range = sheet.Range(sheet.Cells(start_row, start_col), sheet.Cells(end_row, end_col))
            _write_lower_tri_matrix(new_range, adj, labels, matrix_name=new_name)
            _set_named_range(app, new_name, new_range)
            # 更新映射：权重矩阵现在关联到新矩阵
            _update_corr_weight_mapping(app, weight_name, new_name)
            xlcAlert(f"已创建新矩阵 {new_name} 并根据权重调整。")
    except Exception as e:
        xlcAlert(f"执行 DriskWeight 时出错：{str(e)}")
        traceback.print_exc()

# ---------- 辅助函数：从模拟对象获取单元格数据 ----------
def _normalize_cell_key(cell):
    """将 Excel 单元格对象转换为规范化的键（工作表名!地址）"""
    try:
        sheet = cell.Worksheet.Name
        addr = cell.Address.replace('$', '')
        # 规范化工作表名（大写，去除可能的中文标点等）
        # 这里简单处理，实际可使用 index_functions.normalize_cell_key
        return f"{sheet.upper()}!{addr.upper()}"
    except:
        return None

def _get_data_from_cell(sim, cell):
    """
    从模拟对象中获取指定单元格的数据。
    优先尝试作为输出单元格获取，如果失败则尝试作为输入单元格获取。
    返回数据列表，若失败返回 None。
    """
    # 先尝试输出数据
    data = sim.get_output_data_by_range(cell)
    if data is not None:
        return data
    # 再尝试输入数据
    # 获取规范化键
    norm_key = _normalize_cell_key(cell)
    if norm_key is None:
        return None
    # 检查输入缓存中是否有该键
    if hasattr(sim, 'input_cache') and norm_key in sim.input_cache:
        return sim.input_cache[norm_key]
    # 尝试带索引的输入键（例如 "Sheet1!A1_1"） - 简单起见，遍历所有键
    # 假设分布单元格的输入键格式为 "Sheet1!A1_1" 或 "Sheet1!A1_MAKEINPUT"
    # 我们匹配以 norm_key 开头的键
    if hasattr(sim, 'input_cache'):
        prefix = norm_key + '_'
        for key in sim.input_cache:
            if key.startswith(prefix):
                return sim.input_cache[key]
    return None

# ---------- 函数实现 ----------
@xl_func("xl_cell cell1, xl_cell cell2, int sim_num: float", volatile=True)
@xl_arg("cell1", "range")
@xl_arg("cell2", "range")
def DriskCorr(cell1, cell2, sim_num=1):
    """返回两个单元格模拟数据的皮尔逊相关系数"""
    try:
        sim = get_simulation(sim_num)
        if sim is None:
            from simulation_manager import get_all_simulations
            all_sims = get_all_simulations()
            if all_sims:
                sim = next(iter(all_sims.values()))
            else:
                return float('nan')
        data1 = _get_data_from_cell(sim, cell1)
        data2 = _get_data_from_cell(sim, cell2)
        if data1 is None or data2 is None:
            return float('nan')
        # 过滤有效数据
        valid1 = []
        valid2 = []
        for a, b in zip(data1, data2):
            try:
                f1 = float(a)
                f2 = float(b)
                if not (np.isnan(f1) or np.isnan(f2)):
                    valid1.append(f1)
                    valid2.append(f2)
            except:
                continue
        if len(valid1) < 2:
            return float('nan')
        corr = np.corrcoef(valid1, valid2)[0, 1]
        return float(corr) if not np.isnan(corr) else float('nan')
    except Exception as e:
        print(f"DriskCorr错误: {e}")
        return float('nan')

@xl_func("xl_cell cell1, xl_cell cell2, int sim_num: float", volatile=True)
@xl_arg("cell1", "range")
@xl_arg("cell2", "range")
def DriskCorrRank(cell1, cell2, sim_num=1):
    """返回两个单元格模拟数据的斯皮尔曼秩相关系数"""
    try:
        sim = get_simulation(sim_num)
        if sim is None:
            from simulation_manager import get_all_simulations
            all_sims = get_all_simulations()
            if all_sims:
                sim = next(iter(all_sims.values()))
            else:
                return float('nan')
        data1 = _get_data_from_cell(sim, cell1)
        data2 = _get_data_from_cell(sim, cell2)
        if data1 is None or data2 is None:
            return float('nan')
        # 过滤有效数据
        valid1 = []
        valid2 = []
        for a, b in zip(data1, data2):
            try:
                f1 = float(a)
                f2 = float(b)
                if not (np.isnan(f1) or np.isnan(f2)):
                    valid1.append(f1)
                    valid2.append(f2)
            except:
                continue
        if len(valid1) < 2:
            return float('nan')
        try:
            from scipy.stats import spearmanr
            corr, _ = spearmanr(valid1, valid2)
            return float(corr) if not np.isnan(corr) else float('nan')
        except ImportError:
            def rank(arr):
                s = np.argsort(arr)
                ranks = np.empty(len(arr))
                ranks[s] = np.arange(1, len(arr)+1)
                return ranks
            r1 = rank(valid1)
            r2 = rank(valid2)
            corr = np.corrcoef(r1, r2)[0, 1]
            return float(corr) if not np.isnan(corr) else float('nan')
    except Exception as e:
        print(f"DriskCorrRank错误: {e}")
        return float('nan')

# ==================== 新增宏：移除单个相关性矩阵 ====================
@xl_macro
def DriskRemoveCorr():
    """彻底移除一个相关性矩阵及其权重矩阵，同时从所有分布函数中移除对应的 DriskCorrmat 属性。"""
    app = _get_excel_app()
    try:
        wb = app.ActiveWorkbook

        # 获取所有相关性矩阵名称（命名区域，或从映射表获取）
        corr_names = []
        for name_obj in wb.Names:
            name = name_obj.Name
            # 跳过内部名称
            if name.startswith('_xlfn') or name.startswith('__Drisk'):
                continue
            # 如果名称以“相关矩阵_”开头，认为是相关性矩阵
            if name.startswith('相关矩阵_') or name.startswith('矩阵'):
                corr_names.append(name)
            else:
                # 也可能没有前缀，通过映射表确认
                mapping = _get_mapping_table(app)
                if name in mapping.values():
                    corr_names.append(name)

        if not corr_names:
            xlcAlert("未找到任何相关性矩阵。")
            return

        # 选择矩阵（无确认环节）
        corr_name = None
        if len(corr_names) == 1:
            corr_name = corr_names[0]
        else:
            prompt_lines = ["请选择要移除的相关性矩阵（输入数字）:"]
            for i, name in enumerate(corr_names, 1):
                prompt_lines.append(f"{i}. {name}")
            prompt = "\n".join(prompt_lines)
            try:
                choice = app.InputBox(prompt, "DriskRemoveCorr - 选择矩阵", Type=1, Default="1")
                if choice is False:
                    return
                idx = int(choice) - 1
                if 0 <= idx < len(corr_names):
                    corr_name = corr_names[idx]
                else:
                    xlcAlert("输入的数字无效。")
                    return
            except:
                xlcAlert("输入无效。")
                return

        # 1. 从所有分布函数公式中移除 DriskCorrmat 调用（针对该矩阵）
        _remove_corrmat_from_cells(app, corr_name)

        # 2. 删除相关性矩阵的命名区域及其数据区域（先清空，再删除）
        range_obj = _get_named_range(app, corr_name)
        if range_obj is not None:
            # 清空区域内容（覆盖为空）
            range_obj.ClearContents()
            # 删除命名区域
            try:
                wb.Names(corr_name).Delete()
                print(f"已删除命名区域: {corr_name}")
            except Exception as e:
                print(f"删除命名区域 {corr_name} 失败: {e}")

        # 3. 删除关联的权重矩阵（如果存在）
        mapping = _get_mapping_table(app)
        weight_name = None
        for w_name, c_name in mapping.items():
            if c_name == corr_name:
                weight_name = w_name
                break
        if weight_name:
            weight_range = _get_named_range(app, weight_name)
            if weight_range is not None:
                weight_range.ClearContents()
                try:
                    wb.Names(weight_name).Delete()
                    print(f"已删除权重矩阵命名区域: {weight_name}")
                except Exception as e:
                    print(f"删除权重矩阵 {weight_name} 失败: {e}")
            # 从映射表中删除
            del mapping[weight_name]
            _save_mapping_table(app, mapping)

        # 4. 从静音列表中移除（如果存在）
        _remove_muted_corr_matrix(app, corr_name)

        xlcAlert(f"相关性矩阵“{corr_name}”及其关联的权重矩阵已彻底移除，且所有分布函数中的 DriskCorrmat 已清理。")
    except Exception as e:
        xlcAlert(f"执行 DriskRemoveCorr 时出错：{str(e)}")
        traceback.print_exc()

# ==================== 新增宏：移除所有相关性矩阵 ====================
@xl_macro
def DriskRemoveCorrAll():
    """彻底移除所有相关性矩阵和权重矩阵，同时从所有分布函数中移除对应的 DriskCorrmat 属性。"""
    app = _get_excel_app()
    try:
        wb = app.ActiveWorkbook

        # 获取所有相关性矩阵名称
        corr_names = []
        for name_obj in wb.Names:
            name = name_obj.Name
            if name.startswith('_xlfn') or name.startswith('__Drisk'):
                continue
            if name.startswith('相关矩阵_') or name.startswith('矩阵'):
                corr_names.append(name)
            else:
                mapping = _get_mapping_table(app)
                if name in mapping.values():
                    corr_names.append(name)

        if not corr_names:
            xlcAlert("未找到任何相关性矩阵。")
            return

        # 1. 从所有分布函数公式中移除所有 DriskCorrmat 调用（针对每个矩阵）
        _remove_all_corrmat_from_cells(app, corr_names)

        # 2. 删除所有矩阵及其权重
        mapping = _get_mapping_table(app)
        for corr_name in corr_names:
            # 删除相关性矩阵数据区域和命名区域
            range_obj = _get_named_range(app, corr_name)
            if range_obj is not None:
                range_obj.ClearContents()
                try:
                    wb.Names(corr_name).Delete()
                    print(f"已删除命名区域: {corr_name}")
                except Exception as e:
                    print(f"删除命名区域 {corr_name} 失败: {e}")

            # 删除关联的权重矩阵
            weight_name = None
            for w_name, c_name in mapping.items():
                if c_name == corr_name:
                    weight_name = w_name
                    break
            if weight_name:
                weight_range = _get_named_range(app, weight_name)
                if weight_range is not None:
                    weight_range.ClearContents()
                    try:
                        wb.Names(weight_name).Delete()
                        print(f"已删除权重矩阵命名区域: {weight_name}")
                    except Exception as e:
                        print(f"删除权重矩阵 {weight_name} 失败: {e}")
                # 从映射表中删除
                del mapping[weight_name]

        # 清空映射表
        mapping.clear()
        _save_mapping_table(app, mapping)

        # 清空静音列表
        _save_muted_corr_matrices(app, [])

        # 可选：清空隐藏工作表中的矩阵数据（保留工作表，只清空内容）
        try:
            sheet = wb.Worksheets("DriskHiddenMapping")
            sheet.UsedRange.ClearContents()
        except:
            pass

        xlcAlert(f"已移除所有相关性矩阵（共 {len(corr_names)} 个）及其关联的权重矩阵，且所有分布函数中的 DriskCorrmat 已清理。")
    except Exception as e:
        xlcAlert(f"执行 DriskRemoveCorrAll 时出错：{str(e)}")
        traceback.print_exc()

# ==================== 新增宏：切换矩阵静音状态 ====================
@xl_macro
def DriskMuteCorr():
    """切换相关性矩阵的静音状态（静音时模拟不使用该矩阵）"""
    app = _get_excel_app()
    try:
        wb = app.ActiveWorkbook

        # 获取所有相关性矩阵名称
        corr_names = []
        for name_obj in wb.Names:
            name = name_obj.Name
            if name.startswith('_xlfn') or name.startswith('__Drisk'):
                continue
            if name.startswith('相关矩阵_') or name.startswith('矩阵'):
                corr_names.append(name)
            else:
                mapping = _get_mapping_table(app)
                if name in mapping.values():
                    corr_names.append(name)

        if not corr_names:
            xlcAlert("未找到任何相关性矩阵。")
            return

        # 选择矩阵
        corr_name = None
        if len(corr_names) == 1:
            corr_name = corr_names[0]
        else:
            prompt_lines = ["请选择要切换静音状态的相关性矩阵（输入数字）:"]
            for i, name in enumerate(corr_names, 1):
                prompt_lines.append(f"{i}. {name}")
            prompt = "\n".join(prompt_lines)
            try:
                choice = app.InputBox(prompt, "DriskMuteCorr - 选择矩阵", Type=1, Default="1")
                if choice is False:
                    return
                idx = int(choice) - 1
                if 0 <= idx < len(corr_names):
                    corr_name = corr_names[idx]
                else:
                    xlcAlert("输入的数字无效。")
                    return
            except:
                xlcAlert("输入无效。")
                return

        # 检查当前状态并切换
        if _is_corr_matrix_muted(app, corr_name):
            _remove_muted_corr_matrix(app, corr_name)
            status = "取消静音"
        else:
            _add_muted_corr_matrix(app, corr_name)
            status = "静音"

        xlcAlert(f"相关性矩阵“{corr_name}”已{status}。")
    except Exception as e:
        xlcAlert(f"执行 DriskMuteCorr 时出错：{str(e)}")
        traceback.print_exc()