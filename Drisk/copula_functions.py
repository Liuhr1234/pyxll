# copula_functions.py
"""
Copula 管理模块
提供 DriskMakeCopula, DriskCheckCopula, DriskRemoveCopula, DriskRemoveCopulaAll, DriskMuteCopula 宏
以及模拟时所需的辅助函数，与现有 corrmat 机制无缝集成。
新增：Empirical Copula（经验 Copula）支持
新增：旋转 Copula（ClaytonRX, ClaytonRY, GumbelRX, GumbelRY）支持
"""

import re
import traceback
import numpy as np
import win32com.client
from pyxll import xl_macro, xl_func, xl_app, xlcAlert
from typing import List, Tuple, Optional, Dict, Any

# 导入现有系统模块
from constants import DISTRIBUTION_FUNCTION_NAMES
from formula_parser import is_distribution_function, parse_args_with_nested_functions
from attribute_functions import ERROR_MARKER
from simulation_manager import get_simulation

# ---------- 常量与辅助 ----------
MAPPING_NAME = "__DriskCopulaMapping__"      # 隐藏名称，存储 {weight_name: copula_name}
MUTED_NAMES = "__DriskMutedCopulas__"       # 存储静音 copula 名称列表
COPLUA_TYPES = {
    1: "Gaussian",
    2: "t",
    3: "Clayton",
    4: "ClaytonR",
    5: "Gumbel",
    6: "GumbelR",
    7: "Frank",
    8: "Empirical",
    9: "ClaytonRX",      # 新增：旋转 Clayton（X轴）
    10: "ClaytonRY",     # 新增：旋转 Clayton（Y轴）
    11: "GumbelRX",      # 新增：旋转 Gumbel（X轴）
    12: "GumbelRY",      # 新增：旋转 Gumbel（Y轴）
}

# 旋转 Copula 映射：新类型 -> (基础类型, 要旋转的列索引)
ROTATION_MAP = {
    "FrankRX": ("Frank", 1),      # 垂直翻转 (Y轴) -> 第二列变换
    "ClaytonRX": ("Clayton", 0),  # 水平翻转 (X轴) -> 第一列变换
    "ClaytonRY": ("Clayton", 1),  # 垂直翻转 (Y轴) -> 第二列变换
    "GumbelRX": ("Gumbel", 0),    # 水平翻转 (X轴)
    "GumbelRY": ("Gumbel", 1),    # 垂直翻转 (Y轴)
}

# 需要参数 theta 的 Copula 类型（阿基米德及旋转版本）
PARAM_THETA_TYPES = {"Clayton", "ClaytonR", "Gumbel", "GumbelR", "Frank",
                     "ClaytonRX", "ClaytonRY", "GumbelRX", "GumbelRY"}

def _get_excel_app():
    """安全获取 Excel 应用对象"""
    try:
        from com_fixer import _safe_excel_app
        return _safe_excel_app()
    except ImportError:
        return xl_app()

def _get_named_range(app, name: str):
    """返回命名区域对象，若不存在返回 None"""
    try:
        return app.ActiveWorkbook.Names(name).RefersToRange
    except:
        return None

def _set_named_range(app, name: str, range_obj):
    """创建命名区域"""
    try:
        try:
            app.ActiveWorkbook.Names(name).Delete()
        except:
            pass
        app.ActiveWorkbook.Names.Add(Name=name, RefersTo=range_obj)
    except Exception as e:
        raise ValueError(f"无法创建命名区域 {name}: {e}")

def _read_matrix_from_range(range_obj):
    """从 Excel 区域读取数值矩阵（下三角格式），返回完整对称矩阵。

    支持标准下三角矩阵格式、纯数值方阵、以及 Copula 自定义格式
    （第一行矩阵名称，第二行模式信息，第三行“矩阵系数”标签）。
    对于 Empirical Copula，还支持读取数据源信息行。
    """
    try:
        vals = range_obj.Value2
        if vals is None:
            return None
        if not isinstance(vals, (list, tuple)):
            vals = [[vals]]
        rows = len(vals)
        cols = len(vals[0])
        if rows < 2 or cols < 2:
            return None

        def _is_number(value):
            if value is None:
                return False
            if isinstance(value, (int, float, np.integer, np.floating)):
                return True
            if isinstance(value, str):
                try:
                    float(value)
                    return True
                except ValueError:
                    return False
            return False

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

        # 纯数值方阵或纯数值下三角矩阵（没有标签行/列）
        if rows == cols:
            all_numeric_or_blank = True
            for r in vals:
                for v in r:
                    if v is not None and not _is_number(v):
                        all_numeric_or_blank = False
                        break
                if not all_numeric_or_blank:
                    break
            if all_numeric_or_blank:
                n = rows
                mat = np.zeros((n, n), dtype=float)
                for i in range(n):
                    for j in range(n):
                        val = vals[i][j]
                        if val is None and j > i:
                            val = vals[j][i]
                        try:
                            mat[i, j] = float(val) if val is not None else 0.0
                        except Exception:
                            mat[i, j] = 0.0
                return mat

        # 标准下三角矩阵格式：第一行和第一列为标签
        data_rows = rows - 1
        data_cols = cols - 1
        if data_rows < 2 or data_cols < 2:
            return None
        data = [[0.0] * data_cols for _ in range(data_rows)]
        for i in range(data_rows):
            for j in range(data_cols):
                try:
                    val = vals[i + 1][j + 1]
                    data[i][j] = float(val) if val is not None else 0.0
                except Exception:
                    data[i][j] = 0.0
        mat = np.array(data, dtype=float)
        if mat.shape[0] != mat.shape[1]:
            return None
        mat = (mat + mat.T) / 2.0
        return mat
    except Exception as e:
        print(f"读取矩阵失败: {e}")
        return None

def _write_lower_tri_matrix(range_obj, matrix, labels=None, matrix_name=None):
    """将对称矩阵写入区域（下三角格式），带标签行/列。

    支持标准下三角矩阵和 Copula 自定义格式。
    对于 Empirical Copula，使用特殊格式（3行2列）。
    """
    n = matrix.shape[0]
    rows = range_obj.Rows.Count
    cols = range_obj.Columns.Count

    def safe_get_row(row_vals, idx):
        return row_vals[idx] if isinstance(row_vals, (list, tuple)) and idx < len(row_vals) else None

    # 标准下三角格式 (n+1) x (n+1)
    if rows == n + 1 and cols == n + 1:
        data = [[None] * (n + 1) for _ in range(n + 1)]
        if matrix_name:
            data[0][0] = matrix_name
        if labels:
            for j in range(n):
                data[0][j + 1] = labels[j]
            for i in range(n):
                data[i + 1][0] = labels[i]
        else:
            for j in range(n):
                data[0][j + 1] = f"变量{j+1}"
            for i in range(n):
                data[i + 1][0] = f"变量{i+1}"
        for i in range(n):
            for j in range(i + 1):
                data[i + 1][j + 1] = matrix[i, j]
        range_obj.Value2 = data
        return

    # Copula 自定义格式 (n+3) x (n+1)
    if rows == n + 3 and cols == n + 1:
        vals = range_obj.Value2
        if vals is None or not isinstance(vals, (list, tuple)):
            vals = [[None] * cols for _ in range(rows)]
        data = [[None] * cols for _ in range(rows)]
        if isinstance(vals, (list, tuple)):
            for c in range(min(cols, len(vals[0]))):
                data[0][c] = safe_get_row(vals[0], c)
            if len(vals) > 1:
                for c in range(min(cols, len(vals[1]))):
                    data[1][c] = safe_get_row(vals[1], c)
        if matrix_name:
            data[0][0] = matrix_name
        data[2][0] = "矩阵系数"
        if labels:
            for j, label in enumerate(labels):
                if 1 + j < cols:
                    data[2][1 + j] = label
        else:
            if len(vals) > 2:
                for j in range(1, min(cols, len(vals[2]))):
                    data[2][j] = safe_get_row(vals[2], j)
        for i in range(n):
            if labels and i < len(labels) and labels[i] is not None:
                data[3 + i][0] = labels[i]
            elif len(vals) > 3 + i:
                data[3 + i][0] = safe_get_row(vals[3 + i], 0) or f"变量{i+1}"
            else:
                data[3 + i][0] = f"变量{i+1}"
        for i in range(n):
            for j in range(i + 1):
                data[3 + i][1 + j] = matrix[i, j]
        range_obj.Value2 = data
        return

    # 如果是 Empirical Copula 的 3x2 格式，特殊处理
    if rows == 3 and cols == 2:
        # 直接覆盖写入
        data = [[None] * cols for _ in range(rows)]
        # 第一行：矩阵名称
        data[0][0] = matrix_name if matrix_name else ""
        # 第二行：模式、插值标志
        # 这里 matrix 实际上不是矩阵，而是传递的 (interpolate_flag, data_range) 元组
        if isinstance(matrix, tuple) and len(matrix) == 2:
            interpolate_flag, data_range = matrix
            data[1][0] = "模式"
            data[1][1] = "Empirical"
            # 插值标志存放在第二行第三列（如果列数不够，放在第二行第二列之后？但列数只有2，所以无法存放）
            # 因为列数只有2，需要将插值标志作为第二行第二列的值，而第三列不存在。调整：第二行第一列"模式"，第二列"Empirical"，插值标志放在第三行？根据需求，参数位置为 TRUE/FALSE，最后一行第一列为"数据源"，第二列为数据的位置。
            # 这里我们改为：第二行第一列"模式"，第二列"Empirical"；第三行第一列"数据源"，第二列区域地址；插值标志放在第三行第三列？没有第三列。
            # 重新设计：使用3行3列格式，保留第二行第三列放插值标志，第三行第一列"数据源"，第二列区域地址，第三列留空。
            # 由于我们在调用此函数时传入的是3x3区域，因此需要修改调用处。但为了兼容，我们直接在调用时写入3x3区域。
            # 为简化，在 DriskMakeCopula 中直接处理 Empirical 的写入，不调用此函数。
            pass
        range_obj.Value2 = data
        return

    raise ValueError(f"区域大小 ({rows}x{cols}) 与矩阵大小 ({n}x{n}) 不匹配，需要 ({n+1}x{n+1}) 或 ({n+3}x{n+1})")

def _is_positive_semidefinite(matrix, tol=1e-12):
    try:
        eigvals = np.linalg.eigvalsh(matrix)
        return np.all(eigvals >= -tol)
    except:
        return False

def _nearest_psd(matrix, max_iter=1000, tol=1e-12):
    """Higham 算法投影到半正定锥，并缩放对角线为1"""
    sym = (matrix + matrix.T) / 2.0
    eigvals, eigvecs = np.linalg.eigh(sym)
    eigvals = np.maximum(eigvals, 0)
    psd = eigvecs @ np.diag(eigvals) @ eigvecs.T
    psd = (psd + psd.T) / 2.0
    diag = np.diag(psd)
    diag = np.maximum(diag, 1e-12)
    inv_sqrt = 1.0 / np.sqrt(diag)
    psd_scaled = psd * np.outer(inv_sqrt, inv_sqrt)
    psd_scaled = (psd_scaled + psd_scaled.T) / 2.0
    np.fill_diagonal(psd_scaled, 1.0)
    psd_scaled = np.clip(psd_scaled, -1.0, 1.0)
    return psd_scaled

def _adjust_to_psd_with_weights(target_matrix, weight_matrix):
    """加权最小二乘调整（复用 corrmat 函数）"""
    try:
        from corrmat_functions import _adjust_to_psd_with_weights as corr_adjust
        return corr_adjust(target_matrix, weight_matrix)
    except ImportError:
        # 简单实现 fallback
        n = target_matrix.shape[0]
        w = (weight_matrix + weight_matrix.T) / 2.0
        # 权重归一化
        max_w = np.max(w)
        if max_w == 0:
            max_w = 1.0
        w = w / max_w
        # 加权平均
        mat = target_matrix.copy()
        for i in range(n):
            for j in range(i):
                w_ij = w[i,j]
                if w_ij > 0:
                    mat[i,j] = mat[j,i] = target_matrix[i,j]  # 保持原值（权重影响步长？简单处理）
        return _nearest_psd(mat)

def _sample_positive_stable(alpha, size):
    """生成正稳定分布样本，alpha=1/theta，theta>=1"""
    if alpha <= 0.0 or alpha > 1.0:
        raise ValueError("alpha must be in (0,1]")
    if np.isclose(alpha, 1.0):
        return np.ones(size, dtype=float)
    u = np.random.uniform(0.0, np.pi, size=size)
    e = np.random.exponential(1.0, size=size)
    sin_u = np.sin(u)
    sin_alpha_u = np.sin(alpha * u)
    sin_1_alpha_u = np.sin((1.0 - alpha) * u)
    sin_u = np.where(np.abs(sin_u) < 1e-16, 1e-16, sin_u)
    w = sin_alpha_u / (sin_u ** (1.0 / alpha))
    w *= (sin_1_alpha_u / e) ** ((1.0 - alpha) / alpha)
    w = np.maximum(w, 1e-16)
    return w

# ========== 新增：Empirical Copula 辅助函数 ==========
def _read_data_range(app, range_str: str) -> np.ndarray:
    """读取数据区域，返回二维 numpy 数组（浮点数），验证所有单元格均为数字，否则抛出异常"""
    try:
        if '!' in range_str:
            sheet_name, addr = range_str.split('!', 1)
            # 去掉可能的外层单引号（Excel 中工作表名含空格或特殊字符时会被加上单引号）
            if sheet_name.startswith("'") and sheet_name.endswith("'"):
                sheet_name = sheet_name[1:-1]
            sheet = app.ActiveWorkbook.Worksheets(sheet_name)
        else:
            sheet = app.ActiveSheet
            addr = range_str
        range_obj = sheet.Range(addr)
        values = range_obj.Value2
        # 后续代码保持不变...
        if values is None:
            raise ValueError("区域为空")
        # 转换为二维列表
        if not isinstance(values, (list, tuple)):
            values = [[values]]
        else:
            if not isinstance(values[0], (list, tuple)):
                values = [values]
        # 转换为浮点数并检查有效性
        data = []
        for row in values:
            row_data = []
            for val in row:
                if val is None:
                    raise ValueError("区域包含空单元格")
                try:
                    num = float(val)
                    row_data.append(num)
                except:
                    raise ValueError(f"单元格 {val} 不是数字")
            data.append(row_data)
        arr = np.array(data, dtype=float)
        if arr.ndim != 2:
            raise ValueError("区域不是二维数据")
        return arr
    except Exception as e:
        raise ValueError(f"读取数据区域失败: {e}")
    
def _compute_pseudo_obs(data: np.ndarray) -> np.ndarray:
    """
    计算伪观测值（empirical copula 的均匀样本）
    使用 rank/(n+1) 方法，避免边界 0 和 1。
    返回形状 (n, d) 的数组，每列在 (0,1) 内。
    """
    n, d = data.shape
    pseudo = np.zeros((n, d))
    for j in range(d):
        # 使用 scipy 的 rankdata 处理 ties（平均秩）
        try:
            from scipy.stats import rankdata
            ranks = rankdata(data[:, j], method='average')
        except ImportError:
            # 后备：numpy 的 argsort 处理（不处理 ties 平均值）
            ranks = np.argsort(np.argsort(data[:, j])) + 1.0
        # 归一化到 (0,1)
        pseudo[:, j] = ranks / (n + 1.0)
    return pseudo

def _generate_empirical_copula_uniforms(interpolate: bool, pseudo_obs: np.ndarray, n_samples: int) -> np.ndarray:
    """
    从经验 Copula 生成均匀样本。
    interpolate: True 时添加均匀噪声平滑；False 时直接有放回抽样。
    返回形状 (n_samples, d) 的数组，值在 (0,1) 内。
    """
    n_hist, d = pseudo_obs.shape
    if n_hist == 0:
        raise ValueError("历史数据为空")
    if interpolate:
        # 平滑插值：抽样后添加均匀噪声，噪声幅度 = 1/(n_hist+1) 的一半
        delta = 0.5 / (n_hist + 1.0)
        # 随机选择索引（有放回）
        indices = np.random.randint(0, n_hist, size=n_samples)
        base = pseudo_obs[indices, :]
        noise = np.random.uniform(-delta, delta, size=(n_samples, d))
        u = base + noise
        # 裁剪到 (0,1) 并避免极端边界
        u = np.clip(u, 1e-12, 1.0 - 1e-12)
        return u
    else:
        # 直接有放回抽样
        indices = np.random.randint(0, n_hist, size=n_samples)
        return pseudo_obs[indices, :]

# ---------- 映射表管理 ----------
def _get_mapping_table(app):
    """读取 copula 权重矩阵映射表 {weight_name: copula_name}"""
    try:
        range_obj = _get_named_range(app, MAPPING_NAME)
        if range_obj is None:
            return {}
        vals = range_obj.Value2
        if vals is None:
            return {}
        if not isinstance(vals, (list, tuple)):
            vals = [vals]
        mapping = {}
        for row in vals:
            if isinstance(row, (list, tuple)) and len(row) >= 2:
                w_name = str(row[0]).strip()
                c_name = str(row[1]).strip()
                if w_name and c_name:
                    mapping[w_name] = c_name
        return mapping
    except:
        return {}

def _save_mapping_table(app, mapping: Dict[str, str]):
    """保存映射表到隐藏工作表"""
    if not mapping:
        return
    try:
        wb = app.ActiveWorkbook
        sheet_name = "DriskHiddenMapping"
        try:
            sheet = wb.Worksheets(sheet_name)
        except:
            sheet = wb.Worksheets.Add()
            sheet.Name = sheet_name
            sheet.Visible = 0
        # 清空原有内容（使用固定区域）
        used = sheet.UsedRange
        if used:
            used.ClearContents()
        # 写入两列数据
        data = [[w, c] for w, c in mapping.items()]
        if data:
            start_cell = sheet.Cells(1, 1)
            end_cell = sheet.Cells(len(data), 2)
            range_obj = sheet.Range(start_cell, end_cell)
            range_obj.Value2 = data
            _set_named_range(app, MAPPING_NAME, range_obj)
    except Exception as e:
        print(f"保存 copula 映射表失败: {e}")

def _update_copula_weight_mapping(app, weight_name: str, copula_name: str):
    mapping = _get_mapping_table(app)
    mapping[weight_name] = copula_name
    _save_mapping_table(app, mapping)

def _remove_mapping_entry(app, weight_name: str):
    mapping = _get_mapping_table(app)
    if weight_name in mapping:
        del mapping[weight_name]
        _save_mapping_table(app, mapping)

# ---------- 静音列表管理 ----------
def _get_muted_copulas(app):
    try:
        range_obj = _get_named_range(app, MUTED_NAMES)
        if range_obj is None:
            return []
        vals = range_obj.Value2
        if vals is None:
            return []
        if isinstance(vals, str):
            text = vals.strip()
        elif isinstance(vals, (list, tuple)):
            text = ",".join(str(v).strip() for row in vals for v in (row if isinstance(row, (list, tuple)) else [row]) if v is not None)
        else:
            text = str(vals).strip()
        if not text:
            return []
        return [n.strip() for n in text.split(',') if n.strip()]
    except:
        return []

def _save_muted_copulas(app, muted_list):
    if not muted_list:
        try:
            app.ActiveWorkbook.Names(MUTED_NAMES).Delete()
        except:
            pass
        return
    muted_list = list(dict.fromkeys(muted_list))
    text = ",".join(muted_list)
    try:
        wb = app.ActiveWorkbook
        sheet_name = "DriskHiddenMapping"
        try:
            sheet = wb.Worksheets(sheet_name)
        except:
            sheet = wb.Worksheets.Add()
            sheet.Name = sheet_name
            sheet.Visible = 0
        cell = sheet.Cells(1, 3)   # 使用第三列
        cell.Value = text
        _set_named_range(app, MUTED_NAMES, cell)
    except Exception as e:
        print(f"保存静音列表失败: {e}")

def _add_muted_copula(app, copula_name):
    muted = _get_muted_copulas(app)
    if copula_name not in muted:
        muted.append(copula_name)
        _save_muted_copulas(app, muted)

def _remove_muted_copula(app, copula_name):
    muted = _get_muted_copulas(app)
    if copula_name in muted:
        muted.remove(copula_name)
        _save_muted_copulas(app, muted)

def _is_copula_muted(app, copula_name) -> bool:
    return copula_name in _get_muted_copulas(app)

# ---------- 辅助对话框 ----------
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
    except Exception as e:
        # 回退到 InputBox
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

def _show_list_choice_dialog(title, message, items):
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
    except:
        # 回退 InputBox
        app = _get_excel_app()
        prompt = f"{message}\n\n"
        for i, item in enumerate(items, 1):
            prompt += f"{i}. {item}\n"
        prompt += "\n请输入数字选择："
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

# ---------- 在分布公式中管理 DriskCopula ----------
def _ensure_copula_in_formula(formula, copula_name, position):
    """向分布公式中添加或更新 DriskCopula 属性"""
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
                new_attr = f'DriskCopula("{copula_name}", {position})'
                copula_present = False
                for idx, arg in enumerate(args):
                    if arg.strip().upper().startswith('DRISKCOPULA'):
                        copula_present = True
                        args[idx] = new_attr
                        break
                if not copula_present:
                    if args:
                        args.append(new_attr)
                    else:
                        args = [new_attr]
                new_args_str = ','.join(args)
                new_func_call = f"{dist_name}({new_args_str})"
                new_formula_clean = formula_clean[:start] + new_func_call + formula_clean[end+1:]
                return '=' + new_formula_clean
    return formula

def _remove_copula_from_cells(app, copula_name):
    """从所有单元格中移除指定 copula 的 DriskCopula 调用"""
    wb = app.ActiveWorkbook
    pattern = re.compile(
        r'DriskCopula\s*\(\s*["\']?' + re.escape(copula_name) + r'["\']?\s*,\s*\d+\s*\)',
        re.IGNORECASE
    )
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
                    new_formula = pattern.sub('', formula)
                    if new_formula != formula:
                        # 清理多余逗号
                        new_formula = re.sub(r',\s*,', ',', new_formula)
                        new_formula = re.sub(r'\(\s*,', '(', new_formula)
                        new_formula = re.sub(r',\s*\)', ')', new_formula)
                        cell.Formula = new_formula
                except:
                    pass
        except:
            pass

def _remove_all_copula_from_cells(app, copula_names):
    for name in copula_names:
        _remove_copula_from_cells(app, name)

# ---------- 从分布单元格中提取 Copula 信息 ----------
def _get_copula_info_from_cells(distribution_cells, app=None):
    """返回 {copula_name: {'type': ..., 'param': ..., 'matrix': ..., 'inputs': [...]}} 
       对于 Empirical 额外包含 'interpolate' (bool), 'pseudo_obs' (np.ndarray), 'data_range' (str)
    """
    if app is None:
        app = _get_excel_app()
    copula_info = {}
    for cell_addr, dist_funcs in distribution_cells.items():
        for func in dist_funcs:
            input_key = func.get('input_key')
            if not input_key:
                continue
            markers = func.get('markers', {})
            copula_val = markers.get('copula')
            if copula_val is not None and isinstance(copula_val, str):
                parts = copula_val.split(',')
                if len(parts) >= 2:
                    copula_name = parts[0].strip()
                    # 去除引号
                    if (copula_name.startswith('"') and copula_name.endswith('"')) or \
                       (copula_name.startswith("'") and copula_name.endswith("'")):
                        copula_name = copula_name[1:-1]
                    try:
                        pos = int(parts[1].strip())
                    except:
                        pos = 0
                    input_key_upper = input_key.upper()
                    if _is_copula_muted(app, copula_name):
                        continue
                    if copula_name not in copula_info:
                        # 获取 copula 矩阵的定义（命名区域）
                        range_obj = _get_named_range(app, copula_name)
                        if range_obj is None:
                            print(f"警告: 找不到 copula 命名区域 {copula_name}")
                            continue
                        vals = range_obj.Value2
                        if vals is None or not isinstance(vals, (list, tuple)):
                            continue
                        # 解析类型和参数
                        try:
                            # 第二行第一列为"模式"，第二列为类型
                            copula_type = str(vals[1][1]).strip() if len(vals) > 1 and len(vals[1]) > 1 else ""
                            param = None
                            if copula_type in ("t", "Clayton", "ClaytonR", "Gumbel", "GumbelR", "Frank",
                                               "ClaytonRX", "ClaytonRY", "GumbelRX", "GumbelRY"):
                                param = float(vals[1][2]) if len(vals[1]) > 2 else None
                            # 对于 Empirical，第二行第二列为插值标志（字符串 TRUE/FALSE）
                            interpolate = False
                            data_range = None
                            if copula_type == "Empirical":
                                # 读取插值标志（可能在第二行第二列或第三列）
                                if len(vals[1]) > 2:
                                    flag_str = str(vals[1][2]).strip().upper()
                                    interpolate = flag_str == "TRUE"
                                # 读取数据源：最后一行第一列为"数据源"，第二列为地址
                                last_row_idx = len(vals) - 1
                                if last_row_idx >= 0 and len(vals[last_row_idx]) >= 2:
                                    data_source_label = str(vals[last_row_idx][0]).strip()
                                    if data_source_label == "数据源":
                                        data_range = str(vals[last_row_idx][1]).strip()
                        except Exception as e:
                            copula_type = ""
                            param = None
                            interpolate = False
                            data_range = None
                        # 对于 Gaussian 和 t，还需要读取矩阵系数
                        matrix = None
                        if copula_type in ("Gaussian", "t"):
                            matrix = _read_matrix_from_range(range_obj)
                            if matrix is None:
                                print(f"无法读取 copula 矩阵 {copula_name}")
                                continue
                        # 对于 Empirical，读取历史数据并计算伪观测值
                        pseudo_obs = None
                        if copula_type == "Empirical":
                            if data_range:
                                try:
                                    hist_data = _read_data_range(app, data_range)
                                    pseudo_obs = _compute_pseudo_obs(hist_data)
                                    print(f"Empirical Copula {copula_name}: 历史数据形状 {hist_data.shape}, 伪观测值计算完成")
                                except Exception as e:
                                    print(f"读取 Empirical 历史数据失败: {e}")
                                    continue
                            else:
                                print(f"Empirical Copula {copula_name} 缺少数据源")
                                continue
                        copula_info[copula_name] = {
                            'type': copula_type,
                            'param': param,
                            'matrix': matrix,
                            'interpolate': interpolate,
                            'data_range': data_range,
                            'pseudo_obs': pseudo_obs,
                            'inputs': []
                        }
                    copula_info[copula_name]['inputs'].append((input_key_upper, pos))
    # 对每个 copula 的 inputs 按位置排序
    for info in copula_info.values():
        info['inputs'].sort(key=lambda x: x[1])
    return copula_info

# ---------- 旋转 Copula 辅助函数 ----------
def _apply_rotation_to_uniforms(u: np.ndarray, rot_type: str) -> np.ndarray:
    """
    对生成的 uniform 样本应用旋转（列变换）
    rot_type: "FrankRX", "ClaytonRX", "ClaytonRY", "GumbelRX", "GumbelRY"
    返回变换后的数组。
    """
    if rot_type not in ROTATION_MAP:
        return u
    base_type, col_idx = ROTATION_MAP[rot_type]
    rotated = u.copy()
    rotated[:, col_idx] = 1.0 - rotated[:, col_idx]
    rotated = np.clip(rotated, 1e-12, 1.0 - 1e-12)
    return rotated

# ---------- 应用 copula 调整（复用秩相关）----------
def _apply_copula_rank_correlation(samples_dict, copula_info, n_iterations):
    """
    对每个 copula 应用样本生成调整，支持六种理论 Copula 以及 Empirical Copula。
    """
    for copula_name, info in copula_info.items():
        inputs = info['inputs']
        if len(inputs) < 2:
            continue
        copula_type = info['type']
        param = info.get('param')
        matrix = info.get('matrix')
        try:
            if copula_type in ("Gaussian", "t"):
                if matrix is None:
                    print(f"跳过 {copula_name}：没有有效矩阵")
                    continue
                _apply_copula_to_samples(samples_dict, inputs, copula_type, param, matrix, n_iterations)
                print(f"已应用 {copula_type} copula 调整")
            elif copula_type in ("Clayton", "ClaytonR", "Gumbel", "GumbelR", "Frank",
                                 "ClaytonRX", "ClaytonRY", "GumbelRX", "GumbelRY"):
                if param is None:
                    print(f"跳过 {copula_name}：缺少参数")
                    continue
                _apply_copula_to_samples(samples_dict, inputs, copula_type, param, matrix, n_iterations)
                print(f"已应用 {copula_type} copula 调整")
            elif copula_type == "Empirical":
                pseudo_obs = info.get('pseudo_obs')
                if pseudo_obs is None:
                    print(f"跳过 {copula_name}：没有伪观测值数据")
                    continue
                interpolate = info.get('interpolate', False)
                # 生成 uniform 样本
                u_matrix = _generate_empirical_copula_uniforms(interpolate, pseudo_obs, n_iterations)
                if u_matrix is None:
                    continue
                # 应用排序映射（与现有 Copula 相同逻辑）
                inputs_sorted = sorted(inputs, key=lambda x: x[1])
                keys = [key for key, _ in inputs_sorted]
                try:
                    sample_matrix = np.column_stack([samples_dict[key].astype(float) for key in keys])
                except KeyError as e:
                    print(f"Empirical Copula 处理失败: 找不到输入键 {e}")
                    continue
                nan_mask = np.any(np.isnan(sample_matrix), axis=1)
                valid_matrix = sample_matrix[~nan_mask]
                if valid_matrix.shape[0] < 2:
                    print(f"跳过 {copula_name}：有效样本不足")
                    continue
                # 对每个变量进行排序映射
                for col, key in enumerate(keys):
                    orig = samples_dict[key].astype(float)
                    valid_orig = orig[~nan_mask]
                    if valid_orig.size == 0:
                        continue
                    sorted_vals = np.sort(valid_orig)
                    idx = np.floor(u_matrix[:, col] * valid_orig.size).astype(int)
                    idx = np.clip(idx, 0, valid_orig.size - 1)
                    new_vals = np.full_like(orig, np.nan, dtype=float)
                    new_vals[~nan_mask] = sorted_vals[idx]
                    samples_dict[key] = new_vals
                print(f"已应用 Empirical copula 调整 (interpolate={interpolate})")
            else:
                print(f"跳过未知 Copula 类型 {copula_type}")
        except Exception as e:
            print(f"{copula_name} Copula 调整失败：{e}")

def _apply_copula_to_samples(samples_dict, inputs, copula_type, param, matrix, n_iterations):
    """将指定 copula 的 Uniform 相关结构映射回原始样本分布（用于理论 Copula）"""
    if len(inputs) < 2:
        return
    inputs_sorted = sorted(inputs, key=lambda x: x[1])
    keys = [key for key, _ in inputs_sorted]
    try:
        sample_matrix = np.column_stack([samples_dict[key].astype(float) for key in keys])
    except KeyError as e:
        print(f"Copula 处理失败: 找不到输入键 {e}")
        return
    nan_mask = np.any(np.isnan(sample_matrix), axis=1)
    valid_matrix = sample_matrix[~nan_mask]
    if valid_matrix.shape[0] < 2:
        print(f"跳过 {copula_type} copula：有效样本不足")
        return
    n_valid = valid_matrix.shape[0]
    u_matrix = _generate_copula_uniforms(copula_type, n_valid, len(keys), param=param, matrix=matrix)
    if u_matrix is None:
        return
    u_matrix = np.clip(u_matrix, 0.0, 1.0 - 1e-12)
    for col, key in enumerate(keys):
        orig = samples_dict[key].astype(float)
        valid_orig = orig[~nan_mask]
        if valid_orig.size == 0:
            continue
        sorted_vals = np.sort(valid_orig)
        idx = np.floor(u_matrix[:, col] * valid_orig.size).astype(int)
        idx = np.clip(idx, 0, valid_orig.size - 1)
        new_vals = np.full_like(orig, np.nan, dtype=float)
        new_vals[~nan_mask] = sorted_vals[idx]
        samples_dict[key] = new_vals

def _generate_copula_uniforms(copula_type, n, d, param=None, matrix=None):
    """生成指定 copula 类型的二维 Uniform(0,1) 样本矩阵（理论 Copula）
       支持旋转类型：FrankRX, ClaytonRX, ClaytonRY, GumbelRX, GumbelRY
    """
    from scipy.stats import norm, t

    # 处理旋转 Copula
    if copula_type in ROTATION_MAP:
        base_type, col_idx = ROTATION_MAP[copula_type]
        # 递归生成基础 Copula 样本
        u_base = _generate_copula_uniforms(base_type, n, d, param=param, matrix=matrix)
        if u_base is None:
            return None
        # 应用列旋转
        u_rot = u_base.copy()
        u_rot[:, col_idx] = 1.0 - u_rot[:, col_idx]
        u_rot = np.clip(u_rot, 1e-12, 1.0 - 1e-12)
        return u_rot

    if n <= 0 or d <= 0:
        raise ValueError("样本数量和维度必须为正")

    if copula_type == "Gaussian":
        if matrix is None:
            raise ValueError("Gaussian copula 需要相关矩阵")
        C = matrix.copy()
        C = _nearest_psd(C)
        np.fill_diagonal(C, 1.0)
        try:
            L = np.linalg.cholesky(C)
        except np.linalg.LinAlgError:
            raise ValueError("Gaussian copula 相关矩阵无法 Cholesky 分解")
        z = np.random.standard_normal(size=(n, d))
        x = z @ L.T
        return norm.cdf(x)

    if copula_type == "t":
        if matrix is None:
            raise ValueError("t copula 需要相关矩阵")
        if param is None or param <= 0:
            raise ValueError("t copula 需要正的自由度参数")
        C = matrix.copy()
        C = _nearest_psd(C)
        np.fill_diagonal(C, 1.0)
        try:
            L = np.linalg.cholesky(C)
        except np.linalg.LinAlgError:
            raise ValueError("t copula 相关矩阵无法 Cholesky 分解")
        z = np.random.standard_normal(size=(n, d))
        x = z @ L.T
        w = np.random.chisquare(param, size=n) / param
        x = x / np.sqrt(w[:, None])
        return t.cdf(x, df=param)

    if copula_type == "Clayton":
        if param is None or param <= 0:
            raise ValueError("Clayton copula 需要正参数")
        v = np.random.gamma(1.0 / param, 1.0, size=n)
        e = np.random.exponential(1.0, size=(n, d))
        u = (1.0 + e / v[:, None]) ** (-1.0 / param)
        return np.clip(u, 0.0, 1.0)

    if copula_type == "ClaytonR":
        u = _generate_copula_uniforms("Clayton", n, d, param=param)
        return 1.0 - u

    if copula_type == "Gumbel":
        if param is None or param < 1.0:
            raise ValueError("Gumbel copula 需要参数 theta >= 1")
        alpha = 1.0 / param
        if np.isclose(alpha, 1.0):
            w = np.ones(n, dtype=float)
        else:
            w = _sample_positive_stable(alpha, n)
        e = np.random.exponential(1.0, size=(n, d))
        u = np.exp(- (e / w[:, None]) ** alpha)
        return np.clip(u, 0.0, 1.0)

    if copula_type == "GumbelR":
        u = _generate_copula_uniforms("Gumbel", n, d, param=param)
        return 1.0 - u

    if copula_type == "Frank":
        if param is None or param <= 0:
            raise ValueError("Frank copula 需要正参数 theta")
        # 使用条件方法生成 Frank copula 样本
        u = np.random.uniform(0, 1, (n, d))
        for i in range(1, d):
            w = np.random.uniform(0, 1, n)
            exp_theta_u_prev = np.exp(-param * u[:, i-1])
            exp_theta = np.exp(-param)
            numerator = w * (exp_theta - 1)
            denominator = 1 - w + w * exp_theta_u_prev
            # 避免分母为0或log参数无效
            denominator = np.maximum(denominator, 1e-12)
            arg = 1 + numerator / denominator
            arg = np.clip(arg, 1e-12, 1 - 1e-12)  # 确保在(0,1)内
            u[:, i] = -1.0 / param * np.log(arg)
        return np.clip(u, 0.0, 1.0)

    raise ValueError(f"不支持的 Copula 类型: {copula_type}")

# ---------- 宏实现 ----------
@xl_macro
def DriskMakeCopula():
    """创建 Copula 矩阵，并为所选分布添加 DriskCopula 属性"""
    app = _get_excel_app()
    try:
        # 1. 选择分布区域
        try:
            range_obj = app.InputBox(
                Prompt="请选择包含分布函数的单元格区域（每个单元格应为 Drisk 分布）",
                Title="DriskMakeCopula - 选择分布区域",
                Type=8
            )
            if range_obj is False:
                return
        except:
            xlcAlert("选择区域操作取消或失败")
            return

        # 验证每个单元格是否包含有效的分布函数
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

        # 检查是否已有 DriskCopula 属性
        copula_cells = []
        for cell in cells_list:
            try:
                formula = cell.Formula
                if formula and isinstance(formula, str) and re.search(r'DriskCopula\s*\(', formula, re.IGNORECASE):
                    copula_cells.append(cell.Address)
            except:
                pass
        if copula_cells:
            xlcAlert(
                f"以下单元格已经包含 DriskCopula 属性，请先删除或修改后再执行 DriskMakeCopula：\n"
                f"{', '.join(copula_cells)}\n\n建议：手动删除公式中的 DriskCopula(...) 部分。"
            )
            return

        n_vars = len(cells_list)

        # 2. 选择矩阵位置
        try:
            matrix_top_left = app.InputBox(
                Prompt="请选择矩阵表格的左上角单元格（矩阵将向右下扩展）",
                Title="DriskMakeCopula - 矩阵位置",
                Type=8
            )
            if matrix_top_left is False:
                return
        except:
            xlcAlert("选择矩阵位置取消或失败")
            return

        # 3. 选择 Copula 模式（更新提示，包含新的旋转类型）
        try:
            mode_in = app.InputBox(
                Prompt="请输入 Copula 模式编号：\n"
                      "1=Gaussian\n2=t\n3=Clayton\n4=ClaytonR\n5=Gumbel\n6=GumbelR\n7=Frank\n8=Empirical\n"
                      "9=ClaytonRX\n10=ClaytonRY\n11=GumbelRX\n12=GumbelRY",
                Title="DriskMakeCopula - 模式",
                Type=1,
                Default="1"
            )
            if mode_in is False:
                return
            mode = int(mode_in)
            if mode not in COPLUA_TYPES:
                xlcAlert("无效的模式编号，请输入 1-12")
                return
        except:
            xlcAlert("输入无效")
            return

        copula_type = COPLUA_TYPES[mode]
        param = None
        interpolate = False
        data_range = None

        # 对于 Empirical，特殊处理
        if copula_type == "Empirical":
            # 询问是否采用平滑插值（默认是）
            try:
                import tkinter as tk
                from tkinter import ttk
                root = tk.Tk()
                root.title("Empirical Copula 选项")
                root.geometry("300x150")
                root.attributes('-topmost', True)
                root.update_idletasks()
                x = (root.winfo_screenwidth() // 2) - (300 // 2)
                y = (root.winfo_screenheight() // 2) - (150 // 2)
                root.geometry(f"+{x}+{y}")

                label = ttk.Label(root, text="是否采用平滑插值？", padding=10)
                label.pack()
                result = tk.BooleanVar()
                result.set(True)
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
                interpolate = result.get()
            except Exception as e:
                # 回退到 InputBox
                choice = app.InputBox("是否采用平滑插值？\n输入 1 表示是，0 表示否：", "插值选项", Type=1, Default="1")
                if choice is False:
                    return
                interpolate = (int(choice) == 1)

            # 选择历史数据区域
            try:
                data_range_obj = app.InputBox(
                    Prompt="请选择历史数据区域（纯数字，每列代表一个变量）",
                    Title="Empirical Copula - 数据区域",
                    Type=8
                )
                if data_range_obj is False:
                    return
                # 验证数据区域全为数字
                data_range_addr = data_range_obj.Address
                # 获取完整地址（含工作表名）
                sheet_name = data_range_obj.Worksheet.Name
                # 处理工作表名含空格时加引号
                if ' ' in sheet_name or any(c in sheet_name for c in "[]:?*"):
                    safe_sheet = f"'{sheet_name}'"
                else:
                    safe_sheet = sheet_name
                data_range = f"{safe_sheet}!{data_range_addr}"
                # 读取数据并验证
                hist_data = _read_data_range(app, data_range)
                # 验证列数等于变量数
                if hist_data.shape[1] != n_vars:
                    xlcAlert(f"历史数据列数 ({hist_data.shape[1]}) 与所选分布变量数 ({n_vars}) 不一致。")
                    return
                print(f"Empirical Copula 历史数据形状: {hist_data.shape}")
            except Exception as e:
                xlcAlert(f"数据区域无效: {str(e)}")
                return

        # 对于 t，需要 v 参数
        if copula_type == "t":
            try:
                v_in = app.InputBox(
                    Prompt="请输入 t Copula 的自由度参数 v (必须 >0)：",
                    Title="DriskMakeCopula - t 参数",
                    Type=1,
                    Default="3"
                )
                if v_in is False:
                    return
                param = float(v_in)
                if param <= 0:
                    xlcAlert("自由度必须为正数")
                    return
            except:
                xlcAlert("无效的自由度")
                return
        elif copula_type in PARAM_THETA_TYPES:  # 包括所有阿基米德及其旋转版本
            try:
                theta_in = app.InputBox(
                    Prompt=f"请输入 {copula_type} Copula 的参数 θ (必须 >0)：",
                    Title="DriskMakeCopula - 参数",
                    Type=1,
                    Default="2"
                )
                if theta_in is False:
                    return
                param = float(theta_in)
                if param <= 0:
                    xlcAlert("θ 必须为正数")
                    return
            except:
                xlcAlert("无效的参数")
                return

        # 生成矩阵名称
        wb = app.ActiveWorkbook
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

        # 准备写入区域
        sheet = matrix_top_left.Worksheet
        start_row = matrix_top_left.Row
        start_col = matrix_top_left.Column

        # 计算需要的行数和列数
        if copula_type in ("Gaussian", "t"):
            rows = n_vars + 3
            cols = n_vars + 1
        elif copula_type == "Empirical":
            # 使用 3 行 3 列：第一行名称，第二行模式+插值标志，第三行数据源
            rows = 3
            cols = 3
        else:
            rows = 3
            cols = 3

        end_row = start_row + rows - 1
        end_col = start_col + cols - 1
        range_matrix = sheet.Range(sheet.Cells(start_row, start_col), sheet.Cells(end_row, end_col))

        # 准备标签
        labels = []
        for cell in cells_list:
            try:
                addr = cell.Address.replace('$', '')
                labels.append(addr)
            except:
                labels.append(f"变量{len(labels)+1}")

        # 写入内容
        data = [[None] * cols for _ in range(rows)]
        # 第一行：矩阵名称
        data[0][0] = matrix_name

        if copula_type in ("Gaussian", "t"):
            # 第二行：模式及参数
            data[1][0] = "模式"
            data[1][1] = copula_type
            if param is not None:
                data[1][2] = param
            # 第三行及以后：矩阵系数区域
            data[2][0] = "矩阵系数"
            for j, label in enumerate(labels):
                if 1 + j < cols:
                    data[2][1 + j] = label
            for i, label in enumerate(labels):
                if 3 + i < rows:
                    data[3 + i][0] = label
            # 填充单位矩阵
            for i in range(n_vars):
                for j in range(n_vars):
                    if i == j:
                        if 3 + i < rows and 1 + j < cols:
                            data[3 + i][1 + j] = 1.0
                    elif i > j:
                        if 3 + i < rows and 1 + j < cols:
                            data[3 + i][1 + j] = 0.0
            range_matrix.Value2 = data
        elif copula_type == "Empirical":
            # 第二行：模式、插值标志
            data[1][0] = "模式"
            data[1][1] = copula_type
            data[1][2] = "TRUE" if interpolate else "FALSE"
            # 第三行：数据源
            data[2][0] = "数据源"
            data[2][1] = data_range
            range_matrix.Value2 = data
        else:
            # 其他 Copula（包括旋转类型 ClaytonRX, ClaytonRY, GumbelRX, GumbelRY）
            data[1][0] = "模式"
            data[1][1] = copula_type
            if param is not None:
                data[1][2] = param
            data[2][0] = "维度"
            data[2][1] = n_vars
            range_matrix.Value2 = data

        # 创建命名区域
        _set_named_range(app, matrix_name, range_matrix)

        # 为每个选中的分布添加 DriskCopula 属性
        pos = 1
        modified = []
        for cell in cells_list:
            formula = cell.Formula
            new_formula = _ensure_copula_in_formula(formula, matrix_name, pos)
            if new_formula != formula:
                cell.Formula = new_formula
                modified.append(cell.Address)
            pos += 1

        if modified:
            xlcAlert(f"已为 {len(modified)} 个单元格添加/更新 DriskCopula 属性。\nCopula 名称：{matrix_name}\n类型：{copula_type}\n参数：{param if param else '无'}")
        else:
            xlcAlert("所有单元格已包含 DriskCopula，未作修改。")

    except Exception as e:
        xlcAlert(f"执行 DriskMakeCopula 时出错：{str(e)}")
        traceback.print_exc()

@xl_macro
def DriskCheckCopula():
    """检查 Copula 的有效性（Gaussian/t 检查矩阵正定性，其他检查参数范围，Empirical 检查数据区域有效性）"""
    app = _get_excel_app()
    try:
        wb = app.ActiveWorkbook
        # 获取所有 copula 命名区域
        copula_names = []
        for name_obj in wb.Names:
            name = name_obj.Name
            if name.startswith('Copula_') or name.startswith('copula'):
                copula_names.append(name)
        if not copula_names:
            xlcAlert("未找到任何 Copula 矩阵命名区域。请先使用 DriskMakeCopula 创建 Copula。")
            return

        copula_name = None
        if len(copula_names) == 1:
            copula_name = copula_names[0]
        else:
            prompt_lines = ["请选择要检查的 Copula（输入数字）:"]
            for i, name in enumerate(copula_names, 1):
                prompt_lines.append(f"{i}. {name}")
            prompt = "\n".join(prompt_lines)
            try:
                choice = app.InputBox(prompt, "选择 Copula", Type=1, Default="1")
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

        # 读取矩阵内容，获取类型和参数
        vals = range_obj.Value2
        if vals is None or not isinstance(vals, (list, tuple)):
            xlcAlert("Copula 数据为空或格式不正确。")
            return
        try:
            copula_type = str(vals[1][1]).strip() if len(vals) > 1 and len(vals[1]) > 1 else ""
            param = None
            if len(vals[1]) > 2:
                param = vals[1][2]
            interpolate = False
            data_range = None
            if copula_type == "Empirical":
                # 插值标志
                if len(vals[1]) > 2:
                    flag_str = str(vals[1][2]).strip().upper()
                    interpolate = (flag_str == "TRUE")
                # 数据源
                last_row_idx = len(vals) - 1
                if last_row_idx >= 0 and len(vals[last_row_idx]) >= 2:
                    data_source_label = str(vals[last_row_idx][0]).strip()
                    if data_source_label == "数据源":
                        data_range = str(vals[last_row_idx][1]).strip()
        except:
            copula_type = ""
            param = None

        if copula_type in ("Gaussian", "t"):
            full_mat = _read_matrix_from_range(range_obj)
            if full_mat is None:
                xlcAlert("矩阵数据读取失败。")
                return
            is_psd = _is_positive_semidefinite(full_mat)
            if is_psd:
                xlcAlert(f"Copula {copula_name} 的矩阵是半正定的，有效。")
                return
            else:
                choice = _show_choice_dialog(
                    title="DriskCheckCopula - 矩阵调整",
                    message=f"Copula {copula_name} 的矩阵不是半正定。请选择处理方式：",
                    option1="自动调整",
                    option2="手动调整"
                )
                if choice == 1:
                    adj = _nearest_psd(full_mat)
                    overwrite = app.InputBox("是否覆盖原矩阵？(Y/N)", "覆盖选项", Type=2, Default="Y")
                    if overwrite and overwrite.upper() == 'Y':
                        labels = []
                        if len(vals) > 2 and isinstance(vals[2], (list, tuple)) and len(vals[2]) > 1:
                            labels = [vals[2][j] for j in range(1, len(vals[2]))]
                        elif len(vals) > 0 and len(vals[0]) > 1:
                            labels = [vals[0][j] for j in range(1, len(vals[0]))]
                        _write_lower_tri_matrix(range_obj, adj, labels, matrix_name=copula_name)
                        xlcAlert(f"Copula {copula_name} 的矩阵已调整为半正定。")
                    else:
                        new_name = f"{copula_name}_adj"
                        sheet = range_obj.Worksheet
                        start_row = range_obj.Row + range_obj.Rows.Count + 2
                        start_col = range_obj.Column
                        end_row = start_row + range_obj.Rows.Count - 1
                        end_col = start_col + range_obj.Columns.Count - 1
                        new_range = sheet.Range(sheet.Cells(start_row, start_col), sheet.Cells(end_row, end_col))
                        labels = None
                        if len(vals) > 2 and isinstance(vals[2], (list, tuple)) and len(vals[2]) > 1:
                            labels = [vals[2][j] for j in range(1, len(vals[2]))]
                        elif len(vals) > 0 and len(vals[0]) > 1:
                            labels = [vals[0][j] for j in range(1, len(vals[0]))]
                        _write_lower_tri_matrix(new_range, adj, labels, matrix_name=new_name)
                        _set_named_range(app, new_name, new_range)
                        xlcAlert(f"已创建新 Copula {new_name} 并调整为半正定。")
                elif choice == 2:
                    weight_top_left = app.InputBox("请选择权重矩阵的左上角单元格", "权重矩阵位置", Type=8)
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
                    labels = None
                    if len(vals) > 2 and isinstance(vals[2], (list, tuple)) and len(vals[2]) > 1:
                        labels = [vals[2][j] for j in range(1, len(vals[2]))]
                    elif len(vals) > 0 and len(vals[0]) > 1:
                        labels = [vals[0][j] for j in range(1, len(vals[0]))]
                    _write_lower_tri_matrix(weight_range, weight_mat, labels, matrix_name="权重矩阵")
                    weight_name = f"WeightMat_{copula_name}"
                    _set_named_range(app, weight_name, weight_range)
                    _update_copula_weight_mapping(app, weight_name, copula_name)
                    xlcAlert(f"已创建权重矩阵 {weight_name}，请手动修改权重值（0-100），然后运行 DriskWeight 进行调整。")
                else:
                    xlcAlert("未选择任何操作。")
        elif copula_type == "Empirical":
            # 检查数据源是否有效
            if not data_range:
                xlcAlert(f"Empirical Copula {copula_name} 缺少数据源信息。")
                return
            try:
                hist_data = _read_data_range(app, data_range)
                # 验证列数与分布变量数一致（从矩阵区域获取维度？由于矩阵只有3行，无法获取变量数，但可以尝试从分布单元格推断，这里仅验证数据可读）
                xlcAlert(f"Empirical Copula {copula_name} 数据源有效，数据形状 {hist_data.shape}，插值模式 = {interpolate}")
            except Exception as e:
                xlcAlert(f"Empirical Copula {copula_name} 数据源无效：{str(e)}")
        else:
            param_val = float(param) if param is not None else None
            valid = True
            if copula_type in PARAM_THETA_TYPES:
                if param_val is None or param_val <= 0:
                    valid = False
            if valid:
                xlcAlert(f"Copula {copula_name} 参数有效 (θ={param_val})")
            else:
                choice = _show_choice_dialog(
                    title="DriskCheckCopula - 参数调整",
                    message=f"Copula {copula_name} 的参数无效（θ={param_val}）。请选择处理方式：",
                    option1="自动调整（设为默认值 2）",
                    option2="手动输入新参数"
                )
                if choice == 1:
                    new_param = 2.0
                    sheet = range_obj.Worksheet
                    cell = sheet.Cells(range_obj.Row + 1, range_obj.Column + 2)
                    cell.Value = new_param
                    xlcAlert(f"已将参数调整为 {new_param}")
                elif choice == 2:
                    try:
                        new_param_in = app.InputBox("请输入新的 θ 值 ( >0 ):", "输入参数", Type=1, Default="2")
                        if new_param_in is False:
                            return
                        new_param = float(new_param_in)
                        if new_param <= 0:
                            xlcAlert("θ 必须为正数")
                            return
                        sheet = range_obj.Worksheet
                        cell = sheet.Cells(range_obj.Row + 1, range_obj.Column + 2)
                        cell.Value = new_param
                        xlcAlert(f"已将参数更新为 {new_param}")
                    except:
                        xlcAlert("输入无效")
                else:
                    xlcAlert("未选择任何操作。")

    except Exception as e:
        xlcAlert(f"执行 DriskCheckCopula 时出错：{str(e)}")
        traceback.print_exc()

@xl_macro
def DriskRemoveCopula():
    """彻底移除一个 Copula 矩阵及其权重矩阵，并从所有分布函数中移除 DriskCopula 属性"""
    app = _get_excel_app()
    try:
        wb = app.ActiveWorkbook
        # 获取所有 copula 命名区域
        copula_names = []
        for name_obj in wb.Names:
            name = name_obj.Name
            if name.startswith('Copula_') or name.startswith('copula'):
                copula_names.append(name)
        if not copula_names:
            xlcAlert("未找到任何 Copula。")
            return

        copula_name = None
        if len(copula_names) == 1:
            copula_name = copula_names[0]
        else:
            prompt_lines = ["请选择要移除的 Copula（输入数字）:"]
            for i, name in enumerate(copula_names, 1):
                prompt_lines.append(f"{i}. {name}")
            prompt = "\n".join(prompt_lines)
            try:
                choice = app.InputBox(prompt, "DriskRemoveCopula - 选择 Copula", Type=1, Default="1")
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

        # 1. 从公式中移除 DriskCopula
        _remove_copula_from_cells(app, copula_name)

        # 2. 删除命名区域及数据
        range_obj = _get_named_range(app, copula_name)
        if range_obj is not None:
            range_obj.ClearContents()
            try:
                wb.Names(copula_name).Delete()
            except:
                pass

        # 3. 删除关联的权重矩阵
        mapping = _get_mapping_table(app)
        weight_name = None
        for w_name, c_name in mapping.items():
            if c_name == copula_name:
                weight_name = w_name
                break
        if weight_name:
            weight_range = _get_named_range(app, weight_name)
            if weight_range is not None:
                weight_range.ClearContents()
                try:
                    wb.Names(weight_name).Delete()
                except:
                    pass
            del mapping[weight_name]
            _save_mapping_table(app, mapping)

        # 4. 从静音列表中移除
        _remove_muted_copula(app, copula_name)

        xlcAlert(f"Copula“{copula_name}”及其关联的权重矩阵已彻底移除，所有分布函数中的 DriskCopula 已清理。")
    except Exception as e:
        xlcAlert(f"执行 DriskRemoveCopula 时出错：{str(e)}")
        traceback.print_exc()

@xl_macro
def DriskRemoveCopulaAll():
    """移除所有 Copula 矩阵和权重矩阵，并清理所有 DriskCopula 属性"""
    app = _get_excel_app()
    try:
        wb = app.ActiveWorkbook
        copula_names = []
        for name_obj in wb.Names:
            name = name_obj.Name
            if name.startswith('Copula_') or name.startswith('copula'):
                copula_names.append(name)
        if not copula_names:
            xlcAlert("未找到任何 Copula。")
            return

        # 1. 移除所有 DriskCopula
        _remove_all_copula_from_cells(app, copula_names)

        # 2. 删除所有矩阵和权重
        mapping = _get_mapping_table(app)
        for copula_name in copula_names:
            range_obj = _get_named_range(app, copula_name)
            if range_obj is not None:
                range_obj.ClearContents()
                try:
                    wb.Names(copula_name).Delete()
                except:
                    pass
            # 删除关联权重
            weight_name = None
            for w_name, c_name in mapping.items():
                if c_name == copula_name:
                    weight_name = w_name
                    break
            if weight_name:
                weight_range = _get_named_range(app, weight_name)
                if weight_range is not None:
                    weight_range.ClearContents()
                    try:
                        wb.Names(weight_name).Delete()
                    except:
                        pass
                del mapping[weight_name]

        # 清空映射表
        mapping.clear()
        _save_mapping_table(app, mapping)
        # 清空静音列表
        _save_muted_copulas(app, [])

        xlcAlert(f"已移除所有 Copula（共 {len(copula_names)} 个）及其关联的权重矩阵，且所有分布函数中的 DriskCopula 已清理。")
    except Exception as e:
        xlcAlert(f"执行 DriskRemoveCopulaAll 时出错：{str(e)}")
        traceback.print_exc()

@xl_macro
def DriskMuteCopula():
    """切换 Copula 的静音状态（静音时模拟不使用该 Copula）"""
    app = _get_excel_app()
    try:
        wb = app.ActiveWorkbook
        copula_names = []
        for name_obj in wb.Names:
            name = name_obj.Name
            if name.startswith('Copula_') or name.startswith('copula'):
                copula_names.append(name)
        if not copula_names:
            xlcAlert("未找到任何 Copula。")
            return

        copula_name = None
        if len(copula_names) == 1:
            copula_name = copula_names[0]
        else:
            prompt_lines = ["请选择要切换静音状态的 Copula（输入数字）:"]
            for i, name in enumerate(copula_names, 1):
                prompt_lines.append(f"{i}. {name}")
            prompt = "\n".join(prompt_lines)
            try:
                choice = app.InputBox(prompt, "DriskMuteCopula - 选择 Copula", Type=1, Default="1")
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

        if _is_copula_muted(app, copula_name):
            _remove_muted_copula(app, copula_name)
            status = "取消静音"
        else:
            _add_muted_copula(app, copula_name)
            status = "静音"

        xlcAlert(f"Copula“{copula_name}”已{status}。")
    except Exception as e:
        xlcAlert(f"执行 DriskMuteCopula 时出错：{str(e)}")
        traceback.print_exc()