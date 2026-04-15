# ui_stats.py
"""
【UI 统计数据展示层核心模块】

模块定位：
本模块专门负责统计结果在 UI 表格（QTableWidget）中的渲染与展示，不参与任何底层统计学计算。
所有的计算逻辑已剥离至 ui_workers.py 或契约层处理。

核心机制与特性：
1. 智能数据适配：自动识别并适配“建模模式 (Modeler)”与“分析模式 (Results)”的数据差异。
2. 动态空行过滤（隐身机制）：当某个统计指标（如“模拟次数”或“90%CI”）在当前数据字典中完全缺失时，
   渲染引擎会自动隐藏该数据行及其关联的分隔线，确保前端 UI 始终保持高信噪比和极简纯净。
3. 视效增强渲染：提供自定义代理 (Delegate) 以支持特定单元格的高亮，支持高级悬浮滚动条，
   并包含一套“按列分组对齐小数点”的视觉优化算法。
"""

import math
import re
from typing import Any, Dict, List, Optional, Sequence, Tuple

import numpy as np
from PySide6.QtCore import Qt
from PySide6.QtGui import QColor, QFont, QFontMetrics
from PySide6.QtWidgets import QHeaderView, QStyledItemDelegate, QTableWidget, QTableWidgetItem
import ui_stats_contract as stats_contract

# =================================================================
#  Part 1: 通用常量与配置定义 (Constants & Schema Definitions)
# =================================================================

# ----------------- 数据层 Schema 配置 -----------------
# 使用契约层定义的默认百分位数集合（如 1%, 5%, 95%, 99% 等）
DEFAULT_PERCENTILES = list(stats_contract.DEFAULT_PERCENTILES)

# 统计指标 Schema 映射表：定义前端 UI 显示名称与后端数据字典 Key 的对应关系。
# 格式：(UI 显示名称, [可能的后端 Key 列表（按优先级排列）])
STATS_SCHEMA_MAPPING = [
    ("最小值", [stats_contract.MIN_KEY]),
    ("最大值", [stats_contract.MAX_KEY]),
    ("均值", [stats_contract.MEAN_KEY]),
    ("90%CI", [stats_contract.CI90_KEY]),     # 置信区间（仅在实证抽样模式下存在）
    ("中位数", [stats_contract.MEDIAN_KEY]),
    ("标准差", [stats_contract.STD_KEY]),
    ("偏度", [stats_contract.SKEW_KEY]),
    ("峰度", [stats_contract.KURT_KEY]),
]

# 动态扩展 Schema：根据默认百分位数集合，自动生成百分位数的映射关系
for p in DEFAULT_PERCENTILES:
    p_val = float(p)
    label = stats_contract.percentile_display_label(p_val)
    STATS_SCHEMA_MAPPING.append((label, [stats_contract.canonical_percentile_key(p_val)]))

# 定义用于在 UI 中渲染空白分隔行的特殊标记对象
SEP_ROW_MARKER = ("SEQ_LINE", [], True)

# ----------------- 视图层 UI 样式配置 -----------------
COLOR_GRID = "#e0e0e0"       # 表格网格线颜色（浅灰）
COLOR_TEXT_MAIN = "#333333"  # 主体文本颜色（深灰黑）
ROW_HEIGHT_SEP = 2           # 分隔线行的高度（像素）
ROW_HEIGHT_NORMAL = 18       # 常规数据行的高度（像素）
HEADER_HEIGHT = 24           # 表头高度（像素）
FONT_FAMILY = "'Arial','Microsoft YaHei','Segoe UI'" # 字体回退栈
CELL_PAD_V = 1               # 单元格垂直内边距
CELL_PAD_H = 3               # 单元格水平内边距


# =================================================================
#  Part 2: 核心数据处理工具 (Core Utility Functions)
# =================================================================

def normalize_cell_label(label: str) -> str:
    """
    清理并标准化单元格标签名称。
    去除标签中附加的括号说明（例如将 "均值(Mean)" 清理为 "均值"）。
    """
    if not label:
        return ""
    return label.split("(")[0].strip() if "(" in label else label.strip()

def smart_format_number(x: Any, sig: int = 7) -> str:
    """
    智能数字格式化引擎：
    负责将后端传入的各种复杂数据类型转换为适合前端展示的友好字符串。
    
    处理逻辑流：
    1. 范围型数组处理：将长度为2的数组格式化为 [下限, 上限]。
    2. 字符串正则解析：提取字符串形态的数组并递归格式化。
    3. 特殊浮点数拦截：处理无穷大 (Inf)、非数字 (NaN) 及 0 值。
    4. 科学计数法转换：针对极大值 (>=1e8) 或极小值 (<1e-4) 自动启用科学计数法。
    5. 精度动态修剪：根据有效数字计算合理的小数位数，并移除末尾多余的零。
    """
    # 步骤 1：处理范围型数组输入 (Tuple/List/NDArray)
    if isinstance(x, (tuple, list, np.ndarray)) and len(x) == 2:
        lo, hi = x[0], x[1]
        return f"[{smart_format_number(lo, sig)}, {smart_format_number(hi, sig)}]"

    # 步骤 2：解析由其他模块传入的字符串形式的数组范围 (例如 "[1.23, 4.56]")
    if isinstance(x, str):
        s = x.strip()
        # 匹配带有可选科学计数法的浮点数区间
        m = re.match(
            r"^[\[(]\s*([-+]?\d*\.?\d+(?:[eE][-+]?\d+)?)\s*,\s*([-+]?\d*\.?\d+(?:[eE][-+]?\d+)?)\s*[\])]\s*$",
            s
        )
        if m:
            lo = float(m.group(1))
            hi = float(m.group(2))
            return f"[{smart_format_number(lo, sig)}, {smart_format_number(hi, sig)}]"
        try:
            x = float(s)
        except ValueError:
            return s # 若无法转换为浮点数，保留原始字符串返回

    # 步骤 3：尝试转化为标准浮点数并处理极限值
    try:
        xf = float(x)
    except (TypeError, ValueError):
        return str(x)

    if np.isposinf(xf):
        return "∞"
    if np.isneginf(xf):
        return "-∞"
    if np.isnan(xf):
        return ""
    if xf == 0:
        return "0"

    ax = abs(xf)
    # 步骤 4：科学计数法拦截（触发阈值：大于等于1亿，或小于万分之一）
    if ax >= 1e8 or ax < 1e-4:
        return f"{xf:.4e}"  # 使用 .4e (保留四位小数的科学计数法) 兼顾精度与 UI 空间

    # 步骤 5：基于有效数字 (sig) 动态计算并修剪小数位数
    exp = int(math.floor(math.log10(ax)))
    decimals = max(sig - 1 - exp, 0)

    s = f"{xf:,.{decimals}f}" # 使用千分位逗号分隔符
    if "." in s:
        s = s.rstrip("0").rstrip(".") # 剔除小数点后无意义的 0
    return s

def _safe_extract(d: dict, possible_keys: list) -> Any:
    """
    安全的数据提取器：
    遍历 possible_keys 列表尝试从字典中获取有效值。
    遇到 NaN 值会自动跳过，避免将无效数据污染至 UI。
    """
    for k in possible_keys:
        if k in d:
            val = d[k]
            # 过滤 numpy 的 NaN 类型
            if isinstance(val, (float, np.floating)) and np.isnan(val):
                continue
            return val
    return ""

def _normalize_stats_maps(stats_map_list: Sequence[Dict[Any, Any]]) -> List[Dict[str, Any]]:
    """
    数据字典预处理：
    将来自不同业务场景、格式不一的数据字典列表，统一归一化为契约层定义的标准键名空间 (Canonical key-space)。
    """
    normalized: List[Dict[str, Any]] = []
    for stats_map in stats_map_list:
        normalized.append(
            stats_contract.normalize_stats_map(
                stats_map,
                percentiles=DEFAULT_PERCENTILES,
            )
        )
    return normalized


# =================================================================
#  Part 3: 视图代理组件 (View Delegates)
# =================================================================

class StatsDelegate(QStyledItemDelegate):
    """
    统计表格专用渲染代理类：
    主要解决 Qt 表格在部分主题下，默认渲染机制会覆盖或忽略后台设置的背景颜色的问题。
    通过重写 paint 方法，强制优先绘制 Qt.BackgroundRole 指定的背景色。
    """
    def paint(self, painter, option, index):
        bg_brush = index.data(Qt.BackgroundRole)
        if bg_brush:
            painter.save()
            painter.fillRect(option.rect, bg_brush) # 强制填充背景
            painter.restore()
        super().paint(painter, option, index)


# =================================================================
#  Part 4: 数据行装配工厂 (Data Row Assembly Logic)
# =================================================================

def assemble_simulation_rows(
    stats_map_list: Sequence[Dict[Any, Any]],
    col_names: Optional[Sequence[str]] = None,
    cell_labels: Optional[Sequence[str]] = None,
) -> List[Tuple[str, List[Any], bool]]:
    """
    通用型表格数据行装配器（内置空行隐身过滤机制）：
    将后端传递的原始字典列表转化为前端可以直接渲染的 (Label, [Values], IsSeparator) 结构。
    如果某一行的所有列均无实质数据，该行将被彻底丢弃。
    """
    rows: List[Tuple[str, List[Any], bool]] = []
    normalized_stats_map_list = _normalize_stats_maps(stats_map_list)

    def get_vals(possible_keys: list, fmt_func=None) -> List[Any]:
        """内部闭包：批量提取某一指标在所有数据列中的值，并支持传入特定的格式化函数。"""
        vals: List[Any] = []
        for sm in normalized_stats_map_list:
            v = _safe_extract(sm, possible_keys)
            if fmt_func and v != "":
                v = fmt_func(v)
            vals.append(v)
        return vals

    def has_data(vals: List[Any]) -> bool:
        """空行探测：判断当前提取的一行值中是否包含任何非空字符串。"""
        return any(v != "" for v in vals)

    # ---------------- 组装块 1：核心静态指标组 ----------------
    added_static = False
    for label, keys in STATS_SCHEMA_MAPPING[:8]:
        vals = get_vals(keys)
        # 触发隐身机制：若如 "90%CI" 此类指标全体为空，则不添加到 rows 列表
        if has_data(vals):
            rows.append((label, vals, False))
            added_static = True
            
    if added_static:
        rows.append(SEP_ROW_MARKER)

    # ---------------- 组装块 2：计数与异常指标组 ----------------
    def fmt_count(v):
        """内部闭包：专门处理数量型数据的千分位格式化，丢弃小数位。"""
        if isinstance(v, (int, float, np.integer, np.floating)):
            if np.isnan(v):
                return ""
            return f"{int(v):,}"
        return v

    numeric_vals = get_vals([stats_contract.COUNT_KEY], fmt_count)
    error_vals = get_vals([stats_contract.ERROR_COUNT_KEY], fmt_count)
    filtered_vals = get_vals([stats_contract.FILTERED_COUNT_KEY], fmt_count)
    
    if has_data(numeric_vals) or has_data(error_vals) or has_data(filtered_vals):
        rows.append(("数值数", numeric_vals, False))
        rows.append(("错误数", error_vals, False))
        rows.append(("过滤数", filtered_vals, False))
        rows.append(SEP_ROW_MARKER)

    # ---------------- 组装块 3：动态边界与概率指标组 ----------------
    def fmt_p(v):
        """内部闭包：将小数格式化为百分比展示。"""
        return f"{v:.2%}" if isinstance(v, (int, float)) else v
        
    pl_vals = get_vals([stats_contract.DYN_PL_KEY], fmt_p)
    lx_vals = get_vals([stats_contract.DYN_LX_KEY]) 
    pm_vals = get_vals([stats_contract.DYN_PM_KEY], fmt_p)
    rx_vals = get_vals([stats_contract.DYN_RX_KEY]) 
    pr_vals = get_vals([stats_contract.DYN_PR_KEY], fmt_p)

    if has_data(pl_vals) or has_data(lx_vals):
        rows.append(("左侧概率", pl_vals, False))
        rows.append(("左侧X", lx_vals, False))
        rows.append(("中间概率", pm_vals, False))
        rows.append(("右侧X", rx_vals, False))
        rows.append(("右侧概率", pr_vals, False))
        rows.append(SEP_ROW_MARKER)

    # ---------------- 组装块 4：动态百分位数指标组 ----------------
    for label, keys in STATS_SCHEMA_MAPPING[8:]:
        vals = get_vals(keys)
        if has_data(vals):
            rows.append((label, vals, False))

    # ---------------- 清理块：移除末尾多余的分隔线 ----------------
    if rows and rows[-1][2]:
        rows.pop()

    return rows

def assemble_cdf_rows(
    stats_map_list: Sequence[Dict[Any, Any]],
    keys: List[str],
    ci_level: float = 0.95,
    show_ci: bool = False,
    mean_errors: Optional[List[str]] = None
) -> List[Tuple[str, List[Any], bool]]:
    """
    累积分布函数 (CDF) 模式专属表单装配器：
    基于通用装配器构建，但增加了对置信区间 (CI) 表单项以及均值误差项的单独处理逻辑。
    注意：此函数仅处理 UI 层面的数据装配，实际指标计算在 Worker 线程完成。
    """
    # 步骤 1: 调用基础装配器，保证指标计算口径与常规视图一致
    base_rows = assemble_simulation_rows(stats_map_list, keys, keys)
    
    # 步骤 2: 剔除静态的 "90%CI" 行
    # 避免通用模块提取出的固定置信区间与 CDF 模式特有的动态 CI 参数配置框发生视觉与逻辑冲突
    base_rows = [row for row in base_rows if row[0] != "90%CI"]
    
    # 步骤 3: 依据是否开启 CI 渲染决定返回结构
    if not show_ci:
        return base_rows
        
    # 步骤 4: 重组带有动态 CI 配置行和均值误差行的新表单结构
    ci_row = ("置信区间", [f"{ci_level:.2f}"] * len(keys), False)
    final_rows = [ci_row]
    if base_rows:
        final_rows.append(SEP_ROW_MARKER)

    for row in base_rows:
        label, vals, is_sep = row
        final_rows.append(row)

        # 针对特殊需求：将工作线程计算完毕的误差区间 (±误差值) 挂载并紧贴在均值行下方
        if label == "均值" and mean_errors:
            final_rows.append(("", mean_errors, False))

    if final_rows and final_rows[-1][2]:
        final_rows.pop()
    return final_rows


# =================================================================
#  Part 5: 核心表格渲染引擎 (Core Render Engine)
# =================================================================

def render_stats_table(
    table: QTableWidget,
    rows: Sequence[Tuple[str, Sequence[Any], bool]],
    column_headers: Optional[Sequence[str]] = None,
    header_colors: Optional[Any] = None,
    first_col_fill_color: str = "#f2f2f2",
):
    """
    Qt QTableWidget 接管渲染与视觉控制中心：
    负责将装配好的数据流映射至 GUI 组件，同时应用大量的视觉增强算法。
    主要模块：
    1. 小数点对齐预扫描：确保同组指标（如不同列的均值）在视觉上的小数点位数绝对对齐。
    2. 表头与表格初始化：建立网格、注入表头样式及自适应文字颜色。
    3. QSS 样式表注入：绘制现代化的悬浮胶囊式滚动条。
    4. 单元格填充与合并：完成数据的格式化落表，以及特殊空白行的单元格合并 (Span)。
    5. 列宽自适应与交互绑定：监听用户拖拽事件，动态刷新 Tooltip 和自适应列宽。
    """
    if not rows:
        table.setRowCount(0)
        return

    # 尝试保留上一次渲染时用户可能调整过的列宽记录
    prev_col_widths = []
    try:
        prev_col_widths = [int(table.columnWidth(i)) for i in range(int(table.columnCount()))]
    except Exception:
        prev_col_widths = []

    # 探测数据列数量
    n_val_cols = 1
    for r in rows:
        if not r[2]:
            n_val_cols = len(r[1])
            break
    
    # ---------------- 5.1 核心视觉优化：分组对齐预扫描 (Pre-scan for decimal alignment) ----------------
    # 目标：解决传统表格中多列数据由于数量级不同导致小数点参差不齐的问题。
    # 逻辑：预先扫描整个数据矩阵，按业务逻辑分组计算出每列所需保留的最短有效小数位数，并在后续渲染时统一强制应用。
    group_map = {
        "X 均值": "X", "X 标准差": "X", "X 分界线": "X",
        "Y 均值": "Y", "Y 标准差": "Y", "Y 分界线": "Y"
    }
    # 默认需要参与对齐扫描的核心指标
    default_targets = {"最小值", "最大值", "均值", "中位数", "左侧X", "右侧X", "标准差", "90%CI"} 

    # 记录每个分组在每一列上的最少所需小数位数及命中状态
    col_min_decimals = {
        "X": [7] * n_val_cols,
        "Y": [7] * n_val_cols,
        "Default": [7] * n_val_cols
    }
    col_has_target = {
        "X": [False] * n_val_cols,
        "Y": [False] * n_val_cols,
        "Default": [False] * n_val_cols
    }

    # 执行第一遍全表扫描：更新最小小数位数记录字典
    for label, values, is_sep in rows:
        if is_sep:
            continue
            
        grp = group_map.get(label)
        if not grp and (label in default_targets or (isinstance(label, str) and label.endswith("%")) or label == ""):
            grp = "Default"
            
        if grp:
            for c_idx, val in enumerate(values):
                try:
                    # 剔除正负号，解析绝对数值
                    if isinstance(val, str) and val.startswith("±"):
                        xf = float(val.replace("±", ""))
                    else:
                        xf = float(val)
                        
                    if xf == 0 or np.isnan(xf) or np.isinf(xf): continue
                    ax = abs(xf)
                    if ax >= 1e8 or ax < 1e-4: continue 
                        
                    # 结合科学计数法逻辑，计算当前数值合理的保留位数
                    exp = int(math.floor(math.log10(ax)))
                    dec = max(7 - 1 - exp, 0)
                    
                    # 取同列该组内的最小值，以确保显示最粗略的那个指标不超长，而精细指标补零对齐
                    col_min_decimals[grp][c_idx] = min(col_min_decimals[grp][c_idx], dec)
                    col_has_target[grp][c_idx] = True
                except (ValueError, TypeError):
                    pass

    # ---------------- 5.2 表格框架与表头基础配置 (Table & Header Layout Configuration) ----------------
    total_cols = 1 + n_val_cols
    table.setColumnCount(total_cols)
    table.setRowCount(len(rows))
    table.setAlternatingRowColors(False)

    table.horizontalHeader().show()

    def _header_text_color(bg: QColor) -> QColor:
        """视觉优化函数：根据表头背景颜色的灰度亮度，自适应计算使用白色还是黑色的文字，保证高对比度。"""
        r, g, b, _ = bg.getRgb()
        return QColor("#ffffff") if (0.299 * r + 0.587 * g + 0.114 * b) < 140 else QColor("#000000")

    def _get_header_bg(i: int, name: str):
        """安全获取特定列表头的背景颜色配置。"""
        if not header_colors:
            return None
        if isinstance(header_colors, (list, tuple)):
            return header_colors[i] if 0 <= i < len(header_colors) else None
        if isinstance(header_colors, dict):
            return header_colors.get(name)
        return None

    headers = column_headers if (column_headers and len(column_headers) == n_val_cols) else ([""] * n_val_cols)

    # 设定表头统一样式字体
    header_font = QFont()
    header_font.setFamilies(["Arial", "Microsoft YaHei", "Segoe UI"])
    header_font.setBold(True)
    header_font.setPixelSize(12)

    # 初始化第0列列头（指标名称列）
    h0 = QTableWidgetItem("指标")
    h0.setFont(header_font)
    table.setHorizontalHeaderItem(0, h0)

    # 初始化动态数据列的列头
    for i, name in enumerate(headers):
        hi = QTableWidgetItem(str(name))
        hi.setFont(header_font)
        bg_hex = _get_header_bg(i, str(name))
        if bg_hex:
            bg = QColor(bg_hex)
            hi.setData(Qt.BackgroundRole, bg)
            hi.setData(Qt.ForegroundRole, _header_text_color(bg))
        table.setHorizontalHeaderItem(i + 1, hi)

    # 水平与垂直表头行为配置：控制拉伸、点击、隐藏等特性
    h_header = table.horizontalHeader()
    h_header.setFixedHeight(HEADER_HEIGHT)
    h_header.setSectionResizeMode(QHeaderView.Interactive)
    h_header.setStretchLastSection(False)
    h_header.setSectionsMovable(False)
    h_header.setSectionsClickable(True)
    h_header.setTextElideMode(Qt.ElideRight)
    h_header.setSectionResizeMode(0, QHeaderView.Fixed) # 指标列宽度冻结，由后续代码动态计算
    for c_idx in range(1, total_cols):
        h_header.setSectionResizeMode(c_idx, QHeaderView.Interactive)

    v_header = table.verticalHeader()
    v_header.hide()
    v_header.setMinimumSectionSize(1)
    v_header.setSectionResizeMode(QHeaderView.Fixed)
    v_header.setSectionsMovable(False)
    
    table.setWordWrap(False)
    table.setTextElideMode(Qt.ElideRight)
    table.setMouseTracking(True)
    table.setShowGrid(True)
    
    # ---------------- 5.3 QSS 样式配置：高级悬浮胶囊式滚动条 (QSS Styling Integration) ----------------
    table.setStyleSheet(f"""
        QTableWidget {{
            gridline-color: {COLOR_GRID};
            border: none;
            background-color: white;
            font-family: {FONT_FAMILY};
            font-size: 12px;
        }}
        QTableWidget::item {{
            padding: {CELL_PAD_V}px {CELL_PAD_H}px;
            border: none;
        }}
        QTableWidget::item:selected {{
            background-color: #eaf2ff;
            color: #243447;
        }}
        QTableWidget::item:selected:active {{
            background-color: #e5eeff;
            color: #243447;
        }}
        QTableWidget::item:selected:!active {{
            background-color: #eff5ff;
            color: #2e3e52;
        }}
        
        /* ---------------- 高级悬浮胶囊滚动条 ---------------- */
        /* 水平滚动条轨道 */
        QScrollBar:horizontal {{
            border: none;
            background: #f5f5f5;       /* 极浅灰轨道 */
            height: 12px;              /* 稍微细一点更精致 */
            margin: 0px;
            border-radius: 6px;        /* 轨道两端半圆处理 */
        }}
        /* 水平滚动条滑块 */
        QScrollBar::handle:horizontal {{
            background: #bcbcbc;       /* 默认状态颜色 */
            min-width: 30px;
            border-radius: 4px;        /* 滑块圆角 */
            margin: 2px;               /* 核心：四周留白 2px，产生悬浮内嵌感 */
        }}
        QScrollBar::handle:horizontal:hover {{ background: #999999; }}
        QScrollBar::handle:horizontal:pressed {{ background: #777777; }}
        
        /* 隐藏传统两端箭头，符合现代 UI 规范 */
        QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal {{ width: 0px; }}
        QScrollBar::add-page:horizontal, QScrollBar::sub-page:horizontal {{ background: none; }}

        /* 垂直滚动条（逻辑同理） */
        QScrollBar:vertical {{
            border: none;
            background: #f5f5f5;
            width: 12px;
            margin: 0px;
            border-radius: 6px;
        }}
        QScrollBar::handle:vertical {{
            background: #bcbcbc;
            min-height: 30px;
            border-radius: 4px;
            margin: 2px;
        }}
        QScrollBar::handle:vertical:hover {{ background: #999999; }}
        QScrollBar::handle:vertical:pressed {{ background: #777777; }}
        
        QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{ height: 0px; }}
        QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical {{ background: none; }}
    """)

    last_valid_label_row = -1  
    
    # ---------------- 5.4 填充数据单元格并执行对齐逻辑 (Cell Population & Formatting execution) ----------------
    for r_idx, (label, values, is_sep) in enumerate(rows):
        # 场景 A：渲染分隔行
        if is_sep:
            for c in range(total_cols):
                it = QTableWidgetItem("")
                it.setFlags(Qt.NoItemFlags) # 取消交互属性
                if c == 0:
                    it.setBackground(QColor(first_col_fill_color))
                table.setItem(r_idx, c, it)
            table.setRowHeight(r_idx, ROW_HEIGHT_SEP)
            last_valid_label_row = -1  
            continue

        # 场景 B：渲染标准数据行 - 首列 (标签项) 设定
        item_lbl = QTableWidgetItem(str(label))
        item_lbl.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
        item_lbl.setTextAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        item_lbl.setForeground(QColor(COLOR_TEXT_MAIN))
        item_lbl.setBackground(QColor(first_col_fill_color))
        table.setItem(r_idx, 0, item_lbl)
        
        # 处理误差项的合并单元格需求 (例如：将附加在均值下方的正负误差空白标签列与上方均值标签进行纵向合并 span)
        if label == "" and last_valid_label_row != -1:
            span_count = r_idx - last_valid_label_row + 1
            table.setSpan(last_valid_label_row, 0, span_count, 1)
        else:
            last_valid_label_row = r_idx

        # 场景 C：渲染标准数据行 - 动态数据列遍历与对齐输出
        for c_idx, val in enumerate(values):
            txt = ""
            # 分支 1：某些极度敏感于精度的统计量，强制锁定保留 4 位小数，不参与预扫描动态对齐
            if label in ("偏度", "峰度", "相关系数 Pearson", "相关系数 Spearman") and isinstance(val, (float, int)):
                txt = f"{float(val):.4f}"
            else:
                # 分支 2：套用预扫描获取的最佳对齐位数
                grp = group_map.get(label)
                if not grp and (label in default_targets or (isinstance(label, str) and label.endswith("%")) or label == ""):
                    grp = "Default"

                if grp and col_has_target[grp][c_idx]:
                    try:
                        is_pm = (label == "90%CI") or (isinstance(val, str) and val.startswith("±"))
                        
                        if isinstance(val, str) and val.startswith("±"):
                            xf = float(val.replace("±", ""))
                        else:
                            xf = float(val)

                        # 对异常值或极端大小值直接回退至 smart_format_number，不强制固定位数
                        if np.isnan(xf) or np.isinf(xf) or xf == 0:
                            if is_pm:
                                txt = val if isinstance(val, str) else f"±{smart_format_number(val)}"
                            else:
                                txt = smart_format_number(val)
                        elif abs(xf) >= 1e8 or abs(xf) < 1e-4:
                            if is_pm:
                                txt = val if isinstance(val, str) else f"±{smart_format_number(val)}"
                            else:
                                txt = smart_format_number(val)
                        else:
                            # 【核心执行层】：使用预先扫描得到的该组最小小数位数 (dec) 统一格式化并对齐
                            dec = col_min_decimals[grp][c_idx]
                            formatted_num = f"{xf:,.{dec}f}"
                            txt = f"±{formatted_num}" if is_pm else formatted_num
                    except (ValueError, TypeError):
                        txt = val if isinstance(val, str) and val.startswith("±") else smart_format_number(val)
                        if label == "90%CI" and txt and not txt.startswith("±"):
                            txt = f"±{txt}"
                else:
                    # 分支 3：未能归入任何对齐组的其他指标，使用默认的智能格式化规则
                    txt = smart_format_number(val)
                    if isinstance(val, str) and val.startswith("±"):
                        txt = val
                    if label == "90%CI" and txt and not txt.startswith("±"):
                        txt = f"±{txt}"

            # 实例化数据单元格对象并置入表格
            item_val = QTableWidgetItem(txt)
            item_val.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
            item_val.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
            item_val.setForeground(QColor(COLOR_TEXT_MAIN))
            item_val.setBackground(QColor("white"))
            table.setItem(r_idx, c_idx + 1, item_val)

        table.setRowHeight(r_idx, ROW_HEIGHT_NORMAL)

    # 注入自定义渲染代理（以保证 Qt.BackgroundRole 在所有状态下正确生效）
    table.setItemDelegate(StatsDelegate(table))

    # ---------------- 5.5 表格自适应尺寸与事件绑定机制 (Dynamic Layout & Interaction Hooks) ----------------
    # 校验当前数据的宽度指纹是否变化，若变化重置用户拖拽列宽的历史记录
    width_profile_key = (total_cols, tuple(str(x) for x in headers))
    if getattr(table, "_stats_width_profile_key", None) != width_profile_key:
        table._stats_width_profile_key = width_profile_key
        table._stats_user_resized_cols = set()

    user_resized_cols = set(getattr(table, "_stats_user_resized_cols", set()) or set())

    def _compute_first_col_width() -> int:
        """基于 QFontMetrics 动态计算首列（指标名称列）能够完整显示文本所需的最佳宽度。"""
        max_w = QFontMetrics(header_font).horizontalAdvance(str(h0.text() or ""))
        for r_idx in range(table.rowCount()):
            item = table.item(r_idx, 0)
            if item is None or not (item.flags() & Qt.ItemIsEnabled):
                continue
            txt = str(item.text() or "")
            if not txt:
                continue
            item_font = item.font() if item.font() is not None else table.font()
            max_w = max(max_w, QFontMetrics(item_font).horizontalAdvance(txt))
        return max(36, min(320, int(max_w + (CELL_PAD_H * 2) + 14)))

    def _compute_numeric_col_fixed_width() -> int:
        """为数值数据列提供一个统一且相对视觉友好的基准固定宽度（以展示8位数字作为参考）。"""
        fm = QFontMetrics(table.font())
        digit_width = int(fm.horizontalAdvance("88888888"))
        return max(56, min(220, int(digit_width + (CELL_PAD_H * 2) + 18)))

    def _refresh_header_tooltips() -> None:
        """Tooltip 自适应刷新逻辑：仅当列宽不足以完全展示列头文本导致发生裁剪时，才提供悬浮提示以补充信息。"""
        for col_idx in range(1, table.columnCount()):
            hi = table.horizontalHeaderItem(col_idx)
            if hi is None:
                continue
            txt = str(hi.text() or "")
            if not txt:
                hi.setToolTip("")
                continue
            hi_font = hi.font() if hi.font() is not None else header_font
            text_w = QFontMetrics(hi_font).horizontalAdvance(txt)
            avail_w = max(0, int(h_header.sectionSize(col_idx)) - 10)
            hi.setToolTip(txt if text_w > avail_w else "")

    table._stats_refresh_header_tooltips = _refresh_header_tooltips

    # 执行底层列宽应用计算：优先使用用户拉伸记录，否则使用基准宽度
    table._stats_internal_resize = True
    try:
        table.setColumnWidth(0, _compute_first_col_width())
        fixed_numeric_w = _compute_numeric_col_fixed_width()
        for col_idx in range(1, total_cols):
            if col_idx in user_resized_cols and col_idx < len(prev_col_widths) and int(prev_col_widths[col_idx]) > 0:
                table.setColumnWidth(col_idx, int(prev_col_widths[col_idx]))
            else:
                table.setColumnWidth(col_idx, fixed_numeric_w)
    finally:
        table._stats_internal_resize = False

    _refresh_header_tooltips()

    # 注册表头的宽度拖拽及双击重置信号槽：确保用户操作可记忆且能联动 Tooltip。
    if not getattr(table, "_stats_resize_hooks_connected", False):
        def _on_section_resized(section: int, _old: int, _new: int, _table: QTableWidget = table):
            try:
                # 记录用户主动操作的列宽，防止下次刷新时被内部逻辑覆写
                if not bool(getattr(_table, "_stats_internal_resize", False)) and int(section) >= 1:
                    resized = set(getattr(_table, "_stats_user_resized_cols", set()) or set())
                    resized.add(int(section))
                    _table._stats_user_resized_cols = resized
                refresh_cb = getattr(_table, "_stats_refresh_header_tooltips", None)
                if callable(refresh_cb):
                    refresh_cb()
            except Exception:
                pass

        def _on_section_handle_double_clicked(section: int, _table: QTableWidget = table):
            try:
                # 双击表头边缘自动恢复基准宽度的功能
                section = int(section)
                if section < 1:
                    return
                _table._stats_internal_resize = True
                try:
                    _table.setColumnWidth(section, _compute_numeric_col_fixed_width())
                finally:
                    _table._stats_internal_resize = False
                resized = set(getattr(_table, "_stats_user_resized_cols", set()) or set())
                resized.add(section)
                _table._stats_user_resized_cols = resized
                refresh_cb = getattr(_table, "_stats_refresh_header_tooltips", None)
                if callable(refresh_cb):
                    refresh_cb()
            except Exception:
                pass

        h_header.sectionResized.connect(_on_section_resized)
        h_header.sectionHandleDoubleClicked.connect(_on_section_handle_double_clicked)
        table._stats_resize_hooks_connected = True
        table._stats_on_section_resized = _on_section_resized
        table._stats_on_section_dblclick = _on_section_handle_double_clicked