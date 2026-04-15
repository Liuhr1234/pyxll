# ui_tornado.py
"""
本模块提供龙卷风图（敏感性分析）及高级情景分析的视图组件与数据计算服务。

主要功能：
1. 依赖关系追踪 (Dependency Tracking)：基于 Excel 公式树反向追踪，过滤与输出无关的输入变量，确保敏感性分析的准确性。
2. 配置设置面板 (Settings Dialogs)：提供情景分析、敏感性分析参数配置的 UI 弹窗与输入合法性校验。
3. 核心计算引擎 (Calculation Engine)：实现分箱统计、多元回归、斯皮尔曼秩相关、方差贡献度及极端情景分析等底层算法。
4. 视图渲染容器 (View Rendering Container)：基于 PlotlyHost 封装，负责各类敏感性与情景分析图表的渲染调度、配置应用与绘制逻辑。
"""

import math
import re
import numpy as np
import pandas as pd
import plotly.graph_objects as go
from scipy import stats
import statsmodels.api as sm
from scipy.stats import rankdata, spearmanr
from typing import List, Dict

from PySide6.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QLabel, QFrame,QSpinBox, QDoubleSpinBox, QComboBox, QPushButton,
                               QDialog, QGridLayout, QDialogButtonBox, QTableWidget, QTableWidgetItem, QHeaderView, QMessageBox, QLineEdit)
from PySide6.QtCore import Qt
from PySide6.QtGui import QIntValidator, QDoubleValidator
from PySide6.QtWebEngineWidgets import QWebEngineView

from plotly_host import PlotlyHost
from ui_shared import DriskMath, drisk_spinbox_qss, drisk_combobox_qss, DRISK_COLOR_CYCLE, DRISK_DIALOG_BTN_QSS, get_si_unit_config
from drisk_charting import DriskChartFactory  # 引入以复用透明度与十六进制颜色融合工具


# =======================================================
# 1. 颜色与主题常量 (Theme & Style Constants)
# =======================================================
class ThemeColors:
    """定义图表使用的基础主题颜色"""
    BASELINE_COLOR = "#2c3e50"
    GRID_COLOR = "#e0e0e0"

# 龙卷风图通用字体配置
TORNADO_TITLE_FONT = dict(size=14, family="Arial, 'Microsoft YaHei', sans-serif", color="#333333")
TORNADO_AXIS_FONT = dict(size=12, family="Arial, 'Microsoft YaHei', sans-serif", color="#333333")


# =======================================================
# 2. 依赖关系追踪服务 (Dependency Tracking Service)
# =======================================================
# 注意(交接提示)：当前文件中存在两个 filter_inputs_by_dependency 函数的定义。
# 此处为首个简略版实现，在 Python 运行时会被紧接着的第二个完整版定义所覆盖。
# 保留此段落是为了维持源代码的一致性，建议后续重构时清理该冗余代码。
def filter_inputs_by_dependency(
    xl_app,
    output_cell_key: str,
    input_mapping: Dict[str, str],
    *,
    stop_at_makeinput: bool = False,
) -> List[str]:
    """
    [核心辅助函数 - 简略版] 基于 Excel 公式依赖树过滤输入变量
    反向追踪并剔除与输出变量在数学公式上无关的输入变量，避免生成无意义的敏感性分析。
    
    参数:
    :param xl_app: Excel COM 实例对象
    :param output_cell_key: 真实的底层输出单元格地址 (如 'Sheet1!C10')
    :param input_mapping: 输入变量映射字典 {前端展示列名: 真实的底层输入单元格 (如 'Sheet1!A1')}
    
    返回:
    :return: 经过依赖过滤后的前端展示列名列表
    """
    if not xl_app or not output_cell_key:
        return list(input_mapping.keys()) # 兜底策略：若 COM 接口不可用，则跳过过滤直接全量返回

    try:
        from formula_parser import parse_formula_references_tornado
    except ImportError:
        return list(input_mapping.keys())

    visited = set()
    queue = [output_cell_key.upper()]
    visited.add(output_cell_key.upper())

    # 构建反向查找字典：{大写单元格地址: 前端展示列名}
    target_inputs = {}
    for col_name, raw_key in input_mapping.items():
        # 处理可能的模拟引擎后缀 (如将 'Sheet1!A1_1' 还原为 'Sheet1!A1')
        pure_key = raw_key.split('_')[0] if '_' in raw_key else raw_key 
        target_inputs[pure_key.upper()] = col_name

    valid_columns = set()

    # 广度优先搜索 (BFS) 遍历公式依赖树
    while queue:
        current_node = queue.pop(0)

        # 1. 如果当前节点是受监控的输入变量，记录并保留它
        if current_node in target_inputs:
            valid_columns.add(target_inputs[current_node])

        # 2. 解析当前节点的公式寻找上游依赖
        try:
            if '!' in current_node:
                sheet_name, cell_addr = current_node.split('!')
                ws = xl_app.ActiveWorkbook.Worksheets(sheet_name)
                cell = ws.Range(cell_addr)
            else:
                ws = xl_app.ActiveSheet
                cell = ws.Range(current_node)

            formula = cell.Formula
            # 如果是静态值或非公式（不以等号开头），停止该分支的溯源
            if not isinstance(formula, str) or not formula.startswith('='):
                continue
            
            # 当调用方（如敏感性/情景分析模块）要求时，将 DriskMakeInput 节点视为依赖树追踪的边界
            if stop_at_makeinput and "DRISKMAKEINPUT" in formula.upper():
                continue

            # 3. 提取公式中的上游引用单元格并压入队列
            refs = parse_formula_references_tornado(formula)
            for ref in refs:
                ref_upper = ref.upper()
                # 若引用未指定工作表，则自动补全为当前工作表名
                if '!' not in ref_upper:
                    ref_upper = f"{ws.Name.upper()}!{ref_upper}"

                if ref_upper not in visited:
                    visited.add(ref_upper)
                    queue.append(ref_upper)

        except Exception:
            # 忽略受保护/合并的无法读取的单元格
            continue 

    # 如果追踪成功且找到了关联变量，则返回交集；
    # 否则为防止由于跨工作簿引用导致追踪断裂，触发兜底全量返回。
    return list(valid_columns) if valid_columns else list(input_mapping.keys())


# -------------------------------------------------------
# 依赖关系追踪服务 (完整版实现)
# -------------------------------------------------------
def filter_inputs_by_dependency(
    xl_app,
    output_cell_key: str,
    input_mapping: Dict[str, str],
    *,
    stop_at_makeinput: bool = False,
) -> List[str]:
    """
    [核心辅助函数 - 完整版] 通过从输出端反向追踪 Excel 公式依赖，过滤有效的输入列。

    遍历过程保留了完整的单元格标识（包括工作表名和地址），
    并维护了由公式和名称范围引用的所有依赖分支。
    """
    if not xl_app or not output_cell_key:
        return list(input_mapping.keys())

    try:
        from formula_parser import parse_formula_references_tornado
    except ImportError:
        return list(input_mapping.keys())

    def _normalize_sheet_name(sheet_text: str) -> str:
        """规范化工作表名称，去除多余的引号包裹"""
        name = str(sheet_text).strip()
        if name.startswith("'") and name.endswith("'"):
            name = name[1:-1].replace("''", "'")
        return name

    def _normalize_cell_key(raw_key: str, default_sheet: str = None):
        """规范化单元格键名，统一去除绝对引用符号并处理模拟引擎附加后缀"""
        if raw_key is None:
            return None
        text = str(raw_key).strip().replace("$", "")
        if not text:
            return None

        sheet_name = None
        addr_text = text
        if "!" in text:
            sheet_part, addr_part = text.rsplit("!", 1)
            sheet_name = _normalize_sheet_name(sheet_part)
            addr_text = addr_part
        elif default_sheet:
            sheet_name = _normalize_sheet_name(default_sheet)

        addr_text = addr_text.strip().upper()
        # 规范化模拟引擎的单元格后缀:
        # 如 A1_1 / A1_MAKEINPUT / A1_NESTED_1 / A1_AT_1 -> 统一映射为 A1
        # 这有助于保持追踪到的 A1 样式引用匹配稳定性。
        m = re.fullmatch(r"([A-Z]{1,7}\d+)(?:_.+)?", addr_text)
        if m:
            addr_text = m.group(1)

        if sheet_name:
            return f"{sheet_name.upper()}!{addr_text}"
        return addr_text

    def _collect_precedent_refs_from_com(cell_obj) -> List[str]:
        """
        使用 Excel COM 的前导单元格依赖图 (Precedent graph) 
        解析在公式文本中不直接可见的依赖关系（例如自定义名称定义）。
        """
        refs: List[str] = []
        precedents = None

        try:
            precedents = cell_obj.DirectPrecedents
        except Exception:
            try:
                precedents = cell_obj.Precedents
            except Exception:
                precedents = None

        if precedents is None:
            return refs

        areas = []
        try:
            for area in precedents.Areas:
                areas.append(area)
        except Exception:
            areas = [precedents]

        for area in areas:
            try:
                cells_iter = area.Cells
            except Exception:
                cells_iter = [area]

            for dep_cell in cells_iter:
                try:
                    dep_sheet = _normalize_sheet_name(dep_cell.Worksheet.Name)
                    dep_addr = str(dep_cell.Address).replace("$", "")
                    dep_key = _normalize_cell_key(f"{dep_sheet}!{dep_addr}")
                    if dep_key:
                        refs.append(dep_key)
                except Exception:
                    continue

        # 保留发现顺序并去重
        deduped: List[str] = []
        seen = set()
        for item in refs:
            key = str(item).upper()
            if key in seen:
                continue
            seen.add(key)
            deduped.append(item)
        return deduped

    def _mask_makeinput_calls(expr: str) -> str:
        """
        将 DriskMakeInput(...) 替换为空格，以便在启用 stop_at_makeinput 时，
        提取标记 (token) 的逻辑可以忽略内部的分支引用。
        """
        if not expr:
            return expr

        token = "DRISKMAKEINPUT"
        upper_expr = expr.upper()
        masked_chars = list(expr)
        search_pos = 0

        while True:
            idx = upper_expr.find(token, search_pos)
            if idx < 0:
                break

            prev = upper_expr[idx - 1] if idx > 0 else ""
            if prev and (prev.isalnum() or prev == "_"):
                search_pos = idx + len(token)
                continue

            cursor = idx + len(token)
            while cursor < len(expr) and expr[cursor].isspace():
                cursor += 1
            if cursor >= len(expr) or expr[cursor] != "(":
                search_pos = idx + len(token)
                continue

            depth = 0
            quote_char = None
            end = cursor
            while end < len(expr):
                ch = expr[end]
                if quote_char:
                    if ch == quote_char:
                        if end + 1 < len(expr) and expr[end + 1] == quote_char:
                            end += 1
                        else:
                            quote_char = None
                else:
                    if ch in ("'", '"'):
                        quote_char = ch
                    elif ch == "(":
                        depth += 1
                    elif ch == ")":
                        depth -= 1
                        if depth == 0:
                            end += 1
                            break
                end += 1

            if end <= idx:
                search_pos = idx + len(token)
                continue

            for i in range(idx, min(end, len(masked_chars))):
                if masked_chars[i] not in ("\r", "\n", "\t"):
                    masked_chars[i] = " "
            search_pos = end

        return "".join(masked_chars)

    def _extract_defined_name_tokens(formula_text: str, *, exclude_makeinput_inner: bool) -> List[str]:
        """从公式字符串中提取可能是自定义名称的词法标记"""
        if not isinstance(formula_text, str) or not formula_text.startswith("="):
            return []

        body = formula_text[1:]
        if exclude_makeinput_inner:
            body = _mask_makeinput_calls(body)

        # 屏蔽字符串字面量和带引号的工作表名称，避免产生错误的匹配
        body = re.sub(r'"(?:[^"]|"")*"', " ", body)
        body = re.sub(r"'(?:[^']|'')*'", " ", body)

        tokens: List[str] = []
        seen = set()
        for m in re.finditer(r"[A-Za-z_\\][A-Za-z0-9_.]*", body):
            token = m.group(0)
            token_upper = token.upper()

            # 忽略布尔值和类似单元格坐标的引用
            if token_upper in {"TRUE", "FALSE"}:
                continue
            if re.fullmatch(r"[A-Z]{1,7}\d+", token_upper):
                continue

            # 忽略明显是点号命名空间一部分的标识符
            start, end = m.span()
            if start > 0 and body[start - 1] == ".":
                continue
            if end < len(body) and body[end:end + 1] == ".":
                continue

            # 忽略函数标识符
            j = end
            while j < len(body) and body[j].isspace():
                j += 1
            if j < len(body) and body[j] == "(":
                continue

            if token_upper in seen:
                continue
            seen.add(token_upper)
            tokens.append(token)
        return tokens

    def _resolve_defined_name_refs(ws, name_token: str) -> List[str]:
        """通过 Excel 自定义名称 (Defined Name) 解析实际依赖的单元格地址"""
        refs: List[str] = []
        workbook = getattr(xl_app, "ActiveWorkbook", None)
        if workbook is None:
            return refs

        name_objs = []
        seen_names = set()

        def _add_name_obj(obj):
            if obj is None:
                return
            try:
                obj_name = str(getattr(obj, "Name", "") or "").upper()
            except Exception:
                obj_name = ""
            key = obj_name or str(id(obj))
            if key in seen_names:
                return
            seen_names.add(key)
            name_objs.append(obj)

        for scoped in (name_token, f"{ws.Name}!{name_token}", f"'{ws.Name}'!{name_token}"):
            try:
                _add_name_obj(workbook.Names(scoped))
            except Exception:
                continue

        try:
            for n in ws.Names:
                try:
                    n_name = str(getattr(n, "Name", "") or "")
                except Exception:
                    n_name = ""
                if n_name.split("!")[-1].upper() == str(name_token).upper():
                    _add_name_obj(n)
        except Exception:
            pass

        for name_obj in name_objs:
            # 首选路径：解析范围引用 (适用于命名范围)
            try:
                rng = name_obj.RefersToRange
                for dep_cell in rng.Cells:
                    try:
                        dep_sheet = _normalize_sheet_name(dep_cell.Worksheet.Name)
                        dep_addr = str(dep_cell.Address).replace("$", "")
                        dep_key = _normalize_cell_key(f"{dep_sheet}!{dep_addr}")
                        if dep_key:
                            refs.append(dep_key)
                    except Exception:
                        continue
                continue
            except Exception:
                pass

            # 后备路径：当不存在直接范围时，解析 RefersTo 公式文本
            try:
                refers_to = str(getattr(name_obj, "RefersTo", "") or "").strip()
            except Exception:
                refers_to = ""
            if not refers_to:
                continue
            if not refers_to.startswith("="):
                refers_to = f"={refers_to}"

            for ref in parse_formula_references_tornado(refers_to):
                dep_key = _normalize_cell_key(ref, default_sheet=ws.Name)
                if dep_key:
                    refs.append(dep_key)

        deduped: List[str] = []
        seen = set()
        for item in refs:
            key = str(item).upper()
            if key in seen:
                continue
            seen.add(key)
            deduped.append(item)
        return deduped

    # 依赖追踪主循环初始化
    start_node = _normalize_cell_key(output_cell_key)
    if not start_node:
        return list(input_mapping.keys())

    visited = {start_node}
    queue = [start_node]

    # 构建反向查找字典：规范化输入键 -> 一个或多个展示列名
    target_inputs: Dict[str, set] = {}
    for col_name, raw_key in input_mapping.items():
        canonical_key = _normalize_cell_key(raw_key)
        if not canonical_key:
            continue
        target_inputs.setdefault(canonical_key, set()).add(col_name)
        # 保留不带工作表名称的本地备用键，以兼容调用方传入纯地址的场景
        if "!" in canonical_key:
            _, addr_only = canonical_key.split("!", 1)
            target_inputs.setdefault(addr_only, set()).add(col_name)

    valid_columns = set()
    while queue:
        current_node = queue.pop(0)

        if current_node in target_inputs:
            valid_columns.update(target_inputs[current_node])

        try:
            if "!" in current_node:
                sheet_name, cell_addr = current_node.split("!", 1)
                ws = xl_app.ActiveWorkbook.Worksheets(_normalize_sheet_name(sheet_name))
                cell = ws.Range(cell_addr)
            else:
                ws = xl_app.ActiveSheet
                cell = ws.Range(current_node)

            formula = cell.Formula
            if not isinstance(formula, str) or not formula.startswith("="):
                continue
            formula_upper = formula.upper()
            refs = parse_formula_references_tornado(
                formula,
                exclude_makeinput_inner=bool(stop_at_makeinput),
            )
            # 通过 COM 解析定义名称的前导依赖，使得工作簿层级的命名引用
            # (例如 Provisions -> 'Risk factors'!F3) 可以继续追踪至上游输入单元格
            if not (stop_at_makeinput and "DRISKMAKEINPUT" in formula_upper):
                refs.extend(_collect_precedent_refs_from_com(cell))

            # 名称标记后备方案，以保持在 COM 前导不可用或工作簿状态下部分解析时的覆盖率
            name_tokens = _extract_defined_name_tokens(
                formula,
                exclude_makeinput_inner=bool(stop_at_makeinput),
            )
            for token in name_tokens:
                refs.extend(_resolve_defined_name_refs(ws, token))

            for ref in refs:
                normalized_ref = _normalize_cell_key(ref, default_sheet=ws.Name)
                if not normalized_ref or normalized_ref in visited:
                    continue
                visited.add(normalized_ref)
                queue.append(normalized_ref)
        except Exception:
            continue

    # 保持下游渲染稳定性的原始输入列排序
    if valid_columns:
        return [col for col in input_mapping.keys() if col in valid_columns]
    return list(input_mapping.keys())


# =======================================================
# 3. 配置与设置对话框组件 (Settings Dialog Components)
# =======================================================
class ScenarioSettingsDialog(QDialog):
    """
    [UI 组件] 情景分析模式的设置弹窗
    负责收集用户关于情景分析的边界（最小值/最大值）、显示模式及阈值的配置。
    """
    def __init__(
        self,
        current_display_idx=0,
        current_min=0.0,
        current_max=25.0,
        current_threshold=0.5,
        current_display_limit=None,
        parent=None
    ):
        super().__init__(parent)
        self.setWindowTitle("Drisk - 情景设置")
        # 1. 缩小窗体整体长宽，使其更加紧凑
        self.resize(440, 250)
        self.setStyleSheet("""
            QDialog { background-color: #f0f0f0; }
            QLabel { font-size: 12px; color: #333; }
            QComboBox { padding: 4px 6px; border: 1px solid #ccc; background: white; font-weight: normal; }
            QDoubleSpinBox, QLineEdit { padding: 4px 6px; border: 1px solid #ccc; background: white; font-weight: normal; }
        """)

        layout = QVBoxLayout(self)

        # --- 顶部：显示模式选择 ---
        top_layout = QHBoxLayout()
        top_layout.addWidget(QLabel("显示模式:"))
        self.combo_display = QComboBox()
        self.combo_display.addItems([
            "显著性（中位数变化/标准差）",
            "子集中位数百分位",
            "子集中位数实际值"
        ])
        self.combo_display.setCurrentIndex(current_display_idx)
        top_layout.addWidget(self.combo_display, 1)
        layout.addLayout(top_layout)

        # 2. 增加下拉框与下方表格之间的间距
        layout.addSpacing(15)

        # --- 中部：情景边界设置 (简化版表格/表单) ---
        grid_layout = QGridLayout()
        # 增加横向间距，让列与列之间更透气
        grid_layout.setHorizontalSpacing(15)
        
        grid_layout.addWidget(QLabel("情景"), 0, 0)
        grid_layout.addWidget(QLabel("模式"), 0, 1)
        grid_layout.addWidget(QLabel("最小值(%)"), 0, 2)
        grid_layout.addWidget(QLabel("最大值(%)"), 0, 3)

        grid_layout.addWidget(QLabel("当前情景"), 1, 0)
        grid_layout.addWidget(QLabel("百分位区间"), 1, 1)

        # 最小值微调框
        self.spin_min = QDoubleSpinBox()
        self.spin_min.setRange(0, 100)
        self.spin_min.setValue(current_min)
        self.spin_min.setDecimals(2)
        self.spin_min.setButtonSymbols(QDoubleSpinBox.NoButtons) 
        self.spin_min.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        # 3. 强制截断横向长度，防止其被网格布局拉扯
        self.spin_min.setFixedWidth(80)
        
        # 最大值微调框
        self.spin_max = QDoubleSpinBox()
        self.spin_max.setRange(0, 100)
        self.spin_max.setValue(current_max)
        self.spin_max.setDecimals(2)
        self.spin_max.setButtonSymbols(QDoubleSpinBox.NoButtons) 
        self.spin_max.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        # 3. 强制截断横向长度
        self.spin_max.setFixedWidth(80)

        # 挂载到网格时加入靠左对齐参数 (Qt.AlignLeft)，确保它们不会漂移
        grid_layout.addWidget(self.spin_min, 1, 2, Qt.AlignLeft)
        grid_layout.addWidget(self.spin_max, 1, 3, Qt.AlignLeft)

        # 显著性水平阈值
        grid_layout.addWidget(QLabel("显著性水平:"), 2, 0)
        self.edit_significance_threshold = QLineEdit()
        self.edit_significance_threshold.setPlaceholderText("空表示不限制")
        self.edit_significance_threshold.setValidator(QDoubleValidator(0.0, 1e9, 6, self))
        if current_threshold is None:
            self.edit_significance_threshold.setText("")
        else:
            self.edit_significance_threshold.setText(f"{float(current_threshold):g}")
        self.edit_significance_threshold.setFixedWidth(120)
        grid_layout.addWidget(self.edit_significance_threshold, 2, 1, Qt.AlignLeft)

        # 显示上限
        grid_layout.addWidget(QLabel("显示上限:"), 3, 0)
        self.edit_display_limit = QLineEdit()
        self.edit_display_limit.setPlaceholderText("空表示不限制")
        self.edit_display_limit.setValidator(QIntValidator(1, 1000000, self))
        if current_display_limit in (None, ""):
            self.edit_display_limit.setText("")
        else:
            self.edit_display_limit.setText(str(int(current_display_limit)))
        self.edit_display_limit.setFixedWidth(120)
        grid_layout.addWidget(self.edit_display_limit, 3, 1, Qt.AlignLeft)
        layout.addLayout(grid_layout)

        layout.addStretch()

        # --- 底部：按钮区 ---
        self.setStyleSheet(self.styleSheet() + DRISK_DIALOG_BTN_QSS)

        btn_layout = QHBoxLayout()
        btn_layout.addStretch()

        btn_ok = QPushButton("确定")
        btn_ok.setObjectName("btnOk")
        btn_ok.clicked.connect(self._validate_and_accept)
        
        btn_cancel = QPushButton("取消")
        btn_cancel.setObjectName("btnCancel")
        btn_cancel.clicked.connect(self.reject)

        btn_layout.addWidget(btn_ok)     # 确定在左
        btn_layout.addWidget(btn_cancel) # 取消在右
        layout.addLayout(btn_layout)

    def _validate_and_accept(self):
        """[生命周期] 确定按钮触发：执行输入数据的合法性校验"""
        if self.spin_min.value() >= self.spin_max.value():
            QMessageBox.warning(self, "设置错误", "最小值必须小于最大值。")
            return
            
        threshold_text = self.edit_significance_threshold.text().strip()
        if threshold_text:
            try:
                threshold_val = float(threshold_text)
            except Exception:
                QMessageBox.warning(self, "设置错误", "显著性水平必须是非负数字。")
                return
            if threshold_val < 0:
                QMessageBox.warning(self, "设置错误", "显著性水平必须大于等于 0。")
                return
                
        limit_text = self.edit_display_limit.text().strip()
        if limit_text:
            try:
                limit_val = int(limit_text)
            except Exception:
                QMessageBox.warning(self, "设置错误", "显示上限必须是正整数。")
                return
            if limit_val <= 0:
                QMessageBox.warning(self, "设置错误", "显示上限必须大于 0。")
                return
                
        self.accept()

    def get_settings(self):
        """对外接口：返回当前弹窗的设置字典"""
        threshold_text = self.edit_significance_threshold.text().strip()
        limit_text = self.edit_display_limit.text().strip()
        return {
            'display_mode': self.combo_display.currentIndex(),
            'min_pct': self.spin_min.value(),
            'max_pct': self.spin_max.value(),
            'significance_threshold': (float(threshold_text) if threshold_text else None),
            'display_limit': (int(limit_text) if limit_text else None),
        }

class TornadoSettingsDialog(QDialog):
    """
    [UI 组件] 敏感性分析设置弹窗
    负责收集用户关于输入分组个数、显示上限及目标输出变量统计量的配置。
    """
    def __init__(self, current_config=None, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Drisk - 敏感性分析设置")
        self.resize(400, 220)
        self.setStyleSheet("""
            QDialog { background-color: #f0f0f0; }
            QLabel { font-size: 12px; color: #333; }
            QComboBox, QSpinBox, QDoubleSpinBox, QLineEdit { padding: 4px 6px; border: 1px solid #ccc; background: white; font-weight: normal; }
        """)
        
        if current_config is None:
            current_config = {'num_bins': 10, 'max_vars': 16, 'stat_type': 'mean', 'percentile_val': 90.0}

        layout = QVBoxLayout(self)
        grid_layout = QGridLayout()
        grid_layout.setSpacing(15)

        # 1. 分组个数 (Bins)
        grid_layout.addWidget(QLabel("输入变量数据分组:"), 0, 0)
        self.spin_bins = QSpinBox()
        self.spin_bins.setRange(2, 100)
        self.spin_bins.setValue(current_config.get('num_bins', 10))
        self.spin_bins.setButtonSymbols(QSpinBox.NoButtons) # 新增：隐藏上下箭头
        self.spin_bins.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        grid_layout.addWidget(self.spin_bins, 0, 1)

        # 2. 最大展示数限制
        grid_layout.addWidget(QLabel("输入变量展示上限:"), 1, 0)
        self.edit_max_vars = QLineEdit()
        self.edit_max_vars.setValidator(QIntValidator(1, 1000000, self))
        self.edit_max_vars.setPlaceholderText("空表示不限制")
        max_vars_val = current_config.get('max_vars', 16)
        if max_vars_val in (None, ""):
            self.edit_max_vars.setText("")
        else:
            self.edit_max_vars.setText(str(int(max_vars_val)))
        grid_layout.addWidget(self.edit_max_vars, 1, 1)

        # 3. 输出变量的统计量选择
        grid_layout.addWidget(QLabel("输出变量统计量:"), 2, 0)
        self.combo_stat = QComboBox()
        self.combo_stat.addItems(["均值", "中位数", "百分位数"])
        
        stat_map = {'mean': 0, 'median': 1, 'percentile': 2}
        self.combo_stat.setCurrentIndex(stat_map.get(current_config.get('stat_type', 'mean'), 0))
        self.combo_stat.currentIndexChanged.connect(self._toggle_pct)
        grid_layout.addWidget(self.combo_stat, 2, 1)

        # 4. 百分位数值微调 (仅在选择了“百分位数”时显示)
        self.lbl_pct = QLabel("百分位区间(%):")
        self.spin_pct = QDoubleSpinBox()
        self.spin_pct.setRange(0, 100)
        self.spin_pct.setValue(current_config.get('percentile_val', 90.0))
        self.spin_pct.setDecimals(2)
        self.spin_pct.setButtonSymbols(QDoubleSpinBox.NoButtons)
        self.spin_pct.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)

        grid_layout.addWidget(self.lbl_pct, 3, 0)
        grid_layout.addWidget(self.spin_pct, 3, 1)

        layout.addLayout(grid_layout)
        layout.addStretch()

        # 底部按钮区
        self.setStyleSheet(self.styleSheet() + DRISK_DIALOG_BTN_QSS)
        btn_layout = QHBoxLayout()
        btn_layout.addStretch()

        btn_ok = QPushButton("确定")
        btn_ok.setObjectName("btnOk")
        btn_ok.setFixedSize(80, 28)
        btn_ok.clicked.connect(self._validate_and_accept)

        btn_cancel = QPushButton("取消")
        btn_cancel.setObjectName("btnCancel")
        btn_cancel.setFixedSize(80, 28)
        btn_cancel.clicked.connect(self.reject)

        btn_layout.addWidget(btn_ok)
        btn_layout.addWidget(btn_cancel)
        layout.addLayout(btn_layout)

        self._toggle_pct()

    def _toggle_pct(self):
        """控制百分位微调框的动态显示与隐藏"""
        is_pct = (self.combo_stat.currentIndex() == 2)
        self.lbl_pct.setVisible(is_pct)
        self.spin_pct.setVisible(is_pct)

    def _validate_and_accept(self):
        """校验输入展示上限的合法性"""
        limit_text = self.edit_max_vars.text().strip()
        if limit_text:
            try:
                limit_val = int(limit_text)
            except Exception:
                QMessageBox.warning(self, "设置错误", "输入变量展示上限必须是正整数。")
                return
            if limit_val <= 0:
                QMessageBox.warning(self, "设置错误", "输入变量展示上限必须大于 0。")
                return
        self.accept()

    def get_settings(self):
        """对外接口：返回敏感性分析设置字典"""
        stat_map_rev = {0: 'mean', 1: 'median', 2: 'percentile'}
        limit_text = self.edit_max_vars.text().strip()
        return {
            'num_bins': self.spin_bins.value(),
            'max_vars': (int(limit_text) if limit_text else None),
            'stat_type': stat_map_rev[self.combo_stat.currentIndex()],
            'percentile_val': self.spin_pct.value()
        }

# =======================================================
# 4. 核心数据处理与计算引擎 (Core Calculation Engine)
# =======================================================
class TornadoDataProvider:
    """
    [核心业务类] 提供各类敏感性分析与情景分析所需的底层数学/统计计算。
    封装了数据预处理与各项统计指标的生成逻辑。
    """
    def __init__(self, full_df: pd.DataFrame, output_col_name: str, input_col_names: list[str]):
        self.df = full_df
        self.target_col = output_col_name
        self.input_cols = input_col_names
        self.baseline = self.df[self.target_col].mean()
        self.global_min = self.baseline
        self.global_max = self.baseline

    # ---------------------------------------------------
    # 4.1 基础分箱计算 (传统龙卷风图)
    # ---------------------------------------------------
    def calculate(self, config=None):
        """
        分箱统计法 (Binning Strategy)：
        将各个输入变量依据数值大小进行排序并划分为若干箱，观察并在每箱内统计目标输出变量的波动表现。
        """
        if config is None: config = {}
        
        num_bins = config.get('num_bins', 10)
        max_vars = config.get('max_vars', 16)
        stat_type = config.get('stat_type', 'mean')
        percentile_val = config.get('percentile_val', 90.0)

        def calc_stat(arr):
            """内部闭包：根据配置的类型计算具体的统计指标"""
            if len(arr) == 0: return 0.0
            if stat_type == 'median':
                return float(np.median(arr))
            elif stat_type == 'percentile':
                return float(np.percentile(arr, percentile_val))
            else:
                return float(np.mean(arr))

        tornado_data = []
        y_values = self.df[self.target_col].values
        n_total = len(y_values)
        
        self.baseline = calc_stat(y_values)
        self.global_min = self.baseline
        self.global_max = self.baseline
        global_var = np.var(y_values)
        
        # 异常拦截：如果输出无方差或样本过少则直接返回
        if pd.isna(global_var) or global_var == 0 or n_total < 5:
            return [], self.baseline, (self.global_min, self.global_max)

        for input_col in self.input_cols:
            if input_col not in self.df.columns: continue
            
            # 对输入变量排序并切分为指定的箱数 (num_bins)
            sorted_df = self.df.sort_values(by=input_col)
            chunks = [c for c in np.array_split(sorted_df[self.target_col].values, num_bins) if len(c) > 0]
            if len(chunks) < 2: continue
            
            # 白盒已过滤噪音，直接计算各箱的统计表现
            bin_stats = [calc_stat(c) for c in chunks]
            r_min, r_max = min(bin_stats), max(bin_stats)
            r_width = r_max - r_min

            self.global_min = min(self.global_min, r_min)
            self.global_max = max(self.global_max, r_max)

            # 通过相关系数判断颜色走向 (正负相关性决定哪一端使用高亮色)
            correlation = self.df[input_col].corr(self.df[self.target_col])
            
            if correlation >= 0:
                seg_high_val, color_high = r_max, DRISK_COLOR_CYCLE[0]
                seg_low_val, color_low = r_min, DRISK_COLOR_CYCLE[1]
            else:
                seg_high_val, color_high = r_min, DRISK_COLOR_CYCLE[0]
                seg_low_val, color_low = r_max, DRISK_COLOR_CYCLE[1]

            tornado_data.append({
                'input_name': input_col,
                'range_width': r_width,
                'seg_high_val': seg_high_val,
                'seg_high_color': color_high,
                'seg_low_val': seg_low_val,
                'seg_low_color': color_low,
                'bin_stats': bin_stats,
            })

        # 按影响跨度 (range_width) 降序排序
        tornado_data.sort(key=lambda x: x['range_width'], reverse=True)
        if max_vars is not None:
            tornado_data = tornado_data[:max(1, int(max_vars))]
        return tornado_data, self.baseline, (self.global_min, self.global_max)

    # ---------------------------------------------------
    # 4.2 高级统计与回归分析 (Advanced Statistics & Regression)
    # ---------------------------------------------------
    def calculate_regression(self):
        """基于 statsmodels 的直接多元回归标准化系数 (Beta) 计算"""
        tornado_data = []
        y = self.df[self.target_col].values
        n_samples = len(y)

        if n_samples < 5 or np.std(y) == 0:
            return []

        # 标准化输出变量
        y_std = (y - np.mean(y)) / np.std(y)
        
        valid_cols = []
        X_dict = {}
        
        # 标准化输入变量
        for col in self.input_cols:
            if col not in self.df.columns: continue
            x = self.df[col].values
            if np.std(x) == 0: continue
            
            x_std = (x - np.mean(x)) / np.std(x)
            X_dict[col] = x_std
            valid_cols.append(col)
            
        if not valid_cols:
            return []

        X_df = pd.DataFrame(X_dict)

        # 白盒已过滤伪相关，直接使用 OLS 拟合提取偏回归系数
        X_final = sm.add_constant(X_df)
        final_model = sm.OLS(y_std, X_final).fit()
        
        for col in valid_cols:
            beta = final_model.params[col]
            tornado_data.append({
                'input_name': col,
                'weight': abs(beta),  
                'raw_beta': beta      
            })

        tornado_data.sort(key=lambda x: x['weight'], reverse=True)
        return tornado_data[:15]

    def calculate_regression_mapped(self):
        """基于 statsmodels 的多元回归映射值 (带原始单位映射) 计算"""
        tornado_data = []
        y_raw = self.df[self.target_col].values
        n_samples = len(y_raw)
        
        if n_samples < 5 or np.std(y_raw) == 0:
            return []
            
        valid_cols = []
        X_dict = {}
        
        for input_col in self.input_cols:
            if input_col not in self.df.columns: continue
            
            x = self.df[input_col].values
            if np.std(x) == 0: continue
            
            x_std = (x - np.mean(x)) / np.std(x)
            X_dict[input_col] = x_std
            valid_cols.append(input_col)
            
        if not valid_cols:
            return []

        X_df = pd.DataFrame(X_dict)

        # 仅对输入标准化，不标准化输出，拟合 OLS 提取带映射单位的偏回归系数
        X_final = sm.add_constant(X_df)
        final_model = sm.OLS(y_raw, X_final).fit()
        
        for col in valid_cols:
            m_val = final_model.params[col]
            tornado_data.append({
                'input_name': col,
                'weight': abs(m_val),
                'mapped_val': m_val
            })

        tornado_data.sort(key=lambda x: x['weight'], reverse=True)
        return tornado_data[:15]

    def calculate_spearman_correlation(self):
        """基于斯皮尔曼秩相关系数 (Spearman's rank correlation) 的敏感性计算"""
        tornado_data = []
        y = self.df[self.target_col].values
        
        for input_col in self.input_cols:
            if input_col not in self.df.columns: continue
            
            x = self.df[input_col].values
            if len(np.unique(x)) <= 1: continue
            
            corr, p_value = stats.spearmanr(x, y)
            
            if np.isnan(corr): continue

            tornado_data.append({
                'input_name': input_col,
                'weight': abs(corr),
                'corr_val': corr,
                'p_value': p_value
            })

        tornado_data.sort(key=lambda x: x['weight'], reverse=True)
        return tornado_data[:15]
    
    def calculate_variance_contribution(self):
        """基于 statsmodels 的顺序前向回归计算方差贡献度"""
        tornado_data = []
        y = self.df[self.target_col].values
        n_samples = len(y)

        if n_samples < 5 or np.std(y) == 0:
            return []

        y_rank = rankdata(y)
        valid_cols_info = []
        X_dict = {}
        signs = {}

        # 数据预处理：转换为秩序列
        for col in self.input_cols:
            if col not in self.df.columns: continue
            x = self.df[col].values
            if len(np.unique(x)) <= 1: continue

            corr, _ = spearmanr(x, y)
            if np.isnan(corr): continue

            X_dict[col] = rankdata(x)
            signs[col] = np.sign(corr)
            valid_cols_info.append((col, abs(corr)))

        if not valid_cols_info:
            return []

        # 按照相关性绝对值降序，依次进入模型计算纯 R² 增量 (Incremental R²)
        valid_cols_info.sort(key=lambda x: x[1], reverse=True)
        ordered_features = [item[0] for item in valid_cols_info]
        X_df = pd.DataFrame({k: X_dict[k] for k in ordered_features})

        raw_contributions = {}
        current_r2 = 0.0
        seq_features = []
        
        for feature in ordered_features:
            seq_features.append(feature)
            X_subset = sm.add_constant(X_df[seq_features])
            model = sm.OLS(y_rank, X_subset).fit()
            r2 = model.rsquared
            r2_increase = max(0.0, r2 - current_r2) 
            raw_contributions[feature] = r2_increase
            current_r2 = r2

        # 挂载正负号标识
        for col, r2_inc in raw_contributions.items():
            abs_contribution = r2_inc 
            signed_pct = abs_contribution * signs[col]
            
            tornado_data.append({
                'input_name': col,
                'weight': abs_contribution,
                'variance_pct': signed_pct
            })

        tornado_data.sort(key=lambda x: x['weight'], reverse=True)
        return tornado_data[:15]

    def calculate_scenario_analysis(
        self,
        min_pct: float,
        max_pct: float,
        significance_threshold: float | None = 0.5,
        display_limit: int | None = None,
    ):
        """极端情景分析：计算目标输出变量落在特定百分位区间（子集）时，各个输入变量的偏离度"""
        tornado_data = []
        y = self.df[self.target_col].values
        n_samples = len(y)
        
        if n_samples < 5 or min_pct >= max_pct: return []
            
        y_low = np.percentile(y, min_pct)
        y_high = np.percentile(y, max_pct)
        subset_mask = (y >= y_low) & (y <= y_high)
        if not np.any(subset_mask): return []

        for input_col in self.input_cols:
            if input_col not in self.df.columns: continue
            
            x_all = self.df[input_col].values
            std_all = np.std(x_all)
            if std_all == 0: continue
                
            median_all = np.median(x_all)
            x_sub = x_all[subset_mask]
            if len(x_sub) == 0: continue
            
            median_sub = np.median(x_sub)
            # 计算显著性：(子集中位数 - 全局中位数) / 全局标准差
            significance = (median_sub - median_all) / std_all
            pct_rank = stats.percentileofscore(x_all, median_sub)
            
            tornado_data.append({
                'input_name': input_col,
                'weight': abs(significance),    
                'significance': significance,   
                'median_sub': median_sub,       
                'pct_rank': pct_rank            
            })

        if significance_threshold is not None:
            tornado_data = [d for d in tornado_data if abs(d['significance']) >= float(significance_threshold)]

        tornado_data.sort(key=lambda x: x['weight'], reverse=True)

        if display_limit is not None:
            tornado_data = tornado_data[:max(1, int(display_limit))]
        return tornado_data

# =======================================================
# 5. 视图渲染容器 (View Rendering Container)
# =======================================================
class UITornadoView(QWidget):
    """
    [视图层] 龙卷风与敏感性分析图表的呈现容器。
    集成 PlotlyHost 与 QWebEngineView，负责将计算结果渲染为可交互的 Web 图表。
    """
    
    # ---------------------------------------------------
    # 5.1 生命周期与初始化
    # ---------------------------------------------------
    def __init__(self, parent=None, tmp_dir=""):
        super().__init__(parent)
        self.layout = QVBoxLayout(self)
        self.layout.setContentsMargins(0, 0, 0, 0)

        # 占位提示标签 (数据加载中或因过滤导致无法生成图表时显示)
        self.placeholder_label = QLabel("正在计算敏感性分析数据...", self)
        self.placeholder_label.setAlignment(Qt.AlignCenter)
        self.placeholder_label.setStyleSheet("color: #666; font-size: 14px;")
        self.layout.addWidget(self.placeholder_label)

        self.web_view = QWebEngineView(self)
        self.layout.addWidget(self.web_view)
        self.web_view.hide()

        # 实例化 Plotly 图表托管对象
        self.plot_host = PlotlyHost(self.web_view, tmp_dir, use_qt_overlay=False)

    # ---------------------------------------------------
    # 5.2 图表样式与标记生成辅助方法
    # ---------------------------------------------------
    def _get_single_marker(self, style_map, key, default_color):
        """生成单色渲染器样式对象 (用于高低值组的基础分箱龙卷风图)"""
        st = style_map.get(key, {})
        fill = DriskChartFactory._hex_to_rgba(st.get("color", default_color), st.get("fill_opacity", 1.0))
        lc_raw = st.get("outline_color", "rgba(255,255,255,0.5)")
        lc = DriskChartFactory._hex_to_rgba(lc_raw, 1.0) if lc_raw.startswith("#") else lc_raw
        return dict(color=fill, line=dict(color=lc, width=st.get("outline_width", 0.5)))

    def _get_dual_marker_array(self, values, style_map):
        """生成正负双色数组渲染器 (用于标准回归、相关性图等含有正负方向属性的图表)"""
        st_pos = style_map.get("正值", {})
        st_neg = style_map.get("负值", {})
        c_pos = DriskChartFactory._hex_to_rgba(st_pos.get("color", DRISK_COLOR_CYCLE[0]), st_pos.get("fill_opacity", 1.0))
        c_neg = DriskChartFactory._hex_to_rgba(st_neg.get("color", DRISK_COLOR_CYCLE[1]), st_neg.get("fill_opacity", 1.0))
        
        lc_pos_raw = st_pos.get("outline_color", "rgba(255,255,255,0.5)")
        lc_neg_raw = st_neg.get("outline_color", "rgba(255,255,255,0.5)")
        lc_pos = DriskChartFactory._hex_to_rgba(lc_pos_raw, 1.0) if lc_pos_raw.startswith("#") else lc_pos_raw
        lc_neg = DriskChartFactory._hex_to_rgba(lc_neg_raw, 1.0) if lc_neg_raw.startswith("#") else lc_neg_raw
        
        w_pos = st_pos.get("outline_width", 0.5)
        w_neg = st_neg.get("outline_width", 0.5)
        
        return dict(
            color=[c_pos if val >= 0 else c_neg for val in values],
            line=dict(
                color=[lc_pos if val >= 0 else lc_neg for val in values],
                width=[w_pos if val >= 0 else w_neg for val in values]
            )
        )
    
    # ---------------------------------------------------
    # 5.3 具体图表渲染通道 (Render Channels)
    # ---------------------------------------------------
    def render_chart(self, full_df: pd.DataFrame, output_col: str, input_cols: list[str], config: dict = None, style_map: dict = None,
                     xl_app=None, output_raw_key=None, input_mapping=None):
        """主渲染通道：基础分箱龙卷风图"""
        if xl_app and output_raw_key and input_mapping:
            strict_input_cols = filter_inputs_by_dependency(
                xl_app, output_raw_key, input_mapping, stop_at_makeinput=True
            )
            input_cols = [col for col in strict_input_cols if col in input_cols]

        if config is None: config = {}
        if style_map is None: style_map = {}
        
        if full_df is None or full_df.empty or not input_cols:
            self.placeholder_label.setText("未探测到有效的输入变量，无法生成龙卷风图。")
            self.placeholder_label.show()
            self.web_view.hide()
            return

        provider = TornadoDataProvider(full_df, output_col, input_cols)
        chart_data, baseline, global_range = provider.calculate(config)

        if not chart_data:
            self.placeholder_label.setText("所有变量对当前输出的影响均未超过噪声阈值。")
            self.placeholder_label.show()
            self.web_view.hide()
            return

        fig = self._build_figure(chart_data, baseline, global_range, output_col, config, style_map or {})
        self.placeholder_label.hide()
        self.web_view.show()
        self.plot_host.load_plot(plot_json=fig.to_dict(), js_mode="histogram", static_plot=False)

    def _build_figure(self, data, baseline, global_range, output_name, config, style_map):
        """内部构建器：拼装分箱龙卷风图的 Plotly Figure 实例"""
        stat_map_lbl = {"mean": "均值", "median": "中位数", "percentile": f"{config.get('percentile_val', 90):g}%分位数"}
        stat_label = stat_map_lbl.get(config.get('stat_type', 'mean'), "均值")

        def smart_format(x, sig=7):
            """智能数字格式化：避免冗长的科学计数法或显示过多小数位"""
            try: xf = float(x)
            except: return str(x)
            if np.isnan(xf): return ""
            if xf == 0: return "0"
            ax = abs(xf)
            if ax >= 1e8 or ax < 1e-8: return f"{xf:.3e}"
            exp = int(math.floor(math.log10(ax)))
            decimals = max(sig - 1 - exp, 0)
            s = f"{xf:,.{decimals}f}"
            if "." in s: s = s.rstrip("0").rstrip(".")
            return s

        fig = go.Figure()
        data_reversed = data[::-1]
        n_vars = len(data_reversed)

        y_indices = list(range(1, n_vars + 1))
        y_labels = [d['input_name'] for d in data_reversed]

        marker_high = self._get_single_marker(style_map, "高值组", DRISK_COLOR_CYCLE[0])
        marker_low = self._get_single_marker(style_map, "低值组", DRISK_COLOR_CYCLE[1])

        # 构建高值组 Bar 轨迹
        fig.add_trace(go.Bar(
            y=y_indices, x=[d['seg_high_val'] - baseline for d in data_reversed], base=baseline,
            orientation='h', marker=marker_high, name='高值组',
            text=[smart_format(d['seg_high_val']) for d in data_reversed],
            textposition='inside', insidetextanchor='end', textfont=dict(color="white", size=11),
            hovertemplate=f"变量: %{{customdata}}<br>{stat_label}: %{{text}}<extra></extra>", customdata=y_labels
        ))
        
        # 构建低值组 Bar 轨迹
        fig.add_trace(go.Bar(
            y=y_indices, x=[d['seg_low_val'] - baseline for d in data_reversed], base=baseline,
            orientation='h', marker=marker_low, name='低值组',
            text=[smart_format(d['seg_low_val']) for d in data_reversed],
            textposition='inside', insidetextanchor='end', textfont=dict(color="white", size=11),
            hovertemplate=f"变量: %{{customdata}}<br>{stat_label}: %{{text}}<extra></extra>", customdata=y_labels
        ))

        # X 轴刻度推导与智能格式化
        g_min, g_max = global_range
        span = g_max - g_min if g_max > g_min else 1.0
        dtick_x = DriskMath.calc_smart_step(span)
        x_min = math.floor(g_min / dtick_x) * dtick_x
        x_max = math.ceil(g_max / dtick_x) * dtick_x

        # 注入与普通直方图相同的 X 轴刻度文本单位生成逻辑
        forced_mag = config.get('forced_mag')
        x_mag, x_suffix, x_div, x_hint = get_si_unit_config((x_min, x_max), dtick_x, force_m=forced_mag)
        
        first_tick = math.ceil(x_min / dtick_x - 1e-9) * dtick_x
        x_vals = []
        x_texts = []
        current = first_tick
        epsilon = dtick_x * 1e-3
        while current <= x_max + epsilon:
            if abs(current) < epsilon: current = 0.0
            x_vals.append(current)
            display_num = current / x_div if x_div else current
            txt = "0" if display_num == 0 else f"{display_num:g}{x_suffix}"
            x_texts.append(txt)
            current += dtick_x

        # 绘制基线
        fig.add_vline(x=baseline, line_width=1.5, line_dash="dash", line_color="#444")

        tick_vals = list(range(1, n_vars + 2)) 
        tick_text = y_labels + [""] 

        # 动态拼接带单位的坐标轴标题
        unit_str = f" ({DriskChartFactory.VALUE_AXIS_UNIT})" if DriskChartFactory.VALUE_AXIS_UNIT else ""
        final_x_title = f"输出变量{stat_label}变化{unit_str}"
        if x_hint:
            final_x_title += f" ({x_hint})"

        fig.update_layout(
            title={'text': f"{output_name}", 'x': 0.5, 'xref': 'paper', 'xanchor': 'center', 'font': TORNADO_TITLE_FONT},
            plot_bgcolor='#fcfcfc', paper_bgcolor='#ffffff',
            xaxis={
                'title': {'text': final_x_title.strip(), 'font': TORNADO_AXIS_FONT},
                'range': [x_min, x_max], 'dtick': dtick_x,
                'tickmode': 'array', 'tickvals': x_vals, 'ticktext': x_texts, # 覆盖默认自适应刻度，应用自定义文本
                'showline': True, 'linecolor': 'black', 'ticks': 'outside', 
                'tickcolor': 'black', 'ticklen': 5, 'gridcolor': '#dddddd', 
                'gridwidth': 0.5, 'zeroline': False, 'fixedrange': True},
            yaxis={'tickmode': 'array', 'tickvals': tick_vals, 'ticktext': tick_text, 'range': [0, n_vars + 1], 
                   'showline': True, 'linecolor': 'black', 'ticks': 'outside', 'tickcolor': 'black', 'ticklen': 5, 
                   'gridcolor': '#dddddd', 'gridwidth': 0.5, 'automargin': True, 'zeroline': False, 'fixedrange': True},
            barmode='overlay', bargap=0.3, margin=dict(l=10, r=20, t=40, b=35), showlegend=True,
            legend=dict(orientation="h", yanchor="bottom", y=1.02, x=1, xanchor="right")
        )
        return fig
    
    def render_regression_chart(self, full_df: pd.DataFrame, output_col: str, input_cols: list[str], forced_mag=None, style_map: dict = None,
                                xl_app=None, output_raw_key=None, input_mapping=None):
        """渲染通道：标准化多元回归图"""
        if xl_app and output_raw_key and input_mapping:
            strict_input_cols = filter_inputs_by_dependency(
                xl_app, output_raw_key, input_mapping, stop_at_makeinput=True
            )
            input_cols = [col for col in strict_input_cols if col in input_cols]

        if not input_cols:
            self.placeholder_label.setText("未探测到有效的输入变量。")
            self.placeholder_label.show()
            self.web_view.hide()
            return

        provider = TornadoDataProvider(full_df, output_col, input_cols)
        reg_data = provider.calculate_regression()

        if not reg_data:
            self.placeholder_label.setText("未探测到显著的回归关系。")
            self.placeholder_label.show()
            self.web_view.hide()
            return

        fig = go.Figure()
        data_rev = reg_data[::-1]
        y_labels = [d['input_name'] for d in data_rev]
        values = [d['raw_beta'] for d in data_rev]
        marker_dict = self._get_dual_marker_array(values, style_map or {})
        
        fig.add_trace(go.Bar(y=list(range(1, len(data_rev) + 1)), x=values, orientation='h', marker=marker_dict,
            text=[f"{d['raw_beta']:.3f}" for d in data_rev], textposition='auto', customdata=y_labels,
            hovertemplate="变量: %{customdata}<br>标准化系数: %{x}<extra></extra>"))
            
        min_val = min(0, min([d['raw_beta'] for d in data_rev]) if data_rev else 0)
        max_val = max(0, max([d['raw_beta'] for d in data_rev]) if data_rev else 1)
        span = max_val - min_val if max_val > min_val else 1.0
        dtick_x = DriskMath.calc_smart_step(span)
        x_min = math.floor(min_val / dtick_x) * dtick_x
        x_max = math.ceil(max_val / dtick_x) * dtick_x
        if abs(x_min - min_val) < 1e-9 and min_val < 0: x_min -= dtick_x
        if abs(x_max - max_val) < 1e-9 and max_val > 0: x_max += dtick_x
        n_vars = len(data_rev)
        tick_vals = list(range(1, n_vars + 2))
        tick_text = y_labels + [""]
        fig.add_vline(x=0, line_width=1.5, line_dash="dash", line_color="#444")
        fig.update_layout(
            title={'text': f"{output_col}", 'x': 0.5, 'xref': 'paper', 'xanchor': 'center', 'font': TORNADO_TITLE_FONT},
            plot_bgcolor='#fcfcfc', paper_bgcolor='white',
            xaxis={'title': {'text': "标准化回归系数", 'font': TORNADO_AXIS_FONT}, 'range': [x_min, x_max], 'dtick': dtick_x, 'showline': True, 'linecolor': 'black', 'ticks': 'outside', 'tickcolor': 'black', 'ticklen': 5, 'gridcolor': '#dddddd', 'gridwidth': 0.5, 'zeroline': False, 'fixedrange': True},
            yaxis={'tickmode': 'array', 'tickvals': tick_vals, 'ticktext': tick_text, 'range': [0, n_vars + 1], 'showline': True, 'linecolor': 'black', 'ticks': 'outside', 'tickcolor': 'black', 'ticklen': 5, 'gridcolor': '#dddddd', 'gridwidth': 0.5, 'automargin': True, 'zeroline': False, 'fixedrange': True},
            bargap=0.3, margin=dict(l=10, r=20, t=35, b=35), showlegend=False)
        self.placeholder_label.hide()
        self.web_view.show()
        self.plot_host.load_plot(fig.to_dict(), js_mode="histogram", static_plot=False)
    
    def render_regression_mapped_chart(self, full_df: pd.DataFrame, output_col: str, input_cols: list[str], forced_mag=None, style_map: dict = None,
                                       xl_app=None, output_raw_key=None, input_mapping=None):
        """渲染通道：带原始单位映射的回归图"""
        if xl_app and output_raw_key and input_mapping:
            strict_input_cols = filter_inputs_by_dependency(
                xl_app, output_raw_key, input_mapping, stop_at_makeinput=True
            )
            input_cols = [col for col in strict_input_cols if col in input_cols]

        if not input_cols:
            self.placeholder_label.setText("未探测到有效的输入变量。")
            self.placeholder_label.show()
            self.web_view.hide()
            return

        provider = TornadoDataProvider(full_df, output_col, input_cols)
        reg_data = provider.calculate_regression_mapped()

        if not reg_data:
            self.placeholder_label.setText("未探测到显著的回归关系。")
            self.placeholder_label.show()
            self.web_view.hide()
            return

        def smart_format(x, sig=7):
            try: xf = float(x)
            except: return str(x)
            if np.isnan(xf): return ""
            if xf == 0: return "0"
            ax = abs(xf)
            if ax >= 1e8 or ax < 1e-8: return f"{xf:.3e}"
            exp = int(math.floor(math.log10(ax)))
            decimals = max(sig - 1 - exp, 0)
            s = f"{xf:,.{decimals}f}"
            if "." in s: s = s.rstrip("0").rstrip(".")
            return s

        fig = go.Figure()
        data_rev = reg_data[::-1]
        
        y_labels = [d['input_name'] for d in data_rev]
        values = [d['mapped_val'] for d in data_rev]
        marker_dict = self._get_dual_marker_array(values, style_map or {})

        fig.add_trace(go.Bar(
            y=list(range(1, len(data_rev) + 1)),
            x=values,
            orientation='h',
            marker=marker_dict,
            text=[smart_format(d['mapped_val']) for d in data_rev],
            textposition='auto',
            customdata=y_labels,
            hovertemplate=f"变量: %{{customdata}}<br>回归映射值: %{{text}}<extra></extra>"
        ))

        # --- 智能 X 轴区间计算 ---
        min_val = min(0, min([d['mapped_val'] for d in data_rev]) if data_rev else 0)
        max_val = max(0, max([d['mapped_val'] for d in data_rev]) if data_rev else 1)
        
        span = max_val - min_val if max_val > min_val else 1.0
        dtick_x = DriskMath.calc_smart_step(span)
        
        x_min = math.floor(min_val / dtick_x) * dtick_x
        x_max = math.ceil(max_val / dtick_x) * dtick_x
        
        if abs(x_min - min_val) < 1e-9 and min_val < 0: x_min -= dtick_x
        if abs(x_max - max_val) < 1e-9 and max_val > 0: x_max += dtick_x

        # 注入单位换算与后缀
        x_mag, x_suffix, x_div, x_hint = get_si_unit_config((x_min, x_max), dtick_x, force_m=forced_mag)
        
        first_tick = math.ceil(x_min / dtick_x - 1e-9) * dtick_x
        x_vals = []
        x_texts = []
        current = first_tick
        epsilon = dtick_x * 1e-3
        while current <= x_max + epsilon:
            if abs(current) < epsilon: current = 0.0
            x_vals.append(current)
            display_num = current / x_div if x_div else current
            txt = "0" if display_num == 0 else f"{display_num:g}{x_suffix}"
            x_texts.append(txt)
            current += dtick_x

        n_vars = len(data_rev)
        tick_vals = list(range(1, n_vars + 2))
        tick_text = y_labels + [""]

        fig.add_vline(x=0, line_width=1.5, line_dash="dash", line_color="#444")

        # 动态拼接带说明信息的 X 轴标题
        unit_str = f" (输出变量变化/输入变量标准差)"
        final_x_title = f"回归映射系数"
        if x_hint:
            final_x_title += f" ({x_hint})"

        fig.update_layout(
            title={'text': f"{output_col}", 'x': 0.5, 'xref': 'paper', 'xanchor': 'center', 'font': TORNADO_TITLE_FONT},
            plot_bgcolor='#fcfcfc',
            paper_bgcolor='white',
            xaxis={
                'title': {'text': final_x_title.strip(), 'font': TORNADO_AXIS_FONT},
                'range': [x_min, x_max], 'dtick': dtick_x,       
                'tickmode': 'array', 'tickvals': x_vals, 'ticktext': x_texts,
                'showline': True, 'linecolor': 'black',
                'ticks': 'outside', 'tickcolor': 'black', 'ticklen': 5,
                'gridcolor': '#dddddd', 'gridwidth': 0.5, 'zeroline': False, 'fixedrange': True
            },
            yaxis={
                'tickmode': 'array', 'tickvals': tick_vals, 'ticktext': tick_text,
                'range': [0, n_vars + 1],
                'showline': True, 'linecolor': 'black',
                'ticks': 'outside', 'tickcolor': 'black', 'ticklen': 5,
                'gridcolor': '#dddddd', 'gridwidth': 0.5, 'automargin': True, 'zeroline': False, 'fixedrange': True
            },
            bargap=0.3,
            margin=dict(l=10, r=20, t=35, b=35), 
            showlegend=False
        )

        self.placeholder_label.hide()
        self.web_view.show()
        self.plot_host.load_plot(fig.to_dict(), js_mode="histogram", static_plot=False)
    
    def render_spearman_chart(self, full_df: pd.DataFrame, output_col: str, input_cols: list[str], style_map: dict = None,
                              xl_app=None, output_raw_key=None, input_mapping=None):
        """渲染通道：斯皮尔曼秩相关系数图"""
        if xl_app and output_raw_key and input_mapping:
            strict_input_cols = filter_inputs_by_dependency(
                xl_app, output_raw_key, input_mapping, stop_at_makeinput=True
            )
            input_cols = [col for col in strict_input_cols if col in input_cols]

        if not input_cols:
            self.placeholder_label.setText("未探测到有效的输入变量。")
            self.placeholder_label.show()
            self.web_view.hide()
            return

        provider = TornadoDataProvider(full_df, output_col, input_cols)
        spearman_data = provider.calculate_spearman_correlation()

        if not spearman_data:
            self.placeholder_label.setText("未探测到显著的秩相关关系。")
            self.placeholder_label.show()
            self.web_view.hide()
            return

        fig = go.Figure()
        data_rev = spearman_data[::-1] 
        
        y_labels = [d['input_name'] for d in data_rev]
        values = [d['corr_val'] for d in data_rev]
        marker_dict = self._get_dual_marker_array(values, style_map or {})

        fig.add_trace(go.Bar(
            y=list(range(1, len(data_rev) + 1)), x=values, orientation='h', marker=marker_dict,
            text=[f"{d['corr_val']:.3f}" for d in data_rev], textposition='auto', customdata=y_labels,
            hovertemplate="变量: %{customdata}<br>秩相关系数: %{text}<extra></extra>"
        ))

        # --- 智能 X 轴区间计算 ---
        min_val = min(0, min([d['corr_val'] for d in data_rev]) if data_rev else 0)
        max_val = max(0, max([d['corr_val'] for d in data_rev]) if data_rev else 1)
        
        span = max_val - min_val if max_val > min_val else 1.0
        dtick_x = DriskMath.calc_smart_step(span)
        
        x_min = math.floor(min_val / dtick_x) * dtick_x
        x_max = math.ceil(max_val / dtick_x) * dtick_x
        
        if abs(x_min - min_val) < 1e-9 and min_val < 0: x_min -= dtick_x
        if abs(x_max - max_val) < 1e-9 and max_val > 0: x_max += dtick_x

        n_vars = len(data_rev)
        tick_vals = list(range(1, n_vars + 2))
        tick_text = y_labels + [""]

        fig.add_vline(x=0, line_width=1.5, line_dash="dash", line_color="#444")

        fig.update_layout(
            title={'text': f"{output_col}", 'x': 0.5, 'xref': 'paper', 'xanchor': 'center', 'font': TORNADO_TITLE_FONT},
            plot_bgcolor='#fcfcfc', paper_bgcolor='white',
            xaxis={
                'title': {'text': "斯皮尔曼秩相关系数", 'font': TORNADO_AXIS_FONT},
                'range': [x_min, x_max], 'dtick': dtick_x,
                'showline': True, 'linecolor': 'black',
                'ticks': 'outside', 'tickcolor': 'black', 'ticklen': 5,
                'gridcolor': '#dddddd', 'gridwidth': 0.5, 'zeroline': False, 'fixedrange': True
            },
            yaxis={
                'tickmode': 'array', 'tickvals': tick_vals, 'ticktext': tick_text, 'range': [0, n_vars + 1],
                'showline': True, 'linecolor': 'black', 'ticks': 'outside', 'tickcolor': 'black', 'ticklen': 5,
                'gridcolor': '#dddddd', 'gridwidth': 0.5, 'automargin': True, 'zeroline': False, 'fixedrange': True
            },
            bargap=0.3, margin=dict(l=10, r=20, t=35, b=35), showlegend=False
        )

        self.placeholder_label.hide()
        self.web_view.show()
        self.plot_host.load_plot(fig.to_dict(), js_mode="histogram", static_plot=False)

    def render_variance_contribution_chart(self, full_df: pd.DataFrame, output_col: str, input_cols: list[str], style_map: dict = None,
                                           xl_app=None, output_raw_key=None, input_mapping=None):
        """渲染通道：方差贡献度图"""
        if xl_app and output_raw_key and input_mapping:
            strict_input_cols = filter_inputs_by_dependency(
                xl_app, output_raw_key, input_mapping, stop_at_makeinput=True
            )
            input_cols = [col for col in strict_input_cols if col in input_cols]

        if not input_cols:
            self.placeholder_label.setText("未探测到有效的输入变量。")
            self.placeholder_label.show()
            self.web_view.hide()
            return

        provider = TornadoDataProvider(full_df, output_col, input_cols)
        var_data = provider.calculate_variance_contribution()

        if not var_data:
            self.placeholder_label.setText("未探测到显著的方差贡献关系。")
            self.placeholder_label.show()
            self.web_view.hide()
            return

        fig = go.Figure()
        data_rev = var_data[::-1]
        
        y_labels = [d['input_name'] for d in data_rev]
        values = [d['variance_pct'] for d in data_rev]
        marker_dict = self._get_dual_marker_array(values, style_map or {})
        
        # 格式化文本为带符号的百分比形式
        text_labels = [f"{d['variance_pct']*100:.1f}%" for d in data_rev]

        fig.add_trace(go.Bar(
            y=list(range(1, len(data_rev) + 1)), x=values, orientation='h', marker=marker_dict,
            text=text_labels, textposition='auto', customdata=y_labels,
            hovertemplate="变量: %{customdata}<br>方差贡献度: %{text}<extra></extra>"
        ))

        # --- 智能 X 轴区间计算 ---
        min_val = min(0, min([d['variance_pct'] for d in data_rev]) if data_rev else 0)
        max_val = max(0, max([d['variance_pct'] for d in data_rev]) if data_rev else 1)
        
        span = max_val - min_val if max_val > min_val else 1.0
        dtick_x = DriskMath.calc_smart_step(span)
        
        x_min = math.floor(min_val / dtick_x) * dtick_x
        x_max = math.ceil(max_val / dtick_x) * dtick_x
        
        if abs(x_min - min_val) < 1e-9 and min_val < 0: x_min -= dtick_x
        if abs(x_max - max_val) < 1e-9 and max_val > 0: x_max += dtick_x

        n_vars = len(data_rev)
        tick_vals = list(range(1, n_vars + 2))
        tick_text = y_labels + [""]

        fig.add_vline(x=0, line_width=1.5, line_dash="dash", line_color="#444")

        fig.update_layout(
            title={'text': f"{output_col}", 'x': 0.5, 'xref': 'paper', 'xanchor': 'center', 'font': TORNADO_TITLE_FONT},
            plot_bgcolor='#fcfcfc', paper_bgcolor='white',
            xaxis={
                'title': {'text': "方差贡献度", 'font': TORNADO_AXIS_FONT},
                'range': [x_min, x_max], 'dtick': dtick_x, 'tickformat': '.0%',     
                'showline': True, 'linecolor': 'black', 'ticks': 'outside', 'tickcolor': 'black', 'ticklen': 5,
                'gridcolor': '#dddddd', 'gridwidth': 0.5, 'zeroline': False, 'fixedrange': True
            },
            yaxis={
                'tickmode': 'array', 'tickvals': tick_vals, 'ticktext': tick_text, 'range': [0, n_vars + 1],
                'showline': True, 'linecolor': 'black', 'ticks': 'outside', 'tickcolor': 'black', 'ticklen': 5,
                'gridcolor': '#dddddd', 'gridwidth': 0.5, 'automargin': True, 'zeroline': False, 'fixedrange': True
            },
            bargap=0.3, margin=dict(l=10, r=20, t=35, b=35), showlegend=False
        )

        self.placeholder_label.hide()
        self.web_view.show()
        self.plot_host.load_plot(fig.to_dict(), js_mode="histogram", static_plot=False)

    def render_line_chart(self, full_df: pd.DataFrame, output_col: str, input_cols: list[str], config: dict = None, style_map: dict = None,
                          xl_app=None, output_raw_key=None, input_mapping=None):
        """渲染通道：趋势线图"""
        if xl_app and output_raw_key and input_mapping:
            strict_input_cols = filter_inputs_by_dependency(
                xl_app, output_raw_key, input_mapping, stop_at_makeinput=True
            )
            input_cols = [col for col in strict_input_cols if col in input_cols]

        if config is None: config = {}
        if style_map is None: style_map = {}
        
        if full_df is None or full_df.empty or not input_cols:
            self.placeholder_label.setText("未探测到有效的输入变量，无法生成趋势图。")
            self.placeholder_label.show()
            self.web_view.hide()
            return

        provider = TornadoDataProvider(full_df, output_col, input_cols)
        chart_data, baseline, global_range = provider.calculate(config)

        if not chart_data:
            self.placeholder_label.setText("所有变量对当前输出的影响均未超过噪声阈值。")
            self.placeholder_label.show()
            self.web_view.hide()
            return

        self.current_line_vars = [d['input_name'] for d in chart_data]

        fig = self._build_line_figure(chart_data, baseline, global_range, output_col, config, style_map)
        self.placeholder_label.hide()
        self.web_view.show()
        self.plot_host.load_plot(plot_json=fig.to_dict(), js_mode="histogram", static_plot=False)

    def _build_line_figure(self, data, baseline, global_range, output_name, config, style_map):
        """内部构建器：拼装趋势折线图的 Plotly Figure 实例"""
        stat_map_lbl = {"mean": "均值", "median": "中位数", "percentile": f"{config.get('percentile_val', 90):g}%分位数"}
        stat_label = stat_map_lbl.get(config.get('stat_type', 'mean'), "均值")

        fig = go.Figure()
        
        # 定义线型循环（当变量过多时，复用不同样式的虚线进行区分）
        dash_styles = ['solid', 'dash', 'dot', 'dashdot']

        # 遍历每一条趋势线
        for idx, d in enumerate(data):
            bin_stats = d['bin_stats']
            n_bins = len(bin_stats)
            if n_bins == 0: continue
            
            # 计算 X 轴百分位节点
            x_vals = [(i + 0.5) * (100.0 / n_bins) for i in range(n_bins)]
            
            # 动态应用用户设置的线型与颜色样式
            var_name = d['input_name']
            st = style_map.get(var_name, {})
            
            c_raw = st.get("curve_color", DRISK_COLOR_CYCLE[idx % len(DRISK_COLOR_CYCLE)])
            c_rgba = DriskChartFactory._hex_to_rgba(c_raw, st.get("curve_opacity", 1.0))
            w = st.get("curve_width", 2.0)
            dash = st.get("dash", dash_styles[(idx // len(DRISK_COLOR_CYCLE)) % len(dash_styles)])

            fig.add_trace(go.Scatter(
                x=x_vals, y=bin_stats, mode='lines', name=var_name,
                line=dict(color=c_rgba, width=w, dash=dash),
                hovertemplate=f"变量: %{{data.name}}<br>百分位: %{{x:.1f}}%<br>{stat_label}: %{{y:,.2f}}<extra></extra>"
            ))

        # --- Y 轴智能刻度计算 ---
        g_min, g_max = global_range
        span = g_max - g_min if g_max > g_min else 1.0
        dtick_y = DriskMath.calc_smart_step(span)
        y_min = math.floor(g_min / dtick_y) * dtick_y
        y_max = math.ceil(g_max / dtick_y) * dtick_y
        
        # 添加基准水平虚线
        fig.add_hline(y=baseline, line_width=1, line_dash="dash", line_color="#888", opacity=0.6)
        
        # 动态拼接 Y 轴单位标签
        unit_str = f" ({DriskChartFactory.VALUE_AXIS_UNIT})" if DriskChartFactory.VALUE_AXIS_UNIT else ""
        final_y_title = f"输出变量 {stat_label}{unit_str}"

        fig.update_layout(
            title={'text': f"{output_name}", 'x': 0.5, 'xref': 'paper', 'xanchor': 'center', 'font': TORNADO_TITLE_FONT},
            plot_bgcolor='#fcfcfc', paper_bgcolor='#ffffff',
            xaxis={
                'title': {'text': "输入变量百分位", 'font': TORNADO_AXIS_FONT},
                'range': [0, 100], 'dtick': 10, 'ticksuffix': '%',
                'showline': True, 'linecolor': 'black', 'ticks': 'outside', 'tickcolor': 'black', 'ticklen': 5, 
                'gridcolor': '#dddddd', 'gridwidth': 0.5, 'zeroline': False, 'fixedrange': True
            },
            yaxis={
                'title': {'text': final_y_title, 'font': TORNADO_AXIS_FONT},
                'range': [y_min, y_max], 'dtick': dtick_y,
                'showline': True, 'linecolor': 'black', 'ticks': 'outside', 'tickcolor': 'black', 'ticklen': 5, 
                'gridcolor': '#dddddd', 'gridwidth': 0.5, 'zeroline': False, 'fixedrange': True
            },
            margin=dict(l=60, r=20, t=50, b=35),
            showlegend=True,
            legend=dict(orientation="v", yanchor="middle", y=0.5, x=1.02, xanchor="left", bordercolor="#ddd", borderwidth=1)
        )
        return fig
    
    def render_scenario_chart(self, full_df: pd.DataFrame, output_col: str, input_cols: list[str], config: dict, style_map: dict = None,
                              xl_app=None, output_raw_key=None, input_mapping=None):
        """渲染通道：极端情景分析图"""
        if xl_app and output_raw_key and input_mapping:
            strict_input_cols = filter_inputs_by_dependency(
                xl_app, output_raw_key, input_mapping, stop_at_makeinput=True
            )
            input_cols = [col for col in strict_input_cols if col in input_cols]

        if not input_cols:
            self.placeholder_label.setText("未探测到有效的输入变量。")
            self.placeholder_label.show()
            self.web_view.hide()
            return

        min_pct = config.get('min_pct', 0.0)
        max_pct = config.get('max_pct', 25.0)
        display_mode = config.get('display_mode', 0) 
        significance_threshold = config.get('significance_threshold', 0.5)
        display_limit = config.get('display_limit', None)

        provider = TornadoDataProvider(full_df, output_col, input_cols)
        scenario_data = provider.calculate_scenario_analysis(
            min_pct=min_pct, max_pct=max_pct, significance_threshold=significance_threshold, display_limit=display_limit,
        )

        if not scenario_data:
            self.placeholder_label.setText("未找到符合该情景区间的有效数据。")
            self.placeholder_label.show()
            self.web_view.hide()
            return

        fig = go.Figure()
        data_rev = scenario_data[::-1] 
        y_labels = [d['input_name'] for d in data_rev]
        values = [d['significance'] for d in data_rev]
        
        marker_dict = self._get_dual_marker_array(values, style_map or {})

        # --- 动态构建数据标签文本 (Data Labels) ---
        text_labels = []
        for d in data_rev:
            sig_val = d['significance']
            if display_mode == 0:
                text_labels.append(f"{sig_val:.4f}")
            elif display_mode == 1:
                # 展现格式：-1.5617 (40.0%) 或直接显示百分位
                text_labels.append(f"({d['pct_rank']:.1f}%) {sig_val:.4f}" if sig_val > 0 else f"{sig_val:.4f} ({d['pct_rank']:.1f}%)")
            else:
                # 展现格式：-1.5617 (35,000) 
                abs_val_str = f"{d['median_sub']:,.2f}" 
                text_labels.append(f"({abs_val_str}) {sig_val:.4f}" if sig_val > 0 else f"{sig_val:.4f} ({abs_val_str})")

        fig.add_trace(go.Bar(
            y=list(range(1, len(data_rev) + 1)), x=values, orientation='h', marker=marker_dict,
            text=text_labels, textposition='auto', customdata=y_labels,
            hovertemplate="变量: %{customdata}<br>显著性: %{x:.4f}<extra></extra>"
        ))

        # --- 智能 X 轴区间计算 ---
        min_val = min(0, min([d['significance'] for d in data_rev]) if data_rev else 0)
        max_val = max(0, max([d['significance'] for d in data_rev]) if data_rev else 1)
        
        span = max_val - min_val if max_val > min_val else 1.0
        dtick_x = DriskMath.calc_smart_step(span)
        
        x_min = math.floor(min_val / dtick_x) * dtick_x
        x_max = math.ceil(max_val / dtick_x) * dtick_x
        
        # 防贴边处理，确保数据可视化边界留白
        if abs(x_min - min_val) < 1e-9 and min_val < 0: x_min -= dtick_x
        if abs(x_max - max_val) < 1e-9 and max_val > 0: x_max += dtick_x
        
        tick_vals = list(range(1, len(data_rev) + 2))
        fig.add_vline(x=0, line_width=1.5, line_color="#444")

        mode_titles = ["", "(括号内为子集中位数的百分位)", "(括号内为子集中位数实际值)"]
        
        # 动态构建副标题
        sub_text = f"{min_pct}% ＜ 输出变量 ＜ {max_pct}%"
        if mode_titles[display_mode]:
            sub_text += f"  {mode_titles[display_mode]}"

        fig.update_layout(
            # 自适应垂直居中：只保留 text 和 x 属性
            title={
                'text': f"{output_col}<br><sup>{sub_text}</sup>", 
                'x': 0.5, 'xref': 'paper', 'xanchor': 'center', 'font': TORNADO_TITLE_FONT
            },
            plot_bgcolor='#fcfcfc', paper_bgcolor='#ffffff',
            xaxis={
                'title': {'text': "输入变量显著性", 'font': TORNADO_AXIS_FONT},
                'range': [x_min, x_max], 'dtick': dtick_x,        
                'zeroline': True, 'zerolinecolor': '#444', 'zerolinewidth': 1.5,
                'showline': True, 'linecolor': 'black', 'gridcolor': '#dddddd', 'gridwidth': 0.5,
                'ticks': 'outside', 'tickcolor': 'black', 'ticklen': 5, 'fixedrange': True
            },
            yaxis={
                'tickmode': 'array', 'tickvals': tick_vals, 'ticktext': y_labels + [""],
                'range': [0, len(data_rev) + 1],
                'showline': True, 'linecolor': 'black', 'gridcolor': '#dddddd', 'gridwidth': 0.5,
                'ticks': 'outside', 'tickcolor': 'black', 'ticklen': 5, 'automargin': True, 'fixedrange': True
            },
            bargap=0.3, margin=dict(l=10, r=20, t=55, b=35), showlegend=False
        )

        self.placeholder_label.hide()
        self.web_view.show()
        self.plot_host.load_plot(fig.to_dict(), js_mode="histogram", static_plot=False)