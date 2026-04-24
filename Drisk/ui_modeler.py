# ui_modeler.py
# -*- coding: utf-8 -*-
"""
本模块提供核心的 UI 建模器服务 (Pro修复版：完美对齐 ui_results 布局与双轴边距)。
主要功能与历史修复记录：
1) 彻底肃清底层不支持的幽灵分布（如旧版对数正态、二项分布等），严格对齐 9 大核心分布。
2) 全局植入了精巧的“小号倒三角” QComboBox 样式，替代默认的粗大箭头，提升UI质感。
3) 双轴对齐：在“叠加累积概率”模式下，主动增加右侧边距 (ADDITIONAL_Y2_MARGIN) 防止图表越界。
4) [核心升级]：彻底拥抱第一组 Backend Wrapper 的高性能向量化接口 (pdf_vec, cdf_vec)，
   保证前台UI绘图与后台模拟的数学口径、截断平移逻辑达到 100% 像素级一致。
5) [交互升级]：彻底移除老旧的复选框，采用与高级分析界面一致的次级菜单按钮进行视图切换。
6) [架构适配]：解耦旧版 formula_parser 依赖，通过 backend_bridge 安全解析模型属性，防止崩溃。

【交接人注】：此文件是定义概率分布的核心入口，包含分布选择、参数调节、高级截断/平移以及与Plotly图表的双向绑定。请重点关注 recalc_distribution 中的数学边界计算逻辑。
"""

# =======================================================
# 0. 基础环境与第三方库导入 (Environment & Imports)
# =======================================================
# 🔴 核心警告：drisk_env 必须放在所有第三方库（尤其是 PySide6 和 pyxll）之前！
# 它是环境变量和路径初始化的基石，千万不能删除，导入即生效。
import drisk_env

import sys
import os
import json
import re
import math
import time
import tempfile
import traceback
from typing import Any, Optional

import numpy as np
import plotly.graph_objects as go

from com_fixer import _safe_excel_app
from pyxll import xl_app
from PySide6.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QGridLayout, QMessageBox,
                               QLabel, QFrame, QApplication, QWidget,
                               QTableWidget, QTableWidgetItem, QComboBox, QSplitter,
                               QListWidget, QListWidgetItem, QStackedWidget,
                               QPushButton, QFormLayout, QTabWidget, QScrollArea, QGroupBox, QLineEdit, QSizePolicy, QMenu)
from PySide6.QtCore import Qt, QTimer, QThreadPool, QEvent
from PySide6.QtGui import QDoubleValidator, QIntValidator, QFont, QColor, QPixmap
from PySide6.QtWebEngineCore import QWebEnginePage

# --- Drisk 内部核心模块 (Drisk Core Modules) ---
import backend_bridge
from plotly_host import PlotlyHost
from ui_overlay import DriskVLineMaskOverlay
from ui_workers import ModelerStatsJob
# 混入类 (Mixins)：用于拆分 DistributionBuilderDialog 过于庞大的逻辑代码
from ui_modeler_render_stats_mixin import DistributionBuilderRenderStatsMixin
from ui_modeler_param_validation_mixin import DistributionBuilderParamValidationMixin
from ui_modeler_param_editor_mixin import DistributionBuilderParamEditorMixin
from ui_modeler_formula_flow_mixin import DistributionBuilderFormulaFlowMixin
from ui_modeler_orchestration_mixin import DistributionBuilderOrchestrationMixin
from ui_modeler_recalc_orchestration_mixin import DistributionBuilderRecalcOrchestrationMixin
from ui_modeler_runtime_interaction_mixin import DistributionBuilderRuntimeInteractionMixin
from ui_modeler_visual_sync_mixin import DistributionBuilderVisualSyncMixin
from drisk_charting import DriskChartFactory
from drisk_export import DriskExportManager
from ui_shared import (AutoSelectLineEdit, SimpleRangeSlider, infer_si_mag, drisk_combobox_qss,
                       drisk_scrollbar_qss,
                       build_plotly_axis_style, DriskMath, SnapUtils, ChartOverlayController,
                       update_floating_value_edits_pos, create_floating_value_with_mag, set_drisk_icon, find_cell_name_ui,
                       get_shared_webview, recycle_shared_webview, ChartSkeleton, get_plotly_cache_dir, DRISK_COLOR_CYCLE_THEO,
                       resolve_visible_variable_name,
                       apply_toolbar_button_icon)
from ui_label_settings import (
    LabelSettingsDialog,
    create_default_label_settings_config,
    get_axis_display_unit_override,
    get_axis_numeric_flags,
)
from constants import DISTRIBUTION_CARD_TOOLTIPS

# =======================================================
# 1. 模块声明与全局配置 (Global Config & Registry)
# 交接说明：这里维护了整个UI的分布字典、最近使用记录以及中英文映射。
# =======================================================

DEBUG_WEB = True  
class DebugPage(QWebEnginePage):
    """用于拦截和打印 Web 引擎中 JavaScript 调试信息的类。开启DEBUG_WEB时生效。"""
    def javaScriptConsoleMessage(self, level, message, lineNumber, sourceID):
        print(f"[WEB_CONSOLE][level={int(level)}][L{lineNumber}] {sourceID}: {message}")

# === 注册表（唯一权威来源，调用后端分布） ===
DIST_REGISTRY = {}

# 定义分布界面展示的分布函数名称
# 结构：后端类名 -> 中英文字典。用于下拉菜单和界面的多语言支持。
_DISPLAY_NAME_OVERRIDES = {
    "DriskNormal": {"zh": "正态分布", "en": "Normal"},
    "DriskUniform": {"zh": "均匀分布", "en": "Uniform"},
    "DriskErf": {"zh": "误差函数分布", "en": "Erf"},
    "DriskExtvalue": {"zh": "极值分布（最大值）", "en": "Extvalue"},
    "DriskExtvalueMin": {"zh": "极值分布（最小值）", "en": "ExtvalueMin"},
    "DriskFatigueLife": {"zh": "疲劳寿命分布", "en": "FatigueLife"},
    "DriskCauchy": {"zh": "柯西分布", "en": "Cauchy"},
    "DriskDagum": {"zh": "达古姆分布", "en": "Dagum"},
    "DriskDoubleTriang": {"zh": "双三角分布", "en": "DoubleTriang"},
    "DriskGamma": {"zh": "伽马分布", "en": "Gamma"},
    "DriskErlang": {"zh": "埃尔朗分布", "en": "Erlang"},
    "DriskPoisson": {"zh": "泊松分布", "en": "Poisson"},
    "DriskBeta": {"zh": "贝塔分布", "en": "Beta"},
    "DriskChiSq": {"zh": "卡方分布", "en": "ChiSq"},
    "DriskF": {"zh": "F分布", "en": "F"},
    "DriskStudent": {"zh": "学生t分布", "en": "Student"},
    "DriskExpon": {"zh": "指数分布", "en": "Expon"},
    "DriskInvgauss": {"zh": "逆高斯分布", "en": "InvGauss"},
    "DriskGeomet": {"zh": "几何分布", "en": "Geomet"},
    "DriskHypergeo": {"zh": "超几何分布", "en": "Hypergeo"},
    "DriskIntuniform": {"zh": "整数均匀分布", "en": "IntUniform"},
    "DriskNegbin": {"zh": "负二项分布", "en": "Negbin"},
    "DriskBernoulli": {"zh": "伯努利分布", "en": "Bernoulli"},
    "DriskTriang": {"zh": "三角分布", "en": "Triang"},
    "DriskBinomial": {"zh": "二项分布", "en": "Binomial"},
    "DriskTrigen": {"zh": "广义三角分布", "en": "Trigen"},
    "DriskCumul": {"zh": "累积表分布", "en": "Cumul"},
    "DriskDUniform": {"zh": "离散均匀分布", "en": "DUniform"},
    "DriskDiscrete": {"zh": "离散表分布", "en": "Discrete"},
    "DriskFrechet": {"zh": "弗雷歇分布", "en": "Frechet"},
    "DriskGeneral": {"zh": "通用分布", "en": "General"},
    "DriskHistogrm": {"zh": "直方图分布", "en": "Histogrm"},
    "DriskHypSecant": {"zh": "双曲正割分布", "en": "HypSecant"},
    "DriskJohnsonSB": {"zh": "约翰逊SB分布", "en": "JohnsonSB"},
    "DriskJohnsonSU": {"zh": "约翰逊SU分布", "en": "JohnsonSU"},
    "DriskKumaraswamy": {"zh": "库马拉斯瓦米分布", "en": "Kumaraswamy"},
    "DriskLaplace": {"zh": "拉普拉斯分布", "en": "Laplace"},
    "DriskLevy": {"zh": "莱维分布", "en": "Levy"},
    "DriskLogistic": {"zh": "逻辑分布", "en": "Logistic"},
    "DriskLoglogistic": {"zh": "对数逻辑分布", "en": "Loglogistic"},
    "DriskLognorm": {"zh": "对数正态分布", "en": "Lognorm"},
    "DriskLognorm2": {"zh": "对数正态2分布", "en": "Lognorm2"},
    "DriskPareto": {"zh": "帕累托分布", "en": "Pareto"},
    "DriskPareto2": {"zh": "帕累托Ⅱ型分布", "en": "Pareto2"},
    "DriskPearson5": {"zh": "皮尔逊Ⅴ型分布", "en": "Pearson5"},
    "DriskPearson6": {"zh": "皮尔逊Ⅵ型分布", "en": "Pearson6"},
    "DriskPert": {"zh": "Pert分布", "en": "Pert"},
    "DriskRayleigh": {"zh": "瑞利分布", "en": "Rayleigh"},
    "DriskReciprocal": {"zh": "倒数分布", "en": "Reciprocal"},
    "DriskCompound": {"zh": "复合分布", "en": "Compound"},
    "DriskSplice": {"zh": "拼接分布", "en": "Splice"},
}

# --- 最近使用记录 (Recent Histories Config) ---
_RECENT_RETENTION_SECONDS = 30 * 24 * 60 * 60 # 缓存保留时间：30天
_RECENT_MAX_RECORDS = 40                      # 磁盘最大保留记录数
_RECENT_UI_LIMIT = 12                         # UI 面板最多显示记录数
_RECENT_FILE_NAME = "recent_distributions.json" # 存储文件名
_DIST_FAMILY_RECENT = "recent"
_DIST_FAMILY_CONTINUOUS = "continuous"
_DIST_FAMILY_DISCRETE = "discrete"
_DIST_FAMILY_SPECIAL = "special"
_SPECIAL_DISTRIBUTION_KEYS = {"Compound", "Splice"}
# --- 分布图标映射 (Icon Mapping) ---
_DISTRIBUTION_POPUP_ICON_MAP = {
    "Bernoulli": "Bernoulli_icon",
    "Binomial": "Binomial_icon",
    "Discrete": "Discrete_icon", 
    "DUniform": "DUniform_icon",
    "Geomet": "Geomet_icon",
    "Hypergeo": "HyperGeo_icon",
    "Intuniform": "IntUniform_icon",
    "Negbin": "Negbin_icon",
    "Poisson": "Poisson_icon",
    "Beta": "Beta_icon",
    "BetaGeneral": "BetaGeneral_icon",
    "BetaSubj": "dist_betasubj",
    "Burr12": "Burr12_icon",
    "Cauchy": "Cauchy_icon",
    "ChiSq": "ChiSq_icon", 
    "Cumul": "Cumul_icon",
    "Dagum": "Dagum_icon",
    "DoubleTriang": "DoubleTriang_icon",
    "Erf": "Erf_icon",
    "Erlang": "Erlang_icon",
    "Expon": "Expon_icon",
    "Extvalue": "ExtValue_icon",
    "ExtvalueMin": "ExtValueMin_icon",
    "F": "F_icon",
    "FatigueLife": "FatigueLife_icon",
    "Frechet": "Frechet_icon",
    "Gamma": "Gamma_icon",
    "General": "General_icon",
    "Histogrm": "Histogrm_icon",
    "HypSecant": "HypSecant_icon",
    "Invgauss": "InvGauss_icon",
    "JohnsonSB": "JohnsonSB_icon",
    "JohnsonSU": "JohnsonSU_icon",
    "Kumaraswamy": "Kumaraswamy_icon",
    "Laplace": "Laplace_icon",
    "Levy": "Levy_icon",
    "Logistic": "Logistic_icon",
    "Loglogistic": "LogLogistic_icon",
    "Lognorm": "Lognorm_icon",
    "Lognorm2": "Lognorm2_icon",
    "Normal": "Normal_icon",
    "Pareto": "Pareto_icon",
    "Pareto2": "Pareto2_icon",
    "Pearson5": "Pearson5_icon",
    "Pearson6": "Pearson6_icon",
    "Pert": "Pert_icon",
    "Rayleigh": "Rayleigh_icon",
    "Reciprocal": "Reciprocal_icon",
    "Student": "Student_icon",
    "Triang": "Triang_icon",
    "Trigen": "Trigen_icon",
    "Uniform": "Uniform_icon",
    "Weibull": "Weibull_icon",
    "Compound": "Compound_icon",
    "Splice": "Splice_icon",

}


# =======================================================
# 1.1 辅助工具函数 (Helper Functions)
# =======================================================
def _derive_alias_from_func_name(func_name: str) -> str:
    """提取函数名别名，例如去掉 'Drisk' 前缀。"""
    func_name = str(func_name or "").strip()
    if func_name.lower().startswith("drisk"):
        func_name = func_name[5:]
    return func_name or "Distribution"


def _extract_legacy_name_parts(raw_display_name: str) -> tuple[str, str]:
    """解析历史遗留的分布名称格式 (中英文拆分)。"""
    raw = str(raw_display_name or "").strip()
    if not raw:
        return "", ""
    core = raw.split(" - ", 1)[0].strip()
    if "(" in core and ")" in core:
        left, right = core.split("(", 1)
        zh_part = left.strip()
        en_part = right.split(")", 1)[0].strip()
        return zh_part, en_part
    return "", core


def _resolve_distribution_display_name(func_name: str, raw_display_name: str) -> tuple[str, str]:
    """根据优先级（覆写字典 > 历史遗留解析 > 函数名推导）解析分布的显示名称。"""
    override = _DISPLAY_NAME_OVERRIDES.get(func_name, {})
    legacy_zh, legacy_en = _extract_legacy_name_parts(raw_display_name)
    zh_part = str(override.get("zh") or legacy_zh or "").strip()
    en_part = str(override.get("en") or legacy_en or _derive_alias_from_func_name(func_name)).strip()
    if zh_part:
        return f"{zh_part} ({en_part})", en_part
    return en_part, en_part


def _resolve_distribution_descriptor(raw_display_name: str, is_discrete: bool) -> str:
    """解析分布的描述（连续或离散）。"""
    raw = str(raw_display_name or "").strip()
    if " - " in raw:
        return raw.split(" - ", 1)[1].strip()
    return "离散分布" if is_discrete else "连续分布"


def _distribution_sort_key(dist_key: str) -> tuple[str, str]:
    """为分布排序提供 Key。"""
    cfg = DIST_REGISTRY.get(dist_key, {})
    alias = str(cfg.get("ui_alias", "")).strip().lower()
    display = str(cfg.get("ui_name", dist_key)).strip().lower()
    return alias, display


def _get_distribution_keys_by_type(is_discrete: bool) -> list[str]:
    """按类型（离散/连续）获取已注册的分布列表。"""
    keys = [k for k, cfg in DIST_REGISTRY.items() if bool(cfg.get("is_discrete", False)) == bool(is_discrete)]
    return sorted(keys, key=_distribution_sort_key)


def _get_recent_store_path() -> str:
    """
    选取第一个可写的目录，确保历史记录持久化存储的大小可控且安全。
    探测顺序：LocalAppData -> AppData -> UserHome -> CWD -> TempDir。
    """
    candidates = []
    local_appdata = os.environ.get("LOCALAPPDATA", "").strip()
    appdata = os.environ.get("APPDATA", "").strip()
    home_dir = os.path.expanduser("~")
    if local_appdata:
        candidates.append(os.path.join(local_appdata, "Drisk"))
    if appdata:
        candidates.append(os.path.join(appdata, "Drisk"))
    if home_dir:
        candidates.append(os.path.join(home_dir, ".drisk"))
    candidates.append(os.path.join(os.getcwd(), ".drisk"))
    candidates.append(os.path.join(tempfile.gettempdir(), "drisk"))

    seen = set()
    for base_dir in candidates:
        norm = os.path.normcase(os.path.normpath(str(base_dir)))
        if not norm or norm in seen:
            continue
        seen.add(norm)
        try:
            os.makedirs(base_dir, exist_ok=True)
            test_path = os.path.join(base_dir, ".write_probe")
            with open(test_path, "w", encoding="utf-8") as fp:
                fp.write("ok")
            os.remove(test_path)
            return os.path.join(base_dir, _RECENT_FILE_NAME)
        except Exception:
            continue
    return ""


def _load_recent_distribution_records() -> list[dict]:
    """从磁盘加载最近使用的分布记录。"""
    path = _get_recent_store_path()
    if not path or (not os.path.exists(path)):
        return []
    try:
        with open(path, "r", encoding="utf-8") as fp:
            data = json.load(fp)
        return data if isinstance(data, list) else []
    except Exception:
        return []


def _save_recent_distribution_records(records: list[dict]) -> None:
    """保存最近使用的分布记录至磁盘。"""
    path = _get_recent_store_path()
    if not path:
        return
    try:
        with open(path, "w", encoding="utf-8") as fp:
            json.dump(records, fp, ensure_ascii=False, indent=2)
    except Exception:
        pass


def _normalize_recent_distribution_records(records: list[dict], now_ts: float | None = None) -> list[dict]:
    """清洗最近记录：去重、清理超过 30 天的记录、并截断最大数量。"""
    now_ts = float(now_ts if now_ts is not None else time.time())
    cutoff = now_ts - _RECENT_RETENTION_SECONDS
    seen = set()
    cleaned: list[dict] = []
    sorted_records = sorted(
        [rec for rec in records if isinstance(rec, dict)],
        key=lambda rec: float(rec.get("ts", 0) or 0),
        reverse=True,
    )
    for rec in sorted_records:
        key = str(rec.get("key", "")).strip()
        try:
            ts = float(rec.get("ts", 0))
        except Exception:
            continue
        if key not in DIST_REGISTRY:
            continue
        if ts < cutoff:
            continue
        if key in seen:
            continue
        seen.add(key)
        cleaned.append({"key": key, "ts": ts})
        if len(cleaned) >= _RECENT_MAX_RECORDS:
            break
    return cleaned


def _record_recent_distribution_use(dist_key: str) -> None:
    """插入一条新使用记录。"""
    dist_key = str(dist_key or "").strip()
    if dist_key not in DIST_REGISTRY:
        return
    now_ts = float(time.time())
    old_records = _load_recent_distribution_records()
    records = [{"key": dist_key, "ts": now_ts}] + old_records
    cleaned = _normalize_recent_distribution_records(records, now_ts=now_ts)
    _save_recent_distribution_records(cleaned)


def _get_recent_distribution_keys(limit: int = _RECENT_UI_LIMIT) -> list[str]:
    """获取UI中显示的最近分布键列表。"""
    records = _normalize_recent_distribution_records(_load_recent_distribution_records())
    _save_recent_distribution_records(records)
    safe_limit = max(0, int(limit))
    return [rec["key"] for rec in records[:safe_limit]]


def _distribution_family_to_is_discrete(family_key: str) -> bool:
    """根据族群标识判断是否为离散分布。"""
    return str(family_key or "").strip().lower() == _DIST_FAMILY_DISCRETE


def _get_distribution_keys_by_family(family_key: str) -> list[str]:
    """按族群（最近/连续/离散/特殊）获取分布列表。"""
    family = str(family_key or "").strip().lower()
    if family == _DIST_FAMILY_RECENT:
        return _get_recent_distribution_keys()
    if family == _DIST_FAMILY_SPECIAL:
        return [k for k in sorted(_SPECIAL_DISTRIBUTION_KEYS) if k in DIST_REGISTRY]

    keys = _get_distribution_keys_by_type(_distribution_family_to_is_discrete(family))
    return [k for k in keys if k not in _SPECIAL_DISTRIBUTION_KEYS]



def _format_distribution_choice_label(dist_key: str) -> str:
    """格式化分布下拉框中的显示文本，去掉括号内部分让界面更整洁。"""
    cfg = DIST_REGISTRY.get(dist_key, {})
    alias = str(cfg.get("ui_alias", "")).strip()
    if alias:
        return alias
    title = str(cfg.get("ui_name",dist_key)).strip()
    label = re.sub(r"\s*\([^)]*\)\s*$", "", title).strip()
    return label or title or str(dist_key)
def _build_distribution_card_tooltip(dist_key: str) -> str:
    """构建分布卡片悬浮提示。"""
    override = DISTRIBUTION_CARD_TOOLTIPS.get(str(dist_key or "").strip())
    if override:
       return f"{override.get('summary', '')}\n{override.get('detail', '')}".strip()


    cfg = DIST_REGISTRY.get(dist_key, {})
    name = str(cfg.get("ui_alias", "") or cfg.get("ui_name", dist_key)).strip()
    descriptor = str(cfg.get("ui_descriptor", "")).strip()
    description = str(cfg.get("description", "")).strip()

    first_line = name
    if descriptor:
        first_line = f"{name} - {descriptor}"

    if description:
        return f"{first_line}\n{description}"
    return first_line


def _resolve_distribution_popup_icon_path(dist_key: str) -> str:
    """解析分布的图标绝对路径，用于渲染卡片。"""
    icon_name = str(_DISTRIBUTION_POPUP_ICON_MAP.get(dist_key, "")).strip()
    if not icon_name:
        return ""

    base_dir = os.path.dirname(os.path.abspath(__file__))
    icon_dir = os.path.join(base_dir, "icons", "distributions")

    candidates = []
    stem, ext = os.path.splitext(icon_name)
    if ext:
        candidates.append(icon_name)
    else:
        candidates.append(f"{icon_name}.svg")
        if not icon_name.lower().endswith("_icon"):
            candidates.append(f"{icon_name}_icon.svg")
        if icon_name.lower().startswith("dist_"):
            candidates.append(f"{dist_key}_icon.svg")

    for filename in candidates:
        path = os.path.join(icon_dir, filename)
        if os.path.exists(path):
            return path
    return ""


def _populate_distribution_combo_by_family(combo: QComboBox, family_key: str, selected_key: str) -> None:
    """将分布列表填充入对应的下拉框组件中。"""
    combo.clear()
    for key in _get_distribution_keys_by_family(family_key):
        combo.addItem(_format_distribution_choice_label(key), key)

    target_index = combo.findData(selected_key)
    if target_index < 0:
        target_index = 0 if combo.count() > 0 else -1
    if target_index >= 0:
        combo.setCurrentIndex(target_index)


class _ElidedTextLabel(QLabel):
    """单行省略号标签，用于分布选择卡片。"""

    def __init__(self, text: str = "", parent=None):
        super().__init__(parent)
        self._full_text = ""
        self.set_full_text(text)

    def set_full_text(self, text: str) -> None:
        self._full_text = str(text or "")
        self._refresh_text()

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self._refresh_text()

    def _refresh_text(self) -> None:
        metrics = self.fontMetrics()
        available_width = max(8, self.contentsRect().width())
        elided = metrics.elidedText(self._full_text, Qt.ElideRight, available_width)
        super().setText(elided)


class _DistributionCardGrid(QWidget):
    """自适应卡片网格：根据可用宽度自动调整列数。"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self._buttons: list[QPushButton] = []
        self._last_cols = 0
        self._card_width = 110
        self._grid = QGridLayout(self)
        self._grid.setContentsMargins(8, 8, 8, 8)
        self._grid.setHorizontalSpacing(14)
        self._grid.setVerticalSpacing(14)
        self._grid.setAlignment(Qt.AlignTop | Qt.AlignLeft)

    def set_buttons(self, buttons: list[QPushButton]) -> None:
        self._buttons = list(buttons or [])
        self._last_cols = 0
        self._rebuild_grid(force=True)

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self._rebuild_grid()

    def _calculate_cols(self) -> int:
        margins = self._grid.contentsMargins()
        spacing = self._grid.horizontalSpacing()
        usable_width = max(
            0,
            self.contentsRect().width() - margins.left() - margins.right(),
        )
        slot_width = self._card_width + max(0, spacing)
        return max(1, usable_width // max(1, slot_width))

    def _clear_grid(self) -> None:
        while self._grid.count():
            item = self._grid.takeAt(0)
            widget = item.widget()
            if widget is not None:
                self._grid.removeWidget(widget)

    def _rebuild_grid(self, force: bool = False) -> None:
        cols = self._calculate_cols()
        if not force and cols == self._last_cols:
            return

        self._last_cols = cols
        self._clear_grid()

        for idx, button in enumerate(self._buttons):
            row = idx // cols
            col = idx % cols
            self._grid.addWidget(button, row, col)

        self._grid.setRowStretch((len(self._buttons) + cols - 1) // cols, 1)

def _initialize_ui_registry():
    """
    [初始化核心逻辑]：从 backend_bridge 动态构建 UI 分布注册表。
    交接说明：旧版系统分布是硬编码在UI里的，现已改为从后台桥接动态拉取配置，
    保证前后台算法库一致性，同时剥离 "Drisk" 前缀以便于历史数据兼容。
    """
    specs = backend_bridge.get_supported_distributions()
    for spec in specs:
        # 为了兼容历史 UI 逻辑（使用 "Normal", "Uniform" 作为键），剥离 Drisk 前缀
        ui_key = spec.func_name.replace("Drisk", "")
        
        # 兼容旧版的 params 格式: [("mean", "均值", 0), ("std", "标准差", 1)]
        params_config = []
        for i in range(len(spec.param_names)):
            p_name = spec.param_names[i]
            # 从 spec 中提取中文标签和默认值，处理越界保护
            p_label = spec.param_labels[i] if hasattr(spec, 'param_labels') and i < len(spec.param_labels) else p_name
            p_def = spec.defaults[i] if i < len(spec.defaults) else 0.0
            params_config.append((p_name, p_label, p_def))
            
        primary_name, alias_name = _resolve_distribution_display_name(spec.func_name, spec.display_name)
        descriptor = _resolve_distribution_descriptor(spec.display_name, spec.is_discrete)
        DIST_REGISTRY[ui_key] = {
            "ui_name": primary_name,
            "ui_alias": alias_name,
            "ui_descriptor": descriptor,
            "func_name": spec.func_name,
            "is_discrete": spec.is_discrete,
            "params": params_config
        }

# 模块加载时立即初始化注册表
_initialize_ui_registry()

def get_dist_config(key):
    """根据UI字典Key获取分布配置。"""
    return DIST_REGISTRY.get(key)

def get_dist_key_by_func_name(n):
    """根据真实的函数名反查UI的字典Key。"""
    for k, v in DIST_REGISTRY.items():
        if v['func_name'].upper() == n.upper():
            return k
    return None


# =======================================================
# 2. 公式解析与底层桥接层 (Formula Parsing & Backend Bridge)
# =======================================================

# ✅ 适配第2组底层架构：通过 backend_bridge 安全解析属性，替代旧版公式解析器
# 这种降级导入结构确保在系统重构期间兼容新老核心。
try:
    from formula_parser import extract_all_attributes_from_formula
except ImportError:
    def extract_all_attributes_from_formula(f):
        return {}

def find_first_distribution_in_formula(formula):
    """
    安全提取公式中的首个分布参数，优先使用 backend_bridge 的智能解析，
    若不可用则 Fallback (降级) 到通用正则提取。
    """
    try:
        # 直接使用已导入的 backend_bridge
        res = backend_bridge.parse_first_distribution_in_formula(formula)
        if res:
            func_name = res.get("func_name", "")
            k = get_dist_key_by_func_name(func_name)
            if k:
                params = res.get("dist_params", [])
                raw_args = []
                for arg in (res.get("args_list", []) or []):
                    text = str(arg).strip()
                    if not text:
                        continue

                    if k not in ("Compound", "Splice"):
                        if re.match(r"^@?\s*Drisk[A-Za-z0-9_]*\s*\(", text, re.IGNORECASE):
                            continue

                    raw_args.append(text)

                if k in ("Cumul", "Discrete", "DUniform", "General", "Histogrm", "Compound", "Splice") and raw_args:
                    params = raw_args
                elif (not params) and raw_args:
                    params = raw_args

                return k, params
    except Exception as e:
        print(f"[Drisk][ui_modeler] backend_bridge 解析出错，退回到正则: {e}")

    # 兜底方案：将识别范围严格限制在已注册的分布函数内，避免正则误伤。
    dist_func_names = [
        str(v.get("func_name", "")).strip()
        for v in DIST_REGISTRY.values()
        if isinstance(v, dict) and str(v.get("func_name", "")).strip()
    ]
    if not dist_func_names:
        return None, []

    pattern = r'@?\s*(' + "|".join(sorted((re.escape(n) for n in dist_func_names), key=len, reverse=True)) + r')\s*\('
    match = re.search(pattern, formula, re.IGNORECASE)
    if match:
        func_name = match.group(1)
        k = get_dist_key_by_func_name(func_name)
        if k:
            return k, []
    return None, []


# =======================================================
# 3. UI 视图组件：分布选择器 (Distribution Selector UI)
# 交接说明：这是点击“新建分布”时弹出的带图标的九宫格选择界面。
# =======================================================
class DistributionSelector(QDialog):
    """步骤 1：分布模型选取器对话框。"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Drisk - 选择分布模型")
        set_drisk_icon(self, "simu_icon.svg")
        self.setWindowFlag(Qt.WindowMaximizeButtonHint, True)
        self.setWindowFlag(Qt.WindowMinimizeButtonHint, True)
        self.setMinimumSize(760, 520)
        self.resize(980, 640)
        # 应用统一样式：白底黑字、微软雅黑以及全局滚动条样式
        self.setStyleSheet(
            "background-color: white; font-family: 'Microsoft YaHei';"
            + drisk_scrollbar_qss()
        )

        root_layout = QVBoxLayout(self)
        title_label = QLabel("选择概率分布模型")
        title_label.setStyleSheet("font-size:14px; font-weight:bold; margin:6px 8px;")
        root_layout.addWidget(title_label)

        body_layout = QHBoxLayout()
        body_layout.setContentsMargins(4, 0, 4, 3)
        body_layout.setSpacing(8)
        root_layout.addLayout(body_layout)

        # 左侧导航栏 (导航列表)
        self.nav_list = QListWidget()
        self.nav_list.setFixedWidth(100)
        self.nav_list.setStyleSheet(
            """
            QListWidget {
                border: 1px solid #e4e8ee;
                border-radius: 8px;
                background: #fbfcfd;
                outline: none;
                padding: 2px;
            }
            QListWidget::item {
                height: 38px;
                padding: 0 10px;
                font-size: 13px;
                color: #566577;
                border-bottom: 1px solid #edf1f4;
            }
            QListWidget::item:hover {
                background: #f5f7fa;
                color: #3f4c5b;
            }
            QListWidget::item:selected {
                background: #f0f4f8;
                color: #334455;
                font-weight: 600;
                border: none;
            }
            """
            + drisk_scrollbar_qss()
        )
        self.nav_list.addItem(QListWidgetItem("最近使用"))
        self.nav_list.addItem(QListWidgetItem("连续分布"))
        self.nav_list.addItem(QListWidgetItem("离散分布"))
        body_layout.addWidget(self.nav_list)
        self.nav_list.addItem(QListWidgetItem("特殊分布"))

        # 右侧堆叠卡片区 (页面容器)
        self.page_stack = QStackedWidget()
        body_layout.addWidget(self.page_stack, 1)

        recent_keys = _get_distribution_keys_by_family(_DIST_FAMILY_RECENT)
        cont_keys = _get_distribution_keys_by_family(_DIST_FAMILY_CONTINUOUS)
        disc_keys = _get_distribution_keys_by_family(_DIST_FAMILY_DISCRETE)
        special_keys = _get_distribution_keys_by_family(_DIST_FAMILY_SPECIAL)
        self.page_stack.addWidget(self._build_section_page(recent_keys, empty_hint="暂无最近使用记录（最近30天）"))
        self.page_stack.addWidget(self._build_section_page(cont_keys))
        self.page_stack.addWidget(self._build_section_page(disc_keys))
        self.page_stack.addWidget(self._build_section_page(special_keys, empty_hint="暂无特殊分布"))

        self.nav_list.currentRowChanged.connect(self._on_nav_changed)
        self.nav_list.setCurrentRow(0)
        self.selected_dist = None

    def _on_nav_changed(self, row: int):
        """同步左侧导航点击与右侧页面切换。"""
        row = max(0, min(row, self.page_stack.count() - 1))
        self.page_stack.setCurrentIndex(row)

    def _build_section_page(self, dist_keys: list[str], empty_hint: str = "") -> QWidget:
        """构建单个子页面的卡片网格系统。"""
        page = QWidget()
        outer_layout = QVBoxLayout(page)
        outer_layout.setContentsMargins(0, 0, 0, 0)
        outer_layout.setSpacing(0)

        if not dist_keys:
            empty_label = QLabel(empty_hint or "暂无可用分布")
            empty_label.setAlignment(Qt.AlignCenter)
            empty_label.setStyleSheet("color: #8c8c8c; font-size: 13px;")
            outer_layout.addWidget(empty_label, 1)
            return page

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.NoFrame)
        content = _DistributionCardGrid()
        buttons = [self._create_dist_card(key) for key in dist_keys]
        content.set_buttons(buttons)
        scroll.setWidget(content)
        outer_layout.addWidget(scroll, 1)
        return page

    def _create_dist_card(self, key: str) -> QPushButton:
        """生成带图标和文本的分布卡片按钮。"""
        card_text = _format_distribution_choice_label(key)
        icon_path = _resolve_distribution_popup_icon_path(key)

        btn = QPushButton()
        btn.setFixedSize(110, 120)
        btn.setToolTip(_build_distribution_card_tooltip(key))
        btn.setStyleSheet(
            """
            QPushButton {
                border: 1px solid #e4e8ee;
                border-radius: 7px;
                background: #fbfcfd;
                color: #2f3a46;
                padding: 0;
            }
            QPushButton:hover {
                background: #f0f4f8;
                border: 1px solid #8c9eb5;
                color: #334455;
            }
            QPushButton:pressed {
                background: #e8edf3;
            }
            """
        )

        column = QVBoxLayout(btn)
        column.setContentsMargins(8, 8, 8, 8)
        column.setSpacing(8)

        icon_label = QLabel()
        icon_label.setFixedSize(64, 64)
        icon_label.setAlignment(Qt.AlignCenter)
        if icon_path:
            pixmap = QPixmap(icon_path)
            if not pixmap.isNull():
                icon_label.setPixmap(pixmap.scaled(64, 64, Qt.KeepAspectRatio, Qt.SmoothTransformation))
        icon_label.setAttribute(Qt.WA_TransparentForMouseEvents, True)
        column.addWidget(icon_label, 0, Qt.AlignHCenter | Qt.AlignBottom)

        text_label = _ElidedTextLabel(card_text)
        text_label.setFixedWidth(80)
        text_label.setFixedHeight(18)
        text_label.setAlignment(Qt.AlignHCenter | Qt.AlignTop)
        text_label.setStyleSheet("background: transparent;")
        text_label.setAttribute(Qt.WA_TransparentForMouseEvents, True)
        column.addWidget(text_label, 1, Qt.AlignHCenter | Qt.AlignTop)

        btn.clicked.connect(lambda checked=False, dist_key=key: self.select_dist(dist_key))
        return btn

    def select_dist(self, key):
        """确认选择并将结果回传。"""
        self.selected_dist = key
        _record_recent_distribution_use(key)
        self.accept()


# =======================================================
# 4. 核心组件：分布建模主控面板 (Core Component: Main Dialog)
# 交接说明：这是整个文件最核心、最复杂的类。它继承自多个 Mixin，
# 分别处理统计面板、参数校验、编辑流、Web渲染交互等。
# =======================================================
class DistributionBuilderDialog(DistributionBuilderRenderStatsMixin, DistributionBuilderParamValidationMixin, DistributionBuilderParamEditorMixin, DistributionBuilderFormulaFlowMixin, DistributionBuilderOrchestrationMixin, DistributionBuilderRecalcOrchestrationMixin, DistributionBuilderRuntimeInteractionMixin, DistributionBuilderVisualSyncMixin, QDialog):
    """分布定义器（主工作区）"""

    # -------------------------------------------------------
    # 4.1 初始化与布局构建 (Initialization & UI Layout)
    # -------------------------------------------------------
    def __init__(self, dist_type="Normal", initial_params=None, initial_attrs=None,
                 full_formula="", cell_address="", parent=None, display_name_hint=None):
        super().__init__(parent)

        # 解析原始公式中显式写死的属性
        explicit_attrs = extract_all_attributes_from_formula(full_formula if full_formula else f"={dist_type}()")
        # 将合并后的属性赋给 initial_attrs，供 setup_attribute_inputs 自动渲染到输入框
        # 仅允许将公式中显式存在的属性预填至输入框。
        self.initial_attrs = explicit_attrs

        # 仅在显示层保留手动修改的图表标题；绝对不要修改底层的元数据属性。
        self._chart_title_display_override = None
        # 从相邻单元格提取的仅供显示的名称提示，用作兜底命名。
        self._cell_detected_display_hint = str(display_name_hint).strip() if display_name_hint else ""

        # 设置全局兜底名称（用户手动清空输入框时恢复的默认标题）
        self.fallback_chart_name = (
            self._cell_detected_display_hint
            if self._cell_detected_display_hint
            else (cell_address if cell_address else "新建建模变量")
        )
        self.chart_display_name = self._resolve_modeler_display_chart_title()
             
        # 实例化图名标签
        self.chart_title_label = QLabel(self.chart_display_name)
        self.chart_title_label.setObjectName("chartTitleLabel") # 🟢 新增对象标识
        title_font = QFont()
        title_font.setFamilies(["Arial", "Microsoft YaHei"]) 
        title_font.setPointSize(14)
        self.chart_title_label.setFont(title_font)
        self.chart_title_label.setAlignment(Qt.AlignCenter)
        self.chart_title_label.setStyleSheet("color: #333333; margin: 0px; padding: 0px;")

        # [参数合法性规则字典]：用于防呆和输入框校验变红
        def _rules_by_position(dist_key, indexed_rules, cross=None):
           cfg = get_dist_config(dist_key) or {}
           params = cfg.get("params", [])
           rules = {}
           for idx, rule in indexed_rules.items():
              if idx < len(params):
                  rules[params[idx][0]] = rule
           if cross:
               rules["__cross__"] = cross
           return rules

        _PARAM_RULES = {
            "Normal": _rules_by_position("Normal", {
               1: {"type": "pos", "min": 0.0, "exclusive_min": True},
    }),
            "Uniform": _rules_by_position("Uniform", {}, "uniform_min_lt_max"),
            "Gamma": _rules_by_position("Gamma", {
                0: {"type": "pos", "min": 0.0, "exclusive_min": True},
                1: {"type": "pos", "min": 0.0, "exclusive_min": True},
    }),
            "Poisson": _rules_by_position("Poisson", {
                0: {"type": "pos", "min": 0.0, "exclusive_min": True},
    }),
            "Bernoulli": _rules_by_position("Bernoulli", {
                0: {"type": "prob", "min": 0.0, "max": 1.0, "exclusive_min": True, "exclusive_max": True},
    }),
            "Binomial": _rules_by_position("Binomial", {
                0: {"type": "int", "min": 1},
                1: {"type": "prob", "min": 0.0, "max": 1.0, "exclusive_min": True, "exclusive_max": True},
    }),
            "Geomet": _rules_by_position("Geomet", {
                0: {"type": "prob", "min": 0.0, "max": 1.0, "exclusive_min": True},
    }),
            "Hypergeo": _rules_by_position("Hypergeo", {
                0: {"type": "int", "min": 1},
                1: {"type": "int", "min": 1},
                2: {"type": "int", "min": 1},
    }, "hypergeo_bounds"),

            "Beta": _rules_by_position("Beta", {
                0: {"type": "pos", "min": 0.0, "exclusive_min": True},
                1: {"type": "pos", "min": 0.0, "exclusive_min": True},
               }),
            "ChiSq": _rules_by_position("ChiSq", {
                0: {"type": "pos", "min": 0.0, "exclusive_min": True},
                  }),
            "F": _rules_by_position("F", {
                0: {"type": "pos", "min": 0.0, "exclusive_min": True},
                1: {"type": "pos", "min": 0.0, "exclusive_min": True},
    }),
            "Student": _rules_by_position("Student", {
                0: {"type": "pos", "min": 0.0, "exclusive_min": True},
    }),
            "Expon": _rules_by_position("Expon", {
                0: {"type": "pos", "min": 0.0, "exclusive_min": True},
    }),
            "Negbin": _rules_by_position("Negbin", {
                0: {"type": "pos", "min": 0.0, "exclusive_min": True},
                1: {"type": "prob", "min": 0.0, "max": 1.0, "exclusive_min": True, "exclusive_max": True},
    }),
            "Compound": _rules_by_position("Compound", {
                0: {"type": "formula"},
                1: {"type": "formula"},
                2: {"type": "optional_nonneg"},
                3: {"type": "optional_nonneg", "allow_inf": True},
    }),
            "Splice": _rules_by_position("Splice", {
                0: {"type": "formula"},
                1: {"type": "formula"},
                2: {"type": "finite"},
    }),
}



        self._PARAM_RULES = _PARAM_RULES

        self.cell_address = cell_address or ""

        title = "定义分布"
        if cell_address:
            title += f" : {cell_address}"
        self.setWindowTitle(title)
        set_drisk_icon(self)

        # 🔴 自适应屏幕大小逻辑 (完美适配 13 寸 ~ 32 寸)
        screen_geo = QApplication.primaryScreen().availableGeometry()
        # 设定理想大小，但绝不跨越屏幕可用宽度的 90% 和高度的 80%
        target_w = min(1200, int(screen_geo.width() * 0.90))
        target_h = min(800, int(screen_geo.height() * 0.80))
        
        self.resize(target_w, target_h)
        self.setWindowFlags(
            self.windowFlags()
            | Qt.WindowMinimizeButtonHint
            | Qt.WindowMaximizeButtonHint
            | Qt.WindowCloseButtonHint
        )

        # 彻底重写 base_qss，使用高优先级选择器，实现精细排版控制
        base_qss = """
            QDialog { background-color: white; font-family: 'Microsoft YaHei'; }
            
            /* 🟢【顶部大标题】：绝对锁定 14px 不加粗 */
            QLabel#chartTitleLabel { 
                font-family: 'Microsoft YaHei'; 
                font-size: 14px; 
                font-weight: normal; 
                color: #333333; 
            }
            
            /* 🟢【蓝字模块标题】（分布模型、参数设置等）：锁定 12px 不加粗 */
            QGroupBox { 
                border: 1px solid #eee; 
                border-radius: 5px; 
                margin-top: 15px; 
                padding-top: 10px; 
            }
            QGroupBox::title { 
                subcontrol-origin: margin; 
                left: 4px; 
                padding: 0 3px; 
                color: #0050b3; 
                font-size: 12px; 
                font-weight: normal; 
            }

            /* 强制限制左侧输入框和下拉框的最大宽度，防止被 Layout 拉伸得过长 */
            QGroupBox QComboBox, QGroupBox QLineEdit {
                min-width: 35px;
            }
            
            /* 🟢【红框内所有元素】：终极强制锁定 11px 不加粗 */
            /* 加上 QGroupBox 前缀以提升优先级，压制系统默认和 shared 里的 12px */
            QGroupBox QLabel, QGroupBox QLineEdit, QGroupBox QComboBox, QGroupBox QPushButton, QGroupBox QAbstractItemView {
                font-family: 'Microsoft YaHei';
                font-size: 11px; 
                font-weight: normal; 
                color: #333333;
            }
            
            /* 保证输入框的基础外观，校验失败变红 */
            QLineEdit { 
                border: 1px solid #d9d9d9; 
                border-radius: 3px; 
                padding: 3px; 
            }
            QLineEdit[invalid="true"] { 
                border: 1px solid #ff4d4f; 
                background: #fff1f0; 
            }
            QLineEdit:focus { border-color: #40a9ff; }
        """
        # 🟢【核心修复】：调换拼接顺序！把 base_qss 放后面，利用层叠规则覆盖掉 shared 里的 12px 设定
        self.setStyleSheet(drisk_combobox_qss() + drisk_scrollbar_qss() + base_qss)

        self.cell_address = cell_address

        # 初始化公式推导逻辑
        if not full_formula or not full_formula.startswith("="):
            cfg = get_dist_config(dist_type) or {}
            func = cfg.get('func_name', 'DriskNormal')
            
            # 🔴 修复 1：初始化公式时，如果默认参数里带有逗号且没有括号（比如数组），强行穿上 {} 防弹衣！避免解析错乱。
            defaults = []
            for p in cfg.get('params', []):
                d_val = str(p[2]).strip()
                if ',' in d_val and not any(d_val.startswith(c) for c in ['{', '[', '"', "'"]):
                    d_val = f"{{{d_val}}}"
                defaults.append(d_val)
                
            if not defaults: defaults = ["0", "1"]
            self.full_formula = f"={func}({','.join(defaults)})"
        else:
            self.full_formula = full_formula

        self.dist_type = dist_type
        self.config = get_dist_config(dist_type) or get_dist_config("Normal")
        self.is_discrete = self.config.get('is_discrete', False)
        self._dist_func_name = str(self.config.get('func_name', 'DriskNormal'))
        self._dist_params = []
        self._dist_markers = {}
        self.view_mode = "auto"

        self.initial_params = initial_params or {}
        # 保留显式的公式属性；不要使用仅供展示的提示信息进行覆盖。
        self.initial_attrs = self.initial_attrs or {}

        # 状态控制锁
        self.is_updating_formula = False
        self.current_segment_idx = -1
        self.manual_mag = None  
        self._label_settings_config = create_default_label_settings_config()
        self._label_axis_numeric = {"x": True, "y": True}
        self.formula_segments = []

        # 绘图区域边距设定
        self.MARGIN_L = 80
        self.MARGIN_R = 40
        self.MARGIN_T = 0  
        self.MARGIN_B = 60
        self.ADDITIONAL_Y2_MARGIN = 55 # 当开启 CDF 叠加（双Y轴）时，留给右侧的额外边距

        self._computed_margins = (self.MARGIN_L, self.MARGIN_R, self.MARGIN_T, self.MARGIN_B)
        self._has_y2 = False

        self.inputs = {}
        self.attr_inputs = {}
        self.dist_attr_inputs = {}
        self._seed_custom_rng_type = None
        self._seed_custom_value = None
        self._seed_custom_confirmed = False
        self._seed_mode_syncing = False
        self.dist_obj = None

        self._tmp_dir = get_plotly_cache_dir()  

        self.x_dtick = 1.0
        self.view_min = 0.0
        self.view_max = 1.0

        # 初始化指令：离散默认为 PMF，连续默认为 PDF
        self._current_view_data = "pmf" if self.is_discrete else "pdf"

        # 启动界面初始化构建
        self.init_ui()
        self._apply_modeler_chart_title_label_style()
        self.parse_formula_structure()

        self.throttle = QTimer()
        self.throttle.setInterval(30)
        self.throttle.setSingleShot(True)
        self.throttle.timeout.connect(self.perform_buffered_update)

        self._stats_pool = QThreadPool(self)
        self._stats_pool.setMaxThreadCount(1)
        self._stats_job_token = 0

        self._drag_snap_enabled = True
        self._drag_grid_step = None
        self._drag_grid_xmin = None
        self._drag_grid_xmax = None
        self._in_slider_sync = False

        self._param_debounce = QTimer(self)
        self._param_debounce.setSingleShot(True)
        self._param_debounce.setInterval(160)
        self._param_debounce.timeout.connect(self._flush_param_change)

        self._is_slider_dragging = False
        self._block_slider_events = False
        self._stats_table_ever_rendered = False

        self._stats_cache_key = None
        self._stats_cache_rows = None
        self._stats_render_key = None
        self._stats_table_timer = QTimer()
        self._stats_table_timer.setSingleShot(True)
        self._stats_table_timer.setInterval(0)
        self._stats_table_timer.timeout.connect(self._flush_stats_table_update)

        self._stats_force_next = False

        # --- 新增: 初始化防抖与节流定时器 (防止 resize 时 JS 引擎拥塞导致卡顿) ---
        self._last_plotly_resize_time = 0
        self._plotly_resize_timer = QTimer(self)
        self._plotly_resize_timer.setSingleShot(True)
        self._plotly_resize_timer.setInterval(150)
        self._plotly_resize_timer.timeout.connect(self._force_plotly_resize)

        # 初次触发数学引擎计算
        self.recalc_distribution()

    def init_ui(self):
        """核心界面初始化构建（划分为 5 个大区：公式栏、分离器、左参数、中图表、右统计、底操作）"""
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)

        # --- Zone 1: 顶部公式栏 (Formula Bar) ---
        top_frame = QFrame()
        top_frame.setFixedHeight(50)
        top_frame.setStyleSheet("background:#fcfcfc; border-bottom:1px solid #eee;")
        tl = QHBoxLayout(top_frame)
        tl.setContentsMargins(15, 8, 15, 8)

        self.formula_edit = QLineEdit(self.full_formula)
        self.formula_edit.setStyleSheet("border:1px solid #ccc; padding:4px; font-family:Consolas; font-size:14px;")
        self.formula_edit.cursorPositionChanged.connect(self.on_formula_cursor_move)
        tl.addWidget(self.formula_edit)

        self.seg_hint = QLabel("")
        self.seg_hint.setStyleSheet("color:#d46b08; font-size:12px;")
        self.seg_hint.setVisible(False)
        tl.addWidget(self.seg_hint)

        self.btn_sync_seg = QPushButton("同步当前段")
        self.btn_sync_seg.setFixedHeight(30)
        self.btn_sync_seg.setVisible(False)
        self.btn_sync_seg.clicked.connect(self.sync_current_segment_to_formula)
        tl.addWidget(self.btn_sync_seg)

        main_layout.addWidget(top_frame)

        # --- Zone 2: 中间主工作区 (QSplitter 动态分割，支持左右拖拽) ---
        middle_splitter = QSplitter(Qt.Horizontal)
        middle_splitter.setHandleWidth(2)
        middle_splitter.setStyleSheet("QSplitter::handle { background-color: #c0c0c0; }")

        # --- Zone 3: 左侧参数控制面板 (Left Panel) ---
        left_scroll = QScrollArea()
        left_scroll.setMinimumWidth(150)
        left_scroll.setWidgetResizable(True)
        left_scroll.setFrameShape(QFrame.NoFrame)
        left_content = QWidget()
        left_vbox = QVBoxLayout(left_content)
        left_vbox.setContentsMargins(15, 20, 15, 20)
        left_vbox.setSpacing(15)
        left_scroll.setStyleSheet("background:#fcfcfc;" + drisk_scrollbar_qss())

        gb1 = QGroupBox("分布模型")
        gl1 = QVBoxLayout(gb1)
        self.combo_dist_family = QComboBox()
        self.combo_dist_family.addItem("最近使用", _DIST_FAMILY_RECENT)
        self.combo_dist_family.addItem("连续分布", _DIST_FAMILY_CONTINUOUS)
        self.combo_dist_family.addItem("离散分布", _DIST_FAMILY_DISCRETE)
        self.combo_dist_family.addItem("特殊分布", _DIST_FAMILY_SPECIAL)
        gl1.addWidget(self.combo_dist_family)

        self.combo = QComboBox()
        self._sync_dist_family_with_key(self.dist_type)

        self.combo_dist_family.currentIndexChanged.connect(self.on_dist_family_changed)
        self.combo.currentIndexChanged.connect(self.on_model_changed)
        gl1.addWidget(self.combo)
        left_vbox.addWidget(gb1)

        self.gb_params = QGroupBox("参数设置")
        # 🔴 字体/字号等已在 base_qss 中统一拦截控制，此处不硬编码
        self.form = QFormLayout(self.gb_params)
        self.form.setSpacing(12)
        left_vbox.addWidget(self.gb_params)
        
        # 👇 在此处调用高级形态UI构建函数（用于管理截断和平移逻辑）
        self.setup_advanced_morphology_ui(left_vbox)

        self.gb_attrs = QGroupBox("分布描述")
        self.attr_form = QFormLayout(self.gb_attrs)
        self.attr_form.setSpacing(10)
        self.create_static_attribute_fields()
        left_vbox.addWidget(self.gb_attrs)

        self.gb_dist_attrs = QGroupBox("分布属性")
        self.dist_attr_form = QFormLayout(self.gb_dist_attrs)
        self.dist_attr_form.setSpacing(10)
        self.create_distribution_attribute_fields()
        left_vbox.addWidget(self.gb_dist_attrs)
        left_vbox.addStretch()

        left_scroll.setWidget(left_content)
        middle_splitter.addWidget(left_scroll)

        # --- 中间核心交互区 (Plotly 图表与游标滑块) ---
        self.center_widget = QWidget() 
        self.center_widget.setMinimumWidth(400) 
        self.center_widget.setStyleSheet("background-color: #ffffff;") 

        # ✅ 恢复原生单层垂直布局
        center_vbox = QVBoxLayout(self.center_widget)
        center_vbox.setContentsMargins(0, 0, 0, 0)
        center_vbox.setSpacing(0)

        # 滑块顶层包装容器
        self.slider_wrapper = QWidget(self.center_widget)
        self.slider_wrapper.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)

        top_layout = QVBoxLayout(self.slider_wrapper)
        top_layout.setContentsMargins(0, 0, 0, 0)
        top_layout.setSpacing(0)

        # 修改：为标题创建一个独立的水平布局以控制边距，防止图名越界
        self.title_container_layout = QHBoxLayout()
        self.title_container_layout.setContentsMargins(self.MARGIN_L, 0, self.MARGIN_R, 0)
        self.title_container_layout.addWidget(self.chart_title_label)
        top_layout.addLayout(self.title_container_layout)

        # 悬浮的极值输入框 (滑块两端)
        self.float_x_l = create_floating_value_with_mag(self.slider_wrapper, height=22, opacity=0.95)
        self.float_x_r = create_floating_value_with_mag(self.slider_wrapper, height=22, opacity=0.95)

        try:
            _float_h = int(max(self.float_x_l.sizeHint().height(), self.float_x_r.sizeHint().height()))
        except Exception:
            _float_h = 22
        top_pad = int(_float_h + 2)
        
        top_layout.addSpacing(top_pad)

        # Drisk定制的游标滑块
        self.slider = SimpleRangeSlider()
        self.slider.setMargins(self.MARGIN_L, self.MARGIN_R)
        self.slider.rangeChanged.connect(self.on_slider_change)
        if hasattr(self.slider, "dragStarted"):
            self.slider.dragStarted.connect(self.on_slider_drag_started)
        if hasattr(self.slider, "dragFinished"):
            self.slider.dragFinished.connect(self.on_slider_drag_finished)
        if hasattr(self.slider, "inputConfirmed"):
            self.slider.inputConfirmed.connect(self._on_slider_input_confirmed)

        try:
            if hasattr(self.slider, "bind_line_edits"):
                self.slider.bind_line_edits(self.float_x_l.edit, self.float_x_r.edit)
        except Exception:
            pass
        top_layout.addWidget(self.slider)

        try:
            _slider_h = int(self.slider.sizeHint().height())
            _title_h = self.chart_title_label.sizeHint().height() or 24
            self.slider_wrapper.setMinimumHeight(_title_h + int(top_pad) + _slider_h)
        except Exception:
            self.slider_wrapper.setMinimumHeight(24 + int(top_pad) + 40)
            
        center_vbox.addWidget(self.slider_wrapper, 0)
        
        try:
            self.float_x_l.edit.returnPressed.connect(lambda: self._apply_floating_input("l"))
            self.float_x_r.edit.returnPressed.connect(lambda: self._apply_floating_input("r"))
        except Exception:
            pass

        # 加载共享的 Webview 渲染 Plotly
        self.web = get_shared_webview(parent_widget=self.center_widget)
        if DEBUG_WEB:
            self.web.setPage(DebugPage(self.web))
        self._install_chart_context_menu_hook()

        self.web.setStyleSheet("background:transparent;")
        center_vbox.addWidget(self.web, 1)

        # ✅ 将骨架屏作为 center_widget 的独立悬浮子控件（加载中遮罩层）
        self.chart_skeleton = ChartSkeleton(self.center_widget)
        self.chart_skeleton.hide()

        # ✅ 开启事件监听，处理尺寸缩放
        self.center_widget.installEventFilter(self)

        self._use_qt_overlay = True

        l0, r0, t0, b0 = self._computed_margins
        self.overlay = DriskVLineMaskOverlay(
            self.web, margin_l=int(l0), margin_r=int(r0), margin_t=int(t0), margin_b=int(b0)
        )

        self.overlay.setGeometry(self.web.rect())
        self.overlay.show()

        self.chart_skeleton.raise_()

        self._chart_ctrl = ChartOverlayController(self.web, self.overlay)
        
        self._plot_host = PlotlyHost(
            web_view=self.web,
            tmp_dir=self._tmp_dir,
            use_qt_overlay=True,
            overlay=self.overlay
        )

        self._webview_loaded = False
        self._data_sent = False
        self._render_token = 0
        self._waiting_token = 0
        self._js_inflight = False
        self._ready_deadline = 0.0

        self._ready_poll_timer = QTimer(self)
        self._ready_poll_timer.setInterval(50)
        self._ready_poll_timer.timeout.connect(self._poll_plotly_ready)

        try:
            self.web.loadFinished.connect(self._on_webview_load_finished)
        except Exception:
            pass
        
        QTimer.singleShot(0, self._sync_plot_geom_from_plotly)

        middle_splitter.addWidget(self.center_widget)

        # --- Zone 4: 右侧统计指标面板 (Stats Table) ---
        right_panel = QFrame()
        right_panel.setMinimumWidth(130)
        right_panel.setStyleSheet("background-color: #ffffff;")
        rl = QVBoxLayout(right_panel)
        rl.setContentsMargins(0, 0, 0, 0)

        header = QLabel("   统计指标（理论）")
        header.setFixedHeight(36)
        header.setStyleSheet(
            "background-color: #fafafa; color: #333; font-weight: bold; font-size: 14px; border-bottom: 1px solid #f0f0f0;")
        rl.addWidget(header)

        self.tbl = QTableWidget()
        self.tbl.setColumnCount(2)
        self.tbl.verticalHeader().hide()
        self.tbl.horizontalHeader().show()
        self.tbl.setColumnWidth(0, 90)
        self.tbl.horizontalHeader().setStretchLastSection(True)
        self.tbl.setAlternatingRowColors(False)
        self.stats_table = self.tbl
        rl.addWidget(self.tbl)
        self.tbl.cellClicked.connect(self.on_stats_cell_clicked)

        middle_splitter.addWidget(right_panel)

        # 设置 Splitter 的初始比例与拉伸因子
        middle_splitter.setSizes([200, 720, 150])
        middle_splitter.setStretchFactor(0, 0)
        middle_splitter.setStretchFactor(1, 1)
        middle_splitter.setStretchFactor(2, 0)

        main_layout.addWidget(middle_splitter)

        # --- Zone 5: 底部工具栏与操作按钮 (Bottom Action Bar) ---
        bottom_bar = QFrame()
        bottom_bar.setFixedHeight(34)
        bottom_bar.setStyleSheet("""
                    QFrame { background:#f5f5f5; border-top:1px solid #e0e0e0; }
                    QLabel { color: #555; font-size: 12px; }
                    QPushButton { font-size: 12px; border-radius: 3px; padding: 2px 10px; } 
                """)
        bl = QHBoxLayout(bottom_bar)
        bl.setContentsMargins(15, 1, 15, 1)
        bl.setSpacing(10)

        self.btn_view_mode = QPushButton("分布图")
        self.btn_view_mode.setObjectName("btnViewMode")
        self.btn_view_mode.setStyleSheet("""
            QPushButton { background-color: #f0f0f0; border: 1px solid #d9d9d9; border-radius: 3px; padding: 0px 8px; font-size: 12px; height: 20px;}
            QPushButton:hover { border-color: #40a9ff; color: #40a9ff; background-color: #e6f7ff; }
        """)
        self.btn_view_mode.clicked.connect(self._show_view_mode_menu)
        bl.addWidget(self.btn_view_mode)

        self.btn_label_settings = QPushButton("文本设置")
        self.btn_label_settings.setFixedHeight(20)
        self.btn_label_settings.setStyleSheet("""
            QPushButton { background-color: #f0f0f0; border: 1px solid #d9d9d9; border-radius: 3px; padding: 0px 8px; font-size: 12px; height: 20px;}
            QPushButton:hover { border-color: #40a9ff; color: #40a9ff; background-color: #e6f7ff; }
        """)
        self.btn_label_settings.clicked.connect(self._open_label_settings_dialog)
        bl.addWidget(self.btn_label_settings)
        apply_toolbar_button_icon(self.btn_view_mode, "distribution_icon.svg", icon_px=24, icon_only=True, button_px=28)
        self.btn_view_mode.setToolTip("分布图：切换并配置分布图显示模式。")
        apply_toolbar_button_icon(self.btn_label_settings, "text_icon.svg", icon_px=24, icon_only=True, button_px=28)
        self.btn_label_settings.setToolTip("文本设置：调整图名、轴标题和刻度文本样式。")

        bl.addSpacing(10)
        bl.addStretch()

        sep = QFrame()
        sep.setFrameShape(QFrame.VLine)
        sep.setFixedHeight(16)
        sep.setStyleSheet("color:#ccc;")
        bl.addWidget(sep)

        self.btn_export = QPushButton("导出")
        self.btn_export.setFixedHeight(20)
        self.btn_export.setStyleSheet("""
                    QPushButton { background:#0050b3; color:white; font-weight:bold; border:none; border-radius: 3px; }
                    QPushButton:hover { background:#40a9ff; }
                    QPushButton:pressed { background:#003a8c; }
                """)
        self.btn_export.clicked.connect(self._on_export_clicked)
        bl.addWidget(self.btn_export)

        self.btn_insert = QPushButton("插入")
        self.btn_insert.setFixedHeight(20)
        self.btn_insert.setStyleSheet("""
                    QPushButton { background:#0050b3; color:white; font-weight:bold; border:none; border-radius: 3px; }
                    QPushButton:hover { background:#40a9ff; }
                    QPushButton:disabled { background:#d9d9d9; color:#999; }
                """)
        self.btn_insert.clicked.connect(self.on_accept)
        bl.addWidget(self.btn_insert)

        self.btn_close = QPushButton("关闭")
        self.btn_close.setFixedHeight(20)
        self.btn_close.setStyleSheet("""
                    QPushButton { background:#555555; color:white; font-weight:bold; border:none; border-radius: 3px; }
                    QPushButton:hover { background:#777777; }
                    QPushButton:pressed { background:#333333; }
                """)
        self.btn_close.clicked.connect(self.reject)
        bl.addWidget(self.btn_close)

        self.combo_y = QComboBox()
        self.combo_y.setVisible(False)

        main_layout.addWidget(bottom_bar)

    # -------------------------------------------------------
    # 4.2 后端数学向量化安全接口 (Backend Vectorized Math Safe Wrappers)
    # 交接说明：所有的数学引擎调用都必须经过这些 Wrapper。
    # 它们将零散的报错封装掉，并强制转为 1D numpy 数组，
    # 彻底终结因返回标量或 None 导致底层 Plotly JSON 序列化翻车的致命 Bug。
    # -------------------------------------------------------
    def _safe_pdf_vec(self, dist, xs):
        if hasattr(dist, "pdf_vec"):
            res = dist.pdf_vec(xs)
        elif hasattr(dist, "pdf"):
            try: res = dist.pdf(xs)
            except: res = np.array([float(dist.pdf(x)) for x in xs])
        else:
            res = np.zeros_like(xs)
            
        # 强制转换为 1D numpy 数组防崩
        res = np.asarray(res, dtype=float)
        if res.ndim == 0:
            res = np.full_like(xs, res)
        # 过滤非有限数值为 0
        return np.where(np.isfinite(res), res, 0.0)

    def _safe_pmf_vec(self, dist, xs):
        if hasattr(dist, "pmf_vec"):
            res = dist.pmf_vec(xs)
        elif hasattr(dist, "pmf"):
            try: res = dist.pmf(xs)
            except: res = np.array([float(dist.pmf(x)) for x in xs])
        elif hasattr(dist, "pdf_vec"):
            res = dist.pdf_vec(xs)
        else:
            res = np.zeros_like(xs)
            
        res = np.asarray(res, dtype=float)
        if res.ndim == 0:
            res = np.full_like(xs, res)
        return np.where(np.isfinite(res), res, 0.0)

    def _safe_cdf_vec(self, dist, xs):
        if hasattr(dist, "cdf_vec"):
            res = dist.cdf_vec(xs)
        elif hasattr(dist, "cdf"):
            try: res = dist.cdf(xs)
            except: res = np.array([float(dist.cdf(x)) for x in xs])
        else:
            res = np.zeros_like(xs)
            
        res = np.asarray(res, dtype=float)
        if res.ndim == 0:
            res = np.full_like(xs, res)
        return np.where(np.isfinite(res), res, 0.0)

    def _safe_cdf_scalar(self, dist, x):
        res = self._safe_cdf_vec(dist, np.array([float(x)]))
        return float(res[0]) if len(res) > 0 else 0.0

    def _safe_ppf_scalar(self, dist, p):
        if hasattr(dist, "inv_cdf"):
            return float(dist.inv_cdf(p))
        if hasattr(dist, "ppf"):
            return float(dist.ppf(p))
        return float('nan')

    def _current_dist_family_key(self) -> str:
        if hasattr(self, "combo_dist_family") and self.combo_dist_family is not None:
            return str(self.combo_dist_family.currentData() or _DIST_FAMILY_CONTINUOUS)
        return _DIST_FAMILY_CONTINUOUS

    def _sync_dist_family_with_key(self, dist_key: str) -> None:
       cfg = get_dist_config(dist_key) or {}
       if dist_key in _SPECIAL_DISTRIBUTION_KEYS:
           target_family = _DIST_FAMILY_SPECIAL
       else:
           target_family = _DIST_FAMILY_DISCRETE if bool(cfg.get("is_discrete", False)) else _DIST_FAMILY_CONTINUOUS

       if hasattr(self, "combo_dist_family") and self.combo_dist_family is not None:
            idx = self.combo_dist_family.findData(target_family)
            if idx >= 0 and self.combo_dist_family.currentIndex() != idx:
                # 阻塞信号，避免触发无意义的模型重绘循环
                self.combo_dist_family.blockSignals(True)
                self.combo_dist_family.setCurrentIndex(idx)
                self.combo_dist_family.blockSignals(False)
       if hasattr(self, "combo") and self.combo is not None:
            self.combo.blockSignals(True)
            try:
                _populate_distribution_combo_by_family(self.combo, target_family, dist_key)
            finally:
                self.combo.blockSignals(False)

    def on_dist_family_changed(self, idx):
        family_key = self._current_dist_family_key()
        selected_key = ""
        if self.dist_type in DIST_REGISTRY:
           if family_key == _DIST_FAMILY_RECENT:
              same_family = self.dist_type in set(_get_distribution_keys_by_family(_DIST_FAMILY_RECENT))
           elif family_key == _DIST_FAMILY_SPECIAL:
              same_family = self.dist_type in set(_get_distribution_keys_by_family(_DIST_FAMILY_SPECIAL))
           else:
              same_family = (
                self.dist_type not in _SPECIAL_DISTRIBUTION_KEYS
                and bool(DIST_REGISTRY[self.dist_type].get("is_discrete", False))
                == _distribution_family_to_is_discrete(family_key)
    )

           if same_family:
                selected_key = self.dist_type

        self.combo.blockSignals(True)
        _populate_distribution_combo_by_family(self.combo, family_key, selected_key)
        self.combo.blockSignals(False)

        if self.combo.count() <= 0:
            return
        new_key = self.combo.currentData()
        if new_key and new_key != self.dist_type:
            self.on_model_changed(self.combo.currentIndex())

    def _record_recent_distribution_use(self, dist_key: str):
        _record_recent_distribution_use(dist_key)

    # -------------------------------------------------------
    # 4.3 高级形态：截断与平移交互 (Advanced Morphology: Truncation & Shift)
    # 交接说明：该模块动态生成和销毁左侧参数面板中的截断/平移输入框。
    # 特别注意 combo_shift 和 combo_trunc 之间的互斥与联机逻辑。
    # -------------------------------------------------------
    def setup_advanced_morphology_ui(self, parent_layout):
        """建立平移与截断（高级形态）UI及其表单交互"""
        self.gb_morph = QGroupBox("平移与截断")
        morph_layout = QVBoxLayout(self.gb_morph)
        
        # 1. 顶部开关按钮
        self.btn_adv_morph = QPushButton("展开设置 ▼")
        self.btn_adv_morph.setCheckable(True)
        self.btn_adv_morph.setStyleSheet("""
            QPushButton { 
                border: 1px solid #d9d9d9; 
                border-radius: 4px; 
                background: #f0f0f0; 
                padding: 4px; 
                font-size: 11px;  /* 🟢 强制锁定 11px */
                font-weight: normal;
                font-family: 'Microsoft YaHei';
            }
            QPushButton:checked { 
                background: #e6f7ff; 
                border-color: #1890ff; 
                color: #1890ff; 
            }
        """)

        self.btn_adv_morph.toggled.connect(self._toggle_adv_panel)
        morph_layout.addWidget(self.btn_adv_morph)
        
        # 2. 下拉选项控制面板
        self.adv_options_panel = QWidget()
        options_layout = QFormLayout(self.adv_options_panel)
        options_layout.setContentsMargins(0, 5, 0, 0)
        
        self.combo_shift = QComboBox()
        self.combo_shift.addItems(["无平移", "平移"])
        
        self.combo_trunc = QComboBox()
        self.combo_trunc.addItems(["无截断", "值截断", "分位数截断"])
        
        options_layout.addRow("平移  ", self.combo_shift)
        options_layout.addRow("截断  ", self.combo_trunc)
        morph_layout.addWidget(self.adv_options_panel)
        
        # 3. 动态输入框容器
        self.adv_inputs_container = QWidget()
        self.adv_inputs_layout = QVBoxLayout(self.adv_inputs_container)
        self.adv_inputs_layout.setContentsMargins(0, 5, 0, 0)
        self.adv_inputs_layout.setSpacing(8)
        
        # --- [平移模块] ---
        self.shift_container = QWidget()
        self.shift_form = QFormLayout(self.shift_container)
        self.shift_form.setContentsMargins(0, 0, 0, 0)
        self.shift_form.setSpacing(10)
        
        self.input_shift = AutoSelectLineEdit()
        self.input_shift.setPlaceholderText("输入平移值...")
        self.input_shift.editingFinished.connect(self.on_param_changed)
        self.shift_form.addRow("平移值", self.input_shift)
        
        # --- [截断模块] ---
        self.trunc_container = QWidget()
        self.trunc_form = QFormLayout(self.trunc_container)
        self.trunc_form.setContentsMargins(0, 0, 0, 0)
        self.trunc_form.setSpacing(10)
        
        self.label_trunc1 = QLabel("最小值")
        self.input_trunc1 = AutoSelectLineEdit()
        self.input_trunc1.editingFinished.connect(self.on_param_changed)
        
        self.label_trunc2 = QLabel("最大值")
        self.input_trunc2 = AutoSelectLineEdit()
        self.input_trunc2.editingFinished.connect(self.on_param_changed)
        
        self.trunc_form.addRow(self.label_trunc1, self.input_trunc1)
        self.trunc_form.addRow(self.label_trunc2, self.input_trunc2)
        
        morph_layout.addWidget(self.adv_inputs_container)
        parent_layout.addWidget(self.gb_morph)
        
        # 绑定核心联动信号
        self.combo_shift.currentTextChanged.connect(self._on_shift_mode_changed)
        self.combo_trunc.currentTextChanged.connect(self._update_adv_inputs_layout)
        
        # 初始化状态隐藏
        self.adv_options_panel.setVisible(False)
        self.adv_inputs_container.setVisible(False)

    def _toggle_adv_panel(self, checked):
        """展开/收起平移截断控制面板。"""
        self.adv_options_panel.setVisible(checked)
        self.btn_adv_morph.setText("收起设置 ▲" if checked else "展开设置 ▼")
        shift_mode = self.combo_shift.currentText()
        trunc_mode = self.combo_trunc.currentText()
        self.adv_inputs_container.setVisible(checked and (shift_mode == "平移" or trunc_mode != "无截断"))
    
    def _on_shift_mode_changed(self, text):
        """根据平移模式自动更新截断模式候选项（处理截断与平移的计算顺序）。"""
        self.combo_trunc.blockSignals(True)
        current_trunc = self.combo_trunc.currentText()
        self.combo_trunc.clear()
        
        if text == "平移":
            self.combo_trunc.addItems(["无截断", "值截断后平移", "平移后值截断", "分位数截断"])
            if current_trunc in ["值截断", "值截断后平移"]:
                self.combo_trunc.setCurrentText("值截断后平移")
            elif current_trunc == "分位数截断":
                self.combo_trunc.setCurrentText("分位数截断")
            else:
                self.combo_trunc.setCurrentText("无截断")
        else:
            self.combo_trunc.addItems(["无截断", "值截断", "分位数截断"])
            if current_trunc in ["值截断", "值截断后平移", "平移后值截断"]:
                self.combo_trunc.setCurrentText("值截断")
            elif current_trunc == "分位数截断":
                self.combo_trunc.setCurrentText("分位数截断")
            else:
                self.combo_trunc.setCurrentText("无截断")
                
        self.combo_trunc.blockSignals(False)
        self._update_adv_inputs_layout()
        self.on_param_changed() # 触发曲线更新
    
    def _update_adv_inputs_layout(self, *args):
        """动态更新高级输入框布局并应用验证器，防呆用户的非法输入。"""
        shift_mode = self.combo_shift.currentText()
        trunc_mode = self.combo_trunc.currentText()
        show_shift = (shift_mode == "平移")
        show_trunc = (trunc_mode != "无截断")
        
        if show_trunc:
            if "分位数" in trunc_mode:
                self.label_trunc1.setText("最小分位")
                self.label_trunc2.setText("最大分位")
                self.input_trunc1.setPlaceholderText("0.00~1.00")
                self.input_trunc2.setPlaceholderText("0.00~1.00")
                self.input_trunc1.setToolTip("请输入 0 到 1 之间的小数，留空代表不截断下界")
                self.input_trunc2.setToolTip("请输入 0 到 1 之间的小数，留空代表不截断上界")
                
                # 设定只能输入 0~1 数字的严格验证器
                validator = QDoubleValidator(0.0, 1.0, 6)
                validator.setNotation(QDoubleValidator.StandardNotation)
                self.input_trunc1.setValidator(validator)
                self.input_trunc2.setValidator(validator)
            else:
                self.label_trunc1.setText("最小值")
                self.label_trunc2.setText("最大值")
                self.input_trunc1.setPlaceholderText("Min")
                self.input_trunc2.setPlaceholderText("Max")
                self.input_trunc1.setToolTip("请输入截断最小值，留空代表使用原始下界")
                self.input_trunc2.setToolTip("请输入截断最大值，留空代表使用原始上界")
                
                # 解除验证器限制，允许输入任意实数
                self.input_trunc1.setValidator(None)
                self.input_trunc2.setValidator(None)
                
        # 安全清空当前外层排序布局，准备重新注入
        while self.adv_inputs_layout.count():
            item = self.adv_inputs_layout.takeAt(0)
            if item.widget():
                item.widget().setParent(None)
                
        # 重新按照逻辑插入模块容器，实现UI上下掉换（值截断后平移 vs 平移后值截断）
        if trunc_mode == "值截断后平移":
            if show_trunc: self.adv_inputs_layout.addWidget(self.trunc_container)
            if show_shift: self.adv_inputs_layout.addWidget(self.shift_container)
        else:
            if show_shift: self.adv_inputs_layout.addWidget(self.shift_container)
            if show_trunc: self.adv_inputs_layout.addWidget(self.trunc_container)
            
        # 设置模块显示状态
        self.shift_container.setVisible(show_shift)
        self.trunc_container.setVisible(show_trunc)
        
        is_expanded = self.btn_adv_morph.isChecked()
        self.adv_inputs_container.setVisible(is_expanded and (show_shift or show_trunc))
        
        if args: self.on_param_changed()

    # -------------------------------------------------------
    # 4.4 Web/Plotly 渲染同步与视图控制 (Web/Plotly Sync & View Control)
    # -------------------------------------------------------
    def eventFilter(self, obj, event):
        """事件拦截器：同步骨架屏大小以匹配容器尺寸。"""
        if obj == getattr(self, "center_widget", None) and event.type() == QEvent.Resize:
            if hasattr(self, "chart_skeleton") and self.chart_skeleton:
                self.chart_skeleton.setGeometry(self.center_widget.rect())
        return super().eventFilter(obj, event)

    def _install_chart_context_menu_hook(self):
        """安装图表区右键菜单钩子。"""
        if not hasattr(self, "web") or self.web is None:
            return
        self.web.setContextMenuPolicy(Qt.CustomContextMenu)
        try:
            self.web.customContextMenuRequested.disconnect(self._on_chart_context_menu_requested)
        except Exception:
            pass
        self.web.customContextMenuRequested.connect(self._on_chart_context_menu_requested)

    def _on_chart_context_menu_requested(self, pos):
        """处理图表右键菜单的弹出与功能绑定。"""
        source = self.sender()
        global_pos = source.mapToGlobal(pos) if source is not None else self.mapToGlobal(pos)
        menu = QMenu(self)
        menu.setStyleSheet(
            "QMenu { background: white; border: 1px solid #d9d9d9; } "
            "QMenu::item:selected { background: #e6f7ff; color: #0050b3; }"
        )

        act_view = menu.addAction("视图")
        act_text = menu.addAction("文本设置")
        act_export = menu.addAction("导出")
        act_insert = menu.addAction("插入")

        act_view.triggered.connect(self.btn_view_mode.click)
        act_text.triggered.connect(self.btn_label_settings.click)
        act_export.triggered.connect(self._on_export_clicked)
        act_insert.triggered.connect(self.btn_insert.click)
        menu.exec(global_pos)

    def _on_export_clicked(self):
        """调用导出管理器，将当前图像导出。"""
        if hasattr(self, "web") and self.web:
            self.web.setFocus()
        try:
            export_ctx = DriskExportManager.build_export_context(
                self,
                target_widget=getattr(self, "center_widget", self),
                web_view=getattr(self, "web", None),
                chart_mode=str(getattr(self, "_current_view_data", "pdf")),
                current_key=str(getattr(self, "dist_type", "modeler")),
            )
            DriskExportManager.export_from_dialog(self, context=export_ctx)
        except Exception as e:
            QMessageBox.critical(self, "导出失败", f"图表导出发生错误：\n{str(e)}")

    def _show_skeleton(self, text: str = "正在绘图..."):
        """显示图表加载骨架屏（遮罩层），用于掩盖JS渲染延迟。"""
        if hasattr(self, "chart_skeleton") and self.chart_skeleton:
            try:
                self.chart_skeleton.lbl.setText(text)
                if hasattr(self, "center_widget") and self.center_widget:
                    self.chart_skeleton.setGeometry(self.center_widget.rect())
                self.chart_skeleton.raise_() 
                self.chart_skeleton.show()
            except Exception:
                pass

    def _show_view_mode_menu(self):
        """弹出建模器层级视图菜单，响应底部栏按钮。"""
        menu = QMenu(self)
        menu.setStyleSheet("QMenu { background: white; border: 1px solid #d9d9d9; } QMenu::item:selected { background: #e6f7ff; color: #0050b3; }")
        
        if getattr(self, 'is_discrete', False):
            m_pmf = menu.addMenu("离散概率")
            m_pmf.addAction("不叠加").triggered.connect(lambda: self._trigger_view_cmd("离散概率", "pmf"))
            m_pmf.addAction("叠加累积概率").triggered.connect(lambda: self._trigger_view_cmd("离散概率", "pmf_cdf"))
            menu.addAction("累积概率").triggered.connect(lambda: self._trigger_view_cmd("累积概率", "cdf"))
        else:
            m_pdf = menu.addMenu("概率密度")
            m_pdf.addAction("不叠加").triggered.connect(lambda: self._trigger_view_cmd("概率密度", "pdf"))
            m_pdf.addAction("叠加累积概率").triggered.connect(lambda: self._trigger_view_cmd("概率密度", "pdf_cdf"))
            menu.addAction("相对频率").triggered.connect(lambda: self._trigger_view_cmd("相对频率", "rel_freq"))
            menu.addAction("累积概率").triggered.connect(lambda: self._trigger_view_cmd("累积概率", "cdf"))

        menu.exec(self.btn_view_mode.mapToGlobal(self.btn_view_mode.rect().bottomLeft()))

    def _trigger_view_cmd(self, label, cmd):
        """同步视图按钮命令并触发重绘。"""
        self._current_view_data = cmd
        self.on_preview_mode_changed()

    def _resolve_label_settings_y_title(self) -> str:
        """根据当前图表模式自动推导 Y 轴应显示的默认标题。"""
        mode = str(getattr(self, "_current_view_data", "pdf") or "pdf").lower()
        if mode == "cdf":
            return "累积概率"
        if mode in ("pmf", "pmf_cdf"):
            return "概率"
        if mode == "rel_freq":
            return "相对频率 (%)"
        return "密度"

    def _resolve_modeler_metadata_chart_title(self) -> str:
        """从元数据、周边单元格等线索解析出图表的底板标题。"""
        attrs = {}
        markers = getattr(self, "_dist_markers", None)
        if isinstance(markers, dict):
            attrs.update(markers)
        elif isinstance(getattr(self, "initial_attrs", None), dict):
            attrs.update(getattr(self, "initial_attrs", {}) or {})

        raw_cell_key = str(getattr(self, "cell_address", "") or "").strip()
        fallback_label = str(getattr(self, "_cell_detected_display_hint", "") or "").strip()
        if not fallback_label:
            fallback_label = raw_cell_key if raw_cell_key else "新建建模变量"

        xl_ctx = None
        try:
            xl_ctx = _safe_excel_app()
        except Exception:
            xl_ctx = None

        resolved = resolve_visible_variable_name(
            raw_cell_key or fallback_label,
            attrs,
            excel_app=xl_ctx,
            fallback_label=fallback_label,
        )
        text = str(resolved or "").strip()
        if text:
            return text
        return fallback_label

    def _get_modeler_chart_title_display_override(self) -> Optional[str]:
        raw = getattr(self, "_chart_title_display_override", None)
        if raw is None:
            return None
        text = str(raw).strip()
        return text if text else None

    def _set_modeler_chart_title_display_override(self, title_text: Optional[str]) -> None:
        text = str(title_text or "").strip()
        self._chart_title_display_override = text if text else None

    def _resolve_modeler_display_chart_title(self) -> str:
        override = self._get_modeler_chart_title_display_override()
        if override:
            return override
        return self._resolve_modeler_metadata_chart_title()

    def _resolve_modeler_chart_title_font_from_config(self) -> tuple[str, float]:
        """从标签设置配置中解析建模器图表标题的字体系列/大小"""
        cfg = getattr(self, "_label_settings_config", None)
        chart_cfg = cfg.get("chart_title", {}) if isinstance(cfg, dict) else {}

        family = str(chart_cfg.get("font_family", "") or "").strip() or "Arial"
        try:
            size = float(chart_cfg.get("font_size")) if chart_cfg.get("font_size") is not None else 14.0
        except Exception:
            size = 14.0
        if size <= 0:
            size = 14.0
        return family, size

    def _apply_modeler_chart_title_label_style(self) -> None:
        """将文本设置字体选择应用于建模器可见的图表标题标签"""
        if not hasattr(self, "chart_title_label") or self.chart_title_label is None:
            return
        family, size = self._resolve_modeler_chart_title_font_from_config()
        safe_family = family.replace("\\", "\\\\").replace('"', '\\"')
        self.chart_title_label.setStyleSheet(
            "color: #333333; "
            "margin: 0px; "
            "padding: 0px; "
            f"font-size: {size:.1f}px; "
            "font-weight: normal; "
            f'font-family: "{safe_family}", "Microsoft YaHei", sans-serif;'
        )

    def _build_label_settings_context(self) -> dict:
        """构建图表标签设置的上下文配置（传递给文本设置对话框）。"""
        x_mag_recommended = infer_si_mag(
            value_range=(float(getattr(self, "view_min", 0.0)), float(getattr(self, "view_max", 1.0))),
            dtick=float(getattr(self, "x_dtick", 1.0) or 1.0),
            forced_mag=None,
        )
        chart_default_text = self._resolve_modeler_metadata_chart_title()
        return {
            "chart_title": {
                "default_text": chart_default_text,
                "default_font_family": "Arial",
                "default_font_size": 14,
            },
            "axes": {
                "x": {
                    "default_title": DriskChartFactory.VALUE_AXIS_TITLE,
                    "default_title_font_family": "Arial",
                    "default_title_font_size": 12,
                    "default_tick_font_family": "Arial",
                    "default_tick_font_size": 12,
                    "recommended_mag": int(x_mag_recommended),
                    "data_candidates": [
                        {"key": "x_main", "label": "当前X数据", "is_numeric": True},
                    ],
                },
                "y": {
                    "default_title": self._resolve_label_settings_y_title(),
                    "default_title_font_family": "Arial",
                    "default_title_font_size": 12,
                    "default_tick_font_family": "Arial",
                    "default_tick_font_size": 12,
                    "recommended_mag": 0,
                    "data_candidates": [
                        {"key": "y_main", "label": "当前Y数据", "is_numeric": True},
                    ],
                },
            },
        }

    def _open_label_settings_dialog(self):
        """打开文本设置对话框，并将结果应用到 UI 标题与 Plotly 配置。"""
        try:
            context = self._build_label_settings_context()
            dlg = LabelSettingsDialog(
                config=getattr(self, "_label_settings_config", None),
                context=context,
                parent=self,
            )
            if dlg.exec():
                new_cfg = dlg.get_config()
                metadata_title_text = self._resolve_modeler_metadata_chart_title()
                chart_cfg = new_cfg.get("chart_title", {}) if isinstance(new_cfg.get("chart_title", {}), dict) else {}
                chart_title_delta = chart_cfg.get("text_override", None) if isinstance(chart_cfg, dict) else None
                title_changed = False
                if chart_title_delta is not None:
                    desired_title = str(chart_title_delta or "").strip()
                    if (not desired_title) or desired_title == metadata_title_text:
                        self._set_modeler_chart_title_display_override(None)
                    else:
                        self._set_modeler_chart_title_display_override(desired_title)
                    title_changed = True
                if isinstance(chart_cfg, dict):
                    chart_cfg["text_override"] = None
                if title_changed and hasattr(self, "chart_title_label"):
                    shown_title = self._resolve_modeler_display_chart_title()
                    self.chart_title_label.setText(shown_title)
                self._label_settings_config = new_cfg
                self._apply_modeler_chart_title_label_style()
                self._label_axis_numeric = get_axis_numeric_flags(
                    new_cfg,
                    fallback={"x": True, "y": True},
                )
                self.manual_mag = get_axis_display_unit_override(new_cfg, "x")
                self.render_preview()
                self._update_floating_edits_pos()
        except Exception as exc:
            QMessageBox.critical(self, "错误", f"打开文本设置失败：{exc}")

    def _on_global_mag_changed(self, mag_int):
        """处理 X 轴量级 (Mag) 手动切换。"""
        try:
            self.manual_mag = int(mag_int)
            self.render_preview()
            self._update_floating_edits_pos()
        except Exception as e:
            pass

    def get_preview_mode(self) -> str:
        """获取绘图模式 (PDF/PMF/CDF等)。"""
        return getattr(self, "_current_view_data", "pdf" if not self.is_discrete else "pmf")

    def _debug_dump_web_state(self, tag=""):
        """诊断 Webview JS 环境与 Plotly 实例状态。"""
        if not DEBUG_WEB:
            return

        js = r"""
        (function(){
        const info = {};
        info.tag = %s;
        info.href = location.href;
        info.readyState = document.readyState;
        info.hasPlotly = (typeof Plotly !== 'undefined');
        const gd = document.querySelector('.js-plotly-plot');
        info.hasGD = !!gd;
        info.hasFullLayout = !!(gd && gd._fullLayout);
        info.bodyChildren = document.body ? document.body.children.length : -1;
        info.viewport = {w: window.innerWidth, h: window.innerHeight};
        return info;
        })();
        """ % (repr(tag))

        self.web.page().runJavaScript(js, lambda v: print("[WEB_STATE]", v))

    def _on_webview_load_finished(self, ok: bool):
        """将 Webview 加载完成的生命周期通过对话框转接：处理骨架屏消隐与后续同步。"""
        try:
            if getattr(self, "_plot_host", None):
                self._plot_host.on_webview_load_finished(ok)
        except Exception:
            pass

        try:
            if hasattr(self, "_chart_ctrl"):
                self._chart_ctrl.schedule_rect_sync()
        except Exception:
            pass

        self._webview_loaded = bool(ok)
        if not ok:
            self._show_skeleton("绘图区加载失败（请重试）")
            return

        self._maybe_hide_skeleton()
        self._after_plot_ready_visual_sync()

    def _maybe_hide_skeleton(self):
        """使用令牌 (Token) /截止期限 (Deadline) 检测机制，确保不会因为过期的渲染请求错误地隐藏骨架屏。"""
        if not (self._webview_loaded and self._data_sent):
            return
        self._waiting_token = self._render_token
        self._ready_deadline = time.monotonic() + 10.0
        if not self._ready_poll_timer.isActive():
            self._ready_poll_timer.start()

    def _poll_plotly_ready(self):
        """持续轮询 JavaScript，直到当前的渲染令牌对应的 Plotly 布局稳定，方可移除遮罩层。"""
        if self._waiting_token != self._render_token:
            self._ready_poll_timer.stop()
            return

        if time.monotonic() > self._ready_deadline:
            self._show_skeleton("绘图超时（请重试）")
            self._ready_poll_timer.stop()
            return

        if self._js_inflight:
            return
        self._js_inflight = True

        js = r"""
        (function(){
          try{
            const gd = document.querySelector('.js-plotly-plot');
            return !!(gd && gd._fullLayout && gd._fullLayout.width && gd._fullLayout.height);
          }catch(e){ return false; }
        })();
        """

        def _cb(val):
            self._js_inflight = False
            if self._waiting_token != self._render_token:
                return
            if bool(val):
                self._ready_poll_timer.stop()
                if hasattr(self, "chart_skeleton") and self.chart_skeleton:
                    try:
                        self.chart_skeleton.hide()
                    except Exception:
                        pass
                self._after_plot_ready_visual_sync()

        try:
            self.web.page().runJavaScript(js, _cb)
        except Exception:
            self._js_inflight = False

    def _flow_get_dist_registry(self):
        return DIST_REGISTRY

    def _flow_get_dist_config(self, key):
        return get_dist_config(key)

    def _flow_get_dist_key_by_func_name(self, name):
        return get_dist_key_by_func_name(name)

    def _flow_find_first_distribution_in_formula(self, formula):
        return find_first_distribution_in_formula(formula)

    def _flow_extract_all_attributes_from_formula(self, formula):
        return extract_all_attributes_from_formula(formula)

    def on_stats_cell_clicked(self, row, col):
        """处理统计面板表格单元格点击（展示长内容，防止被截断）。"""
        item = self.tbl.item(row, col)
        if not item: return
        text = item.text().strip()
        if not text: return
        if len(text) > 18:
            QMessageBox.information(self, "完整内容", text)

    def _get_plot_pixel_width_for_snap(self) -> int:
        """获取图表区域的实际像素宽度，用于计算拖动捕捉的步长。"""
        try:
            w = int(self.slider.width())
        except Exception:
            w = 800

        l = self.slider.margin_left
        r = self.slider.margin_right
        return max(100, w - int(l) - int(r))

    def _rebuild_drag_snap_grid(self):
        """重构拖动时捕捉值的网格间距 (Snap Grid)。"""
        if self.is_discrete: return
        xmin = float(getattr(self, "view_min", 0.0))
        xmax = float(getattr(self, "view_max", 1.0))
        if xmax <= xmin:
            self._drag_grid_step = None
            self._drag_grid_xmin = None
            self._drag_grid_xmax = None
            return

        px_w = self._get_plot_pixel_width_for_snap()
        self._drag_grid_step = SnapUtils.calc_grid_step(xmin, xmax, px_w, px_per_step=2)
        self._drag_grid_xmin = xmin
        self._drag_grid_xmax = xmax

    def _snap_value_drag(self, x: float) -> float:
        """在用户拖拽滑块时，将连续的浮点值吸附 (Snap) 到合理的整齐刻度上。"""
        x = float(x)
        if self.is_discrete:
            # 离散分布复用连续分布的拖拽取值逻辑。
            return x

        step = getattr(self, "_drag_grid_step", None)
        gxmin = getattr(self, "_drag_grid_xmin", None)
        gxmax = getattr(self, "_drag_grid_xmax", None)

        xmin = float(getattr(self, "view_min", 0.0))
        xmax = float(getattr(self, "view_max", 1.0))

        stale = (step is None) or (gxmin is None) or (gxmax is None) or (step <= 0)
        if not stale:
            # 防止因坐标范围重算导致捕捉网格陈旧。
            tol = max(abs(xmin), abs(xmax), abs(float(gxmin)), abs(float(gxmax)), 1.0) * 1e-12
            stale = (abs(float(gxmin) - xmin) > tol) or (abs(float(gxmax) - xmax) > tol)

        if stale:
            self._rebuild_drag_snap_grid()
            step = getattr(self, "_drag_grid_step", None)
            gxmin = getattr(self, "_drag_grid_xmin", None)

        if step is None or gxmin is None or step <= 0: return x
        return SnapUtils.snap_to_grid(x, gxmin, step)

    # -------------------------------------------------------
    # 4.5 核心物理与数学引擎调度 (Core Physics & Math Engine Recalculation)
    # 交接说明：这是整个分布视图的"心脏"。它决定了图表的边界、步长以及是否因为极端的截断而抛出异常。
    # -------------------------------------------------------
    def recalc_distribution(self):
        """
        [核心引擎]：计算分布边界，确定图表显示范围、离散化步长，以及处理截断边界对有效概率区域的影响。
        """
        try:
            if not self._prepare_recalc_orchestration():
                return

            def _finite(x) -> bool:
                try:
                    return x is not None and math.isfinite(float(x))
                except Exception:
                    return False

            def _safe_ppf(q: float):
                """安全的百分位提取，带有重试与容错。"""
                qs = [q, 0.01, 0.05, 0.1, 0.9, 0.95, 0.99]
                for qq in qs:
                    try:
                        v = self._safe_ppf_scalar(self.dist_obj, qq)
                        if _finite(v):
                            return v
                    except Exception:
                        pass
                return None

            # ==========================================================
            # 🚀 智能边界与极端截断防御机制 (Smart Boundary & Extreme Truncation Defense)
            # 交接说明：这是整个绘图引擎最脆弱的地方。如果用户设置了极其极端的截断范围，
            # 会导致分布概率“坍缩”（即有效概率极小），从而引发底层数学计算返回 NaN 或 Inf。
            # ==========================================================
            # 判断用户是否实际配置了截断参数，绝不误伤无辜分布
            has_trunc = any(k in self._dist_markers for k in ['truncate', 'truncatep', 'truncate2', 'truncatep2'])
            
            try:
                # 1. 只有在实际启用了截断时，才执行严苛的坍缩审查
                if has_trunc:
                    p25 = self._safe_ppf_scalar(self.dist_obj, 0.25)
                    p50 = self._safe_ppf_scalar(self.dist_obj, 0.5)
                    p75 = self._safe_ppf_scalar(self.dist_obj, 0.75)
                    
                    if not (_finite(p25) and _finite(p50) and _finite(p75)):
                        raise ValueError("分位数函数(PPF)退化，返回了 NaN 或 Inf")
                        
                    # 探测 IQR 坍缩 (排除离散分布，因其 IQR 可能天然为 0)
                    if not self.is_discrete and abs(p75 - p25) <= 1e-12:
                        raise ValueError("分布已坍缩为奇点 (四分位距 IQR 接近 0)")

                    # 探测 CDF 下溢出 (排除离散分布，因其阶跃特性 CDF 无法精准落在 0.5)
                    if not self.is_discrete:
                        p50_cdf = self._safe_cdf_scalar(self.dist_obj, p50)
                        # 放宽到 0.2~0.8，只要不是极端 Underflow 导致的 0.0 或 NaN 即可放行
                        if not (0.2 <= p50_cdf <= 0.8): 
                            raise ValueError(f"累积概率函数(CDF)发生数学坍缩 (计算得出的 CDF={p50_cdf})")

                # 2. 提取核心概率区边界 (0.1% ~ 99.9%)
                core_low = self._safe_ppf_scalar(self.dist_obj, 0.01)
                core_high = self._safe_ppf_scalar(self.dist_obj, 0.99)
                
                # 降级容错
                if not _finite(core_low) or not _finite(core_high):
                    if has_trunc:
                        core_low = self._safe_ppf_scalar(self.dist_obj, 0.25)
                        core_high = self._safe_ppf_scalar(self.dist_obj, 0.75)
                    else:
                        core_low, core_high = -1.0, 1.0
                        
            except Exception as e:
                # 💡 核心修复：使用 QTimer 延后 100 毫秒异步弹窗！
                # 彻底解决 __init__ 期间因同步弹窗截断生命周期，导致父级 UI 闪退崩溃的终极 Bug。
                QTimer.singleShot(100, lambda: QMessageBox.warning(self, "截断参数无效", "您设置的截断区间位于极低概率的尾部（或区间过窄），\n导致该截断下的有效概率极度微小，分布已完全坍缩！\n\n请调整截断边界，使其靠近分布的主要概率集中区。"))
                self._show_skeleton("截断范围无效，无法绘图")
                self.tbl.setRowCount(0)
                return
                
            # 3. 探测绝对物理边界 (0% 和 100%)
            try:
                bound_low = self._safe_ppf_scalar(self.dist_obj, 0.0)
            except Exception:
                bound_low = float('-inf')
                
            try:
                bound_high = self._safe_ppf_scalar(self.dist_obj, 1.0)
            except Exception:
                bound_high = float('inf')

            if self.is_discrete:
                # 对离散定律，ppf(0) 可能返回一个低于真实支撑的哨兵值（如泊松分布返回 -1）
                try:
                    phys_min = float(self.dist_obj.min_val())
                    if math.isfinite(phys_min):
                        bound_low = max(float(bound_low), phys_min)
                except Exception:
                    pass
                try:
                    phys_max = float(self.dist_obj.max_val())
                    if math.isfinite(phys_max):
                        bound_high = min(float(bound_high), phys_max)
                except Exception:
                    pass

            core_span = core_high - core_low
            if core_span <= 1e-9:
                core_span = 1.0
                
            # 下界判定：如果存在非常接近核心区的物理下界（如对数正态的 0），则将其包含
            if _finite(bound_low) and bound_low >= core_low - 0.3 * core_span:
                raw_low = float(bound_low)
            else:
                raw_low = core_low  
                
            # 上界判定：同理
            if _finite(bound_high) and bound_high <= core_high + 0.3 * core_span:
                raw_high = float(bound_high)
            else:
                raw_high = core_high  

            # 确保最终范围合法
            if raw_low == raw_high:
                raw_low -= 1.0
                raw_high += 1.0

            if raw_low > raw_high:
                raw_low, raw_high = raw_high, raw_low

            span = raw_high - raw_low

            # 智能刻度推算
            self.x_dtick = float(DriskMath.calc_smart_step(span))
            if self.is_discrete:
                self.x_dtick = max(1.0, float(round(self.x_dtick)))

            def _align_floor(val, step):
                return math.floor(val / step) * step

            def _align_ceil(val, step):
                return math.ceil(val / step) * step

            # ✅ 增加 0.01*步长的呼吸感边距，防止边缘数据紧贴图框
            # 注意：raw_low 和 raw_high 已经是应用了截断和平移后的理论边界
            padded_low = raw_low - 0.01 * self.x_dtick
            padded_high = raw_high + 0.01 * self.x_dtick

            axis_min = _align_floor(padded_low, self.x_dtick)
            axis_max = _align_ceil(padded_high, self.x_dtick)

            discrete_align_step = 1.0
            if self.is_discrete:
                try:
                    if hasattr(self.dist_obj, "x_vals"):
                        xv = np.asarray(getattr(self.dist_obj, "x_vals", []), dtype=float)
                        xv = xv[np.isfinite(xv)]
                        if xv.size >= 2:
                            dv = np.diff(np.unique(np.sort(xv)))
                            dv = dv[dv > 1e-12]
                            if dv.size > 0:
                                discrete_align_step = float(np.min(dv))
                except Exception:
                    discrete_align_step = 1.0
                if not math.isfinite(discrete_align_step) or discrete_align_step <= 0:
                    discrete_align_step = 1.0

                # 仅限制极端的离散扩散，同时保留智能步长的边缘延展
                guard_step = max(float(self.x_dtick), float(discrete_align_step))
                min_guard = math.floor(raw_low / guard_step) * guard_step - guard_step
                max_guard = math.ceil(raw_high / guard_step) * guard_step + guard_step
                axis_min = max(axis_min, min_guard)
                axis_max = min(axis_max, max_guard)

            if self.is_discrete and raw_low >= 0 and axis_min < 0:
                axis_min = max(axis_min, -discrete_align_step)
            if self.is_discrete and raw_low > -0.5 * discrete_align_step and axis_min < -discrete_align_step:
                # 防呆：阻止极小的负浮点数噪声引发向外扩散一个多余离散步长的问题。
                axis_min = -discrete_align_step

            if axis_max <= axis_min:
                axis_max = axis_min + self.x_dtick

            if self.is_discrete:
                axis_min = math.floor(axis_min)
                axis_max = math.ceil(axis_max)

            self.view_min = float(axis_min)
            self.view_max = float(axis_max)

            self.manual_mag = get_axis_display_unit_override(getattr(self, "_label_settings_config", None), "x")
            self._label_axis_numeric = get_axis_numeric_flags(
                getattr(self, "_label_settings_config", None),
                fallback={"x": True, "y": True},
            )

            # ==========================================================
            # 🚀 数据采样生成区 (Data Sampling Generation)
            # ==========================================================
            if self.is_discrete:
                s = int(self.view_min)
                e = int(self.view_max)
                if e < s:
                    s, e = e, s
                n = (e - s + 1)
                MAX_PLOT_POINTS = 1000
                SNAP_MAX_POINTS = 5000
                self._discrete_sampling_note = ""

                if n <= MAX_PLOT_POINTS:
                    raw_xs = np.arange(s, e + 1, 1.0)
                else:
                    raw_xs = np.linspace(s, e, num=MAX_PLOT_POINTS)
                    raw_xs = np.round(raw_xs).astype(int)

                # 🔴 修复 3：将分布自带的真实自定义坐标（可能包含小数）强行并入探测网格！
                # 解决离散表包含如 1.5, 2.5 这种小数时被 int 强制抹去的问题。
                if hasattr(self.dist_obj, "x_vals"):
                    custom_xs = np.array(self.dist_obj.x_vals) + getattr(self.dist_obj, "shift_amount", 0.0)
                    custom_xs = custom_xs[(custom_xs >= self.view_min) & (custom_xs <= self.view_max)]
                    raw_xs = np.concatenate([raw_xs, custom_xs])
                    
                self.x_data = np.unique(np.sort(raw_xs)).astype(float)
                self.y_data = self._safe_pmf_vec(self.dist_obj, self.x_data)

                if n <= SNAP_MAX_POINTS:
                    self._snap_xs = np.arange(s, e + 1, 1.0)
                    self._snap_ys = self._safe_pmf_vec(self.dist_obj, self._snap_xs.astype(float))
                else:
                    self._snap_xs = None
                    self._snap_ys = None

                ql = None
                qr = None
                try:
                    ql = float(self._safe_ppf_scalar(self.dist_obj, 0.05))
                    qr = float(self._safe_ppf_scalar(self.dist_obj, 0.95))
                except Exception:
                    ql = None
                    qr = None

                if not _finite(ql):
                    ql = _safe_ppf(0.05)
                if not _finite(qr):
                    qr = _safe_ppf(0.95)

                if not _finite(ql) or not _finite(qr):
                    ql, qr = float(s), float(e)

                ql = float(round(ql))
                qr = float(round(qr))

                ql = max(float(s), min(float(e), ql))
                qr = max(float(s), min(float(e), qr))

                if ql > qr:
                    ql, qr = qr, ql
                if ql == qr:
                    ql = max(float(s), ql - 1.0)
                    qr = min(float(e), qr + 1.0)
                    if ql == qr:
                        ql, qr = float(s), float(e)

                self.curr_left = float(ql)
                self.curr_right = float(qr)

            else:
                # --- 连续分布数据生成 ---
                # 1. 基础均匀采样
                MAX_PLOT_POINTS = 800
                x_raw = np.linspace(self.view_min, self.view_max, num=MAX_PLOT_POINTS)
                
                # 2. [核心手术] 针对有界分布（Uniform/Triang/Trigen/Cumul）进行边界点硬性注入
                # 如果不这么做，直方图的直线或者三角的尖顶会被曲线平滑算法“削平”。
                extra_points = []
                try:
                    # 获取分布的物理边界
                    d_min = self.dist_obj.min_val()
                    d_max = self.dist_obj.max_val()
                    
                    # 🔴 修复 2：加入 Cumul，并向 Plotly 注入极其精确的边界点与内部转折点
                    if self.dist_type in ["Uniform", "Triang", "Trigen", "Cumul"]:
                        eps = (self.view_max - self.view_min) * 1e-3 if self.view_max > self.view_min else 1e-3
                        for val in [d_min - eps, d_min, d_min + eps, d_max - eps, d_max, d_max + eps]:
                            if self.view_min <= val <= self.view_max:
                                extra_points.append(val)
                        
                        # 专为 Cumul 注入内部断崖转折点，确保阶跃函数在图表上能直上直下！
                        if self.dist_type == "Cumul" and hasattr(self.dist_obj, "x_vals"):
                            shift = getattr(self.dist_obj, "shift_amount", 0.0)
                            for xv in self.dist_obj.x_vals:
                                val = xv + shift
                                if self.view_min <= val <= self.view_max:
                                    extra_points.extend([val - eps, val, val + eps])
                        
                        # 针对三角分布，还需注入众数点（Mode）以保住尖角
                        if hasattr(self.dist_obj, "mode"):
                            m = self.dist_obj.mode()
                            if self.view_min <= m <= self.view_max:
                                extra_points.append(m)
                except:
                    pass

                if extra_points:
                    self.x_data = np.unique(np.sort(np.concatenate([x_raw, extra_points])))
                else:
                    self.x_data = x_raw

                # 3. 计算 PDF 并清理越界概率
                self.y_data = self._safe_pdf_vec(self.dist_obj, self.x_data)
                # 防御底层分布（如均匀分布）的 PDF 在物理界外返回非零错误值的问题
                try:
                    d_min = float(self.dist_obj.min_val())
                    d_max = float(self.dist_obj.max_val())
                    # 强制将分布边界外的概率密度置为 0.0
                    self.y_data = np.where((self.x_data >= d_min) & (self.x_data <= d_max), self.y_data, 0.0)
                except Exception:
                    pass

                ql = None
                qr = None
                try:
                    ql = float(self._safe_ppf_scalar(self.dist_obj, 0.05))
                    qr = float(self._safe_ppf_scalar(self.dist_obj, 0.95))
                except Exception:
                    ql = None
                    qr = None

                if not _finite(ql):
                    ql = _safe_ppf(0.05)
                if not _finite(qr):
                    qr = _safe_ppf(0.95)

                if not _finite(ql) or not _finite(qr):
                    ql, qr = float(self.view_min), float(self.view_max)

                ql = max(float(self.view_min), min(float(self.view_max), float(ql)))
                qr = max(float(self.view_min), min(float(self.view_max), float(qr)))

                if ql > qr:
                    ql, qr = qr, ql
                if abs(qr - ql) <= 1e-12:
                    ql = float(self.view_min)
                    qr = float(self.view_max)

                self.curr_left = float(ql)
                self.curr_right = float(qr)

            self._apply_recalc_post_refresh()

        except Exception as e:
            pass
            self.tbl.setRowCount(0)

    def update_y_combo_items(self):
        """同步更新 Y 轴下拉列表框状态。"""
        text = "概率" if self.is_discrete else "密度"
        self.combo_y.blockSignals(True)
        if self.combo_y.count() == 0:
            self.combo_y.addItem(text)
        else:
            self.combo_y.setItemText(0, text)
        self.combo_y.setCurrentIndex(0)
        self.combo_y.blockSignals(False)


# =======================================================
# 5. 模块全局主入口 (Module Global Entry)
# =======================================================
def show_distribution_gallery(initial_dist=None, initial_params=None, initial_attrs=None, full_formula="", cell_address=""):
    """
    启动分布定义 UI 界面的全剧入口函数。
    包含智能上下文探测逻辑：如果缺少名称或种类，会自动向上/向左扫描周围的表头文本。
    
    参数说明:
    initial_dist: 初始分布类名 (e.g. "Normal")
    initial_params: 初始参数列表
    initial_attrs: 初始分布属性字典
    full_formula: 单元格完整公式字符串
    cell_address: 目标单元格地址 (用于命名溯源)
    """
    app = QApplication.instance() or QApplication(sys.argv)
    try:
        formula_text = str(full_formula or "").strip()
        is_makeinput_formula = bool(
            re.match(r"^\s*=\s*@?\s*DriskMakeInput\s*\(", formula_text, re.IGNORECASE)
        )

        # 将 DriskMakeInput 视为分布定义流程中未解析的占位符。
        # 这确保了“定义/编辑分布”行为与处理空/未定义单元格时保持一致。
        if is_makeinput_formula:
            initial_dist = None
            initial_params = None
            initial_attrs = {}
            full_formula = ""

        # =================================================================
        # 🚀 智能上下文探测：直接调用 dependency_tracker 单一信源寻找可能的命名
        # =================================================================
        if initial_attrs is None:
            initial_attrs = {}
        display_name_hint = ""
            
        if (not initial_attrs.get('name') or not initial_attrs.get('category')) and cell_address:
            try:
                xl_app = _safe_excel_app()
                if '!' in cell_address:
                    sheet_name, addr_str = cell_address.split('!', 1)
                    sheet = xl_app.ActiveWorkbook.Worksheets(sheet_name)
                else:
                    sheet = xl_app.ActiveSheet
                    addr_str = cell_address
                    
                cell = sheet.Range(addr_str)
                
                # 安全解包，现已精简为 2 个返回值，防止索引越界报错
                name_info = find_cell_name_ui(xl_app, sheet, cell)
                full_name = name_info[0] if len(name_info) > 0 else ""
                up_name = name_info[1] if len(name_info) > 1 else ""

                # 仅供显示的提示名，从周围单元格派生；不要将其写入元数据输入框。
                if full_name:
                    display_name_hint = full_name
                elif up_name:
                    display_name_hint = up_name
                    
            except ImportError:
                print("[UI Modeler] 无法导入依赖模块，跳过自动命名。")
            except Exception as e:
                print(f"[UI Modeler] 自动探测发生错误: {e}")
        # =================================================================

        # 🚀 [核心修复]：如果外部调用方（如老旧的 Excel 宏）没有传递 initial_dist，
        # 但传递了 full_formula（如 "=DriskTriang(10,13,15)"），Python 会在此处拦截并自行解析重构！
        # 在编辑现有单元格时，始终优先从 full_formula 重建。
        if isinstance(full_formula, str) and full_formula.strip().startswith("="):
            detected_key, detected_params = find_first_distribution_in_formula(full_formula)
            if detected_key:
                initial_dist = detected_key
                if detected_params:
                    initial_params = detected_params

            try:
                parsed_attrs = extract_all_attributes_from_formula(full_formula) or {}
                if parsed_attrs:
                    merged_attrs = dict(initial_attrs or {})
                    merged_attrs.update(parsed_attrs)
                    initial_attrs = merged_attrs
            except Exception:
                pass
                
        # 仅当当前单元格公式被解析为真实的分布函数时，才进入直接编辑模式。
        resolved_dist_key = None
        if initial_dist:
            dist_key = initial_dist
            try:
                # 将例如 "DriskNormal" 的函数名逆向映射为 UI 字典键 "Normal"。
                if isinstance(initial_dist, str) and initial_dist.upper().startswith("DRISK"):
                    k = get_dist_key_by_func_name(initial_dist.upper())
                    if k:
                        dist_key = k
            except Exception:
                pass

            if dist_key in DIST_REGISTRY:
                resolved_dist_key = dist_key

        # 场景 A：如果传入了可识别的初始分布（即“双击修改现有公式”），直接弹主面板
        if resolved_dist_key:
            dlg = DistributionBuilderDialog(
                resolved_dist_key,
                initial_params,
                initial_attrs,
                full_formula,
                cell_address,
                display_name_hint=display_name_hint,
            )
            if dlg.exec():
                return dlg.result_formula
            return None

        # 场景 B：如果没有传入初始分布（即“新建变量”），先弹出九宫格分布选择器
        sel = DistributionSelector()
        if sel.exec():
            try:
                dlg = DistributionBuilderDialog(
                    sel.selected_dist, 
                    initial_attrs=initial_attrs, 
                    full_formula=full_formula, 
                    cell_address=cell_address,
                    display_name_hint=display_name_hint
                )
                if dlg.exec():
                    return dlg.result_formula
            except Exception as e:
                err_msg = traceback.format_exc()
                QMessageBox.critical(None, "初始化错误", f"无法打开分布定义窗口:\n{str(e)}\n\n{err_msg}")

        return None

    except Exception as e:
        QMessageBox.critical(None, "运行错误", f"程序运行发生异常:\n{str(e)}")
        return None
