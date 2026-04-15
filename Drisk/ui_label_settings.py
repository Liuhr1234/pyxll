# ui_label_settings.py
"""
本模块提供图表文本与标签设置（Label Settings）的对话框界面及相关状态管理逻辑。

模块职责概述：
1. 配置归一化（Config Normalization）：
   将外部传入的用户自定义配置清洗为内部可稳定使用的标准结构，
   区分“用户主动覆盖的配置”与“图表运行时默认上下文”。

2. 上下文归一化（Context Normalization）：
   将底层图表传入的默认标题、默认字体、默认字号、推荐显示单位、
   可选数据序列等上下文信息整理为统一结构，供界面初始化与差异比对使用。

3. 设置面板构建（UI Construction）：
   使用 QStackedWidget 构建“图名”“轴标题”“轴刻度”三类设置页，
   并通过顶部下拉框在不同设置类别之间切换。

4. 草稿状态缓存（Draft State Management）：
   由于 X 轴与 Y 轴共用同一套表单控件，用户在切换轴对象时，
   需要先保存当前表单输入，再恢复目标轴的草稿内容，避免输入丢失。

5. 差异收集（Diff Collection）：
   用户点击“确定”时，将当前输入与默认上下文逐项比对，
   仅保留相对于默认值发生变化的配置项，避免生成冗余配置。

适用场景：
- 图名文本、字体、字号调整
- X/Y 轴标题文本、字体、字号调整
- X/Y 轴刻度字体、字号、数字格式、显示单位调整
- 针对多数据序列场景，为不同轴绑定当前操作的数据对象
"""

from __future__ import annotations

import copy
from typing import Any, Dict, List, Optional, Tuple

from PySide6.QtCore import Qt
from PySide6.QtGui import QFontDatabase, QIntValidator
from PySide6.QtWidgets import (
    QComboBox,
    QDialog,
    QDialogButtonBox,
    QFormLayout,
    QLabel,
    QLineEdit,
    QStackedWidget,
    QVBoxLayout,
    QWidget,
)

from ui_shared import DRISK_DIALOG_BTN_QSS, SmartFormatter, drisk_combobox_qss


# =======================================================
# 1. 全局常量与预设定义
# =======================================================
# 默认字体：当外部未提供字体信息，或提供的字体为空时，使用该默认字体。
DEFAULT_FONT_FAMILY = "Arial"

# 图名默认字号。
DEFAULT_CHART_TITLE_FONT_SIZE = 14

# 坐标轴标题默认字号。
DEFAULT_AXIS_TITLE_FONT_SIZE = 12

# 坐标轴刻度默认字号。
DEFAULT_AXIS_TICK_FONT_SIZE = 12

# 数字格式选项映射表：界面显示文字 -> 内部配置值。
# 说明：
# - 左侧中文文本用于在下拉框中展示；
# - 右侧英文标识用于内部持久化与后续渲染逻辑识别。
NUMBER_FORMAT_OPTIONS: List[Tuple[str, str]] = [
    ("自动", "auto"),
    ("数字", "number"),
    ("货币", "currency"),
    ("百分比", "percent"),
    ("科学", "scientific"),
    ("整数", "integer"),
]

# 支持的坐标轴键集合。
# 当前模块仅支持 x / y 两个坐标轴。
_AXIS_KEYS = ("x", "y")

# 当 Qt 字体数据库不可用时的回退字体列表。
# 仅作为兜底，不代表最终界面一定只使用这些字体。
_FONT_FALLBACKS = [
    "Arial",
    "Microsoft YaHei",
]


# =======================================================
# 2. 基础数据转换与格式化辅助函数
# =======================================================
def _to_int(value: Any, default: int) -> int:
    """
    安全整型转换。

    设计目的：
    - 外部配置和上下文中可能传入字符串、浮点数、None 或非法值；
    - 本函数统一负责做“尽量转换，失败回退”的防呆处理。

    参数：
    - value：待转换对象
    - default：转换失败时返回的默认值

    返回：
    - 转换后的整数；若失败则返回 default
    """
    try:
        return int(value)
    except Exception:
        return int(default)


def _to_str(value: Any) -> str:
    """
    安全字符串转换，并清理首尾空白。

    设计目的：
    - 避免 None、空对象等直接参与字符串比较或写入 UI；
    - 保证内部处理文本时始终使用统一格式。

    参数：
    - value：任意输入值

    返回：
    - 去除首尾空白后的字符串；
    - 若 value 为 None 或等价空值，则返回空字符串。
    """
    return str(value or "").strip()


def _format_mag_label(mag: int) -> str:
    """
    将数量级整数格式化为可读的显示单位文本。

    例如：
    - 0  -> "单位 (1)"
    - 3  -> "k (10^3)"
    - 6  -> "M (10^6)"
    - 若 SmartFormatter 中未定义对应 SI 前缀，则退回显示为 "10^n"

    参数：
    - mag：数量级指数，例如 0、3、6、-3 等

    返回：
    - 用于下拉框展示的显示单位标签文本
    """
    suffix = str(SmartFormatter.SI_MAP.get(int(mag), "") or "")
    if int(mag) == 0:
        return "单位 (1)"
    return f"{suffix} (10^{int(mag)})" if suffix else f"10^{int(mag)}"


def build_nearby_magnitude_options(center_mag: int, levels: int = 3) -> List[Tuple[Optional[int], str]]:
    """
    围绕推荐数量级构建一组显示单位候选项。

    设计逻辑：
    - 以 center_mag 为中心；
    - 每一级按 3 为步长上下展开；
    - 例如 center_mag=0, levels=3 时，可生成：
      -9, -6, -3, 0, 3, 6, 9
    - 中心项会额外加上“推荐:”前缀，便于用户识别系统建议值。

    参数：
    - center_mag：推荐数量级
    - levels：向上和向下扩展的层数，默认 3

    返回：
    - [(mag, label), ...] 形式的列表，
      可直接用于“显示单位”下拉框填充。
    """
    center = _to_int(center_mag, 0)
    items: List[Tuple[Optional[int], str]] = []
    for step in range(-int(levels), int(levels) + 1):
        mag = center + step * 3
        label = _format_mag_label(mag)
        if step == 0:
            label = f"推荐: {label}"
        items.append((int(mag), label))
    return items


# =======================================================
# 3. 核心业务字典对象防呆与规范化
#    Config：用户主动覆盖的配置
#    Context：图表运行时提供的默认上下文
# =======================================================
def create_default_label_settings_config() -> Dict[str, Any]:
    """
    创建一份空的标签设置配置模板。

    该模板代表“没有任何用户自定义覆盖”的初始状态。
    其中：
    - None 表示该项没有主动覆盖值，应回退使用默认上下文；
    - "auto" 表示数字格式默认自动判断；
    - selected_data_key / is_numeric 用于刻度配置中记录当前绑定的数据对象及其类型。

    返回：
    - 标准结构的配置字典
    """
    return {
        "chart_title": {
            "text_override": None,
            "font_family": None,
            "font_size": None,
        },
        "axis_title": {
            "x": {"text_override": None, "font_family": None, "font_size": None},
            "y": {"text_override": None, "font_family": None, "font_size": None},
        },
        "axis_tick": {
            "x": {
                "font_family": None,
                "font_size": None,
                "number_format": "auto",
                "display_unit_mag": None,
                "selected_data_key": "",
                "is_numeric": True,
            },
            "y": {
                "font_family": None,
                "font_size": None,
                "number_format": "auto",
                "display_unit_mag": None,
                "selected_data_key": "",
                "is_numeric": True,
            },
        },
    }


def normalize_label_settings_config(raw: Optional[Dict[str, Any]]) -> Dict[str, Any]:
    """
    将外部传入的原始配置字典清洗为严格的内部配置结构。

    设计目标：
    1. 防止调用方传入缺字段、字段类型不正确或包含非法键值的字典；
    2. 保证下游 UI 加载与最终收集逻辑始终面对同一份标准 Schema；
    3. 对不存在的字段填入默认模板中的安全值。

    处理规则：
    - 仅识别 chart_title / axis_title / axis_tick 三类配置；
    - 对 x / y 轴分别独立读取；
    - number_format / selected_data_key 会强制转换为字符串；
    - is_numeric 会强制转换为布尔值。

    参数：
    - raw：外部传入的原始配置，允许为 None

    返回：
    - 经过归一化后的标准配置字典
    """
    cfg = create_default_label_settings_config()
    src = raw if isinstance(raw, dict) else {}

    # -------------------------------
    # 3.1 图名配置提取
    # -------------------------------
    chart_src = src.get("chart_title", {})
    if isinstance(chart_src, dict):
        cfg["chart_title"]["text_override"] = chart_src.get("text_override", None)
        cfg["chart_title"]["font_family"] = chart_src.get("font_family", None)
        cfg["chart_title"]["font_size"] = chart_src.get("font_size", None)

    # -------------------------------
    # 3.2 轴标题配置提取
    # -------------------------------
    axis_title_src = src.get("axis_title", {})
    if isinstance(axis_title_src, dict):
        for axis in _AXIS_KEYS:
            section = axis_title_src.get(axis, {})
            if not isinstance(section, dict):
                continue
            cfg["axis_title"][axis]["text_override"] = section.get("text_override", None)
            cfg["axis_title"][axis]["font_family"] = section.get("font_family", None)
            cfg["axis_title"][axis]["font_size"] = section.get("font_size", None)

    # -------------------------------
    # 3.3 轴刻度配置提取
    # -------------------------------
    axis_tick_src = src.get("axis_tick", {})
    if isinstance(axis_tick_src, dict):
        for axis in _AXIS_KEYS:
            section = axis_tick_src.get(axis, {})
            if not isinstance(section, dict):
                continue
            cfg["axis_tick"][axis]["font_family"] = section.get("font_family", None)
            cfg["axis_tick"][axis]["font_size"] = section.get("font_size", None)
            cfg["axis_tick"][axis]["number_format"] = _to_str(section.get("number_format", "auto")) or "auto"
            cfg["axis_tick"][axis]["display_unit_mag"] = section.get("display_unit_mag", None)
            cfg["axis_tick"][axis]["selected_data_key"] = _to_str(section.get("selected_data_key", ""))
            cfg["axis_tick"][axis]["is_numeric"] = bool(section.get("is_numeric", True))

    return cfg


def get_axis_display_unit_override(config: Optional[Dict[str, Any]], axis_key: str) -> Optional[int]:
    """
    从配置中提取某个坐标轴的显示单位覆盖值。

    设计目的：
    - 为外部调用方提供便捷访问接口；
    - 避免外部重复编写 config 归一化与异常类型转换逻辑。

    处理规则：
    - axis_key 非 "y" 时一律按 "x" 处理；
    - 若未设置覆盖值，返回 None；
    - 若值存在但无法转换为整数，也返回 None。

    参数：
    - config：原始或已归一化配置
    - axis_key：坐标轴标识，支持 "x" / "y"

    返回：
    - 覆盖的显示单位数量级整数，或 None
    """
    cfg = normalize_label_settings_config(config)
    axis = "y" if _to_str(axis_key).lower() == "y" else "x"
    raw = cfg.get("axis_tick", {}).get(axis, {}).get("display_unit_mag", None)
    if raw is None:
        return None
    try:
        return int(raw)
    except Exception:
        return None


def get_axis_numeric_flags(config: Optional[Dict[str, Any]], fallback: Optional[Dict[str, bool]] = None) -> Dict[str, bool]:
    """
    获取当前配置中各坐标轴是否为数值型数据的标志。

    设计目的：
    - 外部有时需要快速知道 x / y 轴是否允许使用数字格式与显示单位；
    - 若配置中无有效信息，则可使用 fallback 提供兜底值。

    参数：
    - config：原始或已归一化配置
    - fallback：默认标志位字典，例如 {"x": True, "y": False}

    返回：
    - {"x": bool, "y": bool}
    """
    base = {"x": True, "y": True}
    if isinstance(fallback, dict):
        base["x"] = bool(fallback.get("x", True))
        base["y"] = bool(fallback.get("y", True))

    cfg = normalize_label_settings_config(config)
    for axis in _AXIS_KEYS:
        base[axis] = bool(cfg.get("axis_tick", {}).get(axis, {}).get("is_numeric", base[axis]))
    return base


def _normalize_candidates(raw: Any, axis: str) -> List[Dict[str, Any]]:
    """
    将坐标轴可选数据序列列表清洗为统一结构。

    设计背景：
    - 某些图表场景下，一个坐标轴可能对应多条数据线或多个候选数据对象；
    - 为了在“轴刻度”页中通过下拉框切换当前操作对象，需要统一候选项格式。

    目标结构：
    [
        {
            "key": "唯一标识",
            "label": "显示名称",
            "is_numeric": True / False
        },
        ...
    ]

    处理规则：
    - 如果 raw 不是 list，则返回一个默认候选项；
    - 如果候选项本身是 dict，则尽量读取 key / label / is_numeric；
    - 如果候选项不是 dict，则将其转为文本标签，并自动生成 key；
    - 如果最终没有有效候选项，则补一个默认项。

    参数：
    - raw：原始候选项列表
    - axis：当前轴标识，仅用于生成默认 key

    返回：
    - 归一化后的候选项列表
    """
    if not isinstance(raw, list):
        return [{"key": axis, "label": "当前数据", "is_numeric": True}]

    normalized: List[Dict[str, Any]] = []
    for idx, item in enumerate(raw):
        if isinstance(item, dict):
            key = _to_str(item.get("key", "")) or f"{axis}_{idx}"
            label = _to_str(item.get("label", "")) or key
            is_numeric = bool(item.get("is_numeric", True))
        else:
            key = f"{axis}_{idx}"
            label = _to_str(item)
            is_numeric = True
        normalized.append({"key": key, "label": label, "is_numeric": is_numeric})

    if not normalized:
        normalized.append({"key": axis, "label": "当前数据", "is_numeric": True})
    return normalized


def normalize_label_settings_context(raw: Optional[Dict[str, Any]]) -> Dict[str, Any]:
    """
    将图表运行时传入的上下文信息归一化为内部统一结构。

    Context 与 Config 的区别：
    - Config：用户主动覆盖的设置；
    - Context：当前图表环境的默认值，用于界面初始显示和最终差异比对。

    Context 包含的信息主要有：
    1. 图名默认文本、默认字体、默认字号
    2. X/Y 轴默认标题、标题字体、标题字号
    3. X/Y 轴默认刻度字体、刻度字号
    4. X/Y 轴推荐显示单位量级
    5. X/Y 轴可选数据候选列表

    参数：
    - raw：外部传入的上下文字典，允许为 None

    返回：
    - 标准化后的上下文字典
    """
    src = raw if isinstance(raw, dict) else {}
    chart_src = src.get("chart_title", {}) if isinstance(src.get("chart_title", {}), dict) else {}
    axes_src = src.get("axes", {}) if isinstance(src.get("axes", {}), dict) else {}

    ctx = {
        "chart_title": {
            "default_text": _to_str(chart_src.get("default_text", "")),
            "default_font_family": _to_str(chart_src.get("default_font_family", DEFAULT_FONT_FAMILY)) or DEFAULT_FONT_FAMILY,
            "default_font_size": _to_int(chart_src.get("default_font_size", DEFAULT_CHART_TITLE_FONT_SIZE), DEFAULT_CHART_TITLE_FONT_SIZE),
        },
        "axes": {
            "x": {
                "default_title": "",
                "default_title_font_family": DEFAULT_FONT_FAMILY,
                "default_title_font_size": DEFAULT_AXIS_TITLE_FONT_SIZE,
                "default_tick_font_family": DEFAULT_FONT_FAMILY,
                "default_tick_font_size": DEFAULT_AXIS_TICK_FONT_SIZE,
                "recommended_mag": 0,
                "data_candidates": [{"key": "x", "label": "当前X数据", "is_numeric": True}],
            },
            "y": {
                "default_title": "",
                "default_title_font_family": DEFAULT_FONT_FAMILY,
                "default_title_font_size": DEFAULT_AXIS_TITLE_FONT_SIZE,
                "default_tick_font_family": DEFAULT_FONT_FAMILY,
                "default_tick_font_size": DEFAULT_AXIS_TICK_FONT_SIZE,
                "recommended_mag": 0,
                "data_candidates": [{"key": "y", "label": "当前Y数据", "is_numeric": True}],
            },
        },
    }

    # -------------------------------
    # 3.4 遍历填充 X / Y 轴上下文
    # -------------------------------
    for axis in _AXIS_KEYS:
        axis_src = axes_src.get(axis, {}) if isinstance(axes_src.get(axis, {}), dict) else {}
        ctx["axes"][axis]["default_title"] = _to_str(axis_src.get("default_title", ""))
        ctx["axes"][axis]["default_title_font_family"] = _to_str(
            axis_src.get("default_title_font_family", DEFAULT_FONT_FAMILY)
        ) or DEFAULT_FONT_FAMILY
        ctx["axes"][axis]["default_title_font_size"] = _to_int(
            axis_src.get("default_title_font_size", DEFAULT_AXIS_TITLE_FONT_SIZE),
            DEFAULT_AXIS_TITLE_FONT_SIZE,
        )
        ctx["axes"][axis]["default_tick_font_family"] = _to_str(
            axis_src.get("default_tick_font_family", DEFAULT_FONT_FAMILY)
        ) or DEFAULT_FONT_FAMILY
        ctx["axes"][axis]["default_tick_font_size"] = _to_int(
            axis_src.get("default_tick_font_size", DEFAULT_AXIS_TICK_FONT_SIZE),
            DEFAULT_AXIS_TICK_FONT_SIZE,
        )
        ctx["axes"][axis]["recommended_mag"] = _to_int(axis_src.get("recommended_mag", 0), 0)
        ctx["axes"][axis]["data_candidates"] = _normalize_candidates(axis_src.get("data_candidates", []), axis)

    return ctx


# =======================================================
# 4. 文本与标签设置主界面类
# =======================================================
class LabelSettingsDialog(QDialog):
    """
    图表文本设置主对话框。

    本类负责完成以下工作：
    1. 接收并归一化外部配置与上下文；
    2. 构建图名、轴标题、轴刻度三类设置界面；
    3. 管理 X / Y 轴共用表单下的草稿切换；
    4. 收集用户输入并生成最终配置。

    界面层面包含三类设置页：
    - 图名：图名文字、字体、字号
    - 轴标题：X/Y 轴标题文字、字体、字号
    - 轴刻度：X/Y 轴数据对象、字体、字号、数字格式、显示单位

    关键内部状态：
    - self._config：当前生效配置（标准结构）
    - self._context：图表默认上下文（标准结构）
    - self._axis_title_draft：轴标题页草稿缓存
    - self._axis_tick_draft：轴刻度页草稿缓存
    - self._active_axis_title_key：当前轴标题页正在编辑的轴
    - self._active_axis_tick_key：当前轴刻度页正在编辑的轴
    """

    # ---------------------------------------------------
    # 4.1 初始化
    # ---------------------------------------------------
    def __init__(
        self,
        config: Optional[Dict[str, Any]],
        context: Optional[Dict[str, Any]],
        parent=None,
        include_chart_title: bool = True,
    ):
        """
        初始化对话框。

        参数：
        - config：用户已有的自定义配置
        - context：当前图表环境提供的默认上下文
        - parent：父级窗口
        - include_chart_title：
            是否显示“图名”设置页。
            某些场景下如果不允许修改图名，可将其设为 False。
        """
        super().__init__(parent)
        self.setWindowTitle("Drisk - 文本设置")
        self.resize(300, 200)
        self._include_chart_title = bool(include_chart_title)

        # 归一化外部输入，确保后续逻辑面对的是标准数据结构。
        self._config = normalize_label_settings_config(config)
        self._context = normalize_label_settings_context(context)

        # -------------------------------
        # 草稿缓存机制
        # -------------------------------
        # 原因：
        # - 轴标题页和轴刻度页中，X轴 / Y轴共用同一组输入控件；
        # - 用户切换当前轴对象时，若不先保存当前值，表单输入会丢失。
        #
        # 处理方式：
        # - 初始化时复制现有配置为 draft；
        # - 每次切换轴时，先保存当前表单，再加载目标轴的草稿内容。
        self._axis_title_draft = copy.deepcopy(self._config.get("axis_title", {}))
        self._axis_tick_draft = copy.deepcopy(self._config.get("axis_tick", {}))
        self._active_axis_title_key = "x"
        self._active_axis_tick_key = "x"

        self._init_ui()
        self._load_all_controls()

    def get_config(self) -> Dict[str, Any]:
        """
        获取当前对话框内部保存的最终配置。

        返回：
        - self._config 的深拷贝，避免外部直接修改内部状态。
        """
        return copy.deepcopy(self._config)

    # ---------------------------------------------------
    # 4.2 UI 界面构建
    # ---------------------------------------------------
    def _init_ui(self) -> None:
        """
        构建整个对话框的界面结构。

        界面布局包括：
        1. 顶部“配置类别”下拉框
        2. 中部堆叠页（图名 / 轴标题 / 轴刻度）
        3. 底部确定 / 取消按钮

        备注：
        - 三类设置页使用 QStackedWidget 管理；
        - 分类下拉框切换时，仅切换堆叠页，不销毁控件；
        - “确定”按钮不会直接 accept，而是先执行最终配置收集。
        """
        self.setStyleSheet(
            """
            QDialog {
                background-color: #f9f9f9;
                font-family: 'Microsoft YaHei';
                font-weight: normal;
            }
            QLabel {
                font-size: 12px;
                color: #333333;
                font-weight: normal;
            }
            QLineEdit, QComboBox {
                background-color: #ffffff;
                border: 1px solid #d9d9d9;
                border-radius: 3px;
                min-height: 24px;
                padding: 0px 6px;
                font-size: 12px;
                font-weight: normal;
            }
            QLineEdit:focus, QComboBox:focus { border-color: #40a9ff; }
            QPushButton { font-weight: normal; }
            """
        )

        layout = QVBoxLayout(self)
        layout.setContentsMargins(14, 14, 14, 14)
        layout.setSpacing(10)

        # 表单统一列宽：
        # 左侧标签列较窄，右侧输入控件列固定宽度，保证多页布局一致。
        self._label_col_width = 60
        self._field_col_width = 170

        # -------------------------------
        # 顶部分类选择区域
        # -------------------------------
        top_form = QFormLayout()
        self._configure_form_layout(top_form)
        top_form.setContentsMargins(8, 0, 8, 0)

        self.combo_category = QComboBox()
        self.combo_category.setStyleSheet(drisk_combobox_qss())
        self.combo_category.setFixedWidth(self._field_col_width)

        if self._include_chart_title:
            self.combo_category.addItem("图名", "chart_title")
        self.combo_category.addItem("轴标题", "axis_title")
        self.combo_category.addItem("轴刻度", "axis_tick")
        self.combo_category.currentIndexChanged.connect(self._on_category_changed)

        self._add_form_row(top_form, "配置类别", self.combo_category)
        layout.addLayout(top_form)

        # -------------------------------
        # 中部堆叠页
        # -------------------------------
        self.stack = QStackedWidget()
        self._category_index_map: Dict[str, int] = {}

        self.page_chart = None
        if self._include_chart_title:
            self.page_chart = self._build_chart_page()
            self.stack.addWidget(self.page_chart)
            self._category_index_map["chart_title"] = self.stack.indexOf(self.page_chart)

        self.page_axis_title = self._build_axis_title_page()
        self.stack.addWidget(self.page_axis_title)
        self._category_index_map["axis_title"] = self.stack.indexOf(self.page_axis_title)

        self.page_axis_tick = self._build_axis_tick_page()
        self.stack.addWidget(self.page_axis_tick)
        self._category_index_map["axis_tick"] = self.stack.indexOf(self.page_axis_tick)

        layout.addWidget(self.stack, 1)

        # -------------------------------
        # 底部按钮
        # -------------------------------
        btns = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        btn_ok = btns.button(QDialogButtonBox.Ok)
        btn_cancel = btns.button(QDialogButtonBox.Cancel)
        if btn_ok is not None:
            btn_ok.setText("确定")
            btn_ok.setObjectName("btnOk")
        if btn_cancel is not None:
            btn_cancel.setText("取消")
            btn_cancel.setObjectName("btnCancel")

        btns.setStyleSheet(
            DRISK_DIALOG_BTN_QSS
            + "\nQPushButton#btnOk, QPushButton#btnCancel { font-weight: normal; }"
        )

        btns.accepted.connect(self._accept_with_collect)
        btns.rejected.connect(self.reject)
        layout.addWidget(btns)

    def _configure_form_layout(self, form: QFormLayout) -> None:
        """
        统一配置表单布局样式。

        目的：
        - 保证多页表单在标签对齐、间距、整体排版上保持一致；
        - 减少每个页面重复设置布局属性的代码。

        规则：
        - 标签右对齐、垂直居中
        - 横向与纵向间距固定
        - 表单整体左上对齐
        """
        form.setLabelAlignment(Qt.AlignRight | Qt.AlignVCenter)
        form.setHorizontalSpacing(10)
        form.setVerticalSpacing(10)
        form.setFormAlignment(Qt.AlignLeft | Qt.AlignTop)

    def _create_form_label(self, text: str) -> QLabel:
        """
        创建标准表单标签控件。

        设计目的：
        - 所有表单左侧标签使用统一宽度、统一对齐方式；
        - 便于后续全局调整。

        参数：
        - text：标签文本

        返回：
        - QLabel 实例
        """
        label = QLabel(text)
        label.setFixedWidth(self._label_col_width)
        label.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        return label

    def _add_form_row(self, form: QFormLayout, label_text: str, field_widget: QWidget) -> None:
        """
        向表单中插入一行“标签 + 控件”。

        参数：
        - form：目标表单布局
        - label_text：左侧标签文字
        - field_widget：右侧输入控件
        """
        form.addRow(self._create_form_label(label_text), field_widget)

    def _create_font_size_input(self) -> QLineEdit:
        """
        创建字号输入框。

        设计要求：
        - 字号只允许输入 8~72 的整数；
        - 与其他输入控件使用统一宽度。

        返回：
        - 配置好校验器的 QLineEdit
        """
        edit = QLineEdit()
        edit.setValidator(QIntValidator(8, 72, self))
        edit.setFixedWidth(self._field_col_width)
        return edit

    def _build_chart_page(self) -> QWidget:
        """
        构建“图名”设置页。

        页面字段：
        - 图名文字
        - 字体
        - 字号

        返回：
        - 图名页 QWidget
        """
        page = QWidget()
        form = QFormLayout(page)
        self._configure_form_layout(form)
        form.setContentsMargins(8, 8, 8, 8)

        self.edt_chart_text = QLineEdit()
        self.edt_chart_text.setFixedWidth(self._field_col_width)

        self.cmb_chart_font = QComboBox()
        self.cmb_chart_font.setEditable(True)
        self.cmb_chart_font.setFixedWidth(self._field_col_width)

        self.edt_chart_size = self._create_font_size_input()

        self._add_form_row(form, "图名文字", self.edt_chart_text)
        self._add_form_row(form, "字体", self.cmb_chart_font)
        self._add_form_row(form, "字号", self.edt_chart_size)
        return page

    def _build_axis_title_page(self) -> QWidget:
        """
        构建“轴标题”设置页。

        页面字段：
        - 当前轴（X轴 / Y轴）
        - 标题文字
        - 字体
        - 字号

        说明：
        - X 轴与 Y 轴共用同一套输入控件；
        - 实际数据通过 self._axis_title_draft 分别缓存。
        """
        page = QWidget()
        form = QFormLayout(page)
        self._configure_form_layout(form)
        form.setContentsMargins(8, 8, 8, 8)

        self.cmb_axis_title_axis = QComboBox()
        self.cmb_axis_title_axis.addItem("X轴", "x")
        self.cmb_axis_title_axis.addItem("Y轴", "y")
        self.cmb_axis_title_axis.currentIndexChanged.connect(self._on_axis_title_axis_changed)
        self.cmb_axis_title_axis.setFixedWidth(self._field_col_width)

        self.edt_axis_title_text = QLineEdit()
        self.edt_axis_title_text.setFixedWidth(self._field_col_width)

        self.cmb_axis_title_font = QComboBox()
        self.cmb_axis_title_font.setEditable(True)
        self.cmb_axis_title_font.setFixedWidth(self._field_col_width)

        self.edt_axis_title_size = self._create_font_size_input()

        self._add_form_row(form, "轴", self.cmb_axis_title_axis)
        self._add_form_row(form, "标题文字", self.edt_axis_title_text)
        self._add_form_row(form, "字体", self.cmb_axis_title_font)
        self._add_form_row(form, "字号", self.edt_axis_title_size)
        return page

    def _build_axis_tick_page(self) -> QWidget:
        """
        构建“轴刻度”设置页。

        页面字段：
        - 当前轴（X轴 / Y轴）
        - 当前绑定的 X 数据对象
        - 当前绑定的 Y 数据对象
        - 刻度字体
        - 刻度字号
        - 数字格式
        - 显示单位
        - 非数值数据提示信息

        说明：
        1. X/Y 数据对象下拉框用于多数据序列场景下选择“当前操作目标”；
        2. 数字格式和显示单位仅对数值型数据有效；
        3. 若当前绑定对象是文本数据，则相关控件会被禁用，并显示提示语。
        """
        page = QWidget()
        form = QFormLayout(page)
        self._configure_form_layout(form)
        form.setContentsMargins(8, 8, 8, 8)

        self.cmb_axis_tick_axis = QComboBox()
        self.cmb_axis_tick_axis.addItem("X轴", "x")
        self.cmb_axis_tick_axis.addItem("Y轴", "y")
        self.cmb_axis_tick_axis.currentIndexChanged.connect(self._on_axis_tick_axis_changed)
        self.cmb_axis_tick_axis.setFixedWidth(self._field_col_width)

        # 多数据序列场景下的“当前数据对象”选择器。
        self.cmb_tick_data_x = QComboBox()
        self.cmb_tick_data_y = QComboBox()
        self.cmb_tick_data_x.currentIndexChanged.connect(self._on_tick_data_selection_changed)
        self.cmb_tick_data_y.currentIndexChanged.connect(self._on_tick_data_selection_changed)
        self.cmb_tick_data_x.setFixedWidth(self._field_col_width)
        self.cmb_tick_data_y.setFixedWidth(self._field_col_width)

        self.cmb_axis_tick_font = QComboBox()
        self.cmb_axis_tick_font.setEditable(True)
        self.cmb_axis_tick_font.setFixedWidth(self._field_col_width)

        self.edt_axis_tick_size = self._create_font_size_input()

        self.cmb_number_format = QComboBox()
        for label, value in NUMBER_FORMAT_OPTIONS:
            self.cmb_number_format.addItem(label, value)
        self.cmb_number_format.setFixedWidth(self._field_col_width)

        self.cmb_display_unit = QComboBox()
        self.cmb_display_unit.setFixedWidth(self._field_col_width)

        # 当前轴数据不是数值型时，用于给出说明。
        self.lbl_numeric_hint = QLabel("")
        self.lbl_numeric_hint.setStyleSheet("color: #888888;")
        self.lbl_numeric_hint.setWordWrap(True)

        self._add_form_row(form, "轴", self.cmb_axis_tick_axis)
        self._add_form_row(form, "X数据", self.cmb_tick_data_x)
        self._add_form_row(form, "Y数据", self.cmb_tick_data_y)
        self._add_form_row(form, "字体", self.cmb_axis_tick_font)
        self._add_form_row(form, "字号", self.edt_axis_tick_size)
        self._add_form_row(form, "数字格式", self.cmb_number_format)
        self._add_form_row(form, "显示单位", self.cmb_display_unit)
        self._add_form_row(form, "", self.lbl_numeric_hint)
        return page

    # ---------------------------------------------------
    # 4.3 控件数据填充与草稿同步
    # ---------------------------------------------------
    def _list_supported_font_families(self) -> List[str]:
        """
        获取当前 Qt 运行环境可用的字体列表。

        处理逻辑：
        1. 尝试从 QFontDatabase 读取系统字体；
        2. 过滤空字符串；
        3. 过滤 Windows 下的垂直字体别名（如 '@微软雅黑'）；
        4. 按大小写无关规则去重；
        5. 最终按字母序排序；
        6. 若 Qt 字体数据库不可用，则退回 _FONT_FALLBACKS。

        返回：
        - 可用于字体下拉框填充的字体名称列表
        """
        try:
            families = list(QFontDatabase().families())
        except Exception:
            families = []

        result: List[str] = []
        seen: set[str] = set()
        for fam in families:
            text = _to_str(fam)
            if not text:
                continue
            if text.startswith("@"):
                continue
            key = text.casefold()
            if key in seen:
                continue
            seen.add(key)
            result.append(text)

        result.sort(key=lambda x: x.casefold())
        if not result:
            result = [fam for fam in _FONT_FALLBACKS if _to_str(fam)]
        return result

    def _fill_font_combo(self, combo: QComboBox, current_family: str) -> None:
        """
        用可用字体列表填充字体下拉框，并保证当前字体值一定可见。

        设计原因：
        - 某些旧配置中的字体名，可能当前系统字体库中不存在；
        - 为避免界面加载后丢失该值，需要将当前值补充进下拉框。

        参数：
        - combo：目标字体下拉框
        - current_family：当前应显示的字体名称
        """
        combo.blockSignals(True)
        combo.clear()

        current = _to_str(current_family) or DEFAULT_FONT_FAMILY
        for fam in self._list_supported_font_families():
            combo.addItem(fam, fam)

        idx = combo.findData(current)
        if idx < 0:
            combo.addItem(current, current)
            idx = combo.findData(current)
        if idx < 0:
            combo.addItem(DEFAULT_FONT_FAMILY, DEFAULT_FONT_FAMILY)
            idx = combo.findData(DEFAULT_FONT_FAMILY)
        if idx >= 0:
            combo.setCurrentIndex(idx)

        combo.blockSignals(False)

    def _load_all_controls(self) -> None:
        """
        初始化加载所有控件的显示值。

        执行顺序：
        1. 若启用图名页，则加载图名文本 / 字体 / 字号；
        2. 预先填充轴标题页与轴刻度页的字体下拉框；
        3. 加载轴刻度页的数据对象下拉框；
        4. 默认加载 X 轴标题设置；
        5. 默认加载 X 轴刻度设置。

        注意：
        - 这里是界面首次初始化入口；
        - 后续 X/Y 轴切换时，不会再次进入本方法，而是走各自的 save/load 逻辑。
        """
        if self._include_chart_title:
            chart_ctx = self._context.get("chart_title", {})
            chart_cfg = self._config.get("chart_title", {})

            chart_text = chart_cfg.get("text_override", None)
            if chart_text is None:
                chart_text = chart_ctx.get("default_text", "")
            self.edt_chart_text.setText(_to_str(chart_text))

            chart_family = chart_cfg.get("font_family", None) or chart_ctx.get("default_font_family", DEFAULT_FONT_FAMILY)
            chart_size = chart_cfg.get("font_size", None)
            if chart_size is None:
                chart_size = chart_ctx.get("default_font_size", DEFAULT_CHART_TITLE_FONT_SIZE)

            self._fill_font_combo(self.cmb_chart_font, _to_str(chart_family))
            self.edt_chart_size.setText(str(_to_int(chart_size, DEFAULT_CHART_TITLE_FONT_SIZE)))

        self._fill_font_combo(self.cmb_axis_title_font, DEFAULT_FONT_FAMILY)
        self._fill_font_combo(self.cmb_axis_tick_font, DEFAULT_FONT_FAMILY)
        self._load_data_combos()
        self._load_axis_title_controls("x")
        self._load_axis_tick_controls("x")

    def _load_data_combos(self) -> None:
        """
        加载轴刻度页中的数据对象下拉框。

        设计背景：
        - 一个轴在某些场景下可能有多个候选数据序列；
        - 用户需要明确指定当前“刻度设置”作用在哪个数据对象上。

        处理逻辑：
        1. 从 context 中读取候选项；
        2. 填入 X 数据 / Y 数据下拉框；
        3. 如果 draft 中已有 selected_data_key，则优先恢复该值；
        4. 若未命中任何值且下拉框非空，则默认选中第一个候选项。
        """
        def _fill(combo: QComboBox, axis: str) -> None:
            combo.blockSignals(True)
            combo.clear()

            candidates = list(self._context.get("axes", {}).get(axis, {}).get("data_candidates", []))
            for item in candidates:
                key = _to_str(item.get("key", "")) or axis
                label = _to_str(item.get("label", "")) or key
                combo.addItem(label, key)

            draft_key = _to_str(self._axis_tick_draft.get(axis, {}).get("selected_data_key", ""))
            if draft_key:
                idx = combo.findData(draft_key)
                if idx >= 0:
                    combo.setCurrentIndex(idx)

            if combo.currentIndex() < 0 and combo.count() > 0:
                combo.setCurrentIndex(0)

            combo.blockSignals(False)

        _fill(self.cmb_tick_data_x, "x")
        _fill(self.cmb_tick_data_y, "y")

    def _save_axis_title_controls(self, axis: str) -> None:
        """
        将当前轴标题表单中的输入保存到对应轴的草稿缓存。

        设计目的：
        - X / Y 轴切换前，先将当前表单内容写回 draft；
        - 避免用户未点击“确定”时，中间输入内容丢失。

        保存字段：
        - text_override
        - font_family
        - font_size
        """
        axis_key = "y" if axis == "y" else "x"
        section = self._axis_title_draft.setdefault(axis_key, {})
        section["text_override"] = _to_str(self.edt_axis_title_text.text())
        section["font_family"] = _to_str(self.cmb_axis_title_font.currentText())
        section["font_size"] = self._parse_font_size_input(self.edt_axis_title_size.text(), DEFAULT_AXIS_TITLE_FONT_SIZE)

    def _load_axis_title_controls(self, axis: str) -> None:
        """
        将指定轴的标题草稿加载到表单控件中。

        加载规则：
        - 若 draft 中有值，则优先显示 draft；
        - 若 draft 对应字段为 None，则回退使用 context 默认值。

        参数：
        - axis：目标轴，支持 "x" / "y"
        """
        axis_key = "y" if axis == "y" else "x"
        section = self._axis_title_draft.get(axis_key, {})
        axis_ctx = self._context.get("axes", {}).get(axis_key, {})

        text_val = section.get("text_override", None)
        if text_val is None:
            text_val = axis_ctx.get("default_title", "")
        self.edt_axis_title_text.setText(_to_str(text_val))

        fam = section.get("font_family", None) or axis_ctx.get("default_title_font_family", DEFAULT_FONT_FAMILY)
        size = section.get("font_size", None)
        if size is None:
            size = axis_ctx.get("default_title_font_size", DEFAULT_AXIS_TITLE_FONT_SIZE)

        self._fill_font_combo(self.cmb_axis_title_font, _to_str(fam))
        self.edt_axis_title_size.setText(str(_to_int(size, DEFAULT_AXIS_TITLE_FONT_SIZE)))

    def _resolve_axis_selected_data_key(self, axis: str) -> str:
        """
        获取当前轴刻度页中所选的数据对象 key。

        参数：
        - axis：目标轴，支持 "x" / "y"

        返回：
        - 当前下拉框选中的数据对象标识
        """
        if axis == "y":
            return _to_str(self.cmb_tick_data_y.currentData())
        return _to_str(self.cmb_tick_data_x.currentData())

    def _resolve_axis_data_is_numeric(self, axis: str) -> bool:
        """
        判断当前轴所绑定的数据对象是否为数值型。

        用途：
        - 控制“数字格式”“显示单位”是否允许编辑；
        - 写入 axis_tick 草稿中的 is_numeric 状态。

        处理逻辑：
        1. 获取当前轴选中的 selected_data_key；
        2. 在 context 的 data_candidates 中查找对应对象；
        3. 若找到则返回其 is_numeric；
        4. 若找不到，则回退使用候选列表第一个对象的类型；
        5. 若候选列表为空，则默认按数值型处理。
        """
        axis_key = "y" if axis == "y" else "x"
        selected_key = self._resolve_axis_selected_data_key(axis_key)
        candidates = list(self._context.get("axes", {}).get(axis_key, {}).get("data_candidates", []))
        default_flag = True
        if candidates:
            default_flag = bool(candidates[0].get("is_numeric", True))
        for item in candidates:
            if _to_str(item.get("key", "")) == selected_key:
                return bool(item.get("is_numeric", True))
        return default_flag

    def _rebuild_display_unit_combo(self, axis: str) -> None:
        """
        根据当前轴的推荐数量级，重建“显示单位”下拉框。

        设计逻辑：
        - 以 context 中的 recommended_mag 作为中心；
        - 自动生成附近若干个数量级候选；
        - 若草稿中已有 display_unit_mag，则优先恢复该值；
        - 若草稿值无效，则回退选中推荐值。

        参数：
        - axis：目标轴，支持 "x" / "y"
        """
        axis_key = "y" if axis == "y" else "x"
        axis_ctx = self._context.get("axes", {}).get(axis_key, {})
        recommended_mag = _to_int(axis_ctx.get("recommended_mag", 0), 0)
        preferred_mag = self._axis_tick_draft.get(axis_key, {}).get("display_unit_mag", None)

        self.cmb_display_unit.blockSignals(True)
        self.cmb_display_unit.clear()
        for mag, label in build_nearby_magnitude_options(recommended_mag, levels=3):
            self.cmb_display_unit.addItem(label, mag)

        target_mag = recommended_mag if preferred_mag is None else _to_int(preferred_mag, recommended_mag)
        idx = self.cmb_display_unit.findData(target_mag)
        if idx < 0:
            idx = self.cmb_display_unit.findData(recommended_mag)
        if idx >= 0:
            self.cmb_display_unit.setCurrentIndex(idx)

        self.cmb_display_unit.blockSignals(False)

    def _save_axis_tick_controls(self, axis: str) -> None:
        """
        将当前轴刻度表单中的输入保存到对应轴的草稿缓存。

        保存字段：
        - selected_data_key：当前绑定的数据对象
        - font_family：刻度字体
        - font_size：刻度字号
        - number_format：数字格式
        - display_unit_mag：显示单位数量级
        - is_numeric：当前绑定对象是否为数值型

        注意：
        - is_numeric 不是直接由用户输入，而是根据当前数据对象自动判定。
        """
        axis_key = "y" if axis == "y" else "x"
        section = self._axis_tick_draft.setdefault(axis_key, {})
        section["selected_data_key"] = self._resolve_axis_selected_data_key(axis_key)
        section["font_family"] = _to_str(self.cmb_axis_tick_font.currentText())
        section["font_size"] = self._parse_font_size_input(self.edt_axis_tick_size.text(), DEFAULT_AXIS_TICK_FONT_SIZE)
        section["number_format"] = _to_str(self.cmb_number_format.currentData()) or "auto"
        section["display_unit_mag"] = self.cmb_display_unit.currentData()
        section["is_numeric"] = self._resolve_axis_data_is_numeric(axis_key)

    def _load_axis_tick_controls(self, axis: str) -> None:
        """
        将指定轴的刻度草稿加载到表单控件中。

        加载顺序：
        1. 字体、字号
        2. 数字格式
        3. 显示单位
        4. 根据是否为数值型，刷新相关控件启用状态和提示语

        参数：
        - axis：目标轴，支持 "x" / "y"
        """
        axis_key = "y" if axis == "y" else "x"
        section = self._axis_tick_draft.get(axis_key, {})
        axis_ctx = self._context.get("axes", {}).get(axis_key, {})

        fam = section.get("font_family", None) or axis_ctx.get("default_tick_font_family", DEFAULT_FONT_FAMILY)
        size = section.get("font_size", None)
        if size is None:
            size = axis_ctx.get("default_tick_font_size", DEFAULT_AXIS_TICK_FONT_SIZE)

        self._fill_font_combo(self.cmb_axis_tick_font, _to_str(fam))
        self.edt_axis_tick_size.setText(str(_to_int(size, DEFAULT_AXIS_TICK_FONT_SIZE)))

        number_format = _to_str(section.get("number_format", "auto")) or "auto"
        idx_fmt = self.cmb_number_format.findData(number_format)
        if idx_fmt < 0:
            idx_fmt = self.cmb_number_format.findData("auto")
        if idx_fmt >= 0:
            self.cmb_number_format.setCurrentIndex(idx_fmt)

        self._rebuild_display_unit_combo(axis_key)
        self._sync_numeric_control_state(axis_key)

    def _sync_numeric_control_state(self, axis: str) -> None:
        """
        根据当前轴绑定对象是否为数值型，刷新相关控件的可编辑状态。

        规则：
        - 若为数值型：
          启用“数字格式”“显示单位”，提示可配置；
        - 若为文本型：
          禁用“数字格式”“显示单位”，提示当前不可用。

        参数：
        - axis：目标轴，支持 "x" / "y"
        """
        axis_key = "y" if axis == "y" else "x"
        is_numeric = self._resolve_axis_data_is_numeric(axis_key)
        self.cmb_number_format.setEnabled(is_numeric)
        self.cmb_display_unit.setEnabled(is_numeric)

        if is_numeric:
            self.lbl_numeric_hint.setText("当前轴为数值数据，可配置数字格式和显示单位。")
        else:
            self.lbl_numeric_hint.setText("当前轴为文本数据，数字格式和显示单位不可用。")

    # ---------------------------------------------------
    # 4.4 槽函数与交互响应
    # ---------------------------------------------------
    def _on_category_changed(self) -> None:
        """
        顶部“配置类别”下拉框切换事件。

        功能：
        - 根据当前选中的配置类别 key，
          切换 QStackedWidget 中对应的页面索引。

        说明：
        - 这里只负责切换页面显示，不负责保存数据；
        - 轴标题和轴刻度页内部的数据保存由各自轴切换逻辑处理。
        """
        key = _to_str(self.combo_category.currentData())
        index = self._category_index_map.get(key, 0)
        self.stack.setCurrentIndex(index)

    def _on_axis_title_axis_changed(self) -> None:
        """
        轴标题页中 X轴 / Y轴 切换事件。

        工作流：
        1. 将当前轴的输入内容保存到草稿缓存；
        2. 更新当前活动轴标识；
        3. 加载新轴对应的草稿内容到表单控件。

        这是本模块中“共用表单 + 双轴缓存”机制的关键入口之一。
        """
        self._save_axis_title_controls(self._active_axis_title_key)
        axis = "y" if self.cmb_axis_title_axis.currentData() == "y" else "x"
        self._active_axis_title_key = axis
        self._load_axis_title_controls(axis)

    def _on_axis_tick_axis_changed(self) -> None:
        """
        轴刻度页中 X轴 / Y轴 切换事件。

        工作流：
        1. 保存当前轴刻度表单到草稿；
        2. 更新当前活动轴标识；
        3. 加载新轴草稿；
        4. 加载后会同步刷新数字格式与显示单位控件状态。
        """
        self._save_axis_tick_controls(self._active_axis_tick_key)
        axis = "y" if self.cmb_axis_tick_axis.currentData() == "y" else "x"
        self._active_axis_tick_key = axis
        self._load_axis_tick_controls(axis)

    def _on_tick_data_selection_changed(self) -> None:
        """
        轴刻度页中“当前数据对象”切换事件。

        作用：
        - 当用户切换当前绑定的数据对象后，
          其 is_numeric 可能发生变化；
        - 因此需要立即刷新“数字格式”“显示单位”的启用状态与提示语。

        注意：
        - 此处不直接保存草稿；
        - 保存动作仍由轴切换或最终提交时统一完成。
        """
        axis = "y" if self.cmb_axis_tick_axis.currentData() == "y" else "x"
        self._sync_numeric_control_state(axis)

    # ---------------------------------------------------
    # 4.5 数据收集与提交保存
    # ---------------------------------------------------
    def _collect_final_config(self) -> Dict[str, Any]:
        """
        收集当前界面输入，生成最终配置。

        这是本模块最核心的提交逻辑，其目标不是简单“把所有值原样保存”，
        而是生成一份“仅保留用户相对默认值发生变化的配置”。

        处理流程：
        1. 先将当前界面中尚未写回的轴标题 / 轴刻度表单内容保存到 draft；
        2. 新建一份空模板 cfg；
        3. 图名、轴标题、轴刻度分别与 context 默认值比较；
        4. 若用户输入值与默认值相同，则保存为 None；
        5. 若不同，则保存实际覆盖值；
        6. 最后再次走 normalize_label_settings_config，确保结构标准化。

        这样设计的好处：
        - 配置文件更精简；
        - 下游渲染逻辑只需处理“真正有覆盖”的项；
        - 更容易判断某项是“用户改了”还是“沿用默认”。
        """
        # 先保存当前界面上正在编辑但尚未写入 draft 的数据，避免遗漏。
        current_axis_title = "y" if self.cmb_axis_title_axis.currentData() == "y" else "x"
        current_axis_tick = "y" if self.cmb_axis_tick_axis.currentData() == "y" else "x"
        self._save_axis_title_controls(current_axis_title)
        self._save_axis_tick_controls(current_axis_tick)

        cfg = create_default_label_settings_config()

        # -------------------------------
        # 4.5.1 图名配置差异收集
        # -------------------------------
        if self._include_chart_title:
            chart_ctx = self._context.get("chart_title", {})
            chart_default_text = _to_str(chart_ctx.get("default_text", ""))
            chart_input_text = _to_str(self.edt_chart_text.text())
            cfg["chart_title"]["text_override"] = chart_input_text if chart_input_text != chart_default_text else None

            chart_default_family = _to_str(chart_ctx.get("default_font_family", DEFAULT_FONT_FAMILY)) or DEFAULT_FONT_FAMILY
            chart_default_size = _to_int(chart_ctx.get("default_font_size", DEFAULT_CHART_TITLE_FONT_SIZE), DEFAULT_CHART_TITLE_FONT_SIZE)
            chart_input_family = _to_str(self.cmb_chart_font.currentText()) or chart_default_family
            chart_input_size = self._parse_font_size_input(self.edt_chart_size.text(), chart_default_size)

            cfg["chart_title"]["font_family"] = chart_input_family if chart_input_family != chart_default_family else None
            cfg["chart_title"]["font_size"] = chart_input_size if chart_input_size != chart_default_size else None
        else:
            # 如果当前场景不允许编辑图名，则保留原 config 中已有的图名配置。
            existing_chart_cfg = self._config.get("chart_title", {})
            if isinstance(existing_chart_cfg, dict):
                cfg["chart_title"]["text_override"] = existing_chart_cfg.get("text_override", None)
                cfg["chart_title"]["font_family"] = existing_chart_cfg.get("font_family", None)
                cfg["chart_title"]["font_size"] = existing_chart_cfg.get("font_size", None)

        # -------------------------------
        # 4.5.2 各坐标轴配置差异收集
        # -------------------------------
        for axis in _AXIS_KEYS:
            axis_ctx = self._context.get("axes", {}).get(axis, {})

            # ---------------------------
            # 轴标题差异收集
            # ---------------------------
            title_draft = self._axis_title_draft.get(axis, {})
            default_title = _to_str(axis_ctx.get("default_title", ""))
            title_text = _to_str(title_draft.get("text_override", ""))
            cfg["axis_title"][axis]["text_override"] = title_text if title_text != default_title else None

            default_title_family = _to_str(axis_ctx.get("default_title_font_family", DEFAULT_FONT_FAMILY)) or DEFAULT_FONT_FAMILY
            default_title_size = _to_int(axis_ctx.get("default_title_font_size", DEFAULT_AXIS_TITLE_FONT_SIZE), DEFAULT_AXIS_TITLE_FONT_SIZE)
            input_title_family = _to_str(title_draft.get("font_family", default_title_family)) or default_title_family
            input_title_size = _to_int(title_draft.get("font_size", default_title_size), default_title_size)

            cfg["axis_title"][axis]["font_family"] = input_title_family if input_title_family != default_title_family else None
            cfg["axis_title"][axis]["font_size"] = input_title_size if input_title_size != default_title_size else None

            # ---------------------------
            # 轴刻度差异收集
            # ---------------------------
            tick_draft = self._axis_tick_draft.get(axis, {})
            default_tick_family = _to_str(axis_ctx.get("default_tick_font_family", DEFAULT_FONT_FAMILY)) or DEFAULT_FONT_FAMILY
            default_tick_size = _to_int(axis_ctx.get("default_tick_font_size", DEFAULT_AXIS_TICK_FONT_SIZE), DEFAULT_AXIS_TICK_FONT_SIZE)
            input_tick_family = _to_str(tick_draft.get("font_family", default_tick_family)) or default_tick_family
            input_tick_size = _to_int(tick_draft.get("font_size", default_tick_size), default_tick_size)

            cfg["axis_tick"][axis]["font_family"] = input_tick_family if input_tick_family != default_tick_family else None
            cfg["axis_tick"][axis]["font_size"] = input_tick_size if input_tick_size != default_tick_size else None

            # selected_data_key 属于“当前绑定对象”的运行配置，不参与默认值差异置空。
            selected_data_key = _to_str(tick_draft.get("selected_data_key", ""))
            cfg["axis_tick"][axis]["selected_data_key"] = selected_data_key

            # is_numeric 是绑定对象类型判断结果，同样直接保存。
            is_numeric = bool(tick_draft.get("is_numeric", True))
            cfg["axis_tick"][axis]["is_numeric"] = is_numeric

            # 若当前对象不是数值型，则强制 number_format 回退为 auto。
            number_format = _to_str(tick_draft.get("number_format", "auto")) or "auto"
            cfg["axis_tick"][axis]["number_format"] = number_format if is_numeric else "auto"

            # ---------------------------
            # 显示单位差异收集
            # ---------------------------
            recommended_mag = _to_int(axis_ctx.get("recommended_mag", 0), 0)
            display_mag = tick_draft.get("display_unit_mag", None)
            if not is_numeric or display_mag is None:
                cfg["axis_tick"][axis]["display_unit_mag"] = None
            else:
                try:
                    mag_val = int(display_mag)
                except Exception:
                    mag_val = recommended_mag
                cfg["axis_tick"][axis]["display_unit_mag"] = mag_val if mag_val != recommended_mag else None

        return normalize_label_settings_config(cfg)

    def _parse_font_size_input(self, text: Any, default: int) -> int:
        """
        解析字号输入并限制其有效范围。

        设计规则：
        - 空值：回退 default
        - 非法值：回退 default
        - 小于 8：强制设为 8
        - 大于 72：强制设为 72

        参数：
        - text：用户输入内容
        - default：回退默认值

        返回：
        - 合法字号整数
        """
        raw = _to_str(text)
        if not raw:
            return _to_int(default, default)
        try:
            value = int(float(raw))
        except Exception:
            value = _to_int(default, default)
        if value < 8:
            value = 8
        if value > 72:
            value = 72
        return value

    def _accept_with_collect(self) -> None:
        """
        “确定”按钮的最终提交入口。

        处理流程：
        1. 调用 _collect_final_config 收集并生成最终配置；
        2. 将结果写回 self._config；
        3. 调用 accept() 关闭对话框并返回 Accepted 状态。

        说明：
        - 这里不直接依赖控件当前状态作为对外输出；
        - 对外输出前，必须经过统一的差异收集和结构归一化。
        """
        self._config = self._collect_final_config()
        self.accept()