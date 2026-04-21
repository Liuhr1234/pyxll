# sim_engine.py
"""
本模块提供 Drisk 模拟工作流的传统功能区（Ribbon）与主入口编排功能。

模块定位：
1. 作为 Excel Ribbon 层与底层模拟执行层之间的“总调度入口”。
2. 负责维护功能区状态（迭代次数、情景数、引擎模式、采样方式、随机数设置等）。
3. 在运行模拟前，对当前工作簿执行模型扫描、依赖识别与安全性判断。
4. 根据用户选择与模型特征，调度不同模拟引擎（自动 / numpy / index / 稳健迭代）。
5. 为分布构建器、结果分析窗口、散点分析窗口、模拟设置窗口等 UI 提供宏入口。

主要功能分层：
- Ribbon 状态管理：保存与读取功能区输入状态。
- 模型扫描与校验：识别模型中的分布输入、MakeInput、输出、情景表等对象。
- 引擎决策与执行：根据模型规模、采样方法、引擎兼容性选择执行路径。
- 模拟结果落库：将输入样本与输出结果写入 SimulationResult 缓存。
- UI 触发入口：对接 Qt 弹窗与结果分析界面。

说明：
本文件既承担 UI 入口职责，也承担部分运行时编排职责，因此是 Drisk 模拟主流程中的关键中枢文件。

重要：
接入了三个模拟引擎——numpy_functions->极速引擎//index_functions->标准引擎//simulation_engine->稳健引擎
其中稳健引擎基本失效，长久未维护-->后续可能会删除//需要与后端同学商讨具体安排
"""

# =================================================================
#  1. 核心依赖与全局状态
# =================================================================
# 说明：
# 本节负责导入 Ribbon 编排所需的核心依赖，并定义模块级全局状态。
# 这些全局状态用于保存 Excel Ribbon 的当前选项、模拟设置、缓存扫描结果等。
#
# 特别注意：
# drisk_env 必须在绑定 UI 或 Excel 环境前导入，这是当前工程的既定约束。
import drisk_env

import sys
import traceback
import re
import numpy as np
import threading
from contextlib import contextmanager
from pyxll import xl_macro, xl_app, xlcAlert

from com_fixer import _safe_excel_app

import os as _os, importlib.util as _ilu
_ps6 = _ilu.find_spec('PySide6')
if _ps6 and _ps6.submodule_search_locations:
    _os.add_dll_directory(list(_ps6.submodule_search_locations)[0])
_os.add_dll_directory(r'C:\Windows\System32')
del _os, _ilu, _ps6
from PySide6.QtCore import Qt
from PySide6.QtWidgets import QApplication, QDialog, QHBoxLayout, QLabel, QPushButton, QVBoxLayout

from ui_shared import DRISK_DIALOG_BTN_QSS, set_drisk_icon

from simulation_manager import get_simulation, clear_simulations, _SIMULATION_CACHE, SimulationResult, get_current_sim_id
from dependency_tracker import find_all_simulation_cells_in_workbook
from simulation_engine import iterative_simulation_workbook
from attribute_functions import set_static_mode

# 常量模块采用“可选导入”策略：
# 如果 constants.py 导入失败，则使用当前文件内定义的后备默认值。
try:
    import constants as _constants
except ImportError:
    _constants = None

# 全局默认常量：
# ERROR_MARKER：模拟中的错误占位符
# DEFAULT_SEED：默认随机种子
# DEFAULT_RNG_TYPE：默认随机数生成器类型
# RNG_TYPE_NAMES：随机数引擎枚举与显示名称映射
ERROR_MARKER = str(getattr(_constants, "ERROR_MARKER", "#ERROR!"))
DEFAULT_SEED = int(getattr(_constants, "DEFAULT_SEED", 42))
DEFAULT_RNG_TYPE = int(getattr(_constants, "DEFAULT_RNG_TYPE", 1))
RNG_TYPE_NAMES = dict(
    getattr(
        _constants,
        "RNG_TYPE_NAMES",
        {1: "Mersenne Twister (MT19937)"},
    )
)

# 已接入的分布函数名列表，用于识别公式中出现的分布函数。
try:
    from constants import DISTRIBUTION_FUNCTION_NAMES
except Exception:
    DISTRIBUTION_FUNCTION_NAMES = []

# 三类核心执行/解析依赖：
# 1. index_functions：索引引擎路径
# 2. numpy_functions：numpy 快速模拟路径
# 3. 公式与输入暴露判断工具
import index_functions
try:
    import numpy_functions as numpy_engine
except Exception:
    numpy_engine = None
from cell_utils import normalize_cell_address
from formula_parser import extract_all_attributes_from_formula, extract_input_attributes, parse_formula_references_tornado
from input_sample_exposure import is_input_key_exposed

# backend_bridge 为可编辑桥接层：
# 其作用是将 sim_engine 与可迭代升级的后端兼容逻辑解耦。
# 若导入失败，则 sim_engine 会自动走传统兜底路径。
try:
    import backend_bridge as bridge
except Exception as _bridge_error:
    bridge = None
    print(f"[Drisk][sim_engine] backend_bridge 导入失败: {_bridge_error}")

# --- UI 组件导入 ---
# 这些模块为 Ribbon 宏提供图形界面入口。
# 任一模块导入失败时，会保留入口但弹出“UI 不可用”提示。
try:
    from ui_modeler import show_distribution_gallery
    from ui_results import show_results_dialog
except ImportError as e:
    print(f"Drisk UI 在 sim_engine 中导入失败: {e}")
    show_distribution_gallery = None
    show_results_dialog = None
from ui_scatter import show_scatter_dialog_from_macro, resolve_scatter_sim_id
try:
    from ui_simulation_settings import show_simulation_settings_dialog
except Exception:
    show_simulation_settings_dialog = None

# 迭代次数允许范围：
# DEFAULT_N 为常用默认值；
# MIN_N / MAX_N 控制 Ribbon 输入的合法区间。
DEFAULT_N = 5000
MIN_N = 100
MAX_N = 1000000

# -----------------------------------------------------------------
# Ribbon UI 的全局状态缓存
# -----------------------------------------------------------------
# 说明：
# 这部分变量代表 Ribbon 当前选择状态，是整个 sim_engine 运行时的重要共享状态。
# 这些状态通常由 Ribbon 回调函数写入，由模拟入口读取。
#
# 引擎模式说明：
# auto   : 自动模式（优先 index，不可用则回退迭代）
# numpymc: numpy 快速引擎
# index  : 索引引擎
# robust : 稳健迭代引擎
_ENGINE_AUTO = "auto"
_ENGINE_NUMPY = "numpymc"
_ENGINE_INDEX = "index"
_ENGINE_ROBUST = "robust"
_ENGINE_INDEX_TO_KEY = {
    0: _ENGINE_AUTO,
    1: _ENGINE_NUMPY,
    2: _ENGINE_INDEX,
    3: _ENGINE_ROBUST,
}
_ENGINE_KEY_TO_INDEX = {v: k for k, v in _ENGINE_INDEX_TO_KEY.items()}
_ENGINE_DEFAULT_KEY = _ENGINE_NUMPY

# Ribbon 运行时状态缓存：
# _ribbon_ui：Ribbon 对象句柄，用于主动刷新控件
# _ribbon_iterations：功能区迭代次数文本
# _ribbon_scenarios：功能区情景数量文本
# _ribbon_engine_key / idx：当前选中引擎
# _ribbon_sampling_mode_idx：采样方式（0=MC, 1=LHC）
# _ribbon_rng_type：随机数生成器类型
# _ribbon_seed_mode：随机种子模式（fixed/default）
# _ribbon_seed_value：随机种子值
# _ribbon_sim_scope：样本收集范围（all/collect/none）
# _cached_scan_data：智能检测结果缓存，供下一次模拟直接复用
_ribbon_ui = None
_ribbon_iterations = "1000"
_ribbon_scenarios = "1"
_ribbon_engine_key = _ENGINE_DEFAULT_KEY
_ribbon_engine_idx = _ENGINE_KEY_TO_INDEX[_ribbon_engine_key]
_ribbon_sampling_mode_idx = 0  # 0: MC, 1: LHC
_ribbon_rng_type = int(DEFAULT_RNG_TYPE)
_ribbon_seed_mode = "fixed"  # fixed | default
_ribbon_seed_value = str(int(DEFAULT_SEED))
_ribbon_sim_scope = "all"  # all | collect | none
_cached_scan_data = None  # 智能检测的缓存数据

# 引擎标记到标准引擎 key 的统一注册表。
# 作用：
# 1. 兼容 Ribbon 回调传入的不同格式（索引、文本、旧 token）。
# 2. 保证 UI 层解析与执行层分派使用同一套映射。
_ENGINE_TOKEN_TO_KEY = {
    "engauto": _ENGINE_AUTO,
    "auto": _ENGINE_AUTO,
    "automatic": _ENGINE_AUTO,
    "自动选择": _ENGINE_AUTO,
    "自动": _ENGINE_AUTO,
    "engnumpy": _ENGINE_NUMPY,
    "numpy": _ENGINE_NUMPY,
    "numpymc": _ENGINE_NUMPY,
    "极速引擎numpy": _ENGINE_NUMPY,
    "极速模式": _ENGINE_NUMPY,
    "engindex": _ENGINE_INDEX,
    "index": _ENGINE_INDEX,
    "fast": _ENGINE_INDEX,
    "极速引擎index": _ENGINE_INDEX,
    "标准引擎index": _ENGINE_INDEX,
    "标准模式": _ENGINE_INDEX,
    "engiter": _ENGINE_ROBUST,
    "engrobust": _ENGINE_ROBUST,
    "iter": _ENGINE_ROBUST,
    "iterative": _ENGINE_ROBUST,
    "robust": _ENGINE_ROBUST,
    "稳健引擎": _ENGINE_ROBUST,
    "稳健模式": _ENGINE_ROBUST,
}


def _show_drisk_info_dialog(message: str) -> None:
    """
    显示 Drisk 风格的轻量级单按钮信息弹窗。

    设计目的：
    1. 用统一的 Qt 风格弹窗替代 Excel 原生 xlcAlert，保证视觉风格一致。
    2. 作为 sim_engine 内部统一的提示出口，减少各处散乱调用 Excel 原生弹窗。
    3. 若 Qt 弹窗创建失败，则自动回退到 xlcAlert，保证不会因为 UI 创建失败而丢失提示信息。

    参数：
    - message: 要展示的提示文本。

    说明：
    该函数是本文件内所有“信息提示类弹窗”的统一入口。
    """
    text = str(message or "")
    try:
        # 复用已有 QApplication；若不存在则创建。
        app = QApplication.instance() or QApplication(sys.argv)
        _ = app

        dlg = QDialog()
        dlg.setWindowTitle("Drisk")
        set_drisk_icon(dlg, "simu_icon.svg")
        dlg.setModal(True)
        dlg.setFixedWidth(340)

        # 主布局：上方消息文本 + 下方确定按钮
        layout = QVBoxLayout(dlg)
        layout.setContentsMargins(18, 16, 18, 14)
        layout.setSpacing(12)

        msg_label = QLabel(text)
        msg_label.setWordWrap(True)
        # 允许鼠标选择文本，便于复制错误信息。
        msg_label.setTextInteractionFlags(Qt.TextSelectableByMouse)
        layout.addWidget(msg_label)

        btn_row = QHBoxLayout()
        btn_row.addStretch(1)
        btn_ok = QPushButton("确定")
        btn_ok.setObjectName("btnOk")
        btn_ok.setDefault(True)
        btn_ok.clicked.connect(dlg.accept)
        btn_row.addWidget(btn_ok)
        layout.addLayout(btn_row)

        # 弹窗样式：背景、文本样式 + 项目统一按钮样式
        dlg.setStyleSheet(
            """
            QDialog {
                background-color: #f5f6f8;
            }
            QDialog QLabel {
                color: #333333;
                font-size: 12px;
                font-family: 'Microsoft YaHei';
            }
            """
            + DRISK_DIALOG_BTN_QSS
        )
        dlg.exec()
    except Exception:
        # 若 Qt 弹窗失败，退回 Excel 原生提示，保证信息一定能弹出。
        xlcAlert(text)


@contextmanager
def _capture_engine_alerts():
    """
    捕获底层引擎内发出的 xlcAlert 调用，使 sim_engine 统一接管提示样式。

    设计目的：
    1. 某些引擎模块内部仍直接调用 xlcAlert。
    2. 为了让提示风格统一，本上下文会暂时把这些 xlcAlert 替换成“消息收集器”。
    3. 调用结束后恢复原始 xlcAlert，避免影响其他模块。

    返回：
    - captured_messages: 一个列表，收集引擎运行期间弹出的所有消息文本。

    用法：
    with _capture_engine_alerts() as alert_msgs:
        某引擎执行...
    """
    captured_messages = []

    def _capture_alert(msg):
        captured_messages.append(str(msg or ""))

    # 尝试读取各引擎模块当前持有的 xlcAlert 引用。
    original_index_alert = getattr(index_functions, "xlcAlert", None)
    original_numpy_alert = getattr(numpy_engine, "xlcAlert", None) if numpy_engine is not None else None

    try:
        if original_index_alert is not None:
            index_functions.xlcAlert = _capture_alert
        if numpy_engine is not None and original_numpy_alert is not None:
            numpy_engine.xlcAlert = _capture_alert
        yield captured_messages
    finally:
        # 无论执行成功或失败，最终都恢复原始弹窗函数。
        if original_index_alert is not None:
            index_functions.xlcAlert = original_index_alert
        if numpy_engine is not None and original_numpy_alert is not None:
            numpy_engine.xlcAlert = original_numpy_alert


# 采样模式 token 到索引的映射表。
# 0 = MC（Monte Carlo）
# 1 = LHC（Latin Hypercube）
_SAMPLING_MODE_TOKEN_TO_INDEX = {
    "0": 0,
    "mc": 0,
    "montecarlo": 0,
    "monte carlo": 0,
    "1": 1,
    "lhc": 1,
    "latin": 1,
    "hypercube": 1,
}


# =================================================================
#  2. Excel 功能区 (Ribbon) 交互回调
# =================================================================
# 说明：
# 本节函数主要用于响应 Excel Ribbon 控件事件。
# 这些函数的职责包括：
# 1. 提供当前状态给 Ribbon 显示；
# 2. 接收 Ribbon 输入并写入全局状态；
# 3. 将 UI 的松散输入解析为内部统一的标准状态。

def get_engine_index(control):
    """
    获取当前引擎对应的 Ribbon 下拉项索引。

    返回值：
    - int：当前引擎在 Ribbon 中对应的序号。
    """
    _ = control
    return int(_ENGINE_KEY_TO_INDEX.get(_ribbon_engine_key, _ENGINE_KEY_TO_INDEX[_ENGINE_DEFAULT_KEY]))


def _normalize_engine_index(value, source_hint=""):
    """
    将“数值形式”的引擎输入归一化为标准引擎 key。

    参数：
    - value: 可能来自 Ribbon、配置恢复或旧版本持久化数据的数值。
    - source_hint: 输入来源提示，用于区分新旧语义。

    逻辑说明：
    1. 若来源明确是 Ribbon selected-item-index，则按当前新的四项顺序解释。
    2. 否则走旧版本兼容语义：
       0 -> auto
       1 -> index
       2 / 3 -> robust
    """
    source = str(source_hint or "").strip().lower()

    # Ribbon 的 selected-item-index 应严格按当前 4 项顺序解释。
    if source in ("selected_item_index", "selected_index", "selecteditemindex", "ribbon_state"):
        return _ENGINE_INDEX_TO_KEY.get(int(value))

    # 旧语义兼容路径：兼容历史保存下来的数值含义。
    if int(value) == 0:
        return _ENGINE_AUTO
    if int(value) == 1:
        return _ENGINE_INDEX
    if int(value) in (2, 3):
        return _ENGINE_ROBUST
    return None


def _normalize_engine_choice(value, source_hint=""):
    """
    将一个候选引擎输入归一化为标准引擎 key。

    支持输入类型：
    - None / bool / int / float / str / 可转 int 的对象

    说明：
    该函数是所有“引擎选择解析”的底层公共方法。
    """
    if value is None:
        return None

    if isinstance(value, bool):
        return None

    if isinstance(value, int):
        return _normalize_engine_index(value, source_hint=source_hint)

    if isinstance(value, float):
        if value.is_integer():
            as_int = int(value)
            return _normalize_engine_index(as_int, source_hint=source_hint)
        return None

    if isinstance(value, str):
        text = value.strip()
        if not text:
            return None
        if text.lstrip("+-").isdigit():
            as_int = int(text)
            return _normalize_engine_index(as_int, source_hint=source_hint)

        # 对文本进行 token 匹配，兼容中英文标识及旧名称。
        lowered = text.lower()
        for token, key in _ENGINE_TOKEN_TO_KEY.items():
            if token in lowered:
                return key
        return None

    try:
        as_int = int(value)
    except Exception:
        return None
    return _normalize_engine_index(as_int, source_hint=source_hint)


def _resolve_engine_choice_from_args(args, fallback=_ENGINE_DEFAULT_KEY):
    """
    从 Ribbon 回调的复杂参数集合中解析出最终引擎选择。

    参数：
    - args: Ribbon 回调传入的参数元组，可能包含对象、字典、列表、文本等。
    - fallback: 当解析失败时使用的默认引擎 key。

    解析策略：
    1. 收集所有可能候选值；
    2. 优先从对象属性（selected_item_index / id / value / text 等）中提取；
    3. 倒序遍历候选项，以后出现的值优先；
    4. 无法解析时退回 fallback。

    说明：
    该函数用于适配 Ribbon 回调参数格式不稳定、来源多样的问题。
    """
    candidates = []
    for arg in args:
        candidates.append((arg, "arg"))
        if isinstance(arg, (list, tuple)):
            candidates.extend((item, "arg") for item in arg)
        elif isinstance(arg, dict):
            candidates.extend((item, "arg") for item in arg.values())

        for attr_name in (
            "selected_item_index",
            "selected_index",
            "selectedItemIndex",
            "selected_item_id",
            "selected_id",
            "selectedItemId",
            "id",
            "Id",
            "ID",
            "value",
            "Value",
            "text",
            "Text",
        ):
            try:
                if not hasattr(arg, attr_name):
                    continue
                attr_value = getattr(arg, attr_name)
                if callable(attr_value):
                    attr_value = attr_value()
                candidates.append((attr_value, attr_name))
            except Exception:
                continue

    for candidate, source_hint in reversed(candidates):
        normalized = _normalize_engine_choice(candidate, source_hint=source_hint)
        if normalized is not None:
            return normalized

    fallback_normalized = _normalize_engine_choice(fallback, source_hint="ribbon_state")
    return fallback_normalized if fallback_normalized is not None else _ENGINE_DEFAULT_KEY


def _validate_engine_scope_contract(engine_choice, sim_scope):
    """
    校验当前引擎与模拟范围设置是否满足契约约束。

    参数：
    - engine_choice: 当前引擎 key
    - sim_scope: 当前样本收集范围设置

    返回：
    - (bool, str)：是否通过校验、若失败则返回失败原因

    说明：
    当前版本中始终返回 True，属于预留扩展点。
    后续如需限制“某些引擎不支持某类输入收集范围”，可在此集中增加约束。
    """
    _ = (engine_choice, sim_scope)
    return True, ""


def set_engine(*args):
    """
    设置当前计算引擎。

    说明：
    该函数由 Ribbon 回调触发，将外部输入解析为标准引擎 key，
    并同步更新：
    - _ribbon_engine_key
    - _ribbon_engine_idx
    """
    global _ribbon_engine_key
    global _ribbon_engine_idx
    try:
        _ribbon_engine_key = _resolve_engine_choice_from_args(args, fallback=_ribbon_engine_key)
        _ribbon_engine_idx = int(_ENGINE_KEY_TO_INDEX.get(_ribbon_engine_key, _ENGINE_KEY_TO_INDEX[_ENGINE_DEFAULT_KEY]))
    except Exception as e:
        print(f"Drisk set_engine 失败: {e}")
        pass  # 保证 Ribbon 回调的安全性，避免异常传播到 Excel UI 层。


def get_sampling_mode_index(control):
    """
    获取当前采样方式索引。

    返回：
    - 0：Monte Carlo
    - 1：Latin Hypercube
    """
    return _ribbon_sampling_mode_idx


def _normalize_sampling_mode_choice(value):
    """
    将一个候选采样方式输入归一化为合法索引（0 或 1）。

    支持输入：
    - 整数 / 浮点整数 / 文本 / 可转 int 的对象
    - 文本支持 mc / montecarlo / lhc / latin / hypercube 等关键字
    """
    if value is None:
        return None

    if isinstance(value, bool):
        return None

    if isinstance(value, int):
        return value if value in (0, 1) else None

    if isinstance(value, float):
        if value.is_integer():
            as_int = int(value)
            return as_int if as_int in (0, 1) else None
        return None

    if isinstance(value, str):
        text = value.strip()
        if not text:
            return None
        if text.lstrip("+-").isdigit():
            as_int = int(text)
            return as_int if as_int in (0, 1) else None

        lowered = text.lower()
        for token, idx in _SAMPLING_MODE_TOKEN_TO_INDEX.items():
            if token in lowered:
                return idx
        return None

    try:
        as_int = int(value)
    except Exception:
        return None
    return as_int if as_int in (0, 1) else None


def _resolve_sampling_mode_choice_from_args(args, fallback=0):
    """
    从 Ribbon 回调参数中解析采样方式。

    参数：
    - args: Ribbon 回调传入的各类参数对象
    - fallback: 默认值

    解析思路与引擎选择解析一致：
    1. 收集原始参数本体；
    2. 收集其列表/字典内部值；
    3. 读取常见属性名（selected_index、value、text 等）；
    4. 倒序优先；
    5. 无匹配则退回 fallback。
    """
    candidates = []
    for arg in args:
        candidates.append(arg)
        if isinstance(arg, (list, tuple)):
            candidates.extend(arg)
        elif isinstance(arg, dict):
            candidates.extend(arg.values())

        for attr_name in (
            "selected_item_index",
            "selected_index",
            "selectedItemIndex",
            "selected_item_id",
            "selected_id",
            "selectedItemId",
            "id",
            "Id",
            "ID",
            "value",
            "Value",
            "text",
            "Text",
        ):
            try:
                if not hasattr(arg, attr_name):
                    continue
                attr_value = getattr(arg, attr_name)
                if callable(attr_value):
                    attr_value = attr_value()
                candidates.append(attr_value)
            except Exception:
                continue

    for candidate in reversed(candidates):
        normalized = _normalize_sampling_mode_choice(candidate)
        if normalized is not None:
            return normalized

    fallback_normalized = _normalize_sampling_mode_choice(fallback)
    return fallback_normalized if fallback_normalized is not None else 0


def set_sampling_mode(*args):
    """
    设置当前模拟采样方式（MC / LHC）。

    说明：
    由 Ribbon 下拉框写入，最终更新全局变量 _ribbon_sampling_mode_idx。
    """
    global _ribbon_sampling_mode_idx
    try:
        _ribbon_sampling_mode_idx = _resolve_sampling_mode_choice_from_args(
            args,
            fallback=_ribbon_sampling_mode_idx,
        )
    except Exception as e:
        print(f"Drisk set_sampling_mode 失败: {e}")
        pass  # 保证 Ribbon 回调安全，不中断 Excel 交互。


def on_ribbon_load(ribbon):
    """
    在 Excel 加载 Ribbon 时保存 Ribbon 句柄。

    作用：
    - 保存 _ribbon_ui，以便后续主动刷新指定控件（如 InvalidateControl）。
    """
    global _ribbon_ui
    _ribbon_ui = ribbon


def get_iterations(control):
    """
    获取当前 Ribbon 中缓存的迭代次数文本。
    """
    return _ribbon_iterations


def set_iterations(*args):
    """
    设置 Ribbon 迭代次数文本。

    说明：
    当前实现仅从参数中提取第一个字符串值。
    """
    global _ribbon_iterations
    # 从 Ribbon 回调接收第一个字符串参数
    for arg in args:
        if isinstance(arg, str):
            _ribbon_iterations = arg.strip()
            break


def get_scenarios(control):
    """
    获取当前 Ribbon 中缓存的情景数量文本。
    """
    return _ribbon_scenarios


def set_scenarios(*args):
    """
    设置 Ribbon 情景数量文本。

    说明：
    当前实现仅从参数中提取第一个字符串值。
    """
    global _ribbon_scenarios
    for arg in args:
        if isinstance(arg, str):
            _ribbon_scenarios = arg.strip()
            break


def ribbon_smart_detect(control):
    """
    Ribbon 智能检测入口：扫描当前模型，估计最大情景数量。

    执行流程：
    1. 调用 _perform_model_scan(..., deep_scan=True) 进行深度扫描；
    2. 计算当前模型中最大依赖情景数；
    3. 缓存扫描结果，供下一次模拟直接复用；
    4. 将检测结果写回 Ribbon 的情景输入框；
    5. 在状态栏与提示框中反馈检测结果。

    返回：
    - max_deps：检测得到的最大情景数
    - 若失败则返回 None
    """
    global _ribbon_scenarios, _cached_scan_data
    app = xl_app()
    try:
        app.StatusBar = "正在扫描模型..."
        # 执行深度扫描以估计情景数量上限
        cells_data, limits, max_deps = _perform_model_scan(app, deep_scan=True)

        # 缓存扫描输出，以便下一次模拟运行复用，避免重复扫描
        _cached_scan_data = (cells_data, limits)

        # 将检测到的最大情景数同步到 Ribbon 输入框
        _ribbon_scenarios = str(max_deps)
        if _ribbon_ui:
            _ribbon_ui.InvalidateControl("ebScenarios")

        app.StatusBar = f"检测到 {max_deps} 个情景。"
        _show_drisk_info_dialog(f"检测到 {max_deps} 个情景。")
        return max_deps
    except Exception as e:
        app.StatusBar = False
        _show_drisk_info_dialog(f"智能检测失败: {e}")
        return None


def get_simulation_settings_state() -> dict:
    """
    读取当前模拟设置，并转换为适合 UI 弹窗使用的字典结构。

    返回字段：
    - sampling_mode: "MC" / "LHC"
    - rng_type: 随机数生成器类型编号
    - seed_mode: "fixed" / "default"
    - seed_value: 种子值（字符串形式）
    - sim_scope: 输入收集范围（all / collect / none）
    """
    sampling_mode = "LHC" if int(_ribbon_sampling_mode_idx) == 1 else "MC"
    return {
        "sampling_mode": sampling_mode,
        "rng_type": int(_ribbon_rng_type),
        "seed_mode": str(_ribbon_seed_mode or "fixed"),
        "seed_value": str(_ribbon_seed_value),
        "sim_scope": str(_ribbon_sim_scope or "all"),
    }


def _normalize_simulation_settings_payload(settings: dict) -> dict:
    """
    对外部传入的模拟设置字典进行归一化。

    设计目的：
    1. 允许 UI 对话框返回部分字段；
    2. 对字段进行类型转换与合法性校验；
    3. 保留当前状态中未被修改的项。

    参数：
    - settings: 外部设置字典

    返回：
    - 规范化后的设置字典
    """
    state = get_simulation_settings_state()
    if not isinstance(settings, dict):
        return state

    result = dict(state)

    if "rng_type" in settings:
        try:
            rng_type = int(settings.get("rng_type"))
            if rng_type in RNG_TYPE_NAMES:
                result["rng_type"] = rng_type
        except Exception:
            pass
    if "sampling_mode" in settings:
        sampling_mode = str(settings.get("sampling_mode", "")).strip()
        resolved_mode_idx = _resolve_sampling_mode_choice_from_args(
            (sampling_mode,),
            fallback=_ribbon_sampling_mode_idx,
        )
        result["sampling_mode"] = "LHC" if int(resolved_mode_idx) == 1 else "MC"
    if "seed_mode" in settings:
        seed_mode = str(settings.get("seed_mode", "")).strip().lower()
        if seed_mode in ("fixed", "default"):
            result["seed_mode"] = seed_mode
    if "seed_value" in settings:
        try:
            seed_value = int(float(str(settings.get("seed_value")).strip()))
            result["seed_value"] = str(max(0, seed_value))
        except Exception:
            pass
    if "sim_scope" in settings:
        scope = str(settings.get("sim_scope", "")).strip().lower()
        if scope in ("all", "collect", "none"):
            result["sim_scope"] = scope

    return result


def apply_simulation_settings_state(settings: dict) -> None:
    """
    将规范化后的模拟设置写回 Ribbon 全局状态。

    会更新的全局变量：
    - _ribbon_sampling_mode_idx
    - _ribbon_rng_type
    - _ribbon_seed_mode
    - _ribbon_seed_value
    - _ribbon_sim_scope
    """
    global _ribbon_sampling_mode_idx
    global _ribbon_rng_type
    global _ribbon_seed_mode
    global _ribbon_seed_value
    global _ribbon_sim_scope

    normalized = _normalize_simulation_settings_payload(settings)
    _ribbon_sampling_mode_idx = _resolve_sampling_mode_choice_from_args(
        (normalized.get("sampling_mode", "MC"),),
        fallback=_ribbon_sampling_mode_idx,
    )
    _ribbon_rng_type = int(normalized["rng_type"])
    _ribbon_seed_mode = str(normalized["seed_mode"])
    _ribbon_seed_value = str(normalized["seed_value"])
    _ribbon_sim_scope = str(normalized["sim_scope"])


# =================================================================
#  3. 模型扫描与安全校验（非 UI 逻辑）
# =================================================================
# 说明：
# 本节负责模拟前的“模型识别”与“运行可行性判断”。
# 主要解决两个问题：
# 1. 当前工作簿中有哪些单元格参与 Drisk 模拟？
# 2. 当前模型能否安全地使用 index 高速引擎？

def _is_safe_for_index_engine(app, cells_data, n_iters):
    """
    判断当前模型是否适合使用基于索引（Index）的高速引擎。

    判断规则：
    1. 迭代次数不得超过 100000；
    2. 输入变量总数（分布输入 + MakeInput）不得超过 100；
    3. 输出公式与分布公式中不得包含易变函数。

    易变函数说明：
    - INDIRECT / OFFSET / RAND / NOW / TODAY / CELL
    这些函数会破坏索引模拟对“确定性结构”的假设。

    返回：
    - (bool, str)：是否安全、对应说明信息
    """
    if n_iters > 100000:
        return False, "迭代次数过大，无法使用索引引擎。"

    distribution_cells, simtable_cells, makeinput_cells, output_cells, _, _ = cells_data
    total_inputs = len(distribution_cells) + len(makeinput_cells)

    if total_inputs > 100:
        return False, f"随机输入过多 ({total_inputs} > 100)"

    # 易变函数会破坏确定性的索引模拟
    volatile_funcs = ['INDIRECT', 'OFFSET', 'RAND', 'NOW', 'TODAY', 'CELL']

    def check_formulas(cell_dict):
        """
        检查一组单元格公式中是否存在易变函数。

        参数：
        - cell_dict: 以完整地址为键的单元格集合
        """
        for addr in cell_dict:
            try:
                sheet_name = addr.split('!')[0] if '!' in addr else app.ActiveSheet.Name
                cell_addr = addr.split('!')[-1] if '!' in addr else addr
                formula = str(app.Worksheets(sheet_name).Range(cell_addr).Formula).upper()
                for v_func in volatile_funcs:
                    if v_func in formula:
                        return False, f"检测到易变函数 ({v_func})"
            except Exception:
                # 单个单元格读取失败时，不在此中断，继续检查其他项
                pass
        return True, ""

    # 先检查输出公式，再检查分布输入公式
    safe, msg = check_formulas(output_cells)
    if not safe:
        return False, msg

    safe, msg = check_formulas(distribution_cells)
    if not safe:
        return False, msg

    return True, "模型适用于索引引擎。"


def _safe_find_distribution(formula_str):
    """
    安全解析公式中第一个出现的分布函数，并尽量提取其参数。

    参数：
    - formula_str: Excel 单元格公式文本

    返回：
    - (func_name, dist_params)
      func_name: 分布函数名；未找到时为 None
      dist_params: 解析出的参数列表或桥接层给出的参数结构

    设计思路：
    1. 优先调用 backend_bridge 的标准解析器；
    2. 若 bridge 不可用或失败，则使用当前文件中的正则 + 括号层级扫描兜底；
    3. 永远只识别已登记在 DISTRIBUTION_FUNCTION_NAMES 中的分布函数；
    4. 忽略 DriskName / DriskCategory 等属性类辅助调用。

    使用场景：
    - “打开已有公式进行编辑”时，需要恢复当前单元格中原有分布及参数。
    """
    if not isinstance(formula_str, str):
        return None, {}

    allowed_dist_funcs = {
        str(name).strip().upper(): str(name).strip()
        for name in (DISTRIBUTION_FUNCTION_NAMES or [])
        if str(name).strip()
    }

    # 优先使用 bridge 解析器，以便最大限度恢复用户原始参数表达式。
    try:
        if bridge is not None and hasattr(bridge, "parse_first_distribution_in_formula"):
            parsed = bridge.parse_first_distribution_in_formula(formula_str)
            if parsed:
                func_name = str(parsed.get("func_name", "") or "").strip()
                # 优先保留原始参数文本（除去 Drisk* 属性调用），
                # 以保证“打开已有公式”时能恢复字面表达式，而不是数值化后的默认值。
                raw_args = parsed.get("args_list", []) or []
                dist_params = [
                    str(arg).strip()
                    for arg in raw_args
                    if str(arg).strip() and not str(arg).strip().startswith("Drisk")
                ]
                if not dist_params:
                    dist_params = parsed.get("dist_params", []) or {}
                func_name_upper = func_name.upper()
                if func_name_upper in allowed_dist_funcs:
                    return allowed_dist_funcs[func_name_upper], dist_params
    except Exception:
        pass

    # 兜底逻辑：
    # 仅识别已登记分布函数，不接受泛化的 Drisk* 辅助函数。
    if not allowed_dist_funcs:
        return None, {}

    sorted_names = sorted(
        (re.escape(name) for name in allowed_dist_funcs.values()),
        key=len,
        reverse=True,
    )
    pattern = r'@?\s*(' + "|".join(sorted_names) + r')\s*\('
    match = re.search(pattern, formula_str, re.IGNORECASE)
    if match:
        func_name = match.group(1)
        canonical_name = allowed_dist_funcs.get(func_name.strip().upper(), func_name.strip())
        start = match.end()
        depth = 1
        idx = start

        # 从函数起始括号后开始扫描，找到与之匹配的右括号。
        while idx < len(formula_str) and depth > 0:
            ch = formula_str[idx]
            if ch == "(":
                depth += 1
            elif ch == ")":
                depth -= 1
            idx += 1

        if depth == 0 and idx - 1 >= start:
            args_text = formula_str[start:idx - 1]
            parts = []
            current = []
            paren_level = 0
            brace_level = 0
            bracket_level = 0
            in_quotes = False
            quote_char = ""

            # 参数分隔器：只在“最外层逗号”处分割参数，
            # 避免打断括号、数组、字符串内部内容。
            for ch in args_text:
                if in_quotes:
                    if ch == quote_char:
                        in_quotes = False
                    current.append(ch)
                    continue

                if ch in ('"', "'"):
                    in_quotes = True
                    quote_char = ch
                    current.append(ch)
                    continue
                if ch == "(":
                    paren_level += 1
                elif ch == ")":
                    paren_level -= 1
                elif ch == "{":
                    brace_level += 1
                elif ch == "}":
                    brace_level -= 1
                elif ch == "[":
                    bracket_level += 1
                elif ch == "]":
                    bracket_level -= 1

                if ch == "," and paren_level == 0 and brace_level == 0 and bracket_level == 0:
                    parts.append("".join(current).strip())
                    current = []
                else:
                    current.append(ch)

            if current:
                parts.append("".join(current).strip())

            # 过滤掉属性类辅助函数参数，仅保留真正分布参数。
            dist_params = [
                p for p in parts
                if p and not p.strip().startswith("Drisk")
            ]
            return canonical_name, dist_params

        # 找到函数名但参数未能完整闭合时，至少返回函数名
        return canonical_name, {}
    return None, {}


def _perform_model_scan(app, deep_scan=False):
    """
    执行模型扫描，识别当前工作簿中的 Drisk 相关单元格，并可选进行深度依赖分析。

    参数：
    - app: Excel 应用对象
    - deep_scan: 是否执行深度扫描（用于情景上限估计）

    返回：
    - cells_data:
        (
            distribution_cells,
            simtable_cells,
            makeinput_cells,
            output_cells,
            all_input_keys,
            all_related_cells
        )
    - cell_scenario_limits:
        每个单元格可追溯到的最大情景数上限
    - global_max_deps:
        深度扫描模式下，当前选中区域相关的全局最大情景数

    本函数职责较多，是模拟前的数据准备核心：
    1. 调用 dependency_tracker 识别模型参与对象；
    2. 清理“幽灵引用”（已不再含有对应 Drisk 公式，但缓存中仍残留的地址）；
    3. 调用 bridge 对输出元数据做统一补全；
    4. 若启用深度扫描，则递归追踪依赖链，估计情景数量。
    """

    cells_data = find_all_simulation_cells_in_workbook(app)
    distribution_cells, simtable_cells, makeinput_cells, output_cells, all_input_keys, all_related_cells = cells_data

    # ================= [清理幽灵引用] =================
    # 说明：
    # 某些单元格可能曾经带有 DriskOutput / Drisk / MakeInput / SimTable 标记，
    # 但后续用户已修改公式，导致扫描缓存残留失效引用。
    # 此处通过重新读取公式文本，将“当前已不包含关键字”的条目剔除。
    def _purge_ghosts(cell_dict, keyword):
        if not cell_dict:
            return
        ghosts = []
        for full_addr in cell_dict.keys():
            try:
                sheet_name = full_addr.split('!')[0] if '!' in full_addr else app.ActiveSheet.Name
                cell_addr = full_addr.split('!')[-1] if '!' in full_addr else full_addr
                c = app.Worksheets(sheet_name).Range(cell_addr)
                formula = str(c.Formula).upper()
                if keyword.upper() not in formula:
                    ghosts.append(full_addr)
            except Exception:
                # 读取失败也视为失效引用，后续删除
                ghosts.append(full_addr)
        for g in ghosts:
            del cell_dict[g]

    _purge_ghosts(output_cells, "OUTPUT")
    _purge_ghosts(distribution_cells, "DRISK")
    _purge_ghosts(makeinput_cells, "MAKEINPUT")
    _purge_ghosts(simtable_cells, "SIMTABLE")
    # ==============================================================

    # 输出元数据补全：
    # 若 bridge 可用，则调用 bridge 统一整理 output_cells 元数据，
    # 确保不同入口路径下输出元数据格式一致。
    if bridge is not None and hasattr(bridge, "enrich_output_cells_metadata"):
        try:
            output_cells = bridge.enrich_output_cells_metadata(app, output_cells)
            cells_data = (distribution_cells, simtable_cells, makeinput_cells, output_cells, all_input_keys, all_related_cells)
        except Exception as e:
            print(f"[Drisk][sim_engine] output metadata normalization skipped: {e}")

    scenario_count = 1
    cell_scenario_limits = {}
    global_max_deps = 1

    # 深度扫描逻辑：
    # 用于估计“某个单元格或当前所选区域，最大能展开到多少个情景值”。
    if deep_scan and simtable_cells:
        try:
            simtable_keys = set(simtable_cells.keys())
            visited_for_depth = {}

            def _check_deps_limit(c, current_sheet):
                """
                递归检查一个单元格依赖链上的最大情景数。

                递归规则：
                1. 若当前单元格本身是 SimTable，则其 values 长度可决定局部上限；
                2. 若单元格公式依赖其他单元格，则继续递归其前驱；
                3. 同时解析公式文本中的引用，尽可能覆盖 DirectPrecedents 不完整的情况；
                4. 使用 visited_for_depth 做记忆化，避免重复递归与循环依赖放大。
                """
                try:
                    addr = c.Address.replace('$', '').upper()
                    full_addr = f"{current_sheet}!{addr}"

                    if full_addr in visited_for_depth:
                        return visited_for_depth[full_addr]

                    local_max = 1

                    # 当前格若本身就是 simtable，则 values 数量可视为情景数量候选上限
                    if full_addr in simtable_keys:
                        funcs = simtable_cells.get(full_addr, [])
                        if funcs:
                            vals = funcs[0].get('values', [])
                            local_max = max(local_max, len(vals))

                    formula = str(c.Formula)
                    if formula.startswith('='):
                        # 路径一：通过 Excel DirectPrecedents 追踪直接前驱
                        try:
                            for p in c.DirectPrecedents.Cells:
                                local_max = max(local_max, _check_deps_limit(p, current_sheet))
                        except:
                            pass

                        # 路径二：通过公式解析器补充引用追踪
                        try:
                            refs = parse_formula_references_tornado(formula)
                            for ref in refs:
                                target_sheet, target_addr = ref.upper().split('!') if '!' in ref else (current_sheet, ref.upper())
                                ref_full = f"{target_sheet}!{target_addr}"

                                if ref_full in simtable_keys:
                                    funcs = simtable_cells.get(ref_full, [])
                                    if funcs:
                                        vals = funcs[0].get('values', [])
                                        local_max = max(local_max, len(vals))

                                try:
                                    ref_cell = app.Worksheets(target_sheet).Range(target_addr)
                                    local_max = max(local_max, _check_deps_limit(ref_cell, target_sheet))
                                except:
                                    pass
                        except:
                            pass

                    visited_for_depth[full_addr] = local_max
                    return local_max
                except:
                    return 1

            cells_to_check = []

            # 默认检查对象：输出、分布输入、MakeInput
            for dict_cells in [output_cells, distribution_cells, makeinput_cells]:
                for full_addr in dict_cells.keys():
                    try:
                        sheet_name, cell_addr = full_addr.split('!')
                        c = app.Worksheets(sheet_name).Range(cell_addr)
                        cells_to_check.append(c)
                    except:
                        pass

            # 若用户有当前选区，则优先将选区加入待检查列表，
            # 并记录其地址，用于计算“对当前选区相关”的全局最大情景数。
            selected_addrs = set()
            if app.Selection:
                for c in app.Selection.Cells:
                    cells_to_check.append(c)
                    try:
                        selected_addrs.add(f"{c.Worksheet.Name}!{c.Address.replace('$', '').upper()}")
                    except:
                        pass

            for cell in cells_to_check:
                try:
                    sheet = cell.Worksheet.Name
                    full_addr = f"{sheet}!{cell.Address.replace('$', '').upper()}"

                    if full_addr in cell_scenario_limits:
                        continue

                    limit = _check_deps_limit(cell, sheet)
                    cell_scenario_limits[full_addr] = limit

                    # global_max_deps 只对“当前选区中”且上限 > 1 的单元格进行提升
                    if full_addr in selected_addrs:
                        if limit > 1:
                            global_max_deps = max(global_max_deps, limit)
                except:
                    pass

        except Exception as e:
            print(f"模型深度扫描检查失败: {e}")

    return cells_data, cell_scenario_limits, max(1, global_max_deps)


# =================================================================
#  4. 核心模拟执行管线（Batch Simulation Entry）
# =================================================================
# 说明：
# 本节是 sim_engine 的主执行链路，负责：
# 1. 读取 Ribbon 当前设置；
# 2. 扫描模型；
# 3. 选择模拟引擎；
# 4. 调用对应执行路径；
# 5. 将结果写入 SimulationResult 缓存；
# 6. 恢复 Excel 运行时状态并给出结束提示。

def _run_batch_simulations(method="MC"):
    """
    执行批量模拟主流程。

    参数：
    - method: 采样方式，支持 "MC" 或 "LHC"

    主流程概览：
    A. 获取 Excel 句柄，读取 Ribbon 参数并校验；
    B. 读取缓存扫描结果或执行模型扫描；
    C. 根据引擎模式与模型特征进行引擎决策；
    D. 若走 index / numpy 路径，则直接调用对应引擎；
    E. 若走稳健迭代路径，则按情景循环执行 iterative_simulation_workbook；
    F. 将输入样本与输出结果写入 _SIMULATION_CACHE；
    G. 恢复 Excel 状态并弹出运行结果提示。
    """
    global _cached_scan_data
    try:
        app = _safe_excel_app()  # 获取 COM 安全的 Excel 句柄

        method = str(method or "MC").strip().upper()
        if method not in ("MC", "LHC"):
            method = "MC"

        # ----------------------------------------------------------
        # 1. 读取并校验 Ribbon 参数
        # ----------------------------------------------------------
        try:
            n_iters = int(_ribbon_iterations)
            scenario_count = int(_ribbon_scenarios)
        except ValueError:
            _show_drisk_info_dialog("功能区上的迭代/情景输入无效。")
            return

        if n_iters < MIN_N or n_iters > MAX_N:
            _show_drisk_info_dialog(f"迭代次数必须在 {MIN_N} 和 {MAX_N} 之间。")
            return
        if scenario_count < 1:
            _show_drisk_info_dialog("情景数量必须至少为 1。")
            return

        # 取消事件：当前版本中未暴露取消 UI，但迭代引擎已保留取消事件接口。
        cancel_event = threading.Event()

        # ----------------------------------------------------------
        # 2. 模型扫描：优先复用智能检测缓存
        # ----------------------------------------------------------
        try:
            if _cached_scan_data:
                cells_data, cell_scenario_limits = _cached_scan_data
                _cached_scan_data = None
            else:
                cells_data, cell_scenario_limits, _ = _perform_model_scan(app, deep_scan=False)

            distribution_cells, simtable_cells, makeinput_cells, output_cells, all_input_keys, all_related_cells = cells_data
        except Exception as e:
            _show_drisk_info_dialog(f"模型扫描失败: {e}\n{traceback.format_exc()}")
            return

        cells_data = (distribution_cells, simtable_cells, makeinput_cells, output_cells, all_input_keys, all_related_cells)

        # 若既没有输入也没有输出，则没有模拟目标。
        if not distribution_cells and not makeinput_cells and not output_cells:
            _show_drisk_info_dialog("未找到模拟目标。请先添加 Drisk 公式或输出。")
            return

        # 如果 bridge 提供迭代引用兼容补丁，则在迭代引擎运行前先安装。
        if bridge is not None and hasattr(bridge, "ensure_iterative_ref_compat_patch"):
            try:
                bridge.ensure_iterative_ref_compat_patch()
            except Exception as e:
                print(f"[Drisk][sim_engine] iterative ref compat patch skipped: {e}")

        # ======================================================================
        # 四引擎决策：处理自动 / numpy / index / 稳健路径
        # ======================================================================
        engine_choice = _resolve_engine_choice_from_args((_ribbon_engine_key,), fallback=_ribbon_engine_key)

        is_safe, reason = _is_safe_for_index_engine(app, cells_data, n_iters)
        is_contract_ok, contract_reason = _validate_engine_scope_contract(engine_choice, _ribbon_sim_scope)

        # ----------------------------------------------------------
        # 2A. 强制 Index 引擎
        # ----------------------------------------------------------
        if engine_choice == _ENGINE_INDEX:
            if not is_contract_ok:
                app.StatusBar = f"{contract_reason}，本次运行已停止。"
                _show_drisk_info_dialog(contract_reason)
                return
            # 强制索引模式：直接执行索引引擎
            app.StatusBar = "Drisk 正在运行索引引擎... (强制)"
            try:
                with _capture_engine_alerts() as alert_msgs:
                    index_functions.run_index_simulation(n_iters, scenario_count)
                if alert_msgs:
                    _show_drisk_info_dialog(alert_msgs[-1])
                return
            except ImportError:
                _show_drisk_info_dialog("索引引擎模块不可用 (index_functions.py)。")
                return
            except Exception as e:
                _show_drisk_info_dialog(f"强制使用索引引擎失败: {e}")
                return

        # ----------------------------------------------------------
        # 2B. 强制 numpy 引擎
        # ----------------------------------------------------------
        if engine_choice == _ENGINE_NUMPY:
            if not is_contract_ok:
                app.StatusBar = f"{contract_reason}，本次运行已停止。"
                _show_drisk_info_dialog(contract_reason)
                return
            app.StatusBar = "Drisk 正在运行 numpy 引擎... (强制)"
            if numpy_engine is None or not hasattr(numpy_engine, "run_numpy_simulation"):
                _show_drisk_info_dialog("numpy 引擎模块不可用 (numpy_functions.py)。")
                return
            try:
                with _capture_engine_alerts() as alert_msgs:
                    numpy_engine.run_numpy_simulation(n_iters, scenario_count, method=method)
                if alert_msgs:
                    _show_drisk_info_dialog(alert_msgs[-1])
                return
            except Exception as e:
                _show_drisk_info_dialog(f"强制使用 numpy 引擎失败: {e}")
                return

        # ----------------------------------------------------------
        # 2C. 自动引擎
        # ----------------------------------------------------------
        if engine_choice == _ENGINE_AUTO:
            # 自动模式规则：
            # - 模型安全且采样方式不是 LHC 时，优先尝试 index
            # - index 失败则自动回退到迭代引擎
            allow_index = is_safe and method != "LHC"
            if allow_index:
                app.StatusBar = f"Drisk 正在运行索引引擎... (自动: {reason})"
                try:
                    # 捕获 index 引擎中的警告消息；
                    # 若消息中出现失败/错误/取消等关键词，则认定 index 运行失败并回退。
                    with _capture_engine_alerts() as alert_msgs:
                        index_functions.run_index_simulation(n_iters, scenario_count)

                    failed = any(
                        ("fail" in str(m).lower())
                        or ("error" in str(m).lower())
                        or ("cancel" in str(m).lower())
                        or ("失败" in str(m))
                        or ("错误" in str(m))
                        or ("取消" in str(m))
                        or ("中止" in str(m))
                        for m in alert_msgs
                    )
                    if not failed:
                        if alert_msgs:
                            _show_drisk_info_dialog(alert_msgs[-1])
                        return
                    app.StatusBar = "索引引擎失败。回退到迭代引擎。"
                except ImportError:
                    app.StatusBar = "索引引擎不可用。回退到迭代引擎。"
                except Exception:
                    app.StatusBar = "索引引擎出现错误。回退到迭代引擎。"
            else:
                if method == "LHC":
                    reason = "LHC 采样需使用迭代引擎"
                app.StatusBar = f"Drisk 正在运行迭代引擎... (自动回退: {reason})"

        # ----------------------------------------------------------
        # 2D. 强制稳健迭代引擎
        # ----------------------------------------------------------
        if engine_choice == _ENGINE_ROBUST:
            # 强制迭代模式
            app.StatusBar = "Drisk 正在运行迭代引擎... (强制)"

        # ======================================================================

        # 在新的运行前清除模拟缓存，避免旧结果污染当前运行。
        clear_simulations()

        # 关闭静态模式，使模拟计算可以正常动态刷新。
        set_static_mode(False)
        total_runs = 0

        # 保存 Excel 现场状态，稍后恢复。
        old_screen = app.ScreenUpdating
        old_calc = app.Calculation
        old_events = app.EnableEvents

        try:
            app.ScreenUpdating = False
            app.EnableEvents = False

            # 强制整本工作簿重新计算，以刷新依赖公式。
            app.CalculateFull()

            # 切换到手动计算，避免模拟过程中 Excel 自动重算造成性能损失。
            app.Calculation = -4135  # xlCalculationManual（手动计算）

            # ------------------------------------------------------
            # 3. 按情景逐个执行迭代引擎
            # ------------------------------------------------------
            for scenario_idx in range(scenario_count):
                if cancel_event.is_set():
                    break

                # 情景 ID 从 1 开始，以保持与旧缓存键兼容。
                sim_id = scenario_idx + 1
                if sim_id not in _SIMULATION_CACHE:
                    sim = SimulationResult(sim_id, n_iters, method)
                    _SIMULATION_CACHE[sim_id] = sim
                else:
                    sim = _SIMULATION_CACHE[sim_id]
                    sim.n_iterations = max(sim.n_iterations, n_iters)

                # 为当前情景重置结果容器。
                sim.output_cache = {}
                sim.input_cache = {}
                sim.output_attributes = {}
                sim.input_attributes = {}

                # 写入本次情景的全局运行信息。
                sim.set_scenario_info(scenario_count, scenario_idx)
                try:
                    sim.workbook_name = app.ActiveWorkbook.Name
                except Exception:
                    pass
                try:
                    sim.global_rng_type = int(_ribbon_rng_type)
                except Exception:
                    sim.global_rng_type = int(DEFAULT_RNG_TYPE)
                sim.global_seed_mode = str(_ribbon_seed_mode or "fixed")
                try:
                    sim.global_seed_value = int(float(str(_ribbon_seed_value).strip()))
                except Exception:
                    sim.global_seed_value = int(DEFAULT_SEED)
                sim.input_scope_mode = str(_ribbon_sim_scope or "all")
                sim.input_collection_scope_mode = str(_ribbon_sim_scope or "all")

                # 将模型扫描结果附加到 SimulationResult，供后续分析界面使用。
                sim.distribution_cells = distribution_cells
                sim.simtable_cells = simtable_cells
                sim.makeinput_cells = makeinput_cells
                sim.output_cells = output_cells
                sim.all_input_keys = all_input_keys

                # ======================================================================
                # 进度回调：更新状态栏显示当前情景的完成百分比
                # ======================================================================
                def update_progress(completed_iterations):
                    try:
                        pct = (completed_iterations / n_iters) * 100
                        app.StatusBar = f"Drisk: 情景 {scenario_idx+1}/{scenario_count} - {pct:.1f}%"
                    except:
                        pass

                # 针对当前情景执行工作簿级模拟。
                input_samples, output_values, output_info = iterative_simulation_workbook(
                    app=app,
                    distribution_cells=distribution_cells,
                    simtable_cells=simtable_cells,
                    makeinput_cells=makeinput_cells,
                    output_cells=output_cells,
                    n_iterations=n_iters,
                    sampling_method=method,
                    scenario_count=scenario_count,
                    scenario_index=scenario_idx,
                    progress_callback=update_progress,
                    cancel_event=cancel_event
                )

                # --------------------------------------------------
                # 4. 若有结果，则将输入/输出写入 SimulationResult
                # --------------------------------------------------
                if input_samples or output_values:
                    was_cancelled = cancel_event.is_set()
                    if was_cancelled:
                        # 若中途取消，则按实际拿到的输出长度估算有效迭代次数。
                        actual_iterations = 0
                        for data in output_values.values():
                            actual_iterations = max(actual_iterations, len(data))
                        sim.n_iterations = actual_iterations
                    else:
                        sim.n_iterations = max(sim.n_iterations, n_iters)

                    # ==============================================================
                    # A. 收集“分布输入（非 MakeInput）”的采样结果
                    # ==============================================================
                    for cell_addr, samples_list in input_samples.items():
                        # MakeInput 由后续单独处理，此处跳过
                        if cell_addr in makeinput_cells:
                            continue
                        sheet_name = cell_addr.split('!')[0] if '!' in cell_addr else app.ActiveSheet.Name
                        c_addr_only = cell_addr.split('!')[-1] if '!' in cell_addr else cell_addr

                        if cell_addr in distribution_cells:
                            for i, samples in enumerate(samples_list):
                                if i >= len(distribution_cells[cell_addr]):
                                    continue
                                dist_func = distribution_cells[cell_addr][i]
                                is_nested = dist_func.get('is_nested', False)

                                # 输入键名策略：
                                # - 嵌套分布尽量沿用已有 key
                                # - 非嵌套分布使用“单元格_索引”命名
                                if is_nested:
                                    input_key = dist_func.get('key', f"{c_addr_only}_{dist_func.get('index', i+1)}")
                                    if '!' in input_key:
                                        input_key = input_key.split('!')[1]
                                else:
                                    input_key = f"{c_addr_only}_{dist_func.get('index', i+1)}"

                                # 从 full_match 中恢复输入属性
                                attrs = extract_input_attributes(f"={dist_func['full_match']}") if 'full_match' in dist_func else {}
                                if 'full_match' in dist_func:
                                    attrs['formula'] = f"={dist_func['full_match']}"
                                if is_nested:
                                    attrs['is_nested'] = True
                                    if dist_func.get('parent_makeinput'):
                                        attrs['parent_makeinput'] = dist_func.get('parent_makeinput')

                                if isinstance(samples, np.ndarray) and len(samples) > 0:
                                    sim.add_input_result(input_key, samples, sheet_name, attrs)

                    # ==============================================================
                    # B. 收集 MakeInput 的采样结果
                    # ==============================================================
                    for cell_addr, samples_list in input_samples.items():
                        if cell_addr not in makeinput_cells:
                            continue
                        sheet_name = cell_addr.split('!')[0] if '!' in cell_addr else app.ActiveSheet.Name
                        c_addr_only = cell_addr.split('!')[-1] if '!' in cell_addr else cell_addr

                        samples = samples_list[0]
                        if isinstance(samples, np.ndarray) and len(samples) > 0:
                            # 将 ERROR 标记统一映射为 ERROR_MARKER，确保后续分析层识别一致。
                            valid_arr = np.array([ERROR_MARKER if v == ERROR_MARKER or str(v).strip().upper() == "#ERROR!" else v for v in samples], dtype=object)
                            attrs = {'is_makeinput': True}
                            makeinput_func = makeinput_cells[cell_addr][0] if len(makeinput_cells[cell_addr]) > 0 else None
                            if makeinput_func:
                                attrs['expression'] = makeinput_func.get('formula', '')
                                attrs.update(makeinput_func.get('attributes', {}))

                            # 兜底属性提取：
                            # 当 parser 侧对 MakeInput 嵌套属性切片不完整时，
                            # 再直接从当前单元格公式中提取 name / units / category。
                            formula_text = ""
                            try:
                                formula_text = str(app.ActiveWorkbook.Worksheets(sheet_name).Range(c_addr_only).Formula)
                            except Exception:
                                formula_text = ""
                            if formula_text:
                                formula_attrs = extract_all_attributes_from_formula(formula_text) or {}
                                for attr_key in ("name", "units", "category"):
                                    attr_value = str(formula_attrs.get(attr_key, "") or "").strip()
                                    if attr_value and not str(attrs.get(attr_key, "") or "").strip():
                                        attrs[attr_key] = attr_value

                            # 标记“名称来源于显式属性”，供 UI 层后续优先级判断使用。
                            if str(attrs.get("name", "") or "").strip():
                                attrs["_name_from_attr"] = True
                            sim.add_input_result(f"{c_addr_only}_MakeInput", valid_arr, sheet_name, attrs)

                    # ==============================================================
                    # C. 收集计算输出的采样结果
                    # ==============================================================
                    for cell_addr, data in output_values.items():
                        # 为保持兼容性，当前策略是保留所有输出数据
                        sheet_name = cell_addr.split('!')[0] if '!' in cell_addr else app.ActiveSheet.Name
                        c_addr_only = cell_addr.split('!')[-1] if '!' in cell_addr else cell_addr

                        out_info = output_info.get(cell_addr, {})
                        if cell_addr in output_cells:
                            out_info.update(output_cells[cell_addr])
                        if cell_addr in cell_scenario_limits:
                            out_info['max_scenarios'] = cell_scenario_limits[cell_addr]

                        if isinstance(data, np.ndarray) and len(data) > 0:
                            sim.add_output_result(c_addr_only, data, sheet_name, out_info)

                else:
                    _show_drisk_info_dialog(f"情景 {scenario_idx + 1} 模拟失败。")

                total_runs += 1

        finally:
            # ------------------------------------------------------
            # 5. 恢复静态模式与 Excel 运行时状态
            # ------------------------------------------------------
            set_static_mode(True)
            try:
                app.ScreenUpdating = old_screen
                app.EnableEvents = old_events
                app.Calculation = old_calc
                app.CalculateFull()

                # 恢复 Excel 状态栏控制权
                app.StatusBar = False
            except:
                pass

        # ----------------------------------------------------------
        # 6. 运行结束提示
        # ----------------------------------------------------------
        if not cancel_event.is_set():
            if scenario_count > 1:
                _show_drisk_info_dialog(f"已完成 {scenario_count} 个情景。")
            else:
                _show_drisk_info_dialog("模拟已完成。")
        else:
            _show_drisk_info_dialog("模拟已取消。")

    except Exception as e:
        # 最外层兜底：若整个模拟流程异常，尽量恢复 Excel 到可操作状态。
        try:
            set_static_mode(True)
            xl = xl_app()
            xl.ScreenUpdating = True
            xl.EnableEvents = True
            xl.Calculation = -4105
            xl.CalculateFull()
            xl.StatusBar = False
        except:
            pass
        _show_drisk_info_dialog(f"模拟失败。\n{e}\n{traceback.format_exc()}")


# -----------------------------------------------------------------
# Ribbon 绑定宏包装器
# -----------------------------------------------------------------
# 说明：
# 这些函数是暴露给 Excel Ribbon 的宏入口。
# 它们通常只承担“根据当前 Ribbon 状态调用主流程”的职责。

@xl_macro()
def drisk_run_mc(control=None):
    """
    从 Ribbon 直接触发 Monte Carlo 模拟。
    """
    _run_batch_simulations("MC")


@xl_macro()
def drisk_run_lhc(control=None):
    """
    从 Ribbon 直接触发 Latin Hypercube 模拟。
    """
    _run_batch_simulations("LHC")


@xl_macro()
def drisk_run_selected(control=None):
    """
    根据当前 Ribbon 上已选采样方式，自动触发 MC 或 LHC 模拟。

    规则：
    - mode_idx == 1 -> LHC
    - 其他情况 -> MC
    """
    mode_idx = _resolve_sampling_mode_choice_from_args((_ribbon_sampling_mode_idx,), fallback=0)
    if mode_idx == 1:
        _run_batch_simulations("LHC")
        return
    _run_batch_simulations("MC")


@xl_macro()
def drisk_clear_cache(control=None):
    """
    清除模拟缓存。

    优先级：
    1. 若 macros.DriskClear 可用，则优先走统一宏实现；
    2. 否则直接调用 clear_simulations() 兜底。
    """
    try:
        from macros import DriskClear
        DriskClear()
    except Exception:
        try:
            clear_simulations()
            _show_drisk_info_dialog("Drisk: 模拟缓存已清除。")
        except Exception as e2:
            _show_drisk_info_dialog(f"Drisk: 清除缓存失败: {e2}")


@xl_macro()
def drisk_fix_com(control=None):
    """
    触发 COM 修复宏。

    说明：
    当前通过 macros.DriskFixCOM 实现；若失败则提示错误信息。
    """
    try:
        from macros import DriskFixCOM
        DriskFixCOM()
    except Exception as e:
        _show_drisk_info_dialog(f"COM 修复失败: {e}")


# =================================================================
#  5. 建模器与分析视图入口（UI Triggers）
# =================================================================
# 说明：
# 本节提供多个 Ribbon/宏入口，用于打开各类 UI：
# - 模拟设置
# - 输出标记
# - 分布构建器
# - 结果分析对话框
# - 散点分析对话框

@xl_macro()
def open_simulation_settings(control=None):
    """
    打开“模拟设置”对话框。

    流程：
    1. 读取当前全局模拟设置；
    2. 传入 ui_simulation_settings 对话框；
    3. 若用户确认并返回更新后的配置，则写回 Ribbon 全局状态。
    """
    if show_simulation_settings_dialog is None:
        _show_drisk_info_dialog("Drisk UI 不可用: 缺少 ui_simulation_settings.py")
        return

    try:
        updated = show_simulation_settings_dialog(current_settings=get_simulation_settings_state())
        if isinstance(updated, dict):
            apply_simulation_settings_state(updated)
            try:
                xl_app().StatusBar = "模拟设置已更新。"
            except Exception:
                pass
    except Exception as e:
        _show_drisk_info_dialog(f"打开模拟设置失败: {e}\n{traceback.format_exc()}")


@xl_macro()
def drisk_mark_output(control=None):
    """
    为当前选中单元格追加 DriskOutput 输出标记。

    设计目的：
    - 将普通计算单元格标记为 Drisk 输出，供模拟后统一收集与结果展示。

    行为说明：
    1. 若当前单元格无公式，则提示失败；
    2. 若公式已包含 DriskOutput，则仅更新状态栏，不重复添加；
    3. 否则在原公式尾部追加 +DriskOutput("地址")。
    """
    try:
        xl = xl_app()
        cell = xl.ActiveCell
        if not cell:
            return

        formula = str(cell.Formula)
        addr = cell.Address.replace("$", "")

        if not formula:
            _show_drisk_info_dialog("所选单元格中未找到公式。")
            return

        if "DriskOutput" in formula:
            xl.StatusBar = f"{addr} 已标记为输出。"
            return

        if formula.startswith("="):
            cell.Formula = f'{formula}+DriskOutput("{addr}")'
        else:
            cell.Formula = f'={formula}+DriskOutput("{addr}")'

        xl.StatusBar = f"输出标记已添加到 {addr}。"

    except Exception as e:
        _show_drisk_info_dialog(f"标记输出失败: {e}\n{traceback.format_exc()}")


@xl_macro()
def open_distribution_gallery(control=None):
    """
    打开“分布库 / 分布构建器”弹窗。

    典型用途：
    - 在当前选中单元格上新建分布公式；
    - 打开已有分布公式并恢复其参数、属性进行编辑。

    流程：
    1. 读取当前单元格公式；
    2. 解析已有分布函数与参数；
    3. 提取公式中的属性（名称、类别、单位等）；
    4. 打开 ui_modeler 对话框；
    5. 若用户返回新公式，则写回 Excel。
    """
    if show_distribution_gallery is None:
        _show_drisk_info_dialog("Drisk UI 不可用: 缺少 ui_modeler.py")
        return

    try:
        xl = xl_app()
        cell = xl.ActiveCell
        if cell is None:
            xl.StatusBar = "Drisk: "
            return

        # 优先读取 Formula2，兼容动态数组/新公式语法；失败则退回 Formula。
        try:
            current_formula = str(cell.Formula2)
        except Exception:
            current_formula = str(cell.Formula)

        cell_addr = cell.Address.replace("$", "")

        # 解析当前公式中已有的分布函数与属性
        dist_key, params = _safe_find_distribution(current_formula)
        try:
            attrs = extract_all_attributes_from_formula(current_formula)
        except Exception:
            attrs = {}

        # 打开分布构建器
        new_formula = show_distribution_gallery(
            initial_dist=dist_key,
            initial_params=params,
            initial_attrs=attrs,
            full_formula=current_formula,
            cell_address=cell_addr
        )

        # 将新生成的公式写回 Excel
        if new_formula:
            try:
                cell.Formula2 = new_formula
            except Exception:
                cell.Formula = new_formula

            try:
                cell.Calculate()
            except Exception:
                pass

            xl.StatusBar = f"Drisk: {cell_addr} 中的公式已更新"
        else:
            xl.StatusBar = "Drisk: 公式未更新。"

    except Exception as e:
        _show_drisk_info_dialog(f"打开分布库失败: {str(e)}\n{traceback.format_exc()}")


@xl_macro()
def smart_plot_macro(control=None):
    """
    从当前 Excel 选区打开结果分析对话框（智能图表入口）。

    功能目标：
    - 用户先选中一个或多个单元格，再触发本宏；
    - 系统自动判断这些单元格对应的是“输出结果”还是“输入样本”；
    - 整理标签与键名后，打开统一的结果分析窗口。

    说明：
    该函数是“Excel 选择 -> 模拟缓存 -> 结果窗口”的关键桥接入口。
    """

    def _legacy_prepare_result_targets(sim_id_value, requested_keys):
        """
        传统兜底匹配逻辑：当 bridge 侧选择解析失败时，保留旧版匹配语义。

        输入：
        - sim_id_value: 当前模拟缓存 ID
        - requested_keys: 用户选区对应的完整单元格地址列表

        输出：
        - valid_keys_local: 可用于结果对话框的数据键
        - labels_local: 展示标签
        - final_kind_local: "output" 或 "input"
        - missing_local: 未匹配到缓存数据的键

        匹配策略：
        1. 优先在 output_cache 中按完整键匹配；
        2. 若未命中，则尝试在 input_cache 中按输入暴露规则匹配；
        3. 输入键支持“地址前缀”匹配（如 A1 与 A1_1、A1_2）。
        """
        sim = get_simulation(sim_id_value)
        if sim is None:
            return [], {}, "output", list(requested_keys or [])

        output_cache = getattr(sim, "output_cache", {}) or {}
        input_cache = getattr(sim, "input_cache", {}) or {}
        output_attrs = getattr(sim, "output_attributes", {}) or {}
        input_attrs = getattr(sim, "input_attributes", {}) or {}

        valid_keys_local = []
        labels_local = {}
        final_kind_local = "output"
        missing_local = []

        for full_key in requested_keys or []:
            key_clean = str(full_key).replace("$", "").upper()
            matched_output = None

            # 先匹配输出缓存
            for cache_key in output_cache.keys():
                if str(cache_key).replace("$", "").upper() == key_clean:
                    matched_output = cache_key
                    break

            if matched_output is not None:
                valid_keys_local.append(matched_output)
                attrs = output_attrs.get(matched_output, {}) or {}
                name = attrs.get("name")
                display_addr = str(matched_output).split("!")[-1] if "!" in str(matched_output) else str(matched_output)
                labels_local[matched_output] = f"{display_addr} ({name})" if name else display_addr
                final_kind_local = "output"
                continue

            found_input = False
            for input_key in input_cache.keys():
                if not is_input_key_exposed(sim, input_key):
                    continue
                input_clean = str(input_key).replace("$", "").upper()
                if input_clean == key_clean or input_clean.startswith(f"{key_clean}_"):
                    found_input = True
                    valid_keys_local.append(input_key)
                    attrs = input_attrs.get(input_key, {}) or {}
                    name = attrs.get("name")
                    display_addr = str(input_key).split("!")[-1] if "!" in str(input_key) else str(input_key)
                    labels_local[input_key] = f"{display_addr} ({name})" if name else display_addr
                    final_kind_local = "input"

            if not found_input:
                missing_local.append(full_key)

        return valid_keys_local, labels_local, final_kind_local, missing_local

    try:
        app = _safe_excel_app()
        if app is None:
            _show_drisk_info_dialog("Excel 应用程序不可用。")
            return

        selection = app.Selection
        if not selection:
            return

        try:
            sheet_name = str(app.ActiveSheet.Name)
        except Exception:
            sheet_name = ""

        # 先尝试从 bridge 获取默认 sim_id
        sim_id = None
        if bridge is not None and hasattr(bridge, "get_default_result_sim_id"):
            try:
                sim_id = bridge.get_default_result_sim_id()
            except Exception:
                sim_id = None

        # 若 bridge 未提供，则退回到缓存中的首个 sim_id
        if sim_id is None:
            try:
                if isinstance(_SIMULATION_CACHE, dict) and _SIMULATION_CACHE:
                    sim_id = next(iter(_SIMULATION_CACHE.keys()))
            except Exception:
                sim_id = None

        if sim_id is None:
            _show_drisk_info_dialog("没有可用的模拟缓存。请先运行模拟。")
            return

        # 将当前选区转为完整键名列表
        requested_keys = []
        for cell in selection.Cells:
            addr = normalize_cell_address(cell)
            if not addr:
                continue
            full_key = addr if "!" in addr else f"{sheet_name}!{addr}"
            requested_keys.append(full_key)

        valid_keys = []
        labels = {}
        kind = "output"
        missing_keys = []

        # 优先使用 bridge 的统一选择解析逻辑
        if bridge is not None and hasattr(bridge, "prepare_result_dialog_selection"):
            try:
                valid_keys, labels, kind, missing_keys = bridge.prepare_result_dialog_selection(
                    int(sim_id),
                    requested_keys,
                    default_sheet=sheet_name,
                    default_kind="output",
                )
            except Exception:
                valid_keys, labels, kind, missing_keys = _legacy_prepare_result_targets(sim_id, requested_keys)
        else:
            valid_keys, labels, kind, missing_keys = _legacy_prepare_result_targets(sim_id, requested_keys)

        for miss_key in missing_keys:
            print(f"[Drisk][结果入口] 找不到键名: {miss_key}")

        if not valid_keys:
            _show_drisk_info_dialog(
                "未找到所选单元格的模拟数据。\n\n"
                "1. 确保已选择输出单元格或分布公式。\n"
                "2. 打开智能图表前请先运行模拟。\n"
                "3. 调用智能图表前请先选择目标单元格。"
            )
            return

        show_results_dialog(sim_id=sim_id, cell_keys=valid_keys, labels=labels, kind=kind)
        app.StatusBar = "图表已就绪"

    except Exception as e:
        err_msg = traceback.format_exc()
        print(f"[Drisk][结果入口] 打开结果对话框失败。\n{err_msg}")
        _show_drisk_info_dialog(f"打开结果对话框失败。\n{str(e)}")

#压力测试入口，对应代码在stress_testing目录下，包含ui和引擎逻辑
@xl_macro()
def open_stress_test(control=None):
    """打开压力测试配置窗口并执行测试。"""
    try:
        #ensure_qapp()
        from stress_testing.ui_stress_test import StressTestInputDialog
        from PySide6.QtWidgets import QDialog
        #dlg:前端交互获取数据
        dlg = StressTestInputDialog()
        if dlg.exec() == QDialog.Accepted:
            cfg = dlg.get_config()
            try:
                from stress_testing.stress_test_engine import run_stress_test
                run_stress_test(cfg)
            except Exception as e:
                import traceback as _tb
                _show_drisk_info_dialog(f"压力测试执行失败：\n{e}\n\n{_tb.format_exc()}")
    except Exception as e:
        try:
            xlcAlert(f"打开压力测试窗口失败: {e}")
        except Exception:
            pass

@xl_macro()
def open_scatter_viewer(control=None):
    """
    打开散点分析对话框。

    流程：
    1. 解析当前可用的模拟缓存 ID；
    2. 若无有效缓存，则提示先运行模拟；
    3. 确保 QApplication 已存在；
    4. 将后续 UI 编排交给 ui_scatter 模块处理。

    说明：
    sim_engine 在这里仅承担入口职责，不负责散点图界面的后续逻辑。
    """
    global _global_scatter_dialog
    try:
        sim_id = resolve_scatter_sim_id(get_current_sim_id())
        if sim_id is None:
            _show_drisk_info_dialog("未找到有效的模拟缓存。请先运行模拟。")
            return

        # 确保 Qt 应用程序实例存在
        app = QApplication.instance() or QApplication(sys.argv)
        _ = app

        # 将所有散点图 UI 编排委托给 ui_scatter 模块
        show_scatter_dialog_from_macro(sim_id)

    except Exception as e:
        error_msg = traceback.format_exc()
        print(f"打开散点图对话框失败。\n{error_msg}")
        try:
            _show_drisk_info_dialog(f"打开散点分析对话框失败。\n{e}")
        except Exception:
            pass