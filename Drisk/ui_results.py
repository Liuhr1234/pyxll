# ui_results.py
# 这一版实现了统计指标更新线程化操作
"""
===================================================================
UI 结果展示模块 (Drisk Pro版)
===================================================================
特性:
1. 1-2-5 智能刻度步长
2. 智能视野范围 (自动过滤 1%-99% 之外的长尾)
3. 完美直方图切割 + 叠加分析
4. 极简现代 UI 风格

增强:
- y轴采用1-2-5刻度规则，并且上界对齐刻度网格:
  y_axis_max = ceil(raw_y_max / y_dtick) * y_dtick
  例如 y_dtick=2, raw_y_max=11 -> y_axis_max=12

✅ [本次修复核心]:
1. 彻底移除了原先的 combo_chart_type 和 chk_overlay_cdf 控件残留。
2. 删除了由于控件废弃而变得冗余的 _refresh_chart_type_options 等死代码。
3. 实现了通过按钮弹出层级菜单 (Submenu) 进行视图切换与复合叠加指令下发。
4. 【第2组架构升级】将底层历史模拟扫描与数据直接提取重定向至 cache_functions。
===================================================================
"""

# =======================================================
# 1. 模块级依赖与环境初始化
# =======================================================
# -----------------------------------------------------------------
# 本模块架构说明
# -----------------------------------------------------------------
# 1) UltimateRiskDialog: 模拟结果“直方图分析”主窗口（支持单图 + 多图叠加）。
#    - 单图模式: 显示 PDF 曲线（KDE 估计），并显示 1 行概率输入框。
#    - 多图叠加: 不显示 PDF；最多 3 组数据显示“Left/Mid/Right”填充分区；第 4 组及以后只画轮廓。
#    - 不论画不画填充，所有“展示/生效”的数据都会参与:
#      * x/y 坐标轴范围与刻度计算
#      * 右侧统计指标（多列）
# 2) OverlaySelectionDialog: 选择“哪些变量要叠加展示”的勾选窗口。
# 3) CDFDialog: 单独弹出的 CDF 查看窗口（当前逻辑: 从顶部下拉进入）。

# 🔴 核心警告: 必须放在所有第三方库（尤其是 PySide6 和 pyxll）之前！ -- 不能删除，导入即生效
import drisk_env

import re
import sys
import os
import json
import math
import time
import numpy as np
import pandas as pd
import traceback
import plotly.graph_objects as go
import plotly.utils

from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QMessageBox, QLabel, QFrame,
    QApplication, QWidget, QLineEdit,
    QTableWidget, QTableWidgetItem, QComboBox, QAbstractItemView, QFormLayout,
    QPushButton, QSplitter, QSizePolicy, QStackedWidget, QMenu
)
from PySide6.QtCore import Qt, QTimer, Slot, QSignalBlocker, QEvent
from PySide6.QtWebEngineCore import QWebEnginePage
from PySide6.QtGui import QColor

# 是否开启 Web 调试模式（不想输出日志请改为 False）
DEBUG_WEB = True  


from dataclasses import dataclass
from typing import Any, Dict, List, Optional as _Optional

from com_fixer import _safe_excel_app
from formula_parser import extract_all_attributes_from_formula

# 前后端关键桥梁适配器
import backend_bridge as bridge
import theory_adapter as theory

# 引入绘图宿主与遮罩层组件
from ui_overlay import DriskVLineMaskOverlay
from plotly_host import PlotlyHost
# 引入统计指标计算管理器
import ui_stats
# 引入样式配置管理器
from ui_style_manager import StyleManagerDialog
# 引入标准后台计算任务（Worker）
from ui_workers import ModelerStatsJob
from ui_results_data_service import ResultsDataService
from ui_results_job_controller import ResultsJobController

# 重新导出已剥离的对话框组件，保持 ui_results 作为对外兼容的统一门面接口 (Compatibility Facade)
from ui_results_dialogs import AnalysisObjectDialog, AdvancedOverlayDialog
from ui_results_menu import ResultsMenuRouter
from ui_results_modes import ResultsViewCommandHelper, ResultsRenderDispatcher
from ui_results_advanced import ResultsAdvancedRouter
from ui_results_advanced_prepare import (
    AdvancedPreparationError,
    AdvancedPreparedPayload,
    build_advanced_payload,
    resolve_output_data,
)
from ui_results_advanced_render import ResultsAdvancedRenderService
from ui_results_main_render import ResultsMainRenderService
from ui_results_interactions import ResultsInteractionService
from ui_results_runtime_state import ResultsRuntimeStateHelper
from ui_results_advanced_state import (
    build_default_tornado_styles,
)
# 引入底层绘图工厂
from drisk_charting import DriskChartFactory
# 引入高级视图组件：龙卷风图、箱线图
from ui_tornado import UITornadoView, ScenarioSettingsDialog, TornadoSettingsDialog
from ui_boxplot import UIBoxplotView
# 引入底层模拟数据仓库，用于探测历史模拟版本
from simulation_manager import get_all_simulations
# 引入共享 UI 界面常量与工具方法
from ui_shared import (
    SimpleRangeSlider, drisk_combobox_qss, SnapUtils, ChartOverlayController,
    infer_si_mag, create_floating_value_with_mag, update_floating_value_edits_pos, set_drisk_icon, 
    get_shared_webview, recycle_shared_webview, ChartSkeleton, get_plotly_cache_dir,
    DRISK_COLOR_CYCLE, DRISK_COLOR_CYCLE_THEO, LayoutRefreshCoordinator, apply_toolbar_button_icon,
    extract_clean_cell_address, resolve_visible_variable_name,
)
from ui_label_settings import (
    LabelSettingsDialog,
    create_default_label_settings_config,
    get_axis_display_unit_override,
    get_axis_numeric_flags,
)
from ui_sim_display_names import (
    extract_sim_id_from_series_key,
    get_sim_display_name,
)


# =======================================================
# 2. 基础设施与数据模型类
# =======================================================

class DebugPage(QWebEnginePage):
    """
    用于拦截和捕获 QtWebEngine 内核 JavaScript 日志的调试类。
    主要用于在开发模式下（DEBUG_WEB=True）捕获前端 HTML/JS 运行时的控制台输出，
    方便排查 Plotly 图表渲染与通信时的错误。
    """
    def javaScriptConsoleMessage(self, level, message, lineNumber, sourceID):
        """处理 JS 控制台消息并打印到 Python 控制台"""
        print(f"[WEB_CONSOLE][level={int(level)}][L{lineNumber}] {sourceID}: {message}")


@dataclass
class BackendResultsInput:
    """
    用于描述结果对话框的数据源（基于后端驱动模式的数据类）。
    包含了模拟 ID、单元格键值列表以及显示标签，替代了旧版直接向 UI 传递臃肿的 Pandas DataFrame，以极大地减少内存占用。
    """
    sim_id: int
    cell_keys: List[str]
    labels: Dict[str, str]  # 单元格键值 -> 显示标签
    kind: str = "output"


class ResultsMenuFactory:
    """
    绘图模式菜单路由代理的工厂类。
    为了保持向后兼容性而设计的门面类，主要负责将主窗口对各类菜单（分布图、龙卷风图、摘要图等）
    的调用请求，安全地转发给新版架构中的 ui_results_menu 模块。
    """

    @staticmethod
    def show_tornado_menu(dialog):
        """显示敏感性(龙卷风)图的弹出菜单"""
        ResultsMenuRouter.show_tornado_menu(dialog)

    @staticmethod
    def show_boxplot_menu(dialog):
        """显示摘要(箱线)图的弹出菜单"""
        ResultsMenuRouter.show_boxplot_menu(dialog)

    @staticmethod
    def show_scenario_menu(dialog):
        """显示情景分析图的弹出菜单"""
        ResultsMenuRouter.show_scenario_menu(dialog)

    @staticmethod
    def show_view_mode_menu(dialog):
        """显示标准分布视图(PDF/CDF/直方图)的弹出菜单"""
        ResultsMenuRouter.show_view_mode_menu(dialog)


# =======================================================
# 3. 核心主窗口视图类
# =======================================================

class UltimateRiskDialog(QDialog):
    """
    核心结果分析主窗口。
    支持单图展示、多组数据叠加对比，以及高级视图分析（龙卷风图、情景分析、箱线图）。
    负责桥接后端数据抽取与前端 Plotly/WebEngine 渲染。
    """

    # ---------------------------------------------------------
    # 3.1 初始化与基础 UI 构建 (Initialization & Base UI)
    # ---------------------------------------------------------

    def __init__(self, data_input, parent=None):
        """
        窗口的构造函数。
        负责解析传入的数据源，初始化系统缓存、线程池调度器、以及底层状态变量，最后调用 init_ui() 完成渲染。
        """
        super().__init__(parent)
        self.setWindowTitle("Drisk - 模拟结果分析")

        # ✅ 在设置标题后，调用设置设定自定义图标
        set_drisk_icon(self)

        # 🔴 自适应屏幕大小逻辑 (适配屏幕 90% 宽，80% 高)
        screen_geo = QApplication.primaryScreen().availableGeometry()
        target_w = min(1200, int(screen_geo.width() * 0.90))
        target_h = min(800, int(screen_geo.height() * 0.80))
        self.resize(target_w, target_h)
        self.setWindowFlags(
            self.windowFlags()
            | Qt.WindowMinimizeButtonHint
            | Qt.WindowMaximizeButtonHint
            | Qt.WindowCloseButtonHint
        )

        # 移除硬编码的 QComboBox 样式，保留其他组件样式
        base_qss = """
            QDialog { background-color: white; font-family: 'Microsoft YaHei'; }
            /* 紧凑布局: 减小输入框的内边距 padding */
            QLineEdit { border: 1px solid #d9d9d9; border-radius: 3px; padding: 2px 4px; font-weight: bold; }
            QLineEdit[invalid="true"] { border: 1px solid #ff4d4f; background: #fff1f0; }
            QLineEdit:focus { border-color: #40a9ff; }
            QGroupBox { font-weight: bold; border: 1px solid #eee; border-radius: 5px; margin-top: 10px; padding-top: 10px; }
            QGroupBox::title { subcontrol-origin: margin; left: 10px; padding: 0 3px; }
        """
        
        # 将基础样式与 ui_shared 中的高级下拉框样式拼接，实现全局统一
        self.setStyleSheet(base_qss + drisk_combobox_qss())

        # 实例化图名标签
        self.chart_title_label = QLabel()
        self.chart_title_label.setAlignment(Qt.AlignCenter)
        self._apply_chart_title_label_style()

        # --- 1. 数据解析（两种模式: 旧版直接数据 / 新版后端驱动） ---
        self.backend_input: BackendResultsInput | None = None
        self.sim_id: int | None = None
        self.label_to_cell_key: dict[str, str] = {}
        self.attrs_map: dict[str, dict] = {}

        if isinstance(data_input, BackendResultsInput):
            # 新版后端驱动模式（sim_id + cell_keys）
            if bridge is None:
                raise RuntimeError("backend_bridge 模块不可用，无法加载后端结果输入。")
            self.backend_input = data_input
            self.sim_id = int(data_input.sim_id)

            # 记录查询凭证，并探测内存中有哪些包含当前变量的历史版本
            self._base_kind = data_input.kind
            self._base_labels = data_input.labels or {}
            self._base_cell_keys = []
            for ck in (data_input.cell_keys or []):
                try:
                    self._base_cell_keys.append(bridge.normalize_cell_key(ck))
                except Exception:
                    self._base_cell_keys.append(ck)
            
            # 探测可用的模拟 ID
            self.available_sims = []
            try:
                all_sims = get_all_simulations()
                for sid, sim in all_sims.items():
                    cache = sim.output_cache if self._base_kind == "output" else getattr(sim, 'input_cache', {})
                    
                    is_valid = False
                    for ck in self._base_cell_keys:
                        if ck in cache:
                            data = cache[ck]
                            # ✅ 1. 拦截基础报错与越界无效数据
                            if len(data) > 0:
                                v = str(data[0]).strip().upper()
                                if not (v.startswith('#') or 'NAN' in v):
                                    is_valid = True
                                    break
                    if is_valid:
                        self.available_sims.append(sid)
            except Exception as e:
                print(f"探测历史版本失败: {e}")
                self.available_sims = [self.sim_id]

            # 拉取并缓存样本数据到 all_datasets（UI 内部仍用 label 作为 key）
            all_datasets: dict[str, object] = {}
            for ck in (data_input.cell_keys or []):
                try:
                    nck = bridge.normalize_cell_key(ck)
                    label = (data_input.labels or {}).get(nck) or (data_input.labels or {}).get(ck) or nck
                    arr = bridge.get_series(self.sim_id, nck, kind=data_input.kind)
                    all_datasets[label] = arr
                    self.label_to_cell_key[label] = nck
                    try:
                        self.attrs_map[label] = bridge.get_attributes(self.sim_id, nck, kind=data_input.kind) or {}
                    except Exception:
                        self.attrs_map[label] = {}
                except Exception:
                    continue

            # 兜底：如果没有任何可用数据，分配零值占位符
            if not all_datasets:
                all_datasets = {"当前单元格": np.array([0.0] * 10)}

            self.all_datasets = all_datasets

        elif isinstance(data_input, dict):
            # 旧版模式: 外部直接传入 dict[label -> Series/ndarray/list]
            self.all_datasets = data_input

        else:
            # 旧版模式: 单组 Series/ndarray/list
            self.all_datasets = {"当前单元格": data_input}

        self.dataset_keys = list(self.all_datasets.keys())

        # --- 2. 缓存与状态变量初始化 ---
        self._kde_cache = {}
        self._cdf_sorted_cache = {}
        self._cdf_ci_cache = {}  # ✅ 新增：CDF 置信区间专属缓存

        self.current_key = self.dataset_keys[0]
        self.overlay_keys = []

        self.manual_mag = None  # 用于存储用户手动选择的量级
        self._label_settings_config = create_default_label_settings_config()
        self._label_axis_numeric = {"x": True, "y": True}
        self._chart_title_display_overrides = {}

        # 坐标轴与边界控制
        self.x_dtick = 1.0
        self.view_min = 0.0
        self.view_max = 1.0
        self.MARGIN_L = 80
        self.MARGIN_R = 40
        self.MARGIN_T = 0
        self.MARGIN_B = 60
        self.ADDITIONAL_Y2_MARGIN = 55  # 定义副 Y 轴（yaxis2）占据的额外物理宽度

        # 拖拽与吸附交互参数
        self._drag_snap_enabled = True
        self._drag_snap_px_per_step = 2
        self._cont_fixed_steps = 500
        self._cont_fixed_enabled = True
        self._drag_grid_step = None
        self._drag_grid_xmin = None
        self._drag_grid_xmax = None
        self._discrete_points = None
        self._last_sent_lr = None
        self._in_slider_sync = False
        self._first_open = True

        # ✅ 必须最先创建临时目录，防止 init_ui 中的组件调用报错
        self._tmp_dir = get_plotly_cache_dir()

        # ✅ 接入全新的并发调度控制器
        self.job_controller = ResultsJobController(self)
        self.job_controller.pdf_kde_done.connect(self._on_pdf_kde_done)
        self.job_controller.hist_kde_done.connect(self._on_hist_kde_done)
        self.job_controller.stats_done.connect(self._on_stats_done)
        self.job_controller.cdf_stats_done.connect(self._on_cdf_stats_done)

        # 保留必要的 UI 状态与绘图缓存
        self._stats_cache_valid = False
        self._stats_val_maps = {}
        self._pdf_x_grid = None
        self._pdf_y_map = None
        self._pdf_keys_sig = None

        # 滑块拖动防抖定时器
        self.throttle_timer = QTimer(self)
        self.throttle_timer.setSingleShot(True)
        self.throttle_timer.setInterval(30)
        self.throttle_timer.timeout.connect(self.perform_buffered_update)

        # ✅ 初始化视图指令与叠加状态标志
        self._current_view_data = "auto"
        self._show_cdf_overlay = False
        self._show_kde_overlay = False
        self.chart_mode = "hist"
        self.theory_dist_obj = None
        self._theory_func_name = ""
        self._theory_params = []
        self._theory_markers = {}
        self._theory_is_discrete = None

        # ✅ CDF专属置信度缓存（默认95%）
        self.cdf_ci_level = 0.95
        
        # 🔴 新增：敏感性与情景分析专属的样式存储槽
        self._tornado_styles = build_default_tornado_styles(DRISK_COLOR_CYCLE)
        self._tornado_line_styles = {} # 趋势图模式动态加载
        self._last_major_mode = "main" # 🔴 记录当前的四大模式状态，用于判断是否需要重置
        
        # 初始化高级视图配置
        ResultsRuntimeStateHelper.init_advanced_configs(self)

        # =========================================================
        # --- 3. 初始化 UI 布局 (此时 _tmp_dir 和线程变量都已就绪) ---
        # =========================================================
        self.init_ui()

        # 初始化 Plotly 绘图宿主
        self._plot_host = PlotlyHost(
            web_view=self.web_view,
            tmp_dir=self._tmp_dir,
            use_qt_overlay=getattr(self, "_use_qt_overlay", False),
            overlay=getattr(self, "overlay", None),
        )

        self._webview_loaded = False
        self._data_sent = False
        self._render_token = 0
        self._waiting_token = 0
        self._js_inflight = False
        self._ready_deadline = 0.0

        # Web 引擎就绪状态轮询
        self._ready_poll_timer = QTimer(self)
        self._ready_poll_timer.setInterval(50)
        self._ready_poll_timer.timeout.connect(self._poll_plotly_ready)

        try:
            self.web_view.loadFinished.connect(self._on_webview_load_finished)
        except Exception:
            pass

        # 加载初始分析数据集
        self.load_dataset(self.current_key)

        self._first_open = False
        
        # --- 初始化防抖与节流定时器 (用于窗口 Resize) ---
        self._last_plotly_resize_time = 0
        self._plotly_resize_timer = QTimer(self)
        self._plotly_resize_timer.setSingleShot(True)
        self._plotly_resize_timer.setInterval(150)
        self._plotly_resize_timer.timeout.connect(self._force_plotly_resize)

        self._first_open = False
        self.web_view.setFocus()

    def init_ui(self):
        """
        主 UI 布局构建函数。
        将窗口分为：顶部概率输入区、左侧 Web 绘图区、右侧多列统计指标表、底部功能按钮控制栏。
        使用 QSplitter 实现左右两部分宽度的自适应拖拽。
        """
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)

        # 主内容区 (使用 QSplitter 进行动态宽度分割)
        content_splitter = QSplitter(Qt.Horizontal)
        self.content_splitter = content_splitter
        content_splitter.setHandleWidth(2)
        content_splitter.setStyleSheet("QSplitter::handle { background-color: #c0c0c0; }")
        content_splitter.splitterMoved.connect(self._on_content_splitter_moved)

        # --- 左侧绘图区 ---
        # ✅ 1. 恢复纯净的宿主面板
        self.chart_container = QWidget()
        self.chart_container.setMinimumWidth(400)
        self.chart_container.setStyleSheet("background-color: #ffffff;")

        # ✅ 2. 恢复原生单层垂直布局
        left_vbox = QVBoxLayout(self.chart_container)
        left_vbox.setContentsMargins(0, 0, 0, 0)
        left_vbox.setSpacing(0)

        # ---------------- 顶栏滑块区 ----------------
        self.top_wrapper = QWidget(self.chart_container)
        self.top_wrapper.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        top_layout = QVBoxLayout(self.top_wrapper)
        top_layout.setContentsMargins(0, 0, 0, 0)
        top_layout.setSpacing(0)

        top_layout.addWidget(self.chart_title_label)

        # 创建带量级指示的浮动输入框
        self.float_x_l = create_floating_value_with_mag(self.top_wrapper, height=22, opacity=0.95)
        self.float_x_r = create_floating_value_with_mag(self.top_wrapper, height=22, opacity=0.95)
        self.float_x_l.raise_()
        self.float_x_r.raise_()

        self.edt_x_l = self.float_x_l.edit
        self.lbl_x_suffix_l = self.float_x_l.mag
        self.edt_x_l.returnPressed.connect(lambda: self.update_logic('xl'))

        self.edt_x_r = self.float_x_r.edit
        self.lbl_x_suffix_r = self.float_x_r.mag
        self.edt_x_r.returnPressed.connect(lambda: self.update_logic('xr'))

        # 初始化范围边界滑块 (Range Slider)
        self.range_slider = SimpleRangeSlider()
        self.range_slider.setMargins(self.MARGIN_L, self.MARGIN_R)
        self.range_slider.rangeChanged.connect(self.on_range_slider_changed)
        self.range_slider.dragStarted.connect(self.on_slider_drag_started)
        self.range_slider.dragFinished.connect(self.on_slider_drag_finished)
        self.range_slider.bind_line_edits(self.edt_x_l, self.edt_x_r)
        self.range_slider.inputConfirmed.connect(self._on_slider_input_confirmed)

        try:
            _float_h = int(max(self.float_x_l.sizeHint().height(), self.float_x_r.sizeHint().height()))
        except Exception:
            _float_h = 22
        _top_pad = int(_float_h + 2)  
        
        top_layout.addSpacing(_top_pad)
        top_layout.addWidget(self.range_slider)

        try:
            _slider_h = int(self.range_slider.sizeHint().height())
            _title_h = self.chart_title_label.sizeHint().height() or 24
            self.top_wrapper.setMinimumHeight(_title_h + _top_pad + _slider_h)
        except Exception:
            self.top_wrapper.setMinimumHeight(24 + _top_pad + 40)
            
        left_vbox.addWidget(self.top_wrapper, 0)

        # ---------------- 图表栈区 ----------------
        self.view_stack = QStackedWidget(self.chart_container)
        left_vbox.addWidget(self.view_stack, 1)

        # =========================================================
        # ✅ 3. 【真·悬浮遮罩核心】将加载动画骨架屏设为独立悬浮子控件，绝对不加入任何 layout！
        self.chart_skeleton = ChartSkeleton(self.chart_container)
        self.chart_skeleton.hide()
        
        # ✅ 4. 开启事件拦截器，保障悬浮层尺寸完美同步
        self.chart_container.installEventFilter(self)
        # =========================================================

        # 3.1 组装主直方图包装器
        self.main_chart_wrapper = QWidget()
        main_chart_layout = QVBoxLayout(self.main_chart_wrapper)
        main_chart_layout.setContentsMargins(0, 0, 0, 0)
        main_chart_layout.setSpacing(0)

        # 获取共享的 Web 引擎视图
        self.web_view = get_shared_webview(parent_widget=self.main_chart_wrapper)
        if DEBUG_WEB:
            self.web_view.setPage(DebugPage(self.web_view))

        self.web_view.setStyleSheet("background: transparent;")
        self.web_view.page().setBackgroundColor(Qt.white)

        # 初始化 Qt 原生绘图遮罩（取代以前性能低下的 JS 虚线绘图）
        self._use_qt_overlay = True
        self.overlay = DriskVLineMaskOverlay(
            self.web_view,
            margin_l=self.MARGIN_L, margin_r=self.MARGIN_R,
            margin_t=self.MARGIN_T, margin_b=self.MARGIN_B
        )
        self.overlay.setGeometry(self.web_view.rect())
        self.overlay.raise_()
        self.overlay.show()

        self.chart_skeleton.raise_()
        self._chart_ctrl = ChartOverlayController(self.web_view, self.overlay)
        self._layout_refresh = LayoutRefreshCoordinator(
            self,
            frame_cb=self._layout_refresh_frame,
            final_cb=self._layout_refresh_final,
            frame_ms=50,
            settle_ms=180,
        )

        main_chart_layout.addWidget(self.web_view)

        # 3.2 实例化高级分析视图（龙卷风图、箱线图）
        self.tornado_view = UITornadoView(self, tmp_dir=self._tmp_dir)
        self.boxplot_view = UIBoxplotView(self, tmp_dir=self._tmp_dir)

        # 3.3 按层级推入 Stack (由 _activate_analysis_view 路由调度)
        self.view_stack.addWidget(self.main_chart_wrapper) # 默认索引 0 (基础分布图)
        self.view_stack.addWidget(self.tornado_view)       # 索引 1 (敏感性分析)
        self.view_stack.addWidget(self.boxplot_view)       # 索引 2 (摘要图)

        # 将左侧绘图面板加入 Splitter
        content_splitter.addWidget(self.chart_container)

        # --- 右侧统计表区 ---
        right_panel = QFrame()
        self.right_panel = right_panel  # ✅ [新增这一行] 保存引用以便后续隐藏
        right_panel.setMinimumWidth(130)  
        right_panel.setStyleSheet("background-color: #ffffff;")
        right_layout = QVBoxLayout(right_panel)
        right_layout.setContentsMargins(0, 0, 0, 0)

        header = QLabel("   统计指标（模拟）")
        header.setFixedHeight(36)
        header.setStyleSheet(
            "background-color: #fafafa; color: #333; font-weight: bold; font-size: 14px; border-bottom: 1px solid #f0f0f0;")
        right_layout.addWidget(header)

        self.stats_table = QTableWidget()
        self.stats_table.setVerticalScrollMode(QAbstractItemView.ScrollPerPixel)
        self.stats_table.setHorizontalScrollMode(QAbstractItemView.ScrollPerPixel)
        self.stats_table.setColumnCount(1)
        self.stats_table.verticalHeader().hide()
        self.stats_table.horizontalHeader().setVisible(True)
        self.stats_table.setAlternatingRowColors(False) 

        right_layout.addWidget(self.stats_table)
        self.stats_table.cellClicked.connect(self.on_stats_cell_clicked)

        # 将右侧统计表面板加入 Splitter
        content_splitter.addWidget(right_panel)

        # 设置 Splitter 的初始比例与拉伸因子
        content_splitter.setSizes([720, 150])
        content_splitter.setStretchFactor(0, 1)  # 左侧吸收多余空间
        content_splitter.setStretchFactor(1, 0)  # 右侧尽可能保持自身设定

        main_layout.addWidget(content_splitter)

        # ==========================================
        # 底部控制栏（全新分组与带间距布局）
        # ==========================================
        bottom_bar = QFrame()
        bottom_bar.setFixedHeight(34)  
        bottom_bar.setStyleSheet("""
            QFrame { background-color: #f5f5f5; border-top: 1px solid #e0e0e0; }
            QLabel { font-size: 12px; color: #555; }
        """)
        bottom_layout = QHBoxLayout(bottom_bar)
        bottom_layout.setContentsMargins(15, 1, 15, 1) 
        
        # ✅ 设定组内小间距
        bottom_layout.setSpacing(6) 

        # 统一的标准按钮样式定义
        standard_btn_style = """
            QPushButton { background-color: #f0f0f0; border: 1px solid #d9d9d9; border-radius: 3px; padding: 0px 8px; font-size: 12px; height: 20px;}
            QPushButton:hover { border-color: #40a9ff; color: #40a9ff; background-color: #e6f7ff; }
            QPushButton:checked { background-color: #e6f7ff; border-color: #1890ff; color: #1890ff; font-weight: bold; }
        """

        # --- 【第1组】：分析对象、叠加对比、图形样式、文本设置 ---
        self.btn_analysis_obj = QPushButton("分析对象")
        self.btn_analysis_obj.setObjectName("btnAnalysisObj")
        self.btn_analysis_obj.setStyleSheet(standard_btn_style)
        self.btn_analysis_obj.clicked.connect(self._open_analysis_obj_dialog)
        bottom_layout.addWidget(self.btn_analysis_obj)

        self.btn_overlay = QPushButton("叠加对比")
        self.btn_overlay.setStyleSheet(standard_btn_style)
        self.btn_overlay.clicked.connect(self.open_overlay_selector)
        bottom_layout.addWidget(self.btn_overlay)

        self.btn_style = QPushButton("图形样式")
        self.btn_style.setStyleSheet(standard_btn_style)
        self.btn_style.clicked.connect(self._open_style_manager)
        bottom_layout.addWidget(self.btn_style)

        self.btn_label_settings = QPushButton("文本设置")
        self.btn_label_settings.setStyleSheet(standard_btn_style)
        self.btn_label_settings.clicked.connect(self._open_label_settings_dialog)
        bottom_layout.addWidget(self.btn_label_settings)

        # ✅ 组间距 = 组内间距 + 额外增加部分 -- 暂定为 2.5 倍内间距
        bottom_layout.addSpacing(9)

        # --- 【第2组】：分布图、摘要图、敏感性图、情景图 ---
        self.btn_view_mode = QPushButton("分布图")
        self.btn_view_mode.setObjectName("btnViewMode")
        self.btn_view_mode.setStyleSheet(standard_btn_style)
        self.btn_view_mode.clicked.connect(lambda: ResultsMenuFactory.show_view_mode_menu(self))
        bottom_layout.addWidget(self.btn_view_mode)

        self.btn_ci_settings = QPushButton("CI设置")
        self.btn_ci_settings.setStyleSheet(standard_btn_style)
        self.btn_ci_settings.clicked.connect(self._open_cdf_ci_settings_dialog)
        bottom_layout.addWidget(self.btn_ci_settings)
        apply_toolbar_button_icon(
            self.btn_ci_settings,
            "settings_icon.svg",
            icon_px=20,
            icon_only=True,
            button_px=24,
        )
        self.btn_ci_settings.setToolTip("设置：配置CI引擎和置信区间。")
        self.btn_ci_settings.hide()

        self.btn_boxplot = QPushButton("摘要图")
        self.btn_boxplot.setCheckable(True) 
        self.btn_boxplot.setStyleSheet(standard_btn_style)
        self.btn_boxplot.clicked.connect(lambda: ResultsMenuFactory.show_boxplot_menu(self))
        bottom_layout.addWidget(self.btn_boxplot)

        self.btn_tornado = QPushButton("敏感性图")
        self.btn_tornado.setCheckable(True)
        self.btn_tornado.setStyleSheet(standard_btn_style)
        self.btn_tornado.clicked.connect(lambda: ResultsMenuFactory.show_tornado_menu(self))
        bottom_layout.addWidget(self.btn_tornado)

        # 敏感性图的专属设置按钮紧跟其后
        self.btn_tornado_settings = QPushButton("设置")
        self.btn_tornado_settings.setStyleSheet(standard_btn_style)
        self.btn_tornado_settings.clicked.connect(self._open_tornado_settings)
        self.btn_tornado_settings.hide() 
        bottom_layout.addWidget(self.btn_tornado_settings)

        self.btn_scenario = QPushButton("情景图")
        self.btn_scenario.setCheckable(True)
        self.btn_scenario.setStyleSheet(standard_btn_style)
        self.btn_scenario.clicked.connect(lambda: ResultsMenuFactory.show_scenario_menu(self))
        bottom_layout.addWidget(self.btn_scenario)

        # 情景图的专属设置按钮紧跟其后
        self.btn_scenario_settings = QPushButton("设置")
        self.btn_scenario_settings.setStyleSheet(standard_btn_style)
        self.btn_scenario_settings.clicked.connect(lambda: self._open_scenario_settings(is_custom=False))
        self.btn_scenario_settings.hide()
        bottom_layout.addWidget(self.btn_scenario_settings)

        # 应用底部工具栏图标
        toolbar_specs = [
            (self.btn_analysis_obj, "target_icon.svg", "分析对象：切换当前用于分析的目标变量。", 24, 28),
            (self.btn_overlay, "overlay_icon.svg", "叠加对比：选择并叠加其它情景进行对比。", 24, 28),
            (self.btn_style, "style_icon.svg", "图形样式：配置颜色、线条和图形外观。", 24, 28),
            (self.btn_label_settings, "text_icon.svg", "文本设置：调整图名、轴标题和刻度文本样式。", 24, 28),
            (self.btn_view_mode, "distribution_icon.svg", "分布图：切换并配置分布图显示模式。", 24, 28),
            (self.btn_boxplot, "boxplot_icon.svg", "摘要图：查看样本分布的摘要统计图。", 24, 28),
            (self.btn_tornado, "tornado_icon.svg", "敏感性图：查看输入对结果影响程度。", 24, 28),
            (self.btn_tornado_settings, "settings_icon.svg", "设置：配置敏感性图参数。", 20, 24),
            (self.btn_scenario, "scene_icon.svg", "情景图：查看不同情景下的结果对比。", 24, 28),
            (self.btn_scenario_settings, "settings_icon.svg", "设置：配置情景图参数。", 20, 24),
        ]
        for _btn, _icon_name, _tooltip, _icon_px, _btn_px in toolbar_specs:
            apply_toolbar_button_icon(_btn, _icon_name, icon_px=_icon_px, icon_only=True, button_px=_btn_px)
            _btn.setToolTip(_tooltip)

        # --- 尾部弹性空间，将后续组件推到右侧 ---
        bottom_layout.addStretch()

        # --- 【第3组】：引擎选项 与 数据量级 ---
        self.lbl_ci_engine = QLabel("CI引擎:")
        bottom_layout.addWidget(self.lbl_ci_engine)
        
        self.combo_ci_engine = QComboBox()
        self.combo_ci_engine.addItems(["快速 (Fast)", "标准 (Bootstrap)", "科学 (BCa)"])
        self.combo_ci_engine.setFixedHeight(20)
        self.ci_engine_mode = "fast"
        self.combo_ci_engine.currentIndexChanged.connect(self._on_ci_engine_changed)
        bottom_layout.addWidget(self.combo_ci_engine)
        
        self.lbl_ci_engine.hide()
        self.combo_ci_engine.hide()

        # ✅ 组间距 = 组内间距 + 额外增加部分 -- 暂定为 2.5 倍内间距
        bottom_layout.addSpacing(9)

        # --- 【第4组】：导出 与 关闭 ---
        self.btn_export = QPushButton("导出")
        self.btn_export.setStyleSheet("""
            QPushButton { 
                background-color: #0050b3; color: white; border: none; border-radius: 3px; font-weight: bold;
                font-size: 12px; padding: 0px 10px; height: 20px; 
            }
            QPushButton:hover { background-color: #40a9ff; }
            QPushButton:pressed { background-color: #003a8c; }
        """)
        self.btn_export.clicked.connect(self.on_export_clicked)
        bottom_layout.addWidget(self.btn_export)

        btn_close = QPushButton("关闭")
        btn_close.setStyleSheet("""
            QPushButton { 
                background-color: #555555; color: white; border: none; border-radius: 3px; font-weight: bold;
                font-size: 12px; padding: 0px 10px; height: 20px;
            }
            QPushButton:hover { background-color: #777777; }
            QPushButton:pressed { background-color: #333333; }
        """)
        btn_close.clicked.connect(self.reject) 
        bottom_layout.addWidget(btn_close)

        # 防止回车键隐式触发底部栏的任何按钮（取消默认按钮行为）
        for _btn in bottom_bar.findChildren(QPushButton):
            _btn.setAutoDefault(False)
            _btn.setDefault(False)

        main_layout.addWidget(bottom_bar)
        self._install_chart_context_menu_hooks()

    # ---------------------------------------------------------
    # 3.2 布局协调与事件处理 (Layout Coordination & Events)
    # ---------------------------------------------------------

    def eventFilter(self, obj, event):
        """
        事件过滤器：拦截左侧绘图容器尺寸变化 (Resize) 的事件。
        用于同步更新骨架屏的大小及依赖布局的交互组件，确保视觉效果不脱节。
        """
        if obj == getattr(self, "chart_container", None) and event.type() == QEvent.Resize:
            if hasattr(self, "chart_skeleton") and self.chart_skeleton:
                self.chart_skeleton.setGeometry(self.chart_container.rect())
            self._refresh_layout_bound_inputs()
        return super().eventFilter(obj, event)

    def _install_chart_context_menu_hooks(self):
        """安装图表右键菜单的事件挂钩"""
        self._bind_chart_context_menu(getattr(self, "web_view", None))
        tornado_view = getattr(getattr(self, "tornado_view", None), "web_view", None)
        boxplot_view = getattr(getattr(self, "boxplot_view", None), "web_view", None)
        self._bind_chart_context_menu(tornado_view)
        self._bind_chart_context_menu(boxplot_view)

    def _bind_chart_context_menu(self, web_view):
        """为指定的 Web 引擎视图绑定自定义右键菜单"""
        if web_view is None:
            return
        web_view.setContextMenuPolicy(Qt.CustomContextMenu)
        try:
            web_view.customContextMenuRequested.disconnect(self._on_chart_context_menu_requested)
        except Exception:
            pass
        web_view.customContextMenuRequested.connect(self._on_chart_context_menu_requested)

    def _on_chart_context_menu_requested(self, pos):
        """响应右键点击，构建并弹出对应的工具菜单项"""
        source = self.sender()
        global_pos = source.mapToGlobal(pos) if source is not None else self.mapToGlobal(pos)
        menu = QMenu(self)
        menu.setStyleSheet(
            "QMenu { background: white; border: 1px solid #d9d9d9; } "
            "QMenu::item:selected { background: #e6f7ff; color: #0050b3; } "
            "QMenu::separator { height: 1px; background: #d9d9d9; margin: 4px 8px; }"
        )

        act_analysis = menu.addAction("分析对象")
        act_overlay = menu.addAction("叠加对比")
        act_style = menu.addAction("图形样式")
        act_text = menu.addAction("文本设置")
        menu.addSeparator()
        act_view = menu.addAction("分布图")
        act_box = menu.addAction("摘要图")
        act_tornado = menu.addAction("敏感性图")
        act_scenario = menu.addAction("情景图")
        menu.addSeparator()
        act_export = menu.addAction("导出")

        act_analysis.triggered.connect(self.btn_analysis_obj.click)
        act_overlay.triggered.connect(self.btn_overlay.click)
        act_style.triggered.connect(self.btn_style.click)
        act_text.triggered.connect(self.btn_label_settings.click)
        act_view.triggered.connect(self.btn_view_mode.click)
        act_box.triggered.connect(self.btn_boxplot.click)
        act_tornado.triggered.connect(self.btn_tornado.click)
        act_scenario.triggered.connect(self.btn_scenario.click)
        act_export.triggered.connect(self.btn_export.click)
        menu.exec(global_pos)

    def _refresh_layout_bound_inputs(self):
        """
        当 Splitter 被拖动或窗口尺寸改变时，通知协调器或底层图表遮罩重新计算并对齐坐标边界。
        """
        coordinator = getattr(self, "_layout_refresh", None)
        if coordinator is not None:
            coordinator.notify()
            return
        try:
            if hasattr(self, "_chart_ctrl") and self._chart_ctrl:
                self._chart_ctrl.schedule_rect_sync()
        except Exception:
            pass
        self._refresh_layout_bound_inputs_deferred()
        QTimer.singleShot(120, self._refresh_layout_bound_inputs_deferred)

    def _refresh_layout_bound_inputs_deferred(self):
        """
        延迟刷新绑定布局的输入组件，避免阻塞 UI 主线程。
        """
        try:
            self._poll_and_sync_geometry()
        except Exception:
            pass
        try:
            self._update_x_edits_pos()
        except Exception:
            pass
    
    def _layout_refresh_frame(self):
        """
        布局刷新协调器的帧回调，执行高频同步操作。
        """
        try:
            self._force_plotly_resize()
        except Exception:
            pass
        try:
            if hasattr(self, "_chart_ctrl") and self._chart_ctrl:
                self._chart_ctrl.schedule_rect_sync()
        except Exception:
            pass
        try:
            self._poll_and_sync_geometry()
        except Exception:
            pass
    
    def _layout_refresh_final(self):
        """
        布局刷新协调器的结束回调，更新精确的悬浮框位置。
        """
        self._layout_refresh_frame()
        try:
            self._update_x_edits_pos()
        except Exception:
            pass
    
    def _on_content_splitter_moved(self, pos, index):
        """
        当用户拖动中间分割线时触发，重新计算布局并同步底层画板。
        """
        self._refresh_layout_bound_inputs()

    def resizeEvent(self, event):
        """
        窗口尺寸变化事件：触发定时器执行高频轮询，同步 Qt 与 HTML/Plotly 容器的长宽比例。
        """
        super().resizeEvent(event)
        
        # 1. 开启高频轮询 (30fps)，强制绑定 Plotly 的渲染网格宽度
        if not hasattr(self, "_geom_poll_timer"):
            self._geom_poll_timer = QTimer(self)
            self._geom_poll_timer.setInterval(33)
            self._geom_poll_timer.timeout.connect(self._poll_and_sync_geometry)
            
        if not hasattr(self, "_geom_poll_stop_timer"):
            self._geom_poll_stop_timer = QTimer(self)
            self._geom_poll_stop_timer.setSingleShot(True)
            self._geom_poll_stop_timer.timeout.connect(self._geom_poll_timer.stop)
            
        self._geom_poll_timer.start()
        # 拖拽停止后 500ms 自动关闭高频轮询，释放性能
        self._geom_poll_stop_timer.start(500)

        # 2. 防抖控制 Plotly 的底层重绘，避免 WebEngine IPC 通信拥塞卡死
        if not hasattr(self, "_plotly_resize_timer"):
            self._plotly_resize_timer = QTimer(self)
            self._plotly_resize_timer.setInterval(150)
            self._plotly_resize_timer.setSingleShot(True)
            self._plotly_resize_timer.timeout.connect(self._force_plotly_resize)
        self._plotly_resize_timer.start()

        self._update_x_edits_pos()
        try:
            if not getattr(self, "is_discrete_view", False):
                self._rebuild_drag_snap_grid()
        except Exception:
            pass
    
    def _poll_and_sync_geometry(self):
        """
        运行 JS 脚本高频抓取 Plotly 实际渲染网格的物理像素尺寸 (w, l)，
        并将参数同步给顶部的 Qt 悬浮输入框与滑块。
        """
        js = """
        (function(){
            try{
                var gd = document.getElementById('chart_div');
                if(!gd || !gd._fullLayout || !gd._fullLayout._size) return null;
                var sz = gd._fullLayout._size;
                return { w: sz.w, l: sz.l };
            }catch(e){ return null; }
        })();
        """
        def _cb(res):
            if not res or not isinstance(res, dict): return
            w = float(res.get("w", 0))
            if w > 0:
                # 1. 强制绑定滑块的真实绘制宽度
                if hasattr(self, "range_slider") and hasattr(self.range_slider, "setPlotWidth"):
                    self.range_slider.setPlotWidth(w)
                # 2. 刷新悬浮框位置
                self._update_x_edits_pos()
                # 3. 刷新遮罩
                if hasattr(self, "_chart_ctrl"):
                    self._chart_ctrl.schedule_rect_sync()

        if hasattr(self, "web_view") and self.web_view is not None:
            self.web_view.page().runJavaScript(js, _cb)

    def _force_plotly_resize(self):
        """
        防抖兜底机制：当窗口停止拖拽后，向 Web 发送终极对齐指令 Plotly.Plots.resize()。
        """
        if hasattr(self, "web_view") and self.web_view is not None:
            self.web_view.page().runJavaScript(
                "try { var gd = document.getElementById('chart_div'); "
                "if (gd && window.Plotly) { Plotly.Plots.resize(gd); } } catch(e) {}"
            )

    def _show_skeleton(self, text: str = "正在绘图..."):
        """
        在图表数据计算与 WebEngine 渲染期间屏蔽用户的操作，展示加载骨架屏，避免并发冲突。
        """
        if hasattr(self, "chart_skeleton") and self.chart_skeleton:
            try:
                self.chart_skeleton.lbl.setText(text)
                if hasattr(self, "chart_container") and self.chart_container:
                    self.chart_skeleton.setGeometry(self.chart_container.rect())
                self.chart_skeleton.raise_() 
                self.chart_skeleton.show()
            except Exception:
                pass
    
    def _maybe_hide_skeleton(self):
        """
        检查数据是否发送完毕以及图表是否加载完成，满足条件则自动隐藏加载骨架屏。
        """
        if not (self._webview_loaded and self._data_sent):
            return

        self._waiting_token = self._render_token
        self._ready_deadline = time.monotonic() + 10.0  
        if not self._ready_poll_timer.isActive():
            self._ready_poll_timer.start()

    def closeEvent(self, event):
        """
        窗口关闭事件：安全销毁定时器、回收共享的 Web 引擎页面，防止内存泄漏。
        """
        try:
            if hasattr(self, "_ready_poll_timer") and self._ready_poll_timer.isActive():
                self._ready_poll_timer.stop()
        except Exception:
            pass

        if hasattr(self, "overlay") and self.overlay is not None:
            self.overlay.hide()
            self.overlay.setParent(None)
            self.overlay.deleteLater()
            self.overlay = None
        if hasattr(self, "web_view") and self.web_view is not None:
            try:
                self.web_view.loadFinished.disconnect(self._on_webview_load_finished)
            except Exception:
                pass
            try:
                self.web_view.hide()
            except Exception:
                pass
            recycle_shared_webview(self.web_view)
            self.web_view = None
        super().closeEvent(event)
   
    # ---------------------------------------------------------
    # 3.3 数据加载与变量解析 (Data Loading & Variable Parsing)
    # ---------------------------------------------------------

    @staticmethod
    def _normalize_input_title_fallback_cell(raw_key) -> str:
        """
        获取输入变量的备用标题（剥离工作表和引擎的杂音，仅保留纯粹的单元格地址）。
        """
        return extract_clean_cell_address(raw_key)

    @staticmethod
    def _enrich_input_attrs_from_formula(base_key, attrs: dict) -> dict:
        """
        尽力从 Excel 工作表的公式文本中丰富输入变量的属性（作为元数据缓存不完整时的备用手段）。
        """
        merged = dict(attrs or {})
        key_text = str(base_key or "").strip().replace("$", "")
        if not key_text:
            return merged

        if "!" in key_text:
            sheet_name, raw_addr = key_text.split("!", 1)
        else:
            sheet_name, raw_addr = "", key_text

        lookup_addr = UltimateRiskDialog._normalize_input_title_fallback_cell(raw_addr)
        if not lookup_addr:
            return merged

        try:
            xl_app = _safe_excel_app()
            if xl_app is None:
                return merged
            sheet = xl_app.ActiveWorkbook.Worksheets(sheet_name) if sheet_name else xl_app.ActiveSheet
            formula_text = str(sheet.Range(lookup_addr).Formula)
        except Exception:
            return merged

        # 优先保留公式文本，以便后续理论分布能够重建
        if formula_text and not str(merged.get("formula", "") or "").strip():
            merged["formula"] = formula_text

        parsed_attrs = extract_all_attributes_from_formula(formula_text) or {}
        for attr_key in ("name", "units", "category"):
            val = str(parsed_attrs.get(attr_key, "") or "").strip()
            if val and not str(merged.get(attr_key, "") or "").strip():
                merged[attr_key] = val

        if str(merged.get("name", "") or "").strip() and "DRISKNAME" in str(formula_text).upper():
            merged["_name_from_attr"] = True
        return merged

    @staticmethod
    def _normalize_title_override_key(cell_key: str) -> str:
        """标准化用于覆盖标题的字典键"""
        return str(cell_key or "").replace("$", "").strip().upper()

    def _get_chart_title_display_override(self, cell_key: str) -> _Optional[str]:
        """获取图表标题的自定义覆盖文本"""
        key_norm = self._normalize_title_override_key(cell_key)
        if not key_norm:
            return None
        raw = (getattr(self, "_chart_title_display_overrides", {}) or {}).get(key_norm, None)
        if raw is None:
            return None
        text = str(raw).strip()
        return text if text else None

    def _set_chart_title_display_override(self, cell_key: str, title_text: _Optional[str]) -> None:
        """设置图表标题的自定义覆盖文本"""
        key_norm = self._normalize_title_override_key(cell_key)
        if not key_norm:
            return
        text = str(title_text or "").strip()
        if not text:
            if hasattr(self, "_chart_title_display_overrides"):
                self._chart_title_display_overrides.pop(key_norm, None)
            return
        if not hasattr(self, "_chart_title_display_overrides"):
            self._chart_title_display_overrides = {}
        self._chart_title_display_overrides[key_norm] = text

    def _resolve_chart_title_from_metadata(self, base_ck: str) -> str:
        """从元数据（名称、公式标签等）解析出图表的主标题"""
        attrs = self.attrs_map.get(self.current_key, {})
        xl_ctx = None
        try:
            xl_ctx = _safe_excel_app()
        except Exception:
            xl_ctx = None
        return resolve_visible_variable_name(
            base_ck,
            attrs,
            excel_app=xl_ctx,
            fallback_label=self.current_key,
        )

    def _resolve_display_chart_title(self, base_ck: str, metadata_title: str) -> str:
        """解析最终显示的图表标题（判断是否有自定义覆盖项）"""
        override = self._get_chart_title_display_override(base_ck)
        if override:
            return override
        return str(metadata_title or "").strip()

    def load_dataset(self, key):
        """
        核心数据加载方法。
        根据单元格 key，从后端抽取自身数据及叠加数据，合并为 DataFrame，
        重建全局颜色样式与坐标轴极值，并触发重绘。
        """
        self.current_key = key
        self.manual_mag = None
        self.theory_dist_obj = None
        self._theory_func_name = ""
        self._theory_params = []
        self._theory_markers = {}
        self._theory_is_discrete = None
        base_ck = self.label_to_cell_key.get(self.current_key, self.current_key)

        # 针对输入变量重建理论分布
        if getattr(self, "_base_kind", "output") == "input":
            attrs = self.attrs_map.get(self.current_key, {})
            if not attrs and hasattr(self, 'sim_id') and self.sim_id is not None:
                try:
                    pure_ck = self.label_to_cell_key.get(self.current_key, self.current_key)
                    attrs = bridge.get_attributes(self.sim_id, pure_ck, kind="input") or {}
                    self.attrs_map[self.current_key] = attrs
                except Exception as e:
                    print(f"获取属性失败，目标键 {self.current_key}: {e}")
            attrs = self._enrich_input_attrs_from_formula(base_ck, attrs)
            self.attrs_map[self.current_key] = attrs
            formula = attrs.get('formula', '')
            if not formula:
                expr = str(attrs.get('expression', '') or '').strip()
                if expr:
                    formula = expr if expr.startswith("=") else f"={expr}"
            if formula:
                self.theory_dist_obj = self._create_theory_dist_from_formula(formula, source_cell_key=base_ck)
            if self.theory_dist_obj is None:
                pure_ck = self.label_to_cell_key.get(self.current_key, self.current_key)
                self.theory_dist_obj = self._create_theory_dist_from_sim_metadata(pure_ck)

        attrs = self.attrs_map.get(self.current_key, {})
        units_val = attrs.get('units', '')
        metadata_title = self._resolve_chart_title_from_metadata(base_ck)
        full_title = self._resolve_display_chart_title(base_ck, metadata_title)

        self.chart_title_label.setText(full_title)
        self._base_chart_title = full_title
        DriskChartFactory.VALUE_AXIS_TITLE = ""
        DriskChartFactory.VALUE_AXIS_UNIT = units_val.strip() if units_val else ""

        self.display_keys = []
        self.series_map = {}
        cleaned_list = []
        seen_names = set()

        # 添加主数据数列
        base_name = f"{self.current_key} (sim{self.sim_id})"
        self.display_keys.append(base_name)
        seen_names.add(base_name)

        base_ck = self.label_to_cell_key.get(self.current_key, self.current_key)
        try:
            arr = bridge.get_series(self.sim_id, base_ck, kind=getattr(self, "_base_kind", "output"))
            s_base = ResultsDataService.clean_series(arr)
        except Exception:
            s_base = pd.Series([0.0] * 10)

        self.data = s_base
        self.series_map[base_name] = s_base
        cleaned_list.append(s_base)

        # 添加叠加层数据数列
        if hasattr(self, "overlay_items"):
            for item in self.overlay_items:
                if len(item) == 4:
                    sid, ck, var_lbl, knd = item
                else:
                    sid, ck, var_lbl = item
                    knd = getattr(self, "_base_kind", "output")

                ck_pure = ck.split('!')[-1] if '!' in ck else ck
                base_ck_pure = base_ck.split('!')[-1] if '!' in base_ck else base_ck
                ck_sheet = ck.split('!')[0] if '!' in ck else ""
                base_ck_sheet = base_ck.split('!')[0] if '!' in base_ck else ""
                same_cell = ck_pure == base_ck_pure and ck_sheet == base_ck_sheet
                use_lbl = self.current_key if same_cell else var_lbl
                overlay_name = f"{use_lbl} (sim{sid})"
                # 不同表格同一单元格地址时追加表格标识以避免键冲突
                if ck_pure == base_ck_pure and not same_cell:
                    sheet_tag = ck_sheet if ck_sheet else ck
                    overlay_name = f"{use_lbl}[{sheet_tag}] (sim{sid})"

                if overlay_name not in seen_names:
                    self.display_keys.append(overlay_name)
                    seen_names.add(overlay_name)
                    try:
                        arr = bridge.get_series(sid, ck, kind=knd)
                        s_over = ResultsDataService.clean_series(arr)
                    except Exception:
                        s_over = pd.Series([0.0] * 10)
                    self.series_map[overlay_name] = s_over
                    cleaned_list.append(s_over)

        self._build_series_style(self.display_keys)
        self._invalidate_stats_cache()
        self._cdf_sorted_cache.clear()
        if hasattr(self, "_cdf_ci_cache"):
            self._cdf_ci_cache.clear()
        if hasattr(self, "_cdf_table_cache"):
            self._cdf_table_cache.clear()

        # 生成全局并集数据以推断极值边界
        self.data_all = pd.concat(cleaned_list, ignore_index=True)
        self.n_all = len(self.data_all)

        if self.n_all == 0:
            self.data_all = pd.Series([0.0, 1.0])
            self.n_all = 2

        if len(self.data) == 0:
            self.data = pd.Series([0.0, 1.0])
        self.n = len(self.data)

        axis_params = ResultsDataService.calc_x_axis_params(self.data_all, detect_series=self.data)
        self.__dict__.update(axis_params)

        _dt = self.x_dtick if self.x_dtick > 0 else 1.0
        try:
            _mag = math.floor(math.log10(_dt))
        except Exception:
            _mag = 0
        self.manual_mag = get_axis_display_unit_override(getattr(self, "_label_settings_config", None), "x")
        self._label_axis_numeric = get_axis_numeric_flags(
            getattr(self, "_label_settings_config", None),
            fallback=self._resolve_label_settings_axis_numeric_defaults(),
        )
        self._current_view_data = self._normalize_view_command_for_dataset(
            getattr(self, "_current_view_data", "auto")
        )
        current_cmd = getattr(self, "_current_view_data", "auto")
        self._apply_view_command_state(current_cmd)

        self.current_left = float(np.percentile(self.data_all, 5))
        self.current_right = float(np.percentile(self.data_all, 95))

        if self.is_discrete_view:
            self.current_left = self._snap_discrete_nearest(self.current_left)
            self.current_right = self._snap_discrete_nearest(self.current_right)

        self._clamp_and_sync_slider()
        self._sync_discrete_points_from_data_all()
        self._rebuild_drag_snap_grid()

        self.init_chart_via_tempfile()
        self.update_stats_ui()

        if getattr(self, "chart_mode", "hist") == "cdf":
            self.rebuild_cdf_prob_boxes()

        def _delayed_ui_sync():
            self._update_x_edits_pos()
            self._update_vlines_only()

        QTimer.singleShot(0, _delayed_ui_sync)

    def _get_series(self, key: str) -> pd.Series:
        """
        从内存缓存中安全提取对应变量名称的 Pandas Series 数据，若缺失则返回占位数据防止崩溃。
        """
        if not hasattr(self, "series_map") or self.series_map is None:
            self.series_map = {}

        if key in self.series_map:
            return self.series_map[key]

        raw = self.all_datasets.get(key, [])
        s = ResultsDataService.clean_series(raw)
        if len(s) == 0:
            s = pd.Series([0.0] * 10)

        self.series_map[key] = s
        return s

    def _build_series_style(self, keys):
        """
        根据当前显示的变量集合，智能分配调色板颜色（区分输入变量/输出变量的配色方案）。
        """
        cycle = DRISK_COLOR_CYCLE_THEO if getattr(self, "_base_kind", "output") == "input" else DRISK_COLOR_CYCLE
        n = len(cycle)
        self._series_style = {}
        for i, k in enumerate(keys):
            self._series_style[k] = {
                "color": cycle[i % n],
                "cdf_color": cycle[i % n],  # 让 CDF 曲线颜色与直方图主体调色板保持一致
                "dash": "solid" if i < n else "dash"
            }

    def _sync_discrete_points_from_data_all(self):
        """
        提取全部数据中出现过的去重离散点阵列，存入缓存供后续的网格吸附逻辑使用。
        """
        if self.is_discrete_view:
            try:
                arr = np.asarray(self.data_all.values, dtype=float)
                arr = arr[np.isfinite(arr)]
                pts = np.unique(arr)
                pts.sort()
                self._discrete_points = pts
            except Exception:
                self._discrete_points = None
        else:
            self._discrete_points = None

    # ---------------------------------------------------------
    # 3.4 理论分布引擎适配 (Theory Distribution Engine)
    # ---------------------------------------------------------

    def _normalize_theory_cell_key(self, cell_key):
        """
        抹平相对引用与绝对引用 ($) 的差异，便于与后台匹配理论分布的元数据。
        """
        txt = str(cell_key or "").strip()
        if not txt:
            return ""
        if bridge is not None:
            try:
                return bridge.normalize_cell_key(txt)
            except Exception:
                pass
        return txt.replace("$", "").upper()

    def _cache_theory_signature(self, func_name, params, markers):
        """
        缓存当前解析出的理论分布特征（包括函数名、参数、截断/平移等），避免重复执行高昂的解析操作。
        """
        self._theory_func_name = str(func_name or "")
        self._theory_params = list(params) if params is not None else []
        self._theory_markers = dict(markers or {})
        self._theory_is_discrete = None
        if bridge is not None and self._theory_func_name:
            try:
                spec = bridge.get_distribution_spec_by_func_name(self._theory_func_name)
                if spec is not None:
                    self._theory_is_discrete = bool(spec.is_discrete)
            except Exception:
                pass

    @staticmethod
    def _normalize_theory_truncate_markers(markers):
        """标准化理论分布的截断标记位（统一转换为标准的参数键名）"""
        normalized = dict(markers or {})

        # 将下划线别名统一转换为标准截断键
        for src_key, dst_key in (("truncate_p", "truncatep"), ("truncate_p2", "truncatep2")):
            src_val = normalized.get(src_key, None)
            dst_val = normalized.get(dst_key, None)
            src_is_value = (src_val is not None) and (not isinstance(src_val, bool))
            dst_is_placeholder = isinstance(dst_val, bool) or dst_val is None
            if src_is_value and dst_is_placeholder:
                normalized[dst_key] = src_val
            normalized.pop(src_key, None)

        # 丢弃单纯的布尔值占位符，防止参数污染
        for key in ("truncate", "truncate2", "truncatep", "truncatep2"):
            if isinstance(normalized.get(key, None), bool):
                normalized.pop(key, None)

        return normalized

    @staticmethod
    def _coerce_param_float_value(raw_val) -> _Optional[float]:
        """尝试将各种类型的原始参数强制转换为浮点数以供计算使用"""
        if isinstance(raw_val, (list, tuple, np.ndarray)):
            if len(raw_val) == 0:
                return None
            return UltimateRiskDialog._coerce_param_float_value(raw_val[0])

        if raw_val is None:
            return None

        if isinstance(raw_val, str):
            txt = raw_val.strip()
            if not txt or txt.startswith("#"):
                return None
            txt = txt.strip("{}").strip()
            if not txt:
                return None
            try:
                return float(txt)
            except Exception:
                return None

        try:
            return float(raw_val)
        except Exception:
            return None

    @staticmethod
    def _split_ref_token_for_theory(token: str, default_sheet: str = "") -> tuple[str, str]:
        """切分引用字符串（Token），提取用于理论分布解析的工作表名与单元格地址"""
        txt = str(token or "").strip().lstrip("@")
        if txt.startswith("="):
            txt = txt[1:].strip()
        txt = txt.replace("$", "")
        if not txt:
            return "", ""

        if "!" in txt:
            sheet_name, addr = txt.rsplit("!", 1)
            sheet_name = sheet_name.strip()
            if len(sheet_name) >= 2 and sheet_name[0] == sheet_name[-1] == "'":
                sheet_name = sheet_name[1:-1].replace("''", "'").strip()
        else:
            sheet_name, addr = default_sheet, txt

        addr = str(addr or "").strip().upper()
        if not re.match(r"^[A-Za-z]{1,3}\d{1,7}$", addr):
            return "", ""
        return str(sheet_name or "").strip(), addr

    def _resolve_simtable_reference_value(self, full_ref_key: str, sim_obj) -> _Optional[float]:
        """尝试解析 simtable 模拟表引用的数值"""
        if sim_obj is None:
            return None

        simtable_cells = getattr(sim_obj, "simtable_cells", {}) or {}
        if not isinstance(simtable_cells, dict) or not simtable_cells:
            return None

        target_key = self._normalize_theory_cell_key(full_ref_key)
        if not target_key:
            return None

        funcs = simtable_cells.get(target_key)
        if funcs is None:
            for raw_key, raw_funcs in simtable_cells.items():
                if self._normalize_theory_cell_key(raw_key) == target_key:
                    funcs = raw_funcs
                    break

        if not isinstance(funcs, list) or not funcs:
            return None

        first = funcs[0] if isinstance(funcs[0], dict) else {}
        values = first.get("values", []) if isinstance(first, dict) else []
        if not isinstance(values, (list, tuple, np.ndarray)) or len(values) == 0:
            return None

        try:
            scenario_idx = int(getattr(sim_obj, "scenario_index", 0) or 0)
        except Exception:
            scenario_idx = 0

        if scenario_idx < 0 or scenario_idx >= len(values):
            scenario_idx = 0
        if scenario_idx < 0 or scenario_idx >= len(values):
            return None

        return self._coerce_param_float_value(values[scenario_idx])

    def _resolve_numeric_param_token(self, token: str, *, default_sheet: str = "", sim_obj=None) -> _Optional[float]:
        """解析数值类型的参数字符串（提取硬编码值或从 Excel 底层实时求值）"""
        text = str(token or "").strip()
        if not text:
            return None

        text = text.lstrip("@")
        if text.startswith("="):
            text = text[1:].strip()

        direct = self._coerce_param_float_value(text)
        if direct is not None:
            return direct

        sheet_name, addr = self._split_ref_token_for_theory(text, default_sheet=default_sheet)
        if addr:
            full_ref_key = f"{sheet_name}!{addr}" if sheet_name else addr
            simtable_val = self._resolve_simtable_reference_value(full_ref_key, sim_obj)
            if simtable_val is not None:
                return simtable_val

            try:
                xl_app = _safe_excel_app()
                if xl_app is not None:
                    sheet_obj = xl_app.ActiveWorkbook.Worksheets(sheet_name) if sheet_name else xl_app.ActiveSheet
                    cell_obj = sheet_obj.Range(addr)
                    raw_val = cell_obj.Value2
                    if raw_val is None:
                        raw_val = cell_obj.Value
                    resolved = self._coerce_param_float_value(raw_val)
                    if resolved is not None:
                        return resolved
            except Exception:
                pass

        try:
            xl_app = _safe_excel_app()
            if xl_app is None:
                return None
            if default_sheet:
                sheet_obj = xl_app.ActiveWorkbook.Worksheets(default_sheet)
                raw_eval = sheet_obj.Evaluate(text)
            else:
                raw_eval = xl_app.Evaluate(text)
            return self._coerce_param_float_value(raw_eval)
        except Exception:
            return None

    def _resolve_effective_theory_params(
        self,
        raw_tokens,
        *,
        expected_count: int = 0,
        default_sheet: str = "",
        sim_obj=None,
    ) -> list[float]:
        """批量解析理论分布的参数，并统一转换为标准浮点数数组供后台算法调用"""
        tokens = list(raw_tokens or [])
        if expected_count > 0 and len(tokens) > expected_count:
            tokens = tokens[:expected_count]

        resolved: list[float] = []
        for token in tokens:
            val = self._resolve_numeric_param_token(
                token,
                default_sheet=default_sheet,
                sim_obj=sim_obj,
            )
            if val is None:
                return []
            resolved.append(float(val))

        if expected_count > 0 and len(resolved) < expected_count:
            return []
        return resolved

    def _create_theory_dist_from_formula(self, formula, source_cell_key=""):
        """
        根据提取到的 Excel 公式文本构建底层理论分布对象。
        通过正则与桥接接口，将公式字符串转换为后端对应的物理运算对象。
        """
        if not theory or not bridge: return None
        try:
            formula_text = str(formula or "").strip()
            if not formula_text:
                return None
            if not formula_text.startswith("="):
                formula_text = "=" + formula_text
            res = bridge.parse_first_distribution_in_formula(formula_text)
            if not res: return None
            func_name = str(res.get("func_name", "") or "").strip()
            if not func_name:
                return None

            # 从列表驱动分布（例如 Cumul 数组）中提取无标记原始参数。
            args_list = list(res.get("args_list", []) or [])
            raw_main_args = []
            for arg in args_list:
                arg_text = str(arg or "").strip()
                if not arg_text:
                    continue
                if arg_text.lower().startswith("drisk"):
                    continue
                raw_main_args.append(arg_text)

            params = list(res.get("dist_params", []) or [])
            spec = bridge.get_distribution_spec_by_func_name(func_name)
            expected_param_count = len(getattr(spec, "param_names", ()) or ())
            fn_core = func_name.lstrip("@").upper().replace("DRISK", "").strip()
            list_payload_families = {"CUMUL", "DISCRETE", "DUNIFORM"}
            source_key = str(source_cell_key or "").replace("$", "").strip()
            default_sheet = source_key.split("!", 1)[0].strip() if "!" in source_key else ""
            sim_obj = None
            if getattr(self, "sim_id", None) is not None:
                try:
                    sim_obj = bridge.get_simulation(int(self.sim_id))
                except Exception:
                    sim_obj = None

            if fn_core in list_payload_families and raw_main_args:
                params = list(raw_main_args)
            else:
                if expected_param_count > 0 and len(params) > expected_param_count:
                    params = params[:expected_param_count]
                needs_runtime_resolution = (
                    fn_core not in list_payload_families
                    and expected_param_count > 0
                    and (
                        len(params) < expected_param_count
                        or any(
                            self._coerce_param_float_value(p) is None
                            for p in list(params)[:expected_param_count]
                        )
                    )
                )
                if needs_runtime_resolution:
                    resolved_params = self._resolve_effective_theory_params(
                        raw_main_args or params,
                        expected_count=expected_param_count,
                        default_sheet=default_sheet,
                        sim_obj=sim_obj,
                    )
                    if resolved_params:
                        params = resolved_params

            # 如果非列表系列的数值提取不完整，则回退依靠元数据解析。
            if (
                fn_core not in list_payload_families
                and expected_param_count > 0
                and (
                    len(params) < expected_param_count
                    or any(
                        self._coerce_param_float_value(p) is None
                        for p in list(params)[:expected_param_count]
                    )
                )
            ):
                return None
            if fn_core not in list_payload_families and expected_param_count > 0:
                params = [float(self._coerce_param_float_value(p)) for p in list(params)[:expected_param_count]]
            markers = {}
            
            # 1. 尝试使用公式解析器提取
            try:
                attrs = extract_all_attributes_from_formula(formula_text)
                if isinstance(attrs, dict):
                    markers.update(attrs)
            except Exception:
                pass
                
            # 2. 强力正则解析平移与截断，保证与 ui_modeler 行为 100% 对齐
            m_shift = re.search(r'DriskShift\s*\(\s*([-+]?(?:\d*\.\d+|\d+))\s*\)', formula_text, re.IGNORECASE)
            if m_shift: markers['shift'] = float(m_shift.group(1))
            
            def parse_trunc_args(fname):
                m = re.search(rf'{fname}\s*\(\s*([^,]*)\s*,\s*([^)]*)\s*\)', formula_text, re.IGNORECASE)
                if m:
                    def _p(s):
                        s = s.strip()
                        return float(s) if s else None
                    try:
                        return (_p(m.group(1)), _p(m.group(2)))
                    except Exception:
                        pass
                return None
                
            t_res = parse_trunc_args('DriskTruncate')
            if t_res: markers['truncate'] = t_res
            t2_res = parse_trunc_args('DriskTruncate2')
            if t2_res: markers['truncate2'] = t2_res
            tp_res = parse_trunc_args('DriskTruncateP')
            if tp_res: markers['truncatep'] = tp_res
            tp2_res = parse_trunc_args('DriskTruncateP2')
            if tp2_res: markers['truncatep2'] = tp2_res
            markers = self._normalize_theory_truncate_markers(markers)

            # ✅ 核心：将这些属性保存到类变量，供后面的 ModelerStatsJob 完美复用
            self._cache_theory_signature(func_name, params, markers)

            # 🔴 手术式修复 2：抛弃旧的 theory 工厂，直接调用 backend_bridge 的精确接口
            dist_obj = bridge.get_backend_distribution(
                func_name, params, markers
            )
            if dist_obj is None:
                return None
            if bool(getattr(dist_obj, "_invalid", False)):
                return None
            return dist_obj
        except Exception as e:
            print(f"创建理论分布发生错误: {e}")
            return None

    def _create_theory_dist_from_sim_metadata(self, cell_key):
        """
        作为兜底方案：从当前模拟的元数据缓存中，查找并构建对应单元格的理论分布对象。
        """
        if not bridge or not hasattr(self, "sim_id") or self.sim_id is None:
            return None

        target_key = self._normalize_theory_cell_key(cell_key)
        if not target_key:
            return None

        try:
            sim = bridge.get_simulation(int(self.sim_id))
        except Exception:
            return None

        distribution_cells = getattr(sim, "distribution_cells", {}) or {}
        matched = None

        for cell_addr, dist_funcs in distribution_cells.items():
            if not isinstance(dist_funcs, list):
                continue
            for dist_func in dist_funcs:
                if not isinstance(dist_func, dict):
                    continue
                candidate_keys = []
                for raw_key in (dist_func.get("key"), dist_func.get("input_key")):
                    if raw_key:
                        candidate_keys.append(raw_key)
                idx = dist_func.get("index", None)
                if idx is not None:
                    candidate_keys.append(f"{cell_addr}_{idx}")

                for cand in candidate_keys:
                    if self._normalize_theory_cell_key(cand) == target_key:
                        matched = dist_func
                        break
                if matched is not None:
                    break
            if matched is not None:
                break

        if matched is None:
            raw = str(cell_key or "")
            if "!" in raw:
                sheet, addr = raw.split("!", 1)
                base_addr = addr.split("_", 1)[0]
                base_key = f"{sheet}!{base_addr}"
            else:
                base_key = raw.split("_", 1)[0]
            base_key = self._normalize_theory_cell_key(base_key)
            if base_key:
                for cell_addr, dist_funcs in distribution_cells.items():
                    if self._normalize_theory_cell_key(cell_addr) != base_key:
                        continue
                    if isinstance(dist_funcs, list) and dist_funcs:
                        if len(dist_funcs) == 1:
                            matched = dist_funcs[0]
                        else:
                            first = [d for d in dist_funcs if isinstance(d, dict) and int(d.get("index", 0) or 0) == 1]
                            matched = first[0] if first else None
                    break

        if not isinstance(matched, dict):
            return None

        func_name = str(matched.get("func_name") or matched.get("function_name") or "").strip()
        params = matched.get("parameters")
        if params is None:
            params = matched.get("dist_params", [])
        markers = matched.get("markers")
        if not isinstance(markers, dict):
            markers = {}

        full_match = str(matched.get("full_match") or "").strip()
        need_formula_reparse = (not func_name or not params)
        if func_name and not need_formula_reparse and bridge is not None:
            try:
                spec = bridge.get_distribution_spec_by_func_name(func_name)
            except Exception:
                spec = None
            expected_param_count = len(getattr(spec, "param_names", ()) or ()) if spec is not None else 0
            if expected_param_count > 0 and len(list(params or [])) < expected_param_count:
                # 元数据可能携带了针对列表驱动分布截断的参数数组，需要完全重解析以对齐理论图层与原公式。
                need_formula_reparse = True

        if need_formula_reparse and full_match:
            if full_match:
                if not full_match.startswith("="):
                    full_match = "=" + full_match
                try:
                    parsed = bridge.parse_first_distribution_in_formula(full_match) or {}
                except Exception:
                    parsed = {}
                if not func_name:
                    func_name = str(parsed.get("func_name") or "").strip()
                if not params or need_formula_reparse:
                    parsed_args = list(parsed.get("args_list", []) or [])
                    raw_main_args = []
                    for arg in parsed_args:
                        arg_text = str(arg or "").strip()
                        if not arg_text or arg_text.lower().startswith("drisk"):
                            continue
                        raw_main_args.append(arg_text)
                    params = raw_main_args or parsed.get("dist_params", [])
                if not markers:
                    parsed_markers = parsed.get("markers", {})
                    if isinstance(parsed_markers, dict):
                        markers = dict(parsed_markers)

        if not func_name:
            return None

        default_sheet = ""
        raw_cell_key = str(cell_key or "").replace("$", "").strip()
        if "!" in raw_cell_key:
            default_sheet = raw_cell_key.split("!", 1)[0].strip()

        try:
            spec = bridge.get_distribution_spec_by_func_name(func_name)
        except Exception:
            spec = None
        expected_param_count = len(getattr(spec, "param_names", ()) or ()) if spec is not None else 0
        fn_core = func_name.lstrip("@").upper().replace("DRISK", "").strip()
        list_payload_families = {"CUMUL", "DISCRETE", "DUNIFORM"}

        param_tokens = params if isinstance(params, (list, tuple, np.ndarray)) else [params]
        if fn_core not in list_payload_families and expected_param_count > 0:
            resolved_params = self._resolve_effective_theory_params(
                param_tokens,
                expected_count=expected_param_count,
                default_sheet=default_sheet,
                sim_obj=sim,
            )
            if not resolved_params:
                return None
            params = resolved_params
        elif expected_param_count > 0 and len(list(param_tokens)) > expected_param_count:
            params = list(param_tokens)[:expected_param_count]

        try:
            self._cache_theory_signature(func_name, params, markers)
            return bridge.get_backend_distribution(func_name, params, markers)
        except Exception:
            return None

    def _theory_label(self) -> str:
        """
        返回统一的理论分布标签文本（"理论分布"）。
        """
        return "\u7406\u8bba\u5206\u5e03"  # 返回 "理论分布" 的 Unicode 编码，避免源文件编码引起的乱码

    def _is_single_series_view(self) -> bool:
        """
        判断当前视图是否为单变量展示，多变量叠加对比时不应展示理论分布层。
        """
        keys = list(getattr(self, "display_keys", [self.current_key]))
        return len(keys) == 1

    def _theory_overlay_trigger_active(self) -> bool:
        """
        判断是否满足触发理论分布图层覆盖的条件：当前必须是输入变量，且理论公式有效，且为单变量视图。
        """
        return (
            getattr(self, "_base_kind", "output") == "input"
            and getattr(self, "theory_dist_obj", None) is not None
            and self._is_single_series_view()
        )

    def _is_theory_source_discrete(self) -> bool:
        """
        判断当前关联的理论分布是否为离散分布（如 Binomial），以便在绘图时选择适合离散值的策略。
        """
        cached = getattr(self, "_theory_is_discrete", None)
        if isinstance(cached, bool):
            return cached

        func_name = getattr(self, "_theory_func_name", "")
        if bridge is not None and func_name:
            try:
                spec = bridge.get_distribution_spec_by_func_name(func_name)
                if spec is not None:
                    self._theory_is_discrete = bool(spec.is_discrete)
                    return bool(spec.is_discrete)
            except Exception:
                pass

        dist = getattr(self, "theory_dist_obj", None)
        if dist is not None:
            has_pmf = hasattr(dist, "pmf_vec") or hasattr(dist, "pmf")
            has_pdf = hasattr(dist, "pdf_vec") or hasattr(dist, "pdf")
            if has_pmf and not has_pdf:
                self._theory_is_discrete = True
                return True

        fallback = bool(getattr(self, "is_discrete_view", False))
        self._theory_is_discrete = fallback
        return fallback

    def _theory_visibility_flags(self):
        """
        判断在当前的图表模式下，理论分布（PDF / CDF）是否应该显示叠加层。
        """
        if not self._theory_overlay_trigger_active():
            return False, False
        is_discrete_source = self._is_theory_source_discrete()
        mode = getattr(self, "chart_mode", "hist")
        raw_cmd = str(getattr(self, "_current_view_data", "") or "").strip().lower()
        base_cmd = raw_cmd.split("_")[0] if raw_cmd else ""
        is_relfreq_mode = base_cmd == "relfreq"
        is_hist_like_mode = mode in ("hist", "kde") and (not is_relfreq_mode)
        is_discrete_overlay_mode = is_discrete_source and mode == "discrete"
        show_pdf = ((not is_discrete_source) and is_hist_like_mode) or is_discrete_overlay_mode
        show_cdf = mode == "cdf"
        return show_pdf, show_cdf

    def _should_show_theory_stats(self) -> bool:
        """
        判断右侧的统计表中是否应该追加一列展示“理论分布”的真值计算结果。
        """
        show_pdf, show_cdf = self._theory_visibility_flags()
        return bool(show_pdf or show_cdf)

    def _safe_theory_cdf_scalar(self, x: float) -> float:
        """
        安全计算理论分布在特定值 x 处的累积概率 (CDF)。
        """
        dist = getattr(self, "theory_dist_obj", None)
        if dist is None:
            return 0.0
        try:
            if hasattr(dist, "cdf_vec"):
                res = np.asarray(dist.cdf_vec(np.array([float(x)])), dtype=float).ravel()
                if res.size > 0 and np.isfinite(res[0]):
                    return float(res[0])
            if hasattr(dist, "cdf"):
                val = float(dist.cdf(float(x)))
                if np.isfinite(val):
                    return val
        except Exception:
            pass
        return 0.0

    def _safe_theory_ppf_scalar(self, p: float) -> float:
        """
        安全计算理论分布在特定概率 p 下的逆累积概率 (分位点 PPF)。
        包含降级的二分查找算法防备底层 PPF 方法不可用。
        """
        dist = getattr(self, "theory_dist_obj", None)
        if dist is None:
            return float("nan")

        p = float(max(0.0, min(1.0, p)))
        x_min = float(getattr(self, "x_range_min", 0.0))
        x_max = float(getattr(self, "x_range_max", 1.0))

        if p <= 0.0:
            return x_min
        if p >= 1.0:
            return x_max

        try:
            if hasattr(dist, "inv_cdf"):
                val = float(dist.inv_cdf(p))
                if np.isfinite(val):
                    return val
            if hasattr(dist, "ppf"):
                val = float(dist.ppf(p))
                if np.isfinite(val):
                    return val
        except Exception:
            pass

        try:
            lo = float(x_min)
            hi = float(x_max)
            for _ in range(64):
                mid = 0.5 * (lo + hi)
                cdf_mid = self._safe_theory_cdf_scalar(mid)
                if cdf_mid < p:
                    lo = mid
                else:
                    hi = mid
            return 0.5 * (lo + hi)
        except Exception:
            return float("nan")

    # ---------------------------------------------------------
    # 3.5 视图指令与图表模式切换 (View Commands & Chart Modes)
    # ---------------------------------------------------------

    def _parse_view_command_flags(self, cmd):
        """
        解析复合视图指令（如 'histogram_kde'），提取具体的展示状态开关。
        """
        state = ResultsViewCommandHelper.parse_command(cmd)
        return state.base_view, state.show_cdf, state.show_kde, state.show_dkw

    def _apply_view_command_state(self, cmd):
        """
        应用视图指令状态：将解析得到的命令配置到内部环境中，并同步通知至专用辅助模块。
        """
        return ResultsViewCommandHelper.apply_command_state(self, cmd)

    def _normalize_view_command_for_dataset(self, cmd):
        """
        根据当前数据集的离散度特征、单多序列情况，自动降级并纠正不兼容的视图命令以防报错。
        """
        return ResultsViewCommandHelper.normalize_for_dataset(
            cmd,
            first_open=bool(getattr(self, "_first_open", False)),
            is_discrete_view=bool(getattr(self, "is_discrete_view", False)),
            display_count=len(getattr(self, "display_keys", [])),
        )

    def _on_view_cmd_triggered(self, label, cmd):
        """
        视图菜单点击事件的路由接收器，记录指令并移交图表重绘流程。
        """
        self._current_view_data = cmd 
        self.on_chart_type_changed(0)
    
    def on_chart_type_changed(self, index=0):
        """
        图表类型切换的主流程入口：
        重置分析模式，重绘图形，并立刻弃用旧的统计指标缓存以保证数据准确。
        """
        cmd = getattr(self, "_current_view_data", "auto")
        
        # 解析复合指令后缀，更新内部状态
        self._apply_view_command_state(cmd)

        # 👇 切换图表视图时，务必关闭所有高级分析模式 (修复诸如箱线图无法切回直方图的 Bug)
        self._restore_main_view_from_advanced_if_needed()
            
        self.apply_chart_mode_ui()
        self.init_chart_via_tempfile()
        
        # ✅ 核心：视图切换会导致指标表结构大变，必须废弃旧缓存，重新触发后台计算
        self._invalidate_stats_cache() 
        
        self.update_stats_ui()
        self._update_bottom_bar_visibility()

    def apply_chart_mode_ui(self):
        """
        适配当前图表模式的 UI 配置，例如恢复并显示底部的范围滑块两端把手。
        """
        try:
            self.range_slider.setSingleHandleMode(False)
        except Exception:
            pass

        # 始终确保顶部左右两个输入框可见
        self.edt_x_l.show()
        self.edt_x_r.show()

        self._update_vlines_only()
        self._update_x_edits_pos()

    def _set_chart_mode_from_view(self, view_mode: str):
        """
        根据底层视图指令设置相应的内部图表模式标识。
        """
        ResultsRenderDispatcher.set_chart_mode_from_view(self, view_mode)

    def _open_style_manager(self):
        """
        唤起“图形样式”配置管理器，支持调节颜色的填充、不透明度与线条等属性。
        点击确认后，自动触发重绘机制应用新样式。
        """
        mode = ResultsRuntimeStateHelper.get_current_analysis_mode(self, "")
        
        show_bar = False
        show_curve = False
        show_mean = False 
        is_unified = False 
        style_profile = None
        
        # 🔴 动态绑定目标键值和目标样式字典
        active_keys = getattr(self, "display_keys", [])
        active_styles = getattr(self, "_series_style", {})
        series_style_proxy_map: dict[str, str] | None = None

        if mode in ["boxplot", "letter_value", "violin"]:
            show_bar = True
            show_curve = True
            show_mean = (mode != "letter_value") 
            is_unified = True 
        elif mode == "trend":
            show_bar = False
            show_curve = True
            show_mean = True 
            is_unified = True 
        elif mode == "tornado":
            t_mode = ResultsRuntimeStateHelper.get_tornado_mode(self, "bins")
            if t_mode == "bins_line":
                show_bar = False
                show_curve = True
                # 从视图组件中提取当前正在展示的线型变量名称
                active_keys = getattr(self.tornado_view, "current_line_vars", [])
                if not active_keys: return
                # 初始化缺失的线型配置
                for i, k in enumerate(active_keys):
                    if k not in self._tornado_line_styles:
                        self._tornado_line_styles[k] = {"curve_color": DRISK_COLOR_CYCLE[i % len(DRISK_COLOR_CYCLE)], "curve_width": 2.0, "dash": "solid", "curve_opacity": 1.0}
                active_styles = self._tornado_line_styles
            elif t_mode == "bins":
                show_bar = True
                active_keys = ["高值组", "低值组"]
                active_styles = self._tornado_styles
            else:
                show_bar = True
                active_keys = ["正值", "负值"]
                active_styles = self._tornado_styles
        elif mode == "scenario":
            show_bar = True
            active_keys = ["正值", "负值"]
            active_styles = self._tornado_styles
        else:
            cmd = getattr(self, "_current_view_data", "auto")
            if cmd == "auto": cmd = "discrete" if getattr(self, "is_discrete_view", False) else "histogram"
            if cmd in ["histogram", "relfreq", "discrete"]: show_bar = True
            elif cmd in ["pdfcurve", "cdf", "cdf_dkw"]: show_curve = True
            else: show_bar = True; show_curve = True
            base_cmd = str(cmd or "").strip().lower().split("_")[0]
            if show_bar:
                if base_cmd == "discrete":
                    style_profile = {
                        "fill_section_title": "柱体",
                        "fill_color_label": "柱体颜色:",
                        "fill_opacity_label": "柱体不透明度:",
                        "show_outline_controls": False,
                        "defaults": {
                            "fill_opacity": 1.0,
                            "outline_color": "rgba(0,0,0,0)",
                            "outline_width": 0.0,
                        },
                    }
                else:
                    style_profile = {
                        "fill_section_title": "填充",
                        "fill_color_label": "填充颜色:",
                        "fill_opacity_label": "填充不透明度:",
                        "show_outline_controls": True,
                        "defaults": {
                            "fill_opacity": 0.75,
                            "outline_color": "rgba(255,255,255,0.5)",
                            "outline_width": 0.8,
                        },
                    }

        if mode not in ["boxplot", "letter_value", "violin", "trend", "tornado", "scenario"]:
            series_style_proxy_map = self._build_series_display_name_map(active_keys)
            proxy_keys: list[str] = []
            proxy_styles: dict[str, dict] = {}
            for raw_key in list(active_keys or []):
                key = str(raw_key)
                display_key = series_style_proxy_map.get(key, key)
                proxy_keys.append(display_key)
                proxy_styles[display_key] = dict(active_styles.get(key, {}))
            active_keys = proxy_keys
            active_styles = proxy_styles

        dlg = StyleManagerDialog(
            active_keys,
            active_styles,
            show_bar,
            show_curve,
            show_mean,
            is_unified,
            self,
            style_profile=style_profile,
        )
        if dlg.exec():
            if series_style_proxy_map:
                remapped_styles: dict[str, dict] = {}
                for internal_key, display_key in series_style_proxy_map.items():
                    if display_key in active_styles:
                        remapped_styles[internal_key] = dict(active_styles.get(display_key, {}))
                if remapped_styles:
                    self._series_style.update(remapped_styles)
            # 用户点击确认后，根据当前的图表模式触发相应的重绘操作
            if mode in ["boxplot", "letter_value", "violin", "trend"]:
                self._activate_analysis_view()
            elif mode in ["tornado", "scenario"]:
                self._activate_analysis_view()
            else:
                self.init_chart_via_tempfile()

    def rebuild_cdf_prob_boxes(self):
        """
        兼容性存根：清空并卸载因架构升级而废弃的旧版 CDF 概率框布局。
        """
        if getattr(self, "cdf_prob_layout", None):
            while self.cdf_prob_layout.count():
                item = self.cdf_prob_layout.takeAt(0)
                w = item.widget()
                if w:
                    w.deleteLater()
        return

    # ---------------------------------------------------------
    # 3.6 WebEngine 渲染通信 (WebEngine Rendering & IPC)
    # ---------------------------------------------------------

    def init_chart_via_tempfile(self):
        """
        非高级视图标准绘图指令下发入口。
        根据当前的视图指令（直方图、CDF、核密度等），通过专用调度路由器进行数据封包及渲染。
        """
        try:
            cmd = getattr(self, "_current_view_data", "auto")
            view_mode = self._apply_view_command_state(cmd)
            ResultsRenderDispatcher.dispatch_non_advanced_render(self, view_mode)
            return

        except Exception as e:
            tb = traceback.format_exc()
            print("\n[init_chart_via_tempfile 发生异常]\n", tb)

            fig = go.Figure()
            fig.add_annotation(
                text=f"渲染错误：{type(e).__name__}<br>{str(e)}",
                xref="paper", yref="paper", x=0.5, y=0.5, showarrow=False
            )
            fig.update_layout(
                template="plotly_white",
                margin=dict(l=self.MARGIN_L, r=self.MARGIN_R, t=self.MARGIN_T, b=self.MARGIN_B),
            )
            plot_json = json.dumps(fig.to_dict(), cls=plotly.utils.PlotlyJSONEncoder)
            self._load_plotly_html(plot_json, "histogram", "")
            return

    def _load_plotly_html(self, plot_json, js_mode, js_logic, *args):
        """
        底层核心通信：将 Python 侧生成的 Plotly JSON 数据字典封包，
        通过 Qt WebEngine 注入并在内置的 Chromium 浏览器中执行。
        """
        initial_lr = None
        # 如果启用了 Qt 原生 Overlay，则禁止向 JS 传递 initial_lr，以彻底根除残影与双重遮罩 Bug
        if not getattr(self, "_use_qt_overlay", False):
            try:
                if hasattr(self, "current_left") and hasattr(self, "current_right"):
                    initial_lr = (float(self.current_left), float(self.current_right))
            except Exception:
                initial_lr = None
            
        self._show_skeleton("正在绘图...")
        self._render_token += 1

        if DEBUG_WEB: self._debug_dump_web_state("before_load_plot")
        self._plot_host.load_plot(
            plot_json=plot_json,
            js_mode=js_mode,
            js_logic=js_logic,
            initial_lr=initial_lr,
        )

        self._data_sent = True
        self._maybe_hide_skeleton()

        try:
            self._chart_ctrl.schedule_rect_sync()
        except Exception:
            pass
        QTimer.singleShot(120, self._chart_ctrl.schedule_rect_sync)

    def _on_webview_load_finished(self, ok: bool):
        """
        WebEngine 页面加载完毕的信号回调：解绑 JS 安全限制，准备开始接收并渲染数据。
        """
        if DEBUG_WEB:
            self._debug_dump_web_state(tag=f"loadFinished ok={ok}")
        
        try:
            if hasattr(self, "_plot_host") and self._plot_host is not None:
                self._plot_host.on_webview_load_finished(ok)
        except Exception:
            pass

        try:
            if hasattr(self, "_chart_ctrl") and self._chart_ctrl is not None:
                self._chart_ctrl.schedule_rect_sync()
                QTimer.singleShot(120, self._chart_ctrl.schedule_rect_sync)
        except Exception:
            pass

        self._webview_loaded = bool(ok)
        if not ok:
            self._show_skeleton("绘图区加载失败（请重试）")
            return

        self._maybe_hide_skeleton()

    def _debug_dump_web_state(self, tag=""):
        """
        WebEngine 内核探针（仅 DEBUG_WEB = True 生效）。
        抓取当前 HTML 的 DOM 状态、Plotly 对象就绪情况及实际物理渲染参数，回传并输出至 Python 控制台。
        """
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

        self.web_view.page().runJavaScript(js, lambda v: print("[WEB_STATE]", v))
        print("[WEB_SIZE]", self.web_view.size(), "isVisible=", self.web_view.isVisible(), "url=", self.web_view.url().toString())

    def _poll_plotly_ready(self):
        """
        轮询探测 Plotly 实例状态：防止库文件未加载完成时强行下发绘图指令导致崩溃。
        """
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

        try:
            self.web_view.page().runJavaScript(js, _cb)
        except Exception:
            self._js_inflight = False

    def _render_channel_hist(self, view_mode: str):
        """通道渲染器：负责渲染连续分布（直方图）及混合叠加态。"""
        ResultsMainRenderService.render_channel_hist(self, view_mode)

    def _render_channel_discrete(self):
        """通道渲染器：负责渲染离散分布（柱状图模式）。"""
        ResultsMainRenderService.render_channel_discrete(self)

    def _render_channel_pdf(self):
        """通道渲染器：负责渲染平滑的核密度估计 (KDE) PDF 曲线。"""
        ResultsMainRenderService.render_channel_pdf(self)

    def _render_channel_cdf(self):
        """通道渲染器：负责渲染累积概率分布 (CDF) 曲线及置信带。"""
        ResultsMainRenderService.render_channel_cdf(self)

    def _render_channel_auto(self):
        """通道渲染器：基于上下文数据类型自动推断最适合的渲染模式。"""
        ResultsMainRenderService.render_channel_auto(self)

    @Slot(object, object)
    def _on_pdf_kde_done(self, x_grid, y_map):
        """后台 Worker：处理完 KDE PDF 的繁重计算后，将结果装载到视图并触发呈现。"""
        self._pdf_x_grid = x_grid
        self._pdf_y_map = y_map
        self._render_pdf_from_cache()

    @Slot(str, object, object)
    def _on_hist_kde_done(self, cache_key, x_grid, y_map):
        """后台 Worker：完成平滑 KDE 曲线拟合，将结果写入内存缓存并安排更新画布。"""
        try:
            # 1. 字典为空说明计算引擎崩溃，安全退出
            if not y_map: 
                return

            # 2. 忽略键名错乱，强制提取字典内首项结果
            y = list(y_map.values())[0] if len(y_map) > 0 else None
            
            # 3. 数据为空则安全退出
            if y is None: 
                return

            # 4. 存入内存缓存并触发最终形态的重绘
            self._kde_cache[cache_key] = (x_grid, y)
            if getattr(self, "chart_mode", "hist") in ("hist", "histogram"):
                self.init_chart_via_tempfile()
                
        except Exception as e:
            # 安全兜底：将隐性报错写入本地日志系统，便于日后排查定位
            ResultsMainRenderService.log_kde_overlay_error()

    def _render_pdf_from_cache(self):
        """优化策略：直接从内存缓存中抽取复用现成的 KDE 数据并渲染，显著缩短重复计算耗时。"""
        ResultsMainRenderService.render_pdf_from_cache(self)

    # ---------------------------------------------------------
    # 3.7 高级分析视图调度 (Advanced Analysis Views)
    # ---------------------------------------------------------

    def _is_advanced_mode(self):
        """
        判断当前是否处于高级分析的专属视图模式下（如摘要图、敏感性龙卷风图、情景分析等）。
        """
        return ResultsInteractionService.is_advanced_mode(self)

    def _check_and_prepare_advanced_mode(self):
        """
        拦截与数据准备：在正式进入三大高级模式前，预检查数据集样本量是否达标，并完成前置清洗。
        """
        return ResultsInteractionService.check_and_prepare_advanced_mode(self)

    def _activate_analysis_view(self):
        """
        高级分析视图的总路由激活中心。
        收集请求，调用下层服务生成结果，并负责在 UI 的 StackedWidget 中切换正确的层级进行显示。
        """
        mode = ResultsRuntimeStateHelper.get_current_analysis_mode(self, "")
        ResultsAdvancedRouter.sync_view_state(self, mode)

        try:
            ResultsAdvancedRouter.dispatch_render_entry(self, mode)
        except Exception as e:
            is_boxplot_family = ResultsAdvancedRouter.is_boxplot_family(mode)
            print(traceback.format_exc())
            if mode in ["tornado", "scenario"]:
                self.tornado_view.placeholder_label.setText(f"渲染图表失败: {str(e)}")
                self.tornado_view.placeholder_label.show()
                self.tornado_view.web_view.hide()
            elif is_boxplot_family:
                self.boxplot_view.placeholder_label.setText(f"渲染摘要图失败: {str(e)}")
                self.boxplot_view.placeholder_label.show()
                self.boxplot_view.web_view.hide()

    def _get_advanced_runtime_state(self):
        """
        统一构建高级模式的运行时状态对象（集成情景参数、龙卷风配置等）。
        """
        runtime = ResultsAdvancedRenderService.build_runtime_state(self)
        ResultsRuntimeStateHelper.apply_runtime_configs(
            self,
            scenario_cfg=runtime.scenario_config,
            tornado_cfg=runtime.tornado_config,
        )
        return runtime

    def _check_and_reset_styles(self, new_major_mode):
        """
        高级样式的自治重置机制：在不同的基础与高级主模式间跨越切换时，自动清除上一个模式留存的临时图形样式。
        """
        old_mode = getattr(self, "_last_major_mode", "main")
        if old_mode != new_major_mode:
            # 1. 还原基础模式专属样式 (直方图、CDF、箱线图共用基础配置)
            if hasattr(self, "display_keys"):
                self._build_series_style(self.display_keys)
            
            # 2. 格式化龙卷风/情景分析的专属临时样式
            self._tornado_styles = build_default_tornado_styles(DRISK_COLOR_CYCLE)
            self._tornado_line_styles = {}
            
            # 3. 更新系统状态标记
            self._last_major_mode = new_major_mode

    def _prepare_advanced_payload(self, min_rows: int) -> AdvancedPreparedPayload:
        """
        数据制备引擎核心层：按需从底层驱动提取相关性特征与输入输出对应矩阵，并具备容错保护拦截能力。
        """
        if self.sim_id is None or not self.current_key:
            raise AdvancedPreparationError("missing_context", "当前模拟上下文无效。")

        all_sims = get_all_simulations()
        sim_result = all_sims.get(self.sim_id)
        if not sim_result:
            raise AdvancedPreparationError("missing_sim", f"找不到 ID 为 {self.sim_id} 的模拟数据。")

        output_key = self.label_to_cell_key.get(self.current_key, self.current_key)
        output_data = resolve_output_data(sim_result, output_key, self.sim_id)
        if output_data is None:
            raise AdvancedPreparationError("missing_output", f"丢失输出单元格 {output_key} 的数据。")

        return build_advanced_payload(
            sim_result,
            output_key=output_key,
            output_data=output_data,
            min_rows=min_rows,
        )

    def _restore_main_view_from_advanced_if_needed(self):
        """
        UI 退场清理机制：如果用户在某个高级视图内点击了需要返回基础视图的按钮，则清除高亮状态并执行环境切换。
        """
        return ResultsRenderDispatcher.restore_main_view_from_advanced_if_needed(self)

    def _switch_boxplot_mode(self, mode):
        """
        箱线图内部子模式（箱线图、小提琴图、信值图）的切换与重绘命令下发。
        """
        ResultsRuntimeStateHelper.set_boxplot_mode(self, mode)
        self._activate_analysis_view()

    def _render_advanced_boxplot_mode(self, mode):
        """
        装配参数，调用底层绘图工厂渲染选定的摘要图形式。
        """
        runtime = self._get_advanced_runtime_state()
        summary_label_map = self._build_summary_display_label_map(getattr(self, "display_keys", []))
        ResultsAdvancedRenderService.render_boxplot(
            self,
            mode,
            runtime,
            display_label_map=summary_label_map,
        )

    def _run_scenario(self, min_p, max_p):
        """
        执行预定义参数的情景分析（例如排查特定尾部 5% 坏账情景发生时的变量偏度表现）。
        """
        current_cfg = ResultsRuntimeStateHelper.get_scenario_config(self)
        ResultsRuntimeStateHelper.set_scenario_mode(
            self,
            {
                "display_mode": 0,
                "min_pct": min_p,
                "max_pct": max_p,
                "significance_threshold": current_cfg.get("significance_threshold", 0.5),
                "display_limit": current_cfg.get("display_limit", None),
            },
        )
        self._activate_analysis_view()

    def _open_scenario_settings(self, is_custom=False):
        """
        调出情景分析高级设置弹窗，允许用户微调截断占比、过滤阈值等，确认后直接刷新画布。
        """
        current_cfg = ResultsRuntimeStateHelper.get_scenario_config(self)
        
        dlg = ScenarioSettingsDialog(
            current_display_idx=current_cfg['display_mode'],
            current_min=current_cfg['min_pct'],
            current_max=current_cfg['max_pct'],
            current_threshold=current_cfg.get('significance_threshold', 0.5),
            current_display_limit=current_cfg.get('display_limit', None),
            parent=self
        )
        
        if dlg.exec():
            ResultsRuntimeStateHelper.set_scenario_mode(self, dlg.get_settings())
            self._activate_analysis_view()
        else:
            if is_custom and ResultsRuntimeStateHelper.get_current_analysis_mode(self, "") != "scenario":
                self.btn_scenario.setChecked(False)

    def _prepare_scenario_render_payload(self) -> AdvancedPreparedPayload:
        """
        特定于情景分析的数据提取：复用高级准备载荷模块，并配置所需样本参数。
        """
        return self._prepare_advanced_payload(min_rows=5)

    def _render_advanced_scenario_mode(self):
        """
        组装参数流，将预处理好的情景分析矩阵推送至 HTML 渲染管线。
        """
        payload = self._prepare_scenario_render_payload()
        runtime = self._get_advanced_runtime_state()
        ResultsAdvancedRenderService.render_scenario(self, payload, runtime)

    def _toggle_tornado_view(self, checked):
        """
        龙卷风敏感性图表的总控开关及激活跳转。
        """
        ResultsInteractionService.toggle_tornado_view(self, checked)

    def _switch_tornado_mode(self, mode):
        """
        切换龙卷风内部子模式（分箱平均法、相关系数法等）。
        """
        ResultsRuntimeStateHelper.set_tornado_mode(self, mode)
        self._activate_analysis_view()

    def _prepare_tornado_render_payload(self) -> AdvancedPreparedPayload:
        """
        特定于龙卷风的数据提取：默认采用 10 组离散分箱法，要求样本池数据更加充足。
        """
        return self._prepare_advanced_payload(min_rows=10)

    def _render_tornado_payload(self, mode: str, payload: AdvancedPreparedPayload):
        """
        完成数据计算后，最终传递负载给底层 JS 绘图模块绘制敏感性柱状堆积图。
        """
        runtime = self._get_advanced_runtime_state()
        self._update_bottom_bar_visibility()
        ResultsAdvancedRenderService.render_tornado(self, mode, payload, runtime)

    def _load_and_render_tornado(self):
        """
        装配整个龙卷风绘图的生命周期，提供报错降级处理（如无关联输入变量时安全拦截并友好提示）。
        """
        try:
            payload = self._prepare_tornado_render_payload()
            mode = ResultsRuntimeStateHelper.get_tornado_mode(self, "bins")
            self._render_tornado_payload(mode, payload)
        except AdvancedPreparationError as e:
            if e.code == "missing_input":
                QMessageBox.information(self, "提示", str(e))
            else:
                QMessageBox.warning(self, "数据错误", str(e))
            self.btn_tornado.setChecked(False)
            self._toggle_tornado_view(False)
        except Exception as e:
            print(traceback.format_exc())
            QMessageBox.critical(self, "渲染错误", f"渲染敏感性图失败: {str(e)}")
            self.btn_tornado.setChecked(False)
            self._toggle_tornado_view(False)

    def _open_tornado_settings(self):
        """
        提供龙卷风图自定义阈值的控制面板。
        """
        current_cfg = ResultsRuntimeStateHelper.get_tornado_config(self)
        
        dlg = TornadoSettingsDialog(current_cfg, parent=self)
        if dlg.exec():
            ResultsRuntimeStateHelper.set_tornado_config(self, dlg.get_settings())
            self._load_and_render_tornado() # 确认修改后直接重绘

    # ---------------------------------------------------------
    # 3.8 后台统计指标计算与表格 UI (Stats Calculation & UI)
    # ---------------------------------------------------------

    def _invalidate_stats_cache(self):
        """
        释放右侧指标表的陈旧缓存，并将任务状态标记为待更新，以防展示过期数据。
        """
        self._stats_cache_valid = False
        self.job_controller.invalidate_stats()

    def _series_key_to_display_label(self, series_key: str) -> str:
        """从内部字典的键值中智能提取含有模拟 ID 尾缀的友好展示标签"""
        key_text = str(series_key or "")
        sid = extract_sim_id_from_series_key(key_text)
        if sid is None:
            return key_text
        head = key_text[: key_text.rfind("(sim")].strip() if "(sim" in key_text else key_text
        sim_label = get_sim_display_name(sid)
        if head:
            return f"{head} ({sim_label})"
        return sim_label

    def _build_series_display_name_map(self, keys) -> dict[str, str]:
        """构建键名到界面文字的安全映射表，带有冲突自动编号保护机制"""
        mapping: dict[str, str] = {}
        used: set[str] = set()
        for raw_key in list(keys or []):
            key = str(raw_key)
            base_label = str(self._series_key_to_display_label(key) or key).strip()
            if not base_label:
                base_label = key
            final_label = base_label
            sid = extract_sim_id_from_series_key(key)
            if final_label in used and sid is not None:
                final_label = f"{base_label} [ID {sid}]"
            idx = 2
            while final_label in used:
                final_label = f"{base_label} #{idx}"
                idx += 1
            used.add(final_label)
            mapping[key] = final_label
        return mapping

    def _resolve_series_visible_name(self, series_key: str) -> tuple[str, str]:
        """解析核心变量名及来源工程的显示名称"""
        k = str(series_key or "")
        label_part = k[:k.rfind('(sim')].strip() if '(sim' in k else k
        sim_id_for_label = extract_sim_id_from_series_key(k)
        if sim_id_for_label is None:
            sim_id_for_label = self.sim_id
        sim_part = get_sim_display_name(sim_id_for_label)

        ck = str(self.label_to_cell_key.get(label_part, label_part) or "")
        base_ck = str(self.label_to_cell_key.get(self.current_key, self.current_key) or "")
        ck_norm = ck.replace("$", "").upper()
        base_ck_norm = base_ck.replace("$", "").upper()

        if ck_norm == base_ck_norm and str(getattr(self, "_base_chart_title", "") or "").strip():
            visible_name = str(self._base_chart_title).strip()
        else:
            attrs = {}
            if sim_id_for_label == self.sim_id:
                attrs = self.attrs_map.get(label_part, {}) or {}
            if (not attrs) and sim_id_for_label is not None:
                try:
                    attrs = bridge.get_attributes(
                        int(sim_id_for_label),
                        ck,
                        kind=getattr(self, "_base_kind", "output"),
                    ) or {}
                except Exception:
                    attrs = {}

            xl_ctx = None
            try:
                xl_ctx = _safe_excel_app()
            except Exception:
                xl_ctx = None

            visible_name = resolve_visible_variable_name(
                ck,
                attrs,
                excel_app=xl_ctx,
                fallback_label=label_part,
            )

        return str(visible_name or "").strip(), str(sim_part or get_sim_display_name(self.sim_id))

    def _build_summary_display_label_map(self, keys) -> dict[str, str]:
        """特别为摘要图设计的映射器：在不产生冗余的前提下为同源不同次运算添加前缀"""
        label_map: dict[str, str] = {}
        sim_ids = [extract_sim_id_from_series_key(str(k)) for k in list(keys or [])]
        unique_sim_ids = {sid for sid in sim_ids if sid is not None}
        show_sim_suffix = len(unique_sim_ids) > 1
        for series_key in list(keys or []):
            visible_name, sim_name = self._resolve_series_visible_name(series_key)
            if not visible_name:
                continue
            final_name = visible_name
            if show_sim_suffix and sim_name:
                final_name = f"{visible_name} [{sim_name}]"
            label_map[str(series_key)] = final_name
        return label_map

    def _resolve_stats_header_label(self, series_key: str) -> str:
        """
        输出最终供表头 UI 呈现场景的高级解析标签。
        """
        visible_name, sim_part = self._resolve_series_visible_name(series_key)
        return f"{visible_name} [{sim_part}]"

    def _render_stats_placeholder(self, keys):
        """
        异步平滑过渡：在后台线程艰难计算庞大统计指标的间隙期，立刻在表格区生成占位符，避免程序假死。
        """
        try:
            local_keys = list(keys)
            
            # ✅ 同步判断理论分布列是否需要占位
            is_theory_overlay = self._should_show_theory_stats()
            theory_label = self._theory_label()
            if is_theory_overlay:
                local_keys.append(theory_label)

            self.stats_table.clear()
            self.stats_table.setColumnCount(1 + len(local_keys))
            self.stats_table.setRowCount(1)
            
            clean_headers = []
            for k in local_keys:
                if k == theory_label:
                    clean_headers.append(theory_label)
                else:
                    clean_headers.append(self._resolve_stats_header_label(k))
                    
            self.stats_table.setHorizontalHeaderLabels(["指标"] + clean_headers)
            
            item = QTableWidgetItem("统计计算中…")
            item.setFlags(Qt.NoItemFlags)
            self.stats_table.setItem(0, 0, item)
            self.stats_table.resizeRowsToContents()
            self.stats_table.setRowHeight(0, max(24, self.stats_table.rowHeight(0))) 
            for j in range(1, 1 + len(local_keys)):
                it = QTableWidgetItem("")
                it.setFlags(Qt.NoItemFlags)
                self.stats_table.setItem(0, j, it)
        except Exception as e:
            print(f"占位符渲染错误: {e}")

    def _request_stats_job(self, keys):
        """
        异步引擎驱动器：打包当前选定区域的参数状态，提交至高并发后台线程池执行复杂的统计与聚合计算。
        """
        lx = getattr(self, 'current_left', None)
        rx = getattr(self, 'current_right', None)
        
        series_map = {}
        for k in keys:
            s = self.series_map.get(k)
            if s is None or len(s) == 0:
                s = pd.Series([0.0] * 10)
            series_map[k] = s

        self.job_controller.request_stats(
            keys=list(keys),
            sim_id=self.sim_id,
            label_to_cell_key=self.label_to_cell_key,
            series_map=series_map,
            left_x=lx,
            right_x=rx,
            series_kind=getattr(self, "_base_kind", "output"),
        )

    def _on_ci_engine_changed(self, idx):
        """
        置信区间计算引擎切换回调：支持快速法、重抽样法等不同科学要求。
        """
        modes = ["fast", "bootstrap", "bca"]
        safe_idx = int(idx)
        if safe_idx < 0 or safe_idx >= len(modes):
            safe_idx = 0
        self.ci_engine_mode = modes[safe_idx]
        if hasattr(self, "combo_ci_engine") and self.combo_ci_engine.currentIndex() != safe_idx:
            with QSignalBlocker(self.combo_ci_engine):
                self.combo_ci_engine.setCurrentIndex(safe_idx)
        
        # ✅ 触发图表重新绘制，让 CDF 置信带跟随引擎变化实时更新
        self.init_chart_via_tempfile() 
        
        self._invalidate_stats_cache()
        self.update_stats_ui(update_stats_table=True)
    
    def update_stats_ui(self, update_stats_table: bool = True):
        """
        统计数据更新的中枢控制台。
        无论是主动拖拽、还是被动参数更改，都在这里进行物理坐标计算与视图响应，并根据参数按需派发表格任务。
        """

        # 1. 始终优先计算顶部滑块所需的物理切分面概率，保持 UI 阻滞感为 0
        l, r = self.current_left, self.current_right
        keys = list(getattr(self, "display_keys", [self.current_key]))[:3]
        layers_data = []

        def _calc_probs(series):
            if series is None or len(series) == 0: return 0.0, 0.0, 0.0
            arr = np.sort(series.values)
            n = len(arr)
            # 采用右闭合的离散边界语义：左侧尾部包含边界值 <= L
            idx_l = np.searchsorted(arr, l, side='right')
            idx_r = np.searchsorted(arr, r, side='right')
            pl = idx_l / n
            pr = (n - idx_r) / n
            pm = 1.0 - pl - pr
            return pl, pm, pr

        for i, k in enumerate(keys):
            s = self.series_map.get(k)
            style = self._series_style.get(k, {})
            color = style.get("color", "#333")

            if s is None:
                probs = (0.0, 0.0, 0.0)
            else:
                probs = _calc_probs(s)

            layers_data.append({
                'key': k,
                'color': color,
                'probs': probs
            })

        # 输入变量具备“天生神力”，可以叠加显示真值概率层
        if self._should_show_theory_stats():
            try:
                cdf_l = self._safe_theory_cdf_scalar(l)
                cdf_r = self._safe_theory_cdf_scalar(r)
                pl = float(cdf_l)          
                pr = 1.0 - float(cdf_r)    
                pm = 1.0 - pl - pr         
                
                layers_data.append({
                    'key': self._theory_label(),
                    'color': '#000000',
                    'probs': (pl, max(0.0, pm), max(0.0, pr))
                })
            except Exception as e:
                print(f"理论概率层计算异常: {e}")
                pass

        if hasattr(self.range_slider, "set_layers"):
            self.range_slider.set_layers(layers_data)

        # 2. 对于连续拖拽场景下的高频抛弃，直接中止执行底层表格重算
        if not update_stats_table: return

        # ==========================================
        # 3. 如果放行，进行统一指标拦截与装载
        # ==========================================
        keys = list(getattr(self, "display_keys", [self.current_key]))
        if not keys:
            self.stats_table.setRowCount(0)
            return

        if not getattr(self, "_stats_cache_valid", False) or self.stats_table.rowCount() == 0:
            self._render_stats_placeholder(keys)
            
        self._request_stats_job(keys)

    @Slot(list, list)
    def _on_stats_done(self, keys, stats_list):
        """
        当底层指标字典返回主线程后的装配终点：在此对齐数据矩阵、应用颜色并注入表格视图。
        如果当前激活了理论对比模式，会同步插入真值验证列供直观审计。
        """
        self._stats_cache_valid = True
        try:
            # 深层复制键名与结果列表，规避脏引用
            local_keys = list(keys)
            local_stats = list(stats_list)

            # ✅ 1. 拦截检查：是否有资格补充展现理论值
            is_theory_overlay = self._should_show_theory_stats()
            theory_label = self._theory_label()
            
            if is_theory_overlay:
                theo_stats = {}
                try:
                    lx = getattr(self, "current_left", 0.0)
                    rx = getattr(self, "current_right", 0.0)
                    
                    # 设立临时队列接收同步回调的结果集
                    result_container = []
                    def _hook(token, rows):
                        result_container.append(rows)

                    try:
                        # ✅ 首选带截断和平移标记的新版 API，保证理论与模拟样本绝对一致
                        job = ModelerStatsJob(
                            token=0,
                            func_name=getattr(self, "_theory_func_name", ""),
                            params=list(getattr(self, "_theory_params", [])),
                            markers=dict(getattr(self, "_theory_markers", {}) or {}),
                            cell_label=theory_label,
                            left_x=lx,
                            right_x=rx
                        )
                    except TypeError:
                        # 兼容处理：降级返回旧版调用接口
                        job = ModelerStatsJob(
                            0,
                            self.theory_dist_obj,
                            theory_label,
                            lx,
                            rx
                        )
                    
                    job.signals.done.connect(_hook)
                    
                    # 🚀 神来之笔：纯数学公式计算性能极高（< 5ms），直接主线程同步阻塞执行即可，免去排队负担
                    job.run() 
                    
                    if result_container and len(result_container) > 0:
                        rows = result_container[0]
                        if isinstance(rows, list) and len(rows) > 0:
                            theo_stats = rows[0]
                        elif isinstance(rows, dict):
                            theo_stats = rows
                except Exception as e:
                    print(f"调用 ModelerStatsJob 同步计算理论指标失败: {e}")
                    traceback.print_exc()

                if isinstance(theo_stats, dict):
                    # 剥离无用的占位指标名，保持图表清爽
                    for _k in (
                        "count", "模拟次数", "数值数", "n_iterations",
                        "error_count", "错误数",
                        "filtered_count", "过滤数",
                    ):
                        theo_stats.pop(_k, None)

                # 将洗净的纯真值数学列附在末尾
                local_keys.append(theory_label)
                local_stats.append(theo_stats)

            # 调用统用核心装配器，重构矩阵为可视化二维表数据
            rows = ui_stats.assemble_simulation_rows(
                stats_map_list=local_stats,
                col_names=local_keys,
                cell_labels=local_keys
            )
            
            clean_headers = []
            header_colors = []
            st_map = getattr(self, "_series_style", {})
            
            for k in local_keys:
                if k == theory_label:
                    clean_headers.append(theory_label)
                    header_colors.append("#000000")
                else:
                    header_colors.append(st_map.get(k, {}).get("color", "#f0f0f0"))
                    clean_headers.append(self._resolve_stats_header_label(k))

            ui_stats.render_stats_table(self.stats_table, rows, column_headers=clean_headers, header_colors=header_colors)
            
        except Exception as e:
            # 终极防卡死防御：捕获组装异常并暴露于表格首行，为使用者提供最高级别的安全退出
            traceback.print_exc()
            self.stats_table.setColumnCount(1)
            self.stats_table.setRowCount(1)
            item = QTableWidgetItem(f"表格渲染失败: {e}")
            item.setForeground(QColor("red"))
            self.stats_table.setItem(0, 0, item)

    @Slot(list)
    def _on_cdf_stats_done(self, rows):
        """
        接收并呈现带有一级置信区间的进阶 CDF 统计算法结果。
        """
        self._stats_cache_valid = True
        
        # ✅ 新增：在内存中永久保存高成本算法的结果供反向读取
        keys = list(getattr(self, "display_keys", [self.current_key]))
        engine_mode = getattr(self, "ci_engine_mode", "fast")
        ci_level = getattr(self, "cdf_ci_level", 0.95)
        table_cache_key = (tuple(keys), engine_mode, round(ci_level, 6))
        
        if not hasattr(self, "_cdf_table_cache"):
            self._cdf_table_cache = {}
        self._cdf_table_cache[table_cache_key] = rows

        clean_headers = []
        for k in keys:
            clean_headers.append(self._resolve_stats_header_label(k))
            
        header_colors = [(self._series_style.get(k, {}) or {}).get("color", "#f0f0f0") for k in keys]
        
        ui_stats.render_stats_table(self.stats_table, rows, column_headers=clean_headers, header_colors=header_colors)
    
    def _on_cdf_ci_changed(self, val):
        """
        截获用户对于统计区间水平精度（如 95% -> 99%）的更改操作并应用。
        """
        self.cdf_ci_level = float(val)
        
        # 1. 指挥绘图区刷新视觉区域
        self.init_chart_via_tempfile()
        # 2. 清空缓存阻碍，强制底层数据重新计算
        self._invalidate_stats_cache()
        # 3. 接管刷新
        self.update_stats_ui(update_stats_table=True)

    def on_stats_cell_clicked(self, row, col):
        """
        UI 体贴设计：针对较长字段文本展示被挤压的问题，允许用户点击对应方格弹出完整摘要。
        """
        item = self.stats_table.item(row, col)
        if not item:
            return

        text = item.text().strip()
        if not text:
            return

        if len(text) > 18:
            QMessageBox.information(self, "完整内容", text)

    # ---------------------------------------------------------
    # 3.9 滑块交互与数值吸附引擎 (Slider Interaction & Snapping)
    # ---------------------------------------------------------

    def _clamp_and_sync_slider(self):
        """
        双轨约束系统：核验区间值的合理范围并将校正值传递给 QRangeSlider。
        """
        if not hasattr(self, "current_left") or not hasattr(self, "current_right"):
            mid = (self.x_range_min + self.x_range_max) / 2.0
            self.current_left = mid
            self.current_right = mid

        self.current_left = max(self.x_range_min, min(self.current_left, self.x_range_max))
        self.current_right = max(self.x_range_min, min(self.current_right, self.x_range_max))
        if self.current_left > self.current_right:
            self.current_left, self.current_right = self.current_left, self.current_left

        self.range_slider.setRangeLimit(self.x_range_min, self.x_range_max)
        self.range_slider.setRangeValues(self.current_left, self.current_right)

        self._update_vlines_only()

    def _maybe_send_vlines(self, l, r):
        """
        智能渲染管线拦截器：如果判断出 Qt 原生层已被激活，便直接交由原生系统闪电响应视觉更新，阻断高延时的 JS 发包通信。
        """
        if getattr(self, "_use_qt_overlay", False):
            # 虽然免去发包，但必须将边界线位置推送给相应的处理模块展示
            if hasattr(self, "_chart_ctrl") and hasattr(self, "overlay"):
                self._chart_ctrl.update_visuals(
                    l, r, 
                    getattr(self, "x_range_min", l), 
                    getattr(self, "x_range_max", r), 
                    getattr(self, "chart_mode", "hist")
                )
            return

        lr = (float(l), float(r))
        if self._last_sent_lr is not None:
            if abs(lr[0] - self._last_sent_lr[0]) < 1e-12 and abs(lr[1] - self._last_sent_lr[1]) < 1e-12:
                return

        self._last_sent_lr = lr
        try:
            self.web_view.page().runJavaScript(f"updateVLines({lr[0]}, {lr[1]});")
        except Exception:
            pass

    def _update_vlines_only(self):
        """
        精益局部刷新法：跳出底层 HTML 刷新流，只处理最外侧悬浮指示虚线。
        """
        if getattr(self, "_use_qt_overlay", False) and hasattr(self, "overlay"):
            self._chart_ctrl.update_visuals(
                self.current_left, self.current_right,
                self.x_range_min, self.x_range_max,
                getattr(self, "chart_mode", "hist")
            )
            return

        try:
            self.web_view.page().runJavaScript(
                f"updateVLines({float(self.current_left)}, {float(self.current_right)});"
            )
        except Exception:
            pass

    def perform_buffered_update(self, force_full: bool = False):
        """
        带有防抖的渐进更新器：在拖拽过程中暂时阻断复杂的表格指标算力消耗，等释手操作再一并收网计算。
        """
        self._update_vlines_only()
        full = force_full or (not getattr(self, "_slider_dragging", False))
        if full:
            self.update_stats_ui(update_stats_table=True)
        else:
            self.update_stats_ui(update_stats_table=False)

    def on_range_slider_changed(self, l, r):
        """
        核心物理与 UI 交互防抖桥：接受物理世界的坐标变动、施加磁性吸附对齐策略，并激活微秒级更新节流阀。
        """
        l = float(l)
        r = float(r)

        dragging = getattr(self, "_slider_dragging", False)
        discrete = getattr(self, "is_discrete_view", False)

        if getattr(self, "_drag_snap_enabled", True):
            l2 = self._snap_value_drag(l)
            r2 = self._snap_value_drag(r)
        else:
            l2 = l
            r2 = r

        l2 = max(self.x_range_min, min(float(l2), self.x_range_max))
        r2 = max(self.x_range_min, min(float(r2), self.x_range_max))
        if l2 > r2:
            l2, r2 = r2, l2

        need_snap_visual = (
                dragging
                and getattr(self, "_cont_fixed_enabled", False)
                and (not discrete)
                and (abs(l2 - l) > 1e-12 or abs(r2 - r) > 1e-12)
        )

        if not getattr(self, "_in_slider_sync", False):
            if need_snap_visual or ((not dragging) and (abs(l2 - l) > 1e-12 or abs(r2 - r) > 1e-12)):
                self._in_slider_sync = True
                try:
                    self.range_slider.blockSignals(True)
                    self.range_slider.setRangeValues(l2, r2)
                finally:
                    self.range_slider.blockSignals(False)
                    self._in_slider_sync = False

        self.current_left = float(l2)
        self.current_right = float(r2)

        self._update_vlines_only()
        if not self.throttle_timer.isActive():
            self.throttle_timer.start()

        self._update_x_edits_pos()

    def on_slider_drag_started(self):
        """
        释放捕捉点标记，记录此时的操作边界，阻断外部表格干扰行为。
        """
        self._slider_dragging = True
        if getattr(self, "is_discrete_view", False):
            pts = getattr(self, "_discrete_points", None)
            if pts is not None and len(pts) > 0:
                l0 = float(getattr(self, "current_left", pts[0]))
                r0 = float(getattr(self, "current_right", pts[-1]))
                self._drag_last_left_idx = int(np.searchsorted(pts, l0, side="left"))
                self._drag_last_right_idx = int(np.searchsorted(pts, r0, side="left"))
            else:
                self._drag_last_left_idx = None
                self._drag_last_right_idx = None

    def on_slider_drag_finished(self, l, r):
        """
        归位结算器：将滑动模块安放于终局位置，重算坐标，重载指标结果层。
        """
        self._drag_last_left_idx = None
        self._drag_last_right_idx = None

        self._slider_dragging = False

        l = float(l)
        r = float(r)

        if getattr(self, "_drag_snap_enabled", True):
            l = self._snap_value_drag(l)
            r = self._snap_value_drag(r)
        else:
            l = float(l)
            r = float(r)

        l = max(self.x_range_min, min(l, self.x_range_max))
        r = max(self.x_range_min, min(r, self.x_range_max))
        if l > r:
            l, r = r, l

        self.current_left = l
        self.current_right = r

        self.range_slider.setRangeValues(l, r)
        self._update_x_edits_pos()
        QTimer.singleShot(0, self._update_x_edits_pos)
        self.perform_buffered_update(force_full=True)
        self.update_stats_ui()

    def _get_plot_pixel_width_for_snap(self) -> int:
        """
        计算真实的绘图横轴物理宽度空间，从而换算出步进时的精确像素偏差。
        """
        try:
            w = int(self.range_slider.width())
        except Exception:
            w = 800

        w = max(100, w - int(self.MARGIN_L) - int(self.MARGIN_R))
        return w

    def _rebuild_drag_snap_grid(self):
        """
        构造虚拟力场：根据当前比例尺建立滑块的拖拽吸附网格系统，提供顿挫手感且保证数值完美对齐。
        """
        if not getattr(self, "_drag_snap_enabled", True):
            self._drag_grid_step = None
            self._drag_grid_xmin = None
            self._drag_grid_xmax = None
            return

        xmin = float(getattr(self, "x_range_min", 0.0))
        xmax = float(getattr(self, "x_range_max", 1.0))
        if not np.isfinite(xmin) or not np.isfinite(xmax) or (xmax <= xmin):
            self._drag_grid_step = None
            self._drag_grid_xmin = None
            self._drag_grid_xmax = None
            return

        if getattr(self, "_cont_fixed_enabled", False) and (not getattr(self, "is_discrete_view", False)):
            steps = int(getattr(self, "_cont_fixed_steps", 500))
            steps = max(1, steps)
            self._drag_grid_step = (xmax - xmin) / float(steps)
        else:
            px_w = self._get_plot_pixel_width_for_snap()
            self._drag_grid_step = SnapUtils.calc_grid_step(
                xmin, xmax, px_w,
                px_per_step=int(getattr(self, "_drag_snap_px_per_step", 2))
            )

        self._drag_grid_xmin = xmin
        self._drag_grid_xmax = xmax

    def _snap_value(self, x: float) -> float:
        """
        标准数值规整器。强制将任意游离的数值落回网格标线上。
        """
        x = float(x)
        xmin = float(getattr(self, "x_range_min", x))
        xmax = float(getattr(self, "x_range_max", x))
        if np.isfinite(xmin) and np.isfinite(xmax):
            x = max(xmin, min(x, xmax))

        if getattr(self, "is_discrete_view", False):
            return self._snap_discrete_nearest(x)

        step = getattr(self, "_drag_grid_step", None)
        gxmin = getattr(self, "_drag_grid_xmin", None)

        if step is None or gxmin is None:
            return x

        return SnapUtils.snap_to_grid(x, gxmin, step)

    def _snap_value_drag(self, x: float) -> float:
        """
        运动态保护屏障：拖拽移动中实施的特殊规整逻辑，专门避免离散数据越界带来的抖动错乱。
        """
        x = float(x)
        xmin = float(getattr(self, "x_range_min", x))
        xmax = float(getattr(self, "x_range_max", x))
        if np.isfinite(xmin) and np.isfinite(xmax):
            x = max(xmin, min(x, xmax))

        if getattr(self, "is_discrete_view", False):
            # 重新利用连续拖动值路径来处理离散数据集
            return x

        step = getattr(self, "_drag_grid_step", None)
        gxmin = getattr(self, "_drag_grid_xmin", None)

        if step is None or gxmin is None:
            return x

        return SnapUtils.snap_to_grid(x, gxmin, step)

    def _snap_discrete_nearest_sticky(self, x: float, last_idx: int | None, deadband_ratio: float = 0.25):
        """
        解决微小跳动问题的“带死区磁性吸附”核心算法。
        """
        x = float(x)
        pts = getattr(self, "_discrete_points", None)
        if pts is None or len(pts) == 0:
            v = round(x)
            return float(v), None

        n = len(pts)

        i = int(np.searchsorted(pts, x))
        if i <= 0:
            cand = 0
        elif i >= n:
            cand = n - 1
        else:
            left = float(pts[i - 1])
            right = float(pts[i])
            cand = (i - 1) if abs(x - left) <= abs(x - right) else i

        if last_idx is None or last_idx < 0 or last_idx >= n:
            return float(pts[cand]), int(cand)

        if cand == last_idx:
            return float(pts[last_idx]), int(last_idx)

        if abs(cand - last_idx) == 1:
            lo = min(cand, last_idx)
            hi = max(cand, last_idx)
            a = float(pts[lo])
            b = float(pts[hi])
            mid = 0.5 * (a + b)
            dead = deadband_ratio * (b - a)

            if (mid - dead) <= x <= (mid + dead):
                return float(pts[last_idx]), int(last_idx)

        return float(pts[cand]), int(cand)

    def _snap_discrete_nearest(self, x: float) -> float:
        """
        向最近点的物理回拉器。
        """
        x = float(x)
        pts = getattr(self, "_discrete_points", None)
        # 传入 dtick 和界限，开启混合吸附
        return SnapUtils.snap_to_discrete(
            x, pts, method='nearest',
            grid_step=getattr(self, "x_dtick", 1.0),
            xmin=getattr(self, "x_range_min", 0.0),
            xmax=getattr(self, "x_range_max", 1.0)
        )

    def _snap_discrete_floor(self, x: float) -> float:
        """
        向下取整模式：确保只捕捉边界内部的坚实点。
        """
        x = float(x)
        pts = getattr(self, "_discrete_points", None)
        return SnapUtils.snap_to_discrete(
            x, pts, method='floor',
            grid_step=getattr(self, "x_dtick", 1.0),
            xmin=getattr(self, "x_range_min", 0.0),
            xmax=getattr(self, "x_range_max", 1.0)
        )

    def _update_x_edits_pos(self):
        """
        坐标重载系统：计算滑块与图表的对应映射，将界面上悬浮指示框精准落在像素正确的水平点上。
        """
        update_floating_value_edits_pos(
            slider=self.range_slider,
            group_l=self.float_x_l,
            group_r=self.float_x_r,
            val_l=self.current_left,
            val_r=self.current_right,
            range_min=self.view_min,
            range_max=self.view_max,
            margin_l=self.MARGIN_L,
            margin_r=self.MARGIN_R,
            dtick=self.x_dtick,
            forced_mag=self.manual_mag,
            gap=0,  
            single_handle=False
        )

    def update_logic(self, mode):
        """
        键盘回车截获器：当用户自行键入精准数值后，负责清理附加字符、推导并应用结果，最终反转驱动图表更新。
        """
        try:
            if mode == 'xl':
                mag = self.lbl_x_suffix_l.current_mag
                divisor = 10 ** mag
                suffix = getattr(self.edt_x_l, "suffix", "")
                txt = self.edt_x_l.text().replace(',', '').replace(suffix, '')
                self.current_left = float(txt) * divisor
                # 彻底删除 CDF 模式下右侧强制等于左侧的联动逻辑

            elif mode == 'xr':
                mag = self.lbl_x_suffix_r.current_mag
                divisor = 10 ** mag
                suffix = getattr(self.edt_x_r, "suffix", "")
                txt = self.edt_x_r.text().replace(',', '').replace(suffix, '')
                self.current_right = float(txt) * divisor

            # 安全性边界修正拦截
            self.current_left = max(self.x_range_min, min(self.current_left, self.x_range_max))
            self.current_right = max(self.x_range_min, min(self.current_right, self.x_range_max))

            if self.current_left > self.current_right:
                self.current_left, self.current_right = self.current_right, self.current_left

            self.range_slider.setRangeValues(self.current_left, self.current_right)
            self._maybe_send_vlines(self.current_left, self.current_right)

            if not self.throttle_timer.isActive():
                self.throttle_timer.start()

            QTimer.singleShot(0, self._update_x_edits_pos)
            
            if mode == 'xl': self.edt_x_l.clearFocus()
            if mode == 'xr': self.edt_x_r.clearFocus()

        except Exception as e:
            print(f"模式 {mode} 的逻辑更新错误: {e}")

    def _on_slider_input_confirmed(self, mode_str: str, val_str: str):
        """
        概率倒置推演引擎：允许用户修改顶部倒三角内的百分比数值，程序自动反求对应分位物理值并驱动界面渲染。
        """
        try:
            prefix, key = mode_str.split(":", 1)

            txt = val_str.strip().replace("%", "")
            try:
                # 统一不做智能判断，严格作为百分比数值处理
                p = float(txt) / 100.0
            except ValueError:
                return

            p = max(0.0, min(1.0, p))

            new_l, new_r = self.current_left, self.current_right
            theory_key = self._theory_label()
            use_theory = (key == theory_key and self._should_show_theory_stats())

            if use_theory:
                if prefix == 'pl':
                    new_l = self._safe_theory_ppf_scalar(p)
                elif prefix == 'pm':
                    tail = (1 - p) / 2.0
                    new_l = self._safe_theory_ppf_scalar(tail)
                    new_r = self._safe_theory_ppf_scalar(1.0 - tail)
                elif prefix == 'pr':
                    new_r = self._safe_theory_ppf_scalar(1.0 - p)
            else:
                s = self._get_series(key)
                if s is None or len(s) == 0:
                    return
                arr = s.values

                if prefix == 'pl':
                    new_l = float(np.percentile(arr, p * 100))
                elif prefix == 'pm':
                    tail = (1 - p) / 2
                    new_l = float(np.percentile(arr, tail * 100))
                    new_r = float(np.percentile(arr, (1 - tail) * 100))
                elif prefix == 'pr':
                    new_r = float(np.percentile(arr, (1 - p) * 100))

            if not np.isfinite(new_l) or not np.isfinite(new_r):
                return

            new_l = max(self.x_range_min, min(new_l, self.x_range_max))
            new_r = max(self.x_range_min, min(new_r, self.x_range_max))

            if new_l > new_r:
                new_l, new_r = new_r, new_l

            self.current_left = new_l
            self.current_right = new_r

            self.range_slider.setRangeValues(self.current_left, self.current_right)
            self._maybe_send_vlines(self.current_left, self.current_right)
            
            # ✅ 同步更新倒三角悬浮输入框的物理位置
            self._update_x_edits_pos()
            QTimer.singleShot(0, self._update_x_edits_pos)
            
            self.perform_buffered_update(force_full=True)

        except Exception as e:
            print(f"滑块输入确认异常: {e}")

    def _restore_focus_after_input(self, mode: str):
        """
        体验优化：强行扣留焦点。在回车更新后重新定位到文本框并全选，使用户可以轻松进行反复实验微调。
        """
        try:
            target = None
            if mode == "xl":
                target = self.edt_x_l
            elif mode == "xr":
                target = self.edt_x_r

            if target is not None:
                target.setFocus(Qt.FocusReason.OtherFocusReason)
                try:
                    target.selectAll()
                except Exception:
                    pass
        except Exception:
            pass

    # ---------------------------------------------------------
    # 3.10 UI 弹窗与附属功能 (Dialogs & Auxiliary Functions)
    # ---------------------------------------------------------

    def _on_analysis_obj_changed(self, sid, ck, lbl):
        """
        上下文切换中心。收到用户更改标的的指令后，清空叠加项并整体重载底层数据池及视图属性。
        """
        self.sim_id = sid
        self.current_key = ck
        if hasattr(self, 'overlay_items'):
            self.overlay_items = []

        # 数据重载前捕捉高级模式状态的快照，并在重载后还原
        mode_snapshot = ResultsAdvancedRouter.capture_mode_snapshot(self)

        self._current_view_data = "auto"
        self._set_chart_mode_from_view("auto")
        ResultsAdvancedRouter.apply_mode_snapshot(self, mode_snapshot)

        self.apply_chart_mode_ui()
        self.load_dataset(self.current_key)

        if mode_snapshot.is_tornado_active or mode_snapshot.is_scenario_active or mode_snapshot.is_boxplot_active:
            self._activate_analysis_view()

    def _open_analysis_obj_dialog(self):
        """
        召唤跨维选定器：让用户得以跨越不同的历史版本或不同的输出单元进行比对基准切换。
        """
        ResultsInteractionService.open_analysis_obj_dialog(self)

    def open_overlay_selector(self):
        """
        激活多轨对比弹窗，允许选中多个同量级数据一并在画布中展示相对关系。
        """
        ResultsInteractionService.open_overlay_selector(self)

    def _is_cdf_ci_settings_mode_active(self) -> bool:
        """检查当下图形形态，判断置信区间设置按钮是否该被禁用。"""
        is_main_view = self.view_stack.currentWidget() == getattr(self, "main_chart_wrapper", None)
        return bool(
            is_main_view
            and getattr(self, "_show_dkw_overlay", False)
            and str(getattr(self, "chart_mode", "hist") or "hist").lower() == "cdf"
        )

    def _open_cdf_ci_settings_dialog(self):
        """置信精度控制仪：单独为累积概率图提供的严格边界容错配置窗口。"""
        if not self._is_cdf_ci_settings_mode_active():
            return

        dlg = QDialog(self)
        dlg.setWindowTitle("Drisk - 置信区间设置")
        set_drisk_icon(dlg, "simu_icon.svg")
        dlg.setModal(True)
        dlg.setMinimumWidth(420)

        layout = QVBoxLayout(dlg)
        layout.setContentsMargins(14, 12, 14, 12)
        layout.setSpacing(8)

        field_w = 220
        label_w = 130
        form = QFormLayout()
        form.setLabelAlignment(Qt.AlignRight | Qt.AlignVCenter)
        form.setHorizontalSpacing(10)
        form.setVerticalSpacing(8)

        def _form_label(text: str) -> QLabel:
            label = QLabel(text)
            label.setFixedWidth(label_w)
            label.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
            return label

        combo_engine = QComboBox()
        combo_engine.setStyleSheet(drisk_combobox_qss())
        combo_engine.addItems(["快速 (Fast)", "标准 (Bootstrap)", "科学 (BCa)"])
        combo_engine.setFixedWidth(field_w)
        mode_to_idx = {"fast": 0, "bootstrap": 1, "bca": 2}
        combo_engine.setCurrentIndex(mode_to_idx.get(str(getattr(self, "ci_engine_mode", "fast")), 0))
        form.addRow(_form_label("置信区间计算引擎:"), combo_engine)

        edit_ci = QLineEdit(f"{float(getattr(self, 'cdf_ci_level', 0.95)):.2f}")
        edit_ci.setStyleSheet("font-weight: normal; padding: 4px 6px;")
        edit_ci.setFixedWidth(field_w)
        form.addRow(_form_label("置信度:"), edit_ci)
        layout.addLayout(form)

        btn_row = QHBoxLayout()
        btn_row.addStretch()
        btn_ok = QPushButton("确定")
        btn_cancel = QPushButton("取消")
        btn_row.addWidget(btn_ok)
        btn_row.addWidget(btn_cancel)
        layout.addLayout(btn_row)

        def _apply_and_accept():
            text = str(edit_ci.text() or "").strip()
            try:
                ci_value = float(text)
            except Exception:
                QMessageBox.warning(dlg, "无效输入", "置信度必须是在 0.01 到 0.99 之间的数字。")
                return
            if not (0.01 <= ci_value <= 0.99):
                QMessageBox.warning(dlg, "无效范围", "置信度范围必须在 [0.01, 0.99] 内。")
                return
            ci_value = round(ci_value, 2)

            mode_idx = int(combo_engine.currentIndex())
            mode_idx = min(2, max(0, mode_idx))
            modes = ["fast", "bootstrap", "bca"]
            new_mode = modes[mode_idx]
            old_mode = str(getattr(self, "ci_engine_mode", "fast"))
            old_ci = float(getattr(self, "cdf_ci_level", 0.95))
            changed = False
            if new_mode != old_mode:
                self.ci_engine_mode = new_mode
                if hasattr(self, "combo_ci_engine") and self.combo_ci_engine.currentIndex() != mode_idx:
                    with QSignalBlocker(self.combo_ci_engine):
                        self.combo_ci_engine.setCurrentIndex(mode_idx)
                changed = True
            if not math.isclose(ci_value, old_ci, rel_tol=0.0, abs_tol=1e-12):
                self.cdf_ci_level = ci_value
                changed = True
            if changed:
                self.init_chart_via_tempfile()
                self._invalidate_stats_cache()
                self.update_stats_ui(update_stats_table=True)
            dlg.accept()

        btn_ok.clicked.connect(_apply_and_accept)
        btn_cancel.clicked.connect(dlg.reject)
        dlg.exec()

    def _resolve_label_settings_y_title(self) -> str:
        """根据当前不同的物理数学模式动态分配合理的纵轴文字标签。"""
        mode = str(getattr(self, "chart_mode", "hist") or "hist").lower()
        if mode == "cdf":
            return "累积概率"
        if mode == "discrete":
            return "相对频率 (%)"
        if mode == "rel_freq":
            return "相对频率 (%)"
        return "密度"

    def _resolve_label_settings_axis_numeric_defaults(self) -> Dict[str, bool]:
        """
        探测坐标轴格式：由于箱线等摘要图使用的是分类标签横轴，此处拦截并宣告其非数值性质。
        """
        mode = str(ResultsRuntimeStateHelper.get_current_analysis_mode(self, "") or "").lower()
        summary_modes = {"boxplot", "letter_value", "violin", "trend"}
        if mode in summary_modes:
            return {"x": False, "y": True}
        return {"x": True, "y": True}

    def _resolve_chart_title_font_from_config(self) -> tuple[str, float]:
        """解析配置，获取图表主标题的字体族与字号设置。"""
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

    def _apply_chart_title_label_style(self) -> None:
        """将文本设置应用到当前界面的图表主标题组件上以保持高度统一。"""
        family, size = self._resolve_chart_title_font_from_config()
        safe_family = family.replace("\\", "\\\\").replace('"', '\\"')
        self.chart_title_label.setStyleSheet(
            "color: #333333; "
            "margin: 0px; "
            "padding: 0px; "
            f"font-size: {size:.1f}px; "
            "font-weight: normal; "
            f'font-family: "{safe_family}", "Microsoft YaHei", sans-serif;'
        )

    def _build_label_settings_context(self) -> Dict[str, Any]:
        """预处理并将所有的画布细节构建打包，推给“文本设置”模块使用"""
        axis_numeric_defaults = self._resolve_label_settings_axis_numeric_defaults()
        x_is_numeric = bool(axis_numeric_defaults.get("x", True))
        y_is_numeric = bool(axis_numeric_defaults.get("y", True))
        if x_is_numeric:
            x_mag_recommended = infer_si_mag(
                value_range=(float(getattr(self, "x_range_min", 0.0)), float(getattr(self, "x_range_max", 1.0))),
                dtick=float(getattr(self, "x_dtick", 1.0) or 1.0),
                forced_mag=None,
            )
        else:
            x_mag_recommended = 0
        label_widget = getattr(self, "chart_title_label", None)
        widget_text = label_widget.text() if label_widget is not None else ""
        current_label = str(
            getattr(self, "_base_chart_title", "")
            or widget_text
            or getattr(self, "current_key", "")
            or ""
        )
        return {
            "chart_title": {
                "default_text": current_label,
                "default_font_family": "Arial",
                "default_font_size": 14,
            },
            "axes": {
                "x": {
                    "default_title": DriskChartFactory.VALUE_AXIS_TITLE if x_is_numeric else "分类",
                    "default_title_font_family": "Arial",
                    "default_title_font_size": 12,
                    "default_tick_font_family": "Arial",
                    "default_tick_font_size": 12,
                    "recommended_mag": int(x_mag_recommended),
                    "data_candidates": [
                        {"key": "x_main", "label": "当前 X 数据", "is_numeric": x_is_numeric},
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
                        {"key": "y_main", "label": "当前 Y 数据", "is_numeric": y_is_numeric},
                    ],
                },
            },
        }

    def _refresh_after_label_settings(self) -> None:
        """应用更改后依据所处工作面板的位置发起精确制导的重绘"""
        mode = ResultsRuntimeStateHelper.get_current_analysis_mode(self, "")
        if mode == "tornado":
            self._load_and_render_tornado()
            return
        if mode in ["boxplot", "letter_value", "violin", "trend", "scenario"]:
            self._activate_analysis_view()
            return
        self.init_chart_via_tempfile()
        self._update_x_edits_pos()

    def _open_label_settings_dialog(self):
        """核心交互：打开并接管包含图名、轴题在内的完整富文本编辑器对话框。"""
        try:
            context = self._build_label_settings_context()
            dlg = LabelSettingsDialog(
                config=getattr(self, "_label_settings_config", None),
                context=context,
                parent=self,
            )
            if dlg.exec():
                new_cfg = dlg.get_config()
                base_ck = self.label_to_cell_key.get(self.current_key, self.current_key)
                metadata_title = self._resolve_chart_title_from_metadata(base_ck)
                metadata_title_text = str(metadata_title or "").strip()
                chart_cfg = new_cfg.get("chart_title", {}) if isinstance(new_cfg.get("chart_title", {}), dict) else {}
                chart_title_delta = chart_cfg.get("text_override", None) if isinstance(chart_cfg, dict) else None
                title_changed = False
                if chart_title_delta is not None:
                    desired_title = str(chart_title_delta or "").strip()
                    if (not desired_title) or desired_title == metadata_title_text:
                        self._set_chart_title_display_override(base_ck, None)
                    else:
                        self._set_chart_title_display_override(base_ck, desired_title)
                    title_changed = True
                if isinstance(chart_cfg, dict):
                    chart_cfg["text_override"] = None
                if title_changed:
                    shown_title = self._resolve_display_chart_title(base_ck, metadata_title_text)
                    self.chart_title_label.setText(shown_title)
                    self._base_chart_title = shown_title
                self._label_settings_config = new_cfg
                self._apply_chart_title_label_style()
                self._label_axis_numeric = get_axis_numeric_flags(
                    new_cfg,
                    fallback=self._resolve_label_settings_axis_numeric_defaults(),
                )
                self.manual_mag = get_axis_display_unit_override(new_cfg, "x")
                self._refresh_after_label_settings()
        except Exception as exc:
            QMessageBox.critical(self, "错误", f"打开文本设置失败：{exc}")

    def on_export_clicked(self):
        """
        数据变现口：捕捉画布中所有的细节状态直接生成报表图文导出方案。
        """
        ResultsInteractionService.on_export_clicked(self)

    def _update_bottom_bar_visibility(self):
        """
        场景自适应系统：探测高级功能是否正在活跃，并据此折叠或铺开相应的操作按钮栏。
        """
        ResultsInteractionService.update_bottom_bar_visibility(self)

    def _on_global_mag_changed(self, mag_int):
        """
        量级倍增器的回调接口，为未来可能支持的全局 K/M/B 数据缩放准备。
        """
        ResultsInteractionService.on_global_mag_changed(self, mag_int)

# =======================================================
# 4. 外部调用接口 (External Public APIs)
# =======================================================
def show_results_dialog(
    data=None,
    *,
    sim_id: int | None = None,
    cell_keys: list[str] | None = None,
    labels: dict[str, str] | None = None,
    kind: str = "output",
):
    """
    一键弹出模拟结果分析窗口（UltimateRiskDialog）的外部唯一官方入口。
    强制以【模态】(Modal) 方式运行，阻塞用户其他 Excel 操作直到窗口关闭。

    支持两种数据挂载架构:
    1. 新版后端驱动架构（推荐）: 
       仅传递 `sim_id` 和 `cell_keys`，由弹窗内部通过 backend_bridge 按需向底层缓存索要数据。
       极大地降低了大数据量下的内存拷贝开销。
    2. 旧版直传数据架构（兼容）: 
       直接通过 `data` 参数传递 Pandas DataFrame 或 dict 数据集。

    参数:
      - data: [旧版参数] 直接传递的数据集对象 (dict 或 array)。
      - sim_id: [新版参数] 目标模拟任务的唯一标识 ID。
      - cell_keys: [新版参数] 关注的 Excel 单元格绝对地址列表 (如 ["Sheet1!$A$1"])。
      - labels: [新版参数] 单元格地址到自定义 UI 展示名称的映射字典。
      - kind: [新版参数] 数据类型标志，通常为 "output" (输出结果) 或 "input" (输入变量)。
    """
    
    # ---------------------------------------------------------
    # Step 1. 确保 Qt 应用程序实例存在 (兼容无头或独立运行环境)
    # ---------------------------------------------------------
    app = QApplication.instance()
    if app is None:
        app = QApplication(sys.argv)

    # ---------------------------------------------------------
    # Step 2. 路由分发：基于新版后端驱动模式 (Backend-Driven)
    # ---------------------------------------------------------
    if sim_id is not None and cell_keys is not None:
        # 校验桥接层是否就绪
        if bridge is None:
            QMessageBox.critical(None, "Drisk", "backend_bridge 适配模块不可用，无法从后端缓存读取数据")
            return
            
        # 2.1 单元格地址标准化 (抹平大小写、相对/绝对引用的差异)
        norm_keys = []
        for ck in (cell_keys or []):
            try:
                norm_keys.append(bridge.normalize_cell_key(ck))
            except Exception:
                norm_keys.append(ck)

        # 2.2 核心防御层：底层无效数据拦截机制 (Surgical Fix 1)
        # 目的：防止模型中存在底层不支持的特殊公式/数组，导致返回全空的 #VALUE! 从而引发前端绘图崩溃
        has_valid_data = False
        for ck in norm_keys:
            # 试探性提取数据并进行脏数据清洗
            arr = bridge.get_series(int(sim_id), ck, kind=kind or "output")
            clean_arr = ResultsDataService.clean_series(arr)
            # 只要有任意一个单元格包含有效数值，即可放行
            if len(clean_arr) > 0:
                has_valid_data = True
                break
                
        # 若全部数据均被判定为无效，立即阻断弹窗并向用户发出警告
        if not has_valid_data:
            QMessageBox.warning(None, "数据警告", "未能提取到有效的模拟数据！\n这通常是因为模型中包含了底层暂不支持的数组/区域参数（如 Cumul/Discrete 的引用），导致整个模拟过程中断。")
            return

        # 2.3 封装为轻量级数据源对象，实例化主窗口
        bi = BackendResultsInput(
            sim_id=int(sim_id),
            cell_keys=norm_keys,
            labels=labels or {},
            kind=kind or "output",
        )
        dlg = UltimateRiskDialog(bi)
        
    # ---------------------------------------------------------
    # Step 3. 路由分发：基于旧版直传数据模式 (Legacy Direct Data)
    # ---------------------------------------------------------
    else:
        # 3.1 核心防御层：旧版数据的有效性校验 (Surgical Fix 2)
        has_valid_data = False
        if isinstance(data, dict):
            # 如果是字典源，检查任意 Value 中是否存在有效数据
            has_valid_data = any(len(ResultsDataService.clean_series(v)) > 0 for v in data.values())
        elif data is not None:
            # 如果是单一数组源，直接检查清洗后的长度
            has_valid_data = len(ResultsDataService.clean_series(data)) > 0
            
        # 同样阻断无效数据的加载
        if not has_valid_data:
            QMessageBox.warning(None, "数据警告", "未能提取到有效的模拟数据！\n这通常是因为模型中包含了底层暂不支持的数组/区域参数（如 Cumul/Discrete 的引用），导致整个模拟过程中断。")
            return

        # 3.2 直接将数据对象送入主窗口进行实例化
        dlg = UltimateRiskDialog(data)

    # ---------------------------------------------------------
    # Step 4. 以模态方式启动窗口并阻塞当前线程
    # ---------------------------------------------------------
    dlg.exec()