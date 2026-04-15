# ui_shared.py
"""
本模块提供跨模块复用的 UI 组件、图表样式辅助工具以及数学计算方法。

主要功能分组：
1. 缓存与共享视图管理 (Caching & Shared View Management)：管理 WebEngineView 的全局复用与 Plotly 的本地缓存目录。
2. 全局样式与 QSS 构建 (Global Styles & QSS Builders)：定义下拉框、滚动条、数值调节钮等组件的统一样式与颜色序列。
3. 基础显示组件 (Basic Display Widgets)：如量级标签 (MagnitudeSelector) 与加载骨架屏 (ChartSkeleton)。
4. 数学计算与交互辅助 (Math & Interaction Utilities)：处理坐标轴步长、滑块网格吸附与 SI 量级推导。
5. 自定义输入与范围控制组件 (Custom Input & Range Slider Components)：提供带单位后缀的智能输入框及多层概率滑块。
6. 图表配置与单位转换器 (Chart Configuration & Unit Formatters)：统一管理 Plotly 图表的轴样式生成。
7. 渲染调度与视图同步器 (Render Dispatching & View Synchronization)：处理 Qt 遮罩层与 Web 视图的尺寸与重绘同步。
8. Excel 交互与图标管理工具 (Excel Interaction & Icon Tools)：解析 Excel 单元格命名引用、加载图标等辅助方法。
"""
import math
import os
import re
import tempfile
import numpy as np

from PySide6.QtGui import QIcon, QPixmap, QPainter, QColor, QPen, QBrush, QPolygonF, QFontMetrics
from PySide6.QtCore import Qt, QTimer, Signal, QRectF, QPointF, QObject, QEvent, QEventLoop, QUrl, QSize
from PySide6.QtWidgets import QLineEdit, QWidget, QLabel, QPushButton, QMenu, QComboBox, QHBoxLayout, QVBoxLayout, QSizePolicy
from PySide6.QtWebEngineWidgets import QWebEngineView

# =================================================================
# 第 1 组：缓存与共享视图管理 (Caching & Shared View Management)
# =================================================================

# 全局 WebEngineView 缓存池与停放区
_GLOBAL_WEBVIEW_CACHE = []
_WEBVIEW_PARKING_LOT = None

# Plotly 本地缓存目录（进程级别）
_PLOTLY_CACHE_DIR = os.path.join(
    tempfile.gettempdir(),
    f"drisk_plotly_cache_{os.getpid()}",
)
try:
    os.makedirs(_PLOTLY_CACHE_DIR, exist_ok=True)
except Exception:
    pass

def get_plotly_cache_dir() -> str:
    """返回进程级别的 Plotly 缓存目录路径。"""
    return _PLOTLY_CACHE_DIR

def _get_parking_lot():
    """获取或创建一个隐藏的停放区组件 (Parking Lot)，用于暂存后台 WebView。"""
    global _WEBVIEW_PARKING_LOT
    if _WEBVIEW_PARKING_LOT is None:
        _WEBVIEW_PARKING_LOT = QWidget()
        _WEBVIEW_PARKING_LOT.setAttribute(Qt.WA_QuitOnClose, False)
        _WEBVIEW_PARKING_LOT.setWindowTitle("Drisk 缓存停放区")
        _WEBVIEW_PARKING_LOT.hide()
    return _WEBVIEW_PARKING_LOT

def get_shared_webview(parent_widget=None):
    """创建一个新的 WebEngineView 用于图表渲染，背景设为透明与白色。"""
    view = QWebEngineView()
    view.setStyleSheet("background: transparent;")
    view.page().setBackgroundColor(Qt.white)
    if parent_widget:
        view.setParent(parent_widget)
    view.show()
    return view

def recycle_shared_webview(view):
    """安全地销毁与回收 WebView 实例，释放相关资源。"""
    if view is None:
        return
    try:
        view.stop()
        try:
            view.page().setWebChannel(None)
        except Exception:
            pass
        view.setParent(None)
        view.deleteLater()
    except Exception as e:
        print(f"Destroy WebView error: {e}")


# =================================================================
# 第 2 组：全局样式与 QSS 构建 (Global Styles & QSS Builders)
# =================================================================

# 全局图表循环配色表 (常规与理论对比)
# 重要：前两个颜色最好是红蓝经典对比色 // 两组色板仅前两颜色调换顺序，用于区别输入输出数据模拟绘图
DRISK_COLOR_CYCLE = [
    "#C65D5D", "#4F89A8", "#57874C", "#8361A7", "#C59544",
    "#5FA6A1", "#8B6E5E", "#7B8089", "#B27C95", "#4E6A7B",
]
DRISK_COLOR_CYCLE_THEO = [
    "#4F89A8", "#C65D5D", "#57874C", "#8361A7", "#C59544",
    "#5FA6A1", "#8B6E5E", "#7B8089", "#B27C95", "#4E6A7B",
]

# 全局弹窗按钮标准 QSS
DRISK_DIALOG_BTN_QSS = """
    QPushButton#btnOk, QPushButton#btnCancel { 
        min-width: 80px; 
        max-width: 80px; 
        height: 28px; 
        font-weight: bold; 
        font-size: 12px;
        text-align: center; /* 全局强制居中 */
        padding: 0px;       /* 全局清除偏移 */
    }
    QPushButton#btnOk { 
        background-color: #0050b3; color: white; border: none; border-radius: 4px; 
    }
    QPushButton#btnCancel { 
        background-color: white; border: 1px solid #d9d9d9; color: #555; border-radius: 4px; 
    }
    QPushButton#btnCancel:hover { background-color: #f5f5f5; border-color: #40a9ff; color: #40a9ff; }
    QPushButton#btnCancel:pressed { background-color: #e6e6e6; }
"""

def drisk_scrollbar_qss(scope: str = "") -> str:
    """
    构建精美的滚动条样式，支持可选的作用域选择器。
    示例作用域:
    - ""  -> 作用于当前样式表作用域内的所有滚动条
    - "QComboBox QAbstractItemView" -> 仅作用于下拉框的弹出列表
    """
    prefix = f"{scope.strip()} " if str(scope).strip() else ""
    return f"""
    {prefix}QScrollBar:horizontal {{
        border: none;
        background: #f5f5f5;
        height: 12px;
        margin: 0px;
        border-radius: 6px;
    }}
    {prefix}QScrollBar::handle:horizontal {{
        background: #bcbcbc;
        min-width: 30px;
        border-radius: 4px;
        margin: 2px;
    }}
    {prefix}QScrollBar::handle:horizontal:hover {{ background: #999999; }}
    {prefix}QScrollBar::handle:horizontal:pressed {{ background: #777777; }}
    {prefix}QScrollBar::add-line:horizontal, {prefix}QScrollBar::sub-line:horizontal {{ width: 0px; }}
    {prefix}QScrollBar::add-page:horizontal, {prefix}QScrollBar::sub-page:horizontal {{ background: none; }}

    {prefix}QScrollBar:vertical {{
        border: none;
        background: #f5f5f5;
        width: 12px;
        margin: 0px;
        border-radius: 6px;
    }}
    {prefix}QScrollBar::handle:vertical {{
        background: #bcbcbc;
        min-height: 30px;
        border-radius: 4px;
        margin: 2px;
    }}
    {prefix}QScrollBar::handle:vertical:hover {{ background: #999999; }}
    {prefix}QScrollBar::handle:vertical:pressed {{ background: #777777; }}
    {prefix}QScrollBar::add-line:vertical, {prefix}QScrollBar::sub-line:vertical {{ height: 0px; }}
    {prefix}QScrollBar::add-page:vertical, {prefix}QScrollBar::sub-page:vertical {{ background: none; }}
    """

def drisk_combobox_qss() -> str:
    """
    全局统一的下拉框样式 (加载本地特定的实心倒三角 SVG)
    """
    # 动态获取当前 ui_shared.py 所在的绝对路径
    base_dir = os.path.dirname(os.path.abspath(__file__))
    icon_path = os.path.join(base_dir, "icons", "arrow_down.svg").replace('\\', '/')

    base_qss = f"""
    QComboBox {{
        border: 1px solid #d9d9d9;
        border-radius: 3px;
        /* 核心修改 1：右侧内边距从 24px 减小到 16px，让文字紧贴下拉区 */
        padding: 0px 12px 0px 8px; 
        background-color: #ffffff;
        color: #333333;
        font-size: 12px;
        font-family: "Arial", "Microsoft YaHei";
        min-height: 20px;
        max-height: 20px;
    }}
    QComboBox:hover, QComboBox:focus {{
        border: 1px solid #40a9ff;
    }}
    
    QComboBox::drop-down {{
        subcontrol-origin: padding;
        subcontrol-position: top right;
        /* 核心修改 2：去除之前宽大的占位，将宽度极限压缩到 16px */
        width: 12px; 
        border: none;
        background-color: transparent;
    }}
    
    QComboBox::down-arrow {{
        image: url('{icon_path}');
        width: 10px;  /* 保持精巧的 10x10 */
        height: 10px;
        /* 核心修改 3：微调 margin，使箭头在紧凑的 16px 空间内完美居中 */
        margin-right: 4px; 
    }}
    
    QComboBox::down-arrow:on {{
        top: 1px;
    }}
    
    QComboBox QAbstractItemView {{
        border: 1px solid #d9d9d9;
        border-radius: 3px;
        background-color: #ffffff;
        selection-background-color: #f0f5ff;
        selection-color: #333333;
        outline: none;
    }}
    """
    # 在下拉框的弹出列表中复用经过打磨的滚动条样式
    return base_qss + drisk_scrollbar_qss("QComboBox QAbstractItemView")

def drisk_spinbox_qss() -> str:
    """
    全局统一的 SpinBox 高级样式 (修复上下按钮无法点击的 bug，采用扁平化设计 + 本地 SVG 图标)
    """
    # 动态获取当前目录并拼接图标路径，确保使用正斜杠 (Qt 的 url 要求)
    base_dir = os.path.dirname(os.path.abspath(__file__))
    up_icon = os.path.join(base_dir, "icons", "arrow_increase.svg").replace('\\', '/')
    down_icon = os.path.join(base_dir, "icons", "arrow_decrease.svg").replace('\\', '/')

    # 注意外层使用了 f-string，所以 CSS 原本的大括号需要用双大括号 {{ }} 转义
    return f"""
    QSpinBox, QDoubleSpinBox {{
        border: 1px solid #d9d9d9;
        border-radius: 3px;
        padding: 2px 22px 2px 6px; 
        background: #ffffff;
        selection-background-color: #1890ff;
        min-height: 20px;
    }}
    QSpinBox:hover, QDoubleSpinBox:hover {{ border-color: #40a9ff; }}
    QSpinBox:focus, QDoubleSpinBox:focus {{ border-color: #1890ff; }}
    
    /* 上下按钮的背景和边框区域 */
    QSpinBox::up-button, QDoubleSpinBox::up-button {{
        subcontrol-origin: border;
        subcontrol-position: top right;
        width: 20px;
        border-left: 1px solid #d9d9d9;
        border-bottom: 1px solid #d9d9d9;
        border-top-right-radius: 3px;
        background: #fafafa;
    }}
    
    QSpinBox::down-button, QDoubleSpinBox::down-button {{
        subcontrol-origin: border;
        subcontrol-position: bottom right;
        width: 20px;
        border-left: 1px solid #d9d9d9;
        border-bottom-right-radius: 3px;
        background: #fafafa;
    }}

    /* ========================================
       核心修改：应用本地 SVG 图标
       ======================================== */
    QSpinBox::up-arrow, QDoubleSpinBox::up-arrow {{
        image: url('{up_icon}');
        width: 10px;
        height: 10px;
    }}
    
    QSpinBox::down-arrow, QDoubleSpinBox::down-arrow {{
        image: url('{down_icon}');
        width: 10px;
        height: 10px;
    }}
    
    /* 按钮禁用或不可点击时的图标透明度 (可选) */
    QSpinBox::up-arrow:off, QSpinBox::down-arrow:off, 
    QDoubleSpinBox::up-arrow:off, QDoubleSpinBox::down-arrow:off {{
        opacity: 0.3; 
    }}
    /* ======================================== */
    
    /* 按钮悬停和按下时的颜色反馈 */
    QSpinBox::up-button:hover, QDoubleSpinBox::up-button:hover,
    QSpinBox::down-button:hover, QDoubleSpinBox::down-button:hover {{
        background: #e6f7ff;
    }}
    
    QSpinBox::up-button:pressed, QDoubleSpinBox::up-button:pressed,
    QSpinBox::down-button:pressed, QDoubleSpinBox::down-button:pressed {{
        background: #bae0ff;
    }}
    """

# =================================================================
# 第 3 组：基础显示组件 (Basic Display Widgets)
# =================================================================

class MagnitudeSelector(QLabel):
    """展示量级的非交互式标签 -- 仅显示数据量级"""
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.current_mag = 0
        # 样式改为纯文本展示，移除 hover 效果
        self.setStyleSheet("""
            QLabel {
                background: transparent; color: #333333;
                font-weight: normal; font-family: 'Arial';
                font-size: 12px; border: none; padding: 0px;
            }
        """)
        self.setAlignment(Qt.AlignLeft | Qt.AlignVCenter) # 保持左对齐
        
    def set_magnitude(self, mag: int):
        """更新当前量级并修改显示文字"""
        self.current_mag = mag
        # 从 SmartFormatter 获取符号，例如 6 -> 'M'
        suffix = SmartFormatter.SI_MAP.get(mag, "")
        self.setText(suffix if suffix else "")

        # 如果没有后缀（量级为1），则隐藏标签以节省空间
        if not suffix:
            self.hide()
        else:
            self.show()

class ChartSkeleton(QWidget):
    """
    绘图区骨架屏 (ChartSkeleton) - 全局遮罩版。
    提供图表加载过程中的白板遮罩与状态提示。
    """
    def __init__(self, parent=None, *, default_margins=None):
        super().__init__(parent)
        self.setAutoFillBackground(True)
        # 强制使用样式背景，使骨架屏遮罩始终被绘制
        self.setAttribute(Qt.WA_StyledBackground, True)
        
        # 保持纯白色遮罩以提供确定的渲染反馈
        self.setStyleSheet("background-color: rgba(255, 255, 255, 1);")

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        
        self.lbl = QLabel("正在渲染图表...", self)
        self.lbl.setAlignment(Qt.AlignCenter)
        self.lbl.setStyleSheet("""
            color: #555555; 
            font-family: "Microsoft YaHei"; 
            font-size: 16px; 
            font-weight: norm;
            background: transparent;
        """)
        layout.addWidget(self.lbl, alignment=Qt.AlignCenter)

    def set_margins(self, margins):
        pass


# =================================================================
# 第 4 组：数学计算与交互辅助 (Math & Interaction Utilities)
# =================================================================

class DriskMath:
    """数学与图表刻度计算工具"""
    
    @staticmethod
    def calc_smart_step(x_span: float) -> float:
        """1-2-5 智能步长算法 (用于坐标轴刻度计算)"""
        if x_span <= 0:
            return 1.0
        exponent = math.floor(math.log10(x_span))
        base = x_span / (10 ** exponent)

        # 左闭右开区间判定
        if base < 2.0:
            dt = 0.2 
        elif base < 5.0:
            dt = 0.5
        else: # 5.0 <= base < 10.0
            dt = 1.0

        step = dt * (10 ** exponent)
        return float(step)

    @staticmethod
    def ceil_to_tick(value: float, tick: float) -> float:
        """将 value 向上对齐到 tick 的整数倍"""
        if tick <= 0:
            return value
        return math.ceil(value / tick) * tick

    @staticmethod
    def floor_to_tick(value: float, tick: float) -> float:
        """将 value 向下对齐到 tick 的整数倍"""
        if tick <= 0:
            return value
        return math.floor(value / tick) * tick

class SnapUtils:
    """处理拖动吸附的核心数学逻辑 (用于滑块交互)"""
    
    @staticmethod
    def calc_grid_step(xmin: float, xmax: float, plot_width_px: int,
                       px_per_step: int = 2, min_steps: int = 200, max_steps: int = 4000) -> float:
        """根据绘图区像素宽度，计算连续分布的吸附步长"""
        if xmax <= xmin or plot_width_px <= 0:
            return 0.0
        steps = int(plot_width_px / max(1, px_per_step))
        steps = max(min_steps, min(max_steps, steps))
        return (xmax - xmin) / float(steps)

    @staticmethod
    def snap_to_grid(val: float, xmin: float, step: float) -> float:
        """连续网格吸附"""
        if step <= 0:
            return val
        idx = int(round((val - xmin) / step))
        return xmin + idx * step

    @staticmethod
    def snap_to_discrete(val: float, points: np.ndarray, method='nearest', 
                         grid_step: float = None, xmin: float = None, xmax: float = None) -> float:
        """
        离散点吸附 (增强版：混合磁性吸附)
        加入 grid_step, xmin, xmax 可以让滑块同时吸附到刻度网格上，解决数据点过少导致的滑动受限。
        """
        # 1. 找最近的实际数据点
        closest_point = None
        dist_point = float('inf')
        
        if points is not None and len(points) > 0:
            if method == 'nearest':
                i = int(np.searchsorted(points, val))
                if i <= 0: 
                    closest_point = float(points[0])
                elif i >= len(points): 
                    closest_point = float(points[-1])
                else:
                    left = float(points[i - 1])
                    right = float(points[i])
                    closest_point = left if abs(val - left) <= abs(val - right) else right
            elif method == 'floor':
                i = int(np.searchsorted(points, val, side='right')) - 1
                if i < 0: i = 0
                closest_point = float(points[i])
            
            if closest_point is not None:
                dist_point = abs(val - closest_point)
                
        # 2. 如果没有提供网格参数，直接返回传统数据点吸附结果
        if grid_step is None or grid_step <= 0 or xmin is None or xmax is None:
            return closest_point if closest_point is not None else round(val)
            
        # 3. 找最近的网格刻度点 (将步长分为 2 份，允许 0.5 级别的半点吸附)
        sub_step = grid_step / 2.0
        closest_grid = round((val - xmin) / sub_step) * sub_step + xmin
        closest_grid = max(xmin, min(closest_grid, xmax))
        dist_grid = abs(val - closest_grid)
        
        # 4. 决出胜负：哪个离鼠标当前位置更近。给予实际数据点 1e-5 的微小优先权防冲突
        if closest_point is not None and dist_point <= dist_grid + 1e-5:
            return closest_point
        else:
            return closest_grid

class SmartFormatter:
    """智能量级格式化工具：计算从轴步长映射出的 SI 后缀及小数位精度"""

    SI_MAP = {
        12: "T", 9: "G", 6: "M", 3: "k", 0: "",
        -3: "m", -6: "u", -9: "n", -12: "p", -15: "f", -18: "a"
    }

    @staticmethod
    def get_format_config(dtick: float):
        """
        返回 (divisor, suffix, decimals)
        规则：始终显示为 dtick 小两个量级的形式 (decimals = 2 - log10(norm_step))
        """
        if dtick <= 0: return 1.0, "", 2

        # 1. 确定量级 (base magnitude)
        mag = math.floor(math.log10(dtick))
        # 归一化到最近的 3 的倍数 (SI单位)
        si_mag = (mag // 3) * 3

        suffix = SmartFormatter.SI_MAP.get(si_mag, "")
        divisor = 10 ** si_mag

        # 2. 计算归一化后的步长 (例如 20M -> 20)
        norm_step = dtick / divisor

        # 3. 计算建议小数位：目标是分辨出步长的 1/100
        # log10(20) ≈ 1.3; 2 - 1.3 = 0.7 -> 1位小数
        # log10(1) = 0;    2 - 0 = 2   -> 2位小数
        # log10(0.1) = -1; 2 - (-1) = 3 -> 3位小数
        needed_precision = 2 - math.log10(norm_step)
        decimals = max(0, math.ceil(needed_precision))

        return divisor, suffix, decimals


# =================================================================
# 第 5 组：自定义输入与范围控制组件 (Custom Input & Range Slider)
# =================================================================

class AutoSelectLineEdit(QLineEdit):
    """获得焦点或点击时自动全选内容的输入框"""
    def focusInEvent(self, event):
        super().focusInEvent(event)
        QTimer.singleShot(0, self.selectAll)

    def mousePressEvent(self, event):
        # 当点击输入框时，强制将当前 Qt 窗口提升为系统激活窗口，切断 Excel 的焦点劫持
        window = self.window()
        if window:
            window.activateWindow()
            
        super().mousePressEvent(event)
        if not self.hasFocus():
            QTimer.singleShot(0, self.selectAll)

    def keyPressEvent(self, event):
        if event.key() in (Qt.Key_Return, Qt.Key_Enter):
            self.clearFocus()
            self.returnPressed.emit()
            event.accept()
        else:
            super().keyPressEvent(event)

class SuffixAutoSelectLineEdit(AutoSelectLineEdit):
    """带后缀保护机制的自动全选输入框"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.suffix = ""

    def set_suffix(self, s):
        self.suffix = s

    def focusInEvent(self, event):
        # 用户编辑时去除后缀，以便输入纯数字
        txt = self.text()
        if self.suffix and txt.endswith(self.suffix):
            clean_txt = txt[:-len(self.suffix)]
            self.setText(clean_txt)
        super().focusInEvent(event)

    def focusOutEvent(self, event):
        # 失去焦点时恢复后缀
        txt = self.text()
        if self.suffix and not txt.endswith(self.suffix) and txt:
             self.setText(txt + self.suffix)
        super().focusOutEvent(event)

class FloatingValueWithMag(QWidget):
    """包含编辑器与隐藏量级管理器的复合组件"""
    def __init__(self, parent: QWidget, *, height: int = 24, opacity: float = 0.8):
        super().__init__(parent)

        self.edit = SuffixAutoSelectLineEdit(self)
        self.edit.setFixedHeight(int(height))
        self.edit.setAlignment(Qt.AlignCenter)
        self.edit.setStyleSheet(floating_value_edit_qss(opacity))

        self.mag = MagnitudeSelector(self)
        self.mag.hide() 

        lay = QHBoxLayout(self)
        lay.setContentsMargins(0, 0, 0, 0)
        lay.setSpacing(0)
        lay.addWidget(self.edit)
        self.setFixedHeight(int(height))
        self.hide()
        self.raise_()

    def set_magnitude(self, mag: int):
        self.mag.set_magnitude(mag)
        self.mag.hide()
        
        suffix = SmartFormatter.SI_MAP.get(mag, "")
        self.edit.set_suffix(suffix)
        
    def set_edit_width(self, w: int):
        self.edit.setFixedWidth(int(w))
        self.setFixedWidth(int(w))
        
    def ensure_visible(self):
        if not self.isVisible():
            self.show()

def floating_value_edit_qss(opacity: float = 0.8) -> str:
    """悬浮输入框透明闲置状态样式"""
    return (
        "QLineEdit {"
        "  background: transparent;"
        "  border: 1px solid transparent;"
        "  border-radius: 2px;"
        "  color: #333333;"
        "  font-weight: normal;"
        "  font-size: 12px;"
        "  font-family: Arial;"
        "  padding-left: 0px;"
        "  padding-right: 0px;"
        "}"
        "QLineEdit:focus {"
        "  background: rgba(255,255,255, 0.9);"
        "  border: 1px solid #40a9ff;"
        "}"
    )

def create_floating_value_with_mag(parent: QWidget, *, height: int = 24, opacity: float = 0.8) -> FloatingValueWithMag:
    return FloatingValueWithMag(parent, height=height, opacity=opacity)

def update_floating_value_edits_pos(*, slider: QWidget,
                                    group_l,
                                    group_r,
                                    val_l: float,
                                    val_r: float,
                                    range_min: float,
                                    range_max: float,
                                    margin_l: int,
                                    margin_r: int,
                                    dtick: float = 1.0,
                                    forced_mag: int = None,
                                    gap: int = 1,
                                    clamp: bool = True,
                                    single_handle: bool = False):
    """悬浮输入框 -- 滑块正上方坐标输入的布局与格式化更新方法"""
    if slider is None:
        return

    if hasattr(slider, "margin_left") and hasattr(slider, "margin_right"):
        margin_l = int(slider.margin_left)
        margin_r = int(slider.margin_right)

    # 获取量级配置
    std_mag, suffix, divisor, _hint = get_si_unit_config(
        value_range=(range_min, range_max),
        dtick=dtick,
        force_m=forced_mag
    )
    forced_mag = std_mag

    norm_step = (dtick / divisor) if divisor else 0.0
    decimals = max(0, math.ceil(2 - math.log10(norm_step))) if norm_step > 0 else 2
    fmt_str = f"{{:.{decimals}f}}"

    if hasattr(slider, 'get_valid_width'):
        track_w = max(1, slider.get_valid_width())
    else:
        track_w = max(1, slider.width() - margin_l - margin_r)

    def _val_to_center_x(v: float) -> float:
        if range_max <= range_min:
            t = 0.0
        else:
            t = (v - range_min) / (range_max - range_min)
        t = max(0.0, min(1.0, t))
        return margin_l + t * track_w

    cx_l = _val_to_center_x(val_l)
    cx_r = _val_to_center_x(val_r)

    base_x = slider.x()
    base_y = slider.y()

    def _update_and_place_group(group, v: float, cx: float):
        if group is None:
            return

        # 设置量级 (内部会更新 edit.suffix)
        if hasattr(group, "set_magnitude"):
            group.set_magnitude(forced_mag)

        edit = getattr(group, "edit", None)
        if edit is None:
            return

        # 核心修改：非聚焦状态下，设置文本为 "数值+后缀"
        if not edit.hasFocus():
            text_val = fmt_str.format(v / divisor if divisor else v)
            edit.setText(f"{text_val}{suffix}")

        # 计算宽度 (基于当前文本，包含后缀)
        fm = QFontMetrics(edit.font())
        text = edit.text() or "0"
        text_w = fm.horizontalAdvance(text) + 15 
        final_edit_w = max(35, min(160, text_w))

        if hasattr(group, "set_edit_width"):
            group.set_edit_width(final_edit_w)
        
        group.setFixedHeight(edit.height())
        group.adjustSize()

        total_w = group.width()
        x = base_x + cx - total_w / 2.0
        if clamp:
            min_x = base_x + 2
            max_x = base_x + slider.width() - total_w - 2
            x = max(min_x, min(x, max_x))

        y = base_y - group.height() - gap
        group.move(int(x), int(y))
        if not group.isVisible():
            group.show()
        group.raise_()

    _update_and_place_group(group_l, val_l, cx_l)

    if single_handle:
        if group_r is not None:
            group_r.hide()
    else:
        _update_and_place_group(group_r, val_r, cx_r)

class SimpleRangeSlider(QWidget):
    """
    顶部倒三角手柄 + 分层概率显示的滑动条控件。
    """
    rangeChanged = Signal(float, float)
    dragStarted = Signal()
    dragFinished = Signal(float, float)
    inputConfirmed = Signal(str, str)

    def __init__(self, parent=None):
        super().__init__(parent)
        # 开启点击获取焦点策略
        self.setFocusPolicy(Qt.ClickFocus)
        self._row_height = 24
        self._handle_height = 14
        self._max_layers = 3
        self.setMinimumHeight(self._handle_height + self._row_height)

        self._min = 0.0
        self._max = 100.0
        self._low = 20.0
        self._high = 80.0
        self.margin_left = 10
        self.margin_right = 10

        self._layers = []
        self._dragging = None
        self._is_dragging = False
        self._single_handle = False

        # 存储绑定的输入框引用
        self._bound_edits = []

        self.handle_color = QColor("#333333")
        self.handle_border = QColor("#ffffff")

        self.setMouseTracking(True)
        self._active_editor = None
        self._coalesce_ms = 16
        self._pending_emit = None
        self._emit_timer = QTimer(self)
        self._emit_timer.setSingleShot(True)
        self._emit_timer.timeout.connect(self._flush_pending_emit)

    def set_layers(self, layers: list):
        self._layers = layers
        n = max(1, len(self._layers))
        new_h = self._handle_height + (n * self._row_height) + 4
        self.setFixedHeight(new_h)
        self.update()

    def setMargins(self, left, right):
        self.margin_left = left
        self.margin_right = right
        self.update()

    def setRangeLimit(self, min_val, max_val):
        self._min = float(min_val)
        self._max = float(max_val)
        self.update()

    def setSingleHandleMode(self, enabled: bool):
        self._single_handle = bool(enabled)
        if self._single_handle: self._high = self._low
        self.update()

    def setRangeValues(self, low, high):
        self._low = max(self._min, min(float(low), self._max))
        self._high = min(self._max, max(float(high), self._min))
        if self._low > self._high: self._low = self._high
        self.update()
        if self._active_editor: self._close_editor()

    def setSnapValues(self, values, enabled: bool = True):
        pass

    def setEmitCoalesceMs(self, ms: int):
        self._coalesce_ms = max(0, int(ms))

    # 实现绑定逻辑
    def bind_line_edits(self, l, r):
        """绑定左右悬浮输入框，以便在点击滑块时让它们失去焦点"""
        self._bound_edits = [l, r]

    def setPlotWidth(self, w: float):
        self._plot_w = w
        self.update()

    def get_valid_width(self):
        """优先使用 Plotly 的宽度，如果没有才降级使用 Qt 宽度"""
        if getattr(self, "_plot_w", None) and self._plot_w > 0:
            return self._plot_w
        return self.width() - self.margin_left - self.margin_right

    def val_to_pixel(self, val):
        w = self.get_valid_width()
        if w <= 0 or abs(self._max - self._min) < 1e-9: return self.margin_left
        return self.margin_left + ((val - self._min) / (self._max - self._min)) * w

    def pixel_to_val(self, px):
        w = self.get_valid_width()
        if w <= 0: return self._min
        val = self._min + ((px - self.margin_left) / w) * (self._max - self._min)
        return max(self._min, min(val, self._max))

    def _emit_range_changed(self, low, high, immediate=False):
        if immediate or self._coalesce_ms <= 0:
            self.rangeChanged.emit(float(low), float(high))
        else:
            self._pending_emit = (float(low), float(high))
            if not self._emit_timer.isActive(): self._emit_timer.start(self._coalesce_ms)

    def _flush_pending_emit(self):
        if self._pending_emit:
            self.rangeChanged.emit(*self._pending_emit)
            self._pending_emit = None

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)

        w = self.width()
        valid_w = self.get_valid_width()
        x_low = self.val_to_pixel(self._low)
        x_high = self.val_to_pixel(self._high)

        # 1. 绘制顶部手柄
        handle_zone_h = self._handle_height
        painter.setPen(QPen(self.handle_border, 1.0))
        painter.setBrush(QBrush(self.handle_color))

        size = 12
        hw, hh = size * 0.6, size * 0.6
        tip_y = handle_zone_h
        base_y = max(0.0, handle_zone_h - hh * 2)

        left_tri = QPolygonF([
            QPointF(x_low - hw, base_y),
            QPointF(x_low + hw, base_y),
            QPointF(x_low, tip_y)
        ])
        painter.drawPolygon(left_tri)

        if not self._single_handle:
            right_tri = QPolygonF([
                QPointF(x_high - hw, base_y),
                QPointF(x_high + hw, base_y),
                QPointF(x_high, tip_y)
            ])
            painter.drawPolygon(right_tri)

        # 2. 绘制分层显示
        draw_layers = self._layers if self._layers else [{'key': '', 'color': '#ccc', 'probs': (0, 0, 0)}]
        font = painter.font()
        font.setPixelSize(12)
        font.setBold(False)
        painter.setFont(font)

        for i, layer in enumerate(draw_layers):
            if i >= self._max_layers: break
            y = self._handle_height + i * self._row_height
            h = self._row_height
            c_hex = layer.get('color', '#333')
            theme_c = QColor(c_hex)

            border_pen = QPen(QColor('#000000'))
            border_pen.setWidth(1)

            probs = layer.get('probs', (0, 0, 0))

            rect_l = QRectF(self.margin_left, y + 1, max(0, x_low - self.margin_left), h - 2)
            rect_m = QRectF(x_low, y + 1, max(0, x_high - x_low), h - 2)
            rect_r = QRectF(x_high, y + 1, max(0, (self.margin_left + valid_w) - x_high), h - 2)

            def draw_segment(rect, prob):
                painter.setPen(border_pen)
                painter.setBrush(Qt.NoBrush)
                painter.drawRoundedRect(rect, 2, 2)
                if rect.width() >= 20:
                    # ==== 概率框显示两位小数 ====
                    txt = f"{prob * 100:.2f}%"
                    painter.setPen(theme_c)
                    painter.drawText(rect, Qt.AlignCenter, txt)

            draw_segment(rect_l, probs[0])
            draw_segment(rect_m, probs[1])
            draw_segment(rect_r, probs[2])

    def mousePressEvent(self, event):
        self.setFocus() 
        
        if event.button() != Qt.LeftButton: return

        # 点击滑块时，强制清除悬浮输入框的焦点
        for edt in self._bound_edits:
            if edt and edt.isVisible():
                edt.clearFocus()

        click_y = event.position().y()
        click_x = event.position().x()
        val = self.pixel_to_val(click_x)

        handle_zone_limit = self._handle_height + 2
        is_handle_hit = (click_y <= handle_zone_limit)

        if is_handle_hit:
            self.setCursor(Qt.ClosedHandCursor)
            if not self._is_dragging:
                self._is_dragging = True
                self.dragStarted.emit()

            if self._single_handle:
                self._dragging = 'single'
                self._low = val;
                self._high = val
            else:
                d_low = abs(val - self._low)
                d_high = abs(val - self._high)
                if d_low < d_high:
                    self._dragging = 'low'
                    self._low = min(val, self._high)
                else:
                    self._dragging = 'high'
                    self._high = max(val, self._low)
            self.update()
            self._emit_range_changed(self._low, self._high, immediate=True)
            return

        # 检测分层输入点击 (保持原有逻辑)
        layer_y_start = self._handle_height
        if click_y > layer_y_start:
            rel_y = click_y - layer_y_start
            row_idx = int(rel_y / self._row_height)

            if 0 <= row_idx < len(self._layers):
                layer = self._layers[row_idx]
                if not layer.get('key'): return
                x_low = self.val_to_pixel(self._low)
                x_high = self.val_to_pixel(self._high)
                y_pos = layer_y_start + row_idx * self._row_height
                h = self._row_height

                region = None
                rect = None
                if self.margin_left <= click_x < x_low:
                    region = 'l'
                    rect = QRectF(self.margin_left, y_pos, x_low - self.margin_left, h)
                elif x_low <= click_x < x_high:
                    region = 'm'
                    rect = QRectF(x_low, y_pos, x_high - x_low, h)
                elif x_high <= click_x <= (self.width() - self.margin_right):
                    region = 'r'
                    rect = QRectF(x_high, y_pos, (self.width() - self.margin_right) - x_high, h)

                if region and rect and rect.width() > 20:
                    self._spawn_editor(rect, layer, region, row_idx)

    def mouseMoveEvent(self, event):
        val = self.pixel_to_val(event.position().x())
        if self._dragging:
            if self._dragging == 'single':
                self._low = val; self._high = val
            elif self._dragging == 'low':
                self._low = min(val, self._high)
            else:
                self._high = max(val, self._low)
            self.update()
            if self._active_editor: self._close_editor()
            self._emit_range_changed(self._low, self._high, immediate=False)
        else:
            if event.position().y() <= self._handle_height + 2:
                self.setCursor(Qt.OpenHandCursor)
            else:
                self.setCursor(Qt.IBeamCursor)

    def mouseReleaseEvent(self, event):
        self._flush_pending_emit()
        self._dragging = None
        self.setCursor(Qt.ArrowCursor)
        if self._is_dragging:
            self._is_dragging = False
            self.dragFinished.emit(float(self._low), float(self._high))

    def _spawn_editor(self, rect: QRectF, layer, region, row_idx):
        self._close_editor()
        probs = layer.get('probs', (0, 0, 0))
        idx_map = {'l': 0, 'm': 1, 'r': 2}
        val = probs[idx_map[region]]

        edt = AutoSelectLineEdit(f"{val * 100:.2f}%")
        edt.setParent(self)
        edt.setGeometry(rect.toRect().adjusted(1, 2, -1, -2))
        c_hex = layer.get('color', '#333')
        edt.setStyleSheet(f"background: white; color: {c_hex}; border: 1px solid {c_hex}; font-weight: normal; font-family: 'Arial'; font-size: 12px;")
        edt.setAlignment(Qt.AlignCenter)

        def _confirm():
            txt = edt.text()
            mode_prefix = {'l': 'pl', 'm': 'pm', 'r': 'pr'}[region]
            mode_str = f"{mode_prefix}:{layer['key']}"
            self.inputConfirmed.emit(mode_str, txt)
            self._close_editor()

        edt.returnPressed.connect(_confirm)
        edt.show()
        edt.setFocus()
        self._active_editor = edt

    def _close_editor(self):
        if self._active_editor:
            self._active_editor.deleteLater()
            self._active_editor = None


# =================================================================
# 第 6 组：图表配置与单位转换器 (Chart Configuration & Unit Formatters)
# =================================================================

SI_MAG_MAP = {
    12: "T", 9: "G", 6: "M", 3: "k", 0: "",
    -3: "m", -6: "u", -9: "n", -12: "p", -15: "f", -18: "a"
}

_SUPERSCRIPT_MAP = {
    "-": "\u207b",
    "0": "\u2070",
    "1": "\u00b9",
    "2": "\u00b2",
    "3": "\u00b3",
    "4": "\u2074",
    "5": "\u2075",
    "6": "\u2076",
    "7": "\u2077",
    "8": "\u2078",
    "9": "\u2079",
}

def infer_si_mag(value_range=None, dtick: float = None, forced_mag: int = None) -> int:
    """推断最接近的 SI 指数 (3 的倍数)"""
    if forced_mag is not None:
        return int(forced_mag)

    anchor = None
    try:
        if value_range and len(value_range) >= 2:
            a, b = float(value_range[0]), float(value_range[1])
            anchor = max(abs(a), abs(b))
        if (anchor is None or anchor <= 0) and dtick is not None:
            anchor = abs(float(dtick))
        if anchor is None or anchor <= 0:
            return 0

        mag = math.floor(math.log10(anchor))
        return int((mag // 3) * 3)
    except Exception:
        return 0

def get_si_unit_config(value_range=None, dtick: float = None, force_m: int = None):
    """返回规范化的单位配置 (standard_mag, suffix, divisor, title_hint)"""
    standard_mag = infer_si_mag(value_range=value_range, dtick=dtick, forced_mag=force_m)

    suffix = SI_MAG_MAP.get(standard_mag, "")
    if suffix and standard_mag != 0:
        divisor = 10 ** standard_mag
        exp_str = str(standard_mag)
        exp_display = "".join(_SUPERSCRIPT_MAP.get(c, c) for c in exp_str)
        title_hint = f"{suffix}=10{exp_display}"
    else:
        divisor = 1.0
        title_hint = ""
        suffix = ""

    return standard_mag, suffix, divisor, title_hint


def build_plotly_axis_style(
        *,
        x_range=None,
        y_range=None,
        x_dtick=None,
        y_dtick=None,
        x_title=None,
        y_title=None,
        x_unit=None,
        y_unit=None,
        y_visible=True,
        fixedrange=True,
        grid_color="#dddddd",
        border_color="#000000",
        border_width=1,
        forced_mag=None,
        y_forced_mag=None,
        label_overrides=None,
        x_axis_numeric: bool = True,
        y_axis_numeric: bool = True,
):
    """构建包含 SI 单位缩放及可选标签重写的 Plotly 坐标轴样式字典"""
    overrides = label_overrides if isinstance(label_overrides, dict) else {}
    axis_title_overrides = overrides.get("axis_title", {}) if isinstance(overrides.get("axis_title", {}), dict) else {}
    axis_tick_overrides = overrides.get("axis_tick", {}) if isinstance(overrides.get("axis_tick", {}), dict) else {}

    def _axis_title_cfg(axis_key: str) -> dict:
        sec = axis_title_overrides.get(axis_key, {})
        return sec if isinstance(sec, dict) else {}

    def _axis_tick_cfg(axis_key: str) -> dict:
        sec = axis_tick_overrides.get(axis_key, {})
        return sec if isinstance(sec, dict) else {}

    def _resolve_axis_numeric(axis_key: str, default_flag: bool) -> bool:
        sec = _axis_tick_cfg(axis_key)
        if "is_numeric" in sec:
            return bool(sec.get("is_numeric", default_flag))
        return bool(default_flag)

    def _resolve_mag_override(raw_value):
        if raw_value is None:
            return None
        try:
            return int(raw_value)
        except Exception:
            return None

    def _numeric_tickformat(fmt_key: str) -> str:
        mapping = {
            "number": ",.2f",
            "currency": "$,.2f",
            "percent": ".2%",
            "scientific": ".2e",
            "integer": ",.0f",
        }
        return mapping.get(str(fmt_key or "").strip().lower(), "")

    def _format_numeric_text(display_num: float, fmt_key: str) -> str:
        key = str(fmt_key or "auto").strip().lower()
        if key == "number":
            txt = f"{display_num:,.2f}"
        elif key == "currency":
            txt = f"${display_num:,.2f}"
        elif key == "percent":
            txt = f"{display_num * 100:.2f}%"
        elif key == "scientific":
            txt = f"{display_num:.2e}"
        elif key == "integer":
            txt = f"{display_num:,.0f}"
        else:
            txt = f"{display_num:g}"
        return txt

    def _format_tick_text(value: float, divisor: float, suffix: str, fmt_key: str, is_numeric_axis: bool) -> str:
        if not is_numeric_axis:
            return f"{value:g}"
        display_num = value / divisor if divisor else value
        if abs(display_num) < 1e-15:
            display_num = 0.0
        body = _format_numeric_text(display_num, fmt_key)
        if suffix and not str(body).endswith("%"):
            return f"{body}{suffix}"
        return body

    def generate_manual_ticks(val_range, dtick, divisor, suffix, fmt_key, is_numeric_axis):
        if not val_range or dtick is None or dtick <= 0:
            return None, None
        start, end = val_range[0], val_range[1]
        epsilon = dtick * 1e-3
        first_tick = math.ceil(start / dtick - 1e-9) * dtick
        vals, texts = [], []
        current = first_tick
        while current <= end + epsilon:
            if abs(current) < epsilon:
                current = 0.0
            vals.append(current)
            texts.append(_format_tick_text(current, divisor, suffix, fmt_key, is_numeric_axis))
            current += dtick
        return vals, texts

    x_title_cfg = _axis_title_cfg("x")
    x_tick_cfg = _axis_tick_cfg("x")
    y_title_cfg = _axis_title_cfg("y")
    y_tick_cfg = _axis_tick_cfg("y")

    x_axis_is_numeric = _resolve_axis_numeric("x", x_axis_numeric)
    y_axis_is_numeric = _resolve_axis_numeric("y", y_axis_numeric)
    x_number_format = str(x_tick_cfg.get("number_format", "auto") or "auto").strip().lower()
    y_number_format = str(y_tick_cfg.get("number_format", "auto") or "auto").strip().lower()
    if not x_axis_is_numeric:
        x_number_format = "auto"
    if not y_axis_is_numeric:
        y_number_format = "auto"

    x_mag_override = _resolve_mag_override(x_tick_cfg.get("display_unit_mag", None))
    y_mag_override = _resolve_mag_override(y_tick_cfg.get("display_unit_mag", None))
    x_force_m = forced_mag if forced_mag is not None else x_mag_override
    y_force_m = y_forced_mag if y_forced_mag is not None else y_mag_override

    x_mag, x_suffix, x_div, x_hint = get_si_unit_config(x_range, x_dtick, force_m=x_force_m)
    y_mag, y_suffix, y_div, y_hint = get_si_unit_config(y_range, y_dtick, force_m=y_force_m)
    if not x_axis_is_numeric:
        x_mag, x_suffix, x_div, x_hint = 0, "", 1.0, ""
    if not y_axis_is_numeric:
        y_mag, y_suffix, y_div, y_hint = 0, "", 1.0, ""

    x_title_base = x_title_cfg.get("text_override", None)
    if x_title_base is None:
        x_title_base = x_title
    y_title_base = y_title_cfg.get("text_override", None)
    if y_title_base is None:
        y_title_base = y_title

    x_title_parts = []
    if x_title_base and str(x_title_base).strip():
        x_title_parts.append(str(x_title_base).strip())
    if x_unit and str(x_unit).strip():
        x_title_parts.append(str(x_unit).strip())
    full_x_title = " ".join(x_title_parts)
    if x_hint:
        final_x_title = f"{full_x_title} ({x_hint})" if full_x_title else x_hint
    else:
        final_x_title = full_x_title

    y_title_parts = []
    base_y = str(y_title_base).strip() if y_title_base is not None else "密度"
    if base_y:
        y_title_parts.append(base_y)
    if y_unit and str(y_unit).strip():
        y_title_parts.append(str(y_unit).strip())
    full_y_title = " ".join(y_title_parts)
    if y_hint:
        final_y_title = f"{full_y_title} ({y_hint})" if full_y_title else y_hint
    else:
        final_y_title = full_y_title

    axis_font = dict(family="Arial", size=12.5, color="#333333")
    title_font = dict(family="Arial, 'Microsoft YaHei', sans-serif", size=12, color="#333333")

    x_tick_family = str(x_tick_cfg.get("font_family", "") or "").strip()
    y_tick_family = str(y_tick_cfg.get("font_family", "") or "").strip()
    x_title_family = str(x_title_cfg.get("font_family", "") or "").strip()
    y_title_family = str(y_title_cfg.get("font_family", "") or "").strip()

    axis_font_x = dict(axis_font, family=x_tick_family) if x_tick_family else dict(axis_font)
    axis_font_y = dict(axis_font, family=y_tick_family) if y_tick_family else dict(axis_font)

    try:
        x_tick_size = float(x_tick_cfg.get("font_size")) if x_tick_cfg.get("font_size") is not None else None
    except Exception:
        x_tick_size = None
    try:
        y_tick_size = float(y_tick_cfg.get("font_size")) if y_tick_cfg.get("font_size") is not None else None
    except Exception:
        y_tick_size = None
    if x_tick_size is not None and x_tick_size > 0:
        axis_font_x["size"] = x_tick_size
    if y_tick_size is not None and y_tick_size > 0:
        axis_font_y["size"] = y_tick_size

    title_font_x = dict(title_font, family=x_title_family or title_font["family"])
    title_font_y = dict(title_font, family=y_title_family or title_font["family"])
    try:
        x_title_size = float(x_title_cfg.get("font_size")) if x_title_cfg.get("font_size") is not None else None
    except Exception:
        x_title_size = None
    try:
        y_title_size = float(y_title_cfg.get("font_size")) if y_title_cfg.get("font_size") is not None else None
    except Exception:
        y_title_size = None
    if x_title_size is not None and x_title_size > 0:
        title_font_x["size"] = x_title_size
    if y_title_size is not None and y_title_size > 0:
        title_font_y["size"] = y_title_size

    x_vals, x_texts = generate_manual_ticks(x_range, x_dtick, x_div, x_suffix, x_number_format, x_axis_is_numeric)
    y_vals, y_texts = generate_manual_ticks(y_range, y_dtick, y_div, y_suffix, y_number_format, y_axis_is_numeric)
    if x_texts:
        offset_space = "&nbsp;" * 4
        x_texts = [f"{offset_space}{t}" for t in x_texts]
    if y_texts:
        y_texts = [f"<br>{t}" for t in y_texts]

    xaxis = dict(
        title={"text": final_x_title, "font": title_font_x} if final_x_title else None,
        range=x_range,
        tickfont=axis_font_x,
        fixedrange=fixedrange,
        showgrid=True,
        gridcolor=grid_color,
        gridwidth=0.5,
        showline=True,
        mirror=False,
        linecolor=border_color,
        linewidth=border_width,
        zeroline=False,
        ticks="outside",
        ticklen=4,
        tickwidth=1,
        tickcolor=border_color,
        tickmode="array" if x_vals else "linear",
        tickvals=x_vals,
        ticktext=x_texts,
        dtick=x_dtick if not x_vals else None,
        hoverformat=".4g",
    )
    x_tickformat = _numeric_tickformat(x_number_format)
    if x_tickformat and not x_vals and x_axis_is_numeric:
        xaxis["tickformat"] = x_tickformat

    yaxis = dict(
        title={"text": final_y_title, "font": title_font_y} if final_y_title else None,
        range=y_range,
        tickfont=axis_font_y,
        visible=y_visible,
        fixedrange=fixedrange,
        showgrid=True,
        gridcolor=grid_color,
        gridwidth=0.5,
        showline=True,
        mirror=True,
        linecolor=border_color,
        linewidth=border_width,
        zeroline=False,
        ticks="outside",
        ticklen=4,
        tickwidth=1,
        tickcolor=border_color,
        tickmode="array" if y_vals else "linear",
        tickvals=y_vals,
        ticktext=y_texts,
        dtick=y_dtick if not y_vals else None,
        hoverformat=".4g",
    )
    y_tickformat = _numeric_tickformat(y_number_format)
    if y_tickformat and not y_vals and y_axis_is_numeric:
        yaxis["tickformat"] = y_tickformat

    return {"xaxis": xaxis, "yaxis": yaxis, "x_mag": x_mag, "y_mag": y_mag}


# =================================================================
# 第 7 组：渲染调度与视图同步器 (Render Dispatching & View Sync)
# =================================================================

class ChartOverlayController(QObject):
    """
    统一管理 WebView、Overlay 和 Slider 的几何同步
    """
    def __init__(self, web_view, overlay):
        super().__init__(web_view)
        self.web_view = web_view
        self.overlay = overlay
        self.web_view.installEventFilter(self)
        
        self._sync_timer = QTimer(self)
        self._sync_timer.setSingleShot(True)
        self._sync_timer.setInterval(0)
        self._sync_timer.timeout.connect(self._do_sync_plot_rect)
        self._settle_timer = QTimer(self)
        self._settle_timer.setSingleShot(True)
        self._settle_timer.setInterval(120)
        self._settle_timer.timeout.connect(self._do_sync_plot_rect)
        
        self._poll_timer = QTimer(self)
        self._poll_timer.setInterval(100) 
        self._poll_timer.timeout.connect(self._do_sync_plot_rect)
        self._poll_timer.start()

    def eventFilter(self, obj, event):
        if obj == self.web_view:
            if event.type() in (QEvent.Resize, QEvent.Show):
                if self.overlay:
                    self.overlay.setGeometry(self.web_view.rect())
                    self.overlay.raise_()
                    self.schedule_rect_sync()
        return super().eventFilter(obj, event)

    def schedule_rect_sync(self):
        self._sync_timer.start()
        self._settle_timer.start()

    def _do_sync_plot_rect(self):
        """执行具体的矩形同步操作并执行 JS 以获取内部图表尺寸"""
        if not self.overlay or not self.web_view: return
        js = """
        (function(){
            var gd = document.getElementById('chart_div');
            if(!gd || !gd._fullLayout || !gd._fullLayout._size) return null;
            var sz = gd._fullLayout._size;
            return { l: sz.l, t: sz.t, w: sz.w, h: sz.h, dpr: window.devicePixelRatio || 1.0 };
        })()
        """
        def _callback(res):
            if not res or not isinstance(res, dict): return
            try:
                l, t = float(res.get("l", 0)), float(res.get("t", 0))
                w, h = float(res.get("w", 0)), float(res.get("h", 0))
                dpr = float(res.get("dpr", 1.0))
                
                if w > 0 and h > 0:
                    current_dpr = getattr(self.overlay, 'dpr', 1.0)
                    dpr_changed = (abs(current_dpr - dpr) > 0.01)
                    
                    if hasattr(self.overlay, 'set_dpr'): 
                        self.overlay.set_dpr(dpr)
                    
                    # 性能优化防抖机制：只有当底层图表的物理边界真正发生伸缩改变时，才触发 Qt 遮罩重绘
                    old_override = getattr(self.overlay, '_plot_rect_override', None)
                    new_override = (l, t, w, h)
                    
                    if old_override != new_override or dpr_changed:
                        self.overlay.set_plot_rect(l, t, w, h)
                        self.overlay.update()
            except Exception: pass
            
        try:
            self.web_view.page().runJavaScript(js, _callback)
        except Exception: pass

    def update_visuals(self, l: float, r: float, x_min: float, x_max: float, mode: str = "hist"):
        if not self.overlay: return
        self.overlay.set_axis_range(x_min, x_max)
        self.overlay.set_mode(mode)
        self.overlay.set_lr(l, r)
        self.overlay.update()

class LayoutRefreshCoordinator(QObject):
    """布局与刷新协调器"""
    def __init__(self, parent, frame_cb, final_cb=None, frame_ms: int = 50, settle_ms: int = 180):
        super().__init__(parent)
        self._frame_cb = frame_cb
        self._final_cb = final_cb if final_cb is not None else frame_cb

        self._live_timer = QTimer(self)
        self._live_timer.setInterval(max(16, int(frame_ms)))
        self._live_timer.timeout.connect(self._on_live_tick)

        self._settle_timer = QTimer(self)
        self._settle_timer.setSingleShot(True)
        self._settle_timer.setInterval(max(60, int(settle_ms)))
        self._settle_timer.timeout.connect(self._on_settled)

    def notify(self):
        self._run_frame()
        if not self._live_timer.isActive():
            self._live_timer.start()
        self._settle_timer.start()

    def _on_live_tick(self):
        self._run_frame()

    def _on_settled(self):
        self._live_timer.stop()
        self._run_final()

    def _run_frame(self):
        try:
            if self._frame_cb is not None:
                self._frame_cb()
        except Exception:
            pass

    def _run_final(self):
        try:
            if self._final_cb is not None:
                self._final_cb()
        except Exception:
            pass


# =================================================================
# 第 8 组：遗留 UI 与 Excel 交互工具 (Legacy UI & Excel Tools)
# =================================================================

class XAxisMagnitudeComboBox(QComboBox):
    """提供围绕建议中心值附近三个 SI 量级的下拉选择框"""
    magnitudeSelected = Signal(int)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setFixedWidth(80)
        self.setStyleSheet(drisk_combobox_qss())
        self.currentIndexChanged.connect(self._on_index_changed)

    def populate_around(self, center_mag: int):
        self.blockSignals(True)
        self.clear()

        candidates = [center_mag - 3, center_mag, center_mag + 3]

        for mag in candidates:
            suffix = SmartFormatter.SI_MAP.get(mag, "")
            mag_str = str(mag)
            if mag == 0:
                label = "单位 (1)"
            else:
                label = f"{suffix} (10^{mag_str})" if suffix else f"10^{mag_str}"

            self.addItem(label, mag)

        self.setCurrentIndex(1)
        self.blockSignals(False)

    def _on_index_changed(self, index):
        if index < 0: return
        val = self.currentData()
        if val is not None:
            self.magnitudeSelected.emit(int(val))

# 通用单元格字符串查找 -- 默认图名正则规则
_UI_NAME_CELL_WITH_SUFFIX_RE = re.compile(r"^(?P<addr>[A-Za-z]{1,3}\d{1,7})(?:_[A-Za-z0-9]+)+$")
_UI_NAME_PLAIN_CELL_RE = re.compile(r"^[A-Za-z]{1,3}\d{1,7}$")
_UI_NAME_CELL_TOKEN_RE = re.compile(r"([A-Za-z]{1,3}\d{1,7})")
_UI_NAME_SHEET_CELL_REF_RE = re.compile(
    r"^(?:(?:'[^']+'|[A-Za-z_][A-Za-z0-9_\.]*)!)?\$?[A-Za-z]{1,3}\$?\d{1,7}$",
    re.IGNORECASE,
)
_UI_NAME_DRISK_NAME_CALL_RE = re.compile(
    r"DriskName\s*\(\s*(?P<arg>(?:\"(?:[^\"]|\"\")*\"|'(?:[^']|'')*'|[^)]*))\s*\)",
    re.IGNORECASE,
)

def extract_clean_cell_address(raw_key) -> str:
    """提取干净的 A1 样式地址作为显示后备方案"""
    text = str(raw_key or "").strip().replace("$", "")
    if not text:
        return ""

    addr = text.split("!", 1)[1] if "!" in text else text
    addr = addr.strip().upper()
    if not addr:
        return ""

    matched = _UI_NAME_CELL_WITH_SUFFIX_RE.match(addr)
    if matched:
        return matched.group("addr")
    if _UI_NAME_PLAIN_CELL_RE.match(addr):
        return addr

    left = addr.split("_", 1)[0] if "_" in addr else addr
    if _UI_NAME_PLAIN_CELL_RE.match(left):
        return left

    token = _UI_NAME_CELL_TOKEN_RE.search(addr)
    if token:
        return token.group(1).upper()
    return left or addr

def _unquote_formula_token(token: str) -> str:
    text = str(token or "").strip()
    if len(text) >= 2 and text[0] == text[-1] and text[0] in ("'", '"'):
        q = text[0]
        inner = text[1:-1]
        doubled = q * 2
        return inner.replace(doubled, q).strip()
    return text

def _split_sheet_and_addr_for_name_ref(raw_ref: str, default_sheet: str = "") -> tuple[str, str]:
    token = str(raw_ref or "").strip()
    if not token:
        return "", ""
    token = token.lstrip("=").strip()

    if not _UI_NAME_SHEET_CELL_REF_RE.match(token):
        return "", ""

    if "!" in token:
        sheet_name, addr = token.rsplit("!", 1)
        sheet_name = sheet_name.strip()
        if len(sheet_name) >= 2 and sheet_name[0] == sheet_name[-1] == "'":
            sheet_name = sheet_name[1:-1].replace("''", "'").strip()
    else:
        sheet_name, addr = default_sheet, token

    addr_clean = str(addr or "").replace("$", "").strip().upper()
    if not _UI_NAME_PLAIN_CELL_RE.match(addr_clean):
        return "", ""
    return str(sheet_name or "").strip(), addr_clean

def _resolve_name_ref_cell_text(excel_app, raw_ref: str, default_sheet: str = "") -> str:
    if excel_app is None:
        return ""

    sheet_name, addr = _split_sheet_and_addr_for_name_ref(raw_ref, default_sheet=default_sheet)
    if not addr:
        return ""

    try:
        if sheet_name:
            sheet = excel_app.ActiveWorkbook.Worksheets(sheet_name)
        else:
            sheet = excel_app.ActiveSheet
        val = sheet.Range(addr).Value
    except Exception:
        return ""

    if val is None:
        return ""
    text = str(val).strip()
    return text

def _extract_explicit_name_from_formula(formula_text: str, *, excel_app=None, default_sheet: str = "") -> str:
    """从公式文本中读取显式设定的 DriskName(...) 并解析单元格引用名称"""
    text = str(formula_text or "").strip()
    if not text:
        return ""

    match = _UI_NAME_DRISK_NAME_CALL_RE.search(text)
    if not match:
        return ""

    raw_arg = str(match.group("arg") or "").strip()
    if not raw_arg:
        return ""

    if len(raw_arg) >= 2 and raw_arg[0] == raw_arg[-1] and raw_arg[0] in ("'", '"'):
        return _unquote_formula_token(raw_arg)

    resolved = _resolve_name_ref_cell_text(excel_app, raw_arg, default_sheet=default_sheet)
    if resolved:
        return resolved
    return ""

def _resolve_explicit_name_from_attrs(raw_name: str, *, excel_app=None, default_sheet: str = "") -> str:
    """解析由 attrs 承载的显式名称"""
    token = _unquote_formula_token(raw_name)
    if not token:
        return ""

    looks_like_ref = bool(_UI_NAME_SHEET_CELL_REF_RE.match(token))
    if looks_like_ref and ("!" in token or "$" in token):
        resolved = _resolve_name_ref_cell_text(excel_app, token, default_sheet=default_sheet)
        if resolved:
            return resolved
        return ""
    return token

def resolve_visible_variable_name(cell_key, attrs=None, *, excel_app=None, fallback_label="") -> str:
    """
    解析 UI 层面可见的变量名称，按以下优先级统一处理：
    显式元数据名称 -> find_cell_name_ui 探测 -> 清理后的单元格地址兜底。
    """
    attrs_map = attrs if isinstance(attrs, dict) else {}

    key_text = str(cell_key or "").strip().replace("$", "")
    lookup_addr = extract_clean_cell_address(key_text)
    default_sheet = key_text.split("!", 1)[0].strip() if "!" in key_text else ""

    explicit_name_raw = str(attrs_map.get("name", "") or "").strip()
    should_probe_formula_name = (not explicit_name_raw) or bool(_UI_NAME_SHEET_CELL_REF_RE.match(explicit_name_raw))

    if excel_app is not None and lookup_addr and should_probe_formula_name:
        try:
            if default_sheet:
                sheet = excel_app.ActiveWorkbook.Worksheets(default_sheet)
            else:
                sheet = excel_app.ActiveSheet
            formula_text = str(sheet.Range(lookup_addr).Formula)
            formula_name = _extract_explicit_name_from_formula(
                formula_text,
                excel_app=excel_app,
                default_sheet=default_sheet,
            )
            if formula_name:
                return formula_name
        except Exception:
            pass

    explicit_name = _resolve_explicit_name_from_attrs(
        explicit_name_raw,
        excel_app=excel_app,
        default_sheet=default_sheet,
    )
    if explicit_name:
        return explicit_name

    if excel_app is not None and lookup_addr:
        try:
            if default_sheet:
                sheet = excel_app.ActiveWorkbook.Worksheets(default_sheet)
            else:
                sheet = excel_app.ActiveSheet
            cell = sheet.Range(lookup_addr)
            full_name, _ = find_cell_name_ui(excel_app, sheet, cell)
            full_name = str(full_name or "").strip()
            if full_name:
                return full_name
        except Exception:
            pass

    if lookup_addr:
        return lookup_addr

    fb = str(fallback_label or "").strip()
    if fb:
        return fb
    return key_text.split("!")[-1] if "!" in key_text else key_text

def find_cell_name_ui(app, sheet, cell) -> str:
    """
    为单元格自动命名：向上查找第一个非空字符串，向左查找第一个非空字符串。
    格式：左方字符串_上方字符串 或找到的单个字符串
    """
    # 获取单元格的行列
    row = cell.Row
    col = cell.Column
    
    # 向上查找（最大10行）
    up_name = ""
    for i in range(1, 11):
        if row - i < 1:
            break
        up_cell = sheet.Cells(row - i, col)
        val = up_cell.Value
        if val is not None and isinstance(val, str) and val.strip():
            up_name = str(val).strip()
            break
    
    # 向左查找（最大10列）
    left_name = ""
    for i in range(1, 11):
        if col - i < 1:
            break
        left_cell = sheet.Cells(row, col - i)
        val = left_cell.Value
        if val is not None and isinstance(val, str) and val.strip():
            left_name = str(val).strip()
            break
    
    # 组合名称
    if up_name and left_name:
        full_name = f"{left_name}_{up_name}"
    elif up_name:
        full_name = up_name
    elif left_name:
        full_name = left_name
    else:
        full_name = ""
        
    # 简化返回值，只返回图名兜底和类别兜底
    return full_name, up_name

def apply_toolbar_button_icon(
    button: QPushButton,
    icon_name: str,
    icon_px: int = 14,
    icon_only: bool = False,
    button_px: int = 0,
) -> bool:
    """应用图标资源目录中的图标至工具栏按钮"""
    try:
        if button is None:
            return False
        base_dir = os.path.dirname(os.path.abspath(__file__))
        icon_file = str(icon_name or "").strip()
        if not icon_file:
            return False
        icon_path = os.path.join(base_dir, "icons", icon_file)
        if not os.path.exists(icon_path):
            return False
        icon = QIcon(icon_path)
        if icon.isNull():
            return False
        button.setAutoDefault(False)
        button.setDefault(False)
        button.setCursor(Qt.PointingHandCursor)
        if icon_only:
            button.setText("")
            button.setFlat(True)
            if int(button_px) > 0:
                edge = int(button_px)
                button.setFixedSize(edge, edge)
            button.setStyleSheet(
                """
                QPushButton {
                    background-color: transparent;
                    border: none;
                    border-radius: 4px;
                    padding: 0px;
                    margin: 0px;
                }
                QPushButton:hover {
                    background-color: rgba(0, 80, 179, 0.10);
                }
                QPushButton:pressed {
                    background-color: rgba(0, 80, 179, 0.18);
                }
                QPushButton:checked {
                    background-color: rgba(0, 80, 179, 0.14);
                }
                QPushButton:focus {
                    border: none;
                    outline: none;
                    background-color: rgba(0, 80, 179, 0.12);
                }
                """
            )
        button.setIcon(icon)
        try:
            px = max(10, min(32, int(icon_px)))
            if icon_only and int(button_px) > 0:
                px = max(10, min(px, int(button_px) - 2))
            button.setIconSize(QSize(px, px))
        except Exception:
            pass
        return True
    except Exception:
        return False

def apply_excel_select_button_icon(button: QPushButton, icon_name: str = "select_icon.svg") -> bool:
    """应用共享的 Excel 单元格选取图标至指定按钮"""
    try:
        if button is None:
            return False

        try:
            w = int(button.width() or button.size().width() or 0)
            h = int(button.height() or button.size().height() or 0)
            if w > 0 and h > 0 and w != h:
                edge = min(w, h)
                button.setFixedSize(edge, edge)
        except Exception:
            pass

        button.setAutoDefault(False)
        button.setDefault(False)
        button.setCursor(Qt.PointingHandCursor)
        button.setFlat(True)
        button.setStyleSheet(
            """
            QPushButton {
                background-color: transparent;
                border: none;
                border-radius: 999px;
                padding: 0px;
                margin: 0px;
            }
            QPushButton:hover {
                background-color: rgba(0, 80, 179, 0.10);
            }
            QPushButton:pressed {
                background-color: rgba(0, 80, 179, 0.18);
            }
            QPushButton:focus {
                border: none;
                outline: none;
                background-color: rgba(0, 80, 179, 0.12);
            }
            QPushButton:disabled {
                background-color: transparent;
            }
            """
        )

        base_dir = os.path.dirname(os.path.abspath(__file__))
        icon_path = os.path.join(base_dir, "icons", str(icon_name or "").strip() or "select_icon.svg")
        if not os.path.exists(icon_path):
            return False

        icon = QIcon(icon_path)
        if icon.isNull():
            return False

        button.setText("")
        button.setIcon(icon)
        try:
            w = int(button.width() or button.size().width() or 0)
            h = int(button.height() or button.size().height() or 0)
            edge = min(w, h) if (w > 0 and h > 0) else 24
            icon_px = max(14, min(24, edge - 2))
            button.setIconSize(QSize(icon_px, icon_px))
        except Exception:
            pass
        return True
    except Exception:
        return False

def set_drisk_icon(window_widget, icon_name="simu_icon.svg"):
    """从图标库加载并设置统一标准的方形窗口应用图标"""
    try:
        base_dir = os.path.dirname(os.path.abspath(__file__))
        raw_icon_name = str(icon_name or "").strip() or "simu_icon.svg"
        icon_candidates = []
        lowered = raw_icon_name.lower()
        if lowered in ("drisk_icon.png", "simu_icon.svg"):
            icon_candidates.extend(["simu_icon.svg", "DRISK_icon.png", "drisk_icon.png"])
        else:
            icon_candidates.append(raw_icon_name)

        icon_path = ""
        for name in icon_candidates:
            candidate = os.path.join(base_dir, "icons", name)
            if os.path.exists(candidate):
                icon_path = candidate
                break

        if not icon_path:
            return

        original_pixmap = QPixmap(icon_path)
        if original_pixmap.isNull():
            try:
                probe = QIcon(icon_path).pixmap(256, 256)
                if not probe.isNull():
                    original_pixmap = probe
            except Exception:
                pass
        if original_pixmap.isNull():
            return

        max_side = max(original_pixmap.width(), original_pixmap.height())
        square_pixmap = QPixmap(max_side, max_side)
        square_pixmap.fill(QColor(0, 0, 0, 0))

        x_offset = (max_side - original_pixmap.width()) // 2
        y_offset = (max_side - original_pixmap.height()) // 2

        painter = QPainter(square_pixmap)
        painter.setRenderHint(QPainter.Antialiasing, True)
        painter.setRenderHint(QPainter.SmoothPixmapTransform, True)
        painter.drawPixmap(x_offset, y_offset, original_pixmap)
        painter.end()

        window_widget.setWindowIcon(QIcon(square_pixmap))
    except Exception as e:
        print(f"[UI_SHARED] Failed to set icon '{icon_name}': {e}")