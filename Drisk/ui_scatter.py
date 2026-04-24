# -*- coding: utf-8 -*-
"""
ui_scatter.py
Drisk 散点图与相关性分析视图组件。
=========================================================
本模块提供散点图的配置、数据清洗、多情景叠加以及 Web/Qt 混合渲染服务。
特性：分为配置窗口与绘图窗口，底层接口与 ui_results 对齐。
"""

import drisk_env

import copy
import json
import math
import os
import time
import traceback
from typing import Any, Dict, List, Optional

import numpy as np
import pandas as pd
from pyxll import xl_app, xlcAlert
from PySide6.QtCore import Qt, QEvent, QEventLoop, QPoint, QPointF, QRectF, QSignalBlocker, QTimer, Signal
from PySide6.QtGui import QBrush, QColor, QFontMetrics, QPainter, QPen, QPolygonF
from PySide6.QtWidgets import (QAbstractItemView, QApplication, QCheckBox, QDialog, QFormLayout, QFrame, QGridLayout, QGroupBox,
                               QHBoxLayout, QLabel, QInputDialog, QLineEdit, QListWidget, QListWidgetItem, QMenu, QMessageBox, QPushButton,
                               QSplitter, QTableWidget, QVBoxLayout, QWidget)

# =======================================================
# 后端功能 -- 模拟数据获取与暴露接口
# =======================================================
from simulation_manager import get_all_simulations, get_simulation
from input_sample_exposure import is_input_key_exposed

# =======================================================
# 前端功能 -- 绘图工厂、导出管理器、界面设定与样式
# =======================================================
from drisk_charting import DriskChartFactory, SCATTER_SYMBOLS
from drisk_export import DriskExportManager
from plotly_host import PlotlyHost
from ui_style_manager import StyleManagerDialog
from ui_shared import (ChartSkeleton, ChartOverlayController, DRISK_COLOR_CYCLE, DriskMath, LayoutRefreshCoordinator, SI_MAG_MAP,
                       SmartFormatter, SuffixAutoSelectLineEdit, create_floating_value_with_mag,
                       floating_value_edit_qss, get_plotly_cache_dir, get_shared_webview, infer_si_mag, recycle_shared_webview,
                       set_drisk_icon, resolve_visible_variable_name, apply_excel_select_button_icon,
                       apply_toolbar_button_icon)
from ui_label_settings import (
    LabelSettingsDialog,
    create_default_label_settings_config,
    get_axis_display_unit_override,
    get_axis_numeric_flags,
)
from ui_stats import SEP_ROW_MARKER, render_stats_table
import backend_bridge as bridge
from ui_sim_display_names import (
    extract_sim_id_from_series_key,
    get_custom_sim_display_name,
    get_sim_display_name,
    get_sim_display_name_version,
    set_sim_display_name,
)


# =======================================================
# 0. 全局通用辅助函数
# =======================================================
def _sync_application_icon_from(window_widget: QWidget) -> None:
    """
    同步应用图标：
    提取传入窗口的图标，并将其设置到全局 QApplication 实例中，确保任务栏图标一致性。
    """
    try:
        app = QApplication.instance()
        if app is None:
            return
        icon = window_widget.windowIcon()
        if icon is not None and not icon.isNull():
            app.setWindowIcon(icon)
    except Exception:
        print("_sync_application_icon_from failed", exc_info=True)


def _rename_sim_with_dialog(parent, sim_id) -> bool:
    """
    情景重命名弹窗服务：
    提供标准输入对话框，允许用户在散点图界面直接重命名选定的模拟情景（如 sim1 命名为 Base Case）。
    留空则恢复为默认名称。若名称发生实质改变，返回 True 通知上层刷新 UI。
    """
    current_custom = get_custom_sim_display_name(sim_id)
    current_effective = get_sim_display_name(sim_id)
    initial_text = current_custom if current_custom else current_effective
    text, ok = QInputDialog.getText(
        parent,
        "重命名情景",
        "输入用于绘图界面的情景显示名（留空可恢复默认）：",
        text=initial_text,
    )
    if not ok:
        return False
    before = get_sim_display_name(sim_id)
    after = set_sim_display_name(sim_id, text)
    return str(before) != str(after) or (not str(text or "").strip() and bool(current_custom))



# =======================================================
# 1. 基础 UI 组件与交互控件 (滑块、十字光标、输入框)
# =======================================================

class VerticalRotatableValue(QWidget):
    """
    Y 轴专用的垂直旋转数值悬浮输入框。
    
    工作机制：
    1. 折叠状态（默认）：为了节省垂直空间的横向占用，利用 QPainter 将数值文本旋转 90 度绘制。
    2. 展开状态（交互）：鼠标点击后，将原有的文本隐藏，动态显示一个标准的水平 QLineEdit 供用户精准输入数值，失去焦点后自动回退折叠。
    """
    valueChanged = Signal(float)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setFocusPolicy(Qt.ClickFocus)
        # 数值状态与量级控制
        self.mag_div = 1.0
        self.suffix = ""
        self.val = 0.0
        self.decimals = 2  

        self.edit = SuffixAutoSelectLineEdit(self)
        self.edit.setAlignment(Qt.AlignCenter)
        self.edit.setStyleSheet(floating_value_edit_qss(opacity=0.95))
        
        font = self.edit.font()
        font.setPixelSize(12) 
        font.setFamily("Arial")
        self.edit.setFont(font)
        
        self.edit.hide()
        
        self.edit.editingFinished.connect(self._on_edit_finished)
        self.edit.installEventFilter(self)

    def set_value(self, val, mag_div, suffix, decimals=2):
        """设定当前要显示的数值及其缩放倍数和后缀（例如 1000, 10^3, 'K'）。"""
        self.val = val
        self.mag_div = mag_div
        self.suffix = suffix
        self.decimals = decimals
        if not self.edit.hasFocus():
            self.edit.setText(f"{val / mag_div:.{decimals}f}{suffix}")
            self.update()

    def _collapse(self):
        """收起输入框，恢复为紧凑的垂直绘制模式。"""
        self.edit.hide()
        self.setFixedSize(20, 50)
        self.update()

    def _expand(self):
        """展开为水平输入框，计算文本所需的宽度，并获取键盘焦点。"""
        print("expand???")
        fm = QFontMetrics(self.edit.font())
        text_w = fm.horizontalAdvance(self.edit.text()) + 15

        # 限制展开的最大与最小宽度
        final_w = max(35, min(160, text_w))

        self.setFixedSize(final_w, 20)
        self.edit.setGeometry(0, 0, final_w, 20)

        self.edit.show()
        self.edit.setFocus(Qt.MouseFocusReason)

        # 与 AutoSelectLineEdit 对齐：用 Win32 确保 Qt 窗口成为前台窗口
        # 否则 Excel 在 Windows 消息层仍持有键盘输入，导致无法输入
        window = self.topLevelWidget()
        if window:
            window.raise_()
            window.activateWindow()
            try:
                import ctypes
                ctypes.windll.user32.AllowSetForegroundWindow(-1)
                ctypes.windll.user32.SetForegroundWindow(int(window.winId()))
            except Exception:
                pass

    def eventFilter(self, obj, event):
        """事件过滤器：当输入框失去焦点时自动触发折叠。"""
        if obj == self.edit and event.type() == QEvent.FocusOut:
            self._collapse() 
        return super().eventFilter(obj, event)

    def _on_edit_finished(self):
        """用户敲击回车完成输入后：剥离单位后缀，乘回量级系数恢复为真实底层数值，并抛出信号。"""
        try:
            txt = self.edit.text().replace(self.suffix, '')
            new_val = float(txt) * self.mag_div
            self.valueChanged.emit(new_val)
        except ValueError:
            pass
        self._collapse()

    def mousePressEvent(self, event):
        """非编辑状态下点击任意区域均触发输入框展开。"""
        if not self.edit.isVisible():
            self._expand()
        super().mousePressEvent(event)

    def paintEvent(self, event):
        """自定义绘制事件：处理旋转 90 度的文本渲染。"""
        if self.edit.isVisible(): 
            return
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)
        painter.setPen(QColor("#333333"))
        font = painter.font()
        font.setPixelSize(12)
        painter.setFont(font)

        txt = f"{self.val / self.mag_div:.{self.decimals}f}{self.suffix}"
        
        # 将原点平移至控件中心后旋转坐标系
        painter.translate(self.width() / 2, self.height() / 2)
        painter.rotate(90)
        text_rect = QRectF(-self.height() / 2, -self.width() / 2, self.height(), self.width())
        painter.drawText(text_rect, Qt.AlignCenter, txt)


class CrosshairLineOverlay(QWidget):
    """
    底层混合渲染遮罩：叠加在 WebEngine (Plotly 图表) 上的原生 Qt 绘图层。
    职责：
    1. 绘制一条横穿 X/Y 轴的十字对齐线。
    2. 在画布的四个角落动态渲染各情景数据落入四个象限的样本数百分比。
    """
    rectUpdated = Signal(float, float, float, float)

    def __init__(self, parent=None):
        super().__init__(parent)
        # 允许鼠标事件穿透，确保用户可以与底层的 Plotly 进行缩放/悬停交互
        self.setAttribute(Qt.WidgetAttribute.WA_TransparentForMouseEvents, True)
        self.setAttribute(Qt.WidgetAttribute.WA_NoSystemBackground, True)
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground, True)
        
        # 内部状态缓存
        self.x_val = 0.0
        self.y_val = 0.0
        self.xmin = 0.0
        self.xmax = 1.0
        self.ymin = 0.0
        self.ymax = 1.0
        
        self.dpr = 1.0
        self.margin_l = 0.0
        self.margin_r = 0.0
        self.margin_t = 0.0
        self.margin_b = 0.0
        self._plot_rect_override = None
        self.quadrant_stats = []

    def set_margins(self, l: float, r: float, t: float, b: float):
        """接收图表边距更新（由 Plotly 的 layout.margin 回传）。"""
        self.margin_l, self.margin_r = float(l), float(r)
        self.margin_t, self.margin_b = float(t), float(b)
        self.update()

    def set_dpr(self, dpr: float):
        self.dpr = float(dpr) if dpr > 0 else 1.0

    def set_quadrant_stats(self, stats):
        """接收四个象限的统计结果数组并重绘。"""
        self.quadrant_stats = stats
        self.update()

    def set_axis_range(self, xmin: float, xmax: float, ymin: float, ymax: float):
        """设置坐标系数据极值范围，用于辅助计算数值到屏幕像素的映射关系。"""
        self.xmin = float(xmin)
        self.xmax = float(xmax)
        self.ymin = float(ymin)
        self.ymax = float(ymax)
        self.update()

    def set_crosshair_values(self, x_val: float, y_val: float):
        """设置当前十字线的真实物理数据落点。"""
        self.x_val = float(x_val)
        self.y_val = float(y_val)
        self.update()

    def set_plot_rect(self, l: float, t: float, w: float, h: float):
        """接收协调器下发的精准图表内部渲染矩形边界，同步遮罩覆盖范围。"""
        if w <= 0 or h <= 0:
            self._plot_rect_override = None
            return
        
        self._plot_rect_override = (float(l), float(t), float(w), float(h))
        self.rectUpdated.emit(float(l), float(t), float(w), float(h))
        self.update()

    def clear_plot_rect(self):
        self._plot_rect_override = None

    def _plot_rect(self) -> QRectF:
        """计算最终的生效渲染边界。如果底层精确回传失败，则使用边距相减作为 Fallback 兜底方案。"""
        if self._plot_rect_override is not None:
            l, t, w, h = self._plot_rect_override
            if w > 0 and h > 0:
                return QRectF(l, t, w, h)
                
        # Fallback 兜底计算
        w = float(self.width())
        h = float(self.height())
        pl = max(0.0, self.margin_l)
        pr = max(pl + 1.0, w - max(0.0, self.margin_r))
        pt = max(0.0, self.margin_t)
        pb = max(pt + 1.0, h - max(0.0, self.margin_b))
        return QRectF(pl, pt, pr - pl, pb - pt)

    def _x_to_px(self, x: float, rect: QRectF) -> float:
        """【坐标映射】将 X 轴的数据值转化为屏幕中的物理 X 像素坐标。"""
        if self.xmax <= self.xmin: 
            return rect.left()
        t = (float(x) - self.xmin) / (self.xmax - self.xmin)
        t = max(0.0, min(1.0, t))
        return rect.left() + t * rect.width()

    def _y_to_px(self, y: float, rect: QRectF) -> float:
        """【坐标映射】将 Y 轴的数据值转化为屏幕中的物理 Y 像素坐标（注意屏幕 Y 轴是向下的）。"""
        if self.ymax <= self.ymin: 
            return rect.bottom()
        t = (float(y) - self.ymin) / (self.ymax - self.ymin)
        t = max(0.0, min(1.0, t))
        return rect.bottom() - t * rect.height()

    def paintEvent(self, event):
        """核心绘制入口：执行十字线与象限百分比面板的绘制。"""
        rect = self._plot_rect()
        if rect.width() <= 1 or rect.height() <= 1:
            return
            
        px_x = self._x_to_px(self.x_val, rect)
        px_y = self._y_to_px(self.y_val, rect)
            
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing, True) 
        painter.setClipRect(rect)
        
        # 1. 绘制黑色极细十字线
        pen_w = 0.5     
        pen = QPen(QColor("#000000"))
        pen.setWidthF(pen_w)
        pen.setStyle(Qt.SolidLine)
        painter.setPen(pen)

        painter.drawLine(QPointF(px_x, rect.top()), QPointF(px_x, rect.bottom()))
        painter.drawLine(QPointF(rect.left(), px_y), QPointF(rect.right(), px_y))

        if not self.quadrant_stats:
            painter.end()
            return
  
        # 2. 绘制四个角的象限百分比
        font = painter.font()
        font.setFamily("Arial")
        font.setPixelSize(12)
        painter.setFont(font)
        fm = painter.fontMetrics()

        pad_x, pad_y = 4, 2
        line_height = fm.height() + pad_y * 2
        margin = 1  
        
        # 最多绘制前 4 组叠图数据的占比，避免画面过于拥挤
        draw_stats = self.quadrant_stats[:4]  
        count = len(draw_stats)

        for i, stat in enumerate(draw_stats):
            q1, q2, q3, q4 = stat
            # 严格对齐 DRISK 的全局情景循环色板
            bg_color = QColor(DRISK_COLOR_CYCLE[i % len(DRISK_COLOR_CYCLE)])
            
            text_q1 = f"{q1 * 100:.1f}%"
            text_q2 = f"{q2 * 100:.1f}%"
            text_q3 = f"{q3 * 100:.1f}%"
            text_q4 = f"{q4 * 100:.1f}%"

            w_q1 = fm.horizontalAdvance(text_q1) + pad_x * 2
            w_q2 = fm.horizontalAdvance(text_q2) + pad_x * 2
            w_q3 = fm.horizontalAdvance(text_q3) + pad_x * 2
            w_q4 = fm.horizontalAdvance(text_q4) + pad_x * 2

            y_top = rect.top() + margin + i * (line_height + 1)
            y_bot = rect.bottom() - margin - (count - i) * (line_height + 1)
            
            x_left = rect.left() + margin
            x_q1_right = rect.right() - margin - w_q1
            x_q4_right = rect.right() - margin - w_q4

            def draw_box(x, y, w, text):
                """内部闭包函数：绘制带底色的小方块文本标签。"""
                bg_rect = QRectF(x, y, w, line_height)
                painter.setPen(Qt.NoPen)
                painter.setBrush(bg_color)
                painter.drawRect(bg_rect)
                painter.setPen(QColor("#FFFFFF")) 
                painter.drawText(bg_rect, Qt.AlignCenter, text)

            draw_box(x_q1_right, y_top, w_q1, text_q1) # 第一象限 (右上)
            draw_box(x_left, y_top, w_q2, text_q2)     # 第二象限 (左上)
            draw_box(x_left, y_bot, w_q3, text_q3)     # 第三象限 (左下)
            draw_box(x_q4_right, y_bot, w_q4, text_q4) # 第四象限 (右下)

        painter.end()


class HorizontalSingleSlider(QWidget):
    """
    X 轴底部的水平控制滑块。
    负责控制 X 轴十字交叉线的物理基准值，提供鼠标按压和移动的防抖拖拽交互（Debounce）。
    """
    valueChanged = Signal(float)
    dragFinished = Signal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setFixedHeight(32) 
        self.margin_left = 10
        self.margin_right = 10
        self._plot_w = 0.0
        self._min = 0.0
        self._max = 100.0
        self._val = 50.0
        self._dragging = False
        self.handle_color = QColor("#333333")

    def setMargins(self, left, right):
        self.margin_left = float(left)
        self.margin_right = float(right)
        self.update()

    def setRangeLimit(self, min_val, max_val):
        self._min = float(min_val)
        self._max = float(max_val)
        self.update()

    def setValue(self, val):
        self._val = max(self._min, min(float(val), self._max))
        self.update()

    def setPlotWidth(self, w: float):
        """接收实际的图表内部宽度，并更新滑块的总长跨度。"""
        self._plot_w = float(w)
        self.update()

    def get_valid_width(self):
        if self._plot_w > 0:
            return self._plot_w
        return max(1, self.width() - self.margin_left - self.margin_right)

    def val_to_pixel(self, val):
        w = self.get_valid_width()
        if w <= 0 or abs(self._max - self._min) < 1e-9: return self.margin_left
        t = (val - self._min) / (self._max - self._min)
        return self.margin_left + t * w

    def pixel_to_val(self, px):
        w = self.get_valid_width()
        if w <= 0: return self._min
        t = (px - self.margin_left) / w
        val = self._min + t * (self._max - self._min)
        return max(self._min, min(val, self._max))

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)
        
        # 绘制滑轨背景底线
        painter.setPen(QPen(QColor(0, 0, 0, 0), 1, Qt.SolidLine, Qt.FlatCap))
        track_y = self.height() - 2
        painter.drawLine(int(self.margin_left - 1), int(track_y), int(self.margin_left + self.get_valid_width()), int(track_y))

        # 绘制滑块把手 (朝上的小三角形)
        x_px = self.val_to_pixel(self._val)
        painter.setPen(Qt.NoPen)
        painter.setBrush(QBrush(self.handle_color))
        size = 12
        hw, hh = size * 0.6, size * 0.6
        y_base = self.height() 
        poly = QPolygonF([
            QPointF(x_px, y_base),                  
            QPointF(x_px - hw, y_base - hh * 2),    
            QPointF(x_px + hw, y_base - hh * 2)     
        ])
        painter.drawPolygon(poly)

    def mousePressEvent(self, event):
        self.setFocus()
        if event.button() == Qt.LeftButton:
            self._dragging = True
            self._update_from_mouse(event.position().x())

    def mouseMoveEvent(self, event):
        if self._dragging:
            self._update_from_mouse(event.position().x())

    def mouseReleaseEvent(self, event):
        """拖拽释放时抛出信号，通常在这里触发计算量较大的统计指标面板刷新。"""
        if self._dragging:
            self._dragging = False
            self.dragFinished.emit()

    def _update_from_mouse(self, x_px):
        self._val = self.pixel_to_val(x_px)
        self.update()
        self.valueChanged.emit(self._val)


class VerticalSingleSlider(QWidget):
    """
    Y 轴左侧的垂直控制滑块。
    包含倒置的 Y 轴物理像素映射逻辑（屏幕坐标向下为正，图表数据向上为正）。
    与 HorizontalSingleSlider 接口保持高度统一。
    """
    
    valueChanged = Signal(float)
    dragFinished = Signal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setFocusPolicy(Qt.ClickFocus)
        self.setFixedWidth(65)
        self.margin_top = 10
        self.margin_bottom = 10
        self._plot_h = 0.0
        self._min = 0.0
        self._max = 100.0
        self._val = 50.0
        self._dragging = False
        self.handle_color = QColor("#333333")

    def setMargins(self, top, bottom):
        self.margin_top = float(top)
        self.margin_bottom = float(bottom)
        self.update()

    def setRangeLimit(self, min_val, max_val):
        self._min = float(min_val)
        self._max = float(max_val)
        self.update()

    def setValue(self, val):
        self._val = max(self._min, min(float(val), self._max))
        self.update()

    def setPlotHeight(self, h: float):
        self._plot_h = float(h)
        self.update()

    def get_valid_height(self):
        if self._plot_h > 0:
            return self._plot_h
        return max(1, self.height() - self.margin_top - self.margin_bottom)

    def val_to_pixel(self, val):
        h = self.get_valid_height()
        if h <= 0 or abs(self._max - self._min) < 1e-9: 
            return self.margin_top
        t = (val - self._min) / (self._max - self._min)
        # 反转 Y 轴方向
        return self.margin_top + h - (t * h)

    def pixel_to_val(self, px):
        h = self.get_valid_height()
        if h <= 0: 
            return self._min
        t = (self.margin_top + h - px) / h
        val = self._min + t * (self._max - self._min)
        return max(self._min, min(val, self._max))
    
    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)
        
        # 将垂直滑块轨道紧贴靠左
        painter.setPen(QPen(QColor(0, 0, 0, 0), 1, Qt.SolidLine, Qt.FlatCap))
        track_x = 2 
        painter.drawLine(int(track_x), int(self.margin_top), int(track_x), int(self.margin_top + self.get_valid_height() + 1))

        y_px = self.val_to_pixel(self._val)
        
        # 绘制滑块把手 (朝右的小三角形)
        painter.setPen(Qt.NoPen)
        painter.setBrush(QBrush(self.handle_color))
        size = 12
        hw, hh = size * 0.6, size * 0.6
        
        poly = QPolygonF([
            QPointF(0, y_px),                
            QPointF(hw * 2, y_px - hh),      
            QPointF(hw * 2, y_px + hh)       
        ])
        painter.drawPolygon(poly)

    def mousePressEvent(self, event):
        self.setFocus()
        if event.button() == Qt.LeftButton:
            self._dragging = True
            self._update_from_mouse(event.position().y())

    def mouseMoveEvent(self, event):
        if self._dragging:
            self._update_from_mouse(event.position().y())

    def mouseReleaseEvent(self, event):
        if self._dragging:
            self._dragging = False
            self.dragFinished.emit()

    def _update_from_mouse(self, y_px):
        self._val = self.pixel_to_val(y_px)
        self.update()
        self.valueChanged.emit(self._val)


# =======================================================
# 2. 数据与配置选择弹窗类 (交互控制与前置路由)
# =======================================================

class ScatterAnalysisObjectDialog(QDialog):
    """
    分析对象控制弹窗：
    供用户切换当前分析的 Y 轴变量到底绑定在哪个版本的模拟情景上。
    同时提供了一个“选取其他数据”的后门，将用户打回至初始选区配置窗口重选数据。
    """
    requestSelectOtherData = Signal()

    def __init__(self, current_sim_id, y_addr, y_name, available_sims, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Drisk - 分析对象")
        set_drisk_icon(self)
        self.resize(320, 380)

        self.current_sim_id = str(current_sim_id)
        self.available_sims = [str(sid) for sid in available_sims]

        # 统一注入标准样式
        self.setStyleSheet("""
            QDialog { background-color: #f9f9f9; font-family: 'Microsoft YaHei'; }
            QLabel { font-weight: bold; color: #444; font-size: 12px; }
            QListWidget { 
                border: 1px solid #ccc; 
                border-radius: 3px; 
                outline: none; 
                font-size: 12px; 
                background-color: #ffffff;
            }
            QListWidget::item { padding: 8px 6px; border-bottom: 1px solid #f5f5f5; }
            QListWidget::item:hover { background-color: #f0f7ff; }
            QListWidget::item:selected { 
                background-color: #f0f4f8; 
                color: #334455; 
                font-weight: bold; 
                border-left: 3px solid #8c9eb5;
            }
            QPushButton { 
                background-color: white; border: 1px solid #ccc; border-radius: 3px; 
                padding: 4px 12px; font-size: 12px; 
            }
            QPushButton:hover { color: #40a9ff; border-color: #40a9ff; background-color: #f0f7ff;}
            QPushButton#btnOk { background-color: #0050b3; color: white; border: none; }
            QPushButton#btnOther { color: #0050b3; border-color: #91caff; background-color: #e6f7ff; }
            QPushButton#btnOther:hover { color: white; background-color: #40a9ff; }
        """)

        layout = QVBoxLayout(self)
        layout.setSpacing(10)

        lbl_info = QLabel(f"当前变量：{y_name} ({y_addr})")
        lbl_info.setStyleSheet("color: #0050b3; font-size: 13px; margin-bottom: 5px;")
        layout.addWidget(lbl_info)
        layout.addWidget(QLabel("切换当前变量的模拟情景版本："))

        # 情景单选列表
        self.list_sims = QListWidget()
        self.list_sims.setSelectionMode(QListWidget.SingleSelection)
        self.list_sims.setContextMenuPolicy(Qt.CustomContextMenu)
        self.list_sims.customContextMenuRequested.connect(self._open_sim_context_menu)
        layout.addWidget(self.list_sims)

        self._refresh_sim_list()

        # 底部按钮组
        btn_layout = QHBoxLayout()
        self.btn_other = QPushButton("选取其他数据")
        self.btn_other.setObjectName("btnOther")
        self.btn_other.clicked.connect(self._on_other_data_clicked)
        btn_layout.addWidget(self.btn_other)
        
        btn_layout.addStretch()
        
        self.btn_ok = QPushButton("确定")
        self.btn_ok.setObjectName("btnOk")
        self.btn_ok.setFixedWidth(60)  
        self.btn_ok.clicked.connect(self.accept)
        
        btn_cancel = QPushButton("取消")
        btn_cancel.setFixedWidth(60)  
        btn_cancel.clicked.connect(self.reject)
        
        btn_layout.addWidget(self.btn_ok)
        btn_layout.addWidget(btn_cancel)
        layout.addLayout(btn_layout)

    def _on_other_data_clicked(self):
        """抛出重新选区的信号，交由主窗口隐藏自身并唤起 Excel 选区框。"""
        self.requestSelectOtherData.emit()
        self.reject()

    def _refresh_sim_list(self, selected_sid: Optional[str] = None):
        """拉取有效情景数据并刷新 ListView，保持当前激活项。"""
        selected_text = str(selected_sid) if selected_sid is not None else None
        if selected_text is None:
            items = self.list_sims.selectedItems()
            if items:
                selected_text = str(items[0].data(Qt.UserRole))
        self.list_sims.clear()
        
        # 按情景 ID 数字排序
        sorted_sims = sorted(self.available_sims, key=lambda x: int(x) if x.isdigit() else x)
        for sid in sorted_sims:
            sim_label = get_sim_display_name(sid)
            is_current = sid == self.current_sim_id
            display_text = f"{sim_label}（当前）" if is_current else sim_label
            item = QListWidgetItem(display_text)
            item.setData(Qt.UserRole, sid)
            
            if is_current:
                font = item.font()
                font.setBold(True)
                item.setFont(font)
            self.list_sims.addItem(item)
            
            if selected_text is not None and sid == selected_text:
                item.setSelected(True)
            elif selected_text is None and is_current:
                item.setSelected(True)

    def _open_sim_context_menu(self, pos):
        """通过右键菜单直接唤出情景重命名功能。"""
        item = self.list_sims.itemAt(pos)
        if item is None:
            return
        sid = item.data(Qt.UserRole)
        if sid is None:
            return
        
        menu = QMenu(self)
        act_rename = menu.addAction("重命名情景")
        chosen = menu.exec(self.list_sims.mapToGlobal(pos))
        if chosen != act_rename:
            return
            
        if _rename_sim_with_dialog(self, sid):
            self._refresh_sim_list(selected_sid=str(sid))

    def get_selected_sim(self):
        """返回被选定的目标模拟情景 ID。"""
        items = self.list_sims.selectedItems()
        if items:
            return items[0].data(Qt.UserRole)
        return None


class ScatterOverlayDialog(QDialog):
    """
    叠加对比配置弹窗：
    提供列表框供用户多选其它历史情景版本，这些勾选项将在后端组装成独立的数据流共同送入同一张散点图进行渲染。
    """
    def __init__(self, current_sim_id, y_addr, y_name, available_sims, current_overlays, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Drisk - 叠加对比")
        set_drisk_icon(self)  
        self.resize(320, 380)
        
        self.current_sim_id = str(current_sim_id)
        self.available_sims = [str(sid) for sid in available_sims]
        self.current_overlays = [str(sid) for sid in current_overlays]
        
        base_dir = os.path.dirname(os.path.abspath(__file__))
        icon_path = os.path.join(base_dir, "icons", "Selected_Overlay.svg").replace('\\', '/')

        # 使用基于图片的自定义复选框样式
        self.setStyleSheet("""
            QDialog { background-color: #f9f9f9; font-family: 'Microsoft YaHei'; }
            QLabel { font-weight: bold; color: #444; font-size: 12px; }
            QListWidget { 
                border: 1px solid #ccc; 
                border-radius: 3px; 
                outline: none; 
                font-size: 12px; 
                background-color: #ffffff;
            }
            QListWidget::item { padding: 8px 6px; border-bottom: 1px solid #f5f5f5; }
            QListWidget::item:hover { background-color: #f0f7ff; }
            QPushButton { 
                background-color: white; border: 1px solid #ccc; border-radius: 3px; 
                padding: 4px 12px; font-size: 12px; 
            }
            QPushButton:hover { color: #40a9ff; border-color: #40a9ff; background-color: #f0f7ff;}
            QPushButton#btnOk { background-color: #0050b3; color: white; border: none; }
            QCheckBox { font-size: 12px; color: #333; }

            QListWidget::indicator, QCheckBox::indicator {
                width: 16px;
                height: 16px;
                border-radius: 3px;       
                border: 1px solid #cbd0d6; 
                background-color: #ffffff;
            }
            
            QListWidget::indicator:hover, QCheckBox::indicator:hover {
                border: 1px solid #8c9eb5; 
            }
            
            QListWidget::indicator:checked, QCheckBox::indicator:checked {
                background-color: #8c9eb5; 
                border: 1px solid #8c9eb5;
                image: url('""" + icon_path + """');
            }
            
            QListWidget::item:selected { 
                background-color: #f0f4f8; 
                color: #334455; 
                font-weight: bold; 
                border-left: 3px solid #8c9eb5;
            }
        """)
        
        layout = QVBoxLayout(self)
        layout.setSpacing(10)
        
        lbl_info = QLabel(f"当前变量：{y_name} ({y_addr})")
        lbl_info.setStyleSheet("color: #0050b3; font-size: 13px; margin-bottom: 5px;")
        layout.addWidget(lbl_info)
        layout.addWidget(QLabel("勾选需要叠加对比的其它模拟情景："))
        
        self.chk_all = QCheckBox("全选其它情景")
        self.chk_all.stateChanged.connect(self._on_check_all)
        layout.addWidget(self.chk_all)
        
        self.list_sims = QListWidget()
        self.list_sims.setContextMenuPolicy(Qt.CustomContextMenu)
        self.list_sims.customContextMenuRequested.connect(self._open_sim_context_menu)
        layout.addWidget(self.list_sims)
        self._refresh_sim_list()
            
        btn_layout = QHBoxLayout()
        btn_layout.addStretch()
        self.btn_ok = QPushButton("确定")
        self.btn_ok.setObjectName("btnOk")
        self.btn_ok.setFixedWidth(60)  
        self.btn_ok.clicked.connect(self.accept)
        
        btn_cancel = QPushButton("取消")
        btn_cancel.setFixedWidth(60)  
        btn_cancel.clicked.connect(self.reject)
        
        btn_layout.addWidget(self.btn_ok)
        btn_layout.addWidget(btn_cancel)
        layout.addLayout(btn_layout)

    def _refresh_sim_list(self):
        """载入列表并剔除当前正在作为底版的 sim ID。"""
        selected = set(str(sid) for sid in self.current_overlays)
        self.list_sims.clear()
        
        valid_sims = [s for s in self.available_sims if s != self.current_sim_id]
        valid_sims = sorted(valid_sims, key=lambda x: int(x) if x.isdigit() else x)
        
        for sid in valid_sims:
            sim_label = get_sim_display_name(sid)
            item = QListWidgetItem(sim_label)
            item.setData(Qt.UserRole, sid)
            item.setFlags(item.flags() | Qt.ItemIsUserCheckable)
            
            if sid in selected:
                item.setCheckState(Qt.Checked)
            else:
                item.setCheckState(Qt.Unchecked)
                # Fallback: 若全空且存在 sim 1，则默认勾上 sim 1 以作为参照系
                if not selected and sid == "1":
                    item.setCheckState(Qt.Checked)
            self.list_sims.addItem(item)

    def _open_sim_context_menu(self, pos):
        item = self.list_sims.itemAt(pos)
        if item is None:
            return
        sid = item.data(Qt.UserRole)
        if sid is None:
            return
            
        menu = QMenu(self)
        act_rename = menu.addAction("重命名情景")
        chosen = menu.exec(self.list_sims.mapToGlobal(pos))
        if chosen != act_rename:
            return
            
        if _rename_sim_with_dialog(self, sid):
            self.current_overlays = self.get_selected_sims()
            self._refresh_sim_list()

    def _on_check_all(self, state):
        """全选与反选级联触发器。"""
        check_state = Qt.Checked if state == Qt.Checked.value else Qt.Unchecked
        for i in range(self.list_sims.count()):
            self.list_sims.item(i).setCheckState(check_state)
            
    def get_selected_sims(self):
        """遍历 UI，返回所有已勾选态的叠加项列表。"""
        return [self.list_sims.item(i).data(Qt.UserRole) for i in range(self.list_sims.count()) if self.list_sims.item(i).checkState() == Qt.Checked]

"""Drisk - 散点图设置"""
class ScatterClipSettingsDialog(QDialog):
    """
    数据裁剪设置窗：
    允许用户手动设置 X 和 Y 轴数值截断上下界。
    裁剪发生于前端提交数据给 Plotly 引擎之前，所以裁剪能够同时影响并刷新右侧的统计分析数据。
    """

    def __init__(self, current_config: Optional[Dict[str, Any]] = None, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Drisk - 散点图设置")
        set_drisk_icon(self)
        self.resize(430, 290)
        self._config = self.normalize_config(current_config)
        self._init_ui()
        self._load_from_config()

    @staticmethod
    def _normalize_axis_config(axis_cfg: Any) -> Dict[str, Any]:
        """将用户输入的文本类型边界清洗为标准浮点数，执行类型防呆保护。"""
        axis_cfg = axis_cfg if isinstance(axis_cfg, dict) else {}
        enabled = bool(axis_cfg.get("enabled", False))
        lower = axis_cfg.get("lower", None)
        upper = axis_cfg.get("upper", None)
        try:
            lower = None if lower in (None, "") else float(lower)
        except Exception:
            lower = None
        try:
            upper = None if upper in (None, "") else float(upper)
        except Exception:
            upper = None
        return {"enabled": enabled, "lower": lower, "upper": upper}

    @classmethod
    def normalize_config(cls, config: Any) -> Dict[str, Dict[str, Any]]:
        config = config if isinstance(config, dict) else {}
        return {
            "x": cls._normalize_axis_config(config.get("x")),
            "y": cls._normalize_axis_config(config.get("y")),
        }

    def _init_ui(self):
        self.setStyleSheet("""
            QDialog { background-color: #f9f9f9; font-family: 'Microsoft YaHei'; }
            QLabel { color: #333; font-size: 12px; }
            QGroupBox {
                border: 1px solid #d9d9d9;
                border-radius: 4px;
                margin-top: 8px;
                padding-top: 12px;
                font-weight: bold;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 8px;
                padding: 0 4px;
                color: #444;
            }
            QLineEdit {
                background-color: #ffffff;
                border: 1px solid #d9d9d9;
                border-radius: 3px;
                padding: 4px 6px;
            }
            QPushButton {
                background-color: white;
                border: 1px solid #ccc;
                border-radius: 3px;
                padding: 4px 12px;
                font-size: 12px;
            }
            QPushButton:hover { color: #40a9ff; border-color: #40a9ff; background-color: #f0f7ff; }
            QPushButton#btnOk { background-color: #0050b3; color: white; border: none; }
            QPushButton#btnReset { color: #8c8c8c; }
        """)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(12, 10, 12, 10)
        layout.setSpacing(8)

        info = QLabel("设置散点图数据裁剪条件。启用后将过滤超出边界的点。")
        info.setStyleSheet("color: #555;")
        layout.addWidget(info)

        x_group, self.chk_x_enabled, self.edit_x_low, self.edit_x_high = self._build_axis_group("X 轴裁剪")
        y_group, self.chk_y_enabled, self.edit_y_low, self.edit_y_high = self._build_axis_group("Y 轴裁剪")
        layout.addWidget(x_group)
        layout.addWidget(y_group)
        layout.addStretch()

        btn_layout = QHBoxLayout()
        self.btn_reset = QPushButton("重置")
        self.btn_reset.setObjectName("btnReset")
        self.btn_reset.clicked.connect(self._reset_fields)
        btn_layout.addWidget(self.btn_reset)
        btn_layout.addStretch()

        self.btn_ok = QPushButton("确定")
        self.btn_ok.setObjectName("btnOk")
        self.btn_ok.setFixedWidth(60)
        self.btn_cancel = QPushButton("取消")
        self.btn_cancel.setFixedWidth(60)
        
        self.btn_ok.clicked.connect(self._on_accept)
        self.btn_cancel.clicked.connect(self.reject)
        
        btn_layout.addWidget(self.btn_ok)
        btn_layout.addWidget(self.btn_cancel)
        layout.addLayout(btn_layout)

    def _build_axis_group(self, title: str):
        group = QGroupBox(title)
        form = QFormLayout(group)
        form.setLabelAlignment(Qt.AlignRight)
        form.setFormAlignment(Qt.AlignTop)

        chk_enabled = QCheckBox("启用该轴裁剪")
        form.addRow("", chk_enabled)

        edit_lower = QLineEdit()
        edit_lower.setPlaceholderText("留空表示不限制下界")
        form.addRow("下界", edit_lower)

        edit_upper = QLineEdit()
        edit_upper.setPlaceholderText("留空表示不限制上界")
        form.addRow("上界", edit_upper)

        # 联动禁用编辑框
        chk_enabled.toggled.connect(lambda checked, lo=edit_lower, hi=edit_upper: self._toggle_axis_inputs(checked, lo, hi))
        return group, chk_enabled, edit_lower, edit_upper

    def _toggle_axis_inputs(self, enabled: bool, lower_edit: QLineEdit, upper_edit: QLineEdit):
        lower_edit.setEnabled(bool(enabled))
        upper_edit.setEnabled(bool(enabled))

    def _load_from_config(self):
        """回显传入的裁剪设置。"""
        x_cfg = self._config.get("x", {})
        y_cfg = self._config.get("y", {})
        self.chk_x_enabled.setChecked(bool(x_cfg.get("enabled", False)))
        self.chk_y_enabled.setChecked(bool(y_cfg.get("enabled", False)))

        self.edit_x_low.setText("" if x_cfg.get("lower") is None else str(x_cfg.get("lower")))
        self.edit_x_high.setText("" if x_cfg.get("upper") is None else str(x_cfg.get("upper")))
        self.edit_y_low.setText("" if y_cfg.get("lower") is None else str(y_cfg.get("lower")))
        self.edit_y_high.setText("" if y_cfg.get("upper") is None else str(y_cfg.get("upper")))

        self._toggle_axis_inputs(self.chk_x_enabled.isChecked(), self.edit_x_low, self.edit_x_high)
        self._toggle_axis_inputs(self.chk_y_enabled.isChecked(), self.edit_y_low, self.edit_y_high)

    def _reset_fields(self):
        self.chk_x_enabled.setChecked(False)
        self.chk_y_enabled.setChecked(False)
        self.edit_x_low.clear()
        self.edit_x_high.clear()
        self.edit_y_low.clear()
        self.edit_y_high.clear()

    def _parse_bound(self, text: str, axis_name: str, bound_name: str) -> Optional[float]:
        raw = (text or "").strip()
        if not raw:
            return None
        try:
            return float(raw)
        except Exception:
            QMessageBox.warning(self, "输入无效", f"{axis_name}{bound_name}必须是数值。")
            return None

    def _collect_axis_config(self, axis_label: str, enabled: bool, low_edit: QLineEdit, high_edit: QLineEdit):
        """严格校验提取用户的输入，若发现上下界填写反了，抛出弹窗拦截。"""
        low = self._parse_bound(low_edit.text(), axis_label, "下界")
        if low_edit.text().strip() and low is None:
            return None
        high = self._parse_bound(high_edit.text(), axis_label, "上界")
        if high_edit.text().strip() and high is None:
            return None

        if enabled and low is None and high is None:
            QMessageBox.warning(self, "设置无效", f"{axis_label}已启用裁剪，请至少填写一个边界。")
            return None
        if low is not None and high is not None and low > high:
            QMessageBox.warning(self, "设置无效", f"{axis_label}下界不能大于上界。")
            return None

        return {"enabled": bool(enabled), "lower": low, "upper": high}

    def _on_accept(self):
        x_cfg = self._collect_axis_config("X 轴", self.chk_x_enabled.isChecked(), self.edit_x_low, self.edit_x_high)
        if x_cfg is None:
            return
        y_cfg = self._collect_axis_config("Y 轴", self.chk_y_enabled.isChecked(), self.edit_y_low, self.edit_y_high)
        if y_cfg is None:
            return

        self._config = {"x": x_cfg, "y": y_cfg}
        self.accept()

    def get_config(self) -> Dict[str, Dict[str, Any]]:
        return self.normalize_config(self._config)


class ScatterConfigDialog(QDialog):
    """
    初始数据源选区配置对话框：
    利用 pyxll 底层通信，将交互推向 Excel 本体。
    弹窗允许用户直接在 Excel 中使用鼠标框选需要分析的 Y 轴（单变量）和 X 轴（可按 Ctrl 多选）数据区域。
    """
    
    def __init__(self, sim_id, parent=None):
        super().__init__(parent)
        self.sim_id = sim_id
        
        self.selected_y_cells = []
        self.selected_x_cells = []
        
        self.y_addr = None
        self.x_addrs = []

        self.setWindowTitle("Drisk - 散点分析")
        set_drisk_icon(self)  
        _sync_application_icon_from(self)
        self.resize(500, 220)
        self._init_ui()

    def _init_ui(self):
        """左右分栏的配置布局，点击选择靶心图标时触发 Excel 拦截器。"""
        main_layout = QVBoxLayout(self)
        main_layout.setSpacing(15)

        lbl_info = QLabel("选择要在散点图中显示的变量。\nY 轴可绘制单一变量，X 轴可绘制一个或多个变量。")
        lbl_info.setStyleSheet("color: #333333; font-size: 13px; margin-bottom: 10px;")
        main_layout.addWidget(lbl_info)

        # Y 轴选择区
        y_layout = QHBoxLayout()
        y_label = QLabel("Y 轴变量")
        y_label.setFixedWidth(140)
        self.y_edit = QLineEdit()
        self.y_edit.setReadOnly(True)
        self.y_edit.setStyleSheet("background-color: #f5f5f5; border: 1px solid #d9d9d9; padding: 4px;")
        
        self.y_btn = QPushButton()
        self.y_btn.setFixedSize(30, 26)
        self.y_btn.setToolTip("从 Excel 选择一个单元格作为 Y 轴变量")
        if not apply_excel_select_button_icon(self.y_btn, "select_icon.svg"):
            self.y_btn.setText("\U0001F3AF") # Fallback emoji if icon fails
        self.y_btn.clicked.connect(lambda: self._select_range(self.selected_y_cells, self.y_edit, single=True))
        
        y_layout.addWidget(y_label)
        y_layout.addWidget(self.y_edit)
        y_layout.addWidget(self.y_btn)
        main_layout.addLayout(y_layout)

        # X 轴选择区
        x_layout = QHBoxLayout()
        x_label = QLabel("X 轴数据（可多选）")
        x_label.setFixedWidth(140)
        self.x_edit = QLineEdit()
        self.x_edit.setReadOnly(True)
        self.x_edit.setStyleSheet("background-color: #f5f5f5; border: 1px solid #d9d9d9; padding: 4px;")
        
        self.x_btn = QPushButton()
        self.x_btn.setFixedSize(30, 26)
        self.x_btn.setToolTip("从 Excel 选择一个或多个单元格作为 X 轴变量")
        if not apply_excel_select_button_icon(self.x_btn, "select_icon.svg"):
            self.x_btn.setText("\U0001F3AF")
        self.x_btn.clicked.connect(lambda: self._select_range(self.selected_x_cells, self.x_edit, single=False))
        
        x_layout.addWidget(x_label)
        x_layout.addWidget(self.x_edit)
        x_layout.addWidget(self.x_btn)
        main_layout.addLayout(x_layout)

        main_layout.addStretch()

        btn_layout = QHBoxLayout()
        btn_layout.addStretch()
        self.btn_ok = QPushButton("确定")
        self.btn_cancel = QPushButton("取消")
        self.btn_ok.setFixedWidth(60)
        self.btn_cancel.setFixedWidth(60)
        
        self.btn_cancel.setStyleSheet("""
            background-color: white; border: 1px solid #ccc; 
            padding: 4px 0px; border-radius: 3px; font-size: 12px;
        """)
        self.btn_ok.setStyleSheet("""
            background-color: #0050b3; color: white; border: none; 
            padding: 4px 0px; border-radius: 3px; font-size: 12px; font-weight: bold;
        """)
        
        self.btn_cancel.clicked.connect(self.reject)
        self.btn_ok.clicked.connect(self._check_and_accept)
        
        btn_layout.addWidget(self.btn_ok)
        btn_layout.addWidget(self.btn_cancel)
        main_layout.addLayout(btn_layout)

    def _select_range(self, target_list, target_lineedit, single=False):
        """调用 pyxll 的 InputBox 弹出 Excel 选区，解析用户的选框结果并将单元格地址提取到行编辑框。"""
        app = xl_app()

        # 选区时，将当前 Qt 弹窗完全透明化隐藏，使用户能够看清下层的 Excel 表格
        self.setWindowOpacity(0)
        try:
            prompt = "请选择一个单元格（Y轴）:" if single else "请选择一个或多个单元格（X轴，按住 Ctrl 多选）:"
            # Type=8 是 Excel Range 对象的特有标识
            rng = app.InputBox(Prompt=prompt, Title="选择数据单元格", Type=8)
            
            if rng:
                cells = []
                for area in rng.Areas:
                    for cell in area.Cells:
                        sheet = cell.Worksheet.Name
                        addr = cell.Address.replace('$', '')
                        cells.append(f"{sheet}!{addr}")
                        if single: break
                    if single and cells: break
                
                target_list.clear()
                target_list.extend(cells)
                target_lineedit.setText(", ".join(cells))
        except Exception:
            pass
        finally:
            self.setWindowOpacity(1)
            self.raise_()
            self.activateWindow()

    def _check_and_accept(self):
        """严格拦截，确认用户选好了有效格子再放行进入主分析台。"""
        if not self.selected_y_cells:
            QMessageBox.warning(self, "提示", "请选择 Y 轴数据单元格。")
            return
        if not self.selected_x_cells:
            QMessageBox.warning(self, "提示", "请至少选择一个 X 轴数据单元格。")
            return
        
        self.y_addr = self.selected_y_cells[0]
        self.x_addrs = self.selected_x_cells
        self.accept()


# =======================================================
# 3. 核心散点图主窗口视图类 (集成大心脏)
# =======================================================
class ScatterPlotlyDialog(QDialog):
    """
    核心散点图分析台主窗口。
    此控件整合了左右滑动分割、异步 JS 注入引擎交互、图表状态树维护以及统计指标表的自动演算。
    """

    # ---------------------------------------------------
    # Part 3.1: 初始化与基础属性绑定
    # ---------------------------------------------------
    def __init__(self, initial_sim_id, y_addr, x_addrs, parent=None):
        super().__init__(parent)
        self.setWindowFlags(
            self.windowFlags()
            | Qt.WindowMinimizeButtonHint
            | Qt.WindowMaximizeButtonHint
            | Qt.WindowCloseButtonHint
        )

        # 挂载核心业务上下文
        try:
            self.sim_id = int(initial_sim_id)
        except Exception:
            self.sim_id = initial_sim_id
        self.y_addr = y_addr
        self.x_addrs = x_addrs

        self.chart_mode = "scatter" 
        self.current_key = y_addr.split('!')[-1] if '!' in y_addr else y_addr

        # 维护绘图的样式、标签与叠加层级状态树
        self.overlay_items = [] 
        self._scatter_clip_config = ScatterClipSettingsDialog.normalize_config(None)
        # 共享的样式记录，与样式管理器互通，以便记住不同序列分配的颜色和符号
        self._scatter_style_map: Dict[str, Dict[str, Any]] = {}
        self._scatter_display_keys: List[str] = []
        
        self._tmp_dir = get_plotly_cache_dir()
        self._name_resolve_excel_app = None
        self._name_resolve_excel_app_ready = False
        
        self.x_mag = None
        self.y_mag = None
        self._label_settings_config = create_default_label_settings_config()
        self._label_axis_numeric = {"x": True, "y": True}
        
        sim = get_simulation(self.sim_id)
        self.y_name, _ = self._get_sim_data_dynamic(self.y_addr, sim)

        set_drisk_icon(self)  
        _sync_application_icon_from(self)
        self._refresh_window_title()

        # 计算并设置适配屏幕的合理初始大小
        screen_geo = QApplication.primaryScreen().availableGeometry()
        target_w = min(1200, int(screen_geo.width() * 0.90))
        target_h = min(800, int(screen_geo.height() * 0.80))
        self.resize(target_w, target_h)
        
        self._init_ui()

        # == 复杂的异步页面加载轮询锁群 == 
        # 防止页面没加载完就强行给前端传 JSON 导致界面崩溃
        self._webview_loaded = False
        self._data_sent = False
        self._loading_mask_visible = True
        self._render_token = 0
        self._waiting_token = 0
        self._js_inflight = False
        self._ready_deadline = 0.0
        self._ready_poll_timer = QTimer(self)
        self._ready_poll_timer.setInterval(50)
        self._ready_poll_timer.timeout.connect(self._poll_plotly_ready)
        try:
            self.web_view.loadFinished.connect(self._on_webview_load_finished)
        except Exception:
            pass

        self._show_skeleton("正在绘图...")
        # 让出 UI 线程给 Qt 把壳子画完，50ms后再去干脏活累活
        QTimer.singleShot(50, self._update_plot)

    def _refresh_window_title(self):
        self.setWindowTitle("Drisk - 散点分析")

    # ---------------------------------------------------
    # Part 3.2: 复杂的 UI 模块拼装
    # ---------------------------------------------------
    def _init_ui(self):
        """
        组建视图躯干：
        - 左侧区 (grid_widget): PlotlyHost Web 画板为核心，四周边缘包抄吸附着滑块和独立遮罩。
        - 右侧区 (right_panel): 统计数据透视表格。
        - 底部区 (bottom_bar): 横向功能控制按钮条。
        """
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)
        
        self.content_splitter = QSplitter(Qt.Horizontal)
        self.content_splitter.setHandleWidth(2)
        self.content_splitter.setStyleSheet("QSplitter::handle { background-color: #c0c0c0; }")
        
        self.grid_widget = QWidget()
        self.grid_widget.setStyleSheet("background-color: white;")
        self.chart_container = self.grid_widget 
        
        grid_layout = QGridLayout(self.grid_widget)
        grid_layout.setContentsMargins(0, 3, 0, 0)
        grid_layout.setSpacing(0)
        
        self.slider_x = HorizontalSingleSlider(self)     
        self.slider_y = VerticalSingleSlider(self)
        
        self.web_view = get_shared_webview(parent_widget=self.grid_widget)
        self._install_chart_context_menu_hook()
        self.crosshair = CrosshairLineOverlay(self.web_view) 
        # 连接桥，将 Qt 的悬浮控件映射并叠加到 WebEngine 上
        self.host = PlotlyHost(
            web_view=self.web_view, 
            tmp_dir=self._tmp_dir, 
            use_qt_overlay=True, 
            overlay=self.crosshair
        )
        self.crosshair.raise_() 
        
        grid_layout.addWidget(self.slider_x, 0, 0, 1, 2) 
        grid_layout.addWidget(self.web_view, 1, 0)
        grid_layout.addWidget(self.slider_y, 0, 1, 2, 1)
        grid_layout.setColumnStretch(0, 1)
        grid_layout.setColumnStretch(1, 0)
        grid_layout.setRowStretch(0, 0)
        grid_layout.setRowStretch(1, 1)
        
        self.slider_x.show()
        self.slider_y.show()
        self.content_splitter.addWidget(self.grid_widget)

        self.chart_skeleton = ChartSkeleton(self.grid_widget)
        self.chart_skeleton.setGeometry(self.grid_widget.rect())
        self.chart_skeleton.hide()
        self.chart_skeleton.raise_()
        
        self.content_splitter.splitterMoved.connect(self._on_splitter_moved)
        self.grid_widget.installEventFilter(self)

        # 给 WebEngine 内部 Chromium 子控件装上 eventFilter，拦截其抢焦点行为
        QTimer.singleShot(100, lambda: self._install_webview_child_filters())

        # == 右侧统计指标面板 ==
        self.right_panel = QFrame()
        self.right_panel.setMinimumWidth(180)
        self.right_panel.setStyleSheet("background-color: #ffffff;")
        right_layout = QVBoxLayout(self.right_panel)
        right_layout.setContentsMargins(0, 0, 0, 0)

        header = QLabel("   统计指标")
        header.setFixedHeight(36)
        header.setStyleSheet("background-color: #fafafa; color: #333; font-weight: bold; font-size: 14px; border-bottom: 1px solid #f0f0f0;")
        right_layout.addWidget(header)

        self.stats_table = QTableWidget()
        self.stats_table.setVerticalScrollMode(QAbstractItemView.ScrollPerPixel)
        self.stats_table.setHorizontalScrollMode(QAbstractItemView.ScrollPerPixel)
        self.stats_table.setColumnCount(1)
        self.stats_table.verticalHeader().hide()
        self.stats_table.horizontalHeader().setVisible(True)
        self.stats_table.setAlternatingRowColors(False) 
        right_layout.addWidget(self.stats_table)
        
        self.content_splitter.addWidget(self.right_panel)
        
        self.content_splitter.setSizes([780, 220])
        self.content_splitter.setStretchFactor(0, 1)
        self.content_splitter.setStretchFactor(1, 0)
        
        main_layout.addWidget(self.content_splitter, stretch=1)
        
        # == 悬浮输入框注入网格容器环境 ==
        self.float_x = create_floating_value_with_mag(self.grid_widget, height=22, opacity=0.95)
        self.float_y = VerticalRotatableValue(self.grid_widget)
        self.float_y._collapse()

        # 交互信使互通
        self.slider_x.valueChanged.connect(self._on_slider_x_moved)
        self.slider_y.valueChanged.connect(self._on_slider_y_moved) 
        self.slider_x.dragFinished.connect(self.update_stats_ui)
        self.slider_y.dragFinished.connect(self.update_stats_ui)
        self.float_y.valueChanged.connect(self._on_edit_y_confirmed_rotated)
        self.float_x.edit.editingFinished.connect(self._on_edit_x_confirmed)

        self.current_x_val = 0.0
        self.current_y_val = 0.0
        self.x_range_min = 0.0
        self.x_range_max = 1.0
        self.y_range_min = 0.0
        self.y_range_max = 1.0
        self._raw_data_cache = None
        
        # Layout 协同防抖系统
        self.overlay_controller = ChartOverlayController(self.web_view, self.crosshair)
        self._layout_refresh = LayoutRefreshCoordinator(
            self,
            frame_cb=self._layout_refresh_frame,
            final_cb=self._layout_refresh_final,
            frame_ms=50,
            settle_ms=180,
        )
        self._layout_refresh_coalesce_timer = QTimer(self)
        self._layout_refresh_coalesce_timer.setSingleShot(True)
        self._layout_refresh_coalesce_timer.setInterval(0)
        self._layout_refresh_coalesce_timer.timeout.connect(self._flush_layout_refresh_request)
        self._last_polled_plot_size = None
        self.crosshair.rectUpdated.connect(self._on_plot_rect_updated)
        
        # == 底部工具条挂载 ==
        bottom_bar = QFrame()
        bottom_bar.setFixedHeight(34)
        bottom_bar.setStyleSheet("""
            QFrame { background-color: #f5f5f5; border-top: 1px solid #e0e0e0; }
            QLabel { font-size: 12px; color: #555; }
        """)
        bottom_layout = QHBoxLayout(bottom_bar)
        bottom_layout.setContentsMargins(15, 1, 15, 1)
        bottom_layout.setSpacing(6)
        
        standard_btn_style = """
            QPushButton { background-color: #f0f0f0; border: 1px solid #d9d9d9; border-radius: 3px; padding: 0px 8px; font-size: 12px; height: 20px; outline: none;}
            QPushButton:hover { border-color: #40a9ff; color: #40a9ff; background-color: #e6f7ff; }
            QPushButton:pressed { background-color: #bae0ff; border-color: #096dd9; color: #096dd9; font-weight: bold; }
        """
        
        self.btn_analysis_obj = QPushButton("分析对象")
        self.btn_analysis_obj.setStyleSheet(standard_btn_style)
        self.btn_analysis_obj.clicked.connect(self.open_analysis_selector)
        bottom_layout.addWidget(self.btn_analysis_obj)

        self.available_sims = self._get_available_sims_for(self.y_addr)
        self.btn_overlay = QPushButton("叠加对比")
        self.btn_overlay.setStyleSheet(standard_btn_style)
        self.btn_overlay.clicked.connect(self.open_overlay_selector)
        bottom_layout.addWidget(self.btn_overlay)

        # 多自变量的情形下屏蔽层级叠加按钮，防止画面逻辑混乱
        if len(self.available_sims) <= 1 or len(self.x_addrs) > 1:
            self.btn_overlay.hide()

        self.btn_settings = QPushButton("设置")
        self.btn_settings.setStyleSheet(standard_btn_style)
        self.btn_settings.clicked.connect(self.open_scatter_settings)
        bottom_layout.addWidget(self.btn_settings)

        self.btn_style = QPushButton("图形样式")
        self.btn_style.setStyleSheet(standard_btn_style)
        self.btn_style.clicked.connect(self.open_scatter_style_manager)
        bottom_layout.addWidget(self.btn_style)

        self.btn_label_settings = QPushButton("文本设置")
        self.btn_label_settings.setStyleSheet(standard_btn_style)
        self.btn_label_settings.clicked.connect(self.open_label_settings_dialog)
        bottom_layout.addWidget(self.btn_label_settings)

        # 统一注入底栏图标及 Tooltip
        toolbar_specs = [
            (self.btn_analysis_obj, "target_icon.svg", "分析对象：切换当前用于分析的目标变量。", 24, 28),
            (self.btn_settings, "scatter_set_icon.svg", "设置：配置散点图裁剪与显示规则。", 24, 28),
            (self.btn_style, "style_icon.svg", "图形样式：配置颜色、线条和图形外观。", 24, 28),
            (self.btn_label_settings, "text_icon.svg", "文本设置：调整轴标题和刻度文本样式。", 24, 28),
            (self.btn_overlay, "overlay_icon.svg", "叠加对比：选择并叠加其它情景进行对比。", 24, 28),
        ]
        for _btn, _icon_name, _tooltip, _icon_px, _btn_px in toolbar_specs:
            apply_toolbar_button_icon(_btn, _icon_name, icon_px=_icon_px, icon_only=True, button_px=_btn_px)
            _btn.setToolTip(_tooltip)

        bottom_layout.addStretch()
        bottom_layout.addSpacing(12)

        self.btn_export = QPushButton("导出")
        self.btn_export.setStyleSheet("""
            QPushButton { 
                background-color: #0050b3; color: white; border: none; border-radius: 3px; font-weight: bold;
                font-size: 12px; padding: 0px 10px; height: 20px; 
            }
            QPushButton:hover { background-color: #40a9ff; }
            QPushButton:pressed { background-color: #003a8c; }
        """)
        self.btn_export.clicked.connect(self._on_export_clicked)
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
        
        main_layout.addWidget(bottom_bar)

    # ---------------------------------------------------
    # Part 3.3: Resize 与 Layout 状态同步防抖机制
    # ---------------------------------------------------
    def _on_plot_rect_updated(self, l: float, t: float, w: float, h: float):
        """
        跨维同步器：
        当纯 HTML 侧的 Plotly 发生边缘挤压（例如轴标签变长）时，回传实际内部 Margin，
        在 Qt 侧通过 MapTo 将屏幕坐标算出，进而让原生的长条形控件滑块完美包裹着图表。
        """
        if not hasattr(self, 'web_view') or not self.web_view: return

        offset_x = self.web_view.mapTo(self.slider_x, QPoint(0, 0))
        offset_y = self.web_view.mapTo(self.slider_y, QPoint(0, 0))

        ml = offset_x.x() + float(l)
        mr = self.slider_x.width() - (offset_x.x() + float(l) + float(w))

        mt = offset_y.y() + float(t)
        mb = self.slider_y.height() - (offset_y.y() + float(t) + float(h))

        self.slider_x.setMargins(ml, mr)
        self.slider_x.setPlotWidth(w)

        self.slider_y.setMargins(mt, mb)
        self.slider_y.setPlotHeight(h)
        
        if hasattr(self, 'crosshair') and self.crosshair:
            self.crosshair.setGeometry(self.web_view.rect())
            self.crosshair.raise_()
            self.crosshair.show() 
            
        self._update_crosshair_ui(recompute_quadrants=False)

    def resizeEvent(self, event):
        """主窗口拉伸事件：引发防抖重刷，避免持续派发 JS 让前端卡死。"""
        super().resizeEvent(event)
        self._request_layout_refresh()

        if not hasattr(self, "_geom_poll_timer"):
            self._geom_poll_timer = QTimer(self)
            self._geom_poll_timer.setInterval(33)
            self._geom_poll_timer.timeout.connect(self._poll_and_sync_geometry)

        if not hasattr(self, "_geom_poll_stop_timer"):
            self._geom_poll_stop_timer = QTimer(self)
            self._geom_poll_stop_timer.setSingleShot(True)
            self._geom_poll_stop_timer.timeout.connect(self._geom_poll_timer.stop)

        self._geom_poll_timer.start()
        self._geom_poll_stop_timer.start(500)

    def _poll_and_sync_geometry(self):
        """在缩放拉伸停止后，派发一次对 gd._fullLayout 的探针以确保彻底收尾。"""
        js = """
        (function(){
            try{
                var gd = document.getElementById('chart_div');
                if(!gd || !gd._fullLayout || !gd._fullLayout._size) return null;
                var sz = gd._fullLayout._size;
                return { w: sz.w, h: sz.h };
            }catch(e){ return null; }
        })();
        """

        def _cb(res):
            if not res or not isinstance(res, dict):
                return
            try:
                w = float(res.get("w", 0))
                h = float(res.get("h", 0))
            except Exception:
                return
            if w <= 0 or h <= 0:
                return
            size_sig = (round(w, 2), round(h, 2))
            if self._last_polled_plot_size == size_sig:
                return
            self._last_polled_plot_size = size_sig
            try:
                self.slider_x.setPlotWidth(w)
                self.slider_y.setPlotHeight(h)
            except Exception:
                pass
            try:
                if hasattr(self, "overlay_controller"):
                    self.overlay_controller.schedule_rect_sync()
            except Exception:
                pass

        if hasattr(self, "web_view") and self.web_view is not None:
            self.web_view.page().runJavaScript(js, _cb)

    def eventFilter(self, obj, event):
        """Resize 拦截 + 阻止 WebEngine 子控件在输入框活跃时抢焦点。"""
        if obj == getattr(self, 'grid_widget', None) and event.type() == QEvent.Resize:
            try:
                if hasattr(self, "chart_skeleton") and self.chart_skeleton:
                    self.chart_skeleton.setGeometry(self.grid_widget.rect())
            except Exception:
                pass
            self._request_layout_refresh()

        # 阻止 WebEngine 内部子控件在输入框活跃时抢焦点
        if event.type() == QEvent.FocusIn:
            web_view = getattr(self, 'web_view', None)
            if web_view and obj is not web_view and isinstance(obj, QWidget):
                if web_view.isAncestorOf(obj):
                    active = QApplication.focusWidget()
                    float_x_edit = getattr(getattr(self, 'float_x', None), 'edit', None)
                    float_y = getattr(self, 'float_y', None)
                    if active in [float_x_edit, float_y]:
                        return True  # 吞掉焦点转移，保留输入框焦点

        return super().eventFilter(obj, event)

    def _install_webview_child_filters(self):
        """延迟给 WebEngine 内部子控件安装 eventFilter（Chromium 子窗口需要等待初始化）。"""
        if not hasattr(self, 'web_view') or self.web_view is None:
            return
        for child in self.web_view.findChildren(QWidget):
            child.installEventFilter(self)

    def _install_chart_context_menu_hook(self):
        """强行修改掉内置 Chromium 的网页右键菜单，映射到散点工具条功能上。"""
        if not hasattr(self, "web_view") or self.web_view is None:
            return
        self.web_view.setContextMenuPolicy(Qt.CustomContextMenu)
        previous_view = getattr(self, "_chart_context_menu_bound_view", None)
        if previous_view is self.web_view:
            return
        if previous_view is not None:
            try:
                previous_view.customContextMenuRequested.disconnect(self._on_chart_context_menu_requested)
            except Exception:
                pass
        self.web_view.customContextMenuRequested.connect(self._on_chart_context_menu_requested)
        self._chart_context_menu_bound_view = self.web_view

    def _on_chart_context_menu_requested(self, pos):
        source = self.sender()
        global_pos = source.mapToGlobal(pos) if source is not None else self.mapToGlobal(pos)
        menu = QMenu(self)
        menu.setStyleSheet(
            "QMenu { background: white; border: 1px solid #d9d9d9; } "
            "QMenu::item:selected { background: #e6f7ff; color: #0050b3; } "
            "QMenu::separator { height: 1px; background: #d9d9d9; margin: 4px 8px; }"
        )

        button_specs = [
            ("分析对象", getattr(self, "btn_analysis_obj", None)),
            ("设置", getattr(self, "btn_settings", None)),
            ("图形样式", getattr(self, "btn_style", None)),
            ("文本设置", getattr(self, "btn_label_settings", None)),
            ("叠加对比", getattr(self, "btn_overlay", None)),
            ("导出", getattr(self, "btn_export", None)),
        ]

        visible_items = []
        for label, button in button_specs:
            if button is None:
                continue
            if not button.isVisible():
                continue
            visible_items.append((label, button))

        if not visible_items:
            return

        for idx, (label, button) in enumerate(visible_items):
            if label == "导出" and idx > 0:
                menu.addSeparator()
            action = menu.addAction(label)
            action.triggered.connect(button.click)
        menu.exec(global_pos)

    def _force_plotly_resize(self):
        """【防御性渲染】强行使用 JS 让 Plotly 回正自己的响应式外框。"""
        if hasattr(self, "web_view") and self.web_view is not None:
            self.web_view.page().runJavaScript(
                "try { "
                "var gd = document.getElementById('chart_div'); "
                "if (gd && window.Plotly) { "
                "  var _fixOpenRightEdge = function(){ "
                "    try { Plotly.relayout(gd, {'yaxis.mirror': false, 'yaxis.side': 'left'}); } catch(_e) {} "
                "  }; "
                "  var p = Plotly.Plots.resize(gd); "
                "  if (p && typeof p.then === 'function') { p.then(_fixOpenRightEdge); } "
                "  else { _fixOpenRightEdge(); } "
                "} "
                "} catch(e) {}"
            )

    def _request_plotly_resize_fast(self):
        """发送轻量无回调的 resize（省去繁重的 Relayout 重计算）。"""
        if hasattr(self, "web_view") and self.web_view is not None:
            self.web_view.page().runJavaScript(
                "try { "
                "var gd = document.getElementById('chart_div'); "
                "if (gd && window.Plotly) { Plotly.Plots.resize(gd); } "
                "} catch(e) {}"
            )

    def _enforce_scatter_open_right_edge(self):
        """散点图需要保持向右无限延展的开口视觉，此处的 JS 强行去掉图表右边框。"""
        if hasattr(self, "web_view") and self.web_view is not None:
            self.web_view.page().runJavaScript(
                "try { "
                "var gd = document.getElementById('chart_div'); "
                "if (gd && window.Plotly) { "
                "  Plotly.relayout(gd, {'yaxis.mirror': false, 'yaxis.side': 'left'}); "
                "} "
                "} catch(e) {}"
            )

    def _layout_refresh_frame(self):
        """协调器 Frame：每执行一次物理 UI 推送。"""
        try:
            self._request_plotly_resize_fast()
        except Exception:
            pass
        try:
            if hasattr(self, 'overlay_controller'):
                self.overlay_controller.schedule_rect_sync()
        except Exception:
            pass
        try:
            self._update_crosshair_ui(recompute_quadrants=False)
        except Exception:
            pass

    def _layout_refresh_final(self):
        """协调器 Final：防抖结束后收尾巩固。"""
        try:
            self._force_plotly_resize()
        except Exception:
            pass
        try:
            if hasattr(self, 'overlay_controller'):
                self.overlay_controller.schedule_rect_sync()
        except Exception:
            pass
        try:
            self._update_crosshair_ui(recompute_quadrants=False)
        except Exception:
            pass

    def _flush_layout_refresh_request(self):
        coordinator = getattr(self, "_layout_refresh", None)
        if coordinator is not None:
            coordinator.notify()
            return
        self._layout_refresh_frame()

    def _request_layout_refresh(self):
        timer = getattr(self, "_layout_refresh_coalesce_timer", None)
        if timer is not None:
            timer.start()
            return
        self._flush_layout_refresh_request()

    def _on_splitter_moved(self, pos, index):
        self._request_layout_refresh()

    # ---------------------------------------------------
    # Part 3.4: WebEngine 生命锁监控与骨架屏
    # ---------------------------------------------------
    def _show_skeleton(self, text: str = "正在绘图..."):
        """使用包含动画 Label 的 Qt 原生层，遮住下方可能处于白屏或紊乱阶段的网页。"""
        if hasattr(self, "chart_skeleton") and self.chart_skeleton:
            try:
                self._loading_mask_visible = True
                self.chart_skeleton.lbl.setText(text)
                self.chart_skeleton.setGeometry(self.grid_widget.rect())
                self.chart_skeleton.raise_()
                self.chart_skeleton.show()
                if hasattr(self, "float_x") and self.float_x:
                    self.float_x.hide()
                if hasattr(self, "float_y") and self.float_y:
                    self.float_y.hide()
            except Exception:
                pass

    def _on_webview_load_finished(self, ok: bool):
        self._webview_loaded = bool(ok)
        if not ok:
            self._show_skeleton("绘图区加载失败（请重试）")
            return
        self._maybe_hide_skeleton()

    def _maybe_hide_skeleton(self):
        """
        触发收起骨架的逻辑判定门。
        当且仅当 [网页文件彻底装载完毕] 且 [JSON指令完成下发] 时，才会给一个 10秒 倒计时，并放行准备撤去骨架。
        """
        if not (self._webview_loaded and self._data_sent):
            return
        self._waiting_token = self._render_token
        self._ready_deadline = time.monotonic() + 10.0
        if not self._ready_poll_timer.isActive():
            self._ready_poll_timer.start()

    def _poll_plotly_ready(self):
        """
        [最核心脏跳动] 每 50ms 一次。
        检查 Plotly 所挂靠的 DOM div 是不是真正的长出来了宽高属性，一旦长出，立刻扯掉遮羞布（Skeleton）。
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
          }catch(e){ return null; }
        })();
        """

        def _cb(val):
            self._js_inflight = False
            if self._waiting_token != self._render_token:
                return
            if bool(val):
                self._ready_poll_timer.stop()
                try:
                    if hasattr(self, "chart_skeleton") and self.chart_skeleton:
                        self.chart_skeleton.hide()
                    self._loading_mask_visible = False
                except Exception:
                    pass
                try:
                    self._update_crosshair_ui()
                except Exception:
                    pass
                return

            try:
                if hasattr(self, "overlay_controller"):
                    self.overlay_controller.schedule_rect_sync()
            except Exception:
                pass

        try:
            self.web_view.page().runJavaScript(js, _cb)
        except Exception:
            self._js_inflight = False

    # ---------------------------------------------------
    # Part 3.5: 控制手柄、悬浮框反向互通的 UI 驱动
    # ---------------------------------------------------
    def _on_slider_x_moved(self, val):
        self.current_x_val = val
        self._update_crosshair_ui()

    def _on_slider_y_moved(self, val):
        self.current_y_val = val
        self._update_crosshair_ui()
        
    def _on_edit_x_confirmed(self):
        """将用户键盘手敲出的数值强制钳制在滑块极限范围内，并更新滑块以跟手。"""
        mag = self.float_x.mag.current_mag
        txt = self.float_x.edit.text().replace(self.float_x.edit.suffix, '')
        try:
            val = float(txt) * (10 ** mag)
            self.current_x_val = max(self.x_range_min, min(val, self.x_range_max))
            self.slider_x.setValue(self.current_x_val) 
            self._update_crosshair_ui()  
            self.update_stats_ui()
            self.float_x.edit.clearFocus()
        except ValueError:
            pass

    def _on_edit_y_confirmed_rotated(self, new_val):
        self.current_y_val = max(self.y_range_min, min(new_val, self.y_range_max))
        self.slider_y.setValue(self.current_y_val)
        self._update_crosshair_ui()  
        self.update_stats_ui()

    def _update_crosshair_ui(self, recompute_quadrants: bool = True):
        """
        极具控制力的中心枢纽函数。
        根据当前的 (X,Y) 数据点，实时推断其需要落在屏幕物理宽高的第几像素位置。
        然后驱动水平/垂直悬浮小气泡跟随在十字线端点旁边移动，还要兼顾防止气泡框越界。
        """
        if getattr(self, "_loading_mask_visible", False):
            if hasattr(self, 'crosshair'):
                self.crosshair.hide()
            if hasattr(self, "float_x") and self.float_x:
                self.float_x.hide()
            if hasattr(self, "float_y") and self.float_y:
                self.float_y.hide()
            return
        
        if hasattr(self, 'crosshair'):
            self.crosshair.set_crosshair_values(self.current_x_val, self.current_y_val)
            self.crosshair.show()

        slider_x_pos = self.slider_x.mapTo(self.grid_widget, QPoint(0, 0))
        slider_y_pos = self.slider_y.mapTo(self.grid_widget, QPoint(0, 0))

        px_x = self.slider_x.val_to_pixel(self.current_x_val)
        px_y = self.slider_y.val_to_pixel(self.current_y_val)

        # -- X 悬浮气泡控制流 --
        if self.float_x:
            self.float_x.set_magnitude(self.x_mag if self.x_mag is not None else 0)
            
            x_span = self.x_range_max - self.x_range_min
            x_dtick = DriskMath.calc_smart_step(x_span)
            x_mag_div = 10 ** self.float_x.mag.current_mag if self.float_x.mag.current_mag else 1
            x_norm_step = (x_dtick / x_mag_div) if x_mag_div else 0.0
            x_decimals = max(0, math.ceil(2 - math.log10(x_norm_step))) if x_norm_step > 0 else 2

            if not self.float_x.edit.hasFocus():
                self.float_x.edit.setText(f"{self.current_x_val / x_mag_div:.{x_decimals}f}{self.float_x.edit.suffix}")
            
            self.float_x.setFixedWidth(50) 
            self.float_x.setFixedHeight(20)
            
            gap_ref = 15
            fx_x = slider_x_pos.x() + px_x - self.float_x.width() / 2
            top_gap = 4
            fx_y = slider_x_pos.y() - self.float_x.height() - top_gap
            min_visible_top = 2
            fx_y = max(min_visible_top, fx_y)
            
            # X气泡防撞界限设计
            safe_right_bound = min(slider_x_pos.x() + self.slider_x.margin_left + self.slider_x.get_valid_width(), 
                                   slider_x_pos.x() + self.slider_x.width() - 5)
            safe_left_bound = max(slider_x_pos.x() + 5, 
                                  slider_x_pos.x() + self.slider_x.margin_left)
                                  
            fx_x = max(safe_left_bound - self.float_x.width() / 2, 
                       min(fx_x, safe_right_bound - self.float_x.width() / 2))
            
            if self.float_x.pos().x() != int(fx_x) or self.float_x.pos().y() != int(fx_y):
                self.float_x.move(int(fx_x), int(fx_y))
            if not self.float_x.isVisible(): self.float_x.show()
            self.float_x.raise_()

        # -- Y 悬浮气泡控制流 --
        if self.float_y:
            mag_div = 10 ** self.y_mag if self.y_mag is not None else 1
            suffix = SmartFormatter.SI_MAP.get(self.y_mag if self.y_mag is not None else 0, "")
            
            y_span = self.y_range_max - self.y_range_min
            y_dtick = DriskMath.calc_smart_step(y_span)
            y_norm_step = (y_dtick / mag_div) if mag_div else 0.0
            y_decimals = max(0, math.ceil(2 - math.log10(y_norm_step))) if y_norm_step > 0 else 2
            
            self.float_y.set_value(self.current_y_val, mag_div, suffix, decimals=y_decimals)
            
            fy_x = slider_y_pos.x() + gap_ref
            fy_y = slider_y_pos.y() + px_y - self.float_y.height() / 2
            
            # Y气泡防撞界限设计
            safe_bottom_bound = min(slider_y_pos.y() + self.slider_y.margin_top + self.slider_y.get_valid_height(), 
                                    slider_y_pos.y() + self.slider_y.height() - 5)
            safe_top_bound = max(slider_y_pos.y() + 5, 
                                 slider_y_pos.y() + self.slider_y.margin_top)
                                 
            fy_y = max(safe_top_bound - self.float_y.height() / 2, 
                       min(fy_y, safe_bottom_bound - self.float_y.height() / 2))
                       
            if self.float_y.pos().x() != int(fy_x) or self.float_y.pos().y() != int(fy_y):
                self.float_y.move(int(fy_x), int(fy_y))
            if not self.float_y.isVisible(): self.float_y.show()
            self.float_y.raise_()

        if recompute_quadrants:
            self._calc_quadrant_stats()

    # ---------------------------------------------------
    # Part 3.6: 图表引擎与数据的装配车间 (Payload Assembly)
    # ---------------------------------------------------
    def _get_name_resolve_excel_app(self):
        if self._name_resolve_excel_app_ready:
            return self._name_resolve_excel_app
        self._name_resolve_excel_app_ready = True
        try:
            self._name_resolve_excel_app = xl_app()
        except Exception:
            self._name_resolve_excel_app = None
        return self._name_resolve_excel_app

    def _get_sim_data_dynamic(self, full_addr, sim_result):
        """深入缓存层，依据单元格 Address (比如 A1) 甚至其自带的 RangeName 取出 1D 数据阵列。"""
        if not sim_result: return None, None
        clean_addr = full_addr.replace('$', '').upper()
        cell_part = clean_addr.split('!')[-1] if '!' in clean_addr else clean_addr
        xl_ctx = self._get_name_resolve_excel_app()

        for k, v in (sim_result.output_cache or {}).items():
            k_clean = k.replace('$', '').upper()
            if k_clean == clean_addr or k_clean.startswith(f"{clean_addr}_"):
                attrs = sim_result.output_attributes.get(k, {}) or {}
                name = resolve_visible_variable_name(
                    k,
                    attrs,
                    excel_app=xl_ctx,
                    fallback_label=cell_part,
                )
                return name, np.ravel(v)
                
        for k, v in sim_result.input_cache.items():
            if not is_input_key_exposed(sim_result, k):
                continue
            k_clean = k.replace('$', '').upper()
            k_base = k_clean.split('_')[0] if '_' in k_clean else k_clean
            if k_base == clean_addr:
                attrs = sim_result.input_attributes.get(k) or sim_result.input_attributes.get(k_base) or {}
                name = resolve_visible_variable_name(
                    k,
                    attrs,
                    excel_app=xl_ctx,
                    fallback_label=cell_part,
                )
                return name, np.ravel(v)
                 
        return cell_part, None

    def _resolve_scatter_metadata_chart_title(self) -> str:
        sim = get_simulation(self.sim_id)
        title_text, _ = self._get_sim_data_dynamic(getattr(self, "y_addr", ""), sim)
        resolved = str(title_text or "").strip()
        if resolved:
            return resolved
        fallback = str(getattr(self, "y_addr", "") or "").replace("$", "").strip()
        return fallback.split("!")[-1] if "!" in fallback else (fallback or "Y")

    def _resolve_scatter_display_chart_title(self) -> str:
        metadata_title = self._resolve_scatter_metadata_chart_title()
        cfg = getattr(self, "_label_settings_config", None)
        chart_cfg = cfg.get("chart_title", {}) if isinstance(cfg, dict) else {}
        if isinstance(chart_cfg, dict):
            raw_text = chart_cfg.get("text_override", None)
            if raw_text is not None:
                manual_title = str(raw_text).strip()
                if manual_title:
                    return manual_title
        return metadata_title

    def _build_scatter_render_label_overrides(self, display_title: str) -> Dict[str, Any]:
        """将外界注入给 Y 轴的主名称封装入覆盖用的设置包中，发配给绘图中心。"""
        src_cfg = getattr(self, "_label_settings_config", None)
        if isinstance(src_cfg, dict):
            render_cfg = copy.deepcopy(src_cfg)
        else:
            render_cfg = create_default_label_settings_config()

        chart_cfg = render_cfg.setdefault("chart_title", {})
        if not isinstance(chart_cfg, dict):
            chart_cfg = {}
            render_cfg["chart_title"] = chart_cfg
        chart_cfg["text_override"] = str(display_title or "").strip() or None
        return render_cfg

    def _align_clean_data(self, arr_x, arr_y):
        """
        【防呆拦截】由于不同版本间数据或因空算导致长度参差不齐，或者自带非数值型垃圾，
        必须利用 Pandas 宽容强转并进行一一对应的成对截断清洗（DropNA），防止前端图形崩溃。
        """
        min_len = min(len(arr_x), len(arr_y))
        x_trim = arr_x[:min_len]
        y_trim = arr_y[:min_len]
        
        df_temp = pd.DataFrame({"x": x_trim, "y": y_trim})
        df_temp['x'] = pd.to_numeric(df_temp['x'], errors='coerce')
        df_temp['y'] = pd.to_numeric(df_temp['y'], errors='coerce')
        df_temp = df_temp.dropna()
        
        return df_temp['x'].values, df_temp['y'].values

    def _apply_scatter_clipping(self, arr_x: np.ndarray, arr_y: np.ndarray):
        """如果应用配置了边界裁切规则，将越界的噪点直接杀掉。"""
        cfg = ScatterClipSettingsDialog.normalize_config(getattr(self, "_scatter_clip_config", None))
        if len(arr_x) == 0 or len(arr_y) == 0:
            return arr_x, arr_y

        mask = np.ones(len(arr_x), dtype=bool)

        x_cfg = cfg.get("x", {})
        if x_cfg.get("enabled", False):
            x_low = x_cfg.get("lower", None)
            x_high = x_cfg.get("upper", None)
            if x_low is not None:
                mask &= (arr_x >= x_low)
            if x_high is not None:
                mask &= (arr_x <= x_high)

        y_cfg = cfg.get("y", {})
        if y_cfg.get("enabled", False):
            y_low = y_cfg.get("lower", None)
            y_high = y_cfg.get("upper", None)
            if y_low is not None:
                mask &= (arr_y >= y_low)
            if y_high is not None:
                mask &= (arr_y <= y_high)

        if not np.any(mask):
            return np.array([], dtype=float), np.array([], dtype=float)
        return arr_x[mask], arr_y[mask]

    def _is_scatter_clipping_enabled(self) -> bool:
        cfg = ScatterClipSettingsDialog.normalize_config(getattr(self, "_scatter_clip_config", None))
        return bool(cfg.get("x", {}).get("enabled")) or bool(cfg.get("y", {}).get("enabled"))

    def _calc_quadrant_stats(self):
        """向量化极速切分所有组数据的四象限样本量百分占比。"""
        if not getattr(self, "_raw_data_cache", None): return
        
        stats = []
        display_groups = self._raw_data_cache[:4]
        
        for group in display_groups:
            x_data = group['x']
            y_data = group['y']
            total = len(x_data)
            
            if total == 0: 
                stats.append((0.0, 0.0, 0.0, 0.0))
                continue

            q1 = np.sum((x_data >= self.current_x_val) & (y_data >= self.current_y_val)) / total
            q2 = np.sum((x_data <  self.current_x_val) & (y_data >= self.current_y_val)) / total
            q3 = np.sum((x_data <  self.current_x_val) & (y_data <  self.current_y_val)) / total
            q4 = np.sum((x_data >= self.current_x_val) & (y_data <  self.current_y_val)) / total
            
            stats.append((q1, q2, q3, q4))
            
        if hasattr(self, 'crosshair'):
            self.crosshair.set_quadrant_stats(stats)

    def _update_plot(self):
        """
        图表装载主流程。极其庞大但条理清晰：
        1. 搜集并清洗主版本和所有勾选了叠加情景的数据。
        2. 若因裁剪全部数据被剔除则给出空图警告。
        3. 进行 Pearson/Spearman 统计指标推算。
        4. 构建图表专属 JSON Payload，强制配置坐标系样式并清理冗余。
        5. 将组装好的字典通过 pywebchannel 抛入 WebEngine。
        6. 完成后同步十字准星、激活表格属性更新。
        """
        try:
            self._refresh_window_title()
            self._show_skeleton("正在绘图...")
            QApplication.processEvents(QEventLoop.ExcludeUserInputEvents)
            self._render_token += 1
            self._data_sent = False

            all_sims = get_all_simulations()
            raw_groups = []
            primary_x_names = []
            
            # --- 内部闭包装配厂 ---
            def _add_to_raw_groups(sid, y_address, y_display_name, is_overlay=False):
                sim = all_sims.get(int(sid))
                if not sim: return
                
                y_name, y_data = self._get_sim_data_dynamic(y_address, sim)
                if y_data is None: return
                if (not is_overlay) and str(sid) == str(self.sim_id):
                    if str(y_name or "").strip():
                        self.y_name = str(y_name).strip()
                 
                for x_addr in self.x_addrs:
                    x_name, x_data = self._get_sim_data_dynamic(x_addr, sim)
                    if x_data is None: continue
                    if (not is_overlay) and str(sid) == str(self.sim_id):
                        primary_x_names.append(str(x_name) if x_name else str(x_addr))
                    
                    clean_x, clean_y = self._align_clean_data(x_data, y_data)
                    clean_x, clean_y = self._apply_scatter_clipping(clean_x, clean_y)
                    if len(clean_x) < 2: continue

                    x_label = str(x_name) if str(x_name or "").strip() else str(x_addr)
                    trace_name = x_label
                    trace_display_name = x_label
                    if is_overlay or len(self.overlay_items) > 0:
                        sim_label = get_sim_display_name(sid)
                        trace_name = f"{x_label} (sim{sid})"
                        trace_display_name = f"{x_label} ({sim_label})"

                    raw_groups.append({
                        "name": trace_name,
                        "display_name": trace_display_name,
                        "x": clean_x,
                        "y": clean_y
                    })

            # 装载本体
            _add_to_raw_groups(self.sim_id, self.y_addr, self.y_name, is_overlay=False)
            # 装载叠加件
            for sid in self.overlay_items:
                _add_to_raw_groups(int(sid), self.y_addr, self.y_name, is_overlay=True)

            # 极端空图处理
            if not raw_groups:
                if self._is_scatter_clipping_enabled():
                    self._show_skeleton("裁剪后无有效数据点，请调整设置。")
                empty_scatter_payload = {
                    "data": [],
                    "layout": {
                        "title": {"text": ""},
                        "annotations": [],
                        "showlegend": False,
                    },
                }
                self.host.load_plot(
                    plot_json=json.dumps(empty_scatter_payload),
                    js_mode="scatter",
                    static_plot=False,
                )
                self._data_sent = True
                self._maybe_hide_skeleton()
                return

            self._scatter_display_keys = [str(g.get("name", "")) for g in raw_groups if g.get("name")]
            self._sync_scatter_style_map(self._scatter_display_keys)

            cat_x = np.concatenate([g["x"] for g in raw_groups])
            cat_y = np.concatenate([g["y"] for g in raw_groups])
            
            self._raw_data_cache = raw_groups[:10]
            
            # --- 构建统计指标库 ---
            self._static_stats_cache = {}
            for group in self._raw_data_cache:
                df = pd.DataFrame({"x": group["x"], "y": group["y"]})
                self._static_stats_cache[group["name"]] = {
                    "x_mean": df["x"].mean(),
                    "x_std": df["x"].std(ddof=1) if len(df) > 1 else 0.0,
                    "y_mean": df["y"].mean(),
                    "y_std": df["y"].std(ddof=1) if len(df) > 1 else 0.0,
                    "pearson": df["x"].corr(df["y"], method="pearson"),
                    "spearman": df["x"].corr(df["y"], method="spearman"),
                }
            
            x_override_mag = get_axis_display_unit_override(getattr(self, "_label_settings_config", None), "x")
            y_override_mag = get_axis_display_unit_override(getattr(self, "_label_settings_config", None), "y")
            if x_override_mag is not None:
                self.x_mag = int(x_override_mag)
            elif self.x_mag is None:
                self.x_mag = infer_si_mag(value_range=(float(np.min(cat_x)), float(np.max(cat_x))))

            if y_override_mag is not None:
                self.y_mag = int(y_override_mag)
            elif self.y_mag is None:
                self.y_mag = infer_si_mag(value_range=(float(np.min(cat_y)), float(np.max(cat_y))))

            runtime_axis_flags = self._resolve_scatter_axis_numeric_flags()
            self._label_axis_numeric = get_axis_numeric_flags(
                getattr(self, "_label_settings_config", None),
                fallback=runtime_axis_flags,
            )

            # --- 图表工厂开始烧制 ---
            display_chart_title = self._resolve_scatter_display_chart_title()
            render_label_overrides = self._build_scatter_render_label_overrides(display_chart_title)
            fig_dict = DriskChartFactory.build_scatter_figure(
                data_groups=raw_groups, 
                y_title=self.y_name,
                x_forced_mag=self.x_mag,
                y_forced_mag=self.y_mag,
                style_map=self._scatter_style_map,
                label_overrides=render_label_overrides,
                axis_numeric_flags=getattr(self, "_label_axis_numeric", None),
            )
            
            try:
                _plot_dict = json.loads(fig_dict["plot_json"])
                _layout_dict = _plot_dict.setdefault("layout", {})
                _layout_dict["title"] = {"text": ""}
                _layout_dict["annotations"] = []
                fig_dict["plot_json"] = json.dumps(_plot_dict)
            except Exception:
                pass

            x_axis_title = "相关变量"
            if len(self.x_addrs) == 1 and primary_x_names:
                x_axis_title = primary_x_names[0]
            x_axis_title_override = None
            try:
                x_axis_title_override = (
                    ((self._label_settings_config or {}).get("axis_title", {}) or {}).get("x", {}) or {}
                ).get("text_override", None)
            except Exception:
                x_axis_title_override = None
                
            if x_axis_title_override is None:
                try:
                    plot_dict = json.loads(fig_dict["plot_json"])
                    layout_dict = plot_dict.setdefault("layout", {})
                    xaxis_dict = layout_dict.setdefault("xaxis", {})
                    title_obj = xaxis_dict.get("title")
                    existing_title_text = ""
                    if isinstance(title_obj, dict):
                        existing_title_text = str(title_obj.get("text", "") or "")
                    else:
                        existing_title_text = str(title_obj or "")

                    suffix_hint = ""
                    existing_title_text = existing_title_text.strip()
                    if existing_title_text.endswith(")") and "(" in existing_title_text:
                        candidate = existing_title_text[existing_title_text.rfind("(") + 1:-1].strip()
                        if candidate and "=" in candidate:
                            suffix_hint = candidate

                    final_x_title = x_axis_title
                    if suffix_hint:
                        final_x_title = f"{x_axis_title} ({suffix_hint})"

                    if isinstance(title_obj, dict):
                        title_obj["text"] = final_x_title
                    else:
                        xaxis_dict["title"] = {"text": final_x_title}
                    fig_dict["plot_json"] = json.dumps(plot_dict)
                except Exception:
                    pass
            
            # 暴力反抽内部边距给 Qt
            try:
                _layout = json.loads(fig_dict["plot_json"]).get("layout", {})
                _m = _layout.get("margin", {})
                
                _ml, _mr = float(_m.get("l", 80)), float(_m.get("r", 0))
                _mt, _mb = float(_m.get("t", 0)), float(_m.get("b", 60))
                
                self.slider_x.setMargins(_ml, _mr + self.slider_y.width())
                self.slider_y.setMargins(_mt + self.slider_x.height(), _mb)
            except Exception as e:
                print(f"Margin Injection Failed: {e}")
            
            
            if "x_range" in fig_dict and "y_range" in fig_dict:
                self.x_range_min, self.x_range_max = fig_dict["x_range"]
                self.y_range_min, self.y_range_max = fig_dict["y_range"]
            else:
                self.x_range_min, self.x_range_max = float(np.min(cat_x)), float(np.max(cat_x))
                self.y_range_min, self.y_range_max = float(np.min(cat_y)), float(np.max(cat_y))
            
            if hasattr(self, 'crosshair'):
                self.crosshair.set_axis_range(self.x_range_min, self.x_range_max, self.y_range_min, self.y_range_max)

            # 起飞落点为首组坐标云团的物理中心
            self.current_x_val = (self.x_range_min + self.x_range_max) / 2.0
            self.current_y_val = (self.y_range_min + self.y_range_max) / 2.0
            first_group_name = raw_groups[0].get("name") if raw_groups else None
            if first_group_name:
                first_stats = self._static_stats_cache.get(first_group_name, {})
                try:
                    first_x_mean = float(first_stats.get("x_mean"))
                    if np.isfinite(first_x_mean):
                        self.current_x_val = first_x_mean
                except Exception:
                    pass
                try:
                    first_y_mean = float(first_stats.get("y_mean"))
                    if np.isfinite(first_y_mean):
                        self.current_y_val = first_y_mean
                except Exception:
                    pass
            self.current_x_val = max(self.x_range_min, min(self.current_x_val, self.x_range_max))
            self.current_y_val = max(self.y_range_min, min(self.current_y_val, self.y_range_max))
            
            self.slider_x.setRangeLimit(self.x_range_min, self.x_range_max)
            self.slider_y.setRangeLimit(self.y_range_min, self.y_range_max)
            self.slider_x.setValue(self.current_x_val)
            self.slider_y.setValue(self.current_y_val)
            
            # --- 指令下放完毕，发送 ---
            self.host.load_plot(plot_json=fig_dict["plot_json"], js_mode=fig_dict["js_mode"], static_plot=True)
            QTimer.singleShot(60, self._enforce_scatter_open_right_edge)
            self._data_sent = True
            self._maybe_hide_skeleton()
            
            if hasattr(self, 'overlay_controller'):
                self.overlay_controller.schedule_rect_sync()
                QTimer.singleShot(120, self.overlay_controller.schedule_rect_sync)

            # 只在没有输入控件持有焦点时才转移焦点到 web_view
            active = QApplication.focusWidget()
            float_x_edit = getattr(getattr(self, 'float_x', None), 'edit', None)
            float_y = getattr(self, 'float_y', None)
            if active not in [float_x_edit, float_y]:
                self.float_x.edit.clearFocus()
                self.web_view.setFocus()

            self._update_crosshair_ui()
            self.update_stats_ui()
            
        except Exception as e:
            empty_scatter_payload = {
                "data": [],
                "layout": {
                    "title": {"text": ""},
                    "annotations": [],
                    "showlegend": False,
                },
            }
            self.host.load_plot(
                plot_json=json.dumps(empty_scatter_payload),
                js_mode="scatter",
                static_plot=True,
            )
            self._data_sent = True
            self._maybe_hide_skeleton()
            print(f"Plot Error: {traceback.format_exc()}")

    def closeEvent(self, event):
        """窗口被销毁时回收共享轮子，掐断所有 Timer 以免 C++ 悬挂引发的静默闪退。"""
        for attr in ("_ready_poll_timer", "_layout_refresh_coalesce_timer", "_geom_poll_timer", "_geom_poll_stop_timer"):
            try:
                t = getattr(self, attr, None)
                if t is not None and t.isActive():
                    t.stop()
            except Exception:
                print("closeEvent: failed to stop %s", attr, exc_info=True)
        try:
            if hasattr(self, "web_view") and self.web_view is not None:
                recycle_shared_webview(self.web_view)
                self.web_view = None
        except Exception:
            print("closeEvent: failed to recycle web_view", exc_info=True)
        super().closeEvent(event)


    # ---------------------------------------------------
    # Part 3.7: 周边功能支持中心与事件总线
    # ---------------------------------------------------
    def update_stats_ui(self):
        """将底层计算出的特征数值铺陈到 UI 层的 QTableWidget 表格。"""
        if not getattr(self, "_raw_data_cache", None) or not getattr(self, "_static_stats_cache", None):
            self.stats_table.setRowCount(0)
            return

        x_means, x_stds, y_means, y_stds = [], [], [], []
        pearsons, spearmans = [], []
        delims_x, delims_y = [], []
        q1s, q2s, q3s, q4s = [], [], [], []
        
        headers = []
        header_colors = []

        for i, group in enumerate(self._raw_data_cache):
            name = group["name"]
            display_name = str(group.get("display_name", name) or name)
            stats = self._static_stats_cache.get(name, {})
            
            x_means.append(stats.get("x_mean", ""))
            x_stds.append(stats.get("x_std", ""))
            y_means.append(stats.get("y_mean", ""))
            y_stds.append(stats.get("y_std", ""))
            pearsons.append(stats.get("pearson", ""))
            spearmans.append(stats.get("spearman", ""))

            delims_x.append(self.current_x_val)
            delims_y.append(self.current_y_val)

            x_data = group['x']
            y_data = group['y']
            total = len(x_data)
            if total > 0:
                q1 = np.sum((x_data >= self.current_x_val) & (y_data >= self.current_y_val)) / total
                q2 = np.sum((x_data <  self.current_x_val) & (y_data >= self.current_y_val)) / total
                q3 = np.sum((x_data <  self.current_x_val) & (y_data <  self.current_y_val)) / total
                q4 = np.sum((x_data >= self.current_x_val) & (y_data <  self.current_y_val)) / total
            else:
                q1 = q2 = q3 = q4 = 0.0

            q1s.append(f"{q1*100:.1f}%")
            q2s.append(f"{q2*100:.1f}%")
            q3s.append(f"{q3*100:.1f}%")
            q4s.append(f"{q4*100:.1f}%")

            if i >= len(DRISK_COLOR_CYCLE):
                shape_idx = (i // len(DRISK_COLOR_CYCLE)) + 1
                display_name = f"{display_name} (形状 {shape_idx})"

            headers.append(display_name)
            header_colors.append(DRISK_COLOR_CYCLE[i % len(DRISK_COLOR_CYCLE)])

        rows = [
            ("X 均值", x_means, False),
            ("X 标准差", x_stds, False),
            ("Y 均值", y_means, False),
            ("Y 标准差", y_stds, False),
            SEP_ROW_MARKER,
            ("相关系数 Pearson", pearsons, False),
            ("相关系数 Spearman", spearmans, False),
            SEP_ROW_MARKER,
            ("X 分界线", delims_x, False),
            ("Y 分界线", delims_y, False),
            ("第一象限（右上）", q1s, False),
            ("第二象限（左上）", q2s, False),
            ("第三象限（左下）", q3s, False),
            ("第四象限（右下）", q4s, False),
        ]

        render_stats_table(
            self.stats_table, 
            rows, 
            column_headers=headers, 
            header_colors=header_colors
        )

    def _get_available_sims_for(self, y_addr):
        available = []
        all_sims = get_all_simulations()
        clean_addr = y_addr.replace('$', '').upper()
        
        for sid, sim in all_sims.items():
            output_cache = getattr(sim, "output_cache", {}) or {}
            for k in output_cache.keys():
                k_clean = k.replace('$', '').upper()
                if k_clean == clean_addr or k_clean.startswith(f"{clean_addr}_"):
                    available.append(sid)
                    break
            else:
                input_cache = getattr(sim, "input_cache", {}) or {}
                for k in input_cache.keys():
                    if not is_input_key_exposed(sim, k):
                        continue
                    k_clean = k.replace('$', '').upper()
                    k_base = k_clean.split('_')[0] if '_' in k_clean else k_clean
                    if k_base == clean_addr:
                        available.append(sid)
                        break
        return available

    def _build_default_scatter_style(self, index: int) -> Dict[str, Any]:
        color = DRISK_COLOR_CYCLE[index % len(DRISK_COLOR_CYCLE)]
        symbol_idx = (index // len(DRISK_COLOR_CYCLE)) % len(SCATTER_SYMBOLS)
        symbol = SCATTER_SYMBOLS[symbol_idx]
        return {
            "fill_color": color,
            "fill_opacity": 0.6,
            "outline_color": "#ffffff",
            "outline_width": 0.0,
            "marker_symbol": symbol,
            "marker_size": 4.5,
            "opacity": 1.0,
        }

    def _sync_scatter_style_map(self, display_keys: List[str]):
        new_map: Dict[str, Dict[str, Any]] = {}
        for i, key in enumerate(display_keys):
            base = self._build_default_scatter_style(i)
            old = self._scatter_style_map.get(key, {})
            merged = dict(base)
            merged.update(old)
            new_map[key] = merged
        self._scatter_style_map = new_map

    def _resolve_scatter_group_display_name(self, internal_key: str) -> str:
        key = str(internal_key or "")
        for group in list(getattr(self, "_raw_data_cache", []) or []):
            if str(group.get("name", "")) == key:
                display_name = str(group.get("display_name", "") or "").strip()
                if display_name:
                    return display_name
                break
        sid = extract_sim_id_from_series_key(key)
        if sid is None:
            return key
        head = key[: key.rfind("(sim")].strip() if "(sim" in key else key
        sim_label = get_sim_display_name(sid)
        if head:
            return f"{head} ({sim_label})"
        return sim_label

    def _build_scatter_display_name_map(self, internal_keys: List[str]) -> Dict[str, str]:
        mapping: Dict[str, str] = {}
        used: set[str] = set()
        for raw_key in list(internal_keys or []):
            key = str(raw_key or "")
            base = str(self._resolve_scatter_group_display_name(key) or key).strip()
            if not base:
                base = key
            final = base
            sid = extract_sim_id_from_series_key(key)
            if final in used and sid is not None:
                final = f"{base} [ID {sid}]"
            idx = 2
            while final in used:
                final = f"{base} #{idx}"
                idx += 1
            used.add(final)
            mapping[key] = final
        return mapping

    def _detect_data_is_numeric(self, values) -> bool:
        try:
            arr = np.asarray(values).ravel()
        except Exception:
            return False
        if arr.size <= 0:
            return False

        series = pd.Series(arr)
        if series.empty:
            return False
        series = series.mask(series.map(lambda value: isinstance(value, str) and (not value.strip()))).dropna()
        if series.empty:
            return False
        numeric_series = pd.to_numeric(series, errors="coerce")
        return bool(numeric_series.notna().all())

    def _build_label_settings_context(self) -> Dict[str, Any]:
        sim = get_simulation(self.sim_id)

        x_candidates: List[Dict[str, Any]] = []
        for addr in list(getattr(self, "x_addrs", []) or []):
            label, data = self._get_sim_data_dynamic(addr, sim)
            fallback_label = str(addr).split("!")[-1] if str(addr) else "X"
            x_candidates.append(
                {
                    "key": str(addr),
                    "label": str(label or fallback_label),
                    "is_numeric": self._detect_data_is_numeric(data),
                }
            )
        if not x_candidates:
            x_candidates.append({"key": "x_main", "label": "当前X数据", "is_numeric": True})

        y_label, y_data = self._get_sim_data_dynamic(getattr(self, "y_addr", ""), sim)
        y_candidates: List[Dict[str, Any]] = [
            {
                "key": str(getattr(self, "y_addr", "") or "y_main"),
                "label": str(y_label or getattr(self, "y_name", "当前Y数据")),
                "is_numeric": self._detect_data_is_numeric(y_data),
            }
        ]

        x_title_default = x_candidates[0]["label"] if len(x_candidates) == 1 else "相关变量"
        y_title_default = str(getattr(self, "y_name", "") or y_candidates[0]["label"] or "Y轴")
        chart_default_text = self._resolve_scatter_metadata_chart_title()

        x_rec_mag = infer_si_mag(
            value_range=(float(getattr(self, "x_range_min", 0.0)), float(getattr(self, "x_range_max", 1.0))),
            dtick=None,
            forced_mag=None,
        )
        y_rec_mag = infer_si_mag(
            value_range=(float(getattr(self, "y_range_min", 0.0)), float(getattr(self, "y_range_max", 1.0))),
            dtick=None,
            forced_mag=None,
        )

        return {
            "chart_title": {
                "default_text": chart_default_text,
                "default_font_family": "Arial",
                "default_font_size": 14,
            },
            "axes": {
                "x": {
                    "default_title": x_title_default,
                    "default_title_font_family": "Arial",
                    "default_title_font_size": 12,
                    "default_tick_font_family": "Arial",
                    "default_tick_font_size": 12,
                    "recommended_mag": int(x_rec_mag),
                    "data_candidates": x_candidates,
                },
                "y": {
                    "default_title": y_title_default,
                    "default_title_font_family": "Arial",
                    "default_title_font_size": 12,
                    "default_tick_font_family": "Arial",
                    "default_tick_font_size": 12,
                    "recommended_mag": int(y_rec_mag),
                    "data_candidates": y_candidates,
                },
            },
        }

    def _resolve_scatter_axis_numeric_flags(self) -> Dict[str, bool]:
        ctx = self._build_label_settings_context()
        cfg = getattr(self, "_label_settings_config", {}) or {}
        axis_tick_cfg = cfg.get("axis_tick", {}) if isinstance(cfg, dict) else {}

        result = {"x": True, "y": True}
        for axis in ("x", "y"):
            candidates = list(ctx.get("axes", {}).get(axis, {}).get("data_candidates", []))
            selected = ""
            if isinstance(axis_tick_cfg, dict):
                selected = str((axis_tick_cfg.get(axis, {}) or {}).get("selected_data_key", "") or "")
            if selected:
                for item in candidates:
                    if str(item.get("key", "")) == selected:
                        result[axis] = bool(item.get("is_numeric", True))
                        break
                else:
                    result[axis] = bool(candidates[0].get("is_numeric", True)) if candidates else True
            else:
                result[axis] = bool(candidates[0].get("is_numeric", True)) if candidates else True
        return result

    def open_label_settings_dialog(self):
        try:
            context = self._build_label_settings_context()
            dlg = LabelSettingsDialog(
                config=getattr(self, "_label_settings_config", None),
                context=context,
                parent=self,
                include_chart_title=False,
            )
            if dlg.exec():
                new_cfg = dlg.get_config()
                self._label_settings_config = new_cfg
                self._label_axis_numeric = get_axis_numeric_flags(
                    new_cfg,
                    fallback=self._resolve_scatter_axis_numeric_flags(),
                )
                self.x_mag = get_axis_display_unit_override(new_cfg, "x")
                self.y_mag = get_axis_display_unit_override(new_cfg, "y")
                self._update_plot()
        except Exception as e:
            traceback.print_exc()
            QMessageBox.critical(self, "错误", f"打开文本设置失败：{e}")

    def open_scatter_style_manager(self):
        try:
            import inspect
            import importlib
            import ui_style_manager as style_module

            display_keys = list(getattr(self, "_scatter_display_keys", []) or [])
            if not display_keys and getattr(self, "_raw_data_cache", None):
                display_keys = [str(g.get("name", "")) for g in self._raw_data_cache if g.get("name")]
            if not display_keys:
                QMessageBox.information(self, "提示", "当前无可配置的散点数据组，请先完成绘图。")
                return

            self._sync_scatter_style_map(display_keys)
            display_key_map = self._build_scatter_display_name_map(display_keys)
            proxy_keys: List[str] = []
            proxy_styles: Dict[str, Dict[str, Any]] = {}
            for key in display_keys:
                text_key = str(key)
                display_key = display_key_map.get(text_key, text_key)
                proxy_keys.append(display_key)
                proxy_styles[display_key] = dict(self._scatter_style_map.get(text_key, {}))

            dialog_cls = getattr(style_module, "StyleManagerDialog", StyleManagerDialog)
            required_marker_args = {"show_marker", "marker_shapes", "marker_size_range"}

            def _has_marker_args(cls) -> bool:
                try:
                    params = inspect.signature(cls.__init__).parameters
                except Exception:
                    return False
                return required_marker_args.issubset(set(params.keys()))

            if not _has_marker_args(dialog_cls):
                style_module = importlib.reload(style_module)
                dialog_cls = getattr(style_module, "StyleManagerDialog", dialog_cls)

            if not _has_marker_args(dialog_cls):
                QMessageBox.critical(
                    self,
                    "错误",
                    "当前运行时 StyleManagerDialog 版本不支持散点样式参数，请重启宿主后重试。",
                )
                return

            dlg = dialog_cls(
                proxy_keys,
                proxy_styles,
                show_bar=True,
                show_curve=False,
                show_mean=False,
                is_unified=False,
                parent=self,
                show_marker=True,
                marker_shapes=SCATTER_SYMBOLS,
                marker_size_range=(1.0, 20.0),
            )
            if dlg.exec():
                remapped_styles: Dict[str, Dict[str, Any]] = {}
                for key in display_keys:
                    text_key = str(key)
                    display_key = display_key_map.get(text_key, text_key)
                    if display_key in proxy_styles:
                        remapped_styles[text_key] = dict(proxy_styles.get(display_key, {}))
                if remapped_styles:
                    self._scatter_style_map.update(remapped_styles)
                self._update_plot()
        except Exception as e:
            traceback.print_exc()
            QMessageBox.critical(self, "错误", f"打开图形样式管理器失败：{e}")

    def open_scatter_settings(self):
        try:
            dlg = ScatterClipSettingsDialog(current_config=self._scatter_clip_config, parent=self)
            if dlg.exec():
                new_cfg = ScatterClipSettingsDialog.normalize_config(dlg.get_config())
                if new_cfg != self._scatter_clip_config:
                    self._scatter_clip_config = new_cfg
                    self._update_plot()
        except Exception as e:
            traceback.print_exc()
            QMessageBox.critical(self, "错误", f"打开散点图设置失败：{e}")

    def open_analysis_selector(self):
        try:
            name_rev_before = get_sim_display_name_version()
            dlg = ScatterAnalysisObjectDialog(
                current_sim_id=self.sim_id,
                y_addr=self.y_addr,
                y_name=self.y_name,
                available_sims=self.available_sims,
                parent=self
            )
            
            request_other_data = []
            dlg.requestSelectOtherData.connect(lambda: request_other_data.append(True))
            should_refresh_plot = False

            if dlg.exec():
                selected_sim = dlg.get_selected_sim()
                if selected_sim and selected_sim != str(self.sim_id):
                    self.sim_id = int(selected_sim)  
                    self.overlay_items = []
                    self.x_mag = None
                    self.y_mag = None
                    should_refresh_plot = True

            if get_sim_display_name_version() != name_rev_before:
                should_refresh_plot = True

            if request_other_data:
                self._open_config_dialog()
                if should_refresh_plot:
                    self._refresh_window_title()
                    self._update_plot()
            elif should_refresh_plot:
                self._refresh_window_title()
                self._update_plot()
                
        except Exception as e:
            traceback.print_exc()
            QMessageBox.critical(self, "错误", f"打开分析对象选择器失败：{e}")

    def _open_config_dialog(self):
        self.hide()
        
        try:
            config_dlg = ScatterConfigDialog(self.sim_id, parent=self)
            
            if config_dlg.exec():
                old_y_addr = str(getattr(self, "y_addr", "") or "")
                old_x_addrs = [str(addr) for addr in list(getattr(self, "x_addrs", []) or [])]
                self.y_addr = config_dlg.y_addr
                self.x_addrs = list(config_dlg.x_addrs or [])
                new_y_addr = str(getattr(self, "y_addr", "") or "")
                new_x_addrs = [str(addr) for addr in list(getattr(self, "x_addrs", []) or [])]
                source_changed = (old_y_addr != new_y_addr) or (old_x_addrs != new_x_addrs)
                
                sim = get_simulation(self.sim_id)
                self.y_name, _ = self._get_sim_data_dynamic(self.y_addr, sim)
                
                self.available_sims = self._get_available_sims_for(self.y_addr)
                self.btn_overlay.setVisible(len(self.available_sims) > 1 and len(self.x_addrs) == 1)
                
                self.overlay_items = []
                if source_changed:
                    self._scatter_style_map = {}
                    self._scatter_display_keys = []
                self.x_mag = None
                self.y_mag = None
                self._refresh_window_title()
                self._update_plot()
        except Exception as e:
            traceback.print_exc()
            QMessageBox.critical(self, "错误", f"选取数据失败：{e}")
        finally:
            self.show()

    def open_overlay_selector(self):
        try:
            name_rev_before = get_sim_display_name_version()
            dlg = ScatterOverlayDialog(
                current_sim_id=self.sim_id,
                y_addr=self.y_addr,
                y_name=self.y_name,
                available_sims=self.available_sims,
                current_overlays=self.overlay_items,
                parent=self
            )
            should_refresh_plot = False
            if dlg.exec():
                self.overlay_items = dlg.get_selected_sims()
                self.x_mag = None
                self.y_mag = None
                should_refresh_plot = True
            if get_sim_display_name_version() != name_rev_before:
                should_refresh_plot = True
            if should_refresh_plot:
                self._refresh_window_title()
                self._update_plot()
        except Exception as e:
            traceback.print_exc()
            QMessageBox.critical(self, "错误", f"打开叠加选择器失败：{e}")

    # 以下为你刚才指出的被省略的底栏量级控制回调
    def _on_x_mag_changed(self, mag):
        """处理底栏 X 轴下拉框强行更改变变量量级 (如改为 M) 时的重绘。"""
        self.x_mag = mag
        self._update_plot()

    def _on_y_mag_changed(self, mag):
        """处理底栏 Y 轴下拉框强行更改变变量量级 (如改为 M) 时的重绘。"""
        self.y_mag = mag
        self._update_plot()

    def _on_export_clicked(self):
        """挂载全局导出系统，剥离当前 Qt 画板截图输出 PNG / JPEG。"""
        if hasattr(self, 'web_view') and self.web_view:
            self.web_view.setFocus()
            
        try:
            DriskExportManager.export_from_dialog(self)
        except Exception as e:
            QMessageBox.critical(self, "导出失败", f"图表导出发生错误：\n{str(e)}")

# =======================================================
# 4. 全局辅助函数与外部调用接口 (入口枢纽)
# =======================================================
def _iter_scatter_sim_candidates(preferred_sim_id: Optional[int]) -> List[int]:
    """生成查找有效数据的 ID 尝试序列（由于不同情景的计算结果并非全部都在）。"""
    candidates: List[int] = []
    seen = set()

    def _push(raw_sid: Any) -> None:
        try:
            sid = int(raw_sid)
        except Exception:
            return
        if sid in seen:
            return
        seen.add(sid)
        candidates.append(sid)

    _push(preferred_sim_id)

    if bridge is not None and hasattr(bridge, "get_current_sim_id"):
        try:
            _push(bridge.get_current_sim_id())
        except Exception:
            pass

    if bridge is not None and hasattr(bridge, "get_default_result_sim_id"):
        try:
            _push(bridge.get_default_result_sim_id())
        except Exception:
            pass

    sims = get_all_simulations()
    if isinstance(sims, dict):
        for sid in sims.keys():
            _push(sid)

    return candidates

def resolve_scatter_sim_id(preferred_sim_id: Optional[int]) -> Optional[int]:
    """顺藤摸瓜，解析验证究竟哪一个 ID 在库里真实留有可用于散点分析的数据。"""
    for sid in _iter_scatter_sim_candidates(preferred_sim_id):
        sim = get_simulation(sid)
        if sim is None:
            continue
        output_cache = getattr(sim, "output_cache", {}) or {}
        input_cache = getattr(sim, "input_cache", {}) or {}
        has_exposed_input = any(is_input_key_exposed(sim, key) for key in input_cache.keys())
        if output_cache or has_exposed_input:
            return sid
    return None

def show_scatter_dialog_from_macro(sim_id):
    """
    【宏入口】暴露给 Excel Ribbon 工具栏的终极起步方法。
    执行拦截验证 -> 唤起选区配置窗（ConfigDialog） -> 配置通过后，将数据抛入分析主台（PlotlyDialog）。
    """
    global _global_scatter_plot_dialog

    resolved_sim_id = resolve_scatter_sim_id(sim_id)
    if resolved_sim_id is None:
        xlcAlert("未识别到有效的模拟缓存数据，请先运行模拟。")
        return
    
    config_dlg = ScatterConfigDialog(resolved_sim_id)
    if config_dlg.exec():
        y_addr = config_dlg.y_addr
        x_addrs = config_dlg.x_addrs
        
        _global_scatter_plot_dialog = ScatterPlotlyDialog(resolved_sim_id, y_addr, x_addrs)
        _global_scatter_plot_dialog.exec()
        _global_scatter_plot_dialog = None