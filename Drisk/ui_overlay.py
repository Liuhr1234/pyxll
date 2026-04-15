# ui_overlay.py
"""
本模块提供结果视图（Results Views）的高性能交互覆盖层 DriskVLineMaskOverlay。

主要功能：
1. 零延迟视觉反馈：在 QWebEngineView 上方叠加一个完全由 Qt 原生渲染的透明图层。当用户拖拽范围滑块时，
   直接在此图层上重绘定界线与半透明遮罩，避免了高频触发浏览器内核重绘带来的卡顿。
2. 鼠标事件穿透：通过设置特定的 Window 属性，使该遮罩层对鼠标操作完全透明，不阻碍用户与底层图表（如 Hover 提示）的交互。
3. 高精度坐标映射：内部强制使用浮点数 (float) 维护绘图坐标与边距，配合 DPR (设备像素比) 缩放，确保在 4K/高分屏下划线依然精准锐利，无像素抖动。
"""

from PySide6.QtWidgets import QWidget
from PySide6.QtCore import Qt, QRectF, QPointF
from PySide6.QtGui import QColor, QPainter, QPen, QBrush


# =======================================================
# 高性能原生绘制覆盖层
# =======================================================
class DriskVLineMaskOverlay(QWidget):
    """
    覆盖在 WebEngineView 上方的透明绘制层。
    负责绘制两条竖直的定界线（V-Lines）以及界外的半透明遮罩（Masks）。
    """

    def __init__(
            self,
            parent: QWidget,
            margin_l: int = 80,
            margin_r: int = 40,
            margin_t: int = 30,
            margin_b: int = 50,
    ):
        super().__init__(parent)
        
        # [核心机制] 设置 Widget 属性以实现覆盖层效果：
        # WA_TransparentForMouseEvents: 鼠标事件穿透，点按和滑动直接作用于底层的浏览器图表
        # WA_NoSystemBackground & WA_TranslucentBackground: 移除系统默认背景，允许背景完全透明
        self.setAttribute(Qt.WidgetAttribute.WA_TransparentForMouseEvents, True)
        self.setAttribute(Qt.WidgetAttribute.WA_NoSystemBackground, True)
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground, True)

        # 初始化物理边距（图表绘图区到 Widget 边缘的距离）
        self.margin_l = float(margin_l)
        self.margin_r = float(margin_r)
        self.margin_t = float(margin_t)
        self.margin_b = float(margin_b)

        # 图表 X 轴的数据极值范围
        self.xmin = 0.0
        self.xmax = 1.0
        
        # 用户当前选定的左/右区间截断值 (数据坐标)
        self.l = 0.0
        self.r = 0.0
        
        # 视图模式控制：某些特定图表（如 CDF）不需要绘制半透明遮罩
        self.mode = "hist"
        self.enabled = True
        
        # 视觉样式定义
        self.mask_color = QColor(255, 255, 255, 178)        # 遮罩颜色：白色，约 70% 不透明度
        self.line_color = QColor(0, 0, 0, int(0.8 * 255))   # 定界线颜色：黑色，80% 不透明度
        self.line_width = 1
        self.line_style = Qt.PenStyle.SolidLine

        # 缩放倍率 (Device Pixel Ratio)，用于适配高分屏 (Retina/4K)
        self.dpr = 1.0

        # [高精度支持] 覆盖模式下的实际绘图矩形区域。
        # 使用浮点数元组 (Left, Top, Width, Height) 保存，防止 int 截断导致边缘 1px 的对齐抖动。
        self._plot_rect_override: tuple[float, float, float, float] | None = None

    # =======================================================
    # 2. 状态与坐标同步接口
    # =======================================================
    def set_mode(self, mode: str):
        """同步当前的图表模式，并触发重绘以应用不同的遮罩策略"""
        self.mode = mode or "hist"
        self.update()

    def set_axis_range(self, xmin: float, xmax: float):
        """同步底层图表 X 轴的物理极值范围，用于坐标映射计算"""
        try:
            self.xmin = float(xmin)
            self.xmax = float(xmax)
        except Exception:
            pass

    def set_lr(self, l: float, r: float):
        """同步用户当前拖拽的左右截断值 (数据坐标)"""
        try:
            self.l = float(l)
            self.r = float(r)
        except Exception:
            pass

    def set_margins(self, l: float, r: float, t: float, b: float):
        """
        动态更新物理边距。
        [解决痛点]: 在叠加副 Y 轴（例如叠加 CDF 曲线）的模式下，图表的右边距会动态变大。
        如果 Overlay 不实时同步此 margin_r，右侧的半透明遮罩就会错误地延伸覆盖到副 Y 轴的文字上。
        """
        self.margin_l = float(l)
        self.margin_r = float(r)
        self.margin_t = float(t)
        self.margin_b = float(b)
        # 强制请求 Qt 事件循环立即进行重绘
        self.update()

    def set_plot_rect(self, l: float, t: float, w: float, h: float):
        """
        外部强制注入图表实际渲染的矩形坐标。
        接收并保留浮点数，不做 int() 截断以保持亚像素级别的对齐精度。
        """
        try:
            l = float(l)
            t = float(t)
            w = float(w)
            h = float(h)
        except Exception:
            return

        if w <= 0 or h <= 0:
            self._plot_rect_override = None
            return

        # 简单 Clamp 防止极端负数错误，但严格保留浮点精度
        if l < 0: l = 0.0
        if t < 0: t = 0.0

        self._plot_rect_override = (l, t, w, h)

    def clear_plot_rect(self):
        """清除覆盖坐标，回退到基于边距 (margins) 的推导计算"""
        self._plot_rect_override = None

    def _plot_rect(self) -> QRectF:
        """
        [核心计算] 获取当前允许绘制遮罩与线条的有效矩形区域。
        返回 QRectF (浮点矩形) 确保后续画笔的 Antialiasing (抗锯齿) 生效。
        """
        if self._plot_rect_override is not None:
            l, t, w, h = self._plot_rect_override
            if w > 0 and h > 0:
                return QRectF(l, t, w, h)

        # 兜底逻辑：如果外部没有注入精确矩形，则利用当前 Widget 的宽高扣除边距推导
        w = float(self.width())
        h = float(self.height())
        pl = max(0.0, self.margin_l)
        pr = max(pl + 1.0, w - max(0.0, self.margin_r))
        pt = max(0.0, self.margin_t)
        pb = max(pt, h - max(0.0, self.margin_b))

        return QRectF(pl, pt, pr - pl, pb - pt)

    def _x_to_px(self, x: float, rect: QRectF) -> float:
        """
        坐标映射引擎：将数据维度的 X 值转化为屏幕像素的 X 坐标。
        基于现有的 [xmin, xmax] 数据极值和可用绘图宽度 (rect.width()) 线性插值。
        """
        if self.xmax <= self.xmin:
            return rect.left()
        
        t = (float(x) - self.xmin) / (self.xmax - self.xmin)
        # 将比例严格约束在 [0.0, 1.0] 之间，防止线条飞出图表区
        t = max(0.0, min(1.0, t))
        return rect.left() + t * rect.width()

    def set_dpr(self, dpr: float):
        """设置设备像素比 (Device Pixel Ratio)，用于保证高分屏下线条不发虚"""
        self.dpr = float(dpr) if dpr > 0 else 1.0

    # =======================================================
    # 3. 渲染引擎
    # =======================================================
    def paintEvent(self, event):
        """
        Qt 核心绘图回调。
        在此处利用 QPainter 将浮点坐标系下的图形实际光栅化绘制到屏幕上。
        """
        if not self.enabled: return
        rect = self._plot_rect()
        # 防御性判断：绘图区太小则直接跳过
        if rect.width() <= 1 or rect.height() <= 1: return

        # 将逻辑截断值转换为左右像素坐标，并确保 a 是左侧，b 是右侧
        pxL = self._x_to_px(self.l, rect)
        pxR = self._x_to_px(self.r, rect)
        a = min(pxL, pxR)
        b = max(pxL, pxR)

        painter = QPainter(self)
        # 开启抗锯齿，使非整数像素坐标的直线和边缘更加平滑
        painter.setRenderHint(QPainter.RenderHint.Antialiasing, True)

        # [高分屏适配]：计算真实的 1 物理像素的线宽，防止在 200% 缩放的屏幕上画出粗糙的 2px 线条
        pen_w = 0.8 / self.dpr  
        offset = 0.5 * pen_w

        top_y = rect.top()
        bottom_y = rect.bottom()

        # 1) 强力清洗模式字符串，防止外部传入的空格或大小写干扰
        current_mode = str(self.mode).strip().lower()

        # 2) 遮罩渲染策略：
        # 排除 CDF(累积概率)、KDE(核密度) 以及 离散图(discrete/pmf)。
        # 对于这些曲线图或散点图，涂抹界外半透明遮罩会破坏视觉主体；
        # 只要不在排除列表中（例如直方图 hist），就画出左右两块灰白色的失效区遮罩。
        exclude_modes = ["cdf", "kde", "discrete", "pmf"]
        if current_mode not in exclude_modes:
            # 绘制左侧遮罩 (从坐标系左边缘到线 a)
            painter.fillRect(QRectF(rect.left(), top_y, a - rect.left(), rect.height()+0.5), QBrush(self.mask_color))
            # 绘制右侧遮罩 (从线 b 到坐标系右边缘)
            painter.fillRect(QRectF(b, top_y, rect.right() - b, rect.height()+0.5), QBrush(self.mask_color))

        # 3) 绘制精准的垂直定界线 (V-Lines)
        pen = QPen(self.line_color)
        pen.setWidthF(pen_w)  
        painter.setPen(pen)

        # 无论在什么视图模式下，指示截断位置的两条垂直黑线始终都会被绘制出来
        painter.drawLine(QPointF(a, top_y + offset), QPointF(a, bottom_y - offset))
        painter.drawLine(QPointF(b, top_y + offset), QPointF(b, bottom_y - offset))

        painter.end()