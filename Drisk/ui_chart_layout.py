# -*- coding: utf-8 -*-
"""
ui_chart_layout.py

【单一事实源】用于计算并同步“图表边距 margins”，确保以下三处完全一致：
- Plotly Figure 的 layout.margin (l/r/t/b)
- Qt Overlay（如 DriskVLineMaskOverlay）绘制定界线/遮罩所用的边距
- Qt Skeleton（ChartSkeleton）绘制“模拟 Plotly 绘图区边框”所用的边距

设计目标（方案B）：
1) 不再使用任何写死边距（例如过去常见的 L=80, R=40, T=30, B=50）。
2) ui_modeler.py 与 ui_results.py 必须复用同一套边距计算逻辑。
3) 计算逻辑尽量“纯数据”，Qt 同步逻辑可选（便于单测/便于替换）。

本次改造重点（对接修复）：
- ✅ 修复历史 bug：Overlay 的 set_margins 参数顺序为 (l, r, t, b)，Skeleton 的 set_margins 只接收一个对象/tuple/dict
  旧实现把顺序传错、也把 skeleton 的参数传错，导致遮罩/骨架框与真实绘图区不一致。
- ✅ 支持“额外几何因素”：控件高度/标题/字体等导致的额外安全边距（尤其 top）
- ✅ Controller 节流参数 debounce_ms 生效（旧实现固定 30ms）

注意：
- 本模块不依赖业务计算；只负责 UI 布局安全策略。
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Any, Dict, Optional, Tuple, Union


# ============================================================
# 数据结构
# ============================================================

@dataclass(frozen=True)
class Margins:
    """Plotly / Overlay 使用的边距（单位：像素）"""
    l: int
    r: int
    t: int
    b: int

    def as_plotly_dict(self) -> Dict[str, int]:
        """用于 Plotly layout.margin"""
        return {"l": int(self.l), "r": int(self.r), "t": int(self.t), "b": int(self.b)}

    def as_tuple(self) -> Tuple[int, int, int, int]:
        """(l, r, t, b)"""
        return (int(self.l), int(self.r), int(self.t), int(self.b))


@dataclass(frozen=True)
class MarginInputs:
    """
    计算 margins 的输入参数（不带 Qt 依赖）。

    说明：
    - base_*：应用的“设计意图”基础边距（尽量小且稳定）
    - safe_min_*：为了避免截断（tick/边框/抗锯齿）需要的最低安全边距
    - overlay_y2_extra_r：开启 y2 / overlay CDF 时，右侧额外留白
    - extra_*：外部几何提示（例如顶部控件高度、标题高度、额外 padding）
    """
    mode: str                     # e.g. "pdf"/"discrete"/"cdf"/"histogram"
    overlay_y2_enabled: bool      # 是否启用 y2（叠加 CDF 轴）

    # -------------------------
    # 设计意图：基础边距
    # -------------------------
    base_l: int = 60
    base_r: int = 40
    base_t: int = 8     # ✅ top margin 默认不允许为 0（避免缺上边框/截断 y-max tick）
    base_b: int = 45

    # -------------------------
    # y2/overlay：右侧额外留白
    # -------------------------
    overlay_y2_extra_r: int = 60

    # -------------------------
    # 安全下限（像素）
    # top 的安全值建议落在 6~12 量级（本默认：6 + border_pad）
    # -------------------------
    safe_min_t: int = 6
    safe_min_b: int = 35
    safe_min_l: int = 45
    safe_min_r: int = 30

    # -------------------------
    # 边框/抗锯齿补偿（像素）
    # -------------------------
    border_pad: int = 2

    # -------------------------
    # 外部额外边距（几何提示，像素）
    # -------------------------
    extra_l: int = 0
    extra_r: int = 0
    extra_t: int = 0
    extra_b: int = 0


# ============================================================
# Margin 计算
# ============================================================

def _clamp_int(v: Union[int, float], lo: int, hi: int) -> int:
    """安全取整 + 限幅"""
    try:
        iv = int(round(float(v)))
    except Exception:
        iv = int(lo)
    return max(lo, min(hi, iv))


def compute_margins(
    inputs: MarginInputs,
    *,
    # 可选几何提示（不会引入 Qt）
    font_px: Optional[int] = None,
    has_top_title: bool = False,
    allow_top_margin_zero: bool = False,
) -> Margins:
    """
    计算图表边距（margins）。

    核心规则：
    - 默认情况下，top margin 绝不允许为 0（allow_top_margin_zero=False）。
      这是为了避免：Plotly 顶边框缺失、y 轴最大 tick 被截断、annotation 被裁剪等问题。

    参数：
    - font_px：字体像素高度（可由外部 Qt 计算），会小幅增加 t/b 安全边距
    - has_top_title：若图中有 title/顶部注释，增加额外 top 安全边距
    - allow_top_margin_zero：仅用于调试，不建议生产开启
    """
    # 1) 从 base 开始
    l = int(inputs.base_l)
    r = int(inputs.base_r)
    t = int(inputs.base_t)
    b = int(inputs.base_b)

    # 2) 外部几何提示（例如顶部控件、额外 padding）
    l += int(inputs.extra_l)
    r += int(inputs.extra_r)
    t += int(inputs.extra_t)
    b += int(inputs.extra_b)

    # 3) y2/overlay：右侧额外留白（右轴刻度/标题/百分号更宽）
    if inputs.overlay_y2_enabled:
        r += int(inputs.overlay_y2_extra_r)

    # 4) mode 的轻量修正（尽量少、可预期）
    mode = (inputs.mode or "").lower().strip()
    if mode in ("hist", "histogram", "rel_freq", "percent"):
        # 直方图通常 x tick 更密一些，底部略保守
        b = max(b, inputs.safe_min_b)
    elif mode in ("cdf",):
        # CDF：y 轴为 0~1 的百分比标签，top 更敏感，保持稳定
        pass
    elif mode in ("pdf",):
        # PDF：顶部峰值可能接近上边界，top 保守一点
        t = max(t, inputs.safe_min_t)

    # 5) 字体影响：字体越大，tick label/标题越容易溢出
    if font_px is not None:
        fp = _clamp_int(font_px, 8, 28)
        # 12px 作为基准，最大额外加 ~6px
        extra = _clamp_int((fp - 12) * 0.5, 0, 6)
        t += extra
        b += extra

    if has_top_title:
        t += 12  # 保守估计 1 行 title 的高度

    # 6) 施加安全下限 + 边框补偿
    l = max(l, inputs.safe_min_l) + inputs.border_pad
    r = max(r, inputs.safe_min_r) + inputs.border_pad
    b = max(b, inputs.safe_min_b) + inputs.border_pad

    if allow_top_margin_zero:
        t = max(t, 0)
    else:
        # ✅ top 默认必须 > 0，并且满足安全下限
        t = max(t, inputs.safe_min_t) + inputs.border_pad
        if t <= 0:
            t = inputs.safe_min_t + inputs.border_pad

    return Margins(l=l, r=r, t=t, b=b)


# ============================================================
# Plotly Layout 应用器（统一防截断策略）
# ============================================================

def apply_plotly_layout(
    layout: Dict[str, Any],
    axis_style: Optional[Dict[str, Any]],
    margins: Margins,
    *,
    ensure_four_sided_border: bool = True,
    prefer_mirror_axes: bool = True,
) -> Dict[str, Any]:
    """
    将 computed margins 和“轴防截断策略”统一应用到 Plotly layout dict。

    轴防截断策略（尽量温和）：
    - xaxis/yaxis: automargin=True（Plotly 在必要时可略增边距）
    - ticks='outside'（避免 inside 导致边框/label 视觉冲突）
    - mirror=True + showline=True（保证四边框完整）
    """
    if layout is None:
        layout = {}

    # 1) 应用 margins
    layout["margin"] = margins.as_plotly_dict()

    # 2) 轴样式 merge + 安全默认
    if axis_style:
        xaxis = dict(axis_style.get("xaxis", {}))
        yaxis = dict(axis_style.get("yaxis", {}))

        # 安全：允许 Plotly 自适应增加 margin，防止某些平台/字体导致的截断
        xaxis.setdefault("automargin", True)
        yaxis.setdefault("automargin", True)

        # 安全：刻度朝外
        xaxis.setdefault("ticks", "outside")
        yaxis.setdefault("ticks", "outside")

        # 显示轴线（边框）
        if ensure_four_sided_border and prefer_mirror_axes:
            xaxis["showline"] = True
            yaxis["showline"] = True
            xaxis["mirror"] = True
            yaxis["mirror"] = True

        layout["xaxis"] = xaxis
        layout["yaxis"] = yaxis

        # 可选：把轴算法产物缓存到 layout，方便 UI 层读取（例如 FloatingValueWithMag）
        if "x_mag" in axis_style:
            layout["_x_mag"] = axis_style["x_mag"]

    return layout


# ============================================================
# Overlay / Skeleton 同步器（修复参数顺序与签名）
# ============================================================

def sync_overlay_and_skeleton(
    *,
    overlay: Optional[Any],
    skeleton: Optional[Any],
    margins: Margins,
) -> None:
    """
    将 margins 同步给 overlay/skeleton。

    Overlay（DriskVLineMaskOverlay）期望签名：
        set_margins(l, r, t, b)  ← 注意顺序！

    Skeleton（ChartSkeleton）期望签名：
        set_margins(margins_obj_or_tuple_or_dict)

    本函数做最大兼容，不强制某一种实现。
    """
    # 1) overlay
    if overlay is not None:
        if hasattr(overlay, "set_margins"):
            try:
                # ✅ 修复：正确顺序 (l, r, t, b)
                overlay.set_margins(margins.l, margins.r, margins.t, margins.b)
            except Exception:
                pass
        elif hasattr(overlay, "setMargins"):
            # 兜底：极少数情况下是 Qt 的 setMargins(l,t,r,b)
            try:
                overlay.setMargins(margins.l, margins.t, margins.r, margins.b)
            except Exception:
                pass

    # 2) skeleton
    if skeleton is not None:
        # ✅ 首选：ChartSkeleton.set_margins(margins)
        if hasattr(skeleton, "set_margins"):
            try:
                skeleton.set_margins(margins)  # 传对象/tuple/dict均可
                return
            except Exception:
                pass

        # 次选：某些实现可能叫 set_plot_margins
        if hasattr(skeleton, "set_plot_margins"):
            try:
                skeleton.set_plot_margins(margins)
                return
            except Exception:
                pass

        # 兜底：若是 Qt setMargins(l,t,r,b)
        if hasattr(skeleton, "setMargins"):
            try:
                skeleton.setMargins(margins.l, margins.t, margins.r, margins.b)
                return
            except Exception:
                pass

        # 最后兜底：挂属性让 skeleton 自行读取
        try:
            setattr(skeleton, "_plot_margins", margins)
            if hasattr(skeleton, "update"):
                skeleton.update()
        except Exception:
            pass


def sync_all(
    *,
    layout: Optional[Dict[str, Any]] = None,
    axis_style: Optional[Dict[str, Any]] = None,
    figure: Optional[Any] = None,
    webview: Optional[Any] = None,
    overlay: Optional[Any] = None,
    skeleton: Optional[Any] = None,
    margins: Margins,
) -> Optional[Dict[str, Any]]:
    """
    统一一键入口：
    1) 可选：应用 Plotly layout（传 layout/axis_style 或 figure）
    2) 同步 overlay/skeleton

    返回：
    - 若传入 layout，则返回 layout（同一对象）
    - 否则返回 None
    """
    # 1) Plotly layout（可选）
    if figure is not None:
        try:
            # plotly.graph_objects.Figure 支持 update_layout
            figure.update_layout(margin=margins.as_plotly_dict())
        except Exception:
            pass

    if layout is not None:
        apply_plotly_layout(layout, axis_style, margins)

    # 2) overlay/skeleton
    sync_overlay_and_skeleton(overlay=overlay, skeleton=skeleton, margins=margins)

    return layout


# ============================================================
# 可选：Qt Controller（监听 resize/mode/y2/font，节流同步）
# ============================================================

try:
    from PySide6.QtCore import QObject, QTimer, QEvent
except Exception:  # pragma: no cover
    QObject = object  # type: ignore
    QTimer = None     # type: ignore
    QEvent = None     # type: ignore


class ChartLayoutController(QObject):
    """
    可选的小控制器：
    - webview resize / show / font change 时重算 margins
    - mode / y2 开关变化时重算
    - debounce 节流避免频繁重算造成卡顿

    使用建议：
    - 在 ui_results/ui_modeler 初始化时创建
    - 在切换“模式/叠加CDF/y2/字体”时调用 set_mode/set_overlay_y2/set_font_px
    """

    def __init__(
        self,
        *,
        webview: Any,
        overlay: Optional[Any],
        skeleton: Optional[Any],
        inputs: MarginInputs,
        debounce_ms: int = 30,
        font_px: Optional[int] = None,
        has_top_title: bool = False,
    ):
        super().__init__(webview)
        self.webview = webview
        self.overlay = overlay
        self.skeleton = skeleton
        self.inputs = inputs

        self.debounce_ms = int(debounce_ms) if debounce_ms is not None else 30
        self.font_px = font_px
        self.has_top_title = has_top_title

        self._last_margins: Optional[Margins] = None

        # 节流计时器
        if QTimer is None:
            self._timer = None
        else:
            self._timer = QTimer(self)
            self._timer.setSingleShot(True)
            self._timer.timeout.connect(self._do_sync)

        # 监听 resize/show/font change
        try:
            self.webview.installEventFilter(self)
        except Exception:
            pass

        self.request_sync()

    # -------------------------
    # 外部可调用的“状态更新”
    # -------------------------
    def set_mode(self, mode: str) -> None:
        self.inputs = MarginInputs(
            mode=mode,
            overlay_y2_enabled=self.inputs.overlay_y2_enabled,
            base_l=self.inputs.base_l,
            base_r=self.inputs.base_r,
            base_t=self.inputs.base_t,
            base_b=self.inputs.base_b,
            overlay_y2_extra_r=self.inputs.overlay_y2_extra_r,
            safe_min_t=self.inputs.safe_min_t,
            safe_min_b=self.inputs.safe_min_b,
            safe_min_l=self.inputs.safe_min_l,
            safe_min_r=self.inputs.safe_min_r,
            border_pad=self.inputs.border_pad,
            extra_l=self.inputs.extra_l,
            extra_r=self.inputs.extra_r,
            extra_t=self.inputs.extra_t,
            extra_b=self.inputs.extra_b,
        )
        self.request_sync()

    def set_overlay_y2(self, enabled: bool) -> None:
        self.inputs = MarginInputs(
            mode=self.inputs.mode,
            overlay_y2_enabled=bool(enabled),
            base_l=self.inputs.base_l,
            base_r=self.inputs.base_r,
            base_t=self.inputs.base_t,
            base_b=self.inputs.base_b,
            overlay_y2_extra_r=self.inputs.overlay_y2_extra_r,
            safe_min_t=self.inputs.safe_min_t,
            safe_min_b=self.inputs.safe_min_b,
            safe_min_l=self.inputs.safe_min_l,
            safe_min_r=self.inputs.safe_min_r,
            border_pad=self.inputs.border_pad,
            extra_l=self.inputs.extra_l,
            extra_r=self.inputs.extra_r,
            extra_t=self.inputs.extra_t,
            extra_b=self.inputs.extra_b,
        )
        self.request_sync()

    def set_font_px(self, font_px: Optional[int]) -> None:
        self.font_px = font_px
        self.request_sync()

    def set_has_top_title(self, has_top_title: bool) -> None:
        self.has_top_title = bool(has_top_title)
        self.request_sync()

    def set_extra_margins(self, *, extra_l: int = 0, extra_r: int = 0, extra_t: int = 0, extra_b: int = 0) -> None:
        self.inputs = MarginInputs(
            mode=self.inputs.mode,
            overlay_y2_enabled=self.inputs.overlay_y2_enabled,
            base_l=self.inputs.base_l,
            base_r=self.inputs.base_r,
            base_t=self.inputs.base_t,
            base_b=self.inputs.base_b,
            overlay_y2_extra_r=self.inputs.overlay_y2_extra_r,
            safe_min_t=self.inputs.safe_min_t,
            safe_min_b=self.inputs.safe_min_b,
            safe_min_l=self.inputs.safe_min_l,
            safe_min_r=self.inputs.safe_min_r,
            border_pad=self.inputs.border_pad,
            extra_l=int(extra_l),
            extra_r=int(extra_r),
            extra_t=int(extra_t),
            extra_b=int(extra_b),
        )
        self.request_sync()

    # -------------------------
    # 节流同步
    # -------------------------
    def request_sync(self) -> None:
        if self._timer is None:
            self._do_sync()
            return
        self._timer.start(max(0, int(self.debounce_ms)))

    def eventFilter(self, obj, event):  # type: ignore[override]
        if obj is self.webview and QEvent is not None:
            et = event.type()
            # Resize/Show/Move：窗口大小变化
            # FontChange/StyleChange：字体/缩放变化（部分平台会触发）
            if et in (QEvent.Resize, QEvent.Show, QEvent.Move, QEvent.FontChange, QEvent.StyleChange, QEvent.LayoutRequest):
                self.request_sync()
        return False

    def _do_sync(self) -> None:
        margins = compute_margins(
            self.inputs,
            font_px=self.font_px,
            has_top_title=self.has_top_title,
        )
        if margins != self._last_margins:
            self._last_margins = margins
            sync_all(webview=self.webview, overlay=self.overlay, skeleton=self.skeleton, margins=margins)
