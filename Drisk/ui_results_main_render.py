# ui_results_main_render.py
"""
本模块提供结果视图（Results Views）中基础/核心图表的主渲染服务 ResultsMainRenderService。

主要功能：
1. 视图通道路由（Render Channels）：负责处理直方图 (Histogram)、相对频数图 (RelFreq)、离散柱状图 (Discrete)、概率密度图 (PDF) 和累积分布图 (CDF) 的数据组装与渲染触发。
2. 曲线安全截断（Bounded KDE）：提供数学方法，将平滑的核密度估计（KDE）曲线严格限制在实际数据的极值范围内，防止视觉上出现虚假的“拖尾”。
3. 理论分布融合：安全地从统计分布对象中提取理论 PDF/CDF 曲线，并与实际模拟数据进行对齐绘制。
4. 高级非参置信区间：在 CDF 视图中，集成 Beta 分布精确解与 Bootstrap BCa（偏差校正及加速）算法，绘制统计置信带（CI Overlays）。
"""

from __future__ import annotations

import drisk_env
import json
import math
import os
import traceback

import numpy as np
import plotly
import plotly.graph_objects as go
from scipy.stats import beta, norm

from drisk_charting import DriskChartFactory
from ui_results_data_service import ResultsDataService
from ui_shared import DRISK_COLOR_CYCLE, DriskMath, build_plotly_axis_style


# =======================================================
# 1. 主渲染服务 (无状态核心渲染管线)
# =======================================================
class ResultsMainRenderService:
    """
    负责基础视图层的单向数据流渲染。
    将 UI 层的状态字典 (series_map, options) 转化为图表工厂 (DriskChartFactory) 
    可消费的标准结构，并最终生成前端 Plotly 所需的 JSON 数据。
    """

    # =======================================================
    # 2. 布局同步与边界处理工具
    # =======================================================
    @staticmethod
    def _apply_chart_margins(dialog, margin_right: int):
        """
        [UI 联动逻辑] 统一图表边缘留白：
        在向浏览器引擎推入新图表之前，强制同步上方滑块、标题容器与图表区自身的边距。
        确保多图层（如叠加双 Y 轴时）X 轴两端完美垂直对齐。
        """
        dialog.range_slider.setMargins(dialog.MARGIN_L, margin_right)
        dialog.chart_title_label.setContentsMargins(dialog.MARGIN_L, 0, margin_right, 0)
        dialog._update_x_edits_pos()
        
        # 同步半透明遮罩层（Overlay）以匹配图表实际绘图区
        if hasattr(dialog, "overlay") and getattr(dialog, "overlay", None):
            if hasattr(dialog.overlay, "set_margins"):
                dialog.overlay.set_margins(dialog.MARGIN_L, margin_right, dialog.MARGIN_T, dialog.MARGIN_B)
        dialog._chart_ctrl.schedule_rect_sync()

    @staticmethod
    def _build_series_display_name_map(dialog, keys) -> dict[str, str]:
        try:
            if hasattr(dialog, "_build_series_display_name_map"):
                mapping = dialog._build_series_display_name_map(keys)
                if isinstance(mapping, dict):
                    return {str(k): str(v) for k, v in mapping.items()}
        except Exception:
            pass
        return {str(k): str(k) for k in list(keys or [])}

    @staticmethod
    def _display_series_name(name_map: dict[str, str], key: object) -> str:
        key_text = str(key)
        return str(name_map.get(key_text, key_text))

    @staticmethod
    def _extract_series_bounds(dialog, key):
        """
        提取特定数据序列的绝对最小/最大值边界。
        用于后续的曲线截断，确保平滑曲线不会延伸到没有实际数据的区域。
        """
        series_map = getattr(dialog, "series_map", {})
        raw = series_map.get(key, None)
        cleaned = ResultsDataService.clean_series(raw)
        if cleaned is None or len(cleaned) == 0:
            return None
        try:
            data_min = float(cleaned.min())
            data_max = float(cleaned.max())
        except Exception:
            return None
            
        if not np.isfinite(data_min) or not np.isfinite(data_max):
            return None
        if data_min > data_max:
            data_min, data_max = data_max, data_min
        return data_min, data_max

    @staticmethod
    def _build_bounded_kde_curve(x_grid, y_vals, bounds):
        """
        [视觉优化核心] KDE 曲线截断与闭合算法：
        底层的核密度估算（KDE）在两端会自然产生平滑拖尾（哪怕数据严格 > 0，KDE 也可能延伸到负轴）。
        此方法执行以下操作：
        1. 剔除无效点，并保证密度 Y 值 >= 0。
        2. 若提供了绝对边界 (bounds)，剔除边界外的估计点。
        3. [关键体验] 在曲线两头各追加一个坐标为 (x_min, 0) 和 (x_max, 0) 的锚点。
           这使得渲染出的曲线能在两端形成完美的垂直“落地”闭合效果，而不是悬停在半空。
        """
        try:
            x_arr = np.asarray(x_grid, dtype=float).ravel()
            y_arr = np.asarray(y_vals, dtype=float).ravel()
        except Exception:
            return np.array([], dtype=float), np.array([], dtype=float)

        if x_arr.size == 0 or y_arr.size == 0:
            return np.array([], dtype=float), np.array([], dtype=float)

        n = min(x_arr.size, y_arr.size)
        x_arr = x_arr[:n]
        y_arr = y_arr[:n]

        # 过滤非有限数字并规整负密度（KDE 浮点误差导致的小于 0 的值）
        finite_mask = np.isfinite(x_arr) & np.isfinite(y_arr)
        if not np.any(finite_mask):
            return np.array([], dtype=float), np.array([], dtype=float)

        x_arr = x_arr[finite_mask]
        y_arr = np.maximum(y_arr[finite_mask], 0.0)
        if x_arr.size == 0:
            return np.array([], dtype=float), np.array([], dtype=float)

        # 保证数组单调递增
        order = np.argsort(x_arr)
        x_arr = x_arr[order]
        y_arr = y_arr[order]

        if bounds is None:
            return x_arr, y_arr

        data_min, data_max = bounds
        if not np.isfinite(data_min) or not np.isfinite(data_max):
            return x_arr, y_arr
        if data_min > data_max:
            data_min, data_max = data_max, data_min

        span = float(x_arr[-1] - x_arr[0]) if x_arr.size > 1 else 0.0
        eps = max(1e-12, span * 1e-12)

        y_min = float(np.interp(data_min, x_arr, y_arr, left=0.0, right=0.0))
        y_max = float(np.interp(data_max, x_arr, y_arr, left=0.0, right=0.0))
        y_min = max(0.0, y_min)
        y_max = max(0.0, y_max)

        if abs(data_max - data_min) <= eps:
            return (
                np.array([data_min], dtype=float),
                np.array([max(y_min, y_max)], dtype=float),
            )

        inner = (x_arr >= data_min - eps) & (x_arr <= data_max + eps)
        x_inner = x_arr[inner]
        y_inner = y_arr[inner]
        if x_inner.size == 0:
            return (
                np.array([data_min, data_max], dtype=float),
                np.array([y_min, y_max], dtype=float),
            )

        x_plot = np.asarray(x_inner, dtype=float)
        y_plot = np.asarray(y_inner, dtype=float)

        # Keep the displayed curve strictly inside the sample interval without hard zero anchors.
        x_plot[0] = data_min
        y_plot[0] = y_min
        x_plot[-1] = data_max
        y_plot[-1] = y_max
        return x_plot, y_plot

    # =======================================================
    # 3. 理论分布处理适配器
    # =======================================================
    @staticmethod
    def _should_show_theory_pdf(dialog) -> bool:
        """检查用户是否勾选了展示理论 PDF 曲线（针对拟合功能）"""
        if hasattr(dialog, "_theory_visibility_flags"):
            try:
                show_pdf, _ = dialog._theory_visibility_flags()
                return bool(show_pdf)
            except Exception:
                return False
        return False

    @staticmethod
    def _should_show_theory_cdf(dialog) -> bool:
        """检查用户是否勾选了展示理论 CDF 曲线"""
        if hasattr(dialog, "_theory_visibility_flags"):
            try:
                _, show_cdf = dialog._theory_visibility_flags()
                return bool(show_cdf)
            except Exception:
                return False
        return False

    @staticmethod
    def _is_theory_source_discrete(dialog) -> bool:
        """检查理论分布对象是否为离散型（如二项、泊松分布），影响绘图连线模式"""
        if hasattr(dialog, "_is_theory_source_discrete"):
            try:
                return bool(dialog._is_theory_source_discrete())
            except Exception:
                return False
        return False

    @staticmethod
    def _safe_theory_pdf_vec(dist, xs):
        """
        分布计算适配器：向下兼容 scipy.stats 对象及内部封装的拟合分布对象。
        支持向量化运算(pdf_vec/pmf_vec) 以提升大批量 x_grid 计算性能；
        若不支持，则回退到列表推导式的单点求值。
        """
        if dist is None:
            return np.array([], dtype=float)
        try:
            if hasattr(dist, "pdf_vec"):
                res = dist.pdf_vec(xs)
            elif hasattr(dist, "pdf"):
                # 内部分布对象的 pdf() 是标量接口，必须逐点调用
                res = np.array([float(dist.pdf(x)) for x in xs], dtype=float)
            elif hasattr(dist, "pmf_vec"):
                res = dist.pmf_vec(xs)
            elif hasattr(dist, "pmf"):
                res = np.array([float(dist.pmf(x)) for x in xs], dtype=float)
            else:
                return np.array([], dtype=float)

            arr = np.asarray(res, dtype=float).ravel()
            # 过滤非数字并强制密度 >= 0
            return np.where(np.isfinite(arr), np.maximum(arr, 0.0), 0.0)
        except Exception:
            return np.array([], dtype=float)

    @staticmethod
    def _safe_theory_cdf_vec(dist, xs):
        """同上，针对累积分布函数 (CDF) 的安全向量化提取"""
        if dist is None:
            return np.array([], dtype=float)
        try:
            if hasattr(dist, "cdf_vec"):
                res = dist.cdf_vec(xs)
            elif hasattr(dist, "cdf"):
                res = np.array([float(dist.cdf(x)) for x in xs], dtype=float)
            else:
                return np.array([], dtype=float)
            arr = np.asarray(res, dtype=float).ravel()
            return np.where(np.isfinite(arr), arr, 0.0)
        except Exception:
            return np.array([], dtype=float)

    @staticmethod
    def _build_theory_pdf_trace(dialog, x_grid=None):
        """为当前视图构建一条理论概率密度/质量黑线 (Trace)"""
        dist = getattr(dialog, "theory_dist_obj", None)
        if dist is None:
            return None

        # 如果没有传入复用的 X 轴网格，则根据视图区间动态生成 500 个打点
        if x_grid is None:
            x_vals = np.linspace(float(dialog.x_range_min), float(dialog.x_range_max), 500)
        else:
            x_vals = np.asarray(x_grid, dtype=float).ravel()
            if x_vals.size == 0:
                x_vals = np.linspace(float(dialog.x_range_min), float(dialog.x_range_max), 500)

        # 若分布有有限支撑（如 uniform），裁剪到支撑内并加密，再加边界 y=0 锚点形成矩形框
        _bounded = False
        try:
            d_min = float(dist.min_val())
            d_max = float(dist.max_val())
            if np.isfinite(d_min) and np.isfinite(d_max):
                _bounded = True
                eps = max(1e-10, (d_max - d_min) * 1e-6)
                inner = np.linspace(d_min, d_max, 200)
                x_vals = np.unique(np.concatenate([
                    inner,
                    x_vals[(x_vals >= d_min) & (x_vals <= d_max)],
                ]))
        except Exception:
            pass

        y_vals = ResultsMainRenderService._safe_theory_pdf_vec(dist, x_vals)
        if y_vals.size == 0:
            return None

        n = min(x_vals.size, y_vals.size)
        x_vals = x_vals[:n]
        y_vals = y_vals[:n]
        if n == 0:
            return None

        if _bounded:
            x_vals = np.concatenate([[d_min - eps], x_vals, [d_max + eps]])
            y_vals = np.concatenate([[0.0], y_vals, [0.0]])

        print("_build_theory_pdf_trace: 理论 PDF 曲线已构建，打点数=%d，Y 值范围=[%.4g, %.4g]" % (len(x_vals), np.min(y_vals), np.max(y_vals)))
        return DriskChartFactory.get_pdf_line_trace(
            x=x_vals,
            y=y_vals,
            color="#000000",
            name="\u7406\u8bba\u5206\u5e03",  # "理论分布"
        )

    @staticmethod
    def _build_theory_discrete_group(dialog, support_x):
        dist = getattr(dialog, "theory_dist_obj", None)
        if dist is None:
            return None

        xs = np.asarray(support_x, dtype=float).ravel()
        if xs.size == 0:
            return None
        xs = xs[np.isfinite(xs)]
        if xs.size == 0:
            return None
        xs = np.unique(np.sort(xs))
        if xs.size == 0:
            return None

        ys = ResultsMainRenderService._safe_theory_pdf_vec(dist, xs)
        if ys.size == 0:
            return None

        n = min(xs.size, ys.size)
        if n == 0:
            return None
        xs = xs[:n]
        ys = ys[:n]

        return {
            "x": xs,
            "y": ys,
            "name": "\u7406\u8bba\u5206\u5e03",
            "color": "#000000",
            "fill_color": "#000000",
            "fill_opacity": 0.5,
            "cdf_color": "#000000",
            "filled": True,
        }

    @staticmethod
    def _build_theory_cdf_group(dialog):
        """为 CDF 视图构建对应的理论 CDF 黑线字典结构"""
        dist = getattr(dialog, "theory_dist_obj", None)
        if dist is None:
            return None

        is_discrete_source = ResultsMainRenderService._is_theory_source_discrete(dialog)
        # 离散分布必须打点在整数上，连续分布使用均匀划分的高密网格
        if is_discrete_source:
            xs = np.arange(
                int(math.floor(float(dialog.view_min))),
                int(math.ceil(float(dialog.view_max))) + 1,
                dtype=int,
            )
        else:
            xs = np.linspace(float(dialog.view_min), float(dialog.view_max), 800)

        ys = ResultsMainRenderService._safe_theory_cdf_vec(dist, xs)
        if ys.size == 0:
            return None

        n = min(xs.size, ys.size)
        xs = xs[:n]
        ys = ys[:n]

        # 如果理论分布具备严格上下界，裁切掉越界的点（设为 0% 或 100%）
        try:
            d_min = float(dist.min_val())
            d_max = float(dist.max_val())
            ys = np.where(xs < d_min, 0.0, ys)
            ys = np.where(xs > d_max, 1.0, ys)
        except Exception:
            pass

        return {
            "x": xs,
            "y": ys,
            "name": "\u7406\u8bba\u5206\u5e03",
            "color": "#000000",
            "curve_color": "#000000",
            "curve_opacity": 1.0,
            "curve_width": 1.8,
            "dash": "solid",
            "line_shape": "hv" if is_discrete_source else "linear",  # 离散分布使用阶梯线(hv)
        }

    # =======================================================
    # 4. 核心渲染通道 (Render Channels)
    # =======================================================
    @staticmethod
    def render_channel_hist(dialog, view_mode: str):
        """
        [通道] 渲染直方图 / 相对频数图。
        负责协调主条形图与可能叠加在其上的两条线：
        1. 实验数据的平滑 KDE 曲线（需查询异步缓存或发起请求）。
        2. 拟合的理论 PDF 曲线。
        """
        if view_mode not in ("histogram", "relfreq"):
            view_mode = "histogram"
        display_keys = list(getattr(dialog, "display_keys", [dialog.current_key]))
        display_name_map = ResultsMainRenderService._build_series_display_name_map(dialog, display_keys)

        # 处理因叠加副 Y 轴（如叠加展示 CDF 线）导致的右侧边距膨胀
        is_cdf_overlay = getattr(dialog, "_show_cdf_overlay", False)
        is_kde_overlay = getattr(dialog, "_show_kde_overlay", False)
        current_margin_r = dialog.MARGIN_R + (dialog.ADDITIONAL_Y2_MARGIN if is_cdf_overlay else 0)
        ResultsMainRenderService._apply_chart_margins(dialog, current_margin_r)

        pdf_traces = []

        # 挂载：实验数据平滑线 KDE
        if is_kde_overlay and len(display_keys) == 1:
            key = display_keys[0]
            if key not in dialog._kde_cache:
                # 缓存穿透：向线程池发送异步请求，此时暂时不挂载曲线，等待回调触发重绘
                arr = dialog.series_map[key].dropna().values
                if len(arr) > 0:
                    real_min, real_max = float(np.min(arr)), float(np.max(arr))
                else:
                    real_min, real_max = dialog.x_range_min, dialog.x_range_max
                dialog.job_controller.request_hist_kde(key, dialog.series_map[key], real_min, real_max, 500)
            else:
                # 命中缓存：提取数据，执行边界截断修剪，组装 Trace
                x_grid, y_vals = dialog._kde_cache[key]
                if x_grid is not None and y_vals is not None:
                    bounds = ResultsMainRenderService._extract_series_bounds(dialog, key)
                    x_plot, y_plot = ResultsMainRenderService._build_bounded_kde_curve(x_grid, y_vals, bounds)
                    if x_plot.size == 0 or y_plot.size == 0:
                        x_plot, y_plot = np.asarray(x_grid, dtype=float), np.asarray(y_vals, dtype=float)
                    trace = DriskChartFactory.get_pdf_line_trace(
                        x=x_plot,
                        y=y_plot,
                        color="#808080",
                        name="PDF",
                    )
                    if trace:
                        pdf_traces = [trace]
                        
        # 挂载：理论分布曲线
        if ResultsMainRenderService._should_show_theory_pdf(dialog):
            print("render_channel_hist: 用户已启用理论 PDF 曲线，正在构建叠加 Trace")
            trace = ResultsMainRenderService._build_theory_pdf_trace(dialog)
            if trace:
                pdf_traces.append(trace)

        # 调取底层库生成图形配置
        res = DriskChartFactory.build_simulation_histogram(
            series_map=getattr(dialog, "series_map", {}),
            display_keys=display_keys,
            fill_keys=display_keys[:3],  # 限制最多只有前3个层级拥有色彩填充防遮挡
            style_map=getattr(dialog, "_series_style", {}),
            view_mode=view_mode,
            x_range=[dialog.x_range_min, dialog.x_range_max],
            x_dtick=float(dialog.x_dtick),
            margins=(dialog.MARGIN_L, current_margin_r, dialog.MARGIN_T, dialog.MARGIN_B),
            show_overlay_cdf=is_cdf_overlay,
            pdf_traces=pdf_traces,
            forced_mag=dialog.manual_mag,
            display_name_map=display_name_map,
            label_overrides=getattr(dialog, "_label_settings_config", None),
            axis_numeric_flags=getattr(dialog, "_label_axis_numeric", None),
        )
        # 将生成的 Plotly JSON 推入浏览器组件
        dialog._load_plotly_html(res["plot_json"], res["js_mode"], res["js_logic"])

    @staticmethod
    def render_channel_discrete(dialog):
        """
        [通道] 渲染离散柱状图。
        专为分类标识或离散度高（唯一整数值较少）的数据集设计。
        利用 value_counts 自动规约相同数值的频数并计算百分比，保留了情景叠加机制。
        """
        def _build_discrete_group_points(series):
            vals = np.asarray(series.values, dtype=float)
            vals = vals[np.isfinite(vals)]
            if vals.size == 0:
                return np.array([], dtype=float), np.array([], dtype=float)

            # Output-series discrete data can carry floating noise; snap to the detected
            # discrete step first so shared bar-width logic sees stable support spacing.
            if str(getattr(dialog, "_base_kind", "output") or "output").lower() == "output":
                step_val = float(getattr(dialog, "_discrete_step", 1.0) or 1.0)
                if (not np.isfinite(step_val)) or step_val <= 0.0:
                    step_val = 1.0
                vals = np.round(vals / step_val) * step_val
                vals = np.round(vals, 12)
                if step_val >= (1.0 - 1e-9):
                    vals = np.round(vals)

            xs, cnt = np.unique(vals, return_counts=True)
            if xs.size == 0:
                return np.array([], dtype=float), np.array([], dtype=float)
            probs = (cnt.astype(float) / max(1.0, float(np.sum(cnt)))) * 100.0
            return xs.astype(float), probs.astype(float)

        display_keys = list(getattr(dialog, "display_keys", [dialog.current_key]))
        display_name_map = ResultsMainRenderService._build_series_display_name_map(dialog, display_keys)
        groups = []

        is_cdf_overlay = getattr(dialog, "_show_cdf_overlay", False)
        current_margin_r = dialog.MARGIN_R + (dialog.ADDITIONAL_Y2_MARGIN if is_cdf_overlay else 0)

        for key in display_keys:
            series = dialog._get_series(key)
            if len(series) == 0:
                continue
            xs, ys = _build_discrete_group_points(series)
            if xs.size == 0 or ys.size == 0:
                continue

            style = dialog._series_style.get(key, {})
            groups.append(
                {
                    "x": xs,
                    "y": ys,
                    "name": ResultsMainRenderService._display_series_name(display_name_map, key),
                    "color": style.get("color", DRISK_COLOR_CYCLE[0]),
                    "fill_color": style.get("fill_color", style.get("color", DRISK_COLOR_CYCLE[0])),
                    "fill_opacity": style.get("fill_opacity", 1.0),
                    "cdf_color": style.get("cdf_color", DRISK_COLOR_CYCLE[0]),
                    "filled": key in display_keys[:3],
                }
            )

        if (
            groups
            and ResultsMainRenderService._should_show_theory_pdf(dialog)
            and ResultsMainRenderService._is_theory_source_discrete(dialog)
        ):
            print(
                "render_channel_discrete: 准备叠加理论分布，当前数据组数=%d" % len(groups)
            )
            support_parts = []
            for group in groups:
                x_arr = np.asarray(group.get("x", []), dtype=float).ravel()
                if x_arr.size > 0:
                    support_parts.append(x_arr)
            if support_parts:
                support_x = np.concatenate(support_parts)
                theory_group = ResultsMainRenderService._build_theory_discrete_group(dialog, support_x)
                if theory_group is not None:
                    print(
                        "render_channel_discrete: 理论分布组已生成，支撑点数=%d，概率值范围=[%.4g, %.4g]" % (
                            len(theory_group.get("x", [])),
                            float(np.min(theory_group.get("y", [0]))) if len(theory_group.get("y", [])) > 0 else 0.0,
                            float(np.max(theory_group.get("y", [0]))) if len(theory_group.get("y", [])) > 0 else 0.0
                        )
                    )
                    groups.append(theory_group)
                else:
                    print("render_channel_discrete: 理论分布组生成失败（返回 None）")

        res = DriskChartFactory.build_discrete_bar(
            data_groups=groups,
            x_range=[dialog.x_range_min, dialog.x_range_max],
            x_dtick=float(dialog.x_dtick),
            margins=(dialog.MARGIN_L, current_margin_r, dialog.MARGIN_T, dialog.MARGIN_B),
            show_overlay_cdf=is_cdf_overlay,
            forced_mag=dialog.manual_mag,
            label_overrides=getattr(dialog, "_label_settings_config", None),
            axis_numeric_flags=getattr(dialog, "_label_axis_numeric", None),
        )

        ResultsMainRenderService._apply_chart_margins(dialog, current_margin_r)
        dialog._load_plotly_html(res["plot_json"], res["js_mode"], res["js_logic"])

    @staticmethod
    def render_channel_pdf(dialog):
        """
        [通道] 纯概率密度函数视图 (KDE PDF)。
        触发防抖保护的后台 KDE 核密度估计计算。
        如果当前缓存签名匹配直接绘制，否则下发立即执行命令 `immediate=True`。
        """
        if getattr(dialog, "_pdf_x_grid", None) is not None and getattr(dialog, "_pdf_y_map", None):
            ResultsMainRenderService.render_pdf_from_cache(dialog)
            return

        keys = list(getattr(dialog, "display_keys", [dialog.current_key]))
        x_min = float(getattr(dialog, "x_range_min", getattr(dialog, "abs_min", 0.0)))
        x_max = float(getattr(dialog, "x_range_max", getattr(dialog, "abs_max", 1.0)))
        
        # 兜底：如果外部轴计算有误，根据全部序列的合并池求取安全的 Min/Max
        if not np.isfinite(x_min) or not np.isfinite(x_max) or x_max <= x_min:
            all_vals = []
            for key in keys:
                all_vals.extend(dialog.series_map[key].dropna().values)
            if len(all_vals) > 0:
                x_min, x_max = float(np.min(all_vals)), float(np.max(all_vals))
            else:
                x_min, x_max = 0.0, 1.0

        # 保存本次请求签名，防止异步回调脏写
        dialog._pdf_keys_sig = (tuple(keys), x_min, x_max, 800)
        dialog.job_controller.request_pdf_kde(
            keys,
            getattr(dialog, "series_map", {}),
            x_min,
            x_max,
            800,  # 800 个插值打点以保证平滑度
            immediate=True,
        )

    @staticmethod
    def render_channel_cdf(dialog):
        """
        [通道] 渲染累积分布图 (CDF)。
        [核心统计逻辑]：包含极其重要的置信区间 (DKW / Bootstrap) 计算体系。
        """
        display_keys = list(getattr(dialog, "display_keys", [dialog.current_key]))
        display_name_map = ResultsMainRenderService._build_series_display_name_map(dialog, display_keys)
        groups = []
        is_dkw_overlay = getattr(dialog, "_show_dkw_overlay", False) # 用户是否启用了置信带
        alpha = 1.0 - getattr(dialog, "cdf_ci_level", 0.95)          # 显著性水平 (如 0.05)
        is_discrete_cdf = bool(getattr(dialog, "is_discrete_view", False))

        for key in display_keys:
            series = dialog._get_series(key)
            if len(series) == 0:
                continue
            
            # 使用 DataService 提供的高性能 CDF 降采样算法生成点集
            xs, ys = ResultsDataService.build_cdf_points(series.values, discrete=is_discrete_cdf)
            if xs is None or ys is None:
                continue
            style = dialog._series_style.get(key, {})
            base_color = style.get("color", DRISK_COLOR_CYCLE[0])
            curve_color = style.get("curve_color", base_color)
            xs = np.asarray(xs, dtype=float).ravel()
            ys = np.asarray(ys, dtype=float).ravel()
            if is_discrete_cdf and ys.size > 0:
                ys = np.maximum.accumulate(np.clip(ys, 0.0, 1.0))

            y_upper, y_lower = None, None
            if is_dkw_overlay:
                engine_mode = getattr(dialog, "ci_engine_mode", "fast")
                ci_cache_key = (key, engine_mode, round(alpha, 6))
                
                # 优先查询命中本地已计算好的包络带边界
                if hasattr(dialog, "_cdf_ci_cache") and ci_cache_key in dialog._cdf_ci_cache:
                    y_upper, y_lower = dialog._cdf_ci_cache[ci_cache_key]
                else:
                    n_valid = len(series.dropna())
                    if n_valid > 0:
                        if engine_mode == "fast":
                            # === 精确 Beta 分布置信区间 (Clopper-Pearson法) ===
                            # 由于 CDF 实际上是一系列二项试验(小于阈值/不小于阈值)的集合，
                            # 此处使用精确的 Beta 分布逆函数计算每个点的 95% 置信上下限。
                            k_count = ys * n_valid
                            y_lower = np.zeros_like(ys)
                            y_upper = np.ones_like(ys)
                            
                            # 过滤边界以防除以零或溢出
                            mask_low = k_count > 0
                            y_lower[mask_low] = beta.ppf(
                                alpha / 2,
                                k_count[mask_low],
                                n_valid - k_count[mask_low] + 1,
                            )
                            mask_up = k_count < n_valid
                            y_upper[mask_up] = beta.ppf(
                                1 - alpha / 2,
                                k_count[mask_up] + 1,
                                n_valid - k_count[mask_up],
                            )
                        else:
                            # === 非参数 Bootstrap BCa 重抽样算法 ===
                            # 对于偏态数据或特殊分布，使用 1000 次重抽样计算经验置信区间。
                            data_clean = series.dropna().values
                            B = 1000
                            idx = np.random.choice(n_valid, (B, n_valid), replace=True)
                            boot_data = data_clean[idx]
                            boot_data.sort(axis=1)

                            # 计算重抽样样本群的经验 CDF 矩阵
                            boot_cdf = np.zeros((B, len(xs)))
                            for i in range(B):
                                boot_cdf[i] = np.searchsorted(boot_data[i], xs, side="right") / n_valid

                            if engine_mode == "bca":
                                # BCa (偏差校正及加速) 置信区间计算
                                try:
                                    p_hat = ys
                                    mask_valid = (p_hat > 0) & (p_hat < 1)
                                    # 1. 计算偏差校正因子 z0
                                    p_less = np.sum(boot_cdf < p_hat, axis=0) / B
                                    p_eq = np.sum(boot_cdf == p_hat, axis=0) / B
                                    p_inv = np.clip(p_less + 0.5 * p_eq, 1e-6, 1 - 1e-6)
                                    z0 = norm.ppf(p_inv)

                                    # 2. 计算加速因子 a_factor (用 jackknife 类似理论近似)
                                    p_safe = np.clip(p_hat, 1e-6, 1 - 1e-6)
                                    a_factor = (1.0 - 2.0 * p_safe) / (
                                        6.0 * np.sqrt(n_valid * p_safe * (1.0 - p_safe))
                                    )

                                    # 3. 计算校正后的分位数
                                    z_alpha = norm.ppf(alpha / 2)
                                    z_1_alpha = norm.ppf(1 - alpha / 2)
                                    a1 = z0 + (z0 + z_alpha) / (1 - a_factor * (z0 + z_alpha))
                                    a2 = z0 + (z0 + z_1_alpha) / (1 - a_factor * (z0 + z_1_alpha))
                                    q1 = np.clip(norm.cdf(a1), 0, 1)
                                    q2 = np.clip(norm.cdf(a2), 0, 1)

                                    # 4. 提取对应的边界值
                                    y_lower = np.copy(ys)
                                    y_upper = np.copy(ys)
                                    for j in range(len(ys)):
                                        if mask_valid[j]:
                                            y_lower[j] = np.percentile(boot_cdf[:, j], q1[j] * 100)
                                            y_upper[j] = np.percentile(boot_cdf[:, j], q2[j] * 100)
                                except Exception as exc:
                                    print(f"CDF 带 BCa 异常，降级为普通重抽样: {exc}")
                                    y_lower = np.percentile(boot_cdf, (alpha / 2) * 100, axis=0)
                                    y_upper = np.percentile(boot_cdf, (1 - alpha / 2) * 100, axis=0)
                            else:
                                # 普通百分位数 Bootstrap 降级逻辑
                                y_lower = np.percentile(boot_cdf, (alpha / 2) * 100, axis=0)
                                y_upper = np.percentile(boot_cdf, (1 - alpha / 2) * 100, axis=0)

                        if not hasattr(dialog, "_cdf_ci_cache"):
                            dialog._cdf_ci_cache = {}
                        dialog._cdf_ci_cache[ci_cache_key] = (y_upper, y_lower)

            if y_upper is not None and y_lower is not None:
                y_upper = np.asarray(y_upper, dtype=float).ravel()
                y_lower = np.asarray(y_lower, dtype=float).ravel()
                m = min(xs.size, ys.size, y_upper.size, y_lower.size)
                if m <= 0:
                    y_upper, y_lower = None, None
                else:
                    xs = xs[:m]
                    ys = ys[:m]
                    y_upper = y_upper[:m]
                    y_lower = y_lower[:m]
                    if is_discrete_cdf:
                        # Discrete CDF envelopes must remain monotone and ordered on support points.
                        y_lower = np.clip(np.where(np.isfinite(y_lower), y_lower, 0.0), 0.0, 1.0)
                        y_upper = np.clip(np.where(np.isfinite(y_upper), y_upper, 1.0), 0.0, 1.0)
                        y_lower = np.maximum.accumulate(y_lower)
                        y_upper = np.maximum.accumulate(y_upper)
                        y_upper = np.maximum(y_upper, y_lower)

            groups.append(
                {
                    "x": xs,
                    "y": ys,
                    "y_upper": y_upper,
                    "y_lower": y_lower,
                    "name": ResultsMainRenderService._display_series_name(display_name_map, key),
                    "color": curve_color,
                    "curve_color": curve_color,
                    "curve_opacity": style.get("curve_opacity", 1.0),
                    "curve_width": style.get("curve_width", 1.8),
                    "dash": style.get("dash", "solid"),
                    "line_shape": "hv" if is_discrete_cdf else "linear",
                }
            )

        if ResultsMainRenderService._should_show_theory_cdf(dialog):
            theory_group = ResultsMainRenderService._build_theory_cdf_group(dialog)
            if theory_group is not None:
                groups.append(theory_group)

        ResultsMainRenderService._apply_chart_margins(dialog, dialog.MARGIN_R)
        res = DriskChartFactory.build_cdf(
            data_groups=groups,
            x_range=[dialog.view_min, dialog.view_max],
            x_dtick=float(dialog.x_dtick),
            margins=(dialog.MARGIN_L, dialog.MARGIN_R, dialog.MARGIN_T, dialog.MARGIN_B),
            forced_mag=dialog.manual_mag,
            label_overrides=getattr(dialog, "_label_settings_config", None),
            axis_numeric_flags=getattr(dialog, "_label_axis_numeric", None),
        )
        dialog._load_plotly_html(res["plot_json"], res["js_mode"], res["js_logic"])

    @staticmethod
    def render_channel_auto(dialog):
        """[通道] 自动判定模式。基于数据离散特征智能回退到离散柱状图或直方图。"""
        if getattr(dialog, "is_discrete_view", False):
            ResultsMainRenderService.render_channel_discrete(dialog)
            return
        ResultsMainRenderService.render_channel_hist(dialog, "histogram")

    @staticmethod
    def render_pdf_from_cache(dialog):
        """
        [子过程] 基于后台线程池已返回的网格数据组装生成 Plotly PDF 视图。
        处理多线叠加时的颜色/线型匹配，并自适应计算 Y 轴高度范围。
        """
        if dialog._pdf_x_grid is None or not dialog._pdf_y_map:
            return

        keys = list(getattr(dialog, "display_keys", [dialog.current_key]))
        display_name_map = ResultsMainRenderService._build_series_display_name_map(dialog, keys)
        x_min = float(getattr(dialog, "x_range_min", dialog.abs_min))
        x_max = float(getattr(dialog, "x_range_max", dialog.abs_max))
        pdf_points = len(dialog._pdf_x_grid)
        sig = (tuple(keys), x_min, x_max, pdf_points)
        
        # 缓存防御：若缓存签名过期（如图表被外部缩放），重新下发任务
        if dialog._pdf_keys_sig != sig:
            dialog._pdf_keys_sig = sig
            dialog.job_controller.request_pdf_kde(
                keys,
                getattr(dialog, "series_map", {}),
                x_min,
                x_max,
                pdf_points,
            )
            return

        fig = go.Figure()
        y_max = 0.0
        for key in keys:
            y_vals = dialog._pdf_y_map.get(key)
            if y_vals is None:
                continue

            # 使用安全截断算法裁剪曲线
            bounds = ResultsMainRenderService._extract_series_bounds(dialog, key)
            x_plot, y_plot = ResultsMainRenderService._build_bounded_kde_curve(dialog._pdf_x_grid, y_vals, bounds)
            if x_plot.size == 0 or y_plot.size == 0:
                continue
                
            if len(y_plot) > 0:
                y_max = max(y_max, float(np.max(y_plot)))

            style = getattr(dialog, "_series_style", {}).get(key, {})
            color = style.get("color", "#0050b3")
            dash = style.get("dash", "solid")
            
            # 追加 Plotly 连线 Trace
            fig.add_trace(
                go.Scatter(
                    x=x_plot,
                    y=y_plot,
                    mode="lines",
                    name=ResultsMainRenderService._display_series_name(display_name_map, key),
                    line=dict(color=color, width=1.5, dash=dash),
                    hoverinfo="skip",
                )
            )

        if ResultsMainRenderService._should_show_theory_pdf(dialog):
            theory_trace = ResultsMainRenderService._build_theory_pdf_trace(dialog, x_grid=dialog._pdf_x_grid)
            if theory_trace is not None:
                try:
                    y_arr = np.asarray(theory_trace.y, dtype=float).ravel()
                    if y_arr.size > 0:
                        y_max = max(y_max, float(np.max(y_arr)))
                except Exception:
                    pass
                fig.add_trace(theory_trace)

        # 智能推导 Y 轴顶部对齐高度
        if y_max <= 1e-12:
            y_max = 1.0
        y_dtick = DriskMath.calc_smart_step(y_max)
        if y_dtick <= 0:
            y_dtick = y_max / 5.0

        padded_y_max = y_max + 0.01 * y_dtick
        y_axis_max = math.ceil(padded_y_max / y_dtick) * y_dtick
        axis_style = build_plotly_axis_style(
            x_title=DriskChartFactory.VALUE_AXIS_TITLE,
            x_unit=DriskChartFactory.VALUE_AXIS_UNIT,
            y_title="密度",
            x_range=[dialog.x_range_min, dialog.x_range_max],
            y_range=[0, y_axis_max],
            x_dtick=dialog.x_dtick,
            y_dtick=y_dtick,
            fixedrange=True,
            forced_mag=dialog.manual_mag,
            label_overrides=getattr(dialog, "_label_settings_config", None),
            x_axis_numeric=bool((getattr(dialog, "_label_axis_numeric", {}) or {}).get("x", True)),
            y_axis_numeric=bool((getattr(dialog, "_label_axis_numeric", {}) or {}).get("y", True)),
        )
        
        fig.update_layout(
            autosize=True,
            template="plotly_white",
            margin=dict(l=dialog.MARGIN_L, r=dialog.MARGIN_R, t=dialog.MARGIN_T, b=dialog.MARGIN_B),
            showlegend=False,
            xaxis=axis_style.get("xaxis", {}),
            yaxis=axis_style.get("yaxis", {}),
        )
        plot_json = json.dumps(fig.to_dict(), cls=plotly.utils.PlotlyJSONEncoder)
        dialog._load_plotly_html(plot_json, "pdf", "")

    @staticmethod
    def log_kde_overlay_error():
        """
        [诊断日志] 捕获并落盘底层 C 库在生成 KDE 或 Bootstrap 时的极端异常，
        供开发者排查问题，输出至用户桌面。
        """
        desktop_path = os.path.join(os.path.expanduser("~"), "Desktop", "drisk_kde_crash_log.txt")
        with open(desktop_path, "a", encoding="utf-8") as file_obj:
            file_obj.write(f"KDE Overlay Error:\n{traceback.format_exc()}\n")
