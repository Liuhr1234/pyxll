# drisk_charting.py
# -*- coding: utf-8 -*-
"""
本模块提供结果分析链路中的图形构建工厂 DriskChartFactory。

模块定位：
1. 负责将上游整理后的数值序列、样式配置、坐标轴参数转换为 Plotly 可直接消费的图形描述；
2. 覆盖结果分析弹窗中的主要图形类型，包括直方图、理论 PDF、离散概率柱状图、CDF、箱线/小提琴/趋势图以及散点图；
3. 在部分特殊图形中，同时负责补充前端渲染所需的附加 JS 逻辑，例如离散柱宽的像素级校正；
4. 仅负责“图形描述构建”，不直接操作 Qt 控件，也不维护界面状态。

设计原则：
1. 输入尽量宽容：允许 list、numpy、pandas 等多种数据容器；
2. 输出尽量稳定：所有公开构建函数统一返回 plot_json / js_mode / js_logic 结构；
3. 样式与数学逻辑分离：坐标轴样式复用 ui_shared 中的公共工具，步长/取整复用 DriskMath；
4. 面向交接：本文件中的方法按照“基础工具 → 分布图 → 箱型图 → 散点图”的顺序组织。
"""
from __future__ import annotations

import re
import json
import math
from sklearn.cluster import MiniBatchKMeans
from typing import Any, Dict, List, Optional, Sequence, Tuple

import numpy as np
import plotly.graph_objects as go
import plotly.utils

from ui_shared import build_plotly_axis_style, DriskMath, DRISK_COLOR_CYCLE, DRISK_COLOR_CYCLE_THEO

# ============================================================
# 1. 模块级常量
# ============================================================
SCATTER_SYMBOLS = [
    "circle",        
    "square",        
    "diamond",       
    "triangle-up",   
    "pentagon",      
    "hexagram",      
    "star",          
    "cross",         
    "x",             
    "triangle-down"  
]
ERROR_MARKER = "#ERROR!"


# ============================================================
# 2. 基础数据转换工具
# ============================================================

def _to_float_array(seq: Any) -> np.ndarray:
    """
    将任意常见输入容器清洗为一维 float 数组。

    处理规则：
    1. 兼容 pandas / numpy / list 等常见输入；
    2. 自动跳过 None、空字符串、错误标记以及不可转为数值的内容；
    3. 仅保留有限数值（finite numbers），为后续绘图逻辑提供统一输入。

    说明：
    该函数是本模块最基础的防御性入口，几乎所有图形构建函数都会先通过它做一次标准化。
    """
    if seq is None:
        return np.array([], dtype=float)
    try:
        if hasattr(seq, "to_numpy"):
            raw = seq.to_numpy()
        elif hasattr(seq, "values"):
            raw = seq.values
        elif hasattr(seq, "tolist"):
            raw = seq.tolist()
        else:
            raw = list(seq)
    except Exception:
        raw = [seq]

    out: List[float] = []
    for v in raw:
        if v is None:
            continue
        if isinstance(v, str):
            s = v.strip()
            if (not s) or (s == ERROR_MARKER):
                continue
            try:
                fv = float(s)
            except Exception:
                continue
        else:
            try:
                fv = float(v)
            except Exception:
                continue
        if np.isfinite(fv):
            out.append(float(fv))

    if not out:
        return np.array([], dtype=float)
    return np.asarray(out, dtype=float)


# ============================================================
# 3. 图形构建工厂
# ============================================================

class DriskChartFactory:
    """
    结果分析图形工厂。

    约定：
    1. 所有方法均为静态方法，不依赖实例状态；
    2. VALUE_AXIS_TITLE / VALUE_AXIS_UNIT 作为公共值轴标题与单位入口，由上层调用方在运行时覆盖；
    3. 每个 build_* 方法返回的都是“可交给前端宿主直接渲染”的结构化结果。
    """
    VALUE_AXIS_TITLE = ""
    VALUE_AXIS_UNIT = ""
    
    # ------------------------------------------------------------
    # ------------------------------------------------------------
    
    @staticmethod
    def _hex_to_rgba(hex_color: str, a: float) -> str:
        """
        将十六进制颜色转换为 rgba 字符串。

        处理说明：
        1. 支持 6 位和 8 位十六进制输入；
        2. transparent 直接映射为全透明；
        3. 若输入异常，则回退为默认淡紫色，避免图形构建因颜色格式问题中断。
        """
        if not hex_color or hex_color == "transparent":
            return "rgba(0,0,0,0)"
            
        h = str(hex_color).lstrip("#")
        
        if len(h) in (6, 8):
            try:
                if len(h) == 8:
                    r = int(h[2:4], 16)
                    g = int(h[4:6], 16)
                    b = int(h[6:8], 16)
                else:
                    r = int(h[0:2], 16)
                    g = int(h[2:4], 16)
                    b = int(h[4:6], 16)
                return f"rgba({r},{g},{b},{a})"
            except ValueError:
                pass 
        return f"rgba(216,191,216,{a})"

    @staticmethod
    def _compute_cdf_points(arr: np.ndarray) -> Tuple[np.ndarray, np.ndarray]:
        """
        为累计分布曲线构建点集。

        策略：
        1. 小样本直接保留逐点 CDF，确保精确性；
        2. 大样本压缩为 1001 个分位点，兼顾前端渲染性能；
        3. 若输入为空，则返回空数组。
        """
        arr = np.asarray(arr, dtype=float)
        arr = arr[np.isfinite(arr)]
        n = arr.size
        if n == 0:
            return np.array([]), np.array([])

        arr.sort()

        if n >= 1001:
            ys = np.linspace(0.0, 1.0, 1001)
            if hasattr(np, "quantile"):
                try:
                    xs = np.quantile(arr, ys, method="linear")
                except TypeError:
                    xs = np.quantile(arr, ys)
            else:
                xs = np.percentile(arr, ys * 100, interpolation="linear")
            return xs, ys
        else:
            xs = arr
            ys = np.arange(1, n + 1, dtype=float) / n
            return xs, ys

    @staticmethod
    def _inject_overlay_cdf(
        fig: go.Figure,
        layout_dict: dict,
        x_data,
        y_data,
        name: str = "累积概率",
        color: str = "#000000",
        add_trace: bool = True,
        line_shape: str = "hv",
        width: float = 1.8,
        dash: str = "solid",
        opacity: float = 1.0,
    ) -> None:
        """
        向现有图形注入右侧累积概率轴及对应的 CDF 叠加曲线。

        典型用途：
        1. 在直方图上叠加累计概率；
        2. 在离散概率柱图上叠加累计概率；
        3. 在仅需注册右轴但暂不立即加线时，通过 add_trace=False 只写入布局。
        """
        # 在需要真实叠加曲线时，先把颜色透明度转换为 rgba。
        if add_trace:
            if color.startswith("#") and opacity < 1.0:
                color = DriskChartFactory._hex_to_rgba(color, opacity)
                
            fig.add_trace(
                go.Scatter(
                    x=x_data,
                    y=y_data,
                    name=name,
                    mode="lines",
                    yaxis="y2",
                    line=dict(color=color, width=width, dash=dash, shape=line_shape),
                    hoverinfo="skip",
                )
            )
        axis_font = dict(family="Arial", size=12.5, color="#333333")

        if "yaxis2" not in layout_dict:
            layout_dict["yaxis2"] = {
                "title": "累积概率",
                "overlaying": "y",
                "side": "right",
                "range": [0, 1.05],
                "tickformat": ".0%",
                "tickfont": axis_font, 
                "showgrid": False,
                "fixedrange": True,
                "ticks": "outside",
                "ticklen": 4, 
                "tickcolor": "#333",
                "showline": True,
                "linecolor": "#d9d9d9",
            }

    @staticmethod
    def _inject_chart_title_override(layout_dict: dict, label_overrides: Optional[Dict[str, Any]]) -> None:
        """
        将外部传入的标题重写配置注入到当前布局字典中。

        说明：
        - 若未提供标题配置，则保持原布局不变；
        - 若显式传入空标题，则删除已有标题；
        - 仅负责写入布局，不做额外样式推导。
        """
        if not isinstance(layout_dict, dict):
            return
        if not isinstance(label_overrides, dict):
            return
        chart_cfg = label_overrides.get("chart_title", {})
        if not isinstance(chart_cfg, dict):
            return
        raw_text = chart_cfg.get("text_override", None)
        if raw_text is None:
            return

        title_text = str(raw_text).strip()
        # 显式传入空标题时，删除原布局中的标题设置。
        if not title_text:
            layout_dict.pop("title", None)
            return

        chart_font_family = str(chart_cfg.get("font_family") or "Arial, 'Microsoft YaHei', sans-serif")
        try:
            chart_font_size = float(chart_cfg.get("font_size")) if chart_cfg.get("font_size") is not None else 14.0
        except Exception:
            chart_font_size = 14.0

        layout_dict["title"] = dict(
            text=title_text,
            x=0.5,
            xref="paper",
            xanchor="center",
            font=dict(size=chart_font_size, family=chart_font_family, color="#333333"),
        )

    @staticmethod
    def get_pdf_line_trace(x, y, color, name: str = "概率密度"):
        """
        构建一条标准概率密度折线 trace。

        该方法主要用于直方图场景下叠加理论 PDF 或其他外部生成的密度曲线。
        """
        return go.Scatter(
            x=_to_float_array(x),
            y=_to_float_array(y),
            mode="lines",
            name=name,
            line=dict(color=color, width=1.5),
            hoverinfo="skip",
        )

    # ------------------------------------------------------------
    # ------------------------------------------------------------

    @staticmethod
    def build_simulation_histogram(
        series_map: Dict[str, Any],
        display_keys: List[str],
        fill_keys: List[str],
        style_map: Dict[str, Dict[str, Any]],
        view_mode: str,
        x_range: Tuple[float, float],
        x_dtick: float,
        margins: Tuple[int, int, int, int] = (80, 40, 30, 50),
        show_overlay_cdf: bool = False,
        pdf_traces: Optional[List[Any]] = None,
        forced_mag: Optional[int] = None,
        display_name_map: Optional[Dict[str, str]] = None,
        label_overrides: Optional[Dict[str, Any]] = None,
        axis_numeric_flags: Optional[Dict[str, bool]] = None,
    ) -> dict:
        """
        构建模拟结果直方图。

        支持能力：
        1. 普通相对频数直方图与概率密度直方图；
        2. 多序列叠加显示；
        3. 叠加外部 PDF 曲线；
        4. 叠加右轴 CDF 曲线；
        5. 接入统一坐标轴样式与标题覆盖配置。
        """
        view_mode_norm = (view_mode or "").strip().lower()
        use_density = view_mode_norm in ("histogram", "density")
        histnorm = "probability density" if use_density else "percent"
        y_title = "密度" if use_density else "相对频数 (%)"
        # 以样本量估算分箱数，保证大样本时仍控制在合理上限。
        base_size = 1000
        for k in display_keys:
            arr = _to_float_array(series_map.get(k))
            if arr.size > 0:
                base_size = max(base_size, arr.size) 
        
        n_bins = min(max(int(10 * math.log10(max(1, base_size))), 10), 200)

        fig = go.Figure()

        # 先逐组构建直方图，并同步估算 y 轴最高峰值。
        max_peak = 0.0
        for k in display_keys:
            arr = _to_float_array(series_map.get(k))
            if arr.size == 0: continue
            
            arr_min, arr_max = float(np.min(arr)), float(np.max(arr))
            if abs(arr_max - arr_min) < 1e-9:
                arr_min -= 0.5
                arr_max += 0.5
            bin_width = (arr_max - arr_min) / n_bins
            hy, _ = np.histogram(arr, bins=n_bins, range=(arr_min, arr_max), density=use_density)
            if not use_density: hy = (hy / arr.size) * 100.0
            max_peak = max(max_peak, float(np.max(hy)))
            
            # 
            st = style_map.get(k, {})
            base_color = st.get("color", DRISK_COLOR_CYCLE[1])
            
            fill_color = st.get("fill_color", base_color)
            fill_opacity = st.get("fill_opacity", 0.75)     # 填充透明度限制在 [0, 1] 区间。
            outline_color = st.get("outline_color", "rgba(255,255,255,0.5)")
            outline_width = st.get("outline_width", 0.8)

            if k in fill_keys and fill_color != "transparent":
                fill_c = DriskChartFactory._hex_to_rgba(fill_color, fill_opacity)                
                if outline_color.startswith("#"):
                    out_c = DriskChartFactory._hex_to_rgba(outline_color, 1.0)
                else:
                    out_c = outline_color
                    
                marker = dict(
                    color=fill_c, 
                    line=dict(width=outline_width, color=out_c)
                )
            else:
                if outline_color == "rgba(255,255,255,0.5)":
                    out_c = base_color 
                elif outline_color.startswith("#"):
                    out_c = DriskChartFactory._hex_to_rgba(outline_color, 1.0)
                else:
                    out_c = outline_color

                marker = dict(
                    color="rgba(0,0,0,0)", 
                    line=dict(width=st.get("outline_width", 1.0), color=out_c)
                )
            display_name = str((display_name_map or {}).get(k, (display_name_map or {}).get(str(k), k)))
            fig.add_trace(go.Histogram(
                x=arr, name=display_name, histnorm=histnorm, autobinx=False,
                xbins=dict(start=arr_min, end=arr_max, size=bin_width),
                marker=marker
            ))
        # 若叠加 PDF 曲线，则将其峰值一并纳入 y 轴范围计算。
        pdf_peak = 0.0
        if pdf_traces:
            for trace in pdf_traces:
                fig.add_trace(trace)
                y_vals = _to_float_array(trace.y)
                if y_vals.size: pdf_peak = max(pdf_peak, float(np.max(y_vals)))
        raw_y_max = max(max_peak, pdf_peak)
        y_dtick = DriskMath.calc_smart_step(raw_y_max)
        
        padded_y_max = raw_y_max + 0.01 * y_dtick
        y_axis_max = DriskMath.ceil_to_tick(padded_y_max, y_dtick)

        # 
        axis_flags = axis_numeric_flags if isinstance(axis_numeric_flags, dict) else {}
        axis_style = build_plotly_axis_style(
            x_title=DriskChartFactory.VALUE_AXIS_TITLE,
            x_unit=DriskChartFactory.VALUE_AXIS_UNIT,    
            y_title=y_title, 
            x_range=x_range,
            y_range=[0, y_axis_max], x_dtick=x_dtick, y_dtick=y_dtick,
            fixedrange=True,
            forced_mag=forced_mag,
            label_overrides=label_overrides,
            x_axis_numeric=bool(axis_flags.get("x", True)),
            y_axis_numeric=bool(axis_flags.get("y", True)),
        )
        has_x_title = bool(axis_style["xaxis"].get("title"))
        b_margin = margins[3] if has_x_title else 25

        layout_dict = {
            "template": "plotly_white",
            "margin": dict(l=margins[0], r=margins[1], t=margins[2], b=b_margin),
            "hovermode": "x",
            "xaxis": axis_style["xaxis"],
            "yaxis": axis_style["yaxis"],
            "showlegend": False, "barmode": "overlay"
        }
        DriskChartFactory._inject_chart_title_override(layout_dict, label_overrides)
        if show_overlay_cdf:
            for k in display_keys:
                arr = _to_float_array(series_map.get(k))
                if arr.size == 0: continue
                cx, cy = DriskChartFactory._compute_cdf_points(arr)
                
                st = style_map.get(k, {})
                
                cdf_color = st.get("curve_color", st.get("cdf_color", st.get("color", "#9370DB")))
                curve_width = st.get("curve_width", 1.8)
                curve_dash = st.get("dash", "solid")
                curve_opacity = st.get("curve_opacity", 1.0)
                
                DriskChartFactory._inject_overlay_cdf(
                    fig, layout_dict, cx, cy, 
                    color=cdf_color, width=curve_width, dash=curve_dash, opacity=curve_opacity
                )

        fig.update_layout(layout_dict)
        return {"plot_json": json.dumps(fig.to_dict(), cls=plotly.utils.PlotlyJSONEncoder), "js_mode": "histogram", "js_logic": ""}

    @staticmethod
    def build_theory_pdf(
        x_data: Sequence[float],
        y_data: Sequence[float],
        x_range: Tuple[float, float],
        x_dtick: float,
        y_max: Optional[float] = None,
        note_annotation: str = "",
        margins: Tuple[int, int, int, int] = (80, 40, 30, 50),
        show_overlay_cdf: bool = False,
        cdf_data: Optional[Sequence[float]] = None,
        forced_mag: Optional[int] = None,
        label_overrides: Optional[Dict[str, Any]] = None,
        axis_numeric_flags: Optional[Dict[str, bool]] = None,
    ) -> dict:
        """
        构建理论分布 PDF 图。

        特点：
        1. 使用面积填充强调理论概率密度形状；
        2. 可在右轴叠加理论 CDF；
        3. 支持右上角注释文本，用于展示警告或补充说明。
        """

        x_arr = _to_float_array(x_data)
        y_arr = _to_float_array(y_data)
        
        line_color = "#000000"
        fill_color = DriskChartFactory._hex_to_rgba(DRISK_COLOR_CYCLE_THEO[0], 0.75)
        traces = [
            go.Scatter(
                x=x_arr,
                y=y_arr,
                fill="tozeroy",
                fillcolor=fill_color,
                mode="lines",
                line=dict(color=line_color, width=1.5),
                hoverinfo="skip",
            )
        ]
        fig = go.Figure(data=traces)

        if y_max is None:
            y_max = float(np.max(y_arr)) if y_arr.size else 1.0
        if y_max <= 1e-12:
            y_max = 1.0

        y_dtick = DriskMath.calc_smart_step(y_max)
        
        padded_y_max = y_max + 0.01 * y_dtick
        y_axis_max = DriskMath.ceil_to_tick(padded_y_max, y_dtick)

        axis_flags = axis_numeric_flags if isinstance(axis_numeric_flags, dict) else {}
        axis_style = build_plotly_axis_style(
            x_range=x_range,
            y_range=[0, y_axis_max],
            x_dtick=x_dtick,
            y_dtick=y_dtick,
            x_title=DriskChartFactory.VALUE_AXIS_TITLE,
            x_unit=DriskChartFactory.VALUE_AXIS_UNIT,    
            y_title="密度",
            fixedrange=True,
            forced_mag=forced_mag,
            label_overrides=label_overrides,
            x_axis_numeric=bool(axis_flags.get("x", True)),
            y_axis_numeric=bool(axis_flags.get("y", True)),
        )

        has_x_title = bool(axis_style["xaxis"].get("title"))
        b_margin = margins[3] if has_x_title else 25

        ann = []
        if note_annotation:
            ann = [
                dict(
                    text=note_annotation,
                    x=0.99,
                    xref="paper",
                    xanchor="right",
                    y=0.99,
                    yref="paper",
                    yanchor="top",
                    showarrow=False,
                    font=dict(size=11, color="#cf1322"),
                    bgcolor="rgba(255,255,255,0.7)",
                )
            ]

        layout_dict = {
            "autosize": True,
            "template": "plotly_white",
            "margin": dict(l=margins[0], r=margins[1], t=margins[2], b=b_margin),
            "xaxis": axis_style["xaxis"],
            "yaxis": axis_style["yaxis"],
            "showlegend": False,
            "barmode": "overlay",
            "paper_bgcolor": "rgba(0,0,0,0)",
            "plot_bgcolor": "rgba(0,0,0,0)",
            "annotations": ann,
        }
        DriskChartFactory._inject_chart_title_override(layout_dict, label_overrides)

        if show_overlay_cdf and cdf_data is not None:
            cdf_arr = _to_float_array(cdf_data)
            DriskChartFactory._inject_overlay_cdf(
                fig, layout_dict, x_arr, cdf_arr, color=DRISK_COLOR_CYCLE_THEO[0], line_shape="linear"
            )

        fig.update_layout(layout_dict)

        return {
            "plot_json": json.dumps(fig.to_dict(), cls=plotly.utils.PlotlyJSONEncoder),
            "js_mode": "pdf",
            "js_logic": "",
        }

    # ------------------------------------------------------------
    # ------------------------------------------------------------

    @staticmethod
    def _get_discrete_pixel_perfect_js() -> str:
        """
        返回离散柱图的前端像素修正脚本。

        背景：
        Plotly 在不同缩放层级下，使用数据空间宽度定义柱宽时，视觉宽度并不总是稳定。
        该脚本会在前端读取当前坐标轴像素比例，并在 relayout / afterplot 后自动重算柱宽，
        以尽量保持离散柱在不同视口中的可读性。
        """
        return r"""
        (function(){
          // 计算当前横轴“每单位数据对应多少像素”。
          function __compute_px_per_unit(gd){
            try{
              var xa = gd._fullLayout.xaxis;
              var sz = gd._fullLayout._size;
              if(!xa || !sz || !sz.w || sz.w <= 0) return null;
              var pxSpan = sz.w;
              var r = xa.range;
              if(!r || r.length < 2) return null;
              var span = Math.abs(r[1] - r[0]);
              if(!isFinite(span) || span === 0) return null;
              return pxSpan / span;
            }catch(e){ return null; }
          }

          function __get_gd(){
            return document.getElementById('chart_div') || document.querySelector('.js-plotly-plot');
          }

          // 从柱图支持点中提取代表性间距信息。
          function __collect_support_info(gd, barIdx){
            try{
              var xa = gd._fullLayout && gd._fullLayout.xaxis;
              var r = xa && xa.range;
              var hasRange = !!(r && r.length >= 2 && isFinite(r[0]) && isFinite(r[1]));
              var left = null;
              var right = null;
              if(hasRange){
                left = Math.min(r[0], r[1]);
                right = Math.max(r[0], r[1]);
              }

              var xs = [];
              for(var i=0;i<barIdx.length;i++){
                var trace = gd.data[barIdx[i]];
                var tx = trace && trace.x;
                if(!tx || !tx.length) continue;
                for(var j=0;j<tx.length;j++){
                  var num = Number(tx[j]);
                  if(!isFinite(num)) continue;
                  if(hasRange && (num < left || num > right)) continue;
                  xs.push(num);
                }
              }
              if(xs.length === 0) return {count: 0, minGapData: null, minGapPx: null};

              xs.sort(function(a, b){ return a - b; });
              var uniq = [];
              var eps = 1e-10;
              for(var k=0;k<xs.length;k++){
                if(uniq.length === 0 || Math.abs(xs[k] - uniq[uniq.length - 1]) > eps){
                  uniq.push(xs[k]);
                }
              }

              var minGapData = null;
              var gapDataList = [];
              for(var m=1;m<uniq.length;m++){
                var gap = uniq[m] - uniq[m - 1];
                if(!isFinite(gap) || gap <= eps) continue;
                gapDataList.push(gap);
                if(minGapData === null || gap < minGapData) minGapData = gap;
              }

              var repGapData = null;
              if(gapDataList.length > 0){
                gapDataList.sort(function(a, b){ return a - b; });
                repGapData = gapDataList[Math.floor(gapDataList.length / 2)];
              }

              var minGapPx = null;
              var gapPxList = [];
              if(xa && typeof xa.l2p === 'function'){
                for(var n=1;n<uniq.length;n++){
                  var pxA = Number(xa.l2p(uniq[n - 1]));
                  var pxB = Number(xa.l2p(uniq[n]));
                  if(!isFinite(pxA) || !isFinite(pxB)) continue;
                  var pxGap = Math.abs(pxB - pxA);
                  if(pxGap <= eps) continue;
                  gapPxList.push(pxGap);
                  if(minGapPx === null || pxGap < minGapPx) minGapPx = pxGap;
                }
              }

              var repGapPx = null;
              if(gapPxList.length > 0){
                gapPxList.sort(function(a, b){ return a - b; });
                repGapPx = gapPxList[Math.floor(gapPxList.length / 2)];
              }

              return {count: uniq.length, minGapData: minGapData, minGapPx: minGapPx, repGapData: repGapData, repGapPx: repGapPx};
            }catch(e){
              return {count: 0, minGapData: null, minGapPx: null, repGapData: null, repGapPx: null};
            }
          }

          // 
          function __compute_discrete_px_target(gd, barIdx, pxPerUnit){
            try{
              // Hard upper cap for bar pixel width.
              var maxPxTarget = 8.0;
              // Keep a visible lower bound so bars do not collapse to sub-pixel width.
              var minPxTarget = 3.0;
              var gapLockPx = 1.2;

              var info = __collect_support_info(gd, barIdx);
              // 无法提取支持点信息时，直接返回，不做宽度修正。
              if(!info || !info.count) return null;

              var spacingPx = null;
              if(info.repGapPx !== null && isFinite(info.repGapPx) && info.repGapPx > 0){
                spacingPx = info.repGapPx;
              }else if(info.repGapData !== null && isFinite(info.repGapData) && info.repGapData > 0){
                spacingPx = info.repGapData * pxPerUnit;
              }else{
                var sz = gd._fullLayout && gd._fullLayout._size;
                if(sz && isFinite(sz.w) && info.count > 0){
                  spacingPx = sz.w / info.count;
                }
              }

              var sz2 = gd._fullLayout && gd._fullLayout._size;
              if(sz2 && isFinite(sz2.w) && info.count > 0 && isFinite(spacingPx) && spacingPx > 0){
                // Floor spacing by support-count pitch to avoid min-gap outliers collapsing bars.
                var spacingFloorPx = (sz2.w / info.count) * 0.35;
                if(isFinite(spacingFloorPx) && spacingFloorPx > 0){
                  spacingPx = Math.max(spacingPx, spacingFloorPx);
                }
              }

              if(spacingPx !== null && isFinite(spacingPx) && spacingPx > 0){
                var stageSwitchPitchPx = maxPxTarget + gapLockPx;
                if(spacingPx >= stageSwitchPitchPx){
                  return maxPxTarget;
                }
                var stage2WidthPx = spacingPx - gapLockPx;
                if (stage2WidthPx <= 0) {
                    return Math.max(minPxTarget, spacingPx * 0.65);
                }
                return Math.max(minPxTarget, Math.min(stage2WidthPx, maxPxTarget));
              }
              return null;
            }catch(e){
              return null;
            }
          }

          function logic(){
            try{
              var gd = __get_gd();
              if(!gd || !gd._fullLayout || !gd.data) return;

              if(gd.__drisk_running) return;
              gd.__drisk_running = true;

              var barIdx = [];
              for(var i=0;i<gd.data.length;i++){
                if(gd.data[i].type === 'bar') barIdx.push(i);
              }
              if(barIdx.length === 0) {
                  gd.__drisk_running = false;
                  return;
              }

              var pxPerUnit = __compute_px_per_unit(gd);
              if(!pxPerUnit || pxPerUnit <= 0) {
                  gd.__drisk_running = false;
                  return;
              }

              var pxTarget = __compute_discrete_px_target(gd, barIdx, pxPerUnit);
              if(!pxTarget || !isFinite(pxTarget) || pxTarget <= 0) {
                  gd.__drisk_running = false; // 未能算出有效的目标柱宽。
                  return; 
              }
              
              var w = pxTarget / pxPerUnit;
              if (isNaN(w) || !isFinite(w) || w <= 0) {
                  gd.__drisk_running = false;
                  return;
              }

              // 最终宽度仍需受上限约束（放宽以配合 minPxTarget）。
              w = Math.min(w, 5.0);

              var currentW = gd.data[barIdx[0]].width;
              if (Array.isArray(currentW)) {
                  currentW = currentW[0];
              }
              
              if (typeof currentW === 'number' && !isNaN(currentW)) {
                  var diff = Math.abs(currentW - w);
                  if (diff < (w * 0.05)) {
                      gd.__drisk_running = false;
                      return; 
                  }
              }

              // Update trace data first, then trigger Plotly.react.
              for (var k = 0; k < barIdx.length; k++) {
                  gd.data[barIdx[k]].width = w;
              }
              
              Plotly.react(gd, gd.data, gd.layout).then(function(){
                  setTimeout(function(){ gd.__drisk_running = false; }, 50);
              }).catch(function(){
                  gd.__drisk_running = false;
              });

              // Mirror width update to cached window.figure if present.
              try {
                  if (typeof window.figure !== 'undefined' && window.figure && window.figure.data) {
                      for(var k=0; k<barIdx.length; k++){
                          if(window.figure.data[barIdx[k]]){
                              window.figure.data[barIdx[k]].width = w;
                          }
                      }
                  }
              } catch(e) {}
              
            }catch(e){
               var gd = __get_gd();
               if(gd) gd.__drisk_running = false;
            }
          }

          // Expose a post-init hook for host-side scheduling.
          window.__drisk_post_init = function() {
              setTimeout(logic, 150);
          };

          // 在 relayout / afterplot 事件后重新触发柱宽修正。
          try{
            var gd = __get_gd();
            if(gd && gd.on && !gd.__drisk_hook_bound){
              gd.__drisk_hook_bound = true;
              var schedule = function(){
                if(gd.__drisk_t) clearTimeout(gd.__drisk_t);
                gd.__drisk_t = setTimeout(logic, 100);
              };
              gd.on('plotly_relayout', schedule);
              gd.on('plotly_afterplot', schedule);
            }
          }catch(e){}

          // 
          setTimeout(logic, 150);
        })();
        """

    @staticmethod
    def build_discrete_bar(
        data_groups: List[Dict[str, Any]],
        x_range: Tuple[float, float],
        x_dtick: float,
        y_max: Optional[float] = None,
        note_annotation: str = "",
        margins: Tuple[int, int, int, int] = (80, 40, 30, 50),
        show_overlay_cdf: bool = False,
        cdf_data: Optional[Sequence[float]] = None,
        forced_mag: Optional[int] = None,
        label_overrides: Optional[Dict[str, Any]] = None,
        axis_numeric_flags: Optional[Dict[str, bool]] = None,
    ) -> dict:
        """
        构建离散概率柱状图，并按需叠加累计概率曲线。

        该方法主要服务于离散型分布或离散模拟结果的展示场景。
        """

        # 当调用方未提供有效横轴范围时，先基于全部组数据推导一个可靠范围。
        all_x_raw = []
        for group in data_groups:
            all_x_raw.extend(group.get("x", []))
        all_x_arr = _to_float_array(all_x_raw)
        
        if all_x_arr.size > 0:
            d_min, d_max = float(np.min(all_x_arr)), float(np.max(all_x_arr))
            # 对缺失或非法范围参数，用数据边界回填。
            if x_range is None or len(x_range) < 2 or not np.isfinite(x_range[0]) or not np.isfinite(x_range[1]):
                if d_min == d_max:
                    # 当最小值与最大值相同，主动扩展一个退化区间，避免柱体不可见。
                    x_range = [d_min - 5.0, d_max + 5.0]
                else:
                    span_val = d_max - d_min
                    x_range = [d_min - span_val * 0.05, d_max + span_val * 0.05]
            else:
                # 对接近零宽度的显式范围做兜底修正。
                if abs(float(x_range[1]) - float(x_range[0])) < 1e-9:
                    x_range = [float(x_range[0]) - 5.0, float(x_range[1]) + 5.0]

        fig = go.Figure()
        calc_y_max = 0.0

        def _calc_bar_width_data(xs_arr: np.ndarray) -> Optional[float]:
            """
            基于当前可视范围内的支持点分布，估算适合的柱宽（数据空间单位）。

            目标：
            1. 保证缩放后柱子仍可见；
            2. 避免横轴跨度很大时柱子过宽；
            3. 为前端像素级修正提供一个合理初值。
            """
            try:
                if xs_arr.size == 0:
                    return None
                xs_num = xs_arr.astype(float, copy=False)
                xs_num = xs_num[np.isfinite(xs_num)]
                if xs_num.size == 0:
                    return None

                has_range = (
                    x_range is not None
                    and len(x_range) >= 2
                    and np.isfinite(float(x_range[0]))
                    and np.isfinite(float(x_range[1]))
                )
                
                span = 1.0
                if has_range:
                    left = min(float(x_range[0]), float(x_range[1]))
                    right = max(float(x_range[0]), float(x_range[1]))
                    span = max(1e-9, right - left)
                    xs_num = xs_num[(xs_num >= left) & (xs_num <= right)]
                    if xs_num.size == 0:
                        return None

                uniq = np.unique(np.sort(xs_num))
                if uniq.size <= 0:
                    return None

                diffs = np.diff(uniq)
                diffs = diffs[np.isfinite(diffs) & (diffs > 1e-12)]
                if diffs.size > 0:
                    min_gap = float(np.min(diffs))
                    rep_gap = float(np.median(diffs))
                else:
                    min_gap = float(span / max(1, uniq.size))
                    rep_gap = min_gap

                # 将代表性间距与基于支持点数量估算的平均间距共同作为柱宽依据。
                count_pitch_data = float(span / max(1, uniq.size))
                gap_basis = max(rep_gap, count_pitch_data * 0.35)
                safe_width = float(gap_basis * 0.75)
                if has_range:
                    # 缩放很远时，对柱宽设置视觉上限，避免柱体过宽。
                    visual_cap = float(span * 0.008) 
                    safe_width = min(safe_width, visual_cap)
                
                # 施加最小可见宽度下限，避免柱子收缩到不可辨识。
                if has_range:
                    min_visible_width = float(span * 0.002)
                else:
                    min_visible_width = float(gap_basis * 0.2)
                min_visible_width = max(0.05, min(2.0, min_visible_width))  # 最低宽度 0.05
                return max(min_visible_width, safe_width)
            except Exception:
                return None

        # Add bar traces and optional per-group CDF lines.
        for group in data_groups:
            xs = _to_float_array(group.get("x", []))
            ys_raw = _to_float_array(group.get("y", []))
            if ys_raw.size == 0: continue

            use_percent = np.sum(ys_raw) <= 2.0
            ys = ys_raw * 100.0 if use_percent else ys_raw
            
            calc_y_max = max(calc_y_max, float(np.max(ys)))            
            color = group.get("color", DRISK_COLOR_CYCLE[1])
            fill_color = group.get("fill_color", color)
            try:
                fill_opacity = float(group.get("fill_opacity", 1.0))
            except Exception:
                fill_opacity = 1.0
            fill_opacity = max(0.0, min(1.0, fill_opacity))

            if str(fill_color).lower() == "transparent":
                fill_rgba = "rgba(0,0,0,0)"
            elif isinstance(fill_color, str) and fill_color.startswith("#"):
                fill_rgba = DriskChartFactory._hex_to_rgba(fill_color, fill_opacity)
            else:
                fill_rgba = fill_color
            
            bar_width_data = _calc_bar_width_data(xs)
            bar_kwargs = dict(
                x=xs, y=ys, name=str(group.get("name", "")),
                marker=dict(
                    color=fill_rgba,
                    line=dict(width=0, color="rgba(0,0,0,0)"),
                ),
                hovertemplate="x=%{x}<br>p=%{y:.2f}%<extra></extra>"
            )
            if bar_width_data is not None:
                # 先使用数据空间宽度作为初值，随后由前端脚本进一步修正为更稳定的视觉柱宽。
                bar_kwargs["width"] = bar_width_data
            fig.add_trace(go.Bar(**bar_kwargs))

            if show_overlay_cdf:
                
                if cdf_data is not None:
                    y_cdf = _to_float_array(cdf_data)
                elif "cdf_data" in group:
                    y_cdf = _to_float_array(group["cdf_data"])
                else:
                    y_cdf = np.cumsum(ys_raw)
                    if np.max(y_cdf) > 1.1: y_cdf /= 100.0 
                    
                cdf_color = group.get("cdf_color", group.get("color", "#9370DB"))
                DriskChartFactory._inject_overlay_cdf(fig, {}, xs, y_cdf, color=cdf_color)

        
        final_y_peak = calc_y_max if calc_y_max > 0 else 1.0
        if y_max is not None:
            try:
                final_y_peak = max(final_y_peak, float(y_max))
            except Exception:
                pass
        
        y_dtick = DriskMath.calc_smart_step(final_y_peak)
        
        padded_y_max = final_y_peak + 0.01 * y_dtick
        y_axis_max = DriskMath.ceil_to_tick(padded_y_max, y_dtick)

        
        axis_flags = axis_numeric_flags if isinstance(axis_numeric_flags, dict) else {}
        axis_style = build_plotly_axis_style(
            x_title=DriskChartFactory.VALUE_AXIS_TITLE, 
            x_unit=DriskChartFactory.VALUE_AXIS_UNIT,    
            y_title="相对频数 (%)", 
            x_range=x_range,
            y_range=[0, y_axis_max],
            x_dtick=x_dtick, 
            y_dtick=y_dtick, 
            fixedrange=True, 
            forced_mag=forced_mag,
            label_overrides=label_overrides,
            x_axis_numeric=bool(axis_flags.get("x", True)),
            y_axis_numeric=bool(axis_flags.get("y", True)),
        )

        has_x_title = bool(axis_style["xaxis"].get("title"))
        b_margin = margins[3] if has_x_title else 25

        layout_dict = {
            "template": "plotly_white",
            "margin": dict(l=margins[0], r=margins[1], t=margins[2], b=b_margin),
            "xaxis": axis_style["xaxis"],
            "yaxis": axis_style["yaxis"],
            # 离散概率柱图需要柱子尽量紧邻显示，避免出现误导性的间隙。
            "bargap": 0.0,
            "bargroupgap": 0.0,
            "showlegend": False,
            "barmode": "overlay"
        }
        DriskChartFactory._inject_chart_title_override(layout_dict, label_overrides)
        
        if show_overlay_cdf:
            DriskChartFactory._inject_overlay_cdf(fig, layout_dict, [], [], add_trace=False)

        fig.update_layout(layout_dict)
        return {
            "plot_json": json.dumps(fig.to_dict(), cls=plotly.utils.PlotlyJSONEncoder), 
            "js_mode": "discrete", 
            "js_logic": DriskChartFactory._get_discrete_pixel_perfect_js()
        }

    # ------------------------------------------------------------
    # ------------------------------------------------------------

    @staticmethod
    def build_cdf(
        data_groups: List[Dict[str, Any]],
        x_range: Tuple[float, float],
        x_dtick: float,
        margins: Tuple[int, int, int, int] = (80, 40, 30, 50),
        forced_mag: Optional[int] = None,
        label_overrides: Optional[Dict[str, Any]] = None,
        axis_numeric_flags: Optional[Dict[str, bool]] = None,
    ) -> dict:
        """
        构建累计概率图（CDF）。

        支持能力：
        1. 单组或多组 CDF 曲线；
        2. 离散阶梯线与连续曲线两种语义；
        3. 置信区间上下界填充带；
        4. 统一接入坐标轴与标题覆盖配置。
        """

        fig = go.Figure()

        # 以下三个局部函数用于把离散 CDF 的支持点序列扩展为可视化阶梯线。
        def _expand_step_axis(x_unique: np.ndarray) -> np.ndarray:
            if x_unique.size <= 1:
                return x_unique
            x_step = np.empty(2 * x_unique.size - 1, dtype=float)
            x_step[0::2] = x_unique
            x_step[1::2] = x_unique[1:]
            return x_step

        def _expand_step_values(y_unique: np.ndarray) -> np.ndarray:
            if y_unique.size <= 1:
                return y_unique
            y_step = np.empty(2 * y_unique.size - 1, dtype=float)
            y_step[0::2] = y_unique
            y_step[1::2] = y_unique[:-1]
            return y_step

        def _stepify_series(x_vals: np.ndarray, *y_vals: np.ndarray):
            if x_vals.size == 0:
                return np.array([], dtype=float), [np.array([], dtype=float) for _ in y_vals]
            order = np.argsort(x_vals, kind="stable")
            x_sorted = np.asarray(x_vals, dtype=float)[order]
            y_sorted = [np.asarray(y, dtype=float)[order] for y in y_vals]
            x_unique, first_idx, counts = np.unique(x_sorted, return_index=True, return_counts=True)
            last_idx = first_idx + counts - 1
            y_unique = [arr[last_idx] for arr in y_sorted]
            x_step = _expand_step_axis(x_unique.astype(float))
            y_step = [_expand_step_values(arr.astype(float)) for arr in y_unique]
            return x_step, y_step

        for group in data_groups:
            xs = _to_float_array(group.get("x", []))
            ys = _to_float_array(group.get("y", []))
            base_color = group.get("color", DRISK_COLOR_CYCLE[1])
            line_shape = str(group.get("line_shape", "linear") or "linear").lower()
            step_mode = (line_shape == "hv")

            n_curve = min(xs.size, ys.size)
            if n_curve <= 0:
                continue
            xs_work = np.asarray(xs[:n_curve], dtype=float)
            ys_work = np.asarray(ys[:n_curve], dtype=float)
            xs_plot = xs_work
            ys_plot = ys_work
            y_up_plot = None
            y_dn_plot = None
            y_upper_raw = group.get("y_upper")
            y_lower_raw = group.get("y_lower")
            
            if y_upper_raw is not None and y_lower_raw is not None:
                y_up = _to_float_array(y_upper_raw)
                y_dn = _to_float_array(y_lower_raw)
                
                if y_up.size > 0 and y_dn.size > 0:
                    n_band = min(xs_work.size, ys_work.size, y_up.size, y_dn.size)
                    if n_band > 0:
                        xs_band = np.asarray(xs_work[:n_band], dtype=float)
                        ys_band = np.asarray(ys_work[:n_band], dtype=float)
                        y_up = np.asarray(y_up[:n_band], dtype=float)
                        y_dn = np.asarray(y_dn[:n_band], dtype=float)

                        if step_mode:
                            xs_plot, step_series = _stepify_series(xs_band, ys_band, y_up, y_dn)
                            ys_plot, y_up_plot, y_dn_plot = step_series
                        else:
                            xs_plot = xs_band
                            ys_plot = ys_band
                            y_up_plot = y_up
                            y_dn_plot = y_dn

            elif step_mode:
                xs_plot, step_series = _stepify_series(xs_work, ys_work)
                ys_plot = step_series[0]

            if y_up_plot is not None and y_dn_plot is not None:
                n_fill = min(xs_plot.size, y_up_plot.size, y_dn_plot.size)
                if n_fill > 0:
                    fill_color = DriskChartFactory._hex_to_rgba(base_color, 0.2)
                    fig.add_trace(go.Scatter(
                        x=xs_plot[:n_fill], y=y_up_plot[:n_fill], mode="lines",
                        line=dict(width=0, shape="linear"), showlegend=False, hoverinfo="skip"
                    ))

                    fig.add_trace(go.Scatter(
                        x=xs_plot[:n_fill], y=y_dn_plot[:n_fill], mode="lines",
                        fill="tonexty", fillcolor=fill_color,
                        line=dict(width=0, shape="linear"), showlegend=False, hoverinfo="skip"
                    ))

            curve_color = group.get("curve_color", base_color)
            curve_opacity = group.get("curve_opacity", 1.0)
            curve_width = group.get("curve_width", 1.8)
            curve_dash = group.get("dash", "solid")

            if curve_color.startswith("#") and curve_opacity < 1.0:
                curve_color = DriskChartFactory._hex_to_rgba(curve_color, curve_opacity)
            fig.add_trace(
                go.Scatter(
                    x=xs_plot,
                    y=ys_plot,
                    mode="lines",
                    name=str(group.get("name", "")),
                    line=dict(
                        color=curve_color,
                        width=curve_width,
                        dash=curve_dash,
                        shape="linear" if step_mode else line_shape,
                    ),
                    hoverinfo="skip",
                )
            )

        
        axis_flags = axis_numeric_flags if isinstance(axis_numeric_flags, dict) else {}
        axis_style = build_plotly_axis_style(
            x_title=DriskChartFactory.VALUE_AXIS_TITLE,
            x_unit=DriskChartFactory.VALUE_AXIS_UNIT,    
            y_title="概率",
            x_range=x_range,
            y_range=[0, 1.05],
            x_dtick=x_dtick,
            y_dtick=0.1,
            fixedrange=True,
            forced_mag=forced_mag,
            label_overrides=label_overrides,
            x_axis_numeric=bool(axis_flags.get("x", True)),
            y_axis_numeric=bool(axis_flags.get("y", True)),
        )

        has_x_title = bool(axis_style["xaxis"].get("title"))
        b_margin = margins[3] if has_x_title else 25

        layout_dict = {
            "autosize": True,
            "template": "plotly_white",
            "margin": dict(l=margins[0], r=margins[1], t=margins[2], b=b_margin),
            "xaxis": axis_style["xaxis"],
            "yaxis": axis_style["yaxis"],
            "showlegend": False,
        }
        DriskChartFactory._inject_chart_title_override(layout_dict, label_overrides)
        layout = go.Layout(**layout_dict)

        return {
            "plot_json": json.dumps(go.Figure(fig.data, layout).to_dict(), cls=plotly.utils.PlotlyJSONEncoder),
            "js_mode": "cdf_qt_mode", 
            "js_logic": "",
        }

    # ------------------------------------------------------------
    # ------------------------------------------------------------
    
    @staticmethod
    def build_boxplot_figure(
        series_map: Dict[str, Any],
        display_keys: List[str],
        style_map: Dict[str, Dict[str, Any]],
        title: str = "",
        plot_type: str = "boxplot",
        margins: Tuple[int, int, int, int] = (80, 120, 50, 80),
        forced_mag: Optional[int] = None,
        display_label_map: Optional[Dict[str, str]] = None,
        label_overrides: Optional[Dict[str, Any]] = None,
        axis_numeric_flags: Optional[Dict[str, bool]] = None,
    ) -> dict:
        """
        构建箱线图家族图形。

        支持的 plot_type：
        1. boxplot：标准箱线图；
        2. violin：小提琴图；
        3. letter_value：信度分层箱图；
        4. trend：按序列位置展示 5%/25%/75%/95% 区间与均值的趋势图。

        说明：
        该方法不仅生成主体图形，也会同时生成右侧图例示意所需的 shapes 与 annotations。
        """
        fig = go.Figure()
        calc_y_max = -float('inf')
        calc_y_min = float('inf')

        # 趋势图右侧不需要额外说明区，因此缩小右边距。
        if plot_type == "trend":
            margins = (margins[0], 50, margins[2], margins[3])
        parsed_keys = []
        base_counts = {}
        label_map = display_label_map if isinstance(display_label_map, dict) else {}
        for k in display_keys:
            key_text = str(k)
            match = re.match(r"^(.*?)\s*\(sim(\d+)\)$", key_text)
            if match:
                raw_base = match.group(1).strip()
                sim_suffix = f"sim{match.group(2)}"
            else:
                raw_base = key_text.strip()
                sim_suffix = ""
            mapped_base = str(label_map.get(k, "") or label_map.get(key_text, "")).strip()
            if mapped_base:
                clean_base = mapped_base
            else:
                name_match = re.search(r"\(([^)]+)\)", raw_base)
                clean_base = name_match.group(1).strip() if name_match else (raw_base.split("!")[-1].strip() if "!" in raw_base else raw_base.strip())
            parsed_keys.append({"original": k, "base": clean_base, "sim": sim_suffix})
            base_counts[clean_base] = base_counts.get(clean_base, 0) + 1

        x_labels = [f"{item['base']}{item['sim']}" if base_counts[item["base"]] > 1 and item["sim"] else item["base"] for item in parsed_keys]

        num_series = len(display_keys)
        x_positions = list(range(1, num_series + 1))       
        tick_vals = list(range(0, num_series + 2))         
        tick_texts = [""] + x_labels + [""]                

        primary_key = display_keys[0] if display_keys else None
        p_st = style_map.get(primary_key, {}) if primary_key else {}
        p_base_color = p_st.get("color", DRISK_COLOR_CYCLE[0])
        p_fill_color = p_st.get("fill_color", p_base_color)
        p_fill_opacity = p_st.get("fill_opacity", 0.85)
        
        p_outline_color = p_st.get("outline_color", "#666666")
        if p_outline_color == "rgba(255,255,255,0.5)": p_outline_color = "#666666" 
        p_outline_width = p_st.get("outline_width", 0.5)
        
        p_curve_color = p_st.get("curve_color", "#ffffff") 
        p_curve_width = p_st.get("curve_width", 1.5)
        p_curve_dash = p_st.get("dash", "solid")
        p_curve_opacity = p_st.get("curve_opacity", 1.0)
        
        p_mean_color = p_st.get("mean_color", "#000000")   
        p_mean_width = p_st.get("mean_width", 1.5)
        p_mean_dash = p_st.get("mean_dash", "solid")
        p_mean_opacity = p_st.get("mean_opacity", 1.0)
        
        p_fill_rgba = DriskChartFactory._hex_to_rgba(p_fill_color, p_fill_opacity)
        p_outline_rgba = DriskChartFactory._hex_to_rgba(p_outline_color, 1.0) if p_outline_color.startswith("#") else p_outline_color
        p_curve_rgba = DriskChartFactory._hex_to_rgba(p_curve_color, p_curve_opacity) if p_curve_color.startswith("#") else p_curve_color
        p_mean_rgba = DriskChartFactory._hex_to_rgba(p_mean_color, p_mean_opacity) if p_mean_color.startswith("#") else p_mean_color
        p_box_line_style = dict(width=p_outline_width, color=p_outline_rgba)

        def _get_lighter_color(hex_color, depth, opacity=1.0):
            """
            根据层级生成更浅的填充颜色。
            主要用于 letter value boxplot 中不同深度箱体的视觉区分。
            """
            factor = min((depth - 1) * 0.12, 0.8)
            if hex_color == "transparent": return "rgba(0,0,0,0)"
            h = hex_color.lstrip('#')
            if len(h) not in (6, 8): return hex_color
            if len(h) == 8: h = h[2:] 
            r, g, b = tuple(int(h[i:i+2], 16) for i in (0, 2, 4))
            return f"rgba({int(r + (255-r)*factor)}, {int(g + (255-g)*factor)}, {int(b + (255-b)*factor)}, {opacity})"

        base_width = 0.7 
        outlier_marker = dict(symbol='cross-thin', size=4, color='#888888', line=dict(width=0.8, color='#888888'))
        if plot_type == "trend":
            trend_valid_x = []
            trend_p5, trend_p25, trend_p75, trend_p95, trend_mean = [], [], [], [], []
            
            for i, k in enumerate(display_keys):
                arr = _to_float_array(series_map.get(k))
                if arr.size == 0: continue
                p5_val = float(np.percentile(arr, 5))
                p25_val = float(np.percentile(arr, 25))
                p75_val = float(np.percentile(arr, 75))
                p95_val = float(np.percentile(arr, 95))
                calc_y_max = max(calc_y_max, p95_val)
                calc_y_min = min(calc_y_min, p5_val)
                
                trend_valid_x.append(x_positions[i])
                trend_p5.append(p5_val)
                trend_p25.append(p25_val)
                trend_p75.append(p75_val)
                trend_p95.append(p95_val)
                trend_mean.append(float(np.mean(arr)))
            
            if trend_valid_x:
                if len(trend_valid_x) == 1:
                    # 
                    hw = base_width / 2
                    x_coords = [trend_valid_x[0] - hw, trend_valid_x[0] + hw]
                    trend_p5 = [trend_p5[0], trend_p5[0]]
                    trend_p25 = [trend_p25[0], trend_p25[0]]
                    trend_p75 = [trend_p75[0], trend_p75[0]]
                    trend_p95 = [trend_p95[0], trend_p95[0]]
                    trend_mean = [trend_mean[0], trend_mean[0]]
                else:
                    x_coords = trend_valid_x
                    
                
                c_dark = DriskChartFactory._hex_to_rgba(p_mean_color, 0.4 * p_mean_opacity) if p_mean_color != "transparent" else "rgba(0,0,0,0)"
                c_light = DriskChartFactory._hex_to_rgba(p_mean_color, 0.15 * p_mean_opacity) if p_mean_color != "transparent" else "rgba(0,0,0,0)"
                
                fig.add_trace(go.Scatter(x=x_coords, y=trend_p5, mode='lines', line=dict(width=0), showlegend=False, hoverinfo='skip'))
                fig.add_trace(go.Scatter(x=x_coords, y=trend_p95, mode='lines', fill='tonexty', fillcolor=c_light, line=dict(width=0), showlegend=False, hoverinfo='skip'))
                
                fig.add_trace(go.Scatter(x=x_coords, y=trend_p25, mode='lines', line=dict(width=0), showlegend=False, hoverinfo='skip'))
                fig.add_trace(go.Scatter(x=x_coords, y=trend_p75, mode='lines', fill='tonexty', fillcolor=c_dark, line=dict(width=0), showlegend=False, hoverinfo='skip'))
                
                fig.add_trace(go.Scatter(x=x_coords, y=trend_mean, mode='lines', line=dict(color=p_mean_rgba, width=p_mean_width, dash=p_mean_dash), showlegend=False, hoverinfo='y'))
                
                fig.add_trace(go.Scatter(x=[None], y=[None], mode='markers', marker=dict(symbol='square', size=15, color=c_dark, line=dict(width=0)), name='25% - 75% 区间'))
                fig.add_trace(go.Scatter(x=[None], y=[None], mode='markers', marker=dict(symbol='square', size=15, color=c_light, line=dict(width=0)), name='5% - 95% 区间'))
                
                fig.add_trace(go.Scatter(x=[None], y=[None], mode='lines', line=dict(color=p_mean_rgba, width=p_mean_width, dash=p_mean_dash), name='均值'))
        else:
            for i, k in enumerate(display_keys):
                arr = _to_float_array(series_map.get(k))
                if arr.size == 0: continue
                
                calc_y_max = max(calc_y_max, float(np.max(arr)))
                calc_y_min = min(calc_y_min, float(np.min(arr)))
                
                st = p_st
                base_color = p_base_color
                fill_color = st.get("fill_color", base_color)
                fill_opacity = st.get("fill_opacity", 0.85)
                outline_color = st.get("outline_color", "#666666")
                if outline_color == "rgba(255,255,255,0.5)": outline_color = "#666666"
                outline_width = st.get("outline_width", 0.5)
                
                curve_color = st.get("curve_color", "#ffffff") 
                curve_width = st.get("curve_width", 1.5)
                curve_dash = st.get("dash", "solid")
                curve_opacity = st.get("curve_opacity", 1.0)
                
                mean_color = st.get("mean_color", "#000000")   
                mean_width = st.get("mean_width", 1.5)
                mean_dash = st.get("mean_dash", "solid")
                mean_opacity = st.get("mean_opacity", 1.0)
                
                fill_rgba = DriskChartFactory._hex_to_rgba(fill_color, fill_opacity)
                outline_rgba = DriskChartFactory._hex_to_rgba(outline_color, 1.0) if outline_color.startswith("#") else outline_color
                curve_rgba = DriskChartFactory._hex_to_rgba(curve_color, curve_opacity) if curve_color.startswith("#") else curve_color
                mean_rgba = DriskChartFactory._hex_to_rgba(mean_color, mean_opacity) if mean_color.startswith("#") else mean_color
                
                box_line_style = dict(width=outline_width, color=outline_rgba)
                
                median_v = float(np.median(arr))
                mean_v = float(np.mean(arr))

                # === 信度分层箱图（Letter Value Boxplot）===
                if plot_type == "letter_value":
                    n = len(arr)
                    max_depth = 1
                    
                    while (n * (0.5 ** (max_depth + 1))) > 6 and max_depth < 8: 
                        max_depth += 1
                    
                    for depth in range(max_depth, 0, -1):
                        p_tail = (0.5 ** depth)
                        p_low, p_high = p_tail / 2, 1 - (p_tail / 2)
                        v_low, v_high = np.percentile(arr, p_low * 100), np.percentile(arr, p_high * 100)
                        
                        if depth == max_depth: 
                            final_v_low, final_v_high = v_low, v_high
                        
                        fill_c = _get_lighter_color(fill_color, depth, opacity=1.0)
                        width = base_width * (0.85 ** (depth - 1))
                        
                        fig.add_trace(go.Box(
                            y=[v_low, v_low, median_v, v_high, v_high],
                            x=[x_positions[i]] * 5,
                            name=x_labels[i],
                            width=width,
                            fillcolor=fill_c,
                            line=box_line_style,
                            showlegend=False,
                            boxpoints=False,
                            hoverinfo="y",
                            opacity=1.0
                        ))
                    
                    hw = base_width / 2
                    fig.add_trace(go.Scatter(
                        x=[x_positions[i]-hw, x_positions[i]+hw], y=[median_v, median_v],
                        mode='lines', line=dict(color=curve_rgba, width=curve_width, dash=curve_dash), showlegend=False, hoverinfo='skip'
                    ))
                    
                    outliers = arr[(arr < final_v_low) | (arr > final_v_high)]
                    if len(outliers) > 0:
                        fig.add_trace(go.Scatter(y=outliers, x=[x_positions[i]] * len(outliers), mode='markers', marker=outlier_marker, showlegend=False, hoverinfo="y"))
                        
                # === 小提琴图（Violin）===
                elif plot_type == "violin":
                    fig.add_trace(go.Violin(
                        y=arr, x=[x_positions[i]] * len(arr), name=x_labels[i], width=base_width, 
                        box_visible=False, meanline_visible=False, points='outliers', marker=outlier_marker,
                        line=box_line_style, fillcolor=fill_color, opacity=fill_opacity 
                    ))
                    
                    p5, p25, p75, p95 = np.percentile(arr, 5), np.percentile(arr, 25), np.percentile(arr, 75), np.percentile(arr, 95)
                    
                    fig.add_trace(go.Scatter(x=[x_positions[i], x_positions[i]], y=[p5, p95], mode='lines', line=dict(color=outline_rgba, width=outline_width), showlegend=False, hoverinfo='skip'))
                    fig.add_trace(go.Scatter(x=[x_positions[i], x_positions[i]], y=[p25, p75], mode='lines', line=dict(color=outline_rgba, width=outline_width * 4), showlegend=False, hoverinfo='skip'))
                    
                    vw = 0.15 
                    fig.add_trace(go.Scatter(
                        x=[x_positions[i]-vw, x_positions[i]+vw], y=[median_v, median_v],
                        mode='lines', line=dict(color=curve_rgba, width=curve_width, dash=curve_dash), showlegend=False, hoverinfo='skip'
                    ))
                    fig.add_trace(go.Scatter(
                    x=[x_positions[i]-vw, x_positions[i]+vw], y=[mean_v, mean_v],
                        mode='lines', line=dict(color=mean_rgba, width=mean_width, dash=mean_dash), showlegend=False, hoverinfo='skip'
                    ))
                else:
                    fig.add_trace(go.Box(
                        y=arr, x=[x_positions[i]] * len(arr), name=x_labels[i], width=base_width, 
                        boxpoints='outliers', marker=outlier_marker, line=box_line_style, fillcolor=fill_color, opacity=fill_opacity, boxmean=False 
                    ))
                    
                    hw = base_width / 2
                    fig.add_trace(go.Scatter(
                        x=[x_positions[i]-hw, x_positions[i]+hw], y=[median_v, median_v],
                        mode='lines', line=dict(color=curve_rgba, width=curve_width, dash=curve_dash), showlegend=False, hoverinfo='skip'
                    ))
                    fig.add_trace(go.Scatter(
                    x=[x_positions[i]-hw, x_positions[i]+hw], y=[mean_v, mean_v],
                        mode='lines', line=dict(color=mean_rgba, width=mean_width, dash=mean_dash), showlegend=False, hoverinfo='skip'
                    ))

        if calc_y_max == -float('inf'): calc_y_max, calc_y_min = 1.0, 0.0

        y_span = calc_y_max - calc_y_min
        y_span = y_span if y_span > 0 else 1.0
        y_dtick = DriskMath.calc_smart_step(y_span)
        
        padded_max = calc_y_max + 0.01 * y_dtick
        padded_min = calc_y_min - 0.01 * y_dtick
        
        y_axis_max = DriskMath.ceil_to_tick(padded_max, y_dtick)
        y_axis_min = DriskMath.floor_to_tick(padded_min, y_dtick)

        axis_flags = axis_numeric_flags if isinstance(axis_numeric_flags, dict) else {}
        axis_style = build_plotly_axis_style(
            x_title="", 
            y_title=DriskChartFactory.VALUE_AXIS_TITLE, 
            y_unit=DriskChartFactory.VALUE_AXIS_UNIT,    
            x_range=None,
            y_range=[y_axis_min, y_axis_max],
            x_dtick=None,
            y_dtick=y_dtick,
            fixedrange=True,
            y_forced_mag=forced_mag,
            label_overrides=label_overrides,
            x_axis_numeric=bool(axis_flags.get("x", True)),
            y_axis_numeric=bool(axis_flags.get("y", True)),
        )
        if "yaxis" in axis_style: axis_style["yaxis"].update({"showgrid": True, "gridcolor": "#dddddd", "gridwidth": 0.5, "mirror": False})

        has_x_title = bool(axis_style["xaxis"].get("title"))
        b_margin = margins[3] if has_x_title else 25
        legend_x_base = 1.01 
        ann_font = dict(size=13, color="#333", family="Arial")
        shapes, annotations = [], []
        
        if plot_type == "trend":
            pass
        elif plot_type == "letter_value":
            c_depth1 = _get_lighter_color(p_fill_color, 1, 1.0)
            c_depth2 = _get_lighter_color(p_fill_color, 2, 1.0)
            c_depth3 = _get_lighter_color(p_fill_color, 3, 1.0)
            shapes = [
                dict(type="rect", xref="paper", yref="paper", x0=legend_x_base+0.006, x1=legend_x_base+0.014, y0=0.3, y1=0.7, fillcolor=c_depth3, line=p_box_line_style),
                dict(type="rect", xref="paper", yref="paper", x0=legend_x_base+0.003, x1=legend_x_base+0.017, y0=0.35, y1=0.65, fillcolor=c_depth2, line=p_box_line_style),
                dict(type="rect", xref="paper", yref="paper", x0=legend_x_base, x1=legend_x_base+0.020, y0=0.4, y1=0.6, fillcolor=c_depth1, line=p_box_line_style),
                dict(type="line", xref="paper", yref="paper", x0=legend_x_base, x1=legend_x_base+0.020, y0=0.47, y1=0.47, line=dict(color=p_curve_rgba, width=p_curve_width, dash=p_curve_dash)),
            ]
            annotations = [
                dict(xref="paper", yref="paper", x=legend_x_base+0.028, y=0.7, text="上尾部", showarrow=False, font=ann_font, xanchor="left"),
                dict(xref="paper", yref="paper", x=legend_x_base+0.028, y=0.6, text="75%", showarrow=False, font=ann_font, xanchor="left"),
                dict(xref="paper", yref="paper", x=legend_x_base+0.028, y=0.47, text="中位数", showarrow=False, font=ann_font, xanchor="left"),
                dict(xref="paper", yref="paper", x=legend_x_base+0.028, y=0.4, text="25%", showarrow=False, font=ann_font, xanchor="left"),
                dict(xref="paper", yref="paper", x=legend_x_base+0.028, y=0.3, text="下尾部", showarrow=False, font=ann_font, xanchor="left"),
            ]
            
        elif plot_type == "violin":
            shapes = [
                dict(type="rect", xref="paper", yref="paper", x0=legend_x_base, x1=legend_x_base+0.020, y0=0.3, y1=0.7, fillcolor=p_fill_rgba, line=p_box_line_style),
                dict(type="line", xref="paper", yref="paper", x0=legend_x_base+0.010, x1=legend_x_base+0.010, y0=0.25, y1=0.75, line=dict(color=p_outline_rgba, width=p_outline_width)),
                dict(type="line", xref="paper", yref="paper", x0=legend_x_base+0.010, x1=legend_x_base+0.010, y0=0.4, y1=0.6, line=dict(color=p_outline_rgba, width=p_outline_width * 4)),
                dict(type="line", xref="paper", yref="paper", x0=legend_x_base+0.004, x1=legend_x_base+0.016, y0=0.47, y1=0.47, line=dict(color=p_curve_rgba, width=p_curve_width, dash=p_curve_dash)),
                dict(type="line", xref="paper", yref="paper", x0=legend_x_base+0.004, x1=legend_x_base+0.016, y0=0.53, y1=0.53, line=dict(color=p_mean_rgba, width=p_mean_width, dash=p_mean_dash)), 
            ]
            annotations = [
                dict(xref="paper", yref="paper", x=legend_x_base+0.028, y=0.75, text="95%", showarrow=False, font=ann_font, xanchor="left"),
                dict(xref="paper", yref="paper", x=legend_x_base+0.028, y=0.6,  text="75%", showarrow=False, font=ann_font, xanchor="left"),
                dict(xref="paper", yref="paper", x=legend_x_base+0.028, y=0.53, text="均值", showarrow=False, font=ann_font, xanchor="left"),
                dict(xref="paper", yref="paper", x=legend_x_base+0.028, y=0.47, text="中位数", showarrow=False, font=ann_font, xanchor="left"),
                dict(xref="paper", yref="paper", x=legend_x_base+0.028, y=0.4,  text="25%", showarrow=False, font=ann_font, xanchor="left"),
                dict(xref="paper", yref="paper", x=legend_x_base+0.028, y=0.25, text="5%",  showarrow=False, font=ann_font, xanchor="left"),
            ]
            
        else:
            shapes = [
                dict(type="rect", xref="paper", yref="paper", x0=legend_x_base, x1=legend_x_base+0.020, y0=0.4, y1=0.6, fillcolor=p_fill_rgba, line=p_box_line_style),
                dict(type="line", xref="paper", yref="paper", x0=legend_x_base, x1=legend_x_base+0.020, y0=0.47, y1=0.47, line=dict(color=p_curve_rgba, width=p_curve_width, dash=p_curve_dash)),
                dict(type="line", xref="paper", yref="paper", x0=legend_x_base, x1=legend_x_base+0.020, y0=0.53, y1=0.53, line=dict(color=p_mean_rgba, width=p_mean_width, dash=p_mean_dash)), 
                dict(type="line", xref="paper", yref="paper", x0=legend_x_base+0.010, x1=legend_x_base+0.010, y0=0.6, y1=0.75, line=p_box_line_style),
                dict(type="line", xref="paper", yref="paper", x0=legend_x_base+0.010, x1=legend_x_base+0.010, y0=0.25, y1=0.4, line=p_box_line_style),
            ]
            annotations = [
                dict(xref="paper", yref="paper", x=legend_x_base+0.028, y=0.75, text="95%", showarrow=False, font=ann_font, xanchor="left"),
                dict(xref="paper", yref="paper", x=legend_x_base+0.028, y=0.6,  text="75%", showarrow=False, font=ann_font, xanchor="left"),
                dict(xref="paper", yref="paper", x=legend_x_base+0.028, y=0.53, text="均值", showarrow=False, font=ann_font, xanchor="left"),
                dict(xref="paper", yref="paper", x=legend_x_base+0.028, y=0.47, text="中位数", showarrow=False, font=ann_font, xanchor="left"),
                dict(xref="paper", yref="paper", x=legend_x_base+0.028, y=0.4,  text="25%", showarrow=False, font=ann_font, xanchor="left"),
                dict(xref="paper", yref="paper", x=legend_x_base+0.028, y=0.25, text="5%",  showarrow=False, font=ann_font, xanchor="left"),
            ]

        layout_dict = {
            "title": dict(text=f"{title}", x=0.5, xref="paper", xanchor="center", font=dict(size=14, family="Arial, 'Microsoft YaHei', sans-serif", color="#333333")),  
            "template": "plotly_white",
            "plot_bgcolor": "rgba(0,0,0,0)",
            "paper_bgcolor": "rgba(0,0,0,0)",
            "margin": dict(l=margins[0], r=margins[1], t=margins[2], b=b_margin),
            "boxmode": "overlay",
            "xaxis": dict(
                tickangle=0, showgrid=True, gridcolor='#dddddd', gridwidth=0.5,
                showline=True, linecolor='#000', ticks='outside', 
                tickfont=dict(size=13, family='Arial'), zeroline=False,
                mirror=False, range=[0, num_series + 1],  
                tickmode='array', tickvals=tick_vals, ticktext=tick_texts         
            ),
            "yaxis": axis_style["yaxis"],
            "shapes": shapes, "annotations": annotations,
        }

        if plot_type == "trend":
            layout_dict["showlegend"] = True
            layout_dict["legend"] = dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1, font=dict(size=12, color="#333", family="Arial"))
        else:
            layout_dict["showlegend"] = False

        DriskChartFactory._inject_chart_title_override(layout_dict, label_overrides)
        fig.update_layout(layout_dict)
        return {"plot_json": json.dumps(fig.to_dict(), cls=plotly.utils.PlotlyJSONEncoder), "js_mode": "boxplot", "js_logic": ""}
                    
    # ------------------------------------------------------------
    # 2.6 散点图与微聚类压缩
    # ------------------------------------------------------------

    @staticmethod
    def _extract_scatter_group(group: Dict[str, Any], index: int) -> Tuple[np.ndarray, np.ndarray, str]:
        """
        从散点图输入组中提取 x、y 数据及显示名称。

        兼容键名：
        - x / y
        - x_data / y_data
        - name / label
        """
        name = group.get("name", group.get("label", f"序列 {index + 1}"))
        x_src = group.get("x", group.get("x_data", []))
        y_src = group.get("y", group.get("y_data", []))
        return _to_float_array(x_src), _to_float_array(y_src), str(name)

    @staticmethod
    def build_scatter_figure(data_groups: List[Dict[str, Any]], y_title: str, 
                             margins: Tuple[int, int, int, int] = (80, 40, 50, 60),
                             x_forced_mag: Optional[int] = None,
                             y_forced_mag: Optional[int] = None,
                             x_title: str = "X 轴",
                              x_unit: str = "",              
                              y_unit: str = "",
                             style_map: Optional[Dict[str, Dict[str, Any]]] = None,
                             label_overrides: Optional[Dict[str, Any]] = None,
                             axis_numeric_flags: Optional[Dict[str, bool]] = None) -> Dict[str, Any]: 
        """
        构建散点图。

        特点：
        1. 单次最多渲染 100 组序列，避免界面失控；
        2. 当单组点数过多时，使用 MiniBatchKMeans 做微聚类压缩，平衡性能与空间分布表达；
        3. 支持每组独立颜色、透明度、符号、描边和点大小配置；
        4. 自动根据所有数据推导横纵轴范围与步长。
        """
        fig = go.Figure()
        all_x, all_y = [], []
        style_map = style_map or {}
        data_groups = data_groups[:100]  

        # 逐组构建散点 trace，并在必要时对超大点集做压缩。
        for i, group in enumerate(data_groups):
            x_arr, y_arr, name = DriskChartFactory._extract_scatter_group(group, i)
            display_name = str(group.get("display_name", name) or name)

            min_len = min(len(x_arr), len(y_arr))
            if min_len == 0: continue
            x_arr, y_arr = x_arr[:min_len], y_arr[:min_len]

            all_x.append(x_arr)
            all_y.append(y_arr)

            # 当单组点数过多时，优先使用微聚类替代全量散点，降低渲染压力。
            MAX_POINTS = 2000
            if min_len > MAX_POINTS:
                try:
                    X = np.column_stack((x_arr, y_arr))
                    kmeans = MiniBatchKMeans(n_clusters=MAX_POINTS, random_state=42, batch_size=10000, n_init="auto")
                    kmeans.fit(X)
                    x_plot = kmeans.cluster_centers_[:, 0]
                    y_plot = kmeans.cluster_centers_[:, 1]
                    weights = np.bincount(kmeans.labels_, minlength=MAX_POINTS)
                except ImportError:
                    idx = np.random.choice(min_len, MAX_POINTS, replace=False)
                    x_plot = x_arr[idx]
                    y_plot = y_arr[idx]
                    weights = np.ones(MAX_POINTS)
            else:
                x_plot, y_plot, weights = x_arr, y_arr, np.ones(min_len)

            
            base_color = DRISK_COLOR_CYCLE[i % len(DRISK_COLOR_CYCLE)]
            st = style_map.get(name, {})
            fill_color = st.get("fill_color", base_color)
            fill_opacity = float(st.get("fill_opacity", 0.6))
            fill_opacity = max(0.0, min(1.0, fill_opacity))
            outline_color = st.get("outline_color", "#ffffff")
            outline_width = float(st.get("outline_width", 0.0))
            outline_width = max(0.0, outline_width)
            trace_opacity = float(st.get("opacity", 1.0))
            trace_opacity = max(0.0, min(1.0, trace_opacity))
            base_size = float(st.get("marker_size", 4.5))
            base_size = max(0.5, min(60.0, base_size))
            symbol_idx = (i // len(DRISK_COLOR_CYCLE)) % len(SCATTER_SYMBOLS)
            current_symbol = str(st.get("marker_symbol", SCATTER_SYMBOLS[symbol_idx]))
            if current_symbol not in SCATTER_SYMBOLS:
                current_symbol = SCATTER_SYMBOLS[symbol_idx]

            if np.max(weights) > 1:
                log_weights = np.log1p(weights)
                max_log = np.max(log_weights) if np.max(log_weights) > 0 else 1.0
                sizes = base_size * (0.78 + (log_weights / max_log) * 1.89)
                
                marker_dict = dict(
                    color=DriskChartFactory._hex_to_rgba(fill_color, fill_opacity),
                    symbol=current_symbol,
                    size=sizes,
                    opacity=1.0,
                    line=dict(width=outline_width, color=outline_color),
                )
            else:
                
                marker_dict = dict(
                    color=DriskChartFactory._hex_to_rgba(fill_color, fill_opacity),
                    symbol=current_symbol,
                    size=base_size,
                    opacity=1.0,
                    line=dict(width=outline_width, color=outline_color),
                )

            fig.add_trace(go.Scatter(
                x=x_plot, y=y_plot, mode='markers', name=display_name,
                marker=marker_dict, hoverinfo='skip', opacity=trace_opacity
            ))

        # 用全部有效散点统一反推坐标轴范围，避免不同序列各自裁剪。
        if all_x and all_y:
            cat_x = np.concatenate(all_x)
            cat_y = np.concatenate(all_y)
            def _calc_axis_range_and_dtick(arr: np.ndarray):
                vmin, vmax = float(np.min(arr)), float(np.max(arr))
                span = (vmax - vmin) if vmax > vmin else 1.0
                dtick = DriskMath.calc_smart_step(span)
                padded_min = vmin - 0.01 * dtick
                padded_max = vmax + 0.01 * dtick
                
                axis_min = math.floor(padded_min / dtick) * dtick
                axis_max = math.ceil(padded_max / dtick) * dtick
                
                if abs(axis_min - vmin) < 1e-9: axis_min -= dtick
                if abs(axis_max - vmax) < 1e-9: axis_max += dtick
                return [axis_min, axis_max], dtick

            x_range, x_dtick = _calc_axis_range_and_dtick(cat_x)
            y_range, y_dtick = _calc_axis_range_and_dtick(cat_y)
        else:
            x_range, x_dtick = [0, 1], 1
            y_range, y_dtick = [0, 1], 1

        axis_flags = axis_numeric_flags if isinstance(axis_numeric_flags, dict) else {}
        axis_style = build_plotly_axis_style(
            x_title=x_title,
            x_unit=x_unit,                 
            y_title=f"{y_title}",
            y_unit=y_unit,                 
            x_range=x_range,
            y_range=y_range,
            x_dtick=x_dtick,
            y_dtick=y_dtick,
            fixedrange=True,
            forced_mag=x_forced_mag,
            y_forced_mag=y_forced_mag,
            label_overrides=label_overrides,
            x_axis_numeric=bool(axis_flags.get("x", True)),
            y_axis_numeric=bool(axis_flags.get("y", True)),
        )
        if "yaxis" in axis_style:
            axis_style["yaxis"].update({"mirror": False, "side": "left"})

        safe_margins = (margins[0], 0, 0, margins[3])

        layout_dict = {
            "autosize": True,
            "template": "plotly_white",
            "margin": dict(l=safe_margins[0], r=safe_margins[1], t=safe_margins[2], b=safe_margins[3]),
            "xaxis": axis_style["xaxis"],
            "yaxis": axis_style["yaxis"],
            # 散点图模式不显示标题，同时主动清空可能从其他模式遗留的注释。
            "title": {"text": ""},
            "annotations": [],
            "showlegend": False,
            "plot_bgcolor": "white",
            "paper_bgcolor": "white",
            "hovermode": False,
            "dragmode": False,
        }
        fig.update_layout(go.Layout(**layout_dict))
        return {
            "plot_json": json.dumps(fig.to_dict(), cls=plotly.utils.PlotlyJSONEncoder), 
            "js_mode": "scatter", 
            "js_logic": "",
            "x_range": x_range,
            "y_range": y_range
        }

