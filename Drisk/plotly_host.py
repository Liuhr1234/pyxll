# plotly_host.py
# -*- coding: utf-8 -*-
"""
本模块提供 PlotlyHost（后端对接改造版 / 更稳健的宿主页）。

模块定位：
    本模块属于“第二组 UI 代码”，负责在 Qt 的 QWebEngineView 中承载 Plotly 图表页面，
    并打通 Python -> HTML/JavaScript -> Plotly -> Qt Overlay 之间的数据与状态同步链路。

核心职责：
1. 单页宿主架构：
   在 QWebEngineView 中仅加载一次固定的 host.html（内含 plotly.min.js），
   避免旧方案中频繁生成临时 HTML 文件造成的白屏、闪烁和性能损耗。

2. 动态渲染：
   后续所有绘图更新都通过 runJavaScript 下发 figure JSON，
   再由页面内部调用 Plotly.react / Plotly.newPlot 实现局部高性能刷新。

3. 遮罩联动：
   可选地与 Qt 原生遮罩层（Overlay）联动，
   自动探测 Plotly 页面内部真实绘图区矩形（plotRect），
   使原生阈值线、遮罩区与网页图表保持像素级对齐。

4. 诊断能力：
   提供前端异常捕获与读取接口，
   便于定位“图表白屏但 Python 侧无异常”的前端 JavaScript 问题。

本次改造重点：
- 自动绑定 QWebEngineView 的 loadFinished 信号，避免外部遗漏连接。
- 对图表模式（mode）做统一归一化，消除 hist / histogram 等别名差异。
- 动态获取 HTML 内部真实 plotRect，并同步给 Qt 原生 Overlay。
- JS 侧增加基础容错、错误捕获和调试能力。

说明：
- 本模块不直接负责“生成图表数据”，而是负责“承载并渲染图表”。
- 上游只需提供 Plotly figure JSON、图表模式、阈值位置与附加 JS 逻辑，
  本模块负责将其稳定显示到 QWebEngineView，并在需要时同步到 Qt 原生覆盖层。
"""

from __future__ import annotations

import json
import os
from collections import deque
from typing import Any, Optional, Tuple, Dict, Callable

from PySide6.QtCore import QUrl, QTimer
from plotly.offline import get_plotlyjs


class PlotlyHost:
    """
    Plotly 图表宿主管理器。

    类职责概述：
    1. 管理 plotly.min.js 的本地缓存。
    2. 生成并加载固定的宿主页 plotly_host.html。
    3. 将 Python 侧构建好的 Plotly 图表配置下发到网页端进行渲染。
    4. 在启用 Qt Overlay 时，同步网页绘图区矩形与原生遮罩层。
    5. 提供阈值线更新、错误读取等轻量交互与调试能力。

    与旧方案的差异：
    - 旧方案：每次绘图生成新的 HTML，再让 WebView 重新加载，开销较大。
    - 新方案：HTML 页面只初始化一次，后续通过 JS 注入 figure 数据并调用 Plotly.react 更新。
    """

    # ============================================================
    # 1. 初始化与基础工具
    # ============================================================
    def __init__(
        self,
        web_view,
        tmp_dir: str,
        use_qt_overlay: bool = False,
        overlay=None,
    ):
        """
        初始化 PlotlyHost。

        参数：
        - web_view:
            用于承载 Plotly 页面显示的 WebView，一般为 QWebEngineView。
        - tmp_dir:
            临时目录。用于写入 plotly.min.js 与固定宿主页 plotly_host.html。
        - use_qt_overlay:
            是否启用 Qt 原生 Overlay 模式。
            True 时，阈值线与遮罩主要由 Qt 原生层绘制；
            False 时，阈值线与遮罩主要由 HTML DOM 层绘制。
        - overlay:
            外部传入的原生遮罩对象。通常应支持 set_plot_rect / set_lr / set_margins 等接口。

        初始化阶段的主要工作：
        1. 记录外部依赖对象。
        2. 初始化 Plotly 资源与宿主页路径状态。
        3. 初始化旧架构残留字段（为兼容保留，但当前架构已不依赖）。
        4. 初始化“宿主页是否加载完成”“是否有待执行绘图请求”等运行状态。
        5. 自动尝试绑定 web_view.loadFinished 信号。
        6. 首次同步 overlay 的 DPR（屏幕缩放倍率）。
        """
        self.web_view = web_view
        self.tmp_dir = tmp_dir
        self.use_qt_overlay = bool(use_qt_overlay)
        self.overlay = overlay

        # ---- Plotly.js 本地缓存路径 ----
        # 第一次使用时会写入 tmp_dir/plotly.min.js，后续重复复用。
        self._plotly_js_path: Optional[str] = None

        # ---- 旧架构遗留状态（当前方案已基本不用，保留用于兼容或回溯） ----
        # 旧方案曾为每次绘图单独生成 HTML 文件，这些变量用于记录当前/待删/历史 HTML。
        self._current_html_path: Optional[str] = None
        self._pending_delete_html: Optional[str] = None
        self._html_history = deque(maxlen=32)

        # ---- 新架构核心状态：单宿主页 + JS 动态更新 ----
        # 固定宿主页路径：tmp_dir/plotly_host.html
        self._host_html_path: Optional[str] = None

        # 宿主页是否已加载完成。
        # 仅当 web_view.loadFinished(True) 触发后，才允许安全执行 runJavaScript。
        self._host_loaded: bool = False

        # 在宿主页尚未加载完成前，如果已经收到绘图请求，则先暂存最后一份 payload。
        # 待页面 ready 后自动执行，避免“页面未完成加载导致 runJavaScript 无效”。
        self._pending_payload: Optional[Dict[str, Any]] = None

        # ---- Overlay 同步节流标志 ----
        # 避免短时间连续触发 plotRect 同步，影响性能。
        self._overlay_sync_scheduled: bool = False

        # 自动绑定 loadFinished 回调，减少外部接线遗漏风险。
        try:
            self.web_view.loadFinished.connect(self.on_webview_load_finished)
        except Exception:
            # 兼容某些环境：web_view 不一定是标准 QWebEngineView。
            # 若自动绑定失败，允许外部在合适时机手动调用 on_webview_load_finished。
            pass

        # 初始化时先同步一次 Overlay 的屏幕缩放倍率。
        self._sync_overlay_dpr()

    @staticmethod
    def _normalize_mode(mode: str) -> str:
        """
        图表模式归一化工具。

        作用：
        将上层可能传入的简写、别名、非统一命名的模式字符串，
        统一转换为前端和本类内部约定的标准模式名。

        当前支持的主要归一化规则：
        - "hist" / "histogram" -> "histogram"
        - "disc" / "discrete"  -> "discrete"
        - "cdf"                -> "cdf"
        - "pdf"                -> "pdf"

        为什么需要归一化：
        - 上游调用者可能使用不同写法。
        - 前端遮罩、阈值线、显示逻辑依赖 mode 做分支判断。
        - 若 mode 命名不统一，容易出现图表正常但遮罩/线条逻辑失效的情况。

        返回：
        - 标准模式字符串。
        - 若传入为空，则默认回退为 "histogram"。
        """
        m = (mode or "").strip().lower()
        if m in ("hist", "histogram"):
            return "histogram"
        if m in ("disc", "discrete"):
            return "discrete"
        if m in ("cdf",):
            return "cdf"
        if m in ("pdf",):
            return "pdf"
        # 其他未定义模式维持原样，交由后续逻辑自行处理。
        return m or "histogram"

    def _sync_overlay_dpr(self):
        """
        同步 Overlay 的屏幕缩放倍率（DPR）。

        背景：
        在高分屏环境下，Qt 控件存在 devicePixelRatioF 缩放因素。
        若 Overlay 不知道当前 DPR，则原生绘制的线宽、位置等可能与 WebView 内部显示不一致。

        生效条件：
        - 已启用 Qt Overlay。
        - overlay 对象存在。
        - overlay 实现了 set_dpr 接口。

        设计原则：
        - 失败时静默跳过，不影响主链路。
        - 该方法是辅助同步，不应让绘图主流程因其报错而中断。
        """
        if not (self.use_qt_overlay and self.overlay is not None):
            return
        try:
            dpr = float(self.web_view.devicePixelRatioF())
            if hasattr(self.overlay, "set_dpr"):
                self.overlay.set_dpr(dpr)
        except Exception:
            pass

    # ============================================================
    # 2. 资源与宿主页管理
    # ============================================================
    def ensure_plotlyjs_cached(self) -> str:
        """
        确保 plotly.min.js 已缓存到本地临时目录，并返回该文件路径。

        执行逻辑：
        1. 若之前已经缓存过，且对应文件仍然存在，则直接返回。
        2. 若未缓存，则创建 tmp_dir。
        3. 若 tmp_dir/plotly.min.js 不存在，则从 plotly.offline.get_plotlyjs() 提取完整 JS 内容并写入本地。
        4. 记录缓存路径并返回。

        设计考虑：
        - 避免每次绘图都重复拼装/注入 Plotly JS。
        - 统一改为“本地固定 JS 文件 + 固定宿主页”的轻量结构。
        """
        if self._plotly_js_path and os.path.exists(self._plotly_js_path):
            return self._plotly_js_path

        os.makedirs(self.tmp_dir, exist_ok=True)
        js_path = os.path.join(self.tmp_dir, "plotly.min.js")
        if not os.path.exists(js_path):
            with open(js_path, "w", encoding="utf-8") as f:
                f.write(get_plotlyjs())

        self._plotly_js_path = js_path
        return js_path

    def _write_host_html(self, host_path: str, plotly_js_url: str) -> None:
        """
        写入固定宿主页 host.html。

        该 HTML 页面是整个宿主架构的核心：
        - 页面本身不固化具体图表数据；
        - 页面仅负责：
          1）加载 plotly.min.js；
          2）提供一个 chart_div 作为 Plotly 图表容器；
          3）提供阈值线与遮罩的 DOM 元素；
          4）暴露若干全局 JS 函数，供 Python 侧通过 runJavaScript 调用。

        关键对外 JS 接口：
        - window.__driskSetFigure(...)
            Python 侧每次绘图/刷新时调用的核心入口。
        - window.__qtGetPlotRect()
            用于将 Plotly 实际绘图区矩形返回给 Qt Overlay。
        - window.updateVLines(...) / window.updateChart(...)
            用于轻量更新阈值线和遮罩，而非整图重绘。

        关于 use_qt_overlay：
        - 该标志会直接在 HTML 中固化为布尔值。
        - 这样 JS 运行时无需反复判断 Python 侧参数，简化前端分支逻辑。
        """
        html = f"""<!DOCTYPE html>
<html>
<head>
  <meta charset="utf-8"/>
  <meta name="viewport" content="width=device-width, initial-scale=1.0"/>
  <style>
    /* ================================
       页面基础样式与全屏容器设置
       ================================ */

    /* 页面整体去边距，并让 html/body 铺满可视区域 */
    html, body {{
      margin:0; padding:0; width:100%; height:100%;
      overflow:hidden; font-family:'Arial','Microsoft YaHei','Segoe UI';
    }}

    /* 图表外层容器：作为 Plotly 图表与遮罩 DOM 元素的共同父容器 */
    #chart_wrap {{
      position: relative;
      width: 100%;
      height: 100%;
    }}

    /* ================================
       DOM 模式下的阈值竖线样式
       仅在未启用 Qt Overlay 时使用
       ================================ */
    .drisk-vline {{
      position:absolute;
      top:0; height:0; left:0; width:0;
      border-left:1.6px dashed rgba(0,0,0,0.85);
      pointer-events:none; z-index:50; display:none;
    }}

    /* ================================
       DOM 模式下的两侧遮罩样式
       仅在 histogram / discrete 等模式下启用
       ================================ */
    .drisk-mask {{
      position:absolute;
      top:0; height:0; left:0; width:0;
      background: rgba(255,255,255,0.7);
      pointer-events:none; z-index:20; display:none;
    }}
  </style>
  <script src="{plotly_js_url}"></script>
</head>
<body>
  <div id="chart_wrap">
    <!-- Plotly 图表挂载容器 -->
    <div id="chart_div" style="width:100%;height:100%;"></div>

    <!-- 左右遮罩层与左右阈值线（仅 DOM 模式下生效） -->
    <div id="mask_left" class="drisk-mask"></div>
    <div id="mask_right" class="drisk-mask"></div>
    <div id="vline_left" class="drisk-vline"></div>
    <div id="vline_right" class="drisk-vline"></div>
  </div>

  <script>
    // ==========================================================
    // 1. 全局错误捕获与诊断状态
    // ==========================================================

    // 保存最近一次未处理前端错误，供 Python 侧主动读取。
    window.__drisk_last_error = null;

    // 捕获同步错误（例如脚本运行时报错）。
    window.addEventListener('error', function(e){{
      try{{
        var msg = (e && e.message) ? e.message : String(e);
        window.__drisk_last_error = msg;
        console.error("[drisk][window.error]", msg);
      }}catch(_){{}}
    }});

    // 捕获未处理 Promise 异常。
    window.addEventListener('unhandledrejection', function(e){{
      try{{
        var msg = (e && e.reason) ? String(e.reason) : String(e);
        window.__drisk_last_error = msg;
        console.error("[drisk][unhandledrejection]", msg);
      }}catch(_){{}}
    }});

    // ==========================================================
    // 2. 页面运行时状态
    // ==========================================================

    // 当前图表对象（Python 侧下发后解析得到）。
    var figure = null;

    // 当前图表模式。默认直方图。
    var mode = "histogram";

    // 是否启用 Qt 原生 Overlay。
    // true：遮罩/线主要由 Qt 原生层处理；
    // false：遮罩/线主要由 HTML DOM 层处理。
    var __use_qt_overlay = {str(self.use_qt_overlay).lower()};

    // 图表是否已至少完成过一次成功绘制。
    var __drisk_ready = false;

    // Plotly 内部事件是否已绑定，避免重复绑定。
    var __events_bound = false;

    // requestAnimationFrame 节流期间，待处理的最新阈值请求。
    var __pending = null;

    // 当前 requestAnimationFrame 句柄。
    var __raf = 0;

    // 最近一次左右阈值位置 [l, r]。
    var __last_lr = null;

    // 记录最近一次绘图完成时间戳，便于调试或状态判断。
    window.__drisk_last_draw_ts = 0;

    // ==========================================================
    // 3. Plotly 内部绘图区尺寸工具
    // ==========================================================

    // 获取 Plotly 图表内部真实绘图区尺寸（不含外边距）。
    function __getPlotSize(gd){{
      if(!gd || !gd._fullLayout || !gd._fullLayout._size) return null;
      return gd._fullLayout._size;
    }}

    // ==========================================================
    // 4. 供 Qt Overlay 调用：返回绘图区矩形
    // ==========================================================

    window.__qtGetPlotRect = function(){{
        try{{
            var gd = document.getElementById('chart_div');
            if(!gd) return null;
            var sz = __getPlotSize(gd);
            if(!sz) return null;
            return {{l: sz.l, t: sz.t, w: sz.w, h: sz.h}};
        }}catch(e){{
            return null;
        }}
    }};

    // ==========================================================
    // 5. DOM 模式：阈值线与遮罩更新工具
    // ==========================================================

    // 设置单条竖线位置。
    // x 为数据坐标，不是像素坐标；需借助 Plotly 的 xaxis.l2p 转换。
    function __setVLine(gd, el, x){{
      if(!gd || !el) return;
      var fl = gd._fullLayout;
      if(!fl || !fl.xaxis) return;
      var sz = __getPlotSize(gd);
      if(!sz) return;

      // l2p: logical/data -> pixel
      var px = sz.l + fl.xaxis.l2p(x);

      el.style.left = px + "px";
      el.style.top = sz.t + "px";
      el.style.height = sz.h + "px";
      el.style.display = "block";
    }}

    // 更新 DOM 阈值线。
    // CDF 模式下通常只显示左侧线；其余模式通常显示双侧线。
    function __updateDOMLines(l, r){{
      if(__use_qt_overlay) return;
      var gd = document.getElementById("chart_div");
      if(!gd) return;

      var elL = document.getElementById("vline_left");
      var elR = document.getElementById("vline_right");

      if(mode === "cdf"){{
        __setVLine(gd, elL, l);
        if(elR) elR.style.display = "none";
      }} else {{
        __setVLine(gd, elL, l);
        __setVLine(gd, elR, r);
        if(elR) elR.style.display = "block";
      }}
    }}

    // 更新 DOM 遮罩层。
    // 仅对 histogram / discrete 模式启用遮罩，其余模式隐藏。
    function __updateMasks(l, r){{
      if(__use_qt_overlay) return;

      if(!(mode === "histogram" || mode === "discrete")){{
        var ml0 = document.getElementById("mask_left");
        var mr0 = document.getElementById("mask_right");
        if(ml0) ml0.style.display = "none";
        if(mr0) mr0.style.display = "none";
        return;
      }}

      var gd = document.getElementById("chart_div");
      if(!gd) return;

      var fl = gd._fullLayout;
      if(!fl || !fl.xaxis) return;

      var sz = __getPlotSize(gd);
      if(!sz) return;

      var ml = document.getElementById("mask_left");
      var mr = document.getElementById("mask_right");
      if(!ml || !mr) return;

      var plotL = sz.l;
      var plotR = sz.l + sz.w;
      var plotT = sz.t;
      var plotH = sz.h;

      // 将左右阈值从数据坐标转换为像素坐标
      var pxL = plotL + fl.xaxis.l2p(l);
      var pxR = plotL + fl.xaxis.l2p(r);

      // 取有序区间 [a, b]
      var a = Math.min(pxL, pxR);
      var b = Math.max(pxL, pxR);

      // 裁剪到绘图区范围之内，避免越界
      a = Math.max(plotL, Math.min(a, plotR));
      b = Math.max(plotL, Math.min(b, plotR));

      // 左侧遮罩：绘图区左边界 -> a
      ml.style.left = plotL + "px";
      ml.style.top = plotT + "px";
      ml.style.height = plotH + "px";
      ml.style.width = Math.max(0, a - plotL) + "px";
      ml.style.display = (a > plotL) ? "block" : "none";

      // 右侧遮罩：b -> 绘图区右边界
      mr.style.left = b + "px";
      mr.style.top = plotT + "px";
      mr.style.height = plotH + "px";
      mr.style.width = Math.max(0, plotR - b) + "px";
      mr.style.display = (b < plotR) ? "block" : "none";
    }}

    // ==========================================================
    // 6. 阈值更新节流：使用 requestAnimationFrame
    // ==========================================================

    function __requestUpdate(l, r){{
      __last_lr = [l, r];
      __pending = {{l:l, r:r}};
      if(__raf) return;

      __raf = requestAnimationFrame(function(){{
        __raf = 0;
        var p = __pending;
        __pending = null;
        if(!p) return;

        __updateDOMLines(p.l, p.r);
        __updateMasks(p.l, p.r);
      }});
    }}

    // 外部调用：仅刷新阈值线与遮罩，不触发整图 Plotly 重绘。
    window.updateVLines = function(l, r){{
      if(!__drisk_ready){{
        __last_lr = [l, r];
        return;
      }}
      __requestUpdate(l, r);
    }};

    // 兼容命名：保留 updateChart 作为 updateVLines 的别名。
    window.updateChart = function(l, r){{
      window.updateVLines(l, r);
    }};

    // ==========================================================
    // 7. 核心整图更新入口：由 Python 调用
    // ==========================================================
    // 参数说明：
    // - fig_json_str : Plotly 图表 JSON 字符串
    // - new_mode     : 图表模式
    // - staticPlot   : 是否关闭原生交互
    // - lr           : 左右阈值 [l, r]，或 null
    // - js_logic_str : 图表初始化完成后需要补充执行的 JS 逻辑字符串
    window.__driskSetFigure = function(fig_json_str, new_mode, staticPlot, lr, js_logic_str){{
      try{{
        figure = JSON.parse(fig_json_str);
        if(new_mode) mode = new_mode;

        // 构建“初始化后附加逻辑”函数。
        // 这里通常用于修补局部前端细节，而不污染通用宿主页逻辑。
        window.__drisk_post_init = null;
        if(js_logic_str && js_logic_str.trim()){{
          try {{
            window.__drisk_post_init = new Function(js_logic_str);
          }} catch(e) {{
            window.__drisk_post_init = null;
          }}
        }}

        var gd = document.getElementById('chart_div');
        if(typeof Plotly === 'undefined'){{
          window.__drisk_last_error = "Plotly 未定义（plotly.min.js 未成功加载）";
          console.error("[drisk] Plotly is undefined (plotly.min.js not loaded)");
          return;
        }}
        if(!gd || !figure) return;

        var config = {{ responsive: true, displayModeBar: false, staticPlot: !!staticPlot }};

        // 首次绘图使用 newPlot；
        // 后续刷新使用 react，减少整页销毁重建的开销。
        var p;
        if(!__drisk_ready){{
          p = Plotly.newPlot(gd, figure.data, figure.layout, config);
        }} else {{
          p = Plotly.react(gd, figure.data, figure.layout, config);
        }}

        // 绘制完成后的回调
        p.then(function(){{
          __drisk_ready = true;
          window.__drisk_last_draw_ts = Date.now();

          // 初始绘制完成后，同步阈值线与遮罩
          if(lr && lr.length === 2){{
            __requestUpdate(lr[0], lr[1]);
          }}

          // 执行外部附加逻辑
          try {{
            if(typeof window.__drisk_post_init === 'function') window.__drisk_post_init();
          }}catch(e){{
            try{{
              var msg = (e && e.stack) ? String(e.stack) : String(e);
              window.__drisk_last_error = msg;
              console.error("[drisk][__driskSetFigure exception]", msg);
            }}catch(_){{}}
          }}

          // 首次绘图后绑定 Plotly 内部事件。
          // 目的：当图表缩放、平移、重算布局后，阈值线与遮罩仍可跟随最新绘图区。
          if(!__events_bound){{
            __events_bound = true;
            gd.on('plotly_relayout', function(){{
              if(__last_lr) __requestUpdate(__last_lr[0], __last_lr[1]);
            }});
            gd.on('plotly_afterplot', function(){{
              if(__last_lr) __requestUpdate(__last_lr[0], __last_lr[1]);
            }});
          }}
        }});
      }}catch(e){{
        // 这里保持静默。
        // 若需要更强诊断，可后续补充 window.__drisk_last_error 赋值逻辑。
      }}
    }};
  </script>
</body>
</html>"""

        os.makedirs(self.tmp_dir, exist_ok=True)
        with open(host_path, "w", encoding="utf-8") as f:
            f.write(html)

    def ensure_host_page_loaded(self) -> None:
        """
        确保固定宿主页已生成并发起加载。

        执行逻辑：
        1. 先确保 plotly.min.js 已缓存。
        2. 决定 host.html 的固定路径。
        3. 若 host.html 不存在或文件异常过小，则重新生成。
        4. 若当前尚未加载宿主页，则向 web_view 发起 load 请求。

        关键约束：
        - 本方法只负责“保证页面被加入加载流程”。
        - 不保证调用结束时页面已 ready。
        - 页面是否真正可执行 JS，应以 self._host_loaded 为准。
        """
        plotly_js_path = self.ensure_plotlyjs_cached()
        plotly_js_url = QUrl.fromLocalFile(plotly_js_path).toString()

        os.makedirs(self.tmp_dir, exist_ok=True)
        if self._host_html_path is None:
            self._host_html_path = os.path.join(self.tmp_dir, "plotly_host.html")

        # 文件不存在，或过小到明显异常时，重建宿主页。
        if (not os.path.exists(self._host_html_path)) or os.path.getsize(self._host_html_path) < 1024:
            self._write_host_html(self._host_html_path, plotly_js_url)

        # 若还未完成加载，则发起页面加载。
        if not self._host_loaded:
            self.web_view.load(QUrl.fromLocalFile(self._host_html_path))

    # ============================================================
    # 3. 数据下发与 Qt 遮罩同步
    # ============================================================
    def _apply_payload(self, payload: Dict[str, Any]) -> None:
        """
        将 Python 侧构造好的绘图数据包下发给网页端执行渲染。

        payload 约定字段：
        - fig_json:
            Plotly 图表 JSON 字符串。
        - mode:
            图表模式，如 histogram / cdf / pdf / discrete。
        - static_plot:
            是否禁用 Plotly 原生交互。
        - lr:
            左右阈值元组，或 None。
        - js_logic:
            图表初始化完成后的附加 JS 逻辑字符串。

        执行流程：
        1. 提取并规范化 payload 中的字段。
        2. 使用 json.dumps 对字符串参数再包一层，确保进入 JS 字符串字面量时不因引号、换行等报错。
        3. 组装为一段对 window.__driskSetFigure(...) 的 JS 调用。
        4. 使用 runJavaScript 下发到当前页面。
        5. 下发后，异步安排一次 Overlay 与 plotRect 的同步。

        注意：
        - 本方法假设宿主页已 ready。
        - 若页面未 ready，应由外部先缓存 payload，而不是直接调用本方法。
        """
        fig_json = payload.get("fig_json", "")
        mode = self._normalize_mode(payload.get("mode", "histogram"))
        static_plot = bool(payload.get("static_plot", True))
        lr = payload.get("lr", None)
        js_logic = payload.get("js_logic", "") or ""

        # 对字符串做二次 JSON 安全封装，避免直接拼接进 JS 时破坏语法。
        fig_json_js = json.dumps(fig_json)
        mode_js = json.dumps(mode)
        js_logic_js = json.dumps(js_logic)
        static_js = "true" if static_plot else "false"

        # 处理左右阈值参数
        if lr is None:
            lr_js = "null"
        else:
            try:
                l = float(lr[0])
                r = float(lr[1])
            except Exception:
                l, r = 0.0, 0.0
            lr_js = f"[{l}, {r}]"

        js = f"window.__driskSetFigure({fig_json_js}, {mode_js}, {static_js}, {lr_js}, {js_logic_js});"

        try:
            self.web_view.page().runJavaScript(js)
        except Exception:
            return

        # 图表请求下发后，再异步同步一次网页绘图区与 Qt Overlay 的位置关系。
        self._schedule_overlay_plot_rect_sync()

    def _schedule_overlay_plot_rect_sync(self) -> None:
        """
        以节流方式安排一次 Overlay 绘图区矩形同步。

        目的：
        Plotly.react / newPlot 完成布局计算需要一定时间。
        若立即读取 plotRect，可能拿到旧值或空值。
        因此这里采用“延迟 + 防重复调度”的方式，稍后再执行一次同步。

        节流规则：
        - 仅在启用 Qt Overlay 且 overlay 存在时生效。
        - 若当前已安排过一次同步，则不重复安排。
        - 使用 QTimer.singleShot(60ms) 延后执行。

        60ms 的含义：
        - 给 Plotly 一段较短但通常足够的布局计算时间。
        - 不追求绝对即时，而追求稳定与低抖动。
        """
        if not (self.use_qt_overlay and self.overlay is not None):
            return
        if self._overlay_sync_scheduled:
            return
        self._overlay_sync_scheduled = True

        QTimer.singleShot(60, self._sync_overlay_plot_rect_once)

    def _sync_overlay_plot_rect_once(self) -> None:
        """
        单次执行：同步网页绘图区矩形到 Qt Overlay。

        执行步骤：
        1. 取消“已安排”标记。
        2. 检查启用条件：Overlay 存在、宿主页已加载完成。
        3. 让 Overlay 的几何尺寸先对齐整个 WebView，并同步当前 DPR。
        4. 通过 JS 调用 window.__qtGetPlotRect() 获取网页内部真实绘图区矩形。
        5. 将矩形写入 overlay.set_plot_rect(...)。
        6. 请求 Overlay 重绘。

        为什么要区分“整个 WebView rect”和“内部 plotRect”：
        - WebView 是整个网页区域；
        - Plotly 真正绘图区域会受到标题、边距、坐标轴、图例等影响；
        - 阈值线和遮罩必须贴合 plotRect，而非简单贴合整个 WebView。

        注意：
        - Qt 的 runJavaScript 回调运行在主线程。
        - 这里不能做重计算，只做轻量赋值与刷新。
        """
        self._overlay_sync_scheduled = False

        if not (self.use_qt_overlay and self.overlay is not None):
            return
        if not self._host_loaded:
            return

        try:
            self.overlay.setGeometry(self.web_view.rect())
            self._sync_overlay_dpr()
        except Exception:
            pass

        js = "window.__qtGetPlotRect && window.__qtGetPlotRect();"

        def _on_rect(rect):
            """
            JS 返回 plotRect 后的回调。

            预期 rect 结构：
            {{
                "l": 左偏移,
                "t": 上偏移,
                "w": 绘图区宽,
                "h": 绘图区高
            }}

            处理原则：
            - 仅做安全解析与 Overlay 更新；
            - 若数据异常，则静默跳过，不影响主流程。
            """
            try:
                if rect and isinstance(rect, dict):
                    l = float(rect.get("l", 0.0))
                    t = float(rect.get("t", 0.0))
                    w = float(rect.get("w", 0.0))
                    h = float(rect.get("h", 0.0))
                    if hasattr(self.overlay, "set_plot_rect"):
                        self.overlay.set_plot_rect(l, t, w, h)
                    self.overlay.update()
            except Exception:
                pass

        try:
            self.web_view.page().runJavaScript(js, _on_rect)
        except Exception:
            pass

    # ============================================================
    # 4. 主入口与生命周期
    # ============================================================
    def load_plot(
        self,
        plot_json: Any,
        js_mode: str,
        js_logic: str = "",
        initial_lr: Optional[Tuple[float, float]] = None,
        static_plot: bool = True,
    ):
        """
        主渲染入口：加载并显示一张 Plotly 图表。

        参数：
        - plot_json:
            Plotly 图表配置。
            支持两种形式：
            1）Figure.to_json() 产生的 JSON 字符串；
            2）Python 字典格式的 figure。
        - js_mode:
            图表模式，如：
            "histogram"、"cdf"、"pdf"、"discrete" 等。
        - js_logic:
            图表首次绘制完成后，额外执行的前端 JS 逻辑。
            常用于局部补丁或特殊显示修正。
        - initial_lr:
            初始左右阈值。
            若不为 None，则图表绘制完成后会立刻同步阈值线/遮罩位置。
        - static_plot:
            是否关闭 Plotly 原生交互。
            True：更像静态报表。
            False：允许拖拽、缩放等原生交互。

        方法内主要分组：
        A. 规范化输入 figure 数据。
        B. 注入统一视觉背景配置。
        C. 从图表 JSON 中嗅探 margin，并同步给 Overlay。
        D. 构造 payload。
        E. 清空上一张图残留的 Overlay 状态。
        F. 确保宿主页已进入加载流程。
        G. 若页面未 ready，则缓存 payload；若已 ready，则直接应用。
        """
        # ---------------------------------------------------------
        # A. 规范化 plot_json 输入格式
        # ---------------------------------------------------------
        if isinstance(plot_json, str):
            try:
                plot_dict = json.loads(plot_json)
            except Exception:
                plot_dict = {}
        else:
            plot_dict = plot_json

        # ---------------------------------------------------------
        # B. 注入统一视觉背景
        # ---------------------------------------------------------
        # 约定：
        # - plot_bgcolor：绘图区内部浅灰
        # - paper_bgcolor：外部画布白色
        # 这样可以保证图表主体与外围背景层次清晰。
        if "layout" not in plot_dict:
            plot_dict["layout"] = {}
        plot_dict["layout"]["plot_bgcolor"] = "#fcfcfc"

        if "paper_bgcolor" not in plot_dict["layout"]:
            plot_dict["layout"]["paper_bgcolor"] = "#ffffff"

        # ---------------------------------------------------------
        # C. 从图表 JSON 中同步 margin 到 Overlay
        # ---------------------------------------------------------
        # 目的：
        # 某些图表标题、说明、顶部元素会动态出现或消失，导致 Plotly 实际 margin 变化。
        # 如果 Overlay 仍沿用旧 margin，会出现遮罩错位、图下方空白等问题。
        if self.use_qt_overlay and self.overlay is not None:
            layout_margin = plot_dict["layout"].get("margin", {})
            ml = float(layout_margin.get("l", self.overlay.margin_l))
            mr = float(layout_margin.get("r", self.overlay.margin_r))
            mt = float(layout_margin.get("t", self.overlay.margin_t))
            mb = float(layout_margin.get("b", self.overlay.margin_b))

            if hasattr(self.overlay, "set_margins"):
                self.overlay.set_margins(ml, mr, mt, mb)

        # ---------------------------------------------------------
        # D. 转换为最终要传给前端的 payload
        # ---------------------------------------------------------
        plot_json_str = json.dumps(plot_dict)

        payload = {
            "fig_json": plot_json_str,
            "mode": self._normalize_mode(js_mode),
            "static_plot": bool(static_plot),
            "lr": None if initial_lr is None else (float(initial_lr[0]), float(initial_lr[1])),
            "js_logic": js_logic,
        }

        # ---------------------------------------------------------
        # E. 预清空 Overlay 状态，避免图表切换时残留旧遮罩
        # ---------------------------------------------------------
        if self.use_qt_overlay and self.overlay is not None:
            try:
                if hasattr(self.overlay, "clear_plot_rect"):
                    self.overlay.clear_plot_rect()
                self.overlay.setGeometry(self.web_view.rect())
                self.overlay.update()
            except Exception:
                pass

        # ---------------------------------------------------------
        # F. 确保宿主页已生成并发起加载
        # ---------------------------------------------------------
        self.ensure_host_page_loaded()

        # ---------------------------------------------------------
        # G. 根据页面状态决定“缓存”还是“立即执行”
        # ---------------------------------------------------------
        if not self._host_loaded:
            # 页面还没 ready，先缓存这次最新请求。
            # 只保留最后一份，避免页面 ready 后重复渲染过期图。
            self._pending_payload = payload
            return

        # 页面已 ready，立即应用 payload。
        self._apply_payload(payload)

    def on_webview_load_finished(self, ok: bool):
        """
        WebView 页面加载完成回调。

        调用来源：
        - 一般由 self.web_view.loadFinished 自动触发。
        - 若外部环境不支持自动绑定，也可手动调用。

        参数：
        - ok:
            页面是否成功加载。

        主要职责：
        1. 更新 self._host_loaded 状态。
        2. 若启用 Overlay，则初始化其几何、显示层级与可见性。
        3. 若此前有挂起的 payload，则在此时执行补发。
        4. 页面 ready 后主动发起一次 plotRect 同步。

        典型场景：
        - 程序首次启动后第一次绘图。
        - 调用了 load_plot，但 host.html 还没加载完，此时 payload 被挂起。
        - 页面加载完成后，由这里兜底执行挂起图表。
        """
        self._host_loaded = bool(ok)

        # 页面就绪后，先让 Overlay 就位并置顶。
        if self.use_qt_overlay and self.overlay is not None:
            try:
                self.overlay.setGeometry(self.web_view.rect())
                self.overlay.show()
                self.overlay.raise_()
                self.overlay.update()
            except Exception:
                pass

        # 若之前已经收到绘图请求，但页面尚未 ready，则此时执行。
        if self._host_loaded and self._pending_payload is not None:
            payload = self._pending_payload
            self._pending_payload = None
            self._apply_payload(payload)

        # 页面加载完成后，再主动安排一次 plotRect 同步，确保 Overlay 与网页对齐。
        self._schedule_overlay_plot_rect_sync()

    # ============================================================
    # 5. 交互更新与诊断辅助
    # ============================================================
    def update_lr(self, left: float, right: float):
        """
        轻量更新左右阈值线与遮罩位置。

        与 load_plot 的区别：
        - load_plot：整图数据更新，可能触发 Plotly.react / newPlot。
        - update_lr：只更新左右阈值位置，不刷新图表数据本身。

        适用场景：
        - 用户拖动滑块、输入边界值时，需要实时移动阈值线。
        - 不希望因为只是更新边界位置，就重新绘制整张图。

        实现方式：
        - 对于纯 DOM 模式：通过 JS 调用页面内 updateChart(l, r)。
        - 对于 Qt Overlay 模式：理论上也可由外部 overlay 直接处理；
          这里仍保留 JS 侧调用能力，便于兼容非 Overlay 方案。

        设计原则：
        - 入参先做 float 安全转换。
        - 出错时静默，不打断主界面交互。
        """
        try:
            l = float(left)
            r = float(right)
        except Exception:
            l, r = 0.0, 0.0

        try:
            self.web_view.page().runJavaScript(f"updateChart({l}, {r});")
        except Exception:
            pass

    def get_last_error(self, callback: Callable[[Optional[str]], None]):
        """
        获取前端页面最近一次未处理异常信息。

        作用：
        当界面出现“图表区域空白、但 Python 没有直接报错”时，
        可通过本方法读取前端页面内部捕获到的错误信息，辅助定位问题。

        参数：
        - callback:
            回调函数，形如 callback(msg)。
            msg 为字符串或 None。

        行为：
        - 若宿主页尚未加载完成，则直接 callback(None)。
        - 若已加载，则通过 runJavaScript 读取 window.__drisk_last_error。

        示例：
            host.get_last_error(lambda msg: print("前端最后错误:", msg))
        """
        if not self._host_loaded:
            callback(None)
            return
        js = "window.__drisk_last_error || null;"
        try:
            self.web_view.page().runJavaScript(js, callback)
        except Exception:
            callback(None)

    def request_overlay_plot_rect_sync(self):
        """
        手动请求一次 Overlay 与 Plotly 绘图区的矩形同步。

        典型调用场景：
        - 主窗口尺寸变化后；
        - WebView 尺寸变化后；
        - 某些图表布局元素显隐导致绘图区实际位置变化后。

        说明：
        - 本方法本身不立即执行，而是进入已有的节流调度流程。
        - 适合作为外部 Resize Event 或重排事件中的统一接口。
        """
        self._schedule_overlay_plot_rect_sync()
    