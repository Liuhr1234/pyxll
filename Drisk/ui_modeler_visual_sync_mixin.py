# ui_modeler_visual_sync_mixin.py
"""
本模块提供分布构建器（DistributionBuilderDialog）渲染后的图表几何参数与视觉层同步辅助功能（Mixin）。

主要功能：
1. 视图范围锁定 (View Range Enforcement)：强制限制 Plotly 的 X/Y 轴缩放与平移，保持前端图表与外部 Qt UI 控制逻辑（如滑块范围）的一致性。
2. 几何坐标同步 (Geometry Synchronization)：通过 JavaScript 注入，获取 Webview 内实际数据绘图区 (Plot Area) 的物理像素尺寸与边距。
3. 跨层对齐 (Cross-layer Alignment)：利用获取到的真实图表物理坐标，精确调整外部 Qt 组件（如双向滑块、标题栏、悬浮遮罩、骨架屏）的位置和尺寸，使其与内嵌网页图表完全重合。
4. 渲染生命周期与防抖 (Render Lifecycle & Debounce)：处理异步渲染时序，通过多重定时器交错执行，确保图表重绘时的布局最终稳定，防止 UI 组件闪烁或错位。
"""

import drisk_env
import math

from PySide6.QtCore import QTimer
from PySide6.QtWidgets import QMessageBox


class DistributionBuilderVisualSyncMixin:
    """
    [混合类] 渲染后图表几何坐标提取与视觉同步辅助工具。
    设计目的：解决前端图表（Webview/Plotly）与后端原生 UI（PySide6/Qt）跨层叠加时的坐标对齐问题。
    为提升代码交接后的可维护性，此处对跨环境（Python -> JS -> Python）的异步通信与时序控制逻辑进行了详细标注。
    """

    # =======================================================
    # 1. 视图范围强制锁定 (View Range Enforcement)
    # =======================================================
    def _enforce_plotly_view_xrange(self):
        """
        强制锁定 Plotly 的物理视图范围：
        通过向 Webview 注入并执行 JavaScript 代码，直接修改 Plotly 实例的 layout 属性。
        此举旨在禁用 Plotly 内置的平移/缩放交互 (fixedrange=true)，并强制将其 X 轴范围对齐到底层数学模型的物理边界 (view_min 和 view_max)。
        """
        # 校验 Webview 实例是否存在，防止在未初始化或已销毁时调用报错
        if not getattr(self, "web", None):
            return
        
        try:
            # 获取底层数学模型设定的逻辑显示边界
            xmin = float(getattr(self, "view_min", 0.0))
            xmax = float(getattr(self, "view_max", 1.0))
            
            # 防御性校验：确保边界数值合法（非无穷大、非 NaN，且最大值必须大于最小值）
            if not math.isfinite(xmin) or not math.isfinite(xmax) or xmax <= xmin:
                return
        except Exception:
            return

        # 检测当前图表是否启用了双 Y 轴（例如在直方图上叠加了 CDF 曲线）
        # 如果存在，则需要同步锁定右侧 Y 轴 (yaxis2) 的交互
        has_y2 = bool(getattr(self, "_has_y2", False))
        y2_line = "update['yaxis2.fixedrange']=true;" if has_y2 else ""
        
        # 组装待注入的 JavaScript 更新脚本
        # 逻辑说明：
        # 1. 获取名为 'chart_div' 的 DOM 节点
        # 2. 构建 Plotly.relayout 所需的 update 字典
        # 3. 强制覆盖 xaxis.range，并关闭自动调整 (autorange)、拖拽缩放 (fixedrange) 和范围滑块 (rangeslider)
        js = f"""
        (function(){{
            try{{
                var gd = document.getElementById('chart_div');
                if(!gd || !window.Plotly) return false;
                var update = {{}};
                
                // 锁定 X 轴的显示范围，与 Python 端的模型边界保持一致
                update['xaxis.range'] = [{xmin}, {xmax}];
                update['xaxis.autorange'] = false;
                update['xaxis.fixedrange'] = true;
                update['xaxis.rangeslider.visible'] = false;

                // 锁定主 Y 轴的缩放与平移
                update['yaxis.fixedrange'] = true;
                {y2_line}

                // 调用 Plotly API 轻量级更新布局
                Plotly.relayout(gd, update);
                return true;
            }}catch(e){{ return false; }}
        }})();
        """
        
        try:
            # 向 Webview 异步注入并执行 JS 脚本
            self.web.page().runJavaScript(js)
        except Exception:
            pass

    # =======================================================
    # 2. 跨层几何坐标同步与 UI 对齐 (Geometry Synchronization)
    # =======================================================
    def _sync_plot_geom_from_plotly(self):
        """
        核心几何同步引擎：
        向 Webview 注入 JS 脚本，深挖 Plotly 的内部私有对象 (_fullLayout)，
        提取出实际数据绘图区 (Plot Area) 的精确物理像素偏移量、尺寸以及设备像素比 (DPR)。
        随后通过异步回调函数将这些几何数据传回 Python 端，用于驱动外部 Qt 控件（滑块、标题、遮罩）的重绘与对齐。
        """
        if not getattr(self, "web", None):
            return

        # 在提取坐标前，先确保视图范围已被正确锁定，以获取最准确的物理尺寸
        self._enforce_plotly_view_xrange()

        # JS 脚本：提取绘图区域的物理尺寸与边距
        # 注意：此处读取的是 Plotly 渲染后生成的内部属性 (_offset, _length 等)
        js = r"""
        (function(){
            try{
                var gd = document.getElementById('chart_div');
                // 确保图表已完成渲染且存在 _fullLayout 内部属性
                if(!gd || !gd._fullLayout) return null;
                var fl = gd._fullLayout;

                var xa = fl.xaxis;
                var ya = fl.yaxis;
                if(!xa || !ya) return null;

                // 计算内部实际网格（绘图区）相对于容器的物理坐标与尺寸
                var plotL = (xa._offset != null) ? xa._offset : (fl._size ? fl._size.l : 0);
                var plotW = (xa._length != null) ? xa._length : (fl._size ? fl._size.w : 0);
                var plotT = (ya._offset != null) ? ya._offset : (fl._size ? fl._size.t : 0);
                var plotH = (ya._length != null) ? ya._length : (fl._size ? fl._size.h : 0);

                // 获取图表容器的整体宽高
                var fullW = (fl.width != null) ? fl.width : gd.clientWidth;
                var fullH = (fl.height != null) ? fl.height : gd.clientHeight;

                // 计算四周的边距 (Left, Right, Top, Bottom)
                var l = plotL;
                var r = Math.max(0, fullW - (plotL + plotW));
                var t = plotT;
                var b = Math.max(0, fullH - (plotT + plotH));

                // 将几何数据与设备像素比 (DPR) 序列化返回给 Python 环境
                return {
                    plotL: plotL, plotT: plotT, plotW: plotW, plotH: plotH,
                    l: l, r: r, t: t, b: b,
                    dpr: (window.devicePixelRatio || 1.0)
                };
            }catch(e){ return null; }
        })();
        """

        # JS 异步执行完毕后的 Python 端回调函数
        def _cb(res):
            # 校验回传数据结构是否合法
            if not res or not isinstance(res, dict):
                return
            try:
                # 解析绘图区的绝对物理位置与尺寸
                plotL = float(res.get("plotL", 0.0))
                plotT = float(res.get("plotT", 0.0))
                plotW = float(res.get("plotW", 0.0))
                plotH = float(res.get("plotH", 0.0))

                # ---------------------------------------------------
                # 步骤 A：同步底部双向滑块 (Slider) 的宽度
                # 将真实的绘图区宽度反馈给滑块，确保滑块的控制句柄与图表 X 轴的数据点能产生视觉上的精准垂直对齐。
                # ---------------------------------------------------
                if getattr(self, "slider", None) is not None:
                    if hasattr(self.slider, "setPlotWidth"):
                        self.slider.setPlotWidth(plotW)

                # 解析绘图区的四周边距及系统缩放比例
                l = float(res.get("l", plotL))
                r = float(res.get("r", 0.0))
                t = float(res.get("t", plotT))
                b = float(res.get("b", 0.0))
                dpr = float(res.get("dpr", 1.0))

                # 获取当前逻辑显示的极值边界
                self._display_xmin = getattr(self, "view_min", 0.0)
                self._display_xmax = getattr(self, "view_max", 1.0)

                # ---------------------------------------------------
                # 步骤 B：同步滑块的物理边距与逻辑边界
                # ---------------------------------------------------
                if getattr(self, "slider", None) is not None:
                    # 确保滑块的两侧留白区域与 Web 图表的左右边距等宽
                    self.slider.setMargins(int(round(l)), int(round(r)))
                    # 设定滑块组件内部的数值映射区间
                    self.slider.setRangeLimit(self._display_xmin, self._display_xmax)

                # ---------------------------------------------------
                # 步骤 C：同步顶部标题容器，并钳制滑块越界值
                # 保持顶层标题容器与下方实际绘图矩形区域在水平方向上的对齐。
                # ---------------------------------------------------
                if hasattr(self, "title_container_layout"):
                    self.title_container_layout.setContentsMargins(int(round(l)), 0, int(round(r)), 0)

                    # 对齐的同时，执行数值钳制 (Clamp) 逻辑：防止滑块当前的左右控制柄游标超出新计算的边界范围
                    try:
                        # 限制左侧游标不越界
                        self.curr_left = max(self._display_xmin, min(float(getattr(self, "curr_left", self._display_xmin)), self._display_xmax))
                        # 限制右侧游标不越界
                        self.curr_right = max(self._display_xmin, min(float(getattr(self, "curr_right", self._display_xmax)), self._display_xmax))
                        
                        # 保证数学逻辑顺序：左侧游标值绝对不能大于右侧游标值
                        if self.curr_left > self.curr_right:
                            self.curr_left = self.curr_right
                            
                        # 将校验修正后的合法值重新挂载到滑块实例
                        if hasattr(self.slider, "setRangeValues"):
                            self.slider.setRangeValues(self.curr_left, self.curr_right)
                    except Exception:
                        pass

                # ---------------------------------------------------
                # 步骤 D：同步 Qt 浮动透明遮罩层 (Overlay)
                # 遮罩层通常用于显示加载动画、特殊状态提示或拦截底层图表鼠标事件
                # ---------------------------------------------------
                if getattr(self, "overlay", None) is not None:
                    if hasattr(self.overlay, "set_dpr"):
                        self.overlay.set_dpr(dpr)
                    if hasattr(self.overlay, "set_plot_rect"):
                        self.overlay.set_plot_rect(plotL, plotT, plotW, plotH)
                    if hasattr(self.overlay, "set_margins"):
                        self.overlay.set_margins(l, r, t, b)
                    # 数据注入完成后，强制触发遮罩层 Qt 引擎的重绘事件
                    self.overlay.update()

                # ---------------------------------------------------
                # 步骤 E：同步加载过程中的骨架屏 (Skeleton Screen) 边距
                # ---------------------------------------------------
                if getattr(self, "chart_skeleton", None) is not None:
                    if hasattr(self.chart_skeleton, "set_margins"):
                        self.chart_skeleton.set_margins((l, r, t, b))

                # ---------------------------------------------------
                # 步骤 F：尾部清理与子组件位置刷新
                # 核心布局几何参数全部更新并稳定后，触发最后一次的垂直参考线与悬浮输入框的位置刷新。
                # 利用 QTimer.singleShot(0, ...) 将具体任务推迟到当前事件循环栈清空后执行，防止主线程 UI 卡顿。
                # ---------------------------------------------------
                QTimer.singleShot(0, self._update_vlines_only)
                QTimer.singleShot(0, self._update_floating_edits_pos)

            except Exception:
                pass

        try:
            # 执行 JS 并在执行完毕后安全地调用 _cb 回调进行 UI 同步
            self.web.page().runJavaScript(js, _cb)
        except Exception:
            pass

    # =======================================================
    # 3. 异步渲染生命周期控制与防抖 (Lifecycle & Debounce)
    # =======================================================
    def _after_plot_ready_visual_sync(self):
        """
        图表初次就绪后的同步调度：
        专门处理 Webview 页面初次加载完毕后的视觉对齐初始化。
        由于 Plotly 的首次渲染计算（包括坐标轴刻度推导、图例计算等）需要一定耗时，
        直接同步可能会读取到尚未完全展开的旧坐标，因此需要采取防抖和延迟策略。
        """
        # 首先确保视图在尚未完全稳定时，不被用户的误操作拖拽变形
        self._enforce_plotly_view_xrange()
        
        try:
            # 通知底层图表事件控制器，安排一次矩形区域的同步检查
            if getattr(self, "_chart_ctrl", None):
                self._chart_ctrl.schedule_rect_sync()
        except Exception:
            pass

        # 防抖与延迟重试策略：
        # 采用两次延迟同步 (0ms 和 160ms) 有助于吸收 Webview 与 Plotly 内部较晚的 DOM 重排 (Reflow) 过程带来的坐标偏差。
        QTimer.singleShot(0, self._sync_plot_geom_from_plotly)
        QTimer.singleShot(160, self._sync_plot_geom_from_plotly)
        
        # 同步更新外部悬浮编辑框的位置
        QTimer.singleShot(0, self._update_floating_edits_pos)
        QTimer.singleShot(120, self._update_floating_edits_pos)

    def _after_plot_update(self):
        """
        图表数据更新后的同步调度：
        负责在用户调整模型参数触发重绘时，管理加载动画的切换，并触发交错的布局对齐操作。
        由于数据变化会引起极值变动和坐标轴重算，此步骤的同步频率相对初次加载会更频密。
        """
        # 显示前端骨架屏，向用户提供“正在绘图...”的过渡视觉反馈
        self._show_skeleton("正在绘图...")
        
        # 更新内部状态的事务标识位，防止乱序更新
        self._data_sent = True
        self._render_token += 1
        
        # 绘图指令发送完毕后，尝试隐藏骨架屏
        self._maybe_hide_skeleton()
        # 更新期间，强制保持 X 轴锁定状态
        self._enforce_plotly_view_xrange()

        try:
            if getattr(self, "_chart_ctrl", None):
                self._chart_ctrl.schedule_rect_sync()
        except Exception:
            pass

        # 时序防抖策略：多级交错同步（Ticks）
        # 连续进行 0ms, 100ms, 300ms 间隔的几何坐标提取同步。
        # 此策略旨在应对浏览器内部不可控的异步渲染微任务时序，能够有效保留历史交互行为的连续性，
        # 极大程度降低了由于图表瞬间重绘导致的 UI 联动控件（如滑块、悬浮框）在视觉上的跳跃与抖动。
        QTimer.singleShot(0, self._sync_plot_geom_from_plotly)
        QTimer.singleShot(100, self._sync_plot_geom_from_plotly)
        QTimer.singleShot(300, self._sync_plot_geom_from_plotly)
        
        # 确保悬浮输入框能够正确追踪到最新的垂直参考线位置
        QTimer.singleShot(0, self._update_floating_edits_pos)

    # =======================================================
    # 4. 异常拦截与主界面弹窗提示 (Exception Handling)
    # =======================================================
    def _show_histogram_error_popup(self, err: Exception, axis_style=None, layout_dict=None):
        """
        展示图表渲染错误的系统级界面弹窗：
        用于捕获底端 Python 数学计算引擎或数据序列化阶段抛出的未期严重异常，
        并在主 Qt 界面上展示友好的、具可读性的错误提示面板，确保程序容错机制的完整，避免静默崩溃。
        """
        try:
            # 提取异常堆栈信息的第一行作为核心摘要进行展示
            lines = str(err).splitlines()
            summary = lines[0] if lines else repr(err)

            # 构建并弹出模态级别的系统错误提示窗口
            msg = QMessageBox(self)
            msg.setIcon(QMessageBox.Critical)
            # 统一应用窗体错误级别的标题命名
            msg.setWindowTitle("Drisk - 绘图失败")
            # 呈现具体的异常类型名称及摘要明细
            msg.setText(f"{type(err).__name__}: {summary}")
            msg.exec()
        except Exception:
            pass