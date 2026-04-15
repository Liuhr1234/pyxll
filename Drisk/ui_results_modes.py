# ui_results_modes.py
"""
本模块提供结果视图（Results Views）的模式状态管理与渲染调度服务。

主要功能：
1. 视图指令解析 (View Command Parsing)：将用户从下拉菜单触发的字符串指令（如 "histogram_cdf"）解析为结构化的状态对象。
2. 状态规范化与防呆 (Normalization)：根据当前数据的特征（连续/离散、单曲线/多曲线），自动修正不合法的视图指令，回退到安全模式。
3. 渲染调度 (Render Dispatching)：将规范化后的指令路由至具体的渲染通道，并处理主视图与高级分析视图（如龙卷风图）之间的状态回退。
"""

from dataclasses import dataclass


from ui_results_runtime_state import ResultsRuntimeStateHelper


# =======================================================
# 1. 视图指令状态结构
# =======================================================
@dataclass(frozen=True)
class ViewCommandState:
    """
    [不可变数据类] 规范化后的视图指令与叠加层标志位。
    将复杂的字符串组合拆解为布尔开关，方便下游逻辑直接调用，避免重复进行字符串包含判断。
    """
    raw: str          # 原始触发字符串，如 "histogram_kde"
    base_view: str    # 基础视图类型，如 "histogram"
    show_cdf: bool    # 是否叠加累积概率曲线 (CDF)
    show_kde: bool    # 是否叠加核密度曲线 (KDE)
    show_dkw: bool    # 是否叠加非参置信区间 (DKW)


# =======================================================
# 2. 视图指令解析与应用服务
# =======================================================
class ResultsViewCommandHelper:
    """
    处理视图指令字符串的解析、状态映射以及安全性降级的辅助类。
    """

    @staticmethod
    def parse_command(cmd) -> ViewCommandState:
        """
        指令解析器：
        通过约定好的下划线后缀语义（如 '_cdf', '_kde', '_all'）提取出基础图表类型及需要激活的叠加图层。
        """
        raw = (cmd or "auto").strip().lower()
        if not raw:
            raw = "auto"
            
        # 提取下划线前的第一部分作为基础视图 (例如 "histogram_kde" -> "histogram")
        base_view = raw.split("_")[0]
        
        return ViewCommandState(
            raw=raw,
            base_view=base_view,
            show_cdf=("_cdf" in raw or "_all" in raw),
            show_kde=("_kde" in raw or "_all" in raw),
            show_dkw=("_dkw" in raw),
        )

    @staticmethod
    def apply_command_state(dialog, cmd) -> str:
        """
        状态应用：
        解析命令并将结果同步挂载到主对话框 (dialog) 的内部状态属性上。
        同时调用调度器更新当前的全局图表模式 (chart_mode)。
        返回解析后的基础视图模式名称。
        """
        state = ResultsViewCommandHelper.parse_command(cmd)
        
        # 将叠加层开关同步至对话框运行时环境
        dialog._show_cdf_overlay = state.show_cdf
        dialog._show_kde_overlay = state.show_kde
        dialog._show_dkw_overlay = state.show_dkw
        
        # 同步底层绘图模式
        ResultsRenderDispatcher.set_chart_mode_from_view(dialog, state.base_view)
        return state.base_view

    @staticmethod
    def normalize_for_dataset(cmd, *, first_open: bool, is_discrete_view: bool, display_count: int) -> str:
        """
        [核心防呆逻辑] 指令规范化与冲突降级：
        根据当前渲染上下文（是否离散、叠加的情景数量等）校验指令合法性。
        如果不合法，强制回退 (Fallback) 到兼容的默认模式。
        """
        # 如果是首次打开弹窗，强制使用自动模式推断最佳图表
        if first_open:
            return "auto"
            
        current_cmd = cmd or "auto"
        
        # 规则 1：离散数据 (Discrete) 强制屏蔽连续概率分布特有的渲染指令 (如 KDE, Histogram)
        if is_discrete_view and current_cmd in ["pdfcurve", "histogram", "histogram_kde", "histogram_cdf", "histogram_all"]:
            return "auto"
            
        # 规则 2：多情景对比 (Multiple Series) 渲染防重叠控制
        # 多个情景同时叠加 KDE 曲线会导致视觉极其混乱且性能低下，
        # 因此强制剥离 KDE 渲染，降级为叠加 CDF 或纯直方图。
        if display_count > 1 and current_cmd in ["histogram_kde", "histogram_all"]:
            return "histogram_cdf" if current_cmd == "histogram_all" else "histogram"
            
        return current_cmd


# =======================================================
# 3. 渲染调度与视图切换服务
# =======================================================
class ResultsRenderDispatcher:
    """
    负责非高级渲染路径的路由分发，以及在不同分析模式间跳转时的状态重置与 UI 恢复。
    """

    @staticmethod
    def set_chart_mode_from_view(dialog, view_mode: str):
        """
        状态映射：将具体的视图指令名称转化为底层导图及 UI 控制流所需的 chart_mode 标识。
        """
        if view_mode == "cdf":
            dialog.chart_mode = "cdf"
        elif view_mode == "pdfcurve":
            dialog.chart_mode = "kde"
        elif getattr(dialog, "is_discrete_view", False) or view_mode == "discrete":
            dialog.chart_mode = "discrete"
        else:
            # 默认为直方图模式
            dialog.chart_mode = "hist"
            # 当返回直方图时，需要重置可能残留在内存中的龙卷风图模式
            ResultsRuntimeStateHelper.reset_tornado_mode(dialog, "bins")

    @staticmethod
    def dispatch_non_advanced_render(dialog, view_mode: str):
        """
        渲染分发中心：
        根据最终确认的 view_mode，调用 dialog 上挂载的具体渲染通道方法（这些方法由 ResultsMainRenderService 提供实际实现）。
        """
        # 自动推导模式路由
        if view_mode == "auto":
            ResultsRenderDispatcher.set_chart_mode_from_view(dialog, "auto")
            dialog._render_channel_auto()
            return

        # 同步状态并分发给具体通道
        ResultsRenderDispatcher.set_chart_mode_from_view(dialog, view_mode)
        
        if view_mode == "cdf":
            dialog._render_channel_cdf()
            return
        if view_mode == "pdfcurve":
            dialog._render_channel_pdf()
            return
        if view_mode == "discrete":
            dialog._render_channel_discrete()
            return

        # 处理带变体的直方图系列指令（如 relfreq 相对频数）
        if view_mode not in ("histogram", "relfreq"):
            view_mode = "histogram"
        dialog._render_channel_hist(view_mode)

    @staticmethod
    def restore_main_view_from_advanced_if_needed(dialog) -> bool:
        """
        高级视图退出与状态恢复系统：
        当用户在处于“高级分析视图”（龙卷风、情景分析、箱线图）时，触发了其他需要回到基础图表的操作，
        此方法将自动重置所有高级按钮的按下状态，隐藏高级配置悬浮窗，并将界面切回主控面板。
        返回 True 表示确实发生了一次从高级视图的回退拦截。
        """
        need_restore = False
        
        # 检测并弹起已激活的高级模式按钮
        if hasattr(dialog, "btn_tornado") and dialog.btn_tornado.isChecked():
            dialog.btn_tornado.setChecked(False)
            need_restore = True
            
        if hasattr(dialog, "btn_scenario") and dialog.btn_scenario.isChecked():
            dialog.btn_scenario.setChecked(False)
            need_restore = True
            
        if hasattr(dialog, "btn_boxplot") and dialog.btn_boxplot.isChecked():
            dialog.btn_boxplot.setChecked(False)
            need_restore = True

        # 如果发生回退，执行深入的 UI 现场清理
        if need_restore:
            # 清除记录的当前高级模式名称
            ResultsRuntimeStateHelper.clear_advanced_mode(dialog)
            
            # 隐藏龙卷风专属设置项及子面板
            if hasattr(dialog, "btn_tornado_settings"):
                dialog.btn_tornado_settings.hide()
                dialog.btn_tornado_settings.setChecked(False)
                if hasattr(dialog, "tornado_view") and hasattr(dialog.tornado_view, "config_panel"):
                    dialog.tornado_view.config_panel.hide()
                    
            # 隐藏情景分析专属设置项
            if hasattr(dialog, "btn_scenario_settings"):
                dialog.btn_scenario_settings.hide()
                
            # 重新激活并渲染基础图表视图
            dialog._activate_analysis_view()
            
        return need_restore