# ui_advanced.py
"""
高级分析视图（摘要箱线图、敏感性龙卷风图、情景分析图）的核心路由调度与状态管理器。
负责统筹高级模式下底部工具栏按钮的互斥高亮、前端堆叠视图（QStackedWidget）的层级切换、图表样式的全局重置与应用。
提供了一套轻量级的状态快照（Snapshot）机制，确保用户在跨版本切换“分析对象”或重新加载底层数据时，能够无缝保留并恢复当前的高级分析环境与配置。
"""

from dataclasses import dataclass
from ui_results_runtime_state import ResultsRuntimeStateHelper

# =======================================================
# 1. 数据模型与状态快照类
# =======================================================
@dataclass
class AdvancedModeSnapshot:
    """
    高级模式状态快照数据类 (Lightweight Snapshot)。
    用于在用户跨模拟 ID 切换“分析对象”或重新加载底层数据集前，
    轻量级地捕获并保存当前高级分析视图（龙卷风、情景分析、箱线图）的激活状态与子模式标记。
    这样在数据刷新完成后，可以将界面无缝恢复到用户刚才停留的高级分析环境。
    """
    is_tornado_active: bool   # 敏感性（龙卷风）图是否处于激活状态
    is_scenario_active: bool  # 情景分析图是否处于激活状态
    is_boxplot_active: bool   # 摘要（箱线）图是否处于激活状态
    tornado_mode: str         # 龙卷风图当前的子模式（如 "bins", "bins_line" 等）
    boxplot_mode: str         # 箱线图当前的子模式（如 "boxplot", "violin" 等）

# =======================================================
# 2. 高级视图路由与状态调度器
# =======================================================
class ResultsAdvancedRouter:
    """
    高级视图路由与状态调度器。
    将原本杂糅在主窗口中的 UI 状态流转逻辑进行了解耦。负责统筹高级模式下的
    UI 互斥联动、QStackedWidget 视图层级切换、以及渲染入口的安全分发。
    """

    # ============================
    # Part 1: 模式常量与基础状态判别
    # ============================
    # 定义摘要图(Boxplot)家族的子模式集合
    _BOXPLOT_FAMILY = {"boxplot", "letter_value", "violin", "trend"}
    # 定义所有高级分析模式的集合（龙卷风 + 情景分析 + 箱线图家族）
    _ADVANCED_MODES = {"tornado", "scenario"} | _BOXPLOT_FAMILY

    @staticmethod
    def is_boxplot_family(mode: str) -> bool:
        """检查给定模式字符串是否属于摘要(箱线)图家族。"""
        return mode in ResultsAdvancedRouter._BOXPLOT_FAMILY

    @staticmethod
    def is_advanced_mode(mode: str) -> bool:
        """检查给定模式字符串是否属于任何一种高级分析模式。"""
        return mode in ResultsAdvancedRouter._ADVANCED_MODES

    @staticmethod
    def resolve_major_mode(mode: str) -> str:
        """
        核心映射函数：将非常细分的子模式（如 'violin' 小提琴图）归类并映射到
        粗粒度的四大基础主模式（'main', 'scenario', 'tornado', 'boxplot'）中。
        主要用于触发全局样式重置机制，判断是否发生了跨主模式的切换。
        """
        if mode == "scenario":
            return "scenario"
        if mode == "tornado":
            return "tornado"
        if ResultsAdvancedRouter.is_boxplot_family(mode):
            return "boxplot"
        return "main"

    # ============================
    # Part 2: UI 组件交互与视觉层级同步
    # ============================
    @staticmethod
    def _sync_button_states(dialog, mode: str):
        """
        同步并维护底部工具栏中高级模式控制按钮的互斥状态与显隐逻辑。
        确保同一时间只会高亮按下一个模式按钮（如激活情景图时，自动弹起龙卷风图和摘要图按钮），
        并智能控制该模式专属的“设置”按钮（如 btn_scenario_settings）的可见性。
        """
        is_boxplot_family = ResultsAdvancedRouter.is_boxplot_family(mode)
        
        if mode == "scenario":
            if hasattr(dialog, "btn_tornado"): dialog.btn_tornado.setChecked(False)
            if hasattr(dialog, "btn_boxplot"): dialog.btn_boxplot.setChecked(False)
            if hasattr(dialog, "btn_scenario"): dialog.btn_scenario.setChecked(True)
            if hasattr(dialog, "btn_scenario_settings"): dialog.btn_scenario_settings.show()
            return

        if mode == "tornado":
            if hasattr(dialog, "btn_scenario"): dialog.btn_scenario.setChecked(False)
            if hasattr(dialog, "btn_boxplot"): dialog.btn_boxplot.setChecked(False)
            if hasattr(dialog, "btn_tornado"): dialog.btn_tornado.setChecked(True)
            if hasattr(dialog, "btn_scenario_settings"): dialog.btn_scenario_settings.hide()
            return

        if is_boxplot_family:
            if hasattr(dialog, "btn_scenario"): dialog.btn_scenario.setChecked(False)
            if hasattr(dialog, "btn_tornado"): dialog.btn_tornado.setChecked(False)
            if hasattr(dialog, "btn_boxplot"): dialog.btn_boxplot.setChecked(True)
            if hasattr(dialog, "btn_scenario_settings"): dialog.btn_scenario_settings.hide()

    @staticmethod
    def _switch_view_stack(dialog, mode: str):
        """
        路由调度底层的 QStackedWidget。根据传入的模式，将前端显示层透明地切换至
        对应的主图表区 (main_chart_wrapper)、龙卷风视图区 (tornado_view) 或 箱线图视图区 (boxplot_view)。
        """
        if mode in {"tornado", "scenario"}:
            dialog.view_stack.setCurrentWidget(dialog.tornado_view)
            return
        if ResultsAdvancedRouter.is_boxplot_family(mode):
            dialog.view_stack.setCurrentWidget(dialog.boxplot_view)
            return
        dialog.view_stack.setCurrentWidget(dialog.main_chart_wrapper)

    @staticmethod
    def sync_view_state(dialog, mode: str):
        """
        UI 视觉状态同步的综合入口。
        一次性执行：样式分类重置 -> 按钮互斥同步 -> 层级面板切换。
        随后，根据是否切回了基础主图表区，智能隐藏/显示标准模式下的专有组件
        （如顶部悬浮滑块区域 top_wrapper 和 右侧独立统计指标面板 right_panel），最后刷新底部控制栏布局。
        """
        dialog._check_and_reset_styles(ResultsAdvancedRouter.resolve_major_mode(mode))
        ResultsAdvancedRouter._sync_button_states(dialog, mode)
        ResultsAdvancedRouter._switch_view_stack(dialog, mode)

        is_main_view = dialog.view_stack.currentWidget() == dialog.main_chart_wrapper
        
        if hasattr(dialog, "top_wrapper"):
            dialog.top_wrapper.setVisible(is_main_view)
        if hasattr(dialog, "right_panel"):
            dialog.right_panel.setVisible(is_main_view)

        # 切回主视图时，恢复顶部的常规变量名称标题
        if is_main_view and hasattr(dialog, "_base_chart_title") and hasattr(dialog, "chart_title_label"):
            dialog.chart_title_label.setText(dialog._base_chart_title)

        dialog._update_bottom_bar_visibility()

    # ============================
    # Part 3: 渲染任务算法入口分发
    # ============================
    @staticmethod
    def dispatch_render_entry(dialog, mode: str):
        """
        高级图表渲染任务分发器。
        在此将请求“解复用”，反向回调（Callback）主窗口（dialog）中对应图表架构的底层计算与渲染方法，
        从而将渲染的具体业务逻辑与状态路由保持分离。
        """
        if ResultsAdvancedRouter.is_boxplot_family(mode):
            dialog._render_advanced_boxplot_mode(mode)
            return
        if mode == "tornado":
            dialog._load_and_render_tornado()
            return
        if mode == "scenario":
            dialog._render_advanced_scenario_mode()

    # ============================
    # Part 4: 状态快照的捕获与恢复
    # ============================
    @staticmethod
    def capture_mode_snapshot(dialog) -> AdvancedModeSnapshot:
        """
        在数据源发生更替前，抓取并生成当前高级模式下各类组件的开关与属性状态快照。
        """
        return AdvancedModeSnapshot(
            is_tornado_active=hasattr(dialog, "btn_tornado") and dialog.btn_tornado.isChecked(),
            is_scenario_active=hasattr(dialog, "btn_scenario") and dialog.btn_scenario.isChecked(),
            is_boxplot_active=hasattr(dialog, "btn_boxplot") and dialog.btn_boxplot.isChecked(),
            tornado_mode=ResultsRuntimeStateHelper.get_tornado_mode(dialog, "bins"),
            boxplot_mode=ResultsRuntimeStateHelper.get_current_analysis_mode(dialog, "boxplot"),
        )

    @staticmethod
    def apply_mode_snapshot(dialog, snapshot: AdvancedModeSnapshot):
        """
        在数据源重新加载完毕后，应用之前保存的快照恢复高级模式环境。
        通过代理类将记录的子模式参数重新写入运行环境（Runtime State），并做安全性拦截
        （如非法子模式字符将被兜底恢复为 "boxplot"）。
        """
        if snapshot.is_tornado_active:
            ResultsRuntimeStateHelper.set_tornado_mode(dialog, snapshot.tornado_mode)
            return
        if snapshot.is_scenario_active:
            ResultsRuntimeStateHelper.set_scenario_mode(dialog)
            return
        if snapshot.is_boxplot_active:
            ResultsRuntimeStateHelper.set_boxplot_mode(
                dialog,
                snapshot.boxplot_mode
                if snapshot.boxplot_mode in {"boxplot", "letter_value", "violin", "trend"}
                else "boxplot"
            )