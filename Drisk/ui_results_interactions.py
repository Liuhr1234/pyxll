# ui_results_interactions.py
"""
结果分析界面的交互编排辅助模块。

职责定位：
1. 统一承接结果分析弹窗中的“交互型操作”；
2. 负责在不同视图/模式之间做状态判断与简单切换；
3. 负责打开分析对象选择器、叠加选择器、导出等交互入口；
4. 负责根据当前模式同步底部工具条/按钮显隐状态。

说明：
- 本文件更偏“交互编排层”，不直接承担复杂的数据计算。
- 这里的方法大多是静态方法，调用方式简单，便于从主界面类中转发。
- 具体的高级分析模式识别、运行时状态读取、导出逻辑、弹窗实现，
  分别委托给其他模块处理。
"""

from __future__ import annotations

import traceback

from PySide6.QtWidgets import QMessageBox

# 导出管理器：统一构建导出上下文，并执行实际导出
from drisk_export import DriskExportManager
# 获取当前所有模拟结果对象，用于弹出“分析对象选择器/叠加选择器”
from simulation_manager import get_all_simulations
# 高级分析模式路由器：用于判断当前模式是否属于高级分析
from ui_results_advanced import ResultsAdvancedRouter
# 交互弹窗：分析对象选择器、叠加对象选择器
from ui_results_dialogs import AnalysisObjectDialog, AdvancedOverlayDialog
# 运行时状态辅助器：统一读取当前模式、龙卷风模式等运行态信息
from ui_results_runtime_state import ResultsRuntimeStateHelper
# 模拟显示名版本号：用于判断显示名是否发生变化，决定是否需要刷新界面
from ui_sim_display_names import get_sim_display_name_version


class ResultsInteractionService:
    """
    结果分析界面的交互服务类。

    设计说明：
    - 采用静态方法形式，不维护自身状态；
    - 所有状态都从 dialog（结果分析主对话框实例）读取或回写；
    - 主要作用是把“按钮点击/弹窗打开/视图切换/底栏状态更新”这类
      交互编排逻辑从主界面类中拆出来，降低主类复杂度。
    """

    @staticmethod
    def is_advanced_mode(dialog) -> bool:
        """
        判断当前是否处于高级分析模式。

        处理流程：
        1. 从运行时状态中读取当前分析模式；
        2. 交由 ResultsAdvancedRouter 判断该模式是否属于高级分析。

        返回：
        - True：当前模式属于高级分析；
        - False：当前模式属于普通分析。
        """
        mode = ResultsRuntimeStateHelper.get_current_analysis_mode(dialog, "")
        return ResultsAdvancedRouter.is_advanced_mode(mode)

    @staticmethod
    def check_and_prepare_advanced_mode(dialog) -> bool:
        """
        在进入高级分析前做前置校验与必要清理。

        当前约束：
        - 高级分析仅允许基于输出变量执行；
        - 若当前基础分析对象是输入变量，则直接拦截并提示；
        - 若当前叠加对象中包含输入变量，则清空叠加项并重新加载数据，
          避免高级分析继续沿用不合法的叠加状态。

        返回：
        - True：校验通过，可以进入高级分析；
        - False：校验失败，不应继续执行后续高级分析操作。
        """
        # 高级分析不允许以输入变量作为当前主分析对象
        if getattr(dialog, "_base_kind", "output") == "input":
            QMessageBox.warning(
                dialog,
                "无效操作",
                "当前分析对象为输入数据，请先切换到输出变量后再执行高级分析。",
            )
            return False

        # 若当前存在叠加项，则检查其中是否混入了输入变量
        if hasattr(dialog, "overlay_items") and dialog.overlay_items:
            has_input = False

            # overlay item 兼容两种格式：
            # - 长度为 4：第 4 位记录 kind（input/output）
            # - 否则默认视为 output（兼容旧结构）
            for item in dialog.overlay_items:
                kind = item[3] if len(item) == 4 else "output"
                if kind == "input":
                    has_input = True
                    break

            # 高级分析下若叠加项中含 input，则整体清空并刷新当前数据
            if has_input:
                dialog.overlay_items = []
                dialog.load_dataset(dialog.current_key)

        return True

    @staticmethod
    def open_analysis_obj_dialog(dialog):
        """
        打开“分析对象选择器”弹窗。

        主要用途：
        - 切换当前分析对象；
        - 在高级分析模式下，限制只能选择输出变量；
        - 若只是显示名版本发生变化，也会在取消弹窗后触发一次刷新。

        处理流程：
        1. 读取当前所有模拟对象；
        2. 记录当前模拟显示名版本号；
        3. 判断当前是否处于高级分析模式；
        4. 构造并打开 AnalysisObjectDialog；
        5. 若用户确认选择，则调用 dialog._on_analysis_obj_changed(...)；
        6. 若用户取消，但显示名版本有变化，则刷新当前数据集；
        7. 若过程中报错，则打印堆栈并弹出错误提示。
        """
        try:
            # 获取当前所有模拟，用于在弹窗中展示可选对象
            all_sims = get_all_simulations()

            # 记录弹窗打开前的显示名版本号；
            # 若弹窗期间名称映射发生变化，取消后也需要刷新界面
            name_rev_before = get_sim_display_name_version()

            # 判断当前是否为高级分析模式；
            # 高级分析下会锁定只能选择 output
            is_adv = ResultsInteractionService.is_advanced_mode(dialog)

            # current_key 可能是 label 映射后的展示键；
            # 这里尽量还原成真实 cell key 传给选择器
            real_ck = dialog.label_to_cell_key.get(dialog.current_key, dialog.current_key)

            selector = AnalysisObjectDialog(
                all_sims=all_sims,
                current_kind="output" if is_adv else getattr(dialog, "_base_kind", "output"),
                current_ck=real_ck,
                current_sid=dialog.sim_id,
                parent=dialog,
                lock_to_output=is_adv,
            )

            # 用户点击“确定”
            if selector.exec():
                res = selector.get_selection()
                if res:
                    kind, sid, ck, lbl = res
                    # 回写当前分析对象类型（input/output）
                    dialog._base_kind = kind
                    # 统一走主对话框内部的分析对象切换处理逻辑
                    dialog._on_analysis_obj_changed(sid, ck, lbl)

            # 用户取消，但显示名版本发生变化，则刷新当前展示
            elif get_sim_display_name_version() != name_rev_before:
                dialog.load_dataset(dialog.current_key)

        except Exception as exc:
            traceback.print_exc()
            QMessageBox.critical(dialog, "错误", f"无法打开分析对象选择器: {exc}")

    @staticmethod
    def open_overlay_selector(dialog):
        """
        打开“叠加对象选择器”弹窗。

        主要用途：
        - 配置当前图表的叠加分析对象；
        - 在高级分析模式下，只允许输出变量参与叠加；
        - 用户确认后刷新当前数据及相关缓存；
        - 若高级分析模式处于激活状态，则刷新对应高级分析视图。

        处理流程：
        1. 获取所有模拟对象；
        2. 记录打开前的显示名版本；
        3. 若 overlay_items 尚未初始化，则先创建为空列表；
        4. 打开 AdvancedOverlayDialog；
        5. 若用户确认：
           - 回写 overlay_items；
           - 重新加载当前数据；
           - 清空 CDF 排序缓存与 CDF CI 缓存；
           - 若当前是高级分析模式，则重新激活高级分析视图；
        6. 若用户取消但显示名版本变化，则刷新当前数据；
        7. 异常时打印堆栈并弹框提示。
        """
        try:
            all_sims = get_all_simulations()
            name_rev_before = get_sim_display_name_version()

            # 某些旧状态下 overlay_items 可能尚未初始化，这里做兜底
            if not hasattr(dialog, "overlay_items"):
                dialog.overlay_items = []

            is_adv = ResultsInteractionService.is_advanced_mode(dialog)
            real_ck = dialog.label_to_cell_key.get(dialog.current_key, dialog.current_key)

            selector = AdvancedOverlayDialog(
                all_sims=all_sims,
                current_kind="output" if is_adv else getattr(dialog, "_base_kind", "output"),
                current_ck=real_ck,
                current_sid=dialog.sim_id,
                current_overlays=dialog.overlay_items,
                parent=dialog,
                lock_to_output=is_adv,
            )

            if selector.exec():
                # 用用户在弹窗中最终确认的叠加项替换当前状态
                dialog.overlay_items = selector.get_results()

                # 重新加载当前数据，以反映最新叠加对象
                dialog.load_dataset(dialog.current_key)

                # 若主对话框维护了 CDF 相关缓存，则在叠加项变化后清空，
                # 避免使用旧缓存造成显示不一致
                if hasattr(dialog, "_cdf_sorted_cache"):
                    dialog._cdf_sorted_cache.clear()
                if hasattr(dialog, "_cdf_ci_cache"):
                    dialog._cdf_ci_cache.clear()

                # 若当前是高级分析模式，则需要重新激活高级分析视图，
                # 让界面与当前叠加状态保持一致
                if ResultsInteractionService.is_advanced_mode(dialog):
                    dialog._activate_analysis_view()

            elif get_sim_display_name_version() != name_rev_before:
                # 用户取消，但显示名版本有变化，仍应刷新界面
                dialog.load_dataset(dialog.current_key)

        except Exception as exc:
            traceback.print_exc()
            QMessageBox.critical(dialog, "错误", f"打开叠加选择器失败: {exc}")

    @staticmethod
    def _resolve_export_view_context(dialog):
        """
        根据当前界面所处视图，解析导出应使用的 web_view 与 chart_mode。

        设计原因：
        - 结果分析弹窗并不只有一个图表承载视图；
        - 主图、龙卷风图、情景图、箱线图等，可能分别拥有自己的 web_view；
        - 导出时必须准确找到“当前正在展示的那一个视图”对应的上下文。

        返回：
        - export_web_view：当前应导出的图表视图对象；
        - export_mode：当前应导出的图表模式字符串。
        """
        # 当前 view_stack 中正在展示的 widget
        current_widget = (
            getattr(dialog, "view_stack", None).currentWidget() if hasattr(dialog, "view_stack") else None
        )

        # 判断底部几个关键模式按钮当前是否处于选中状态
        is_tornado = hasattr(dialog, "btn_tornado") and dialog.btn_tornado.isChecked()
        is_scenario = hasattr(dialog, "btn_scenario") and dialog.btn_scenario.isChecked()
        is_boxplot = hasattr(dialog, "btn_boxplot") and dialog.btn_boxplot.isChecked()

        # 默认导出主图 web_view 与当前 chart_mode
        export_web_view = getattr(dialog, "web_view", None)
        export_mode = getattr(dialog, "chart_mode", "hist")

        # 若当前不是主图，而是龙卷风/情景/箱线等特殊视图，
        # 则需要改用对应子视图的 web_view 与 mode
        if (is_tornado or is_scenario or is_boxplot) and current_widget:
            # 当前是龙卷风/情景图共用的视图容器
            if current_widget == getattr(dialog, "tornado_view", None):
                export_web_view = getattr(dialog.tornado_view, "web_view", export_web_view)
                export_mode = "tornado" if is_tornado else "scenario"

            # 当前是箱线图相关视图容器
            elif current_widget == getattr(dialog, "boxplot_view", None):
                export_web_view = getattr(dialog.boxplot_view, "web_view", export_web_view)
                export_mode = ResultsRuntimeStateHelper.get_current_analysis_mode(dialog, "boxplot")

        return export_web_view, export_mode

    @staticmethod
    def on_export_clicked(dialog):
        """
        导出按钮点击后的统一处理入口。

        处理流程：
        1. 先根据当前视图解析实际导出上下文；
        2. 用导出管理器构建 export context；
        3. 若存在 web_view，则先让其获得焦点，减少导出时焦点错位问题；
        4. 调用导出管理器执行真正的导出动作。
        """
        export_web_view, export_mode = ResultsInteractionService._resolve_export_view_context(dialog)
        export_ctx = DriskExportManager.build_export_context(
            dialog,
            web_view=export_web_view,
            chart_mode=export_mode,
        )

        # 某些导出链路可能依赖当前视图焦点，这里在导出前显式聚焦
        if export_ctx.web_view:
            export_ctx.web_view.setFocus()

        DriskExportManager.export_from_dialog(dialog, context=export_ctx)

    @staticmethod
    def toggle_tornado_view(dialog, checked: bool):
        """
        切换龙卷风视图开关时的界面处理。

        参数：
        - checked=True：进入龙卷风视图；
        - checked=False：退出龙卷风视图，回到主图视图。

        退出龙卷风视图时的处理包括：
        1. 切回主图容器；
        2. 恢复右侧统计面板；
        3. 隐藏龙卷风设置按钮；
        4. 隐藏情景设置按钮；
        5. 重新应用主图 chart mode 的 UI 状态；
        6. 最后统一刷新底部工具条显隐状态。
        """
        # 顶部主图包装区在龙卷风模式下通常需要隐藏
        if hasattr(dialog, "top_wrapper"):
            dialog.top_wrapper.setHidden(checked)

        # 退出龙卷风模式
        if not checked:
            # 切回主图 widget
            dialog.view_stack.setCurrentWidget(dialog.main_chart_wrapper)

            # 恢复右侧信息区；
            # 新版可能叫 right_panel，旧版可能仍是 stats_table
            if hasattr(dialog, "right_panel"):
                dialog.right_panel.show()
            else:
                dialog.stats_table.show()

            # 隐藏龙卷风设置按钮，并取消其选中状态
            if hasattr(dialog, "btn_tornado_settings"):
                dialog.btn_tornado_settings.hide()
                dialog.btn_tornado_settings.setChecked(False)

            # 隐藏情景设置按钮
            if hasattr(dialog, "btn_scenario_settings"):
                dialog.btn_scenario_settings.hide()

            # 恢复主图模式对应的 UI 状态
            dialog.apply_chart_mode_ui()

        # 无论进入还是退出龙卷风模式，都统一更新底部工具条显隐
        ResultsInteractionService.update_bottom_bar_visibility(dialog)

    @staticmethod
    def update_bottom_bar_visibility(dialog):
        """
        根据当前模式，统一控制底部工具条相关按钮的显隐状态。

        当前涉及的控件主要包括：
        - 情景图设置按钮 btn_scenario_settings
        - 龙卷风设置按钮 btn_tornado_settings
        - CDF 置信区间设置按钮 btn_ci_settings
        - 标签设置按钮 btn_label_settings
        - 以及当前版本中暂时统一隐藏的 CI 引擎选择控件

        设计目标：
        - 不同模式下只显示有意义的按钮，避免用户误操作；
        - 把分散的按钮显隐判断统一收敛到这里。
        """
        is_tornado = hasattr(dialog, "btn_tornado") and dialog.btn_tornado.isChecked()
        is_scenario = hasattr(dialog, "btn_scenario") and dialog.btn_scenario.isChecked()

        # 当前是否处于主图容器
        is_main_view = dialog.view_stack.currentWidget() == getattr(dialog, "main_chart_wrapper", None)

        # 情景图设置按钮：仅在情景图模式下显示
        if hasattr(dialog, "btn_scenario_settings"):
            dialog.btn_scenario_settings.setVisible(is_scenario)

        # 龙卷风设置按钮：仅在龙卷风模式且模式为 bins / bins_line 时显示
        if hasattr(dialog, "btn_tornado_settings"):
            tornado_mode = ResultsRuntimeStateHelper.get_tornado_mode(dialog, "bins")
            dialog.btn_tornado_settings.setVisible(is_tornado and tornado_mode in ["bins", "bins_line"])

        # DKW 相关 CI 设置按钮：
        # 只有在主图视图 + 开启 DKW 叠加 + 当前图模式为 CDF 时才显示
        is_dkw_active = (
            is_main_view
            and getattr(dialog, "_show_dkw_overlay", False)
            and getattr(dialog, "chart_mode", "hist") == "cdf"
        )

        # 当前版本中，CI 引擎标签与下拉框统一隐藏
        if hasattr(dialog, "lbl_ci_engine"):
            dialog.lbl_ci_engine.setVisible(False)
        if hasattr(dialog, "combo_ci_engine"):
            dialog.combo_ci_engine.setVisible(False)

        # 仅保留 CI 设置按钮在特定条件下显示
        if hasattr(dialog, "btn_ci_settings"):
            dialog.btn_ci_settings.setVisible(is_dkw_active)

        # 标签设置按钮默认显示，但在某些模式下需要隐藏或受限
        show_label_settings = True

        # 情景图模式下不显示标签设置
        if is_scenario:
            show_label_settings = False

        # 龙卷风模式下，仅部分龙卷风子模式允许显示标签设置
        elif is_tornado:
            tornado_mode = ResultsRuntimeStateHelper.get_tornado_mode(dialog, "bins")
            show_label_settings = tornado_mode in ["bins", "reg_mapped"]

        if hasattr(dialog, "btn_label_settings"):
            dialog.btn_label_settings.setVisible(show_label_settings)

    @staticmethod
    def on_global_mag_changed(dialog, mag_int: int):
        """
        处理“全局数量级/倍率（manual_mag）”变更后的刷新逻辑。

        参数：
        - mag_int：新的全局数量级整数值。

        处理思路：
        1. 先把新值写回 dialog.manual_mag；
        2. 再根据当前模式决定采用哪一种刷新方式：
           - tornado：重载并重新绘制龙卷风图；
           - boxplot / letter_value / violin / trend / scenario：
             统一通过高级分析视图刷新；
           - 其他普通模式：
             重新初始化主图，并同步更新 X 轴输入框位置。
        """
        dialog.manual_mag = int(mag_int)
        mode = ResultsRuntimeStateHelper.get_current_analysis_mode(dialog, "")

        if mode == "tornado":
            dialog._load_and_render_tornado()
        elif mode in ["boxplot", "letter_value", "violin", "trend", "scenario"]:
            dialog._activate_analysis_view()
        else:
            dialog.init_chart_via_tempfile()
            dialog._update_x_edits_pos()