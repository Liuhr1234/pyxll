# ui_modeler_recalc_orchestration_mixin.py
"""
本模块提供分布构建器（DistributionBuilderDialog）重新计算（Recalculation）流程的编排辅助功能。

主要功能：
1. 重算前置准备 (Pre-Recalc Orchestration)：在触发底层耗时的数学重算前，进行严格的参数校验、数据清洗（如处理数组大括号）、UI 标题同步，并尝试获取后端底层分布对象。
2. 重算后置刷新 (Post-Recalc Refresh)：在底层数据更新后，安全地恢复和同步 UI 控件状态（如双向滑块边界、悬浮输入框位置），并触发最终的图表渲染。
"""

import drisk_env
import math
from PySide6.QtCore import QTimer

import backend_bridge
from drisk_charting import DriskChartFactory


class DistributionBuilderRecalcOrchestrationMixin:
    """
    [混合类] 分布构建器重算编排混合类
    
    设计意图：
    围绕核心重算方法 `recalc_distribution` 提供轻量级的流程编排辅助工具。
    将复杂的“UI状态读取 -> 参数校验与清洗 -> 底层数学对象实例化 -> UI状态写回”整个生命周期拆分为清晰的前置和后置两部分，降低主类的代码耦合度。
    """

    # =======================================================
    # 1. 重算前置准备与数据编排 (Pre-Recalc Orchestration)
    # =======================================================
    
    def _validate_truncation_against_distribution_support(self):
        """
        [防呆与校验] 验证截断边界是否合法：
        校验基于数值的截断边界是否在底层分布的支持集（Support，即有效定义域）范围内。
        当截断区间完全处于支持集之外时，返回错误提示信息；若校验通过或暂无数据，返回 None。
        """
        markers = getattr(self, "_dist_markers", {}) or {}

        trunc_key = None
        trunc_bounds = None
        
        # 提取当前激活的截断类型与边界值（支持 truncate 基础截断 和 truncate2 第二种截断模式）
        for key in ("truncate", "truncate2"):
            val = markers.get(key)
            if isinstance(val, (tuple, list)) and len(val) >= 2:
                trunc_key = key
                trunc_bounds = val
                break
                
        # 如果没有设置截断边界，则直接跳过校验
        if trunc_bounds is None:
            return None

        def _to_finite_or_none(v):
            """内部辅助函数：将输入安全转换为有限的浮点数，若无效则返回 None"""
            try:
                if v is None:
                    return None
                fv = float(v)
                return fv if math.isfinite(fv) else None
            except Exception:
                return None

        # 清洗并获取用户输入的截断极值
        trunc_min = _to_finite_or_none(trunc_bounds[0])
        trunc_max = _to_finite_or_none(trunc_bounds[1])
        if trunc_min is None and trunc_max is None:
            return None

        # 复制基础标记，并剔除所有截断相关的属性，以便获取一个“纯净”的基础分布用于比对
        base_markers = dict(markers)
        for key in ("truncate", "truncate2", "truncatep", "truncatep2"):
            base_markers.pop(key, None)

        # 核心逻辑："truncate" 指的是平移（shift）发生前的数值截断。
        # 因此，必须剥离 "shift" 属性，与平移前的基础分布支持集进行校验。
        if trunc_key == "truncate":
            base_markers.pop("shift", None)

        # 尝试通过后端桥接层获取基础分布的数学实例对象
        try:
            support_dist = backend_bridge.get_backend_distribution(
                self._dist_func_name,
                self._dist_params,
                base_markers
            )
        except Exception:
            support_dist = None
            
        if support_dist is None:
            return None

        # 提取基础分布的理论最小值和最大值
        try:
            support_min = _to_finite_or_none(getattr(support_dist, "min_val", lambda: None)())
        except Exception:
            support_min = None
        try:
            support_max = _to_finite_or_none(getattr(support_dist, "max_val", lambda: None)())
        except Exception:
            support_max = None

        # 边界碰撞检测并返回易读的中文错误提示
        if trunc_min is not None and support_max is not None and trunc_min > support_max:
            return f"截断最小值不能大于分布最大值（当前最大值约为 {support_max:g}）。"
        if trunc_max is not None and support_min is not None and trunc_max < support_min:
            return f"截断最大值不能小于分布最小值（当前最小值约为 {support_min:g}）。"
            
        return None

    def _prepare_recalc_orchestration(self):
        """
        重算前置准备中心：
        将重算前的参数校验、状态初始化，与耗费性能的底层数学采样逻辑分离开来。
        确保只有在所有参数合法的情况下，才会去请求后端引擎实例化分布对象。
        
        返回值: 
            bool: 返回 True 表示准备就绪，可以进行后续的重算逻辑；返回 False 表示校验失败或实例化失败并已中断流程。
        """
        # 清空图表Y轴的裁剪提示信息
        self._y_clip_note = ""
        
        # 1. 严格参数校验
        ok, vals, err = self._validate_params_strict()
        if not ok:
            # 如果校验失败，尝试弹出对应的错误提示（如针对累积分布的特殊提示），并中断重算
            self._maybe_show_cumul_validation_prompt(err)
            return False

        # 2. 提取并清洗分布参数
        self._dist_func_name = str(self.config.get("func_name", "DriskNormal"))
        self._dist_params = []
        for k, _, _ in self.config.get("params", []):
            raw_val = vals[k]
            # 规范化处理类数组的输入形式，以供 backend_bridge（后端桥接层）消费
            if isinstance(raw_val, str):
                cleaned_val = raw_val.strip()
                # 如果用户输入包含大括号 {}，则将其剥离，保留内部干净的数据
                if cleaned_val.startswith("{") and cleaned_val.endswith("}"):
                    cleaned_val = cleaned_val[1:-1].strip()
                self._dist_params.append(cleaned_val)
            else:
                self._dist_params.append(raw_val)

        # 3. 收集高级属性标记 (如 Name 名称, Shift 平移, Truncate 截断等)
        self._dist_markers = self._collect_markers()
        if self._dist_markers is None:
            return False

        # 4. 同步 UI 标题与图表坐标轴元数据
        units_val = self._dist_markers.get('units', '')

        # 更新弹窗顶部图表标题
        if hasattr(self, "chart_title_label"):
            if hasattr(self, "_resolve_modeler_display_chart_title"):
                full_title = self._resolve_modeler_display_chart_title()
            else:
                full_title = getattr(self, "fallback_chart_name", "新建建模变量")
            self.chart_title_label.setText(full_title)

        # 配置底层图表工厂的坐标轴单位名称
        DriskChartFactory.VALUE_AXIS_TITLE = ""
        DriskChartFactory.VALUE_AXIS_UNIT = units_val.strip() if units_val else ""

        # 5. 调用后端桥接层，获取真正的底层分布数学实例
        self.dist_obj = None
        if backend_bridge is not None:
            try:
                self.dist_obj = backend_bridge.get_backend_distribution(
                    self._dist_func_name,
                    self._dist_params,
                    self._dist_markers
                )
            except Exception as e:
                # 捕获并打印底层实例化过程中的异常，便于控制台排障
                print(f"[Drisk][ui_modeler] 获取底层分布实例失败: {e}")

        # 实例化失败则中断后续数学计算
        if self.dist_obj is None:
            return False

        # 6. 进行截断边界与分布支持集的二次冲突校验
        trunc_err = self._validate_truncation_against_distribution_support()
        if trunc_err:
            # 若发现冲突，在画布区域抛出视觉错误提示，并清空数据表避免误导用户
            if hasattr(self, "_show_param_validation_error_in_canvas"):
                self._show_param_validation_error_in_canvas(trunc_err)
            if hasattr(self, "tbl") and self.tbl is not None:
                self.tbl.setRowCount(0)
            return False
             
        # 所有前置条件达成，准许放行重算
        return True


    # =======================================================
    # 2. 重算后的 UI 刷新与状态同步 (Post-Recalc Refresh)
    # =======================================================
    
    def _apply_recalc_post_refresh(self):
        """
        重算后置刷新处理：
        在底层数据重算完毕后调用，负责将最新的计算边界和视图范围同步到前端 UI 控件。
        """
        # 注意：此处必须严格保持刷新顺序，以避免由于预览重绘或悬浮层叠加引发的难以察觉的 UI 回退或状态混乱（Regressions）。
        
        # 1. 更新底部双向滑块 (Slider) 的物理极限范围，确保与新的视图边界匹配
        self.slider.setRangeLimit(self.view_min, self.view_max)
        
        # 保持拖拽吸附网格（Snap Grid）与重新计算后的视图范围同步
        try:
            self._rebuild_drag_snap_grid()
        except Exception:
            pass

        # 2. 更新滑块的当前选中区间，同时屏蔽信号防止由于值改变触发死循环回调
        self._block_slider_events = True
        try:
            self.slider.blockSignals(True)
            self.slider.setRangeValues(self.curr_left, self.curr_right)
            self.slider.blockSignals(False)
        finally:
            self._block_slider_events = False

        # 3. 更新依附于滑块上的悬浮输入框（例如百分位数输入框）的屏幕坐标位置
        self._update_floating_edits_pos()
        
        # 4. 触发预览区的快速渲染（仅更新图表路径，不处理复杂的布局变更）
        self.render_preview()
        
        # 5. 使用定时器在当前 Qt 事件循环结束后，异步调度一次全量 UI 更新
        # 这种做法可以确保前面所有的尺寸调整（Resize）和状态变更（State Mutation）都能被 Qt 布局引擎正确处理和消化后，再绘制最终形态的图表。
        QTimer.singleShot(0, self.perform_full_update)