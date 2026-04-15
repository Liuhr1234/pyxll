# ui_modeler_runtime_interaction_mixin.py
"""
本模块提供分布构建器（DistributionBuilderDialog）的运行时交互辅助功能（Mixin）。

主要功能：
1. 悬浮输入与反向计算 (Floating Inputs & Reverse Prob): 处理直接输入 X 坐标或输入概率百分比倒推 X 坐标的逻辑。
2. 滑块拖拽事件流 (Slider Drag Events): 处理底部双向滑块的按下、拖拽（防抖联动）、释放等完整生命周期。
3. 状态同步与更新管线 (State Sync & Updates): 在内部状态 (curr_left, curr_right) 与外部 UI 之间同步数据，并分发轻量/全量更新。
4. 视觉层联动 (Visual Coordination): 控制图表上的垂直区间参考线 (V-lines) 的移动以及视图模式切换。
"""

import drisk_env
from PySide6.QtCore import QTimer

from ui_shared import update_floating_value_edits_pos, DRISK_COLOR_CYCLE_THEO


class DistributionBuilderRuntimeInteractionMixin:
    """
    [混合类] 滑块、概率计算及悬浮输入框的运行时交互辅助工具。
    """

    # =======================================================
    # 1. 悬浮输入与反向计算 (精确控制)
    # =======================================================
    def _apply_floating_input(self, which: str):
        """
        处理悬浮输入框的精确数值输入：
        用户在滑块上方的手动输入框中输入具体数值时触发，将数值解析、格式化并应用到当前滑块区间。
        """
        which = (which or "").lower()
        group = self.float_x_l if which == "l" else self.float_x_r
        try:
            # 清理格式化字符（如后缀、千分位逗号）
            suffix = getattr(group.edit, "suffix", "")
            txt = (group.edit.text() or "").strip()
            if suffix:
                txt = txt.replace(suffix, "")
            txt = txt.replace(",", "")

            if not txt:
                return
            
            # 解析数值并结合当前的量级 (Magnitude) 进行还原计算
            v = float(txt)
            mag = 0
            try:
                mag = int(getattr(getattr(group, "mag", None), "current_mag", 0))
            except Exception:
                mag = 0
            raw = v * (10 ** mag)
        except Exception:
            return

        # 获取当前的物理显示边界，并对输入值进行钳制 (Clamp)
        xmin = getattr(self, "_display_xmin", self.view_min)
        xmax = getattr(self, "_display_xmax", self.view_max)
        raw = max(xmin, min(raw, xmax))

        # 根据操作的是左滑块还是右滑块，更新对应的状态
        if which == "l":
            new_l, new_r = raw, float(getattr(self, "curr_right", raw))
        else:
            new_l, new_r = float(getattr(self, "curr_left", raw)), raw

        # 防护：确保左侧始终小于等于右侧
        if new_l > new_r:
            new_l, new_r = new_r, new_l

        # 同步更新底层滑块组件
        try:
            self.slider.setRangeValues(float(new_l), float(new_r))
        except Exception:
            pass

        self.curr_left, self.curr_right = float(new_l), float(new_r)

        # 触发视觉与数据刷新
        self._update_vlines_only()
        self._update_floating_edits_pos()
        
        # 保持手动输入的行为与拖拽完成后的行为一致：触发完整的统计和概率刷新。
        self.perform_full_update()

        # 输入完成后清除焦点，避免持续占用键盘输入
        if which == "l":
            self.float_x_l.edit.clearFocus()
        else:
            self.float_x_r.edit.clearFocus()

        self.curr_left, self.curr_right = float(new_l), float(new_r)
        self._update_vlines_only()
        self._update_floating_edits_pos()

    def _update_floating_edits_pos(self):
        """
        更新悬浮输入框的屏幕物理位置。
        调用 ui_shared 中的通用方法，使其始终紧跟滑块的句柄移动。
        """
        try:
            if not hasattr(self, "float_x_l") or self.view_max <= self.view_min:
                return

            update_floating_value_edits_pos(
                slider=self.slider,
                group_l=self.float_x_l,
                group_r=self.float_x_r,
                val_l=self.curr_left,
                val_r=self.curr_right,
                range_min=self.view_min,
                range_max=self.view_max,
                margin_l=0,
                margin_r=0,
                dtick=self.x_dtick,
                forced_mag=self.manual_mag,
                gap=0,
                clamp=True,
                single_handle=False
            )
        except Exception:
            pass

    def _on_slider_input_confirmed(self, mode_str, val_str):
        """
        反向概率推导：
        用户在滑块的概率标签（如 5%）上直接输入期望的百分比时触发。
        利用底层分布的 PPF（百分位函数 / 逆累积分布函数）反推对应的 X 轴坐标。
        """
        try:
            # 清洗输入的百分比文本
            txt = val_str.strip().replace("%", "")
            try:
                p = float(txt) / 100.0
            except ValueError:
                return

            p = max(0.0, min(1.0, p))

            if not self.dist_obj:
                return

            # 根据操作的是哪一段概率（左尾、右尾、中间）计算 PPF
            if mode_str.startswith("pl"):
                # 左尾概率 -> 直接查 PPF
                self.curr_left = self._safe_ppf_scalar(self.dist_obj, p)
            elif mode_str.startswith("pr"):
                # 右尾概率 -> 反向查 1-p 的 PPF
                self.curr_right = self._safe_ppf_scalar(self.dist_obj, 1.0 - p)
            elif mode_str.startswith("pm"):
                # 中间概率 -> 假定对称，从两端各切除一半的尾部
                tail = (1.0 - p) / 2.0
                self.curr_left = self._safe_ppf_scalar(self.dist_obj, tail)
                self.curr_right = self._safe_ppf_scalar(self.dist_obj, 1.0 - tail)

            # 更新 UI 状态
            self.update_sliders_from_state()

        except Exception:
            pass


    # =======================================================
    # 2. 滑块拖拽事件流 (实时响应与防抖)
    # =======================================================
    def on_slider_drag_started(self):
        """滑块开始拖拽：标记拖拽状态，停止部分耗时更新。"""
        self._is_slider_dragging = True

    def on_slider_change(self, l, r):
        """
        滑块实时拖拽中：
        进行数据吸附（Snap）、边界钳制，并触发轻量级的视觉更新（参考线移动）。
        由于拖拽触发频率极高，繁重的图表重算会被推入节流定时器（throttle）。
        """
        try:
            if self._block_slider_events:
                return

            # 若开启了网格吸附，对坐标进行就近取整
            if self._drag_snap_enabled:
                l = self._snap_value_drag(l)
                r = self._snap_value_drag(r)

            # 获取显示范围并钳制拖拽值
            xmin = getattr(self, "_display_xmin", None)
            xmax = getattr(self, "_display_xmax", None)
            if xmin is None or xmax is None:
                xmin, xmax = self.view_min, self.view_max

            l = max(xmin, min(l, xmax))
            r = max(xmin, min(r, xmax))
            if l > r:
                l, r = r, l

            self.curr_left, self.curr_right = float(l), float(r)

            # 实时更新悬浮框位置与图表上的垂直参考线
            QTimer.singleShot(0, self._update_floating_edits_pos)
            self._update_vlines_only()

            # 在拖拽移动期间，将耗时的更新任务推迟到节流定时器 (throttle timer) 中执行。
            if not self.throttle.isActive():
                self.throttle.start()

        except Exception:
            pass

    def on_slider_drag_finished(self, l, r):
        """
        滑块拖拽结束 (松开鼠标)：
        确认最终数值，停止节流定时器，并强制执行一次全量的深度更新（包括统计表格重算）。
        """
        try:
            self._is_slider_dragging = False

            if self._drag_snap_enabled:
                l = self._snap_value_drag(l)
                r = self._snap_value_drag(r)

            self.curr_left, self.curr_right = float(l), float(r)

            self._block_slider_events = True
            self.slider.setRangeValues(self.curr_left, self.curr_right)
            self._block_slider_events = False

            # 清除并停止防抖定时器
            if self.throttle.isActive():
                self.throttle.stop()

            # 触发深度全量更新
            self.perform_full_update()

        except Exception:
            pass


    # =======================================================
    # 3. 状态同步与核心更新管线
    # =======================================================
    def update_sliders_from_state(self):
        """
        状态下发：将类中存储的内部状态 (curr_left, curr_right) 反向应用到 UI 滑块上。
        常用于初始化或代码层面修改了界限后同步 UI。
        """
        if self.curr_left > self.curr_right:
            self.curr_left, self.curr_right = self.curr_right, self.curr_left

        self.curr_left = max(self.view_min, min(self.curr_left, self.view_max))
        self.curr_right = max(self.view_min, min(self.curr_right, self.view_max))

        # 屏蔽事件流，防止死循环
        self._block_slider_events = True
        try:
            self.slider.setRangeValues(self.curr_left, self.curr_right)
        finally:
            self._block_slider_events = False

        self.perform_full_update()
        QTimer.singleShot(0, self._update_floating_edits_pos)

    def perform_buffered_update(self):
        """轻量级更新：仅更新参考线与底部概率输入框，不请求耗时的统计后台任务。"""
        if getattr(self, "_use_qt_overlay", False):
            self._update_vlines_only()
        else:
            self._plot_host.update_lr(self.curr_left, self.curr_right)
        self.update_prob_boxes()

    def perform_full_update(self):
        """全量重量级更新：更新所有视觉元素，并强制触发右侧统计表格的重新渲染。"""
        self._update_vlines_only()
        self.update_prob_boxes()
        # 只有当统计表格从未渲染过时，才强制刷新
        self.request_stats_table_update(force=(not self._stats_table_ever_rendered))

    def update_prob_boxes(self):
        """
        核心概率计算与 UI 更新：
        基于当前滑块的位置，向底层引擎查询其 CDF（累积概率），
        进而计算出左尾 (pl)、中间 (pm) 和右尾 (pr) 面积的真实概率，并将其推送到滑块组件进行显示。
        """
        if not self.dist_obj:
            return
            
        if self.is_discrete:
            # Use right-closed boundaries for discrete mode:
            # left tail = P(X <= L), right tail = P(X > R).
            cdf_l = self._safe_cdf_scalar(self.dist_obj, self.curr_left)
            cdf_r = self._safe_cdf_scalar(self.dist_obj, self.curr_right)
            pl = float(cdf_l)
            pr = 1.0 - float(cdf_r)
            if pr < 0:
                pr = 0
        else:
            # 连续模式直接使用当前坐标查询 CDF
            pl = float(self._safe_cdf_scalar(self.dist_obj, self.curr_left))
            pr = 1.0 - float(self._safe_cdf_scalar(self.dist_obj, self.curr_right))

        pl = max(0.0, min(1.0, pl))
        pr = max(0.0, min(1.0, pr))
        pm = 1.0 - pl - pr
        if pm < 0:
            pm = 0.0

        # 构建图层数据推送给滑块组件
        layers = [{
            'key': 'Model',
            'color': DRISK_COLOR_CYCLE_THEO[0],
            'probs': (pl, pm, pr)
        }]

        if hasattr(self.slider, "set_layers"):
            self.slider.set_layers(layers)


    # =======================================================
    # 4. 视觉层联动
    # =======================================================
    def _update_vlines_only(self):
        """
        仅更新垂直区间参考线 (V-lines)：
        根据当前的视图模式，指挥 Plotly 内部 JS 或 Qt Overlay 重新绘制区间高亮范围。
        """
        try:
            if getattr(self, "_use_qt_overlay", False) and getattr(self, "_chart_ctrl", None):
                pmode = self.get_preview_mode()

                if pmode == "cdf":
                    mode_str = "cdf"
                elif pmode.startswith("pdf"):
                    mode_str = "pdf"
                elif getattr(self, "is_discrete", False) or pmode.startswith("pmf"):
                    mode_str = "discrete"
                else:
                    mode_str = "hist"

                self._chart_ctrl.update_visuals(
                    l=float(self.curr_left),
                    r=float(self.curr_right),
                    x_min=getattr(self, "_display_xmin", self.view_min),
                    x_max=getattr(self, "_display_xmax", self.view_max),
                    mode=mode_str
                )
                return

            # 如果没有启用 Qt 覆盖层，则向 Webview 注入 JS 代码来更新 Plotly 范围
            self.web.page().runJavaScript(
                f"updateVLines({float(self.curr_left)}, {float(self.curr_right)});"
            )
        except Exception:
            pass

    def on_preview_mode_changed(self, *args):
        """
        图表预览模式切换处理：
        当用户在下拉菜单改变图表类型（如 PDF 切换至 CDF）时，重置视觉状态并重新渲染。
        """
        try:
            cmd = self.get_preview_mode()
            self._has_y2 = ("_cdf" in cmd or "+" in cmd)

            QTimer.singleShot(0, self._sync_plot_geom_from_plotly)

            # 确保滑块为双向模式
            if hasattr(self, "slider"):
                self.slider.setSingleHandleMode(False)

            self.render_preview()
            self.perform_full_update()

            QTimer.singleShot(0, self._update_floating_edits_pos)

        except Exception:
            pass

