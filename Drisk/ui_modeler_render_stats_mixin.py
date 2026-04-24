# -*- coding: utf-8 -*-
# ui_modeler_render_stats_mixin.py
"""
本模块为 DistributionBuilderDialog 提供“统计表 + 图表预览”两组能力的混合类实现，
主要承担建模弹窗中的理论分布预览渲染、统计指标异步计算、窗口生命周期协同等职责。

模块职责概述：
1. 统计表异步更新
   - 将均值、方差、分位数等统计量计算放入后台线程执行；
   - 通过缓存键判断参数是否真实变化，避免重复计算；
   - 在参数频繁变动时，通过定时器防抖，减少无意义刷新。

2. 图表窗口生命周期管理
   - 在窗口 resize 时触发 Plotly 自适应；
   - 在窗口关闭时回收 overlay、webview、线程池等资源；
   - 协调图表区、滑块区、浮动输入框区的几何同步。

3. 理论分布图表渲染
   - 支持 PDF（概率密度）、CDF（累积分布）、PMF（离散概率质量）、
     理论直方图 / 相对频率直方图等预览模式；
   - 支持主图叠加右侧 Y 轴的 CDF 曲线；
   - 对接 DriskChartFactory / Plotly，统一下发 JSON 到前端渲染宿主。

4. 预览模式路由
   - 根据界面下拉菜单返回的模式指令，路由到不同渲染通道；
   - 在连续 / 离散分布间自动选择适配的默认图形。

说明：
- 本文件是 Mixin，不独立工作，依赖宿主对话框提供大量运行时属性与辅助方法；
- 本文件的重点不是“定义分布”，而是“把已配置好的分布对象渲染出来，并展示统计结果”；
- 若后续接手人定位渲染问题，应优先联动查看：
  1) drisk_charting.py
  2) ui_workers.py
  3) ui_stats.py
  4) ui_shared.py 中与坐标轴、webview 回收相关的方法
"""

import drisk_env
import json
import math
import traceback

import numpy as np
import plotly.graph_objects as go
import plotly.utils

from PySide6.QtCore import Qt, QTimer
from PySide6.QtGui import QColor
from PySide6.QtWidgets import QTableWidgetItem

import ui_stats
from drisk_charting import DriskChartFactory
from ui_workers import ModelerStatsJob
from ui_shared import (
    DRISK_COLOR_CYCLE_THEO,
    DriskMath,
    build_plotly_axis_style,
    recycle_shared_webview,
)


class DistributionBuilderRenderStatsMixin:
    """
    [混合类]
    从 DistributionBuilderDialog 中拆出的“统计表渲染 + 图表预览渲染”方法集合。

    设计定位：
    - 不负责弹窗整体初始化；
    - 不负责分布对象本身的构造；
    - 专门处理“分布对象已经准备好之后”的预览与统计展示工作。

    对宿主对象的主要依赖包括但不限于：
    - self.dist_obj：当前分布对象
    - self.tbl：统计表格控件
    - self.web / self._plot_host：前端图表承载对象
    - self.x_data / self.y_data：理论曲线采样数据
    - self.curr_left / self.curr_right：当前左右概率锚点 / 滑块位置
    - self._safe_cdf_vec(...)：安全计算 CDF 向量的方法
    - self._sync_plot_geom_from_plotly()：同步 Plotly 绘图区几何信息
    - self._show_skeleton(...) / self._maybe_hide_skeleton()：绘图遮罩控制
    """

    # =======================================================
    # 1. 统计数据表与后台任务管理
    # =======================================================
    def _make_stats_cache_key(self):
        """
        生成统计表缓存键。

        目的：
        - 判断当前统计表对应的“分布配置状态”是否发生变化；
        - 若未变化，则直接复用缓存结果，不重复启动后台统计任务；
        - 若变化，则触发重新计算。

        参与缓存键的核心内容包括：
        1. 分布类型 self.dist_type
        2. 分布函数名 self._dist_func_name
        3. 参数列表 self._dist_params
        4. 标记信息 self._dist_markers（如平移、截断等运行态标记）
        5. dist_obj 的 args / kwds
        6. 单元格地址 self.cell_address
        7. 左右边界 curr_left / curr_right

        注意：
        - 这里专门做了 _norm_value 归一化，是为了把 dict / list / numpy 数值等
          转成可比较、可哈希、且结构稳定的 tuple / 基础类型；
        - 尤其是 markers、kwds 这类嵌套结构，如果不归一化，缓存会失真，
          可能出现“参数已变但缓存未失效”的问题。
        """
        if not self.dist_obj:
            return None

        # 将嵌套运行时状态归一化为“可比较、可缓存”的稳定结构，
        # 确保诸如“平移 / 截断标记变化”也能体现在缓存键中。
        def _norm_value(v):
            # 字典：按 key 排序后递归展开，避免键顺序不同导致缓存误判
            if isinstance(v, dict):
                pairs = []
                for k in sorted(v.keys(), key=lambda x: str(x)):
                    pairs.append((str(k), _norm_value(v.get(k))))
                return tuple(pairs)

            # 列表 / 元组：递归归一化为 tuple
            if isinstance(v, (list, tuple)):
                return tuple(_norm_value(x) for x in v)

            # numpy 标量：转为 Python 原生数值
            if isinstance(v, (np.integer, np.floating)):
                v = v.item()

            # 普通数值：仅在有限值时保留 float 形式
            if isinstance(v, (int, float)) and not isinstance(v, bool):
                try:
                    fv = float(v)
                    if math.isfinite(fv):
                        return fv
                except Exception:
                    pass

            # 空值 / 布尔 / 字符串：直接保留
            if v is None or isinstance(v, (bool, str)):
                return v

            # 其余对象：兜底转字符串，避免不可哈希对象破坏缓存键构造
            return str(v)

        # dist_obj.args：部分分布对象内部会保留位置参数
        try:
            args = tuple(getattr(self.dist_obj, "args", ()) or ())
        except Exception:
            args = ()

        # dist_obj.kwds：部分分布对象内部会保留关键字参数
        try:
            kwds = getattr(self.dist_obj, "kwds", {}) or {}
            kwds = tuple(sorted((str(k), _norm_value(v)) for k, v in kwds.items()))
        except Exception:
            kwds = ()

        # 读取当前运行态中的分布函数名、参数、标记与左右锚点
        func_name = str(getattr(self, "_dist_func_name", "") or "")
        params = _norm_value(list(getattr(self, "_dist_params", []) or []))
        markers = _norm_value(dict(getattr(self, "_dist_markers", {}) or {}))
        lx = getattr(self, 'curr_left', 0)
        rx = getattr(self, 'curr_right', 0)

        # 返回一个足够完整的缓存键元组
        return (
            str(self.dist_type),
            func_name,
            params,
            markers,
            args,
            kwds,
            str(self.cell_address or ""),
            lx,
            rx,
        )

    def request_stats_table_update(self, force: bool = False):
        """
        请求更新统计表（防抖入口）。

        调用语义：
        - 该方法通常不直接计算统计量；
        - 它只负责“标记需要更新”，并启动一个防抖定时器；
        - 真正执行更新的是 _flush_stats_table_update。

        参数：
        - force=True：表示即使缓存键未变化，也要强制刷新一次。
        """
        if force:
            self._stats_force_next = True

        if not self._stats_table_timer.isActive():
            self._stats_table_timer.start()

    def _flush_stats_table_update(self):
        """
        执行真正的统计表更新逻辑。

        执行顺序：
        1. 检查当前是否存在分布对象；
        2. 生成缓存键；
        3. 若当前界面已渲染过相同 key，且无强制刷新，则直接跳过；
        4. 若命中统计缓存，则直接渲染缓存结果；
        5. 若未命中缓存，则启动新的后台任务进行统计计算；
        6. 在任务启动前先将表格切到“计算中...”状态。

        这是统计表更新链路的核心调度方法。
        """
        if not self.dist_obj:
            self.tbl.setRowCount(0)
            return

        key = self._make_stats_cache_key()
        if key is None:
            return

        # 情况 1：
        # 本次 key 与当前已渲染 key 一致，且没有强制刷新需求，
        # 说明界面展示已经是最新结果，无需任何动作。
        if (not self._stats_force_next) and (self._stats_render_key == key):
            return

        # 情况 2：
        # 命中缓存：说明此前已经完成过相同参数的后台统计计算，
        # 直接将缓存行渲染出来即可，无需重新起线程。
        if (not self._stats_force_next) and (self._stats_cache_key == key) and (self._stats_cache_rows is not None):
            self._render_stats_rows(self._stats_cache_rows)
            self._stats_render_key = key
            self._stats_force_next = False
            return

        # 情况 3：
        # 缓存未命中，必须启动新的后台统计任务。
        # 这里通过 token 机制区分“新旧任务”，防止慢任务回写覆盖新结果。
        self._stats_job_token += 1
        current_token = self._stats_job_token
        self._pending_stats_key = key

        # 先把表格切换到“加载中”状态，给用户明确反馈
        self._render_loading_state()

        lx = getattr(self, 'curr_left', None)
        rx = getattr(self, 'curr_right', None)

        try:
            # 兼容不同版本的 ui_workers.ModelerStatsJob 构造签名。
            # 新版本优先传入 func_name / params / markers；
            # 若抛出 TypeError，则降级回老版本接口。
            if getattr(self, "_dist_func_name", None):
                try:
                    job = ModelerStatsJob(
                        token=current_token,
                        func_name=self._dist_func_name,
                        params=list(getattr(self, "_dist_params", [])),
                        markers=dict(getattr(self, "_dist_markers", {}) or {}),
                        cell_label=self.cell_address or "",
                        left_x=lx,
                        right_x=rx
                    )
                except TypeError as e:
                    print(f"[警告] ui_workers 接口版本未同步（{e}），启用兼容降级路径。")
                    job = ModelerStatsJob(
                        current_token,
                        self.dist_obj,
                        self.cell_address or "",
                        lx,
                        rx
                    )
            else:
                # 如果没有显式函数名，则直接传 dist_obj 给后台任务
                job = ModelerStatsJob(
                    current_token,
                    self.dist_obj,
                    self.cell_address or "",
                    lx,
                    rx
                )

            # 连接任务完成 / 失败信号，并压入线程池执行
            job.signals.done.connect(self._on_stats_job_done)
            job.signals.error.connect(self._on_stats_job_error)
            self._stats_pool.start(job)

            # 当前强制刷新标记已消费
            self._stats_force_next = False

        except Exception as ex:
            # 任务连启动都失败时，走统一错误回调，让表格上能显示错误信息
            print(f"[错误] 启动统计任务失败：{ex}")
            traceback.print_exc()
            self._on_stats_job_error(current_token, str(ex))

    def _on_stats_job_done(self, token, rows):
        """
        后台统计任务完成后的回调。

        关键机制：
        - 使用 token 判断是否为“当前最新任务”；
        - 若不是，则说明这是一个过期任务结果，直接丢弃；
        - 若是，则写入缓存并刷新表格。

        参数：
        - token：本次任务令牌
        - rows：后台返回的统计结果行数据
        """
        # 忽略过期任务，避免旧任务覆盖新参数下的最新界面
        if token != self._stats_job_token:
            return

        # 写入缓存
        if hasattr(self, '_pending_stats_key'):
            self._stats_cache_key = self._pending_stats_key
            self._stats_cache_rows = rows

        # 正式渲染到统计表
        self._render_stats_rows(rows)

        # 更新“当前已渲染 key”
        if hasattr(self, '_pending_stats_key'):
            self._stats_render_key = self._pending_stats_key

        self._stats_table_ever_rendered = True

    def _on_stats_job_error(self, token, error_msg):
        """
        后台统计任务失败后的回调。

        行为：
        - 仅处理当前最新 token 的报错；
        - 清空表格并在第一行写入错误信息；
        - 使用红色字体强调错误状态。
        """
        print(f"\n[后台统计报错] {error_msg}\n")

        if token != self._stats_job_token:
            return

        self.tbl.clearContents()
        self.tbl.setRowCount(1)

        item = QTableWidgetItem(f"错误：{error_msg}")
        item.setForeground(QColor("red"))
        self.tbl.setItem(0, 0, item)

    def _render_stats_rows(self, rows):
        """
        将后台返回的统计结果渲染到 QTableWidget 中。

        处理逻辑：
        1. 兼容 rows 为 dict 或 list 两种结构；
        2. 调用 ui_stats.assemble_simulation_rows 整理为统一表格行结构；
        3. 提取分布显示名称，作为统计表列标题；
        4. 调用 ui_stats.render_stats_table 执行统一绘制。

        注意：
        - 虽然这里叫 assemble_simulation_rows，但在理论分布建模界面也复用了同一套表格结构；
        - 列标题使用当前分布名，而不是固定写死。
        """
        try:
            # 兼容“单个 dict”与“列表”两种后台返回格式
            stats_list = [rows] if isinstance(rows, dict) else rows
            formatted_rows = ui_stats.assemble_simulation_rows(stats_list)

            # 从配置中提取界面显示名称；若名称带括号，只取主名部分
            raw_name = str(self.config.get("ui_name", self.dist_type))
            dist_name = raw_name.split("(")[0].strip()

            # 理论分布颜色采用理论图统一色板首色
            theme_color = DRISK_COLOR_CYCLE_THEO[0]

            ui_stats.render_stats_table(
                self.tbl,
                formatted_rows,
                column_headers=[dist_name],
                header_colors=[theme_color]
            )
        except Exception as e:
            # 统计表渲染失败时，仍在表格内给出可见反馈，避免“静默失败”
            print("\n" + "=" * 60)
            print(f"[错误] 统计面板渲染失败（ui_modeler -> ui_stats）\n{e}")
            traceback.print_exc()
            print("=" * 60 + "\n")

            self.tbl.clearContents()
            self.tbl.setRowCount(1)

            item = QTableWidgetItem("统计表渲染失败，请查看日志")
            item.setForeground(QColor("red"))
            self.tbl.setItem(0, 0, item)

    def _render_loading_state(self):
        """
        在统计表区域显示“计算中”占位状态。

        使用场景：
        - 参数刚改变；
        - 后台统计任务已提交但尚未返回；
        - 通过禁用单元格交互，明确这是临时占位行。
        """
        self.tbl.clearContents()
        self.tbl.setRowCount(1)

        item = QTableWidgetItem("计算中...")
        item.setFlags(Qt.NoItemFlags)
        item.setForeground(Qt.gray)
        self.tbl.setItem(0, 0, item)

    def update_stats(self):
        """
        统计区全局同步入口。

        当前行为：
        1. 先同步概率输入框状态；
        2. 再强制刷新统计表。

        适用场景：
        - 关键参数变更后需要确保统计表与概率框同时更新；
        - 外部调用方希望跳过缓存、立即重算统计值。
        """
        self.update_prob_boxes()
        self.request_stats_table_update(force=True)

    # =======================================================
    # 2. 窗口事件与生命周期管理
    # =======================================================
    def _force_plotly_resize(self):
        """
        强制 Plotly 图表根据当前 WebView 容器尺寸执行重排。

        说明：
        - Qt 容器 resize 后，Plotly 并不总会自动感知；
        - 因此这里通过 runJavaScript 主动调用 Plotly.Plots.resize(gd)；
        - chart_div 为前端图表根节点的约定 ID。
        """
        if hasattr(self, "web") and self.web is not None:
            self.web.page().runJavaScript(
                "try { var gd = document.getElementById('chart_div'); "
                "if (gd && window.Plotly) { Plotly.Plots.resize(gd); } } catch(e) {}"
            )

    def resizeEvent(self, event):
        """
        窗口尺寸变化事件。

        处理目标：
        1. 在窗口拖拽过程中，持续跟踪 Plotly 实际绘图区几何信息；
        2. 使用防抖机制，在拖拽结束后执行 Plotly resize；
        3. 同步浮动输入框位置；
        4. 在连续型模式下，重建拖拽吸附网格。

        这里同时用了两个定时器：
        - _geom_poll_timer：高频轮询绘图区几何，防止坐标与遮罩错位；
        - _plotly_resize_timer：低频防抖 resize，减少拖拽过程中的闪烁与崩溃风险。
        """
        super().resizeEvent(event)

        # 1) 高频几何同步定时器：约 33ms 一次，接近 30fps
        if not hasattr(self, "_geom_poll_timer"):
            self._geom_poll_timer = QTimer(self)
            self._geom_poll_timer.setInterval(33)
            self._geom_poll_timer.timeout.connect(self._sync_plot_geom_from_plotly)

        # 2) 停止轮询的单次定时器：窗口停止变化一段时间后结束高频轮询
        if not hasattr(self, "_geom_poll_stop_timer"):
            self._geom_poll_stop_timer = QTimer(self)
            self._geom_poll_stop_timer.setSingleShot(True)
            self._geom_poll_stop_timer.timeout.connect(self._geom_poll_timer.stop)

        self._geom_poll_timer.start()
        self._geom_poll_stop_timer.start(500)

        # 3) Plotly resize 防抖：避免用户持续拖拽时反复触发重排
        if not hasattr(self, "_plotly_resize_timer"):
            self._plotly_resize_timer = QTimer(self)
            self._plotly_resize_timer.setInterval(150)
            self._plotly_resize_timer.setSingleShot(True)
            self._plotly_resize_timer.timeout.connect(self._force_plotly_resize)
        self._plotly_resize_timer.start()

        # 4) 同步浮动编辑框（例如概率输入框）的位置
        self._update_floating_edits_pos()

        # 5) 连续型模式下，拖拽吸附网格可能依赖当前尺寸，需要重建
        try:
            if not getattr(self, "is_discrete", False):
                self._rebuild_drag_snap_grid()
        except Exception:
            pass

    def closeEvent(self, event):
        """
        窗口关闭事件。

        回收目标：
        1. 停止 ready 轮询；
        2. 清空统计线程池中的待执行任务；
        3. 销毁 overlay；
        4. 断开 webview 的 loadFinished 信号；
        5. 回收共享 webview，避免重复创建造成资源泄漏。

        这是本混合类的主要资源回收入口。
        """
        try:
            if hasattr(self, "_ready_poll_timer") and self._ready_poll_timer.isActive():
                self._ready_poll_timer.stop()
        except Exception:
            pass

        # 清空统计线程池中的排队任务
        if hasattr(self, '_stats_pool'):
            self._stats_pool.clear()

        # 释放 overlay 覆盖层
        if hasattr(self, "overlay") and self.overlay is not None:
            self.overlay.hide()
            self.overlay.setParent(None)
            self.overlay.deleteLater()
            self.overlay = None

        # 回收 webview
        if hasattr(self, "web") and self.web is not None:
            try:
                self.web.loadFinished.disconnect(self._on_webview_load_finished)
            except Exception:
                pass

            try:
                self.web.hide()
            except Exception:
                pass

            recycle_shared_webview(self.web)
            self.web = None

        super().closeEvent(event)

    # =======================================================
    # 3. 核心图表渲染通道
    # =======================================================
    def _render_channel_pdf_theory(self, y_max: float):
        """
        渲染通道：理论概率密度曲线（PDF）。

        功能特点：
        - 使用 x_data / y_data 作为理论 PDF 曲线；
        - 可选叠加右侧 Y 轴的 CDF 曲线；
        - 根据是否启用右轴动态调整右边距；
        - 通过 DriskChartFactory.build_theory_pdf 统一生成前端绘图 JSON。

        参数：
        - y_max：主 Y 轴上限，由外部估算或路由层传入。
        """
        try:
            # 是否启用右侧 Y 轴（叠加 CDF）
            show_cdf = getattr(self, "_has_y2", False)
            cdf_vals = None
            if show_cdf and self.dist_obj:
                cdf_vals = self._safe_cdf_vec(self.dist_obj, self.x_data)

            # 将当前状态回写到实例属性
            self._has_y2 = bool(show_cdf)

            # 下一轮事件循环中同步 Plotly 绘图区几何信息
            QTimer.singleShot(0, self._sync_plot_geom_from_plotly)

            # 若存在右轴，则为其额外预留右边距
            current_margin_r = self.MARGIN_R + (self.ADDITIONAL_Y2_MARGIN if show_cdf else 0)

            # 同步滑块容器边距
            if hasattr(self, "slider"):
                self.slider.setMargins(self.MARGIN_L, current_margin_r)

            # 同步 overlay 边距
            if hasattr(self, "overlay") and self.overlay:
                if hasattr(self.overlay, "set_margins"):
                    self.overlay.set_margins(self.MARGIN_L, current_margin_r, self.MARGIN_T, self.MARGIN_B)

            l, r, t, b = self.MARGIN_L, current_margin_r, self.MARGIN_T, self.MARGIN_B

            # 构造理论 PDF 图表
            res = DriskChartFactory.build_theory_pdf(
                x_data=self.x_data,
                y_data=self.y_data,
                x_range=[self.view_min, self.view_max],
                x_dtick=float(self.x_dtick),
                y_max=y_max,
                note_annotation="",
                margins=(l, r, t, b),
                show_overlay_cdf=show_cdf,
                cdf_data=cdf_vals,
                forced_mag=self.manual_mag,
                label_overrides=getattr(self, "_label_settings_config", None),
                axis_numeric_flags=getattr(self, "_label_axis_numeric", None),
            )

            # 下发给前端图表宿主进行渲染
            if self._plot_host:
                self._show_skeleton("正在绘图...")
                self._render_token += 1
                self._data_sent = True

                self._plot_host.load_plot(
                    plot_json=res["plot_json"],
                    js_mode=res["js_mode"],
                    js_logic=res["js_logic"],
                    initial_lr=None if getattr(self, "_use_qt_overlay", False)
                    else (float(self.curr_left), float(self.curr_right)),
                    static_plot=True,
                )

                self._after_plot_update()
                self._maybe_hide_skeleton()

        except Exception:
            # 当前版本保留原有静默吞错策略；
            # 若后续交接中需要增强可观测性，可在这里补充日志。
            pass

    def _render_channel_cdf(self):
        """
        渲染通道：累积分布函数曲线（CDF）。

        连续 / 离散处理差异：
        - 连续型：直接对 x_data 计算 CDF；
        - 离散型：将 x 采样改为整数序列，并在前端使用 hv 阶梯线渲染。

        额外保护：
        - 若底层分布在定义域之外的 CDF 返回异常值，
          这里会根据 min_val / max_val 强制裁正到 [0, 1]。
        """
        try:
            xs = self.x_data

            if self.is_discrete:
                # 离散型分布的 CDF 应按整数点绘制阶梯线
                xs = np.arange(
                    int(math.floor(self.view_min)),
                    int(math.ceil(self.view_max)) + 1,
                    dtype=int
                )
                ys = self._safe_cdf_vec(self.dist_obj, xs)
            else:
                ys = self._safe_cdf_vec(self.dist_obj, xs)

            # 对定义域外的 CDF 做物理意义上的归正：
            # 小于最小值应为 0，大于最大值应为 1
            try:
                d_min = float(self.dist_obj.min_val())
                d_max = float(self.dist_obj.max_val())
                ys = np.where(xs < d_min, 0.0, ys)
                ys = np.where(xs > d_max, 1.0, ys)
            except Exception:
                pass

            group = {
                'x': xs,
                'y': ys,
                'color': DriskChartFactory._hex_to_rgba(DRISK_COLOR_CYCLE_THEO[0], 0.90),
                'dash': 'solid',
                'name': '累积概率',
                'line_shape': 'hv' if self.is_discrete else 'linear'  # 离散：阶梯线；连续：普通折线
            }

            l, r, t, b = self.MARGIN_L, self.MARGIN_R, self.MARGIN_T, self.MARGIN_B

            # 该模式下不额外启用右轴，因此使用默认边距
            if hasattr(self, "slider"):
                self.slider.setMargins(self.MARGIN_L, self.MARGIN_R)

            if hasattr(self, "overlay") and self.overlay:
                if hasattr(self.overlay, "set_margins"):
                    self.overlay.set_margins(l, r, t, b)

            res = DriskChartFactory.build_cdf(
                data_groups=[group],
                x_range=[self.view_min, self.view_max],
                x_dtick=float(self.x_dtick),
                margins=(l, r, t, b),
                forced_mag=self.manual_mag,
                label_overrides=getattr(self, "_label_settings_config", None),
                axis_numeric_flags=getattr(self, "_label_axis_numeric", None),
            )

            if self._plot_host:
                self._show_skeleton("正在绘图...")
                self._render_token += 1
                self._data_sent = True

                self._plot_host.load_plot(
                    plot_json=res["plot_json"],
                    js_mode=res["js_mode"],
                    js_logic=res["js_logic"],
                    initial_lr=None if getattr(self, "_use_qt_overlay", False)
                    else (float(self.curr_left), float(self.curr_right)),
                    static_plot=True,
                )

                self._after_plot_update()
                self._maybe_hide_skeleton()

        except Exception:
            pass

    def _render_channel_discreteprob(self, y_max: float):
        """
        渲染通道：离散型概率质量函数（PMF）柱状图。

        功能特点：
        - 直接使用 self.x_data / self.y_data 作为离散点与概率质量；
        - 可选叠加右侧 Y 轴的 CDF；
        - 由 DriskChartFactory.build_discrete_bar 统一生成离散柱状图。

        参数：
        - y_max：主 Y 轴最大值。
        """
        try:
            show_cdf = getattr(self, "_has_y2", False)
            self._has_y2 = bool(show_cdf)

            # 推迟到下一轮事件循环再同步几何，保证 Plotly DOM 已完成更新
            QTimer.singleShot(0, self._sync_plot_geom_from_plotly)

            # 如果启用右侧 Y 轴，则额外加大右边距
            current_margin_r = self.MARGIN_R + (self.ADDITIONAL_Y2_MARGIN if show_cdf else 0)

            if hasattr(self, "slider"):
                self.slider.setMargins(self.MARGIN_L, current_margin_r)

            if hasattr(self, "overlay") and self.overlay:
                if hasattr(self.overlay, "set_margins"):
                    self.overlay.set_margins(self.MARGIN_L, current_margin_r, self.MARGIN_T, self.MARGIN_B)

            l, r, t, b = self.MARGIN_L, current_margin_r, self.MARGIN_T, self.MARGIN_B

            group = {
                'x': self.x_data,
                'y': self.y_data,
                'color': DRISK_COLOR_CYCLE_THEO[0],
                'filled': True,
                'name': '概率质量'
            }

            # 若启用右轴叠加，则提前准备同一批 x 点上的 CDF 数据
            cdf_vals = None
            if show_cdf and self.dist_obj:
                cdf_vals = self._safe_cdf_vec(self.dist_obj, self.x_data)

            res = DriskChartFactory.build_discrete_bar(
                data_groups=[group],
                x_range=[self.view_min, self.view_max],
                x_dtick=float(self.x_dtick),
                y_max=float(y_max),
                note_annotation=getattr(self, "_discrete_sampling_note", ""),
                margins=(l, r, t, b),
                show_overlay_cdf=show_cdf,
                cdf_data=cdf_vals,
                forced_mag=self.manual_mag,
                label_overrides=getattr(self, "_label_settings_config", None),
                axis_numeric_flags=getattr(self, "_label_axis_numeric", None),
            )

            if self._plot_host:
                self._show_skeleton("正在绘图...")
                self._render_token += 1
                self._data_sent = True

                self._plot_host.load_plot(
                    plot_json=res["plot_json"],
                    js_mode=res["js_mode"],
                    js_logic=res["js_logic"],
                    initial_lr=None if getattr(self, "_use_qt_overlay", False)
                    else (float(self.curr_left), float(self.curr_right)),
                    static_plot=True,
                )

                self._after_plot_update()
                self._maybe_hide_skeleton()

        except Exception:
            pass

    def _render_channel_sample_hist(self, y_max_target=None, mode="density"):
        """
        渲染通道：理论分组直方图。

        适用对象：
        - 连续型理论分布；
        - 通过对理论 CDF 做分箱差分，构造“理论直方图”；
        - 可用于展示：
          1) 密度直方图（mode="density"）
          2) 相对频率直方图（mode="rel_freq"）

        核心思路：
        - 并非对样本做 histogram；
        - 而是对“理论分布”在各箱区间内的概率面积做近似表达；
        - 因此这里通过 CDF(edges) 的差值获取每箱概率 probs。

        参数：
        - y_max_target：外部指定的 Y 轴参考上限，用于与自动估算值取 max；
        - mode：
            * "density"：柱高 = 概率 / 箱宽，同时叠加 PDF 曲线
            * "rel_freq"：柱高 = 概率 * 100，单位为百分比，不叠加 CDF 右轴
        """
        try:
            if self.dist_obj is None or not hasattr(self, "_plot_host") or self._plot_host is None:
                return

            # 防御性处理：
            # 若当前其实是离散分布，不应进入连续型直方图链路，
            # 这里直接回退到 PMF 柱状图模式。
            if self.is_discrete:
                y_max = 1.0
                if getattr(self, "y_data", None) is not None and len(self.y_data) > 0:
                    y_max = float(np.max(self.y_data))

                old_note = getattr(self, "_discrete_sampling_note", "")
                self._discrete_sampling_note = "理论分布（PMF）"

                self._render_channel_discreteprob(float(y_max))

                self._discrete_sampling_note = old_note
                return

            # ---------------------------------------------------
            # 1. 分箱与理论概率准备
            # ---------------------------------------------------
            n_bins = 50
            edges = np.linspace(self.view_min, self.view_max, n_bins + 1)
            bin_width = edges[1] - edges[0]
            bin_centers = (edges[:-1] + edges[1:]) / 2.0

            # 用 CDF 差分计算每个箱体区间内的理论概率质量
            cdf_vals = self._safe_cdf_vec(self.dist_obj, edges)
            probs = np.diff(cdf_vals)

            # 根据模式决定柱高与 Y 轴标题
            if mode == "rel_freq":
                density_heights = probs * 100.0
                y_title_str = "相对频率（%）"
            else:
                density_heights = probs / bin_width
                y_title_str = "密度"

            # 理论 PDF 曲线采样数据（仅 density 模式叠加）
            curve_x = getattr(self, "x_data", [])
            curve_y = getattr(self, "y_data", [])

            # ---------------------------------------------------
            # 2. 构建 Plotly 轨迹
            # ---------------------------------------------------
            traces = []

            # 直方图柱体
            traces.append(go.Bar(
                x=bin_centers,
                y=density_heights,
                width=[bin_width] * n_bins,
                name="直方图",
                marker=dict(color=DriskChartFactory._hex_to_rgba(DRISK_COLOR_CYCLE_THEO[0], 0.75))
            ))

            # 密度模式下叠加理论 PDF 曲线
            if mode == "density":
                traces.append(go.Scatter(
                    x=curve_x,
                    y=curve_y,
                    mode="lines",
                    name="概率密度",
                    line=dict(width=2, color=DRISK_COLOR_CYCLE_THEO[0])
                ))

            # ---------------------------------------------------
            # 3. 计算并保护 Y 轴范围
            # ---------------------------------------------------
            y_max = float(np.max(density_heights)) if len(density_heights) else 1.0

            if mode == "density" and curve_y is not None and len(curve_y) > 0:
                y_max = max(y_max, float(np.max(curve_y)))

            if y_max_target is not None:
                y_max = max(y_max, float(y_max_target))

            raw_y_max = float(y_max)
            if not math.isfinite(raw_y_max) or raw_y_max <= 1e-12:
                raw_y_max = 1.0

            # 对极端大值做上限保护，避免坐标轴异常膨胀
            cap = 1e3
            clipped = False
            note = ""

            # 优先使用 DriskMath.safe_y_axis_max 统一计算智能 Y 轴上限
            if hasattr(DriskMath, "safe_y_axis_max"):
                y_axis_max, clipped, note = DriskMath.safe_y_axis_max(raw_y_max, cap=cap)
            else:
                # 兼容旧版本：自行计算“友好刻度 + 轻微顶部留白”
                y_dtick_tmp = float(DriskMath.calc_smart_step(raw_y_max))
                if not math.isfinite(y_dtick_tmp) or y_dtick_tmp <= 0:
                    y_dtick_tmp = raw_y_max / 5.0 if raw_y_max > 0 else 1.0

                # 给顶部增加少量呼吸空间，避免图形紧贴上边界
                padded_y_max = raw_y_max + 0.01 * y_dtick_tmp
                y_axis_max = float(math.ceil(padded_y_max / y_dtick_tmp) * y_dtick_tmp)

                if y_axis_max > cap:
                    clipped = True
                    y_axis_max = float(cap)
                    note = f"Y 轴已截断至 {cap:g}"

            # 若发生截断，将说明保留到实例属性，供其他 UI 使用
            self._y_clip_note = note if clipped else ""
            y_dtick = float(DriskMath.calc_smart_step(y_axis_max))

            # ---------------------------------------------------
            # 4. 构建坐标轴样式与布局
            # ---------------------------------------------------
            axis_style = build_plotly_axis_style(
                x_range=(self.view_min, self.view_max),
                y_range=(0.0, y_axis_max),
                x_dtick=float(self.x_dtick),
                y_dtick=float(y_dtick),
                x_title=DriskChartFactory.VALUE_AXIS_TITLE,
                x_unit=DriskChartFactory.VALUE_AXIS_UNIT,
                y_title=y_title_str,
                fixedrange=True,
                forced_mag=self.manual_mag,
                label_overrides=getattr(self, "_label_settings_config", None),
                x_axis_numeric=bool((getattr(self, "_label_axis_numeric", {}) or {}).get("x", True)),
                y_axis_numeric=bool((getattr(self, "_label_axis_numeric", {}) or {}).get("y", True)),
            )

            # 某些返回值兼容修正：避免 hover 配置写成布尔值导致前端表现异常
            try:
                if isinstance(axis_style, dict):
                    if axis_style.get("hovermode", None) is True:
                        axis_style["hovermode"] = "closest"
                    if axis_style.get("hoversubplots", None) is True:
                        axis_style["hoversubplots"] = False
            except Exception:
                pass

            # 统一设置背景色、网格线等直方图风格
            if isinstance(axis_style, dict):
                axis_style.pop("x_mag", None)
                axis_style.pop("y_mag", None)
                axis_style["plot_bgcolor"] = "#fcfcfc"
                axis_style["paper_bgcolor"] = "#ffffff"

                GRID_COLOR = "#dddddd"

                xaxis = dict(axis_style.get("xaxis", {}))
                yaxis = dict(axis_style.get("yaxis", {}))
                xaxis.update(showgrid=True, gridcolor=GRID_COLOR, gridwidth=0.5, zeroline=False)
                yaxis.update(showgrid=True, gridcolor=GRID_COLOR, gridwidth=0.5, zeroline=False)
                axis_style["xaxis"] = xaxis
                axis_style["yaxis"] = yaxis

            # density 模式允许叠加右轴 CDF；rel_freq 模式禁用右轴
            show_cdf = getattr(self, "_has_y2", False)
            if mode == "rel_freq":
                show_cdf = False

            self._has_y2 = bool(show_cdf)
            QTimer.singleShot(0, self._sync_plot_geom_from_plotly)

            current_margin_r = self.MARGIN_R + (self.ADDITIONAL_Y2_MARGIN if show_cdf else 0)

            if hasattr(self, "slider"):
                self.slider.setMargins(self.MARGIN_L, current_margin_r)

            if hasattr(self, "overlay") and self.overlay:
                if hasattr(self.overlay, "set_margins"):
                    self.overlay.set_margins(self.MARGIN_L, current_margin_r, self.MARGIN_T, self.MARGIN_B)

            l, r, t, b = self.MARGIN_L, current_margin_r, self.MARGIN_T, self.MARGIN_B

            layout_dict = {
                **axis_style,
                "barmode": "overlay",
                "showlegend": False,
                "margin": dict(l=l, r=r, t=t, b=b),
                "annotations": []
            }

            # 如果配置中对图标题做了自定义，则在这里覆盖
            chart_cfg = (
                getattr(self, "_label_settings_config", {}) or {}
            ).get("chart_title", {}) if isinstance(getattr(self, "_label_settings_config", {}), dict) else {}

            if isinstance(chart_cfg, dict) and chart_cfg.get("text_override", None) is not None:
                chart_text = str(chart_cfg.get("text_override", "") or "").strip()
                if chart_text:
                    chart_font_family = str(chart_cfg.get("font_family") or "Arial, 'Microsoft YaHei', sans-serif")
                    try:
                        chart_font_size = float(chart_cfg.get("font_size")) if chart_cfg.get("font_size") is not None else 14.0
                    except Exception:
                        chart_font_size = 14.0

                    layout_dict["title"] = dict(
                        text=chart_text,
                        x=0.5,
                        xref="paper",
                        xanchor="center",
                        font=dict(size=chart_font_size, family=chart_font_family, color="#333333"),
                    )
                else:
                    # 明确传入空标题时，移除 title 字段
                    layout_dict.pop("title", None)

            fig = go.Figure(traces, go.Layout(layout_dict))

            # 若需要叠加 CDF，则注入第二 Y 轴对应曲线
            if show_cdf and self.dist_obj:
                cdf_y = self._safe_cdf_vec(self.dist_obj, curve_x)
                DriskChartFactory._inject_overlay_cdf(
                    fig,
                    layout_dict,
                    curve_x,
                    cdf_y,
                    color=DRISK_COLOR_CYCLE_THEO[0]
                )
                fig.update_layout(layout_dict)

            plot_json = json.dumps(fig.to_dict(), cls=plotly.utils.PlotlyJSONEncoder)

            # ---------------------------------------------------
            # 5. 下发渲染指令
            # ---------------------------------------------------
            self._show_skeleton("正在绘图...")
            self._render_token += 1
            self._data_sent = True

            self._plot_host.load_plot(
                plot_json=plot_json,
                js_mode="histogram",
                js_logic="",
                initial_lr=None if getattr(self, "_use_qt_overlay", False)
                else (float(self.curr_left), float(self.curr_right)),
                static_plot=True,
            )

            self._after_plot_update()
            self._maybe_hide_skeleton()

        except Exception as e:
            # 先弹出错误提示
            self._show_histogram_error_popup(e)

            # 再尝试做一次“图内错误占位渲染”，避免用户看到空白区域
            try:
                if getattr(self, "_plot_host", None):
                    err_msg = f"直方图渲染失败：{type(e).__name__}: {e}"

                    fig = go.Figure()
                    l, r, t, b = self._computed_margins
                    fig.update_layout(
                        margin=dict(l=l, r=r, t=t, b=b),
                        annotations=[
                            dict(
                                text=err_msg,
                                x=0.5,
                                y=0.5,
                                xref="paper",
                                yref="paper",
                                showarrow=False
                            )
                        ]
                    )

                    plot_json = json.dumps(fig.to_dict(), cls=plotly.utils.PlotlyJSONEncoder)

                    self._show_skeleton("正在绘图...")
                    self._render_token += 1
                    self._data_sent = True

                    self._plot_host.load_plot(
                        plot_json=plot_json,
                        js_mode="histogram",
                        js_logic="",
                        initial_lr=None,
                        static_plot=True,
                    )

                    self._after_plot_update()
                    self._maybe_hide_skeleton()
            except Exception:
                pass

    # =======================================================
    # 4. 预览路由与分发中心
    # =======================================================
    def _is_compound_splice_preview(self) -> bool:
        dist_key = str(getattr(self, "dist_type", "") or "").lower()
        func_name = str(getattr(self, "_dist_func_name", "") or "").lower()
        cfg_func = ""
        try:
            cfg_func = str((getattr(self, "config", {}) or {}).get("func_name", "") or "").lower()
        except Exception:
            cfg_func = ""

        text = " ".join([dist_key, func_name, cfg_func])
        return "compound" in text or "splice" in text


    def _get_compound_splice_preview_samples(self, n: int = 3000) -> np.ndarray:
        dist = getattr(self, "dist_obj", None)
        if dist is None:
            return np.array([], dtype=float)

        rng = np.random.default_rng(20260422)
        qs = rng.uniform(1e-6, 1.0 - 1e-6, int(n))

        samples = []
        for q in qs:
            try:
                samples.append(float(self._safe_ppf_scalar(dist, float(q))))
            except Exception:
                pass

        arr = np.asarray(samples, dtype=float)
        arr = arr[np.isfinite(arr)]
        return arr

    def _render_channel_compound_splice_simulation(self, cmd: str = "pdf"):
        try:
            if not getattr(self, "_plot_host", None):
                return

            samples = self._get_compound_splice_preview_samples(3000)
            if samples.size <= 0:
                return

            x_min = float(np.percentile(samples, 1.0))
            x_max = float(np.percentile(samples, 99.0))

            if not math.isfinite(x_min) or not math.isfinite(x_max) or x_max <= x_min:
                x_min = float(np.min(samples))
                x_max = float(np.max(samples))

            if not math.isfinite(x_min) or not math.isfinite(x_max) or x_max <= x_min:
                x_min, x_max = 0.0, 1.0

            n_bins = 50
            bin_width = (x_max - x_min) / float(n_bins)
            if not math.isfinite(bin_width) or bin_width <= 0:
                bin_width = 1.0

            plot_samples = samples[(samples >= x_min) & (samples <= x_max)]
            if plot_samples.size <= 0:
                plot_samples = samples

            is_rel_freq = str(cmd or "").lower() == "rel_freq"
            show_cdf = str(cmd or "").lower().endswith("_cdf")

            histnorm = "percent" if is_rel_freq else "probability density"
            y_title = "\u76f8\u5bf9\u9891\u7387\uff08%\uff09" if is_rel_freq else "\u5bc6\u5ea6"

            counts, _edges = np.histogram(plot_samples, bins=n_bins, range=(x_min, x_max))
            if is_rel_freq:
                raw_y_max = float(np.max(counts) / max(float(plot_samples.size), 1.0) * 100.0)
            else:
                raw_y_max = float(np.max(counts) / max(float(plot_samples.size) * bin_width, 1.0))

            if not math.isfinite(raw_y_max) or raw_y_max <= 1e-12:
                raw_y_max = 1.0

            y_dtick = float(DriskMath.calc_smart_step(raw_y_max))
            if not math.isfinite(y_dtick) or y_dtick <= 0:
                y_dtick = raw_y_max / 5.0 if raw_y_max > 0 else 1.0

            y_axis_max = float(math.ceil((raw_y_max + 0.01 * y_dtick) / y_dtick) * y_dtick)

            x_dtick = float(DriskMath.calc_smart_step(x_max - x_min))
            if not math.isfinite(x_dtick) or x_dtick <= 0:
                x_dtick = 1.0

            self.view_min = float(x_min)
            self.view_max = float(x_max)
            self.x_dtick = float(x_dtick)

            axis_style = build_plotly_axis_style(
                x_range=(x_min, x_max),
                y_range=(0.0, y_axis_max),
                x_dtick=float(x_dtick),
                y_dtick=float(y_dtick),
                x_title=DriskChartFactory.VALUE_AXIS_TITLE,
                x_unit=DriskChartFactory.VALUE_AXIS_UNIT,
                y_title=y_title,
                fixedrange=True,
                forced_mag=getattr(self, "manual_mag", None),
                label_overrides=getattr(self, "_label_settings_config", None),
                x_axis_numeric=bool((getattr(self, "_label_axis_numeric", {}) or {}).get("x", True)),
                y_axis_numeric=bool((getattr(self, "_label_axis_numeric", {}) or {}).get("y", True)),
            )

            current_margin_r = self.MARGIN_R + (self.ADDITIONAL_Y2_MARGIN if show_cdf else 0)
            if hasattr(self, "slider"):
                self.slider.setMargins(self.MARGIN_L, current_margin_r)
            if hasattr(self, "overlay") and self.overlay and hasattr(self.overlay, "set_margins"):
                self.overlay.set_margins(self.MARGIN_L, current_margin_r, self.MARGIN_T, self.MARGIN_B)

            layout_dict = {
                "template": "plotly_white",
                "xaxis": axis_style.get("xaxis", {}),
                "yaxis": axis_style.get("yaxis", {}),
                "barmode": "overlay",
                "showlegend": False,
                "margin": dict(l=self.MARGIN_L, r=current_margin_r, t=self.MARGIN_T, b=self.MARGIN_B),
                "plot_bgcolor": "#fcfcfc",
                "paper_bgcolor": "#ffffff",
            }


            fig = go.Figure()
            fig.add_trace(go.Histogram(
                x=plot_samples,
                name="\u6a21\u62df\u6837\u672c",
                histnorm=histnorm,
                autobinx=False,
                xbins=dict(start=x_min, end=x_max, size=bin_width),
                marker=dict(
                    color=DriskChartFactory._hex_to_rgba(DRISK_COLOR_CYCLE_THEO[0], 0.72),
                    line=dict(width=0.5, color=DriskChartFactory._hex_to_rgba(DRISK_COLOR_CYCLE_THEO[0], 0.95)),
                ),
            ))

            if show_cdf:
                xs = np.sort(plot_samples)
                ys = np.arange(1, xs.size + 1, dtype=float) / float(xs.size)
                DriskChartFactory._inject_overlay_cdf(
                    fig,
                    layout_dict,
                    xs,
                    ys,
                    color=DRISK_COLOR_CYCLE_THEO[0],
                )

            fig.update_layout(layout_dict)
            plot_json = json.dumps(fig.to_dict(), cls=plotly.utils.PlotlyJSONEncoder)

            self._show_skeleton("\u6b63\u5728\u7ed8\u56fe...")
            self._render_token += 1
            self._data_sent = True

            self._plot_host.load_plot(
                plot_json=plot_json,
                js_mode="histogram",
                js_logic="",
                initial_lr=None,
                static_plot=True,
            )

            self._after_plot_update()
            self._maybe_hide_skeleton()

        except Exception as e:
            self._show_histogram_error_popup(e)

    def init_chart(self, y_max):
        """
        图表初始化入口。

        规则：
        - 若当前为离散型分布，则默认进入 PMF 柱状图；
        - 否则默认进入理论 PDF 曲线。

        这是在“未指定细分模式”时的初始图表路由。
        """
        try:
            if getattr(self, "is_discrete", False):
                self._render_channel_discreteprob(float(y_max))
            else:
                self._render_channel_pdf_theory(float(y_max))
        except Exception:
            pass

    def render_preview(self, y_max: float = None):
        """
        预览渲染总入口。

        调度逻辑：
        1. 读取当前界面下拉菜单返回的模式指令；
        2. 若未传入 y_max，则基于当前 y_data 自动估算；
        3. 按指令路由到 CDF / 相对频率直方图 / PDF / PMF；
        4. 若命令无法识别，则回退到 init_chart 自动模式。

        当前支持的典型命令包括：
        - "cdf"
        - "rel_freq"
        - "pdf"
        - "pdf_cdf"
        - "pmf"
        - "pmf_cdf"
        """
        try:
            if not self.dist_obj:
                return

            # 从界面读取当前预览模式指令
            # 例如："cdf"、"rel_freq"、"pdf"、"pdf_cdf"、"pmf" 等
            cmd = self.get_preview_mode()
            if self._is_compound_splice_preview():
                if cmd == "cdf":
                    self._render_channel_cdf()
                    return
                self._render_channel_compound_splice_simulation(cmd)
                return


            # 若外部未显式传入 y_max，则根据当前 y_data 自动估算
            if y_max is None:
                try:
                    if getattr(self, "y_data", None) is not None and len(self.y_data) > 0:
                        y_max = float(np.nanmax(self.y_data))
                    else:
                        y_max = 1.0
                except Exception:
                    y_max = 1.0

            # 1) 累积分布曲线
            if cmd == "cdf":
                self._render_channel_cdf()
                return

            # 2) 相对频率直方图
            if cmd == "rel_freq":
                self._render_channel_sample_hist(mode="rel_freq")
                return

            # 3) 概率密度曲线（也兼容如 pdf_cdf 这类以 pdf 开头的模式）
            if cmd.startswith("pdf"):
                self._render_channel_pdf_theory(float(y_max))
                return

            # 4) 离散概率质量柱状图（也兼容如 pmf_cdf 这类以 pmf 开头的模式）
            if cmd.startswith("pmf"):
                self._render_channel_discreteprob(float(y_max))
                return

            # 5) 兜底：回退到按分布类型自动推导的初始图表
            self.init_chart(float(y_max))

        except Exception:
            pass