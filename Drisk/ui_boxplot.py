# ui_boxplot.py
# -*- coding: utf-8 -*-
"""
本模块提供箱线图视图（Boxplot View）模块。

模块职责概述：
1. 视图初始化（UI Initialization）
   - 构建一个独立的 Qt 容器；
   - 内部包含用于承载 Plotly 图表的共享 WebEngineView；
   - 同时准备一个占位标签，用于显示“加载中”或“渲染失败”等状态。

2. 渲染调度（Render Dispatching）
   - 接收上层整理好的多情景数据；
   - 调用 DriskChartFactory 生成图表 JSON / JS 配置；
   - 再将结果下发给 PlotlyHost 执行真正的前端渲染。

3. 异常处理与反馈（Error Handling & Feedback）
   - 若图表渲染失败，不让界面空白；
   - 而是隐藏图表区，切换为提示文字，方便用户和开发者定位问题。

典型适用场景：
- 多个 simulation / series 的箱线图比较；
- 也可扩展为 violin 等同族分布图模式（由 plot_type 控制）。

后续接手注意：
- 本模块是“视图层容器”，不负责统计计算；
- 数据清洗、series_map 组织、样式字典生成应由更上层完成；
- 若后续出现“图不出来”，优先排查：
  1) series_map / display_keys / style_map 输入是否合法
  2) DriskChartFactory.build_boxplot_figure 是否报错
  3) PlotlyHost.load_plot 是否正常下发
"""
import traceback
from PySide6.QtWidgets import QWidget, QVBoxLayout, QLabel
from PySide6.QtCore import Qt

# 共享组件：
# - get_shared_webview：复用共享 WebEngineView，降低资源占用
# - get_plotly_cache_dir：获取 Plotly 相关缓存目录
from ui_shared import get_shared_webview, get_plotly_cache_dir

# Plotly 宿主控制器：
# 负责把图表 JSON / JS 逻辑真正送入 webview 执行
from plotly_host import PlotlyHost

# 绘图工厂：
# 负责将“结构化数据”转成 Plotly 所需配置
from drisk_charting import DriskChartFactory


class UIBoxplotView(QWidget):
    """
    [视图类] 专门用于展示箱线图的独立容器。

    定位说明：
    - 这是一个轻量视图组件；
    - 核心价值是“承接外部数据 -> 调工厂出图 -> 展示 / 报错反馈”；
    - 不负责业务逻辑判断，也不负责数据统计运算。
    """

    # =======================================================
    # 1. 视图初始化与 UI 构建
    # =======================================================
    def __init__(self, parent=None, tmp_dir=None):
        """
        初始化箱线图视图。

        参数：
        - parent：Qt 父组件
        - tmp_dir：Plotly 渲染缓存目录；若未传，则自动取共享缓存目录

        说明：
        - tmp_dir 传入的意义主要是兼容外层统一缓存策略；
        - 通常不需要每个视图单独自建目录。
        """
        super().__init__(parent)

        # 若外部未指定缓存目录，则使用共享 Plotly 缓存目录
        self.tmp_dir = tmp_dir or get_plotly_cache_dir()

        self._init_ui()

    def _init_ui(self):
        """
        构建基础 UI 结构。

        结构组成：
        1. 垂直布局容器
        2. 共享 WebEngineView：用于承载 Plotly 图表
        3. 占位标签 placeholder_label：用于“生成中 / 渲染失败”提示
        4. PlotlyHost：作为 Python 与前端图表之间的桥接宿主
        """
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        # 共享 webview：
        # 这样做通常比每次新建 QWebEngineView 更省资源，
        # 也有助于提升多个图表视图切换时的加载体验。
        self.web_view = get_shared_webview(self)
        layout.addWidget(self.web_view)

        # 占位标签：
        # 默认文案为“生成箱线图中...”，实际渲染成功后会隐藏；
        # 只有在加载中或报错时才展示。
        self.placeholder_label = QLabel("生成箱线图中...")
        self.placeholder_label.setAlignment(Qt.AlignCenter)
        self.placeholder_label.setStyleSheet("color: #888888; font-size: 14px;")
        self.placeholder_label.hide()
        layout.addWidget(self.placeholder_label)

        # Plotly 宿主：
        # 封装了将 JSON / JS 下发到 webview 的逻辑
        self.plot_host = PlotlyHost(self.web_view, tmp_dir=self.tmp_dir)


    # =======================================================
    # 2. 核心渲染调度入口
    # =======================================================
    def render_boxplot(
        self,
        series_map,
        display_keys,
        style_map,
        title="",
        plot_type="boxplot",
        forced_mag=None,
        display_label_map=None,
    ):
        """
        核心渲染调度入口。

        参数说明：
        - series_map：
            多组数据序列的主数据结构，通常是“系列名 -> 数值数组/列表”的映射。
        - display_keys：
            指定绘图时的系列展示顺序 / 目标系列集合。
        - style_map：
            每个系列的样式配置映射，如颜色、填充等。
        - title：
            图表标题。
        - plot_type：
            图形类型，默认是 "boxplot"；
            也可扩展为 "violin" 等同族模式，前提是绘图工厂支持。
        - forced_mag：
            强制数量级控制，用于统一显示尺度（如 k=10^3 这类量级控制）。
        - display_label_map：
            可选的“内部 key -> 最终展示标签”映射，用于图例 / 轴标签替换。

        渲染流程：
        1. 先隐藏 placeholder，显示 webview；
        2. 调用 DriskChartFactory.build_boxplot_figure(...) 构建图表配置；
        3. 再调用 PlotlyHost.load_plot(...) 将结果送往前端；
        4. 若任一环节报错，则切换为错误提示模式。

        注意：
        - 本方法默认假设输入数据已经由上层准备好；
        - 不在这里做复杂清洗，以保持视图层职责单一。
        """
        try:
            # 进入正常渲染态：
            # 若之前处于“报错占位”状态，这里先把占位隐藏掉，恢复 web 区显示。
            self.placeholder_label.hide()
            self.web_view.show()

            # 调用绘图工厂生成箱线图（或其变体）的 Plotly 配置
            res = DriskChartFactory.build_boxplot_figure(
                series_map=series_map,
                display_keys=display_keys,
                style_map=style_map,
                title=title,
                plot_type=plot_type,  # 向下传递图形模式，如 boxplot / violin
                forced_mag=forced_mag,
                display_label_map=display_label_map
            )

            # 将图表配置交给 PlotlyHost 执行真正的渲染
            self.plot_host.load_plot(
                res["plot_json"],
                res["js_mode"],
                res["js_logic"]
            )

        except Exception as e:
            # 渲染失败的兜底逻辑：
            # 1. 打印完整 traceback，方便开发调试
            # 2. 隐藏 webview，避免显示残缺页面
            # 3. 在占位标签上直观展示错误信息
            print(traceback.format_exc())

            self.web_view.hide()
            self.placeholder_label.setText(f"箱线图渲染失败: {str(e)}")
            self.placeholder_label.show()