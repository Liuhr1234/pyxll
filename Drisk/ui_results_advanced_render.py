"""
作为高级分析视图（摘要箱线图、敏感性龙卷风图、情景分析图）的核心渲染分发服务引擎。

模块核心定位：
- 本文件不直接做数据清洗；
- 不直接做统计计算；
- 不直接定义图表控件；
- 它承担的是“高级分析渲染调度中心”的角色。

核心价值在于“解耦与路由”：
1. 将主窗口（dialog）中杂糅的大量界面运行态信息，例如：
   - 配色体系
   - 手动量级设置
   - 情景分析参数
   - 龙卷风分析参数
   统一打包为轻量级的 Runtime State Snapshot（运行时快照）。

2. 将前置模块已经整理好的纯净数据载荷（Payload）与上述运行时状态组合，
   再根据 mode 指令精确路由到具体底层视图：
   - UIBoxplotView
   - UITornadoView

3. 避免主窗口直接充当“大而全控制器”，降低高级分析模式越来越多之后的维护成本。

为什么这个模块很重要：
- 高级分析模式很多，且每个模式的底层图表函数签名相似但不完全相同；
- 如果全部写在 dialog 内部，会导致：
  1) 分支极多
  2) 界面状态与算法逻辑缠绕
  3) 后续新增模式时极难维护
- 因此本文件将“渲染路由”单独抽离，形成稳定的中间调度层。

后续接手注意：
- 如果图表“画不出来”，优先检查三层：
  1) payload 是否已准备完整（ui_results_advanced_prepare.py）
  2) runtime 中的配置是否正常（ui_results_advanced_state.py）
  3) 下游视图方法签名是否与这里调用一致（ui_boxplot.py / ui_tornado.py）
- 若新增高级分析模式，通常应：
  1) 在 prepare 层确认 payload 已满足需要
  2) 在本文件新增 mode 路由
  3) 在底层视图中补充 render_xxx_chart(...)
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Any

from ui_results_advanced_prepare import AdvancedPreparedPayload
from ui_results_advanced_state import normalize_scenario_config, normalize_tornado_config

# =======================================================
# 1. 渲染运行时状态数据模型
# =======================================================
@dataclass
class ResultsAdvancedRuntimeState:
    """
    高级渲染运行时状态快照（Runtime State Snapshot）。

    设计目的：
    - 将底层视图在渲染高级图表时所需的 UI 运行态统一封装起来；
    - 避免直接把庞大的 dialog 实例原样传入每个底层视图；
    - 让 render 层只依赖“稳定数据结构”，降低耦合。

    为什么要做快照：
    - dialog 上往往挂着大量属性，既有界面控件，也有状态字段；
    - 底层绘图实际只需要其中少数关键字段；
    - 若直接把 dialog 到处传递，会让下游视图无边界地访问主窗口，
      后期非常难维护，也不利于排查“某个字段到底是谁在依赖”。

    字段说明：
    - series_style:
        标准多变量系列的颜色 / 样式映射。
        主要用于摘要图（箱线图、小提琴图、趋势图）这类“多 series 对比”的场景。

    - tornado_styles:
        龙卷风图 / 情景分析专用的颜色映射。
        通常用于区分高低分组、正负影响、不同方向的视觉编码。

    - tornado_line_styles:
        龙卷风相关“折线趋势图”专用样式。
        之所以单独拆出来，是因为线图与条形图对样式字段的需求可能不同。

    - scenario_config:
        情景分析参数字典。
        例如：阈值、尾部比例、分组规则等。
        会先经过 normalize_scenario_config 归一化，保证字段完整性与默认值安全。

    - tornado_config:
        龙卷风图基础配置字典。
        例如：分箱数量、算法参数等。
        会先经过 normalize_tornado_config 归一化。

    - manual_mag:
        用户在界面中手动选定的强制量级（例如 K / M / B）。
        用于控制不同高级图在数值展示尺度上的一致性。
    """
    series_style: dict
    tornado_styles: dict
    tornado_line_styles: dict
    scenario_config: dict
    tornado_config: dict
    manual_mag: Any


# =======================================================
# 2. 高级视图核心渲染调度服务
# =======================================================
class ResultsAdvancedRenderService:
    """
    高级分析图表渲染调度与执行服务中心。

    模块职责：
    1. 从 dialog 中提取渲染所需的运行时状态；
    2. 根据 mode 决定当前属于哪类高级分析；
    3. 将 payload + runtime 组合后，分发给底层视图执行渲染。

    本类本质上是一个“静态服务类”：
    - 不保存长期状态；
    - 不承担界面生命周期；
    - 只负责“按模式调对的方法”。

    这种设计的好处：
    - 新增模式时改动集中；
    - 调用入口统一；
    - dialog 中的复杂分支得以瘦身。
    """

    # 预定义摘要图（Boxplot 家族）的标准中文标题映射表。
    # 这里的 key 是内部 mode 指令，value 是界面标题栏上展示给用户的中文标题。
    _BOXPLOT_TITLE_MAP = {
        "boxplot": "箱形图",
        "letter_value": "字母值图",
        "violin": "小提琴图",
        "trend": "趋势图",
    }

    # ============================
    # Part 1: 运行时状态快照构建
    # ============================
    @staticmethod
    def build_runtime_state(dialog) -> ResultsAdvancedRuntimeState:
        """
        从 dialog 构建标准化运行时状态快照。

        处理逻辑：
        1. 从 dialog 上安全读取若干约定属性；
        2. 若属性缺失，则使用空 dict / None 作为默认值；
        3. 对 scenario_config / tornado_config 先做 normalize，
           以保证即便外层未设置配置，也能获得结构合法、带默认值的配置字典；
        4. 返回 ResultsAdvancedRuntimeState 实例。

        设计意图：
        - 将主窗口上的“散装字段”收敛为统一结构；
        - 避免下游 render_xxx 方法里反复写 getattr(dialog, ...)；
        - 让后续调试时更容易看清“当前渲染到底依赖了哪些运行态”。

        后续接手建议：
        - 若以后新增新的高级图表样式控制项，优先考虑先加到 runtime snapshot；
        - 不要直接让下游视图绕过 runtime 去从 dialog 上乱取字段。
        """
        return ResultsAdvancedRuntimeState(
            series_style=getattr(dialog, "_series_style", {}),
            tornado_styles=getattr(dialog, "_tornado_styles", {}),
            tornado_line_styles=getattr(dialog, "_tornado_line_styles", {}),
            scenario_config=normalize_scenario_config(getattr(dialog, "_scenario_config", None)),
            tornado_config=normalize_tornado_config(getattr(dialog, "_tornado_config", None)),
            manual_mag=getattr(dialog, "manual_mag", None),
        )

    # ============================
    # Part 2: 摘要图 (Boxplot) 家族渲染
    # ============================
    @staticmethod
    def render_boxplot(
        dialog,
        mode: str,
        runtime: ResultsAdvancedRuntimeState,
        display_label_map: dict[str, str] | None = None,
    ):
        """
        摘要图（Boxplot 家族）渲染分发器。

        适用模式示例：
        - boxplot
        - letter_value
        - violin
        - trend

        执行流程：
        1. 根据 mode 从 _BOXPLOT_TITLE_MAP 中匹配中文标题；
        2. 若 dialog 存在 chart_title_label，则同步更新标题文字；
        3. 调用 dialog.boxplot_view.render_boxplot(...) 执行底层绘制。

        参数说明：
        - dialog:
            主窗口实例。这里仍保留 dialog，是因为标题栏更新与视图对象获取都依赖主窗口。
        - mode:
            当前摘要图子模式。
        - runtime:
            已标准化的运行时快照。
        - display_label_map:
            可选的显示名映射表，用于将内部 key 替换为更友好的展示名称。

        注意：
        - 本方法不检查 series_map / display_keys 是否为空，默认认为上游已准备完毕；
        - 真正的图表构造与错误处理逻辑在 boxplot_view 内部。

        后续接手建议：
        - 若未来新增新的摘要图类型，通常只需：
          1) 在 _BOXPLOT_TITLE_MAP 加入标题映射
          2) 确保 boxplot_view.render_boxplot 支持该 plot_type
        """
        title = ResultsAdvancedRenderService._BOXPLOT_TITLE_MAP.get(mode, "分布对比分析")

        # 同步更新主界面标题栏，前提是主窗口暴露了 chart_title_label 控件
        if hasattr(dialog, "chart_title_label"):
            dialog.chart_title_label.setText(title)

        # 调用底层箱线图视图进行绘图
        dialog.boxplot_view.render_boxplot(
            series_map=dialog.series_map,
            display_keys=dialog.display_keys,
            style_map=runtime.series_style,
            title=title,
            plot_type=mode,
            forced_mag=runtime.manual_mag,
            display_label_map=display_label_map,
        )

    # ============================
    # Part 3: 情景分析 (Scenario) 渲染
    # ============================
    @staticmethod
    def render_scenario(dialog, payload: AdvancedPreparedPayload, runtime: ResultsAdvancedRuntimeState):
        """
        情景分析图渲染执行器。

        本方法职责非常单一：
        - 不做数据整理；
        - 不做配置补全；
        - 只把“已经准备好的 payload”与“标准化后的 runtime 配置”推给 tornado_view。

        payload 中通常已包含：
        - full_df:
            完整且已清洗、对齐好的分析数据表
        - output_display_name:
            输出变量展示名称
        - valid_input_cols:
            合法输入列名集合
        - xl_app:
            Excel 应用对象（若底层需要用于回查单元格/映射）
        - output_key:
            输出变量原始 key
        - input_mapping:
            输入变量映射关系

        这里选择复用 tornado_view.render_scenario_chart(...) 的原因：
        - 情景分析在视图表达上与龙卷风图同属“高级分析”家族；
        - 复用同一底层视图容器可以减少重复建设。

        后续接手建议：
        - 若情景分析未来需要新的配置字段，应优先在 normalize_scenario_config 中补足默认值；
        - 不建议在这里临时拼字段，以免 render 层逐渐承担 state 修补职责。
        """
        dialog.tornado_view.render_scenario_chart(
            payload.full_df,
            payload.output_display_name,
            payload.valid_input_cols,
            runtime.scenario_config,
            style_map=runtime.tornado_styles,
            xl_app=payload.xl_app,
            output_raw_key=payload.output_key,
            input_mapping=payload.input_mapping,
        )

    # ============================
    # Part 4: 敏感性分析 (Tornado) 家族渲染路由
    # ============================
    @staticmethod
    def render_tornado(dialog, mode: str, payload: AdvancedPreparedPayload, runtime: ResultsAdvancedRuntimeState):
        """
        敏感性分析（Tornado 家族）核心路由分发器。

        为什么这个方法最关键：
        - 敏感性分析子模式最多；
        - 不同模式虽然都基于同一份 full_df / 输入输出映射，
          但其底层算法视角不同、调用的方法也不同；
        - 因此这里承担“按 mode 精确分发”的职责。

        统一设计原则：
        - 先从 payload 中解包公共数据；
        - 再根据 mode 调不同的 tornado_view.render_xxx_chart(...)；
        - 保持下游方法签名尽量一致，降低维护成本。

        payload 解包字段说明：
        - full_df:
            已清洗好的分析主数据表
        - output_display_name:
            输出变量展示名称
        - valid_input_cols:
            可以进入敏感性分析的输入列
        - output_key:
            输出变量原始 key
        - input_mapping:
            输入变量映射关系
        - xl_app:
            Excel 上下文对象

        当前支持的 mode：
        1. reg
           多元回归系数龙卷风图
        2. reg_mapped
           带量级映射的多元回归龙卷风图
        3. spearman
           斯皮尔曼秩相关系数龙卷风图
        4. variance
           独立方差贡献率图
        5. bins_line
           多输入变量动态分箱折线趋势图
        6. default（兜底）
           传统高低分箱均值龙卷风图

        后续接手建议：
        - 若新增 mode，尽量遵守当前“一个 mode 对应一个清晰的底层 render 方法”；
        - 不建议在这里塞入大量算法逻辑；
        - render 层应始终只做“调度”，而非“计算”。
        """
        # ---------------------------------------------------
        # Step 1. 从标准数据载荷中解包公共参数
        # ---------------------------------------------------
        # 这样做的意义是：
        # - 保持下游调用更简洁；
        # - 让所有分支看到的输入命名一致；
        # - 后续若 payload 字段改名，这里是唯一集中调整点。
        full_df = payload.full_df
        output_display_name = payload.output_display_name
        valid_input_cols = payload.valid_input_cols
        output_key = payload.output_key
        input_mapping = payload.input_mapping
        xl_app = payload.xl_app

        # ---------------------------------------------------
        # 路由 1: 多元回归系数龙卷风图 (Regression)
        # ---------------------------------------------------
        # 使用回归系数反映各输入变量对输出的影响方向与相对强弱。
        if mode == "reg":
            dialog.tornado_view.render_regression_chart(
                full_df,
                output_display_name,
                valid_input_cols,
                style_map=runtime.tornado_styles,
                xl_app=xl_app,
                output_raw_key=output_key,
                input_mapping=input_mapping,
            )
            return

        # ---------------------------------------------------
        # 路由 2: 带量级映射的多元回归龙卷风图 (Regression Mapped)
        # ---------------------------------------------------
        # 与 reg 类似，但额外引入 manual_mag，用于统一数据显示量级。
        if mode == "reg_mapped":
            dialog.tornado_view.render_regression_mapped_chart(
                full_df,
                output_display_name,
                valid_input_cols,
                forced_mag=runtime.manual_mag,
                style_map=runtime.tornado_styles,
                xl_app=xl_app,
                output_raw_key=output_key,
                input_mapping=input_mapping,
            )
            return

        # ---------------------------------------------------
        # 路由 3: 斯皮尔曼秩相关系数龙卷风图 (Spearman)
        # ---------------------------------------------------
        # 适合非线性但单调关系的稳健相关性刻画。
        if mode == "spearman":
            dialog.tornado_view.render_spearman_chart(
                full_df,
                output_display_name,
                valid_input_cols,
                style_map=runtime.tornado_styles,
                xl_app=xl_app,
                output_raw_key=output_key,
                input_mapping=input_mapping,
            )
            return

        # ---------------------------------------------------
        # 路由 4: 独立方差贡献率分析 (Variance Contribution)
        # ---------------------------------------------------
        # 用于分析各输入对输出总体波动的贡献程度。
        if mode == "variance":
            dialog.tornado_view.render_variance_contribution_chart(
                full_df,
                output_display_name,
                valid_input_cols,
                style_map=runtime.tornado_styles,
                xl_app=xl_app,
                output_raw_key=output_key,
                input_mapping=input_mapping,
            )
            return

        # ---------------------------------------------------
        # 路由 5: 多输入变量动态分箱折线趋势图 (Bins Line Chart)
        # ---------------------------------------------------
        # 这一模式与传统条形龙卷风图不同，使用的是“折线趋势图”表达，
        # 因此样式不走 tornado_styles，而是走 tornado_line_styles。
        if mode == "bins_line":
            dialog.tornado_view.render_line_chart(
                full_df,
                output_display_name,
                valid_input_cols,
                runtime.tornado_config,
                style_map=runtime.tornado_line_styles,  # 注意：这里使用 line 专属样式
                xl_app=xl_app,
                output_raw_key=output_key,
                input_mapping=input_mapping,
            )
            return

        # ---------------------------------------------------
        # 路由 6（默认兜底）: 传统两极高低分箱均值龙卷风图
        # ---------------------------------------------------
        # 这是最传统的 tornado 形态：
        # - 依据配置对输入变量进行高低分组 / 分箱
        # - 比较不同分组对输出均值的影响
        #
        # 这里特别 copy 一份 tornado_config，而不是原地改 runtime.tornado_config，
        # 目的是避免把 forced_mag 写回共享配置，污染其他模式的渲染状态。
        config_copy = runtime.tornado_config.copy()
        config_copy["forced_mag"] = runtime.manual_mag

        dialog.tornado_view.render_chart(
            full_df,
            output_display_name,
            valid_input_cols,
            config_copy,
            style_map=runtime.tornado_styles,
            xl_app=xl_app,
            output_raw_key=output_key,
            input_mapping=input_mapping,
        )