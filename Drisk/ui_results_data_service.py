# ui_results_data_service.py
"""
本模块提供结果视图共享的无状态数据服务类 ResultsDataService。

模块定位：
1. 本模块位于结果分析链路中的“数据预处理与坐标参数准备”层，
   主要为上层结果弹窗、统计图形绘制逻辑、坐标轴计算逻辑提供通用辅助能力。
2. 本模块不直接负责界面渲染，也不维护界面状态；
   其职责是对原始数据进行标准化、统计摘要补充、展示参数推导，以及部分绘图基础点集构建。

主要功能：
1. 数据清洗与标准化：
   统一不同输入容器的数据结构，识别错误值，清洗非数值数据，并生成后续可安全计算的数值序列。
2. 数据质量统计：
   统计原始样本数、错误值数量、可用数值数量、被过滤数量，并通过 attrs 传递给上层界面。
3. 坐标轴参数推导：
   根据数据分布自动推导横轴展示范围、刻度步长、离散/连续展示模式等关键参数。
4. 累积分布点集构建：
   为 CDF（累积分布函数）绘图准备点集，在小样本下保留精确性，在大样本下兼顾渲染性能。

设计特点：
1. 无状态设计：
   全部方法均为静态方法，不依赖实例状态，便于在不同视图、不同调用链路中复用。
2. 面向结果展示层：
   输出内容以“安全可绘图”“便于界面使用”为目标，而非追求底层数学处理的完全通用性。
3. 强调容错：
   对空数据、异常输入、非数值内容、无穷值等情况均做了防御性处理，避免上层绘图链路直接报错。
"""

import math
import numpy as np
import pandas as pd

from typing import Optional as _Optional

from ui_shared import DriskMath


# =======================================================
# 1. 核心结果数据服务类（无状态工具集）
# =======================================================
class ResultsDataService:
    """
    结果数据服务辅助类。

    职责说明：
    1. 向结果分析界面提供通用数据预处理能力；
    2. 向坐标轴计算逻辑提供统一的输入数据基础；
    3. 向统计图形构建逻辑提供压缩后的绘图点集；
    4. 保证整个结果视图链路在面对异常数据时仍能稳定运行。

    设计说明：
    - 本类采用无状态设计，不持有实例变量；
    - 所有能力均通过静态方法暴露；
    - 适合被结果弹窗、绘图服务、后台任务结果整理逻辑直接调用。
    """

    # =======================================================
    # 1.1 原始输入遍历与基础统计
    # =======================================================
    @staticmethod
    def _iter_raw_values(raw):
        """
        内部辅助方法：统一原始输入的迭代方式。

        处理目标：
        - pandas.Series
        - numpy.ndarray
        - list / tuple
        - 单个标量
        - 其他不可预期对象

        处理策略：
        1. 若输入为常见容器类型，则尽量展平为一维列表；
        2. 若输入为单个值，则包装为单元素列表；
        3. 若中间转换过程异常，则退化为 [raw]，保证调用方仍可继续执行。

        说明：
        本方法的目标不是做严格类型校验，而是尽可能为后续统计逻辑提供“可遍历对象”。
        """
        try:
            if isinstance(raw, pd.Series):
                return list(raw.values)
            if isinstance(raw, np.ndarray):
                return list(np.asarray(raw, dtype=object).ravel())
            if isinstance(raw, (list, tuple)):
                return list(np.asarray(raw, dtype=object).ravel())
            return [raw]
        except Exception:
            return [raw]

    @staticmethod
    def _count_total_and_error(raw):
        """
        内部辅助方法：统计原始输入中的总样本数与错误值数量。

        统计内容：
        1. total：
           原始输入展开后的总元素数。
        2. error_count：
           以 '#' 开头的字符串数量。

        错误值识别规则：
        - 当前逻辑将以 '#' 开头的字符串视为错误值；
        - 这类值通常来自 Excel 或底层公式计算错误，
          例如：#DIV/0!、#NUM!、#VALUE! 等。

        说明：
        这里的 error_count 仅识别“显式错误标记”；
        其他无法转为数值的普通文本，会在后续 clean_series 中被归入过滤数量。
        """
        values = ResultsDataService._iter_raw_values(raw)
        total = 0
        error_count = 0
        for v in values:
            total += 1
            if isinstance(v, str):
                s = v.strip()
                if s.startswith("#"):
                    error_count += 1
        return int(total), int(error_count)

    # =======================================================
    # 2. 数据清洗与格式化
    # =======================================================
    @staticmethod
    def clean_series(raw) -> pd.Series:
        """
        将原始输入清洗为可安全计算的数值序列。

        主要步骤：
        1. 统计原始样本总数与错误值数量；
        2. 将常见容器输入展平；
        3. 强制转换为 pandas.Series；
        4. 使用 to_numeric 将非数值内容转为 NaN；
        5. 移除 NaN、正无穷、负无穷；
        6. 尝试统一为 float 类型；
        7. 将数据质量统计信息写入返回结果的 attrs。

        返回值：
        - 一个仅包含有效有限数值的 pandas.Series；
        - 若清洗后为空，则返回空 Series；
        - 返回的 Series.attrs 中包含上层界面可直接读取的统计信息。

        attrs 中写入的字段：
        - raw_total_count：原始总样本数
        - numeric_count：清洗后可用数值数
        - error_count：显式错误值数量
        - filtered_count：被过滤掉的非数值数量（不含已识别的 '#' 错误值）

        说明：
        该方法是结果分析链路中最基础的数据入口之一。
        上层若需要绘图、统计或展示样本质量信息，通常都应优先经过该方法。
        """
        # 先统计原始输入中的样本总数与显式错误值数量，
        # 便于后续构建完整的数据质量摘要。
        total_count, error_count = ResultsDataService._count_total_and_error(raw)
        try:
            # 对常见容器做一次展平处理，避免出现多维数组或嵌套结构
            # 影响 pandas.Series 的后续转换效果。
            if isinstance(raw, (list, tuple, np.ndarray)):
                raw = np.asarray(raw).ravel()
        except Exception:
            # 若展平失败，则保留原始输入继续向下处理。
            pass

        # 将输入统一转为数值序列：
        # - 合法数值将被保留；
        # - 普通文本、非法对象等无法解析为数值的内容将被转为 NaN。
        s = pd.to_numeric(pd.Series(raw), errors="coerce")

        # 进一步剔除正无穷、负无穷和空值，确保下游统计与绘图逻辑安全。
        s = s.replace([np.inf, -np.inf], np.nan).dropna()

        try:
            # 尝试统一为 float，方便后续百分位、坐标轴、CDF 等数值计算逻辑。
            s = s.astype(float)
        except Exception:
            # 若类型转换失败，则保留当前结果，不中断主流程。
            pass

        numeric_count = int(len(s))

        # filtered_count 表示：
        # 原始总样本数 - 显式错误值数 - 最终成功保留的数值数。
        # 该值主要对应普通文本、空字符串、无法解析的对象等“被过滤内容”。
        filtered_count = max(0, int(total_count) - int(error_count) - int(numeric_count))

        # 将数据质量摘要附着到 Series.attrs，供上层界面直接读取。
        # 例如：
        # - 结果弹窗中的数据质量说明
        # - 清洗摘要提示
        # - 调试输出或日志记录
        try:
            attrs = dict(getattr(s, "attrs", {}) or {})
            attrs["raw_total_count"] = int(total_count)
            attrs["numeric_count"] = int(numeric_count)
            attrs["error_count"] = int(error_count)
            attrs["filtered_count"] = int(filtered_count)
            s.attrs = attrs
        except Exception:
            # attrs 写入失败不应影响主结果返回。
            pass

        return s

    # =======================================================
    # 3. 坐标轴与展示视图计算
    # =======================================================
    @staticmethod
    def calc_x_axis_params(data_for_axis: pd.Series, detect_series: _Optional[pd.Series] = None) -> dict:
        """
        根据输入数据自动推导横轴展示参数。

        输入参数：
        1. data_for_axis：
           用于实际计算横轴展示范围的数值序列。
        2. detect_series：
           用于辅助判断数据应采用离散视图还是连续视图的检测序列。
           若未提供，则默认使用 data_for_axis 自身进行检测。

        返回结果字典包含：
        - abs_min / abs_max：
          原始绝对最小值、最大值。
        - p1 / p99：
          1% 与 99% 分位数，用于上层做额外判断或展示。
        - is_discrete_view：
          当前数据是否建议按离散模式展示。
        - _discrete_step：
          离散模式下推导出的步长。
        - x_dtick：
          横轴主刻度步长。
        - x_range_min / x_range_max：
          对齐后的横轴显示范围。
        - view_min / view_max：
          与显示范围一致的快捷字段，便于上层统一读取。

        核心逻辑：
        1. 先获取原始极值与分位数；
        2. 再根据唯一值数量、唯一值占比、整数性、网格步长特征等判断是否适合离散展示；
        3. 最后基于智能步长算法与“呼吸边距”规则生成横轴显示范围。

        说明：
        本方法是结果图横轴展示逻辑的核心入口之一，
        其输出通常会被直方图、CDF、离散分布图等多个绘图分支复用。
        """
        # 容错处理：
        # 若输入为空，则使用默认区间 [0.0, 1.0] 作为兜底，
        # 避免后续最值、百分位等计算直接报错。
        if data_for_axis is None or len(data_for_axis) == 0:
            data_for_axis = pd.Series([0.0, 1.0])

        # 计算原始绝对边界与关键分位数。
        # 这些值既用于坐标轴推导，也可能被上层用于数据摘要显示。
        abs_min = float(data_for_axis.min())
        abs_max = float(data_for_axis.max())
        p1 = float(np.percentile(data_for_axis, 1))
        p99 = float(np.percentile(data_for_axis, 99))

        # 离散性检测优先使用 detect_series。
        # 这样可以支持“坐标轴按 A 数据绘制，但离散/连续判断按 B 数据推导”的场景。
        det_series = detect_series if detect_series is not None else data_for_axis
        arr_det = np.asarray(det_series.values, dtype=float)
        arr_det = arr_det[np.isfinite(arr_det)]
        n = arr_det.size if arr_det is not None else 0

        # 若检测序列为空，则同样使用默认值兜底，
        # 避免 unique、比例判断等逻辑失效。
        if n == 0:
            arr_det = np.array([0.0, 1.0])
            n = 2

        # 计算唯一值集合及其规模，
        # 用于判断数据是否具有离散型特征。
        u = np.unique(arr_det)
        unique = u.size
        unique_ratio = unique / max(1, n)

        # 判断数据是否“近似全为整数”。
        # 这里使用容差判断，而非绝对相等，
        # 以适应浮点存储误差带来的微小偏差。
        is_int = np.all(np.isclose(arr_det, np.round(arr_det), atol=1e-6, rtol=0.0))

        # 对非整数数据，进一步尝试识别其是否落在固定网格上，
        # 例如 0.5、0.1、0.05 等常见步长。
        # 若大多数点都满足某一网格步长，可将其视为“离散化数轴上的取值”。
        grid_step = None
        if not is_int:
            candidates = [0.5, 0.2, 0.1, 0.05, 0.02, 0.01, 0.005, 0.002, 0.001]
            for q in candidates:
                scaled = arr_det / q
                ok = np.mean(np.isclose(scaled, np.round(scaled), atol=1e-6, rtol=0.0))
                if ok >= 0.995:
                    grid_step = float(q)
                    break

        # 若唯一值不少于 2，则尝试从相邻唯一值差中推导稳定步长。
        # 这里使用中位数作为代表值，并检查大多数差值是否接近该中位数，
        # 以识别“等步长或近似等步长”的离散数据。
        step = None
        if unique >= 2:
            diffs = np.diff(u)
            diffs = diffs[diffs > 1e-12]
            if diffs.size > 0:
                step0 = float(np.median(diffs))
                step0 = float(np.round(step0, 12))
                ok = np.mean(np.isclose(diffs, step0, rtol=1e-4, atol=1e-9))
                if ok >= 0.90:
                    step = step0

        # 综合判断是否采用离散视图。
        #
        # 判断依据包括：
        # 1. 唯一值占比是否足够低；
        # 2. 唯一值总数是否较少；
        # 3. 数据是否近似整数；
        # 4. 是否存在稳定网格步长；
        # 5. 是否存在较稳定的唯一值差分步长。
        #
        # 这里的阈值采用经验规则，
        # 目标是让结果图在“可读性”和“自动识别稳定性”之间取得平衡。
        is_discrete_view = (
                (unique_ratio <= 0.2 and 1000 <= n < 100000)
                or (unique_ratio <= 0.1 and 100000 <= n < 1000000)
                or (unique_ratio <= 0.05 and 1000000 <= n)
                or (unique_ratio <= 0.3 and n < 1000)
                or (unique <= 30)
                or is_int
                or (grid_step is not None)
                or (step is not None and unique <= 300)
        )

        # 推导离散步长：
        # 1. 若全为整数，则步长优先取 1；
        # 2. 否则优先使用识别到的网格步长；
        # 3. 若无网格步长，则尝试使用唯一值差分步长；
        # 4. 最终兜底为 1.0。
        _discrete_step = 1.0 if is_int else (grid_step if grid_step is not None else (step if step is not None else 1.0))

        # 计算原始跨度。
        # 当最小值与最大值几乎相同，需强制给一个非零跨度，
        # 否则智能步长计算与坐标轴展开都会失去意义。
        span_core = abs_max - abs_min
        if span_core < 1e-9:
            span_core = 1.0

        # 调用统一数学工具计算“智能刻度步长”。
        # 该步长用于横轴刻度显示，也是后续范围对齐的基准。
        dtick = float(DriskMath.calc_smart_step(span_core))
        if not np.isfinite(dtick) or dtick <= 0:
            dtick = 1.0

        # 根据离散/连续模式分别生成最终显示范围。
        if is_discrete_view:
            # 离散视图下的处理顺序为：
            # 1. 先得到刻度步长；
            # 2. 再按刻度步长的 1% 增加轻微呼吸边距；
            # 3. 最后把范围向刻度边界对齐。
            #
            # 这样可以避免离散端点贴边，同时保持坐标轴刻度整齐。
            x_dtick = dtick
            breathing = 0.01 * x_dtick
            expanded_min = abs_min - breathing
            expanded_max = abs_max + breathing
            x_range_min = math.floor(expanded_min / x_dtick) * x_dtick
            x_range_max = math.ceil(expanded_max / x_dtick) * x_dtick

            # 若极端情况下对齐后区间无效，则至少补足一个刻度单位。
            if x_range_max <= x_range_min:
                x_range_max = x_range_min + x_dtick
        else:
            # 连续视图下同样增加 1% 步长的轻微边距，
            # 再按步长进行下取整/上取整对齐。
            x_dtick = dtick
            padded_min = abs_min - 0.01 * dtick
            padded_max = abs_max + 0.01 * dtick
            x_range_min = math.floor(padded_min / dtick) * dtick
            x_range_max = math.ceil(padded_max / dtick) * dtick

        # 返回统一结构的坐标轴参数字典，
        # 供上层绘图、状态保存、界面显示等逻辑直接使用。
        return {
            "abs_min": abs_min, "abs_max": abs_max, "p1": p1, "p99": p99,
            "is_discrete_view": is_discrete_view, "_discrete_step": _discrete_step,
            "x_dtick": x_dtick, "x_range_min": x_range_min, "x_range_max": x_range_max,
            "view_min": x_range_min, "view_max": x_range_max
        }

    # =======================================================
    # 4. 统计图形算法
    # =======================================================
    @staticmethod
    def build_cdf_points(arr: np.ndarray, *, discrete: bool = False):
        """
        为累积分布函数（CDF）绘图构建点集。

        输入参数：
        1. arr：
           原始数值数组。
        2. discrete：
           是否按离散数据模式构建 CDF 点集。

        返回值：
        - xs：横轴点
        - ys：对应的累计概率
        - 若输入为空，则返回 (None, None)

        构建策略：
        1. 先过滤非有限数值并排序；
        2. 若为离散数据，则按支持点（support points）生成阶梯式 CDF；
        3. 若为连续数据：
           - 小样本时保留逐点精确结果；
           - 大样本时压缩为固定数量分位点，以控制前端渲染压力。

        设计考虑：
        - 小样本阶段优先保留精确性；
        - 大样本阶段优先兼顾可视化平滑度与性能；
        - 离散数据必须保持“阶梯语义”，不能简单按连续曲线处理。
        """
        # 强制转为浮点数组，并过滤掉非有限值。
        arr = np.asarray(arr, dtype=float)
        arr = arr[np.isfinite(arr)]
        n = arr.size

        # 空数组直接返回空结果，交由上层决定是否跳过绘图。
        if n == 0:
            return None, None

        # CDF 的构建必须以排序后的数据为基础。
        arr.sort()

        if discrete:
            # 离散数据下，横轴应取唯一支持点，
            # 而不是保留所有重复值。
            # 这样得到的 CDF 才能正确体现“在每个支持点发生跃升”的阶梯特征。
            xs = np.unique(arr)
            if xs.size == 0:
                return None, None

            # 对每个支持点，统计小于等于该点的样本占比，
            # 作为对应的累计概率。
            ys = np.searchsorted(arr, xs, side="right").astype(float) / float(n)
            return xs, ys

        if n >= 1001:
            # 大样本压缩策略：
            # 使用 0.0 到 1.0 的 1001 个等间距分位点，
            # 通过线性分位数计算得到横轴点。
            #
            # 优点：
            # 1. 点数固定，便于前端稳定渲染；
            # 2. 在视觉上仍能较好保留整体分布形态；
            # 3. 避免超大样本逐点绘图带来的性能问题。
            q = np.linspace(0.0, 1.0, 1001)
            xs = np.quantile(arr, q, method="linear")
            ys = q
            return xs, ys
        else:
            # 小样本精确策略：
            # 直接使用排序后的样本值作为横轴，
            # 使用 1/n, 2/n, ..., n/n 作为累计概率。
            #
            # 这样可以完整保留每个样本点的阶梯变化细节。
            xs = arr
            ys = np.arange(1, n + 1, dtype=float) / n
            return xs, ys