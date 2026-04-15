# ui_workers.py
"""
本模块提供 UI 后台任务（Worker Threads）的处理逻辑，属于第 3 组（UI）
与第 2 组（底层架构）完成解耦后的“后台计算层”。

模块总体定位
--------------------------------
它的核心职责不是“画图”，也不是“管理界面控件”，而是：
1. 在后台线程里完成相对耗时的统计 / KDE / CDF / 理论统计计算；
2. 将计算结果通过 Qt Signal 安全回传给 UI 主线程；
3. 在 UI 层与底层 cache / theory 体系之间建立一层稳定、口径一致的桥接。

为什么本模块重要
--------------------------------
UI 中很多统计与图表更新都具有以下特征：
- 计算量不一定大，但可能频繁触发；
- 若直接在主线程做，会造成界面卡顿；
- 且结果分析器与建模器都依赖一致的统计口径，不能各算各的。

因此本模块承担三项关键目标：
1. 数据解耦：
   UI 不再把自己维护的数据当成唯一真源；
   优先通过 backend_bridge 从第 2 组架构的缓存中读取数据与统计结果。

2. 统计口径对齐：
   例如：
   - 标准差 ddof=1
   - 峰度用 Pearson 峰度（fisher=False）
   - 分位数键名规范化
   - 过滤规则与错误计数规则一致
   这些都尽量与底层保持统一。

3. 核心计算能力保留：
   - 保留 UI 侧成熟的 KDE 核密度估计能力；
   - 理论分布则通过 theory_adapter 与 backend_bridge 获取统一语义，
     避免 UI 自己再造一套分布解释逻辑。

本模块包含 4 类主要后台任务（QRunnable）
--------------------------------
1) ModelerStatsJob
   - 建模器理论统计任务；
   - 面向 ui_modeler；
   - 处理理论均值、方差、分位数、动态尾部概率等。

2) StatsJob
   - 结果分析器实证统计任务；
   - 面向 ui_results；
   - 处理模拟样本的统计汇总（均值、标准差、分位数、错误数等）。

3) KDEJob
   - KDE 密度曲线计算任务；
   - 面向结果分析器密度图 / PDF 视图等。

4) CDFStatsWorker
   - 累积概率模式下的统计任务；
   - 除了基础统计外，还可额外计算均值置信区间误差范围。

后续接手人优先关注
--------------------------------
如果发现“统计值不一致 / 图画不出来 / 某些模式下行数据为空”，优先排查：
1. backend_bridge 是否可正常读取 series / statistics；
2. theory_adapter 是否能正确返回理论统计；
3. stats_contract.normalize_stats_map / expand_legacy_aliases 是否覆盖到当前消费者需要的键；
4. _filter_valid_float_array 是否把输入数据过滤得过多；
5. Worker 的 cancel_check 是否过早中断任务。
"""
from __future__ import annotations

import numpy as np
import traceback
from typing import Any, Callable, Dict, List, Optional, Sequence
from scipy.stats import gaussian_kde
from scipy.stats import skew as _skew
from scipy.stats import kurtosis as _kurtosis
from scipy.stats import t as _t

from PySide6.QtCore import QObject, Signal, QRunnable

import ui_stats
from ui_stats import assemble_cdf_rows, smart_format_number
import ui_stats_contract as stats_contract
from ui_results_data_service import ResultsDataService
from statistical_functions import _smart_ci_engine


# ============================================================
# 1. 模块级导入与后端桥接服务
# ============================================================

# --- 后端数据桥 (Backend Bridge) ---
# 作用：
# - 作为 UI 访问底层缓存 / 统计真源的统一入口；
# - 若导入失败，说明“新架构桥接层”不可用，此时部分新版 Worker 功能会退化或直接不可用。
#
# 注意：
# - StatsJob / ModelerStatsJob / CDFStatsWorker 的新版路径都依赖 bridge；
# - 若此处失败，通常不是本文件逻辑错，而是 backend_bridge 模块接入异常。
try:
    import backend_bridge as bridge
except Exception as _e:
    bridge = None
    print(f"[Drisk][ui_workers] backend_bridge 导入失败: {_e}")

# --- 理论分布适配器 (Theory Adapter) ---
# 作用：
# - 让 UI 层能够通过统一接口拿到“理论分布统计量”；
# - 对于部分 SciPy 分布、平移/截断等语义，由 theory_adapter 负责统一解释；
# - 避免 UI 自己拼一套理论统计逻辑，导致与底层脱节。
try:
    import theory_adapter as theory
except Exception as _e:
    theory = None
    print(f"[Drisk][ui_workers] theory_adapter 导入失败: {_e}")


# ============================================================
# 2. 公共工具与数据处理逻辑
# ============================================================

# Excel / 仿 Excel 链路中常见的错误占位标记。
# 当前约定为 "#ERROR!"；后续若产品内错误标记体系调整，这里需要同步。
ERROR_MARKER = "#ERROR!"


def _filter_valid_float_array(data: Any) -> np.ndarray:
    """
    [数据清洗]
    将输入数据过滤为仅包含“有限浮点数”的 NumPy 数组。

    过滤规则（尽量与底层保持一致）：
    1. 跳过 None
    2. 跳过错误标记 "#ERROR!"
    3. 跳过 NaN / Inf / -Inf
    4. 字符串尝试转为 float，失败则跳过
    5. 非可迭代标量也会尝试兜底转换

    返回：
    - dtype=float 的一维 numpy 数组
    - 若无有效值，则返回空数组

    为什么这是关键函数：
    - 几乎所有经验统计、KDE、CDF 计算都要先经过它；
    - 一旦这里过滤规则变动，会影响全局统计口径；
    - 因此不要轻易在其他地方另写一套“只过滤部分脏值”的逻辑。
    """
    if data is None:
        return np.array([], dtype=float)

    out: List[float] = []
    try:
        # 兼容 pandas.Series / numpy.ndarray / list / tuple 等可迭代容器
        for v in data:
            if v is None:
                continue

            if isinstance(v, str):
                if v == ERROR_MARKER:
                    continue
                try:
                    fv = float(v)
                except Exception:
                    continue
            else:
                try:
                    fv = float(v)
                except Exception:
                    continue

            if not np.isfinite(fv):
                continue
            out.append(float(fv))
    except Exception:
        # 非可迭代对象的兜底路径，例如单个数值标量
        try:
            fv = float(data)
            if np.isfinite(fv):
                out.append(float(fv))
        except Exception:
            pass

    return np.asarray(out, dtype=float)


def _to_non_negative_int(value: Any) -> Optional[int]:
    """
    [类型转换]
    尝试将输入转为“非负整数”。

    返回规则：
    - 成功且 >= 0：返回 int
    - 失败或小于 0：返回 None

    使用场景：
    - 解析 attrs / backend stats 中的 count、error_count、filtered_count 等字段；
    - 这些字段理论上都应该是非负整数。
    """
    try:
        iv = int(value)
    except Exception:
        return None
    return iv if iv >= 0 else None


def _fallback_count_total_and_error(data: Any) -> tuple[int, int]:
    """
    [降级处理]
    回退统计方法：在底层服务不可用或失败时，手动统计：
    1. 总元素个数 total
    2. 错误值个数 error_count（规则：以 '#' 开头的字符串视为错误）

    注意：
    - 这是“保底逻辑”，优先级低于 ResultsDataService._count_total_and_error；
    - 这里只能做到粗粒度识别，不能保证比底层更精细。
    """
    values: List[Any]
    if data is None:
        values = []
    else:
        try:
            if hasattr(data, "values"):
                values = list(np.asarray(getattr(data, "values"), dtype=object).ravel())
            elif isinstance(data, np.ndarray):
                values = list(np.asarray(data, dtype=object).ravel())
            elif isinstance(data, (list, tuple)):
                values = list(np.asarray(data, dtype=object).ravel())
            else:
                values = list(data)
        except Exception:
            values = [data]

    total = int(len(values))
    error_count = 0
    for v in values:
        if isinstance(v, str) and v.strip().startswith("#"):
            error_count += 1
    return total, int(error_count)


def _extract_error_and_filtered_count(data: Any, numeric_count: int) -> tuple[int, int]:
    """
    [数据解析]
    结合 data.attrs 或底层服务，提取：
    - error_count：错误值数量
    - filtered_count：被过滤值数量

    优先级：
    1. 若 data.attrs 中已经有 numeric_count / error_count / filtered_count，
       且 numeric_count 与当前清洗结果一致，则优先信任 attrs。
    2. 否则调用 ResultsDataService._count_total_and_error(...)。
    3. 若仍失败，则退回 _fallback_count_total_and_error(...)。

    filtered_count 的计算逻辑：
        total_count - error_count - numeric_count

    这里的意义：
    - 让 UI 表格中的“数值数 / 错误数 / 过滤数”口径尽量统一；
    - 避免单纯依据清洗后的数值数组，丢失“原始总量”信息。
    """
    attrs = getattr(data, "attrs", None)
    if isinstance(attrs, dict):
        numeric_attr = _to_non_negative_int(attrs.get("numeric_count"))
        error_attr = _to_non_negative_int(attrs.get("error_count"))
        filtered_attr = _to_non_negative_int(attrs.get("filtered_count"))
        if (
            numeric_attr is not None
            and error_attr is not None
            and filtered_attr is not None
            and numeric_attr == int(max(0, numeric_count))
        ):
            return int(error_attr), int(filtered_attr)

    total_count: Optional[int] = None
    error_count: Optional[int] = None
    try:
        total_count, error_count = ResultsDataService._count_total_and_error(data)
    except Exception:
        total_count, error_count = _fallback_count_total_and_error(data)

    total_count = int(max(0, total_count if total_count is not None else 0))
    error_count = int(max(0, error_count if error_count is not None else 0))
    filtered_count = max(0, total_count - error_count - int(max(0, numeric_count)))
    return error_count, int(filtered_count)


def _normalize_series_kind(series_kind: Any) -> str:
    """
    [数据规范化]
    将 series_kind 统一规范为：
    - "input"
    - "output"

    任何非 "input" 的值都回退为 "output"。

    使用场景：
    - bridge.get_series(...) / bridge.get_statistics(...) 需要区分读取输入序列还是输出序列；
    - UI 外层有时传入大小写不一致或空值，这里统一收口。
    """
    kind = str(series_kind or "output").strip().lower()
    return "input" if kind == "input" else "output"


def _override_stats_with_backend(
    stats: Dict[Any, Any],
    backend_stats: Optional[Dict[Any, Any]],
    raw_data: Any,
) -> Dict[Any, Any]:
    """
    [合并逻辑]
    用底层后端传来的统计结果覆盖本地统计结果，以尽量保证“真源优先”。

    处理流程：
    1. 若 backend_stats 是合法 dict，则先通过 stats_contract.normalize_stats_map(...) 标准化；
    2. 对以下核心键进行覆盖：
       - count
       - min / max / mean / median / std
       - 所有百分位键 pxx
    3. 若 backend 返回了 count，则进一步重算 error_count / filtered_count；
    4. 最后调用 expand_legacy_aliases(...)，补齐旧版别名，兼容历史消费者。

    为什么要做“后端覆盖”：
    - UI 本地可自行算出一套统计量，但底层可能有更精确、更统一的真源口径；
    - 特别是在缓存 / collect / 过滤规则参与后，底层统计往往更值得信赖。

    注意：
    - 这里不是“完全替换 stats”，而是重点覆盖核心字段；
    - 某些 UI 专属字段（例如 dyn_pl 之类）仍可能保留本地计算结果。
    """
    if not isinstance(stats, dict):
        stats = {}

    source = backend_stats if isinstance(backend_stats, dict) else {}
    if source:
        normalized = stats_contract.normalize_stats_map(
            source,
            percentiles=tuple(ui_stats.DEFAULT_PERCENTILES),
        )
        for key, value in normalized.items():
            if key in (
                stats_contract.COUNT_KEY,
                stats_contract.MIN_KEY,
                stats_contract.MAX_KEY,
                stats_contract.MEAN_KEY,
                stats_contract.MEDIAN_KEY,
                stats_contract.STD_KEY,
            ) or str(key).startswith("p"):
                stats[key] = value

    backend_count = _to_non_negative_int(stats.get(stats_contract.COUNT_KEY))
    if backend_count is not None:
        total_count: Optional[int] = None
        error_count: Optional[int] = None
        try:
            total_count, error_count = ResultsDataService._count_total_and_error(raw_data)
        except Exception:
            total_count, error_count = _fallback_count_total_and_error(raw_data)
        total_count = int(max(0, total_count if total_count is not None else 0))
        error_count = int(max(0, error_count if error_count is not None else 0))
        stats[stats_contract.ERROR_COUNT_KEY] = int(error_count)
        stats[stats_contract.FILTERED_COUNT_KEY] = int(max(0, total_count - error_count - int(backend_count)))

    # 补回历史别名，兼容旧版 UI / 表格消费者
    stats_contract.expand_legacy_aliases(stats, percentiles=tuple(ui_stats.DEFAULT_PERCENTILES))
    return stats


def _mean_ci_half_width(arr: np.ndarray, confidence_level: float = 0.90) -> float:
    """
    [统计计算]
    计算“均值置信区间”的半宽（±margin）。

    核心公式：
        margin = t_critical * (std(ddof=1) / sqrt(n))

    口径说明：
    - 使用样本标准差 ddof=1；
    - 使用 t 分布分位点；
    - 当样本数小于 2 或标准差异常时，返回 NaN。

    适用场景：
    - 结果分析界面中“均值 ± 误差范围”的展示；
    - 建模器理论统计不走这里，理论统计另有独立路径。
    """
    arr = _filter_valid_float_array(arr)
    n = int(arr.size)
    if n < 2:
        return float("nan")

    std = float(np.std(arr, ddof=1))
    if not np.isfinite(std):
        return float("nan")

    std_err = std / float(np.sqrt(n))
    alpha = 1.0 - float(confidence_level)

    try:
        t_critical = float(_t.ppf(1.0 - alpha / 2.0, df=n - 1))
    except Exception:
        return float("nan")

    return float(t_critical * std_err)


def _calculate_empirical_stats(
    data: Any,
    *,
    left_x: Optional[float] = None,
    right_x: Optional[float] = None,
) -> Dict[Any, float]:
    """
    [统计计算]
    计算一组“经验样本”的统计指标，返回字典供 UI 表格与图表使用。

    当前输出可能包含：
    - count / error_count / filtered_count
    - min / max / mean / median / std
    - skew / kurt
    - 分位数 pxx
    - ci90
    - 动态尾部概率 dyn_pl / dyn_lx / dyn_pm / dyn_rx / dyn_pr

    统计口径说明：
    - std：样本标准差 ddof=1
    - kurt：Pearson 峰度（fisher=False, bias=False）
    - skew：SciPy 默认偏度口径
    - 分位数：使用 ui_stats.DEFAULT_PERCENTILES 中定义的百分位点

    动态尾部概率说明：
    - 若提供 left_x / right_x，则根据当前滑块分割点实时计算：
      左侧概率、右侧概率、中间概率。
    - 这与结果分析界面中的左右游标交互直接相关。

    注意：
    - 本函数是经验统计的核心公共方法；
    - StatsJob / CDFStatsWorker 都依赖它；
    - 如果修改输出字段，要同步检查 ui_stats / stats_contract / 表格渲染路径。
    """
    arr = _filter_valid_float_array(data)
    stats: Dict[Any, float] = {}
    n = int(arr.size)
    stats["count"] = n
    error_count, filtered_count = _extract_error_and_filtered_count(data, n)
    stats["error_count"] = int(error_count)
    stats["filtered_count"] = int(filtered_count)

    if n == 0:
        stats_contract.expand_legacy_aliases(stats, percentiles=tuple(ui_stats.DEFAULT_PERCENTILES))
        return stats

    stats["min"] = float(np.min(arr))
    stats["max"] = float(np.max(arr))
    stats["mean"] = float(np.mean(arr))
    stats["median"] = float(np.median(arr))

    stats["std"] = float(np.std(arr, ddof=1)) if n > 1 else 0.0
    stats["skew"] = float(_skew(arr)) if n > 2 else 0.0
    stats["kurt"] = float(_kurtosis(arr, fisher=False, bias=False)) if n > 3 else 0.0

    # 计算默认百分位组，并写入规范化 percentile 键
    pct_keys = list(ui_stats.DEFAULT_PERCENTILES)
    try:
        pct_vals = np.percentile(arr, [p * 100.0 for p in pct_keys])
        for p, v in zip(pct_keys, pct_vals):
            val = float(v)
            stats[stats_contract.canonical_percentile_key(float(p))] = val
    except Exception:
        pass

    # 计算均值 90% 置信区间半宽，UI 往往显示为 ± 值
    stats["ci90"] = _mean_ci_half_width(arr, confidence_level=0.90)

    # 动态尾部概率（由左右滑块坐标触发）
    if left_x is not None and right_x is not None and n > 0:
        try:
            lx = float(left_x)
            rx = float(right_x)
            c_left = int(np.sum(arr <= lx))
            c_right = int(np.sum(arr > rx))
            pl = c_left / n
            pr = c_right / n
            pm = 1.0 - pl - pr
            if pm < 0:
                pm = 0.0
            stats["dyn_pl"] = float(pl)
            stats["dyn_lx"] = float(lx)
            stats["dyn_pm"] = float(pm)
            stats["dyn_rx"] = float(rx)
            stats["dyn_pr"] = float(pr)
        except Exception:
            pass

    # 回填旧版别名，保证历史组件继续可读
    stats_contract.expand_legacy_aliases(stats, percentiles=pct_keys)

    return stats


def _build_theory_stats_rows(
    stats_map: Dict[str, float],
    *,
    left_x: Optional[float] = None,
    right_x: Optional[float] = None,
    cdf_func: Optional[Callable[[float], float]] = None,
) -> List[Any]:
    """
    [理论统计行构建]
    将理论统计字典包装成 UI 可消费的 rows 列表。

    当前策略：
    - 不再手工拼一堆行对象；
    - 而是先把 stats_map 规范化，再注入动态尾部概率；
    - 最终返回 [stats_map] 这种列表形式，
      交给上层 ui_stats.assemble_simulation_rows 走统一排版。

    动态尾部概率（理论版）：
    - 若给定 left_x / right_x 和 cdf_func，则：
        pl = CDF(left_x)
        pr = 1 - CDF(right_x)
        pm = CDF(right_x) - CDF(left_x)
    - 这对应建模器里理论分布滑块分割的实时概率显示。

    注意：
    - 这里服务的是“理论统计路径”，不是经验样本路径；
    - 它通常由 ModelerStatsJob 调用。
    """
    # 标准化理论统计键名
    stats_map = stats_contract.normalize_stats_map(
        stats_map,
        percentiles=getattr(ui_stats, "DEFAULT_PERCENTILES", stats_contract.DEFAULT_PERCENTILES),
    )

    # 若提供了 CDF 函数，则补充动态分割概率
    if left_x is not None and right_x is not None and cdf_func is not None:
        try:
            lx = float(left_x)
            rx = float(right_x)
            cdf_l = float(cdf_func(lx))
            cdf_r = float(cdf_func(rx))
            pl = cdf_l
            pr = 1.0 - cdf_r
            pm = cdf_r - cdf_l
            if pm < 0:
                pm = 0.0
            stats_map["dyn_pl"] = pl
            stats_map["dyn_lx"] = lx
            stats_map["dyn_pm"] = pm
            stats_map["dyn_rx"] = rx
            stats_map["dyn_pr"] = pr
        except Exception:
            pass

    # 兼容旧字段别名，避免老展示链路断裂
    stats_contract.expand_legacy_aliases(
        stats_map,
        percentiles=getattr(ui_stats, "DEFAULT_PERCENTILES", stats_contract.DEFAULT_PERCENTILES),
    )

    # QSignal / 上层 UI 约定返回 List[Dict]
    return [stats_map]


# ============================================================
# 3. 理论分布统计任务 (用于 ui_modeler)
# ============================================================

class ModelerStatsSignals(QObject):
    """
    建模器理论统计任务的 Qt 信号定义。

    done:
    - 参数：token, rows
    - token 用于外层区分“当前回调是否仍有效”
    - rows 为理论统计结果列表

    error:
    - 参数：token, error_message
    """
    done = Signal(int, list)   # token, rows
    error = Signal(int, str)


class ModelerStatsJob(QRunnable):
    """
    [后台任务]
    建模器理论统计任务，服务于 ui_modeler。

    支持两套调用方式：

    1. 新版推荐调用：
        ModelerStatsJob(
            token,
            func_name="DriskNormal",
            params=[0,1],
            markers={"shift":1.0, ...},
            cell_label="Sheet1!A1",
            left_x=..., right_x=...
        )

    2. 旧版兼容调用：
        ModelerStatsJob(token, dist_obj, cell_label, left_x, right_x)

    当前实现说明：
    - 新版路径依赖 theory_adapter + backend_bridge；
    - 旧版 dist_obj 路径当前主要作为兼容保留；
    - 真正 run 时若缺少 func_name 或 theory 不可用，会直接返回空结果列表。

    为什么需要 token：
    - 建模器参数在用户输入时可能高频变化；
    - 外层会启动多个任务，最后只希望最新任务结果落地；
    - token 用于 UI 层丢弃过期结果。
    """

    def __init__(self, token: int, *args, **kwargs):
        super().__init__()
        self.token = token
        self.signals = ModelerStatsSignals()

        # 通用参数：单元格标签与左右分割点
        self.cell_label: str = str(kwargs.get("cell_label", "")) if "cell_label" in kwargs else ""
        self.left_x = kwargs.get("left_x", None)
        self.right_x = kwargs.get("right_x", None)

        # 新版参数：理论函数名、参数列表、标记字典
        self.func_name: Optional[str] = kwargs.get("func_name")
        self.params: Optional[Sequence[float]] = kwargs.get("params")
        self.markers: Optional[Dict[str, Any]] = kwargs.get("markers")

        # 旧版兼容参数：直接传入分布对象
        self.dist_obj = None

        # 动态解析位置参数，兼容新旧调用习惯
        # 新版位置参数约定：
        #   (func_name, params, markers, cell_label, left_x, right_x)
        # 旧版位置参数约定：
        #   (dist_obj, cell_label, left_x, right_x)
        if self.func_name is None and args:
            if isinstance(args[0], str) and str(args[0]).startswith("Drisk"):
                self.func_name = str(args[0])
                self.params = args[1] if len(args) > 1 else []
                self.markers = args[2] if len(args) > 2 else {}
                if len(args) > 3:
                    self.cell_label = str(args[3])
                if len(args) > 4:
                    self.left_x = args[4]
                if len(args) > 5:
                    self.right_x = args[5]
            else:
                self.dist_obj = args[0]
                if len(args) > 1:
                    self.cell_label = str(args[1])
                if len(args) > 2:
                    self.left_x = args[2]
                if len(args) > 3:
                    self.right_x = args[3]

    def run(self):
        """
        理论统计后台执行入口。

        主流程：
        1. 若理论函数名缺失，或 theory_adapter 不可用，则直接回传空 rows；
        2. 调 theory.compute_theory_stats(...) 获取理论统计字典；
        3. 若 bridge 可用，且左右分割点存在，则额外请求底层真实分布对象，
           构造 cdf_func；
        4. 调 _build_theory_stats_rows(...) 注入动态尾部概率并包装为 rows；
        5. emit done。

        失败时：
        - 打印 traceback；
        - emit error(token, str(e))
        """
        try:
            if not self.func_name or theory is None:
                self.signals.done.emit(self.token, [])
                return

            # 1) 获取理论统计量（均值、标准差、分位数等）
            stats = theory.compute_theory_stats(
                func_name=self.func_name,
                params=list(self.params or []),
                markers=self.markers or {},
                percentiles=getattr(ui_stats, "DEFAULT_PERCENTILES", (0.01, 0.05, 0.25, 0.5, 0.75, 0.95, 0.99))
            )

            # 2) 为动态概率滑块准备 CDF 函数
            #    这里直接向 backend_bridge 申请“底层真实分布实例”，
            #    而不是让 UI 自己解释 shift / truncation 等语义。
            cdf_func = None
            if bridge is not None and self.left_x is not None and self.right_x is not None:
                try:
                    dist = bridge.get_backend_distribution(self.func_name, list(self.params or []), self.markers or {})
                    if dist and not getattr(dist, '_invalid', False):
                        cdf_func = lambda x: float(dist.cdf(float(x)))
                except Exception:
                    pass

            # 3) 生成 UI 可消费的理论统计行
            rows = _build_theory_stats_rows(
                stats,
                left_x=self.left_x,
                right_x=self.right_x,
                cdf_func=cdf_func,
            )
            self.signals.done.emit(self.token, rows)

        except Exception as e:
            print(f"[ModelerStatsJob Error]: {traceback.format_exc()}")
            self.signals.error.emit(self.token, str(e))


# ============================================================
# 4. 实证数据统计任务 (用于 ui_results)
# ============================================================

class StatsSignals(QObject):
    """
    实证数据统计任务的 Qt 信号定义。

    done:
    - 参数：token, keys, stats_list
    - keys 通常是 UI 展示顺序下的标签列表
    - stats_list 与 keys 一一对应

    error:
    - 参数：token, error_message
    """
    done = Signal(int, list, list)  # token, keys(labels), stats_list
    error = Signal(int, str)


class StatsJob(QRunnable):
    """
    [后台任务]
    结果分析器中的“实证样本统计任务”，主要服务于 ui_results。

    支持两套路径：

    1. 新版推荐路径（桥接到底层缓存）：
        StatsJob(
            token,
            sim_id=123,
            label_to_cell_key={"A": "Sheet1!A1", ...},
            keys=["A","B",...],
            left_x=..., right_x=...,
            cancel_check=...
        )

    2. 旧版兼容路径（直接传内存 series_map）：
        StatsJob(token, series_map, keys, left_x, right_x, cancel_check=...)

    设计意图：
    - 新版路径优先，保证 UI 侧统计尽量复用底层缓存真源；
    - 旧版路径保留，是为了兼容仍未迁到 bridge 的调用点。

    关键字段：
    - sim_id：当前 simulation 标识
    - label_to_cell_key：UI 展示 key 到底层 cell key 的映射
    - keys：本次要统计的标签顺序
    - series_kind：区分统计输入还是输出
    - cancel_check：任务取消回调，防止过期任务浪费资源
    """

    def __init__(self, token: int, *args, **kwargs):
        super().__init__()
        self.token = token
        self.signals = StatsSignals()

        # 取消检测函数：
        # 外层通常在切换图表 / 关闭窗口 / 新任务覆盖旧任务时使用
        self.cancel_check: Optional[Callable[[], bool]] = kwargs.get("cancel_check", None)

        # 左右游标坐标，用于动态尾部概率统计
        self.left_x = kwargs.get("left_x", None)
        self.right_x = kwargs.get("right_x", None)

        # 新版路径参数
        self.sim_id: Optional[int] = kwargs.get("sim_id", None)
        self.label_to_cell_key: Dict[str, str] = kwargs.get("label_to_cell_key", {}) or {}
        self.keys: List[str] = kwargs.get("keys", None) or []
        self.series_kind: str = _normalize_series_kind(kwargs.get("series_kind", "output"))

        # 旧版路径：直接携带原始序列字典
        self.series_map: Dict[str, Any] = {}

        # 兼容旧版位置参数调用：
        # (series_map, keys, left_x, right_x, cancel_check)
        if not self.keys and len(args) >= 2 and isinstance(args[0], dict):
            self.series_map = args[0]
            self.keys = list(args[1])
            if len(args) >= 3 and self.left_x is None:
                self.left_x = args[2]
            if len(args) >= 4 and self.right_x is None:
                self.right_x = args[3]
            if len(args) >= 5 and self.cancel_check is None:
                self.cancel_check = args[4]

    def run(self):
        """
        实证统计后台执行入口。

        执行逻辑分两条：

        A. 新版 bridge 路径：
           - 逐个 key 从底层缓存读 data；
           - 本地算一遍 _calculate_empirical_stats；
           - 再尝试用 bridge.get_statistics(...) 的真源统计覆盖核心字段；
           - 汇总后 emit done。

        B. 旧版 series_map 路径：
           - 直接对传入序列逐个做 _calculate_empirical_stats；
           - 不使用 bridge 覆盖；
           - 最后 emit done。

        注意：
        - 每个循环点都尽量检查 cancel_check；
        - 避免旧任务在新界面状态下继续回写。
        """
        try:
            stats_list: List[Dict[Any, float]] = []

            # -------- 新版架构：通过 bridge 从底层缓存读取数据 --------
            if self.sim_id is not None:
                if bridge is None:
                    raise RuntimeError("backend_bridge 不可用，无法执行 StatsJob(sim_id=...)")
                sim_id = int(self.sim_id)

                for label in self.keys:
                    if self.cancel_check and self.cancel_check():
                        return

                    ck = self.label_to_cell_key.get(label, label)
                    try:
                        data = bridge.get_series(sim_id, ck, kind=self.series_kind)
                    except Exception:
                        data = None

                    # 先算一遍本地经验统计，再尽量以底层统计覆盖核心字段
                    stats = _calculate_empirical_stats(data, left_x=self.left_x, right_x=self.right_x)
                    try:
                        backend_stats = bridge.get_statistics(sim_id, ck, kind=self.series_kind)
                    except Exception:
                        backend_stats = {}
                    stats = _override_stats_with_backend(stats, backend_stats, data)
                    stats_list.append(stats)

                if self.cancel_check and self.cancel_check():
                    return

                self.signals.done.emit(self.token, self.keys, stats_list)
                return

            # -------- 旧版架构：直接使用传入的内存序列字典 --------
            for k in self.keys:
                if self.cancel_check and self.cancel_check():
                    return

                s = self.series_map.get(k)
                stats = _calculate_empirical_stats(s, left_x=self.left_x, right_x=self.right_x)
                stats_list.append(stats)

            if self.cancel_check and self.cancel_check():
                return

            self.signals.done.emit(self.token, self.keys, stats_list)

        except Exception as e:
            self.signals.error.emit(self.token, str(e))


# ============================================================
# 5. KDE 密度估计任务 (用于 ui_results)
# ============================================================

class KDESignals(QObject):
    """
    KDE 任务的 Qt 信号定义。

    done:
    - 参数：token, x_grid, y_map
    - x_grid 为统一采样网格
    - y_map 为每条 series 的 KDE 密度数组

    error:
    - 参数：token, error_message
    """
    done = Signal(int, object, object)  # token, x_grid, y_map
    error = Signal(int, str)            # token, error message


class KDEJob(QRunnable):
    """
    [后台任务]
    KDE（核密度估计）计算任务。

    主要职责：
    - 对多个数据序列计算 KDE 密度曲线；
    - 输出统一 x_grid 与各序列对应的 y 值；
    - 对脏数据、低样本量、边界支撑集等情况做稳健处理。

    设计特征：
    1. 强健的数据清洗：
       自动跳过 '#ERROR!'、None、非数值字符串等脏值。

    2. 下采样保护：
       当样本过大时，随机无放回抽样到 max_samples，
       减轻 KDE 计算压力。

    3. 边界反射修正：
       针对有限支撑集样本，使用 reflection boundary correction
       改善边界附近的 KDE 形状。

    4. 固定 RNG：
       rng = np.random.default_rng(0)
       避免多次重绘时下采样结果抖动，导致曲线视觉不稳定。
    """

    def __init__(
        self,
        token: int,
        datasets: Dict[str, Any],
        keys: List[str],
        x_min: float,
        x_max: float,
        pdf_points: int,
        max_samples: int = 8000,
        cancel_check: Optional[Callable[[], bool]] = None,
    ):
        super().__init__()
        self.token = token
        self.datasets = datasets
        self.keys = keys
        self.x_min = x_min
        self.x_max = x_max
        self.pdf_points = int(pdf_points)
        self.max_samples = int(max_samples)
        self.cancel_check = cancel_check
        self.signals = KDESignals()

    def run(self):
        """
        KDE 后台执行入口。

        主流程：
        1. 先构造统一 x_grid；
        2. 对每个 key：
           - 提取原始序列；
           - 清洗成有效 float 数组；
           - 样本过少或支撑过窄则跳过；
           - 样本过大则做下采样；
           - 用 gaussian_kde 计算 KDE；
           - 做边界反射修正与平滑增益；
           - 写入 y_map。
        3. emit done(token, x_grid, y_map)

        关键保护逻辑：
        - 任意单条序列的 KDE 失败，不应导致整个图表崩溃；
        - 因此单序列失败时采用 continue。
        """
        try:
            x_grid = np.linspace(float(self.x_min), float(self.x_max), int(self.pdf_points))
            y_map: Dict[str, np.ndarray] = {}

            # 固定随机种子，避免下采样导致多次重绘曲线形状不稳定
            rng = np.random.default_rng(0)

            for k in self.keys:
                if self.cancel_check and self.cancel_check():
                    return

                s = self.datasets.get(k, None)
                if s is None:
                    continue

                # 关键预处理：先清洗出有限浮点数
                arr = _filter_valid_float_array(s)
                if arr.size < 3:
                    continue
                sample_min = float(np.min(arr))
                sample_max = float(np.max(arr))
                if float(sample_max - sample_min) <= 1e-12:
                    continue

                # 样本数过大时做无放回下采样，降低 KDE 计算成本
                if arr.size > self.max_samples:
                    arr = rng.choice(arr, size=self.max_samples, replace=False)

                if self.cancel_check and self.cancel_check():
                    return

                try:
                    kde = gaussian_kde(arr)

                    # 先计算中心 KDE
                    y_core = np.asarray(kde(x_grid), dtype=float)

                    # 边界反射修正：
                    # 对于有限支撑分布，在左右边界进行镜像采样补偿
                    y_left = np.asarray(kde(2.0 * sample_min - x_grid), dtype=float)
                    y_right = np.asarray(kde(2.0 * sample_max - x_grid), dtype=float)

                    # 对反射项附加局部边缘增益，使边缘衰减更自然
                    sample_span = float(sample_max - sample_min)
                    bw = float(np.sqrt(np.ravel(np.asarray(kde.covariance, dtype=float))[0]))
                    if not np.isfinite(bw) or bw <= 0.0:
                        bw = sample_span * 0.05

                    edge_span = max(sample_span * 0.03, 2.0 * bw)
                    edge_span = min(edge_span, sample_span * 0.25)

                    if edge_span > 0.0:
                        dist_to_edge = np.minimum(x_grid - sample_min, sample_max - x_grid)
                        t = np.clip(dist_to_edge / edge_span, 0.0, 1.0)
                        smooth = t * t * (3.0 - 2.0 * t)  # smoothstep
                        reflect_gain = 0.60 + 0.40 * smooth
                    else:
                        reflect_gain = 1.0

                    y = y_core + reflect_gain * (y_left + y_right)
                except Exception:
                    # 单个 series 失败时静默跳过，避免整张图不可用
                    continue

                if y.size != x_grid.size:
                    continue
                y = np.where(np.isfinite(y), y, 0.0)
                y = np.maximum(y, 0.0)
                y_map[k] = y

            if self.cancel_check and self.cancel_check():
                return

            self.signals.done.emit(self.token, x_grid, y_map)

        except Exception as e:
            self.signals.error.emit(self.token, str(e))


# ============================================================
# 6. CDF 累积概率模式统计任务 (用于 ui_results)
# ============================================================

class CDFStatsSignals(QObject):
    """
    CDF 统计任务的 Qt 信号定义。

    done:
    - 参数：token, rows
    - rows 为已经过 assemble_cdf_rows(...) 排版前整理的数据行
    """
    done = Signal(int, list)


class CDFStatsWorker(QRunnable):
    """
    [后台任务]
    累积概率模式下的统计任务，服务于 ui_results 的 CDF 视图。

    它的职责比 StatsJob 多一点：
    1. 先统一计算基础统计；
    2. 若 show_ci=True，则进一步计算均值的置信区间误差范围；
    3. 最终调用 assemble_cdf_rows(...)，将结果整理为 CDF 模式表格所需结构。

    支持新版 bridge 路径：
    - sim_id + label_to_cell_key + series_kind

    也兼容旧版路径：
    - 直接使用 raw_series_map

    参数说明：
    - ci_level：置信水平
    - show_ci：是否显示均值置信区间
    - engine_mode：传给 _smart_ci_engine 的引擎模式
    - left_x / right_x：用于动态尾部概率
    """

    def __init__(
        self, 
        token: int, 
        raw_series_map: Dict[str, Any], 
        keys: List[str], 
        ci_level: float, 
        show_ci: bool, 
        engine_mode: str, 
        cancel_check: Callable[[], bool],
        left_x: Optional[float] = None,
        right_x: Optional[float] = None,
        sim_id: Optional[int] = None,
        label_to_cell_key: Optional[Dict[str, str]] = None,
        series_kind: str = "output",
    ):
        super().__init__()

        self.token = token

        # 防御性拷贝：
        # 避免外部在任务运行过程中修改原字典 / 原序列对象。
        # 这里不强制所有值都支持 copy()，仅在可 copy 时拷贝。
        self.raw_series_map = {
            k: (v.copy() if hasattr(v, "copy") else v) for k, v in raw_series_map.items()
        }

        self.keys = keys
        self.ci_level = ci_level
        self.show_ci = show_ci
        self.engine_mode = engine_mode
        self.cancel_check = cancel_check

        # 游标切分点，用于动态概率统计
        self.left_x = left_x
        self.right_x = right_x

        # 新版 bridge 路径参数
        self.sim_id = sim_id
        self.label_to_cell_key: Dict[str, str] = dict(label_to_cell_key or {})
        self.series_kind = _normalize_series_kind(series_kind)

        self.signals = CDFStatsSignals()
        
    def run(self):
        """
        CDF 统计后台执行入口。

        主流程：
        1. 若任务已取消，则直接返回；
        2. 对每个 key 计算基础经验统计：
           - 优先走 bridge 路径读 series；
           - 必要时再用 bridge.get_statistics 覆盖核心字段；
           - 同时保留原始 data 供后续 CI 计算使用。
        3. 若 show_ci=True，则对每个 key 额外计算均值置信区间误差范围；
        4. 调用 assemble_cdf_rows(...) 交给 UI 排版层；
        5. emit done(token, rows)

        注意：
        - 这里的 CI 计算比普通统计更耗时；
        - 因此 show_ci 开关很重要，避免无谓后台负载。
        """
        if self.cancel_check and self.cancel_check(): 
            return
            
        try:
            # 1) 基础统计计算（与普通经验统计路径保持一致）
            stats_map_list = []
            data_map_for_ci: Dict[str, Any] = {}

            use_backend = bridge is not None and self.sim_id is not None
            sim_id = int(self.sim_id) if self.sim_id is not None else None
            
            for k in self.keys:
                data = self.raw_series_map.get(k)
                ck = self.label_to_cell_key.get(k, k)

                if use_backend and sim_id is not None:
                    try:
                        data = bridge.get_series(sim_id, ck, kind=self.series_kind)
                    except Exception:
                        data = self.raw_series_map.get(k)
                        
                stats = _calculate_empirical_stats(data, left_x=self.left_x, right_x=self.right_x)

                if use_backend and sim_id is not None:
                    try:
                        backend_stats = bridge.get_statistics(sim_id, ck, kind=self.series_kind)
                    except Exception:
                        backend_stats = {}
                    stats = _override_stats_with_backend(stats, backend_stats, data)
                    
                stats_map_list.append(stats)
                data_map_for_ci[k] = data
            
            if self.cancel_check and self.cancel_check(): 
                return

            # 2) 若需要展示均值 CI，则进一步计算误差半宽显示文本
            mean_errors = []
            if self.show_ci:
                for k in self.keys:
                    d = data_map_for_ci.get(k)
                    d_clean = _filter_valid_float_array(d)
                    if len(d_clean) > 0:
                        # 通过 _smart_ci_engine 分别计算均值 CI 下限 / 上限
                        L = _smart_ci_engine(d_clean, 'mean', self.ci_level, True, 0.5, is_sorted=False, engine_mode=self.engine_mode)
                        U = _smart_ci_engine(d_clean, 'mean', self.ci_level, False, 0.5, is_sorted=False, engine_mode=self.engine_mode)
                        if not np.isnan(L) and not np.isnan(U):
                            err = (U - L) / 2.0
                            mean_errors.append(f"±{smart_format_number(err)}")
                            continue
                    mean_errors.append("")
            
            if self.cancel_check and self.cancel_check(): 
                return

            # 3) 将后台结果交给 ui_stats 统一排版
            rows = assemble_cdf_rows(
                stats_map_list=stats_map_list, 
                keys=self.keys, 
                ci_level=self.ci_level, 
                show_ci=self.show_ci,
                mean_errors=mean_errors
            )
            
            self.signals.done.emit(self.token, rows)
            
        except Exception as e:
            print(f"[CDF Stats Error]: {traceback.format_exc()}")
            if not (self.cancel_check and self.cancel_check()):
                self.signals.done.emit(self.token, [])