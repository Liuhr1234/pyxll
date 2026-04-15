"""
本模块定义了工作线程 (Worker) 与统计 UI 层之间共享的统计字段契约 (Contract)。
主要用于统一前后端数据交互时的键名规范，并提供数据清洗和向下兼容的转换方法。
"""

from __future__ import annotations

from typing import Any, Dict, Iterable, Mapping, Sequence, Tuple

import numpy as np

# =======================================================
# 1. 规范定义与标准键名 (Canonical Definitions)
# =======================================================

# 工作线程生成数据和 UI 组装时使用的标准百分位数集合。
DEFAULT_PERCENTILES: Tuple[float, ...] = (
    0.01, 0.025, 0.05, 0.10, 0.20, 0.25, 0.30, 0.35, 0.40, 0.45,
    0.50, 0.55, 0.60, 0.65, 0.70, 0.75, 0.80, 0.85, 0.90, 0.95,
    0.975, 0.99,
)

# 标准化键名 (Canonical keys)
# 这些是系统内部流转和核心逻辑中唯一被认可的字典键名常量
COUNT_KEY = "count"
MIN_KEY = "min"
MAX_KEY = "max"
MEAN_KEY = "mean"
CI90_KEY = "ci90"
MEDIAN_KEY = "median"
STD_KEY = "std"
SKEW_KEY = "skew"
KURT_KEY = "kurt"

# 动态概率计算相关的标准键名（用于滑块区间关联的概率统计）
DYN_PL_KEY = "dyn_pl"
DYN_LX_KEY = "dyn_lx"
DYN_PM_KEY = "dyn_pm"
DYN_RX_KEY = "dyn_rx"
DYN_PR_KEY = "dyn_pr"

# 异常与过滤统计的标准键名
ERROR_COUNT_KEY = "error_count"
FILTERED_COUNT_KEY = "filtered_count"


# =======================================================
# 2. 兼容级别名映射字典 (Alias Mapping Tables)
# =======================================================
# 别名映射表：用于保持与旧版本数据载荷 (payloads) 的向后兼容性，
# 并能识别包含中英文及不同命名习惯的输入。

COUNT_ALIASES: Tuple[str, ...] = ("count", "模拟次数", "数值数", "n_iterations")

CORE_FIELD_ALIASES: Dict[str, Tuple[str, ...]] = {
    MIN_KEY: ("min", "最小值", "min_val", "minimum"),
    MAX_KEY: ("max", "最大值", "max_val", "maximum"),
    MEAN_KEY: ("mean", "均值", "mu", "expectation"),
    CI90_KEY: ("ci90", "90%CI", "ci_90"),
    MEDIAN_KEY: ("median", "中位数", "med", "p50", "p50.0", "p50.00"),
    STD_KEY: ("std", "std_dev", "标准差", "sigma"),
    SKEW_KEY: ("skew", "skewness", "偏度"),
    KURT_KEY: ("kurt", "kurtosis", "峰度"),
}

DYNAMIC_FIELD_ALIASES: Dict[str, Tuple[str, ...]] = {
    DYN_PL_KEY: ("dyn_pl", "左侧概率", "p_left"),
    DYN_LX_KEY: ("dyn_lx", "左侧X", "x_left"),
    DYN_PM_KEY: ("dyn_pm", "中间概率", "p_mid"),
    DYN_RX_KEY: ("dyn_rx", "右侧X", "x_right"),
    DYN_PR_KEY: ("dyn_pr", "右侧概率", "p_right"),
}

COUNT_DETAIL_FIELD_ALIASES: Dict[str, Tuple[str, ...]] = {
    ERROR_COUNT_KEY: ("error_count", "错误数"),
    FILTERED_COUNT_KEY: ("filtered_count", "过滤数"),
}


# =======================================================
# 3. 辅助工具函数 (Helper Functions)
# =======================================================
def _is_nan_value(value: Any) -> bool:
    """检查给定的值是否为 NaN (非数字)。"""
    return isinstance(value, (float, np.floating)) and np.isnan(value)


def percentile_display_label(p: float) -> str:
    """将小数形式的概率转化为 UI 显示用的百分比标签（如 0.025 -> '2.5%'）。"""
    return f"{int(p * 1000) / 10:g}%"


def canonical_percentile_key(p: float) -> str:
    """
    生成百分位数标准键名。
    这种表示方法可确保像 2.5 和 97.5 这样的浮点条目生成唯一且干净的字符串键（如 'p2.5'）。
    """
    return f"p{int(p * 1000) / 10:g}"


def percentile_aliases(p: float) -> Tuple[Any, ...]:
    """为特定的百分位数生成一系列可能存在的历史别名（兼容各种格式）。"""
    pct_int = int(round(p * 100))
    return (
        canonical_percentile_key(p),
        f"p{pct_int}",
        f"p{pct_int:02d}",
        f"p{p}",
        percentile_display_label(p),
        str(p),
        p,  # 遗留的浮点数类型键
    )


def _resolve_first(stats_map: Mapping[Any, Any], keys: Iterable[Any]) -> Any:
    """
    字典键名解析器：
    按提供的别名列表顺序依次查找字典，返回第一个非 NaN 的匹配值。
    如果都没有找到，返回 None。
    """
    for key in keys:
        if key in stats_map:
            value = stats_map[key]
            if _is_nan_value(value):
                continue
            return value
    return None


# =======================================================
# 4. 核心数据转换引擎 (Core Transformation Logic)
# =======================================================
def normalize_stats_map(
    stats_map: Mapping[Any, Any] | None,
    percentiles: Sequence[float] = DEFAULT_PERCENTILES,
) -> Dict[str, Any]:
    """
    [清洗入口] 将可能包含混合键名（别名/中文名/旧字段名）的统计数据字典，
    标准化为完全符合当前契约的、仅包含标准键名（Canonical keys）的新字典。
    """
    source = stats_map or {}
    normalized: Dict[str, Any] = {}

    # 1. 提取计数值
    count_value = _resolve_first(source, COUNT_ALIASES)
    if count_value is not None:
        normalized[COUNT_KEY] = count_value

    # 2. 提取核心统计特征
    for canonical_key, aliases in CORE_FIELD_ALIASES.items():
        value = _resolve_first(source, aliases)
        if value is not None:
            normalized[canonical_key] = value

    # 3. 提取动态区间/概率统计特征
    for canonical_key, aliases in DYNAMIC_FIELD_ALIASES.items():
        value = _resolve_first(source, aliases)
        if value is not None:
            normalized[canonical_key] = value
            
    # 4. 提取异常明细计数
    for canonical_key, aliases in COUNT_DETAIL_FIELD_ALIASES.items():
        value = _resolve_first(source, aliases)
        if value is not None:
            normalized[canonical_key] = value

    # 5. 提取百分位数值
    for p in percentiles:
        aliases = percentile_aliases(float(p))
        value = _resolve_first(source, aliases)
        if value is not None:
            normalized[canonical_percentile_key(float(p))] = value

    return normalized


def expand_legacy_aliases(
    stats_map: Dict[Any, Any],
    percentiles: Sequence[float] = DEFAULT_PERCENTILES,
) -> Dict[Any, Any]:
    """
    [兼容入口] 将标准键名的值回填 (Backfill) 到历史别名键中，
    以此来生成一个包含所有可能命名方式的超集字典，确保旧版本的外部调用代码不会因为找不到键名而报错。
    """
    if COUNT_KEY in stats_map:
        count_value = stats_map[COUNT_KEY]
        for alias in COUNT_ALIASES:
            stats_map.setdefault(alias, count_value)

    for canonical_key, aliases in CORE_FIELD_ALIASES.items():
        if canonical_key in stats_map:
            value = stats_map[canonical_key]
            for alias in aliases:
                stats_map.setdefault(alias, value)

    for canonical_key, aliases in DYNAMIC_FIELD_ALIASES.items():
        if canonical_key in stats_map:
            value = stats_map[canonical_key]
            for alias in aliases:
                stats_map.setdefault(alias, value)
                
    for canonical_key, aliases in COUNT_DETAIL_FIELD_ALIASES.items():
        if canonical_key in stats_map:
            value = stats_map[canonical_key]
            for alias in aliases:
                stats_map.setdefault(alias, value)

    for p in percentiles:
        canonical_key = canonical_percentile_key(float(p))
        if canonical_key not in stats_map:
            continue
        value = stats_map[canonical_key]
        for alias in percentile_aliases(float(p)):
            stats_map.setdefault(alias, value)

    return stats_map
