# ui_results_advanced_state.py
"""

"""

from __future__ import annotations

from typing import Mapping

# =======================================================
# 1. 高级分析全局默认配置常量
# =======================================================
# 龙卷风图 (Tornado) 的默认配置字典
_DEFAULT_TORNADO_CONFIG = {
    "num_bins": 10,           # 默认数据分箱数量 (用于分箱均值等模式)
    "max_vars": 16,           # 绘图区域最多展示的变量条目数，超出的将被隐藏
    "stat_type": "mean",      # 各组内的统计算法，默认使用均值 (mean)
    "percentile_val": 90.0,   # 当统计类型被设为分位数时的默认百分位值
}

# 情景分析图 (Scenario) 的默认配置字典
_DEFAULT_SCENARIO_CONFIG = {
    "display_mode": 0,                # 默认视图模式索引 (决定使用哪种子图表渲染)
    "min_pct": 0.0,                   # 情景尾部筛选下限概率 (0%)
    "max_pct": 25.0,                  # 情景尾部筛选上限概率 (25%)
    "significance_threshold": 0.5,    # 显著性阈值过滤 (低于此方差贡献的变量将被隐藏)
    "display_limit": None,            # 强制显示的变量总数上限 (None 代表无强制限制)
}

# 获取隔离副本的工厂方法
def default_tornado_config() -> dict:
    """获取一份解除引用的龙卷风默认配置副本，防止运行时的局部修改污染全局字典。"""
    return dict(_DEFAULT_TORNADO_CONFIG)

def default_scenario_config() -> dict:
    """获取一份解除引用的情景分析默认配置副本。"""
    return dict(_DEFAULT_SCENARIO_CONFIG)

# =======================================================
# 2. 运行时配置规范化与安全清洗引擎 (Normalization)
# =======================================================
# 龙卷风图参数安全清洗
def normalize_tornado_config(cfg: Mapping | None) -> dict:
    """
    龙卷风图配置清洗器。
    对传入的配置字典进行字段的类型安全转换和极值约束，确保底层绘图算法绝对不会收到非法参数。
    """
    merged = default_tornado_config()
    if cfg:
        merged.update(dict(cfg))

    # 约束 1：分箱数量必须是大于等于 2 的整数，否则无统计意义
    try:
        merged["num_bins"] = max(2, int(merged.get("num_bins", 10)))
    except Exception:
        merged["num_bins"] = _DEFAULT_TORNADO_CONFIG["num_bins"]

    # 约束 2：最大展示变量数，允许为 None（展示所有），否则强制约束为 >= 1 的整数
    raw_max_vars = merged.get("max_vars", _DEFAULT_TORNADO_CONFIG["max_vars"])
    if raw_max_vars in (None, ""):
        merged["max_vars"] = None
    else:
        try:
            merged["max_vars"] = max(1, int(raw_max_vars))
        except Exception:
            merged["max_vars"] = _DEFAULT_TORNADO_CONFIG["max_vars"]

    # 约束 3：标准化统计类型名称。处理缩写兼容性，过滤未知类型
    stat_type = str(merged.get("stat_type", "mean")).strip().lower()
    if stat_type == "pctl":
        stat_type = "percentile"
    if stat_type not in {"mean", "median", "percentile"}:
        stat_type = "mean"
    merged["stat_type"] = stat_type

    # 约束 4：百分位数值必须严格限制在 [0.0, 100.0] 区间内
    try:
        p = float(merged.get("percentile_val", 90.0))
    except Exception:
        p = _DEFAULT_TORNADO_CONFIG["percentile_val"]
    merged["percentile_val"] = min(100.0, max(0.0, p))

    return merged

# 情景分析参数安全清洗
def normalize_scenario_config(cfg: Mapping | None) -> dict:
    """
    情景分析配置清洗器。
    核心难点在于处理各种用户输入的概率百分比阈值，防止越界或产生逻辑冲突（如下限大于上限）。
    """
    merged = default_scenario_config()
    if cfg:
        merged.update(dict(cfg))

    # 约束 1：显示模式防错，确保是 >= 0 的整型索引
    try:
        merged["display_mode"] = max(0, int(merged.get("display_mode", 0)))
    except Exception:
        merged["display_mode"] = _DEFAULT_SCENARIO_CONFIG["display_mode"]

    # 约束 2：提取下限百分比与上限百分比
    try:
        min_pct = float(merged.get("min_pct", 0.0))
    except Exception:
        min_pct = _DEFAULT_SCENARIO_CONFIG["min_pct"]

    try:
        max_pct = float(merged.get("max_pct", 25.0))
    except Exception:
        max_pct = _DEFAULT_SCENARIO_CONFIG["max_pct"]

    # 约束 3：百分比硬边界约束 [0, 100] 与智能防呆翻转（如果用户输入 下限 > 上限，则自动对调两者）
    min_pct = min(100.0, max(0.0, min_pct))
    max_pct = min(100.0, max(0.0, max_pct))
    if min_pct > max_pct:
        min_pct, max_pct = max_pct, min_pct

    merged["min_pct"] = min_pct
    merged["max_pct"] = max_pct

    # 约束 4：显著性过滤阈值安全检查，不能小于 0.0
    raw_threshold = merged.get("significance_threshold", _DEFAULT_SCENARIO_CONFIG["significance_threshold"])
    if raw_threshold in (None, ""):
        merged["significance_threshold"] = None
    else:
        try:
            merged["significance_threshold"] = max(0.0, float(raw_threshold))
        except Exception:
            merged["significance_threshold"] = _DEFAULT_SCENARIO_CONFIG["significance_threshold"]

    # 约束 5：最大显示条目数量限制安全检查
    raw_limit = merged.get("display_limit", _DEFAULT_SCENARIO_CONFIG["display_limit"])
    if raw_limit in (None, ""):
        merged["display_limit"] = None
    else:
        try:
            merged["display_limit"] = max(1, int(raw_limit))
        except Exception:
            merged["display_limit"] = _DEFAULT_SCENARIO_CONFIG["display_limit"]
            
    return merged

# =======================================================
# 3. 高级图表专属配色样式工厂
# =======================================================
def build_default_tornado_styles(color_cycle: list[str]) -> dict:
    """
    构建龙卷风与情景分析专属的默认颜色映射字典。
    将全局的主题调色板数组与特定的对比分组（如高/低、正/负组）安全绑定。
    注：此处原代码使用了 Unicode 字符硬编码中文键名以防止编码乱码：
        \u9ad8\u503c\u7ec4 = 高值组
        \u4f4e\u503c\u7ec4 = 低值组
        \u6b63\u503c = 正值
        \u8d1f\u503c = 负值
    """
    return {
        "高值组": {"color": color_cycle[0], "fill_opacity": 1.0},
        "低值组": {"color": color_cycle[1], "fill_opacity": 1.0},
        "正值": {"color": color_cycle[0], "fill_opacity": 1.0},
        "负值": {"color": color_cycle[1], "fill_opacity": 1.0},
    }