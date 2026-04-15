# theory_adapter.py
"""
主要架构与设计意图：
1. 纯粹的 API 转发层：彻底剥离了前端 UI 对复杂数学计算逻辑以及第三方科学计算库（如 SciPy、NumPy 复杂函数）的直接依赖。
2. 数据溯源统一：前端不再自行估算分布特征，而是全权委托 backend_bridge 向底层的 PyXLL 引擎（C++ 核心）索要实例化后的 DistributionBase 对象。
3. 绝对精度保证：通过直接读取底层引擎预先计算好的精确理论矩、截断边界与支撑集，确保前端展示的数据与后端蒙特卡洛引擎使用的计算模型完全一致，消除双端不一致的隐患。
"""

from __future__ import annotations
from typing import Any, Dict, List, Optional, Sequence, Tuple
import math
import numpy as np

import backend_bridge

class TheoryAdapterError(RuntimeError):
    """理论适配器专属异常类，用于抛出前端与底层引擎交互时的运行时错误"""
    pass

# =============================================================================
# 核心接口 1：获取理论统计面板数据 (Theoretical Statistics Computation)
# =============================================================================
def compute_theory_stats(
    func_name: str,
    params: Sequence[Any],
    markers: Optional[Dict[str, Any]] = None,
    *,
    percentiles: Sequence[float] = (
        0.01, 0.025, 0.05, 0.10, 0.20, 0.25, 0.30, 0.35, 0.40, 0.45,
        0.50, 0.55, 0.60, 0.65, 0.70, 0.75, 0.80, 0.85, 0.90, 0.95,
        0.975, 0.99
    ),
    **kwargs # 吸收并废弃旧版遗留的无用参数（如 prefer_scipy），以保持接口兼容性
) -> Dict[str, float]:
    """
    数据抓取方法：向底层引擎传递参数及高级形态特征（截断、平移），换取该分布绝对精确的理论矩与分位数。
    返回的字典将直接用于渲染 UI 侧的“理论统计特性”面板。
    """
    # 1. 实例委托：向桥接层索取由底层引擎构建的分布实例 (包含截断、平移等全套逻辑)
    dist = backend_bridge.get_backend_distribution(func_name, list(params), markers or {})
    
    # 防呆校验：如果引擎实例化失败或返回无效标记，安全返回 NaN 以保护前端不崩溃
    if dist is None or getattr(dist, '_invalid', False):
        return {"error": float('nan')} 

    # 2. 提取核心理论矩 (Theoretical Moments)
    out: Dict[str, float] = {}
    out["mean"] = float(dist.mean())                 # 均值
    out["variance"] = float(dist.variance())         # 方差
    out["skew"] = float(dist.skewness())             # 偏度
    out["skewness"] = out["skew"]                    # 别名映射
    out["kurt"] = float(dist.kurtosis())             # 峰度
    out["kurtosis"] = out["kurt"]                    # 别名映射
    
    # 标准差由方差开方得出（增加负数方差的防呆保护）
    out["std"] = math.sqrt(out["variance"]) if out["variance"] >= 0 else float('nan')

    try:
        # 众数计算可能在某些特殊分布或截断形态下不可用，故作异常捕获
        out["mode"] = float(dist.mode())
    except Exception:
        pass

    # 3. 提取有效支撑集 (Effective Bounds)
    # 这里获取的是叠加了“截断 (Truncation)”等边界条件后的有效最大/最小值
    lo, hi = dist.get_effective_bounds()
    out["min"] = float(lo) if lo is not None else float(dist.min_val())
    out["max"] = float(hi) if hi is not None else float(dist.max_val())

    # 4. 计算逆累积分布/百分位数 (Percent Point Function / PPF)
    # 根据传入的百分比序列，批量获取理论分位数点
    for p in percentiles:
        try:
            val = float(dist.ppf(p))
        except Exception:
            val = float("nan")
            
        # 为了兼容不同的 UI 组件调用习惯，将同一个值绑定到多种键名格式上
        out[p] = val                             # float 键名: 0.05
        out[str(p)] = val                        # 字符串键名: "0.05"
        out[f"p{p}"] = val                       # 前缀字符串: "p0.05"
        
        pct_int = int(round(p * 100))
        out[f"p{pct_int}"] = val                 # 百分比整数: "p5"
        out[f"p{pct_int:02d}"] = val             # 固定两位数: "p05"
        out[f"p{int(p * 1000) / 10:g}"] = val    # 高精度浮点: "p5.5" (用于 0.055)
        
        # 提取中位数特例 (50th percentile)
        if abs(p - 0.5) < 1e-6:
            out["median"] = val
            out["中位数"] = val

    return out


# =============================================================================
# 核心接口 2：生成用于 Plotly 绘图的曲线点集 (Curve Series Generation)
# =============================================================================
def compute_pdf_or_pmf_series(
    func_name: str,
    params: Sequence[Any],
    markers: Optional[Dict[str, Any]] = None,
    *,
    n_points: int = 401,
    x_min: Optional[float] = None,
    x_max: Optional[float] = None,
    discrete_max_points: int = 120,
    **kwargs # 吸收并废弃旧版参数
) -> Tuple[np.ndarray, np.ndarray]:
    """
    绘图引擎适配：根据底层分布特征自动推导可视化的 X 轴范围，并将其喂给底层引擎，
    取回完美的 Y 轴序列 (连续型为概率密度 PDF，离散型为概率质量 PMF)。
    返回供 Plotly 消费的 (X_array, Y_array) 数据组。
    """
    # 初始化分布引擎
    dist = backend_bridge.get_backend_distribution(func_name, list(params), markers or {})
    
    # 数据异常时的保护性空返回
    if dist is None or getattr(dist, '_invalid', False):
        return np.array([]), np.array([])

    # 1. 确定绘图的 X 轴有效边界 (Determine Visible X-Axis Support)
    lo, hi = dist.get_effective_bounds()
    if lo is None: lo = dist.min_val()
    if hi is None: hi = dist.max_val()

    # 视口裁剪：如果调用方没有强行锁定视图区域，则使用 0.1% 到 99.9% 概率覆盖的区间作为最优展示视口
    # 这样可以避免具有无限长尾的分布（如对数正态）导致核心图形被压扁。
    if x_min is None or x_max is None:
        try:
            x_min = float(lo) if math.isfinite(lo) else float(dist.ppf(0.001))
            x_max = float(hi) if math.isfinite(hi) else float(dist.ppf(0.999))
        except Exception:
            x_min, x_max = -1.0, 1.0

    # 边界数值安全兜底校验
    if not math.isfinite(x_min) or not math.isfinite(x_max) or x_max <= x_min:
        x_min, x_max = -1.0, 1.0

    # 2. 路由分发：判断并计算离散/连续曲线特征 (Discrete PMF vs Continuous PDF)
    # (策略：通过反射检查底层实例是否具有 pmf 接口来判断是否为离散分布)
    if hasattr(dist, "pmf"):
        # 离散分布分支
        start = int(math.floor(x_min))
        end = int(math.ceil(x_max))
        
        # 性能保护机制：限制离散点生成的最大数量（防止用户输入极端的参数导致渲染引擎死锁挂起）
        # 如果超出阈值，强制将绘图焦点收缩并锁定在理论中位数附近。
        if (end - start) > discrete_max_points:
            med = float(dist.ppf(0.5))
            half = discrete_max_points // 2
            start = int(math.floor(med)) - half
            end = int(math.floor(med)) + half
            
        xs = np.arange(start, end + 1, dtype=float)
        try:
            # 安全地将标量求值的 pmf 映射逻辑向量化
            ys = np.array([float(dist.pmf(x)) for x in xs])
        except Exception:
            ys = np.zeros_like(xs)
    else:
        # 连续分布分支
        # 默认生成 401 个线性均匀分布的采样点进行密度评估
        xs = np.linspace(x_min, x_max, n_points)
        try:
            ys = np.array([float(dist.pdf(x)) for x in xs])
        except Exception:
            ys = np.zeros_like(xs)

    # 3. 最终清洗：剔除所有无穷大的异常点 (NaN/Inf 替换为 0.0)，防止图形库渲染报错
    ys = np.where(np.isfinite(ys), ys, 0.0)
    
    return xs, ys