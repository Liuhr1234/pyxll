# dist_trigen.py
"""
三参数三角分布（Trigen）专属模块
基于分位数参数 L, M, U, alpha, beta 定义，内部转换为标准三角分布。
用于 Drisk 系统集成，复用三角分布的随机数生成和理论统计功能。
"""

import math
import numpy as np
from typing import List, Optional, Union, Tuple
from distribution_base import DistributionBase
from dist_triang import (
    triang_generator_single, triang_generator_vectorized,
    triang_cdf, triang_ppf, triang_pdf,
    triang_raw_mean, triang_raw_var, triang_raw_skew, triang_raw_kurt, triang_raw_mode,
    TriangDistribution
)

# -------------------- 精确转换函数（迭代求解，与 Trigen_test&trunc.py 保持一致）--------------------
def _convert_trigen_to_triang(
    L: float, M: float, U: float, alpha: float, beta: float,
    max_iter: int = 100, tol: float = 1e-9
) -> Tuple[float, float, float]:
    """
    将三参数分位数信息转换为三角形分布参数 - 使用迭代方法
    确保返回的 a, c, b 满足 a < c < b
    """
    # 强制转换为浮点数
    try:
        L = float(L)
        M = float(M)
        U = float(U)
        alpha = float(alpha)
        beta = float(beta)
    except (TypeError, ValueError) as e:
        raise ValueError(f"Trigen 参数必须为数值，收到 L={L}, M={M}, U={U}, alpha={alpha}, beta={beta}") from e

    # 验证输入
    if not (L <= M <= U):
        raise ValueError(f"Trigen 参数需满足 L ≤ M ≤ U，收到 L={L}, M={M}, U={U}")
    if not (0.0 <= alpha < 1.0 and 0.0 < beta <= 1.0):
        raise ValueError("alpha 必须在 [0,1) 且 beta 必须在 (0,1]")
    if not (alpha < beta):
        raise ValueError("alpha 必须小于 beta")

    # 边界情况处理
    if alpha == 0.0:
        a_exact = L
    else:
        a_exact = None
    if beta == 1.0:
        b_exact = U
    else:
        b_exact = None

    # 如果都是边界情况，直接返回
    if a_exact is not None and b_exact is not None:
        return (a_exact, M, b_exact)

    # 初始估计（采用 Trigen_test&trunc.py 的方法）
    a_scale = L
    b_scale = U

    # 迭代求解
    def solve_A_given_B(B_val: float) -> float:
        # 从方程 (L-a)² = α(b-a)(c-a) 解出a
        a2 = 1.0 - alpha
        a1 = -2.0 * L + alpha * (B_val + M)
        a0 = L * L - alpha * B_val * M

        if abs(a2) < 1e-15:
            return L

        disc = a1 * a1 - 4.0 * a2 * a0
        if disc < 0:
            return None

        sqrt_d = math.sqrt(disc)
        r1 = (-a1 + sqrt_d) / (2.0 * a2)
        r2 = (-a1 - sqrt_d) / (2.0 * a2)

        candidates = [r1, r2]
        # 筛选出小于 M 的候选
        candidates = [x for x in candidates if x < M - 1e-12]
        if not candidates:
            return None
        # 返回最小的（最左边） - 与 Trigen_test&trunc.py 一致
        return min(candidates)

    def solve_B_given_A(A_val: float) -> float:
        # 从方程 (b-U)² = (1-β)(b-a)(b-c) 解出b
        b2 = beta
        b1 = -2.0 * U + (1.0 - beta) * (A_val + M)
        b0 = U * U - (1.0 - beta) * A_val * M

        if abs(b2) < 1e-15:
            return U

        disc = b1 * b1 - 4.0 * b2 * b0
        if disc < 0:
            return None

        sqrt_d = math.sqrt(disc)
        r1 = (-b1 + sqrt_d) / (2.0 * b2)
        r2 = (-b1 - sqrt_d) / (2.0 * b2)

        candidates = [r1, r2]
        # 筛选出大于 M 的候选
        candidates = [x for x in candidates if x > M + 1e-12]
        if not candidates:
            return None
        # 返回最大的（最右边） - 与 Trigen_test&trunc.py 一致
        return max(candidates)

    A = a_scale
    B = b_scale

    # 确保初始猜测有效（a < M < b）
    if not (A < M < B):
        A = M - (M - L) * 2
        B = M + (U - M) * 2

    # 迭代求解
    converged = False
    for i in range(max_iter):
        # 给定B求解A
        if a_exact is None:
            A_new = solve_A_given_B(B)
            if A_new is None:
                # 如果无解，使用之前的 A
                A_new = A
        else:
            A_new = a_exact

        # 给定A求解B
        if b_exact is None:
            B_new = solve_B_given_A(A_new)
            if B_new is None:
                B_new = B
        else:
            B_new = b_exact

        # 检查收敛
        if abs(A_new - A) < tol and abs(B_new - B) < tol:
            A, B = A_new, B_new
            converged = True
            break

        A, B = A_new, B_new

    if not converged:
        # 如果未收敛，使用初始估计 (L, U) - 与 Trigen_test&trunc.py 一致
        A, B = a_scale, b_scale

    return (A, M, B)

# -------------------- 生成器（直接复用三角分布） --------------------
def trigen_generator_single(rng: np.random.Generator, params: List[float]) -> float:
    L, M, U, alpha, beta = params[0], params[1], params[2], params[3], params[4]
    a, c, b = _convert_trigen_to_triang(L, M, U, alpha, beta)
    # 再次确保 a < c < b（防御性）
    if not (a < c < b):
        width = max(U - L, 1.0)
        a = c - 0.5 * width
        b = c + 0.5 * width
    return triang_generator_single(rng, [a, c, b])

def trigen_generator_vectorized(rng: np.random.Generator, params: List[float], n_samples: int) -> np.ndarray:
    L, M, U, alpha, beta = params[0], params[1], params[2], params[3], params[4]
    a, c, b = _convert_trigen_to_triang(L, M, U, alpha, beta)
    if not (a < c < b):
        width = max(U - L, 1.0)
        a = c - 0.5 * width
        b = c + 0.5 * width
    return triang_generator_vectorized(rng, [a, c, b], n_samples)

def trigen_generator(rng: np.random.Generator, params: List[float], n_samples: Optional[int] = None) -> Union[float, np.ndarray]:
    if n_samples is None:
        return trigen_generator_single(rng, params)
    else:
        return trigen_generator_vectorized(rng, params, n_samples)

# -------------------- 原始分布函数（内部转换后调用三角分布） --------------------
def trigen_cdf(x: float, L: float, M: float, U: float, alpha: float, beta: float) -> float:
    a, c, b = _convert_trigen_to_triang(L, M, U, alpha, beta)
    return triang_cdf(x, a, c, b)

def trigen_ppf(q: float, L: float, M: float, U: float, alpha: float, beta: float) -> float:
    a, c, b = _convert_trigen_to_triang(L, M, U, alpha, beta)
    return triang_ppf(q, a, c, b)

def trigen_pdf(x: float, L: float, M: float, U: float, alpha: float, beta: float) -> float:
    a, c, b = _convert_trigen_to_triang(L, M, U, alpha, beta)
    return triang_pdf(x, a, c, b)

# -------------------- 理论统计类（包装 TriangDistribution，修正截断委托） --------------------
class TrigenDistribution(DistributionBase):
    """
    三参数三角分布理论统计类。
    内部转换为三角分布后复用其功能。
    """

    def __init__(self, params: List[float], markers: dict = None, func_name: str = None):
        if len(params) < 5:
            L, M, U, alpha, beta = 0.0, 0.5, 1.0, 0.25, 0.75
        else:
            L, M, U, alpha, beta = float(params[0]), float(params[1]), float(params[2]), float(params[3]), float(params[4])
        # 转换为三角参数
        a, c, b = _convert_trigen_to_triang(L, M, U, alpha, beta)
        # 存储原始参数以供可能的调试
        self.L, self.M, self.U, self.alpha, self.beta = L, M, U, alpha, beta
        # 调用父类，但传递转换后的三角参数给内部三角分布对象
        super().__init__(params, markers, func_name)
        # 创建内部三角分布对象，明确指定函数名为 "DriskTriang" 以确保正确支撑
        self._triang_dist = TriangDistribution([a, c, b], markers, "DriskTriang")

        # ===== 继承三角分布的正确支持范围 =====
        self.support_low = self._triang_dist.support_low
        self.support_high = self._triang_dist.support_high
        # =====================================

        # 覆盖原始矩
        self._raw_mean = self._triang_dist._raw_mean
        self._raw_var = self._triang_dist._raw_var
        self._raw_skew = self._triang_dist._raw_skew
        self._raw_kurt = self._triang_dist._raw_kurt
        self._raw_mode = self._triang_dist._raw_mode
        # 截断矩的计算由 _triang_dist 处理，直接引用
        self._truncated_stats = self._triang_dist._truncated_stats
        # 继承有效性标记
        self._truncate_invalid = self._triang_dist._truncate_invalid

    # 委托方法：原始分布函数
    def _original_cdf(self, x: float) -> float:
        return self._triang_dist._original_cdf(x)

    def _original_ppf(self, q: float) -> float:
        return self._triang_dist._original_ppf(q)

    def _original_pdf(self, x: float) -> float:
        return self._triang_dist._original_pdf(x)

    # 委托方法：矩（已包含截断和平移）
    def mean(self) -> float:
        if self._truncate_invalid:
            return float('nan')
        return self._triang_dist.mean()

    def variance(self) -> float:
        if self._truncate_invalid:
            return float('nan')
        return self._triang_dist.variance()

    def skewness(self) -> float:
        if self._truncate_invalid:
            return float('nan')
        return self._triang_dist.skewness()

    def kurtosis(self) -> float:
        if self._truncate_invalid:
            return float('nan')
        return self._triang_dist.kurtosis()

    def mode(self) -> float:
        if self._truncate_invalid:
            return float('nan')
        return self._triang_dist.mode()

    def min_val(self) -> float:
        if self._truncate_invalid:
            return float('nan')
        return self._triang_dist.min_val()

    def max_val(self) -> float:
        if self._truncate_invalid:
            return float('nan')
        return self._triang_dist.max_val()

    # 关键修复：显式委托概率方法，确保截断生效
    def ppf(self, p: float) -> float:
        if self._truncate_invalid:
            return float('nan')
        return self._triang_dist.ppf(p)

    def cdf(self, x: float) -> float:
        if self._truncate_invalid:
            return float('nan')
        return self._triang_dist.cdf(x)

    def pdf(self, x: float) -> float:
        if self._truncate_invalid:
            return float('nan')
        return self._triang_dist.pdf(x)

    def pmf(self, x: float) -> float:
        # 对于连续分布，pmf 返回 0；但为了完整性，直接委托
        if self._truncate_invalid:
            return float('nan')
        return self._triang_dist.pmf(x)