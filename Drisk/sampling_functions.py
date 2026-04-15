# sampling_functions.py
"""采样方法模块 - 包含所有采样算法"""

import numpy as np
import threading
import time
import re
import math
from scipy.stats import qmc
from typing import List, Dict, Any, Callable, Tuple

# ==================== 全局常量和辅助函数 ====================

# 全局计数器，确保每次调用都不同
_random_counter = 0
_random_counter_lock = threading.RLock()

def create_rng(rng_type=1, seed=42):
    """
    创建随机数生成器
    
    Args:
        rng_type: 随机数生成器类型
                  1: Mersenne Twister (MT19937)
                  2: MRG32k3a
                  3: PCG64
                  4: Philox
                  5: SFC64
                  6: Threefry
                  7: Xoshiro256++
                  8: Xoroshiro128++
        seed: 随机数种子
    
    Returns:
        numpy.random.Generator对象
    """
    try:
        if rng_type == 1:  # Mersenne Twister
            return np.random.Generator(np.random.MT19937(seed=seed))
        elif rng_type == 2:  # MRG32k3a
            return np.random.Generator(np.random.MRG32k3a(seed=seed))
        elif rng_type == 3:  # PCG64
            return np.random.Generator(np.random.PCG64(seed=seed))
        elif rng_type == 4:  # Philox
            return np.random.Generator(np.random.Philox(seed=seed))
        elif rng_type == 5:  # SFC64
            return np.random.Generator(np.random.SFC64(seed=seed))
        elif rng_type == 6:  # Threefry
            return np.random.Generator(np.random.Threefry(seed=seed))
        elif rng_type == 7:  # Xoshiro256++
            return np.random.Generator(np.random.Xoshiro256(seed=seed))
        elif rng_type == 8:  # Xoroshiro128++
            return np.random.Generator(np.random.Xoroshiro128(seed=seed))
        else:
            return np.random.Generator(np.random.MT19937(seed=seed))
    except Exception as e:
        # 如果指定类型失败，回退到默认生成器
        print(f"创建RNG类型 {rng_type} 失败: {e}, 使用默认MT19937")
        return np.random.Generator(np.random.MT19937(seed=seed))

def get_unique_seed(user_seed=42, counter_offset=0):
    """
    生成唯一的随机数种子
    
    Args:
        user_seed: 用户指定的种子
        counter_offset: 计数器偏移量，用于区分不同分布
    
    Returns:
        唯一的种子值
    """
    with _random_counter_lock:
        global _random_counter
        _random_counter += 1
    
    # 创建唯一的种子
    base_seed = int(time.time() * 1000000) % 1000000000
    base_seed ^= threading.get_ident() % 1000000
    base_seed ^= _random_counter % 1000000
    base_seed ^= counter_offset  # 添加偏移量，使不同分布有不同的种子
    
    # 如果用户指定了种子，使用用户种子
    if user_seed != 42:
        final_seed = user_seed ^ base_seed % 1000000
    else:
        final_seed = base_seed
    
    return final_seed

# ==================== 采样算法 ====================

def generate_latin_hypercube_samples(n_samples: int, n_dimensions: int, seed=42) -> np.ndarray:
    """
    生成拉丁超立方样本 - 使用 SciPy 优化 LatinHypercube (random-cd)
    
    Args:
        n_samples: 样本数量
        n_dimensions: 维度数量
        seed: 随机种子
    
    Returns:
        n_samples × n_dimensions 的样本矩阵，值在 [0,1) 范围内
    """
    if n_samples <= 1:
        # 单样本时直接返回均匀随机数
        return np.random.default_rng(seed).random((n_samples, n_dimensions))

    if SCIPY_QMC_AVAILABLE:
        try:
            # 使用优化拉丁超立方 (random-cd)
            sampler = qmc.LatinHypercube(d=n_dimensions, optimization='random-cd', seed=seed)
            samples = sampler.random(n=n_samples)
            return samples  # 形状 (n_samples, n_dimensions)
        except Exception as e:
            print(f"使用 SciPy 优化拉丁超立方失败: {str(e)}，回退到基础方法。")
    
    # 回退：原始实现（分层随机 + 打乱）
    try:
        rng = np.random.RandomState(seed)
        samples = np.zeros((n_samples, n_dimensions))
        for d in range(n_dimensions):
            strata = np.linspace(0, 1, n_samples + 1)
            for i in range(n_samples):
                samples[i, d] = rng.uniform(strata[i], strata[i + 1])
            rng.shuffle(samples[:, d])
        return samples
    except Exception as e:
        print(f"生成拉丁超立方样本失败: {str(e)}")
        # 最终回退到简单随机抽样
        return np.random.RandomState(seed).random((n_samples, n_dimensions))
    
def _approx_norm_ppf(p):
    """
    正态分布逆累积分布函数的近似（标准正态分布）
    使用Wichura算法近似
    
    Args:
        p: 概率值，范围[0, 1]
    
    Returns:
        标准正态分布的分位数
    """
    p = np.asarray(p)
    mask = p < 0.5
    p = np.where(mask, p, 1.0 - p)
    
    # 近似公式
    t = np.sqrt(-2.0 * np.log(p))
    c0 = 2.515517
    c1 = 0.802853
    c2 = 0.010328
    d1 = 1.432788
    d2 = 0.189269
    d3 = 0.001308
    
    numerator = c0 + c1 * t + c2 * t * t
    denominator = 1.0 + d1 * t + d2 * t * t + d3 * t * t * t
    x = t - numerator / denominator
    
    # 调整符号
    x = np.where(mask, -x, x)
    return x

def generate_latin_hypercube_for_distributions(dist_info_list: List[Dict[str, Any]], 
                                              n_iterations: int, 
                                              seed=42, 
                                              method='random') -> Dict[str, np.ndarray]:
    """
    为多个分布生成拉丁超立方样本
    
    Args:
        dist_info_list: 分布信息列表，每个元素包含:
            - input_key: 输入键
            - dist_type: 分布类型（'normal', 'uniform', 'gamma'等）
            - params: 分布参数列表
            - cell_addr: 单元格地址
            - index: 索引
            - is_nested: 是否嵌套
        n_iterations: 迭代次数
        seed: 随机种子
        method: 采样方法
    
    Returns:
        字典：{input_key: samples_array}
    """
    print(f"为 {len(dist_info_list)} 个分布生成拉丁超立方样本，方法: {method}")
    
    n_distributions = len(dist_info_list)
    
    if n_distributions == 0:
        return {}
    
    # 生成拉丁超立方设计矩阵
    samples = generate_latin_hypercube_samples(n_iterations, n_distributions, seed)
    
    # 为每个分布生成样本
    all_samples = {}
    
    for i, dist_info in enumerate(dist_info_list):
        input_key = dist_info.get('input_key', f'unknown_{i}')
        dist_type = dist_info.get('dist_type', 'normal')
        params = dist_info.get('params', [0.0, 1.0])
        
        # 获取该分布的拉丁超立方样本列
        uniform_samples = samples[:, i]
        
        try:
            if dist_type == 'normal':
                # 正态分布：使用逆正态累积分布函数近似
                mean = float(params[0]) if len(params) > 0 else 0.0
                std = float(params[1]) if len(params) > 1 else 1.0
                
                if std <= 0:
                    all_samples[input_key] = np.full(n_iterations, mean, dtype=np.float64)
                else:
                    # 使用正态分布逆累积分布函数近似
                    normal_samples = _approx_norm_ppf(uniform_samples)
                    all_samples[input_key] = mean + std * normal_samples
            
            elif dist_type == 'uniform':
                # 均匀分布：直接缩放
                a = float(params[0]) if len(params) > 0 else 0.0
                b = float(params[1]) if len(params) > 1 else 1.0
                
                if b <= a:
                    all_samples[input_key] = np.full(n_iterations, a, dtype=np.float64)
                else:
                    all_samples[input_key] = a + uniform_samples * (b - a)
            
            elif dist_type == 'gamma':
                # 伽马分布：使用近似方法
                shape = float(params[0]) if len(params) > 0 else 1.0
                scale = float(params[1]) if len(params) > 1 else 1.0
                
                if shape <= 0 or scale <= 0:
                    all_samples[input_key] = np.zeros(n_iterations, dtype=np.float64)
                else:
                    # 对于伽马分布，使用简单的逆变换近似
                    # 当形状参数较大时，可以使用正态近似
                    if shape > 10:
                        # 正态近似
                        mean = shape * scale
                        std = np.sqrt(shape) * scale
                        normal_samples = _approx_norm_ppf(uniform_samples)
                        all_samples[input_key] = mean + std * normal_samples
                        # 确保非负
                        all_samples[input_key] = np.maximum(all_samples[input_key], 0.0)
                    else:
                        # 对于小形状参数，使用简单的近似
                        # 使用多个指数分布的和来近似
                        k = int(np.floor(shape))
                        remainder = shape - k
                        
                        # 生成k个指数分布的和
                        gamma_samples = np.zeros(n_iterations)
                        
                        # 使用逆变换方法生成指数分布
                        for j in range(k):
                            # 均匀分布转换为指数分布：-ln(1-u)/λ
                            u_sub = uniform_samples  # 使用不同的变换
                            exp_samples = -scale * np.log(1.0 - u_sub)
                            gamma_samples += exp_samples
                        
                        # 处理小数部分
                        if remainder > 0:
                            # 使用另一个均匀分布
                            u_rem = 1.0 - uniform_samples[::-1]  # 反向以增加随机性
                            exp_rem = -scale * np.log(1.0 - u_rem) * remainder
                            gamma_samples += exp_rem
                        
                        all_samples[input_key] = gamma_samples
            
            else:
                # 其他分布：默认使用正态分布
                print(f"警告: 分布类型 {dist_type} 未实现拉丁超立方，使用正态分布替代")
                mean = float(params[0]) if len(params) > 0 else 0.0
                std = float(params[1]) if len(params) > 1 else 1.0
                
                if std <= 0:
                    all_samples[input_key] = np.full(n_iterations, mean, dtype=np.float64)
                else:
                    normal_samples = _approx_norm_ppf(uniform_samples)
                    all_samples[input_key] = mean + std * normal_samples
        
        except Exception as e:
            print(f"为分布 {input_key} 生成拉丁超立方样本失败: {str(e)}")
            # 使用蒙特卡洛抽样作为后备
            if dist_type == 'normal':
                mean = float(params[0]) if len(params) > 0 else 0.0
                std = float(params[1]) if len(params) > 1 else 1.0
                all_samples[input_key] = np.random.normal(mean, std, n_iterations)
            elif dist_type == 'uniform':
                a = float(params[0]) if len(params) > 0 else 0.0
                b = float(params[1]) if len(params) > 1 else 1.0
                all_samples[input_key] = np.random.uniform(a, b, n_iterations)
            elif dist_type == 'gamma':
                shape = float(params[0]) if len(params) > 0 else 1.0
                scale = float(params[1]) if len(params) > 1 else 1.0
                all_samples[input_key] = np.random.gamma(shape, scale, n_iterations)
            else:
                all_samples[input_key] = np.random.normal(0, 1, n_iterations)
    
    print(f"拉丁超立方样本生成完成: {len(all_samples)} 组样本")
    return all_samples

def generate_batch_samples(dist_params: List[float], 
                          dist_type: str, 
                          n_samples: int, 
                          batch_size: int = 1000) -> np.ndarray:
    """
    批量生成样本
    
    Args:
        dist_params: 分布参数列表
        dist_type: 分布类型（'normal', 'uniform', 'gamma'等）
        n_samples: 总样本数
        batch_size: 批量大小
    
    Returns:
        生成的样本数组
    """
    # 计算需要多少个批次
    n_batches = (n_samples + batch_size - 1) // batch_size
    all_samples = []
    
    for batch_idx in range(n_batches):
        batch_start = batch_idx * batch_size
        batch_end = min((batch_idx + 1) * batch_size, n_samples)
        current_batch_size = batch_end - batch_start
        
        if current_batch_size <= 0:
            continue
        
        # 根据分布类型生成样本
        if dist_type == 'normal':
            mean = float(dist_params[0]) if len(dist_params) > 0 else 0.0
            std = float(dist_params[1]) if len(dist_params) > 1 else 1.0
            if std <= 0:
                batch_samples = np.full(current_batch_size, mean, dtype=np.float64)
            else:
                batch_samples = np.random.normal(mean, std, current_batch_size)
        
        elif dist_type == 'uniform':
            a = float(dist_params[0]) if len(dist_params) > 0 else 0.0
            b = float(dist_params[1]) if len(dist_params) > 1 else 1.0
            if b <= a:
                batch_samples = np.full(current_batch_size, a, dtype=np.float64)
            else:
                batch_samples = np.random.uniform(a, b, current_batch_size)
        
        elif dist_type == 'gamma':
            shape = float(dist_params[0]) if len(dist_params) > 0 else 1.0
            scale = float(dist_params[1]) if len(dist_params) > 1 else 1.0
            if shape <= 0 or scale <= 0:
                batch_samples = np.zeros(current_batch_size, dtype=np.float64)
            else:
                batch_samples = np.random.gamma(shape, scale, current_batch_size)
        
        elif dist_type == 'poisson':
            lam = float(dist_params[0]) if len(dist_params) > 0 else 1.0
            if lam <= 0:
                batch_samples = np.zeros(current_batch_size, dtype=np.float64)
            else:
                batch_samples = np.random.poisson(lam, current_batch_size)
        
        elif dist_type == 'beta':
            alpha = float(dist_params[0]) if len(dist_params) > 0 else 1.0
            beta = float(dist_params[1]) if len(dist_params) > 1 else 1.0
            if alpha <= 0 or beta <= 0:
                batch_samples = np.full(current_batch_size, 0.5, dtype=np.float64)
            else:
                batch_samples = np.random.beta(alpha, beta, current_batch_size)
        
        else:
            # 默认正态分布
            mean = float(dist_params[0]) if len(dist_params) > 0 else 0.0
            std = float(dist_params[1]) if len(dist_params) > 1 else 1.0
            batch_samples = np.random.normal(mean, std, current_batch_size)
        
        all_samples.append(batch_samples)
    
    # 合并所有批次的样本
    if all_samples:
        return np.concatenate(all_samples)[:n_samples]
    else:
        return np.zeros(n_samples, dtype=np.float64)

def generate_random_sample(dist_params: List[float], 
                          markers: Dict[str, Any],
                          rng_generator: Callable, 
                          loc_value: float,
                          counter_offset: int = 0) -> float:
    """
    生成随机样本的通用函数
    
    Args:
        dist_params: 分布参数列表
        markers: 标记字典
        rng_generator: 随机数生成函数
        loc_value: 理论均值（用于loc标记）
        counter_offset: 种子偏移量，确保不同分布使用不同种子
    
    Returns:
        随机样本值
    """
    # 检查是否有静态值标记
    static_value = markers.get('static')
    
    # 获取当前模式
    from attribute_functions import get_static_mode
    is_static_mode = get_static_mode()
    
    # 关键修改：如果存在static标记且处于静态模式，直接返回静态值
    if static_value is not None and is_static_mode:
        return float(static_value)
    
    # 检查是否有足够的分布参数
    if len(dist_params) < 1:
        return 0.0
    
    # 获取随机数生成器设置
    rng_type = markers.get('rng_type', 1)
    seed = markers.get('seed', 42)
    
    # 创建唯一的种子
    final_seed = get_unique_seed(seed, counter_offset)
    
    # 创建随机数生成器
    rng = create_rng(rng_type, final_seed)
    
    # 关键修改：在静态模式下，如果没有static标记，返回随机值
    # loc标记只在模拟模式下起作用
    if not is_static_mode:
        use_loc = markers.get('loc', False)
        if use_loc:
            # 模拟模式下且使用loc标记：返回理论均值
            return float(loc_value)

    # 生成原始样本值（可能需要后处理）
    try:
        sample = rng_generator(rng, dist_params)
        # 确保是标量浮点数
        val = float(np.asarray(sample).item())
    except Exception:
        # 如果生成失败，返回0作为后备
        val = 0.0

    # ----- 处理 shift -----
    shift_marker = markers.get('shift')
    if shift_marker is not None:
        try:
            shift_val = float(shift_marker)
            val = val + shift_val
        except Exception:
            # 忽略无效的 shift 标记
            pass

    # ----- 处理 truncate 和相关标记 -----
    # 支持多种形式：数值范围 ('low,high')、列表/元组、或单值
    def _parse_range(marker_val):
        if marker_val is None:
            return None, None
        if isinstance(marker_val, (list, tuple)) and len(marker_val) >= 1:
            try:
                low = float(marker_val[0])
                high = float(marker_val[1]) if len(marker_val) > 1 else low
                return low, high
            except Exception:
                return None, None
        if isinstance(marker_val, str):
            parts = [p.strip() for p in marker_val.split(',') if p.strip() != ""]
            try:
                if len(parts) == 1:
                    v = float(parts[0])
                    return v, v
                elif len(parts) >= 2:
                    low = float(parts[0])
                    high = float(parts[1])
                    return low, high
            except Exception:
                return None, None
        try:
            v = float(marker_val)
            return v, v
        except Exception:
            return None, None

    # 简单绝对截断
    trunc_low, trunc_high = _parse_range(markers.get('truncate'))
    if trunc_low is not None:
        # 直接裁剪
        if trunc_high is None:
            trunc_high = trunc_low
        val = min(max(val, trunc_low), trunc_high)

    # 百分比截断：通过小批量采样估计分位数，再裁剪当前值
    truncp_marker = markers.get('truncatep') or markers.get('truncatep2')
    if truncp_marker is not None:
        p_low, p_high = _parse_range(truncp_marker)
        if p_low is not None:
            # 接受 0-1 或 0-100 的表示
            if p_low > 1 or p_high > 1:
                p_low = max(0.0, min(1.0, p_low / 100.0))
                p_high = max(0.0, min(1.0, p_high / 100.0))

            # 为估计分位数从同一 RNG 中抽取小样本
            try:
                # 使用一个临时 RNG 副本，避免消耗主 RNG 的状态过多
                temp_rng = np.random.Generator(np.random.PCG64(int(final_seed) ^ 0xC0FFEE))
                # 生成一个小批量用于估计分位数
                batch_n = 1024
                batch = []
                for _ in range(batch_n):
                    try:
                        v = float(np.asarray(rng_generator(temp_rng, dist_params)).item())
                        batch.append(v)
                    except Exception:
                        continue
                if len(batch) >= 3:
                    lo_val = float(np.percentile(batch, p_low * 100.0))
                    hi_val = float(np.percentile(batch, p_high * 100.0))
                    # 裁剪当前值
                    val = min(max(val, lo_val), hi_val)
            except Exception:
                pass

    # 二次截断 truncate2：当存在时，先平移（如果有 shift），再按 truncate2 裁剪
    trunc2_low, trunc2_high = _parse_range(markers.get('truncate2'))
    if trunc2_low is not None:
        if trunc2_high is None:
            trunc2_high = trunc2_low
        val = min(max(val, trunc2_low), trunc2_high)

    # 最终确保返回浮点数
    try:
        return float(val)
    except Exception:
        return 0.0

def monte_carlo_sampling(dist_info_list: List[Dict[str, Any]], 
                        n_iterations: int, 
                        seed=42) -> Dict[str, np.ndarray]:
    """
    蒙特卡洛抽样
    
    Args:
        dist_info_list: 分布信息列表
        n_iterations: 迭代次数
        seed: 随机种子
    
    Returns:
        字典：{input_key: samples_array}
    """
    print(f"使用蒙特卡洛抽样生成 {n_iterations} 次迭代的样本")
    np.random.seed(seed)
    
    all_samples = {}
    
    for dist_info in dist_info_list:
        input_key = dist_info.get('input_key', 'unknown')
        dist_type = dist_info.get('dist_type', 'normal')
        params = dist_info.get('params', [0.0, 1.0])
        
        if dist_type == 'normal':
            mean = float(params[0]) if len(params) > 0 else 0.0
            std = float(params[1]) if len(params) > 1 else 1.0
            if std <= 0:
                all_samples[input_key] = np.full(n_iterations, mean, dtype=np.float64)
            else:
                all_samples[input_key] = np.random.normal(mean, std, n_iterations)
        
        elif dist_type == 'uniform':
            a = float(params[0]) if len(params) > 0 else 0.0
            b = float(params[1]) if len(params) > 1 else 1.0
            if b <= a:
                all_samples[input_key] = np.full(n_iterations, a, dtype=np.float64)
            else:
                all_samples[input_key] = np.random.uniform(a, b, n_iterations)
        
        elif dist_type == 'gamma':
            shape = float(params[0]) if len(params) > 0 else 1.0
            scale = float(params[1]) if len(params) > 1 else 1.0
            if shape <= 0 or scale <= 0:
                all_samples[input_key] = np.zeros(n_iterations, dtype=np.float64)
            else:
                all_samples[input_key] = np.random.gamma(shape, scale, n_iterations)
        
        elif dist_type == 'poisson':
            lam = float(params[0]) if len(params) > 0 else 1.0
            if lam <= 0:
                all_samples[input_key] = np.zeros(n_iterations, dtype=np.float64)
            else:
                all_samples[input_key] = np.random.poisson(lam, n_iterations)
        
        else:
            # 默认正态分布
            mean = float(params[0]) if len(params) > 0 else 0.0
            std = float(params[1]) if len(params) > 1 else 1.0
            all_samples[input_key] = np.random.normal(mean, std, n_iterations)
    
    print(f"蒙特卡洛样本生成完成: {len(all_samples)} 组样本")
    return all_samples

def generate_latin_hypercube_1d(n_samples: int, seed=42) -> np.ndarray:
    """
    生成一维拉丁超立方样本（分层均匀随机数），使用优化拉丁超立方。
    
    Args:
        n_samples: 样本数量
        seed: 随机种子
    
    Returns:
        长度为 n_samples 的数组，包含在 [0,1] 内分层均匀的随机数
    """
    if n_samples <= 1:
        return np.random.default_rng(seed).random(n_samples)
    
    if SCIPY_QMC_AVAILABLE:
        try:
            sampler = qmc.LatinHypercube(d=1, optimization='random-cd', seed=seed)
            samples = sampler.random(n=n_samples)
            return samples.flatten()
        except Exception as e:
            print(f"一维优化拉丁超立方失败: {e}，回退到基础方法。")
    
    # 原始回退方法
    rng = np.random.default_rng(seed)
    intervals = np.linspace(0, 1, n_samples + 1)
    u = np.zeros(n_samples)
    for i in range(n_samples):
        u[i] = rng.uniform(intervals[i], intervals[i+1])
    rng.shuffle(u)
    return u

# ==================== Sobol 序列生成 ====================
try:
    from scipy.stats import qmc
    SCIPY_QMC_AVAILABLE = True
except ImportError:
    SCIPY_QMC_AVAILABLE = False
    print("警告：SciPy 不可用，Sobol 采样将回退到均匀随机数（蒙特卡洛）。建议安装 scipy 以获得更好的准蒙特卡洛序列。")

def generate_sobol_1d(n_samples: int, seed=42, scramble=True) -> np.ndarray:
    """
    生成一维 Sobol 序列（准蒙特卡洛），返回在 [0,1] 内均匀分布的点。
    
    Args:
        n_samples: 样本数量
        seed: 随机种子（用于 scramble）
        scramble: 是否对序列进行打乱（默认 True）
    
    Returns:
        长度为 n_samples 的数组，包含 Sobol 点。
    """
    if not SCIPY_QMC_AVAILABLE:
        # 回退到均匀随机数
        return np.random.default_rng(seed).random(n_samples)
    try:
        # 使用 SciPy 的 Sobol 引擎
        engine = qmc.Sobol(d=1, scramble=scramble, seed=seed)
        samples = engine.random(n_samples)
        return samples.flatten()
    except Exception as e:
        print(f"生成 Sobol 序列失败: {e}，回退到均匀随机数。")
        return np.random.default_rng(seed).random(n_samples)