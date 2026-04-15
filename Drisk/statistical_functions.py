# statistical_functions.py
"""统计函数模块 - 修正百分位数范围版本，支持场景模拟ID，增强暂停模拟处理"""

from typing import Optional
import numpy as np
from pyxll import xl_func, xl_arg
from simulation_manager import get_simulation
import warnings
warnings.filterwarnings("ignore", message="Inconsistent type specified for")

def _filter_valid_data(data):
    """过滤有效数据，剔除ERROR标记、NaN和None值"""
    if data is None:
        return np.array([])
    
    valid_data = []
    for val in data:
        # 跳过None
        if val is None:
            continue
        # 跳过字符串错误标记
        if isinstance(val, str) and val == "#ERROR!":
            continue
        # 跳过NaN值
        try:
            if isinstance(val, (int, float, np.number)):
                if np.isnan(val):
                    continue
                valid_data.append(float(val))
            elif isinstance(val, str):
                # 尝试转换其他字符串
                try:
                    float_val = float(val)
                    if np.isnan(float_val):
                        continue
                    valid_data.append(float_val)
                except:
                    continue
            else:
                # 其他类型尝试转换
                try:
                    float_val = float(val)
                    if np.isnan(float_val):
                        continue
                    valid_data.append(float_val)
                except:
                    continue
        except (TypeError, ValueError):
            continue
    
    return np.array(valid_data)

@xl_func("xl_cell cell, int sim_num, var trunc: float", volatile=True)
@xl_arg("cell", "range")
def DriskMean(cell, sim_num: int = 1, trunc=None) -> float:
    """计算模拟分布均值 - 支持截断参数
    
    参数：
        cell: 输出单元格
        sim_num: 模拟ID（场景数）
        trunc: 可选截断参数，支持DriskTruncate或DriskTruncateP标记字符串
               支持单边值截断，如DriskTruncate(, max)或DriskTruncateP(low,)
    """
    try:
        # 获取模拟对象 - 现在sim_num直接对应场景ID
        sim = get_simulation(sim_num)
        if sim is None:
            # 尝试获取任何可用的模拟
            from simulation_manager import get_all_simulations
            all_sims = get_all_simulations()
            if all_sims:
                sim = next(iter(all_sims.values()))
            else:
                return float('nan')
        
        # 检查是否是Output单元格
        if not sim.is_output_cell(cell):
            # 即使不是标记的Output单元格，也尝试获取数据
            data = sim.get_output_data_by_range(cell)
            if data is None:
                return float('nan')
        else:
            # 获取数据
            data = sim.get_output_data_by_range(cell)
            if data is None:
                return float('nan')
        
        # 过滤有效数据
        valid_data = _filter_valid_data(data)
        if len(valid_data) == 0:
            return float('nan')
        
        # 应用截断处理（如果提供了trunc参数）
        if trunc is not None and isinstance(trunc, str):
            processed_data = _apply_truncation(valid_data, trunc)
            if processed_data is None:
                return float('nan')  # 截断参数无效
            valid_data = processed_data
            if len(valid_data) == 0:
                return float('nan')
        
        # 计算均值
        return float(np.mean(valid_data))
        
    except Exception as e:
        print(f"DriskMean错误: {str(e)}")
        return float('nan')


def _apply_truncation(data: np.ndarray, trunc_str: str) -> Optional[np.ndarray]:
    """
    应用截断处理，支持单边值截断
    
    参数：
        data: 原始数据数组
        trunc_str: 截断标记字符串，支持两种格式：
                   - DriskTruncate标记: "__DRISK_TRUNCATE__:min,max" 支持单边值，如",max"或"min,"
                   - DriskTruncateP标记: "__DRISK_TRUNCATEP__:low_prob,up_prob" 
                                         输入必须在[0,1]范围内，支持单边值，如",0.95"或"0.05,"
    
    返回：
        截断后的数据数组，如果截断参数无效则返回None
    """
    try:
        # 检查是否是标记字符串
        if not isinstance(trunc_str, str):
            return data
        
        # 检查标记类型
        if trunc_str.startswith("__DRISK_TRUNCATE__:"):
            # DriskTruncate处理：绝对范围截断
            params_str = trunc_str[len("__DRISK_TRUNCATE__:"):]
            
            # 解析参数，允许空字符串
            params = params_str.split(',')
            if len(params) < 2:
                print(f"DriskTruncate参数不足: {params_str}")
                return data
            
            min_val_str = params[0].strip()
            max_val_str = params[1].strip()
            
            # 处理单边值
            has_min = min_val_str != ''
            has_max = max_val_str != ''
            
            if not has_min and not has_max:
                # 两个参数都为空，返回原始数据
                return data
            
            try:
                min_val = float(min_val_str) if has_min else -np.inf
                max_val = float(max_val_str) if has_max else np.inf
                
                # 验证范围
                if min_val > max_val:
                    min_val, max_val = max_val, min_val
                
                # 应用截断
                mask = np.ones_like(data, dtype=bool)
                if has_min:
                    mask = mask & (data >= min_val)
                if has_max:
                    mask = mask & (data <= max_val)
                
                truncated = data[mask]
                print(f"DriskTruncate应用: 范围[{min_val if has_min else '-∞'}, {max_val if has_max else '∞'}], "
                      f"原始数据{len(data)}条, 截断后{len(truncated)}条")
                return truncated
            except ValueError as ve:
                print(f"DriskTruncate参数解析失败: {params_str}, 错误: {ve}")
                return None
        
        elif trunc_str.startswith("__DRISK_TRUNCATEP__:"):
            # DriskTruncateP处理：概率范围截断，输入必须在[0,1]范围内
            params_str = trunc_str[len("__DRISK_TRUNCATEP__:"):]
            
            # 解析参数，允许空字符串
            params = params_str.split(',')
            if len(params) < 2:
                print(f"DriskTruncateP参数不足: {params_str}")
                return data
            
            low_prob_str = params[0].strip()
            up_prob_str = params[1].strip()
            
            # 处理单边值
            has_low = low_prob_str != ''
            has_up = up_prob_str != ''
            
            if not has_low and not has_up:
                # 两个参数都为空，返回原始数据
                return data
            
            try:
                # 解析概率值（必须是[0,1]范围内）
                low_prob = float(low_prob_str) if has_low else 0.0
                up_prob = float(up_prob_str) if has_up else 1.0
                
                # 验证概率范围必须为[0,1]
                if has_low and (low_prob < 0 or low_prob > 1):
                    print(f"DriskTruncateP低概率值超出[0,1]范围: {low_prob}")
                    return None
                
                if has_up and (up_prob < 0 or up_prob > 1):
                    print(f"DriskTruncateP高概率值超出[0,1]范围: {up_prob}")
                    return None
                
                # 确保低概率值小于高概率值
                if low_prob > up_prob:
                    low_prob, up_prob = up_prob, low_prob
                
                # 转换为百分位数
                low_percentile = low_prob * 100
                up_percentile = up_prob * 100
                
                if len(data) > 0:
                    # 计算分位数
                    low_val = np.percentile(data, low_percentile) if has_low else -np.inf
                    up_val = np.percentile(data, up_percentile) if has_up else np.inf
                    
                    # 应用截断
                    mask = np.ones_like(data, dtype=bool)
                    if has_low:
                        mask = mask & (data >= low_val)
                    if has_up:
                        mask = mask & (data <= up_val)
                    
                    truncated = data[mask]
                    print(f"DriskTruncateP应用: 概率范围[{low_prob if has_low else 0}, {up_prob if has_up else 1}], "
                          f"对应值范围[{low_val:.4f if has_low else '-∞'}, {up_val:.4f if has_up else '∞'}], "
                          f"原始数据{len(data)}条, 截断后{len(truncated)}条")
                    return truncated
            except ValueError as ve:
                print(f"DriskTruncateP参数解析失败: {params_str}, 错误: {ve}")
                return None
        
        # 如果不是有效的截断标记，返回原始数据
        print(f"警告: 无法识别的截断标记: {trunc_str}")
        return data
        
    except Exception as e:
        print(f"截断处理错误: {str(e)}")
        return None
    
@xl_func("xl_cell cell, int sim_num, var trunc: float", volatile=True)
@xl_arg("cell", "range")
def DriskStd(cell, sim_num: int = 1, trunc=None) -> float:
    """计算模拟分布标准差 - 支持截断参数
    
    参数：
        cell: 输出单元格
        sim_num: 模拟ID（场景数）
        trunc: 可选截断参数，支持DriskTruncate或DriskTruncateP标记字符串
               支持单边值截断，如DriskTruncate(, max)或DriskTruncateP(low,)
    """
    try:
        # 获取模拟对象 - 现在sim_num直接对应场景ID
        sim = get_simulation(sim_num)
        if sim is None:
            # 尝试获取任何可用的模拟
            from simulation_manager import get_all_simulations
            all_sims = get_all_simulations()
            if all_sims:
                sim = next(iter(all_sims.values()))
            else:
                return float('nan')
        
        # 获取数据
        data = sim.get_output_data_by_range(cell)
        if data is None:
            return float('nan')
        
        # 过滤有效数据
        valid_data = _filter_valid_data(data)
        if len(valid_data) < 2:  # 至少需要2个数据点计算标准差
            return float('nan')
        
        # 应用截断处理（如果提供了trunc参数）
        if trunc is not None and isinstance(trunc, str):
            processed_data = _apply_truncation(valid_data, trunc)
            if processed_data is None:
                return float('nan')  # 截断参数无效
            valid_data = processed_data
            if len(valid_data) < 2:
                return float('nan')
        
        return float(np.std(valid_data, ddof=1))  # 样本标准差
        
    except Exception as e:
        print(f"DriskStd错误: {str(e)}")
        return float('nan')

@xl_func("xl_cell cell, int sim_num, var trunc: float", volatile=True)
@xl_arg("cell", "range")
def DriskVariance(cell, sim_num: int = 1, trunc=None) -> float:
    """计算模拟分布方差 - 支持截断参数
    
    参数：
        cell: 输出单元格
        sim_num: 模拟ID（场景数）
        trunc: 可选截断参数，支持DriskTruncate或DriskTruncateP标记字符串
               支持单边值截断，如DriskTruncate(, max)或DriskTruncateP(low,)
    """
    try:
        # 获取模拟对象 - 现在sim_num直接对应场景ID
        sim = get_simulation(sim_num)
        if sim is None:
            # 尝试获取任何可用的模拟
            from simulation_manager import get_all_simulations
            all_sims = get_all_simulations()
            if all_sims:
                sim = next(iter(all_sims.values()))
            else:
                return float('nan')
        
        # 获取数据
        data = sim.get_output_data_by_range(cell)
        if data is None:
            return float('nan')
        
        # 过滤有效数据
        valid_data = _filter_valid_data(data)
        if len(valid_data) < 2:  # 至少需要2个数据点计算方差
            return float('nan')
        
        # 应用截断处理（如果提供了trunc参数）
        if trunc is not None and isinstance(trunc, str):
            processed_data = _apply_truncation(valid_data, trunc)
            if processed_data is None:
                return float('nan')  # 截断参数无效
            valid_data = processed_data
            if len(valid_data) < 2:
                return float('nan')
        
        return float(np.var(valid_data, ddof=1))  # 样本方差
        
    except Exception as e:
        print(f"DriskVariance错误: {str(e)}")
        return float('nan')

@xl_func("xl_cell cell, int sim_num, var trunc: float", volatile=True)
@xl_arg("cell", "range")
def DriskMin(cell, sim_num: int = 1, trunc=None) -> float:
    """计算模拟分布最小值 - 支持截断参数
    
    参数：
        cell: 输出单元格
        sim_num: 模拟ID（场景数）
        trunc: 可选截断参数，支持DriskTruncate或DriskTruncateP标记字符串
               支持单边值截断，如DriskTruncate(, max)或DriskTruncateP(low,)
    """
    try:
        # 获取模拟对象 - 现在sim_num直接对应场景ID
        sim = get_simulation(sim_num)
        if sim is None:
            # 尝试获取任何可用的模拟
            from simulation_manager import get_all_simulations
            all_sims = get_all_simulations()
            if all_sims:
                sim = next(iter(all_sims.values()))
            else:
                return float('nan')
        
        # 获取数据
        data = sim.get_output_data_by_range(cell)
        if data is None:
            return float('nan')
        
        # 过滤有效数据
        valid_data = _filter_valid_data(data)
        if len(valid_data) == 0:
            return float('nan')
        
        # 应用截断处理（如果提供了trunc参数）
        if trunc is not None and isinstance(trunc, str):
            processed_data = _apply_truncation(valid_data, trunc)
            if processed_data is None:
                return float('nan')  # 截断参数无效
            valid_data = processed_data
            if len(valid_data) == 0:
                return float('nan')
        
        return float(np.min(valid_data))
        
    except Exception as e:
        print(f"DriskMin错误: {str(e)}")
        return float('nan')

@xl_func("xl_cell cell, int sim_num, var trunc: float", volatile=True)
@xl_arg("cell", "range")
def DriskMax(cell, sim_num: int = 1, trunc=None) -> float:
    """计算模拟分布最大值 - 支持截断参数
    
    参数：
        cell: 输出单元格
        sim_num: 模拟ID（场景数）
        trunc: 可选截断参数，支持DriskTruncate或DriskTruncateP标记字符串
               支持单边值截断，如DriskTruncate(, max)或DriskTruncateP(low,)
    """
    try:
        # 获取模拟对象 - 现在sim_num直接对应场景ID
        sim = get_simulation(sim_num)
        if sim is None:
            # 尝试获取任何可用的模拟
            from simulation_manager import get_all_simulations
            all_sims = get_all_simulations()
            if all_sims:
                sim = next(iter(all_sims.values()))
            else:
                return float('nan')
        
        # 获取数据
        data = sim.get_output_data_by_range(cell)
        if data is None:
            return float('nan')
        
        # 过滤有效数据
        valid_data = _filter_valid_data(data)
        if len(valid_data) == 0:
            return float('nan')
        
        # 应用截断处理（如果提供了trunc参数）
        if trunc is not None and isinstance(trunc, str):
            processed_data = _apply_truncation(valid_data, trunc)
            if processed_data is None:
                return float('nan')  # 截断参数无效
            valid_data = processed_data
            if len(valid_data) == 0:
                return float('nan')
        
        return float(np.max(valid_data))
        
    except Exception as e:
        print(f"DriskMax错误: {str(e)}")
        return float('nan')

@xl_func("xl_cell cell, int sim_num, var trunc: float", volatile=True)
@xl_arg("cell", "range")
def DriskMedian(cell, sim_num: int = 1, trunc=None) -> float:
    """计算模拟分布中位数 - 支持截断参数
    
    参数：
        cell: 输出单元格
        sim_num: 模拟ID（场景数）
        trunc: 可选截断参数，支持DriskTruncate或DriskTruncateP标记字符串
               支持单边值截断，如DriskTruncate(, max)或DriskTruncateP(low,)
    """
    try:
        # 获取模拟对象 - 现在sim_num直接对应场景ID
        sim = get_simulation(sim_num)
        if sim is None:
            # 尝试获取任何可用的模拟
            from simulation_manager import get_all_simulations
            all_sims = get_all_simulations()
            if all_sims:
                sim = next(iter(all_sims.values()))
            else:
                return float('nan')
        
        # 获取数据
        data = sim.get_output_data_by_range(cell)
        if data is None:
            return float('nan')
        
        # 过滤有效数据
        valid_data = _filter_valid_data(data)
        if len(valid_data) == 0:
            return float('nan')
        
        # 应用截断处理（如果提供了trunc参数）
        if trunc is not None and isinstance(trunc, str):
            processed_data = _apply_truncation(valid_data, trunc)
            if processed_data is None:
                return float('nan')  # 截断参数无效
            valid_data = processed_data
            if len(valid_data) == 0:
                return float('nan')
        
        return float(np.median(valid_data))
        
    except Exception as e:
        print(f"DriskMedian错误: {str(e)}")
        return float('nan')

@xl_func("xl_cell cell, float percent_value, int sim_num, var trunc: float", volatile=True)
@xl_arg("cell", "range")
def DriskPtoX(cell, percent_value: float, sim_num: int = 1, trunc=None) -> float:
    """
    返回指定概率对应的x值 - 支持截断参数
    
    参数：
        cell: 输出单元格
        percent_value: 概率值，范围[0,1]
        sim_num: 模拟ID（场景数）
        trunc: 可选截断参数，支持DriskTruncate或DriskTruncateP标记字符串
               支持单边值截断，如DriskTruncate(, max)或DriskTruncateP(low,)
    """
    try:
        # 修正：检查范围改为[0,1]
        if not (0 <= percent_value <= 1):
            return float('nan')
        
        # 获取模拟对象 - 现在sim_num直接对应场景ID
        sim = get_simulation(sim_num)
        if sim is None:
            # 尝试获取任何可用的模拟
            from simulation_manager import get_all_simulations
            all_sims = get_all_simulations()
            if all_sims:
                sim = next(iter(all_sims.values()))
            else:
                return float('nan')
        
        # 获取数据
        data = sim.get_output_data_by_range(cell)
        if data is None:
            return float('nan')
        
        # 过滤有效数据
        valid_data = _filter_valid_data(data)
        if len(valid_data) == 0:
            return float('nan')
        
        # 应用截断处理（如果提供了trunc参数）
        if trunc is not None and isinstance(trunc, str):
            processed_data = _apply_truncation(valid_data, trunc)
            if processed_data is None:
                return float('nan')  # 截断参数无效
            valid_data = processed_data
            if len(valid_data) == 0:
                return float('nan')
        
        sorted_data = np.sort(valid_data)
        percentile = percent_value * 100
        return float(np.percentile(sorted_data, percentile))
        
    except Exception as e:
        print(f"DriskPtoX错误: {str(e)}")
        return float('nan')

@xl_func("xl_cell cell, int iteration_num, int sim_num: float", volatile=True)
@xl_arg("cell", "range")
def DriskData(cell, iteration_num: int, sim_num: int = 1) -> float:
    """返回指定迭代次数下对应单元格的值 - 只能处理Output单元格
    
    参数：
        cell: 输出单元格
        iteration_num: 迭代次数（从1开始）
        sim_num: 模拟ID（场景数）
    """
    try:
        if iteration_num < 1:
            return float('nan')
        
        # 获取模拟对象 - 现在sim_num直接对应场景ID
        sim = get_simulation(sim_num)
        if sim is None:
            # 尝试获取任何可用的模拟
            from simulation_manager import get_all_simulations
            all_sims = get_all_simulations()
            if all_sims:
                sim = next(iter(all_sims.values()))
            else:
                return float('nan')
        
        # 检查是否是Output单元格
        if not sim.is_output_cell(cell):
            # 即使不是标记的Output单元格，也尝试获取数据
            data = sim.get_output_data_by_range(cell)
            if data is None:
                return float('nan')
            
            # 检查迭代范围
            if iteration_num > len(data):
                return float('nan')
            
            # 获取值
            value = data[iteration_num - 1]
        else:
            # 获取数据
            data = sim.get_output_data_by_range(cell)
            if data is None:
                return float('nan')
            
            # 检查迭代范围
            if iteration_num > len(data):
                return float('nan')
            
            # 获取值
            value = data[iteration_num - 1]
        
        # 检查是否是错误标记
        if isinstance(value, str) and value == "#ERROR!":
            return float('nan')
        
        # 尝试转换为数值
        try:
            if isinstance(value, (int, float, np.number)):
                if np.isnan(value):
                    return float('nan')
                return float(value)
            elif isinstance(value, str):
                # 尝试转换字符串
                try:
                    return float(value)
                except:
                    return float('nan')
            else:
                return float('nan')
        except:
            return float('nan')
    except Exception as e:
        print(f"DriskData错误: {str(e)}")
        return float('nan')

# ==================== 其他统计函数 ====================

@xl_func("xl_cell cell, int sim_num, var trunc: float", volatile=True)
@xl_arg("cell", "range")
def DriskSkew(cell, sim_num: int = 1, trunc=None) -> float:
    """计算模拟数据偏度 - 支持截断参数
    
    参数：
        cell: 输出单元格
        sim_num: 模拟ID（场景数）
        trunc: 可选截断参数，支持DriskTruncate或DriskTruncateP标记字符串
               支持单边值截断，如DriskTruncate(, max)或DriskTruncateP(low,)
    """
    try:
        # 获取模拟对象 - 现在sim_num直接对应场景ID
        sim = get_simulation(sim_num)
        if sim is None:
            # 尝试获取任何可用的模拟
            from simulation_manager import get_all_simulations
            all_sims = get_all_simulations()
            if all_sims:
                sim = next(iter(all_sims.values()))
            else:
                return float('nan')
        
        # 获取数据
        data = sim.get_output_data_by_range(cell)
        if data is None:
            return float('nan')
        
        # 过滤有效数据
        valid_data = _filter_valid_data(data)
        if len(valid_data) < 3:
            return float('nan')
        
        # 应用截断处理（如果提供了trunc参数）
        if trunc is not None and isinstance(trunc, str):
            processed_data = _apply_truncation(valid_data, trunc)
            if processed_data is None:
                return float('nan')  # 截断参数无效
            valid_data = processed_data
            if len(valid_data) < 3:
                return float('nan')
        
        # 计算偏度
        try:
            from scipy.stats import skew
            return float(skew(valid_data))
        except ImportError:
            # 备用计算方法
            n = len(valid_data)
            if n < 3:
                return float('nan')
            
            mean = np.mean(valid_data)
            std = np.std(valid_data, ddof=1)
            if std == 0:
                return 0.0
            
            skewness = np.sum(((valid_data - mean) / std) ** 3) / n
            # 调整偏差
            skewness = skewness * (n / ((n - 1) * (n - 2))) ** 0.5
            return float(skewness)
    except Exception as e:
        print(f"DriskSkew错误: {str(e)}")
        return float('nan')


@xl_func("xl_cell cell, int sim_num, var trunc: float", volatile=True)
@xl_arg("cell", "range")
def DriskKurt(cell, sim_num: int = 1, trunc=None) -> float:
    """计算峰度（Pearson，等于excess kurtosis + 3） - 支持截断参数
    
    参数：
        cell: 输出单元格
        sim_num: 模拟ID（场景数）
        trunc: 可选截断参数，支持DriskTruncate或DriskTruncateP标记字符串
               支持单边值截断，如DriskTruncate(, max)或DriskTruncateP(low,)
    """
    try:
        # 获取模拟对象 - 现在sim_num直接对应场景ID
        sim = get_simulation(sim_num)
        if sim is None:
            # 尝试获取任何可用的模拟
            from simulation_manager import get_all_simulations
            all_sims = get_all_simulations()
            if all_sims:
                sim = next(iter(all_sims.values()))
            else:
                return float('nan')
        
        # 获取数据
        data = sim.get_output_data_by_range(cell)
        if data is None:
            return float('nan')
        
        # 过滤有效数据
        valid_data = _filter_valid_data(data)
        if len(valid_data) < 4:
            return float('nan')
        
        # 应用截断处理（如果提供了trunc参数）
        if trunc is not None and isinstance(trunc, str):
            processed_data = _apply_truncation(valid_data, trunc)
            if processed_data is None:
                return float('nan')  # 截断参数无效
            valid_data = processed_data
            if len(valid_data) < 4:
                return float('nan')
        
        try:
            from scipy.stats import kurtosis
            # fisher=False 返回 Pearson kurtosis (excess+3)
            return float(kurtosis(valid_data, fisher=False, bias=False))
        except Exception:
            # 备用计算：使用样本中心化四阶矩估计（近似 Pearson 峰度）
            arr = np.array(valid_data)
            m = np.mean(arr)
            sd = np.std(arr)
            if sd == 0:
                return 3.0
            kurt = np.mean(((arr - m) / sd) ** 4)
            return float(kurt)
    except Exception as e:
        print(f"DriskKurt错误: {str(e)}")
        return float('nan')


@xl_func("xl_cell cell, int sim_num, var trunc: float", volatile=True)
@xl_arg("cell", "range")
def DriskRange(cell, sim_num: int = 1, trunc=None) -> float:
    """计算范围（最大值 - 最小值） - 支持截断参数
    
    参数：
        cell: 输出单元格
        sim_num: 模拟ID（场景数）
        trunc: 可选截断参数，支持DriskTruncate或DriskTruncateP标记字符串
               支持单边值截断，如DriskTruncate(, max)或DriskTruncateP(low,)
    """
    try:
        # 获取模拟对象 - 现在sim_num直接对应场景ID
        sim = get_simulation(sim_num)
        if sim is None:
            # 尝试获取任何可用的模拟
            from simulation_manager import get_all_simulations
            all_sims = get_all_simulations()
            if all_sims:
                sim = next(iter(all_sims.values()))
            else:
                return float('nan')
        
        # 获取数据
        data = sim.get_output_data_by_range(cell)
        if data is None:
            return float('nan')
        
        # 过滤有效数据
        valid_data = _filter_valid_data(data)
        if len(valid_data) == 0:
            return float('nan')
        
        # 应用截断处理（如果提供了trunc参数）
        if trunc is not None and isinstance(trunc, str):
            processed_data = _apply_truncation(valid_data, trunc)
            if processed_data is None:
                return float('nan')  # 截断参数无效
            valid_data = processed_data
            if len(valid_data) == 0:
                return float('nan')
        
        return float(np.max(valid_data) - np.min(valid_data))
        
    except Exception as e:
        print(f"DriskRange错误: {str(e)}")
        return float('nan')


@xl_func("xl_cell cell, int sim_num, var trunc, str mode_method: float", volatile=True)
@xl_arg("cell", "range")
def DriskMode(cell, sim_num: int = 1, trunc=None, mode_method: str = "auto") -> float:
    """
    计算模拟分布的众数（出现频率最高的数值），支持连续分布近似。
    
    参数：
        cell: 输出单元格
        sim_num: 模拟ID（场景数）
        trunc: 可选截断参数，支持DriskTruncate或DriskTruncateP标记字符串
        mode_method: 众数计算方法
            - "discrete" : 直接使用原始值计数（适用于离散或有限取值数据）
            - "histogram": 基于直方图分组，取组中值作为众数
            - "kde"      : 基于核密度估计，取密度最大值对应的数值
            - "auto"     : 自动选择（当唯一值数量 < sqrt(n) 时用 discrete，否则用 kde）
    
    返回：
        众数值（浮点数），若无法计算则返回 NaN
    """
    try:
        # 1. 获取模拟对象和数据
        sim = get_simulation(sim_num)
        if sim is None:
            from simulation_manager import get_all_simulations
            all_sims = get_all_simulations()
            if all_sims:
                sim = next(iter(all_sims.values()))
            else:
                return float('nan')
        
        data = sim.get_output_data_by_range(cell)
        if data is None:
            return float('nan')
        
        # 2. 过滤有效数据
        valid_data = _filter_valid_data(data)
        if len(valid_data) == 0:
            return float('nan')
        
        # 3. 应用截断
        if trunc is not None and isinstance(trunc, str):
            processed_data = _apply_truncation(valid_data, trunc)
            if processed_data is None or len(processed_data) == 0:
                return float('nan')
            valid_data = processed_data
        
        arr = np.array(valid_data)
        n = len(arr)
        
        # 4. 根据方法计算众数
        def _mode_discrete(data):
            """原始离散众数：使用Counter"""
            from collections import Counter
            cnt = Counter(data)
            max_count = max(cnt.values())
            modes = [k for k, v in cnt.items() if v == max_count]
            return float(min(modes))
        
        def _mode_histogram(data, bins='auto'):
            """直方图众数：取频数最高的组中值"""
            hist, bin_edges = np.histogram(data, bins=bins)
            max_bin_idx = np.argmax(hist)
            # 组中值 = (左边界 + 右边界) / 2
            bin_center = (bin_edges[max_bin_idx] + bin_edges[max_bin_idx + 1]) / 2.0
            return float(bin_center)
        
        def _mode_kde(data):
            """核密度估计众数：最大化密度函数"""
            try:
                from scipy.stats import gaussian_kde
                from scipy.optimize import minimize_scalar
            except ImportError:
                # 若scipy不可用，降级为直方图方法
                print("scipy未安装，降级使用直方图方法计算众数")
                return _mode_histogram(data)
            
            # 带宽选择：Scott's rule
            kde = gaussian_kde(data, bw_method='scott')
            # 在数据范围内最大化密度
            x_min, x_max = np.min(data), np.max(data)
            # 增加一点边界扩展
            pad = (x_max - x_min) * 0.05
            x_min -= pad
            x_max += pad
            
            # 负密度用于最小化
            def neg_density(x):
                return -kde(x)[0]
            
            # 先粗网格搜索找到大致位置
            grid = np.linspace(x_min, x_max, 200)
            densities = kde(grid)
            rough_max = grid[np.argmax(densities)]
            
            # 精细优化（局部搜索）
            res = minimize_scalar(neg_density, bracket=(rough_max - (x_max-x_min)*0.1,
                                                        rough_max,
                                                        rough_max + (x_max-x_min)*0.1),
                                  method='brent')
            if res.success:
                return float(res.x)
            else:
                # 优化失败则返回粗网格最大值
                return float(rough_max)
        
        # 自动选择策略
        unique_ratio = len(np.unique(arr)) / n
        if mode_method == "auto":
            # 如果唯一值比例低（< 0.2）或者数据点少（< 1000），用离散方法
            # 否则用 KDE
            if unique_ratio < 0.2 or n < 1000:
                method_used = "discrete"
            else:
                method_used = "kde"
        else:
            method_used = mode_method
        
        # 计算众数
        if method_used == "discrete":
            return _mode_discrete(arr)
        elif method_used == "histogram":
            return _mode_histogram(arr)
        elif method_used == "kde":
            return _mode_kde(arr)
        else:
            # 未知方法，回退离散
            return _mode_discrete(arr)
        
    except Exception as e:
        print(f"DriskMode错误: {str(e)}")
        return float('nan')

@xl_func("xl_cell cell, int sim_num, var trunc: float", volatile=True)
@xl_arg("cell", "range")
def DriskMeanAbsDev(cell, sim_num: int = 1, trunc=None) -> float:
    """计算平均绝对离差（Mean Absolute Deviation） - 支持截断参数
    
    参数：
        cell: 输出单元格
        sim_num: 模拟ID（场景数）
        trunc: 可选截断参数，支持DriskTruncate或DriskTruncateP标记字符串
               支持单边值截断，如DriskTruncate(, max)或DriskTruncateP(low,)
    """
    try:
        # 获取模拟对象 - 现在sim_num直接对应场景ID
        sim = get_simulation(sim_num)
        if sim is None:
            # 尝试获取任何可用的模拟
            from simulation_manager import get_all_simulations
            all_sims = get_all_simulations()
            if all_sims:
                sim = next(iter(all_sims.values()))
            else:
                return float('nan')
        
        # 获取数据
        data = sim.get_output_data_by_range(cell)
        if data is None:
            return float('nan')
        
        # 过滤有效数据
        valid_data = _filter_valid_data(data)
        if len(valid_data) == 0:
            return float('nan')
        
        # 应用截断处理（如果提供了trunc参数）
        if trunc is not None and isinstance(trunc, str):
            processed_data = _apply_truncation(valid_data, trunc)
            if processed_data is None:
                return float('nan')  # 截断参数无效
            valid_data = processed_data
            if len(valid_data) == 0:
                return float('nan')
        
        arr = np.array(valid_data)
        return float(np.mean(np.abs(arr - np.mean(arr))))
    except Exception as e:
        print(f"DriskMeanAbsDev错误: {str(e)}")
        return float('nan')


@xl_func("xl_cell cell, float x_value, int sim_num, var trunc: float", volatile=True)
@xl_arg("cell", "range")
def DriskXtoP(cell, x_value: float, sim_num: int = 1, trunc=None) -> float:
    """返回给定x对应的累计概率（p） - p 返回范围[0,1] - 支持截断参数
    
    参数：
        cell: 输出单元格
        x_value: x值
        sim_num: 模拟ID（场景数）
        trunc: 可选截断参数，支持DriskTruncate或DriskTruncateP标记字符串
               支持单边值截断，如DriskTruncate(, max)或DriskTruncateP(low,)
    """
    try:
        # 获取模拟对象 - 现在sim_num直接对应场景ID
        sim = get_simulation(sim_num)
        if sim is None:
            # 尝试获取任何可用的模拟
            from simulation_manager import get_all_simulations
            all_sims = get_all_simulations()
            if all_sims:
                sim = next(iter(all_sims.values()))
            else:
                return float('nan')
        
        # 获取数据
        data = sim.get_output_data_by_range(cell)
        if data is None:
            return float('nan')
        
        # 过滤有效数据
        valid_data = _filter_valid_data(data)
        if len(valid_data) == 0:
            return float('nan')
        
        # 应用截断处理（如果提供了trunc参数）
        if trunc is not None and isinstance(trunc, str):
            processed_data = _apply_truncation(valid_data, trunc)
            if processed_data is None:
                return float('nan')  # 截断参数无效
            valid_data = processed_data
            if len(valid_data) == 0:
                return float('nan')
        
        # 计算不超过 x_value 的比例（含等于）
        count_le = np.sum(valid_data <= float(x_value))
        return float(count_le / len(valid_data))
    except Exception as e:
        print(f"DriskXtoP错误: {str(e)}")
        return float('nan')

# ==================== 新增统计函数 ====================

# 修改处1 -- 新增置信区间智能计算函数
# ==================== >>> 智能置信区间路由引擎 (修改范围开始) <<< ====================
def _smart_ci_engine(data, stat_type, confidence_level, lower_bound, p_val=0.5, is_sorted=False, engine_mode='fast'):
    """
    金融级多模式置信区间计算引擎
    参数 engine_mode: 'fast' (解析解), 'bootstrap' (标准重抽样), 'bca' (高精度偏差修正)
    """
    n = len(data)
    if n < 2: return float('nan')
    alpha = 1.0 - confidence_level

    # ========================================================
    # 强制保护区：分位数与极值（避免 BCa 奇异矩阵崩溃，强制使用高精度解析解/EVT）
    # ========================================================
    if stat_type == 'percentile':
        sorted_data = data if is_sorted else np.sort(data)
        try:
            from scipy.stats import beta
            L_prob = beta.ppf(alpha / 2, n * p_val, n - n * p_val + 1)
            U_prob = beta.ppf(1 - alpha / 2, n * p_val + 1, n - n * p_val)
            L_idx = max(0, min(n - 1, int(np.floor(L_prob * n))))
            U_idx = max(0, min(n - 1, int(np.ceil(U_prob * n))))
            return float(sorted_data[L_idx]) if lower_bound else float(sorted_data[U_idx])
        except ImportError:
            z_val = 1.96 + (confidence_level - 0.95) * 6.2
            k = int(np.round(n * p_val - z_val * np.sqrt(n * p_val * (1-p_val))))
            m = int(np.round(n * p_val + z_val * np.sqrt(n * p_val * (1-p_val))))
            return float(sorted_data[max(1, min(n, k))-1]) if lower_bound else float(sorted_data[max(1, min(n, m))-1])

    elif stat_type in ('min', 'max'):
        try:
            from scipy.stats import genpareto
            process_data = data if stat_type == 'max' else -data
            u = np.percentile(process_data, 95)
            tail_data = process_data[process_data > u] - u
            if len(tail_data) < 5 or np.std(tail_data) < 1e-12:
                return float(np.min(data)) if stat_type == 'min' else float(np.max(data))
                
            c, loc, scale = genpareto.fit(tail_data, floc=0)
            prob_target = 1 - alpha/2 if lower_bound else 1 - (1 - alpha/2)
            ratio = len(tail_data) / n
            p_exceed = 1 - prob_target
            
            if abs(c) < 1e-6:
                val = u - scale * np.log(p_exceed / ratio)
            else:
                val = u + (scale / c) * (((ratio / p_exceed)**c) - 1)
            return float(-val) if stat_type == 'min' else float(val)
        except Exception:
            return float(np.min(data)) if stat_type == 'min' else float(np.max(data))

    # ========================================================
    # 动态引擎区：统计矩 (Mean, Std, Skew, Kurt)
    # ========================================================
    if engine_mode in ('bootstrap', 'bca'):
        try:
            from scipy.stats import bootstrap
            func_map = {
                'mean': np.mean,
                'std': lambda x: np.std(x, ddof=1),
                'skew': lambda x: float(np.mean(((x - np.mean(x)) / np.std(x, ddof=1))**3)),
                'kurt': lambda x: float(np.mean(((x - np.mean(x)) / np.std(x, ddof=1))**4) + 3)
            }
            stat_func = func_map.get(stat_type)
            if stat_func:
                method = 'BCa' if engine_mode == 'bca' else 'percentile'
                res = bootstrap((data,), stat_func, confidence_level=confidence_level, method=method, n_resamples=1000)
                return float(res.confidence_interval.low) if lower_bound else float(res.confidence_interval.high)
        except Exception as e:
            print(f"[{engine_mode}] 计算 {stat_type} 失败，自动降级为 fast 模式: {e}")
            engine_mode = 'fast' # 平滑降级

    # Fast 极速解析引擎 (包含 Cornish-Fisher 修正)
    if engine_mode == 'fast':
        try:
            from scipy.stats import norm
            z_base = norm.ppf(1 - alpha / 2)
        except ImportError:
            z_base = 1.96 + (confidence_level - 0.95) * 6.2

        if stat_type == 'mean':
            mean_val = np.mean(data)
            se = np.std(data, ddof=1) / np.sqrt(n)
            try:
                from scipy.stats import skew, kurtosis
                s = skew(data, bias=False)
                k = kurtosis(data, fisher=True, bias=False)
                z_cf = z_base + (z_base**2 - 1)*s/6 + (z_base**3 - 3*z_base)*k/24 - (2*z_base**3 - 5*z_base)*(s**2)/36
                used_z = z_cf if np.isfinite(z_cf) else z_base
            except Exception:
                used_z = z_base
            return float(mean_val - used_z * se) if lower_bound else float(mean_val + used_z * se)

        elif stat_type == 'std':
            std_val = np.std(data, ddof=1)
            se = std_val / np.sqrt(2 * n)
            return float(std_val - z_base * se) if lower_bound else float(std_val + z_base * se)
            
        elif stat_type == 'skew':
            skew_val = float(np.mean(((data - np.mean(data)) / np.std(data, ddof=1))**3))
            se = np.sqrt(6.0 / n)
            return float(skew_val - z_base * se) if lower_bound else float(skew_val + z_base * se)
            
        elif stat_type == 'kurt':
            kurt_val = float(np.mean(((data - np.mean(data)) / np.std(data, ddof=1))**4) + 3)
            se = np.sqrt(24.0 / n)
            return float(kurt_val - z_base * se) if lower_bound else float(kurt_val + z_base * se)

    return float('nan')
# ==================== >>> 智能置信区间路由引擎 (修改范围结束) <<< ====================

# 修改处2 -- 同名函数替换
# ==================== >>> 修改后的 DriskCIMean (覆盖替换原函数) <<< ====================
@xl_func("xl_cell cell, float confidence_level, bool lower_bound, int sim_num, var trunc: float", volatile=True)
@xl_arg("cell", "range")
def DriskCIMean(cell, confidence_level: float, lower_bound: bool = True, sim_num: int = 1, trunc=None) -> float:
    """返回模拟分布均值的置信区间下限或上限 - 支持截断参数（已接入智能路由引擎）"""
    try:
        if not (0 < confidence_level < 1):
            return float('nan')
        
        sim = get_simulation(sim_num)
        if sim is None:
            from simulation_manager import get_all_simulations
            all_sims = get_all_simulations()
            if all_sims:
                sim = next(iter(all_sims.values()))
            else:
                return float('nan')
        
        data = sim.get_output_data_by_range(cell)
        if data is None:
            return float('nan')
        
        valid_data = _filter_valid_data(data)
        if len(valid_data) == 0:
            return float('nan')
        
        if trunc is not None and isinstance(trunc, str):
            processed_data = _apply_truncation(valid_data, trunc)
            if processed_data is None or len(processed_data) == 0:
                return float('nan')
            valid_data = processed_data
            
        # 核心修改：一键移交智能引擎处理
        return _smart_ci_engine(valid_data, 'mean', confidence_level, lower_bound)
        
    except Exception as e:
        print(f"DriskCIMean错误: {str(e)}")
        return float('nan')
# ==================== >>> 修改后的 DriskCIMean (修改范围结束) <<< ====================

@xl_func("xl_cell cell, int sim_num, var trunc: float", volatile=True)
@xl_arg("cell", "range")
def DriskCoeffOfVariation(cell, sim_num: int = 1, trunc=None) -> float:
    """
    返回模拟分布的变异系数 - 支持截断参数
    
    参数：
        cell: 要计算其模拟分布变异系数的单元格、输出或输入
        sim_num: 可选参数，用于在存在多次模拟时指定要计算其变异系数的具体模拟
        trunc: 可选截断参数，支持DriskTruncate或DriskTruncateP标记字符串
               支持单边值截断，如DriskTruncate(, max)或DriskTruncateP(low,)
    """
    try:
        # 获取模拟对象
        sim = get_simulation(sim_num)
        if sim is None:
            # 尝试获取任何可用的模拟
            from simulation_manager import get_all_simulations
            all_sims = get_all_simulations()
            if all_sims:
                sim = next(iter(all_sims.values()))
            else:
                return float('nan')
        
        # 获取数据
        data = sim.get_output_data_by_range(cell)
        if data is None:
            return float('nan')
        
        # 过滤有效数据
        valid_data = _filter_valid_data(data)
        if len(valid_data) == 0:
            return float('nan')
        
        # 应用截断处理（如果提供了trunc参数）
        if trunc is not None and isinstance(trunc, str):
            processed_data = _apply_truncation(valid_data, trunc)
            if processed_data is None:
                return float('nan')  # 截断参数无效
            valid_data = processed_data
            if len(valid_data) == 0:
                return float('nan')
        
        # 计算变异系数
        mean_val = np.mean(valid_data)
        if mean_val == 0:
            return float('nan')
        
        std_val = np.std(valid_data, ddof=1)
        cv = std_val / mean_val
        
        return float(cv)
        
    except Exception as e:
        print(f"DriskCoeffOfVariation错误: {str(e)}")
        return float('nan')

@xl_func("xl_cell cell, bool lower_data, int sim_num, var trunc: float", volatile=True)
@xl_arg("cell", "range")
def DriskSemiStdDev(cell, lower_data: bool = True, sim_num: int = 1, trunc=None) -> float:
    """
    返回模拟分布的半标准差 - 支持截断参数
    
    参数：
        cell: 要计算其模拟分布半标准差的单元格、输出或输入分布函数
        lower_data: TRUE表示使用小于或等于均值的数据（默认），FALSE表示使用大于或等于均值的数据
        sim_num: 可选参数，用于在存在多次模拟时指定要计算其半标准差的具体模拟
        trunc: 可选截断参数，支持DriskTruncate或DriskTruncateP标记字符串
               支持单边值截断，如DriskTruncate(, max)或DriskTruncateP(low,)
    """
    try:
        # 获取模拟对象
        sim = get_simulation(sim_num)
        if sim is None:
            # 尝试获取任何可用的模拟
            from simulation_manager import get_all_simulations
            all_sims = get_all_simulations()
            if all_sims:
                sim = next(iter(all_sims.values()))
            else:
                return float('nan')
        
        # 获取数据
        data = sim.get_output_data_by_range(cell)
        if data is None:
            return float('nan')
        
        # 过滤有效数据
        valid_data = _filter_valid_data(data)
        if len(valid_data) == 0:
            return float('nan')
        
        # 应用截断处理（如果提供了trunc参数）
        if trunc is not None and isinstance(trunc, str):
            processed_data = _apply_truncation(valid_data, trunc)
            if processed_data is None:
                return float('nan')  # 截断参数无效
            valid_data = processed_data
            if len(valid_data) == 0:
                return float('nan')
        
        # 计算均值
        mean_val = np.mean(valid_data)
        
        # 根据lower_data参数选择数据
        if lower_data:
            # 小于或等于均值的数据
            selected_data = [x for x in valid_data if x <= mean_val]
        else:
            # 大于或等于均值的数据
            selected_data = [x for x in valid_data if x >= mean_val]
        
        if len(selected_data) == 0:
            return float('nan')
        
        # 计算半标准差
        # 公式：sqrt(Σ(x - mean)² / (n-1))，但只计算选定方向的数据
        deviations = np.array(selected_data) - mean_val
        semi_variance = np.sum(deviations ** 2) / (len(selected_data) - 1) if len(selected_data) > 1 else 0
        semi_std_dev = np.sqrt(semi_variance)
        
        return float(semi_std_dev)
        
    except Exception as e:
        print(f"DriskSemiStdDev错误: {str(e)}")
        return float('nan')

@xl_func("xl_cell cell, bool lower_data, int sim_num, var trunc: float", volatile=True)
@xl_arg("cell", "range")
def DriskSemiVariance(cell, lower_data: bool = True, sim_num: int = 1, trunc=None) -> float:
    """
    返回模拟分布的半方差 - 支持截断参数
    
    参数：
        cell: 要计算其模拟分布半方差的单元格、输出或输入分布函数
        lower_data: TRUE表示使用小于或等于均值的数据（默认），FALSE表示使用大于或等于均值的数据
        sim_num: 可选参数，用于在存在多次模拟时指定要计算其半方差的具体模拟
        trunc: 可选截断参数，支持DriskTruncate或DriskTruncateP标记字符串
               支持单边值截断，如DriskTruncate(, max)或DriskTruncateP(low,)
    """
    try:
        # 获取模拟对象
        sim = get_simulation(sim_num)
        if sim is None:
            # 尝试获取任何可用的模拟
            from simulation_manager import get_all_simulations
            all_sims = get_all_simulations()
            if all_sims:
                sim = next(iter(all_sims.values()))
            else:
                return float('nan')
        
        # 获取数据
        data = sim.get_output_data_by_range(cell)
        if data is None:
            return float('nan')
        
        # 过滤有效数据
        valid_data = _filter_valid_data(data)
        if len(valid_data) == 0:
            return float('nan')
        
        # 应用截断处理（如果提供了trunc参数）
        if trunc is not None and isinstance(trunc, str):
            processed_data = _apply_truncation(valid_data, trunc)
            if processed_data is None:
                return float('nan')  # 截断参数无效
            valid_data = processed_data
            if len(valid_data) == 0:
                return float('nan')
        
        # 计算均值
        mean_val = np.mean(valid_data)
        
        # 根据lower_data参数选择数据
        if lower_data:
            # 小于或等于均值的数据
            selected_data = [x for x in valid_data if x <= mean_val]
        else:
            # 大于或等于均值的数据
            selected_data = [x for x in valid_data if x >= mean_val]
        
        if len(selected_data) == 0:
            return float('nan')
        
        # 计算半方差
        # 公式：Σ(x - mean)² / (n-1)，但只计算选定方向的数据
        deviations = np.array(selected_data) - mean_val
        semi_variance = np.sum(deviations ** 2) / (len(selected_data) - 1) if len(selected_data) > 1 else 0
        
        return float(semi_variance)
        
    except Exception as e:
        print(f"DriskSemiVariance错误: {str(e)}")
        return float('nan')

@xl_func("xl_cell cell, int sim_num, var trunc: float", volatile=True)
@xl_arg("cell", "range")
def DriskStdErrOfMean(cell, sim_num: int = 1, trunc=None) -> float:
    """
    返回模拟分布均值的标准误 - 支持截断参数
    
    参数：
        cell: 要计算其模拟分布均值标准误的单元格、输出或输入分布函数
        sim_num: 可选参数，用于在存在多次模拟时指定要计算其均值标准误的具体模拟
        trunc: 可选截断参数，支持DriskTruncate或DriskTruncateP标记字符串
               支持单边值截断，如DriskTruncate(, max)或DriskTruncateP(low,)
    """
    try:
        # 获取模拟对象
        sim = get_simulation(sim_num)
        if sim is None:
            # 尝试获取任何可用的模拟
            from simulation_manager import get_all_simulations
            all_sims = get_all_simulations()
            if all_sims:
                sim = next(iter(all_sims.values()))
            else:
                return float('nan')
        
        # 获取数据
        data = sim.get_output_data_by_range(cell)
        if data is None:
            return float('nan')
        
        # 过滤有效数据
        valid_data = _filter_valid_data(data)
        n = len(valid_data)
        if n == 0:
            return float('nan')
        
        # 应用截断处理（如果提供了trunc参数）
        if trunc is not None and isinstance(trunc, str):
            processed_data = _apply_truncation(valid_data, trunc)
            if processed_data is None:
                return float('nan')  # 截断参数无效
            valid_data = processed_data
            n = len(valid_data)
            if n == 0:
                return float('nan')
        
        # 计算均值的标准误
        std_val = np.std(valid_data, ddof=1)
        std_err = std_val / np.sqrt(n)
        
        return float(std_err)
        
    except Exception as e:
        print(f"DriskStdErrOfMean错误: {str(e)}")
        return float('nan')

# 修改处3 -- 同名函数替换
# ==================== >>> 修改后的 DriskCIPercentile (覆盖替换原函数) <<< ====================
@xl_func("xl_cell cell, float confidence_level, bool lower_bound, int sim_num, var trunc: float", volatile=True)
@xl_arg("cell", "range")
def DriskCIPercentile(cell, confidence_level: float, lower_bound: bool = True, sim_num: int = 1, trunc=None) -> float:
    """返回模拟分布中位数的分位数置信区间 - 支持截断参数（已接入智能路由引擎）"""
    try:
        if not (0 < confidence_level < 1):
            return float('nan')
        
        sim = get_simulation(sim_num)
        if sim is None:
            from simulation_manager import get_all_simulations
            all_sims = get_all_simulations()
            if all_sims:
                sim = next(iter(all_sims.values()))
            else:
                return float('nan')
        
        data = sim.get_output_data_by_range(cell)
        if data is None:
            return float('nan')
        
        valid_data = _filter_valid_data(data)
        if len(valid_data) == 0:
            return float('nan')
        
        if trunc is not None and isinstance(trunc, str):
            processed_data = _apply_truncation(valid_data, trunc)
            if processed_data is None or len(processed_data) == 0:
                return float('nan')
            valid_data = processed_data
        
        # 核心修改：指定为 'percentile' 类型，引擎将自动使用 Binomial 极速算法
        # 原逻辑硬编码了求中位数的置信区间，所以 p_val=0.5
        return _smart_ci_engine(valid_data, 'percentile', confidence_level, lower_bound, p_val=0.5)
        
    except Exception as e:
        print(f"DriskCIPercentile错误: {str(e)}")
        return float('nan')
# ==================== >>> 修改后的 DriskCIPercentile (修改范围结束) <<< ====================
