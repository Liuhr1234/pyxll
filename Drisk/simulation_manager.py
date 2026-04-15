# simulation_manager.py
"""模拟管理器模块 - 支持跨sheet版本（性能优化）"""

import numpy as np
import pandas as pd
import threading
from collections import OrderedDict
from typing import Dict, List, Optional, Any, Tuple
from constants import DEFAULT_ITERATIONS, SAMPLING_MC

# 全局模拟缓存
_SIMULATION_CACHE = OrderedDict()

# 特殊错误标记
ERROR_MARKER = "#ERROR!"

class SimulationResult:
    """模拟结果对象 - 支持跨sheet"""
    def __init__(self, sim_id: int, n_iterations: int = DEFAULT_ITERATIONS, sampling_method: str = SAMPLING_MC):
        self.sim_id = sim_id
        self.n_iterations = n_iterations
        self.sampling_method = sampling_method
        
        # Input存储：键为"工作表名!单元格地址_序号"，值为结果数组
        self.input_cache = {}
        
        # Output存储：键为"工作表名!单元格地址"，值为结果数组
        self.output_cache = {}
        
        # Output属性存储：键为"工作表名!单元格地址"，值为属性字典
        self.output_attributes = {}
        
        # Input属性存储：键为"工作表名!单元格地址_序号"，值为属性字典
        self.input_attributes = {}
        
        self.workbook_name = None  # 模拟所在的工作簿名称
        self.timestamp = pd.Timestamp.now()
        self.end_time = None  # 添加结束时间
        self.name = f"Sim_{sim_id}_{sampling_method}"
        
        # 存储原始信息
        self.distribution_cells = {}  # {sheet!cell_addr: [dist_func1, dist_func2, ...]}
        self.simtable_cells = {}  # 新增：{sheet!cell_addr: [simtable_func1, simtable_func2, ...]}
        self.output_cells = {}  # {sheet!cell_addr: output_info}
        self.all_input_keys = []  # [sheet!cell_addr_index1, sheet!cell_addr_index2, ...]
        
        # 场景信息 - 现在一个模拟只包含一个场景
        self.scenario_count = 1  # 默认1个场景
        self.scenario_index = 0  # 当前场景索引（0-based）
        self.is_scenario_simulation = False  # 是否为场景模拟
        
        # 缓存
        self._input_statistics_cache = {}
        self._output_statistics_cache = {}
        self._input_sorted_cache = {}
        self._output_sorted_cache = {}
        
        self._lock = threading.RLock()
        
        # 统计计算标志
        self._input_stats_computed = False
        self._output_stats_computed = False
        
        # 性能统计
        self.performance_stats = {
            'total_iterations_completed': 0,
            'simulation_time': 0,
            'data_points': 0
        }
        
    def set_scenario_info(self, scenario_count: int, scenario_index: int):
        """设置场景信息"""
        self.scenario_count = scenario_count
        self.scenario_index = scenario_index
        self.is_scenario_simulation = scenario_count > 1
        if self.is_scenario_simulation:
            self.name = f"Sim_{self.sim_id}_{self.sampling_method}_场景{scenario_index+1}"
        
    def add_input_result(self, input_key: str, data: np.ndarray, sheet_name: str, attributes: Dict = None):
        """添加Input结果 - 修复：确保正确保存多个input"""
        if len(data) > 0:
            # 规范化key - 确保包含序号
            # input_key 格式应为 "单元格地址_序号" 或 "工作表名!单元格地址_序号"
            if '_' not in input_key:
                # 如果没有序号，添加默认序号1
                input_key = f"{input_key}_1"
            
            input_key = input_key.replace('$', '').upper()
            
            # 创建缓存键：工作表名!input_key
            cache_key = f"{sheet_name}!{input_key}"
            
            with self._lock:
                # 存储数据
                self.input_cache[cache_key] = data
                
                # 存储属性
                if attributes:
                    self.input_attributes[cache_key] = attributes
                elif cache_key not in self.input_attributes:
                    self.input_attributes[cache_key] = {}
                
                # 预排序数据 - 需要过滤掉错误值和NaN
                valid_numeric_data = []
                for v in data:
                    # 跳过 None
                    if v is None:
                        continue
                    # 跳过字符串错误标记
                    if isinstance(v, str) and v == ERROR_MARKER:
                        continue
                    # 跳过NaN值
                    try:
                        if isinstance(v, (int, float, np.number)):
                            if np.isnan(v):
                                continue
                        elif isinstance(v, str):
                            # 尝试转换字符串
                            try:
                                v = float(v)
                            except:
                                continue
                    except (TypeError, ValueError):
                        pass
                    valid_numeric_data.append(float(v))
                
                if len(valid_numeric_data) > 0:
                    sorted_data = np.sort(np.array(valid_numeric_data))
                    self._input_sorted_cache[cache_key] = sorted_data
                
                # 标记统计需要重新计算
                self._input_stats_computed = False
                self._input_statistics_cache = {}
                
                # 更新性能统计 - 计算实际的有效数值
                self.performance_stats['data_points'] += len(valid_numeric_data)
                
                # 添加到all_input_keys（如果不存在）
                if cache_key not in self.all_input_keys:
                    self.all_input_keys.append(cache_key)
                
                print(f"成功保存Input数据: {cache_key}, 数据长度: {len(data)}")
                return True
        else:
            print(f"警告: Input数据为空 {input_key}")
        return False
    
    def add_output_result(self, cell_address: str, data: np.ndarray, sheet_name: str, 
                         output_info: Dict = None):
        """添加Output结果"""
        if len(data) > 0:
            # 规范化地址
            cell_address = cell_address.replace('$', '').upper()
            
            # 创建缓存键：工作表名!单元格地址
            cache_key = f"{sheet_name}!{cell_address}"
            
            with self._lock:
                # 存储数据
                self.output_cache[cache_key] = data
                
                # 存储属性
                if output_info:
                    self.output_attributes[cache_key] = output_info
                elif cache_key not in self.output_attributes:
                    self.output_attributes[cache_key] = {}
                
                # 预排序数据 - 需要过滤掉错误值和NaN
                valid_numeric_data = []
                for v in data:
                    # 跳过字符串错误标记
                    if isinstance(v, str) and v == ERROR_MARKER:
                        continue
                    # 跳过NaN值
                    try:
                        if isinstance(v, (int, float, np.number)):
                            if np.isnan(v):
                                continue
                        elif isinstance(v, str):
                            # 尝试转换字符串
                            try:
                                v = float(v)
                            except:
                                continue
                    except (TypeError, ValueError):
                        continue
                    valid_numeric_data.append(float(v))
                
                if len(valid_numeric_data) > 0:
                    sorted_data = np.sort(np.array(valid_numeric_data))
                    self._output_sorted_cache[cache_key] = sorted_data
                
                # 标记统计需要重新计算
                self._output_stats_computed = False
                self._output_statistics_cache = {}
                
                # 更新性能统计 - 计数所有迭代
                self.performance_stats['total_iterations_completed'] = max(
                    self.performance_stats['total_iterations_completed'],
                    len(data)
                )
                self.performance_stats['data_points'] += len(valid_numeric_data)
                
                print(f"成功保存Output数据: {cache_key}, 数据长度: {len(data)}")
                return True
        else:
            print(f"警告: Output数据为空 {cell_address}")
        return False
    
    def get_input_data(self, input_key: str) -> Optional[np.ndarray]:
        """获取Input数据（支持跨sheet）"""
        try:
            # 如果input_key已经包含工作表名
            if '!' in input_key:
                # 直接尝试缓存键
                if input_key in self.input_cache:
                    return self.input_cache[input_key]
                
                # 规范化
                cache_key = input_key.replace('$', '').upper()
                if cache_key in self.input_cache:
                    return self.input_cache[cache_key]
            
            # 尝试在所有工作表中查找
            for key in self.input_cache:
                if key.endswith(f"!{input_key}"):
                    return self.input_cache[key]
            
            # 尝试不带下标的查找
            for key in self.input_cache:
                base_key = key.split('_')[0] if '_' in key else key
                if base_key.endswith(f"!{input_key}"):
                    return self.input_cache[key]
            
            return None
        except Exception as e:
            print(f"获取Input数据失败 {input_key}: {e}")
            return None
    
    def get_output_data(self, cell_address: str) -> Optional[np.ndarray]:
        """获取Output数据（支持跨sheet）"""
        try:
            # 如果cell_address已经包含工作表名
            if '!' in cell_address:
                # 直接尝试缓存键
                if cell_address in self.output_cache:
                    return self.output_cache[cell_address]
                
                # 规范化
                cache_key = cell_address.replace('$', '').upper()
                if cache_key in self.output_cache:
                    return self.output_cache[cache_key]
            
            # 尝试在所有工作表中查找 - 多种匹配策略
            cell_only = cell_address.split('!')[-1] if '!' in cell_address else cell_address
            
            # 策略1：精确匹配
            for key in self.output_cache:
                if key.endswith(f"!{cell_only}"):
                    return self.output_cache[key]
            
            # 策略2：模糊匹配（忽略大小写和$符号）
            cell_only_clean = cell_only.replace('$', '').upper()
            for key in self.output_cache:
                key_cell_only = key.split('!')[-1].replace('$', '').upper() if '!' in key else key.replace('$', '').upper()
                if key_cell_only == cell_only_clean:
                    return self.output_cache[key]
            
            return None
        except Exception as e:
            print(f"获取Output数据失败 {cell_address}: {e}")
            return None
    
    def get_output_attributes(self, cell_address: str) -> Dict:
        """获取Output属性（支持跨sheet）"""
        try:
            # 如果cell_address已经包含工作表名
            if '!' in cell_address:
                # 直接尝试缓存键
                if cell_address in self.output_attributes:
                    return self.output_attributes[cell_address]
                
                # 规范化
                cache_key = cell_address.replace('$', '').upper()
                if cache_key in self.output_attributes:
                    return self.output_attributes[cache_key]
            
            # 尝试在所有工作表中查找
            for key in self.output_attributes:
                if key.endswith(f"!{cell_address}"):
                    return self.output_attributes[key]
            
            return {}
        except Exception as e:
            print(f"获取Output属性失败 {cell_address}: {e}")
            return {}
    
    def get_input_attributes(self, input_key: str) -> Dict:
        """获取Input属性（支持跨sheet）"""
        try:
            # 如果input_key已经包含工作表名
            if '!' in input_key:
                # 直接尝试缓存键
                if input_key in self.input_attributes:
                    return self.input_attributes[input_key]
                
                # 规范化
                cache_key = input_key.replace('$', '').upper()
                if cache_key in self.input_attributes:
                    return self.input_attributes[cache_key]
            
            # 尝试在所有工作表中查找
            for key in self.input_attributes:
                if key.endswith(f"!{input_key}"):
                    return self.input_attributes[key]
            
            return {}
        except Exception as e:
            print(f"获取Input属性失败 {input_key}: {e}")
            return {}
    
    def is_output_cell(self, cell_range) -> bool:
        """检查单元格是否是Output单元格（支持跨sheet）"""
        try:
            sheet_name = cell_range.Worksheet.Name
            address = cell_range.Address.replace('$', '')
            cache_key = f"{sheet_name}!{address}"
            
            return cache_key in self.output_cache
        except Exception as e:
            print(f"检查Output单元格失败: {e}")
            return False
    
    # simulation_manager.py - 在 SimulationResult 类中添加方法

    def get_all_inputs_with_type(self) -> Dict[str, Dict[str, Any]]:
        """获取所有输入及其类型信息"""
        inputs = {}
        
        for cache_key in self.input_cache:
            attrs = self.input_attributes.get(cache_key, {})
            
            # 确定输入类型
            input_type = 'distribution'
            if attrs.get('is_makeinput'):
                input_type = 'makeinput'
            elif attrs.get('is_nested'):
                input_type = 'nested_distribution'
            
            inputs[cache_key] = {
                'type': input_type,
                'attributes': attrs,
                'data': self.input_cache[cache_key]
            }
        
        return inputs


    def compute_input_statistics(self):
        """预计算所有Input的统计量"""
        with self._lock:
            if self._input_stats_computed:
                return
            
            self._input_statistics_cache = {}
            
            for cache_key, data in self.input_cache.items():
                stats = self._compute_statistics(data)
                if stats:  # 只存储非空统计
                    self._input_statistics_cache[cache_key] = stats
            
            self._input_stats_computed = True
    
    def compute_output_statistics(self):
        """预计算所有Output的统计量"""
        with self._lock:
            if self._output_stats_computed:
                return
            
            self._output_statistics_cache = {}
            
            for cache_key, data in self.output_cache.items():
                stats = self._compute_statistics(data)
                if stats:  # 只存储非空统计
                    self._output_statistics_cache[cache_key] = stats
            
            self._output_stats_computed = True
    
    def _compute_statistics(self, data: np.ndarray) -> Dict[str, float]:
        """计算统计量（优化版 - 过滤#ERROR值）"""
        # 过滤掉#ERROR字符串和NaN值
        valid_data = []
        for val in data:
            # 跳过字符串错误标记
            if isinstance(val, str):
                if val == ERROR_MARKER:
                    continue  # 跳过#ERROR
                else:
                    # 尝试转换其他字符串为数值
                    try:
                        float_val = float(val)
                        if np.isnan(float_val):
                            continue
                        valid_data.append(float_val)
                    except:
                        continue
            else:
                # 检查数值型的NaN
                try:
                    if np.isnan(val):
                        continue
                    valid_data.append(float(val))
                except (TypeError, ValueError):
                    continue
        
        if len(valid_data) == 0:
            return {}  # 返回空字典，表示没有有效数据
        
        valid_data_array = np.array(valid_data)
        stats = {}
        
        try:
            # 基础统计量
            stats['mean'] = float(np.mean(valid_data_array))
            stats['median'] = float(np.median(valid_data_array))
            stats['std'] = float(np.std(valid_data_array))
            stats['min'] = float(np.min(valid_data_array))
            stats['max'] = float(np.max(valid_data_array))
            stats['count'] = len(valid_data_array)
            
            # 分位数（使用快速计算方法）
            if len(valid_data_array) > 0:
                percentiles = [5, 25, 50, 75, 95]
                for p in percentiles:
                    try:
                        stats[f'p{p}'] = float(np.percentile(valid_data_array, p))
                    except:
                        stats[f'p{p}'] = 0.0
        except Exception as e:
            print(f"计算统计量失败: {e}")
            # 如果计算失败，返回空统计
            return {}
        
        return stats
    
    def get_input_statistics(self, input_key: str) -> Dict[str, float]:
        """获取Input统计量（支持跨sheet）"""
        if not self._input_stats_computed:
            self.compute_input_statistics()
        
        try:
            # 如果input_key已经包含工作表名
            if '!' in input_key:
                cache_key = input_key.replace('$', '').upper()
                if cache_key in self._input_statistics_cache:
                    return self._input_statistics_cache[cache_key]
            
            # 尝试在所有工作表中查找
            for key in self._input_statistics_cache:
                if key.endswith(f"!{input_key}"):
                    return self._input_statistics_cache[key]
        except Exception as e:
            print(f"获取Input统计量失败 {input_key}: {e}")
            pass
        
        return {}
    
    def get_output_statistics_by_range(self, cell_range) -> Dict[str, float]:
        """获取Output统计量"""
        if not self._output_stats_computed:
            self.compute_output_statistics()
        
        # 创建缓存键
        try:
            sheet_name = cell_range.Worksheet.Name
            address = cell_range.Address.replace('$', '')
            cache_key = f"{sheet_name}!{address}"
            
            # 从缓存获取
            if cache_key in self._output_statistics_cache:
                return self._output_statistics_cache[cache_key]
        except Exception as e:
            print(f"获取Output统计量失败: {e}")
            pass
        
        # 如果没有缓存，计算单个单元格
        data = self.get_output_data_by_range(cell_range)
        if data is None:
            return {}
        
        return self._compute_statistics(data)
    
    def get_output_data_by_range(self, cell_range) -> Optional[np.ndarray]:
        """通过Range对象获取Output数据"""
        try:
            # 从Range对象提取工作表名和地址
            sheet_name = cell_range.Worksheet.Name
            address = cell_range.Address.replace('$', '')
            
            # 创建缓存键
            cache_key = f"{sheet_name}!{address}"
            
            # 1. 直接尝试缓存键
            if cache_key in self.output_cache:
                return self.output_cache[cache_key]
            
            # 2. 尝试模糊匹配 - 多种策略
            short_key = address.upper()
            
            # 策略1：精确匹配
            for key in self.output_cache:
                if key.endswith(f"!{short_key}"):
                    return self.output_cache[key]
            
            # 策略2：只匹配单元格地址（忽略工作表名）
            for key in self.output_cache:
                cell_only = key.split('!')[-1] if '!' in key else key
                if cell_only == short_key:
                    return self.output_cache[key]
            
            # 策略3：清理$符号后匹配
            short_key_clean = short_key.replace('$', '')
            for key in self.output_cache:
                key_clean = key.replace('$', '')
                cell_only = key_clean.split('!')[-1] if '!' in key_clean else key_clean
                if cell_only == short_key_clean:
                    return self.output_cache[key]
            
            return None
            
        except Exception as e:
            print(f"通过Range获取Output数据失败: {e}")
            return None
    
    def get_sorted_output_data_by_range(self, cell_range) -> Optional[np.ndarray]:
        """获取排序后的Output数据"""
        try:
            sheet_name = cell_range.Worksheet.Name
            address = cell_range.Address.replace('$', '')
            cache_key = f"{sheet_name}!{address}"
            
            # 尝试从缓存获取
            if cache_key in self._output_sorted_cache:
                return self._output_sorted_cache[cache_key]
            
            # 获取数据
            data = self.get_output_data_by_range(cell_range)
            if data is None or len(data) == 0:
                return None
            
            # 排序并缓存 - 需要过滤掉错误值和NaN
            valid_numeric_data = []
            for v in data:
                # 跳过字符串错误标记
                if isinstance(v, str) and v == ERROR_MARKER:
                    continue
                # 尝试转换为数值
                try:
                    if isinstance(v, (int, float, np.number)):
                        if np.isnan(v):
                            continue
                        valid_numeric_data.append(float(v))
                    elif isinstance(v, str):
                        # 尝试转换其他字符串
                        try:
                            float_val = float(v)
                            if np.isnan(float_val):
                                continue
                            valid_numeric_data.append(float_val)
                        except:
                            continue
                except (TypeError, ValueError):
                    continue
            
            if len(valid_numeric_data) == 0:
                return None
            
            sorted_data = np.sort(np.array(valid_numeric_data))
            self._output_sorted_cache[cache_key] = sorted_data
            
            return sorted_data
        except Exception as e:
            print(f"获取排序后Output数据失败: {e}")
            return None
    
    def force_recompute_statistics(self):
        """强制重新计算所有统计量"""
        with self._lock:
            self._input_stats_computed = False
            self._output_stats_computed = False
            self._input_statistics_cache = {}
            self._output_statistics_cache = {}
            self.compute_input_statistics()
            self.compute_output_statistics()
    
    def get_performance_stats(self) -> Dict:
        """获取性能统计"""
        return self.performance_stats.copy()
    
    def get_all_input_data(self) -> Dict[str, np.ndarray]:
        """获取所有Input数据"""
        return self.input_cache.copy()
    
    def get_all_output_data(self) -> Dict[str, np.ndarray]:
        """获取所有Output数据"""
        return self.output_cache.copy()
    
    def set_end_time(self):
        """设置模拟结束时间"""
        self.end_time = pd.Timestamp.now()
    
    def get_duration(self) -> float:
        """获取模拟持续时间（秒）"""
        if self.end_time and self.timestamp:
            return (self.end_time - self.timestamp).total_seconds()
        return 0.0
    
    def get_scenario_info(self) -> Dict:
        """获取场景信息"""
        duration = self.get_duration()
        return {
            'scenario_count': self.scenario_count,
            'scenario_index': self.scenario_index,
            'is_scenario_simulation': self.is_scenario_simulation,
            'name': self.name,
            'duration': duration  # 添加持续时间
        }

# 全局模拟缓存管理器
_current_sim_id = 1
_sim_lock = threading.RLock()

def get_simulation(sim_num: int = 1) -> Optional[SimulationResult]:
    """获取模拟对象"""
    return _SIMULATION_CACHE.get(sim_num)

def create_simulation(n_iterations: int = DEFAULT_ITERATIONS, sampling_method: str = SAMPLING_MC) -> int:
    """创建新模拟"""
    global _current_sim_id
    with _sim_lock:
        sim_id = _current_sim_id
        sim = SimulationResult(sim_id, n_iterations, sampling_method)
        _SIMULATION_CACHE[sim_id] = sim
        _current_sim_id += 1
        print(f"创建新模拟: ID={sim_id}, 迭代次数={n_iterations}, 抽样方法={sampling_method}")
        return sim_id

def clear_simulations():
    """清除所有模拟"""
    global _SIMULATION_CACHE, _current_sim_id
    with _sim_lock:
        _SIMULATION_CACHE.clear()
        _current_sim_id = 1
        print("已清除所有模拟数据")

def get_all_simulations():
    """获取所有模拟"""
    return _SIMULATION_CACHE

def get_current_sim_id():
    """获取当前模拟ID"""
    return _current_sim_id