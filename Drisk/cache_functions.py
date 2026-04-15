# cache_functions.py - 简化优化版，解决嵌套函数问题并提高速度
"""缓存模拟函数模块 - DriskCacheMC相关算法"""

import tkinter as tk
from tkinter import ttk
import re
import threading
import traceback
import time
import numpy as np
from typing import Dict, List, Tuple, Any, Optional, Set
from pyxll import xl_macro, xl_func, xlcAlert, xl_app
from formula_parser import (
    extract_nested_distributions_advanced,
    parse_args_with_nested_functions,
    extract_dist_params_and_markers,
    parse_formula_references,
    extract_input_attributes,
    extract_makeinput_attributes  # 新增导入
)

# 从com_fixer导入_safe_excel_app
try:
    from com_fixer import _safe_excel_app
except ImportError:
    def _safe_excel_app():
        from pyxll import xl_app
        return xl_app()

# 导入日志模块用于控制输出
import logging
import sys

# 创建logger，但避免重复添加handler
logger = logging.getLogger(__name__)
if not logger.handlers:
    handler = logging.StreamHandler(sys.stdout)
    handler.setLevel(logging.WARNING)  # 降低默认日志级别
    formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
    handler.setFormatter(formatter)
    logger.addHandler(handler)
    logger.setLevel(logging.WARNING)

# ==================== 全局模拟迭代计数器（简化版） ====================

class GlobalIterationCounter:
    """全局迭代计数器 - 简化版：支持两层嵌套"""
    
    _instance = None
    _lock = threading.RLock()
    
    def __new__(cls):
        with cls._lock:
            if cls._instance is None:
                cls._instance = super().__new__(cls)
                cls._instance._init_counter()
            return cls._instance
    
    def _init_counter(self):
        """初始化计数器"""
        self.current_iteration = 0
        self.total_iterations = 0
        self.simulation_running = False
        self.simulation_id = 0
        self.distribution_registry = {}  # stable_key -> 分布信息
        self.random_cache = {}  # stable_key -> np.array(长度为total_iterations)
        self.call_tracker = {}  # 跟踪每个稳定键在每个迭代中的调用次数
        self.nested_functions = set()  # 嵌套函数的稳定键
        self.iteration_lock = threading.RLock()  # 迭代锁
        self.last_log_time = time.time()  # 最后日志时间
        self.log_interval = 2.0  # 日志输出间隔（秒）
        self.output_cache = {}  # 缓存输出单元格的值 {cell_addr: np.array}
        self.nested_parent_map = {}  # 嵌套函数父子关系映射 {child_key: parent_key}
        self.nested_children_map = {}  # 父函数的子函数映射 {parent_key: [child_key1, child_key2]}
        self.cell_value_cache = {}  # 缓存所有单元格的值 {cell_addr: np.array}
        self.cell_value_locks = {}  # 单元格值锁，确保线程安全
        self.nested_depth_map = {}  # 函数深度映射 {func_key: depth}
        self.debug_log = []  # 调试日志
        self.last_debug_print = 0  # 最后调试打印时间
        self.simtable_values = {}  # 存储Simtable值 {cell_addr: [values]}
        self.makeinput_cache = {}  # 存储MakeInput缓存 {cell_addr: np.array}
        self.scenario_count = 1  # 场景数
        self.scenario_index = 0  # 当前场景索引
        self.distribution_attributes = {}  # 存储分布函数的属性信息 {stable_key: attributes}
        self.makeinput_attributes = {}  # 新增：存储MakeInput的属性信息 {cell_addr: attributes}
        self.cell_formula_cache = {}  # 缓存单元格公式 {cell_addr: formula}
        self.makeinput_processed = set()  # 已处理的MakeInput单元格
        
    def log_debug(self, message: str):
        """记录调试信息"""
        current_time = time.time()
        timestamp = time.strftime("%H:%M:%S", time.localtime(current_time))
        debug_msg = f"[{timestamp}] {message}"
        self.debug_log.append(debug_msg)
        
        # 控制输出频率，避免日志过多
        if current_time - self.last_debug_print > 2.0 or len(self.debug_log) % 20 == 0:
            logger.info(debug_msg)
            self.last_debug_print = current_time
    
    def initialize_simulation(self, total_iterations: int, simulation_id: int = 0, 
                            scenario_count: int = 1, scenario_index: int = 0):
        """初始化模拟"""
        with self._lock:
            self.current_iteration = 0
            self.total_iterations = total_iterations
            self.simulation_running = True
            self.simulation_id = simulation_id
            self.scenario_count = scenario_count
            self.scenario_index = scenario_index
            self.random_cache.clear()
            self.distribution_registry.clear()
            self.call_tracker.clear()
            self.nested_functions.clear()
            self.output_cache.clear()
            self.nested_parent_map.clear()
            self.nested_children_map.clear()
            self.cell_value_cache.clear()
            self.cell_value_locks.clear()
            self.nested_depth_map.clear()
            self.last_log_time = time.time()
            self.debug_log.clear()
            self.last_debug_print = 0
            self.simtable_values.clear()
            self.makeinput_cache.clear()
            self.distribution_attributes.clear()
            self.makeinput_attributes.clear()  # 新增：清除MakeInput属性
            self.cell_formula_cache.clear()    # 清除公式缓存
            self.makeinput_processed.clear()   # 清除处理标记
            self.log_debug(f"全局迭代计数器初始化: 总迭代次数={total_iterations}, 模拟ID={simulation_id}, 场景数={scenario_count}, 当前场景索引={scenario_index}")
    
    def register_distribution(self, stable_key: str, func_name: str, params: List[float], 
                            markers: Dict[str, Any], is_nested: bool = False, 
                            parent_key: Optional[str] = None, depth: int = 0,
                            args_text: str = None, attributes: Dict[str, Any] = None):
        """注册分布函数 - 仅支持两层嵌套"""
        with self._lock:
            if stable_key in self.distribution_registry:
                self.log_debug(f"稳定键已注册，跳过: {stable_key}")
                return
            
            # 检查嵌套深度，超过两层报错
            if depth > 2:
                error_msg = f"错误: 函数 {func_name} 嵌套深度为 {depth}，超过两层嵌套限制！"
                self.log_debug(error_msg)
                raise ValueError(error_msg)
            
            # 为函数创建独立的缓存数组
            cache = np.full(self.total_iterations, np.nan, dtype=float)
            
            self.distribution_registry[stable_key] = {
                'func_name': func_name,
                'params': params,
                'markers': markers,
                'cache': cache,
                'is_nested': is_nested,
                'has_generated': [False] * self.total_iterations,
                'parent_key': parent_key,
                'depth': depth,
                'current_value': np.nan,  # 当前迭代的值
                'call_count': 0,  # 调用次数
                'input_key': None,  # 存储输入键
                'args_text': args_text,  # 存储原始参数文本
                'is_simtable': 'simtable' in func_name.lower(),  # 是否为Simtable函数
                'is_makeinput': 'makeinput' in func_name.lower()  # 是否为MakeInput函数
            }
            
            # 保存属性信息 - 改进：确保所有属性都被保存
            if attributes:
                # 首先保存传入的属性
                self.distribution_attributes[stable_key] = attributes.copy()
                self.log_debug(f"为稳定键 {stable_key} 保存属性: {attributes}")
            
            # 同时将markers中的属性也保存到distribution_attributes
            if markers:
                if stable_key not in self.distribution_attributes:
                    self.distribution_attributes[stable_key] = {}
                self.distribution_attributes[stable_key].update(markers)
            
            # 记录深度
            self.nested_depth_map[stable_key] = depth
            
            if is_nested:
                self.nested_functions.add(stable_key)
                if parent_key:
                    self.nested_parent_map[stable_key] = parent_key
                    if parent_key not in self.nested_children_map:
                        self.nested_children_map[parent_key] = []
                    if stable_key not in self.nested_children_map[parent_key]:
                        self.nested_children_map[parent_key].append(stable_key)
            
            self.log_debug(f"注册分布函数: {stable_key} -> {func_name}{params}, 嵌套={is_nested}, 深度={depth}, 父键={parent_key}, 参数文本={args_text}, 标记={markers}, 属性={attributes}")
                
    def set_input_key_for_stable_key(self, stable_key: str, input_key: str):
        """为稳定键设置对应的输入键"""
        with self._lock:
            if stable_key in self.distribution_registry:
                self.distribution_registry[stable_key]['input_key'] = input_key
                self.log_debug(f"为稳定键 {stable_key} 设置输入键: {input_key}")
    
    def set_makeinput_attributes(self, cell_addr: str, attributes: Dict[str, Any]):
        """设置MakeInput单元格的属性"""
        with self._lock:
            self.makeinput_attributes[cell_addr] = attributes.copy() if attributes else {}
            self.log_debug(f"设置MakeInput属性: {cell_addr} -> {attributes}")
    
    def get_makeinput_attributes(self, cell_addr: str) -> Optional[Dict[str, Any]]:
        """获取MakeInput单元格的属性"""
        with self._lock:
            return self.makeinput_attributes.get(cell_addr, {}).copy()
    
    def set_simtable_values(self, cell_addr: str, values: List[float]):
        """设置Simtable单元格的值"""
        with self._lock:
            self.simtable_values[cell_addr] = values.copy() if values else []
            self.log_debug(f"设置Simtable值: {cell_addr} -> {values}, 长度={len(values)}")
    
    def get_simtable_value(self, cell_addr: str, iteration: int = None) -> Optional[float]:
        """获取Simtable单元格的值 - 改进：修复场景切换逻辑"""
        with self._lock:
            if cell_addr in self.simtable_values:
                values = self.simtable_values[cell_addr]
                if not values:
                    return None
                
                if iteration is None:
                    iteration = self.current_iteration
                
                # 改进的场景切换逻辑
                if self.scenario_count > 1:
                    # 多场景模式：使用场景索引选择值
                    if self.scenario_index < len(values):
                        # 场景索引在值范围内，直接使用
                        return values[self.scenario_index]
                    else:
                        # 场景索引超出范围，循环使用
                        scenario_index = self.scenario_index % len(values)
                        return values[scenario_index]
                else:
                    # 单场景模式：使用迭代次数循环选择值
                    index = iteration % len(values)
                    return values[index]
            return None
    
    def set_current_iteration(self, iteration: int):
        """设置当前迭代次数"""
        with self._lock:
            if 0 <= iteration < self.total_iterations:
                self.current_iteration = iteration
                # 每次设置新迭代时，重置该迭代的所有生成标记
                for stable_key in self.distribution_registry:
                    if iteration < len(self.distribution_registry[stable_key]['has_generated']):
                        self.distribution_registry[stable_key]['has_generated'][iteration] = False
            else:
                if iteration % 5000 == 0:  # 减少日志频率
                    self.log_debug(f"无效的迭代次数 {iteration}, 范围应为 0-{self.total_iterations-1}")
    
    def get_current_iteration(self) -> int:
        """获取当前迭代次数"""
        with self._lock:
            return self.current_iteration
    
    def is_simulation_running(self) -> bool:
        """检查模拟是否运行中"""
        with self._lock:
            return self.simulation_running
    
    def record_random_value(self, stable_key: str, value: float):
        """记录随机数值 - 简化处理"""
        with self._lock:
            if not self.simulation_running:
                return False
            
            if stable_key not in self.distribution_registry:
                self.log_debug(f"稳定键未注册，无法记录: {stable_key}")
                return False
            
            current_iter = self.current_iteration
            if 0 <= current_iter < self.total_iterations:
                try:
                    # 检查是否已经生成过该迭代的随机数
                    dist_info = self.distribution_registry[stable_key]
                    if not dist_info['has_generated'][current_iter]:
                        dist_info['cache'][current_iter] = value
                        dist_info['has_generated'][current_iter] = True
                        dist_info['current_value'] = value
                        
                        # 跟踪调用
                        call_key = f"{stable_key}_{current_iter}"
                        self.call_tracker[call_key] = self.call_tracker.get(call_key, 0) + 1
                        dist_info['call_count'] += 1
                        
                        # 控制日志输出频率
                        current_time = time.time()
                        if current_iter < 3 or (current_time - self.last_log_time >= 10.0 and current_iter % 1000 == 0):
                            input_key = dist_info.get('input_key', 'N/A')
                            args_text = dist_info.get('args_text', 'N/A')
                            self.log_debug(f"记录随机数: {stable_key}[{current_iter}] = {value:.6f}, 输入键={input_key}, 嵌套={dist_info.get('is_nested', False)}, 深度={dist_info.get('depth', 0)}")
                            self.last_log_time = current_time
                        
                        return True
                    else:
                        # 已经生成过，更新当前值但不重复记录
                        dist_info['current_value'] = value
                        return False
                except Exception as e:
                    if current_iter % 5000 == 0:
                        self.log_debug(f"记录随机数失败 {stable_key}[{current_iter}]: {str(e)}")
                    return False
            else:
                if current_iter % 5000 == 0:
                    self.log_debug(f"迭代索引越界 {current_iter}, 范围应为 0-{self.total_iterations-1}")
                return False
    
    def record_makeinput_value(self, cell_addr: str, value: float, attributes: Dict[str, Any] = None):
        """记录MakeInput单元格的值和属性 - 关键修复：确保所有迭代都被记录"""
        with self._lock:
            if not self.simulation_running:
                return False
            
            current_iter = self.current_iteration
            if 0 <= current_iter < self.total_iterations:
                try:
                    if cell_addr not in self.makeinput_cache:
                        self.makeinput_cache[cell_addr] = np.full(self.total_iterations, np.nan, dtype=float)
                    
                    # 记录值到缓存
                    self.makeinput_cache[cell_addr][current_iter] = value
                    
                    # 同时记录到单元格值缓存
                    self.record_cell_value(cell_addr, value)
                    
                    # 保存属性 - 改进：合并属性
                    if attributes:
                        if cell_addr not in self.makeinput_attributes:
                            self.makeinput_attributes[cell_addr] = {}
                        self.makeinput_attributes[cell_addr].update(attributes)
                    
                    # 记录到日志
                    if current_iter < 3 or current_iter % 1000 == 0:
                        self.log_debug(f"记录MakeInput值: {cell_addr}[{current_iter}] = {value}, 属性={attributes}")
                    
                    return True
                except Exception as e:
                    if current_iter % 5000 == 0:
                        self.log_debug(f"记录MakeInput值失败 {cell_addr}[{current_iter}]: {str(e)}")
                    return False
            return False
    
    def get_makeinput_cache(self, cell_addr: str) -> Optional[np.ndarray]:
        """获取MakeInput单元格缓存"""
        with self._lock:
            return self.makeinput_cache.get(cell_addr)
    
    def record_cell_value(self, cell_addr: str, value: float):
        """记录单元格的值"""
        with self._lock:
            if not self.simulation_running:
                return False
            
            current_iter = self.current_iteration
            if 0 <= current_iter < self.total_iterations:
                try:
                    if cell_addr not in self.cell_value_cache:
                        self.cell_value_cache[cell_addr] = np.full(self.total_iterations, np.nan, dtype=float)
                        self.cell_value_locks[cell_addr] = threading.RLock()
                    
                    with self.cell_value_locks[cell_addr]:
                        self.cell_value_cache[cell_addr][current_iter] = value
                    
                    return True
                except Exception as e:
                    if current_iter % 5000 == 0:
                        self.log_debug(f"记录单元格值失败 {cell_addr}[{current_iter}]: {str(e)}")
                    return False
            return False
    
    def record_output_value(self, cell_addr: str, value: float):
        """记录输出单元格的值"""
        with self._lock:
            if not self.simulation_running:
                return False
            
            current_iter = self.current_iteration
            if 0 <= current_iter < self.total_iterations:
                try:
                    if cell_addr not in self.output_cache:
                        self.output_cache[cell_addr] = np.full(self.total_iterations, np.nan, dtype=float)
                    
                    self.output_cache[cell_addr][current_iter] = value
                    
                    # 同时记录到单元格值缓存
                    self.record_cell_value(cell_addr, value)
                    
                    return True
                except Exception as e:
                    if current_iter % 5000 == 0:
                        self.log_debug(f"记录输出值失败 {cell_addr}[{current_iter}]: {str(e)}")
                    return False
            return False
    
    def get_cell_value(self, cell_addr: str) -> Optional[float]:
        """获取单元格当前迭代的值"""
        with self._lock:
            if cell_addr in self.cell_value_cache:
                cache_array = self.cell_value_cache[cell_addr]
                current_iter = self.current_iteration
                if 0 <= current_iter < len(cache_array):
                    with self.cell_value_locks.get(cell_addr, threading.RLock()):
                        value = cache_array[current_iter]
                        if not np.isnan(value):
                            return float(value)
            return None
    
    def get_output_cache(self, cell_addr: str) -> Optional[np.ndarray]:
        """获取输出单元格缓存"""
        with self._lock:
            return self.output_cache.get(cell_addr)
    
    def get_random_cache(self, stable_key: str) -> Optional[np.ndarray]:
        """获取随机数缓存"""
        with self._lock:
            if stable_key in self.distribution_registry:
                return self.distribution_registry[stable_key]['cache'].copy()
            return None
    
    def get_distribution_attributes(self, stable_key: str) -> Optional[Dict[str, Any]]:
        """获取分布函数的属性信息 - 改进：返回完整属性"""
        with self._lock:
            # 获取基础属性
            base_attrs = self.distribution_attributes.get(stable_key, {}).copy()
            
            # 如果分布注册信息中有markers，也合并进去
            if stable_key in self.distribution_registry:
                dist_info = self.distribution_registry[stable_key]
                markers = dist_info.get('markers', {})
                if markers:
                    base_attrs.update(markers)
            
            return base_attrs
    
    def clear_simulation(self):
        """清除模拟状态"""
        with self._lock:
            self.current_iteration = 0
            self.total_iterations = 0
            self.simulation_running = False
            self.simulation_id = 0
            self.scenario_count = 1
            self.scenario_index = 0
            self.random_cache.clear()
            self.distribution_registry.clear()
            self.call_tracker.clear()
            self.nested_functions.clear()
            self.output_cache.clear()
            self.nested_parent_map.clear()
            self.nested_children_map.clear()
            self.cell_value_cache.clear()
            self.cell_value_locks.clear()
            self.nested_depth_map.clear()
            self.debug_log.clear()
            self.last_debug_print = 0
            self.simtable_values.clear()
            self.makeinput_cache.clear()
            self.distribution_attributes.clear()
            self.makeinput_attributes.clear()  # 新增：清除MakeInput属性
            self.cell_formula_cache.clear()
            self.makeinput_processed.clear()
            self.log_debug("全局迭代计数器已清除")
    
    def get_nested_children(self, parent_key: str) -> List[str]:
        """获取嵌套子函数"""
        with self._lock:
            return self.nested_children_map.get(parent_key, []).copy()
    
    def get_parent_key(self, child_key: str) -> Optional[str]:
        """获取父函数键"""
        with self._lock:
            return self.nested_parent_map.get(child_key)
    
    def get_depth(self, stable_key: str) -> int:
        """获取函数深度"""
        with self._lock:
            return self.nested_depth_map.get(stable_key, 0)

# 创建全局实例
_global_iter_counter = GlobalIterationCounter()

# ==================== 稳定键生成器（简化版） ====================

class StableKeyGenerator:
    """稳定键生成器 - 简化版"""
    
    def __init__(self):
        self.cell_function_counts = {}
        self.generated_keys = set()
        
    def generate_stable_key(self, cell_address: str, func_name: str, params: List[float],
                          is_at_function: bool = False, is_nested: bool = False,
                          index_in_cell: int = 1, depth: int = 0, 
                          parent_key: Optional[str] = None, args_text: str = None) -> str:
        """生成稳定键"""
        try:
            if '!' in cell_address:
                sheet_part, cell_addr_only = cell_address.split('!')
                sheet_part = sheet_part.strip()
                cell_addr_only = cell_addr_only.strip()
            else:
                sheet_part = "Sheet1"
                cell_addr_only = cell_address.strip()
            
            sheet_normalized = re.sub(r'\s+', '', sheet_part).upper()
            cell_normalized = re.sub(r'\s+', '', cell_addr_only).upper()
            
            if is_at_function:
                func_type = "AT"
            elif is_nested:
                func_type = "NESTED"
            else:
                func_type = "DIST"
            
            # 生成参数哈希
            param_hash = 0
            if args_text:
                normalized_args = args_text.strip().upper()
                normalized_args = re.sub(r'\s+', '', normalized_args)
                param_hash = abs(hash(normalized_args)) % 1000000
            elif params:
                param_strs = []
                for param in params:
                    if isinstance(param, (int, float)):
                        param_str = f"{param:.6f}"
                        param_str = param_str.rstrip('0').rstrip('.')
                        if not param_str:
                            param_str = "0"
                    else:
                        param_str = str(param)
                    param_strs.append(param_str)
                
                param_str_combined = ",".join(param_strs)
                param_str_combined = re.sub(r'\s+', '', param_str_combined)
                param_hash = abs(hash(param_str_combined)) % 1000000
            
            key_components = [
                sheet_normalized,
                cell_normalized,
                func_type,
                func_name.upper(),
                f"H{param_hash:06d}",
                f"I{index_in_cell:03d}",
                f"D{depth:02d}"
            ]
            
            if parent_key:
                if '_' in parent_key:
                    parent_simple = parent_key.split('_')[-1]
                    if len(parent_simple) > 6:
                        parent_simple = parent_simple[:6]
                else:
                    parent_simple = parent_key[:6] if len(parent_key) > 6 else parent_key
                key_components.append(f"P{parent_simple}")
            
            stable_key = "_".join(key_components)
            
            if stable_key in self.generated_keys:
                counter = 1
                while True:
                    unique_key = f"{stable_key}_V{counter:02d}"
                    if unique_key not in self.generated_keys:
                        stable_key = unique_key
                        break
                    counter += 1
            
            self.generated_keys.add(stable_key)
            
            return stable_key
            
        except Exception as e:
            import random
            fallback_key = f"{cell_address.replace('!', '_')}_{func_name}_{index_in_cell}_{random.randint(10000, 99999)}"
            return fallback_key.upper()

# 全局稳定键生成器实例
_stable_key_generator = StableKeyGenerator()

def get_stable_key(cell_address: str, func_name: str, params: List[float],
                  is_at_function: bool = False, is_nested: bool = False,
                  index_in_cell: int = 1, depth: int = 0, 
                  parent_key: Optional[str] = None, args_text: str = None) -> str:
    """获取稳定键的公共接口"""
    return _stable_key_generator.generate_stable_key(
        cell_address, func_name, params, is_at_function, 
        is_nested, index_in_cell, depth, parent_key, args_text
    )

# ==================== 分布函数UDF（简化版） ====================

@xl_func("string stable_key, var func_with_args: var", volatile=True)
def DriskSimulateWrapper(stable_key: str, func_with_args):
    """
    模拟包装函数 - 简化版
    """
    try:
        # 检查是否在模拟模式下
        from attribute_functions import get_static_mode
        static_mode = get_static_mode()
        
        # 检查模拟是否运行中
        simulation_running = _global_iter_counter.is_simulation_running()
        
        if not simulation_running or static_mode:
            # 静态模式或非模拟模式，直接返回值
            if isinstance(func_with_args, (int, float, np.number)):
                return func_with_args
            else:
                try:
                    return float(func_with_args)
                except:
                    return func_with_args
        
        # 获取当前迭代次数
        current_iter = _global_iter_counter.get_current_iteration()
        
        with threading.RLock():
            # 检查稳定键是否已注册
            if stable_key not in _global_iter_counter.distribution_registry:
                logger.warning(f"稳定键 {stable_key} 未在注册表中找到，尝试动态注册...")
                
                # 动态注册函数
                if isinstance(func_with_args, str):
                    try:
                        # 尝试从字符串中提取分布函数信息
                        funcs = extract_nested_distributions_advanced(f"={func_with_args}", "unknown")
                        
                        if funcs:
                            for func in funcs:
                                # 检查嵌套深度
                                depth = func.get('depth', 0)
                                if depth > 2:
                                    logger.error(f"嵌套深度 {depth} 超过两层限制")
                                    return "#ERROR!"
                                
                                func_name = func.get('func_name', '')
                                params = func.get('parameters', [])
                                markers = func.get('markers', {})
                                is_nested = func.get('is_nested', False)
                                depth = func.get('depth', 0)
                                args_text = func.get('args_text', '')
                                
                                # 提取属性信息 - 改进：使用完整属性提取
                                attributes = {}
                                if args_text:
                                    try:
                                        attributes = extract_input_attributes(f"={func_name}({args_text})")
                                    except:
                                        pass
                                
                                # 将markers也合并到attributes中
                                if markers:
                                    attributes.update(markers)
                                
                                # 注册到全局计数器
                                _global_iter_counter.register_distribution(
                                    stable_key, func_name, params, markers, is_nested,
                                    None, depth, args_text, attributes
                                )
                                
                                logger.info(f"动态注册函数: {stable_key} -> {func_name}{params}, 参数文本={args_text}, 属性={attributes}")
                                break
                    except Exception as e:
                        logger.error(f"动态解析函数失败: {str(e)}")
                
                if stable_key not in _global_iter_counter.distribution_registry:
                    logger.error(f"稳定键 {stable_key} 未在注册表中，返回NaN")
                    return float('nan')
            
            # 继续原有逻辑...
            dist_info = _global_iter_counter.distribution_registry[stable_key]
            cache_array = dist_info['cache']
            
            # 检查是否已生成当前迭代的随机数
            if 0 <= current_iter < len(cache_array) and not np.isnan(cache_array[current_iter]):
                # 已生成，直接返回缓存的值
                cached_value = cache_array[current_iter]
                return float(cached_value)
            
            # 需要生成新的随机数
            func_name = dist_info['func_name']
            params = dist_info['params']
            markers = dist_info['markers']
            
            # 特殊处理：Simtable函数
            if 'simtable' in func_name.lower():
                # 获取单元格地址
                cell_addr = None
                for addr, info in _global_iter_counter.distribution_registry.items():
                    if info.get('stable_key') == stable_key:
                        # 提取单元格地址
                        if '_' in addr:
                            parts = addr.split('_')
                            cell_addr = '_'.join(parts[:-1])  # 去掉最后的稳定键部分
                        else:
                            cell_addr = addr
                        break
                
                if cell_addr:
                    # 从simtable_values获取值
                    simtable_value = _global_iter_counter.get_simtable_value(cell_addr, current_iter)
                    if simtable_value is not None:
                        # 记录到缓存
                        if 0 <= current_iter < len(cache_array):
                            _global_iter_counter.record_random_value(stable_key, simtable_value)
                        return float(simtable_value)
            
            # 特殊处理：MakeInput函数 - 关键修复：确保MakeInput值被正确记录
            if 'makeinput' in func_name.lower():
                # MakeInput函数应该直接返回参数值
                if isinstance(func_with_args, (int, float, np.number)):
                    value = float(func_with_args)
                else:
                    try:
                        value = float(func_with_args)
                    except:
                        value = 0.0
                
                # 记录到MakeInput缓存
                # 获取单元格地址 - 关键修复：直接从稳定键中提取单元格地址
                if '_' in stable_key:
                    # 稳定键格式：SHEET_CELL_FUNCTYPE_FUNCNAME_HASH...
                    parts = stable_key.split('_')
                    if len(parts) >= 2:
                        sheet_name = parts[0]
                        cell_addr_only = parts[1]
                        cell_addr = f"{sheet_name}!{cell_addr_only}"
                    else:
                        cell_addr = None
                else:
                    cell_addr = None
                
                if cell_addr:
                    # 获取MakeInput的属性
                    attributes = {}
                    if args_text := dist_info.get('args_text'):
                        try:
                            attributes = extract_makeinput_attributes(f"=DriskMakeInput({args_text})")
                        except:
                            pass
                    
                    # 记录到MakeInput缓存
                    _global_iter_counter.record_makeinput_value(cell_addr, value, attributes)
                    logger.debug(f"MakeInput值已记录: {cell_addr}[{current_iter}] = {value}")
                
                return value
            
            # 生成随机数
            random_value = _generate_random_value_simple(
                func_name, params, markers, current_iter, stable_key, 
                dist_info.get('is_nested', False), dist_info.get('depth', 0)
            )
            
            # 记录到缓存
            if 0 <= current_iter < len(cache_array):
                success = _global_iter_counter.record_random_value(stable_key, random_value)
                if success:
                    logger.debug(f"生成并记录随机数: {stable_key}[{current_iter}] = {random_value:.6f}")
                else:
                    logger.warning(f"记录随机数失败: {stable_key}[{current_iter}]")
            
            return float(random_value)
            
    except Exception as e:
        logger.error(f"DriskSimulateWrapper错误: {str(e)}")
        import traceback
        traceback.print_exc()
        return "#ERROR!"
    
def _generate_random_value_simple(func_name: str, params: List[float], markers: Dict[str, Any], 
                                 current_iter: int, stable_key: str, 
                                 is_nested: bool = False, depth: int = 0) -> float:
    """简化的随机数生成函数"""
    try:
        # 检查是否有static标记
        static_value = markers.get('static')
        if static_value is not None:
            from attribute_functions import get_static_mode
            if get_static_mode():
                return float(static_value)
        
        # 检查是否有loc标记
        if markers.get('loc'):
            from attribute_functions import get_static_mode
            if not get_static_mode():
                # 计算理论均值
                return _calculate_theoretical_mean(func_name, params, markers)
        
        # 生成唯一种子：基于稳定键和当前迭代
        seed_parts = []
        seed_parts.append(42)  # 基础种子
        seed_parts.append(current_iter * 1000)
        seed_parts.append(abs(hash(stable_key) % 1000000))
        seed_parts.append(depth * 10000)
        
        # 参数相关
        for i, param in enumerate(params):
            if isinstance(param, (int, float)):
                int_part = int(abs(param) * 1000) % 10000
                seed_parts.append(int_part + i * 100)
        
        # 计算综合种子
        final_seed = sum(seed_parts) % 1000000
        
        # 根据分布类型生成随机数
        import numpy as np
        rng = np.random.Generator(np.random.MT19937(seed=final_seed))
        
        func_name_lower = func_name.lower()
        
        if 'normal' in func_name_lower:
            if len(params) >= 2:
                value = rng.normal(loc=params[0], scale=params[1])
            elif len(params) >= 1:
                value = rng.normal(loc=params[0], scale=1.0)
            else:
                value = rng.normal(loc=0, scale=1.0)
        elif 'uniform' in func_name_lower and len(params) >= 2:
            value = rng.uniform(low=params[0], high=params[1])
        elif 'gamma' in func_name_lower and len(params) >= 2:
            value = rng.gamma(shape=params[0], scale=params[1])
        elif 'poisson' in func_name_lower and len(params) >= 1:
            value = rng.poisson(lam=params[0])
        elif 'beta' in func_name_lower and len(params) >= 2:
            value = rng.beta(a=params[0], b=params[1])
        elif 'chisq' in func_name_lower and len(params) >= 1:
            value = rng.chisquare(df=params[0])
        elif 'f' in func_name_lower and len(params) >= 2:
            value = rng.f(dfnum=params[0], dfden=params[1])
        elif 't' in func_name_lower and len(params) >= 1:
            value = rng.standard_t(df=params[0])
        elif 'expon' in func_name_lower and len(params) >= 1:
            value = rng.exponential(scale=params[0])
        else:
            # 默认正态分布
            mean = params[0] if len(params) >= 1 else 0
            std = params[1] if len(params) >= 2 else 1
            value = rng.normal(loc=mean, scale=std)
        
        # 应用标记转换
        value = _apply_markers_to_value(value, markers)
        
        return float(value)
        
    except Exception as e:
        if current_iter % 5000 == 0:
            _global_iter_counter.log_debug(f"生成随机数失败 {func_name}: {str(e)}")
        return float('nan')

def _calculate_theoretical_mean(func_name: str, params: List[float], markers: Dict[str, Any]) -> float:
    """计算理论均值（用于loc标记）"""
    func_name_lower = func_name.lower()
    
    # 基本均值
    if 'normal' in func_name_lower:
        if len(params) >= 2:
            base_mean = params[0]
        elif len(params) >= 1:
            base_mean = params[0]
        else:
            base_mean = 0.0
    elif 'uniform' in func_name_lower and len(params) >= 2:
        base_mean = (params[0] + params[1]) / 2
    elif 'gamma' in func_name_lower and len(params) >= 2:
        base_mean = params[0] * params[1]
    elif 'poisson' in func_name_lower and len(params) >= 1:
        base_mean = params[0]
    elif 'beta' in func_name_lower and len(params) >= 2:
        base_mean = params[0] / (params[0] + params[1])
    elif 'chisq' in func_name_lower and len(params) >= 1:
        base_mean = params[0]
    elif 'f' in func_name_lower and len(params) >= 2:
        if params[1] > 2:
            base_mean = params[1] / (params[1] - 2)
        else:
            base_mean = params[1]
    elif 't' in func_name_lower and len(params) >= 1:
        base_mean = 0.0
    elif 'expon' in func_name_lower and len(params) >= 1:
        base_mean = params[0]
    else:
        base_mean = 0.0
    
    # 应用shift标记
    shift_val = markers.get('shift')
    if shift_val is not None:
        try:
            base_mean += float(shift_val)
        except:
            pass
    
    return base_mean

def _apply_markers_to_value(value: float, markers: Dict[str, Any]) -> float:
    """应用标记转换到值"""
    result = value
    
    # 应用shift标记
    shift_val = markers.get('shift')
    if shift_val is not None:
        try:
            result += float(shift_val)
        except:
            pass
    
    # 应用截断标记
    truncate_val = markers.get('truncate')
    truncate2_val = markers.get('truncate2')
    
    if truncate_val is not None or truncate2_val is not None:
        truncate_str = truncate_val if truncate_val is not None else truncate2_val
        if isinstance(truncate_str, str):
            try:
                truncate_str = truncate_str.strip(' ()')
                parts = truncate_str.split(',')
                if len(parts) >= 2:
                    lower = float(parts[0].strip())
                    upper = float(parts[1].strip())
                    result = max(lower, min(upper, result))
            except:
                pass
    
    return result

# ==================== 单元格值读取UDF ====================

@xl_func("string cell_addr: var", volatile=True)
def DriskGetCellValue(cell_addr: str):
    """
    获取单元格当前迭代的值
    """
    try:
        # 检查是否在模拟模式下
        from attribute_functions import get_static_mode
        static_mode = get_static_mode()
        
        # 检查模拟是否运行中
        simulation_running = _global_iter_counter.is_simulation_running()
        
        if not simulation_running or static_mode:
            # 静态模式或非模拟模式，直接读取单元格值
            try:
                app = xl_app()
                if '!' in cell_addr:
                    sheet_name, addr = cell_addr.split('!')
                    sheet = app.ActiveWorkbook.Worksheets(sheet_name)
                    cell = sheet.Range(addr)
                else:
                    cell = app.ActiveSheet.Range(cell_addr)
                
                value = cell.Value
                if value is None:
                    value = cell.Value2
                
                if isinstance(value, (int, float, np.number)):
                    return float(value)
                else:
                    return value
            except:
                return 0.0
        
        # 模拟模式下，从全局缓存获取单元格值
        cached_value = _global_iter_counter.get_cell_value(cell_addr)
        if cached_value is not None:
            return cached_value
        
        # 如果缓存中没有，尝试读取单元格值并记录到缓存
        try:
            app = xl_app()
            if '!' in cell_addr:
                sheet_name, addr = cell_addr.split('!')
                sheet = app.ActiveWorkbook.Worksheets(sheet_name)
                cell = sheet.Range(addr)
            else:
                cell = app.ActiveSheet.Range(cell_addr)
            
            value = cell.Value
            if value is None:
                value = cell.Value2
            
            if isinstance(value, (int, float, np.number)):
                float_value = float(value)
                # 记录到缓存
                _global_iter_counter.record_cell_value(cell_addr, float_value)
                return float_value
            else:
                return value
        except Exception as e:
            current_iter = _global_iter_counter.get_current_iteration()
            if current_iter % 5000 == 0:
                _global_iter_counter.log_debug(f"读取单元格值失败 {cell_addr}: {str(e)}")
            return 0.0
            
    except Exception as e:
        current_iter = _global_iter_counter.get_current_iteration() if hasattr(_global_iter_counter, 'get_current_iteration') else 0
        if current_iter % 5000 == 0:
            _global_iter_counter.log_debug(f"DriskGetCellValue错误: {str(e)}")
        return 0.0