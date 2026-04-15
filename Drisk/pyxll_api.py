# pyxll_api.py
"""PyXLL C API接口模块 - 简化版本，移除向量化"""

import numpy as np
from pyxll import xl_func, xl_arg, xl_return
from typing import List, Dict, Any, Union, Optional
import json
import base64

# 导入模拟管理器
from simulation_manager import get_simulation, get_current_sim_id

# ==================== 高性能数据交换函数 ====================

@xl_func("float mean, float std, int n_samples: var", 
         category="Drisk High Performance", 
         thread_safe=True,
         disable_function_wizard_calc=True)
@xl_arg("mean", "float")
@xl_arg("std", "float")
@xl_arg("n_samples", "int")
@xl_return("var")
def DriskGenerateNormalSamples(mean: float, std: float, n_samples: int = 1000):
    """
    生成正态分布样本（高性能版本）
    """
    try:
        if n_samples <= 0:
            n_samples = 1000
        
        if std <= 0:
            return np.full(n_samples, mean)
        
        samples = np.random.normal(mean, std, n_samples)
        return samples.reshape(-1, 1)
        
    except Exception as e:
        return np.array([[f"错误: {str(e)}"]])

@xl_func("float min_val, float max_val, int n_samples: var", 
         category="Drisk High Performance", 
         thread_safe=True,
         disable_function_wizard_calc=True)
@xl_arg("min_val", "float")
@xl_arg("max_val", "float")
@xl_arg("n_samples", "int")
@xl_return("var")
def DriskGenerateUniformSamples(min_val: float, max_val: float, n_samples: int = 1000):
    """
    生成均匀分布样本（高性能版本）
    """
    try:
        if n_samples <= 0:
            n_samples = 1000
        
        if max_val <= min_val:
            return np.full(n_samples, min_val)
        
        samples = np.random.uniform(min_val, max_val, n_samples)
        return samples.reshape(-1, 1)
        
    except Exception as e:
        return np.array([[f"错误: {str(e)}"]])

@xl_func("float shape, float scale, int n_samples: var", 
         category="Drisk High Performance", 
         thread_safe=True,
         disable_function_wizard_calc=True)
@xl_arg("shape", "float")
@xl_arg("scale", "float")
@xl_arg("n_samples", "int")
@xl_return("var")
def DriskGenerateGammaSamples(shape: float, scale: float, n_samples: int = 1000):
    """
    生成伽马分布样本（高性能版本）
    """
    try:
        if n_samples <= 0:
            n_samples = 1000
        
        if shape <= 0 or scale <= 0:
            return np.zeros(n_samples)
        
        samples = np.random.gamma(shape, scale, n_samples)
        return samples.reshape(-1, 1)
        
    except Exception as e:
        return np.array([[f"错误: {str(e)}"]])

@xl_func("var data, string operation: var",
         category="Drisk High Performance",
         thread_safe=True)
@xl_arg("data", "var[][]")
@xl_arg("operation", "str")
def DriskBatchProcess(data, operation: str):
    """
    批量处理数据
    """
    try:
        if not data:
            return np.array([])
        
        arr = np.array(data)
        
        if operation.lower() == 'mean':
            result = np.mean(arr, axis=0)
        elif operation.lower() == 'sum':
            result = np.sum(arr, axis=0)
        elif operation.lower() == 'min':
            result = np.min(arr, axis=0)
        elif operation.lower() == 'max':
            result = np.max(arr, axis=0)
        elif operation.lower() == 'std':
            result = np.std(arr, axis=0)
        else:
            return np.array([f"未知操作: {operation}"])
        
        return result.reshape(-1, 1)
        
    except Exception as e:
        return np.array([[f"错误: {str(e)}"]])

# ==================== 新增高性能预生成函数 ====================

@xl_func("float mean, float std, int n_iterations, int scenario_count, int scenario_index: var",
         category="Drisk High Performance",
         thread_safe=True,
         disable_function_wizard_calc=True)
@xl_arg("mean", "float")
@xl_arg("std", "float")
@xl_arg("n_iterations", "int")
@xl_arg("scenario_count", "int")
@xl_arg("scenario_index", "int")
def DriskPreGenerateNormal(mean: float, std: float, n_iterations: int = 1000, 
                           scenario_count: int = 1, scenario_index: int = 0):
    """
    预生成正态分布随机数（C API高性能版本）
    """
    try:
        if n_iterations <= 0:
            n_iterations = 1000
        
        if std <= 0:
            return np.full(n_iterations, mean)
        
        # 使用C API进行高性能随机数生成
        samples = np.random.normal(mean, std, n_iterations)
        
        # 添加场景影响（确保不同场景的随机数不同）
        if scenario_index > 0:
            # 轻微调整随机数以区分不同场景
            adjustment = scenario_index * 0.0001
            samples = samples * (1.0 + adjustment)
        
        return samples.reshape(-1, 1)
        
    except Exception as e:
        return np.array([[f"错误: {str(e)}"]])

@xl_func("float min_val, float max_val, int n_iterations, int scenario_count, int scenario_index: var",
         category="Drisk High Performance",
         thread_safe=True,
         disable_function_wizard_calc=True)
@xl_arg("min_val", "float")
@xl_arg("max_val", "float")
@xl_arg("n_iterations", "int")
@xl_arg("scenario_count", "int")
@xl_arg("scenario_index", "int")
def DriskPreGenerateUniform(min_val: float, max_val: float, n_iterations: int = 1000,
                            scenario_count: int = 1, scenario_index: int = 0):
    """
    预生成均匀分布随机数（C API高性能版本）
    """
    try:
        if n_iterations <= 0:
            n_iterations = 1000
        
        if max_val <= min_val:
            return np.full(n_iterations, min_val)
        
        # 使用C API进行高性能随机数生成
        samples = np.random.uniform(min_val, max_val, n_iterations)
        
        # 添加场景影响
        if scenario_index > 0:
            adjustment = scenario_index * 0.0001
            samples = samples * (1.0 + adjustment)
        
        return samples.reshape(-1, 1)
        
    except Exception as e:
        return np.array([[f"错误: {str(e)}"]])

@xl_func("int n_iterations, int scenario_count, int scenario_index: var",
         category="Drisk High Performance",
         thread_safe=True,
         disable_function_wizard_calc=True)
@xl_arg("n_iterations", "int")
@xl_arg("scenario_count", "int")
@xl_arg("scenario_index", "int")
def DriskPreGenerateRandom(n_iterations: int = 1000, scenario_count: int = 1, scenario_index: int = 0):
    """
    预生成标准正态分布随机数（C API高性能版本）
    """
    try:
        if n_iterations <= 0:
            n_iterations = 1000
        
        # 使用C API进行高性能随机数生成
        samples = np.random.randn(n_iterations)
        
        # 添加场景影响
        if scenario_index > 0:
            adjustment = scenario_index * 0.0001
            samples = samples * (1.0 + adjustment)
        
        return samples.reshape(-1, 1)
        
    except Exception as e:
        return np.array([[f"错误: {str(e)}"]])

# ==================== 高性能批量操作函数 ====================

@xl_func("string range_address, var values: string",
         category="Drisk High Performance",
         thread_safe=True)
@xl_arg("range_address", "str")
@xl_arg("values", "var[]")
def DriskSetRangeValues(range_address: str, values) -> str:
    """
    高性能设置单元格区域值（使用C API）
    """
    try:
        from pyxll import xl_app
        app = xl_app()
        
        if range_address:
            cell_range = app.Range(range_address)
            
            if isinstance(values, (list, tuple, np.ndarray)):
                if len(values) > 0:
                    # 获取数据的维度
                    if isinstance(values, np.ndarray):
                        rows, cols = values.shape if len(values.shape) == 2 else (len(values), 1)
                        value_array = values.tolist()
                    else:
                        # 尝试确定维度
                        if isinstance(values[0], (list, tuple, np.ndarray)):
                            rows = len(values)
                            cols = len(values[0]) if rows > 0 else 0
                            value_array = [list(row) for row in values]
                        else:
                            rows = len(values)
                            cols = 1
                            value_array = [[v] for v in values]
                    
                    # 调整目标区域大小
                    target_range = cell_range.Resize(rows, cols)
                    target_range.Value = value_array
                else:
                    cell_range.Value = values
            else:
                cell_range.Value = values
            
            return f"成功设置 {range_address} 的值 ({rows}x{cols})"
        else:
            return "错误: 区域地址不能为空"
        
    except Exception as e:
        return f"错误: {str(e)}"

@xl_func("string range_address: var",
         category="Drisk High Performance",
         thread_safe=True)
@xl_arg("range_address", "str")
def DriskGetRangeValues(range_address: str):
    """
    高性能获取单元格区域值（使用C API）
    """
    try:
        from pyxll import xl_app
        app = xl_app()
        
        if range_address:
            cell_range = app.Range(range_address)
            values = cell_range.Value
            
            if values is None:
                return np.array([])
            elif isinstance(values, (list, tuple)):
                # 转换为numpy数组
                return np.array(values)
            else:
                return np.array([[values]])
        else:
            return np.array([])
        
    except Exception as e:
        return np.array([[f"错误: {str(e)}"]])

# ==================== 其他函数保持不变 ====================

@xl_func("int sim_id, string cell_address: var", 
         category="Drisk High Performance", 
         thread_safe=True,
         disable_function_wizard_calc=True)
@xl_arg("sim_id", "int")
@xl_arg("cell_address", "str")
@xl_return("var")
def DriskGetOutputArray(sim_id: int, cell_address: str):
    """
    通过PyXLL C API获取Output数组数据
    """
    try:
        sim = get_simulation(sim_id)
        if sim is None:
            return np.array([[0, f"模拟 {sim_id} 不存在"]])
        
        data = sim.get_output_data(cell_address)
        if data is None:
            return np.array([[0, f"单元格 {cell_address} 无数据"]])
        
        n = len(data)
        result = np.zeros((n, 2), dtype=object)
        
        for i in range(n):
            result[i, 0] = i + 1
            result[i, 1] = float(data[i]) if not np.isnan(data[i]) else 0.0
        
        return result
        
    except Exception as e:
        return np.array([[0, f"错误: {str(e)}"]])

@xl_func("int sim_id, string input_key: var", 
         category="Drisk High Performance", 
         thread_safe=True,
         disable_function_wizard_calc=True)
@xl_arg("sim_id", "int")
@xl_arg("input_key", "str")
@xl_return("var")
def DriskGetInputArray(sim_id: int, input_key: str):
    """
    通过PyXLL C API获取Input数组数据
    """
    try:
        sim = get_simulation(sim_id)
        if sim is None:
            return np.array([[0, f"模拟 {sim_id} 不存在"]])
        
        data = sim.get_input_data(input_key)
        if data is None:
            return np.array([[0, f"输入键 {input_key} 无数据"]])
        
        n = len(data)
        result = np.zeros((n, 2), dtype=object)
        
        for i in range(n):
            result[i, 0] = i + 1
            result[i, 1] = float(data[i]) if not np.isnan(data[i]) else 0.0
        
        return result
        
    except Exception as e:
        return np.array([[0, f"错误: {str(e)}"]])

@xl_func("string cell_address, var values: string", 
         category="Drisk High Performance", 
         thread_safe=True)
@xl_arg("cell_address", "str")
@xl_arg("values", "var[]")
def DriskSetValues(cell_address: str, values) -> str:
    """
    批量设置单元格值
    """
    try:
        from pyxll import xl_app
        app = xl_app()
        
        if cell_address:
            cell = app.Range(cell_address)
            
            if isinstance(values, (list, tuple, np.ndarray)):
                if len(values) > 0:
                    if isinstance(values[0], (list, tuple, np.ndarray)):
                        rows = len(values)
                        cols = len(values[0]) if rows > 0 else 0
                        target_range = cell.Resize(rows, cols)
                        value_list = []
                        for row in values:
                            if isinstance(row, (list, tuple, np.ndarray)):
                                value_list.append(list(row))
                            else:
                                value_list.append([row])
                        target_range.Value = value_list
                    else:
                        rows = len(values)
                        target_range = cell.Resize(rows, 1)
                        target_range.Value = [[v] for v in values]
                else:
                    cell.Value = values
            else:
                cell.Value = values
            
            return f"成功设置 {cell_address} 的值"
        else:
            return "错误: 单元格地址不能为空"
        
    except Exception as e:
        return f"错误: {str(e)}"

@xl_func("string cell_address: var", 
         category="Drisk High Performance", 
         thread_safe=True)
@xl_arg("cell_address", "str")
def DriskGetValues(cell_address: str):
    """
    批量获取单元格值
    """
    try:
        from pyxll import xl_app
        app = xl_app()
        
        if cell_address:
            cell = app.Range(cell_address)
            values = cell.Value
            
            if values is None:
                return np.array([])
            elif isinstance(values, (list, tuple)):
                return np.array(values)
            else:
                return np.array([[values]])
        else:
            return np.array([])
        
    except Exception as e:
        return np.array([[f"错误: {str(e)}"]])

@xl_func("int sim_id, string cell_address: var", 
         category="Drisk High Performance", 
         thread_safe=True)
@xl_arg("sim_id", "int")
@xl_arg("cell_address", "str")
def DriskGetStatistics(sim_id: int, cell_address: str):
    """
    获取统计信息
    """
    try:
        sim = get_simulation(sim_id)
        if sim is None:
            return np.array([["统计量", "值"], ["错误", f"模拟 {sim_id} 不存在"]])
        
        stats = sim.get_output_statistics_by_range(cell_address)
        if not stats:
            return np.array([["统计量", "值"], ["错误", f"单元格 {cell_address} 无统计信息"]])
        
        result = []
        result.append(["统计量", "值"])
        
        if 'mean' in stats:
            result.append(["均值", stats['mean']])
        if 'std' in stats:
            result.append(["标准差", stats['std']])
        if 'min' in stats:
            result.append(["最小值", stats['min']])
        if 'max' in stats:
            result.append(["最大值", stats['max']])
        if 'median' in stats:
            result.append(["中位数", stats['median']])
        if 'p5' in stats:
            result.append(["5%分位数", stats['p5']])
        if 'p95' in stats:
            result.append(["95%分位数", stats['p95']])
        if 'count' in stats:
            result.append(["样本数", stats['count']])
        
        return np.array(result)
        
    except Exception as e:
        return np.array([["统计量", "值"], ["错误", f"{str(e)}"]])

@xl_func("int sim_id: var", 
         category="Drisk High Performance", 
         thread_safe=True)
@xl_arg("sim_id", "int")
def DriskGetSimulationInfo(sim_id: int):
    """
    获取模拟信息
    """
    try:
        sim = get_simulation(sim_id)
        if sim is None:
            return np.array([["字段", "值"], ["错误", f"模拟 {sim_id} 不存在"]])
        
        result = []
        result.append(["字段", "值"])
        result.append(["模拟ID", sim.sim_id])
        result.append(["名称", sim.name])
        result.append(["抽样方法", sim.sampling_method])
        result.append(["迭代次数", sim.n_iterations])
        result.append(["Input数量", len(sim.all_input_keys)])
        result.append(["Output数量", len(sim.output_cells)])
        result.append(["工作簿", sim.workbook_name or "未知"])
        result.append(["创建时间", str(sim.timestamp)])
        
        if hasattr(sim, 'input_cache'):
            result.append(["实际Input数据", len(sim.input_cache)])
        if hasattr(sim, 'output_cache'):
            result.append(["实际Output数据", len(sim.output_cache)])
        
        return np.array(result)
        
    except Exception as e:
        return np.array([["字段", "值"], ["错误", f"{str(e)}"]])

@xl_func(": var", 
         category="Drisk High Performance", 
         thread_safe=True)
def DriskListAllSimulations():
    """
    列出所有模拟
    """
    try:
        from simulation_manager import get_all_simulations
        sim_cache = get_all_simulations()
        
        if not sim_cache:
            return np.array([["模拟ID", "名称", "抽样方法", "迭代次数", "Input数", "Output数"]])
        
        result = []
        result.append(["模拟ID", "名称", "抽样方法", "迭代次数", "Input数", "Output数", "创建时间"])
        
        for sim_id, sim in sim_cache.items():
            result.append([
                sim.sim_id,
                sim.name,
                sim.sampling_method,
                sim.n_iterations,
                len(sim.all_input_keys),
                len(sim.output_cells),
                str(sim.timestamp)
            ])
        
        return np.array(result)
        
    except Exception as e:
        return np.array([["错误", f"{str(e)}"]])

@xl_func(": string", 
         category="Drisk High Performance", 
         thread_safe=True)
def DriskInitAPI():
    """
    初始化Drisk API
    """
    try:
        from simulation_manager import get_current_sim_id
        from attribute_functions import set_static_mode
        
        set_static_mode(True)
        sim_id = get_current_sim_id()
        
        return f"Drisk API已初始化。当前模拟ID: {sim_id}"
        
    except Exception as e:
        return f"Drisk API初始化失败: {str(e)}"

@xl_func(": string",
         category="Drisk High Performance",
         thread_safe=True)
def DriskClearCacheAPI():
    """
    清除Drisk API缓存
    """
    try:
        from simulation_engine import clear_dependency_cache
        clear_dependency_cache()
        return "Drisk API缓存已清除"
    except Exception as e:
        return f"Drisk API缓存清除失败: {str(e)}"