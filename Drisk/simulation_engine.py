# simulation_engine.py
"""高性能迭代模拟引擎 - 重构版，使用统一的分布生成器和公式解析器"""

import numpy as np
import time
import traceback
import threading
import statistics
from typing import Dict, List, Optional, Tuple, Any, Set

# 导入统一的公式解析器
from formula_parser import (
    parse_complete_formula,
    parse_args_with_nested_functions,
    extract_dist_params_and_markers,
    parse_marker_function,
    parse_truncate_args,
    extract_cell_references_from_args,
    extract_all_distribution_functions_with_index,
    extract_all_distribution_functions,
    extract_input_attributes,
    parse_formula_references,
    extract_simtable_functions,
    extract_makeinput_functions,
    get_simtable_value_at_index,
    remove_makeinput_function_from_formula
)

# 导入统一的分布生成器
from distribution_functions import create_distribution_generator, DistributionGenerator, ERROR_MARKER

from attribute_functions import set_static_mode, extract_markers_from_args
from constants import SAMPLING_MC, SAMPLING_LHC

# 导入分布注册表
from constants import DISTRIBUTION_REGISTRY, get_distribution_info, get_distribution_type

# 全局依赖分析缓存
_dependency_cache = {}
_dependency_cache_lock = threading.RLock()

class IterativeSimulationEngine:
    """高性能迭代模拟引擎 - 重构版，使用统一的分布生成器和公式解析器"""
    
    def __init__(self, app, distribution_cells: Dict[str, List[Dict]], 
                 simtable_cells: Dict[str, List[Dict]],
                 makeinput_cells: Dict[str, List[Dict]],
                 output_cells: Dict[str, Dict], n_iterations: int,
                 sampling_method: str = SAMPLING_MC,
                 scenario_count: int = 1,
                 scenario_index: int = 0):
        """初始化迭代模拟引擎。
        
        Args:
            app: Excel 应用对象
            distribution_cells: 分布定义单元格字典 {cell_addr: [dist_func_dict, ...]}
            simtable_cells: Simtable单元格字典 {cell_addr: [simtable_func_dict, ...]}
            makeinput_cells: MakeInput单元格字典 {cell_addr: [makeinput_func_dict, ...]}
            output_cells: 输出单元格字典 {cell_addr: output_info_dict}
            n_iterations: 迭代次数
            sampling_method: 抽样方法，支持 'MC' 和 'LHC'
            scenario_count: 场景数（用于Simtable）
            scenario_index: 当前场景索引（0-based）
        """
        self.app = app
        self.workbook = app.ActiveWorkbook
        self.distribution_cells = distribution_cells
        self.simtable_cells = simtable_cells
        self.makeinput_cells = makeinput_cells
        self.output_cells = output_cells
        self.n_iterations = n_iterations
        self.sampling_method = sampling_method
        self.scenario_count = scenario_count
        self.scenario_index = scenario_index
        
        print(f"模拟引擎初始化: 迭代次数={n_iterations:,}, 抽样={sampling_method}, 场景={scenario_index+1}/{scenario_count}")
        
        # 输入输出缓存
        self.input_samples = {}
        self.output_cache = {}
        self.output_info = {}
        
        # 单元格缓存
        self.input_cell_cache = {}
        self.output_cell_cache = {}
        self.simtable_cell_cache = {}
        self.makeinput_cell_cache = {}
        
        # 原始值缓存（用于恢复）
        self.original_values = {}
        self.original_formulas = {}
        
        # 分布生成器缓存
        self.distribution_generators = {}  # {cell_addr: [DistributionGenerator, ...]}
        
        # MakeInput单元格属性缓存
        self.makeinput_attributes = {}
        
        # 内嵌分布函数信息
        self.nested_distributions = {}  # {makeinput_cell_addr: [virtual_addr1, virtual_addr2, ...]}
        
        # 性能统计
        self.performance_stats = {
            'sample_generation_time': 0,
            'simulation_time': 0,
            'total_time': 0,
            'iterations_completed': 0
        }
        
        # 用于自适应批次的时序数据
        self._recent_iteration_times: List[float] = []
        self._recent_calc_durations: List[float] = []
        self._last_calculate_duration: float = 0.001  # 初始估计，避免零除
        
        # 构建输入键列表和依赖关系
        self.input_keys = []
        self.dependencies = {}  # 存储单元格间的依赖关系
        
        # 预生成随机数缓存
        self.pre_generated_samples = {}  # {input_key: np.array}
        
        # 分析依赖关系 - 使用缓存版本
        self._analyze_dependencies_cached()
        self.cancelled = False
        
        # 打印内嵌分布函数信息
        print(f"内嵌分布函数统计:")
        nested_count = 0
        for cell_addr, dist_funcs in distribution_cells.items():
            for dist_func in dist_funcs:
                if dist_func.get('is_nested', False):
                    nested_count += 1
                    print(f"  - {cell_addr}: {dist_func.get('func_name', 'unknown')} (index={dist_func.get('index', 'N/A')})")
        print(f"  总计: {nested_count} 个内嵌分布函数")
    
    def _parse_cell_address(self, cell_addr: str) -> Tuple[str, str]:
        """解析单元格地址为 (工作表名, 单元格地址)。
        
        Args:
            cell_addr: 单元格地址，可以是 'A1' 或 'Sheet1!A1' 格式
            
        Returns:
            (工作表名, 单元格地址) 元组
        """
        if '!' in cell_addr:
            sheet_name, addr = cell_addr.split('!')
            return sheet_name, addr
        else:
            return self.app.ActiveSheet.Name, cell_addr
    
    def _get_cell_object(self, cell_addr: str):
        """获取指定单元格的 Excel Range 对象。
        
        Args:
            cell_addr: 单元格地址
            
        Returns:
            Excel Range 对象，如果获取失败则返回 None
        """
        sheet_name, addr = self._parse_cell_address(cell_addr)
        try:
            sheet = self.workbook.Worksheets(sheet_name)
            return sheet.Range(addr)
        except Exception as e:
            print(f"获取单元格对象失败 {cell_addr}: {str(e)}")
            return None
    
    def _analyze_dependencies_cached(self):
        """分析单元格间的依赖关系并进行拓扑排序，使用缓存提高性能。
        
        构建依赖图，使用拓扑排序确保计算顺序正确。
        """
        # 创建缓存键：工作簿名 + 单元格地址列表的哈希
        try:
            cache_key_parts = []
            
            # 添加工作簿名
            try:
                workbook_name = self.workbook.Name
                cache_key_parts.append(workbook_name)
            except:
                pass
            
            # 添加单元格地址的排序列表
            all_cells = []
            all_cells.extend(self.distribution_cells.keys())
            all_cells.extend(self.makeinput_cells.keys())
            all_cells.sort()
            cache_key_parts.extend(all_cells)
            
            cache_key = hash(tuple(cache_key_parts))
            
            with _dependency_cache_lock:
                if cache_key in _dependency_cache:
                    print("使用缓存的依赖分析结果")
                    self.sorted_cells = _dependency_cache[cache_key]['sorted_cells']
                    self.dependencies = _dependency_cache[cache_key]['dependencies']
                    self.input_keys = _dependency_cache[cache_key]['input_keys']
                    return
        except Exception as e:
            print(f"依赖缓存键生成失败，重新分析: {str(e)}")
        
        # 如果缓存未命中，进行依赖分析
        print("分析单元格依赖关系...")
        
        try:
            # 在状态栏显示分析进度
            self.app.StatusBar = "正在分析依赖关系..."
        except:
            pass
        
        # 第一步：收集所有单元格引用
        cell_refs = {}
        
        for cell_addr, dist_funcs in self.distribution_cells.items():
            cell_refs[cell_addr] = []
            
            for dist_func in dist_funcs:
                args_text = dist_func.get('args_text', '')
                # 在参数中查找单元格引用 - 使用统一的公式解析器
                refs = extract_cell_references_from_args(args_text)
                cell_refs[cell_addr].extend(refs)
        
        # MakeInput单元格的依赖（公式中的引用）
        for cell_addr, makeinput_funcs in self.makeinput_cells.items():
            cell_refs[cell_addr] = []
            
            for makeinput_func in makeinput_funcs:
                formula = makeinput_func.get('formula', '')
                # 从公式中提取单元格引用
                refs = parse_formula_references(f"={formula}") if formula else []
                cell_refs[cell_addr].extend(refs)
        
        # 第二步：建立依赖关系
        for cell_addr, refs in cell_refs.items():
            self.dependencies[cell_addr] = []
            for ref in refs:
                # 检查引用的单元格是否在分布单元格或MakeInput单元格中
                if ref in self.distribution_cells:
                    self.dependencies[cell_addr].append(ref)
                elif ref in self.makeinput_cells:
                    self.dependencies[cell_addr].append(ref)
        
        # 第三步：拓扑排序检测
        self.sorted_cells = self._topological_sort(set(cell_refs.keys()))
        
        # 第四步：构建输入键列表（按依赖顺序）
        for cell_addr in self.sorted_cells:
            if cell_addr in self.distribution_cells:
                for i, dist_func in enumerate(self.distribution_cells[cell_addr]):
                    # 检查是否为内嵌分布函数
                    is_nested = dist_func.get('is_nested', False)
                    
                    if is_nested:
                        # 内嵌分布函数：使用dist_func中的key
                        input_key = dist_func.get('key', f"{cell_addr}_{i+1}")
                    else:
                        # 普通分布函数
                        input_key = f"{cell_addr}_{dist_func.get('index', i+1)}"
                    
                    self.input_keys.append(input_key)
            elif cell_addr in self.makeinput_cells:
                # MakeInput单元格本身作为一个输入键
                input_key = f"{cell_addr}_MakeInput"
                self.input_keys.append(input_key)
        
        print(f"生成的输入键列表 ({len(self.input_keys)} 个):")
        for i, key in enumerate(self.input_keys):
            print(f"  {i+1}. {key}")
        
        # 缓存结果
        try:
            with _dependency_cache_lock:
                _dependency_cache[cache_key] = {
                    'sorted_cells': self.sorted_cells,
                    'dependencies': self.dependencies,
                    'input_keys': self.input_keys
                }
                print(f"依赖分析结果已缓存 (key: {cache_key})")
        except Exception as e:
            print(f"依赖分析结果缓存失败: {str(e)}")
    
    def _topological_sort(self, cells: Set[str]) -> List[str]:
        """使用 DFS 算法进行拓扑排序，确保依赖的单元格先计算。
        
        Args:
            cells: 单元格地址集合
            
        Returns:
            排序后的单元格地址列表
        """
        visited = set()
        result = []
        
        def dfs(cell):
            if cell in visited:
                return
            visited.add(cell)
            
            # 先访问依赖的单元格
            for dep in self.dependencies.get(cell, []):
                if dep in cells:
                    dfs(dep)
            
            result.append(cell)
        
        for cell in cells:
            dfs(cell)
        
        return result
    
    def _precache_cells(self):
        """预缓存所有输入输出单元格及其原始公式和值。"""
        print("预缓存单元格...")
        
        # 按依赖顺序缓存输入单元格（包括内嵌分布函数对应的虚拟地址）
        for cell_addr in self.sorted_cells:
            if cell_addr in self.distribution_cells:
                # 检查是否是内嵌分布函数的虚拟地址
                is_nested = False
                for dist_func in self.distribution_cells[cell_addr]:
                    if dist_func.get('is_nested', False):
                        is_nested = True
                        break
                
                if not is_nested:
                    # 普通分布函数，有对应的实际单元格
                    cell = self._get_cell_object(cell_addr)
                    if cell:
                        self.input_cell_cache[cell_addr] = cell
                        # 保存原始公式和值
                        self.original_formulas[cell_addr] = cell.Formula
                        self.original_values[cell_addr] = cell.Value
                else:
                    # 内嵌分布函数，没有对应的实际单元格
                    print(f"内嵌分布函数虚拟地址: {cell_addr}")
            elif cell_addr in self.makeinput_cells:
                # MakeInput单元格
                cell = self._get_cell_object(cell_addr)
                if cell:
                    self.makeinput_cell_cache[cell_addr] = cell
                    # 保存原始公式和值
                    if cell_addr not in self.original_formulas:
                        self.original_formulas[cell_addr] = cell.Formula
                    if cell_addr not in self.original_values:
                        self.original_values[cell_addr] = cell.Value
                    
                    # 提取MakeInput单元格属性
                    makeinput_funcs = self.makeinput_cells.get(cell_addr, [])
                    if makeinput_funcs:
                        makeinput_func = makeinput_funcs[0]
                        self.makeinput_attributes[cell_addr] = makeinput_func.get('attributes', {})
        
        # 缓存Simtable单元格
        for cell_addr, simtable_funcs in self.simtable_cells.items():
            cell = self._get_cell_object(cell_addr)
            if cell:
                self.simtable_cell_cache[cell_addr] = cell
                # 保存原始公式和值
                if cell_addr not in self.original_formulas:
                    self.original_formulas[cell_addr] = cell.Formula
                if cell_addr not in self.original_values:
                    self.original_values[cell_addr] = cell.Value
        
        # 缓存输出单元格
        for cell_addr, _ in self.output_cells.items():
            cell = self._get_cell_object(cell_addr)
            if cell:
                self.output_cell_cache[cell_addr] = cell
                self.output_info[cell_addr] = self.output_cells[cell_addr]
    
    def _pre_generate_samples(self):
        """预生成所有独立分布函数的随机数样本。
        
        对于不依赖其他单元格的分布函数，提前生成所有迭代的随机数。
        """
        print("预生成随机数样本...")
        
        try:
            # 在状态栏显示预生成进度
            self.app.StatusBar = "正在预生成随机数..."
        except:
            pass
        
        start_time = time.time()
        pre_generated_count = 0
        
        # 检查哪些分布函数是独立的（参数不依赖其他单元格）
        for cell_addr in self.sorted_cells:
            if cell_addr in self.distribution_cells:
                dist_funcs = self.distribution_cells[cell_addr]
                
                for i, dist_func in enumerate(dist_funcs):
                    # 检查是否为内嵌分布函数
                    is_nested = dist_func.get('is_nested', False)
                    
                    # 确定输入键
                    if is_nested:
                        input_key = dist_func.get('key', f"{cell_addr}_{i+1}")
                    else:
                        input_key = f"{cell_addr}_{dist_func.get('index', i+1)}"
                    
                    # 检查参数是否包含单元格引用
                    args_text = dist_func.get('args_text', '')
                    cell_refs = extract_cell_references_from_args(args_text)
                    
                    # 如果参数不依赖其他单元格，预生成随机数
                    if not cell_refs:
                        try:
                            # 获取生成器
                            gens = self.distribution_generators.get(cell_addr, [])
                            if i < len(gens):
                                generator = gens[i]
                                
                                # 预生成所有迭代的随机数
                                pre_generated = np.zeros(self.n_iterations, dtype=np.float64)
                                
                                for iteration in range(self.n_iterations):
                                    # 使用与_evaluate_cell_value相同的种子逻辑
                                    seed_base = 42 + iteration * 1000 + hash(cell_addr) % 10000
                                    seed = seed_base + self.scenario_index * 1000000 + i
                                    sample = generator.generate_sample(seed)
                                    pre_generated[iteration] = float(sample)
                                
                                self.pre_generated_samples[input_key] = pre_generated
                                pre_generated_count += 1
                                print(f"  预生成: {input_key}, 长度: {self.n_iterations}")
                        except Exception as e:
                            print(f"预生成随机数失败 {input_key}: {str(e)}")
        
        end_time = time.time()
        print(f"预生成完成: {pre_generated_count} 个独立分布函数, 耗时: {end_time-start_time:.3f}秒")
    
    def _create_distribution_generators(self):
        """为所有分布单元格创建分布生成器。
        
        从原始公式中解析分布类型、参数和标记，创建对应的生成器。
        包括内嵌在MakeInput中的分布函数。
        """
        print("创建分布生成器...")
        
        # 清除旧的生成器缓存
        self.distribution_generators = {}
        
        for cell_addr in self.sorted_cells:
            if cell_addr in self.distribution_cells:
                dist_funcs = self.distribution_cells[cell_addr]
                cell_generators = []
                
                for i, dist_func in enumerate(dist_funcs):
                    # 检查是否是内嵌分布函数
                    is_nested = dist_func.get('is_nested', False)
                    
                    # dist_func 来自 formula_parser.extract_all_distribution_functions_with_index
                    func_name = dist_func.get('func_name', '')
                    args_text = dist_func.get('args_text', '')
                    
                    # 使用注册表获取分布类型
                    dist_type = get_distribution_type(func_name)
                    
                    # 解析 args_text 为参数列表（考虑嵌套函数）
                    try:
                        args_list = parse_args_with_nested_functions(args_text)
                    except Exception:
                        args_list = []
                    
                    # 提取分布参数和标记
                    dist_params, markers = extract_dist_params_and_markers(args_list)
                    
                    # 如果参数不足，使用默认值
                    if len(dist_params) == 0:
                        # 从注册表获取参数信息
                        info = get_distribution_info(func_name)
                        if info:
                            min_params = info.get('min_params', 0)
                            if min_params > 0:
                                # 使用默认参数
                                if dist_type in ['normal', 'uniform', 'gamma', 'beta', 'f']:
                                    dist_params = [0.0, 1.0]
                                elif dist_type in ['poisson', 'chisq', 't', 'expon']:
                                    dist_params = [1.0]
                    
                    # 创建分布生成器
                    # 修改处1
                    try:
                        # 🔴 修复1：补上 func_name 参数 -- 对应distribution_functions中定义需要接收的四个参数
                        generator = create_distribution_generator(func_name, dist_type, dist_params, markers)
                        cell_generators.append(generator)
                        
                        # 验证截断逻辑
                        self._validate_truncate_logic(generator, dist_type, dist_params, markers)
                        
                        # 打印生成器信息
                        if is_nested:
                            print(f"  内嵌分布生成器: {cell_addr} - {dist_type} ({dist_params}) [index={dist_func.get('index', 'N/A')}]")
                        else:
                            print(f"  普通分布生成器: {cell_addr} - {dist_type} ({dist_params})")
                            
                    except Exception as e:
                        # 创建默认生成器
                        print(f"创建分布生成器失败 {cell_addr}: {str(e)}")
                        default_params = [0.0, 1.0] if dist_type in ['normal', 'uniform'] else [1.0]
                        # 🔴 修复2：降级方案也必须补上 func_name (这里默认用 'DriskNormal')
                        generator = create_distribution_generator('DriskNormal', 'normal', default_params, {})
                        cell_generators.append(generator)
                
                self.distribution_generators[cell_addr] = cell_generators
        
        print(f"分布生成器创建完成: 总共 {len(self.distribution_generators)} 个单元格的生成器")
    
    def _validate_truncate_logic(self, generator, dist_type, dist_params, markers):
        """验证截断逻辑是否正确配置。"""
        if any(key in markers for key in ['truncate', 'truncate2', 'truncatep', 'truncatep2']):
            # 生成一些样本来验证
            test_samples = []
            for i in range(10):
                seed = 42 + i * 1000
                sample = generator.generate_sample(seed)
                test_samples.append(sample)
    
    def _is_excel_error_value(self, value) -> bool:
        """检查值是否为 Excel 错误值。
        
        Args:
            value: 要检查的值
            
        Returns:
            是否为错误值
        """
        if value is None:
            return False
        
        # Excel错误值的常见表示
        if isinstance(value, (int, float)):
            # Excel错误值的数字表示
            excel_errors = {
                -2146826273: True,  # #VALUE!
                -2146826246: True,  # #DIV/0!
                -2146826259: True,  # #NAME?
                -2146826252: True,  # #N/A
                -2146826265: True,  # #REF!
                -2146826281: True,  # #NUM!
                -2146826250: True,  # #NULL!
            }
            return excel_errors.get(int(value), False)
        
        if isinstance(value, str):
            # 字符串形式的错误
            return value.startswith('#') and value.endswith('!')
        
        return False
    
    def _convert_to_valid_value(self, value):
        """将值转换为有效的数值或错误标记。
        
        Args:
            value: 待转换的值
            
        Returns:
            float 或 ERROR_MARKER
        """
        if value is None:
            return ERROR_MARKER
        
        # 检查是否为字符串形式的错误标记
        if isinstance(value, str) and value.strip().upper() == "#ERROR!":
            return ERROR_MARKER
        
        # 检查是否为Excel错误值
        if self._is_excel_error_value(value):
            return ERROR_MARKER
        
        # 检查是否已经是数值
        if isinstance(value, (int, float, np.number)):
            try:
                # 检查是否为NaN或无穷大
                if np.isnan(value) or np.isinf(value):
                    return ERROR_MARKER
                return float(value)
            except (TypeError, ValueError):
                return ERROR_MARKER
        
        # 尝试转换为数值
        if isinstance(value, str):
            # 移除可能的空格
            value_str = value.strip()
            
            # 检查是否为错误标记
            if value_str.upper() == "#ERROR!":
                return ERROR_MARKER
            
            # 检查是否为Excel错误字符串
            if value_str.startswith('#') and value_str.endswith('!'):
                return ERROR_MARKER
            
            # 尝试转换为数值
            try:
                # 尝试直接转换
                num_val = float(value_str)
                if np.isnan(num_val) or np.isinf(num_val):
                    return ERROR_MARKER
                return num_val
            except ValueError:
                # 尝试处理逗号分隔的数字（如1,000）
                try:
                    value_str_no_comma = value_str.replace(',', '')
                    num_val = float(value_str_no_comma)
                    if np.isnan(num_val) or np.isinf(num_val):
                        return ERROR_MARKER
                    return num_val
                except ValueError:
                    # 尝试处理百分比（如50%）
                    if value_str.endswith('%'):
                        try:
                            percent_val = float(value_str[:-1].strip())
                            return percent_val / 100.0
                        except ValueError:
                            return ERROR_MARKER
                    else:
                        return ERROR_MARKER
        
        # 其他类型，尝试转换为字符串再处理
        try:
            str_val = str(value)
            if str_val.upper() == "#ERROR!":
                return ERROR_MARKER
            
            # 再次尝试转换
            try:
                num_val = float(str_val)
                if np.isnan(num_val) or np.isinf(num_val):
                    return ERROR_MARKER
                return num_val
            except ValueError:
                return ERROR_MARKER
        except Exception:
            return ERROR_MARKER
    
    def _evaluate_cell_value(self, cell_addr: str, iteration: int, 
                            cell_values_cache: Dict[str, Any]) -> Any:
        """计算单元格的值，使用统一的分布生成器。
        
        Args:
            cell_addr: 单元格地址
            iteration: 当前迭代次数
            cell_values_cache: 单元格值缓存
            
        Returns:
            计算得到的值或 ERROR_MARKER
        """
        # 如果已经计算过，直接返回
        if cell_addr in cell_values_cache:
            return cell_values_cache[cell_addr]
        
        # 首先检查是否是Simtable单元格
        if cell_addr in self.simtable_cells:
            simtable_funcs = self.simtable_cells[cell_addr]
            if simtable_funcs:
                # 使用第一个Simtable函数
                simtable_func = simtable_funcs[0]
                values = simtable_func.get('values', [])
                
                # 根据场景索引获取值
                if self.scenario_index < len(values):
                    value = values[self.scenario_index]
                    # 转换为数值
                    try:
                        num_value = float(value)
                        cell_values_cache[cell_addr] = num_value
                        return num_value
                    except (ValueError, TypeError):
                        # 如果无法转换为数值，返回原始值
                        cell_values_cache[cell_addr] = value
                        return value
                else:
                    # 场景索引超出范围，返回错误标记
                    cell_values_cache[cell_addr] = ERROR_MARKER
                    return ERROR_MARKER
        
        # 检查是否是MakeInput单元格
        if cell_addr in self.makeinput_cells:
            # MakeInput单元格的值在模拟过程中由Excel计算
            # 我们只需要在每次迭代后读取它的值
            # 这里返回None，表示需要从Excel读取
            cell_values_cache[cell_addr] = None
            return None
        
        # 获取单元格的分布生成器
        if cell_addr not in self.distribution_generators:
            # 如果不是分布单元格，尝试从缓存中获取
            if cell_addr in self.input_cell_cache:
                # 这是一个输入单元格但没有分布生成器？
                cell_values_cache[cell_addr] = ERROR_MARKER
                return ERROR_MARKER
            else:
                # 不是输入单元格，可能是其他引用，返回0
                cell_values_cache[cell_addr] = 0.0
                return 0.0
        
        # 获取该单元格的所有生成器
        cell_generators = self.distribution_generators[cell_addr]
        
        # 如果有多个生成器，只使用第一个（或根据需要处理）
        if not cell_generators:
            cell_values_cache[cell_addr] = ERROR_MARKER
            return ERROR_MARKER
        
        # 使用第一个生成器（假设每个单元格只有一个分布）
        generator = cell_generators[0]
        
        # 检查是否有静态值标记
        static_value = generator.markers.get('static')
        if static_value is not None:
            # static标记仅在非模拟（静态模式）下使用
            from attribute_functions import get_static_mode
            if get_static_mode():  # 静态模式
                cell_values_cache[cell_addr] = float(static_value)
                return float(static_value)
            # 模拟模式下忽略static标记，使用正常随机数
        
        # 检查是否有loc标记
        if generator.markers.get('loc'):
            from attribute_functions import get_static_mode
            if not get_static_mode():  # 模拟模式下使用理论均值
                loc_value = generator._calculate_loc_value()
                cell_values_cache[cell_addr] = float(loc_value)
                return float(loc_value)
            # 非模拟模式下忽略loc标记，使用正常随机数
        
        # 检查是否预生成过随机数
        input_key = None
        for i, dist_func in enumerate(self.distribution_cells[cell_addr]):
            if dist_func.get('is_nested', False):
                input_key = dist_func.get('key', f"{cell_addr}_{i+1}")
            else:
                input_key = f"{cell_addr}_{dist_func.get('index', i+1)}"
            break
        
        if input_key and input_key in self.pre_generated_samples:
            # 从预生成的数组中获取值
            try:
                if iteration < len(self.pre_generated_samples[input_key]):
                    value = self.pre_generated_samples[input_key][iteration]
                    cell_values_cache[cell_addr] = float(value)
                    return float(value)
            except Exception as e:
                print(f"从预生成数组获取值失败 {input_key}: {str(e)}")
        
        # 生成随机样本
        try:
            # 关键修改：加入场景索引，确保不同场景相同迭代次数生成不同的随机数
            # 使用场景索引乘以一个大数来确保不同场景的随机数种子差异足够大
            # 同时使用迭代次数和单元格地址作为种子的一部分，确保可重复性
            seed_base = 42 + iteration * 1000 + hash(cell_addr) % 10000
            # 添加场景索引影响，确保不同场景的随机数不同
            seed = seed_base + self.scenario_index * 1000000
            
            # 关键修改：使用generate_sample方法，该方法内部会正确处理shift和truncate的顺序
            value = generator.generate_sample(seed)
            
            # 确保是有效值
            if value == ERROR_MARKER or (isinstance(value, float) and (np.isnan(value) or np.isinf(value))):
                cell_values_cache[cell_addr] = ERROR_MARKER
                return ERROR_MARKER
            
            cell_values_cache[cell_addr] = float(value)
            return float(value)
            
        except Exception as e:
            print(f"生成随机样本失败 {cell_addr}: {str(e)}")
            return ERROR_MARKER
    
    def _optimize_excel_settings(self):
        """优化 Excel 设置以提高性能。
        
        禁用自动计算、屏幕更新和事件处理。
        """
        try:
            self.original_calculation = self.app.Calculation
            self.original_screen_updating = self.app.ScreenUpdating
            self.original_events = self.app.EnableEvents
            
            # 设置为手动计算
            self.app.Calculation = -4135  # xlCalculationManual
            self.app.ScreenUpdating = False
            self.app.EnableEvents = False
            self.app.DisplayAlerts = False
            
        except Exception as e:
            print(f"Excel优化设置失败: {str(e)}")
    
    def _restore_excel_settings(self):
        """恢复原始 Excel 设置。"""
        try:
            if hasattr(self, 'original_calculation'):
                self.app.Calculation = self.original_calculation
            if hasattr(self, 'original_screen_updating'):
                self.app.ScreenUpdating = self.original_screen_updating
            if hasattr(self, 'original_events'):
                self.app.EnableEvents = self.original_events
            
            self.app.DisplayAlerts = True
            self.app.Calculation = -4105  # xlCalculationAutomatic
            
        except Exception as e:
            print(f"恢复Excel设置失败: {str(e)}")
    
    def _calculate_dynamic_batch_size(self, elapsed_time, completed_iterations):
        """计算自适应批次大小。
        
        基于最近的迭代时间和 Calculate 调用的开销，动态调整批次大小。
        使用指数加权移动平均（EWMA）算法，对最近数据赋予更高权重，更快适应变化。
        
        Args:
            elapsed_time: 已耗时（秒）
            completed_iterations: 已完成的迭代数
            
        Returns:
            下一个批次的大小
        """
        # 如果还没有足够的迭代时间数据，使用保守估计
        if len(self._recent_iteration_times) < 5:
            if completed_iterations > 0:
                # 使用全局平均
                avg_iter_time = elapsed_time / completed_iterations
                # 初始批次大小：根据迭代时间动态调整
                if avg_iter_time < 0.001:  # 非常快
                    return min(200, self.n_iterations - completed_iterations)
                elif avg_iter_time < 0.005:  # 快
                    return min(100, self.n_iterations - completed_iterations)
                elif avg_iter_time < 0.01:  # 中等
                    return min(50, self.n_iterations - completed_iterations)
                else:  # 慢
                    return min(20, self.n_iterations - completed_iterations)
            else:
                # 还没有完成任何迭代，使用保守估计
                return min(10, self.n_iterations)
        
        # 使用指数加权移动平均（EWMA）计算平均迭代时间，对最近数据赋予更高权重
        # alpha = 0.3 表示最近数据权重为30%，历史权重为70%
        alpha = 0.3
        
        # 计算EWMA迭代时间
        ewma_iter_time = 0
        recent_iters = self._recent_iteration_times[-50:]  # 使用最近50个数据点
        for i, t in enumerate(recent_iters):
            # 更近的数据权重更高（指数衰减）
            weight = alpha * ((1 - alpha) ** (len(recent_iters) - i - 1))
            ewma_iter_time += t * weight
        
        # 归一化权重
        total_weight = sum(alpha * ((1 - alpha) ** i) for i in range(len(recent_iters)))
        ewma_iter_time /= total_weight if total_weight > 0 else 1
        
        # 计算Calculate耗时的EWMA
        if len(self._recent_calc_durations) > 0:
            calc_durations = self._recent_calc_durations[-10:]  # 使用最近10个Calculate耗时
            ewma_calc_time = 0
            for i, t in enumerate(calc_durations):
                weight = alpha * ((1 - alpha) ** (len(calc_durations) - i - 1))
                ewma_calc_time += t * weight
            
            total_weight = sum(alpha * ((1 - alpha) ** i) for i in range(len(calc_durations)))
            ewma_calc_time /= total_weight if total_weight > 0 else 1
        else:
            ewma_calc_time = self._last_calculate_duration
        
        # 如果Calculate耗时异常大或小，进行限制
        ewma_calc_time = max(0.001, min(ewma_calc_time, 0.1))  # 限制在1ms到100ms之间
        ewma_iter_time = max(0.0001, min(ewma_iter_time, 0.1))  # 限制在0.1ms到100ms之间
        
        # 目标：让Calculate开销占批次总耗时的比例小于target_fraction
        target_fraction = 0.05  # 目标5%
        
        # 基于EWMA值估算最优批次大小
        try:
            # 基本公式：batch = (calc_time * (1 - target)) / (iter_time * target)
            est_batch = int((ewma_calc_time * (1.0 - target_fraction)) / (ewma_iter_time * target_fraction))
        except Exception:
            est_batch = 10
        
        # 考虑迭代时间的稳定性：如果迭代时间变化很大，减小批次大小
        if len(recent_iters) >= 10:
            try:
                # 计算变异系数（标准差/均值）
                cv = statistics.stdev(recent_iters) / (statistics.mean(recent_iters) + 1e-9)
                # 如果变异系数大，减小批次大小
                if cv > 0.5:  # 变异系数大于50%
                    est_batch = int(est_batch / (1 + cv))
            except:
                pass
        
        # 基于整体速度的调整
        overall_speed = (completed_iterations / elapsed_time) if (elapsed_time > 0 and completed_iterations > 0) else 0
        
        # 根据整体速度调整批次大小
        if overall_speed > 5000:  # 非常快
            est_batch = max(est_batch, 1000)
        elif overall_speed > 1000:  # 快
            est_batch = max(est_batch, 500)
        elif overall_speed > 200:  # 中等
            est_batch = max(est_batch, 200)
        elif overall_speed > 50:  # 较慢
            est_batch = max(est_batch, 100)
        else:  # 慢
            est_batch = max(est_batch, 50)
        
        # 确保最小批次大小
        min_batch = 1
        
        # 基于迭代时间动态调整最小批次
        if ewma_iter_time < 0.001:  # 迭代非常快（<1ms）
            min_batch = 100
        elif ewma_iter_time < 0.005:  # 迭代快（<5ms）
            min_batch = 50
        elif ewma_iter_time < 0.01:  # 迭代中等（<10ms）
            min_batch = 20
        else:  # 迭代慢（>=10ms）
            min_batch = 10
        
        # 考虑剩余迭代数的智能调整
        remaining = max(1, self.n_iterations - completed_iterations)
        
        # 如果剩余迭代数很少，减小批次大小
        if remaining < 100:
            min_batch = max(1, remaining // 10)
        elif remaining < 500:
            min_batch = max(5, remaining // 50)
        
        # 应用最小和最大限制
        est_batch = max(min_batch, est_batch)
        est_batch = min(est_batch, 5000)  # 上限避免过大
        
        # 最终限制：不超过剩余迭代数，且至少为1
        final_batch = min(est_batch, remaining)
        
        # 打印调试信息（可选）
        if completed_iterations % 100 == 0:
            print(f"批次大小计算: ewma_iter={ewma_iter_time:.6f}s, ewma_calc={ewma_calc_time:.6f}s, "
                  f"est_batch={est_batch}, final_batch={final_batch}, remaining={remaining}")
        
        return final_batch

    def _run_simulation_iteratively(self, progress_callback=None, cancel_event=None):
        """迭代运行模拟。
        
        按批次处理迭代，每次迭代后计算 Excel 工作簿。
        
        Args:
            progress_callback: 进度回调函数，接收已完成迭代数
            cancel_event: 取消事件
            
        Returns:
            是否成功运行
        """
        print(f"开始迭代模拟，共 {self.n_iterations:,} 次迭代，场景 {self.scenario_index+1}/{self.scenario_count}...")
        simulation_start = time.time()
        
        # 初始化输出缓存 - 使用object数组以存储字符串和数字
        for cell_addr in self.output_cell_cache:
            self.output_cache[cell_addr] = np.empty(self.n_iterations, dtype=object)
        
        # 初始化输入缓存（包括分布单元格、内嵌分布函数和MakeInput单元格）
        # 对于可能包含多个分布函数的单元格，使用列表保存每个分布的数组
        for cell_addr in self.distribution_cells:
            # 检查该单元格有多少个分布函数
            num_funcs = len(self.distribution_cells[cell_addr])
            self.input_samples[cell_addr] = [np.empty(self.n_iterations, dtype=object) for _ in range(num_funcs)]
        
        # 为 MakeInput 单元格只分配一个数组（用于存储最终计算值）
        for cell_addr in self.makeinput_cells:
            # MakeInput单元格只记录最终计算值，所以只分配一个数组
            self.input_samples[cell_addr] = [np.empty(self.n_iterations, dtype=object)]
        
        # 首次恢复公式
        for cell_addr, cell in self.input_cell_cache.items():
            original_formula = self.original_formulas.get(cell_addr)
            if original_formula:
                try:
                    cell.Formula = original_formula
                except Exception as e:
                    print(f"恢复输入单元格公式失败 {cell_addr}: {str(e)}")
                    pass
        
        # 恢复MakeInput单元格公式
        for cell_addr, cell in self.makeinput_cell_cache.items():
            original_formula = self.original_formulas.get(cell_addr)
            if original_formula:
                try:
                    cell.Formula = original_formula
                except Exception as e:
                    print(f"恢复MakeInput单元格公式失败 {cell_addr}: {str(e)}")
                    pass
        
        # 设置Simtable单元格的值（当前场景）
        for cell_addr, cell in self.simtable_cell_cache.items():
            simtable_funcs = self.simtable_cells.get(cell_addr, [])
            if simtable_funcs:
                simtable_func = simtable_funcs[0]
                values = simtable_func.get('values', [])
                
                if self.scenario_index < len(values):
                    # 设置当前场景的值
                    value = values[self.scenario_index]
                    try:
                        # 尝试转换为数值
                        num_value = float(value)
                        cell.Value2 = num_value
                    except (ValueError, TypeError):
                        # 如果无法转换为数值，设置为字符串
                        cell.Value2 = str(value)
                else:
                    # 场景索引超出范围，设置为错误标记
                    cell.Value2 = "#ERROR!"
        
        # 等待公式恢复
        time.sleep(0.1)
        
        completed_iterations = 0
        
        print("样本预生成完成")
        
        # 按批次处理（批次大小由自适应函数决定，但保证每次迭代后计算）
        while completed_iterations < self.n_iterations:
            # 检查取消
            if cancel_event and cancel_event.is_set():
                self.cancelled = True
                break
            
            # 计算动态批次大小（使用优化的自适应方法）
            current_elapsed = time.time() - simulation_start
            if completed_iterations > 0:
                dynamic_batch_size = self._calculate_dynamic_batch_size(
                    current_elapsed, completed_iterations
                )
            else:
                dynamic_batch_size = 10
            
            # 确保不超过总迭代次数
            batch_size = min(dynamic_batch_size, self.n_iterations - completed_iterations)
            batch_start = completed_iterations
            batch_end = batch_start + batch_size
            
            print(f"处理批次 {batch_start+1}-{batch_end} (批次大小: {batch_size})...")
            
            try:
                # 处理当前批次
                for i in range(batch_size):
                    iteration = batch_start + i
                    iter_start_time = time.time()
                    
                    # 缓存本次迭代的单元格值
                    cell_values_cache = {}
                    has_input_error = False
                    
                    # 按依赖顺序计算所有输入单元格（包括内嵌分布函数）
                    for cell_addr in self.sorted_cells:
                        # 处理分布单元格（包括内嵌分布函数）
                        if cell_addr in self.distribution_cells:
                            # 处理该单元格中所有分布生成器，逐个生成样本并保存
                            gens = self.distribution_generators.get(cell_addr, [])
                            if not gens:
                                # 退回到原有计算路径
                                cell_value = self._evaluate_cell_value(cell_addr, iteration, cell_values_cache)
                                if cell_value == ERROR_MARKER:
                                    has_input_error = True
                                # 存储到单一位置（保持兼容）
                                if cell_addr in self.input_samples:
                                    try:
                                        arr = self.input_samples[cell_addr][0]
                                        arr[iteration] = ERROR_MARKER if cell_value == ERROR_MARKER else float(cell_value)
                                    except Exception:
                                        pass
                            else:
                                # 为每个生成器生成样本
                                any_error = False
                                for j, generator in enumerate(gens):
                                    # 检查是否有预生成的随机数
                                    input_key = None
                                    dist_func = self.distribution_cells[cell_addr][j]
                                    if dist_func.get('is_nested', False):
                                        input_key = dist_func.get('key', f"{cell_addr}_{j+1}")
                                    else:
                                        input_key = f"{cell_addr}_{dist_func.get('index', j+1)}"
                                    
                                    if input_key in self.pre_generated_samples:
                                        # 使用预生成的随机数
                                        try:
                                            if iteration < len(self.pre_generated_samples[input_key]):
                                                value = self.pre_generated_samples[input_key][iteration]
                                                if value == ERROR_MARKER or (isinstance(value, float) and (np.isnan(value) or np.isinf(value))):
                                                    stored = ERROR_MARKER
                                                else:
                                                    stored = float(value)
                                            else:
                                                stored = ERROR_MARKER
                                        except Exception:
                                            stored = ERROR_MARKER
                                    else:
                                        # 动态生成随机数
                                        try:
                                            seed_base = 42 + iteration * 1000 + (hash(cell_addr) % 10000)
                                            seed = seed_base + self.scenario_index * 1000000 + j
                                            value = generator.generate_sample(seed)
                                            if value == ERROR_MARKER or (isinstance(value, float) and (np.isnan(value) or np.isinf(value))):
                                                stored = ERROR_MARKER
                                            else:
                                                stored = float(value)
                                        except Exception:
                                            stored = ERROR_MARKER
                                    
                                    # 存储到对应的数组位置
                                    if cell_addr in self.input_samples and j < len(self.input_samples[cell_addr]):
                                        try:
                                            arr = self.input_samples[cell_addr][j]
                                            arr[iteration] = stored
                                        except Exception:
                                            pass
                                    if stored == ERROR_MARKER:
                                        any_error = True
                                
                                # 如果只有一个生成器且该单元格不是 MakeInput，则把该值放到缓存供依赖使用
                                if len(gens) == 1 and cell_addr not in self.makeinput_cell_cache:
                                    try:
                                        cell_values_cache[cell_addr] = float(self.input_samples[cell_addr][0][iteration])
                                    except Exception:
                                        cell_values_cache[cell_addr] = ERROR_MARKER
                                
                                if any_error:
                                    has_input_error = True
                        
                        # 设置单元格值（如果是数值且不是内嵌分布函数）
                        if cell_addr in self.input_cell_cache:
                            try:
                                # 如果该单元格同时被标记为 MakeInput，则不要覆盖其公式
                                if cell_addr in self.makeinput_cell_cache:
                                    # 跳过对 Value2 的写入，保留公式
                                    pass
                                else:
                                    # 确定要写入 Excel 的值：优先使用 cell_values_cache，然后使用生成器生成的样本
                                    write_value = None
                                    if cell_addr in cell_values_cache and cell_values_cache[cell_addr] is not None:
                                        write_value = cell_values_cache[cell_addr]
                                    else:
                                        gens = self.distribution_generators.get(cell_addr, [])
                                        if gens and cell_addr in self.input_samples:
                                            try:
                                                write_value = self.input_samples[cell_addr][0][iteration]
                                            except Exception:
                                                write_value = None
                                        else:
                                            # 回退到评估函数
                                            try:
                                                write_value = self._evaluate_cell_value(cell_addr, iteration, cell_values_cache)
                                            except:
                                                write_value = None
                                    
                                    if write_value is None or write_value == ERROR_MARKER:
                                        # 输入有错误，设置特殊值以触发计算
                                        try:
                                            self.input_cell_cache[cell_addr].Value2 = "#ERROR!"
                                        except:
                                            pass
                                    else:
                                        try:
                                            if isinstance(write_value, (int, float, np.number)):
                                                if not np.isnan(write_value) and not np.isinf(write_value):
                                                    self.input_cell_cache[cell_addr].Value2 = float(write_value)
                                            else:
                                                # 尝试转换为数字
                                                num_value = float(write_value)
                                                if not np.isnan(num_value) and not np.isinf(num_value):
                                                    self.input_cell_cache[cell_addr].Value2 = num_value
                                                else:
                                                    self.input_cell_cache[cell_addr].Value2 = "#ERROR!"
                                        except Exception:
                                            try:
                                                self.input_cell_cache[cell_addr].Value2 = "#ERROR!"
                                            except:
                                                pass
                            except Exception:
                                pass
                    
                    # 每次迭代后立即计算（保证每次输出不同）
                    try:
                        calc_t0 = time.time()
                        self.app.Calculate()
                        calc_t1 = time.time()
                        calc_dur = calc_t1 - calc_t0
                        # 更新最近的 Calculate 耗时记录
                        self._last_calculate_duration = calc_dur
                        self._recent_calc_durations.append(calc_dur)
                        if len(self._recent_calc_durations) > 50:
                            self._recent_calc_durations.pop(0)
                    except Exception as e:
                        # 忽略单次 Calculate 的异常，继续尝试读取
                        pass
                    
                    # 关键修改：正确读取MakeInput单元格的值（只读取最终计算值）
                    for cell_addr, cell in self.makeinput_cell_cache.items():
                        try:
                            # 强制重新计算当前单元格
                            try:
                                cell.Calculate()
                            except:
                                pass
                            
                            # 获取值
                            value = cell.Value2
                            
                            # 使用统一的转换函数
                            converted_value = self._convert_to_valid_value(value)
                            
                            # 确保数组足够大（input_samples 为 list of arrays）
                            if cell_addr in self.input_samples:
                                arr_list = self.input_samples[cell_addr]
                                # 只有一个数组（用于存储最终计算值）
                                if len(arr_list) > 0:
                                    arr = arr_list[0]
                                    try:
                                        if iteration >= len(arr):
                                            current_len = len(arr)
                                            new_len = max(iteration + 1, current_len * 2)
                                            new_array = np.empty(new_len, dtype=object)
                                            new_array[:current_len] = arr
                                            arr_list[0] = new_array
                                    except Exception:
                                        pass
                                
                                # 存储转换后的值到数组中
                                try:
                                    if len(self.input_samples[cell_addr]) > 0:
                                        self.input_samples[cell_addr][0][iteration] = converted_value
                                except Exception:
                                    pass
                                
                                # 调试信息：第一次迭代时打印
                                if iteration == 0:
                                    print(f"MakeInput单元格 {cell_addr} 的原始值: {value}, 转换后值: {converted_value}")
                                
                        except Exception as e:
                            print(f"读取MakeInput单元格 {cell_addr} 失败: {str(e)}")
                            # 存储错误标记
                            if cell_addr in self.input_samples:
                                arr_list = self.input_samples[cell_addr]
                                if len(arr_list) > 0:
                                    try:
                                        arr = arr_list[0]
                                        if iteration >= len(arr):
                                            current_len = len(arr)
                                            new_len = max(iteration + 1, current_len * 2)
                                            new_array = np.empty(new_len, dtype=object)
                                            new_array[:current_len] = arr
                                            arr_list[0] = new_array
                                    except Exception:
                                        pass
                                try:
                                    if len(self.input_samples[cell_addr]) > 0:
                                        self.input_samples[cell_addr][0][iteration] = ERROR_MARKER
                                except Exception:
                                    pass
                    
                    # 收集输出值 - 处理可能的错误
                    for cell_addr, cell in self.output_cell_cache.items():
                        try:
                            # 强制重新计算当前单元格
                            try:
                                cell.Calculate()
                            except:
                                pass
                            
                            # 获取值
                            value = cell.Value2
                            
                            # 如果输入有错误，输出应该标记为错误
                            if has_input_error:
                                value = ERROR_MARKER
                            elif value is None:
                                value = 0.0  # None视为0而不是错误
                            else:
                                # 使用统一的转换函数
                                value = self._convert_to_valid_value(value)
                            
                            # 确保数组足够大
                            if cell_addr in self.output_cache:
                                if iteration >= len(self.output_cache[cell_addr]):
                                    # 扩展数组
                                    current_len = len(self.output_cache[cell_addr])
                                    new_len = max(iteration + 1, current_len * 2)
                                    new_array = np.empty(new_len, dtype=object)
                                    new_array[:current_len] = self.output_cache[cell_addr]
                                    self.output_cache[cell_addr] = new_array
                                
                                # 存储值
                                self.output_cache[cell_addr][iteration] = value
                                
                        except Exception as e:
                            pass
                    
                    # 记录本次迭代耗时（包括写入、计算、读取）
                    iter_end_time = time.time()
                    iter_dur = iter_end_time - iter_start_time
                    self._recent_iteration_times.append(iter_dur)
                    if len(self._recent_iteration_times) > 200:
                        self._recent_iteration_times.pop(0)
                
                completed_iterations = batch_end
                
                # 更新进度
                if progress_callback:
                    progress_callback(completed_iterations)
                
                # 计算批次速度
                batch_time = time.time() - simulation_start
                if batch_time > 0:
                    speed = completed_iterations / batch_time
            
            except Exception as e:
                print(f"处理批次失败: {str(e)}")
                traceback.print_exc()
                break
        
        simulation_time = time.time() - simulation_start
        self.performance_stats['simulation_time'] = simulation_time
        self.performance_stats['iterations_completed'] = completed_iterations
        
        print(f"迭代模拟完成: {completed_iterations} 次迭代, 耗时: {simulation_time:.3f}秒")
        
        # 统计错误情况
        error_stats = {}
        for cell_addr, data_list in self.input_samples.items():
            if isinstance(data_list, list):
                for i, data in enumerate(data_list):
                    if isinstance(data, np.ndarray):
                        error_count = np.sum(data == ERROR_MARKER)
                        if error_count > 0:
                            error_stats[f"输入 {cell_addr}[{i}]"] = error_count
        
        for cell_addr, data in self.output_cache.items():
            if isinstance(data, np.ndarray):
                error_count = np.sum(data == ERROR_MARKER)
                if error_count > 0:
                    error_stats[f"输出 {cell_addr}"] = error_count
        
        if error_stats:
            print(f"错误统计: {error_stats}")
        
        return completed_iterations > 0
    
    def _restore_input_formulas_safe(self):
        """安全地恢复输入公式。
        
        如果恢复失败，进行重试。
        """
        max_retries = 3
        for cell_addr, cell in self.input_cell_cache.items():
            original_formula = self.original_formulas.get(cell_addr)
            
            if original_formula:
                for retry in range(max_retries):
                    try:
                        cell.Formula = original_formula
                        break
                    except Exception as e:
                        time.sleep(0.1)
            else:
                # 如果没有公式，设置为空
                try:
                    cell.Value2 = ""
                except Exception as e:
                    print(f"清空单元格 {cell_addr} 失败: {str(e)}")
    
    def _restore_makeinput_formulas_safe(self):
        """安全地恢复MakeInput单元格的公式。
        
        如果恢复失败，进行重试。
        """
        max_retries = 3
        for cell_addr, cell in self.makeinput_cell_cache.items():
            original_formula = self.original_formulas.get(cell_addr)
            
            if original_formula:
                for retry in range(max_retries):
                    try:
                        cell.Formula = original_formula
                        break
                    except Exception as e:
                        print(f"恢复MakeInput单元格公式失败 {cell_addr} (尝试{retry+1}/{max_retries}): {str(e)}")
                        time.sleep(0.1)
            else:
                # 如果没有公式，设置为空
                try:
                    cell.Value2 = ""
                except Exception as e:
                    print(f"清空MakeInput单元格 {cell_addr} 失败: {str(e)}")
    
    def _restore_simtable_formulas_safe(self):
        """安全地恢复Simtable单元格的公式。
        
        如果恢复失败，进行重试。
        """
        max_retries = 3
        for cell_addr, cell in self.simtable_cell_cache.items():
            original_formula = self.original_formulas.get(cell_addr)
            
            if original_formula:
                for retry in range(max_retries):
                    try:
                        cell.Formula = original_formula
                        break
                    except Exception as e:
                        print(f"恢复Simtable单元格公式失败 {cell_addr} (尝试{retry+1}/{max_retries}): {str(e)}")
                        time.sleep(0.1)
            else:
                # 如果没有公式，设置为空
                try:
                    cell.Value2 = ""
                except Exception as e:
                    print(f"清空Simtable单元格 {cell_addr} 失败: {str(e)}")
    def _prepare_results(self):
        """准备最终结果。
        
        截断数组到实际完成的迭代数，生成结果字典。
        """
        # 准备输入结果（支持每个单元格含多个样本数组的情况）
        input_results = {}

        for cell_addr, samples_list in self.input_samples.items():
            # samples_list 应为 list of np.ndarray
            if isinstance(samples_list, list) and samples_list:
                trimmed = []
                actual_iterations = self.performance_stats.get('iterations_completed', 0)
                for arr in samples_list:
                    if isinstance(arr, np.ndarray) and len(arr) > 0:
                        if actual_iterations < len(arr):
                            trimmed.append(arr[:actual_iterations])
                        else:
                            trimmed.append(arr)
                    else:
                        trimmed.append(np.array([], dtype=object))
                input_results[cell_addr] = trimmed
            else:
                # 兼容旧格式：单个 ndarray
                arr = samples_list if isinstance(samples_list, np.ndarray) else None
                if arr is not None and len(arr) > 0:
                    actual_iterations = self.performance_stats.get('iterations_completed', 0)
                    if actual_iterations < len(arr):
                        arr = arr[:actual_iterations]
                    input_results[cell_addr] = [arr]
        
        # 准备输出结果
        output_results = {}
        
        for cell_addr, data in self.output_cache.items():
            if isinstance(data, np.ndarray) and len(data) > 0:
                # 截断到实际完成的迭代数
                actual_iterations = self.performance_stats['iterations_completed']
                if actual_iterations < len(data):
                    data = data[:actual_iterations]
                output_results[cell_addr] = data
        
        return input_results, output_results, self.output_info
    
    def run_simulation(self, progress_callback=None, cancel_event=None):
        """运行完整的模拟过程。
        
        Args:
            progress_callback: 进度回调函数，接收已完成迭代数
            cancel_event: 取消事件对象
            
        Returns:
            (input_results, output_results, output_info) 元组，失败时返回 (None, None, None)
        """
        print(f"\n{'='*80}")
        print(f"开始模拟 - 场景 {self.scenario_index+1}/{self.scenario_count}")
        print(f"抽样方法: {self.sampling_method}")
        print(f"迭代次数: {self.n_iterations:,}")
        print(f"{'='*80}")
        
        total_start = time.time()
        
        try:
            # 1. 优化Excel设置
            self._optimize_excel_settings()
            
            # 2. 预缓存单元格
            self._precache_cells()
            
            # 3. 创建分布生成器（包括内嵌分布函数）
            self._create_distribution_generators()
            
            # 4. 预生成独立分布函数的随机数
            self._pre_generate_samples()
            
            # 5. 检查取消
            if cancel_event and cancel_event.is_set():
                print("模拟在开始前被取消")
                self.cancelled = True
                self._restore_excel_settings()
                return None, None, None
            
            # 6. 运行模拟
            success = self._run_simulation_iteratively(progress_callback, cancel_event)
            
            if not success:
                print("模拟失败")
                # 恢复Simtable单元格公式
                self._restore_simtable_formulas_safe()
                # 恢复MakeInput单元格公式
                self._restore_makeinput_formulas_safe()
                self._restore_input_formulas_safe()
                self._restore_excel_settings()
                return None, None, None
            
            # 7. 安全恢复公式
            # 先恢复Simtable单元格公式，再恢复MakeInput单元格公式，最后恢复输入单元格公式
            self._restore_simtable_formulas_safe()
            self._restore_makeinput_formulas_safe()
            self._restore_input_formulas_safe()
            
            # 8. 准备结果
            input_results, output_results, output_info = self._prepare_results()
            
            # 9. 性能统计
            total_time = time.time() - total_start
            self.performance_stats['total_time'] = total_time
            
            completed = self.performance_stats['iterations_completed']
            if completed > 0:
                speed = completed / self.performance_stats['simulation_time'] if self.performance_stats['simulation_time'] > 0 else 0
                print(f"\n模拟统计:")
                print(f"  - 完成迭代: {completed:,}")
                print(f"  - 总耗时: {total_time:.2f}秒")
                print(f"  - 模拟速度: {speed:.1f} 迭代/秒")
                print(f"  - 输入数量: {len(input_results)} (包括分布和MakeInput)")
                print(f"  - 输出数量: {len(output_results)}")
            
            return input_results, output_results, output_info
            
        except Exception as e:
            print(f"模拟失败: {str(e)}")
            traceback.print_exc()
            # 确保恢复公式和设置
            try:
                self._restore_simtable_formulas_safe()
                self._restore_makeinput_formulas_safe()
                self._restore_input_formulas_safe()
            except:
                pass
            return None, None, None
        
        finally:
            # 恢复Excel设置
            self._restore_excel_settings()
# ==================== 主接口函数 ====================

def iterative_simulation_workbook(app, distribution_cells: Dict[str, List[Dict]], 
                                 simtable_cells: Dict[str, List[Dict]],
                                 makeinput_cells: Dict[str, List[Dict]],
                                 output_cells: Dict[str, Dict], 
                                 n_iterations: int, 
                                 sampling_method: str = "MC",
                                 scenario_count: int = 1,
                                 scenario_index: int = 0,
                                 progress_callback=None,
                                 cancel_event=None):
    """
    迭代模拟引擎入口函数。
    
    Args:
        app: Excel 应用对象
        distribution_cells: 分布定义单元格字典
        simtable_cells: Simtable单元格字典
        makeinput_cells: MakeInput单元格字典
        output_cells: 输出单元格字典
        n_iterations: 迭代次数
        sampling_method: 抽样方法，默认为 'MC'
        scenario_count: 场景数
        scenario_index: 当前场景索引（0-based）
        progress_callback: 进度回调函数（可选）
        cancel_event: 取消事件对象（可选）
        
    Returns:
        (input_results, output_results, output_info) 元组
    """
    print(f"\n启动模拟引擎 - 场景 {scenario_index+1}/{scenario_count}")
    
    # 创建模拟引擎
    engine = IterativeSimulationEngine(
        app, distribution_cells, simtable_cells, makeinput_cells, output_cells, n_iterations, 
        sampling_method, scenario_count, scenario_index
    )
    
    # 运行模拟
    return engine.run_simulation(progress_callback, cancel_event)

# 提供向后兼容的接口
iterative_simulation_workbook_fast = iterative_simulation_workbook

# 清除缓存函数
def clear_dependency_cache():
    """清除依赖分析缓存"""
    with _dependency_cache_lock:
        _dependency_cache.clear()
    print("依赖分析缓存已清除")