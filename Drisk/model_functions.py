# model_functions.py - 仅保留和DriskModel有关的算法
"""模型分析函数模块 - DriskModel相关算法"""

import tkinter as tk
from tkinter import ttk, messagebox
import re
import logging
import sys
from typing import Dict, List, Tuple, Any, Optional
import numpy as np

# 导入公式解析器
from formula_parser import (
    extract_nested_distributions_advanced,
    parse_formula_references,
    parse_args_with_nested_functions,
    extract_dist_params_and_markers,
    extract_makeinput_functions,  # 新增导入
    extract_makeinput_attributes,  # 新增导入
    is_makeinput_function  # 新增导入
)

# 导入COM修复器
try:
    from com_fixer import _safe_excel_app
except ImportError:
    def _safe_excel_app():
        from pyxll import xl_app
        return xl_app()

# 导入日志模块
logger = logging.getLogger(__name__)
if not logger.handlers:
    handler = logging.StreamHandler(sys.stdout)
    handler.setLevel(logging.WARNING)
    formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
    handler.setFormatter(formatter)
    logger.addHandler(handler)
    logger.setLevel(logging.WARNING)

# ==================== 稳定键生成器（简化版） ====================

class ModelStableKeyGenerator:
    """模型稳定键生成器 - 用于DriskModel"""
    
    def __init__(self):
        self.generated_keys = set()
        self.nested_relationships = {}
        
    def _generate_param_hash(self, params: List[float], args_text: str = None) -> int:
        """生成参数哈希"""
        if args_text:
            normalized_args = args_text.strip().upper()
            normalized_args = re.sub(r'\s+', '', normalized_args)
            return abs(hash(normalized_args)) % 1000000
        
        if not params:
            return 0
        
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
        return abs(hash(param_str_combined)) % 1000000

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
            
            param_hash = self._generate_param_hash(params, args_text)
            
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
            
            if parent_key:
                if parent_key not in self.nested_relationships:
                    self.nested_relationships[parent_key] = []
                if stable_key not in self.nested_relationships[parent_key]:
                    self.nested_relationships[parent_key].append(stable_key)
            
            return stable_key
            
        except Exception as e:
            import random
            fallback_key = f"{cell_address.replace('!', '_')}_{func_name}_{index_in_cell}_{random.randint(10000, 99999)}"
            return fallback_key.upper()

# 全局稳定键生成器实例
_model_stable_key_generator = ModelStableKeyGenerator()

def get_model_stable_key(cell_address: str, func_name: str, params: List[float],
                        is_at_function: bool = False, is_nested: bool = False,
                        index_in_cell: int = 1, depth: int = 0, 
                        parent_key: Optional[str] = None, args_text: str = None) -> str:
    """获取模型稳定键的公共接口"""
    return _model_stable_key_generator.generate_stable_key(
        cell_address, func_name, params, is_at_function, 
        is_nested, index_in_cell, depth, parent_key, args_text
    )

# ==================== 分布函数模型分析器 ====================

class DistributionModelAnalyzer:
    """分布函数模型分析器"""
    
    def __init__(self, app):
        self.app = app
        self.workbook = app.ActiveWorkbook
        self.distribution_info = []
        self.output_info = []
        self.statistical_info = []  # 新增：统计函数信息
        self.cell_info_cache = {}
        
    def analyze_workbook(self):
        """分析整个工作簿中的分布函数、输出函数和统计函数"""
        self.distribution_info = []
        self.output_info = []
        self.statistical_info = []  # 清空统计函数信息
        self.cell_info_cache = {}
        
        logger.info("开始分析工作簿中的分布函数、输出函数和统计函数...")
        
        from attribute_functions import get_static_mode, set_static_mode
        original_mode = get_static_mode()
        if not original_mode:
            set_static_mode(True)
        
        try:
            for sheet in self.workbook.Worksheets:
                try:
                    sheet_name = sheet.Name
                    
                    used_range = sheet.UsedRange
                    if used_range is None:
                        continue
                    
                    for row in range(1, used_range.Rows.Count + 1):
                        for col in range(1, used_range.Columns.Count + 1):
                            try:
                                cell = used_range.Cells(row, col)
                                cell_address = self._get_cell_address(cell)
                                full_address = f"{sheet_name}!{cell_address}"
                                
                                formula = cell.Formula
                                if not isinstance(formula, str) or not formula.startswith('='):
                                    continue
                                
                                self.cell_info_cache[full_address] = {
                                    'sheet': sheet_name,
                                    'address': cell_address,
                                    'formula': formula,
                                    'value': cell.Value
                                }
                                
                                # 分析分布函数和MakeInput函数
                                self._analyze_input_functions(full_address, formula, sheet_name)
                                
                                # 分析输出函数 - 传递单元格对象以便获取名称
                                self._analyze_output_functions(full_address, formula, sheet_name, cell)
                                
                                # 新增：分析统计函数
                                self._analyze_statistical_functions(full_address, formula, sheet_name, cell)
                                
                            except Exception:
                                continue
                                
                except Exception as e:
                    continue
        
            logger.info(f"分析完成，共找到 {len(self.distribution_info)} 个输入函数，{len(self.output_info)} 个输出函数，{len(self.statistical_info)} 个统计函数")
            return self.distribution_info, self.output_info, self.statistical_info
            
        finally:
            set_static_mode(original_mode)
    
    def _get_cell_address(self, cell):
        """获取单元格地址"""
        try:
            return cell.Address.replace('$', '')
        except:
            try:
                return cell.AddressLocal.replace('$', '')
            except:
                return f"R{cell.Row}C{cell.Column}"
    
    def _analyze_input_functions(self, cell_address: str, formula: str, sheet_name: str):
        """分析单元格公式中的输入函数（包括分布函数和MakeInput函数）"""
        try:
            if not formula.startswith('='):
                return
            
            # 检查是否是MakeInput函数
            if is_makeinput_function(formula):
                self._analyze_makeinput_function(cell_address, formula, sheet_name)
            
            # 分析分布函数（包括可能嵌套在MakeInput中的分布函数）
            all_functions = extract_nested_distributions_advanced(formula, cell_address)
            compound_indices = {
                func.get('index')
                for func in all_functions
                if str(func.get('func_name', '')).lower() == 'driskcompound'
            }
            if compound_indices:
                all_functions = [
                    func for func in all_functions
                    if not any(parent in compound_indices for parent in (func.get('parent_indices', []) or []))
                ]
            
            if all_functions:
                for func_info in all_functions:
                    self._add_distribution_info(cell_address, formula, func_info, sheet_name)
                    
        except Exception as e:
            pass
    
    def _analyze_makeinput_function(self, cell_address: str, formula: str, sheet_name: str):
        """分析MakeInput函数 - 修复版：获取完整的MakeInput公式"""
        try:
            # 提取MakeInput函数信息
            makeinput_funcs = extract_makeinput_functions(formula)
            
            if not makeinput_funcs:
                return
            
            for func_info in makeinput_funcs:
                # 创建稳定键
                stable_key = get_model_stable_key(
                    cell_address=cell_address,
                    func_name='DriskMakeInput',
                    params=[],
                    is_at_function=False,
                    is_nested=False,
                    index_in_cell=1,
                    depth=0,
                    parent_key=None,
                    args_text=func_info.get('args_text', '')
                )
                
                # 获取属性信息
                attributes = func_info.get('attributes', {})
                formula_arg = func_info.get('formula', '')
                
                # 关键修改：使用完整的MakeInput公式，而不是公式参数
                # 完整的MakeInput公式应该包含整个函数调用，例如：=DriskMakeInput(DriskNormal(4,2)+C1, DriskName("输入"))
                # 而不是只提取 DriskNormal(4,2)+C1
                full_makeinput_formula = func_info.get('full_match', '')
                if full_makeinput_formula and full_makeinput_formula.startswith('DriskMakeInput'):
                    # 确保有等号
                    makeinput_expression = f"={full_makeinput_formula}"
                else:
                    # 回退方案：使用完整公式
                    makeinput_expression = formula
                
                # 构建信息字典
                info = {
                    'stable_key': stable_key,
                    'sheet': sheet_name,
                    'cell_address': cell_address,
                    'full_formula': formula,
                    'function_snippet': func_info.get('full_match', ''),
                    'function_name': 'DriskMakeInput',
                    'is_at_function': False,
                    'args_text': func_info.get('args_text', ''),
                    'parameters': [],
                    'markers': {},
                    'cell_references': [],
                    'is_nested': False,
                    'nested_depth': 0,
                    'nested_path': [],
                    'parent_indices': [],
                    'index_in_cell': 1,
                    'start_position': 0,
                    'end_position': 0,
                    'distribution_type': 'MakeInput',
                    'makeinput_attributes': attributes,
                    'expression': makeinput_expression  # 修改：使用完整的MakeInput表达式
                }
                
                # 添加属性信息
                if 'name' in attributes:
                    info['name'] = attributes['name']
                if 'category' in attributes:
                    info['category'] = attributes['category']
                if 'units' in attributes:
                    info['units'] = attributes['units']
                
                self.distribution_info.append(info)
                
                # 提取MakeInput中的内嵌分布函数
                nested_dists = func_info.get('nested_distributions', [])
                for i, nested_dist in enumerate(nested_dists):
                    # 为内嵌分布函数创建信息
                    nested_func_name = nested_dist.get('func_name', '')
                    nested_args_text = nested_dist.get('args_text', '')
                    
                    nested_stable_key = get_model_stable_key(
                        cell_address=cell_address,
                        func_name=nested_func_name,
                        params=nested_dist.get('parameters', []),
                        is_at_function=False,
                        is_nested=True,
                        index_in_cell=i+1,
                        depth=nested_dist.get('depth', 0),
                        parent_key=stable_key,
                        args_text=nested_args_text
                    )
                    
                    nested_info = {
                        'stable_key': nested_stable_key,
                        'sheet': sheet_name,
                        'cell_address': cell_address,
                        'full_formula': formula,
                        'function_snippet': nested_dist.get('full_match', ''),
                        'function_name': nested_func_name,
                        'is_at_function': False,
                        'is_nested': True,
                        'nested_depth': nested_dist.get('depth', 0),
                        'args_text': nested_args_text,
                        'parameters': nested_dist.get('parameters', []),
                        'markers': nested_dist.get('markers', {}),
                        'cell_references': [],
                        'nested_path': [],
                        'parent_indices': [],
                        'index_in_cell': i+1,
                        'start_position': 0,
                        'end_position': 0,
                        'distribution_type': self._get_distribution_type(nested_func_name),
                        'makeinput_attributes': attributes,
                        'parent_makeinput': cell_address,
                        'expression': makeinput_expression  # 修改：使用完整的MakeInput表达式
                    }
                    
                    self.distribution_info.append(nested_info)
                
        except Exception as e:
            pass
    
    def _analyze_output_functions(self, cell_address: str, formula: str, sheet_name: str, cell):
        """分析输出函数"""
        try:
            if not formula.startswith('='):
                return
            
            # 检查是否是输出函数
            if not ('DriskOutput' in formula.upper() or 'DRISKOUTPUT' in formula.upper()):
                return
            
            # 方法1: 尝试使用extract_output_info
            from formula_parser import extract_output_info
            output_info = extract_output_info(formula)
            
            # 方法2: 如果extract_output_info返回空，使用正则表达式直接提取
            if not output_info:
                output_info = self._extract_output_info_by_regex(formula)
            
            if not output_info:
                return
            
            # 如果输出名称为空，尝试从相邻单元格获取名称
            output_name = output_info.get('name', '')
            if not output_name:
                output_name = self._find_cell_name(cell)
            
            # 构建输出信息字典
            info = {
                'sheet': sheet_name,
                'cell_address': cell_address,
                'full_formula': formula,
                'output_name': output_name,
                'output_category': output_info.get('category', ''),
                'output_position': output_info.get('position', 1)
            }
            
            self.output_info.append(info)
                
        except Exception as e:
            logger.error(f"分析输出函数失败: {str(e)}")
            pass
    
    def _analyze_statistical_functions(self, cell_address: str, formula: str, sheet_name: str, cell):
        """分析统计函数（包括理论统计函数和模拟统计函数）"""
        try:
            if not formula.startswith('='):
                return
            
            # 统计函数列表（包括理论统计函数和模拟统计函数）
            statistical_functions = [
                # 理论统计函数
                'DriskTheoMean', 'DriskTheoStdDev', 'DriskTheoVariance', 'DriskTheoSkewness',
                'DriskTheoKurtosis', 'DriskTheoMin', 'DriskTheoMax', 'DriskTheoRange',
                'DriskTheoMode', 'DriskTheoPtoX', 'DriskTheoXtoP', 'DriskTheoXtoY',
                # 模拟统计函数
                'DriskMean', 'DriskStd', 'DriskVariance', 'DriskMin', 'DriskMax',
                'DriskMedian', 'DriskPtoX', 'DriskData', 'DriskSkew', 'DriskKurt',
                'DriskRange', 'DriskMode', 'DriskMeanAbsDev', 'DriskXtoP',
                'DriskCIMean', 'DriskCoeffOfVariation', 'DriskSemiStdDev',
                'DriskSemiVariance', 'DriskStdErrOfMean', 'DriskCIPercentile'
            ]
            
            # 检查公式中是否包含任何统计函数
            is_statistical = False
            func_name = ""
            
            for func in statistical_functions:
                if func.upper() in formula.upper():
                    # 进一步检查是否是函数调用（后面跟着括号）
                    func_pattern = re.compile(re.escape(func) + r'\s*\(', re.IGNORECASE)
                    if func_pattern.search(formula):
                        is_statistical = True
                        func_name = func
                        break
            
            if not is_statistical:
                return
            
            # 获取单元格值（计算值）
            cell_value = cell.Value
            
            # 处理计算值
            calculated_value = self._process_cell_value(cell_value)
            
            # 构建统计函数信息字典
            info = {
                'sheet': sheet_name,
                'cell_address': cell_address,
                'full_formula': formula,
                'function_name': func_name,
                'calculated_value': calculated_value,
                'cell_value': cell_value
            }
            
            self.statistical_info.append(info)
                
        except Exception as e:
            logger.error(f"分析统计函数失败: {str(e)}")
            pass
    
    def _process_cell_value(self, cell_value):
        """处理单元格值，返回适当的表示形式"""
        try:
            if cell_value is None:
                return "#ERROR!"
            
            # 检查是否是错误值
            if isinstance(cell_value, str) and cell_value.startswith('#'):
                return "#ERROR!"
            
            # 检查是否是数字
            if isinstance(cell_value, (int, float, np.number)):
                # 检查是否是NaN或无穷大
                if isinstance(cell_value, float):
                    if np.isnan(cell_value) or np.isinf(cell_value):
                        return "#ERROR!"
                
                # 格式化数值
                try:
                    # 尝试保留合理的小数位数
                    if isinstance(cell_value, int) or cell_value.is_integer():
                        return str(int(cell_value))
                    else:
                        # 保留6位小数
                        return f"{cell_value:.6f}".rstrip('0').rstrip('.')
                except:
                    return str(cell_value)
            
            # 其他类型转换为字符串
            return str(cell_value)
            
        except Exception as e:
            return "#ERROR!"
    
    def _extract_output_info_by_regex(self, formula: str) -> Dict[str, Any]:
        """使用正则表达式直接提取输出函数信息"""
        try:
            # 匹配 DriskOutput 函数
            pattern = r'DRISKOUTPUT\s*\(\s*["\']?([^"\',]+)["\']?\s*,\s*["\']?([^"\',]+)["\']?\s*(?:,\s*(\d+)\s*)?\)'
            
            match = re.search(pattern, formula.upper())
            if match:
                name = match.group(1).strip().strip('"\'')
                category = match.group(2).strip().strip('"\'')
                position = int(match.group(3)) if match.group(3) else 1
                
                return {
                    'name': name,
                    'category': category,
                    'position': position
                }
            
            # 如果没有找到，尝试匹配其他可能的输出函数格式
            alt_pattern = r'Output\s*\(\s*["\']?([^"\',]+)["\']?\s*,\s*["\']?([^"\',]+)["\']?\s*(?:,\s*(\d+)\s*)?\)'
            match = re.search(alt_pattern, formula.upper())
            if match:
                name = match.group(1).strip().strip('"\'')
                category = match.group(2).strip().strip('"\'')
                position = int(match.group(3)) if match.group(3) else 1
                
                return {
                    'name': name,
                    'category': category,
                    'position': position
                }
            
            return {}
        except Exception as e:
            logger.error(f"正则表达式提取输出信息失败: {str(e)}")
            return {}
    
    def _find_cell_name(self, cell) -> str:
        """为单元格自动命名：向上查找第一个非空字符串，向左查找第一个非空字符串"""
        try:
            # 获取单元格的行列
            row = cell.Row
            col = cell.Column
            sheet = cell.Worksheet
            
            # 向上查找（最多10行）
            up_name = ""
            for i in range(1, 11):
                if row - i < 1:
                    break
                up_cell = sheet.Cells(row - i, col)
                val = up_cell.Value
                if val is not None and isinstance(val, str) and val.strip():
                    up_name = str(val).strip()
                    break
            
            # 向左查找（最多10列）
            left_name = ""
            for i in range(1, 11):
                if col - i < 1:
                    break
                left_cell = sheet.Cells(row, col - i)
                val = left_cell.Value
                if val is not None and isinstance(val, str) and val.strip():
                    left_name = str(val).strip()
                    break
            
            # 组合名称
            if up_name and left_name:
                return f"{up_name}_{left_name}"
            elif up_name:
                return up_name
            elif left_name:
                return left_name
            else:
                return f"Output_{row}_{col}"
        except Exception as e:
            return ""
    
    def _add_distribution_info(self, cell_address: str, formula: str, 
                              func_info: Dict[str, Any], sheet_name: str):
        """添加分布函数信息到列表"""
        try:
            func_name = func_info.get('func_name', '')
            args_text = func_info.get('args_text', '')
            full_match = func_info.get('full_match', '')
            is_at_function = func_info.get('is_at_function', False)
            is_nested = func_info.get('is_nested', False)
            start_pos = func_info.get('start_pos', 0)
            end_pos = func_info.get('end_pos', 0)
            depth = func_info.get('depth', 0)
            nested_path = func_info.get('nested_path', [])
            parent_indices = func_info.get('parent_indices', [])
            
            # 检查嵌套深度
            if depth > 2:
                logger.warning(f"警告: 函数 {func_name} 嵌套深度为 {depth}，超过两层嵌套限制！")
            
            # 生成稳定键
            stable_key = get_model_stable_key(
                cell_address=cell_address,
                func_name=func_name,
                params=func_info.get('parameters', []),
                is_at_function=is_at_function,
                is_nested=is_nested,
                index_in_cell=func_info.get('index', 1),
                depth=depth,
                parent_key=None,
                args_text=args_text
            )
            
            # 提取单元格引用
            cell_refs = parse_formula_references(f"={full_match}")
            
            # 确定分布类型
            distribution_type = self._get_distribution_type(func_name)
            
            # 构建信息字典
            info = {
                'stable_key': stable_key,
                'sheet': sheet_name,
                'cell_address': cell_address,
                'full_formula': formula,
                'function_snippet': full_match,
                'function_name': func_name,
                'is_at_function': is_at_function,
                'args_text': args_text,
                'parameters': func_info.get('parameters', []),
                'markers': func_info.get('markers', {}),
                'cell_references': cell_refs,
                'is_nested': is_nested,
                'nested_depth': depth,
                'nested_path': nested_path,
                'parent_indices': parent_indices,
                'index_in_cell': func_info.get('index', 1),
                'start_position': start_pos,
                'end_position': end_pos,
                'distribution_type': distribution_type
            }
            
            self.distribution_info.append(info)
            
        except Exception as e:
            pass
    
    def _get_distribution_type(self, func_name: str) -> str:
        """获取分布类型 - 更新版，支持所有新分布"""
        func_name_lower = func_name.lower()
        
        # 按顺序检查，更长的模式优先
        if 'makeinput' in func_name_lower:
            return 'MakeInput'
        elif 'simtable' in func_name_lower:
            return 'Simtable'
        elif 'trigen' in func_name_lower:
            return '三参数三角分布'
        elif 'doubletriang' in func_name_lower:
            return '\u53cc\u4e09\u89d2\u5206\u5e03'
        elif 'triang' in func_name_lower:
            return '三角分布'
        elif 'binomial' in func_name_lower:
            return '二项分布'
        elif 'negbin' in func_name_lower:
            return '负二项分布'
        elif 'invgauss' in func_name_lower:
            return '逆高斯分布'
        elif 'extvaluemin' in func_name_lower:
            return 'ExtvalueMin'
        elif 'extvalue' in func_name_lower:
            return 'Extvalue'
        elif 'fatiguelife' in func_name_lower:
            return 'FatigueLife'
        elif 'frechet' in func_name_lower:
            return 'Frechet'
        elif 'general' in func_name_lower:
            return 'General'
        elif 'histogrm' in func_name_lower:
            return 'Histogrm'
        elif 'hypsecant' in func_name_lower:
            return 'HypSecant'
        elif 'johnsonsb' in func_name_lower:
            return 'JohnsonSB'
        elif 'johnsonsu' in func_name_lower:
            return 'JohnsonSU'
        elif 'kumaraswamy' in func_name_lower:
            return 'Kumaraswamy'
        elif 'laplace' in func_name_lower:
            return 'Laplace'
        elif 'loglogistic' in func_name_lower:
            return 'Loglogistic'
        elif 'lognorm2' in func_name_lower:
            return 'Lognorm2'
        elif 'betageneral' in func_name_lower:
            return 'BetaGeneral'
        elif 'betasubj' in func_name_lower:
            return 'BetaSubj'
        elif 'burr12' in func_name_lower:
            return 'Burr12'
        elif 'compound' in func_name_lower:
            return 'Compound'
        elif 'splice' in func_name_lower:
            return 'Splice'
        elif 'pert' in func_name_lower:
            return 'Pert'
        elif 'reciprocal' in func_name_lower:
            return 'Reciprocal'
        elif 'rayleigh' in func_name_lower:
            return 'Rayleigh'
        elif 'weibull' in func_name_lower:
            return 'Weibull'
        elif 'pearson5' in func_name_lower:
            return 'Pearson5'
        elif 'pearson6' in func_name_lower:
            return 'Pearson6'
        elif 'pareto2' in func_name_lower:
            return 'Pareto2'
        elif 'pareto' in func_name_lower:
            return 'Pareto'
        elif 'lognorm' in func_name_lower:
            return 'Lognorm'
        elif 'logistic' in func_name_lower:
            return 'Logistic'
        elif 'levy' in func_name_lower:
            return 'Levy'
        elif 'duniform' in func_name_lower:
            return '离散均匀分布'
        elif 'bernoulli' in func_name_lower:
            return '伯努利分布'
        elif 'geomet' in func_name_lower:
            return '几何分布'
        elif 'hypergeo' in func_name_lower:
            return '超几何分布'
        elif 'intuniform' in func_name_lower:
            return '整数均匀分布'
        elif 'dagum' in func_name_lower:
            return '\u8fbe\u683c\u59c6\u5206\u5e03'
        elif 'cumul' in func_name_lower:
            return '累积分布'
        elif 'discrete' in func_name_lower:
            return '离散分布'
        elif 'normal' in func_name_lower:
            return '正态分布'
        elif 'uniform' in func_name_lower:
            return '均匀分布'
        elif 'cauchy' in func_name_lower:
            return '柯西分布'
        elif 'erf' in func_name_lower:
            return '误差函数分布'
        elif 'gamma' in func_name_lower:
            return 'Gamma分布'
        elif 'poisson' in func_name_lower:
            return '泊松分布'
        elif 'beta' in func_name_lower:
            return 'Beta分布'
        elif 'chisq' in func_name_lower:
            return '卡方分布'
        elif 'f' in func_name_lower:
            return 'F分布'
        elif 'student' in func_name_lower:
            return 'student分布'
        elif 'expon' in func_name_lower:
            return '指数分布'
        else:
            return '未知分布'

# ==================== 模型信息窗口（三表格版） ====================

class ModelInfoWindow:
    """模型信息窗口 - 三表格版"""
    
    def __init__(self, distribution_info: List[Dict], output_info: List[Dict], statistical_info: List[Dict]):
        self.distribution_info = distribution_info
        self.output_info = output_info
        self.statistical_info = statistical_info
        self.root = tk.Tk()
        self.root.title("Drisk - 模型分析")
        self.root.geometry("1400x900")
        
        # 创建主框架
        main_frame = ttk.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # 添加标题
        title_label = ttk.Label(main_frame, text="Drisk 模型分析", 
                               font=("Arial", 14, "bold"))
        title_label.pack(pady=10)
        
        # 创建Notebook（选项卡）
        notebook = ttk.Notebook(main_frame)
        notebook.pack(fill=tk.BOTH, expand=True, pady=10)
        
        # 创建Input信息框架
        input_frame = ttk.Frame(notebook)
        notebook.add(input_frame, text="Input信息")
        
        # 创建Output信息框架
        output_frame = ttk.Frame(notebook)
        notebook.add(output_frame, text="Output信息")
        
        # 创建Statistical信息框架
        statistical_frame = ttk.Frame(notebook)
        notebook.add(statistical_frame, text="统计函数信息")
        
        # ========== Input信息表格 ==========
        
        # 添加Input标题
        input_title = ttk.Label(input_frame, text="输入函数信息（包括分布函数和MakeInput）", 
                               font=("Arial", 12, "bold"))
        input_title.pack(pady=10)
        
        # 创建Input Treeview
        input_tree_frame = ttk.Frame(input_frame)
        input_tree_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        # 定义Input列
        input_columns = [
            "工作表", "单元格地址", "Input公式", "稳定键", "分布类型"
        ]
        
        self.input_tree = ttk.Treeview(input_tree_frame, columns=input_columns, show='headings', height=15)
        
        # 设置Input列标题
        for col in input_columns:
            self.input_tree.heading(col, text=col)
            if col == "Input公式":
                self.input_tree.column(col, width=300)
            elif col == "稳定键":
                self.input_tree.column(col, width=200)
            else:
                self.input_tree.column(col, width=100)
        
        # 添加Input滚动条
        input_vsb = ttk.Scrollbar(input_tree_frame, orient="vertical", command=self.input_tree.yview)
        input_hsb = ttk.Scrollbar(input_tree_frame, orient="horizontal", command=self.input_tree.xview)
        self.input_tree.configure(yscrollcommand=input_vsb.set, xscrollcommand=input_hsb.set)
        
        # Input布局
        self.input_tree.grid(row=0, column=0, sticky="nsew")
        input_vsb.grid(row=0, column=1, sticky="ns")
        input_hsb.grid(row=1, column=0, sticky="ew")
        
        # 配置Input网格权重
        input_tree_frame.grid_rowconfigure(0, weight=1)
        input_tree_frame.grid_columnconfigure(0, weight=1)
        
        # 填充Input数据
        self._populate_input_tree()
        
        # Input统计信息
        input_stats_frame = ttk.Frame(input_frame)
        input_stats_frame.pack(fill=tk.X, pady=5)
        
        # 计算Input统计
        total_input_count = len(distribution_info)
        at_function_count = sum(1 for info in distribution_info if info.get('is_at_function', False))
        nested_count = sum(1 for info in distribution_info if info.get('is_nested', False) and not info.get('is_at_function', False))
        normal_count = total_input_count - at_function_count - nested_count
        makeinput_count = sum(1 for info in distribution_info if info.get('distribution_type') == 'MakeInput')
        
        ttk.Label(input_stats_frame, text=f"总输入函数数: {total_input_count}").pack(side=tk.LEFT, padx=10)
        ttk.Label(input_stats_frame, text=f"@函数数量: {at_function_count}").pack(side=tk.LEFT, padx=10)
        ttk.Label(input_stats_frame, text=f"嵌套函数数量: {nested_count}").pack(side=tk.LEFT, padx=10)
        ttk.Label(input_stats_frame, text=f"普通函数数量: {normal_count}").pack(side=tk.LEFT, padx=10)
        ttk.Label(input_stats_frame, text=f"MakeInput数量: {makeinput_count}").pack(side=tk.LEFT, padx=10)
        
        # ========== Output信息表格 ==========
        
        # 添加Output标题
        output_title = ttk.Label(output_frame, text="输出函数信息", 
                                font=("Arial", 12, "bold"))
        output_title.pack(pady=10)
        
        # 创建Output Treeview
        output_tree_frame = ttk.Frame(output_frame)
        output_tree_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        # 定义Output列
        output_columns = [
            "工作表", "单元格地址", "所在单元格公式", "输出名称", "输出类别", "输出位置"
        ]
        
        self.output_tree = ttk.Treeview(output_tree_frame, columns=output_columns, show='headings', height=15)
        
        # 设置Output列标题
        for col in output_columns:
            self.output_tree.heading(col, text=col)
            if col == "所在单元格公式":
                self.output_tree.column(col, width=400)
            else:
                self.output_tree.column(col, width=100)
        
        # 添加Output滚动条
        output_vsb = ttk.Scrollbar(output_tree_frame, orient="vertical", command=self.output_tree.yview)
        output_hsb = ttk.Scrollbar(output_tree_frame, orient="horizontal", command=self.output_tree.xview)
        self.output_tree.configure(yscrollcommand=output_vsb.set, xscrollcommand=output_hsb.set)
        
        # Output布局
        self.output_tree.grid(row=0, column=0, sticky="nsew")
        output_vsb.grid(row=0, column=1, sticky="ns")
        output_hsb.grid(row=1, column=0, sticky="ew")
        
        # 配置Output网格权重
        output_tree_frame.grid_rowconfigure(0, weight=1)
        output_tree_frame.grid_columnconfigure(0, weight=1)
        
        # 填充Output数据
        self._populate_output_tree()
        
        # Output统计信息
        output_stats_frame = ttk.Frame(output_frame)
        output_stats_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(output_stats_frame, text=f"总输出函数数: {len(output_info)}").pack(side=tk.LEFT, padx=10)
        
        # ========== 统计函数信息表格 ==========
        
        # 添加Statistical标题
        statistical_title = ttk.Label(statistical_frame, text="统计函数信息（包括理论统计函数和模拟统计函数）", 
                                     font=("Arial", 12, "bold"))
        statistical_title.pack(pady=10)
        
        # 创建Statistical Treeview
        statistical_tree_frame = ttk.Frame(statistical_frame)
        statistical_tree_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        # 定义Statistical列
        statistical_columns = [
            "工作表", "单元格地址", "统计函数公式", "函数名称", "计算值"
        ]
        
        self.statistical_tree = ttk.Treeview(statistical_tree_frame, columns=statistical_columns, show='headings', height=15)
        
        # 设置Statistical列标题
        for col in statistical_columns:
            self.statistical_tree.heading(col, text=col)
            if col == "统计函数公式":
                self.statistical_tree.column(col, width=400)
            elif col == "计算值":
                self.statistical_tree.column(col, width=150)
            else:
                self.statistical_tree.column(col, width=100)
        
        # 添加Statistical滚动条
        statistical_vsb = ttk.Scrollbar(statistical_tree_frame, orient="vertical", command=self.statistical_tree.yview)
        statistical_hsb = ttk.Scrollbar(statistical_tree_frame, orient="horizontal", command=self.statistical_tree.xview)
        self.statistical_tree.configure(yscrollcommand=statistical_vsb.set, xscrollcommand=statistical_hsb.set)
        
        # Statistical布局
        self.statistical_tree.grid(row=0, column=0, sticky="nsew")
        statistical_vsb.grid(row=0, column=1, sticky="ns")
        statistical_hsb.grid(row=1, column=0, sticky="ew")
        
        # 配置Statistical网格权重
        statistical_tree_frame.grid_rowconfigure(0, weight=1)
        statistical_tree_frame.grid_columnconfigure(0, weight=1)
        
        # 填充Statistical数据
        self._populate_statistical_tree()
        
        # Statistical统计信息
        statistical_stats_frame = ttk.Frame(statistical_frame)
        statistical_stats_frame.pack(fill=tk.X, pady=5)
        
        # 统计理论统计函数和模拟统计函数的数量
        theo_count = sum(1 for info in statistical_info if 'Theo' in info.get('function_name', ''))
        sim_count = len(statistical_info) - theo_count
        
        ttk.Label(statistical_stats_frame, text=f"总统计函数数: {len(statistical_info)}").pack(side=tk.LEFT, padx=10)
        ttk.Label(statistical_stats_frame, text=f"理论统计函数: {theo_count}").pack(side=tk.LEFT, padx=10)
        ttk.Label(statistical_stats_frame, text=f"模拟统计函数: {sim_count}").pack(side=tk.LEFT, padx=10)
        
        # 添加按钮框架
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(pady=10)
        
        ttk.Button(button_frame, text="导出到Excel", command=self.export_to_excel).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="关闭", command=self.root.destroy).pack(side=tk.LEFT, padx=5)
        
        # 添加状态栏
        status_frame = ttk.Frame(main_frame)
        status_frame.pack(fill=tk.X, pady=5)
        
        self.status_label = ttk.Label(status_frame, text=f"共 {len(distribution_info)} 个输入函数，{len(output_info)} 个输出函数，{len(statistical_info)} 个统计函数")
        self.status_label.pack(side=tk.LEFT)
    
    def _populate_input_tree(self):
        """填充Input Treeview数据"""
        for info in self.distribution_info:
            # 提取Input公式（截断过长的公式，并去掉等号）
            input_formula = info.get('function_snippet', '')
            
            # 对于MakeInput函数，使用完整的表达式
            if info.get('distribution_type') == 'MakeInput':
                expression = info.get('expression', '')
                if expression:
                    input_formula = expression
            
            # 去掉等号
            if input_formula.startswith('='):
                input_formula = input_formula[1:]
            
            if len(input_formula) > 80:
                input_formula = input_formula[:77] + "..."
            
            row_data = [
                info.get('sheet', ''),
                info.get('cell_address', ''),
                input_formula,
                info.get('stable_key', ''),
                info.get('distribution_type', '')
            ]
            
            self.input_tree.insert("", tk.END, values=row_data)
    
    def _populate_output_tree(self):
        """填充Output Treeview数据"""
        for info in self.output_info:
            # 提取单元格公式（截断过长的公式，并去掉等号）
            cell_formula = info.get('full_formula', '')
            # 去掉等号
            if cell_formula.startswith('='):
                cell_formula = cell_formula[1:]
            
            if len(cell_formula) > 100:
                cell_formula = cell_formula[:97] + "..."
            
            row_data = [
                info.get('sheet', ''),
                info.get('cell_address', ''),
                cell_formula,
                info.get('output_name', ''),
                info.get('output_category', ''),
                str(info.get('output_position', 1))
            ]
            
            self.output_tree.insert("", tk.END, values=row_data)
    
    def _populate_statistical_tree(self):
        """填充Statistical Treeview数据"""
        for info in self.statistical_info:
            # 提取统计函数公式（截断过长的公式，并去掉等号）
            cell_formula = info.get('full_formula', '')
            # 去掉等号
            if cell_formula.startswith('='):
                cell_formula = cell_formula[1:]
            
            if len(cell_formula) > 100:
                cell_formula = cell_formula[:97] + "..."
            
            # 获取计算值
            calculated_value = info.get('calculated_value', '#ERROR!')
            
            row_data = [
                info.get('sheet', ''),
                info.get('cell_address', ''),
                cell_formula,
                info.get('function_name', ''),
                calculated_value
            ]
            
            self.statistical_tree.insert("", tk.END, values=row_data)
    
    def export_to_excel(self):
        """导出到Excel"""
        try:
            app = _safe_excel_app()
            workbook = app.ActiveWorkbook
            
            # 创建Input工作表
            try:
                input_sheet = workbook.Worksheets("DriskModel_Input")
                input_sheet.Delete()
            except:
                pass
            
            input_sheet = workbook.Worksheets.Add()
            input_sheet.Name = "DriskModel_Input"
            
            # 写入Input表头
            input_headers = [
                "工作表", "单元格地址", "Input公式", "稳定键", "分布类型", "函数名", "参数", "是否嵌套", "嵌套深度"
            ]
            
            for col_idx, header in enumerate(input_headers, start=1):
                cell = input_sheet.Cells(1, col_idx)
                cell.Value = header
                cell.Font.Bold = True
                cell.Interior.Color = 0xCCCCFF
            
            # 写入Input数据
            for row_idx, info in enumerate(self.distribution_info, start=2):
                input_sheet.Cells(row_idx, 1).Value = info.get('sheet', '')
                input_sheet.Cells(row_idx, 2).Value = info.get('cell_address', '')
                
                # Input公式（去掉等号）
                input_formula = info.get('function_snippet', '')
                if info.get('distribution_type') == 'MakeInput':
                    expression = info.get('expression', '')
                    if expression:
                        input_formula = expression
                # 去掉等号
                if input_formula.startswith('='):
                    input_formula = input_formula[1:]
                input_sheet.Cells(row_idx, 3).Value = input_formula
                
                input_sheet.Cells(row_idx, 4).Value = info.get('stable_key', '')
                input_sheet.Cells(row_idx, 5).Value = info.get('distribution_type', '')
                input_sheet.Cells(row_idx, 6).Value = info.get('function_name', '')
                
                # 参数
                params = info.get('parameters', [])
                if params:
                    param_str = ', '.join([str(p) for p in params])
                else:
                    param_str = info.get('args_text', '')
                input_sheet.Cells(row_idx, 7).Value = param_str
                
                input_sheet.Cells(row_idx, 8).Value = info.get('is_nested', False)
                input_sheet.Cells(row_idx, 9).Value = info.get('nested_depth', 0)
            
            # 设置Input列宽
            for i, width in enumerate([15, 15, 40, 30, 15, 15, 20, 10, 10]):
                input_sheet.Columns(i+1).ColumnWidth = width
            
            # 创建Output工作表
            try:
                output_sheet = workbook.Worksheets("DriskModel_Output")
                output_sheet.Delete()
            except:
                pass
            
            output_sheet = workbook.Worksheets.Add()
            output_sheet.Name = "DriskModel_Output"
            
            # 写入Output表头
            output_headers = [
                "工作表", "单元格地址", "所在单元格公式", "输出名称", "输出类别", "输出位置"
            ]
            
            for col_idx, header in enumerate(output_headers, start=1):
                cell = output_sheet.Cells(1, col_idx)
                cell.Value = header
                cell.Font.Bold = True
                cell.Interior.Color = 0xCCFFCC
            
            # 写入Output数据
            for row_idx, info in enumerate(self.output_info, start=2):
                output_sheet.Cells(row_idx, 1).Value = info.get('sheet', '')
                output_sheet.Cells(row_idx, 2).Value = info.get('cell_address', '')
                
                # 去掉等号
                cell_formula = info.get('full_formula', '')
                if cell_formula.startswith('='):
                    cell_formula = cell_formula[1:]
                output_sheet.Cells(row_idx, 3).Value = cell_formula
                
                output_sheet.Cells(row_idx, 4).Value = info.get('output_name', '')
                output_sheet.Cells(row_idx, 5).Value = info.get('output_category', '')
                output_sheet.Cells(row_idx, 6).Value = info.get('output_position', 1)
            
            # 设置Output列宽
            for i, width in enumerate([15, 15, 60, 20, 20, 10]):
                output_sheet.Columns(i+1).ColumnWidth = width
            
            # 创建Statistical工作表
            try:
                statistical_sheet = workbook.Worksheets("DriskModel_Statistical")
                statistical_sheet.Delete()
            except:
                pass
            
            statistical_sheet = workbook.Worksheets.Add()
            statistical_sheet.Name = "DriskModel_Statistical"
            
            # 写入Statistical表头
            statistical_headers = [
                "工作表", "单元格地址", "统计函数公式", "函数名称", "计算值"
            ]
            
            for col_idx, header in enumerate(statistical_headers, start=1):
                cell = statistical_sheet.Cells(1, col_idx)
                cell.Value = header
                cell.Font.Bold = True
                cell.Interior.Color = 0xFFCCCC
            
            # 写入Statistical数据
            for row_idx, info in enumerate(self.statistical_info, start=2):
                statistical_sheet.Cells(row_idx, 1).Value = info.get('sheet', '')
                statistical_sheet.Cells(row_idx, 2).Value = info.get('cell_address', '')
                
                # 去掉等号
                cell_formula = info.get('full_formula', '')
                if cell_formula.startswith('='):
                    cell_formula = cell_formula[1:]
                statistical_sheet.Cells(row_idx, 3).Value = cell_formula
                
                statistical_sheet.Cells(row_idx, 4).Value = info.get('function_name', '')
                statistical_sheet.Cells(row_idx, 5).Value = info.get('calculated_value', '#ERROR!')
            
            # 设置Statistical列宽
            for i, width in enumerate([15, 15, 60, 20, 20]):
                statistical_sheet.Columns(i+1).ColumnWidth = width
            
            input_sheet.Activate()
            messagebox.showinfo("导出成功", "模型信息已导出到工作表 'DriskModel_Input'、'DriskModel_Output' 和 'DriskModel_Statistical'")
            
        except Exception as e:
            error_msg = f"导出到Excel失败: {str(e)}"
            logger.error(error_msg)
            messagebox.showerror("导出失败", error_msg)
    
    def run(self):
        """运行窗口"""
        self.root.mainloop()

# ==================== 主宏函数 ====================

from pyxll import xl_macro, xlcAlert
import traceback

@xl_macro()
def DriskModel():
    """
    DriskModel宏 - 静态模式下分析工作簿中所有分布函数、输出函数和统计函数
    """
    try:
        app = _safe_excel_app()
        
        # 确保在静态模式下运行
        from attribute_functions import get_static_mode, set_static_mode
        current_mode = get_static_mode()
        
        if not current_mode:
            response = app.InputBox(
                "检测到当前处于模拟模式。\n\n" +
                "DriskModel需要在静态模式下运行以正确分析分布函数。\n" +
                "是否切换到静态模式并继续分析？\n\n" +
                "点击'确定'切换到静态模式并继续分析\n" +
                "点击'取消'中止操作",
                "DriskModel - 模式切换", Type=2, Default="确定"
            )
            
            if response == "确定":
                set_static_mode(True)
                logger.info("已切换到静态模式")
            else:
                xlcAlert("操作已取消")
                return
        
        # 创建分析器
        analyzer = DistributionModelAnalyzer(app)
        
        # 分析工作簿
        logger.info("开始分析工作簿中的分布函数、输出函数和统计函数...")
        distribution_info, output_info, statistical_info = analyzer.analyze_workbook()
        
        if not distribution_info and not output_info and not statistical_info:
            xlcAlert("未在工作簿中找到任何分布函数、输出函数或统计函数。")
            return
        
        # 创建并显示窗口
        window = ModelInfoWindow(distribution_info, output_info, statistical_info)
        window.run()
        
    except Exception as e:
        error_msg = f"DriskModel执行失败: {str(e)}\n\n{traceback.format_exc()}"
        logger.error(error_msg)
        xlcAlert(f"DriskModel执行失败: {str(e)}")
