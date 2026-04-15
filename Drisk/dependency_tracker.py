# dependency_tracker.py
"""依赖关系追踪模块 - 增强版，支持向量化依赖分析"""

import re
import traceback
from typing import List, Dict, Tuple, Set, Any, Optional
from cell_utils import extract_address_from_cell_object
from formula_parser import (
    extract_nested_distributions_advanced,
    parse_formula_references,
    is_distribution_function,
    is_simtable_function,
    is_makeinput_function,
    is_statistical_function,
    is_output_cell,
    extract_all_distribution_functions_with_index,
    extract_simtable_functions,
    extract_makeinput_functions,
    extract_output_info,
    extract_nested_distributions_from_makeinput
)

class EnhancedDependencyAnalyzer:
    """增强型依赖关系分析器 - 支持向量化依赖分析"""
    
    def __init__(self, app):
        self.app = app
        self.workbook = app.ActiveWorkbook
    
    def find_all_related_cells(self, distribution_cells: Dict[str, List[Dict]], 
                              output_cells: Dict[str, Dict],
                              makeinput_cells: Dict[str, List[Dict]]) -> Dict[str, Dict[str, Any]]:
        """
        查找所有相关单元格（包括中间计算单元格）
        返回格式：{sheet!cell_addr: cell_info}
        """
        all_cells = {}
        
        # 1. 添加分布函数单元格
        for cell_addr, dist_funcs in distribution_cells.items():
            if cell_addr not in all_cells:
                formula = self._get_cell_formula(cell_addr)
                dependencies = self._get_cell_dependencies(cell_addr, formula)
                
                all_cells[cell_addr] = {
                    'type': 'distribution',
                    'funcs': dist_funcs,
                    'formula': formula,
                    'dependencies': dependencies
                }
        
        # 2. 添加MakeInput单元格
        for cell_addr, makeinput_funcs in makeinput_cells.items():
            if cell_addr not in all_cells:
                formula = self._get_cell_formula(cell_addr)
                dependencies = self._get_cell_dependencies(cell_addr, formula)
                
                all_cells[cell_addr] = {
                    'type': 'makeinput',
                    'funcs': makeinput_funcs,
                    'formula': formula,
                    'dependencies': dependencies
                }
        
        # 3. 添加输出单元格
        for cell_addr, info in output_cells.items():
            if cell_addr not in all_cells:
                formula = self._get_cell_formula(cell_addr)
                dependencies = self._get_cell_dependencies(cell_addr, formula)
                
                all_cells[cell_addr] = {
                    'type': 'output',
                    'info': info,
                    'formula': formula,
                    'dependencies': dependencies
                }
        
        # 4. 深度搜索：从输出单元格、分布单元格和MakeInput单元格开始，查找所有相关单元格
        visited = set(all_cells.keys())
        frontier = list(all_cells.keys())
        
        while frontier:
            current_cell = frontier.pop()
            
            # 获取当前单元格的依赖
            if current_cell in all_cells:
                dependencies = all_cells[current_cell]['dependencies']
            else:
                # 如果不在all_cells中，获取其公式和依赖
                formula = self._get_cell_formula(current_cell)
                if formula and formula.startswith('='):
                    all_cells[current_cell] = {
                        'type': 'intermediate',
                        'formula': formula,
                        'dependencies': self._get_cell_dependencies(current_cell, formula)
                    }
                    dependencies = all_cells[current_cell]['dependencies']
                else:
                    continue
            
            # 检查每个依赖
            for dep in dependencies:
                if dep not in visited:
                    # 检查这个依赖单元格是否包含公式
                    dep_formula = self._get_cell_formula(dep)
                    if dep_formula and dep_formula.startswith('='):
                        # 这是一个新的中间计算单元格
                        cell_info = {
                            'type': 'intermediate',
                            'formula': dep_formula,
                            'dependencies': self._get_cell_dependencies(dep, dep_formula)
                        }
                        all_cells[dep] = cell_info
                        frontier.append(dep)
                    visited.add(dep)
        
        return all_cells
    
    def _get_cell_formula(self, cell_addr: str) -> Optional[str]:
        """获取单元格公式"""
        try:
            if '!' in cell_addr:
                sheet_name, addr = cell_addr.split('!')
                sheet = self.workbook.Worksheets(sheet_name)
                cell = sheet.Range(addr)
            else:
                cell = self.app.ActiveSheet.Range(cell_addr)
            
            formula = cell.Formula
            if isinstance(formula, str):
                return formula
            return None
        except Exception as e:
            return None
    
    def _get_cell_dependencies(self, cell_addr: str, formula: str = None) -> List[str]:
        """获取单元格的直接依赖"""
        if formula is None:
            formula = self._get_cell_formula(cell_addr)
        
        if not formula or not formula.startswith('='):
            return []
        
        # 提取公式中的引用
        refs = parse_formula_references(formula)
        
        # 转换为完整地址（带工作表名）
        full_refs = []
        for ref in refs:
            if '!' in ref:
                full_refs.append(ref.upper())
            else:
                # 确定工作表名
                if '!' in cell_addr:
                    sheet_name, _ = cell_addr.split('!')
                else:
                    sheet_name = self.app.ActiveSheet.Name
                full_refs.append(f"{sheet_name}!{ref.upper()}")
        
        return full_refs
    
    def analyze_dependencies_for_vectorization(self, all_cells: Dict[str, Dict[str, Any]]) -> Dict[str, List[str]]:
        """
        分析依赖关系，为向量化计算做准备
        返回：单元格地址 -> 依赖的输入单元格列表
        """
        dependencies = {}
        
        for cell_addr, cell_info in all_cells.items():
            formula = cell_info.get('formula', '')
            if formula and formula.startswith('='):
                # 提取所有单元格引用
                refs = self._extract_cell_references(formula)
                
                # 找出哪些是输入单元格（包含分布函数或MakeInput）
                input_deps = []
                for ref in refs:
                    if ref in all_cells and all_cells[ref].get('type') == 'distribution':
                        input_deps.append(ref)
                    elif ref in all_cells and all_cells[ref].get('type') == 'makeinput':
                        input_deps.append(ref)
                    elif ref in all_cells and all_cells[ref].get('type') == 'intermediate':
                        # 中间单元格可能依赖其他单元格
                        pass
                
                dependencies[cell_addr] = input_deps
        return dependencies

    def _extract_cell_references(self, formula: str) -> List[str]:
        """提取公式中的所有单元格引用"""
        refs = []
        
        # 提取单个单元格引用
        single_pattern = r'([A-Z]{1,3}\d{1,7})'
        single_matches = re.findall(single_pattern, formula, re.IGNORECASE)
        refs.extend(single_matches)
        
        # 提取区域引用
        range_pattern = r'([A-Z]{1,3}\d{1,7})\s*:\s*([A-Z]{1,3}\d{1,7})'
        range_matches = re.findall(range_pattern, formula, re.IGNORECASE)
        for start, end in range_matches:
            # 提取区域内的所有单元格
            cells = self._expand_range(start, end)
            refs.extend(cells)
        
        return list(set(refs))  # 去重

    def _expand_range(self, start: str, end: str) -> List[str]:
        """展开区域引用为单元格列表"""
        cells = []
        
        # 提取起始和结束的行列
        start_col = re.match(r'([A-Z]+)', start).group(1)
        start_row = int(re.match(r'[A-Z]+(\d+)', start).group(1))
        end_col = re.match(r'([A-Z]+)', end).group(1)
        end_row = int(re.match(r'[A-Z]+(\d+)', end).group(1))
        
        if start_col == end_col and start_row <= end_row:
            # 同一列的区域
            for row in range(start_row, end_row + 1):
                cells.append(f"{start_col}{row}")
        
        return cells

def get_all_dependents_direct(app, cell_address: str) -> List[str]:
    """
    获取直接依赖于指定单元格的所有单元格（直接依赖）
    使用Excel的DirectDependents属性
    """
    try:
        # 获取目标单元格
        target_cell = app.ActiveSheet.Range(cell_address)
        
        # 尝试获取直接依赖
        try:
            # 使用DirectDependents获取直接依赖的单元格
            dependents = target_cell.DirectDependents
        except Exception:
            # 如果DirectDependents失败，尝试使用Dependents
            try:
                dependents = target_cell.Dependents
            except Exception:
                return []
        
        # 如果dependents是None，返回空列表
        if dependents is None:
            return []
        
        # 提取所有单元格地址
        addresses = set()
        
        # 处理可能的多个区域
        if hasattr(dependents, 'Areas'):
            for area in dependents.Areas:
                for cell in area.Cells:
                    addr = extract_address_from_cell_object(cell)
                    if addr:
                        addresses.add(addr)
        else:
            # 单个单元格或单个区域
            for cell in dependents.Cells:
                addr = extract_address_from_cell_object(cell)
                if addr:
                    addresses.add(addr)
        
        return list(addresses)
        
    except Exception as e:
        return []

def is_formula_cell(app, cell_address: str) -> bool:
    """检查单元格是否包含公式"""
    try:
        cell = app.ActiveSheet.Range(cell_address)
        formula = cell.Formula
        return isinstance(formula, str) and formula.startswith('=')
    except:
        return False

def find_distribution_functions_in_sheet(app, sheet) -> Dict[str, List[Dict]]:
    """
    查找指定工作表中所有分布函数单元格
    使用增强的嵌套函数提取
    """
    try:
        distribution_cells = {}
        used_range = sheet.UsedRange
        
        for cell in used_range:
            try:
                formula = cell.Formula
                if isinstance(formula, str) and formula.startswith('='):
                    # 获取带工作表名的完整地址
                    sheet_name = sheet.Name
                    addr = f"{sheet_name}!{extract_address_from_cell_object(cell)}"
                    
                    # 使用增强的嵌套函数提取
                    dist_funcs = extract_nested_distributions_advanced(formula, addr)
                    
                    if dist_funcs:
                        distribution_cells[addr] = dist_funcs
                        print(f"找到分布函数单元格 {addr}: {len(dist_funcs)} 个函数")
            except Exception as e:
                continue
        
        return distribution_cells
    except Exception as e:
        print(f"查找分布函数失败: {str(e)}")
        return {}        

def find_simtable_functions_in_sheet(app, sheet) -> Dict[str, List[Dict]]:
    """
    查找指定工作表中所有Simtable函数单元格
    返回格式：{cell_addr: [simtable_func1, simtable_func2, ...]}
    """
    try:
        simtable_cells = {}
        used_range = sheet.UsedRange
        
        for cell in used_range:
            try:
                formula = cell.Formula
                if isinstance(formula, str) and formula.startswith('='):
                    # 检查是否为Simtable函数
                    if is_simtable_function(formula):
                        # 获取带工作表名的完整地址
                        sheet_name = sheet.Name
                        addr = f"{sheet_name}!{extract_address_from_cell_object(cell)}"
                        
                        # 提取Simtable函数信息（包含区域引用处理）
                        simtable_funcs = extract_simtable_functions_with_range(app, sheet, cell, formula)
                        if simtable_funcs:
                            simtable_cells[addr] = simtable_funcs
            except Exception as e:
                continue
        
        return simtable_cells
    except Exception as e:
        return {}

def extract_simtable_functions_with_range(app, sheet, cell, formula: str) -> List[Dict[str, Any]]:
    """
    提取Simtable函数信息，支持单元格区域引用
    """
    try:
        from formula_parser import extract_simtable_functions, parse_args_with_nested_functions
        import re
        
        # 首先使用原有的extract_simtable_functions函数
        simtable_funcs = extract_simtable_functions(formula)
        if not simtable_funcs:
            return []
        
        # 处理每个Simtable函数
        for simtable_func in simtable_funcs:
            args_text = simtable_func.get('args_text', '')
            
            # 检查参数中是否包含区域引用（如E17:E19）
            range_pattern = r'([A-Z]{1,3}\d{1,7})\s*:\s*([A-Z]{1,3}\d{1,7})'
            range_match = re.search(range_pattern, args_text, re.IGNORECASE)
            
            if range_match:
                # 找到区域引用
                start_cell = range_match.group(1)
                end_cell = range_match.group(2)
                
                # 获取区域内的所有值
                try:
                    range_address = f"{start_cell}:{end_cell}"
                    range_obj = sheet.Range(range_address)
                    
                    # 获取区域内的所有值
                    values = []
                    if hasattr(range_obj, 'Value2'):
                        range_values = range_obj.Value2
                        
                        # 处理可能的一维或二维数组
                        if isinstance(range_values, tuple):
                            # 二维数组
                            for row in range_values:
                                if isinstance(row, tuple):
                                    # 多列
                                    for val in row:
                                        values.append(val)
                                else:
                                    # 单列
                                    values.append(row)
                        else:
                            # 单个值或一维数组
                            values.append(range_values)
                    
                    # 更新Simtable函数的值
                    simtable_func['values'] = values
                    simtable_func['range_reference'] = range_address
                    simtable_func['is_range'] = True
                    
                except Exception as e:
                    print(f"获取Simtable区域引用值失败: {str(e)}")
                    # 如果获取失败，使用原有的值
                    pass
        
        return simtable_funcs
    except Exception as e:
        print(f"提取Simtable函数信息失败: {str(e)}")
        return []

def find_makeinput_functions_in_sheet(app, sheet) -> Dict[str, List[Dict]]:
    """
    查找指定工作表中所有MakeInput函数单元格
    返回格式：{cell_addr: [makeinput_func1, makeinput_func2, ...]}
    
    增强：提取MakeInput中的内嵌分布函数，并将它们合并到分布函数列表中
    """
    try:
        makeinput_cells = {}
        used_range = sheet.UsedRange
        
        for cell in used_range:
            try:
                formula = cell.Formula
                if isinstance(formula, str) and formula.startswith('='):
                    # 检查是否为MakeInput函数
                    if is_makeinput_function(formula):
                        # 获取带工作表名的完整地址
                        sheet_name = sheet.Name
                        addr = f"{sheet_name}!{extract_address_from_cell_object(cell)}"
                        
                        # 提取MakeInput函数信息（包括内嵌分布函数）
                        makeinput_funcs = extract_makeinput_functions(formula)
                        if makeinput_funcs:
                            # 提取内嵌分布函数
                            for makeinput_func in makeinput_funcs:
                                nested_dists = extract_nested_distributions_from_makeinput(formula, addr)
                                if nested_dists:
                                    makeinput_func['nested_distributions'] = nested_dists
                            
                            makeinput_cells[addr] = makeinput_funcs
            except Exception as e:
                continue
        
        return makeinput_cells
    except Exception as e:
        return {}

def find_output_cells_in_sheet(app, sheet) -> Dict[str, Dict]:
    """
    查找指定工作表中所有输出单元格
    返回格式：{cell_addr: output_info}
    
    修改：name向左查找第一个非空字符串，category向上查找第一个非空字符串
    """
    try:
        output_cells = {}
        used_range = sheet.UsedRange
        
        for cell in used_range:
            try:
                formula = cell.Formula
                if isinstance(formula, str) and formula.startswith('='):
                    if is_output_cell(formula):
                        # 获取带工作表名的完整地址
                        sheet_name = sheet.Name
                        addr = f"{sheet_name}!{extract_address_from_cell_object(cell)}"
                        
                        # 提取输出信息
                        output_info = extract_output_info(formula)
                        
                        # 获取单元格行列号
                        row = cell.Row
                        col = cell.Column
                        
                        # 若 name 为空，向左查找第一个非空字符串（最多10列）
                        if not output_info.get('name'):
                            left_name = ""
                            for i in range(1, 11):
                                if col - i < 1:
                                    break
                                left_cell = sheet.Cells(row, col - i)
                                val = left_cell.Value
                                if val is not None and isinstance(val, str) and val.strip():
                                    left_name = str(val).strip()
                                    break
                            output_info['name'] = left_name
                        
                        # 若 category 为空，向上查找第一个非空字符串（最多10行）
                        if not output_info.get('category'):
                            up_category = ""
                            for i in range(1, 11):
                                if row - i < 1:
                                    break
                                up_cell = sheet.Cells(row - i, col)
                                val = up_cell.Value
                                if val is not None and isinstance(val, str) and val.strip():
                                    up_category = str(val).strip()
                                    break
                            output_info['category'] = up_category
                        
                        # 确保所有必要字段都存在
                        if 'position' not in output_info:
                            output_info['position'] = 1
                        
                        output_cells[addr] = output_info
            except Exception as e:
                continue
        
        return output_cells
    except Exception as e:
        return {}

def find_cell_name(app, sheet, cell) -> str:
    """
    为单元格自动命名：向上查找第一个非空字符串，向左查找第一个非空字符串
    格式："上方字符串_左方字符串" 或找到的单个字符串
    """
    try:
        # 获取单元格的行列
        row = cell.Row
        col = cell.Column
        
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
            return ""
    except Exception as e:
        return ""

def get_output_dependencies(app, output_cell_addr: str) -> List[str]:
    """
    获取输出单元格依赖的所有分布函数单元格
    """
    try:
        # 解析工作表名和单元格地址
        if '!' in output_cell_addr:
            sheet_name, cell_addr = output_cell_addr.split('!')
            sheet = app.ActiveWorkbook.Worksheets(sheet_name)
            cell = sheet.Range(cell_addr)
        else:
            cell = app.ActiveSheet.Range(output_cell_addr)
            
        formula = cell.Formula
        
        if not isinstance(formula, str) or not formula.startswith('='):
            return []
        
        # 获取公式中引用的所有单元格
        refs = parse_formula_references(formula)
        
        # 找出哪些是分布函数单元格
        dist_deps = []
        for ref in refs:
            try:
                # 检查ref是否包含工作表名
                if '!' in ref:
                    ref_sheet_name, ref_cell_addr = ref.split('!')
                    ref_sheet = app.ActiveWorkbook.Worksheets(ref_sheet_name)
                    ref_cell = ref_sheet.Range(ref_cell_addr)
                else:
                    ref_cell = cell.Worksheet.Range(ref)
                    
                ref_formula = ref_cell.Formula
                if isinstance(ref_formula, str) and ref_formula.startswith('='):
                    if is_distribution_function(ref_formula):
                        # 添加完整地址（包含工作表名）
                        ref_addr = f"{ref_cell.Worksheet.Name}!{ref_cell_addr if '!' in ref else ref}"
                        dist_deps.append(ref_addr)
            except:
                continue
        
        return dist_deps
    except:
        return []

def find_all_simulation_cells_in_workbook(app) -> Tuple[Dict, Dict, Dict, Dict, List, Dict]:
    """
    查找整个工作簿中所有需要模拟的单元格（跨sheet版本）
    仅处理可见工作表（不包含隐藏或深度隐藏工作表）
    返回：
        distribution_cells: {sheet!cell_addr: [dist_func1, dist_func2, ...]}
        simtable_cells: {sheet!cell_addr: [simtable_func1, simtable_func2, ...]}
        makeinput_cells: {sheet!cell_addr: [makeinput_func1, makeinput_func2, ...]}
        output_cells: {sheet!cell_addr: output_info}
        all_input_keys: [sheet!cell_addr_index1, sheet!cell_addr_index2, ...]
        all_related_cells: {sheet!cell_addr: cell_info} - 新增：所有相关单元格
    
    关键修改：不再为内嵌分布添加虚拟单元格，所有分布函数（包括内嵌）都记录在物理单元格的 dist_funcs 中。
    """
    try:
        workbook = app.ActiveWorkbook
        distribution_cells = {}
        simtable_cells = {}
        makeinput_cells = {}
        output_cells = {}
        all_input_keys = []
        
        # 遍历工作簿中的所有工作表，仅处理可见工作表
        for sheet in workbook.Worksheets:
            # 检查工作表是否可见：-1 表示 xlSheetVisible
            if sheet.Visible != -1:
                continue  # 跳过隐藏或深度隐藏工作表
            
            try:
                # 1. 查找当前工作表中的分布函数单元格
                sheet_dist_cells = find_distribution_functions_in_sheet(app, sheet)
                if sheet_dist_cells:
                    distribution_cells.update(sheet_dist_cells)
                
                # 2. 查找当前工作表中的Simtable函数单元格
                sheet_simtable_cells = find_simtable_functions_in_sheet(app, sheet)
                if sheet_simtable_cells:
                    simtable_cells.update(sheet_simtable_cells)
                
                # 3. 查找当前工作表中的MakeInput函数单元格
                sheet_makeinput_cells = find_makeinput_functions_in_sheet(app, sheet)
                if sheet_makeinput_cells:
                    makeinput_cells.update(sheet_makeinput_cells)
                
                # 4. 查找当前工作表中的输出单元格
                sheet_output_cells = find_output_cells_in_sheet(app, sheet)
                if sheet_output_cells:
                    output_cells.update(sheet_output_cells)
                    
            except Exception as e:
                continue
        
        # --- 删除为内嵌分布添加虚拟单元格的代码 ---
        # 确保所有分布函数（包括内嵌）都已正确标记，但不再创建虚拟单元格
        
        # 5. 生成所有Input键（包含工作表名）
        # 首先，为普通分布函数单元格生成输入键
        for cell_addr, dist_funcs in distribution_cells.items():
            for dist_func in dist_funcs:
                # 检查是否为内嵌分布函数
                is_nested = dist_func.get('is_nested', False)
                is_at_function = dist_func.get('is_at_function', False)
                
                # 内嵌函数使用特殊的键格式
                if is_nested:
                    # 嵌套函数：key = "工作表名!单元格地址_nested_深度_索引"
                    depth = dist_func.get('depth', 0)
                    parent_indices = dist_func.get('parent_indices', [])
                    if parent_indices:
                        input_key = f"{cell_addr}_nested_{depth}_{dist_func.get('index', 1)}"
                    else:
                        input_key = f"{cell_addr}_nested_{dist_func.get('index', 1)}"
                elif is_at_function:
                    input_key = f"{cell_addr}_at_{dist_func.get('index', 1)}"
                else:
                    input_key = f"{cell_addr}_{dist_func.get('index', 1)}"
                
                # 确保键的唯一性
                if input_key not in all_input_keys:
                    all_input_keys.append(input_key)
                
                # 记录到dist_func中，用于后续查找
                dist_func['input_key'] = input_key
                
                # 打印调试信息
                func_type = "普通分布"
                if is_at_function:
                    func_type = "@函数"
                elif is_nested:
                    func_type = "嵌套分布"
                
                print(f"  生成输入键: {input_key} (类型={func_type}, 函数={dist_func.get('func_name', 'N/A')}, "
                    f"index={dist_func.get('index', 'N/A')}, depth={dist_func.get('depth', 'N/A')})")
        
        # 6. 使用增强型依赖分析器查找所有相关单元格
        analyzer = EnhancedDependencyAnalyzer(app)
        all_related_cells = analyzer.find_all_related_cells(distribution_cells, output_cells, makeinput_cells)
        
        # 打印详细的输入信息
        print(f"\n跨sheet模拟: 输入统计")
        print(f"  分布单元格: {len(distribution_cells)} 个（物理单元格）")
        print(f"  Simtable单元格: {len(simtable_cells)} 个")
        print(f"  MakeInput单元格: {len(makeinput_cells)} 个")
        print(f"  输出单元格: {len(output_cells)} 个")
        print(f"  总输入键: {len(all_input_keys)} 个")
        
        # 详细打印每个输入键
        print(f"\n详细输入键列表:")
        for i, key in enumerate(all_input_keys):
            # 获取对应的分布函数信息
            func_info = None
            for cell_addr, funcs in distribution_cells.items():
                for func in funcs:
                    if func.get('key') == key or f"{cell_addr}_{func.get('index')}" == key:
                        func_info = func
                        break
                if func_info:
                    break
            
            func_type = "普通分布"
            if func_info:
                if func_info.get('is_nested'):
                    func_type = "嵌套分布"
            
            print(f"  {i+1}. {key} ({func_type})")
        
        print(f"\n  相关单元格: {len(all_related_cells)} 个")
        
        return distribution_cells, simtable_cells, makeinput_cells, output_cells, all_input_keys, all_related_cells
        
    except Exception as e:
        print(f"查找模拟单元格时出错: {str(e)}")
        traceback.print_exc()
        return {}, {}, {}, {}, [], {}

def find_distribution_functions(app, sheet_name=None) -> Dict[str, List[Dict]]:
    """
    查找工作表中所有分布函数单元格（保持向后兼容）
    """
    if sheet_name:
        sheet = app.ActiveWorkbook.Worksheets(sheet_name)
        return find_distribution_functions_in_sheet(app, sheet)
    else:
        sheet = app.ActiveSheet
        return find_distribution_functions_in_sheet(app, sheet)

def find_simtable_functions(app, sheet_name=None) -> Dict[str, List[Dict]]:
    """
    查找工作表中所有Simtable函数单元格
    """
    if sheet_name:
        sheet = app.ActiveWorkbook.Worksheets(sheet_name)
        return find_simtable_functions_in_sheet(app, sheet)
    else:
        sheet = app.ActiveSheet
        return find_simtable_functions_in_sheet(app, sheet)

def find_makeinput_functions(app, sheet_name=None) -> Dict[str, List[Dict]]:
    """
    查找工作表中所有MakeInput函数单元格
    """
    if sheet_name:
        sheet = app.ActiveWorkbook.Worksheets(sheet_name)
        return find_makeinput_functions_in_sheet(app, sheet)
    else:
        sheet = app.ActiveSheet
        return find_makeinput_functions_in_sheet(app, sheet)

def find_output_cells(app, sheet_name=None) -> Dict[str, Dict]:
    """
    查找工作表中所有输出单元格（保持向后兼容）
    """
    if sheet_name:
        sheet = app.ActiveWorkbook.Worksheets(sheet_name)
        return find_output_cells_in_sheet(app, sheet)
    else:
        sheet = app.ActiveSheet
        return find_output_cells_in_sheet(app, sheet)