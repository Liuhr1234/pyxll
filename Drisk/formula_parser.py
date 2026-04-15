# formula_parser.py
"""公式解析器模块 - 统一版本，包含所有公式解析功能"""

import re
import numpy as np
from typing import List, Dict, Tuple, Any, Optional
from constants import DISTRIBUTION_FUNCTION_NAMES
DISTRIBUTION_FUNCTIONS = DISTRIBUTION_FUNCTION_NAMES

# ==================== 基础常量 ====================

SIMTABLE_FUNCTIONS = [
    "DriskSimtable"
]

# 新增MakeInput函数
MAKEINPUT_FUNCTIONS = [  # 新增
    "DriskMakeInput"
]

STATISTICAL_FUNCTIONS = [
    "DriskMean", "DriskStd", "DriskMin", "DriskMax", 
    "DriskMed", "DriskPtoX", "DriskData"
]

ATTRIBUTE_FUNCTIONS = [
    "DriskName", "DriskLoc", "DriskCategory", "DriskCollect", "DriskConvergence",
    "DriskCopula", "DriskCorrmat", "DriskFit", "DriskIsDate", "DriskIsDiscrete",
    "DriskLock", "DriskSeed", "DriskShift", "DriskStatic", "DriskTruncate",
    "DriskTruncateP", "DriskTruncate2", "DriskTruncateP2", "DriskUnits",
    "DriskMakeInput" 
]

# ==================== 基础公式解析函数 ====================

def parse_formula_references(formula: str) -> List[str]:
    """
    解析公式中的单元格引用
    
    Args:
        formula: Excel公式字符串
        
    Returns:
        单元格引用列表
    """
    if not formula or not isinstance(formula, str) or not formula.startswith('='):
        return []
    
    formula_body = formula[1:]
    
    # 提取单元格引用 - 支持跨工作表引用
    pattern = r'(?:[A-Za-z_][A-Za-z0-9_\.]*!)?\$?[A-Za-z]+\$?\d+(?::\$?[A-Za-z]+\$?\d+)?'
    matches = re.findall(pattern, formula_body, re.IGNORECASE)
    
    # 清理引用
    refs = []
    for ref in matches:
        # 移除$符号
        clean_ref = ref.replace('$', '')
        
        # 处理区域引用
        if ':' in clean_ref:
            # 区域引用，提取起始单元格
            start_cell, end_cell = clean_ref.split(':')
            # 如果是跨工作表引用，处理工作表名
            if '!' in start_cell:
                sheet_name, cell = start_cell.split('!')
                refs.append(f"{sheet_name}!{cell.upper()}")
            else:
                refs.append(start_cell.upper())
        else:
            # 单个单元格引用
            refs.append(clean_ref.upper())
    
    # 移除重复引用
    return list(set(refs))

def extract_cell_references_from_args(args_text: str) -> List[str]:
    """从参数文本中提取单元格引用"""
    refs = []
    # 查找类似A1, B2, C3等单元格引用
    pattern = r'[A-Z]{1,3}\d{1,7}'
    matches = re.findall(pattern, args_text, re.IGNORECASE)
    
    for match in matches:
        # 检查是否是工作表引用（如Sheet1!A1）
        if '!' in args_text:
            # 找到工作表名
            sheet_pattern = r'([A-Za-z0-9_]+)!' + re.escape(match)
            sheet_match = re.search(sheet_pattern, args_text, re.IGNORECASE)
            if sheet_match:
                ref = f"{sheet_match.group(1)}!{match}"
            else:
                ref = match
        
        refs.append(ref.upper())
    
    return list(set(refs))  # 去重

def expand_cell_range(start: str, end: str) -> List[str]:
    """展开区域引用为单元格列表"""
    cells = []
    
    # 提取起始和结束的行列
    try:
        start_col = re.match(r'([A-Z]+)', start).group(1)
        start_row = int(re.match(r'[A-Z]+(\d+)', start).group(1))
        end_col = re.match(r'([A-Z]+)', end).group(1)
        end_row = int(re.match(r'[A-Z]+(\d+)', end).group(1))
        
        if start_col == end_col and start_row <= end_row:
            # 同一列的区域
            for row in range(start_row, end_row + 1):
                cells.append(f"{start_col}{row}")
    except Exception:
        pass
    
    return cells

# ================================================================================
# ui_tornado专用
# 用于敏感性分析、情景分析相关输入变量依赖关系查询 -- 支持一组output同时查询input
def parse_formula_references_tornado(
    formula: str,
    *,
    exclude_makeinput_inner: bool = False,
) -> List[str]:
    """
    Parse Excel formula references into a flat list of cell addresses.

    Range references are expanded into all member cells to preserve full
    dependency coverage (for example, MAX(A1:A10) includes A1..A10).
    """
    if not formula or not isinstance(formula, str) or not formula.startswith('='):
        return []

    formula_body = formula[1:]

    def _mask_makeinput_calls(expr: str) -> str:
        """
        Replace DriskMakeInput(...) spans with spaces so references inside
        MakeInput branches are not parsed as expandable dependencies.
        """
        if not expr:
            return expr

        token = "DRISKMAKEINPUT"
        upper_expr = expr.upper()
        masked_chars = list(expr)
        search_pos = 0

        while True:
            idx = upper_expr.find(token, search_pos)
            if idx < 0:
                break

            # Keep matching strict enough to avoid partial identifier hits.
            prev = upper_expr[idx - 1] if idx > 0 else ""
            if prev and (prev.isalnum() or prev == "_"):
                search_pos = idx + len(token)
                continue

            cursor = idx + len(token)
            while cursor < len(expr) and expr[cursor].isspace():
                cursor += 1
            if cursor >= len(expr) or expr[cursor] != "(":
                search_pos = idx + len(token)
                continue

            depth = 0
            quote_char = None
            end = cursor
            while end < len(expr):
                ch = expr[end]
                if quote_char:
                    if ch == quote_char:
                        # Excel escapes quotes by doubling them.
                        if end + 1 < len(expr) and expr[end + 1] == quote_char:
                            end += 1
                        else:
                            quote_char = None
                else:
                    if ch in ("'", '"'):
                        quote_char = ch
                    elif ch == "(":
                        depth += 1
                    elif ch == ")":
                        depth -= 1
                        if depth == 0:
                            end += 1
                            break
                end += 1

            if end <= idx:
                search_pos = idx + len(token)
                continue

            for i in range(idx, min(end, len(masked_chars))):
                if masked_chars[i] not in ("\r", "\n", "\t"):
                    masked_chars[i] = " "
            search_pos = end

        return "".join(masked_chars)

    if exclude_makeinput_inner:
        formula_body = _mask_makeinput_calls(formula_body)

    # Match A1, Sheet1!A1, 'My Sheet'!A1, and their range forms.
    sheet_pattern = r"(?:'(?:[^']|'')+'|[A-Za-z_][A-Za-z0-9_\.]*)"
    cell_pattern = r"\$?[A-Za-z]{1,7}\$?\d+"
    pattern = rf"(?:{sheet_pattern}!)?{cell_pattern}(?::(?:{sheet_pattern}!)?{cell_pattern})?"
    matches = re.findall(pattern, formula_body, re.IGNORECASE)

    refs: List[str] = []
    for ref in matches:
        clean_ref = ref.replace('$', '')
        if ':' in clean_ref:
            start_ref, end_ref = clean_ref.split(':', 1)
            expanded = expand_cell_range_tornado(start_ref, end_ref)
            if expanded:
                refs.extend(expanded)
            else:
                refs.append(start_ref)
                refs.append(end_ref)
        else:
            refs.append(clean_ref)

    # Keep reference order stable while removing duplicates.
    deduped: List[str] = []
    seen = set()
    for ref in refs:
        ref_upper = ref.upper()
        if ref_upper in seen:
            continue
        seen.add(ref_upper)
        deduped.append(ref_upper)
    return deduped


def expand_cell_range_tornado(start: str, end: str) -> List[str]:
    """Expand a rectangular cell range into member cell references."""
    try:
        def _split_ref(ref_text: str) -> Tuple[Optional[str], str]:
            ref_text = ref_text.strip()
            if "!" not in ref_text:
                return None, ref_text
            sheet_part, cell_part = ref_text.rsplit("!", 1)
            if sheet_part.startswith("'") and sheet_part.endswith("'"):
                sheet_part = sheet_part[1:-1].replace("''", "'")
            return sheet_part, cell_part

        def _col_to_index(col_text: str) -> int:
            value = 0
            for ch in col_text:
                value = value * 26 + (ord(ch) - ord("A") + 1)
            return value

        def _index_to_col(index: int) -> str:
            out = []
            while index > 0:
                index, rem = divmod(index - 1, 26)
                out.append(chr(ord("A") + rem))
            return "".join(reversed(out))

        def _parse_cell(cell_text: str) -> Optional[Tuple[int, int]]:
            m = re.fullmatch(r"([A-Za-z]{1,7})(\d+)", cell_text.strip())
            if not m:
                return None
            col_idx = _col_to_index(m.group(1).upper())
            row_idx = int(m.group(2))
            if row_idx < 1:
                return None
            return col_idx, row_idx

        start_sheet, start_cell = _split_ref(start)
        end_sheet, end_cell = _split_ref(end)

        if start_sheet and end_sheet and start_sheet.upper() != end_sheet.upper():
            return []
        sheet_name = start_sheet or end_sheet

        start_pos = _parse_cell(start_cell)
        end_pos = _parse_cell(end_cell)
        if start_pos is None or end_pos is None:
            return []

        c1, r1 = start_pos
        c2, r2 = end_pos
        left, right = min(c1, c2), max(c1, c2)
        top, bottom = min(r1, r2), max(r1, r2)

        cells: List[str] = []
        for col_idx in range(left, right + 1):
            col_text = _index_to_col(col_idx)
            for row_idx in range(top, bottom + 1):
                cell_ref = f"{col_text}{row_idx}"
                if sheet_name:
                    cells.append(f"{sheet_name}!{cell_ref}")
                else:
                    cells.append(cell_ref)
        return cells
    except Exception:
        return []
# ================================================================================

# ==================== 函数类型检测 ====================

def is_distribution_function(formula: str) -> bool:
    """检查公式是否包含分布函数"""
    if not isinstance(formula, str):
        return False
    
    formula_upper = formula.upper()
    for func in DISTRIBUTION_FUNCTIONS:
        if func.upper() in formula_upper:
            return True
    
    return False

def is_simtable_function(formula: str) -> bool:
    """检查公式是否包含Simtable函数"""
    if not isinstance(formula, str):
        return False
    
    formula_upper = formula.upper()
    for func in SIMTABLE_FUNCTIONS:
        if func.upper() in formula_upper:
            return True
    
    return False

def is_makeinput_function(formula: str) -> bool:  # 新增
    """检查公式是否包含MakeInput函数"""
    if not isinstance(formula, str):
        return False
    
    formula_upper = formula.upper()
    for func in MAKEINPUT_FUNCTIONS:
        if func.upper() in formula_upper:
            return True
    
    return False

def is_statistical_function(formula: str) -> bool:
    """检查公式是否包含统计函数"""
    if not isinstance(formula, str):
        return False
    
    formula_upper = formula.upper()
    for func in STATISTICAL_FUNCTIONS:
        if func.upper() in formula_upper:
            return True
    
    return False

def is_attribute_function(formula: str) -> bool:
    """检查公式是否包含属性函数"""
    if not isinstance(formula, str):
        return False
    
    formula_upper = formula.upper()
    for func in ATTRIBUTE_FUNCTIONS:
        if func.upper() in formula_upper:
            return True
    
    return False

def is_output_cell(formula: str) -> bool:
    """检查公式是否包含DriskOutput函数"""
    if not isinstance(formula, str):
        return False
    
    formula_upper = formula.upper()
    return "DRISKOUTPUT" in formula_upper

def has_static_attribute(formula: str) -> bool:
    """检查公式是否包含DriskStatic属性"""
    if not isinstance(formula, str):
        return False
    
    formula_upper = formula.upper()
    return "DriskStatic" in formula_upper

# ==================== 完整公式解析 ====================

def parse_complete_formula(formula: str):
    """解析完整公式，提取函数名和所有参数。
    
    Args:
        formula: Excel 公式字符串（包含 = 符号）
        
    Returns:
        (函数名, 参数列表) 元组，如果解析失败返回 (None, [])
    """
    if not formula or not formula.startswith('='):
        return None, []
    
    # 移除等号
    formula = formula[1:].strip()
    
    # 查找函数名和参数部分
    func_match = re.match(r'([A-Za-z_][A-Za-z0-9_]*)\s*\((.*)\)', formula, re.DOTALL)
    if not func_match:
        return None, []
    
    func_name = func_match.group(1)
    args_text = func_match.group(2).strip()
    
    # 解析参数列表（考虑嵌套函数）
    args_list = parse_args_with_nested_functions(args_text)
    
    return func_name, args_list

def parse_args_with_nested_functions(args_text: str):
    """解析参数列表，考虑嵌套函数和引号。
    
    Args:
        args_text: 参数文本（不含外层括号）
        
    Returns:
        参数列表
    """
    args_list = []
    current_arg = ""
    paren_depth = 0
    brace_depth = 0
    in_quotes = False
    quote_char = None
    
    for char in args_text:
        if char in ['"', "'"]:
            if not in_quotes:
                in_quotes = True
                quote_char = char
            elif char == quote_char:
                in_quotes = False
            current_arg += char
        elif not in_quotes:
            if char == '(':
                paren_depth += 1
                current_arg += char
            elif char == ')':
                paren_depth -= 1
                current_arg += char
            elif char == '{':
                brace_depth += 1
                current_arg += char
            elif char == '}':
                brace_depth = max(0, brace_depth - 1)
                current_arg += char
            elif char == ',' and paren_depth == 0 and brace_depth == 0:
                args_list.append(current_arg.strip())
                current_arg = ""
            else:
                current_arg += char
        else:
            current_arg += char
    
    if current_arg:
        args_list.append(current_arg.strip())
    
    return args_list

# ==================== 标记函数解析 ====================

def parse_marker_function(func_str: str):
    """解析标记函数（如 DriskShift、DriskTruncate 等）。
    
    Args:
        func_str: 函数字符串
        
    Returns:
        (标记类型, 标记值) 元组，如果解析失败返回 None
    """
    # 匹配函数名和参数
    match = re.match(r'([A-Za-z_][A-Za-z0-9_]*)\s*\((.*)\)', func_str)
    if not match:
        return None
    
    func_name = match.group(1).lower()
    args_text = match.group(2).strip()
    
    # 根据函数名解析标记类型和值
    if 'shift' in func_name:
        try:
            shift_val = float(args_text)
            return 'shift', shift_val
        except:
            return 'shift', 0.0
    
    elif 'truncate' in func_name and 'truncatep' not in func_name:
        # 检查是truncate还是truncate2
        if func_name.endswith('2'):
            # DriskTruncate2 - 先平移后截断
            return parse_truncate_args(args_text, 'truncate2')
        else:
            # DriskTruncate - 先截断后平移
            return parse_truncate_args(args_text, 'truncate')
    
    elif 'truncatep' in func_name:
        # 检查是truncatep还是truncatep2
        if func_name.endswith('2'):
            # DriskTruncateP2 - 先平移后百分比截断
            return parse_truncate_args(args_text, 'truncatep2')
        else:
            # DriskTruncateP - 先百分比截断后平移
            return parse_truncate_args(args_text, 'truncatep')
    
    elif 'static' in func_name:
        try:
            static_val = float(args_text)
            return 'static', static_val
        except:
            return 'static', 0.0
    
    elif 'loc' in func_name:
        return 'loc', True
    
    elif 'seed' in func_name:
        args_list = [arg.strip() for arg in args_text.split(',')]
        if len(args_list) >= 2:
            try:
                rng_type = int(float(args_list[0]))
                seed = int(float(args_list[1]))
                return 'seed', f"{rng_type},{seed}"
            except:
                pass
        elif args_list:
            try:
                seed = int(float(args_list[0]))
                return 'seed', f"1,{seed}"
            except:
                pass
    
    elif 'name' in func_name:
        try:
            name_val = args_text.strip().strip('"\'')
            return 'name', name_val
        except:
            return 'name', ''
    
    elif 'units' in func_name:
        try:
            units_val = args_text.strip().strip('"\'')
            return 'units', units_val
        except:
            return 'units', ''
    
    elif 'category' in func_name:
        try:
            category_val = args_text.strip().strip('"\'')
            return 'category', category_val
        except:
            return 'category', ''
    
    elif 'is_date' in func_name:
        try:
            is_date_val = args_text.strip().lower() == 'true'
            return 'is_date', is_date_val
        except:
            return 'is_date', False
    
    elif 'is_discrete' in func_name:
        try:
            is_discrete_val = args_text.strip().lower() == 'true'
            return 'is_discrete', is_discrete_val
        except:
            return 'is_discrete', False
    
    elif 'makeinput' in func_name:  # 新增
        # MakeInput函数，返回True标记
        return 'makeinput', True
    
    return None

def parse_truncate_args(args_text: str, truncate_type: str):
    """解析截断参数"""
    return truncate_type, args_text.strip()

def extract_dist_params_and_markers(args_list: List[str]):
    """从参数列表中提取分布参数和标记。
    
    Args:
        args_list: 参数列表
        
    Returns:
        (分布参数列表, 标记字典) 元组
    """
    dist_params = []
    markers = {}
    
    for arg in args_list:
        if not arg:
            continue
            
        # 检查是否是Drisk函数（标记函数）
        if arg.startswith('Drisk'):
            # 解析标记函数
            marker_info = parse_marker_function(arg)
            if marker_info:
                marker_type, marker_value = marker_info
                if marker_type and marker_value is not None:
                    markers[marker_type] = marker_value
        else:
            # 尝试解析为数值参数
            try:
                # 移除可能的空格和括号
                arg_clean = arg.strip().strip('()')
                
                # 尝试直接转换
                val = float(arg_clean)
                dist_params.append(val)
            except ValueError:
                # 可能是表达式或单元格引用
                try:
                    # 尝试评估表达式
                    safe_dict = {
                        '__builtins__': {},
                        'exp': np.exp, 'log': np.log, 'sqrt': np.sqrt,
                        'sin': np.sin, 'cos': np.cos, 'tan': np.tan,
                        'pi': np.pi, 'e': np.e
                    }
                    # 替换^为**
                    expr = arg_clean.replace('^', '**')
                    val = eval(expr, {"__builtins__": {}}, safe_dict)
                    dist_params.append(float(val))
                except:
                    # 可能是单元格引用，跳过
                    pass
    
    return dist_params, markers

# ==================== 分布函数解析 ====================

def extract_all_distribution_functions(formula: str) -> List[Dict[str, Any]]:
    """
    提取公式中的所有分布函数（包括嵌套的）
    返回格式：[{'func_name': 'DriskNormal', 'full_match': 'DriskNormal(5,2)', 'args_text': '5,2'}, ...]
    
    增强版：使用括号匹配算法正确处理嵌套函数
    """
    if not isinstance(formula, str) or not formula.startswith('='):
        return []
    
    formula_body = formula[1:]
    distribution_functions = []
    
    # 递归提取分布函数，使用括号匹配算法
    def extract_recursive(expr: str, depth: int = 0, start_pos: int = 0):
        
        # 构建正则表达式模式，匹配函数名（可能带@符号）
        pattern_str = r'(@?)(' + '|'.join(DISTRIBUTION_FUNCTIONS) + r')\s*\('
        pattern = re.compile(pattern_str, re.IGNORECASE)
        
        # 查找所有匹配的函数名
        pos = 0
        while pos < len(expr):
            match = pattern.search(expr, pos)
            if not match:
                break
            
            # 获取匹配信息
            at_symbol = match.group(1)  # 可能为@或空
            func_name = match.group(2)  # 函数名
            func_start_in_expr = match.start()  # 在expr中的起始位置
            args_start = match.end()  # 参数开始位置（在'('之后）
            
            # 使用栈匹配括号，找到对应的右括号
            paren_stack = 1  # 已经有一个左括号
            args_end = args_start
            
            while paren_stack > 0 and args_end < len(expr):
                if expr[args_end] == '(':
                    paren_stack += 1
                elif expr[args_end] == ')':
                    paren_stack -= 1
                args_end += 1
            
            if paren_stack == 0:
                # 成功匹配括号对
                args_text = expr[args_start:args_end-1]  # 去掉最后的')'
                
                # 计算全局位置
                global_start = start_pos + func_start_in_expr
                global_end = start_pos + args_end
                
                # 检查是否带@符号
                is_at_function = bool(at_symbol)
                
                # 检查是否嵌套（通过参数中是否包含其他分布函数）
                is_nested = False
                
                # 递归查找参数中的分布函数
                inner_functions = extract_recursive(
                    args_text, 
                    depth + 1, 
                    global_start + (args_start - func_start_in_expr)
                )
                
                if inner_functions:
                    distribution_functions.extend(inner_functions)
                    is_nested = True
                
                # 添加当前函数
                full_match = expr[func_start_in_expr:args_end]
                
                distribution_functions.append({
                    'func_name': func_name,
                    'args_text': args_text,
                    'full_match': full_match,
                    'is_at_function': is_at_function,
                    'is_nested': is_nested or (depth > 0),
                    'depth': depth,
                    'start_pos': global_start,
                    'end_pos': global_end
                })
                
                # 继续搜索下一个函数（从当前函数结束位置开始）
                pos = args_end
            else:
                # 括号不匹配，跳过这个匹配
                pos = match.end()
    
    extract_recursive(formula_body)
    
    # 按位置排序（从左到右）
    distribution_functions.sort(key=lambda x: x.get('start_pos', 0))
    
    return distribution_functions

def extract_all_distribution_functions_with_index(formula: str, cell_addr: str) -> List[Dict[str, Any]]:
    """
    增强版：提取公式中的所有分布函数（包括嵌套的），并分配索引
    """
    if not isinstance(formula, str) or not formula.startswith('='):
        return []
    
    # 提取所有分布函数（包括嵌套的）
    funcs = extract_all_distribution_functions(formula)
    
    # 为每个函数分配索引和位置信息
    for i, func in enumerate(funcs):
        func['cell_addr'] = cell_addr
        func['index'] = i + 1
        func['key'] = f"{cell_addr}_{i+1}"
        # 标记是否为嵌套函数
        func['is_nested'] = func.get('is_nested', False) or func.get('depth', 0) > 0
        # 确保有完整的函数信息
        if 'func_name' not in func:
            func['func_name'] = func.get('function_name', '')
        if 'args_text' not in func:
            func['args_text'] = func.get('args', '')
        if 'full_match' not in func:
            func['full_match'] = func.get('original_text', '')
        # 确保有起始和结束位置
        if 'start_pos' not in func:
            func['start_pos'] = i * 10  # 默认值
        if 'end_pos' not in func:
            func['end_pos'] = i * 10 + 5  # 默认值
    
    return funcs

# ==================== 增强的嵌套函数提取 ====================

def extract_nested_distributions_advanced(formula: str, cell_addr: str) -> List[Dict[str, Any]]:
    """
    增强的嵌套函数提取：正确处理复杂嵌套情况
    返回格式：[{'func_name': 'DriskNormal', 'args_text': '...', 'full_match': '...', 
               'depth': 0, 'parent_index': None, 'nested_path': []}, ...]
    """
    if not isinstance(formula, str) or not formula.startswith('='):
        return []
    
    formula_body = formula[1:]
    all_functions = []
    
    def extract_recursive_advanced(expr: str, depth: int = 0, start_pos: int = 0, 
                                   parent_indices: List[int] = None, path: List[str] = None):
        if parent_indices is None:
            parent_indices = []
        if path is None:
            path = []
        
        pattern_str = r'(@?)(' + '|'.join(DISTRIBUTION_FUNCTIONS) + r')\s*\('
        pattern = re.compile(pattern_str, re.IGNORECASE)
        
        pos = 0
        while pos < len(expr):
            match = pattern.search(expr, pos)
            if not match:
                break
            
            at_symbol = match.group(1)
            func_name = match.group(2)
            func_start_in_expr = match.start()
            args_start = match.end()
            
            # 匹配括号
            paren_stack = 1
            args_end = args_start
            
            while paren_stack > 0 and args_end < len(expr):
                if expr[args_end] == '(':
                    paren_stack += 1
                elif expr[args_end] == ')':
                    paren_stack -= 1
                args_end += 1
            
            if paren_stack == 0:
                args_text = expr[args_start:args_end-1]
                full_match = expr[func_start_in_expr:args_end]
                
                global_start = start_pos + func_start_in_expr
                global_end = start_pos + args_end
                
                # 为当前函数生成唯一ID
                func_index = len(all_functions) + 1
                
                # 解析参数获取分布参数
                try:
                    args_list = parse_args_with_nested_functions(args_text)
                    dist_params, markers = extract_dist_params_and_markers(args_list)
                except:
                    dist_params, markers = [], {}
                
                # 构建函数信息
                func_info = {
                    'func_name': func_name,
                    'args_text': args_text,
                    'full_match': full_match,
                    'is_at_function': bool(at_symbol),
                    'is_nested': depth > 0,
                    'depth': depth,
                    'start_pos': global_start,
                    'end_pos': global_end,
                    'cell_addr': cell_addr,
                    'index': func_index,
                    'key': f"{cell_addr}_{func_index}",
                    'parameters': dist_params,
                    'markers': markers,
                    'parent_indices': parent_indices.copy(),  # 父函数索引链
                    'nested_path': path.copy(),  # 嵌套路径
                    'global_start': global_start,
                    'global_end': global_end
                }
                
                all_functions.append(func_info)
                
                # 递归提取内层函数
                current_path = path + [func_name]
                extract_recursive_advanced(
                    args_text, 
                    depth + 1, 
                    global_start + (args_start - func_start_in_expr),
                    parent_indices + [func_index],
                    current_path
                )
                
                pos = args_end
            else:
                pos = match.end()
    
    extract_recursive_advanced(formula_body)
    
    # 按全局起始位置排序
    all_functions.sort(key=lambda x: x.get('global_start', 0))
    
    # 重新分配索引，确保顺序正确
    for i, func in enumerate(all_functions):
        func['index'] = i + 1
        func['key'] = f"{cell_addr}_{i+1}"
    
    return all_functions

# ==================== Simtable函数解析 ====================

def extract_simtable_functions(formula: str) -> List[Dict[str, Any]]:
    """
    提取公式中的所有Simtable函数
    返回格式：[{'func_name': 'DriskSimtable', 'full_match': 'DriskSimtable(1,2,3)', 'args_text': '1,2,3'}, ...]
    
    修改：支持单元格区域引用（如E17:E19）和数组常量（如{0.8,1.0,1.2}）
    """
    if not isinstance(formula, str) or not formula.startswith('='):
        return []
    
    formula_body = formula[1:]
    simtable_functions = []
    
    # 查找所有Simtable函数
    pattern = r'(DriskSimtable)\s*\(([^)]+)\)'
    matches = list(re.finditer(pattern, formula_body, re.IGNORECASE))
    
    for match in matches:
        func_name = match.group(1)
        args_text = match.group(2)
        full_match = match.group(0)
        
        # 计算全局位置
        global_start = match.start()
        global_end = match.end()
        
        # 解析参数值
        try:
            # 将参数文本分割成列表，处理可能的花括号和逗号
            args_text_clean = args_text.strip()
            
            # 处理花括号语法：{1,2,3}（Excel数组常量）
            if args_text_clean.startswith('{') and args_text_clean.endswith('}'):
                # 数组常量，提取内部值
                inner_text = args_text_clean[1:-1].strip()
                values = []
                
                # 分割逗号，但要注意可能有空格和嵌套（虽然数组常量通常不会嵌套）
                if ',' in inner_text:
                    parts = [p.strip() for p in inner_text.split(',') if p.strip()]
                else:
                    parts = [inner_text] if inner_text else []
                
                # 将每个部分转换为数值
                for part in parts:
                    try:
                        val = float(part)
                        values.append(val)
                    except:
                        # 如果不是数值，尝试作为字符串
                        try:
                            val = float(part)
                            values.append(val)
                        except:
                            # 无法转换，跳过
                            pass
                
                simtable_functions.append({
                    'func_name': func_name,
                    'args_text': args_text,
                    'full_match': full_match,
                    'start_pos': global_start,
                    'end_pos': global_end,
                    'values': values,
                    'value_count': len(values),
                    'is_array_constant': True,  # 标记为数组常量
                    'is_range': False
                })
                continue  # 处理完毕，继续下一个
                
            # 检查是否包含区域引用（如E17:E19）
            range_pattern = r'([A-Z]{1,3}\d{1,7})\s*:\s*([A-Z]{1,3}\d{1,7})'
            range_match = re.search(range_pattern, args_text_clean, re.IGNORECASE)
            
            if range_match:
                # 区域引用，标记为需要后续处理
                simtable_functions.append({
                    'func_name': func_name,
                    'args_text': args_text,
                    'full_match': full_match,
                    'start_pos': global_start,
                    'end_pos': global_end,
                    'values': [],  # 值将在后续处理中填充
                    'value_count': 0,
                    'range_reference': args_text_clean,  # 区域引用
                    'is_range': True,
                    'is_array_constant': False
                })
            else:
                # 常规参数（逗号分隔的数值或引用）
                if ',' in args_text_clean:
                    arg_list = [arg.strip() for arg in args_text_clean.split(',') if arg.strip()]
                else:
                    # 可能是单个值
                    arg_list = [args_text_clean] if args_text_clean else []
                
                # 尝试将参数转换为数值
                values = []
                for arg in arg_list:
                    try:
                        # 移除可能的引号
                        arg_clean = arg.strip().strip('"\'')
                        # 尝试转换为数值
                        val = float(arg_clean)
                        values.append(val)
                    except:
                        # 如果不是数值，保持原样（可能是单元格引用）
                        values.append(arg_clean)
                
                simtable_functions.append({
                    'func_name': func_name,
                    'args_text': args_text,
                    'full_match': full_match,
                    'start_pos': global_start,
                    'end_pos': global_end,
                    'values': values,
                    'value_count': len(values),
                    'is_range': False,
                    'is_array_constant': False
                })
                
        except Exception:
            # 解析失败，返回空值
            simtable_functions.append({
                'func_name': func_name,
                'args_text': args_text,
                'full_match': full_match,
                'start_pos': global_start,
                'end_pos': global_end,
                'values': [],
                'value_count': 0,
                'is_range': False,
                'is_array_constant': False
            })
    
    return simtable_functions

def extract_simtable_array(formula: str) -> List[Any]:
    """
    从Simtable函数中提取数组值
    返回：值列表，如果没有Simtable函数返回空列表
    """
    simtable_funcs = extract_simtable_functions(formula)
    if not simtable_funcs:
        return []
    
    # 返回第一个Simtable函数的值
    return simtable_funcs[0].get('values', [])

def get_simtable_value_at_index(formula: str, index: int) -> Any:
    """
    获取Simtable函数在指定索引处的值
    如果索引超出范围，返回None
    """
    values = extract_simtable_array(formula)
    if 0 <= index < len(values):
        return values[index]
    return None

# ==================== MakeInput函数解析 ====================

def extract_makeinput_functions(formula: str) -> List[Dict[str, Any]]:
    """
    提取公式中的所有MakeInput函数和内嵌分布函数
    返回格式：[{'func_name': 'DriskMakeInput', 'full_match': '...', 'args_text': '...', 
                'formula': '...', 'attributes': {...}, 'nested_distributions': [...]}]
    
    修改：现在会提取MakeInput中的内嵌分布函数
    """
    if not isinstance(formula, str) or not formula.startswith('='):
        return []
    
    formula_body = formula[1:]
    makeinput_functions = []
    
    # 查找所有MakeInput函数
    pattern = r'(DriskMakeInput)\s*\(([^)]+)\)'
    matches = list(re.finditer(pattern, formula_body, re.IGNORECASE))
    
    for match in matches:
        func_name = match.group(1)
        args_text = match.group(2)
        full_match = match.group(0)
        
        # 计算全局位置
        global_start = match.start()
        global_end = match.end()
        
        try:
            # 解析参数列表
            args_list = parse_args_with_nested_functions(args_text)
            
            # 第一个参数是计算公式
            formula_arg = args_list[0] if args_list else ""
            
            # 提取内嵌分布函数 - 关键修改：从公式参数中提取
            nested_distributions = []
            if formula_arg:
                # 从计算公式中提取所有分布函数
                nested_dists = extract_nested_distributions_advanced(f"={formula_arg}", "")
                for dist_func in nested_dists:
                    nested_distributions.append({
                        'func_name': dist_func.get('func_name', ''),
                        'args_text': dist_func.get('args_text', ''),
                        'full_match': dist_func.get('full_match', ''),
                        'parameters': dist_func.get('parameters', []),
                        'markers': dist_func.get('markers', {}),
                        'is_nested': dist_func.get('is_nested', False),
                        'depth': dist_func.get('depth', 0)
                    })
            
            # 提取属性函数
            attributes = args_list[1:] if len(args_list) > 1 else []
            attr_info = {}
            for attr in attributes:
                marker_info = parse_marker_function(attr)
                if marker_info:
                    marker_type, marker_value = marker_info
                    if marker_type:
                        attr_info[marker_type] = marker_value
            
            makeinput_functions.append({
                'func_name': func_name,
                'args_text': args_text,
                'full_match': full_match,
                'start_pos': global_start,
                'end_pos': global_end,
                'formula': formula_arg,
                'attributes': attr_info,
                'nested_distributions': nested_distributions,  # 新增：内嵌分布列表
                'is_makeinput': True
            })
            
        except Exception as e:
            # 解析失败，返回基本结构
            makeinput_functions.append({
                'func_name': func_name,
                'args_text': args_text,
                'full_match': full_match,
                'start_pos': global_start,
                'end_pos': global_end,
                'formula': '',
                'attributes': {},
                'nested_distributions': [],  # 空列表
                'is_makeinput': True
            })
    
    return makeinput_functions

# ==================== 属性提取函数 ====================

def extract_all_attributes_from_formula(formula: str) -> dict:
    """
    从公式中提取所有属性
    增强版：支持所有属性函数
    """
    if not isinstance(formula, str) or not formula.startswith('='):
        return {}
    
    attributes = {}
    formula_body = formula[1:]
    
    # 定义属性函数正则表达式模式
    attr_patterns = {
        'name': r'DriskName\s*\(\s*"([^"]*)"\s*\)',
        'category': r'DriskCategory\s*\(\s*"([^"]*)"\s*\)',
        'units': r'DriskUnits\s*\(\s*"([^"]*)"\s*\)',
        'is_date': r'DriskIsDate\s*\(\s*([^)]*)\s*\)',
        'is_discrete': r'DriskIsDiscrete\s*\(\s*([^)]*)\s*\)',
        'static': r'DriskStatic\s*\(\s*([-+]?\d*\.?\d+)\s*\)',
        'loc': r'DriskLoc\s*\(\s*\)',
        'seed': r'DriskSeed\s*\(\s*([^)]*)\s*\)',
        'lock': r'DriskLock\s*\(\s*([^)]*)\s*\)',
        'shift': r'DriskShift\s*\(\s*([-+]?\d*\.?\d+)\s*\)',
        'collect': r'DriskCollect\s*\(\s*([^)]*)\s*\)',
        'convergence': r'DriskConvergence\s*\(\s*([^)]*)\s*\)',
        'copula': r'DriskCopula\s*\(\s*([^)]*)\s*\)',
        'corrmat': r'DriskCorrmat\s*\(\s*([^)]*)\s*\)',
        'fit': r'DriskFit\s*\(\s*([^)]*)\s*\)',
        'truncate': r'DriskTruncate\s*\(\s*([^)]*)\s*\)',
        'truncate_p': r'DriskTruncateP\s*\(\s*([^)]*)\s*\)',
        'truncate2': r'DriskTruncate2\s*\(\s*([^)]*)\s*\)',
        'truncate_p2': r'DriskTruncateP2\s*\(\s*([^)]*)\s*\)',
        'output': r'DriskOutput\s*\(\s*([^)]*)\s*\)',
        'makeinput': r'DriskMakeInput\s*\(\s*([^)]*)\s*\)'  # 新增
    }
    
    # 检查每个属性函数
    for attr_type, pattern in attr_patterns.items():
        matches = re.findall(pattern, formula_body, re.IGNORECASE)
        if matches:
            if attr_type == 'static':
                try:
                    attributes['static'] = float(matches[0].strip().strip('"\''))
                except:
                    attributes['static'] = True
            elif attr_type == 'loc':
                attributes['loc'] = True
            elif attr_type == 'is_date':
                value = matches[0].strip().lower()
                attributes['is_date'] = value == 'true' or value == ''
            elif attr_type == 'is_discrete':
                value = matches[0].strip().lower()
                attributes['is_discrete'] = value == 'true' or value == ''
            elif attr_type == 'name':
                attributes['name'] = matches[0].strip().strip('"\'')
            elif attr_type == 'category':
                attributes['category'] = matches[0].strip().strip('"\'')
            elif attr_type == 'units':
                attributes['units'] = matches[0].strip().strip('"\'')
            elif attr_type == 'seed':
                # 解析种子参数
                seed_parts = matches[0].split(',')
                if len(seed_parts) >= 2:
                    try:
                        attributes['rng_type'] = int(seed_parts[0].strip())
                        attributes['seed'] = int(seed_parts[1].strip())
                    except:
                        attributes['rng_type'] = 1
                        attributes['seed'] = 42
            elif attr_type == 'output':
                # 解析输出参数
                output_parts = matches[0].split(',')
                if len(output_parts) >= 3:
                    attributes['output_name'] = output_parts[0].strip().strip('"\'')
                    attributes['output_category'] = output_parts[1].strip().strip('"\'')
                    try:
                        attributes['output_position'] = int(output_parts[2].strip())
                    except:
                        attributes['output_position'] = 1
                elif len(output_parts) >= 1:
                    if output_parts[0].strip():
                        attributes['output_name'] = output_parts[0].strip().strip('"\'')
                    if len(output_parts) >= 2 and output_parts[1].strip():
                        attributes['output_category'] = output_parts[1].strip().strip('"\'')
            elif attr_type == 'makeinput':
                # 解析MakeInput参数
                makeinput_parts = matches[0].split(',')
                if len(makeinput_parts) >= 1:
                    attributes['makeinput_formula'] = makeinput_parts[0].strip().strip('"\'')
            elif attr_type in ['lock', 'shift', 'collect', 'convergence', 
                              'copula', 'corrmat', 'fit', 'truncate', 
                              'truncate_p', 'truncate2', 'truncate_p2']:
                # 其他属性，标记为True
                attributes[attr_type] = True
    
    return attributes

def extract_input_attributes(formula: str) -> Dict[str, Any]:
    """
    从分布函数公式中提取Input属性信息
    返回格式：{'name': '...', 'units': '...', 'category': '...', 'is_date': False, 'is_discrete': False, ...}
    """
    if not isinstance(formula, str) or not formula.startswith('='):
        return {}
    
    formula_body = formula[1:]
    attributes = {}
    
    # 提取所有属性函数
    # DriskName
    name_match = re.search(r'DriskName\s*\(\s*"([^)]*)"\s*\)', formula_body, re.IGNORECASE)
    if name_match:
        attributes['name'] = name_match.group(1).strip().strip('"\'')
    
    # DriskUnits
    units_match = re.search(r'DriskUnits\s*\(\s*"([^"]*)"\s*\)', formula_body, re.IGNORECASE)
    if units_match:
        attributes['units'] = units_match.group(1).strip().strip('"\'')
    
    # DriskCategory
    category_match = re.search(r'DriskCategory\s*\(\s*"([^)]*)"\s*\)', formula_body, re.IGNORECASE)
    if category_match:
        attributes['category'] = category_match.group(1).strip().strip('"\'')
    
    # DriskIsDate
    isdate_match = re.search(r'DriskIsDate\s*\(\s*([^)]*)\s*\)', formula_body, re.IGNORECASE)
    if isdate_match:
        isdate_value = isdate_match.group(1).strip().strip('"\'').lower()
        attributes['is_date'] = isdate_value == 'true' or isdate_value == ''
    
    # DriskIsDiscrete
    isdiscrete_match = re.search(r'DriskIsDiscrete\s*\(\s*([^)]*)\s*\)', formula_body, re.IGNORECASE)
    if isdiscrete_match:
        isdiscrete_value = isdiscrete_match.group(1).strip().strip('"\'').lower()
        attributes['is_discrete'] = isdiscrete_value == 'true' or isdiscrete_value == ''
    
    # DriskStatic
    static_match = re.search(r'DriskStatic\s*\(\s*([-+]?\d*\.?\d+)\s*\)', formula_body, re.IGNORECASE)
    if static_match:
        try:
            attributes['static'] = float(static_match.group(1).strip())
        except:
            attributes['static'] = 0.0
    
    # DriskSeed
    seed_match = re.search(r'DriskSeed\s*\(\s*([^)]*)\s*\)', formula_body, re.IGNORECASE)
    if seed_match:
        seed_parts = seed_match.group(1).split(',')
        if len(seed_parts) >= 2:
            try:
                attributes['rng_type'] = int(seed_parts[0].strip())
                attributes['seed'] = int(seed_parts[1].strip())
            except:
                attributes['rng_type'] = 1
                attributes['seed'] = 42
    
    # DriskLoc
    if re.search(r'DriskLoc\s*\(\s*\)', formula_body, re.IGNORECASE):
        attributes['loc'] = True
    
    # DriskCollect
    if re.search(r'DriskCollect', formula_body, re.IGNORECASE):
        attributes['collect'] = True
    
    # DriskConvergence
    convergence_match = re.search(r'DriskConvergence\s*\(\s*"([^)]*)"\s*\)', formula_body, re.IGNORECASE)
    if convergence_match:
        attributes['convergence'] = convergence_match.group(1).strip().strip('"\'')
    
    # DriskLock
    if re.search(r'DriskLock', formula_body, re.IGNORECASE):
        attributes['lock'] = True
    
    # DriskShift
    shift_match = re.search(r'DriskShift\s*\(\s*([-+]?\d*\.?\d+)\s*\)', formula_body, re.IGNORECASE)
    if shift_match:
        try:
            attributes['shift'] = float(shift_match.group(1).strip())
        except:
            attributes['shift'] = 0.0
    
    return attributes

def extract_makeinput_attributes(formula: str) -> Dict[str, Any]:  # 新增
    """
    从DriskMakeInput公式中提取属性信息
    返回格式：{'formula': '...', 'name': '...', 'units': '...', 'category': '...', 'static': 0.0, 'loc': False}
    """
    if not isinstance(formula, str) or not formula.startswith('='):
        return {}
    
    formula_body = formula[1:]
    attributes = {}
    
    # 提取MakeInput函数
    makeinput_funcs = extract_makeinput_functions(formula)
    if not makeinput_funcs:
        return {}
    
    makeinput_func = makeinput_funcs[0]
    attributes['formula'] = makeinput_func.get('formula', '')
    attributes.update(makeinput_func.get('attributes', {}))
    
    return attributes

def extract_output_info(formula: str) -> Dict[str, Any]:
    """
    从公式中提取DriskOutput信息
    返回格式：{'name': '...', 'category': '...', 'position': 1, 'units': '...', 'is_date': False, 'is_discrete': False, 'convergence': ''}
    """
    if not isinstance(formula, str) or not formula.startswith('='):
        return {}
    
    formula_body = formula[1:]
    
    # 查找DriskOutput函数调用
    pattern = r'DriskOutput\s*\(\s*(?:([^,)]*)\s*,\s*)?(?:([^,)]*)\s*,\s*)?(?:([^,)]*)\s*(?:,\s*(.*))?)?\s*\)'
    match = re.search(pattern, formula_body, re.IGNORECASE)
    
    if not match:
        return {}
    
    groups = match.groups()
    
    # 处理参数
    name = ''
    category = ''
    position = 1
    
    if groups[0] is not None:
        name = groups[0].strip().strip('"\'')
    
    if groups[1] is not None:
        category = groups[1].strip().strip('"\'')
    
    if groups[2] is not None:
        try:
            position_str = groups[2].strip()
            if position_str:
                position = int(position_str)
        except:
            position = 1
    
    # 提取属性函数
    attributes = {}
    
    # 如果有第四个参数（属性函数参数）
    if groups[3] is not None:
        # 提取DriskUnits
        units_match = re.search(r'DriskUnits\s*\(\s*"([^"]*)"\s*\)', groups[3], re.IGNORECASE)
        if units_match:
            attributes['units'] = units_match.group(1).strip().strip('"\'')
        else:
            # 尝试不带引号
            units_match = re.search(r'DriskUnits\s*\(\s*([^)]+)\s*\)', groups[3], re.IGNORECASE)
            if units_match:
                attributes['units'] = units_match.group(1).strip().strip('"\'')
        
        # 提取DriskIsDate
        isdate_match = re.search(r'DriskIsDate\s*\(\s*([^)]*)\s*\)', groups[3], re.IGNORECASE)
        if isdate_match:
            isdate_value = isdate_match.group(1).strip().strip('"\'').lower()
            attributes['is_date'] = isdate_value == 'true' or isdate_value == ''
        
        # 提取DriskIsDiscrete
        isdiscrete_match = re.search(r'DriskIsDiscrete\s*\(\s*([^)]*)\s*\)', groups[3], re.IGNORECASE)
        if isdiscrete_match:
            isdiscrete_value = isdiscrete_match.group(1).strip().strip('"\'').lower()
            attributes['is_discrete'] = isdiscrete_value == 'true' or isdiscrete_value == ''
        
        # 提取DriskConvergence
        convergence_match = re.search(r'DriskConvergence\s*\(\s*"([^"]*)"\s*\)', groups[3], re.IGNORECASE)
        if convergence_match:
            attributes['convergence'] = convergence_match.group(1).strip().strip('"\'')
        else:
            # 尝试不带引号
            convergence_match = re.search(r'DriskConvergence\s*\(\s*([^)]+)\s*\)', groups[3], re.IGNORECASE)
            if convergence_match:
                attributes['convergence'] = convergence_match.group(1).strip().strip('"\'')
    
    # 合并结果
    result = {
        'name': name,
        'category': category,
        'position': position
    }
    
    # 添加属性
    if 'units' in attributes:
        result['units'] = attributes['units']
    if 'is_date' in attributes:
        result['is_date'] = attributes['is_date']
    if 'is_discrete' in attributes:
        result['is_discrete'] = attributes['is_discrete']
    if 'convergence' in attributes:
        result['convergence'] = attributes['convergence']
    
    return result

def extract_output_attributes(formula_body: str) -> Dict[str, Any]:
    """从公式中提取Output的属性函数信息"""
    attributes = {
        'units': '',
        'is_date': False,
        'is_discrete': False,
        'convergence': ''
    }
    
    # 提取DriskUnits
    units_match = re.search(r'DriskUnits\s*\(\s*"([^)]*)"\s*\)', formula_body, re.IGNORECASE)
    if units_match:
        attributes['units'] = units_match.group(1).strip().strip('"\'')
    
    # 提取DriskIsDate
    isdate_match = re.search(r'DriskIsDate\s*\(\s*([^)]*)\s*\)', formula_body, re.IGNORECASE)
    if isdate_match:
        isdate_value = isdate_match.group(1).strip().strip('"\'').lower()
        attributes['is_date'] = isdate_value == 'true' or isdate_value == ''
    
    # 提取DriskIsDiscrete
    isdiscrete_match = re.search(r'DriskIsDiscrete\s*\(\s*([^)]*)\s*\)', formula_body, re.IGNORECASE)
    if isdiscrete_match:
        isdiscrete_value = isdiscrete_match.group(1).strip().strip('"\'').lower()
        attributes['is_discrete'] = isdiscrete_value == 'true' or isdiscrete_value == ''
    
    # 提取DriskConvergence
    convergence_match = re.search(r'DriskConvergence\s*\(\s*"([^"]*)"\s*\)', formula_body, re.IGNORECASE)
    if convergence_match:
        attributes['convergence'] = convergence_match.group(1).strip().strip('"\'')
    
    return attributes

# ==================== 公式处理函数 ====================

def remove_output_function_from_formula(formula: str) -> str:
    """
    从公式中移除DriskOutput函数，只保留计算部分
    例如：=DriskOutput("name","category",1)+DriskNormal(5,2)+100
    返回：=DriskNormal(5,2)+100
    
    修复版：正确处理各种情况
    """
    if not isinstance(formula, str) or not formula.startswith('='):
        return formula
    
    formula_body = formula[1:]
    
    # 查找DriskOutput函数并移除
    # 匹配 DriskOutput(...) 可能后面跟着运算符
    pattern = r'DriskOutput\s*\([^)]*\)\s*([\+\-\*/])?'
    
    # 首先检查是否有DriskOutput
    if not re.search(r'DriskOutput', formula_body, re.IGNORECASE):
        return formula
    
    # 移除DriskOutput函数
    new_body = re.sub(pattern, '', formula_body, flags=re.IGNORECASE)
    
    # 清理可能的运算符问题
    new_body = re.sub(r'^\s*\+\s*', '', new_body)  # 移除开头的+
    new_body = re.sub(r'^\s*\-\s*', '-', new_body)  # 保留开头的-，但去掉多余空格
    new_body = re.sub(r'^\s*\*\s*', '*', new_body)  # 保留开头的*，但去掉多余空格
    new_body = re.sub(r'^\s*/\s*', '/', new_body)  # 保留开头的/，但去掉多余空格
    
    # 如果公式为空，返回=0
    if not new_body.strip():
        return '=0'
    
    # 确保以运算符开头的公式是有效的
    if new_body.startswith('+') or new_body.startswith('-') or new_body.startswith('*') or new_body.startswith('/'):
        # 如果以运算符开头，在前面加0
        new_body = '0' + new_body
    
    return '=' + new_body

def remove_makeinput_function_from_formula(formula: str) -> str:  # 新增
    """
    从公式中移除DriskMakeInput函数，只保留计算部分
    例如：=DriskMakeInput(DriskNormal(4,2)+C1, DriskName("输入"))
    返回：=DriskNormal(4,2)+C1
    
    注意：DriskMakeInput的第一个参数是计算公式，所以移除函数后保留第一个参数
    """
    if not isinstance(formula, str) or not formula.startswith('='):
        return formula
    
    formula_body = formula[1:]
    
    # 查找DriskMakeInput函数
    pattern = r'DriskMakeInput\s*\(([^)]+)\)'
    match = re.search(pattern, formula_body, re.IGNORECASE)
    
    if not match:
        return formula
    
    args_text = match.group(1).strip()
    
    # 解析参数，获取第一个参数（计算公式）
    args_list = parse_args_with_nested_functions(args_text)
    if not args_list:
        return '=0'
    
    # 第一个参数是计算公式
    calc_formula = args_list[0].strip()
    
    # 如果计算公式以等号开头，去掉等号
    if calc_formula.startswith('='):
        calc_formula = calc_formula[1:]
    
    # 返回计算公式
    return '=' + calc_formula

def extract_calculation_part(formula: str) -> str:
    """
    提取公式中的计算部分（移除DriskOutput标记）
    用于计算输出单元格的值
    """
    if not isinstance(formula, str) or not formula.startswith('='):
        return formula
    
    # 直接移除DriskOutput函数
    calc_formula = remove_output_function_from_formula(formula)
    
    # 如果结果为空或只有=，返回=0
    if calc_formula == '=' or not calc_formula.strip():
        return '=0'
    
    return calc_formula

def extract_vectorizable_formula(formula: str) -> str:
    """
    提取适合向量化计算的公式部分
    """
    if not isinstance(formula, str) or not formula.startswith('='):
        return formula
    
    # 移除DriskOutput函数
    formula = remove_output_function_from_formula(formula)
    
    # 移除等号
    if formula.startswith('='):
        formula = formula[1:]
    
    return formula

def is_vectorizable_formula(formula: str) -> bool:
    """
    检查公式是否适合向量化计算
    """
    if not isinstance(formula, str) or not formula.startswith('='):
        return True
    
    formula_upper = formula.upper()
    
    # 检查是否包含复杂函数
    complex_functions = [
        'OFFSET', 'INDIRECT', 'VLOOKUP', 'HLOOKUP', 
        'INDEX', 'MATCH', 'IF', 'IFERROR', 'IFNA',
        'CHOOSE', 'ROW', 'COLUMN', 'ADDRESS'
    ]
    
    for func in complex_functions:
        if func + '(' in formula_upper:
            return False
    
    return True

# ==================== 分布参数解析 ====================

def extract_distribution_params_from_formula(formula_body: str, dist_type: str) -> List[str]:
    """从公式中提取分布参数（简化版本）"""
    try:
        pattern = rf'{dist_type}\s*\(([^)]+)\)'
        match = re.search(pattern, formula_body, re.IGNORECASE)
        if not match:
            return []
        
        params_text = match.group(1)
        # 简单分割参数
        params = [p.strip() for p in params_text.split(',')]
        params = [p for p in params if p]
        
        return params
    except Exception:
        return []

# ==================== 辅助函数 ====================

def parse_range_string(range_str: str):
    """解析范围字符串"""
    if not range_str:
        return None, None
    
    try:
        parts = [p.strip() for p in range_str.split(',') if p.strip() != ""]
        if len(parts) == 1:
            v = float(parts[0])
            return v, v
        elif len(parts) >= 2:
            low = float(parts[0])
            high = float(parts[1])
            return low, high
    except Exception:
        pass
    
    try:
        v = float(range_str)
        return v, v
    except Exception:
        return None, None

def extract_nested_distributions_from_makeinput(formula: str, cell_addr: str) -> List[Dict]:
    """
    从DriskMakeInput函数中提取内嵌分布函数
    
    Args:
        formula: 完整的Excel公式字符串
        cell_addr: 单元格地址（带工作表名）
        
    Returns:
        内嵌分布函数列表，每个元素包含分布函数信息
    """
    try:
        nested_dists = []
        
        # 检查是否为MakeInput函数
        if not is_makeinput_function(formula):
            return nested_dists
        
        # 提取MakeInput函数的参数
        formula_clean = formula.strip()
        if formula_clean.upper().startswith('=DRISKMAKEINPUT('):
            # 提取括号内的内容
            start_idx = formula_clean.find('(') + 1
            end_idx = formula_clean.rfind(')')
            if end_idx == -1:
                end_idx = len(formula_clean)
            
            inner_content = formula_clean[start_idx:end_idx].strip()
            
            # 提取第一个参数（公式）
            # 我们需要解析第一个参数，直到遇到第一个逗号或结束
            # 但注意：公式本身可能包含逗号，所以需要更复杂的解析
            
            # 使用parse_args_with_nested_functions解析所有参数
            args = parse_args_with_nested_functions(inner_content)
            if not args:
                return nested_dists
            
            # 第一个参数是表达式
            expression = args[0] if args else ''
            
            # 从表达式中提取分布函数
            # 使用增强的嵌套函数提取
            nested_dists = extract_nested_distributions_advanced(f"={expression}", cell_addr)
            
            # 标记这些函数来自MakeInput
            for dist in nested_dists:
                dist['is_makeinput_nested'] = True
                dist['parent_makeinput'] = cell_addr
            
            return nested_dists
        
        return nested_dists
        
    except Exception as e:
        return []

def extract_distribution_function_info(dist_text: str, parent_cell_addr: str) -> Dict:
    """
    提取分布函数的详细信息
    
    Args:
        dist_text: 分布函数文本（如 "DriskNormal(5, 2)"）
        parent_cell_addr: 父单元格地址
        
    Returns:
        分布函数信息字典
    """
    try:
        # 提取函数名
        import re
        func_match = re.match(r'([A-Za-z]+)\s*\(', dist_text)
        if not func_match:
            return None
        
        func_name = func_match.group(1)
        
        # 提取参数
        args_start = dist_text.find('(') + 1
        args_end = dist_text.rfind(')')
        args_text = dist_text[args_start:args_end].strip()
        
        # 解析参数
        args = parse_args_with_nested_functions(args_text)
        
        # 构建分布函数信息
        dist_info = {
            'function_name': func_name,
            'func_name': func_name,
            'args_text': args_text,
            'args': args,
            'original_text': dist_text,
            'full_match': dist_text,
            'cell_addr': parent_cell_addr,
            'parent_cell_addr': parent_cell_addr
        }
        
        return dist_info
        
    except Exception as e:
        return None

# ==================== 统一的公式解析器类 ====================

class FormulaParser:
    """统一的公式解析器类"""
    
    def __init__(self, formula: str):
        self.formula = formula
        self.parsed_info = {}
        
    def parse(self) -> Dict[str, Any]:
        """统一解析公式"""
        if not self.formula or not self.formula.startswith('='):
            return {}
        
        # 解析基本信息
        func_name, args_list = parse_complete_formula(self.formula)
        
        # 提取分布参数和标记
        dist_params, markers = extract_dist_params_and_markers(args_list)
        
        # 提取所有分布函数
        dist_functions = extract_all_distribution_functions(self.formula)
        
        # 提取所有Simtable函数
        simtable_functions = extract_simtable_functions(self.formula)
        
        # 提取所有MakeInput函数
        makeinput_functions = extract_makeinput_functions(self.formula)  # 新增
        
        # 提取所有属性
        attributes = extract_all_attributes_from_formula(self.formula)
        
        # 提取单元格引用
        cell_refs = parse_formula_references(self.formula)
        
        # 构建解析结果
        self.parsed_info = {
            'func_name': func_name,
            'args_list': args_list,
            'dist_params': dist_params,
            'markers': markers,
            'dist_functions': dist_functions,
            'simtable_functions': simtable_functions,
            'makeinput_functions': makeinput_functions,  # 新增
            'attributes': attributes,
            'cell_refs': cell_refs,
            'is_distribution': is_distribution_function(self.formula),
            'is_simtable': is_simtable_function(self.formula),
            'is_makeinput': is_makeinput_function(self.formula),  # 新增
            'is_output': is_output_cell(self.formula),
            'has_static': has_static_attribute(self.formula)
        }
        
        return self.parsed_info
