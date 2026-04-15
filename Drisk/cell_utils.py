# cell_utils.py
"""单元格地址处理工具"""

import re
from typing import Any

def extract_address_from_cell_object(cell_obj) -> str:
    """从任何类型的单元格对象中提取地址字符串（如'B13'）"""
    try:
        # 1. 如果是字符串，直接返回
        if isinstance(cell_obj, str):
            # 移除$符号和工作表名
            addr = cell_obj.replace('$', '')
            if '!' in addr:
                addr = addr.split('!')[-1]
            return addr.upper()
        
        # 2. 如果是XLCell对象（PyXLL传入的参数）
        if cell_obj.__class__.__name__ == 'XLCell':
            try:
                # 尝试使用address属性
                if hasattr(cell_obj, 'address'):
                    addr = str(cell_obj.address)
                    addr = addr.replace('$', '')
                    if '!' in addr:
                        addr = addr.split('!')[-1]
                    if re.match(r'^[A-Z]{1,3}\d+$', addr, re.IGNORECASE):
                        return addr.upper()
            except:
                pass
        
        # 3. 如果是Range对象（通过COM接口）
        try:
            if hasattr(cell_obj, 'Address'):
                addr = str(cell_obj.Address)
                addr = addr.replace('$', '')
                if '!' in addr:
                    addr = addr.split('!')[-1]
                if re.match(r'^[A-Z]{1,3}\d+$', addr, re.IGNORECASE):
                    return addr.upper()
        except:
            pass
        
        # 4. 最后尝试字符串表示
        try:
            s = str(cell_obj)
            # 尝试提取地址
            match = re.search(r'([A-Z]{1,3}\d+)', s, re.IGNORECASE)
            if match:
                return match.group(1).upper()
        except:
            pass
        
        # 5. 返回原始字符串表示
        return str(cell_obj)
    except Exception:
        return ""

def normalize_cell_address(cell_input) -> str:
    """规范化单元格地址"""
    try:
        # 提取地址字符串
        addr = extract_address_from_cell_object(cell_input)
        
        # 验证地址格式
        if re.match(r'^[A-Z]{1,3}\d+$', addr, re.IGNORECASE):
            return addr.upper()
        else:
            # 尝试从字符串中提取地址
            match = re.search(r'([A-Z]{1,3}\d+)', addr, re.IGNORECASE)
            if match:
                return match.group(1).upper()
            else:
                return addr
    except Exception:
        return str(cell_input)