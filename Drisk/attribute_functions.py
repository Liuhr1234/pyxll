# attribute_functions.py
"""属性函数模块 - 统一字符串标记版本"""

import re
import threading
from typing import Dict, List, Tuple, Any, Union
from pyxll import xl_func

# ==================== 全局常量 ====================
# 标记字符串前缀（确保唯一性）
_MARKER_PREFIXES = {
    'name': "__DRISK_NAME__:",
    'loc': "__DRISK_LOC__:",
    'category': "__DRISK_CATEGORY__:",
    'collect': "__DRISK_COLLECT__:",
    'convergence': "__DRISK_CONVERGENCE__:",
    'copula': "__DRISK_COPULA__:",
    'corrmat': "__DRISK_CORRMAT__:",
    'fit': "__DRISK_FIT__:",
    'isdate': "__DRISK_ISDATE__:",
    'isdiscrete': "__DRISK_ISDISCRETE__:",
    'lock': "__DRISK_LOCK__:",
    'seed': "__DRISK_SEED__:",
    'shift': "__DRISK_SHIFT__:",
    'static': "__DRISK_STATIC__:",
    'truncate': "__DRISK_TRUNCATE__:",
    'truncatep': "__DRISK_TRUNCATEP__:",
    'truncate2': "__DRISK_TRUNCATE2__:",
    'truncatep2': "__DRISK_TRUNCATEP2__:",
    'units': "__DRISK_UNITS__:",
    'position': "__DRISK_POSITION__:",
    'makeinput': "__DRISK_MAKEINPUT__:",  # 新增
}

ERROR_MARKER = "#ERROR!"

# ==================== 全局状态 ====================
_STATIC_MODE = True  # True: 显示静态值, False: 模拟模式
_STATIC_LOCK = threading.RLock()

def get_static_mode() -> bool:
    """获取当前静态模式状态"""
    return _STATIC_MODE

def set_static_mode(mode: bool):
    """设置静态模式开关"""
    global _STATIC_MODE
    with _STATIC_LOCK:
        _STATIC_MODE = mode

# ==================== 辅助函数 ====================
def is_marker_string(value) -> bool:
    """检查值是否是标记字符串"""
    if isinstance(value, str):
        for prefix in _MARKER_PREFIXES.values():
            if value.startswith(prefix):
                return True
    return False

def extract_marker_info(value_str: str) -> Tuple[str, str]:
    """从标记字符串中提取信息"""
    if not isinstance(value_str, str):
        return None, None
    
    for marker_type, prefix in _MARKER_PREFIXES.items():
        if value_str.startswith(prefix):
            value = value_str[len(prefix):]
            return marker_type, value
    
    return None, None

def create_marker_string(marker_type: str, value: str = "") -> str:
    """创建标记字符串"""
    if marker_type not in _MARKER_PREFIXES:
        raise ValueError(f"未知的标记类型: {marker_type}")
    
    prefix = _MARKER_PREFIXES[marker_type]
    return f"{prefix}{value}"

# ==================== 属性函数定义 ====================
# 注意：所有函数都返回标记字符串，而不是Marker对象

@xl_func("string name: var", category="Drisk Attributes", volatile=True)
def DriskName(name: str):
    """
    为单元格命名
    
    参数：
      name: 单元格名称（字符串）
    """
    return create_marker_string('name', str(name))

@xl_func(": var", category="Drisk Attributes", volatile=True)
def DriskLoc():
    """
    将分布单元格的静态值锁定为理论均值
    """
    return create_marker_string('loc', "True")

@xl_func("var category_name: var", category="Drisk Attributes", volatile=True)
def DriskCategory(category_name):
    """
    指定输入分布的类别
    
    参数：
      category_name: 类别名称
    """
    return create_marker_string('category', str(category_name))

@xl_func("var collect_mode: var", category="Drisk Attributes", volatile=True)
def DriskCollect(collect_mode=None):
    """
    在模拟过程中收集分布样本数据
    
    参数：
      collect_mode: 收集模式（可选）
    """
    if collect_mode is None:
        return create_marker_string('collect', "all")
    return create_marker_string('collect', str(collect_mode))

@xl_func("var*: var", category="Drisk Attributes", volatile=True)
def DriskConvergence(*args):
    """
    控制收敛性检验方式
    
    参数：
      *args: 收敛性参数
    """
    # 参数转为字符串，用逗号分隔
    args_str = ",".join(str(arg) for arg in args)
    return create_marker_string('convergence', args_str)

@xl_func("var*: var", category="Drisk Attributes", volatile=True)
def DriskCopula(*args):
    """
    标识属于一组相关分布函数
    
    参数：
      *args: Copula参数
    """
    args_str = ",".join(str(arg) for arg in args)
    return create_marker_string('copula', args_str)

@xl_func("var*: var", category="Drisk Attributes", volatile=True)
def DriskCorrmat(*args):
    """
    标识相关分布函数中的特定分布
    
    参数：
      *args: 相关矩阵参数
    """
    args_str = ",".join(str(arg) for arg in args)
    return create_marker_string('corrmat', args_str)

@xl_func("var*: var", category="Drisk Attributes", volatile=True)
def DriskFit(*args):
    """
    将数据集及其拟合结果链接到输入分布
    
    参数：
      *args: 拟合参数
    """
    args_str = ",".join(str(arg) for arg in args)
    return create_marker_string('fit', args_str)

@xl_func("var is_date: var", category="Drisk Attributes", volatile=True)
def DriskIsDate(is_date=True):
    """
    指定变量是否为日期类型
    
    参数：
      is_date: 是否为日期（默认True）
    """
    return create_marker_string('isdate', str(bool(is_date)))

@xl_func("var is_discrete: var", category="Drisk Attributes", volatile=True)
def DriskIsDiscrete(is_discrete=True):
    """
    指定变量是否为离散类型
    
    参数：
      is_discrete: 是否为离散（默认True）
    """
    return create_marker_string('isdiscrete', str(bool(is_discrete)))

@xl_func("var lock_mode: var", category="Drisk Attributes", volatile=True)
def DriskLock(lock_mode=None):
    """
    控制分布的抽样方式
    
    参数：
      lock_mode: 锁定模式（可选）
    """
    if lock_mode is None:
        return create_marker_string('lock', "default")
    elif lock_mode == "True":
        return create_marker_string('lock', "True")
    return create_marker_string('lock', "True")

@xl_func("var*: var", category="Drisk Attributes", volatile=True)
def DriskSeed(*args):
    """
    指定随机数生成器和种子值
    
    参数：
      rng_type: 随机数生成器类型（可选，第一个参数）
      seed: 种子值（可选，第二个参数）
    """
    if len(args) == 0:
        return create_marker_string('seed', "1,42")  # 默认值
    elif len(args) == 1:
        try:
            seed = int(float(args[0]))
            return create_marker_string('seed', f"1,{seed}")
        except:
            return create_marker_string('seed', "1,42")
    elif len(args) >= 2:
        try:
            rng_type = int(float(args[0]))
            seed = int(float(args[1]))
            return create_marker_string('seed', f"{rng_type},{seed}")
        except:
            return create_marker_string('seed', "1,42")

@xl_func("var shift: var", category="Drisk Attributes", volatile=True)
def DriskShift(shift):
    """
    使分布的所有抽样值按指定量进行移位
    
    参数：
      shift: 移位量
    """
    try:
        return create_marker_string('shift', str(float(shift)))
    except:
        return create_marker_string('shift', "0.0")

@xl_func("var static_value: var", category="Drisk Attributes", volatile=True)
def DriskStatic(static_value):
    """
    指定输入分布的静态值
    
    参数：
      static_value: 静态值
    """
    try:
        value = float(static_value)
        return create_marker_string('static', str(value))
    except:
        return create_marker_string('static', "0.0")

@xl_func("var*: var", category="Drisk Attributes", volatile=True)
def DriskTruncate(*args):
    """
    对输入分布进行截断处理
    
    参数：
      *args: 截断参数
    """
    args_str = ",".join(str(arg) for arg in args)
    return create_marker_string('truncate', args_str)

@xl_func("var lower, var upper: var", category="Drisk Attributes", volatile=True)
def DriskTruncateP(lower=None, upper=None):
    """
    将抽样范围限制在累积分布函数的区间内
    参数：
      lower: 下界百分比（可选）
      upper: 上界百分比（可选）
    """
    lower_str = "" if lower is None else str(lower)
    upper_str = "" if upper is None else str(upper)
    args_str = f"{lower_str},{upper_str}"
    return create_marker_string('truncatep', args_str)

@xl_func("var*: var", category="Drisk Attributes", volatile=True)
def DriskTruncate2(*args):
    """
    对输入分布进行截断处理
    
    参数：
      *args: 二次截断参数
    """
    args_str = ",".join(str(arg) for arg in args)
    return create_marker_string('truncate2', args_str)
"""
@xl_func("var*: var", category="Drisk Attributes", volatile=True)
def DriskTruncateP2(*args):
    args_str = ",".join(str(arg) for arg in args)
    return create_marker_string('truncatep2', args_str)
"""
@xl_func("var units: var", category="Drisk Attributes", volatile=True)
def DriskUnits(units):
    """
    定义输入或输出的单位
    
    参数：
      units: 单位字符串
    """
    return create_marker_string('units', str(units))

# ==================== 修改：DriskMakeInput函数 ====================
@xl_func("var formula, var 属性1, var 属性2, var 属性3, var 属性4, var 属性5, var*: var", category="Drisk Attributes", volatile=True)
def DriskMakeInput(formula, *attributes):
    """
    DriskMakeInput函数 - 标记当前单元格的值为输入
    
    参数：
      formula: 必须的，计算公式（如DriskNormal(4,2)+C1）
      *attributes: 可选属性函数（DriskName, DriskCategory, DriskLoc, DriskUnits等）
    
    返回：
      静态模式下：返回formula的计算结果
      模拟模式下：formula值会被收集为输入
    
    注意：
      1. 不再支持复杂嵌套形式（如DriskNormal()+DriskMakeInput()+DriskNormal()）
      2. 现在只作为独立的输入标记函数
      3. 在模拟中记录整个单元格的值
      4. 在info中会显示为输入，与分布函数共享编号
    """
    # 检查是否处于静态模式
    static_mode = get_static_mode()

    # 检查是否有静态值属性（DriskStatic）
    static_value = None
    for attr in attributes:
        if is_marker_string(attr):
            marker_type, marker_value = extract_marker_info(attr)
            if marker_type == 'static':
                try:
                    static_value = float(marker_value)
                except:
                    static_value = None
                break

    # 静态模式下，如果有静态值，返回静态值；否则返回公式计算结果（尝试转换为数值以减少后续错误）
    if static_mode:
        if static_value is not None:
            return static_value
        # 否则让Excel计算公式并透传结果
        try:
            # 如果传入的是类似数字的字符串，尝试转换
            if isinstance(formula, str):
                f = formula.strip()
                try:
                    return float(f)
                except:
                    return formula
            return formula
        except Exception:
            return formula

    # 模拟模式下，检查是否有loc标记（由解析器处理理论均值）
    loc_enabled = False
    for attr in attributes:
        if is_marker_string(attr):
            marker_type, _ = extract_marker_info(attr)
            if marker_type == 'loc':
                loc_enabled = True
                break

    if loc_enabled:
        # 有loc标记时，仍然返回当前单元格的计算结果（引擎会处理loc逻辑）
        try:
            if isinstance(formula, str):
                f = formula.strip()
                try:
                    return float(f)
                except:
                    return formula
            return formula
        except Exception:
            return formula

    # 模拟模式下，正常返回公式计算结果（尽量返回数值而非字符串）
    try:
        if isinstance(formula, str):
            f = formula.strip()
            try:
                return float(f)
            except:
                return formula
        return formula
    except Exception:
        return formula

# ==================== 简化的DriskOutput函数 ====================
# 关键修改：DriskOutput函数返回0，所有逻辑在模拟引擎中处理
@xl_func("var name, var category, int position, var*: float", category="Drisk Attributes", volatile=True)
def DriskOutput(name="", category="", position=1, *args):
    """
    标记单元格为输出单元格 - 简化版
    
    参数：
      name: 输出名称（字符串）
      category: 输出类别（字符串）
      position: 输出位置（整数）
      *args: 可选属性函数参数（DriskUnits, DriskIsDate, DriskIsDiscrete, DriskConvergence等）
    
    返回：
      总是返回0，不会影响计算结果
    
    说明：
      1. 这个函数只是标记单元格为输出单元格
      2. 在模拟过程中，该单元格的计算结果会被收集
      3. 属性函数参数用于设置输出的属性
      4. 这个函数返回0，以便与其他计算结合使用（如 =@DriskOutput()+@DriskNormal()）
    """
    # 这个函数在静态模式下总是返回0
    # 在模拟模式下，模拟引擎会收集该单元格的计算结果
    return 0

# ==================== 标记处理工具函数 ====================
def extract_markers_from_args(args) -> Tuple[List[Any], Dict[str, Any]]:
    """
    从参数列表中提取标记和普通参数
    
    返回:
        tuple: (normal_params, marker_dict)
    """
    normal_params = []
    markers = {}
    
    for arg in args:
        if is_marker_string(arg):
            marker_type, marker_value = extract_marker_info(arg)
            
            if marker_type:
                # 根据标记类型处理
                if marker_type == 'name':
                    markers['name'] = marker_value
                elif marker_type == 'loc':
                    markers['loc'] = True
                elif marker_type == 'category':
                    markers['category'] = marker_value
                elif marker_type == 'collect':
                    markers['collect'] = marker_value
                elif marker_type == 'convergence':
                    markers['convergence'] = marker_value
                elif marker_type == 'copula':
                    markers['copula'] = True
                elif marker_type == 'corrmat':
                    markers['corrmat'] = True
                elif marker_type == 'fit':
                    markers['fit'] = True
                elif marker_type == 'isdate':
                    markers['is_date'] = marker_value.lower() == 'true'
                elif marker_type == 'isdiscrete':
                    markers['is_discrete'] = marker_value.lower() == 'true'
                elif marker_type == 'lock':
                    markers['lock'] = marker_value
                elif marker_type == 'seed':
                    # 解析种子参数
                    seed_parts = marker_value.split(',')
                    if len(seed_parts) >= 2:
                        try:
                            markers['rng_type'] = int(seed_parts[0])
                            markers['seed'] = int(seed_parts[1])
                        except:
                            markers['rng_type'] = 1
                            markers['seed'] = 42
                    else:
                        markers['rng_type'] = 1
                        markers['seed'] = 42
                elif marker_type == 'shift':
                    try:
                        markers['shift'] = float(marker_value)
                    except:
                        markers['shift'] = 0.0
                elif marker_type == 'static':
                    try:
                        markers['static'] = float(marker_value)
                    except:
                        markers['static'] = 0.0
                elif marker_type == 'truncate':
                    # 保留截断参数字符串，例如 "2000,4000"
                    markers['truncate'] = marker_value
                elif marker_type == 'truncatep':
                    # 保留百分位截断参数字符串，例如 "5,95" 或 "0.05,0.95"
                    markers['truncate_p'] = marker_value
                elif marker_type == 'truncate2':
                    # 先平移后截断，保留参数字符串
                    markers['truncate2'] = marker_value
                elif marker_type == 'truncatep2': 
                    # 先平移后百分位截断，保留参数字符串
                    markers['truncate_p2'] = marker_value
                elif marker_type == 'units':
                    markers['units'] = marker_value
                elif marker_type == 'position':
                    try:
                        markers['position'] = int(marker_value)
                    except:
                        markers['position'] = 1
                elif marker_type == 'makeinput':  # 新增
                    markers['makeinput'] = True
        else:
            normal_params.append(arg)
    
    return normal_params, markers

# 在attribute_functions.py中添加以下函数

def ensure_nested_functions_generated(parent_key: str, current_iter: int):
    """
    确保嵌套函数的父函数已经生成值
    用于在UDF中处理嵌套函数的依赖关系
    """
    try:
        from model_functions import _global_iter_counter
        
        if parent_key in _global_iter_counter.distribution_registry:
            parent_info = _global_iter_counter.distribution_registry[parent_key]
            if not parent_info['has_generated'][current_iter]:
                # 父函数还未生成，需要先生成父函数的值
                pass
    except:
        pass

def extract_markers_from_output_args(name, category, position, attributes) -> Tuple[Dict[str, Any], Dict[str, Any]]:
    """
    从DriskOutput参数中提取标记和普通参数
    专门为DriskOutput设计，处理前三个固定参数和后续属性函数
    
    返回:
        tuple: (output_info, marker_dict)
    """
    # 基础输出信息
    output_info = {
        'name': str(name).strip().strip('"\''),
        'category': str(category).strip().strip('"\''),
        'position': int(position) if position else 1
    }
    
    # 处理属性函数
    marker_dict = {}
    
    for attr in attributes:
        if is_marker_string(attr):
            marker_type, marker_value = extract_marker_info(attr)
            
            if marker_type:
                # 根据标记类型处理
                if marker_type == 'units':
                    marker_dict['units'] = marker_value
                elif marker_type == 'isdate':
                    marker_dict['is_date'] = marker_value.lower() == 'true'
                elif marker_type == 'isdiscrete':
                    marker_dict['is_discrete'] = marker_value.lower() == 'true'
                elif marker_type == 'convergence':
                    marker_dict['convergence'] = marker_value
                elif marker_type == 'position':
                    try:
                        marker_dict['position'] = int(marker_value)
                    except:
                        marker_dict['position'] = 1
                # 其他属性函数也可以在这里处理
                elif marker_type == 'collect':
                    marker_dict['collect'] = marker_value
                elif marker_type == 'lock':
                    marker_dict['lock'] = marker_value
                elif marker_type == 'static':
                    try:
                        marker_dict['static'] = float(marker_value)
                    except:
                        marker_dict['static'] = 0.0
                elif marker_type == 'shift':
                    try:
                        marker_dict['shift'] = float(marker_value)
                    except:
                        marker_dict['shift'] = 0.0
                elif marker_type == 'seed':
                    seed_parts = marker_value.split(',')
                    if len(seed_parts) >= 2:
                        try:
                            marker_dict['rng_type'] = int(seed_parts[0])
                            marker_dict['seed'] = int(seed_parts[1])
                        except:
                            marker_dict['rng_type'] = 1
                            marker_dict['seed'] = 42
    
    # 合并marker_dict中的position到output_info
    if 'position' in marker_dict:
        output_info['position'] = marker_dict['position']
    
    return output_info, marker_dict

def get_static_value_from_formula(formula: str) -> float:
    """
    从公式中提取DriskStatic的静态值
    """
    if not isinstance(formula, str) or not formula.startswith('='):
        return None
    
    formula_body = formula[1:]
    
    # 查找DriskStatic(数值)模式
    static_matches = re.findall(r'DriskStatic\s*\(\s*([-+]?\d*\.?\d+)\s*\)', formula_body, re.IGNORECASE)
    if static_matches:
        try:
            return float(static_matches[0])
        except:
            pass
    
    return None
