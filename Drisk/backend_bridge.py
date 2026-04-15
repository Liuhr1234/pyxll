# backend_bridge.py
"""
本模块提供不可变的 Drisk 后端（PyXLL/Excel 模拟核心）与可变的 PySide6/Plotly UI 层之间的桥接服务。

设计目标（交接必读）：
---------
- 将 *第一组* (后端) 视为以下功能的唯一事实来源 (Single Source of Truth)：
  - 分布函数集（DriskNormal/Uniform/...）
  - 属性标记协议（DriskShift/Truncate/...）
  - 模拟执行（依赖项扫描 + 迭代引擎）
  - 结果缓存（simulation_manager.SimulationResult）
  - 模拟统计数据的过滤/定义（SimulationResult._compute_statistics）

- UI层（第二组）应该：
  - 仅负责创建窗口和渲染图表。
  - 通过本桥接模块，使用 (sim_id, cell_key) 检索数据序列/属性/统计数据。
  - 构建/回写后端能够解析和执行的 Excel 公式。

注意事项：
---------
- 此文件刻意设计为 **完全不依赖** PySide6/Plotly，以保持后端纯净性。
- 严禁跨工作线程传递 COM/Excel 对象，以防引发死锁或崩溃。请暴露纯文本键值 (keys) 或 Numpy 数组 (arrays)。
"""

from __future__ import annotations

from dataclasses import dataclass
import re
from typing import Any, Callable, Dict, List, Optional, Sequence, Tuple, Union

import numpy as np
from input_sample_exposure import filter_exposed_input_keys, is_input_key_exposed

# =======================================================
# 后端核心功能导入区
# =======================================================
import cell_utils
import formula_parser
import dependency_tracker
import simulation_engine
import simulation_manager
import attribute_functions
from constants import DISTRIBUTION_REGISTRY, get_distribution_type


# =======================================================
# Part 1. 核心异常与错误定义
# =======================================================
class BridgeError(RuntimeError):
    """
    桥接层自定义异常。
    当 UI 层请求的操作由于数据缺失、格式错误或后端引擎拒绝而无法执行时，引发此异常。
    """


# =======================================================
# Part 2. 分布函数元数据管理 (Distribution Registry)
# =======================================================
@dataclass(frozen=True)
class DistributionSpec:
    """
    [不可变数据类] 面向 UI 友好的分布元数据，与后端可执行函数保持一致。
    用于在前端生成下拉列表、参数表单，并指导公式生成。
    """
    key: str                          # 简写键值 (如 'normal')
    func_name: str                    # 完整的后端函数名 (如 'DriskNormal')
    display_name: str                 # UI 界面显示的名称
    param_names: Tuple[str, ...]      # 参数名称列表，用于参数表单渲染
    param_labels: Tuple[str, ...]     # UI 界面显示的参数中文标签列表
    defaults: Tuple[float, ...]       # 默认参数值，用于初始化表单
    is_discrete: bool                 # 标识是否为离散分布，影响渲染逻辑（如 PDF vs PMF）
    theory_backend_supported: bool    # 标识 backend_theo 模块是否支持计算其理论概率密度曲线


def get_supported_distributions() -> List[DistributionSpec]:
    """
    分布元数据工厂：
    从底层的常数注册表 (constants.DISTRIBUTION_REGISTRY) 构建并返回前端所需的分布函数规格说明 (specs)。
    """
    specs = []
    # 明确声明后端支持理论计算的分布名单（后续若底层算法扩充，需同步修改此处）
    _THEO_SUPPORTED = {"DriskNormal", "DriskUniform", "DriskGamma", "DriskBeta"}
    
    for func_name, info in DISTRIBUTION_REGISTRY.items():
        specs.append(DistributionSpec(
            key=info.get("type", func_name.replace("Drisk", "").lower()),
            func_name=func_name,
            display_name=info.get("ui_name", info.get("description", func_name)),
            param_names=tuple(info.get("param_names", [])),
            param_labels=tuple(info.get("ui_param_labels", info.get("param_names", []))),
            defaults=tuple(info.get("default_params", [])),
            is_discrete=info.get("is_discrete", False),
            theory_backend_supported=(func_name in _THEO_SUPPORTED)
        ))
    return specs


def get_distribution_spec_by_func_name(func_name: str) -> Optional[DistributionSpec]:
    """
    根据给定的 Excel 公式函数名（例如 'DriskNormal' 或 '@DriskNormal'）
    查找并返回对应的分布元数据规格。未找到则返回 None。
    """
    func_name = (func_name or "").strip().lstrip("@").lower()
    for spec in get_supported_distributions():
        if spec.func_name.lower() == func_name:
            return spec
    return None


# =======================================================
# Part 3. 单元格键值与地址标准化 (Cell Key Normalization)
# =======================================================
def normalize_sheet_name(sheet_name: str) -> str:
    """
    标准化工作表名称：去除多余的引号和空格。
    注：Excel 工作表名称是区分大小写的，因此保留其原始大小写。
    """
    if sheet_name is None:
        return ""
    return str(sheet_name).strip().strip("'")


def normalize_cell_addr(addr: Any) -> str:
    """调用底层的地址标准化工具（通常处理 '$' 符号去除及 A1 样式转换）。"""
    return cell_utils.normalize_cell_address(addr)


def normalize_cell_key(cell_key: str) -> str:
    """
    规范化全局键值格式。最终形态约定为: 'SheetName!A1'。
    核心防呆：必须保留后端特有的后缀片段（例如时间序列输入的 '_1' 和 '_MakeInput'），
    防止被底层的 normalize_cell_addr 函数误伤截断。
    """
    if not isinstance(cell_key, str) or not cell_key:
        raise BridgeError("单元格键值（cell_key）必须是非空字符串，例如 'Sheet1!A1'。")
    key = cell_key.replace("$", "")
    
    if "!" not in key:
        # 仅有单元格地址时，保护后缀不被破坏
        parts = key.split("_", 1)
        base_addr = normalize_cell_addr(parts[0])
        return f"{base_addr}_{parts[1]}" if len(parts) > 1 else base_addr

    sheet, addr = key.split("!", 1)
    sheet = normalize_sheet_name(sheet)
    
    # 同样在存在工作表名的情况下，保护地址后缀
    parts = addr.split("_", 1)
    base_addr = normalize_cell_addr(parts[0])
    addr = f"{base_addr}_{parts[1]}" if len(parts) > 1 else base_addr
    
    return f"{sheet}!{addr}"


def range_to_cell_key(cell_range: Any) -> str:
    """
    将底层的 COM Range 对象 (或 PyXLL 的 XLCell) 转换为标准化的单元格字符串键值。
    用于在读取 Excel 交互时生成全局唯一的索引键。
    """
    try:
        sheet_name = ""
        try:
            # 尝试通过 COM 属性抓取工作表名称
            if hasattr(cell_range, "Worksheet") and hasattr(cell_range.Worksheet, "Name"):
                sheet_name = str(cell_range.Worksheet.Name)
        except Exception:
            sheet_name = ""

        addr = normalize_cell_addr(cell_range)
        if sheet_name:
            return f"{normalize_sheet_name(sheet_name)}!{addr}"
        # 降级处理：如果因为权限或对象类型拿不到工作表名，就只返回地址
        return addr
    except Exception as e:
        raise BridgeError(f"将 COM 范围对象（range）转换为单元格键值失败: {e}") from e


# =======================================================
# Part 4. 公式解析与构建协议 (Formula Parsing / Building)
# =======================================================

# -------------------------------------------------------
# Part 4.1: 输入分布函数解析与构建
# -------------------------------------------------------
def parse_first_distribution_in_formula(formula: str) -> Optional[Dict[str, Any]]:
    """
    在传入的 Excel 公式字符串中，查找并解析第一个识别到的 Drisk 分布函数。
    
    返回结构化字典，包含：
    - func_name: 函数名
    - dist_params: 基础分布参数
    - markers: 属性标记（如截断、平移）
    - attr_calls: 原始的属性函数调用文本列表
    - full_match: 完整匹配的文本
    - args_list: 参数列表
    """
    if not isinstance(formula, str) or not formula.startswith("="):
        return None

    # 调用底层正则提取公式内所有分布
    dist_funcs = formula_parser.extract_all_distribution_functions(formula)
    if not dist_funcs:
        return None

    first = dist_funcs[0]
    full_match = first.get("full_match", "") or ""
    full_match = full_match.lstrip("@")  # 清除 Office 365 可能带来的隐式交叉引用符
    if not full_match:
        return None

    # 使用后端解析助手解析独立的函数调用
    func_name, args_list = formula_parser.parse_complete_formula("=" + full_match)
    if not func_name:
        return None

    # 分离基础参数与附加的属性标记
    dist_params, markers = formula_parser.extract_dist_params_and_markers(args_list)
    # 提取并保留原始的属性函数调用片段（用于 UI 表单反向回填）
    attr_calls = [a.strip() for a in args_list if isinstance(a, str) and a.strip().startswith("Drisk")]

    return {
        "func_name": func_name,
        "dist_params": dist_params,
        "markers": markers,
        "attr_calls": attr_calls,
        "full_match": full_match,
        "args_list": args_list,
    }


def marker_dict_to_attr_calls(markers: Dict[str, Any]) -> List[str]:
    """
    [序列化辅助] 将后端的标记字典 (Marker Dict) 转化为对应的 Excel 属性函数文本。
    该方法为 UI 构建最终公式提供支持，仅转换后端 formula_parser 能够识别的标记类型。
    """
    if not markers:
        return []

    calls: List[str] = []

    # 优先处理身份/元数据类属性，保证公式可读性时这类标识靠前
    if "name" in markers and markers["name"] not in (None, ""):
        calls.append(f'DriskName("{_escape_excel_string(markers["name"])}")')
    if "units" in markers and markers["units"] not in (None, ""):
        calls.append(f'DriskUnits("{_escape_excel_string(markers["units"])}")')
    if "category" in markers and markers["category"] not in (None, ""):
        calls.append(f'DriskCategory("{_escape_excel_string(markers["category"])}")')

    if "is_date" in markers:
        calls.append(f'DriskIsDate({_excel_bool(markers["is_date"])})')
    if "is_discrete" in markers:
        calls.append(f'DriskIsDiscrete({_excel_bool(markers["is_discrete"])})')

    # 采样变换标记（平移）
    if "shift" in markers:
        calls.append(f"DriskShift({_fmt_num(markers['shift'])})")

    # 处理各种截断变体 (单尾/双尾，数值/百分位)
    for k, fn in (
        ("truncate", "DriskTruncate"),
        ("truncate2", "DriskTruncate2"),
        ("truncatep", "DriskTruncateP"),
        ("truncatep2", "DriskTruncateP2"),
    ):
        if k in markers and markers[k]:
            lower_upper = str(markers[k])
            parts = [p.strip() for p in lower_upper.split(",") if p.strip()]
            if len(parts) == 1:
                calls.append(f"{fn}({_fmt_num(parts[0])})")
            elif len(parts) >= 2:
                calls.append(f"{fn}({_fmt_num(parts[0])},{_fmt_num(parts[1])})")

    # 处理随机数种子标记
    if "seed" in markers and markers["seed"]:
        seed_str = str(markers["seed"])
        parts = [p.strip() for p in seed_str.split(",") if p.strip()]
        if len(parts) >= 2:
            calls.append(f"DriskSeed({_fmt_int(parts[0])},{_fmt_int(parts[1])})")
        elif len(parts) == 1:
            calls.append(f"DriskSeed({_fmt_int(parts[0])})")

    # 静态值设定标记
    if "static" in markers:
        calls.append(f"DriskStatic({_fmt_num(markers['static'])})")

    if "loc" in markers and bool(markers["loc"]):
        calls.append("DriskLoc()")

    return calls


def build_distribution_formula(
    func_name: str,
    params: Union[Sequence[Any], Dict[str, Any]],
    markers: Optional[Dict[str, Any]] = None,
    extra_attr_calls: Optional[Sequence[str]] = None,
) -> str:
    """
    [UI 反写接口] 构建一个可供第一组 (后端) 解析执行的 Excel 公式字符串。

    参数:
        func_name: 分布函数名称，例如 "DriskNormal"
        params: 参数值序列，或者以 param_names 为键的字典对象
        markers: 属性标记字典 (shift/truncate/name/units/seed/...)
        extra_attr_calls: 额外的属性调用文本列表 (如 ['DriskName("X")'])

    返回:
        以 "=" 开头的合法 Excel 公式字符串。
    """
    if not func_name:
        raise BridgeError("必须提供 func_name，例如 'DriskNormal'。")

    func_name = func_name.lstrip("@")

    spec = get_distribution_spec_by_func_name(func_name)
    if spec is None:
        raise BridgeError(
            f"不支持的分布函数: {func_name}。UI 只能使用后端支持的那 9 种基础分布。"
        )

    # 依据元数据的参数定义顺序，将字典平铺为列表，保障位置参数的准确性
    if isinstance(params, dict):
        ordered = []
        for p in spec.param_names:
            if p not in params:
                raise BridgeError(f"构建公式失败: {spec.func_name} 缺少必要的参数 '{p}'。")
            ordered.append(params[p])
        param_list = ordered
    else:
        param_list = list(params)

    # 校验参数数量底线
    if len(param_list) < len(spec.param_names):
        raise BridgeError(f"{spec.func_name} 需要至少 {len(spec.param_names)} 个参数，当前列表: {spec.param_names}")

    args: List[str] = [_fmt_num(v) for v in param_list[: len(spec.param_names)]]

    # 追加序列化后的属性标记
    if markers:
        args.extend(marker_dict_to_attr_calls(markers))

    # 追加用户手工补充的额外属性声明
    if extra_attr_calls:
        for c in extra_attr_calls:
            c = (c or "").strip()
            if c:
                args.append(c)

    arg_text = ",".join(args)
    return f"={spec.func_name}({arg_text})"


# -------------------------------------------------------
# Part 4.2: DriskOutput 与属性调用解析
# -------------------------------------------------------
def _strip_formula_token(token: Any) -> str:
    """
    清洗公式 Token：将带有包裹引号的公式参数字符串规范化为纯文本，
    并处理内部双写引号（Excel 转义风格）的还原。
    """
    text = str(token).strip()
    if len(text) >= 2 and text[0] == text[-1] and text[0] in ("'", '"'):
        quote = text[0]
        inner = text[1:-1]
        doubled = quote * 2
        return inner.replace(doubled, quote)
    return text


def _to_int_or_default(value: Any, default: int = 1) -> int:
    """安全整形转换器，失败时优雅降级为默认值。"""
    try:
        return int(float(str(value).strip()))
    except Exception:
        return default


def _parse_excel_bool(value: Any, default: bool = False) -> bool:
    """解析 Excel 风格的布尔值字面量（True/False/1/0）。"""
    text = str(value).strip().lower()
    if text in {"true", "1"}:
        return True
    if text in {"false", "0"}:
        return False
    return default


def _extract_function_args_text(formula_body: str, func_name: str) -> Optional[str]:
    """
    高级语法解析器：通过字符遍历与深度计数，尊重嵌套括号原则提取参数。
    主要为了解决像 DriskOutput(..., DriskUnits("年")) 这种嵌套调用被简单正则错误截断的问题。
    """
    pattern = re.compile(rf'@?\s*{re.escape(func_name)}\s*\(', re.IGNORECASE)
    match = pattern.search(formula_body)
    if not match:
        return None

    start = match.end()
    i = start
    depth = 1
    in_quotes = False
    quote_char = ""

    while i < len(formula_body):
        ch = formula_body[i]

        if in_quotes:
            if ch == quote_char:
                # 兼容 Excel 内部对引号的双写转义 ("" 或 '')
                if i + 1 < len(formula_body) and formula_body[i + 1] == quote_char:
                    i += 2
                    continue
                in_quotes = False
        else:
            if ch in ("'", '"'):
                in_quotes = True
                quote_char = ch
            elif ch == "(":
                depth += 1
            elif ch == ")":
                depth -= 1
                if depth == 0:
                    return formula_body[start:i]
        i += 1

    return None


def _parse_output_attr_call(arg_text: str) -> Optional[Tuple[str, Any]]:
    """解析包裹在 DriskOutput(...) 内的单体 Drisk 属性函数调用（如 DriskName）。"""
    match = re.match(r'^\s*@?\s*([A-Za-z_][A-Za-z0-9_]*)\s*\((.*)\)\s*$', str(arg_text), re.IGNORECASE | re.DOTALL)
    if not match:
        return None

    func_name = match.group(1).lower()
    inner_text = match.group(2).strip()
    args_list = formula_parser.parse_args_with_nested_functions(inner_text)

    # 路由不同属性的解析逻辑
    if func_name == "driskunits":
        return "units", _strip_formula_token(args_list[0]) if args_list else ""
    if func_name == "driskname":
        return "name", _strip_formula_token(args_list[0]) if args_list else ""
    if func_name == "driskcategory":
        return "category", _strip_formula_token(args_list[0]) if args_list else ""
    if func_name == "driskisdate":
        return "is_date", _parse_excel_bool(args_list[0] if args_list else "TRUE", default=True)
    if func_name == "driskisdiscrete":
        return "is_discrete", _parse_excel_bool(args_list[0] if args_list else "TRUE", default=True)
    if func_name == "driskconvergence":
        if not args_list:
            return "convergence", ""
        joined = ",".join(_strip_formula_token(x) for x in args_list if str(x).strip() != "")
        return "convergence", joined
    return None


def _looks_like_drisk_call(arg_text: str) -> bool:
    """粗略检测一段文本是否形似 Drisk 属性调用，用于前置拦截。"""
    return re.match(r'^\s*@?\s*Drisk[A-Za-z0-9_]*\s*\(.*\)\s*$', str(arg_text), re.IGNORECASE | re.DOTALL) is not None


def parse_output_info_compatible(formula: str) -> Dict[str, Any]:
    """
    带有向下兼容能力的 DriskOutput 元数据解析器。
    旧版本用户习惯使用基于位置的参数（如 =DriskOutput("营收", "利润组", 1)），
    新架构支持现代的属性调用参数（如 =DriskOutput(DriskName("营收"))）。
    该方法将两者平滑混合为同一字典接口。
    """
    base_info = formula_parser.extract_output_info(formula) or {}
    if not isinstance(formula, str) or not formula.startswith("="):
        return base_info

    args_text = _extract_function_args_text(formula[1:], "DriskOutput")
    if args_text is None:
        return base_info

    args_list = formula_parser.parse_args_with_nested_functions(args_text)
    if not args_list:
        if "position" not in base_info:
            base_info["position"] = 1
        return base_info

    plain_args: List[str] = []
    attr_values: Dict[str, Any] = {}
    name_from_attr = False

    for arg in args_list:
        attr_parsed = _parse_output_attr_call(arg)
        if attr_parsed is None:
            if _looks_like_drisk_call(arg):
                # 如果形似内部属性调用但不被支持，主动忽略，防止污染常规位置参数。
                continue
            plain_args.append(str(arg).strip())
            continue
        attr_key, attr_val = attr_parsed
        attr_values[attr_key] = attr_val
        if attr_key == "name" and str(attr_val).strip():
            name_from_attr = True

    # 兼容阶段：对未能解析为属性调用的参数，视作传统位置参数分配含义
    if plain_args:
        first = _strip_formula_token(plain_args[0])
        if first != "":
            base_info["name"] = first
    if len(plain_args) >= 2:
        second = _strip_formula_token(plain_args[1])
        if second != "":
            base_info["category"] = second
    if len(plain_args) >= 3:
        base_info["position"] = _to_int_or_default(plain_args[2], default=base_info.get("position", 1))

    # 覆盖阶段：显式的属性调用具有更高优先级，覆盖位置参数解析出的结果
    for k, v in attr_values.items():
        if isinstance(v, str):
            if v.strip() == "" and k in {"name", "category", "units", "convergence"}:
                continue
            base_info[k] = v.strip()
        else:
            base_info[k] = v

    base_info["position"] = _to_int_or_default(base_info.get("position", 1), default=1)
    base_info["_name_from_attr"] = bool(name_from_attr)
    return base_info


# =======================================================
# Part 5. 模拟执行与引擎调度 (Simulation Execution)
# =======================================================
# 全局扫描缓存，用于降低频繁调用 COM 的性能损耗
_SCAN_CACHE: Dict[str, Tuple[Dict, Dict, Dict, Dict, List, Dict]] = {}
# 标记迭代引擎兼容性补丁是否已挂载
_ITER_REF_COMPAT_PATCHED = False

# -------------------------------------------------------
# Part 5.1: 单元格信息扫描与环境补齐
# -------------------------------------------------------
def enrich_output_cells_metadata(app: Any, output_cells: Dict[str, Dict]) -> Dict[str, Dict]:
    """
    补全输出项元数据：
    通过 COM 访问当前工作表单元格，利用实际的公式文本充实底层扫描出来的基础结构。
    该方法刻意留在桥接适配层，以避免修改底层脆弱受保护的 formula_parser 解析模块。
    """
    if not isinstance(output_cells, dict) or not output_cells:
        return output_cells

    default_sheet = _get_active_sheet_name(app)
    merged_cells: Dict[str, Dict] = {}

    for cell_key, raw_info in output_cells.items():
        info = dict(raw_info) if isinstance(raw_info, dict) else {}
        try:
            sheet_name, addr_only = _split_sheet_key(cell_key, default_sheet=default_sheet)
            sheet_obj = app.Worksheets(sheet_name) if sheet_name else app.ActiveSheet
            formula = str(sheet_obj.Range(addr_only).Formula)
            parsed = parse_output_info_compatible(formula)
            if isinstance(parsed, dict) and parsed:
                info.update(parsed)
        except Exception:
            # 捕获异常：如果目标工作表被删除或公式无法读取，容错并保留原有的扫描信息
            pass

        if "position" not in info:
            info["position"] = 1
        merged_cells[cell_key] = info

    return merged_cells


def _has_output_attribute_calls(app: Any, output_cells: Dict[str, Dict]) -> bool:
    """
    检测输出公式群中是否含有新型的 Drisk* 属性调用。
    在双引擎调度策略下，如果检测到复杂的属性调用，系统可倾向于强制降级至更安全的迭代引擎，
    防止快速索引引擎破坏或遗漏元数据关联。
    """
    if not isinstance(output_cells, dict) or not output_cells:
        return False

    attr_tokens = (
        "DRISKUNITS(",
        "DRISKNAME(",
        "DRISKCATEGORY(",
        "DRISKISDATE(",
        "DRISKISDISCRETE(",
        "DRISKCONVERGENCE(",
    )
    default_sheet = _get_active_sheet_name(app)

    for cell_key in output_cells.keys():
        try:
            sheet_name, addr_only = _split_sheet_key(cell_key, default_sheet=default_sheet)
            sheet_obj = app.Worksheets(sheet_name) if sheet_name else app.ActiveSheet
            formula_upper = str(sheet_obj.Range(addr_only).Formula).upper()
            if "DRISKOUTPUT" in formula_upper and any(token in formula_upper for token in attr_tokens):
                return True
        except Exception:
            continue
    return False


def _safe_extract_cell_refs_fallback(args_text: Any) -> List[str]:
    """
    容灾备用的单元格引用提取器。
    仅在底层的受保护解析器（formula_parser）由于边缘 Case 抛出未绑定错误时作为降级手段调用。
    """
    text = str(args_text or "").strip()
    if not text:
        return []

    refs = formula_parser.parse_formula_references_tornado(f"={text}")
    if refs:
        return refs

    # 如果 parse_formula_references_tornado 也彻底瘫痪，最后通过正则表达式硬抓取引用。
    pattern = r'(?:[A-Za-z_][A-Za-z0-9_\.]*!)?\$?[A-Za-z]{1,3}\$?\d{1,7}(?::\$?[A-Za-z]{1,3}\$?\d{1,7})?'
    out: List[str] = []
    for ref in re.findall(pattern, text, re.IGNORECASE):
        clean_ref = str(ref).replace("$", "")
        if ":" in clean_ref:
            clean_ref = clean_ref.split(":", 1)[0]
        out.append(clean_ref.upper())
    return list(dict.fromkeys(out))


def _merge_output_attr_info(existing: Dict[str, Any], parsed: Dict[str, Any]) -> Dict[str, Any]:
    """
    元数据合并器：将新解析的输出属性合并进现有字典，核心处理了 `name` 属性来源的优先级问题。
    使用显式 `DriskName` 设置的名称决不能被隐式推导的名称覆盖。
    """
    merged = dict(existing or {})
    parsed_info = dict(parsed or {})
    if not parsed_info:
        return merged

    parsed_name = str(parsed_info.get("name", "") or "").strip()
    parsed_name_from_attr = bool(parsed_info.get("_name_from_attr", False))
    existing_name = str(merged.get("name", "") or "").strip()
    existing_name_from_attr = bool(merged.get("_name_from_attr", False))

    if parsed_name:
        if parsed_name_from_attr:
            # 显式的 DriskName(...) 拥有最高优先权，无条件覆盖
            merged["name"] = parsed_name
            merged["_name_from_attr"] = True
        elif not existing_name:
            # 如果是常规位置推断出来的名字，仅当原本没有名字时才应用
            merged["name"] = parsed_name
    elif "_name_from_attr" not in merged:
        merged["_name_from_attr"] = existing_name_from_attr

    for k, v in parsed_info.items():
        if k in {"name", "_name_from_attr"}:
            continue
        merged[k] = v

    if "_name_from_attr" not in merged:
        merged["_name_from_attr"] = False
    return merged


def ensure_iterative_ref_compat_patch() -> None:
    """
    [运行时热修复] 为迭代引擎的引用提取逻辑挂载补丁。
    历史包袱：旧版解析器在遇到极简引用（例如直接写 'A1' 作为参数）时，可能会由于内部逻辑分支未覆盖
    抛出 `UnboundLocalError('ref')`。此补丁拦截该错误并路由到安全的 fallback 解析器。
    """
    global _ITER_REF_COMPAT_PATCHED
    if _ITER_REF_COMPAT_PATCHED:
        return

    extractor = getattr(simulation_engine, "extract_cell_references_from_args", None)
    if not callable(extractor):
        return
    if getattr(extractor, "_drisk_ref_compat", False):
        _ITER_REF_COMPAT_PATCHED = True
        return

    def _compat_extract_cell_references_from_args(args_text: Any):
        try:
            return extractor(args_text)
        except UnboundLocalError as exc:
            msg = str(exc)
            if "local variable 'ref'" not in msg and 'local variable "ref"' not in msg:
                raise
            return _safe_extract_cell_refs_fallback(args_text)

    _compat_extract_cell_references_from_args._drisk_ref_compat = True  # type: ignore[attr-defined]
    simulation_engine.extract_cell_references_from_args = _compat_extract_cell_references_from_args
    formula_parser.extract_cell_references_from_args = _compat_extract_cell_references_from_args
    _ITER_REF_COMPAT_PATCHED = True


def enrich_sim_output_attributes_from_formulas(sim: Any, app: Any = None) -> None:
    """
    模拟结果展示层的元数据保底方案。
    部分引擎逻辑在执行完毕后可能会丢失输出项名称属性，本方法通过二次读取 Excel 公式
    重新将 DriskName 等属性挂载回模拟结果实例中，保障图表渲染层的名称展示正确无误。
    """
    if sim is None:
        return
    output_cache = getattr(sim, "output_cache", {}) or {}
    if not output_cache:
        return

    if app is None:
        try:
            app = get_excel_app()
        except Exception:
            return

    default_sheet = _get_active_sheet_name(app)
    attrs_map = getattr(sim, "output_attributes", None)
    if not isinstance(attrs_map, dict):
        return

    for cache_key in output_cache.keys():
        existing = attrs_map.get(cache_key, {}) or {}
        # 优化策略：若已经成功获取了显式声明的 DriskName，则直接跳过昂贵的 COM 抓取操作
        if str(existing.get("name", "")).strip() and bool(existing.get("_name_from_attr", False)):
            continue
        try:
            sheet_name, addr_only = _split_sheet_key(cache_key, default_sheet=default_sheet)
            sheet_obj = app.Worksheets(sheet_name) if sheet_name else app.ActiveSheet
            formula = str(sheet_obj.Range(addr_only).Formula)
            parsed = parse_output_info_compatible(formula)
            if parsed:
                attrs_map[cache_key] = _merge_output_attr_info(existing, parsed)
        except Exception:
            continue


def enrich_sim_output_attribute_for_key(sim: Any, cell_key: str, app: Any = None) -> None:
    """对单一输出键（支持松散的键名匹配）执行精确的元数据补全。避免全量扫描浪费性能。"""
    if sim is None or not isinstance(cell_key, str):
        return
    output_cache = getattr(sim, "output_cache", {}) or {}
    if not output_cache:
        return

    normalized = cell_key.replace("$", "").upper()
    candidates: List[str] = []
    
    # 模糊匹配寻找候选键
    for cache_key in output_cache.keys():
        cache_norm = str(cache_key).replace("$", "").upper()
        if cache_norm == normalized:
            candidates.append(cache_key)
            continue
        if "!" in normalized and cache_norm.endswith("!" + normalized.split("!", 1)[1]):
            candidates.append(cache_key)
            continue
        if "!" not in normalized and cache_norm.endswith("!" + normalized):
            candidates.append(cache_key)

    if not candidates:
        candidates = [cell_key]

    if app is None:
        try:
            app = get_excel_app()
        except Exception:
            return

    default_sheet = _get_active_sheet_name(app)
    attrs_map = getattr(sim, "output_attributes", None)
    if not isinstance(attrs_map, dict):
        return

    for cache_key in candidates:
        existing = attrs_map.get(cache_key, {}) or {}
        if str(existing.get("name", "")).strip() and bool(existing.get("_name_from_attr", False)):
            continue
        try:
            sheet_name, addr_only = _split_sheet_key(cache_key, default_sheet=default_sheet)
            sheet_obj = app.Worksheets(sheet_name) if sheet_name else app.ActiveSheet
            formula = str(sheet_obj.Range(addr_only).Formula)
            parsed = parse_output_info_compatible(formula)
            if parsed:
                attrs_map[cache_key] = _merge_output_attr_info(existing, parsed)
        except Exception:
            continue


def get_excel_app():
    """桥接层统一获取 Excel COM 对象的入口，内部依托于 com_fixer 保证线程/环境安全性。"""
    try:
        import com_fixer
        return com_fixer._safe_excel_app()
    except Exception as e:
        raise BridgeError(f"无法通过 com_fixer 获取 Excel 应用程序实例: {e}") from e


def scan_workbook(app: Any, force: bool = False) -> Tuple[Dict, Dict, Dict, Dict, List, Dict]:
    """
    工作簿元数据全量扫描器。
    搜集所有参与模拟的分布输入、数据表输入以及需要监控的输出单元格。
    支持对同名工作簿的扫描结果进行内存级缓存 (`_SCAN_CACHE`) 以提升 UI 响应效率。
    """
    wb_name = _get_workbook_name(app)
    if not force and wb_name and wb_name in _SCAN_CACHE:
        return _SCAN_CACHE[wb_name]

    scan = dependency_tracker.find_all_simulation_cells_in_workbook(app)
    if isinstance(scan, tuple) and len(scan) >= 6:
        distribution_cells, simtable_cells, makeinput_cells, output_cells, all_input_keys, all_related_cells = scan
        # 拦截并在桥接层补全输出元数据
        output_cells = enrich_output_cells_metadata(app, output_cells)
        scan = (
            distribution_cells,
            simtable_cells,
            makeinput_cells,
            output_cells,
            all_input_keys,
            all_related_cells,
        )
    if wb_name:
        _SCAN_CACHE[wb_name] = scan
    return scan


# -------------------------------------------------------
# Part 5.2: 引擎运行与结果封装
# -------------------------------------------------------
def run_simulation_and_cache(
    app: Any,
    n_iterations: int,
    sampling_method: str = "MC",
    *,
    scenario_count: int = 1,
    scenario_index: int = 0,
    progress_callback: Optional[Callable[[int], None]] = None,
    cancel_event: Any = None,
    force_rescan: bool = False,
    precompute_statistics: bool = True,
) -> int:
    """
    单情景模拟调度核心：
    调用底层迭代引擎 (iterative_simulation_workbook) 执行一次完整的模拟运算，
    并将生成的巨型矩阵和分布状态打包为 SimulationResult 对象推入全局管理器缓存中。
    
    返回创建成功的 sim_id。
    """
    if n_iterations is None or int(n_iterations) <= 0:
        raise BridgeError("迭代次数必须为正整数。")
    ensure_iterative_ref_compat_patch()

    distribution_cells, simtable_cells, makeinput_cells, output_cells, all_input_keys, all_related_cells = scan_workbook(
        app, force=force_rescan
    )

    # 进入非静态模式，允许引擎产生随机扰动
    attribute_functions.set_static_mode(False)
    try:
        input_results, output_results, output_info = simulation_engine.iterative_simulation_workbook(
            app,
            distribution_cells,
            simtable_cells,
            makeinput_cells,
            output_cells,
            int(n_iterations),
            sampling_method,
            int(scenario_count),
            int(scenario_index),
            progress_callback,
            cancel_event,
        )
    finally:
        # 运算结束立即恢复静态挂机模式，保障用户日常 Excel 交互不受干扰
        attribute_functions.set_static_mode(True)

    if input_results is None or output_results is None:
        raise BridgeError("模拟失败或已被用户主动取消；没有产生有效的数据结果。")

    # 持久化模拟上下文和结果集
    sim_id = simulation_manager.create_simulation(int(n_iterations), sampling_method)
    sim = simulation_manager.get_simulation(sim_id)
    if sim is None:
        raise BridgeError("在 simulation_manager 缓存中心创建模拟对象失败。")

    sim.set_scenario_info(int(scenario_count), int(scenario_index))
    sim.distribution_cells = distribution_cells
    sim.simtable_cells = simtable_cells
    sim.makeinput_cells = makeinput_cells
    sim.output_cells = output_cells
    sim.all_input_keys = all_input_keys
    sim.all_related_cells = all_related_cells

    wb_name = _get_workbook_name(app)
    if wb_name:
        sim.workbook_name = wb_name

    # 封装备选输入特征数据
    for input_key, data in (input_results or {}).items():
        if not isinstance(data, np.ndarray) or data.size == 0:
            continue
        sheet_name, key_only = _split_sheet_key(input_key, default_sheet=_get_active_sheet_name(app))
        sim.add_input_result(key_only, data, sheet_name, attributes=None)

    # 封装核心监控输出数据
    for out_key, data in (output_results or {}).items():
        if not isinstance(data, np.ndarray) or data.size == 0:
            continue
        sheet_name, addr_only = _split_sheet_key(out_key, default_sheet=_get_active_sheet_name(app))

        out_info = {}
        if isinstance(output_info, dict) and out_key in output_info and isinstance(output_info[out_key], dict):
            out_info.update(output_info[out_key])
        if isinstance(output_cells, dict) and out_key in output_cells and isinstance(output_cells[out_key], dict):
            out_info.update(output_cells[out_key])

        sim.add_output_result(addr_only, data, sheet_name, out_info)

    # 离线预热统计量，以便后续结果面板快速呈现
    if precompute_statistics:
        try:
            sim.compute_input_statistics()
            sim.compute_output_statistics()
        except Exception:
            pass

    return sim_id


def run_scenarios_and_cache(
    app: Any,
    n_iterations: int,
    sampling_method: str = "MC",
    *,
    scenario_count: int = 1,
    progress_callback: Optional[Callable[[int], None]] = None,
    cancel_event: Any = None,
    force_rescan: bool = False,
    precompute_statistics: bool = True,
    engine_mode: str = "auto",
) -> List[int]:
    """
    多情景全局执行入口。
    核心特征：实现了双引擎智能回退（Fast Index Engine 与 Robust Iterative Engine）。
    如果在极速模式（Index）下因为复杂的 Excel 引用拓扑崩溃，能够静默地无缝切换至健壮的迭代模式。
    """
    ensure_iterative_ref_compat_patch()

    # 必须前置扫描，辅助智能调度器评估模型规模
    cells_data = scan_workbook(app, force=force_rescan)
    distribution_cells, simtable_cells, makeinput_cells, output_cells, _, _ = cells_data

    # 定义极速引擎的安全阈值边界：迭代次数过高或模型拓扑过大时禁用极速模式
    is_safe = True
    if n_iterations > 100000 or (len(distribution_cells) + len(makeinput_cells)) > 100:
        is_safe = False

    # 优先尝试 Fast Index 极速引擎
    if engine_mode == "fast" or (engine_mode == "auto" and is_safe):
        try:
            import index_functions
            from simulation_manager import get_all_simulations

            old_sims = set(get_all_simulations().keys())

            if engine_mode == "auto":
                # 自动试错期间：动态挂载钩子拦截并丢弃底层警告弹窗，实现对用户的静默回退
                original_alert = index_functions.xlcAlert
                alert_msgs = []
                index_functions.xlcAlert = lambda msg: alert_msgs.append(msg)

                try:
                    index_functions.run_index_simulation(n_iterations, scenario_count)
                finally:
                    # 无论成功失败，恢复原生弹窗系统
                    index_functions.xlcAlert = original_alert

                # 通过 ASCII 保留关键字分析弹窗日志，绕过繁琐的多语言区域判断
                if not any(
                    any(token in str(m).lower() for token in ("fail", "error", "exception", "cancel"))
                    for m in alert_msgs
                ):
                    new_sims = set(get_all_simulations().keys()) - old_sims
                    return list(new_sims)
            else:
                # 若用户强制锁定极速模式，跳过弹窗抑制直接运行
                index_functions.run_index_simulation(n_iterations, scenario_count)
                new_sims = set(get_all_simulations().keys()) - old_sims
                return list(new_sims)
        except Exception:
            # 捕获崩溃，不向上抛出，进入下面的回退循环分支
            pass

    # ===============================
    # 灾难回退：使用健壮引擎逐个执行情景循环
    # ===============================
    sim_ids: List[int] = []
    for idx in range(int(scenario_count)):
        sim_id = run_simulation_and_cache(
            app,
            n_iterations,
            sampling_method,
            scenario_count=int(scenario_count),
            scenario_index=idx,
            progress_callback=progress_callback,
            cancel_event=cancel_event,
            force_rescan=False,  # 此处复用片头的扫描结果
            precompute_statistics=precompute_statistics,
        )
        sim_ids.append(sim_id)
    return sim_ids


# =======================================================
# Part 6. UI 数据与状态检索服务 (Data retrieval for UI)
# =======================================================

# -------------------------------------------------------
# Part 6.1: 基础数据获取 (序列、属性、统计)
# -------------------------------------------------------
def get_simulation(sim_id: int):
    """安全解析实体：根据 sim_id 抽取缓存的模拟实例。"""
    sim = simulation_manager.get_simulation(int(sim_id))
    if sim is None:
        raise BridgeError(f"未找到 ID 为 {sim_id} 的模拟记录，请先运行模拟。")
    return sim


def get_series(sim_id: int, cell_key: str, *, kind: str = "output") -> np.ndarray:
    """提取核心矩阵：从后端缓存提取渲染所需的一维或多维数据序列。"""
    sim = get_simulation(sim_id)
    key = normalize_cell_key(cell_key)

    if kind.lower() == "output":
        data = sim.get_output_data(key)
    elif kind.lower() == "input":
        # 遵守权限隔离，验证输入键是否被暴露给 UI
        if not is_input_key_exposed(sim, key):
            return np.array([])
        data = sim.get_input_data(key)
    else:
        raise BridgeError("参数 kind 必须限定为 'output' 或 'input'。")

    if data is None:
        return np.array([])
    if isinstance(data, np.ndarray):
        return data
    return np.array(data)


def get_attributes(sim_id: int, cell_key: str, *, kind: str = "output") -> Dict[str, Any]:
    """提取状态元数据：获取目标的属性字典（包含用户指定的别名、单位等）。"""
    sim = get_simulation(sim_id)
    key = normalize_cell_key(cell_key)

    if kind.lower() == "output":
        enrich_sim_output_attribute_for_key(sim, key)
        attrs = sim.get_output_attributes(key)
    elif kind.lower() == "input":
        if not is_input_key_exposed(sim, key):
            return {}
        attrs = sim.get_input_attributes(key)
    else:
        raise BridgeError("参数 kind 必须限定为 'output' 或 'input'。")

    return attrs or {}


def get_statistics(sim_id: int, cell_key: str, *, kind: str = "output") -> Dict[str, float]:
    """提取分析聚合指标：通过底层引擎获取带有权威性和边界过滤（如 NaN 剔除）的计算结果。"""
    sim = get_simulation(sim_id)
    data = get_series(sim_id, cell_key, kind=kind)
    if data.size == 0:
        return {}
    try:
        return sim._compute_statistics(data) or {}
    except Exception:
        # 在某些非常规数据结构下，强制包裹 Numpy 降级重试
        try:
            return sim._compute_statistics(np.array(data)) or {}
        except Exception:
            return {}


def list_simulations() -> List[Dict[str, Any]]:
    """为 UI 提供轻量级的模拟记录列表用于渲染状态切换下拉栏。"""
    sims = simulation_manager.get_all_simulations()
    out: List[Dict[str, Any]] = []
    for s in sims:
        out.append(
            {
                "sim_id": getattr(s, "sim_id", None),
                "n_iterations": getattr(s, "n_iterations", None),
                "sampling_method": getattr(s, "sampling_method", None),
                "timestamp": getattr(s, "timestamp", None),
                "workbook_name": getattr(s, "workbook_name", None),
                "scenario_index": getattr(s, "scenario_index", 0),
                "scenario_count": getattr(s, "scenario_count", 1),
            }
        )
    return out


def get_current_sim_id() -> Optional[int]:
    """快捷路由：获取当前处于全局上下文激活状态的模拟实例 ID。"""
    try:
        return simulation_manager.get_current_sim_id()
    except Exception:
        return None


# -------------------------------------------------------
# Part 6.2: 结果面板目标解析 (Result Dialog Resolving)
# -------------------------------------------------------
@dataclass(frozen=True)
class ResultDialogTargetMatch:
    """进入结果面板时所用的经过解析的目标匹配对象，承载展示与寻址所需的关键信息。"""
    requested_key: str
    matched_key: str
    kind: str
    label: str


def get_default_result_sim_id() -> Optional[int]:
    """
    按照内存缓存插入顺序，贪婪返回第一个可用的模拟 ID。
    重要：该逻辑映射并兼容了旧版 smart_plot_macro 的行为习惯。
    """
    sims = simulation_manager.get_all_simulations()
    if isinstance(sims, dict):
        if not sims:
            return None
        try:
            return int(next(iter(sims.keys())))
        except Exception:
            return None
    return None


def _build_result_dialog_label(sim: Any, cache_key: str, kind: str) -> str:
    """利用缓存键与附属的元数据（特别是 name 别名）合成前端图表展示的友好标签。"""
    if str(kind).lower() == "input":
        attrs = getattr(sim, "input_attributes", {}).get(cache_key, {}) or {}
    else:
        attrs = getattr(sim, "output_attributes", {}).get(cache_key, {}) or {}
    display_name = str(attrs.get("name", "") or "").strip()
    name_from_attr = bool(attrs.get("_name_from_attr", False))
    display_addr = cache_key.split("!")[-1] if "!" in cache_key else cache_key
    
    # 标签防呆：如果名字就是地址本身，无需渲染为 "A1 (A1)"
    if display_name and (name_from_attr or display_name.upper() != display_addr.upper()):
        return f"{display_addr} ({display_name})"
    return display_addr


def _normalize_result_requested_key(raw_key: str, default_sheet: str = "") -> str:
    """
    匹配清洗器：
    对用户通过 UI 请求的键值进行粗略规范化以支持缓存碰撞，但绝不更改后缀的本质语义。
    """
    if not isinstance(raw_key, str):
        return ""
    key = raw_key.strip()
    if not key:
        return ""
    key = key.replace("$", "")
    if "!" in key:
        return key
    sheet_name = normalize_sheet_name(default_sheet or "")
    return f"{sheet_name}!{key}" if sheet_name else key


def resolve_result_dialog_targets(
    sim_id: int,
    requested_keys: Sequence[str],
    *,
    default_sheet: str = "",
) -> Tuple[List[ResultDialogTargetMatch], List[str]]:
    """
    目标寻址路由系统：
    将选定的前端单元格键安全映射为缓存内部有效的目标节点。
    
    匹配优先级（保持向后兼容旧版 smart_plot_macro）：
    1. 输出节点精确匹配 (不区分大小写，忽略 '$' 绝对锁定符)
    2. 输入节点的精确匹配及前缀匹配（例如匹配带有时序后缀 '_1' 的输入变量）
    """
    sim = get_simulation(sim_id)
    output_cache = getattr(sim, "output_cache", {}) or {}
    input_cache = getattr(sim, "input_cache", {}) or {}

    matches: List[ResultDialogTargetMatch] = []
    missing: List[str] = []

    output_keys = list(output_cache.keys())
    input_keys = list(input_cache.keys())

    for raw_key in requested_keys or []:
        full_key = _normalize_result_requested_key(raw_key, default_sheet=default_sheet)
        if not full_key:
            continue

        full_key_clean = full_key.replace("$", "").upper()
        matched_output = None
        for cache_key in output_keys:
            if cache_key.replace("$", "").upper() == full_key_clean:
                matched_output = cache_key
                break

        if matched_output is not None:
            matches.append(
                ResultDialogTargetMatch(
                    requested_key=full_key,
                    matched_key=matched_output,
                    kind="output",
                    label=_build_result_dialog_label(sim, matched_output, "output"),
                )
            )
            continue

        found_input = False
        for cache_key in input_keys:
            # 输入权限验证网关
            if not is_input_key_exposed(sim, cache_key):
                continue
            cache_clean = cache_key.replace("$", "").upper()
            if cache_clean == full_key_clean or cache_clean.startswith(f"{full_key_clean}_"):
                found_input = True
                matches.append(
                    ResultDialogTargetMatch(
                        requested_key=full_key,
                        matched_key=cache_key,
                        kind="input",
                        label=_build_result_dialog_label(sim, cache_key, "input"),
                    )
                )

        if not found_input:
            missing.append(full_key)

    return matches, missing


def build_result_dialog_launch_args(
    matches: Sequence[ResultDialogTargetMatch],
    *,
    default_kind: str = "output",
) -> Tuple[List[str], Dict[str, str], str]:
    """为主结果渲染面板弹窗构建初始化参数 Payload 集合 (包含 cell_keys, labels字典, kind)"""
    valid_keys: List[str] = []
    labels: Dict[str, str] = {}
    final_kind = str(default_kind or "output")

    for match in matches or []:
        valid_keys.append(match.matched_key)
        labels[match.matched_key] = match.label
        final_kind = match.kind

    return valid_keys, labels, final_kind


def prepare_result_dialog_selection(
    sim_id: int,
    requested_keys: Sequence[str],
    *,
    default_sheet: str = "",
    default_kind: str = "output",
) -> Tuple[List[str], Dict[str, str], str, List[str]]:
    """
    高层次合成入口：由 UI 的 sim_engine 的结果展现入口直接调用的聚合辅助函数。
    封装了从扫描、解析到组装的一站式逻辑。
    返回: (合法键名列表, 键值到标签的映射字典, 数据类型标识, 丢失或非法的键名列表)
    """
    try:
        sim_obj = get_simulation(sim_id)
        enrich_sim_output_attributes_from_formulas(sim_obj)
    except Exception:
        pass

    matches, missing_keys = resolve_result_dialog_targets(
        sim_id,
        requested_keys,
        default_sheet=default_sheet,
    )
    valid_keys, labels, final_kind = build_result_dialog_launch_args(
        matches,
        default_kind=default_kind,
    )
    return valid_keys, labels, final_kind, missing_keys


# -------------------------------------------------------
# Part 6.3: UI组件键值列表与散点图数据 (Keys List Extracting)
# -------------------------------------------------------
def get_all_input_keys(sim_id: int) -> Tuple[List[str], Dict[str, str]]:
    """
    拉取指定模拟记录中暴露给前端的所有 Input 变量键集合及其名称映射。
    返回: (keys列表, {key: label}字典)
    """
    sim = simulation_manager.get_simulation(sim_id)
    if not sim:
        return [], {}
        
    keys = filter_exposed_input_keys(sim, list(sim.input_cache.keys()))
    labels = {}
    for k in keys:
        # 尝试挂载用户通过属性声明的自定义业务名称        
        attrs = sim.input_attributes.get(k, {})
        name = attrs.get("name")
        # 组装混合显示名称，例如 "Sheet1!A1 (营收增长率)" 或保留 "Sheet1!A1"
        labels[k] = f"{k} ({name})" if name else k
        
    return keys, labels


def get_all_output_keys(sim_id: int) -> Tuple[List[str], Dict[str, str]]:
    """
    拉取指定模拟记录中所有监控的目标 Output 键集合及其名称映射。
    返回: (keys列表, {key: label}字典)
    """
    sim = simulation_manager.get_simulation(sim_id)
    if not sim:
        return [], {}
    try:
        enrich_sim_output_attributes_from_formulas(sim)
    except Exception:
        pass
        
    keys = list(sim.output_cache.keys())
    labels = {}
    for k in keys:
        attrs = sim.output_attributes.get(k, {})
        name = str(attrs.get("name", "") or "").strip()
        name_from_attr = bool(attrs.get("_name_from_attr", False))
        addr = k.split("!")[-1] if "!" in k else k
        if name and (name_from_attr or name.upper() != addr.upper()):
            labels[k] = f"{k} ({name})"
        else:
            labels[k] = k
        
    return keys, labels


# =======================================================
# Part 7. 理论分布工厂 (Theory Distribution Factory)
# 核心架构解耦区：彻底消除早期架构中的“前端脑裂”现象。
# 该部分接管了底层具备强大截断/平移等数学计算能力的实体 Distribution 对象，
# 并封装后直接暴露给前端（专供 theory_adapter 模块等调用进行 PDF/CDF 理论值测算）。
# =======================================================
def _parse_ui_to_float_list(val: Any) -> List[float]:
    """将来自 UI 层繁杂模糊的参数负载转化为清晰统一的纯浮点数列表。"""
    if isinstance(val, (list, tuple, np.ndarray)):
        return [float(x) for x in val if x is not None]
    try:
        # 处理旧版本残留的 Excel 数组字符串风格 "{1,2,3}" -> "1,2,3"
        s = str(val).strip().strip('{}')
        if not s:
            return []
        # 同时兼容中英文逗号/分号分隔的数值流
        parts = s.replace(';', ',').split(',')
        return [float(p.strip()) for p in parts if p.strip()]
    except (ValueError, TypeError):
        return []


def _normalize_triang_params_for_bridge(params: List[Any]) -> Optional[List[Any]]:
    """
    针对三角分布 (DriskTriang) 特殊的参数校验与桥接归一化。
    容忍来自边缘侧输入的 a <= c <= b 顺序关系，但坚决拒绝错乱顺序或无意义的完全坍缩参数面。
    """
    if len(params) < 3:
        return params

    try:
        a = float(params[0])
        c = float(params[1])
        b = float(params[2])
    except (ValueError, TypeError):
        return params

    # 保持显式的业务逻辑约束：拒绝非升序的不合法参数输入。
    if a > b or c < a or c > b:
        return None

    # 若发生理论上的完全坍缩点 (a=c=b)，触发静默抛弃策略，防止底层计算除零崩溃
    if a == c == b:
        return None

    # 数值微调安全垫：底层的 TriangDistribution 模型对于边界严格要求 a < c < b；
    # 针对端点相等但未完全坍缩的退化场景，引入极小位移计算下一刻可用浮点数。
    if c == a and b > a:
        c = float(np.nextafter(a, b))
    elif c == b and b > a:
        c = float(np.nextafter(b, a))

    normalized = list(params)
    normalized[0] = a
    normalized[1] = c
    normalized[2] = b
    return normalized


def get_backend_distribution(func_name: str, params: Sequence[Any], attrs: Optional[Dict[str, Any]] = None):
    """
    中央分布实例化工厂：
    根据给定的分布名、参数矩阵与属性字典（包含平移/截断），反射构建并返回真正的底层数学引擎运算对象。
    若不支持或参数损坏，抛弃并返回 None。
    """
    normalized_func_name = (func_name or "").lstrip("@").strip()
    spec = get_distribution_spec_by_func_name(normalized_func_name)
    canonical_func_name = spec.func_name if spec else normalized_func_name
    fn = canonical_func_name.upper()
    fn_core = fn.replace("DRISK", "").strip()
    dist_class = None
    marker_dict: Dict[str, Any] = {}

    # 构建并分类标记字典，处理复杂的截断序列和平移偏移量
    if attrs and isinstance(attrs, dict):
        for k, v in attrs.items():
            key_lower = str(k).lower()
            if key_lower in {"truncate", "truncatep", "truncate2", "truncatep2"}:
                if isinstance(v, (list, tuple)) and len(v) == 2:
                    marker_dict[key_lower] = (v[0], v[1])
                else:
                    marker_dict[key_lower] = v
            elif key_lower == "shift":
                try:
                    marker_dict[key_lower] = float(v)
                except (ValueError, TypeError):
                    marker_dict[key_lower] = 0.0
            else:
                marker_dict[key_lower] = v

    try:
        # 使用精确切割了 DRISK 前缀的纯净核心名作为路由依据，防范如 'DRISKNORMAL' 子串重叠引发判定灾难
        if fn_core == "NORMAL":
            from distribution_base import NormalDistribution
            dist_class = NormalDistribution
        elif fn_core == "UNIFORM":
            from distribution_base import UniformDistribution
            dist_class = UniformDistribution
        elif fn_core == "ERF":
            from dist_erf import ErfDistribution
            dist_class = ErfDistribution
        elif fn_core == "EXTVALUE":
            from dist_extvalue import ExtvalueDistribution
            dist_class = ExtvalueDistribution
        elif fn_core == "EXTVALUEMIN":
            from dist_extvaluemin import ExtvalueMinDistribution
            dist_class = ExtvalueMinDistribution
        elif fn_core == "FATIGUELIFE":
            from dist_fatiguelife import FatigueLifeDistribution
            dist_class = FatigueLifeDistribution
        elif fn_core == "CAUCHY":
            from dist_cauchy import CauchyDistribution
            dist_class = CauchyDistribution
        elif fn_core == "DAGUM":
            from dist_dagum import DagumDistribution
            dist_class = DagumDistribution
        elif fn_core == "BURR12":
            from dist_burr12 import Burr12Distribution
            dist_class = Burr12Distribution
        elif fn_core == "DOUBLETRIANG":
            from dist_doubletriang import DoubleTriangDistribution
            dist_class = DoubleTriangDistribution
        elif fn_core == "GAMMA":
            from distribution_base import GammaDistribution
            dist_class = GammaDistribution
        elif fn_core == "ERLANG":
            from dist_erlang import ErlangDistribution
            dist_class = ErlangDistribution
        elif fn_core == "POISSON":
            from distribution_base import PoissonDistribution
            dist_class = PoissonDistribution
        elif fn_core == "BETA":
            from distribution_base import BetaDistribution
            dist_class = BetaDistribution
        elif fn_core == "BETAGENERAL":
            from dist_betageneral import BetaGeneralDistribution
            dist_class = BetaGeneralDistribution
        elif fn_core == "BETASUBJ":
            from dist_betasubj import BetaSubjDistribution
            dist_class = BetaSubjDistribution
        elif fn_core == "CHISQ":
            from distribution_base import ChiSquaredDistribution
            dist_class = ChiSquaredDistribution
        elif fn_core == "F":
            from distribution_base import FDistribution
            dist_class = FDistribution
        elif fn_core in {"T", "STUDENT"}:
            from distribution_base import TDistribution
            dist_class = TDistribution
        elif fn_core == "EXPON":
            from distribution_base import ExponentialDistribution
            dist_class = ExponentialDistribution
        elif fn_core == "INVGAUSS":
            from dist_invgauss import InvgaussDistribution
            dist_class = InvgaussDistribution
        elif fn_core == "GEOMET":
            from dist_geomet import GeometDistribution
            dist_class = GeometDistribution
        elif fn_core == "HYPERGEO":
            from dist_hypergeo import HypergeoDistribution
            dist_class = HypergeoDistribution
        elif fn_core == "INTUNIFORM":
            from dist_intuniform import IntuniformDistribution
            dist_class = IntuniformDistribution
        elif fn_core == "NEGBIN":
            from dist_negbin import NegbinDistribution
            dist_class = NegbinDistribution
        elif fn_core == "BERNOULLI":
            from dist_bernoulli import BernoulliDistribution
            dist_class = BernoulliDistribution
        elif fn_core == "TRIANG":
            from dist_triang import TriangDistribution
            dist_class = TriangDistribution
        elif fn_core == "PERT":
            from dist_pert import PertDistribution
            dist_class = PertDistribution
        elif fn_core == "BINOMIAL":
            from dist_binomial import BinomialDistribution
            dist_class = BinomialDistribution
        elif fn_core == "TRIGEN":
            from dist_trigen import TrigenDistribution
            dist_class = TrigenDistribution
        elif fn_core == "CUMUL":
            from dist_cumul import CumulDistribution
            dist_class = CumulDistribution
        elif fn_core == "DUNIFORM":
            from dist_duniform import DUniformDistribution
            dist_class = DUniformDistribution
        elif fn_core == "DISCRETE":
            from dist_discrete import DiscreteDistribution
            dist_class = DiscreteDistribution
        elif fn_core == "FRECHET":
            from dist_frechet import FrechetDistribution
            dist_class = FrechetDistribution
        elif fn_core == "HYPSECANT":
            from dist_hypsecant import HypSecantDistribution
            dist_class = HypSecantDistribution
        elif fn_core == "JOHNSONSB":
            from dist_johnsonsb import JohnsonSBDistribution
            dist_class = JohnsonSBDistribution
        elif fn_core == "JOHNSONSU":
            from dist_johnsonsu import JohnsonSUDistribution
            dist_class = JohnsonSUDistribution
        elif fn_core == "KUMARASWAMY":
            from dist_kumaraswamy import KumaraswamyDistribution
            dist_class = KumaraswamyDistribution
        elif fn_core == "LAPLACE":
            from dist_laplace import LaplaceDistribution
            dist_class = LaplaceDistribution
        elif fn_core == "LEVY":
            from dist_levy import LevyDistribution
            dist_class = LevyDistribution
        elif fn_core == "LOGISTIC":
            from dist_logistic import LogisticDistribution
            dist_class = LogisticDistribution
        elif fn_core == "LOGLOGISTIC":
            from dist_loglogistic import LoglogisticDistribution
            dist_class = LoglogisticDistribution
        elif fn_core == "LOGNORM":
            from dist_lognorm import LognormDistribution
            dist_class = LognormDistribution
        elif fn_core == "LOGNORM2":
            from dist_lognorm2 import Lognorm2Distribution
            dist_class = Lognorm2Distribution
        elif fn_core == "RAYLEIGH":
            from dist_rayleigh import RayleighDistribution
            dist_class = RayleighDistribution
        elif fn_core == "RECIPROCAL":
            from dist_reciprocal import ReciprocalDistribution
            dist_class = ReciprocalDistribution
        elif fn_core == "WEIBULL":
            from dist_weibull import WeibullDistribution
            dist_class = WeibullDistribution
        elif fn_core == "PARETO":
            from dist_pareto import ParetoDistribution
            dist_class = ParetoDistribution
        elif fn_core == "PARETO2":
            from dist_pareto2 import Pareto2Distribution
            dist_class = Pareto2Distribution
        elif fn_core == "PEARSON5":
            from dist_pearson5 import Pearson5Distribution
            dist_class = Pearson5Distribution
        elif fn_core == "PEARSON6":
            from dist_pearson6 import Pearson6Distribution
            dist_class = Pearson6Distribution
        elif fn_core == "GENERAL":
            from dist_general import GeneralDistribution
            dist_class = GeneralDistribution
        elif fn_core in {"HISTOGRM", "HISTGRAM"}:
            from dist_histogrm import HistogrmDistribution
            dist_class = HistogrmDistribution
        # 对无法精确切割匹配的未知遗留旧版别名，执行极宽容的回退模糊识别网
        elif "NORMAL" in fn:
            from distribution_base import NormalDistribution
            dist_class = NormalDistribution
        elif "UNIFORM" in fn:
            from distribution_base import UniformDistribution
            dist_class = UniformDistribution
        elif "GAMMA" in fn:
            from distribution_base import GammaDistribution
            dist_class = GammaDistribution
        elif "BETA" in fn:
            from distribution_base import BetaDistribution
            dist_class = BetaDistribution
        elif "EXPON" in fn:
            from distribution_base import ExponentialDistribution
            dist_class = ExponentialDistribution
        elif "POISSON" in fn:
            from distribution_base import PoissonDistribution
            dist_class = PoissonDistribution
        elif "CHISQ" in fn:
            from distribution_base import ChiSquaredDistribution
            dist_class = ChiSquaredDistribution
        elif fn_core == "F":
            from distribution_base import FDistribution
            dist_class = FDistribution
        elif fn_core == "T" or "STUDENT" in fn:
            from distribution_base import TDistribution
            dist_class = TDistribution
        elif "TRIANG" in fn:
            from dist_triang import TriangDistribution
            dist_class = TriangDistribution
        elif "TRIGEN" in fn:
            from dist_trigen import TrigenDistribution
            dist_class = TrigenDistribution
        elif "DISCRETE" in fn:
            from dist_discrete import DiscreteDistribution
            dist_class = DiscreteDistribution
        elif "CUMUL" in fn:
            from dist_cumul import CumulDistribution
            dist_class = CumulDistribution
        elif "BERNOULLI" in fn:
            from dist_bernoulli import BernoulliDistribution
            dist_class = BernoulliDistribution
        elif "BINOMIAL" in fn:
            from dist_binomial import BinomialDistribution
            dist_class = BinomialDistribution
    except Exception as e:
        print(f"[Backend Bridge] 实例化分布对象时发生错误 ({func_name}): {e}")
        return None

    if not dist_class:
        print(f"[Backend Bridge] 无法识别的分布函数名: {func_name}")
        return None

    list_like_params = list(params) if isinstance(params, (list, tuple, np.ndarray)) else [params]
    cleaned_params: List[Any] = []

    # 针对底层数据表驱动型的高级分布群，放开浮点数强转限制，原封不动透传表格引用矩阵
    if fn_core in {"DISCRETE", "CUMUL", "DUNIFORM", "GENERAL", "HISTOGRM", "HISTGRAM"}:
        cleaned_params = list_like_params
        
        # 定义空负载防呆校验工具
        def _is_empty_payload(value: Any) -> bool:
            if value is None:
                return True
            if isinstance(value, str):
                return value.strip() == ""
            if isinstance(value, (list, tuple, np.ndarray)):
                return len(value) == 0
            return False

        if fn_core == "DISCRETE":
            if len(cleaned_params) < 2:
                print(f"[Backend Bridge] 离散分布参数解析失败: {params}")
                return None
            x_table = cleaned_params[0]
            p_table = cleaned_params[1]
            if _is_empty_payload(x_table) or _is_empty_payload(p_table):
                print(f"[Backend Bridge] 离散分布参数解析失败: {params}")
                return None
        elif fn_core == "CUMUL":
            # Cumul 分布的有效载荷（payload）结构要求为 [最小值, 最大值, x数据表, p数据表]。
            if len(cleaned_params) < 4:
                print(f"[Backend Bridge] Cumul 分布缺少 x/p 数据表参数: {params}")
                return None
            x_table = cleaned_params[2]
            p_table = cleaned_params[3]
            if _is_empty_payload(x_table) or _is_empty_payload(p_table):
                print(f"[Backend Bridge] Cumul 分布的 x/p 数据表负载为空: {params}")
                return None
        elif fn_core == "DUNIFORM":
            first_param = cleaned_params[0] if cleaned_params else None
            if first_param is None or (isinstance(first_param, str) and not first_param.strip()):
                print(f"[Backend Bridge] DUniform 分布参数解析失败: {params}")
                return None
        elif fn_core == "GENERAL":
            # General 需要最小值/最大值以及 x数据表 和 p数据表 的矩阵负载。
            if len(cleaned_params) < 4:
                print(f"[Backend Bridge] General 分布缺少数据表负载参数: {params}")
                return None
            x_table = cleaned_params[2]
            p_table = cleaned_params[3]
            if (x_table is None or str(x_table).strip() == "") or (p_table is None or str(p_table).strip() == ""):
                print(f"[Backend Bridge] General 分布的 x/p 数据表负载为空: {params}")
                return None
        else:
            # Histogrm 需要最小值/最大值以及一个核心的 p数据表 负载。
            if len(cleaned_params) < 3:
                print(f"[Backend Bridge] Histogrm 分布缺少 p-table 负载参数: {params}")
                return None
            p_table = cleaned_params[2]
            if p_table is None or str(p_table).strip() == "":
                print(f"[Backend Bridge] Histogrm 分布的 p-table 负载为空: {params}")
                return None
    elif fn_core == "BERNOULLI":
        # 兼容单独的参数拆包逻辑
        try:
            raw_val = params[0] if isinstance(params, (list, tuple, np.ndarray)) else params
            if isinstance(raw_val, (list, tuple, np.ndarray)):
                raw_val = raw_val[0]
            cleaned_params = [float(raw_val)]
        except Exception:
            cleaned_params = [0.5]
    else:
        # 常规浮点数值类参数清理流程
        for p in list_like_params:
            try:
                if isinstance(p, (list, tuple, np.ndarray)) and len(p) > 0:
                    cleaned_params.append(float(p[0]))
                else:
                    cleaned_params.append(float(p))
            except (ValueError, TypeError):
                cleaned_params.append(0.0)

    # 针对三角分布执行上文声明的归一化调整
    if dist_class is not None and getattr(dist_class, "__name__", "") == "TriangDistribution":
        normalized_triang_params = _normalize_triang_params_for_bridge(cleaned_params)
        if normalized_triang_params is None:
            return None
        cleaned_params = normalized_triang_params

    # 工厂对象终态生成与兼容性装配
    try:
        try:
            return dist_class(cleaned_params, marker_dict, canonical_func_name)
        except TypeError:
            # 向下兼容历史遗留的分布对象构造器，适用于未适配完整 canonical_func_name 传参的老分布基类
            return dist_class(cleaned_params, marker_dict)
    except Exception as e:
        print(f"[Backend Bridge] 创建分布实例时出错 {func_name} 异常: {e} | 参数: {cleaned_params}")
        return None


# =======================================================
# Part 8. 底层工具与辅助函数 (Helpers)
# 此区域集中提供轻量、无状态的文本与对象解析辅助能力。
# =======================================================
def _get_workbook_name(app: Any) -> str:
    """非侵入式尝试从跨进程 COM 对象中提取活动工作簿的名字。"""
    try:
        if hasattr(app, "ActiveWorkbook") and hasattr(app.ActiveWorkbook, "Name"):
            return str(app.ActiveWorkbook.Name)
    except Exception:
        pass
    return ""


def _get_active_sheet_name(app: Any) -> str:
    """非侵入式尝试从跨进程 COM 对象中提取活动工作表的名字。容灾回退至 'Sheet1'。"""
    try:
        if hasattr(app, "ActiveSheet") and hasattr(app.ActiveSheet, "Name"):
            return str(app.ActiveSheet.Name)
    except Exception:
        pass
    return "Sheet1"


def _split_sheet_key(key: str, default_sheet: str) -> Tuple[str, str]:
    """
    通用拆分器：将全局寻址键 'Sheet!A1' 平滑拆解为元组 ('Sheet', 'A1')。
    若无显式工作表名称，自动降级沿用 default_sheet 上下文。
    该方法会忠实保留 input_key 特有的后缀信息（如 '_1'）。
    """
    if not isinstance(key, str):
        return default_sheet, str(key)

    k = key.replace("$", "")
    if "!" in k:
        sheet, rest = k.split("!", 1)
        sheet = normalize_sheet_name(sheet)
        return sheet, rest.upper()
    return normalize_sheet_name(default_sheet), k.upper()


def _excel_bool(v: Any) -> str:
    """常量布尔翻译器：将 Python 环境的布尔态翻转为符合公式写入标准的 Excel 内部布尔保留字。"""
    return "TRUE" if bool(v) else "FALSE"


def _escape_excel_string(s: Any) -> str:
    """防注入转义器：将字符串字面量转义为符合 Excel 标准的双写内部语法，以允许在公式体内呈现双引号。"""
    return str(s).replace('"', '""')


def _fmt_int(v: Any) -> str:
    """防空整数转化器：安全格式化值为整数表示法，异常一律归零兜底。"""
    try:
        return str(int(float(v)))
    except Exception:
        return "0"


def _fmt_num(v: Any) -> str:
    """
    工业级浮点数转化器。
    专门对付前端传入的 NaN 与 Inf 危险数据体，将其转录为合法但数值中性的 '0' 进行兜底；
    其余通过 '.12g' 压缩呈现最友好的紧凑浮点数字面量格式。
    """
    if isinstance(v, str):
        return v.strip()
    try:
        fv = float(v)
        if np.isnan(fv) or np.isinf(fv):
            return "0"
        return f"{fv:.12g}"
    except Exception:
        return str(v).strip()