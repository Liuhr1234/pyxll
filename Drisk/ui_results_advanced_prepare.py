# ui_results_advanced_prepare.py
"""
高级分析数据制备模块 (Advanced Preparation)

核心职责：
作为高级分析（如龙卷风图、情景分析等）的数据制备与清洗引擎。
由于高级多维分析（如计算斯皮尔曼相关性、多元分箱等）对数据质量要求极高，不能有任何缺失值（NaN），
且要求输入和输出的样本矩阵必须严格对齐。
因此，这个独立模块负责在真正交由前端渲染之前：
1. 从底层模型提取输入/输出缓存数据。
2. 通过 Excel COM 智能推断人类可读的变量名。
3. 进行矩阵长度对齐，并剔除无效数据行。
最终将纯净、标准的 DataFrame 打包为 AdvancedPreparedPayload 载荷交付给下游。
"""

from dataclasses import dataclass
import re
import pandas as pd

from com_fixer import _safe_excel_app
from ui_shared import resolve_visible_variable_name

# =======================================================
# 1. 异常处理与数据通信载荷结构 (Exceptions & Payloads)
# =======================================================

class AdvancedPreparationError(RuntimeError):
    """
    [异常类] 高级模式数据制备专属异常。
    提供结构化的错误代码 (code) 与错误信息，便于在 UI 主窗口中进行精准拦截与分类错误提示
    （例如：缺少输入数据、有效行数不足等）。
    """
    def __init__(self, code: str, message: str):
        super().__init__(message)
        self.code = code

@dataclass
class AdvancedPreparedPayload:
    """
    [数据类] 高级分析统一数据载荷 (Payload)。
    封装了经过提取、命名解析、清洗、对齐后的标准数据上下文。
    下游的敏感性分析(Tornado)和情景分析(Scenario)渲染器将直接消费此标准化对象。
    """
    full_df: pd.DataFrame          # 包含所有有效输入变量和唯一输出变量的纯净矩阵 DataFrame (已剔除 NaN)
    output_display_name: str       # 输出变量的最终前端展示名称
    valid_input_cols: list[str]    # 参与分析的所有有效输入变量的列名清单
    output_key: str                # 输出变量的原始底层键值 (如 Sheet1!A1)
    input_mapping: dict[str, str]  # 记录 DataFrame 列名映射回原始底层单元格键值的字典
    xl_app: object                 # Excel COM 实例句柄 (用于后续可能需要的 COM 操作)


# =======================================================
# 2. 数据提取与制备核心逻辑 (Core Preparation Logic)
# =======================================================

def resolve_output_data(sim_result, output_key: str, sim_id: int):
    """
    输出数据解析提取器：
    从当前模拟结果的输出缓存 (output_cache) 中安全提取目标变量阵列，
    并向下兼容带有模拟 ID 后缀 (例如 'Sheet1!A1_1') 的旧版键值命名规则。
    """
    output_data = sim_result.output_cache.get(output_key)
    if output_data is None:
        output_data = sim_result.output_cache.get(f"{output_key}_{sim_id}")
    return output_data


# 预编译正则：匹配缓存键名的后缀模式（提取形如 Sheet1!A1_MAKEINPUT 或 Sheet1!A1_1 的后缀）
_INPUT_CACHE_KEY_SUFFIX_PATTERN = re.compile(
    r"([A-Z]{1,7}\d+)(?:_(\d+|MAKEINPUT))?$",
    re.IGNORECASE,
)


def _split_input_cache_key(raw_key: str) -> tuple[str, str]:
    """
    输入缓存键名拆分器：
    将原始键名（如 "Sheet1!A1_MAKEINPUT"）拆分为基础地址 ("Sheet1!A1") 和特殊后缀 ("MAKEINPUT")。
    便于后续将相同基础地址的多个相关输入变量进行归类或过滤。
    """
    text = str(raw_key or "").replace("$", "").strip()
    if not text:
        return "", ""

    sheet_name = ""
    addr_text = text
    
    # 拆分工作表名与单元格地址
    if "!" in text:
        sheet_name, addr_text = text.rsplit("!", 1)
        sheet_name = str(sheet_name).strip()

    addr_text = str(addr_text).strip().upper()
    suffix = ""
    
    # 提取特殊后缀
    m = _INPUT_CACHE_KEY_SUFFIX_PATTERN.fullmatch(addr_text)
    if m:
        addr_text = str(m.group(1) or "").upper()
        suffix = str(m.group(2) or "").upper()

    base_key = f"{sheet_name}!{addr_text}" if sheet_name else addr_text
    return base_key, suffix


def build_advanced_payload(sim_result, output_key: str, output_data, *, min_rows: int) -> AdvancedPreparedPayload:
    """
    构建高级分析数据载荷的核心工厂函数。
    
    执行流：
    1. 前置校验：检查是否存在输入变量分布数据。
    2. 输出处理：提取输出变量属性并推断展示名称。
    3. 输入迭代与智能命名：遍历所有输入变量，优先读取缓存中的 Name/Category 属性；若缺失，
       则调用 Excel COM 接口从表头自动抓取并组装名称。处理名称冲突并防止重名。
    4. 数组切平与矩阵合并：将不同长度的随机数序列截断对齐到最短长度（n_out），并灌入 DataFrame。
    5. 严格数据清洗：强制进行数值型转换 (to_numeric)，并无情剔除任何含有 NaN / 错误的行。
    6. 样本量防线：校验清洗后剩下的有效数据行数是否满足该高级模式要求的最低样本阈值 (min_rows)。
    """
    
    input_dict = getattr(sim_result, "input_cache", {})
    if not input_dict:
        # 如果当前模拟没有任何输入分布产生，则高级分析无从谈起，抛出精准异常
        raise AdvancedPreparationError("missing_input", "没有可用的输入分布数据。")

    df_data = {}
    valid_input_cols = []
    input_mapping = {}
    makeinput_base_keys = set()
    non_makeinput_counts_by_base = {}
    
    # 以输出数组的长度作为基准线，所有输入数组的长度不能超过此基准
    n_out = len(output_data)

    # 尝试获取 Excel COM 句柄用于命名推断
    xl_app = None
    try:
        xl_app = _safe_excel_app()
    except Exception:
        pass

    # 提取输出变量展示名称
    out_attrs = sim_result.output_attributes.get(output_key, {}) or {}
    output_display_name = resolve_visible_variable_name(
        output_key,
        out_attrs,
        excel_app=xl_app,
        fallback_label=output_key.split("!")[-1] if "!" in output_key else output_key,
    )

    # Pre-scan cache keys once so MakeInput de-duplication can stay compatible
    # without dropping legitimate multi-input series from numpy/vectorized runs.
    for raw_in_key in input_dict.keys():
        base_key, suffix = _split_input_cache_key(raw_in_key)
        if not base_key:
            continue
        base_upper = base_key.upper()
        if suffix == "MAKEINPUT":
            makeinput_base_keys.add(base_upper)
            continue
        non_makeinput_counts_by_base[base_upper] = non_makeinput_counts_by_base.get(base_upper, 0) + 1

    # ---------------------------------------------------------
    # Step A: 遍历所有输入变量，提取数据并执行智能名称映射
    # ---------------------------------------------------------
    for in_key, in_array in input_dict.items():
        # 过滤空数组
        if len(in_array) == 0:
            continue

        base_in_key, key_suffix = _split_input_cache_key(in_key)
        if not base_in_key:
            continue

        # Keep legacy behavior only for a likely duplicate pair:
        # one plain numeric variant (e.g. A1_1) + one A1_MAKEINPUT.
        # If multiple non-MakeInput variants exist under the same base,
        # treat them as independent candidates and keep them.
        base_upper = base_in_key.upper()
        if base_upper in makeinput_base_keys and key_suffix != "MAKEINPUT":
            non_makeinput_count = non_makeinput_counts_by_base.get(base_upper, 0)
            if non_makeinput_count <= 1:
                continue

        # 提取输入变量属性
        in_attrs = (
            sim_result.input_attributes.get(in_key)
            or sim_result.input_attributes.get(base_in_key)
            or {}
        )

        # 解析人类可读的变量展示名称
        in_name = resolve_visible_variable_name(
            base_in_key,
            in_attrs,
            excel_app=xl_app,
            fallback_label=base_in_key.split("!")[-1] if "!" in base_in_key else base_in_key,
        )

        # 冲突防御机制：处理多输入变量推断出相同名字的情况，自动追加计数后缀 (_1, _2...)
        final_col_name = in_name
        counter = 1
        while final_col_name in df_data:
            final_col_name = f"{in_name}_{counter}"
            counter += 1

        # 截取对齐数据长度并存入字典
        df_data[final_col_name] = in_array[:n_out] if len(in_array) > n_out else in_array
        valid_input_cols.append(final_col_name)
        input_mapping[final_col_name] = base_in_key

    # ---------------------------------------------------------
    # Step B: 矩阵合成与严格无情清洗
    # ---------------------------------------------------------
    
    # 安全拦截：若没有提纯出任何输入数据，阻断流程
    if not valid_input_cols:
        raise AdvancedPreparationError("no_valid_inputs", "未找到可用于高级分析的有效输入列。")

    # 避免输出列名恰好与输入列名完全一致产生冲突
    if output_display_name in df_data:
        output_display_name = f"{output_display_name}_结果"
        
    # 将目标输出数据压入字典
    df_data[output_display_name] = output_data[:n_out] if len(output_data) > n_out else output_data

    # .apply(pd.to_numeric) 强制转化为数值，将文本/错误等非数字转化为 NaN
    # .dropna() 彻底剔除任何包含 NaN 的横向样本行，确保相关性计算等矩阵运算不会崩溃
    full_df = pd.DataFrame(df_data).apply(pd.to_numeric, errors="coerce").dropna()
    
    # 样本底线校验：高级分析需要基础样本支撑，否则结果没有统计学意义
    if len(full_df) < min_rows:
        raise AdvancedPreparationError(
            "insufficient_rows",
            f"清洗后的有效数据行数不足: {len(full_df)} < {min_rows}。"
        )

    # 交付最终的标准数据载荷
    return AdvancedPreparedPayload(
        full_df=full_df,
        output_display_name=output_display_name,
        valid_input_cols=valid_input_cols,
        output_key=output_key,
        input_mapping=input_mapping,
        xl_app=xl_app,
    )
