"""模拟完成后，输入变量样本暴露策略（post-simulation input sample exposure policy）的共享辅助函数。

模块定位：
- 本模块不负责“执行模拟”；
- 也不负责“收集输入样本”本身；
- 它只负责在模拟完成后，判断某个输入变量 / 输入项是否应该在后续界面中暴露出来。

典型应用场景：
1. 模拟完成后，某些输入变量允许绘图、比较、进入高级分析；
2. 某些输入变量虽然参与了模拟，但如果未被 collect 标记，则不应向用户暴露；
3. 当 scope 设为 all / collect / none 时，这里提供统一判断口径。

模块核心语义：
- all：
    所有输入变量都可暴露。
- collect：
    仅暴露被 collect 标记的输入变量。
- none：
    所有输入变量都不暴露。

后续接手注意：
- 这里的“暴露”是 UI / 后续分析层面的准入判断，不等于是否参与底层模拟；
- 若未来 collect 语义变化，优先统一修改本模块，而不是在多个绘图模块里各自硬编码。
"""

from __future__ import annotations

import re
from typing import Any, Dict, Iterable, List, Sequence, Tuple

# 匹配输入变量 key 的尾部后缀：
# 例如：
#   A1_3
#   A1_MAKEINPUT
# 这些尾部可能是运行时附加标记，需要在做属性回查时剥离。
_INPUT_SUFFIX_PATTERN = re.compile(r"_(\d+|MAKEINPUT)$", re.IGNORECASE)


def normalize_input_collection_scope(raw_scope: Any) -> str:
    """
    归一化输入收集范围配置。

    合法值：
    - "all"
    - "collect"
    - "none"

    其余任何输入都统一回退为 "all"。

    设计意图：
    - 上层对象属性可能为空、拼写不规范、大小写不一致；
    - 这里统一做收口，保证后续逻辑只面对稳定三值。
    """
    text = str(raw_scope or "").strip().lower()
    if text in ("all", "collect", "none"):
        return text
    return "all"


def get_input_collection_scope(sim_obj: Any) -> str:
    """
    从 simulation 对象中读取“输入样本暴露范围”配置。

    兼容读取顺序：
    1. sim_obj.input_collection_scope_mode
    2. sim_obj.input_scope_mode（旧字段兼容）
    3. 若对象为空或字段缺失，则回退为 "all"

    设计意图：
    - 兼容新旧字段命名；
    - 避免不同模块直接散落字段判断逻辑。
    """
    if sim_obj is None:
        return "all"

    scope = getattr(sim_obj, "input_collection_scope_mode", None)
    if scope is None:
        scope = getattr(sim_obj, "input_scope_mode", None)

    return normalize_input_collection_scope(scope)


def strip_input_suffix(cell_key: Any) -> str:
    """
    去除输入变量 key 末尾的运行时后缀。

    示例：
    - "A1_3" -> "A1"
    - "B5_MAKEINPUT" -> "B5"
    - "C10" -> "C10"

    使用场景：
    - 某些 input_attributes 是按原始 cell key 存储的；
    - 但运行时传入的 key 可能带后缀；
    - 因此查属性前要先尝试剥离后缀。
    """
    text = str(cell_key or "")
    if not text:
        return text
    return _INPUT_SUFFIX_PATTERN.sub("", text)


def normalize_cell_key(cell_key: Any) -> str:
    """
    标准化单元格 key，用于大小写 / $ 符号兼容比较。

    处理规则：
    - 去掉 `$`
    - 去首尾空格
    - 转为大写

    示例：
    - "$a$1" -> "A1"
    - " b$2 " -> "B2"

    设计意图：
    - Excel 风格单元格引用可能有绝对引用符号；
    - 大小写也可能不一致；
    - 属性查找时需要统一口径。
    """
    return str(cell_key or "").replace("$", "").strip().upper()


def _is_collect_truthy(raw_value: Any) -> bool:
    """
    判断某个 collect 标记值是否应视为“真”。

    规则：
    - bool 类型：直接返回
    - None：False
    - 空字符串：False
    - "0" / "false" / "none" / "no" / "off"：False
    - 其他非空文本：True

    设计意图：
    - collect 字段可能来源不统一，既可能是布尔值，也可能是字符串；
    - 此函数提供统一的宽松真值判断。
    """
    if isinstance(raw_value, bool):
        return raw_value

    if raw_value is None:
        return False

    text = str(raw_value).strip().lower()
    if not text:
        return False

    return text not in {"0", "false", "none", "no", "off"}


def _resolve_attr_entry(input_attrs: Any, request_key: Any) -> Dict[str, Any]:
    """
    在 input_attributes 字典中查找某个输入 key 对应的属性条目。

    查找策略：
    1. 先按原始 request_key 直接查；
    2. 若失败，再将 request_key 归一化（去 $ / 大写）后，
       与 input_attrs 中所有 key 做逐项归一化匹配；
    3. 找到且对应值为 dict 时返回该 dict；
    4. 否则返回空 dict。

    参数：
    - input_attrs：通常来自 sim_obj.input_attributes
    - request_key：待查找的输入 key

    设计意图：
    - 兼容 key 的格式差异；
    - 防止上层每次都自己手写 key 归一化逻辑。
    """
    if not isinstance(input_attrs, dict):
        return {}

    req = str(request_key or "")

    # 优先精确匹配，效率最高
    if req in input_attrs and isinstance(input_attrs.get(req), dict):
        return input_attrs.get(req) or {}

    # 再做归一化后的宽松匹配
    req_norm = normalize_cell_key(req)
    for key, attrs in input_attrs.items():
        if normalize_cell_key(key) == req_norm and isinstance(attrs, dict):
            return attrs

    return {}


def get_input_attributes_for_key(sim_obj: Any, input_key: Any) -> Dict[str, Any]:
    """
    获取某个输入 key 对应的属性字典。

    查找顺序：
    1. 直接用 input_key 查；
    2. 若未找到，再对 input_key 去尾部后缀后查一次；
    3. 若仍未找到，返回空 dict。

    使用原因：
    - 运行时 key 可能是 "A1_3"；
    - 但属性表里存的可能只有 "A1"；
    - 因此这里封装“两段式查找”逻辑。
    """
    input_attrs = getattr(sim_obj, "input_attributes", {}) or {}

    attrs = _resolve_attr_entry(input_attrs, input_key)
    if attrs:
        return attrs

    base_key = strip_input_suffix(input_key)
    if base_key and base_key != str(input_key or ""):
        attrs = _resolve_attr_entry(input_attrs, base_key)
        if attrs:
            return attrs

    return {}


def is_input_key_collect_marked(sim_obj: Any, input_key: Any) -> bool:
    """
    判断某个输入 key 是否带有 collect 标记。

    判断过程：
    1. 先获取该 key 对应的属性 dict；
    2. 读取 attrs["collect"]；
    3. 交给 _is_collect_truthy 做统一真值判定。

    返回：
    - True：该输入被视为 collect 标记输入
    - False：未标记或属性缺失
    """
    attrs = get_input_attributes_for_key(sim_obj, input_key)
    if not attrs:
        return False
    return _is_collect_truthy(attrs.get("collect"))


def is_input_key_exposed(sim_obj: Any, input_key: Any) -> bool:
    """
    判断某个输入 key 在当前 scope 规则下是否应被暴露。

    规则：
    - scope == "all"     -> 一律暴露
    - scope == "none"    -> 一律不暴露
    - scope == "collect" -> 仅 collect 标记输入暴露

    这是本模块最核心的准入判断函数之一。
    """
    scope = get_input_collection_scope(sim_obj)

    if scope == "all":
        return True

    if scope == "none":
        return False

    return is_input_key_collect_marked(sim_obj, input_key)


def filter_exposed_input_keys(sim_obj: Any, keys: Sequence[Any]) -> List[str]:
    """
    根据当前暴露策略，从输入 key 列表中过滤出允许暴露的 key。

    返回值：
    - 仅保留 is_input_key_exposed(...) 为 True 的项；
    - 输出统一转为 str 列表。

    使用场景：
    - 某个绘图 / 对象选择面板拿到一串候选输入变量 key；
    - 需要先按 exposure policy 过滤，再展示给用户。
    """
    out: List[str] = []

    for key in keys or []:
        if is_input_key_exposed(sim_obj, key):
            out.append(str(key))

    return out


def filter_exposed_input_items(sim_obj: Any, items: Iterable[Tuple[Any, Any]]) -> List[Tuple[Any, Any]]:
    """
    根据当前暴露策略，从 (key, value) 二元组序列中过滤出允许暴露的项。

    输入示例：
    - [("A1", obj1), ("B2", obj2), ...]

    返回规则：
    - 仅保留其中 key 满足 is_input_key_exposed(...) 的条目；
    - value 不做修改，原样保留。

    使用场景：
    - 候选对象列表不只是 key，还带有配套值 / 元数据时；
    - 可直接用本函数一次性过滤，不需要先拆 key 再重组。
    """
    out: List[Tuple[Any, Any]] = []

    for key, value in items or []:
        if is_input_key_exposed(sim_obj, key):
            out.append((key, value))

    return out