"""
运行时专用的 simulation display-name 注册表，用于绘图类 UI 场景。

模块定位：
- 这里只管理“界面展示名称”，不改变内部 simulation 的真实身份。
- 内部识别仍然统一使用 sim_id。
- 外部元数据（例如公式属性、DriskName 等）不在这里修改。

典型用途：
- 多情景绘图时，内部数据仍按 sim1 / sim2 / sim3 组织；
- 但在界面展示上，允许把 sim1 显示为“基准方案”、sim2 显示为“保守方案”等；
- 这样既不破坏底层逻辑，又能提升前端可读性。

设计特点：
1. 纯运行时：
   - 本模块不负责持久化；
   - 关闭程序后，这里的映射默认消失。
2. UI 范围内生效：
   - 仅服务于展示层，不应被当作底层主键来源。
3. 带版本号：
   - 每次重命名后 version 自增；
   - 方便外层 UI 判断“命名是否变化，需要不要刷新”。

后续接手注意：
- 若未来要支持“重命名持久化到工程 / 配置文件”，不要直接在这里硬加文件读写；
- 更合理的做法是由更高层负责持久化，再在运行时回填到本模块。
"""

from __future__ import annotations

import re
from typing import Optional

# 匹配 series key 中的 sim 标记，例如：
# "XXX (sim2)" -> 提取出 2
# 这里忽略大小写，因此 "(SIM2)" 也能识别。
_SERIES_SIM_PATTERN = re.compile(r"\(sim(\d+)\)", re.IGNORECASE)

# 运行时映射表：
# key   = 标准化后的 sim_id（int）
# value = 自定义展示名（str）
_SIM_DISPLAY_NAME_MAP: dict[int, str] = {}

# 单调递增版本号：
# 每次 set / clear 自定义显示名时都会 +1，
# 用于外部刷新判断。
_SIM_DISPLAY_NAME_VERSION = 0


def normalize_sim_id(sim_id: object) -> Optional[int]:
    """
    将外部传入的 sim_id 归一化为 int。

    返回规则：
    - 成功：返回 int
    - 失败：返回 None

    设计意图：
    - 上层可能传入 "2"、2、numpy.int、甚至其他对象；
    - 本模块内部统一只认 int，便于后续作为 dict key 使用。
    """
    try:
        return int(sim_id)
    except Exception:
        return None


def get_custom_sim_display_name(sim_id: object) -> str:
    """
    获取某个 sim_id 的“自定义展示名”。

    返回规则：
    - 若该 sim_id 存在自定义名称，则返回去首尾空格后的文本；
    - 若不存在，返回空字符串 ""。

    注意：
    - 返回空字符串并不代表 sim_id 无效，只表示“当前没有 override”。
    - 想拿到最终展示名，请调用 get_sim_display_name(...)。
    """
    sid = normalize_sim_id(sim_id)
    if sid is None:
        return ""
    return str(_SIM_DISPLAY_NAME_MAP.get(sid, "") or "").strip()


def get_sim_display_name(sim_id: object) -> str:
    """
    获取某个 sim_id 的“最终生效展示名”。

    回退逻辑：
    1. 若 sim_id 可被标准化：
       - 先查 custom display name（用户手动重命名）；
       - 若存在则返回自定义名称；
       - 再查 SimulationResult.name（由引擎/压力测试等写入）；
       - 若为非默认格式则返回；
       - 否则回退为 "simN"。
    2. 若 sim_id 无法标准化：
       - 先尝试把原值转成字符串；
       - 有文本则直接返回该文本；
       - 否则回退为 "sim"。
    """
    sid = normalize_sim_id(sim_id)
    if sid is None:
        text = str(sim_id or "").strip()
        return text if text else "sim"

    custom = get_custom_sim_display_name(sid)
    if custom:
        return custom

    # 从 SimulationResult.name 读取引擎设置的名称
    try:
        from simulation_manager import get_simulation
        sim = get_simulation(sid)
        if sim is not None:
            sim_name = str(getattr(sim, "name", "") or "").strip()
            default_prefix = f"Sim_{sid}_"
            if sim_name and not sim_name.startswith(default_prefix):
                return sim_name
    except Exception:
        pass

    return f"sim{sid}"

def set_sim_display_name(sim_id: object, display_name: object) -> str:
    """
    设置或清除某个 simulation 的自定义展示名。

    传参规则：
    - sim_id：目标 simulation id
    - display_name：新的展示名称
        * 非空：写入 / 覆盖自定义名称
        * 空字符串或空值：删除 override，恢复为默认 simN

    返回值：
    - 返回本次更新后“最终生效的展示名”。

    关键行为：
    1. sim_id 无法标准化时：
       - 不写入映射表；
       - 直接返回 display_name 的字符串形式。
    2. display_name 非空时：
       - 写入 _SIM_DISPLAY_NAME_MAP
    3. display_name 为空时：
       - 从映射表中移除该 sid
    4. 每次成功走到映射更新路径后：
       - version 自增 1

    设计意图：
    - “空字符串即清除”这套语义，便于 UI 输入框直接复用；
    - 上层不必额外区分“设置”与“删除”两套接口。
    """
    global _SIM_DISPLAY_NAME_VERSION

    sid = normalize_sim_id(sim_id)
    if sid is None:
        return str(display_name or "").strip()

    text = str(display_name or "").strip()

    if text:
        # 非空：写入 / 更新 override
        _SIM_DISPLAY_NAME_MAP[sid] = text
    else:
        # 空值：删除 override，恢复默认 simN
        _SIM_DISPLAY_NAME_MAP.pop(sid, None)

    # 命名有变化，版本号递增，供外层 UI 检测刷新
    _SIM_DISPLAY_NAME_VERSION += 1

    return get_sim_display_name(sid)


def get_sim_display_name_version() -> int:
    """
    获取当前展示名映射的版本号。

    使用场景：
    - 图表界面、图例、统计表等需要判断“重命名后是否需要重绘”；
    - 外层可缓存最近一次 version，发现变化后再刷新，避免无意义更新。

    返回值特点：
    - 单调递增整数；
    - 不保证连续语义，只保证“变了就更大”。
    """
    return int(_SIM_DISPLAY_NAME_VERSION)


def extract_sim_id_from_series_key(series_key: object) -> Optional[int]:
    """
    从序列键名中提取内部 sim_id。

    目标输入示例：
    - "风险值 (sim2)"
    - "Output A (SIM10)"

    返回规则：
    - 匹配成功：返回 int 形式 sim_id
    - 未匹配到 sim token：返回 None

    设计意图：
    - 某些图表系列 key 是“展示文本 + 内部 sim 标记”的复合格式；
    - 外层若要根据系列名反查其内部 simulation 身份，可通过本函数解析。

    注意：
    - 本函数只提取 "(simN)" 这种后缀 token；
    - 如果未来 series_key 编码规则改变，这里需要同步更新正则。
    """
    text = str(series_key or "")
    match = _SERIES_SIM_PATTERN.search(text)
    if not match:
        return None
    return normalize_sim_id(match.group(1))