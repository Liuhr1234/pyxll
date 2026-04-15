# ui_results_runtime_state.py
"""
本模块提供结果视图（Results Views）的运行时状态管理助手 ResultsRuntimeStateHelper。

主要功能：
1. 状态集中管控 (State Centralization)：接管主对话框 (dialog) 中散落的高级图表（龙卷风图、情景分析等）上下文状态的读写。
2. 边界清洗 (Normalization Boundary)：在数据存入 UI 状态机或被读取执行前，强制调用清洗引擎（如 normalize_tornado_config）进行数据校验与默认值回退，形成安全的防腐层。
3. 模式跃迁管理 (Mode Transitions)：提供清晰的接口用于在“基础图表视图”和“高级分析视图”之间安全地切换内部标识位。
"""

from __future__ import annotations

from ui_results_advanced_state import (
    default_scenario_config,
    default_tornado_config,
    normalize_scenario_config,
    normalize_tornado_config,
)


# =======================================================
# 1. 运行时状态管控助手
# =======================================================
class ResultsRuntimeStateHelper:
    """
    统一管理高级/基础模式切换时的运行时状态（Runtime State）。
    所有方法均为无状态的静态方法，通过操作传入的 dialog 实例的动态属性（如 `_scenario_config`）来工作。
    """

    # =======================================================
    # 2. 初始化与配置读写接口 (Getters / Setters)
    # =======================================================
    @staticmethod
    def init_advanced_configs(dialog):
        """
        [生命周期] 弹窗启动初始化：
        在弹窗首次加载时被调用，尝试读取可能存在的历史配置，
        并强制执行一次清洗规范化，将其安全地挂载为 dialog 的内部属性。
        """
        dialog._scenario_config = normalize_scenario_config(
            getattr(dialog, "_scenario_config", default_scenario_config())
        )
        dialog._tornado_config = normalize_tornado_config(
            getattr(dialog, "_tornado_config", default_tornado_config())
        )

    @staticmethod
    def get_scenario_config(dialog) -> dict:
        """
        获取情景分析 (Scenario) 配置：
        读取时强制清洗，确保即使在运行时被意外篡改，返回给渲染引擎的也是绝对合法的参数字典。
        """
        return normalize_scenario_config(getattr(dialog, "_scenario_config", None))

    @staticmethod
    def get_tornado_config(dialog) -> dict:
        """
        获取龙卷风图 (Tornado) 配置：
        返回一份经过数据类型安全校验和极值约束的配置字典副本。
        """
        return normalize_tornado_config(getattr(dialog, "_tornado_config", None))

    @staticmethod
    def set_scenario_config(dialog, cfg):
        """
        写入情景分析配置：
        作为防腐层的统一入口，任何尝试修改情景参数的操作都必须通过此方法，
        确保不合法的阈值（如下限 > 上限）在此处被自动纠正。
        """
        dialog._scenario_config = normalize_scenario_config(cfg)

    @staticmethod
    def set_tornado_config(dialog, cfg):
        """
        写入龙卷风图配置：
        防抖/清洗边界，过滤掉非法的分箱数量或无效的统计算法标识。
        """
        dialog._tornado_config = normalize_tornado_config(cfg)

    # =======================================================
    # 3. 运行模式标识位管理 (Mode Flags)
    # =======================================================
    @staticmethod
    def get_current_analysis_mode(dialog, default: str = "") -> str:
        """
        读取当前激活的高级分析模式 (如 "tornado", "scenario", "boxplot" 等)。
        使用统一的访问器避免硬编码属性名 `_current_analysis_mode` 散落在整个项目中。
        """
        return getattr(dialog, "_current_analysis_mode", default)

    @staticmethod
    def get_tornado_mode(dialog, default: str = "bins") -> str:
        """
        读取当前龙卷风图具体的子模式 (如分箱 "bins" 或 回归 "reg")。
        """
        return getattr(dialog, "_tornado_analysis_mode", default)

    @staticmethod
    def set_boxplot_mode(dialog, mode: str):
        """
        激活箱线图族群视图：
        将当前分析模式切换为箱线图 (boxplot)、小提琴图 (violin) 或字母值图等。
        """
        dialog._current_analysis_mode = mode

    @staticmethod
    def set_tornado_mode(dialog, mode: str):
        """
        激活龙卷风图视图：
        1. 记录具体的子算法模式（如基于方差还是基于相关系数）。
        2. 将主分析模式锁定为 "tornado"。
        """
        dialog._tornado_analysis_mode = mode
        dialog._current_analysis_mode = "tornado"

    @staticmethod
    def reset_tornado_mode(dialog, mode: str = "bins"):
        """
        重置龙卷风子模式：
        通常在用户退出高级分析并返回基础直方图时调用，将子模式重置为安全的默认值（分箱模式），
        以防下次进入时残留旧的不兼容状态。
        """
        dialog._tornado_analysis_mode = mode

    @staticmethod
    def set_scenario_mode(dialog, cfg=None):
        """
        激活情景分析视图：
        可选择性地传入新配置。在更新配置后，将主分析模式锁定为 "scenario"。
        """
        if cfg is not None:
            ResultsRuntimeStateHelper.set_scenario_config(dialog, cfg)
        dialog._current_analysis_mode = "scenario"

    @staticmethod
    def clear_advanced_mode(dialog):
        """
        退出高级分析视图：
        清除高级模式的专属标识符（置为空字符串），UI 的路由调度器在侦测到该状态后，
        将自动回退到基础渲染管线（直方图/CDF/离散图等）。
        """
        dialog._current_analysis_mode = ""

    # =======================================================
    # 4. 批量状态应用
    # =======================================================
    @staticmethod
    def apply_runtime_configs(dialog, *, scenario_cfg, tornado_cfg):
        """
        [事务性更新] 批量写入运行时配置：
        通常在视图状态快照（Snapshot）构建或恢复时使用，一次性将多个清洗后的配置对象挂载到上下文中。
        """
        ResultsRuntimeStateHelper.set_scenario_config(dialog, scenario_cfg)
        ResultsRuntimeStateHelper.set_tornado_config(dialog, tornado_cfg)