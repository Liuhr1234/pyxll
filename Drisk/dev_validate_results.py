# dev_validate_results.py
"""
【开发验证工具：结果分析模块 API 契约与边界测试】

主要功能：
1. 编译与导入验证：确保所有核心结果分析模块没有语法错误且能被正常加载。
2. 签名与契约验证：利用反射 (inspect) 检查关键类、函数是否存在，以及它们的参数列表是否符合预期边界要求。
3. 防止破坏性重构：在提交代码前运行，确保底层引擎 (sim_engine)、数据桥接层 (backend_bridge) 与 UI 层之间的调用纽带未被破坏。
"""

from __future__ import annotations

import importlib
import inspect
import py_compile
import sys
from pathlib import Path


# =======================================================
# 1. 目标文件与模块清单
# =======================================================
# 保持此列表专注于定义当前调用边界的结果分析相关模块。
TARGET_FILES = [
    "backend_bridge.py",
    "sim_engine.py",
    "ui_results.py",
    "ui_results_data_service.py",
    "ui_results_job_controller.py",
    "ui_results_dialogs.py",
    "ui_results_menu.py",
    "ui_results_modes.py",
    "ui_results_main_render.py",
    "ui_results_interactions.py",
    "ui_results_advanced.py",
    "ui_results_advanced_prepare.py",
    "ui_results_advanced_state.py",
    "ui_results_advanced_render.py",
    "ui_results_runtime_state.py",
    "ui_stats_contract.py",
    "ui_workers.py",
    "ui_stats.py",
    "ui_scatter.py",
    "drisk_charting.py",
    "drisk_export.py",
]

TARGET_IMPORTS = [
    "backend_bridge",
    "ui_results",
    "ui_results_data_service",
    "ui_results_job_controller",
    "ui_results_dialogs",
    "ui_results_menu",
    "ui_results_modes",
    "ui_results_main_render",
    "ui_results_interactions",
    "ui_results_advanced",
    "ui_results_advanced_prepare",
    "ui_results_advanced_state",
    "ui_results_advanced_render",
    "ui_results_runtime_state",
    "ui_stats_contract",
    "ui_workers",
    "ui_stats",
    "ui_scatter",
    "sim_engine",
    "drisk_export",
]


# =======================================================
# 2. 基础可用性校验 (编译与导入)
# =======================================================
def check_compile(project_root: Path) -> list[str]:
    """检查清单中的所有文件是否存在，并且没有 Python 语法编译错误。"""
    errors: list[str] = []
    for rel_path in TARGET_FILES:
        file_path = project_root / rel_path
        if not file_path.exists():
            errors.append(f"缺失文件: {rel_path}")
            continue
        try:
            py_compile.compile(str(file_path), doraise=True)
        except Exception as exc:  # pragma: no cover - 本地验证工具跳过覆盖率
            errors.append(f"编译失败: {rel_path} -> {exc}")
    return errors


def check_imports() -> list[str]:
    """检查清单中的模块是否都能被 importlib 成功导入（无循环依赖或运行时报错）。"""
    errors: list[str] = []
    for mod_name in TARGET_IMPORTS:
        try:
            importlib.import_module(mod_name)
        except Exception as exc:  # pragma: no cover - 本地验证工具跳过覆盖率
            errors.append(f"导入失败: {mod_name} -> {exc}")
    return errors


# =======================================================
# 3. 公共入口校验
# =======================================================
def check_public_entry() -> list[str]:
    """验证结果面板主模块提供的对外公共入口 (show_results_dialog) 及其参数签名。"""
    errors: list[str] = []
    module = importlib.import_module("ui_results")
    if not hasattr(module, "show_results_dialog"):
        errors.append("缺失公共入口: show_results_dialog")
        return errors

    # 使用反射检查函数签名，确保关键的必备参数未被意外删除或改名
    sig = inspect.signature(module.show_results_dialog)
    required_params = ["data", "sim_id", "cell_keys", "labels", "kind"]
    for name in required_params:
        if name not in sig.parameters:
            errors.append(f"show_results_dialog 缺失参数: {name}")
    return errors


# =======================================================
# 4. 内部扩展点与类结构校验
# =======================================================
def check_extension_points() -> list[str]:
    """验证拆分后的内部子模块中，核心的控制器、服务类和状态结构是否依然存在。"""
    errors: list[str] = []

    main_modes = importlib.import_module("ui_results_modes")
    for attr in ["ResultsViewCommandHelper", "ResultsRenderDispatcher"]:
        if not hasattr(main_modes, attr):
            errors.append(f"缺失主模式符号: {attr}")

    main_render = importlib.import_module("ui_results_main_render")
    for attr in ["ResultsMainRenderService"]:
        if not hasattr(main_render, attr):
            errors.append(f"缺失主渲染符号: {attr}")

    interaction_service = importlib.import_module("ui_results_interactions")
    for attr in ["ResultsInteractionService"]:
        if not hasattr(interaction_service, attr):
            errors.append(f"缺失交互符号: {attr}")

    advanced_router = importlib.import_module("ui_results_advanced")
    for attr in ["ResultsAdvancedRouter", "AdvancedModeSnapshot"]:
        if not hasattr(advanced_router, attr):
            errors.append(f"缺失高级路由符号: {attr}")

    advanced_prepare = importlib.import_module("ui_results_advanced_prepare")
    for attr in ["AdvancedPreparedPayload", "AdvancedPreparationError", "build_advanced_payload"]:
        if not hasattr(advanced_prepare, attr):
            errors.append(f"缺失高级载荷符号: {attr}")

    advanced_render = importlib.import_module("ui_results_advanced_render")
    for attr in ["ResultsAdvancedRuntimeState", "ResultsAdvancedRenderService"]:
        if not hasattr(advanced_render, attr):
            errors.append(f"缺失高级渲染符号: {attr}")

    runtime_state = importlib.import_module("ui_results_runtime_state")
    for attr in ["ResultsRuntimeStateHelper"]:
        if not hasattr(runtime_state, attr):
            errors.append(f"缺失运行时状态符号: {attr}")

    export_module = importlib.import_module("drisk_export")
    for attr in ["ExportContext", "DriskExportManager"]:
        if not hasattr(export_module, attr):
            errors.append(f"缺失导出符号: {attr}")

    return errors


# =======================================================
# 5. 跨模块边界契约校验
# =======================================================
def check_boundary_contracts() -> list[str]:
    """验证不同业务模块（如桥接层、散点图、导出管理器、底层引擎）之间的调用契约接口。"""
    errors: list[str] = []

    # 1. 验证后端数据桥接层
    bridge_module = importlib.import_module("backend_bridge")
    bridge_fn = getattr(bridge_module, "prepare_result_dialog_selection", None)
    if bridge_fn is None:
        errors.append("缺失边界函数: backend_bridge.prepare_result_dialog_selection")
    else:
        bridge_sig = inspect.signature(bridge_fn)
        expected = ["sim_id", "requested_keys", "default_sheet", "default_kind"]
        for name in expected:
            if name not in bridge_sig.parameters:
                errors.append(f"prepare_result_dialog_selection 缺失参数: {name}")

    # 2. 验证散点图模块
    scatter_module = importlib.import_module("ui_scatter")
    for attr in ["resolve_scatter_sim_id", "show_scatter_dialog_from_macro"]:
        if not hasattr(scatter_module, attr):
            errors.append(f"缺失散点图边界符号: {attr}")

    # 3. 验证报表导出模块
    export_module = importlib.import_module("drisk_export")
    export_mgr = getattr(export_module, "DriskExportManager", None)
    if export_mgr is None:
        errors.append("缺失导出管理器类: DriskExportManager")
    else:
        export_sig = inspect.signature(export_mgr.export_from_dialog)
        for name in ["dialog", "context"]:
            if name not in export_sig.parameters:
                errors.append(f"export_from_dialog 缺失参数: {name}")

        context_sig = inspect.signature(export_mgr.build_export_context)
        for name in ["dialog", "target_widget", "web_view", "chart_mode", "current_key"]:
            if name not in context_sig.parameters:
                errors.append(f"build_export_context 缺失参数: {name}")

    # 4. 验证 Excel 引擎层模块
    sim_module = importlib.import_module("sim_engine")
    for attr in ["smart_plot_macro", "open_scatter_viewer"]:
        if not hasattr(sim_module, attr):
            errors.append(f"缺失引擎入口符号: {attr}")

    return errors


# =======================================================
# 6. 主执行引擎
# =======================================================
def main() -> int:
    """按序执行所有验证项并打印控制台结果。"""
    project_root = Path(__file__).resolve().parent
    checks = [
        ("compile", check_compile(project_root)),
        ("imports", check_imports()),
        ("public-entry", check_public_entry()),
        ("extension-points", check_extension_points()),
        ("boundary-contracts", check_boundary_contracts()),
    ]

    has_error = False
    for name, errors in checks:
        if errors:
            has_error = True
            print(f"[失败 (FAIL)] {name}")
            for err in errors:
                print(f"  - {err}")
        else:
            print(f"[通过 (PASS)] {name}")

    if has_error:
        print("验证失败。")
        return 1

    print("验证通过。")
    return 0


if __name__ == "__main__":
    sys.exit(main())
