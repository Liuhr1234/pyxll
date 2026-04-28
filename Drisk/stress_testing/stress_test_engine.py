# -*- coding: utf-8 -*-
"""
压力测试执行引擎：
对每个 X 自变量单独施加极端值（下限/上限），记录 Y 因变量的响应，
最终将结果写入 Excel 新工作表。

同时支持条件筛选统计分析：根据X的条件筛选蒙特卡洛模拟数据，计算筛选后Y的平均值。
"""
from __future__ import annotations
import traceback
from pyxll import xl_app, xlcAlert
import sys
import os
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
from simulation_manager import get_simulation, get_current_sim_id
import numpy as np

#单元格辅助函数
def _resolve_range(sheet, addr: str):
    """将 'SheetName!A1' 或 'A1' 解析为 COM Range 对象。"""
    if "!" in addr:
        sname, a = addr.split("!", 1)
        s = xl_app().ActiveWorkbook.Sheets(sname)
    else:
        s = sheet
        a = addr
    return s.Range(a)


def _read_cell(sheet, addr: str) -> float:
    """读取单元格数值。"""
    val = _resolve_range(sheet, addr).Value
    return float(val) if val is not None else 0.0

def _restore_cell(sheet, addr: str, formula: str, value: float):
    """恢复单元格：若原来有公式则写回公式，否则写回数值。"""
    rng = _resolve_range(sheet, addr)
    if formula and formula.startswith("="):
        rng.Formula = formula
    else:
        rng.Value = value


def _calc_stressed_value(base: float, lo: float, hi: float, stress_type: str, use_hi: bool, x_data_range: float = None) -> float:
    """
    根据类型和方向计算压力值。

    参数：
        base: 基准值
        lo: 下限
        hi: 上限
        stress_type: "百分比 (%)" 或 "数值"
        use_hi: True 使用上限，False 使用下限
        x_data_range: X 数据的极差（max - min），用于百分比计算
    """
    delta = hi if use_hi else lo
    if stress_type == "百分比 (%)":
        if x_data_range is not None and x_data_range > 0:
            # 统一使用相对数据极差的计算方式
            return base + delta / 100.0 * x_data_range
        else:
            # 降级为相对基准值（兼容无缓存数据的情况）
            return base * (1.0 + delta / 100.0)
    else:
        return base + delta


def _get_latest_sim():
    """获取最新的模拟对象（_current_sim_id 是下一个ID，已有的最大ID是它减1）。"""
    from simulation_manager import _SIMULATION_CACHE
    if not _SIMULATION_CACHE:
        return None
    # 取 ID 最大的那个
    return _SIMULATION_CACHE[max(_SIMULATION_CACHE.keys())]

# 获取单元格模拟信息
def get_cell_cache_info(cell_addr: str, sim_id: int = None) -> dict:
    """
    获取指定单元格的缓存信息，优先查 output_cache，其次 input_cache。

    参数：
        cell_addr: 单元格地址，格式为 'A1' 或 'Sheet1!A1'
        sim_id: 指定模拟ID；None 则取最新模拟

    返回：
        {'found': bool, 'type': str, 'data': ndarray, 'attributes': dict,
         'n_iterations': int, 'method': str}
    """
    if sim_id is not None:
        sim = get_simulation(sim_id)
    else:
        sim = _get_latest_sim()

    if sim is None:
        return {'found': False, 'error': '未找到模拟缓存'}

    # 标准化：去 $ 、大写、补 sheet 前缀
    addr_clean = cell_addr.replace('$', '').upper()
    if '!' not in addr_clean:
        try:
            sheet_name = xl_app().ActiveSheet.Name
            addr_clean = f"{sheet_name}!{addr_clean}"
        except Exception:
            pass

    addr_cell = addr_clean.split('!')[-1]   # 纯单元格部分，如 "A1"

    # --- 查 output_cache（键格式：SHEET!CELL，全大写无$）---
    data = sim.get_output_data(addr_clean)
    if data is not None:
        return {
            'found': True,
            'type': 'output',
            'data': data,
            'attributes': sim.get_output_attributes(addr_clean),
            'n_iterations': sim.n_iterations,
            'method': sim.sampling_method,
        }

    # --- 查 input_cache（键格式：SHEET!CELL_N，全大写无$）---
    # 遍历所有 input key，匹配 SHEET!CELL_N 或仅 CELL_N
    for cache_key, cache_data in (sim.input_cache or {}).items():
        key_clean = str(cache_key).replace('$', '').upper()
        # key_clean 形如 "SHEET1!A1_1"
        key_cell_part = key_clean.split('!')[-1]          # "A1_1"
        key_cell_base = key_cell_part.rsplit('_', 1)[0]   # "A1"
        if key_cell_base == addr_cell:
            # 若地址带 sheet，还要校验 sheet 名一致
            if '!' in addr_clean:
                addr_sheet = addr_clean.split('!')[0]
                key_sheet = key_clean.split('!')[0] if '!' in key_clean else ''
                if key_sheet and key_sheet != addr_sheet:
                    continue
            return {
                'found': True,
                'type': 'input',
                'data': cache_data,
                'attributes': (sim.input_attributes or {}).get(cache_key, {}),
                'n_iterations': sim.n_iterations,
                'method': sim.sampling_method,
            }

    return {'found': False, 'error': f'单元格 {cell_addr} 未找到缓存数据'}


def run_stress_test(cfg: dict) -> list[dict]:
    """
    执行压力测试

    cfg 结构：
        y_cell: str               — 因变量单元格地址
        x_configs: list[dict]     — 每个自变量的配置 {addr, type, lo, hi}

    返回：每个场景的结果列表
    """
    app = xl_app()
    wb = app.ActiveWorkbook
    active_sheet = app.ActiveSheet
    sheet_name = active_sheet.Name

    y_addr = cfg["y_cell"]
    x_configs = cfg["x_configs"]

    # 先检查Y单元格的缓存是否存在，若无则自动触发模拟
    y_cache = get_cell_cache_info(y_addr)

    if not y_cache['found']:
        try:
            import sim_engine
            sim_engine._run_batch_simulations("MC")
        except Exception as e:
            xlcAlert(f"自动运行模拟失败：{e}\n\n请手动运行模拟后再执行压力测试。")
            return []
        y_cache = get_cell_cache_info(y_addr)
        if not y_cache['found']:
            xlcAlert(f"模拟已运行但仍未找到单元格 {y_addr} 的缓存数据。\n请确认该单元格已被标记为输出或含有分布公式。")
            return []

    # 读取基准值
    y_base = _read_cell(active_sheet, y_addr)
    x_bases = {c["addr"]: _read_cell(active_sheet, c["addr"]) for c in x_configs}

    results = []

    # 预先获取所有 X 的数据极差（用于百分比计算）
    x_data_ranges = {}
    for xc in x_configs:
        addr = xc["addr"]
        full_addr = f"{sheet_name}!{addr}" if "!" not in addr else addr
        x_cache = get_cell_cache_info(full_addr)
        if x_cache['found']:
            x_data = np.asarray(x_cache['data'])
            x_data_ranges[addr] = float(np.max(x_data)) - float(np.min(x_data))
        else:
            x_data_ranges[addr] = None

    for xc in x_configs:
        addr = xc["addr"]
        stress_type = xc["type"]
        lo = xc["lo"]
        hi = xc["hi"]
        x_base = x_bases[addr]
        # 读取单元格原始公式（无公式则返回空字符串）。
        x_formula = str(_resolve_range(active_sheet, addr).Formula or "")
    

        x_data_range = x_data_ranges.get(addr)

        # 构建完整的带工作表名的地址
        if "!" not in addr:
            full_addr = f"{sheet_name}!{addr}"
        else:
            full_addr = addr

        for direction, use_hi in [("下限", False), ("上限", True)]:
            x_stressed = _calc_stressed_value(x_base, lo, hi, stress_type, use_hi, x_data_range)
            try:
                """写入单元格数值。"""
                _resolve_range(active_sheet, addr).Value = x_stressed
                app.Calculate()
                y_stressed = _read_cell(active_sheet, y_addr)
            except Exception:
                y_stressed = float("nan")
            finally:
                _restore_cell(active_sheet, addr, x_formula, x_base)

            delta_y = y_stressed - y_base
            delta_y_pct = (delta_y / y_base * 100.0) if y_base != 0 else float("nan")

            delta_x = x_stressed - x_base
            delta_x_pct = (delta_x / x_base * 100.0) if x_base != 0 else float("nan")

            results.append({
                "label": f"{addr} {direction}",
                "x_addr": full_addr,
                "direction": direction,
                "x_base": x_base,
                "x_stressed": x_stressed,
                "delta_x": delta_x,
                "delta_x_pct": delta_x_pct,
                "y_base": y_base,
                "y_stressed": y_stressed,
                "delta_y": delta_y,
                "delta_y_pct": delta_y_pct,
            })

    # 执行多场景条件筛选统计分析
    _perform_multi_scenario_analysis(cfg, y_cache, sheet_name)

    return results

def _perform_multi_scenario_analysis(cfg: dict, y_cache: dict, sheet_name: str):
    """
    执行多场景条件筛选统计分析：
    1. 为每个 X 单独创建筛选场景
    2. 创建全部 X 联合筛选场景

    每个场景写入独立的模拟缓存，名称标识不同。
    """
    y_addr = cfg["y_cell"]
    x_configs = cfg["x_configs"]

    if not x_configs:
        return

    analyze_single = cfg.get("analyze_single", True)
    count = 0

    if analyze_single:
        # 场景1：每个 X 单独筛选
        for xc in x_configs:
            single_cfg = {"y_cell": y_addr, "x_configs": [xc]}
            addr = xc["addr"]
            if "!" in addr:
                sheet_part, cell_part = addr.split("!", 1)
                x_label = f"{sheet_part}!{cell_part}"
            else:
                x_label = addr
            _perform_conditional_analysis(single_cfg, y_cache, sheet_name, scenario_name=f"单变量_{x_label}")
            count += 1
    else:
        # 场景2：全部 X 联合筛选
        _perform_conditional_analysis(cfg, y_cache, sheet_name, scenario_name="全部变量联合")
        count += 1
    y_data = np.asarray(y_cache['data'])
    total_samples = len(y_data)

    title= ["目前的压力测试设置：",
	   f"模拟测试：{count}次",
	   f"每次模拟迭代次数：{total_samples}次",
	   f"全部迭代次数：{count * total_samples}次"]
    xlcAlert("\n\n".join(title))

def _perform_conditional_analysis(cfg: dict, y_cache: dict, sheet_name: str, scenario_name: str = "压力筛选") -> str:
    """
    执行条件筛选统计分析，返回结果消息字符串（由调用方统一弹窗）。
    参数：
        cfg: 压力测试配置
        y_cache: Y单元格的缓存信息
        sheet_name: 工作表名称
        scenario_name: 场景名称，用于缓存命名
    返回：结果消息字符串，失败时返回空字符串
    """
    print("Performing conditional analysis...")
    try:
        y_addr = cfg["y_cell"]
        x_configs = cfg["x_configs"]

        # 获取Y的数据并转为 numpy 数组
        y_data = np.asarray(y_cache['data'])
        total_samples = len(y_data)

        # 初始化筛选掩码
        combined_mask = np.ones(total_samples, dtype=bool)

        # 应用所有X条件
        for xc in x_configs:
            x_addr = xc["addr"]

            # 构建完整地址
            if "!" not in x_addr:
                full_x_addr = f"{sheet_name}!{x_addr}"
            else:
                full_x_addr = x_addr

            x_cache = get_cell_cache_info(full_x_addr)

            if not x_cache['found']:
                print(f"未找到 {full_x_addr} 的缓存数据")
                continue

            x_data = np.asarray(x_cache['data'])

            if len(x_data) != total_samples:
                print(f"{full_x_addr} 样本数不匹配: {len(x_data)} vs {total_samples}")
                continue

            # 根据压力范围创建筛选条件
            stress_type = xc["type"]
            lo = xc["lo"]
            hi = xc["hi"]

            # 读取X的基准值
            try:
                app = xl_app()
                x_base = _read_cell(app.ActiveSheet, x_addr)
            except:
                continue

            # 计算压力值范围
            if stress_type == "百分比 (%)":
                # 百分比相对整个数据集的极差（max - min）计算偏移量
                x_range = float(np.max(x_data)) - float(np.min(x_data))
                x_min = float(np.min(x_data)) + lo / 100.0 * x_range
                x_max = float(np.min(x_data)) + hi / 100.0 * x_range
            else:
                x_min = lo
                x_max = hi

            # 筛选：保留在压力范围内的样本
            x_mask = (x_data >= x_min) & (x_data <= x_max)
            combined_mask &= x_mask

        # 应用筛选
        filtered_y_data = y_data[combined_mask]
        filtered_samples = len(filtered_y_data)

        if filtered_samples == 0:
            return f"【{scenario_name}】可能条件有误，无法生成符合条件的样本。"

        # 有放回重采样：将筛选后数据扩充回原始规模，保持分布形态不变
        rng_gen = np.random.default_rng()
        resample_idx = rng_gen.integers(0, filtered_samples, size=total_samples)
        resampled_y_data = filtered_y_data[resample_idx]

        # 计算筛选后Y的平均值（基于重采样后数据）
        y_mean = float(np.mean(resampled_y_data))

        # ---- 将筛选结果写入新的模拟缓存 ----
        from simulation_manager import create_simulation, _SIMULATION_CACHE
        filtered_sim_id = create_simulation(n_iterations=total_samples, sampling_method="stress_filter")
        filtered_sim = _SIMULATION_CACHE[filtered_sim_id]
        filtered_sim.name = f"sim{filtered_sim_id}({scenario_name})"

        # 写入重采样后的 Y 数据
        y_sheet = y_addr.split("!")[0] if "!" in y_addr else sheet_name
        y_cell = y_addr.split("!")[-1] if "!" in y_addr else y_addr
        y_attrs = dict(y_cache.get("attributes") or {})
        y_attrs["stress_filtered"] = True
        y_attrs["filtered_samples"] = filtered_samples
        y_attrs["total_samples"] = total_samples
        filtered_sim.add_output_result(y_cell, resampled_y_data.astype(float), y_sheet, y_attrs)

        # 写入重采样后的各 X 数据（与 Y 使用同一组 resample_idx 保持行对齐）
        for xc in x_configs:
            x_addr = xc["addr"]
            full_x_addr = f"{sheet_name}!{x_addr}" if "!" not in x_addr else x_addr
            x_cache = get_cell_cache_info(full_x_addr)
            if not x_cache["found"]:
                continue
            x_data_full = np.asarray(x_cache["data"])
            if len(x_data_full) != total_samples:
                continue
            filtered_x_data = x_data_full[combined_mask][resample_idx]
            x_sheet = full_x_addr.split("!")[0]
            x_cell = full_x_addr.split("!")[-1]
            x_attrs = dict(x_cache.get("attributes") or {})
            x_attrs["stress_filtered"] = True
            filtered_sim.add_input_result(f"{x_cell}_1", filtered_x_data.astype(float), x_sheet, x_attrs)

        print(f"[压力测试] {scenario_name} 筛选结果已写入缓存 sim_id={filtered_sim_id}，原始筛选样本数={filtered_samples}，重采样至={total_samples}")


    except Exception as e:
        print(f"[_perform_conditional_analysis] 错误: {e}\n{traceback.format_exc()}")
