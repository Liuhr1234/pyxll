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


def _read_cell(sheet, addr: str) -> float:
    """读取单元格数值，addr 格式为 'SheetName!A1' 或 'A1'。"""
    if "!" in addr:
        sname, a = addr.split("!", 1)
        wb = xl_app().ActiveWorkbook
        s = wb.Sheets(sname)
    else:
        s = sheet
        a = addr
    val = s.Range(a).Value
    return float(val) if val is not None else 0.0


def _write_cell(sheet, addr: str, value):
    if "!" in addr:
        sname, a = addr.split("!", 1)
        wb = xl_app().ActiveWorkbook
        s = wb.Sheets(sname)
    else:
        s = sheet
        a = addr
    s.Range(a).Value = value


def _calc_stressed_value(base: float, lo: float, hi: float, stress_type: str, use_hi: bool) -> float:
    """根据类型和方向计算压力值。"""
    delta = hi if use_hi else lo
    if stress_type == "百分比 (%)":
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
    执行压力测试并将结果写入新工作表。

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

    # 先检查Y单元格的缓存是否存在
    y_cache = get_cell_cache_info(y_addr)

    if not y_cache['found']:
        # 缓存不存在，弹出提示并终止
        xlcAlert(f"未找到单元格 {y_addr} 的模拟缓存数据。\n\n请先运行蒙特卡洛模拟后再执行压力测试。")
        return []

    # 读取基准值
    y_base = _read_cell(active_sheet, y_addr)
    x_bases = {c["addr"]: _read_cell(active_sheet, c["addr"]) for c in x_configs}

    results = []

    for xc in x_configs:
        addr = xc["addr"]
        stress_type = xc["type"]
        lo = xc["lo"]
        hi = xc["hi"]
        x_base = x_bases[addr]

        # 构建完整的带工作表名的地址
        if "!" not in addr:
            full_addr = f"{sheet_name}!{addr}"
        else:
            full_addr = addr

        for direction, use_hi in [("下限", False), ("上限", True)]:
            x_stressed = _calc_stressed_value(x_base, lo, hi, stress_type, use_hi)
            try:
                _write_cell(active_sheet, addr, x_stressed)
                app.Calculate()
                y_stressed = _read_cell(active_sheet, y_addr)
            except Exception:
                y_stressed = float("nan")
            finally:
                _write_cell(active_sheet, addr, x_base)  # 恢复原值

            delta_y = y_stressed - y_base
            delta_y_pct = (delta_y / y_base * 100.0) if y_base != 0 else float("nan")

            # 计算X的变化比例
            delta_x = x_stressed - x_base
            delta_x_pct = (delta_x / x_base * 100.0) if x_base != 0 else float("nan")

            results.append({
                "label": f"{addr} {direction}",
                "x_addr": full_addr,  # 使用完整地址
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
    
    _write_results_to_sheet(wb, y_addr, results)

    # 执行条件筛选统计分析（传递工作表名）
    _perform_conditional_analysis(cfg, y_cache, sheet_name)

    # 显示压力测试分析窗口（传递 Y 地址和缓存）
    try:
        import importlib
        import stress_testing.stress_chart as stress_chart_module
        importlib.reload(stress_chart_module)

        # 构建完整的 Y 地址
        if "!" not in y_addr:
            full_y_addr = f"{sheet_name}!{y_addr}"
        else:
            full_y_addr = y_addr

        stress_chart_module.show_stress_test_dialog(results, full_y_addr, y_cache)
    except Exception as e:
        import traceback
        print(f"显示压力测试分析窗口失败: {e}\n{traceback.format_exc()}")

    return results


def _perform_conditional_analysis(cfg: dict, y_cache: dict, sheet_name: str):
    """
    执行条件筛选统计分析并弹窗显示结果。

    参数：
        cfg: 压力测试配置
        y_cache: Y单元格的缓存信息
        sheet_name: 工作表名称
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
            x_min = _calc_stressed_value(x_base, lo, hi, stress_type, False)
            x_max = _calc_stressed_value(x_base, lo, hi, stress_type, True)

            # 筛选：保留在压力范围内的样本
            x_mask = (x_data >= x_min) & (x_data <= x_max)
            combined_mask &= x_mask

        # 应用筛选
        filtered_y_data = y_data[combined_mask]
        filtered_samples = len(filtered_y_data)

        if filtered_samples == 0:
            xlcAlert("筛选条件过于严格，没有样本满足所有条件。")
            return

        # 计算筛选后Y的平均值
        y_mean = float(np.mean(filtered_y_data))

        # 弹窗显示结果
        msg = f"条件筛选统计分析结果\n\n"
        msg += f"因变量: {y_addr}\n"
        msg += f"总样本数: {total_samples}\n"
        msg += f"筛选后样本数: {filtered_samples}\n"
        msg += f"筛选比例: {filtered_samples/total_samples*100:.2f}%\n\n"
        msg += f"筛选后因变量平均值: {y_mean:.4f}\n"

        xlcAlert(msg)

    except Exception as e:
        print(f"[_perform_conditional_analysis] 错误: {e}\n{traceback.format_exc()}")


def _write_results_to_sheet(wb, y_addr: str, results: list[dict]):
    """将压力测试结果写入新工作表（若已存在则覆盖）。"""
    sheet_name = "压力测试结果"
    try:
        ws = wb.Sheets(sheet_name)
        ws.Cells.Clear()
    except Exception:
        ws = wb.Sheets.Add(After=wb.Sheets(wb.Sheets.Count))
        ws.Name = sheet_name

    headers = ["场景", "自变量单元格", "方向", "X 基准值", "X 压力值", "ΔX", "ΔX (%)", "Y 基准值", "Y 压力值", "ΔY", "ΔY (%)"]
    for col, h in enumerate(headers, start=1):
        ws.Cells(1, col).Value = h
        ws.Cells(1, col).Font.Bold = True

    for row, r in enumerate(results, start=2):
        ws.Cells(row, 1).Value = r["label"]
        ws.Cells(row, 2).Value = r["x_addr"]
        ws.Cells(row, 3).Value = r["direction"]
        ws.Cells(row, 4).Value = r["x_base"]
        ws.Cells(row, 5).Value = r["x_stressed"]
        ws.Cells(row, 6).Value = r["delta_x"]
        ws.Cells(row, 7).Value = r["delta_x_pct"]
        ws.Cells(row, 8).Value = r["y_base"]
        ws.Cells(row, 9).Value = r["y_stressed"]
        ws.Cells(row, 10).Value = r["delta_y"]
        ws.Cells(row, 11).Value = r["delta_y_pct"]

    ws.Columns.AutoFit()
    ws.Activate()
