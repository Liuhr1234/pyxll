# -*- coding: utf-8 -*-
"""
条件筛选统计分析引擎
根据自变量条件筛选蒙特卡洛模拟数据，计算因变量的统计量
"""
import numpy as np
from stress_test_engine import get_cell_cache_info


def apply_filter_conditions(x_data: np.ndarray, conditions: dict) -> np.ndarray:
    """
    根据条件筛选数据索引。

    参数：
        x_data: X变量的numpy数组
        conditions: 筛选条件字典
            {
                'type': 'range',  # 'range', 'percentile', 'value'
                'min': float,     # 最小值（range类型）
                'max': float,     # 最大值（range类型）
                'p_min': float,   # 最小百分位（percentile类型）
                'p_max': float,   # 最大百分位（percentile类型）
                'operator': '>', # 比较运算符（value类型）
                'value': float    # 比较值（value类型）
            }

    返回：
        满足条件的索引数组（布尔数组）
    """
    cond_type = conditions.get('type', 'range')

    if cond_type == 'range':
        # 范围筛选：min <= x <= max
        min_val = float(conditions.get('min', -np.inf))
        max_val = float(conditions.get('max', np.inf))
        mask = (x_data >= min_val) & (x_data <= max_val)

    elif cond_type == 'percentile':
        # 百分位筛选：p_min <= x <= p_max
        p_min = float(conditions.get('p_min', 0))
        p_max = float(conditions.get('p_max', 100))
        min_val = np.percentile(x_data, p_min)
        max_val = np.percentile(x_data, p_max)
        mask = (x_data >= min_val) & (x_data <= max_val)

    elif cond_type == 'value':
        # 值比较筛选：x > value, x < value, x == value
        operator = conditions.get('operator', '>')
        value = float(conditions.get('value', 0))

        if operator == '>':
            mask = x_data > value
        elif operator == '>=':
            mask = x_data >= value
        elif operator == '<':
            mask = x_data < value
        elif operator == '<=':
            mask = x_data <= value
        elif operator == '==':
            mask = np.isclose(x_data, value)
        elif operator == '!=':
            mask = ~np.isclose(x_data, value)
        else:
            mask = np.ones(len(x_data), dtype=bool)

    else:
        # 默认：不筛选
        mask = np.ones(len(x_data), dtype=bool)

    return mask


def conditional_analysis(y_addr: str, x_conditions: list[dict], sim_id: int = None) -> dict:
    """
    执行条件筛选统计分析。

    参数：
        y_addr: 因变量单元格地址
        x_conditions: 自变量筛选条件列表
            [
                {
                    'addr': 'A1',
                    'type': 'range',
                    'min': 80,
                    'max': 120
                },
                ...
            ]
        sim_id: 模拟ID

    返回：
        {
            'success': bool,
            'error': str,
            'y_addr': str,
            'y_cache_found': bool,
            'total_samples': int,
            'filtered_samples': int,
            'filter_ratio': float,
            'y_stats': {
                'mean': float,
                'std': float,
                'min': float,
                'max': float,
                'p5': float,
                'p25': float,
                'p50': float,
                'p75': float,
                'p95': float
            },
            'x_filters': list[dict]
        }
    """
    try:
        # 1. 获取Y的缓存数据
        y_cache = get_cell_cache_info(y_addr, sim_id)

        if not y_cache['found']:
            return {
                'success': False,
                'error': f'未找到单元格 {y_addr} 的模拟缓存数据。\n请先运行蒙特卡洛模拟。',
                'y_addr': y_addr,
                'y_cache_found': False
            }

        y_data = y_cache['data']
        total_samples = len(y_data)

        # 2. 初始化筛选掩码（全部为True）
        combined_mask = np.ones(total_samples, dtype=bool)

        # 3. 应用所有X条件
        x_filter_results = []

        for x_cond in x_conditions:
            x_addr = x_cond.get('addr')
            if not x_addr:
                continue

            # 获取X的缓存数据
            x_cache = get_cell_cache_info(x_addr, sim_id)

            if not x_cache['found']:
                return {
                    'success': False,
                    'error': f'未找到单元格 {x_addr} 的模拟缓存数据。\n请先运行蒙特卡洛模拟。',
                    'y_addr': y_addr,
                    'y_cache_found': True
                }

            x_data = x_cache['data']

            # 检查数据长度是否匹配
            if len(x_data) != total_samples:
                return {
                    'success': False,
                    'error': f'单元格 {x_addr} 的样本数量 ({len(x_data)}) 与 {y_addr} ({total_samples}) 不匹配。',
                    'y_addr': y_addr,
                    'y_cache_found': True
                }

            # 应用筛选条件
            x_mask = apply_filter_conditions(x_data, x_cond)
            combined_mask &= x_mask

            # 记录筛选结果
            x_filter_results.append({
                'addr': x_addr,
                'condition': x_cond,
                'matched_count': int(np.sum(x_mask)),
                'matched_ratio': float(np.sum(x_mask) / total_samples)
            })

        # 4. 应用组合筛选掩码
        filtered_y_data = y_data[combined_mask]
        filtered_samples = len(filtered_y_data)

        if filtered_samples == 0:
            return {
                'success': False,
                'error': '筛选条件过于严格，没有样本满足所有条件。\n请放宽筛选条件。',
                'y_addr': y_addr,
                'y_cache_found': True,
                'total_samples': total_samples,
                'filtered_samples': 0,
                'x_filters': x_filter_results
            }

        # 5. 计算筛选后Y的统计量
        y_stats = {
            'mean': float(np.mean(filtered_y_data)),
            'std': float(np.std(filtered_y_data)),
            'min': float(np.min(filtered_y_data)),
            'max': float(np.max(filtered_y_data)),
            'p5': float(np.percentile(filtered_y_data, 5)),
            'p25': float(np.percentile(filtered_y_data, 25)),
            'p50': float(np.percentile(filtered_y_data, 50)),
            'p75': float(np.percentile(filtered_y_data, 75)),
            'p95': float(np.percentile(filtered_y_data, 95))
        }

        return {
            'success': True,
            'y_addr': y_addr,
            'y_cache_found': True,
            'total_samples': total_samples,
            'filtered_samples': filtered_samples,
            'filter_ratio': float(filtered_samples / total_samples),
            'y_stats': y_stats,
            'x_filters': x_filter_results
        }

    except Exception as e:
        import traceback
        return {
            'success': False,
            'error': f'条件筛选分析失败：\n{e}\n\n{traceback.format_exc()}',
            'y_addr': y_addr,
            'y_cache_found': False
        }
