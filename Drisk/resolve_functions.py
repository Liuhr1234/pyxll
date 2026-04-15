# resolve_functions.py
"""
Excel 内置函数向量化实现模块（包括 NPV、PV、SUBTOTAL、SUMPRODUCT 等财务与数组函数）
供 numpy_functions.py 调用
"""

import math
import datetime
import re
import numpy as np
from pyxll import xl_app

from attribute_functions import ERROR_MARKER
from distribution_functions import _is_excel_error

def _range_to_coords(rng: str) -> list:
    """将区域引用（如 A1:B2）展开为单元格坐标列表（大写字母，已去除 $）"""
    # 去除 $ 符号
    rng = rng.replace('$', '')
    m = re.match(r'([A-Z]+)(\d+):([A-Z]+)(\d+)', rng, re.I)
    if not m:
        return [rng.upper()]
    sc, sr, ec, er = m.groups()
    coords = []
    def col2n(c: str) -> int:
        c = c.replace('$', '')
        n = 0
        for ch in c.upper():
            n = n * 26 + ord(ch) - 64
        return n
    def n2col(n: int) -> str:
        s = ''
        while n > 0:
            n -= 1
            s = chr(n % 26 + 65) + s
            n //= 26
        return s
    for row in range(int(sr), int(er) + 1):
        for cn in range(col2n(sc), col2n(ec) + 1):
            coords.append(f"{n2col(cn)}{row}")
    return coords

# ==================== 日期时间辅助函数（用于内置函数） ====================
_ExcelDateBase = datetime.date(1899, 12, 31)  # Excel 日期系统基准（0 对应 1899-12-31）

def _excel_date_to_date(serial: float) -> datetime.date:
    """将 Excel 日期序列号转换为 datetime.date"""
    return _ExcelDateBase + datetime.timedelta(days=int(serial))

def _date_to_excel_date(d: datetime.date) -> float:
    """将 datetime.date 转换为 Excel 日期序列号"""
    return float((d - _ExcelDateBase).days)

def _excel_datetime_to_float(dt: datetime.datetime) -> float:
    """将 datetime.datetime 转换为 Excel 日期时间浮点数"""
    days = (dt.date() - _ExcelDateBase).days
    seconds = dt.hour * 3600 + dt.minute * 60 + dt.second + dt.microsecond / 1e6
    return days + seconds / 86400.0

# ==================== SUBTOTAL 辅助函数 ====================
# 缓存区域的行隐藏状态和值数组（避免重复读取 Excel）
_SUBTOTAL_CACHE = {}

def _subtotal_core(
    function_num,  # 可以是 int 或字符串
    ref_strs,
    n: int,
    static_values: dict,
    live_arrays: dict,
    default_sheet: str,
    app,
    cell_formulas: dict = None   # 新增参数，默认为空字典
) -> np.ndarray:
    """
    向量化 SUBTOTAL 实现，正确处理错误值传播，修正区域引用为绝对坐标。
    若任何引用的单元格为 #ERROR!，则整个 SUBTOTAL 返回 #ERROR!。
    新增：忽略参数中包含 SUBTOTAL 函数的单元格。
    """
    # 确保 function_num 是整数
    try:
        func_num = int(function_num)
    except:
        return np.full(n, ERROR_MARKER, dtype=object)

    # 确定是否忽略隐藏行
    ignore_hidden = 101 <= func_num <= 111
    if ignore_hidden:
        func_num -= 100
    if not (1 <= func_num <= 11):
        return np.full(n, ERROR_MARKER, dtype=object)

    # 如果 ref_strs 是字符串，转换为列表
    if isinstance(ref_strs, str):
        ref_strs = [ref_strs]

    # 收集所有单元格的值数组（每个数组长度为 n）
    all_values_by_cell = []
    seen_cells = set()   # 记录已处理的单元格

    for ref_str in ref_strs:
        # 解析区域引用
        ref_str = ref_str.strip()
        if '!' in ref_str:
            sheet_part, addr_part = ref_str.split('!', 1)
            sheet_part = sheet_part.strip()
            if sheet_part.startswith("'") and sheet_part.endswith("'"):
                sheet_part = sheet_part[1:-1].replace("''", "'")
            sheet_name = sheet_part
            range_addr = addr_part
        else:
            sheet_name = default_sheet
            range_addr = ref_str
        sheet_name = sheet_name.strip()
        if sheet_name.startswith("'") and sheet_name.endswith("'"):
            sheet_name = sheet_name[1:-1].replace("''", "'")

        # 获取区域对象及起始行列
        try:
            sheet = app.ActiveWorkbook.Worksheets(sheet_name)
            range_obj = sheet.Range(range_addr)
            start_row = range_obj.Row
            start_col = range_obj.Column
            rows = range_obj.Rows.Count
            cols = range_obj.Columns.Count
        except Exception as e:
            # 无效区域，跳过
            continue

        # 获取隐藏行信息（缓存）
        cache_key = f"{sheet_name}!{range_addr}".upper()
        if cache_key in _SUBTOTAL_CACHE:
            hidden_rows = _SUBTOTAL_CACHE[cache_key]['hidden']
        else:
            # 必须用绝对行号构建隐藏掩码
            hidden_rows = []
            for r in range(rows):
                abs_row = start_row + r
                try:
                    cell = sheet.Cells(abs_row, start_col)  # 取区域内第一列判断隐藏
                    hidden_rows.append(cell.EntireRow.Hidden)
                except:
                    hidden_rows.append(False)
            hidden_rows = np.array(hidden_rows, dtype=bool)
            _SUBTOTAL_CACHE[cache_key] = {'hidden': hidden_rows}

        # 收集区域内每个单元格的值数组（使用绝对坐标）
        values_by_cell = []
        for r in range(rows):
            abs_row = start_row + r
            for c in range(cols):
                abs_col = start_col + c
                col_letter = _n2col(abs_col)
                cell_ref = f"{sheet_name}!{col_letter}{abs_row}".upper()
                # 去重检查
                if cell_ref in seen_cells:
                    continue
                seen_cells.add(cell_ref)

                # 检查单元格公式是否包含 SUBTOTAL
                if cell_formulas is not None:
                    formula = cell_formulas.get(cell_ref, '')
                    if 'SUBTOTAL' in formula.upper():
                        # 忽略该单元格
                        continue

                arr = None
                if cell_ref in live_arrays:
                    arr = live_arrays[cell_ref]
                    # 确保是一维数组
                    if arr.ndim != 1:
                        arr = arr.flatten()
                    # 调整长度
                    if len(arr) != n:
                        if len(arr) > n:
                            arr = arr[:n]
                        else:
                            arr = np.pad(arr, (0, n - len(arr)), constant_values=ERROR_MARKER)
                else:
                    # 尝试静态值
                    val = static_values.get(cell_ref)
                    if val is None:
                        # 空单元格 -> 返回 NaN 数组
                        arr = np.full(n, np.nan, dtype=float)
                    elif val == ERROR_MARKER or (isinstance(val, str) and val.upper() == "#ERROR!"):
                        arr = np.full(n, ERROR_MARKER, dtype=object)
                    else:
                        try:
                            fval = float(val)
                            arr = np.full(n, fval, dtype=float)
                        except:
                            # 无法转换为数值（如文本），视为 NaN
                            arr = np.full(n, np.nan, dtype=float)
                values_by_cell.append(arr)

        # 若忽略隐藏行，则过滤掉隐藏行对应的单元格
        if ignore_hidden:
            # 为每个单元格生成掩码：是否可见（基于其绝对行号）
            cell_mask = []
            for r in range(rows):
                row_hidden = hidden_rows[r]
                for _ in range(cols):
                    cell_mask.append(not row_hidden)
            cell_mask = np.array(cell_mask)
            values_by_cell = [arr for idx, arr in enumerate(values_by_cell) if cell_mask[idx]]

        all_values_by_cell.extend(values_by_cell)

    if not all_values_by_cell:
        return np.full(n, ERROR_MARKER, dtype=object)

    # ==================== 将所有数组转换为浮点数数组，错误标记转为 NaN ====================
    float_arrays = []
    for arr in all_values_by_cell:
        if arr.dtype.kind == 'O':  # 对象数组
            # 转换为浮点数，将 "#ERROR!" 等转换为 np.nan
            arr_float = np.full(arr.shape, np.nan, dtype=float)
            try:
                # 找出错误标记的位置
                mask = (arr == ERROR_MARKER)
                # 非错误元素尝试转换为浮点数
                valid_mask = ~mask
                arr_float[valid_mask] = arr[valid_mask].astype(float)
            except Exception:
                # 如果转换失败，全部设为 NaN
                arr_float = np.full(arr.shape, np.nan, dtype=float)
            float_arrays.append(arr_float)
        elif arr.dtype.kind in 'fc':  # 已经是浮点或复数数组，转为浮点
            float_arrays.append(arr.astype(float))
        else:
            # 其他类型（如 bool）转为浮点数
            float_arrays.append(arr.astype(float))

    # 堆叠为二维数组 (num_cells, n)
    col_data = np.vstack(float_arrays)

    # 根据 func_num 应用聚合函数（使用 nan-safe 函数）
    if func_num == 1:      # AVERAGE
        result = np.nanmean(col_data, axis=0)
    elif func_num == 2:    # COUNT
        valid = ~np.isnan(col_data)
        result = np.sum(valid, axis=0).astype(float)
    elif func_num == 3:    # COUNTA
        valid = ~np.isnan(col_data)
        result = np.sum(valid, axis=0).astype(float)
    elif func_num == 4:    # MAX
        result = np.nanmax(col_data, axis=0)
    elif func_num == 5:    # MIN
        result = np.nanmin(col_data, axis=0)
    elif func_num == 6:    # PRODUCT
        result = np.nanprod(col_data, axis=0)
    elif func_num == 7:    # STDEV
        result = np.nanstd(col_data, axis=0, ddof=1)
        if col_data.shape[0] <= 1:
            result = np.zeros(n)
    elif func_num == 8:    # STDEVP
        result = np.nanstd(col_data, axis=0, ddof=0)
        if col_data.shape[0] == 0:
            result = np.zeros(n)
    elif func_num == 9:    # SUM
        result = np.nansum(col_data, axis=0)
    elif func_num == 10:   # VAR
        result = np.nanvar(col_data, axis=0, ddof=1)
        if col_data.shape[0] <= 1:
            result = np.zeros(n)
    elif func_num == 11:   # VARP
        result = np.nanvar(col_data, axis=0, ddof=0)
        if col_data.shape[0] == 0:
            result = np.zeros(n)
    else:
        result = np.full(n, ERROR_MARKER, dtype=object)

    # 确保 result 是数值数组，否则返回错误标记
    if result.dtype.kind not in 'fc':
        try:
            result = result.astype(float)
        except:
            return np.full(n, ERROR_MARKER, dtype=object)

    # 如果结果全为 NaN，返回错误标记
    if np.all(np.isnan(result)):
        return np.full(n, ERROR_MARKER, dtype=object)
    return result

def _n2col(n: int) -> str:
    """将列号转换为字母"""
    s = ''
    while n > 0:
        n -= 1
        s = chr(n % 26 + 65) + s
        n //= 26
    return s

# ==================== SUMPRODUCT 核心函数 ====================
def _sumproduct_core(*arrays) -> np.ndarray:
    """
    向量化 SUMPRODUCT 实现，支持错误标记传播。
    """
    # 0. 确定迭代次数 n
    n = None
    for arr in arrays:
        arr = np.asarray(arr)
        if arr.ndim == 0:
            continue
        elif arr.ndim == 1:
            n = len(arr)
            break
        elif arr.ndim == 2:
            n = arr.shape[1]
            break
    if n is None:
        n = 1

    # 1. 预处理每个参数：检查错误标记，并转换为浮点数组（形状 (m_i, n)）
    processed = []
    for arr in arrays:
        arr = np.asarray(arr)
        # 错误标记检测：如果是对象数组且包含 ERROR_MARKER
        if arr.dtype == object:
            if np.any(arr == ERROR_MARKER):
                return np.full(n, ERROR_MARKER, dtype=object)
        # 转换为浮点数组（对象数组中的非错误值转为 float）
        if arr.dtype == object:
            # 将对象数组转为 float，错误标记已经提前返回
            try:
                arr = arr.astype(float)
            except:
                # 若转换失败，说明仍有非数字项，视为错误
                return np.full(n, ERROR_MARKER, dtype=object)
        elif arr.dtype.kind in 'fc':
            arr = arr.astype(float)
        else:
            # 其他类型（如 bool）转为 float
            arr = arr.astype(float)

        # 统一形状为 (m_i, n)
        if arr.ndim == 0:
            processed.append(np.full((1, n), arr, dtype=float))
        elif arr.ndim == 1:
            if len(arr) == 1:
                processed.append(np.full((1, n), arr[0], dtype=float))
            elif len(arr) == n:
                processed.append(arr.reshape(1, -1))
            else:
                # 长度不匹配，视为错误
                return np.full(n, ERROR_MARKER, dtype=object)
        elif arr.ndim == 2:
            if arr.shape[1] == 1:
                processed.append(np.repeat(arr, n, axis=1))
            elif arr.shape[1] == n:
                processed.append(arr)
            else:
                return np.full(n, ERROR_MARKER, dtype=object)
        else:
            return np.full(n, ERROR_MARKER, dtype=object)

    # 2. 行数一致性检查并广播
    rows = [arr.shape[0] for arr in processed]
    max_rows = max(rows)
    # 如果存在两个以上行数大于 1 且不相等，则返回错误
    if sum(1 for r in rows if r > 1) > 1 and len(set(rows)) > 1:
        return np.full(n, ERROR_MARKER, dtype=object)

    broadcasted = []
    for arr in processed:
        if arr.shape[0] == 1:
            arr = np.repeat(arr, max_rows, axis=0)
        broadcasted.append(arr)

    # 3. 逐元素相乘并求和
    product = np.ones((max_rows, n), dtype=float)
    for arr in broadcasted:
        product = product * arr
    result = np.sum(product, axis=0)

    # 4. 若结果全为 NaN，返回错误标记（但已有错误检测，此处作为兜底）
    if np.all(np.isnan(result)):
        return np.full(n, ERROR_MARKER, dtype=object)
    return result

# ==================== Excel 内置函数命名空间（扩充版） ====================
def _xl_func_ns() -> dict:
    """返回 Excel 内置函数的向量化实现（涵盖常用数学、逻辑、统计、日期、查找、财务函数）"""
    import numpy as _np
    ns = {}

    # 数学函数
    ns['ABS']       = lambda x: _np.abs(x)
    ns['SQRT']      = lambda x: _np.sqrt(_np.maximum(x, 0))
    ns['EXP']       = lambda x: _np.exp(x)
    ns['LN']        = lambda x: _np.log(_np.maximum(x, 1e-300))
    ns['LOG']       = lambda x, b=None: (_np.log10(_np.maximum(x, 1e-300)) if b is None
                                          else _np.log(_np.maximum(x, 1e-300)) /
                                               _np.log(_np.maximum(b, 1e-300)))
    ns['LOG10']     = lambda x: _np.log10(_np.maximum(x, 1e-300))
    ns['POWER']     = lambda x, p: _np.power(x, p)
    ns['MOD']       = lambda x, d: _np.mod(x, d)
    ns['SIGN']      = lambda x: _np.sign(x)
    ns['INT']       = lambda x: _np.floor(x).astype(float)
    ns['TRUNC']     = lambda x, d=None: (_np.trunc(x) if d is None
                                          else _np.trunc(x * 10**d) / 10**d)
    ns['ROUND']     = lambda x, d=0: _np.round(x, int(d) if _np.ndim(d) == 0 else int(_np.mean(d)))
    ns['ROUNDUP']   = lambda x, d=0: _np.ceil(_np.abs(x) * 10**d) / 10**d * _np.sign(x)
    ns['ROUNDDOWN'] = lambda x, d=0: _np.floor(_np.abs(x) * 10**d) / 10**d * _np.sign(x)
    ns['CEILING']   = lambda x, s=1: _np.ceil(x / s) * s
    ns['FLOOR']     = lambda x, s=1: _np.floor(x / s) * s
    ns['MROUND']    = lambda x, m: _np.round(x / m) * m
    ns['QUOTIENT']  = lambda x, d: _np.floor(x / d)
    ns['PRODUCT']   = lambda *args: _np.prod(args, axis=0) if args else 1.0
    ns['SUMSQ']     = lambda *args: _np.sum(_np.square(args), axis=0) if args else 0.0
    ns['SUM']       = lambda *args: _np.sum(args, axis=0) if args else 0.0
    ns['AVERAGE']   = lambda *args: _np.mean(args, axis=0) if args else 0.0

    # 三角函数
    ns['SIN']       = lambda x: _np.sin(x)
    ns['COS']       = lambda x: _np.cos(x)
    ns['TAN']       = lambda x: _np.tan(x)
    ns['ASIN']      = lambda x: _np.arcsin(_np.clip(x, -1, 1))
    ns['ACOS']      = lambda x: _np.arccos(_np.clip(x, -1, 1))
    ns['ATAN']      = lambda x: _np.arctan(x)
    ns['ATAN2']     = lambda y, x: _np.arctan2(y, x)
    ns['DEGREES']   = lambda x: _np.degrees(x)
    ns['RADIANS']   = lambda x: _np.radians(x)

    # 常量
    ns['PI']        = lambda: _np.float64(math.pi)
    ns['E']         = math.e

    # 逻辑函数 (返回值转换为 1.0 / 0.0)
    ns['TRUE']      = lambda: 1.0
    ns['FALSE']     = lambda: 0.0
    ns['AND']       = lambda *args: _np.where(_np.logical_and.reduce(args), 1.0, 0.0) if args else 1.0
    ns['OR']        = lambda *args: _np.where(_np.logical_or.reduce(args), 1.0, 0.0) if args else 0.0
    ns['NOT']       = lambda x: _np.where(_np.logical_not(x), 1.0, 0.0)
    ns['XOR']       = lambda x, y: _np.where(_np.logical_xor(x, y), 1.0, 0.0)
    ns['IF']        = lambda cond, t, f=0.0: _np.where(cond, t, f)
    ns['IFERROR']   = lambda x, v=0.0: _np.where(_np.isnan(x) | _np.isinf(x), v, x)
    ns['IFNA']      = lambda x, v=0.0: _np.where(_np.isnan(x), v, x)
    ns['NA']        = lambda: _np.nan

    # 统计函数（部分）
    ns['AVEDEV']    = lambda *args: _np.mean(_np.abs(args - _np.mean(args, axis=0)), axis=0) if args else 0.0
    ns['DEVSQ']     = lambda *args: _np.sum((args - _np.mean(args, axis=0))**2, axis=0) if args else 0.0
    ns['KURT']      = lambda *args: (_np.mean((args - _np.mean(args, axis=0))**4, axis=0) /
                                     (_np.var(args, axis=0, ddof=0)**2) - 3) if args and _np.any(_np.var(args, axis=0) > 0) else 0.0
    ns['SKEW']      = lambda *args: (_np.mean((args - _np.mean(args, axis=0))**3, axis=0) /
                                     (_np.std(args, axis=0, ddof=0)**3)) if args and _np.any(_np.std(args, axis=0) > 0) else 0.0
    ns['STEYX']     = lambda known_y, known_x: np.sqrt((np.sum((known_y - np.polyval(np.polyfit(known_x.flatten(), known_y.flatten(), 1), known_x))**2) / (len(known_y)-2)) if len(known_y) > 2 else 0.0)

    ns['SUBTOTAL'] = _subtotal_core
    ns['SUMPRODUCT'] = _sumproduct_core


    # ---------- 日期时间函数 ----------
    def DATE(year, month, day):
        """返回指定日期的序列号（向量化）"""
        year = np.asarray(year)
        month = np.asarray(month)
        day = np.asarray(day)
        def _date_one(y, m, d):
            try:
                y_int = int(y)
                m_int = int(m)
                d_int = int(d)
                dt = datetime.date(y_int, m_int, d_int)
                return _date_to_excel_date(dt)
            except:
                return np.nan
        vfunc = np.vectorize(_date_one, otypes=[float])
        return vfunc(year, month, day)

    def DAY(date):
        """提取日期的日（向量化）"""
        def _day_one(d):
            try:
                dt = _excel_date_to_date(float(d))
                return float(dt.day)
            except:
                return np.nan
        vfunc = np.vectorize(_day_one, otypes=[float])
        return vfunc(date)

    def MONTH(date):
        """提取日期的月份（向量化）"""
        def _month_one(d):
            try:
                dt = _excel_date_to_date(float(d))
                return float(dt.month)
            except:
                return np.nan
        vfunc = np.vectorize(_month_one, otypes=[float])
        return vfunc(date)

    def YEAR(date):
        """提取日期的年份（向量化）"""
        def _year_one(d):
            try:
                dt = _excel_date_to_date(float(d))
                return float(dt.year)
            except:
                return np.nan
        vfunc = np.vectorize(_year_one, otypes=[float])
        return vfunc(date)

    def HOUR(datetime_serial):
        """提取时间的小时（向量化）"""
        def _hour_one(d):
            try:
                days = float(d)
                frac = days - int(days)
                seconds = frac * 86400
                hours = seconds // 3600
                return float(hours)
            except:
                return np.nan
        vfunc = np.vectorize(_hour_one, otypes=[float])
        return vfunc(datetime_serial)

    def MINUTE(datetime_serial):
        """提取时间的分钟（向量化）"""
        def _minute_one(d):
            try:
                days = float(d)
                frac = days - int(days)
                seconds = frac * 86400
                hours = seconds // 3600
                minutes = (seconds - hours*3600) // 60
                return float(minutes)
            except:
                return np.nan
        vfunc = np.vectorize(_minute_one, otypes=[float])
        return vfunc(datetime_serial)

    def SECOND(datetime_serial):
        """提取时间的秒（向量化）"""
        def _second_one(d):
            try:
                days = float(d)
                frac = days - int(days)
                seconds = frac * 86400
                hours = seconds // 3600
                minutes = (seconds - hours*3600) // 60
                sec = seconds - hours*3600 - minutes*60
                return float(sec)
            except:
                return np.nan
        vfunc = np.vectorize(_second_one, otypes=[float])
        return vfunc(datetime_serial)

    def TIME(hour, minute, second):
        """返回指定时间的一天中的小数部分（向量化）"""
        hour = np.asarray(hour)
        minute = np.asarray(minute)
        second = np.asarray(second)
        def _time_one(h, m, s):
            try:
                h_int = int(h)
                m_int = int(m)
                s_int = int(s)
                total_seconds = h_int*3600 + m_int*60 + s_int
                return total_seconds / 86400.0
            except:
                return np.nan
        vfunc = np.vectorize(_time_one, otypes=[float])
        return vfunc(hour, minute, second)

    def TODAY():
        """返回当前日期（序列号），所有迭代相同"""
        today = datetime.date.today()
        return np.full(1, _date_to_excel_date(today), dtype=float)[0]

    def NOW():
        """返回当前日期时间（序列号），所有迭代相同"""
        now = datetime.datetime.now()
        return np.full(1, _excel_datetime_to_float(now), dtype=float)[0]

    def EOMONTH(start_date, months):
        """返回起始日期之前/之后几个月的最后一天的序列号（向量化）"""
        start_date = np.asarray(start_date)
        months = np.asarray(months)
        def _eomonth_one(s, m):
            try:
                dt = _excel_date_to_date(float(s))
                year = dt.year
                month = dt.month
                m_int = int(m)
                new_year = year + (month + m_int - 1) // 12
                new_month = (month + m_int - 1) % 12 + 1
                # 计算该月的最后一天
                import calendar
                last_day = calendar.monthrange(new_year, new_month)[1]
                new_date = datetime.date(new_year, new_month, last_day)
                return _date_to_excel_date(new_date)
            except:
                return np.nan
        vfunc = np.vectorize(_eomonth_one, otypes=[float])
        return vfunc(start_date, months)

    def WEEKDAY(date, return_type=1):
        """返回星期几（向量化）"""
        date = np.asarray(date)
        return_type = np.asarray(return_type)
        def _weekday_one(d, rt):
            try:
                dt = _excel_date_to_date(float(d))
                # 0=星期一,6=星期日（Python）
                py_weekday = dt.weekday()  # 0=星期一
                if rt == 1:  # 返回1（星期日）到7（星期六）
                    return float(py_weekday + 2 if py_weekday < 6 else 1)
                elif rt == 2:  # 返回1（星期一）到7（星期日）
                    return float(py_weekday + 1)
                elif rt == 3:  # 返回0（星期一）到6（星期日）
                    return float(py_weekday)
                else:
                    return np.nan
            except:
                return np.nan
        vfunc = np.vectorize(_weekday_one, otypes=[float])
        return vfunc(date, return_type)

    def WEEKNUM(date, return_type=1):
        """返回一年中的第几周（向量化，基于ISO星期）"""
        date = np.asarray(date)
        return_type = np.asarray(return_type)
        def _weeknum_one(d, rt):
            try:
                dt = _excel_date_to_date(float(d))
                # 使用ISO周
                isocal = dt.isocalendar()
                return float(isocal[1])  # 周数
            except:
                return np.nan
        vfunc = np.vectorize(_weeknum_one, otypes=[float])
        return vfunc(date, return_type)

    ns['DATE'] = DATE
    ns['DAY'] = DAY
    ns['MONTH'] = MONTH
    ns['YEAR'] = YEAR
    ns['HOUR'] = HOUR
    ns['MINUTE'] = MINUTE
    ns['SECOND'] = SECOND
    ns['TIME'] = TIME
    ns['TODAY'] = TODAY
    ns['NOW'] = NOW
    ns['EOMONTH'] = EOMONTH
    ns['WEEKDAY'] = WEEKDAY
    ns['WEEKNUM'] = WEEKNUM

    # ---------- 查找函数（简化版）----------
    def INDEX(array, row_num, column_num=None):
        """
        INDEX 函数（向量化简化版）
        array: 可以是区域引用（字符串）或数组
        row_num, column_num: 索引（1-based），可以是向量
        返回对应位置的元素数组。
        """
        # 如果 array 是字符串且看起来像区域引用，则尝试展开
        if isinstance(array, str) and re.match(r'^[A-Z]+\d+(:[A-Z]+\d+)?$', array, re.I):
            # 转换为全单元格列表
            if ':' in array:
                coords = _range_to_coords(array)
            else:
                coords = [array]
            # 假设当前工作表
            sheet = "Sheet1"  # 实际使用时会在上层传递 default_sheet
            # 我们暂时无法在此处获取 default_sheet，所以交给调用者处理？
            # 为了简化，我们在这里假设 array 已经是展开后的数组（在调用前已经解析）
            # 实际上，在求值引擎中，array 参数已经被 eval 过，可能已经是数组了
            # 因此我们期望 array 已经是数组
            pass
        # 将 array 视为数组
        arr = np.asarray(array)
        row_num = np.asarray(row_num)
        if column_num is not None:
            col_num = np.asarray(column_num)
            # 确保维度匹配
            # 如果 arr 是二维，则使用行列索引；否则按一维处理
            if arr.ndim == 2:
                # 将行列索引转换为 0-based
                r = (row_num - 1).astype(int)
                c = (col_num - 1).astype(int)
                # 处理标量或向量
                if r.ndim == 0 and c.ndim == 0:
                    return arr[r, c]
                else:
                    # 向量化索引需要小心，使用 np.take 或迭代
                    # 简单实现：将 r 和 c 广播为相同形状
                    r, c = np.broadcast_arrays(r, c)
                    result = np.empty(r.shape, dtype=arr.dtype)
                    for idx in np.ndindex(r.shape):
                        result[idx] = arr[r[idx], c[idx]]
                    return result
            else:
                # 一维数组，忽略列号（或视 column_num 为无效）
                idx = (row_num - 1).astype(int)
                if idx.ndim == 0:
                    return arr[idx]
                else:
                    return arr[idx]
        else:
            # 只有行号，按一维处理
            idx = (row_num - 1).astype(int)
            if idx.ndim == 0:
                return arr[idx]
            else:
                return arr[idx]

    def MATCH(lookup_value, lookup_array, match_type=0):
        """
        MATCH 函数（向量化简化版，仅支持精确匹配 match_type=0）
        lookup_value: 要查找的值（标量或向量）
        lookup_array: 查找区域（必须是一维数组）
        match_type: 必须为 0（精确匹配），否则返回 #N/A
        返回 lookup_value 在 lookup_array 中的位置（1-based），如果未找到返回 #N/A（np.nan）
        """
        match_type = np.asarray(match_type)
        # 如果 match_type 不是 0，返回 nan
        if not np.all(match_type == 0):
            return np.full_like(lookup_value, np.nan, dtype=float)
        lookup_arr = np.asarray(lookup_array).flatten()
        lookup_val = np.asarray(lookup_value)
        if lookup_val.ndim == 0:
            # 标量
            indices = np.where(lookup_arr == lookup_val)[0]
            if len(indices) > 0:
                return float(indices[0] + 1)  # 1-based
            else:
                return np.nan
        else:
            # 向量
            result = np.full_like(lookup_val, np.nan, dtype=float)
            for i, val in np.ndenumerate(lookup_val):
                idx = np.where(lookup_arr == val)[0]
                if len(idx) > 0:
                    result[i] = idx[0] + 1
            return result

    ns['INDEX'] = INDEX
    ns['MATCH'] = MATCH

    # ---------- 财务函数 NPV 和 PV ----------
    def NPV(rate, *values):
        """
        净现值函数（向量化）
        rate: 贴现率（标量或数组，长度 n）
        values: 现金流序列，可以是：
                - 多个参数，每个参数对应一期的现金流（可以是标量或数组）
                - 一个参数，且该参数是一个二维数组，形状为 (t, n)，表示 t 期现金流
        返回长度为 n 的数组，表示每个迭代的净现值
        """
        # 将 rate 转换为数组
        rate_arr = np.asarray(rate)
        if rate_arr.ndim == 0:
            # 如果 rate 是标量，广播到与现金流相同的长度（取第一个现金流数组的长度）
            if values:
                # 尝试获取第一个现金流数组的长度
                first_val = values[0]
                if isinstance(first_val, np.ndarray):
                    n = len(first_val)
                else:
                    n = 1
                rate_arr = np.full(n, rate_arr)
            else:
                rate_arr = np.array([rate_arr])
        n = len(rate_arr)

        # 处理现金流
        cash_flows = []
        # 特殊情况：只有一个参数，且该参数是二维数组
        if len(values) == 1 and isinstance(values[0], np.ndarray) and values[0].ndim == 2:
            # 二维数组，形状 (t, n)，每行是一期的现金流
            cf_array = values[0]
            t = cf_array.shape[0]
            # 确保 t 期现金流
            if cf_array.shape[1] != n:
                raise ValueError(f"NPV 现金流数组列数 {cf_array.shape[1]} 与贴现率长度 {n} 不匹配")
            # 直接使用 cf_array，形状 (t, n)
            cf_stack = cf_array
        else:
            # 多个参数，每个参数对应一期
            for v in values:
                v_arr = np.asarray(v)
                if v_arr.ndim == 0:
                    v_arr = np.full(n, v_arr)
                elif len(v_arr) != n:
                    # 长度不匹配，尝试广播
                    if len(v_arr) == 1:
                        v_arr = np.full(n, v_arr[0])
                    else:
                        raise ValueError(f"NPV 参数长度 {len(v_arr)} 与贴现率长度 {n} 不匹配")
                cash_flows.append(v_arr)
            if not cash_flows:
                return np.zeros(n, dtype=float)
            cf_stack = np.vstack(cash_flows)  # (t, n)
        t = cf_stack.shape[0]

        # 计算贴现因子：1 / (1+rate)^t，t 从 1 开始
        rate_expanded = rate_arr.reshape(1, -1)
        t_vals = np.arange(1, t+1).reshape(-1, 1)
        denominator = (1 + rate_expanded) ** t_vals
        discount_factors = 1.0 / denominator
        npv = np.sum(cf_stack * discount_factors, axis=0)
        return npv

    def PV(rate, nper, pmt, fv=0.0, type_=0):
        """
        现值函数（向量化，支持二维数组）
        
        参数：
            rate: 贴现率（标量、一维数组或二维数组）
            nper: 期数
            pmt: 每期付款
            fv: 未来值，默认 0
            type_: 付款时间，0 期末，1 期初，默认 0
        
        返回：
            现值数组，形状与输入中最复杂的数组相同
        """
        # 将输入转换为 numpy 数组
        rate_arr = np.asarray(rate)
        nper_arr = np.asarray(nper)
        pmt_arr = np.asarray(pmt)
        fv_arr = np.asarray(fv)
        type_arr = np.asarray(type_)
        
        # 确定迭代次数 n = 所有参数最后一维的最大长度
        def last_dim_len(arr):
            return arr.shape[-1] if arr.ndim >= 1 else 1
        n = max(last_dim_len(rate_arr), last_dim_len(nper_arr),
                last_dim_len(pmt_arr), last_dim_len(fv_arr), last_dim_len(type_arr))
        
        # 广播所有参数到形状 (..., n)
        # 对于标量或一维数组，直接扩展到最后维度；对于二维数组，保持第一维（单元格数）
        def broadcast_to_n(arr):
            if arr.ndim == 0:
                return np.full(n, arr)
            elif arr.ndim == 1:
                if len(arr) == 1:
                    return np.full(n, arr[0])
                elif len(arr) == n:
                    return arr
                else:
                    raise ValueError(f"数组长度 {len(arr)} 与迭代次数 {n} 不匹配")
            elif arr.ndim == 2:
                # 二维数组，形状 (m, n) 或 (m, 1) -> 扩展第二维到 n
                if arr.shape[1] == 1:
                    return np.repeat(arr, n, axis=1)
                elif arr.shape[1] == n:
                    return arr
                else:
                    raise ValueError(f"二维数组第二维长度 {arr.shape[1]} 与迭代次数 {n} 不匹配")
            else:
                raise ValueError(f"不支持 {arr.ndim} 维数组")
        
        rate_b = broadcast_to_n(rate_arr)
        nper_b = broadcast_to_n(nper_arr)
        pmt_b = broadcast_to_n(pmt_arr)
        fv_b = broadcast_to_n(fv_arr)
        type_b = broadcast_to_n(type_arr)
        
        # 确保所有数组形状一致
        # 对于二维数组，我们需将一维数组也扩展为 (1, n) 以兼容二维广播
        if any(arr.ndim == 2 for arr in (rate_b, nper_b, pmt_b, fv_b, type_b)):
            # 找到二维数组的最大第一维
            m = max((arr.shape[0] for arr in (rate_b, nper_b, pmt_b, fv_b, type_b) if arr.ndim == 2), default=1)
            def to_2d(arr):
                if arr.ndim == 0:
                    return np.full((m, n), arr)
                elif arr.ndim == 1:
                    if len(arr) == 1:
                        return np.full((m, n), arr[0])
                    else:
                        return np.tile(arr.reshape(1, -1), (m, 1))
                else:  # ndim == 2
                    if arr.shape[0] == 1:
                        return np.tile(arr, (m, 1))
                    else:
                        return arr
            rate_2d = to_2d(rate_b)
            nper_2d = to_2d(nper_b)
            pmt_2d = to_2d(pmt_b)
            fv_2d = to_2d(fv_b)
            type_2d = to_2d(type_b)
            
            # 统一形状为 (m, n)
            # 对每个单元格（行）独立计算现值
            result = np.zeros_like(rate_2d)
            for i in range(m):
                r = rate_2d[i]
                nper_i = nper_2d[i]
                pmt_i = pmt_2d[i]
                fv_i = fv_2d[i]
                t = type_2d[i]
                # 对于当前行的向量计算
                mask_nonzero = np.abs(r) > 1e-12
                res = np.zeros(n)
                if np.any(mask_nonzero):
                    rnz = r[mask_nonzero]
                    nper_nz = nper_i[mask_nonzero]
                    pmt_nz = pmt_i[mask_nonzero]
                    fv_nz = fv_i[mask_nonzero]
                    t_nz = t[mask_nonzero]
                    factor = (1 + rnz) ** nper_nz
                    pv_nz = -pmt_nz * (1 - 1/factor) / rnz - fv_nz / factor
                    if np.any(t_nz == 1):
                        pv_nz[t_nz == 1] *= (1 + rnz[t_nz == 1])
                    res[mask_nonzero] = pv_nz
                mask_zero = ~mask_nonzero
                if np.any(mask_zero):
                    r_zero = r[mask_zero]
                    nper_zero = nper_i[mask_zero]
                    pmt_zero = pmt_i[mask_zero]
                    fv_zero = fv_i[mask_zero]
                    # 公式：-(pmt * nper + fv)
                    res[mask_zero] = -(pmt_zero * nper_zero + fv_zero)
                result[i] = res
            return result
        else:
            # 所有参数都是一维或标量，按原逻辑计算，返回一维数组
            # 原逻辑略作调整，但核心不变
            mask_nonzero = np.abs(rate_b) > 1e-12
            result = np.zeros(n)
            if np.any(mask_nonzero):
                r = rate_b[mask_nonzero]
                nper_i = nper_b[mask_nonzero]
                pmt_i = pmt_b[mask_nonzero]
                fv_i = fv_b[mask_nonzero]
                t = type_b[mask_nonzero]
                factor = (1 + r) ** nper_i
                pv_nz = -pmt_i * (1 - 1/factor) / r - fv_i / factor
                if np.any(t == 1):
                    pv_nz[t == 1] *= (1 + r[t == 1])
                result[mask_nonzero] = pv_nz
            mask_zero = ~mask_nonzero
            if np.any(mask_zero):
                r = rate_b[mask_zero]
                nper_i = nper_b[mask_zero]
                pmt_i = pmt_b[mask_zero]
                fv_i = fv_b[mask_zero]
                result[mask_zero] = -(pmt_i * nper_i + fv_i)
            return result
    
    ns['NPV'] = NPV
    ns['PV'] = PV

    return ns

# 全局缓存
_XL_NS_CACHE = None
def _get_xl_ns():
    """获取 Excel 内置函数命名空间（单例）"""
    global _XL_NS_CACHE
    if _XL_NS_CACHE is None:
        _XL_NS_CACHE = _xl_func_ns()
    return _XL_NS_CACHE