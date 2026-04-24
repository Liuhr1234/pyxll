# ui_modeler_param_validation_mixin.py
# -*- coding: utf-8 -*-
"""
本模块为分布构建器（DistributionBuilderDialog）提供参数解析与验证的辅助功能（Mixin）。

主要功能：
1. 输入校验器配置 (Validator Setup)：为 UI 输入框生成基于规则的 QValidator（如整数、概率、正数限制等）。
2. 参数值解析 (Value Resolution)：将用户输入的纯文本或 Excel 单元格引用（通过 COM 接口）解析为实际数值。
3. 数组型参数处理 (List Param Handling)：处理离散分布或自定义分布中成对的 X 数组和 P 数组输入。
4. 特殊分布约束校验 (Distribution Constraints)：检查特定分布的内在逻辑错误（如累积概率必须递增，概率值必须在0到1之间）。
5. 全局验证引擎与 UI 反馈 (Validation Engine & UI)：严格校验所有参数，并根据结果在 UI 上标记错误（如红框提示、禁用插入按钮等）。
"""

import drisk_env
import re
from typing import Any
import numpy as np

from com_fixer import _safe_excel_app
from PySide6.QtGui import QDoubleValidator, QIntValidator


class DistributionBuilderParamValidationMixin:
    """
    [混合类] 分布构建器对话框的参数解析与验证辅助工具。
    """

    # =======================================================
    # 1. 输入校验器配置
    # =======================================================
    def _make_validator_for(self, dist_key: str, param_key: str):
        """
        根据底层配置的参数规则，为特定的输入框生成对应的 Qt 校验器。
        确保用户在 UI 层面无法输入非法字符（如在只允许正数的地方输入字母）。
        """
        rules = getattr(self, "_PARAM_RULES", {}).get(dist_key, {})
        r = rules.get(param_key)
        
        # 如果没有规则，默认提供高精度的浮点数校验器
        if not r:
            v = QDoubleValidator()
            v.setNotation(QDoubleValidator.StandardNotation)
            v.setDecimals(12)
            return v

        t = r.get("type")
        # 整数类型校验
        if t == "int":
            return QIntValidator(int(r.get("min", -10 ** 9)), int(r.get("max", 10 ** 9)))

        # 概率类型校验（0.0 到 1.0）
        if t == "prob":
            v = QDoubleValidator(float(r.get("min", 0.0)), float(r.get("max", 1.0)), 12)
            v.setNotation(QDoubleValidator.StandardNotation)
            return v

        # 正数类型校验
        if t == "pos":
            bottom = float(r.get("min", 0.0))
            top = float(r.get("max", 1e18)) if "max" in r else 1e18
            v = QDoubleValidator(bottom, top, 12)
            v.setNotation(QDoubleValidator.StandardNotation)
            return v

        # 兜底的高精度浮点数校验器
        v = QDoubleValidator()
        v.setNotation(QDoubleValidator.StandardNotation)
        v.setDecimals(12)
        return v


    # =======================================================
    # 2. 参数值与 Excel 引用解析
    # =======================================================
    def _resolve_param_value(self, raw_str: str) -> Any:
        """
        核心数值解析器：将原始字符串转换为数值。
        支持：纯数字文本、逗号分隔的数组字符串、大括号包围的数组、以及 Excel 单元格/区域引用。
        """
        raw_str = (raw_str or "").strip()
        if not raw_str:
            raise ValueError("空")

        # 尝试直接转为浮点数，成功则快速返回
        try:
            return float(raw_str)
        except ValueError:
            pass

        # 若纯文本转换失败，则尝试通过 COM 接口连接 Excel 读取数据
        xl_excel = _safe_excel_app()

        def _flatten_numeric_values(value_obj):
            """辅助函数：递归展平从 Excel 读取的多维数据结构（如二维区域）"""
            numbers = []

            def _walk(node):
                if isinstance(node, (tuple, list, np.ndarray)):
                    for sub in node:
                        _walk(sub)
                else:
                    if node is None or isinstance(node, bool):
                        return
                    if isinstance(node, (int, float)):
                        numbers.append(float(node))
                        return
                    # Accept numeric text from worksheet cells, including percentage text.
                    if isinstance(node, str):
                        txt = node.strip().replace(",", "")
                        if not txt:
                            return
                        try:
                            if txt.endswith("%"):
                                numbers.append(float(txt[:-1]) / 100.0)
                            else:
                                numbers.append(float(txt))
                        except Exception:
                            return

            _walk(value_obj)
            return numbers

        def _resolve_single_reference(token: str):
            """辅助函数：解析单个 Excel 单元格或区域引用（如 'Sheet1'!A1:B2）"""
            token = token.strip()
            if not token:
                return []

            # 处理跨工作表引用
            if "!" in token:
                sheet_name, cell_addr_raw = token.split("!", 1)
                sheet_name = sheet_name.strip().replace("'", "").replace('"', "")
                sheet = xl_excel.ActiveWorkbook.Worksheets(sheet_name)
            else:
                sheet = xl_excel.ActiveSheet
                cell_addr_raw = token

            # 正则提取干净的单元格地址
            match = re.search(r"(\$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?)", cell_addr_raw)
            if not match:
                raise ValueError(f"无效的单元格引用: {token}")

            clean_addr = match.group(1).upper()
            val = sheet.Range(clean_addr).Value
            if val is None:
                raise ValueError("单元格为空")

            numbers = _flatten_numeric_values(val)
            if not numbers:
                raise ValueError("没有数值")
            return numbers

        # 预处理：去除大括号
        cleaned = raw_str
        if cleaned.startswith("{") and cleaned.endswith("}"):
            cleaned = cleaned[1:-1].strip()
        tokens = [tok.strip() for tok in cleaned.split(",") if tok.strip()]

        try:
            # 如果存在多个片段（如 A1, A2, 1.5）
            if len(tokens) > 1:
                numeric_only = []
                all_numeric = True
                for token in tokens:
                    try:
                        numeric_only.append(float(token))
                    except ValueError:
                        all_numeric = False
                        break
                
                # 如果全是纯数字，直接拼接返回
                if all_numeric:
                    return ",".join(str(v) for v in numeric_only)

                # 否则混合解析（包含引用的情况）
                flat_numbers = []
                for token in tokens:
                    flat_numbers.extend(_resolve_single_reference(token))
                if len(flat_numbers) == 1:
                    return float(flat_numbers[0])
                return ",".join(str(v) for v in flat_numbers)

            # 单个复杂输入的情况（单引用或单区域）
            one_token = tokens[0] if tokens else raw_str
            numbers = _resolve_single_reference(one_token)
            if len(numbers) == 1:
                return float(numbers[0])
            return ",".join(str(v) for v in numbers)
        except Exception as e:
            print(f"解析参数失败: {raw_str}, 错误: {e}")
            raise ValueError(f"无效的单元格或数字: {raw_str}")

    def _resolve_string_value(self, raw_str: str) -> str:
        """
        解析字符串属性（如 Name，Category 等）。
        如果输入是单元格引用，则通过 COM 获取其实际文本值；如果带有引号，则直接剥离作为字面量。
        """
        raw_str = (raw_str or "").strip()
        if not raw_str:
            return ""
            
        # 1. 检查是否自带引号，如果是则直接剥离作为字面量
        if (raw_str.startswith('"') and raw_str.endswith('"')) or (raw_str.startswith("'") and raw_str.endswith("'")):
            return raw_str[1:-1]
            
        # 2. 尝试判断是否为单元格引用并解析
        try:
            addr = raw_str.split('!')[-1].replace('$', '')
            if re.fullmatch(r'[A-Za-z]{1,3}\d+', addr):
                xl_app = _safe_excel_app()
                
                if '!' in raw_str:
                    sheet_name, cell_addr_raw = raw_str.split('!', 1)
                    sheet_name = sheet_name.replace("'", "").replace('"', "")
                    sheet = xl_app.ActiveWorkbook.Worksheets(sheet_name)
                else:
                    sheet = xl_app.ActiveSheet
                    cell_addr_raw = raw_str
                
                clean_addr = cell_addr_raw.replace('$', '')
                match = re.search(r'([A-Za-z]{1,3}\d+)', clean_addr)
                if match:
                    cell = sheet.Range(match.group(1).upper())
                    val = cell.Value
                    if val is not None:
                        return str(val)
        except Exception as e:
            print(f"[Drisk] 解析字符串单元格失败: {raw_str}, 错误: {e}")
            
        # 3. 兜底：如果不是合法的单元格，或者解析失败，直接作为普通文本返回
        return raw_str

    def _format_excel_ref_for_input(self, app, excel_obj):
        """
        将通过 UI 交互获取的 Excel COM 区域对象，反向格式化为标准的字符串引用形式。
        例如转换回：'Sheet1'!A1
        """
        sheet = excel_obj.Worksheet.Name
        addr = excel_obj.Address.replace('$', '')
        active_sheet = app.ActiveSheet.Name
        # 如果是在当前激活的表，则省略工作表前缀
        if sheet == active_sheet:
            return addr
        return f"'{sheet}'!{addr}" if " " in sheet else f"{sheet}!{addr}"


    # =======================================================
    # 3. 数组/列表型参数处理
    # =======================================================
    def _get_list_param_pair_config(self):
        """识别并返回当前分布下对应的数组参数键名组合（X 数组与 P 数组）。"""
        dist_key = str(self.dist_type or "")
        if dist_key in ("Cumul", "DriskCumul"):
            return ("X-Table", "P-Table", "自定义分布")
        if dist_key in ("Discrete", "DriskDiscrete"):
            return ("X-Table", "P-Table", "自定义分布")
        if dist_key in ("General", "DriskGeneral"):
            return ("X-Table", "P-Table", "自定义分布")
        if dist_key in ("Histogrm", "DriskHistogrm"):
            return (None, "P-Table", "直方图分布")
        if dist_key in ("DUniform", "DriskDUniform"):
            return ("X-Table", None, "离散均匀分布")
        return None

    def _is_list_param_selector(self, param_key):
        """判断给定的参数键是否为数组型参数（例如用来决定是否显示多选输入工具）。"""
        cfg = self._get_list_param_pair_config()
        if not cfg:
            return False
        x_key, p_key, _name = cfg
        return param_key in tuple(k for k in (x_key, p_key) if k)


    def _to_numeric_list_for_validation(self, value):
        """递归展平混合输入，将任意合法输入（引用、数组、字符串）转化为用于逻辑校验的一维浮点数列表。"""
        if value is None:
            return []
        if isinstance(value, (int, float)) and not isinstance(value, bool):
            return [float(value)]
        if isinstance(value, (list, tuple, np.ndarray)):
            out = []
            for item in value:
                out.extend(self._to_numeric_list_for_validation(item))
            return out

        text = str(value).strip()
        if not text:
            return []
        if text.startswith("{") and text.endswith("}"):
            text = text[1:-1].strip()
        if not text:
            return []

        # 如果长得像单元格，尝试深层解析
        if re.search(r'[A-Za-z]{1,3}\d+', text):
            try:
                resolved = self._resolve_param_value(text)
                return self._to_numeric_list_for_validation(resolved)
            except Exception:
                pass

        # 普通字符串按逗号、分号或空格切分
        out = []
        for part in re.split(r'[,;\s]+', text):
            if not part:
                continue
            out.append(float(part))
        return out

    def _get_list_length_pair(self, value_map):
        """获取一对相关联数组的长度信息，用于后续的比对。"""
        cfg = self._get_list_param_pair_config()
        if not cfg:
            return None
        x_key, p_key, dist_name = cfg
        if x_key not in value_map or p_key not in value_map:
            return None

        x_list = self._to_numeric_list_for_validation(value_map.get(x_key))
        p_list = self._to_numeric_list_for_validation(value_map.get(p_key))
        if not x_list or not p_list:
            return None
        return x_key, p_key, dist_name, len(x_list), len(p_list)

    def _show_list_length_mismatch_warning_if_needed(self):
        """实时检查 X 和 P 数组的长度，如果不一致则触发 UI 警告。"""
        cfg = self._get_list_param_pair_config()
        if not cfg:
            return
        x_key, p_key, dist_name = cfg
        x_le = self.inputs.get(x_key) if hasattr(self, "inputs") else None
        p_le = self.inputs.get(p_key) if hasattr(self, "inputs") else None
        if not x_le or not p_le:
            return

        try:
            pair = self._get_list_length_pair({
                x_key: (x_le.text() or "").strip(),
                p_key: (p_le.text() or "").strip(),
            })
        except Exception:
            return

        if pair and pair[3] != pair[4]:
            msg = f"{dist_name} 的 X 数组与 P 数组长度不一致（{pair[3]} vs {pair[4]}）"
            self._show_param_validation_error_in_canvas(msg)


    # =======================================================
    # 4. 特定分布约束校验 (如 Cumul 累积分布)
    # =======================================================
    def _validate_cumul_constraints(self, value_map):
        """
        专门针对累积分布（Cumul）的严格数学规则校验：
        - 最小值必须小于最大值
        - X 数组必须严格递增，且在其 [Min, Max] 范围内
        - P 数组（概率）必须严格递增，且在 (0, 1) 之间
        """
        if str(self.dist_type or "") not in ("Cumul", "DriskCumul"):
            return {}

        err = {}
        min_val = None
        max_val = None
        x_list = None
        p_list = None

        try:
            min_val = float(value_map.get("Min"))
        except Exception:
            err["Min"] = "最小值无效"

        try:
            max_val = float(value_map.get("Max"))
        except Exception:
            err["Max"] = "最大值无效"

        try:
            x_list = self._to_numeric_list_for_validation(value_map.get("X-Table"))
            if not x_list:
                err["X-Table"] = "X 数组不能为空"
        except Exception:
            err["X-Table"] = "X 数组解析失败"

        try:
            p_list = self._to_numeric_list_for_validation(value_map.get("P-Table"))
            if not p_list:
                err["P-Table"] = "P 数组不能为空"
        except Exception:
            err["P-Table"] = "P 数组解析失败"

        # 规则 1：Min 必须 < Max
        if min_val is not None and max_val is not None and min_val >= max_val:
            msg = "最小值必须小于最大值"
            err["Min"] = msg
            err["Max"] = msg

        # 规则 2：X 数组检验
        if x_list and len(x_list) > 1:
            if any(x_list[i] >= x_list[i + 1] for i in range(len(x_list) - 1)):
                err["X-Table"] = "X 数组必须严格递增"
        if x_list and min_val is not None and max_val is not None:
            if any((x <= min_val) or (x >= max_val) for x in x_list):
                err["X-Table"] = "元素必须在 Min 和 Max 之间"

        # 规则 3：P 数组检验
        if p_list and len(p_list) > 1:
            if any(p_list[i] >= p_list[i + 1] for i in range(len(p_list) - 1)):
                err["P-Table"] = "P 数组必须严格递增"
        if p_list:
            if any((p <= 0.0) or (p >= 1.0) for p in p_list):
                err["P-Table"] = "元素必须在 0 到 1 之间"

        return err

    def _maybe_show_cumul_validation_prompt(self, err):
        """管理特殊分布错误信息的 UI 弹窗提示，防止重复或频繁的无用报错干扰用户。"""
        dist_key = str(self.dist_type or "")
        if dist_key not in ("Cumul", "DriskCumul", "Discrete", "DriskDiscrete"):
            self._last_cumul_validation_msg = ""
            return

        # 保证报错显示的确定性顺序，优先展示最具操作性的提示。
        ordered_keys = ("Min", "Max", "X-Table", "P-Table", "x_vals", "p_vals")
        messages = []
        for key in ordered_keys:
            msg = (err or {}).get(key, "")
            if msg:
                messages.append(msg)
        if not messages and err:
            for _k, msg in (err or {}).items():
                if msg:
                    messages.append(msg)
                    break
        if not messages:
            self._last_cumul_validation_msg = ""
            return

        msg = messages[0]
        # 忽略基础的非空提示，专注业务逻辑错误
        if msg == "不能为空":
            return

        last_msg = getattr(self, "_last_cumul_validation_msg", "")
        if msg == last_msg:
            return

        self._last_cumul_validation_msg = msg
        self._show_param_validation_error_in_canvas(msg)


    # =======================================================
    # 5. 核心验证引擎与 UI 反馈
    # =======================================================
    def _show_param_validation_error_in_canvas(self, detail: str):
        """将最终的格式化错误信息推送到画板/骨架屏 (Skeleton) 进行显示。"""
        detail = (detail or "").strip()
        if not detail:
            detail = "输入参数无效"
        text = f"参数错误：{detail}"
        if hasattr(self, "_show_skeleton"):
            try:
                self._show_skeleton(text)
                return
            except Exception:
                pass

    def _validate_params_strict(self):
        """
        [主验证循环] 严格参数校验：
        遍历当前分布所需的所有参数，执行基础解析、上下限验证、类型（整数/概率）验证、
        以及特殊的交叉验证（如统一分布 min < max，数组长度比对），
        最终收集并汇总所有的错误信息。
        """
        vals, err = {}, {}
        if not getattr(self, "inputs", None) or not getattr(self, "config", None):
            return True, vals, err

        params_def = self.config.get("params", [])
        rules = getattr(self, "_PARAM_RULES", {}).get(self.dist_type, {})

        # 遍历检验每个输入字段
        for key, _label, _default in params_def:
            le = self.inputs.get(key)
            if le is None:
                continue
            raw = (le.text() or "").strip()
            if raw == "":
                err[key] = "不能为空"
                continue

            rule = rules.get(key, None)
            if rule and rule.get("type") == "formula":
                formula_raw = raw[1:].strip() if raw.startswith("=") else raw
                cell_ref_pattern = r"(?:'[^']+'!|[A-Za-z_][A-Za-z0-9_ ]*!)?\$?[A-Za-z]{1,3}\$?\d+"
                dist_formula_pattern = r"@?\s*Drisk[A-Za-z0-9_]*\s*\(.+\)\s*"

                if re.fullmatch(cell_ref_pattern, formula_raw) or re.fullmatch(dist_formula_pattern, formula_raw, re.IGNORECASE):
                    vals[key] = formula_raw
                    continue

                err[key] = "\u5fc5\u987b\u662f\u5206\u5e03\u516c\u5f0f\u6216\u5355\u5143\u683c\u5f15\u7528"
                continue

            if rule and rule.get("type") == "optional_nonneg":
                is_inf_text = raw.lower() in ("inf", "infinity") or raw == "\u221e"
                if raw == "" or (rule.get("allow_inf") and is_inf_text):
                    vals[key] = raw
                    continue

                try:
                    resolved = self._resolve_param_value(raw)
                    if isinstance(resolved, (list, tuple)):
                        resolved = resolved[0][0]
                    v = float(resolved)
                except Exception:
                    err[key] = "\u53c2\u6570\u89e3\u6790\u5931\u8d25"
                    continue

                if v < 0:
                    err[key] = "\u5fc5\u987b >= 0"
                    continue

                vals[key] = v
                continue

            if rule and rule.get("type") == "finite":
                try:
                    resolved = self._resolve_param_value(raw)
                    if isinstance(resolved, (list, tuple)):
                        resolved = resolved[0][0]
                    v = float(resolved)
                except Exception:
                    err[key] = "\u53c2\u6570\u89e3\u6790\u5931\u8d25"
                    continue

                if not np.isfinite(v):
                    err[key] = "\u5fc5\u987b\u662f\u6709\u9650\u6570\u503c"
                    continue

                vals[key] = v
                continue

            try:
                # 整数类型校验
                if rule and rule.get("type") == "int":
                    f = self._resolve_param_value(raw)
                    if isinstance(f, (list, tuple)):
                        raise ValueError("整数参数不能使用数组")
                    if not float(f).is_integer():
                        raise ValueError("不是整数")
                    v = int(f)
                else:
                    v = self._resolve_param_value(raw)
            except Exception:
                # Only bypass parse errors for real list-style params of list-driven dists.
                is_list_param = False
                try:
                    is_list_param = bool(self._is_list_param_selector(key))
                except Exception:
                    is_list_param = False

                if "{" in raw or "[" in raw or "," in raw or is_list_param:
                    v = raw
                else:
                    err[key] = "参数解析失败"
                    continue

            # 边界与规则校验
            if rule:
                try:
                    v_scalar = float(v) if not isinstance(v, (list, tuple)) else float(v[0][0])
                except Exception:
                    err[key] = "无法转换为标量"
                    continue

                if "min" in rule:
                    mn = float(rule["min"])
                    if rule.get("exclusive_min", False):
                        if v_scalar <= mn:
                            err[key] = f"必须 > {mn}"
                            continue
                    else:
                        if v_scalar < mn:
                            err[key] = f"必须 >= {mn}"
                            continue

                if "max" in rule:
                    mx = float(rule["max"])
                    if rule.get("exclusive_max", False):
                        if v_scalar >= mx:
                            err[key] = f"必须 < {mx}"
                            continue
                    else:
                        if v_scalar > mx:
                            err[key] = f"必须 <= {mx}"
                            continue

                if rule.get("type") == "prob":
                    if not (0.0 <= v_scalar <= 1.0):
                        err[key] = "必须在 [0,1] 之间"
                        continue

                v = v_scalar
            vals[key] = v

        # 交叉参数校验：如均匀分布的上下限
        if "__cross__" in rules and rules["__cross__"] == "uniform_min_lt_max":
            params_def = self.config.get("params", [])
            if len(params_def) >= 2:
                min_key = params_def[0][0]
                max_key = params_def[1][0]
                if min_key in vals and max_key in vals:
                    if vals[min_key] >= vals[max_key]:
                        err[max_key] = "必须大于最小值"
        if "__cross__" in rules and rules["__cross__"] == "hypergeo_bounds":
            params_def = self.config.get("params", [])
            if len(params_def) >= 3:
                n_key = params_def[0][0]
                d_key = params_def[1][0]
                m_key = params_def[2][0]
                if n_key in vals and m_key in vals and vals[n_key] > vals[m_key]:
                    err[n_key] = "\u5fc5\u987b\u5c0f\u4e8e\u7b49\u4e8e M"
                if d_key in vals and m_key in vals and vals[d_key] > vals[m_key]:
                    err[d_key] = "\u5fc5\u987b\u5c0f\u4e8e\u7b49\u4e8e M"



        # 检查数组长度一致性
        try:
            pair = self._get_list_length_pair(vals)
            if pair and pair[3] != pair[4]:
                msg = f"{pair[2]} 的 X 数组与 P 数组长度不一致（{pair[3]} vs {pair[4]}）"
                err[pair[0]] = msg
                err[pair[1]] = msg
        except Exception:
            pass

        # 合并特定分布（如Cumul）的错误
        cumul_err = self._validate_cumul_constraints(vals)
        for k, msg in cumul_err.items():
            if k not in err:
                err[k] = msg

        ok = (len(err) == 0)
        return ok, vals, err

    def _apply_param_validation_ui(self, ok: bool, err: dict):
        """
        将上一步收集到的错误状态应用到实际的 UI 组件上。
        使用动态属性（invalid）触发 Qt 样式表 (QSS) 变红，设置鼠标悬浮的提示文字 (ToolTip)，
        并控制底层“插入/保存”按钮的可用状态。
        """
        err = err or {}
        for k, le in (getattr(self, "inputs", {}) or {}).items():
            if le is None: continue
            bad = (k in err)
            # 通过 dynamic property 触发 QSS 样式变化（如红框警示）
            le.setProperty("invalid", bool(bad))
            le.setToolTip(err.get(k, ""))
            le.style().unpolish(le)
            le.style().polish(le)

        # 只有在验证完全通过时，才允许点击插入按钮
        if hasattr(self, "btn_insert") and self.btn_insert is not None:
            self.btn_insert.setEnabled(bool(ok))
