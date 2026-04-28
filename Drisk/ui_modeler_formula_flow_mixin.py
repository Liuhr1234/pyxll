# ui_modeler_formula_flow_mixin.py
"""
本模块提供分布构建器对话框（DistributionBuilderDialog）的公式结构解析与多段流控制混合类（Mixin）。

主要功能模块：
1. UI 状态与提示反馈：管理多段公式编辑时的界面提示信息与状态流转。
2. 公式同步与写回：负责将当前处于编辑状态的局部公式段落，安全且准确地缝合回顶部的完整公式字符串中。
3. 结构解析 (Structure Parsing)：利用正则表达式和括号平衡匹配算法，将嵌套的长公式拆解为独立的分布函数调用段（Segments）。
4. 参数与属性提取 (Activation & Extraction)：安全地解析特定段落的参数（能够避开嵌套括号或字符串内部的逗号干扰），并提取高级属性（如名称 Name、偏移 Shift、截断 Truncate 等）。
"""

import drisk_env
import re

from constants import DEFAULT_RNG_TYPE, DEFAULT_SEED

class DistributionBuilderFormulaFlowMixin:
    """
    [混合类] 负责处理复杂的嵌套公式解析、多段分布参数的提取与状态同步。
    设计意图：作为 DistributionBuilderDialog 的功能扩展模块被继承，解耦公式解析与 UI 渲染逻辑。
    """

    # =======================================================
    # 1. UI 状态与提示反馈 (UI State & Prompt Feedback)
    # =======================================================
    def _refresh_segment_mode_ui(self):
        """
        刷新多段公式编辑模式下的 UI 提示信息。
        根据当前公式段落的数量，动态控制提示标签和同步按钮的可见性与文本内容。
        """
        multi = len(self.formula_segments) > 1
        if not hasattr(self, "seg_hint"):
            return

        # 场景一：单段公式
        # 行为：隐藏提示标签和局部同步按钮，保持界面简洁。
        if not multi:
            self.seg_hint.setVisible(False)
            if hasattr(self, "btn_sync_seg"):
                self.btn_sync_seg.setVisible(False)
            return

        # 场景二：多段公式
        # 行为：计算并更新当前编辑进度提示，引导用户进行多段编辑。
        idx = max(0, int(self.current_segment_idx))
        total = len(self.formula_segments)
        dk = self.formula_segments[idx].get("dist_key", "") if self.formula_segments else ""
        
        # 不再显示顶部多段提示，避免占用公式栏右侧空间。
        self.seg_hint.setText("")
        self.seg_hint.setVisible(False)

        # 更新同步按钮的状态
        if hasattr(self, "btn_sync_seg"):
            self.btn_sync_seg.setVisible(True)
            self.btn_sync_seg.setEnabled(idx >= 0 and idx < total)

    # =======================================================
    # 2. 公式段同步与写回 (Formula Segment Sync & Write-back)
    # =======================================================
    def sync_current_segment_to_formula(self):
        """
        将当前 UI 面板中配置好的单段分布公式，局部替换并写回至顶部的完整公式输入框中。
        此过程会处理等号前缀、计算正确的字符串切片位置，并触发全局公式的重新解析。
        """
        # 边界检查：确保当前索引在合法范围内
        if self.current_segment_idx < 0 or self.current_segment_idx >= len(self.formula_segments):
            return

        # 生成当前激活段的最新公式文本
        new_seg_text = self.gen_func_str()
        full = self.formula_edit.text()
        
        # 预处理：剥离等号前缀以便于统一处理偏移量
        full_body = full[1:] if full.startswith("=") else full
        offset = 1 if full.startswith("=") else 0

        # 获取当前编辑段在原字符串中的起止索引，执行局部字符串替换
        seg = self.formula_segments[self.current_segment_idx]
        s, e = seg["start"], seg["end"]
        new_body = full_body[:s] + new_seg_text + full_body[e:]

        # 状态锁：屏蔽信号，防止 setText 触发意外的二次解析循环（避免死循环）
        self.is_updating_formula = True
        try:
            self.formula_edit.blockSignals(True)
            self.formula_edit.setText(("=" if offset else "") + new_body)
            self.formula_edit.blockSignals(False)
        finally:
            self.is_updating_formula = False

        # 数据流同步：更新全局公式状态并强制重新解析整段公式结构
        self.full_formula = self.formula_edit.text()
        self.parse_formula_structure()

        # 状态恢复：由于重新解析可能重置状态，此处恢复之前的激活段落索引并刷新 UI
        idx = min(self.current_segment_idx, len(self.formula_segments) - 1)
        self.activate_segment(idx)
        self._refresh_segment_mode_ui()

    # =======================================================
    # 3. 公式结构解析 (Formula Structure Parsing)
    # =======================================================
    def parse_formula_structure(self):
        """
        解析完整公式的结构：
        扫描公式字符串，找出所有注册的分布函数（例如 DriskNormal），
        并记录它们在字符串中的起止位置，生成公式段落（segments）列表供多段编辑流使用。
        """
        self.formula_segments = []
        formula = self.full_formula
        # 获取所有已注册的分布函数名称，如果为空则默认提供 DriskNormal 作为兜底
        funcs = [cfg['func_name'] for cfg in self._flow_get_dist_registry().values()] or ["DriskNormal"]

        # 构建正则匹配模式：动态匹配所有注册函数，形式为 "函数名(" 
        pattern = "(" + "|".join(funcs) + r")\s*\("
        
        # 核心算法：通过平衡括号（Parenthesis Balancing）的方法解析每个分布函数调用
        # 此算法可以完美兼容嵌套包装器（Nested Wrappers）或嵌套函数的情形
        for match in re.finditer(pattern, formula, re.IGNORECASE):
            func_name = match.group(1)
            start = match.start()
            cnt = 1
            end = match.end()
            
            # 向后线性扫描，利用计数器寻找闭合的最外层右括号
            while end < len(formula) and cnt > 0:
                if formula[end] == '(':
                    cnt += 1
                elif formula[end] == ')':
                    cnt -= 1
                end += 1
                
            # cnt == 0 说明成功找到了完整的函数闭合区间
            if cnt == 0:
                dk = self._flow_get_dist_key_by_func_name(func_name)
                self.formula_segments.append({'start': start, 'end': end, 'text': formula[start:end], 'dist_key': dk})

        # 降级处理 (Fallback)：当未检测到任何分布调用时，将整个公式作为一个完整的可编辑段落处理
        if not self.formula_segments:
            self.formula_segments.append({'start': 0, 'end': len(formula), 'text': formula, 'dist_key': self.dist_type})
            self.current_segment_idx = 0
        else:
            self.current_segment_idx = 0

        # 初始化：默认激活第一段公式并触发 UI 更新
        self.activate_segment(0)
        self._refresh_segment_mode_ui()

    # =======================================================
    # 4. 公式段激活与核心参数提取 (Activation & Core Extraction)
    # =======================================================
    def activate_segment(self, idx):
        """
        激活指定的公式段进行编辑：
        1. 解析该段的基础分布类型。
        2. [核心] 安全提取参数：避开嵌套括号或字符串内的逗号。
        3. [核心] 属性解析：通过正则提取高级属性（如 Name, Truncate 等）。
        4. 视图分发：初始化底层 UI 控件与渲染状态。
        """
        self.current_segment_idx = idx
        seg = self.formula_segments[idx]
        temp_f = "=" + seg['text']

        # 尝试从公式中发现首个分布特征与所有属性字典
        dk, params = self._flow_find_first_distribution_in_formula(temp_f)
        attrs = self._flow_extract_all_attributes_from_formula(temp_f)

        if dk:
            # ---------------------------------------------------
            # 4.1 成功识别分布类型，准备提取主参数
            # ---------------------------------------------------
            self.dist_type = dk
            self.config = self._flow_get_dist_config(dk)

            # 核心修复：手工提取原生参数字符串，拒绝 backend_bridge 层的强制数字转换。
            # 这保证了单元格引用（如 'A1'）或表达式不会被强转而丢失。
            p_dict = {}
            if '(' in temp_f and ')' in temp_f:
                # 提取最外层主函数括号内部的内容
                inner = temp_f[temp_f.find('(')+1: temp_f.rfind(')')]

                # 智能的安全逗号分割器（全方位保护 {}, [], (), "" 内部的逗号）
                # 目的：防止类似 DriskTruncate(0,10) 这样的参数片段被错误切碎
                parts = []
                current_part = []
                paren_level = 0    # 圆括号层级 ()
                brace_level = 0    # 花括号层级 {}
                bracket_level = 0  # 方括号层级 []
                in_q = False       # 引号内标志
                qc = ''            # 记录当前激活的引号类型 (' 或 ")

                # 遍历内部字符，仅在顶层逗号处进行分割
                for char in inner:
                    if in_q:
                        if char == qc:
                            in_q = False
                    else:
                        if char in ('"', "'"):
                            in_q, qc = True, char
                        elif char == '(':
                            paren_level += 1
                        elif char == ')':
                            paren_level -= 1
                        elif char == '{':
                            brace_level += 1
                        elif char == '}':
                            brace_level -= 1
                        elif char == '[':
                            bracket_level += 1
                        elif char == ']':
                            bracket_level -= 1

                    # 判断条件：不在任何括号层级内，且不在引号内，此时的逗号才是参数分隔符
                    if char == ',' and paren_level == 0 and brace_level == 0 and bracket_level == 0 and not in_q:
                        parts.append(''.join(current_part).strip())
                        current_part = []
                    else:
                        current_part.append(char)

                # 追加最后一部分
                if current_part:
                    parts.append(''.join(current_part).strip())

                # 将提取出的纯参数字符串与配置文件中的参数名（p_names）进行映射绑定
                parts = [p.strip() for p in parts]
                p_names = [p[0] for p in self.config.get('params', [])]

                for i, p_val in enumerate(parts):
                    # 过滤掉属于属性调用的部分（例如 @DriskName("xxx") ），只保留真正的分布参数
                    is_attr_call = bool(re.match(r"^@?\s*Drisk[A-Za-z0-9_]*\s*\(", p_val.strip(), re.IGNORECASE))
                    if i < len(p_names) and p_val.strip() and not is_attr_call:
                        p_dict[p_names[i]] = p_val.strip()

            # 缺失值回填逻辑：
            # 如果本地的纯文本切割仅提取了部分参数，则从后端解析好的参数中填补缺失的槽位，
            # 以避免悄无声息地回退到 UI 的默认值。
            if p_dict and params:
                if isinstance(params, dict):
                    for param_name in p_names:
                        if param_name not in p_dict and param_name in params:
                            param_value = str(params.get(param_name, "")).strip()
                            if param_value:
                                p_dict[param_name] = param_value
                elif isinstance(params, (list, tuple)):
                    for idx_param, param_name in enumerate(p_names):
                        if param_name in p_dict:
                            continue
                        if idx_param >= len(params):
                            continue
                        param_value = str(params[idx_param]).strip()
                        if param_value:
                            p_dict[param_name] = param_value

            # 参数应用：如果成功提取到原生字符串（如 'F6', 'E6' 或 '1.5'），则将其置入初始参数中覆盖后端数据
            self.initial_params = p_dict if p_dict else (params or {})

            # ---------------------------------------------------
            # 4.2 属性解析与合并 (Attribute Parsing & Merging)
            # ---------------------------------------------------
            # 提取基础属性（包含在 __init__ 中前置探测回填的 Name 和 Category）
            merged_attrs = self.initial_attrs.copy() if hasattr(self, 'initial_attrs') else {}
            # 合并公式中显式声明的属性（公式声明具有最高优先级）
            if attrs:
                merged_attrs.update(attrs)

            # 手动正则解析所有属性，由于原有方案可能存在盲区，此处彻底接管解析权
            
            # 子步骤 1：解析文本类属性 (Name, Category, Units)
            def _parse_str_attr(func_name, formula_str):
                """辅助函数：兼容带引号的纯文本和不带引号的单元格引用的正则提取"""
                pattern = rf'{func_name}\s*\(\s*(?:["\']([^"\']*)["\']|([^)]+))\s*\)'
                m = re.search(pattern, formula_str, re.IGNORECASE)
                if m:
                    return m.group(1) if m.group(1) is not None else m.group(2).strip()
                return None

            v_name = _parse_str_attr('DriskName', temp_f)
            if v_name:
                merged_attrs['name'] = v_name

            v_cat = _parse_str_attr('DriskCategory', temp_f)
            if v_cat:
                merged_attrs['category'] = v_cat

            v_units = _parse_str_attr('DriskUnits', temp_f)
            if v_units:
                merged_attrs['units'] = v_units

            # 子步骤 2：解析高级形态属性 (Shift 偏移, Truncate 截断等)
            m_shift = re.search(r'DriskShift\s*\(\s*([-+]?(?:\d*\.\d+|\d+))\s*\)', temp_f, re.IGNORECASE)
            if m_shift:
                merged_attrs['shift'] = float(m_shift.group(1))

            def parse_trunc_args(fname):
                """辅助函数：解析具有双参数特征的截断类属性"""
                m = re.search(rf'{fname}\s*\(\s*([^,]*)\s*,\s*([^)]*)\s*\)', temp_f, re.IGNORECASE)
                if m:
                    def _p(s):
                        s = s.strip()
                        return float(s) if s else None
                    try:
                        return (_p(m.group(1)), _p(m.group(2)))
                    except Exception:
                        pass
                return None

            t_res = parse_trunc_args('DriskTruncate')
            if t_res:
                merged_attrs['truncate'] = t_res

            t2_res = parse_trunc_args('DriskTruncate2')
            if t2_res:
                merged_attrs['truncate2'] = t2_res

            tp_res = parse_trunc_args('DriskTruncateP')
            if tp_res:
                merged_attrs['truncatep'] = tp_res

            static_match = re.search(r'DriskStatic\s*\(\s*([^)]*)\s*\)', temp_f, re.IGNORECASE)
            if static_match:
                try:
                    merged_attrs['static'] = float(static_match.group(1).strip())
                except Exception:
                    pass

            # 随机种子与随机数引擎类型的解析
            seed_match = re.search(r'DriskSeed\s*\(\s*([^)]*)\s*\)', temp_f, re.IGNORECASE)
            if seed_match:
                seed_tokens = [tok.strip() for tok in seed_match.group(1).split(',') if tok.strip()]
                if len(seed_tokens) >= 2:
                    try:
                        merged_attrs['rng_type'] = int(float(seed_tokens[0]))
                    except Exception:
                        merged_attrs['rng_type'] = int(DEFAULT_RNG_TYPE)
                    try:
                        merged_attrs['seed'] = int(float(seed_tokens[1]))
                    except Exception:
                        merged_attrs['seed'] = int(DEFAULT_SEED)
                elif len(seed_tokens) == 1:
                    merged_attrs['rng_type'] = int(DEFAULT_RNG_TYPE)
                    try:
                        merged_attrs['seed'] = int(float(seed_tokens[0]))
                    except Exception:
                        merged_attrs['seed'] = int(DEFAULT_SEED)

            # 标志位属性解析
            if re.search(r'DriskLock\s*\(', temp_f, re.IGNORECASE):
                merged_attrs['lock'] = True
            if re.search(r'DriskCollect\s*\(', temp_f, re.IGNORECASE):
                merged_attrs['collect'] = True
            
            # 保存最终的属性集
            self.initial_attrs = merged_attrs
            
        else:
            # ---------------------------------------------------
            # 4.3 兜底分支 (Fallback Branch)
            # 针对无法直接使用主逻辑识别分布类型的情况执行备选方案
            # ---------------------------------------------------
            dk = seg.get('dist_key')
            if dk:
                self.dist_type = dk
                self.config = self._flow_get_dist_config(dk)

            if not getattr(self, "initial_params", None) or idx > 0:
                self.initial_params = {}
                temp_f = self.full_formula if self.full_formula else self.gen_func_str()
                
                if '(' in temp_f and ')' in temp_f:
                    inner = temp_f[temp_f.find('(')+1: temp_f.rfind(')')]

                    # 核心修复：确保兜底分支（else 块）同样使用完整的安全切割逻辑
                    parts = []
                    current_part = []
                    paren_level = 0
                    brace_level = 0
                    bracket_level = 0
                    in_q = False
                    qc = ''
                    for char in inner:
                        if in_q:
                            if char == qc:
                                in_q = False
                        else:
                            if char in ('"', "'"):
                                in_q, qc = True, char
                            elif char == '(':
                                paren_level += 1
                            elif char == ')':
                                paren_level -= 1
                            elif char == '{':
                                brace_level += 1
                            elif char == '}':
                                brace_level -= 1
                            elif char == '[':
                                bracket_level += 1
                            elif char == ']':
                                bracket_level -= 1

                        # 仅在所有括号外层的逗号处执行切割
                        if char == ',' and paren_level == 0 and brace_level == 0 and bracket_level == 0 and not in_q:
                            parts.append(''.join(current_part).strip())
                            current_part = []
                        else:
                            current_part.append(char)
                            
                    if current_part:
                        parts.append(''.join(current_part).strip())

                    parts = [p.strip() for p in parts]

                    # 保留原有的按索引位置赋值的逻辑
                    config = self._flow_get_dist_registry().get(self.dist_type, {})
                    for i, (k, t, r) in enumerate(config.get("params", [])):
                        if i < len(parts):
                            self.initial_params[k] = parts[i]

        # ---------------------------------------------------
        # 4.4 渲染与底层视图分发同步 (Render & View Synchronization)
        # ---------------------------------------------------
        # 判断当前分布是否为离散型 (Discrete) 分布
        self.is_discrete = self.config.get('is_discrete', False)

        if hasattr(self, "_sync_dist_family_with_key"):
            self._sync_dist_family_with_key(self.dist_type)

        # 同步下拉选择框：屏蔽信号，避免在程序修改值时死循环触发重新渲染
        self.combo.blockSignals(True)
        idx_in_combo = self.combo.findData(self.dist_type)
        if idx_in_combo >= 0:
            self.combo.setCurrentIndex(idx_in_combo)
        self.combo.blockSignals(False)

        # 更新 Y 轴相关的可选项菜单
        self.update_y_combo_items()

        # 视图模式锁定：根据连续/离散属性锁定当前展现图表 
        # 离散使用概率质量分布 (PMF)，连续使用概率密度分布 (PDF)
        self._current_view_data = "pmf" if self.is_discrete else "pdf"

        # 触发底层的一系列参数重算与视图刷新工作
        self.setup_dynamic_params()
        self.setup_attribute_inputs()
        self.recalc_distribution()

        # 操作完成，最后刷新顶部的段落编辑提示状态
        self._refresh_segment_mode_ui()
