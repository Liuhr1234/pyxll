# ui_modeler_orchestration_mixin.py
"""
本模块提供分布构建器（DistributionBuilderDialog）的生命周期调度与状态编排混合类（Mixin）。

主要功能模块：
1. 模型与视图切换调度 (Model & View Switch Orchestration)：管理用户在下拉菜单中切换不同概率分布时的上下文清理与 UI 重绘。
2. 属性与高级标记收集 (Attributes & Advanced Markers Collection)：从复杂的 UI 控件群中提取并清洗数据，生成标准化的高级参数字典（如截断、平移、随机种子等）。
3. 公式生成与提交流程 (Formula Generation & Commit Flow)：基于当前 UI 状态，将其序列化并组装为符合 Drisk 语法的 Excel 公式字符串。
"""
import drisk_env

import math
import re

from PySide6.QtWidgets import QMessageBox

from constants import DEFAULT_RNG_TYPE, DEFAULT_SEED


class DistributionBuilderOrchestrationMixin:
    """
    [混合类] 负责处理分布模型切换、属性清洗校验以及最终公式序列化的核心调度逻辑。
    设计意图：将复杂的参数提取与公式拼接逻辑从主对话框类中解耦，提高代码可维护性。
    """

    # =======================================================
    # 1. 模型与视图切换调度 (Model & View Switch Orchestration)
    # =======================================================
    def on_model_changed(self, idx):
        """
        处理分布模型切换事件 (Model Changed Event)：
        当用户选择新的分布类型时触发。负责重置参数缓存、更新 Y 轴标签、重新构建动态参数输入框，并触发图表重绘。
        """
        new_key = self.combo.currentData()
        # 校验新选择的分布键值是否在注册表中，且是否发生了实质性变更
        if new_key not in self._flow_get_dist_registry():
            return
        if new_key == self.dist_type:
            return

        # 更新当前分布上下文与配置字典
        self.dist_type = new_key
        self.config = self._flow_get_dist_config(new_key)
        self.is_discrete = self.config.get('is_discrete', False)
        
        # 记录用户的近期使用偏好（若存在该扩展方法）
        if hasattr(self, '_record_recent_distribution_use'):
            self._record_recent_distribution_use(new_key)
        
        # [核心清理逻辑]：重置当前分布域下的 UI 状态缓存，以防止旧分布的参数值在模型切换时发生泄漏
        self.initial_params = {}
        self.initial_attrs = {}
        
        # 重构依赖于当前分布类型的 UI 控件
        self.update_y_combo_items()
        self.setup_dynamic_params()
        self.setup_attribute_inputs()
        
        # 视图模式推断：离散型分布采用概率质量分布 (PMF)，连续型采用概率密度分布 (PDF)
        self._current_view_data = "pmf" if self.is_discrete else "pdf"

        # 触发底层引擎重算分布特征并渲染图表
        self.recalc_distribution()
        
        # 同步顶部公式编辑栏的状态
        if len(self.formula_segments) == 1:
            self.update_formula_bar()
        else:
            self._refresh_segment_mode_ui()


    # =======================================================
    # 2. 属性与高级标记收集 (Attributes & Advanced Markers Collection)
    # =======================================================
    def _collect_markers(self) -> dict:
        """
        提取并校验 UI 面板中的所有附加标记参数（Markers）。
        涵盖静态描述属性（名称、分类、单位）以及高级形态属性（平移、截断、静态值、随机种子等）。
        返回结构化的字典供底层计算引擎使用。
        """
        markers = {}
        try:
            # 2.1 收集基础描述属性 (Base Attributes)
            if 'name' in self.attr_inputs:
                nm = self.attr_inputs['name'].text().strip()
                if nm: 
                    resolved_nm = self._resolve_string_value(nm)
                    if resolved_nm: markers['name'] = resolved_nm
            
            if 'category' in self.attr_inputs:
                cat = self.attr_inputs['category'].text().strip()
                if cat: 
                    resolved_cat = self._resolve_string_value(cat)
                    if resolved_cat: markers['category'] = resolved_cat
            
            if 'units' in self.attr_inputs:
                un = self.attr_inputs['units'].text().strip()
                if un: 
                    resolved_un = self._resolve_string_value(un)
                    if resolved_un: markers['units'] = resolved_un
        except Exception:
            pass
            
        # 2.2 收集高级形态控制属性 (Advanced Morphological Attributes)
        try:
            if hasattr(self, "combo_shift"):
                shift_mode = self.combo_shift.currentText()
                trunc_mode = self.combo_trunc.currentText()
                
                # 提取平移参数 (Shift)
                if shift_mode == "平移" and self.input_shift.text().strip():
                    markers['shift'] = float(self.input_shift.text().strip())
                        
                # 提取截断参数 (Truncation)
                if trunc_mode != "无截断":
                    t1_txt = self.input_trunc1.text().strip()
                    t2_txt = self.input_trunc2.text().strip()
                    
                    # 逻辑变更：现已支持单边截断，即只需满足上限或下限中任意一侧的输入即可触发截断逻辑。
                    if t1_txt or t2_txt:
                        
                        # [内部辅助函数]：严格的截断值解析。
                        # 留空时返回 None（交由底层引擎处理原始边界）；分位数模式下严格限幅在 0.0 至 1.0 之间。
                        def strict_parse(txt):
                            if not txt: 
                                return None
                            v = float(txt)
                            if "分位数" in trunc_mode:
                                v = max(0.0, min(1.0, v))
                            return v

                        t1, t2 = strict_parse(t1_txt), strict_parse(t2_txt)
                        
                        # [边界防呆校验]：如果下限严格大于上限，则阻断执行并向用户抛出警告。
                        if t1 is not None and t2 is not None and t1 > t2:
                            QMessageBox.warning(self, "截断参数错误", "截断下限不能大于上限！请修改。")
                            return None # 返回 None 以通知调用链中断当前的绘图流程
                        
                        # 根据不同的截断模式，将参数映射到对应的后端执行键
                        if trunc_mode in ["值截断", "值截断后平移"]:
                            markers['truncate'] = (t1, t2)
                        elif trunc_mode == "平移后值截断":
                            markers['truncate2'] = (t1, t2)
                        elif trunc_mode == "分位数截断":
                            # [关键修复]：分位数截断场景下，单侧截断的缺失端点默认值必须严格赋值为 0.0 和 1.0，禁止传入 None。
                            p1 = t1 if t1 is not None else 0.0
                            p2 = t2 if t2 is not None else 1.0
                            markers['truncatep'] = (p1, p2)
                            
        except Exception as e:
            pass

        # 2.3 收集静态替代值参数 (Static Value)
        try:
            if hasattr(self, "input_static_value"):
                static_text = self.input_static_value.text().strip()
                if static_text:
                    resolved_static = self._resolve_param_value(static_text)
                    if isinstance(resolved_static, str):
                        resolved_static = str(resolved_static).split(",")[0].strip()
                    markers["static"] = float(resolved_static)
        except Exception:
            pass

        # 2.4 收集随机数引擎与种子参数 (Random Number Generator & Seed)
        try:
            if hasattr(self, "combo_seed_mode"):
                seed_mode = str(self.combo_seed_mode.currentData() or "standard")
                if seed_mode == "custom":
                    rng_type = int(getattr(self, "_seed_custom_rng_type", DEFAULT_RNG_TYPE))
                    seed_value = int(getattr(self, "_seed_custom_value", DEFAULT_SEED))
                    # 按照后端约定的格式组合为逗号分隔的字符串
                    markers["seed"] = f"{rng_type},{seed_value}"
        except Exception:
            pass

        # 2.5 收集全局标志位控制 (Flags: Lock & Collect)
        try:
            if hasattr(self, "chk_lock") and self.chk_lock.isChecked():
                markers["lock"] = True
            if hasattr(self, "chk_collect") and self.chk_collect.isChecked():
                markers["collect"] = "all"
        except Exception:
            pass
            
        return markers


    # =======================================================
    # 3. 公式生成与提交流程 (Formula Generation & Commit Flow)
    # =======================================================
    def gen_func_str(self):
        """
        公式序列化引擎：
        根据当前 UI 中捕获的所有有效状态，动态组装并返回一条完整的 Drisk 分布公式字符串。
        处理内容包括：核心参数提取、数组括号包裹、单元格引用鉴别、形态调整函数追加等。
        """
        func = self.config['func_name']
        
        # [关键修复]：在序列化主参数时，对包含逗号的数组型输入自动补全大括号 {}。
        # 目的：防止参数在解析层发生非预期的扁平化溢出（Flattening Disaster）。
        vals = []
        for k, _, _ in self.config.get('params', []):
            val = self.inputs[k].text().strip()
            # 检查是否包含逗号且未被合理的数组/字符串标识符包裹
            if ',' in val and not any(val.startswith(c) for c in ['{', '[', '"', "'"]):
                val = f"{{{val}}}"
            vals.append(val)
            
        core = f"{func}({', '.join(vals)}"
        
        # [内部辅助函数]：智能判定字符串类型的属性是否需要强制补充双引号。
        def _format_str_attr(val):
            val = val.strip()
            if not val: return '""'
            
            # 若用户已经显式输入了单/双引号，则原样保留
            if (val.startswith('"') and val.endswith('"')) or (val.startswith("'") and val.endswith("'")):
                return val
                
            # 剥离表名引用符号 "!" 与绝对引用符号 "$" 以提取地址本体
            addr = val.split('!')[-1].replace('$', '')
            
            # 采用正则表达式验证是否符合 Excel 标准单元格引用（如 A1, AB12）。
            # 规则：单元格引用保持裸字符串；纯文本则必须使用双引号包裹。
            if re.fullmatch(r'[A-Za-z]{1,3}\d+', addr):
                return val
            return f'"{val}"'

        # 3.1 拼装描述性扩展函数 (Descriptive Extensions)
        nm = self.attr_inputs['name'].text()
        if nm: core += f', DriskName({_format_str_attr(nm)})'
        
        cat = self.attr_inputs.get('category')
        if cat and cat.text().strip(): 
            core += f', DriskCategory({_format_str_attr(cat.text())})'
            
        un = self.attr_inputs['units'].text()
        if un: core += f', DriskUnits({_format_str_attr(un)})'

        # 3.2 拼装静态值函数 (Static Value Extension)
        static_formula_value = None
        try:
            if hasattr(self, "input_static_value"):
                static_text = self.input_static_value.text().strip()
                if static_text:
                    static_val = self._resolve_param_value(static_text)
                    if isinstance(static_val, str):
                        static_val = str(static_val).split(",")[0].strip()
                    static_formula_value = float(static_val)
        except Exception:
            static_formula_value = None

        if static_formula_value is not None and math.isfinite(static_formula_value):
            core += f", DriskStatic({format(float(static_formula_value), '.15g')})"

        # 3.3 拼装随机引擎种子函数 (Seed Extension)
        try:
            if hasattr(self, "combo_seed_mode"):
                seed_mode = str(self.combo_seed_mode.currentData() or "standard")
                if seed_mode == "custom":
                    rng_type = int(getattr(self, "_seed_custom_rng_type", DEFAULT_RNG_TYPE))
                    seed_value = int(getattr(self, "_seed_custom_value", DEFAULT_SEED))
                    core += f", DriskSeed({rng_type}, {seed_value})"
        except Exception:
            pass

        # 3.4 拼装标志位控制函数 (Flags Extensions)
        try:
            if hasattr(self, "chk_lock") and self.chk_lock.isChecked():
                core += ", DriskLock()"
            if hasattr(self, "chk_collect") and self.chk_collect.isChecked():
                core += ", DriskCollect()"
        except Exception:
            pass

        # 3.5 拼装高级形态函数 (Morphological Extensions: Shift & Truncate)
        if hasattr(self, "combo_shift"):
            shift_mode = self.combo_shift.currentText()
            trunc_mode = self.combo_trunc.currentText()
            
            shift_str = ""
            if shift_mode == "平移" and self.input_shift.text().strip():
                shift_str = f', DriskShift({self.input_shift.text().strip()})'
                
            trunc_str = ""
            if trunc_mode != "无截断":
                t1 = self.input_trunc1.text().strip()
                t2 = self.input_trunc2.text().strip()
                
                # 兼容性逻辑：保留底层引擎支持的单侧空端点语法（例如生成 DriskTruncate( , 10) 是完全合法的）。
                if t1 or t2:
                    if trunc_mode in ["值截断", "值截断后平移"]:
                        trunc_str = f', DriskTruncate({t1}, {t2})'
                    elif trunc_mode == "平移后值截断":
                        trunc_str = f', DriskTruncate2({t1}, {t2})'
                    elif trunc_mode == "分位数截断":
                        # [关键修复]：公式序列化时，分位数模式的缺失端点同样需要显式写入为 "0.0" 或 "1.0"，防止 Excel 计算期崩溃。
                        _t1 = t1 if t1 else "0.0"
                        _t2 = t2 if t2 else "1.0"
                        trunc_str = f', DriskTruncateP({_t1}, {_t2})'
            
            # 将形态调整函数附加至主体字符串末尾
            core += trunc_str + shift_str
            
        core += ")"
        return core

    def on_accept(self):
        """
        对话框确认提交钩子 (Accept Hook)：
        如果用户手动修改了顶部公式编辑栏，则优先保存用户的编辑内容；
        否则调用公式生成器生成标准字符串。随后结束并关闭对话框进程。
        """
        if self.formula_edit.text().strip():
            self.result_formula = self.formula_edit.text()
        else:
            self.result_formula = "=" + self.gen_func_str()
        self.accept()
        