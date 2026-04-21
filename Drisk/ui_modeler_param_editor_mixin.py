# ui_modeler_param_editor_mixin.py
import drisk_env

from pyxll import xl_app
from PySide6.QtGui import QIntValidator
from PySide6.QtWidgets import (
    QCheckBox,
    QComboBox,
    QDialog,
    QDialogButtonBox,
    QFormLayout,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QVBoxLayout,
)

from constants import DEFAULT_RNG_TYPE, DEFAULT_SEED, RNG_TYPE_NAMES, DISTRIBUTION_PARAM_TOOLTIPS
from ui_shared import AutoSelectLineEdit, apply_excel_select_button_icon


class DistributionBuilderParamEditorMixin:
    """
    [混合类] 分布构建器（DistributionBuilderDialog）参数编辑器与公式同步辅助类。
    
    设计意图与职责边界：
    本模块将主对话框中与“输入表单”相关的繁杂逻辑剥离出来，主要负责：
    1. 动态表单渲染：根据不同的分布函数（如正态、对数正态）动态销毁并重建对应的参数输入 UI。
    2. Excel 深度交互：通过 PyXLL 唤起底层的 Excel InputBox，实现单元格或区域的框选绑定。
    3. 状态与属性管理：管理分布的高级属性（如随机种子、截断、平移、静态值等）的 UI 状态。
    4. 双向数据同步：处理用户输入事件，执行输入防抖（Debounce），并在 UI 参数与顶部的 Drisk 公式字符串之间进行双向同步。
    """

    # =======================================================
    # 1. 动态参数渲染 (Dynamic Parameter Setup)
    # =======================================================
    def setup_dynamic_params(self):
        """
        动态参数表单构建器：
        根据当前选择的分布模型（存储在 self.config 中），动态生成对应的参数输入框（如：正态分布的“均值”、“标准差”）。
        为防止内存泄漏，在构建新表单前会严格清理并销毁旧的 UI 组件。
        """
        # 开启正在构建的锁，防止在构建过程中触发 on_param_changed 导致死循环或不必要的重算
        self._is_building_inputs = True

        # 1. 安全清空旧的参数输入组件及布局 (避免内存泄漏与界面重叠)
        while self.form.count():
            item = self.form.takeAt(0)
            w = item.widget()
            if w:
                w.deleteLater()
            else:
                lay = item.layout()
                if lay:
                    while lay.count():
                        sub_item = lay.takeAt(0)
                        if sub_item.widget():
                            sub_item.widget().deleteLater()
                    lay.deleteLater()

        self.inputs = {}

        # 2. 从当前分布配置中提取所需的参数元数据列表
        params_def = self.config.get('params', [])
        for i, (key, label, default) in enumerate(params_def):
            param_tooltip = ""
            try:
                tooltip_map = DISTRIBUTION_PARAM_TOOLTIPS.get(str(getattr(self, "dist_type", "") or ""), {})
                param_tooltip = str(
                    tooltip_map.get(str(key), "") or tooltip_map.get(str(label), "")
                ).strip()
            except Exception:
                param_tooltip = ""

            val = str(default)
            # 尝试回填初始传入的参数值（支持字典通过 key 匹配，或列表通过索引匹配）
            if isinstance(self.initial_params, dict) and key in self.initial_params:
                val = str(self.initial_params[key])
            elif isinstance(self.initial_params, (list, tuple)) and i < len(self.initial_params):
                val = str(self.initial_params[i])

            # 实例化自定义的输入框组件 
            # (注：输入框的尺寸限制和全选逻辑封装在 AutoSelectLineEdit 类内部)
            e = AutoSelectLineEdit(val)
            if param_tooltip:
                e.setToolTip(param_tooltip)
            e.editingFinished.connect(self.on_param_changed)

            # 3. 实例化“靶点”按钮，用于选取 Excel 单元格地址
            btn = QPushButton()
            btn.setFixedSize(28, 24)
            
            # 判断当前参数是否支持多单元格选取（数组参数）
            is_list_param = self._is_list_param_selector(key)
            if is_list_param:
                btn.setToolTip("从 Excel 选择一或多个单元格作为参数")
            else:
                btn.setToolTip("从 Excel 选择一个单元格作为参数")
                
            btn.setStyleSheet("""
                QPushButton { background-color: #f5f5f5; border: 1px solid #d9d9d9; border-radius: 3px; font-size: 11px;}
                QPushButton:hover { background-color: #e6f7ff; border-color: #40a9ff; color: #40a9ff;}
                QPushButton:pressed { background-color: #bae0ff; }
            """)
            
            # 加载选取图标，若图标文件缺失则降级使用 Unicode 靶心符号
            if not apply_excel_select_button_icon(btn, "select_icon.svg"):
                btn.setText("\U0001F3AF")

            # 绑定点击事件，使用 lambda 闭包传递当前输入框对象和参数键名
            btn.clicked.connect(lambda checked=False, le=e, pk=key: self._select_cell_for_input(le, pk))

            # 4. 水平布局组装：将输入框和靶点按钮组合在同一行
            h_layout = QHBoxLayout()
            h_layout.setContentsMargins(0, 0, 0, 0)
            h_layout.setSpacing(6)
            h_layout.addWidget(e)
            h_layout.addWidget(btn)

            # 挂载到主表单布局中，并记录引用
            label_widget = QLabel(str(label))
            if param_tooltip:
                label_widget.setToolTip(param_tooltip)
            self.form.addRow(label_widget, h_layout)
            self.inputs[key] = e

        self._is_building_inputs = False


    # =======================================================
    # 2. Excel 选区交互 (Excel Range Selection)
    # =======================================================
    def _select_cell_for_input(self, target_lineedit, param_key=None):
        """
        Excel 选区拦截器：
        通过 PyXLL 唤起 Excel 原生的 InputBox，允许用户直接在表格画布上框选单元格区域。
        获取到 Range 对象后，解析地址并回填至对应的 Qt 输入框中。
        """
        try:
            app = xl_app()
        except ImportError:
            QMessageBox.warning(self, "环境错误", "PyXLL 未加载，无法与 Excel 进行底层交互。")
            return

        allow_multi = self._is_list_param_selector(param_key)

        # [关键设计]：将当前弹窗透明度设为 0（即隐藏），防止 Qt 窗口遮挡 Excel 的 InputBox 提示框
        self.setWindowOpacity(0)
        try:
            if allow_multi:
                prompt = "请在 Excel 中选择数据区域（可按住 Ctrl 键多选）:"
                title = "选择单元格区域"
            else:
                prompt = "请在 Excel 中选择单个单元格:"
                title = "选择单元格"

            # 唤起 Excel InputBox。Type=8 属于 Excel API 规范，表示要求返回一个 Range 对象
            rng = app.InputBox(Prompt=prompt, Title=title, Type=8)
            if not rng:
                return

            # 解析获取到的单元格地址并写入目标输入框
            if allow_multi:
                cells = []
                # 处理可能存在的不连续多选区域 (Areas)
                for area in rng.Areas:
                    for cell in area.Cells:
                        cells.append(self._format_excel_ref_for_input(app, cell))
                if not cells:
                    return
                target_lineedit.setText(", ".join(cells))
            else:
                # 单选模式下强制只取第一个区域的第一个单元格
                cell = rng.Areas(1).Cells(1)
                target_lineedit.setText(self._format_excel_ref_for_input(app, cell))

            # 数据回填完成后，主动触发一次参数变更校验与图表刷新
            self.on_param_changed()
            if allow_multi:
                self._show_list_length_mismatch_warning_if_needed()
        except Exception:
            # 捕获用户在 InputBox 点击“取消”造成的静默异常
            pass
        finally:
            # [关键恢复]：交互完成后，必须将窗口透明度恢复，并将其重新推至前台获取焦点
            self.setWindowOpacity(1)
            self.raise_()
            self.activateWindow()


    # =======================================================
    # 3. 静态与高级属性渲染 (Static & Advanced Attributes Setup)
    # =======================================================
    def create_static_attribute_fields(self):
        """初始化分布的基础描述性字段 UI（如名称、类别、业务单位）"""
        self.add_attr("名称", "name")
        self.add_attr("类别", "category")
        self.add_attr("单位", "units")

    def create_distribution_attribute_fields(self):
        """初始化分布的底层数学属性与采样控制字段 UI（如静态值替代、随机种子、采样锁定）"""
        self.input_static_value = AutoSelectLineEdit()
        self.input_static_value.setPlaceholderText("留空时默认使用分布均值")
        self.input_static_value.editingFinished.connect(self.on_param_changed)
        self.dist_attr_form.addRow("静态值", self.input_static_value)
        self.dist_attr_inputs["static"] = self.input_static_value

        self.combo_seed_mode = QComboBox()
        self.combo_seed_mode.addItem("标准", "standard")
        self.combo_seed_mode.addItem("自定义", "custom")
        self.combo_seed_mode.currentIndexChanged.connect(self._on_seed_mode_changed)
        self.dist_attr_form.addRow("种子", self.combo_seed_mode)
        self.dist_attr_inputs["seed_mode"] = self.combo_seed_mode

        self.chk_lock = QCheckBox()
        self.chk_lock.stateChanged.connect(lambda *_: self.on_param_changed())
        self.dist_attr_form.addRow("锁定", self.chk_lock)
        self.dist_attr_inputs["lock"] = self.chk_lock

        self.chk_collect = QCheckBox()
        self.chk_collect.stateChanged.connect(lambda *_: self.on_param_changed())
        self.dist_attr_form.addRow("样本收集", self.chk_collect)
        self.dist_attr_inputs["collect"] = self.chk_collect

        # 初始化自定义种子相关内存变量
        if getattr(self, "_seed_custom_rng_type", None) is None:
            self._seed_custom_rng_type = int(DEFAULT_RNG_TYPE)
        if getattr(self, "_seed_custom_value", None) is None:
            self._seed_custom_value = int(DEFAULT_SEED)
            
        self._set_seed_combo_custom_label()
        self._set_seed_mode("standard", emit_change=False)

    def _set_seed_combo_custom_label(self):
        """更新下拉框中‘自定义’选项的显示文本，携带当前的种子数值以便用户直观查看"""
        if not hasattr(self, "combo_seed_mode"):
            return
        if bool(getattr(self, "_seed_custom_confirmed", False)):
            seed_text = str(getattr(self, "_seed_custom_value", DEFAULT_SEED))
            self.combo_seed_mode.setItemText(1, f"自定义({seed_text})")
        else:
            self.combo_seed_mode.setItemText(1, "自定义")

    def _set_seed_mode(self, mode_key: str, emit_change: bool = False):
        """安全地设置种子下拉框状态，使用 _seed_mode_syncing 标志位防止死循环触发信号"""
        if not hasattr(self, "combo_seed_mode"):
            return
        idx = self.combo_seed_mode.findData(mode_key)
        if idx < 0:
            idx = 0
            
        self._seed_mode_syncing = True
        try:
            self.combo_seed_mode.blockSignals(True)
            self.combo_seed_mode.setCurrentIndex(idx)
            self.combo_seed_mode.blockSignals(False)
        finally:
            self._seed_mode_syncing = False
            
        if emit_change:
            self.on_param_changed()

    def _open_custom_seed_dialog(self) -> bool:
        """
        自定义种子配置弹窗：
        提供给用户选择随机数生成器（RNG Type）以及具体种子数值的专用窗口。
        返回 bool 值标识用户是否确保持了修改（点击了 OK）。
        """
        dlg = QDialog(self)
        dlg.setWindowTitle("Drisk - 自定义种子配置")
        layout = QVBoxLayout(dlg)
        form = QFormLayout()

        # 生成器类型下拉框
        rng_combo = QComboBox(dlg)
        for rng_type, rng_name in sorted(RNG_TYPE_NAMES.items(), key=lambda kv: int(kv[0])):
            rng_combo.addItem(str(rng_name), int(rng_type))
            
        target_rng = int(getattr(self, "_seed_custom_rng_type", DEFAULT_RNG_TYPE))
        rng_idx = rng_combo.findData(target_rng)
        if rng_idx >= 0:
            rng_combo.setCurrentIndex(rng_idx)
        form.addRow("随机模式", rng_combo)

        # 种子数值输入框 (限定为正整数)
        seed_edit = QLineEdit(dlg)
        seed_edit.setValidator(QIntValidator(0, 2147483647, seed_edit))
        seed_edit.setText(str(int(getattr(self, "_seed_custom_value", DEFAULT_SEED))))
        form.addRow("随机种子", seed_edit)
        layout.addLayout(form)

        tip = QLabel("确认后主界面“种子”项会显示为当前配置的自定义种子值。", dlg)
        tip.setStyleSheet("color:#666666; font-size:11px;")
        layout.addWidget(tip)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel, parent=dlg)
        buttons.accepted.connect(dlg.accept)
        buttons.rejected.connect(dlg.reject)
        layout.addWidget(buttons)

        if dlg.exec() != QDialog.Accepted:
            return False

        seed_text = (seed_edit.text() or "").strip()
        if not seed_text:
            QMessageBox.warning(self, "种子数值无效", "请输入有效的随机种子数值。")
            return False
            
        try:
            seed_value = int(seed_text)
        except Exception:
            QMessageBox.warning(self, "种子数值无效", "请输入有效的随机种子数值。")
            return False

        # 保存用户配置的状态
        self._seed_custom_rng_type = int(rng_combo.currentData() or DEFAULT_RNG_TYPE)
        self._seed_custom_value = int(seed_value)
        self._seed_custom_confirmed = True
        
        self._set_seed_combo_custom_label()
        return True

    def _on_seed_mode_changed(self, _idx):
        """处理种子模式下拉框变更事件，若选择自定义则自动弹出配置窗口"""
        if bool(getattr(self, "_seed_mode_syncing", False)):
            return
            
        mode_key = str(self.combo_seed_mode.currentData() or "standard")
        if mode_key == "custom":
            ok = self._open_custom_seed_dialog()
            if not ok:
                # 若取消配置，需根据之前是否已确认过来决定回退状态
                fallback_mode = "custom" if bool(getattr(self, "_seed_custom_confirmed", False)) else "standard"
                self._set_seed_mode(fallback_mode, emit_change=False)
                return
            self._set_seed_mode("custom", emit_change=False)
            self.on_param_changed()
            return
            
        self.on_param_changed()

    def add_attr(self, label, key, val=None):
        """辅助方法：生成标准的基础属性输入行（输入框 + Excel靶心选区按钮）并挂载至 UI"""
        e = AutoSelectLineEdit()
        if val: e.setValidator(val)
        e.editingFinished.connect(self.on_param_changed)
        
        btn = QPushButton()
        btn.setFixedSize(28, 24)
        btn.setToolTip("从 Excel 选择单元格作为属性参数")
        btn.setStyleSheet("""
            QPushButton { background-color: #f5f5f5; border: 1px solid #d9d9d9; border-radius: 3px; font-size: 11px;}
            QPushButton:hover { background-color: #e6f7ff; border-color: #40a9ff; color: #40a9ff;}
            QPushButton:pressed { background-color: #bae0ff; }
        """)
        if not apply_excel_select_button_icon(btn, "select_icon.svg"):
            btn.setText("\U0001F3AF")
        btn.clicked.connect(lambda checked=False, le=e: self._select_cell_for_input(le))
        
        h_layout = QHBoxLayout()
        h_layout.setContentsMargins(0, 0, 0, 0)
        h_layout.setSpacing(6)
        h_layout.addWidget(e)
        h_layout.addWidget(btn)
        
        self.attr_form.addRow(label, h_layout)
        self.attr_inputs[key] = e

    def setup_attribute_inputs(self):
        """
        初始化反向解析：
        当从已有公式启动弹窗时，将公式内的各项属性参数（基础属性与高级截断/平移）回填至界面的各个控件中。
        """
        # 1. 恢复基础文本型属性 (如 Name, Units)
        for k, e in self.attr_inputs.items():
            val = self.initial_attrs.get(k, "")
            # 兼容布尔类型属性的字符串呈现
            if val is True: val = "TRUE"
            elif val is False: val = "FALSE"
            e.setText(str(val))
            
        # 2. 恢复底层高级配置属性
        static_val = self.initial_attrs.get("static", None)
        if hasattr(self, "input_static_value"):
            if isinstance(static_val, bool) or static_val is None:
                self.input_static_value.setText("")
            else:
                self.input_static_value.setText(str(static_val))

        # 恢复种子配置状态
        seed_val = self.initial_attrs.get("seed", None)
        rng_type_val = self.initial_attrs.get("rng_type", DEFAULT_RNG_TYPE)
        if seed_val is None:
            self._seed_custom_rng_type = int(DEFAULT_RNG_TYPE)
            self._seed_custom_value = int(DEFAULT_SEED)
            self._seed_custom_confirmed = False
            self._set_seed_combo_custom_label()
            self._set_seed_mode("standard", emit_change=False)
        else:
            try:
                self._seed_custom_value = int(float(seed_val))
            except Exception:
                self._seed_custom_value = int(DEFAULT_SEED)
            try:
                self._seed_custom_rng_type = int(float(rng_type_val))
            except Exception:
                self._seed_custom_rng_type = int(DEFAULT_RNG_TYPE)
            self._seed_custom_confirmed = True
            self._set_seed_combo_custom_label()
            self._set_seed_mode("custom", emit_change=False)

        # 恢复勾选框状态
        lock_flag = bool(self.initial_attrs.get("lock", False))
        collect_flag = bool(self.initial_attrs.get("collect", False))
        if hasattr(self, "chk_lock"):
            self.chk_lock.blockSignals(True)
            self.chk_lock.setChecked(lock_flag)
            self.chk_lock.blockSignals(False)
        if hasattr(self, "chk_collect"):
            self.chk_collect.blockSignals(True)
            self.chk_collect.setChecked(collect_flag)
            self.chk_collect.blockSignals(False)

        if not hasattr(self, "combo_shift"): return
        
        # 3. 恢复分布的形态转换参数 (平移与截断)
        shift_val = self.initial_attrs.get('shift')
        trunc = self.initial_attrs.get('truncate')
        truncp = self.initial_attrs.get('truncatep')
        trunc2 = self.initial_attrs.get('truncate2')
        
        # [防御性修复设计]：由于早期版本的或底层的解析可能返回错误的 bool 值，需要进行清洗拦截
        if isinstance(shift_val, bool): shift_val = None
        if isinstance(trunc, bool): trunc = None
        if isinstance(truncp, bool): truncp = None
        if isinstance(trunc2, bool): trunc2 = None
        
        has_shift = shift_val is not None
        has_trunc = trunc is not None or truncp is not None or trunc2 is not None
        
        self.combo_shift.blockSignals(True)
        self.combo_trunc.blockSignals(True)
        
        # 回填平移数据
        if has_shift:
            self.combo_shift.setCurrentText("平移")
            self.input_shift.setText(str(shift_val))
        else:
            self.combo_shift.setCurrentText("无平移")
            self.input_shift.setText("")
            
        # 依据平移状态动态刷新 Trunc 截断选项列表（截断和平移的顺序是关键业务逻辑）
        if has_shift:
            self.combo_trunc.clear()
            self.combo_trunc.addItems(["无截断", "值截断后平移", "平移后值截断", "分位数截断"])
        else:
            self.combo_trunc.clear()
            self.combo_trunc.addItems(["无截断", "值截断", "分位数截断"])
            
        # 安全转译函数：当截断区间单边为空 (None) 时，返回空字符串而不是 "None"
        def _fmt(val):
            return str(val) if val is not None else ""

        # 分类回填具体截断模式下的边界值
        if trunc is not None:
            self.combo_trunc.setCurrentText("值截断后平移" if has_shift else "值截断")
            self.input_trunc1.setText(_fmt(trunc[0]))
            self.input_trunc2.setText(_fmt(trunc[1]))
        elif trunc2 is not None:
            self.combo_trunc.setCurrentText("平移后值截断")
            self.input_trunc1.setText(_fmt(trunc2[0]))
            self.input_trunc2.setText(_fmt(trunc2[1]))
        elif truncp is not None:
            self.combo_trunc.setCurrentText("分位数截断")
            self.input_trunc1.setText(_fmt(truncp[0]))
            self.input_trunc2.setText(_fmt(truncp[1]))
        else:
            self.combo_trunc.setCurrentText("无截断")
            self.input_trunc1.setText("")
            self.input_trunc2.setText("")
            
        self.combo_shift.blockSignals(False)
        self.combo_trunc.blockSignals(False)
        
        # 触发布局刷新
        self._update_adv_inputs_layout()
        
        # 交互优化：如果当前公式确实包含高级形态，自动展开高级形态配置面板以方便用户查看
        if has_shift or has_trunc:
            self.btn_adv_morph.setChecked(True)
            self._toggle_adv_panel(True)


    # =======================================================
    # 4. 参数变更与公式同步 (Parameter Events & Formula Sync)
    # =======================================================
    def on_param_changed(self):
        """
        统一的 UI 变更触发器：
        处理任何输入框内容修改后的初步校验。
        如果参数非法则阻断绘图并给出视觉提示；如果合法则重启防抖定时器（Debounce Countdown），以防高频更新导致卡顿。
        """
        # 防止在代码批量构建 UI 过程中触发业务逻辑
        if getattr(self, "_is_building_inputs", False):
            return
            
        # 严格校验当前所有的表单输入参数
        ok, _vals, err = self._validate_params_strict()
        self._apply_param_validation_ui(ok, err)

        if not ok:
            # 针对累积分布（Cumul）等特殊分布可能需要弹出的专门指导性提示
            self._maybe_show_cumul_validation_prompt(err)
            # 若已有防抖倒计时则立即终止，彻底阻断无效重算
            if hasattr(self, "_param_debounce") and self._param_debounce.isActive():
                self._param_debounce.stop()
            return

        self._last_cumul_validation_msg = ""
        
        # 参数合法，启动或重启防抖定时器，计时结束后真正执行重绘（_flush_param_change）
        if hasattr(self, "_param_debounce"):
            self._param_debounce.start()

    def _flush_param_change(self):
        """
        防抖延迟执行器 (Debounced Executor)：
        由 _param_debounce 定时器触发。在此方法中执行耗时操作，如底层分布数据重计算、图表全量刷新以及顶部公式同步。
        """
        ok, _vals, err = self._validate_params_strict()
        self._apply_param_validation_ui(ok, err)
        if not ok:
            self._maybe_show_cumul_validation_prompt(err)
            return

        self._last_cumul_validation_msg = ""
        
        # 拖拽冲突保护：如果用户仍在拖拽图表底部的区间滑块，为避免相互干扰，推迟本次绘图
        if getattr(self, "_is_slider_dragging", False):
            self._param_debounce.start()
            return

        # 核心入口：调用底层数学引擎根据新参数重构分布
        self.recalc_distribution()
        
        # 同步更新顶部公式编辑栏
        if len(self.formula_segments) == 1:
            self.update_formula_bar()
        else:
            # 如果是多段公式组合，仅刷新局部 UI 以防覆盖整体公式结构
            self._refresh_segment_mode_ui()

    def update_formula_bar(self):
        """
        公式双向同步器：
        将当前表单里的离散参数拼装转化为完整的 Drisk 字符串函数格式（如 '=DriskNormal(10,2)'），
        并将其写回顶部的公式编辑框内，实现 UI -> 字符串 的同步。
        """
        if len(self.formula_segments) != 1: return
        
        new_str = self.gen_func_str()

        # 使用标志位锁定信号，防止公式框 textChanged 信号二次触发导致死循环回调
        self.is_updating_formula = True
        try:
            self.formula_edit.blockSignals(True)
            self.formula_edit.setText("=" + new_str)
            self.formula_edit.blockSignals(False)
        finally:
            self.is_updating_formula = False

        # 更新内部记录的公式全文及分段游标信息
        self.full_formula = self.formula_edit.text()
        self.formula_segments[0]["text"] = new_str
        self.formula_segments[0]["start"] = 0
        self.formula_segments[0]["end"] = len(new_str)
        self.current_segment_idx = 0

    def on_formula_cursor_move(self, o, n):
        """
        顶部公式编辑框光标移动监听器：
        当支持解析包含多个分布函数组合的复杂公式时（例如 '=DriskNormal(...) + DriskUniform(...)'），
        通过判断当前光标停留在哪个函数的字符区间，自动切片并切换下方表单激活的对应编辑面板。
        """
        if self.is_updating_formula: return
        
        for i, s in enumerate(self.formula_segments):
            # 判断光标是否处于当前片段 [start, end] 区间内
            if s['start'] <= n <= s['end']:
                if i != self.current_segment_idx:
                    # 如果发生跨段落移动，激活新段落的参数面板
                    self.activate_segment(i)
                    
                # 【体验优化】：强制让光标选中当前分段的整段文本，以便用户感知当前操作区域
                self.formula_edit.blockSignals(True)
                self.formula_edit.setSelection(s['start'], s['end'] - s['start'])
                self.formula_edit.blockSignals(False)
                break
