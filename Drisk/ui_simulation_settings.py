# ui_simulation_settings.py
"""
本模块提供“模拟设置”（Simulation Settings）的交互式对话框组件。

主要功能：
1. 界面构建 (UI Construction)：基于 PySide6 构建模块化的设置面板，目前主要包含“抽样”配置选项卡。
2. 状态管理 (State Management)：提供对抽样模式（MC/LHC）、随机数生成器（RNG）类型、种子策略（固定/默认）以及输出样本范围的可视化配置。
3. 数据校验与导出 (Validation & Export)：提供表单防呆校验，并将用户确认的设置封装为标准化字典，供底层仿真引擎调用。
4. 全局调用接口 (Global Entry Point)：提供独立的弹出函数 `show_simulation_settings_dialog`，屏蔽底层窗口实例化的细节。
"""

import drisk_env

import sys
from typing import Any, Dict, Optional

from PySide6.QtCore import Qt
from PySide6.QtGui import QIntValidator
from PySide6.QtWidgets import (
    QApplication,
    QComboBox,
    QDialog,
    QFormLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QPushButton,
    QSizePolicy,
    QTabWidget,
    QVBoxLayout,
    QWidget,
)

from constants import DEFAULT_RNG_TYPE, DEFAULT_SEED, RNG_TYPE_NAMES
from ui_shared import DRISK_DIALOG_BTN_QSS, drisk_combobox_qss, set_drisk_icon


# =======================================================
# 1. 对话框主类定义与初始化
# =======================================================
class SimulationSettingsDialog(QDialog):
    """
    模拟设置主对话框类。
    负责管理所有的模拟配置项（当前包含单一的“抽样”选项卡）。
    """

    def __init__(self, current_settings: Optional[Dict[str, Any]] = None, parent=None):
        """
        初始化模拟设置对话框。
        
        :param current_settings: 当前已有的设置字典，若为空则使用默认初始状态。
        :param parent: 父级窗口指针。
        """
        super().__init__(parent)
        self._current = current_settings or {}

        # --- 窗口基础属性设置 ---
        self.setWindowTitle("Drisk - 模拟设置")
        set_drisk_icon(self, "simu_icon.svg")
        self.resize(500, 360)
        self.setModal(True)  # 设置为模态对话框，阻塞其他窗口交互
        
        # --- 布局尺寸常量预设 ---
        self._label_col_width = 112
        self._field_col_min_width = 260
        
        # --- 窗口全局样式表 (QSS) ---
        self.setStyleSheet(
            """
            QDialog { background-color: #f5f6f8; }
            QGroupBox {
                border: 1px solid #d9d9d9;
                border-radius: 6px;
                margin-top: 10px;
                font-size: 12px;
                font-weight: bold;
                color: #333333;
                background: #ffffff;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 4px;
            }
            QLabel {
                color: #333333;
                font-size: 12px;
            }
            """
        )

        # --- 根布局构建 ---
        root = QVBoxLayout(self)
        root.setContentsMargins(14, 14, 14, 14)
        root.setSpacing(10)

        # 构建选项卡容器
        self.tabs = QTabWidget()
        self.tabs.addTab(self._build_sampling_tab(), "抽样")
        self.tabs.setCurrentIndex(0)
        root.addWidget(self.tabs, 1)

        # 挂载共享的按钮样式
        self.setStyleSheet(self.styleSheet() + DRISK_DIALOG_BTN_QSS)
        
        # --- 底部操作按钮栏 ---
        button_bar = QHBoxLayout()
        button_bar.addStretch()  # 将按钮推至右侧
        
        btn_ok = QPushButton("确定")
        btn_ok.setObjectName("btnOk")
        btn_ok.clicked.connect(self._validate_and_accept)
        
        btn_cancel = QPushButton("取消")
        btn_cancel.setObjectName("btnCancel")
        btn_cancel.clicked.connect(self.reject)
        
        button_bar.addWidget(btn_ok)
        button_bar.addWidget(btn_cancel)
        root.addLayout(button_bar)

        # --- 加载并回显初始数据 ---
        self._load_settings(self._current)


# =======================================================
# 2. 界面辅助构建方法 (UI Helper Methods)
# =======================================================
    def _configure_form_layout(self, form: QFormLayout):
        """
        统一配置表单布局 (QFormLayout) 的样式与对齐规则。
        """
        form.setLabelAlignment(Qt.AlignRight | Qt.AlignVCenter)
        form.setHorizontalSpacing(10)
        form.setVerticalSpacing(8)
        form.setFormAlignment(Qt.AlignLeft | Qt.AlignTop)
        form.setFieldGrowthPolicy(QFormLayout.AllNonFixedFieldsGrow)

    def _create_form_label(self, text: str) -> QLabel:
        """
        创建标准化的表单标签。
        确保所有标签宽度一致，从而使表单列对齐整齐。
        """
        label = QLabel(text)
        label.setFixedWidth(self._label_col_width)
        label.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        return label

    def _configure_form_field(self, field: QWidget):
        """
        统一配置表单输入控件（如 QComboBox, QLineEdit）的尺寸策略。
        """
        field.setMinimumWidth(self._field_col_min_width)
        policy = field.sizePolicy()
        policy.setHorizontalPolicy(QSizePolicy.Expanding)
        field.setSizePolicy(policy)


# =======================================================
# 3. 核心视图构建 (Core View Construction)
# =======================================================
    def _build_sampling_tab(self):
        """
        构建“抽样”选项卡页面。
        包含三个核心配置组：抽样模式、随机数种子、收集数据范围。
        """
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(8, 8, 8, 8)
        layout.setSpacing(10)

        # ----------------------------------------
        # 组 1：抽样模式配置 (Sampling Mode & RNG)
        # ----------------------------------------
        group_rng = QGroupBox("抽样模式")
        form_rng = QFormLayout(group_rng)
        self._configure_form_layout(form_rng)
        
        # 采样方式下拉框 (MC - 蒙特卡洛 / LHC - 拉丁超立方)
        self.combo_sampling_mode = QComboBox()
        self.combo_sampling_mode.setStyleSheet(drisk_combobox_qss())
        self.combo_sampling_mode.addItem("MC", "MC")
        self.combo_sampling_mode.addItem("LHC", "LHC")
        self._configure_form_field(self.combo_sampling_mode)

        # 随机数生成器下拉框
        self.combo_rng_type = QComboBox()
        self.combo_rng_type.setStyleSheet(drisk_combobox_qss())
        self._configure_form_field(self.combo_rng_type)
        # 根据系统常量动态填充 RNG 选项，并确保按 ID 排序
        for rng_type, name in sorted(RNG_TYPE_NAMES.items(), key=lambda kv: int(kv[0])):
            self.combo_rng_type.addItem(str(name), int(rng_type))
            
        form_rng.addRow(self._create_form_label("采样方式："), self.combo_sampling_mode)
        form_rng.addRow(self._create_form_label("随机数生成器："), self.combo_rng_type)
        layout.addWidget(group_rng)

        # ----------------------------------------
        # 组 2：随机数种子配置 (Random Seed)
        # ----------------------------------------
        group_seed = QGroupBox("随机数种子")
        form_seed = QFormLayout(group_seed)
        self._configure_form_layout(form_seed)

        # 种子模式下拉框
        self.combo_seed_mode = QComboBox()
        self.combo_seed_mode.setStyleSheet(drisk_combobox_qss())
        self.combo_seed_mode.addItem("固定种子", "fixed")
        self.combo_seed_mode.addItem("使用当前默认策略", "default")
        self._configure_form_field(self.combo_seed_mode)
        # 绑定状态联动逻辑：当模式改变时，动态启用或禁用数值输入框
        self.combo_seed_mode.currentIndexChanged.connect(self._sync_seed_mode_state)
        form_seed.addRow(self._create_form_label("种子模式："), self.combo_seed_mode)

        # 种子数值输入框
        self.edit_seed = QLineEdit()
        # 限制输入范围为合法正整数 (0 到 2^31 - 1)
        self.edit_seed.setValidator(QIntValidator(0, 2147483647, self))
        self.edit_seed.setPlaceholderText("请输入整数种子")
        self._configure_form_field(self.edit_seed)
        form_seed.addRow(self._create_form_label("种子值："), self.edit_seed)
        layout.addWidget(group_seed)

        # ----------------------------------------
        # 组 3：收集数据范围配置 (Data Collection Scope)
        # ----------------------------------------
        group_scope = QGroupBox("收集数据范围")
        form_scope = QFormLayout(group_scope)
        self._configure_form_layout(form_scope)
        
        self.combo_scope = QComboBox()
        self.combo_scope.setStyleSheet(drisk_combobox_qss())
        self.combo_scope.addItem("全部", "all")
        self.combo_scope.addItem("仅 Collect 标记", "collect")
        self.combo_scope.addItem("无", "none")
        self._configure_form_field(self.combo_scope)
        form_scope.addRow(self._create_form_label("输入样本可用范围："), self.combo_scope)
        layout.addWidget(group_scope)

        # 底部补充说明文案
        hint = QLabel(
            "该设置仅控制模拟完成后输入变量样本的可见/可用范围，不影响模拟执行范围。"
            "“仅 Collect 标记”只暴露带 Collect 标记的输入样本，“无”则不暴露任何输入样本。"
        )
        hint.setWordWrap(True)
        hint.setStyleSheet("color:#666666;")
        layout.addWidget(hint)

        # 增加弹性空间，将所有配置推至页面顶部
        layout.addStretch(1)
        return page


# =======================================================
# 4. 状态同步与数据生命周期 (State Sync & Data Lifecycle)
# =======================================================
    def _sync_seed_mode_state(self):
        """
        交互联动：根据“种子模式”的值，自动启用或禁用“种子值”文本框。
        """
        mode = str(self.combo_seed_mode.currentData() or "fixed").strip().lower()
        self.edit_seed.setEnabled(mode == "fixed")

    def _load_settings(self, settings: Dict[str, Any]):
        """
        数据反填逻辑：将传入的设置字典解析并应用到对应的 UI 控件中。
        利用 `findData` 方法进行安全查找，若未找到对应值则回退至默认索引 0。
        """
        # 1. 恢复随机数生成器 (RNG)
        try:
            rng_type = int(settings.get("rng_type", DEFAULT_RNG_TYPE))
        except Exception:
            rng_type = int(DEFAULT_RNG_TYPE)
        pos_rng = self.combo_rng_type.findData(rng_type)
        self.combo_rng_type.setCurrentIndex(pos_rng if pos_rng >= 0 else 0)

        # 2. 恢复采样方式
        sampling_mode = str(settings.get("sampling_mode", "MC")).strip().upper()
        pos_sampling = self.combo_sampling_mode.findData(sampling_mode)
        self.combo_sampling_mode.setCurrentIndex(pos_sampling if pos_sampling >= 0 else 0)

        # 3. 恢复种子模式
        seed_mode = str(settings.get("seed_mode", "fixed")).strip().lower()
        pos_seed_mode = self.combo_seed_mode.findData(seed_mode)
        self.combo_seed_mode.setCurrentIndex(pos_seed_mode if pos_seed_mode >= 0 else 0)

        # 4. 恢复种子具体数值
        seed_value = settings.get("seed_value", DEFAULT_SEED)
        self.edit_seed.setText(str(seed_value))

        # 5. 恢复收集数据范围
        sim_scope = str(settings.get("sim_scope", "all")).strip().lower()
        pos_scope = self.combo_scope.findData(sim_scope)
        self.combo_scope.setCurrentIndex(pos_scope if pos_scope >= 0 else 0)

        # 执行一次状态同步，确保 UI 组件激活状态正确
        self._sync_seed_mode_state()

    def _validate_and_accept(self):
        """
        表单提交流程与防呆校验：
        在用户点击“确定”时触发。若校验失败，将拦截提交并转移焦点。
        """
        text_seed = self.edit_seed.text().strip()
        mode = str(self.combo_seed_mode.currentData() or "fixed").strip().lower()
        
        # 校验规则：如果是固定种子模式，且输入框为空，则阻止关闭并提示用户输入
        if mode == "fixed" and not text_seed:
            self.edit_seed.setFocus()
            return
            
        # 校验通过，发出 accepted 信号并关闭对话框
        self.accept()

    def get_settings(self) -> Dict[str, Any]:
        """
        数据序列化：
        提取各控件当前的值，并封装为标准的字典对象供下游逻辑使用。
        """
        return {
            "sampling_mode": str(self.combo_sampling_mode.currentData() or "MC"),
            "rng_type": int(self.combo_rng_type.currentData() or int(DEFAULT_RNG_TYPE)),
            "seed_mode": str(self.combo_seed_mode.currentData() or "fixed"),
            "seed_value": self.edit_seed.text().strip() or str(int(DEFAULT_SEED)),
            "sim_scope": str(self.combo_scope.currentData() or "all"),
        }


# =======================================================
# 5. 全局调用接口 (Global Entry Point)
# =======================================================
def show_simulation_settings_dialog(
    current_settings: Optional[Dict[str, Any]] = None,
    parent=None,
) -> Optional[Dict[str, Any]]:
    """
    启动“模拟设置”对话框的快捷全局入口。
    
    内部实现了对 QApplication 生命周期的安全检查。
    如果用户点击了“确定”，则返回最新的配置字典；如果点击了“取消”或直接关闭窗口，则返回 None。
    
    :param current_settings: 初始设置字典。
    :param parent: 可选的父组件实例。
    :return: 包含最新设置的字典，若用户取消操作则返回 None。
    """
    # 确保 Qt 应用程序实例已启动，防呆保护
    app = QApplication.instance()
    if app is None:
        app = QApplication(sys.argv)
        _ = app
        
    dialog = SimulationSettingsDialog(current_settings=current_settings, parent=parent)
    
    # 阻塞式运行并判断返回值状态
    if dialog.exec() == QDialog.Accepted:
        return dialog.get_settings()
        
    return None