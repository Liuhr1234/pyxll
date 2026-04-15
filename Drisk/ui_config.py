# ui_config.py
"""
统一模拟配置面板 UI
"""
from PySide6.QtWidgets import (QApplication, QDialog, QVBoxLayout, QHBoxLayout, QCheckBox, 
                               QLabel, QLineEdit, QPushButton, QMessageBox, QGroupBox, QFormLayout)
from PySide6.QtCore import Qt
from PySide6.QtGui import QIntValidator

class SimulationConfigDialog(QDialog):
    # 【修改点 1】：增加 scan_callback 参数
    def __init__(self, excel_app, scan_callback, default_iter=5000, parent=None):
        super().__init__(parent)
        self.excel_app = excel_app
        self.scan_callback = scan_callback  # 保存回调函数，避免直接依赖 sim_engine
        self.scanned_data = None
        self.cell_scenario_limits = {}
        
        self.setWindowTitle("Drisk - 运行模拟配置")
        self.setMinimumWidth(400)
        self.setWindowFlags(self.windowFlags() | Qt.WindowStaysOnTopHint)

        layout = QVBoxLayout(self)

        group_box = QGroupBox("参数设置")
        form_layout = QFormLayout(group_box)

        self.iter_input = QLineEdit(str(default_iter))
        self.iter_input.setValidator(QIntValidator(1, 10000000))
        form_layout.addRow("模拟迭代次数:", self.iter_input)

        scenario_layout = QHBoxLayout()
        self.scenario_input = QLineEdit("1")
        self.scenario_input.setValidator(QIntValidator(1, 10000))
        
        self.detect_btn = QPushButton("✨ 智能探测")
        self.detect_btn.setToolTip("深度扫描工作簿，自动计算场景数")
        self.detect_btn.clicked.connect(self.run_smart_detect)
        
        scenario_layout.addWidget(self.scenario_input)
        scenario_layout.addWidget(self.detect_btn)
        form_layout.addRow("模拟场景个数:", scenario_layout)

        layout.addWidget(group_box)
        
        self.independent_cb = QCheckBox("🎯 独立缓存模式 (仅更新选中单元格，保留其他历史缓存)")
        self.independent_cb.setChecked(True)
        self.independent_cb.setStyleSheet("color: #0050b3; font-weight: bold; margin-bottom: 5px;")
        layout.addWidget(self.independent_cb)

        self.status_label = QLabel("提示：若您确定场景数，可直接点击运行，跳过扫描避免卡顿。")
        self.status_label.setStyleSheet("color: #666666; font-size: 12px; margin-top: 5px;")
        layout.addWidget(self.status_label)

        btn_layout = QHBoxLayout()
        btn_layout.addStretch()
        self.cancel_btn = QPushButton("取消")
        self.run_btn = QPushButton("🚀 开始运行")
        self.run_btn.setStyleSheet("background-color: #2196F3; color: white; font-weight: bold; padding: 5px 15px;")
        self.cancel_btn.clicked.connect(self.reject)
        self.run_btn.clicked.connect(self.accept)
        btn_layout.addWidget(self.cancel_btn)
        btn_layout.addWidget(self.run_btn)
        layout.addLayout(btn_layout)

    def run_smart_detect(self):
        self.detect_btn.setEnabled(False)
        self.run_btn.setEnabled(False)
        self.status_label.setStyleSheet("color: #E65100; font-weight: bold;")
        self.status_label.setText("正在深度解析模型依赖网络并清理幽灵数据，请稍候...")
        QApplication.processEvents()

        try:
            # 【修改点 2】：调用传入的 callback，而不是写死的 _perform_model_scan
            cells_data, limits, max_deps = self.scan_callback(self.excel_app, deep_scan=True)
            
            self.scanned_data = cells_data
            self.cell_scenario_limits = limits
            self.scenario_input.setText(str(max_deps))
            self.status_label.setStyleSheet("color: #2E7D32; font-weight: bold;")
            self.status_label.setText(f"扫描完成！精准探测到 {max_deps} 个模拟场景。")
        except Exception as e:
            self.status_label.setStyleSheet("color: #D32F2F;")
            self.status_label.setText("扫描模型结构时发生异常，请手动输入。")
            QMessageBox.warning(self, "探测错误", f"无法解析依赖关系:\n{e}")
        finally:
            self.detect_btn.setEnabled(True)
            self.run_btn.setEnabled(True)
    
    def get_values(self):
        return int(self.iter_input.text()), int(self.scenario_input.text()), self.independent_cb.isChecked()