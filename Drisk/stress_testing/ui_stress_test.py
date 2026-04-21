# -*- coding: utf-8 -*-
from pyxll import xl_app
from PySide6.QtCore import Qt, Signal
from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QFormLayout, QGroupBox,
    QLabel, QLineEdit, QPushButton, QTableWidget, QTableWidgetItem,
    QComboBox, QDoubleSpinBox, QHeaderView, QMessageBox, QWidget
)
from ui_shared import DRISK_DIALOG_BTN_QSS

#压力测试输入配置窗口：选择自变量（X）、因变量（Y）及各自变量的压力范围。
class StressTestInputDialog(QDialog):
    """压力测试输入配置窗口：选择自变量（X）、因变量（Y）及各自变量的压力范围。"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("压力测试 - 配置")
        self.setMinimumWidth(560)
        self.x_cells: list[str] = []   # ["Sheet1!A1", ...]
        self.y_cell: str = ""
        self._build_ui()
    # ------------------------------------------------------------------
    def _build_ui(self):
        root = QVBoxLayout(self)
        root.setSpacing(10)

        # ---- Y 因变量 ----
        grp_y = QGroupBox("因变量（Y）")
        ly = QHBoxLayout(grp_y)
        self._y_edit = QLineEdit()
        self._y_edit.setReadOnly(True)
        self._y_edit.setPlaceholderText("点击右侧按钮在 Excel 中选择单元格")
        btn_y = QPushButton("选择…")
        btn_y.setFixedWidth(60)
        btn_y.clicked.connect(self._pick_y)
        ly.addWidget(self._y_edit)
        ly.addWidget(btn_y)
        root.addWidget(grp_y)

        # ---- X 自变量 ----
        grp_x = QGroupBox("自变量（X，可多选）")
        lx = QHBoxLayout(grp_x)
        self._x_edit = QLineEdit()
        self._x_edit.setReadOnly(True)
        self._x_edit.setPlaceholderText("点击右侧按钮在 Excel 中选择单元格（可多选）")
        btn_x = QPushButton("选择…")
        btn_x.setFixedWidth(60)
        btn_x.clicked.connect(self._pick_x)
        lx.addWidget(self._x_edit)
        lx.addWidget(btn_x)
        root.addWidget(grp_x)

        # ---- 压力范围表格 ----
        grp_range = QGroupBox("各自变量压力测试范围")
        lrange = QVBoxLayout(grp_range)
        self._table = QTableWidget(0, 4)
        self._table.setHorizontalHeaderLabels(["单元格", "类型", "下限", "上限"])
        self._table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        self._table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        self._table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeToContents)
        self._table.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeToContents)
        self._table.setEditTriggers(QTableWidget.NoEditTriggers)
        lrange.addWidget(self._table)
        root.addWidget(grp_range)

        # ---- 按钮 ----
        btn_row = QHBoxLayout()
        btn_row.addStretch()
        btn_ok = QPushButton("确定")
        btn_ok.setObjectName("btnOk")
        btn_cancel = QPushButton("取消")
        btn_cancel.setObjectName("btnCancel")
        btn_ok.clicked.connect(self._accept)
        btn_cancel.clicked.connect(self.reject)
        btn_row.addWidget(btn_ok)
        btn_row.addWidget(btn_cancel)
        root.addLayout(btn_row)

        self.setStyleSheet(DRISK_DIALOG_BTN_QSS)

    # ------------------------------------------------------------------
    def _pick_range(self, single: bool) -> list[str]:
        """弹出 Excel InputBox 让用户框选单元格，返回地址列表。"""
        self.setWindowOpacity(0)
        cells = []
        try:
            app = xl_app()
            prompt = "请选择因变量单元格（单个）:" if single else "请选择自变量单元格（可按 Ctrl 多选）:"
            rng = app.InputBox(Prompt=prompt, Title="选择单元格", Type=8)
            if rng:
                for area in rng.Areas:
                    for cell in area.Cells:
                        sheet = cell.Worksheet.Name
                        addr = cell.Address.replace("$", "")
                        cells.append(f"{sheet}!{addr}")
                        if single:
                            break
                    if single and cells:
                        break
        except Exception:
            pass
        finally:
            self.setWindowOpacity(1)
            self.raise_()
            self.activateWindow()
        return cells

    def _pick_y(self):
        cells = self._pick_range(single=True)
        if cells:
            self.y_cell = cells[0]
            self._y_edit.setText(self.y_cell)

    def _pick_x(self):
        cells = self._pick_range(single=False)
        if cells:
            self.x_cells = cells
            self._x_edit.setText(", ".join(cells))
            self._rebuild_table()

    # ------------------------------------------------------------------
    def _rebuild_table(self):
        """根据当前 x_cells 重建压力范围表格，保留已有设置。"""
        old = {}
        for r in range(self._table.rowCount()):
            addr = self._table.item(r, 0).text()
            combo = self._table.cellWidget(r, 1)
            lo = self._table.cellWidget(r, 2)
            hi = self._table.cellWidget(r, 3)
            if combo and lo and hi:
                old[addr] = (combo.currentText(), lo.value(), hi.value())

        self._table.setRowCount(len(self.x_cells))
        for r, addr in enumerate(self.x_cells):
            self._table.setItem(r, 0, QTableWidgetItem(addr))

            combo = QComboBox()
            combo.addItems(["百分比 (%)", "数值"])
            prev_type, prev_lo, prev_hi = old.get(addr, ("百分比 (%)", -20.0, 20.0))
            combo.setCurrentText(prev_type)
            self._table.setCellWidget(r, 1, combo)

            lo_spin = QDoubleSpinBox()
            lo_spin.setRange(-1e9, 1e9)
            lo_spin.setDecimals(2)
            lo_spin.setValue(prev_lo)
            self._table.setCellWidget(r, 2, lo_spin)

            hi_spin = QDoubleSpinBox()
            hi_spin.setRange(-1e9, 1e9)
            hi_spin.setDecimals(2)
            hi_spin.setValue(prev_hi)
            self._table.setCellWidget(r, 3, hi_spin)

    # ------------------------------------------------------------------
    def get_config(self) -> dict:
        """返回用户配置：y_cell, x_configs=[{addr, type, lo, hi}]"""
        x_configs = []
        for r in range(self._table.rowCount()):
            addr = self._table.item(r, 0).text()
            combo = self._table.cellWidget(r, 1)
            lo = self._table.cellWidget(r, 2)
            hi = self._table.cellWidget(r, 3)
            x_configs.append({
                "addr": addr,
                "type": combo.currentText(),
                "lo": lo.value(),
                "hi": hi.value(),
            })
        return {"y_cell": self.y_cell, "x_configs": x_configs}

    def _accept(self):
        if not self.y_cell:
            QMessageBox.warning(self, "提示", "请选择因变量（Y）单元格。")
            return
        if not self.x_cells:
            QMessageBox.warning(self, "提示", "请至少选择一个自变量（X）单元格。")
            return
        self.accept()
