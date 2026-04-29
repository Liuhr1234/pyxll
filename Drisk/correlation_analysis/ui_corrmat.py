# -*- coding: utf-8 -*-
import sys
import os
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

import numpy as np
from pyxll import xl_app
from PySide6.QtCore import Qt
from PySide6.QtGui import QColor
from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit,
    QPushButton, QTableWidget, QTableWidgetItem, QHeaderView,
    QMessageBox, QWidget,
)
from ui_shared import apply_excel_select_button_icon, set_drisk_icon


def _build_stylesheet() -> str:
    return (
        "QDialog { background-color: #f9f9f9; font-family: 'Microsoft YaHei'; }\n"
        "QLabel { color: #444; font-size: 12px; }\n"
        "QLineEdit { font-size: 12px; }\n"
        "QTableWidget { font-size: 12px; gridline-color: #ddd; }\n"
        "QHeaderView::section { background-color: #f0f4f8; color: #334455; font-size: 12px; padding: 4px; border: 1px solid #ddd; }\n"
        "QPushButton {\n"
        "    background-color: white; border: 1px solid #ccc; border-radius: 3px;\n"
        "    padding: 4px 12px; font-size: 12px;\n"
        "}\n"
        "QPushButton:hover { color: #40a9ff; border-color: #40a9ff; background-color: #f0f7ff; }\n"
        "QPushButton#btnOk { background-color: #0050b3; color: white; border: none; min-width: 70px; }\n"
        "QPushButton#btnCheck { background-color: #389e0d; color: white; border: none; }\n"
        "QPushButton#btnCheck:hover { background-color: #52c41a; border-color: #52c41a; }\n"
    )


# ─────────────────────────────────────────────────────────────────────────────
# Step-1 dialog: pick cells
# ─────────────────────────────────────────────────────────────────────────────

class CellPickerDialog(QDialog):
    """第一步：让用户选择要建立相关性矩阵的单元格列表。"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Drisk - 相关性矩阵 · 选择单元格")
        self.resize(420, 200)
        self.cells: list[str] = []
        set_drisk_icon(self)
        self._build_ui()
        self.setStyleSheet(_build_stylesheet())

    def _build_ui(self):
        root = QVBoxLayout(self)
        root.setSpacing(10)
        root.setContentsMargins(14, 14, 14, 14)

        root.addWidget(QLabel("请选择需要建立相关性矩阵的单元格（可多选）："))

        row = QHBoxLayout()
        self._edit = QLineEdit()
        self._edit.setReadOnly(True)
        self._edit.setPlaceholderText("点击右侧按钮从 Excel 选择")
        self._edit.setStyleSheet("background-color: #f5f5f5; border: 1px solid #d9d9d9; padding: 4px;")
        btn_pick = QPushButton()
        btn_pick.setFixedSize(30, 26)
        btn_pick.setToolTip("从 Excel 选择单元格")
        if not apply_excel_select_button_icon(btn_pick, "select_icon.svg"):
            btn_pick.setText("\U0001F3AF")
        btn_pick.clicked.connect(self._pick)
        row.addWidget(self._edit)
        row.addWidget(btn_pick)
        root.addLayout(row)

        root.addStretch()

        btn_row = QHBoxLayout()
        btn_row.addStretch()
        btn_ok = QPushButton("下一步")
        btn_ok.setObjectName("btnOk")
        btn_ok.setFixedSize(80, 28)
        btn_cancel = QPushButton("取消")
        btn_cancel.setFixedSize(80, 28)
        btn_ok.clicked.connect(self._accept)
        btn_cancel.clicked.connect(self.reject)
        btn_row.addWidget(btn_ok)
        btn_row.addWidget(btn_cancel)
        root.addLayout(btn_row)

    def _pick(self):
        self.setWindowOpacity(0)
        try:
            app = xl_app()
            rng = app.InputBox(
                Prompt="请选择包含分布函数的单元格（可按 Ctrl 多选）:",
                Title="选择单元格",
                Type=8,
            )
            if rng:
                cells = []
                for area in rng.Areas:
                    for cell in area.Cells:
                        sheet = cell.Worksheet.Name
                        addr = cell.Address.replace("$", "")
                        cells.append(f"{sheet}!{addr}")
                self.cells = cells
                self._edit.setText(", ".join(cells))
        except Exception:
            pass
        finally:
            self.setWindowOpacity(1)
            self.raise_()
            self.activateWindow()

    def _accept(self):
        if not self.cells:
            QMessageBox.warning(self, "提示", "请先选择至少两个单元格。")
            return
        if len(self.cells) < 2:
            QMessageBox.warning(self, "提示", "相关性矩阵至少需要两个变量。")
            return
        self.accept()


# ─────────────────────────────────────────────────────────────────────────────
# Step-2 dialog: edit correlation matrix
# ─────────────────────────────────────────────────────────────────────────────

_DIAG_COLOR = QColor("#e8f0fe")
_UPPER_COLOR = QColor("#f5f5f5")
_INVALID_COLOR = QColor("#fff1f0")
_VALID_COLOR = QColor("#ffffff")


class CorrMatrixDialog(QDialog):
    """第二步：展示并编辑相关性矩阵（下三角可编辑，对角线固定为1，上三角镜像只读）。"""

    def __init__(self, cells: list[str], parent=None):
        super().__init__(parent)
        self.setWindowTitle("Drisk - 相关性矩阵编辑器")
        self.cells = cells
        n = len(cells)
        self._n = n
        # 初始化为单位矩阵
        self._matrix = np.eye(n)
        set_drisk_icon(self)
        self._build_ui()
        self.setStyleSheet(_build_stylesheet())
        self.resize(max(500, 120 + n * 80), max(420, 200 + n * 40))

    def _build_ui(self):
        root = QVBoxLayout(self)
        root.setSpacing(10)
        root.setContentsMargins(14, 14, 14, 14)

        # 矩阵名称行
        name_row = QHBoxLayout()
        name_row.addWidget(QLabel("矩阵名称："))
        self._name_edit = QLineEdit()
        self._name_edit.setPlaceholderText("例如：相关矩阵_1")
        self._name_edit.setText(self._suggest_name())
        name_row.addWidget(self._name_edit)
        root.addLayout(name_row)

        # 说明
        hint = QLabel("对角线固定为 1；下三角可编辑（−1 到 1）；上三角自动镜像。")
        hint.setStyleSheet("color: #888; font-size: 11px;")
        root.addWidget(hint)

        # 矩阵表格
        n = self._n
        self._table = QTableWidget(n, n)
        labels = [c.split("!")[-1] if "!" in c else c for c in self.cells]
        self._table.setHorizontalHeaderLabels(labels)
        self._table.setVerticalHeaderLabels(labels)
        self._table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self._table.verticalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        self._populate_table()
        self._table.itemChanged.connect(self._on_item_changed)
        root.addWidget(self._table, stretch=1)

        # 按钮行
        btn_row = QHBoxLayout()
        btn_check = QPushButton("验证半正定")
        btn_check.setObjectName("btnCheck")
        btn_check.clicked.connect(self._check_psd)
        btn_row.addWidget(btn_check)
        btn_row.addStretch()
        btn_ok = QPushButton("确定")
        btn_ok.setObjectName("btnOk")
        btn_ok.setFixedSize(80, 28)
        btn_cancel = QPushButton("取消")
        btn_cancel.setFixedSize(80, 28)
        btn_ok.clicked.connect(self._accept)
        btn_cancel.clicked.connect(self.reject)
        btn_row.addWidget(btn_ok)
        btn_row.addWidget(btn_cancel)
        root.addLayout(btn_row)

    def _suggest_name(self) -> str:
        try:
            import re
            app = xl_app()
            wb = app.ActiveWorkbook
            max_num = 0
            for name_obj in wb.Names:
                m = re.match(r'^相关矩阵_(\d+)$', name_obj.Name)
                if m:
                    max_num = max(max_num, int(m.group(1)))
            return f"相关矩阵_{max_num + 1}"
        except Exception:
            return "相关矩阵_1"

    def _populate_table(self):
        n = self._n
        self._table.blockSignals(True)
        for r in range(n):
            for c in range(n):
                val = self._matrix[r, c]
                item = QTableWidgetItem(f"{val:.4f}" if r != c else "1")
                if r == c:
                    item.setFlags(item.flags() & ~Qt.ItemIsEditable)
                    item.setBackground(_DIAG_COLOR)
                elif c > r:
                    item.setFlags(item.flags() & ~Qt.ItemIsEditable)
                    item.setBackground(_UPPER_COLOR)
                    item.setForeground(QColor("#999999"))
                else:
                    item.setBackground(_VALID_COLOR)
                self._table.setItem(r, c, item)
        self._table.blockSignals(False)

    def _on_item_changed(self, item: QTableWidgetItem):
        r, c = item.row(), item.column()
        if c >= r:
            return
        text = item.text().strip()
        try:
            val = float(text)
            if not (-1.0 <= val <= 1.0):
                raise ValueError
            self._matrix[r, c] = val
            self._matrix[c, r] = val
            item.setBackground(_VALID_COLOR)
            # 同步上三角镜像
            self._table.blockSignals(True)
            mirror = self._table.item(c, r)
            if mirror:
                mirror.setText(f"{val:.4f}")
            self._table.blockSignals(False)
        except ValueError:
            item.setBackground(_INVALID_COLOR)

    def _check_psd(self):
        from corrmat_functions import _is_positive_semidefinite, _nearest_psd
        if _is_positive_semidefinite(self._matrix):
            QMessageBox.information(self, "验证结果", "矩阵是半正定的，有效。")
        else:
            reply = QMessageBox.question(
                self, "验证结果",
                "矩阵不是半正定的。\n\n是否自动调整为最近的半正定矩阵？",
                QMessageBox.Yes | QMessageBox.No,
            )
            if reply == QMessageBox.Yes:
                self._matrix = _nearest_psd(self._matrix)
                self._populate_table()
                QMessageBox.information(self, "已调整", "矩阵已调整为最近的半正定矩阵。")

    def get_matrix_name(self) -> str:
        return self._name_edit.text().strip() or self._suggest_name()

    def get_matrix(self) -> np.ndarray:
        return self._matrix.copy()

    def _accept(self):
        from corrmat_functions import _is_positive_semidefinite
        # 检查是否有无效输入（红色单元格）
        n = self._n
        for r in range(n):
            for c in range(r):
                item = self._table.item(r, c)
                if item and item.background().color() == _INVALID_COLOR:
                    QMessageBox.warning(self, "输入错误", f"单元格 ({r+1},{c+1}) 的值无效，请修正后再确定。")
                    return
        if not _is_positive_semidefinite(self._matrix):
            reply = QMessageBox.question(
                self, "矩阵非半正定",
                "当前矩阵不是半正定的，可能导致模拟结果异常。\n\n是否仍然继续？",
                QMessageBox.Yes | QMessageBox.No,
            )
            if reply == QMessageBox.No:
                return
        self.accept()
