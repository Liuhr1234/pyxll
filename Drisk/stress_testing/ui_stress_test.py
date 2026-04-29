# -*- coding: utf-8 -*-
import sys
import os
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
import sim_engine
from pyxll import xl_app
from PySide6.QtCore import Qt
from PySide6.QtGui import QColor
from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout,
    QLabel, QLineEdit, QPushButton, QTableWidget, QTableWidgetItem,
    QComboBox, QDoubleSpinBox, QHeaderView, QMessageBox,
    QListWidget, QListWidgetItem, QSplitter, QWidget, QAbstractItemView,
    QButtonGroup, QRadioButton,
)
from ui_shared import apply_excel_select_button_icon, set_drisk_icon


def _get_output_candidates() -> list[tuple[str, str]]:
    """返回内存中所有模拟缓存的 output 变量列表，每项为 (display_label, full_addr)。"""
    try:
        from simulation_manager import _SIMULATION_CACHE
        from sim_engine import drisk_run_selected
        seen = set()
        candidates = []
        for sim in _SIMULATION_CACHE.values():
            for ck in (sim.output_cache or {}):
                if ck in seen:
                    continue
                seen.add(ck)
                attrs = (sim.output_attributes or {}).get(ck, {}) or {}
                name = attrs.get("name") or attrs.get("DriskName") or ""
                label = f"{name} ({ck})" if name else ck
                candidates.append((label, ck))
        #若没有候选值，则运行一次模拟
        if candidates == []:
            drisk_run_selected()
            for sim in _SIMULATION_CACHE.values():
                for ck in (sim.output_cache or {}):
                    if ck in seen:
                        continue
                    seen.add(ck)
                    attrs = (sim.output_attributes or {}).get(ck, {}) or {}
                    name = attrs.get("name") or attrs.get("DriskName") or ""
                    label = f"{name} ({ck})" if name else ck
                    candidates.append((label, ck))
        return candidates
    except Exception:
        return []


def _get_input_candidates_for_y(y_addr: str) -> list[tuple[str, str]]:
    """
    返回与 Y 地址相关的输入变量候选项（每项为 (display_label, full_addr)）：
    1. Y 单元格的 Excel 公式直接前驱；
    """
    from ui_shared import resolve_visible_variable_name
    candidates = []
    seen = set()

    # 预先从最新模拟缓存中构建 addr -> attrs 映射
    attrs_map: dict[str, dict] = {}
    try:
        from simulation_manager import _SIMULATION_CACHE
        if _SIMULATION_CACHE:
            sim = _SIMULATION_CACHE[max(_SIMULATION_CACHE.keys())]
            for ck in (sim.input_cache or {}):
                attrs = sim.get_input_attributes(ck)
                if attrs:
                    # 标准化：去 $ 、去数字后缀（如 _1、_MAKEINPUT）
                    base = ck.replace("$", "").upper()
                    cell_part = base.split("!")[-1]
                    if "_" in cell_part:
                        base = base.rsplit("_", 1)[0]
                    attrs_map[base] = attrs
    except Exception:
        pass

    # Excel 公式前驱
    try:
        app = xl_app()
        if "!" in y_addr:
            sheet_name, addr = y_addr.split("!", 1)
        else:
            sheet_name = app.ActiveSheet.Name
            addr = y_addr
        ws = app.Sheets(sheet_name)
        rng = ws.Range(addr)
        try:
            for area in rng.DirectPrecedents.Areas:
                for cell in area.Cells:
                    cell_sheet = cell.Worksheet.Name
                    cell_addr = cell.Address.replace("$", "")
                    full = f"{cell_sheet}!{cell_addr}"
                    if full not in seen:
                        seen.add(full)
                        attrs = attrs_map.get(full.upper(), {})
                        name = resolve_visible_variable_name(
                            full, attrs, excel_app=app, fallback_label=full
                        )
                        label = f"{name} ({full})" if name != full else full
                        candidates.append((label, full))
        except Exception:
            pass
    except Exception:
        pass

    return candidates


def _build_dialog_stylesheet() -> str:
    """与 AdvancedOverlayDialog 对齐的整体样式表。"""
    base_dir = os.path.dirname(os.path.abspath(os.path.join(__file__, "..")))
    icon_path = os.path.join(base_dir, "icons", "Selected_Overlay.svg").replace("\\", "/")
    return (
        "QDialog { background-color: #f9f9f9; font-family: 'Microsoft YaHei'; }\n"
        "QLabel { color: #444; font-size: 12px; }\n"
        "QLineEdit { font-size: 12px; }\n"
        "QListWidget {\n"
        "    border: 1px solid #ccc; border-radius: 3px; outline: none;\n"
        "    font-size: 12px; background-color: #ffffff;\n"
        "}\n"
        "QListWidget::item { padding: 6px; border-bottom: 1px solid #f5f5f5; }\n"
        "QListWidget::item:hover { background-color: #f0f7ff; }\n"
        "QListWidget::item:selected {\n"
        "    background-color: #f0f4f8; color: #334455;\n"
        "    font-weight: bold; border-left: 3px solid #8c9eb5;\n"
        "}\n"
        "QPushButton {\n"
        "    background-color: white; border: 1px solid #ccc; border-radius: 3px;\n"
        "    padding: 4px 12px; font-size: 12px;\n"
        "}\n"
        "QPushButton:hover { color: #40a9ff; border-color: #40a9ff; background-color: #f0f7ff; }\n"
        "QPushButton#btnOk { background-color: #0050b3; color: white; border: none; min-width: 70px; }\n"
        "QPushButton#btnRemoveAll:hover { color: #ff4d4f; border-color: #ff4d4f; background-color: #fff1f0; }\n"
        "QRadioButton { font-size: 12px; color: #333; }\n"
        "QRadioButton::indicator {\n"
        "    width: 16px; height: 16px; border-radius: 8px;\n"
        "    border: 1px solid #cbd0d6; background-color: #ffffff;\n"
        "}\n"
        "QRadioButton::indicator:hover { border: 1px solid #8c9eb5; }\n"
        "QRadioButton::indicator:checked { background-color: #8c9eb5; border: 1px solid #8c9eb5; }\n"
        "QListWidget::indicator {\n"
        "    width: 16px; height: 16px; border-radius: 3px;\n"
        "    border: 1px solid #cbd0d6; background-color: #ffffff;\n"
        "}\n"
        "QListWidget::indicator:hover { border: 1px solid #8c9eb5; }\n"
        f"QListWidget::indicator:checked {{ background-color: #8c9eb5; border: 1px solid #8c9eb5; image: url('{icon_path}'); }}\n"
    )


class StressTestInputDialog(QDialog):
    """压力测试配置对话框：左侧选择输出变量 Y，右侧配置输入变量 X 的压力范围。"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Drisk - 压力测试设置")
        self.resize(600, 660)
        self.x_cells: list[str] = []
        self.y_cell: str = ""
        # 地址 -> 显示名称映射
        self._addr_to_name: dict[str, str] = {}
        set_drisk_icon(self)
        self._build_ui()
        self.setStyleSheet(_build_dialog_stylesheet())

    # ------------------------------------------------------------------
    def _build_ui(self):
        root = QVBoxLayout(self)
        root.setSpacing(10)
        root.setContentsMargins(12, 12, 12, 12)

        # 说明
        lbl_info = QLabel("选择要监控的输出变量，在右侧添加要进行测试的输入变量，并在下方设置其测试范围。")
        lbl_info.setStyleSheet("color: #333333; font-size: 13px;")
        root.addWidget(lbl_info)

        # ===== 左右主体：QSplitter =====
        splitter = QSplitter(Qt.Horizontal)
        splitter.setHandleWidth(6)

        # ---------- 左侧面板：输出变量 Y ----------
        left_widget = QWidget()
        left_layout = QVBoxLayout(left_widget)
        left_layout.setSpacing(6)
        left_layout.setContentsMargins(0, 0, 6, 0)

        lbl_y_title = QLabel("1. 选择要监控的输出变量：")
        lbl_y_title.setStyleSheet("font-weight: bold; font-size: 12px; color: #444;")
        left_layout.addWidget(lbl_y_title)

        y_row = QHBoxLayout()
        self._y_edit = QLineEdit()
        self._y_edit.setReadOnly(True)
        self._y_edit.setPlaceholderText("从 Excel 选择或从下方列表双击")
        self._y_edit.setStyleSheet("background-color: #f5f5f5; border: 1px solid #d9d9d9; padding: 4px;")
        btn_y = QPushButton()
        btn_y.setFixedSize(30, 26)
        btn_y.setToolTip("从 Excel 选择输出变量单元格")
        if not apply_excel_select_button_icon(btn_y, "select_icon.svg"):
            btn_y.setText("\U0001F3AF")
        btn_y.clicked.connect(self._pick_y)
        y_row.addWidget(self._y_edit)
        y_row.addWidget(btn_y)
        left_layout.addLayout(y_row)

        lbl_cache = QLabel("输出变量列表（双击选择）：")
        lbl_cache.setStyleSheet("color: #666; font-size: 11px; font-weight: normal;")
        left_layout.addWidget(lbl_cache)

        self._list_y_candidates = QListWidget()
        self._list_y_candidates.setSelectionMode(QAbstractItemView.SingleSelection)
        self._list_y_candidates.itemDoubleClicked.connect(self._on_y_candidate_double_clicked)
        left_layout.addWidget(self._list_y_candidates, stretch=1)

        btn_refresh = QPushButton("刷新列表")
        btn_refresh.clicked.connect(self._refresh_y_candidates)
        left_layout.addWidget(btn_refresh)

        splitter.addWidget(left_widget)

        # ---------- 右侧面板：输入变量 X ----------
        right_widget = QWidget()
        right_layout = QVBoxLayout(right_widget)
        right_layout.setSpacing(6)
        right_layout.setContentsMargins(6, 0, 0, 0)

        lbl_x_title = QLabel("2. 勾选需要测试的输入变量：")
        lbl_x_title.setStyleSheet("font-weight: bold; font-size: 12px; color: #444;")
        right_layout.addWidget(lbl_x_title)

        x_row = QHBoxLayout()
        self._x_edit = QLineEdit()
        self._x_edit.setReadOnly(True)
        self._x_edit.setPlaceholderText("从 Excel 选择或从下方候选列表双击添加")
        self._x_edit.setStyleSheet("background-color: #f5f5f5; border: 1px solid #d9d9d9; padding: 4px;")
        btn_x = QPushButton()
        btn_x.setFixedSize(30, 26)
        btn_x.setToolTip("从 Excel 选择输入变量单元格（可多选）")
        if not apply_excel_select_button_icon(btn_x, "select_icon.svg"):
            btn_x.setText("\U0001F3AF")
        btn_x.clicked.connect(self._pick_x)
        x_row.addWidget(self._x_edit)
        x_row.addWidget(btn_x)
        right_layout.addLayout(x_row)

        lbl_x_cand = QLabel("候选输入变量列表（双击选择或勾选添加）：")
        lbl_x_cand.setStyleSheet("color: #666; font-size: 11px; font-weight: normal;")
        right_layout.addWidget(lbl_x_cand)

        self._list_x_candidates = QListWidget()
        self._list_x_candidates.setSelectionMode(QAbstractItemView.SingleSelection)
        self._list_x_candidates.itemDoubleClicked.connect(self._on_x_candidate_double_clicked)
        right_layout.addWidget(self._list_x_candidates, stretch=1)

        splitter.addWidget(right_widget)
        splitter.setSizes([300, 300])
        root.addWidget(splitter, stretch=1)

        btn_add_checked = QPushButton("↓ 添加到测试输入变量列表 ↓")
        btn_add_checked.clicked.connect(self._add_checked_x_candidates)
        root.addWidget(btn_add_checked)
        # ---- 压力范围表格标题行（含"移除所有项"按钮）----
        table_hdr = QHBoxLayout()
        lbl_table = QLabel("3.输入变量压力测试范围：")
        lbl_table.setStyleSheet("font-weight: bold; font-size: 12px; color: #444;")
        table_hdr.addWidget(lbl_table)
        table_hdr.addStretch()
        
        root.addLayout(table_hdr)

        self._table = QTableWidget(0, 5)
        self._table.setHorizontalHeaderLabels(["输入变量", "类型", "下限", "上限", ""])
        header = self._table.horizontalHeader()
        for col in range(4):
            header.setSectionResizeMode(col, QHeaderView.Stretch)
        header.setSectionResizeMode(4, QHeaderView.Fixed)
        header.resizeSection(4, 50)
        self._table.setEditTriggers(QTableWidget.NoEditTriggers)
        self._table.setFixedHeight(150)
        root.addWidget(self._table)
        btn_remove_all = QPushButton("移除所有项")
        btn_remove_all.setObjectName("btnRemoveAll")
        btn_remove_all.clicked.connect(self._remove_all_rows)
        root.addWidget(btn_remove_all)
        # ---- 分析模式：两个互斥单选项 ----
        lbl_mode = QLabel("压力测试场景生成方式：")
        root.addWidget(lbl_mode)

        self._radio_group = QButtonGroup(self)
        self._radio_single = QRadioButton("依次针对单个输入变量进行压力测试模拟")
        self._radio_all = QRadioButton("一次针对所有输入变量进行压力测试模拟")
        self._radio_all.setChecked(True)
        self._radio_group.addButton(self._radio_single, 0)
        self._radio_group.addButton(self._radio_all, 1)

        radio_layout = QHBoxLayout()
        radio_layout.setSpacing(20)
        radio_layout.addWidget(self._radio_all)
        radio_layout.addWidget(self._radio_single)
        radio_layout.addStretch()
        root.addLayout(radio_layout)

        # ---- 确定 / 取消 ----
        btn_row = QHBoxLayout()
        btn_row.addStretch()
        btn_ok = QPushButton("确定")
        btn_ok.setObjectName("btnOk")
        btn_ok.setFixedSize(80, 28)
        btn_cancel = QPushButton("取消")
        btn_cancel.setObjectName("btnCancel")
        btn_cancel.setFixedSize(80, 28)
        btn_ok.clicked.connect(self._accept)
        btn_cancel.clicked.connect(self.reject)
        btn_row.addWidget(btn_ok)
        btn_row.addWidget(btn_cancel)
        root.addLayout(btn_row)

        self._refresh_y_candidates()

    # ------------------------------------------------------------------
    # Y 候选列表
    # ------------------------------------------------------------------
    def _refresh_y_candidates(self):
        self._list_y_candidates.clear()
        for label, addr in _get_output_candidates():
            # 提取纯名称（去掉括号中的地址）
            if " (" in label and label.endswith(")"):
                name = label.rsplit(" (", 1)[0]
            else:
                name = label
            self._addr_to_name[addr] = name
            item = QListWidgetItem(label)
            item.setData(Qt.UserRole, addr)
            self._list_y_candidates.addItem(item)

    def _on_y_candidate_double_clicked(self, item: QListWidgetItem):
        addr = item.data(Qt.UserRole)
        self.y_cell = addr
        name = self._addr_to_name.get(addr, addr)
        self._y_edit.setText(name)
        self._refresh_x_candidates()

    # ------------------------------------------------------------------
    # X 候选列表
    # ------------------------------------------------------------------
    def _refresh_x_candidates(self):
        self._list_x_candidates.clear()
        if not self.y_cell:
            return
        selected_set = set(self.x_cells)
        self._list_x_candidates.blockSignals(True)
        for label, addr in _get_input_candidates_for_y(self.y_cell):
            # 提取纯名称（去掉括号中的地址）
            if " (" in label and label.endswith(")"):
                name = label.rsplit(" (", 1)[0]
            else:
                name = label
            self._addr_to_name[addr] = name
            item = QListWidgetItem(label)
            item.setData(Qt.UserRole, addr)
            item.setFlags(item.flags() | Qt.ItemIsUserCheckable)
            if addr in selected_set:
                item.setCheckState(Qt.Checked)
                item.setForeground(QColor("#aaaaaa"))
            else:
                item.setCheckState(Qt.Unchecked)
                item.setForeground(QColor("#334455"))
            self._list_x_candidates.addItem(item)
        self._list_x_candidates.blockSignals(False)

    def _on_x_candidate_double_clicked(self, item: QListWidgetItem):
        self._add_x_cells([item.data(Qt.UserRole)])


    def _add_checked_x_candidates(self):
        """将候选列表中所有勾选项批量加入压力测试列表。"""
        checked = []
        for i in range(self._list_x_candidates.count()):
            item = self._list_x_candidates.item(i)
            if item.checkState() == Qt.Checked:
                checked.append(item.data(Qt.UserRole))
        if checked:
            self._add_x_cells(checked)

    # ------------------------------------------------------------------
    # 手动选择 Y / X（从 Excel InputBox）
    # ------------------------------------------------------------------
    def _pick_range(self, single: bool) -> list[str]:
        self.setWindowOpacity(0)
        cells = []
        try:
            app = xl_app()
            prompt = "请选择输出变量单元格（单个）:" if single else "请选择输入变量单元格（可按 Ctrl 多选）:"
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
            addr = cells[0]
            self.y_cell = addr
            # 解析名称并存储
            from ui_shared import resolve_visible_variable_name
            try:
                app = xl_app()
                name = resolve_visible_variable_name(addr, excel_app=app, fallback_label=addr)
                self._addr_to_name[addr] = name
                self._y_edit.setText(name)
            except Exception:
                self._y_edit.setText(addr)
            self._refresh_x_candidates()

    def _pick_x(self):
        cells = self._pick_range(single=False)
        if cells:
            # 解析名称并存储
            from ui_shared import resolve_visible_variable_name
            try:
                app = xl_app()
                for addr in cells:
                    if addr not in self._addr_to_name:
                        name = resolve_visible_variable_name(addr, excel_app=app, fallback_label=addr)
                        self._addr_to_name[addr] = name
            except Exception:
                pass
            self._add_x_cells(cells)

    def _add_x_cells(self, cells: list[str]):
        existing = set(self.x_cells)
        for c in cells:
            if c not in existing:
                self.x_cells.append(c)
                existing.add(c)
        self._x_edit.setText(", ".join(self.x_cells))
        self._rebuild_table()
        self._sync_candidate_check_states()

    # ------------------------------------------------------------------
    # 压力范围表格
    # ------------------------------------------------------------------
    def _rebuild_table(self):
        old = {}
        for r in range(self._table.rowCount()):
            item = self._table.item(r, 0)
            addr = item.data(Qt.UserRole) if item else None
            combo = self._table.cellWidget(r, 1)
            lo = self._table.cellWidget(r, 2)
            hi = self._table.cellWidget(r, 3)
            if addr and combo and lo and hi:
                old[addr] = (combo.currentText(), lo.value(), hi.value())

        self._table.setRowCount(len(self.x_cells))
        for r, addr in enumerate(self.x_cells):
            name = self._addr_to_name.get(addr, addr)+f"({addr})"
            item = QTableWidgetItem(name)
            item.setData(Qt.UserRole, addr)
            self._table.setItem(r, 0, item)

            combo = QComboBox()
            combo.addItems(["百分比 (%)", "数值"])
            prev_type, prev_lo, prev_hi = old.get(addr, ("百分比 (%)", 0.0, 100.0))
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

            btn_del = QPushButton("删除")
            btn_del.setFixedWidth(50)
            btn_del.clicked.connect(lambda _, row=r: self._delete_row(row))
            self._table.setCellWidget(r, 4, btn_del)

    def _delete_row(self, row: int):
        if 0 <= row < len(self.x_cells):
            del self.x_cells[row]
            self._x_edit.setText(", ".join(self.x_cells))
            self._rebuild_table()
            self._sync_candidate_check_states()

    def _remove_all_rows(self):
        self.x_cells.clear()
        self._x_edit.clear()
        self._table.setRowCount(0)
        self._sync_candidate_check_states()

    def _sync_candidate_check_states(self):
        """根据 x_cells 同步候选列表的勾选状态和文字颜色。"""
        selected_set = set(self.x_cells)
        self._list_x_candidates.blockSignals(True)
        for i in range(self._list_x_candidates.count()):
            item = self._list_x_candidates.item(i)
            addr = item.data(Qt.UserRole)
            if addr in selected_set:
                item.setCheckState(Qt.Checked)
                item.setForeground(QColor("#aaaaaa"))
            else:
                item.setCheckState(Qt.Unchecked)
                item.setForeground(QColor("#334455"))
        self._list_x_candidates.blockSignals(False)

    # ------------------------------------------------------------------
    def get_config(self) -> dict:
        x_configs = []
        for r in range(self._table.rowCount()):
            item = self._table.item(r, 0)
            addr = item.data(Qt.UserRole) if item else None
            combo = self._table.cellWidget(r, 1)
            lo = self._table.cellWidget(r, 2)
            hi = self._table.cellWidget(r, 3)
            if addr and combo and lo and hi:
                x_configs.append({
                    "addr": addr,
                    "type": combo.currentText(),
                    "lo": lo.value(),
                    "hi": hi.value(),
                })
        analyze_single = self._radio_group.checkedId() == 0
        return {"y_cell": self.y_cell, "x_configs": x_configs, "analyze_single": analyze_single}

    def _accept(self):
        if not self.y_cell:
            QMessageBox.warning(self, "提示", "请选择输出变量（Y）单元格。")
            return
        if not self.x_cells:
            QMessageBox.warning(self, "提示", "请至少选择一个输入变量（X）单元格。")
            return
        y_name = self._addr_to_name.get(self.y_cell, self.y_cell)
        x_names = [self._addr_to_name.get(c, c)+f"({c})" for c in self.x_cells]
        x_list = ", ".join(f"{n}" for n in x_names)
        selected_id = self._radio_group.checkedId()
        time = len(self.x_cells) + 1 if selected_id == 0 else 2
        msg = (
            f"输出变量（Y）：{y_name}({self.y_cell})\n\n"
            f"输入变量（X）：{x_list}\n\n"
            f"压力测试次数：{time}\n\n"
            f"总计算次数：{time*int(sim_engine._ribbon_iterations)}\n\n"
            "确认后将开始压力测试模拟，取消返回压力测试设置。"
        )
        msg_box = QDialog(self)
        msg_box.setWindowTitle("确认分析配置")
        set_drisk_icon(msg_box)
        msg_box.setStyleSheet(
            "QDialog { background-color: #f5f6f8; }"
            "QLabel { color: #333; font-size: 12px; font-family: 'Microsoft YaHei'; }"
        )
        msg_box.setFixedWidth(420)
        layout = QVBoxLayout(msg_box)
        layout.setContentsMargins(20, 18, 20, 14)
        layout.setSpacing(14)
        lbl = QLabel(msg)
        lbl.setWordWrap(True)
        lbl.setTextInteractionFlags(Qt.TextSelectableByMouse)
        layout.addWidget(lbl)
        btn_row = QHBoxLayout()
        btn_row.addStretch(1)
        btn_cancel = QPushButton("取消")
        btn_cancel.setFixedSize(80, 28)
        btn_cancel.clicked.connect(msg_box.reject)
        btn_ok = QPushButton("确认")
        btn_ok.setObjectName("btnOk")
        btn_ok.setFixedSize(80, 28)
        btn_ok.setDefault(True)
        btn_ok.clicked.connect(msg_box.accept)
        btn_row.addWidget(btn_cancel)
        btn_row.addWidget(btn_ok)
        layout.addLayout(btn_row)
        if msg_box.exec() == QDialog.Accepted:
            self.accept()
