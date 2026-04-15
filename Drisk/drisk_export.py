# -*- coding: utf-8 -*-
"""
drisk_export.py（轮询通信 + 离屏缓冲终极稳定版）

本模块为 ui_results 体系提供极简且直观的高清“导出”能力。

核心技术难点与解决方案：
1. 混合矢量渲染引擎 (Hybrid Rendering)：图表底图是基于 WebEngine 的 HTML/JS (Plotly)，而上层控制面板（如截断滑块、悬浮窗）是 Qt 原生控件。传统的 `widget.grab()` 截图在放大 DPI 时会导致严重模糊。本模块通过分别获取高分辨率的 JS 图表数据和 Qt 离屏缓冲图层，再进行无损叠合。
2. 全平台跨版本兼容：不同版本的 PyQt/PySide 对异步 JS 执行（Promises）的支持差异巨大。本模块放弃了依赖原生回调，采用强健的 “JS 全局变量注入 + Qt 定时器轮询” 通信机制，彻底跨越了版本兼容鸿沟。
3. 图像后处理：自动剥离透明通道，支持物理 DPI 数据（Dots Per Meter）注入并优化压缩比。
"""

from __future__ import annotations

import os
import base64
import csv
import traceback
from dataclasses import dataclass
from typing import Any, Dict, List, Optional, Sequence, Tuple

from PySide6.QtCore import QEventLoop, QTimer, Qt, QRect, QPoint
from PySide6.QtGui import QImage, QPainter, QColor, QPen
from PySide6.QtWidgets import (
    QDialog, QFileDialog, QHBoxLayout, QLabel, QMessageBox,
    QPushButton, QVBoxLayout, QWidget, QSpinBox, QSizePolicy, QFrame, QApplication,
    QCheckBox, QListWidget, QListWidgetItem, QComboBox
)

from simulation_manager import get_simulation
from input_sample_exposure import is_input_key_exposed
from ui_shared import (
    drisk_spinbox_qss,
    extract_clean_cell_address,
    resolve_visible_variable_name,
)
from ui_variable_search import (
    SEARCH_TEXT_ROLE,
    build_input_reference_display_map,
    build_input_variable_label,
    fuzzy_contains_case_insensitive,
    refresh_combo_dropdown_filter,
    resolve_search_text,
    setup_searchable_combo,
)


# =======================================================
# 1. 基础 UI 组件与数据结构
# =======================================================
class ImagePreviewWidget(QWidget):
    """
    自定义底层绘图组件：用于在导出设置弹窗中预览当前图表。
    核心作用是确保图片在不改变原始长宽比（KeepAspectRatio）的前提下，
    彻底居中并填满预览空间，并在周围绘制边框。
    """
    def __init__(self, image: QImage, parent=None):
        super().__init__(parent)
        self.image = image
        # 允许控件在布局中自由伸缩
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.setMinimumSize(200, 200)

    def paintEvent(self, event):
        painter = QPainter(self)
        # 开启平滑缩放与抗锯齿，保证缩略图质量
        painter.setRenderHint(QPainter.SmoothPixmapTransform)
        painter.setRenderHint(QPainter.Antialiasing)

        # 绘制纯白底色和灰色边框
        painter.fillRect(self.rect(), QColor("#ffffff"))
        painter.setPen(QPen(QColor("#d9d9d9"), 1))
        painter.drawRect(0, 0, self.width() - 1, self.height() - 1)

        if self.image.isNull(): return

        # 核心：计算等比例缩放后的尺寸及居中偏移量
        scaled_size = self.image.size().scaled(self.size(), Qt.KeepAspectRatio)
        x = (self.width() - scaled_size.width()) // 2
        y = (self.height() - scaled_size.height()) // 2
        
        target_rect = QRect(x, y, scaled_size.width(), scaled_size.height())
        painter.drawImage(target_rect, self.image)


@dataclass(frozen=True)
class ExportWorkflowOptions:
    """
    [不可变数据类] 导出工作流选项。
    记录用户在统一导出对话框中选择的具体配置参数。
    """
    export_image: bool
    export_data: bool
    export_stats: bool
    export_input_variables: bool
    export_output_variables: bool
    selected_input_keys: Tuple[str, ...]
    selected_output_keys: Tuple[str, ...]


@dataclass(frozen=True)
class ExportContext:
    """
    [不可变数据类] 显式的导出上下文契约。
    将主对话框 (Dialog) 复杂的运行时状态打包，解耦导出逻辑与 UI 界面。
    """
    dialog: QWidget
    target_widget: QWidget                 # 最终需要进行组合渲染的顶层容器
    web_view: Optional[QWidget]            # 承载 JS 引擎的 WebEngineView
    chart_mode: str                        # 决定文件命名后缀的图表模式
    current_key: str                       # 决定文件命名主体的变量键名
    overlay_widgets: Tuple[QWidget, ...] = ()  # 悬浮在浏览器上方需要二次复合的 Qt 原生控件


# =======================================================
# 2. 导出配置交互对话框
# =======================================================
class VariableSelectionDialog(QDialog):
    """
    复用的轻量级复选列表弹窗，提供类似悬浮层的交互方式，用于选取导出变量。
    """

    def __init__(
        self,
        *,
        title: str,
        caption: str,
        candidates: Sequence[Tuple[str, str]],
        selected_keys: Sequence[str] = (),
        parent=None,
    ):
        super().__init__(parent)
        self.setWindowTitle(title)
        self.resize(340, 400)

        self._candidates = list(candidates or [])
        self._selected_keys = set(str(k) for k in (selected_keys or []))

        base_dir = os.path.dirname(os.path.abspath(__file__))
        icon_path = os.path.join(base_dir, "icons", "Selected_Overlay.svg").replace("\\", "/")
        self.setStyleSheet(
            """
            QDialog { background-color: #f9f9f9; font-family: 'Microsoft YaHei'; }
            QLabel { font-weight: bold; color: #444; font-size: 12px; }
            QListWidget {
                border: 1px solid #ccc;
                border-radius: 3px;
                outline: none;
                font-size: 12px;
                background-color: #ffffff;
            }
            QListWidget::item { padding: 8px 6px; border-bottom: 1px solid #f5f5f5; }
            QListWidget::item:hover { background-color: #f0f7ff; }
            QPushButton {
                background-color: white; border: 1px solid #ccc; border-radius: 3px;
                padding: 4px 12px; font-size: 12px;
            }
            QPushButton:hover { color: #40a9ff; border-color: #40a9ff; background-color: #f0f7ff; }
            QPushButton#btnOk { background-color: #0050b3; color: white; border: none; }
            QCheckBox { font-size: 12px; color: #333; }
            QListWidget::indicator, QCheckBox::indicator {
                width: 16px;
                height: 16px;
                border-radius: 3px;
                border: 1px solid #cbd0d6;
                background-color: #ffffff;
            }
            QListWidget::indicator:hover, QCheckBox::indicator:hover {
                border: 1px solid #8c9eb5;
            }
            QListWidget::indicator:checked, QCheckBox::indicator:checked {
                background-color: #8c9eb5;
                border: 1px solid #8c9eb5;
                image: url('"""
            + icon_path
            + """');
            }
            QListWidget::item:selected {
                background-color: #f0f4f8;
                color: #334455;
                font-weight: bold;
                border-left: 3px solid #8c9eb5;
            }
            """
        )

        layout = QVBoxLayout(self)
        layout.setSpacing(10)
        layout.addWidget(QLabel(caption))
        self.combo_search = QComboBox()
        setup_searchable_combo(self.combo_search, "搜索变量名或地址（例如 Sheet1!B1）")
        self.combo_search.currentIndexChanged.connect(self._on_search_pick)
        if self.combo_search.lineEdit() is not None:
            self.combo_search.lineEdit().textChanged.connect(self._on_search_text_changed)
        layout.addWidget(self.combo_search)

        self.chk_all = QCheckBox("全选")
        self.chk_all.stateChanged.connect(self._on_check_all)
        layout.addWidget(self.chk_all)

        self.list_vars = QListWidget()
        self.list_vars.itemChanged.connect(lambda _item: self._sync_check_state())
        layout.addWidget(self.list_vars, 1)

        for label, key in self._candidates:
            item = QListWidgetItem(str(label))
            item.setData(Qt.UserRole, str(key))
            item.setData(SEARCH_TEXT_ROLE, resolve_search_text(label, key))
            item.setFlags(item.flags() | Qt.ItemIsUserCheckable)
            if str(key) in self._selected_keys:
                item.setCheckState(Qt.Checked)
            else:
                item.setCheckState(Qt.Unchecked)
            self.list_vars.addItem(item)
            self.combo_search.addItem(str(label), str(key))
            idx = self.combo_search.count() - 1
            self.combo_search.setItemData(idx, resolve_search_text(label, key), SEARCH_TEXT_ROLE)
        self.combo_search.setCurrentIndex(-1)
        refresh_combo_dropdown_filter(self.combo_search, "")
        self._sync_check_state()

        btn_layout = QHBoxLayout()
        btn_layout.addStretch()
        btn_ok = QPushButton("确定")
        btn_ok.setObjectName("btnOk")
        btn_ok.setFixedWidth(60)
        btn_cancel = QPushButton("取消")
        btn_cancel.setFixedWidth(60)
        btn_ok.clicked.connect(self.accept)
        btn_cancel.clicked.connect(self.reject)
        btn_layout.addWidget(btn_ok)
        btn_layout.addWidget(btn_cancel)
        layout.addLayout(btn_layout)

    def _on_check_all(self, state: int) -> None:
        check_state = Qt.Checked if state == Qt.Checked.value else Qt.Unchecked
        for i in range(self.list_vars.count()):
            item = self.list_vars.item(i)
            if item.isHidden():
                continue
            item.setCheckState(check_state)
        self._sync_check_state()

    def _on_search_pick(self, idx: int) -> None:
        if idx < 0:
            return
        key = self.combo_search.itemData(idx, Qt.UserRole)
        if key is None:
            return
        key_text = str(key)
        for i in range(self.list_vars.count()):
            item = self.list_vars.item(i)
            if str(item.data(Qt.UserRole)) == key_text and not item.isHidden():
                self.list_vars.setCurrentItem(item)
                self.list_vars.scrollToItem(item)
                break

    def _on_search_text_changed(self, text: str) -> None:
        query = str(text or "")
        refresh_combo_dropdown_filter(self.combo_search, query)
        for i in range(self.list_vars.count()):
            item = self.list_vars.item(i)
            search_text = item.data(SEARCH_TEXT_ROLE) or item.text()
            matched = fuzzy_contains_case_insensitive(query, search_text)
            item.setHidden(not matched)
        # 搜索过滤仅在当前导出上下文已圈定的候选范围内生效
        self._sync_check_state()

    def _sync_check_state(self) -> None:
        visible_count = 0
        checked_visible = 0
        for i in range(self.list_vars.count()):
            item = self.list_vars.item(i)
            if item.isHidden():
                continue
            visible_count += 1
            if item.checkState() == Qt.Checked:
                checked_visible += 1

        if visible_count == 0:
            self.chk_all.setChecked(False)
            self.chk_all.setEnabled(False)
            return
        self.chk_all.blockSignals(True)
        self.chk_all.setChecked(checked_visible == visible_count)
        self.chk_all.setEnabled(True)
        self.chk_all.blockSignals(False)

    def get_selected_keys(self) -> List[str]:
        selected: List[str] = []
        for i in range(self.list_vars.count()):
            item = self.list_vars.item(i)
            if item.checkState() == Qt.Checked:
                selected.append(str(item.data(Qt.UserRole)))
        return selected


class ExportPreviewDialog(QDialog):
    """
    导出配置主对话框：
    提供给用户确认导出画面，支持动态调节目标图片的清晰度（DPI），
    并允许用户选择需要同步导出的图表数据、统计数据及相关变量。
    """

    def __init__(
        self,
        base_image: QImage,
        parent=None,
        *,
        supports_data: bool = True,
        supports_stats: bool = True,
        input_candidates: Sequence[Tuple[str, str]] = (),
        output_candidates: Sequence[Tuple[str, str]] = (),
    ):
        super().__init__(parent)
        self.setWindowTitle("Drisk - 导出预览")

        self._input_candidates = list(input_candidates or [])
        self._output_candidates = list(output_candidates or [])
        self._selected_input_keys: Tuple[str, ...] = ()
        self._selected_output_keys: Tuple[str, ...] = ()

        main_layout = QHBoxLayout(self)
        main_layout.setContentsMargins(15, 15, 15, 15)
        main_layout.setSpacing(15)

        # 左侧区域：图片预览
        self.preview_widget = ImagePreviewWidget(base_image)
        main_layout.addWidget(self.preview_widget, 1)

        # 右侧区域：导出工作流参数配置面板
        right_panel = QFrame()
        right_panel.setObjectName("RightPanel")
        right_panel.setFixedWidth(330)
        right_panel.setStyleSheet("#RightPanel { background: #ffffff; border: 1px solid #d9d9d9; border-radius: 4px; }")

        right_layout = QVBoxLayout(right_panel)
        right_layout.setContentsMargins(15, 15, 15, 15)

        lbl_title = QLabel("导出参数设置")
        lbl_title.setStyleSheet("font-size: 15px; font-weight: bold; color: #333; margin-bottom: 5px;")
        right_layout.addWidget(lbl_title)

        dpi_layout = QHBoxLayout()
        lbl_dpi = QLabel("图片清晰度:")
        lbl_dpi.setStyleSheet("color: #555;")
        self.spin_dpi = QSpinBox()
        self.spin_dpi.setRange(72, 2400)
        self.spin_dpi.setValue(300)
        self.spin_dpi.setSuffix(" DPI")
        self.spin_dpi.setStyleSheet(drisk_spinbox_qss())
        dpi_layout.addWidget(lbl_dpi)
        dpi_layout.addWidget(self.spin_dpi)
        right_layout.addLayout(dpi_layout)

        right_layout.addWidget(QLabel("导出内容"))
        self.chk_export_image = QCheckBox("导出图片")
        self.chk_export_data = QCheckBox("导出模拟结果")
        self.chk_export_stats = QCheckBox("导出统计指标")
        self.chk_export_image.setChecked(True)
        self.chk_export_data.setChecked(False)
        self.chk_export_stats.setChecked(False)
        self.chk_export_data.setEnabled(bool(supports_data))
        self.chk_export_stats.setEnabled(bool(supports_stats))
        right_layout.addWidget(self.chk_export_image)
        right_layout.addWidget(self.chk_export_data)
        right_layout.addWidget(self.chk_export_stats)

        right_layout.addWidget(QLabel("导出变量选择"))
        self.chk_export_input_vars = QCheckBox("输入变量")
        self.chk_export_output_vars = QCheckBox("输出变量")
        self.btn_select_input_vars = QPushButton()
        self.btn_select_output_vars = QPushButton()
        self.btn_select_input_vars.clicked.connect(self._select_input_variables)
        self.btn_select_output_vars.clicked.connect(self._select_output_variables)
        self.chk_export_input_vars.toggled.connect(self._on_toggle_input_variables)
        self.chk_export_output_vars.toggled.connect(self._on_toggle_output_variables)

        self.chk_export_input_vars.setEnabled(bool(self._input_candidates))
        self.chk_export_output_vars.setEnabled(bool(self._output_candidates))
        self.btn_select_input_vars.setEnabled(bool(self._input_candidates))
        self.btn_select_output_vars.setEnabled(bool(self._output_candidates))
        self._update_var_button_text("input")
        self._update_var_button_text("output")
        right_layout.addWidget(self.chk_export_input_vars)
        right_layout.addWidget(self.btn_select_input_vars)
        right_layout.addWidget(self.chk_export_output_vars)
        right_layout.addWidget(self.btn_select_output_vars)
        right_layout.addStretch()

        self.btn_export = QPushButton("开始导出")
        self.btn_export.setFixedHeight(36)
        self.btn_export.setStyleSheet(
            """
            QPushButton { background-color: #0050b3; color: white; border: none; border-radius: 4px; font-weight: bold; font-size: 14px; }
            QPushButton:hover { background-color: #40a9ff; }
            QPushButton:pressed { background-color: #003a8c; }
            """
        )
        self.btn_export.clicked.connect(self.accept)

        self.btn_cancel = QPushButton("取消")
        self.btn_cancel.setFixedHeight(36)
        self.btn_cancel.setStyleSheet(
            """
            QPushButton { background-color: #f0f0f0; color: #333; border: 1px solid #d9d9d9; border-radius: 4px; font-size: 14px; margin-top: 5px; }
            QPushButton:hover { background-color: #e6e6e6; border-color: #adadad; }
            """
        )
        self.btn_cancel.clicked.connect(self.reject)

        right_layout.addWidget(self.btn_export)
        right_layout.addWidget(self.btn_cancel)
        main_layout.addWidget(right_panel)

        img_w = base_image.width()
        img_h = base_image.height()
        if img_w > 0 and img_h > 0:
            aspect = img_w / img_h
            target_h = 600
            target_w = int(target_h * aspect)
            if target_w < 420:
                target_w = 420
            if target_w > 1200:
                target_w = 1200
                target_h = int(target_w / aspect)
            self.resize(target_w + 330 + 45, target_h + 30)
        else:
            self.resize(1080, 600)

    def _update_var_button_text(self, kind: str) -> None:
        if kind == "input":
            cnt = len(self._selected_input_keys)
            self.btn_select_input_vars.setText(f"选择输入变量 ({cnt})")
        else:
            cnt = len(self._selected_output_keys)
            self.btn_select_output_vars.setText(f"选择输出变量 ({cnt})")

    def _open_variable_selector(self, kind: str) -> bool:
        if kind == "input":
            candidates = self._input_candidates
            selected = self._selected_input_keys
            title = "选择输入变量"
            caption = "勾选需要导出的输入变量："
        else:
            candidates = self._output_candidates
            selected = self._selected_output_keys
            title = "选择输出变量"
            caption = "勾选需要导出的输出变量："

        if not candidates:
            QMessageBox.information(self, "提示", "当前无可选择变量。")
            return False

        dlg = VariableSelectionDialog(
            title=title,
            caption=caption,
            candidates=candidates,
            selected_keys=selected,
            parent=self,
        )
        if not dlg.exec():
            return False

        selected_keys = tuple(dlg.get_selected_keys())
        if not selected_keys:
            QMessageBox.warning(self, "提示", "请至少选择一个变量。")
            return False

        if kind == "input":
            self._selected_input_keys = selected_keys
        else:
            self._selected_output_keys = selected_keys
        self._update_var_button_text(kind)
        return True

    def _on_toggle_input_variables(self, checked: bool) -> None:
        if not checked:
            self._selected_input_keys = ()
            self._update_var_button_text("input")
            return
        if not self._open_variable_selector("input"):
            self.chk_export_input_vars.blockSignals(True)
            self.chk_export_input_vars.setChecked(False)
            self.chk_export_input_vars.blockSignals(False)

    def _on_toggle_output_variables(self, checked: bool) -> None:
        if not checked:
            self._selected_output_keys = ()
            self._update_var_button_text("output")
            return
        if not self._open_variable_selector("output"):
            self.chk_export_output_vars.blockSignals(True)
            self.chk_export_output_vars.setChecked(False)
            self.chk_export_output_vars.blockSignals(False)

    def _select_input_variables(self) -> None:
        if not self.chk_export_input_vars.isChecked():
            self.chk_export_input_vars.setChecked(True)
            return
        self._open_variable_selector("input")

    def _select_output_variables(self) -> None:
        if not self.chk_export_output_vars.isChecked():
            self.chk_export_output_vars.setChecked(True)
            return
        self._open_variable_selector("output")

    def _validate_options(self) -> bool:
        has_any = (
            self.chk_export_image.isChecked()
            or self.chk_export_data.isChecked()
            or self.chk_export_stats.isChecked()
            or self.chk_export_input_vars.isChecked()
            or self.chk_export_output_vars.isChecked()
        )
        if not has_any:
            QMessageBox.warning(self, "提示", "请至少选择一种导出内容。")
            return False

        if self.chk_export_input_vars.isChecked() and not self._selected_input_keys:
            if not self._open_variable_selector("input"):
                return False
        if self.chk_export_output_vars.isChecked() and not self._selected_output_keys:
            if not self._open_variable_selector("output"):
                return False
        return True

    def accept(self) -> None:
        if not self._validate_options():
            return
        super().accept()

    def get_dpi(self) -> int:
        return self.spin_dpi.value()

    def get_options(self) -> ExportWorkflowOptions:
        return ExportWorkflowOptions(
            export_image=bool(self.chk_export_image.isChecked()),
            export_data=bool(self.chk_export_data.isChecked()),
            export_stats=bool(self.chk_export_stats.isChecked()),
            export_input_variables=bool(self.chk_export_input_vars.isChecked()),
            export_output_variables=bool(self.chk_export_output_vars.isChecked()),
            selected_input_keys=tuple(self._selected_input_keys),
            selected_output_keys=tuple(self._selected_output_keys),
        )


# =======================================================
# 3. 核心辅助工具函数
# =======================================================
def _grab_widget_image(widget: QWidget) -> Optional[QImage]:
    """
    安全截取当前控件的屏幕缓冲。
    为了防止有些 UI 还在动画渲染过程中导致截出黑屏，
    采用局部事件循环（QEventLoop）短暂挂起 80 毫秒以等待渲染队列清空。
    """
    if widget is None: return None
    img = widget.grab().toImage()
    if not img.isNull(): return img
    
    loop = QEventLoop()
    QTimer.singleShot(80, loop.quit)
    loop.exec()
    
    img = widget.grab().toImage()
    return None if img.isNull() else img


# =======================================================
# 4. 导出管理器核心类
# =======================================================
class DriskExportManager:
    """提供统一导出生命周期管理。"""
    
    _last_export_dir: Optional[str] = None
    # 集中管理支持高级导出的模式标识符，便于后续扩展
    _SUPPORTED_MODE_TAGS = (
        "cdf",
        "pdf",
        "scatter",
        "tornado",
        "scenario",
        "boxplot",
        "letter_value",
        "violin",
        "trend",
    )

    # ---------------------------------------------------
    # 4.1 上下文构建与解析
    # ---------------------------------------------------
    @staticmethod
    def build_export_context(
        dialog: QWidget,
        *,
        target_widget: Optional[QWidget] = None,
        web_view: Optional[QWidget] = None,
        chart_mode: Optional[str] = None,
        current_key: Optional[str] = None,
    ) -> ExportContext:
        """
        组装导出上下文。
        向下兼容旧代码：如果外部没有显式传入目标组件，则利用 getattr 在 dialog 中逐级自动嗅探。
        """
        resolved_target = target_widget if isinstance(target_widget, QWidget) else None
        if resolved_target is None:
            resolved_target = getattr(dialog, "chart_container", getattr(dialog, "web_view", dialog))
            if not isinstance(resolved_target, QWidget):
                resolved_target = dialog

        resolved_web = web_view if isinstance(web_view, QWidget) else None
        if resolved_web is None:
            candidate_web = getattr(dialog, "web_view", getattr(dialog, "web", None))
            resolved_web = candidate_web if isinstance(candidate_web, QWidget) else None

        resolved_mode = str(chart_mode if chart_mode is not None else getattr(dialog, "chart_mode", "hist"))
        resolved_key = str(current_key if current_key is not None else getattr(dialog, "current_key", "result"))
        
        # 嗅探所有可能悬浮在 WebEngine 上方的半透明层、滑块和统计面板
        overlays = DriskExportManager._collect_overlay_widgets(dialog)
        
        return ExportContext(
            dialog=dialog,
            target_widget=resolved_target,
            web_view=resolved_web,
            chart_mode=resolved_mode,
            current_key=resolved_key,
            overlay_widgets=overlays,
        )

    @staticmethod
    def _collect_overlay_widgets(dialog: QWidget) -> Tuple[QWidget, ...]:
        """搜集需要在高清导出时被二次复合（Composited）的浮动原生控件。"""
        widgets = []
        top_wrapper = getattr(dialog, "top_wrapper", getattr(dialog, "slider_wrapper", None))
        if isinstance(top_wrapper, QWidget):
            widgets.append(top_wrapper)

        for attr_name in ("slider_x", "slider_y", "float_x", "float_y", "overlay", "crosshair"):
            widget = getattr(dialog, attr_name, None)
            if isinstance(widget, QWidget):
                widgets.append(widget)

        tornado_view = getattr(dialog, "tornado_view", None)
        config_panel = getattr(tornado_view, "config_panel", None) if tornado_view is not None else None
        if isinstance(config_panel, QWidget):
            widgets.append(config_panel)

        return tuple(widgets)

    @staticmethod
    def _resolve_simulation_id(dialog: QWidget) -> Optional[int]:
        raw_sid = getattr(dialog, "sim_id", None)
        if raw_sid is None:
            return None
        try:
            return int(raw_sid)
        except Exception:
            return None

    # ---------------------------------------------------
    # 4.2 变量映射与过滤逻辑
    # ---------------------------------------------------
    @staticmethod
    def _extract_series_values(raw_values: Any) -> List[Any]:
        """将支持的序列型（Series-like）数据统一归一化为扁平的 Python 列表。"""
        if raw_values is None:
            return []

        obj = getattr(raw_values, "values", raw_values)
        if isinstance(obj, (str, bytes)):
            return [obj]

        if hasattr(obj, "tolist"):
            try:
                obj = obj.tolist()
            except Exception:
                pass

        if isinstance(obj, (list, tuple)):
            seq = list(obj)
        else:
            try:
                seq = list(obj)
            except Exception:
                seq = [obj]

        flat: List[Any] = []
        for it in seq:
            if isinstance(it, (list, tuple)):
                flat.extend(it)
            else:
                flat.append(it)
        return flat

    @staticmethod
    def _build_variable_candidates_for_kind(sim_obj: Any, kind: str) -> List[Tuple[str, str]]:
        if sim_obj is None:
            return []

        if kind == "input":
            cache = getattr(sim_obj, "input_cache", {}) or {}
            attrs_map = getattr(sim_obj, "input_attributes", {}) or {}
        else:
            cache = getattr(sim_obj, "output_cache", {}) or {}
            attrs_map = getattr(sim_obj, "output_attributes", {}) or {}

        display_ref_map: Dict[str, str] = {}
        excel_app = None
        if kind == "input":
            display_ref_map = build_input_reference_display_map(cache.keys())
            try:
                from com_fixer import _safe_excel_app  # 局部导入以保持导出模块在导入时的解耦
                excel_app = _safe_excel_app()
            except Exception:
                excel_app = None

        keyed_labels: Dict[str, str] = {}
        for cell_key, raw_values in cache.items():
            if kind == "input" and not is_input_key_exposed(sim_obj, cell_key):
                continue
            values = DriskExportManager._extract_series_values(raw_values)
            if not values:
                continue
            first = str(values[0]).strip().upper()
            if first.startswith("#") or "NAN" in first:
                continue

            attrs = attrs_map.get(cell_key, {}) if isinstance(attrs_map, dict) else {}
            pure_key = cell_key.split("!")[-1] if "!" in str(cell_key) else str(cell_key)
            if kind == "input":
                display_ref = display_ref_map.get(str(cell_key), pure_key)
                visible_name = DriskExportManager._resolve_input_visible_name(
                    cell_key,
                    attrs,
                    excel_app=excel_app,
                )
                keyed_labels[str(cell_key)] = build_input_variable_label(visible_name, display_ref)
            else:
                var_name = str(attrs.get("name", "") or pure_key)
                keyed_labels[str(cell_key)] = f"{var_name} ({cell_key})"

        return sorted(((lbl, key) for key, lbl in keyed_labels.items()), key=lambda t: t[0].lower())

    @staticmethod
    def _normalize_cell_key(raw_key: Any) -> str:
        return str(raw_key or "").replace("$", "").upper()

    @staticmethod
    def _looks_like_raw_input_key_name(name_text: Any, cell_key: Any) -> bool:
        name = str(name_text or "").strip().replace("$", "")
        key = str(cell_key or "").strip().replace("$", "")
        if not name or not key:
            return False

        name_upper = name.upper()
        key_upper = key.upper()
        if name_upper == key_upper:
            return True

        key_tail = key.split("!", 1)[-1].strip().upper()
        if name_upper == key_tail:
            return True

        clean_name = extract_clean_cell_address(name)
        clean_key = extract_clean_cell_address(key)
        if clean_name and clean_key and clean_name.upper() == clean_key.upper():
            if "_" in name or "_MAKEINPUT" in name_upper:
                return True
        return False

    @staticmethod
    def _resolve_input_visible_name(cell_key: Any, attrs: Any, *, excel_app: Any = None) -> str:
        attrs_map = attrs if isinstance(attrs, dict) else {}
        safe_attrs = dict(attrs_map)
        raw_name = str(safe_attrs.get("name", "") or "").strip()
        if DriskExportManager._looks_like_raw_input_key_name(raw_name, cell_key):
            safe_attrs["name"] = ""

        fallback_addr = extract_clean_cell_address(cell_key)
        visible = resolve_visible_variable_name(
            cell_key,
            safe_attrs,
            excel_app=excel_app,
            fallback_label=fallback_addr,
        )
        text = str(visible or "").strip()
        if text:
            return text
        if fallback_addr:
            return fallback_addr
        key = str(cell_key or "").strip()
        return key.split("!", 1)[-1] if "!" in key else key

    @staticmethod
    def _strip_input_suffix(cell_key: str) -> str:
        key = str(cell_key or "")
        return key.split("_", 1)[0] if "_" in key else key

    @staticmethod
    def _filter_input_keys_by_related_graph(sim_obj: Any, output_key: str, input_keys: Sequence[str]) -> List[str]:
        """
        复用仿真依赖图快照，仅保留从当前输出可达的输入变量。
        此路径在无法获取 Excel COM 对象进行公式实时遍历时使用。
        """
        related_cells = getattr(sim_obj, "all_related_cells", {}) or {}
        if not isinstance(related_cells, dict):
            return []

        node_map: Dict[str, Dict[str, Any]] = {}
        for raw_node, info in related_cells.items():
            node_map[DriskExportManager._normalize_cell_key(raw_node)] = info if isinstance(info, dict) else {}

        output_norm = DriskExportManager._normalize_cell_key(output_key)
        if output_norm not in node_map:
            return []

        input_base_map: Dict[str, List[str]] = {}
        for full_key in input_keys or []:
            base = DriskExportManager._normalize_cell_key(DriskExportManager._strip_input_suffix(full_key))
            input_base_map.setdefault(base, []).append(str(full_key))

        matched: List[str] = []
        matched_set = set()
        queue = [output_norm]
        visited = {output_norm}

        while queue:
            node = queue.pop(0)
            info = node_map.get(node, {})
            deps = info.get("dependencies", []) if isinstance(info, dict) else []
            for dep in deps or []:
                dep_norm = DriskExportManager._normalize_cell_key(dep)
                for full_key in input_base_map.get(dep_norm, []):
                    if full_key not in matched_set:
                        matched.append(full_key)
                        matched_set.add(full_key)
                if dep_norm in node_map and dep_norm not in visited:
                    visited.add(dep_norm)
                    queue.append(dep_norm)

        return matched

    @staticmethod
    def _build_results_scoped_input_candidates(dialog: QWidget, sim_obj: Any) -> List[Tuple[str, str]]:
        """
        将仿真输入候选范围限定为与当前输出变量相关的变量。
        优先复用现有的龙卷风图依赖过滤器，如失败则降级使用缓存的关联单元格网络图。
        """
        if sim_obj is None:
            return []

        base_kind = str(getattr(dialog, "_base_kind", "output") or "output").lower()
        if base_kind != "output":
            # 输入范围是相对于某个输出目标定义的；如果目标不是输出，则禁用该列表。
            return []

        current_key = str(getattr(dialog, "current_key", "") or "")
        label_to_cell_key = getattr(dialog, "label_to_cell_key", {}) or {}
        output_key = str(label_to_cell_key.get(current_key, current_key) or "")
        if not output_key:
            return []

        input_cache = getattr(sim_obj, "input_cache", {}) or {}
        input_attrs = getattr(sim_obj, "input_attributes", {}) or {}
        input_ref_map = build_input_reference_display_map(input_cache.keys())
        input_label_map: Dict[str, str] = {}
        input_mapping_for_filter: Dict[str, str] = {}
        filter_token_to_key: Dict[str, str] = {}
        xl_app = None
        try:
            from com_fixer import _safe_excel_app  # 局部导入以保持导出模块在导入时的解耦
            xl_app = _safe_excel_app()
        except Exception:
            xl_app = None

        for cell_key, raw_values in input_cache.items():
            if not is_input_key_exposed(sim_obj, cell_key):
                continue
            values = DriskExportManager._extract_series_values(raw_values)
            if not values:
                continue
            first = str(values[0]).strip().upper()
            if first.startswith("#") or "NAN" in first:
                continue
            attrs = input_attrs.get(cell_key, {}) if isinstance(input_attrs, dict) else {}
            pure_key = cell_key.split("!")[-1] if "!" in str(cell_key) else str(cell_key)
            display_ref = input_ref_map.get(str(cell_key), pure_key)
            visible_name = DriskExportManager._resolve_input_visible_name(
                cell_key,
                attrs,
                excel_app=xl_app,
            )
            label = build_input_variable_label(visible_name, display_ref)
            input_label_map[str(cell_key)] = label
            # 使用内部稳定令牌进行依赖过滤，避免标签冲突导致覆盖。
            token = f"__exp_input__{len(filter_token_to_key)}"
            filter_token_to_key[token] = str(cell_key)
            input_mapping_for_filter[token] = str(cell_key)

        if not input_mapping_for_filter:
            return []

        scoped_keys: List[str] = []

        if xl_app is not None:
            try:
                from ui_tornado import filter_inputs_by_dependency
                scoped_labels = filter_inputs_by_dependency(xl_app, output_key, input_mapping_for_filter)
                for lbl in scoped_labels or []:
                    key = filter_token_to_key.get(str(lbl))
                    if key:
                        scoped_keys.append(key)
            except Exception:
                scoped_keys = []

        if not scoped_keys:
            scoped_keys = DriskExportManager._filter_input_keys_by_related_graph(
                sim_obj,
                output_key,
                list(input_label_map.keys()),
            )

        if not scoped_keys:
            return []

        dedup = []
        seen = set()
        for key in scoped_keys:
            sk = str(key)
            if sk in seen or sk not in input_label_map:
                continue
            seen.add(sk)
            dedup.append((input_label_map[sk], sk))
        return dedup

    @staticmethod
    def _resolve_scatter_addr_key(sim_obj: Any, full_addr: str) -> Tuple[Optional[str], Optional[str]]:
        clean_addr = DriskExportManager._normalize_cell_key(full_addr)
        if not clean_addr:
            return None, None

        output_cache = getattr(sim_obj, "output_cache", {}) or {}
        for cache_key in output_cache.keys():
            cache_norm = DriskExportManager._normalize_cell_key(cache_key)
            if cache_norm == clean_addr or cache_norm.startswith(f"{clean_addr}_"):
                return "output", str(cache_key)

        input_cache = getattr(sim_obj, "input_cache", {}) or {}
        for cache_key in input_cache.keys():
            if not is_input_key_exposed(sim_obj, cache_key):
                continue
            cache_norm = DriskExportManager._normalize_cell_key(cache_key)
            cache_base = cache_norm.split("_", 1)[0] if "_" in cache_norm else cache_norm
            if cache_base == clean_addr:
                return "input", str(cache_key)

        return None, None

    @staticmethod
    def _build_scatter_scoped_candidates(dialog: QWidget, sim_obj: Any) -> Tuple[List[Tuple[str, str]], List[Tuple[str, str]]]:
        """
        将散点图的变量候选范围限定为当前图表选择所实际使用的变量。
        """
        if sim_obj is None:
            return [], []

        used_addrs: List[str] = []
        y_addr = getattr(dialog, "y_addr", None)
        if y_addr:
            used_addrs.append(str(y_addr))
        for x_addr in getattr(dialog, "x_addrs", []) or []:
            if x_addr:
                used_addrs.append(str(x_addr))

        if not used_addrs:
            return [], []

        input_attrs = getattr(sim_obj, "input_attributes", {}) or {}
        output_attrs = getattr(sim_obj, "output_attributes", {}) or {}
        input_ref_map = build_input_reference_display_map(getattr(sim_obj, "input_cache", {}).keys())
        input_candidates: List[Tuple[str, str]] = []
        output_candidates: List[Tuple[str, str]] = []
        seen_input = set()
        seen_output = set()
        xl_app = None
        try:
            from com_fixer import _safe_excel_app  # 局部导入以保持导出模块在导入时的解耦
            xl_app = _safe_excel_app()
        except Exception:
            xl_app = None

        for addr in used_addrs:
            kind, cache_key = DriskExportManager._resolve_scatter_addr_key(sim_obj, addr)
            if not kind or not cache_key:
                continue
            if kind == "input":
                if cache_key in seen_input:
                    continue
                seen_input.add(cache_key)
                attrs = input_attrs.get(cache_key, {}) if isinstance(input_attrs, dict) else {}
                pure_key = cache_key.split("!")[-1] if "!" in cache_key else cache_key
                display_ref = input_ref_map.get(str(cache_key), pure_key)
                visible_name = DriskExportManager._resolve_input_visible_name(
                    cache_key,
                    attrs,
                    excel_app=xl_app,
                )
                input_candidates.append((build_input_variable_label(visible_name, display_ref), cache_key))
            else:
                if cache_key in seen_output:
                    continue
                seen_output.add(cache_key)
                attrs = output_attrs.get(cache_key, {}) if isinstance(output_attrs, dict) else {}
                pure_key = cache_key.split("!")[-1] if "!" in cache_key else cache_key
                var_name = str(attrs.get("name", "") or pure_key)
                output_candidates.append((f"{var_name} ({cache_key})", cache_key))

        return input_candidates, output_candidates

    # ---------------------------------------------------
    # 4.3 导出能力检测与数据生成
    # ---------------------------------------------------
    @staticmethod
    def _collect_workflow_capabilities(dialog: QWidget) -> Dict[str, Any]:
        """从运行时对话框状态中收集导出功能标志和变量候选列表。"""
        supports_data = False
        if getattr(dialog, "_raw_data_cache", None):
            supports_data = True
        elif getattr(dialog, "series_map", None):
            supports_data = True
        elif getattr(dialog, "all_datasets", None):
            supports_data = True

        stats_table = getattr(dialog, "stats_table", None)
        supports_stats = bool(
            stats_table is not None
            and hasattr(stats_table, "rowCount")
            and hasattr(stats_table, "columnCount")
            and stats_table.rowCount() > 0
            and stats_table.columnCount() > 0
        )

        input_candidates: List[Tuple[str, str]] = []
        output_candidates: List[Tuple[str, str]] = []
        sid = DriskExportManager._resolve_simulation_id(dialog)
        if sid is not None:
            sim_obj = get_simulation(sid)
            is_scatter = hasattr(dialog, "x_addrs") and hasattr(dialog, "y_addr")
            is_results = hasattr(dialog, "label_to_cell_key") and hasattr(dialog, "current_key")

            if is_scatter:
                input_candidates, output_candidates = DriskExportManager._build_scatter_scoped_candidates(dialog, sim_obj)
            elif is_results:
                input_candidates = DriskExportManager._build_results_scoped_input_candidates(dialog, sim_obj)
                output_candidates = DriskExportManager._build_variable_candidates_for_kind(sim_obj, "output")
            else:
                input_candidates = DriskExportManager._build_variable_candidates_for_kind(sim_obj, "input")
                output_candidates = DriskExportManager._build_variable_candidates_for_kind(sim_obj, "output")

        return {
            "supports_data": supports_data,
            "supports_stats": supports_stats,
            "input_candidates": input_candidates,
            "output_candidates": output_candidates,
        }

    @staticmethod
    def _ensure_parent_dir(file_path: str) -> None:
        directory = os.path.dirname(file_path)
        if directory and not os.path.exists(directory):
            os.makedirs(directory, exist_ok=True)

    @staticmethod
    def _cell_text(table: Any, row: int, col: int) -> str:
        item = table.item(row, col) if hasattr(table, "item") else None
        if item is not None:
            try:
                return str(item.text())
            except Exception:
                pass

        widget = table.cellWidget(row, col) if hasattr(table, "cellWidget") else None
        if widget is not None:
            try:
                if hasattr(widget, "value"):
                    return str(widget.value())
                if hasattr(widget, "text"):
                    return str(widget.text())
            except Exception:
                pass
        return ""

    @staticmethod
    def _export_chart_data_csv(context: ExportContext, base_root: str) -> str:
        """导出当前活动绘图对话框所使用的图表级数据为 CSV。"""
        dialog = context.dialog
        file_path = f"{base_root}_data.csv"
        DriskExportManager._ensure_parent_dir(file_path)
        rows_written = 0

        with open(file_path, "w", newline="", encoding="utf-8-sig") as f:
            writer = csv.writer(f)
            raw_groups = getattr(dialog, "_raw_data_cache", None)
            if isinstance(raw_groups, list) and raw_groups:
                writer.writerow(["group", "index", "x", "y"])
                for group in raw_groups:
                    name = str(group.get("name", ""))
                    x_vals = DriskExportManager._extract_series_values(group.get("x", []))
                    y_vals = DriskExportManager._extract_series_values(group.get("y", []))
                    n = min(len(x_vals), len(y_vals))
                    for i in range(n):
                        writer.writerow([name, i, x_vals[i], y_vals[i]])
                        rows_written += 1
            else:
                series_map = getattr(dialog, "series_map", {}) or {}
                display_keys = list(getattr(dialog, "display_keys", [])) or list(series_map.keys())
                if not display_keys and getattr(dialog, "current_key", None):
                    display_keys = [str(getattr(dialog, "current_key", ""))]
                all_datasets = getattr(dialog, "all_datasets", {}) or {}

                writer.writerow(["series", "index", "value"])
                for key in display_keys:
                    values = DriskExportManager._extract_series_values(series_map.get(key))
                    if not values and key in all_datasets:
                        values = DriskExportManager._extract_series_values(all_datasets.get(key))
                    for i, val in enumerate(values):
                        writer.writerow([str(key), i, val])
                        rows_written += 1

        if rows_written <= 0:
            raise RuntimeError("当前图表没有可导出的数据内容。")
        return file_path

    @staticmethod
    def _export_stats_csv(context: ExportContext, base_root: str) -> str:
        """从当前活动的绘图对话框中将已渲染的统计信息表导出为 CSV。"""
        dialog = context.dialog
        table = getattr(dialog, "stats_table", None)
        if table is None or table.rowCount() <= 0 or table.columnCount() <= 0:
            raise RuntimeError("当前界面没有可导出的统计结果。")

        file_path = f"{base_root}_stats.csv"
        DriskExportManager._ensure_parent_dir(file_path)

        with open(file_path, "w", newline="", encoding="utf-8-sig") as f:
            writer = csv.writer(f)
            headers: List[str] = []
            for c in range(table.columnCount()):
                header_item = table.horizontalHeaderItem(c)
                headers.append(str(header_item.text()) if header_item is not None else f"col_{c+1}")
            writer.writerow(headers)

            for r in range(table.rowCount()):
                row_vals = [DriskExportManager._cell_text(table, r, c) for c in range(table.columnCount())]
                writer.writerow(row_vals)
        return file_path

    @staticmethod
    def _resolve_cache_value_by_key(cache: Dict[str, Any], key: str) -> Tuple[Optional[str], Any]:
        if key in cache:
            return key, cache.get(key)

        normalized_target = str(key).replace("$", "").upper()
        for cache_key, cache_val in cache.items():
            if str(cache_key).replace("$", "").upper() == normalized_target:
                return str(cache_key), cache_val
        return None, None

    @staticmethod
    def _export_selected_variables_csv(
        context: ExportContext,
        base_root: str,
        *,
        input_keys: Sequence[str] = (),
        output_keys: Sequence[str] = (),
    ) -> str:
        """从仿真缓存中导出所选输入/输出变量的样本序列。"""
        dialog = context.dialog
        sid = DriskExportManager._resolve_simulation_id(dialog)
        if sid is None:
            raise RuntimeError("当前导出上下文没有有效 simulation id，无法导出变量。")

        sim_obj = get_simulation(sid)
        if sim_obj is None:
            raise RuntimeError(f"无法找到 simulation id={sid} 的缓存数据。")

        export_plan = [
            ("input", tuple(str(k) for k in (input_keys or ()))),
            ("output", tuple(str(k) for k in (output_keys or ()))),
        ]
        rows_written = 0
        unresolved_requests: Dict[str, List[str]] = {"input": [], "output": []}
        file_path = f"{base_root}_variables.csv"
        DriskExportManager._ensure_parent_dir(file_path)

        with open(file_path, "w", newline="", encoding="utf-8-sig") as f:
            writer = csv.writer(f)
            writer.writerow(["kind", "sim_id", "variable_name", "cell_key", "index", "value"])

            for kind, keys in export_plan:
                if not keys:
                    continue
                if kind == "input":
                    cache = getattr(sim_obj, "input_cache", {}) or {}
                    attrs_map = getattr(sim_obj, "input_attributes", {}) or {}
                else:
                    cache = getattr(sim_obj, "output_cache", {}) or {}
                    attrs_map = getattr(sim_obj, "output_attributes", {}) or {}

                for request_key in keys:
                    if kind == "input" and not is_input_key_exposed(sim_obj, request_key):
                        unresolved_requests[kind].append(f"{request_key} (not exposed)")
                        continue

                    real_key, raw_values = DriskExportManager._resolve_cache_value_by_key(cache, request_key)
                    if real_key is None:
                        unresolved_requests[kind].append(request_key)
                        continue

                    attrs = attrs_map.get(real_key, {}) if isinstance(attrs_map, dict) else {}
                    pure_key = real_key.split("!")[-1] if "!" in str(real_key) else str(real_key)
                    var_name = str(attrs.get("name", "") or pure_key)
                    values = DriskExportManager._extract_series_values(raw_values)
                    for i, val in enumerate(values):
                        writer.writerow([kind, sid, var_name, real_key, i, val])
                        rows_written += 1

        if rows_written <= 0:
            parts: List[str] = []
            for kind in ("input", "output"):
                bad = unresolved_requests.get(kind, [])
                if bad:
                    shown = ", ".join(bad[:6])
                    if len(bad) > 6:
                        shown += f", ... (total={len(bad)})"
                    parts.append(f"{kind}: {shown}")
            if parts:
                raise RuntimeError(
                    "变量导出失败：所选变量无法解析为有效缓存键。"
                    + " | "
                    + " ; ".join(parts)
                )
            raise RuntimeError("变量导出为空，请检查已选变量是否存在可用样本。")
        return file_path

    # ---------------------------------------------------
    # 4.4 导出生命周期主控
    # ---------------------------------------------------
    @staticmethod
    def export_from_dialog(dialog: QWidget, context: Optional[ExportContext] = None) -> None:
        """
        导出动作主入口。
        协调截图 -> 预览对话框 -> 文件保存对话框 -> 执行导出 的完整业务流。
        """
        ctx = context or DriskExportManager.build_export_context(dialog)
        target_widget = ctx.target_widget
        base_image = _grab_widget_image(target_widget)
        
        if base_image is None:
            QMessageBox.critical(dialog, "截图失败", "无法获取界面截图（组件可能尚未渲染完成）。")
            return

        workflow_caps = DriskExportManager._collect_workflow_capabilities(dialog)
        preview_dlg = ExportPreviewDialog(
            base_image,
            dialog,
            supports_data=bool(workflow_caps.get("supports_data", False)),
            supports_stats=bool(workflow_caps.get("supports_stats", False)),
            input_candidates=workflow_caps.get("input_candidates", []) or [],
            output_candidates=workflow_caps.get("output_candidates", []) or [],
        )
        if not preview_dlg.exec():
            return

        target_dpi = preview_dlg.get_dpi()
        options = preview_dlg.get_options()
        suggested = DriskExportManager._suggest_default_filename(ctx)

        need_multi_payload = (
            options.export_data
            or options.export_stats
            or options.export_input_variables
            or options.export_output_variables
        )
        image_path: Optional[str] = None
        base_root: Optional[str] = None

        if options.export_image and not need_multi_payload:
            file_path, _selected_filter = QFileDialog.getSaveFileName(
                dialog, "保存导出图片", suggested, "PNG 图片 (*.png);;JPG 图片 (*.jpg)"
            )
            if not file_path:
                return
            root, ext = os.path.splitext(file_path)
            if not ext:
                file_path = file_path + ".png"
                root = file_path[:-4]
            image_path = file_path
            base_root = root
        else:
            default_base = os.path.splitext(suggested)[0]
            base_path, _selected_filter = QFileDialog.getSaveFileName(
                dialog, "选择导出文件基名", default_base, "所有文件 (*)"
            )
            if not base_path:
                return
            root, ext = os.path.splitext(base_path)
            base_root = root if ext else base_path

            if options.export_image:
                ext_lower = ext.lower()
                if ext_lower in (".png", ".jpg", ".jpeg"):
                    image_path = base_path
                else:
                    image_path = base_root + ".png"

        if not base_root:
            QMessageBox.critical(dialog, "导出失败", "无法确定导出文件名。")
            return

        last_dir = os.path.dirname(image_path or base_root) or os.getcwd()
        DriskExportManager._last_export_dir = last_dir
        exported_files: List[str] = []

        try:
            # 保持图片导出路径的向后兼容性，并在启用时附加额外的载荷文件。
            if options.export_image and image_path:
                img_path = DriskExportManager._do_hybrid_export(ctx, base_image, image_path, target_dpi)
                if img_path:
                    exported_files.append(img_path)

            if options.export_data:
                exported_files.append(DriskExportManager._export_chart_data_csv(ctx, base_root))

            if options.export_stats:
                exported_files.append(DriskExportManager._export_stats_csv(ctx, base_root))

            if options.export_input_variables or options.export_output_variables:
                exported_files.append(
                    DriskExportManager._export_selected_variables_csv(
                        ctx,
                        base_root,
                        input_keys=options.selected_input_keys if options.export_input_variables else (),
                        output_keys=options.selected_output_keys if options.export_output_variables else (),
                    )
                )
        except Exception as e:
            traceback.print_exc()
            QMessageBox.critical(dialog, "导出失败", f"导出过程中发生内部错误：\n{e}")
            return

        if exported_files:
            msg = "\n".join(exported_files)
            QMessageBox.information(dialog, "导出完成", f"已导出以下文件：\n{msg}")
        else:
            QMessageBox.warning(dialog, "导出提示", "未生成任何导出文件。")

    @staticmethod
    def _suggest_default_filename(context: ExportContext) -> str:
        """根据当前变量名和分析模式智能拼装导出文件的默认名称。"""
        if DriskExportManager._last_export_dir and os.path.exists(DriskExportManager._last_export_dir):
            base_dir = DriskExportManager._last_export_dir
        else:
            base_dir = os.getcwd()

        key = context.current_key
        mode = context.chart_mode

        mode_tag = mode if mode in DriskExportManager._SUPPORTED_MODE_TAGS else "hist"
        
        # 净化文件名：将非法字符替换为下划线
        safe_key = "".join(c if (c.isalnum() or c in "-_") else "_" for c in str(key))
        return os.path.join(base_dir, f"drisk_{safe_key}_{mode_tag}.png")

    # ---------------------------------------------------
    # 4.5 混合渲染与离屏缓冲引擎 (核心技术区)
    # ---------------------------------------------------
    @staticmethod
    def _do_hybrid_export(context: ExportContext, base_image: QImage, file_path: str, target_dpi: int) -> Optional[str]:
        """
        极度稳定的真高清混合导出引擎。
        处理复杂的 JS Plotly 高分图表与 Qt 原生层 (Overlays) 的复合。
        """
        dialog = context.dialog
        target_widget = context.target_widget
        scale_factor = target_dpi / 96.0  # 核心缩放倍率 (300 DPI 时约为 3.125倍)
        web_view = context.web_view
        
        # 若无需放大或未挂载 WebEngine，直接降级处理
        if web_view is None or (not hasattr(web_view, "page")) or scale_factor <= 1.0:
            return DriskExportManager._fallback_export(base_image, file_path, target_dpi)
            
        # -------------------------------------------------------------------------
        # 第一步：规避 Promise 异常，利用 JS 全局变量进行通信
        # 为什么这么做？：不同版本的 PyQt 提供的 runJavaScript 对 Promise(then/catch) 的
        # 返回值支持极不稳定（甚至崩溃）。通过将生成的 Base64 字符串赋值给 window 下的变量，
        # 并通过定时器轮询拉取，可以实现 100% 的兼容与稳定。
        # -------------------------------------------------------------------------
        init_js = f"""
        (function() {{
            window.__drisk_export_b64 = "LOADING";
            try {{
                var gd = document.querySelector('.js-plotly-plot') || document.getElementById('chart_div');
                var PlotlyObj = window.Plotly;
                
                // 处理可能被包裹在 iframe 中的情况
                if (!gd) {{
                    var frames = document.getElementsByTagName('iframe');
                    if (frames.length > 0 && frames[0].contentDocument) {{
                        gd = frames[0].contentDocument.querySelector('.js-plotly-plot') || frames[0].contentDocument.getElementById('chart_div');
                        PlotlyObj = frames[0].contentWindow.Plotly;
                    }}
                }}
                
                if (!gd) {{ window.__drisk_export_b64 = "ERR: NO_GD_FOUND"; return; }}
                if (!PlotlyObj) {{ window.__drisk_export_b64 = "ERR: NO_PLOTLY_OBJ"; return; }}
                
                // 触发底层 Plotly 渲染管线，并附带动态的放大参数 (scale_factor)
                PlotlyObj.toImage(gd, {{
                    format: 'png', 
                    width: gd.clientWidth || 800, 
                    height: gd.clientHeight || 600, 
                    scale: {scale_factor}
                }}).then(function(dataUrl) {{
                    window.__drisk_export_b64 = dataUrl;
                }}).catch(function(err) {{
                    window.__drisk_export_b64 = "ERR: " + err.toString();
                }});
            }} catch(e) {{
                window.__drisk_export_b64 = "ERR: " + e.toString();
            }}
        }})();
        """
        web_view.page().runJavaScript(init_js)
        
        # 使用局部事件循环阻塞主线程，等待 JS 异步渲染完成
        loop = QEventLoop()
        b64_result = ["LOADING"]
        elapsed_time = [0]
        
        def check_result():
            """轮询探针：每 200ms 执行一次，检查 JS 全局变量的状态"""
            elapsed_time[0] += 200
            if elapsed_time[0] > 12000:  # 12秒终极防死锁超时保护
                b64_result[0] = "ERR: TIMEOUT"
                if loop.isRunning(): loop.quit()
                return
                
            def cb(val):
                # 如果获取到了实际内容，则打断事件循环并恢复执行
                if val and str(val) != "LOADING":
                    b64_result[0] = str(val)
                    if loop.isRunning(): loop.quit()
                    
            try:
                web_view.page().runJavaScript("window.__drisk_export_b64;", cb)
            except Exception:
                pass
                
        # 修改鼠标指针以提示用户正在后台导出
        QApplication.setOverrideCursor(Qt.WaitCursor)
        poll_timer = QTimer()
        poll_timer.setInterval(200)
        poll_timer.timeout.connect(check_result)
        poll_timer.start()
        
        loop.exec()
        
        poll_timer.stop()
        QApplication.restoreOverrideCursor()
        
        # 结果解析与降级判断
        b64_data = b64_result[0]
        if not b64_data.startswith('data:image/png;base64,'):
            QMessageBox.warning(
                dialog, "未触发真高清渲染", 
                f"无法调用底层的矢量渲染引擎，已自动降级为截图拉伸模式（图表边缘会模糊）。\n\n调试信息: {b64_data}"
            )
            return DriskExportManager._fallback_export(base_image, file_path, target_dpi)
            
        b64_str = b64_data.split(',')[1]
        img_data = base64.b64decode(b64_str)
        plotly_img = QImage.fromData(img_data)
        
        if plotly_img.isNull():
            return DriskExportManager._fallback_export(base_image, file_path, target_dpi)
            
        # -------------------------------------------------------------------------
        # 第二步：安全的离屏渲染与混合 (Off-screen Compositing)
        # 为什么这么做？：如果在当前激活的 QWidget 的 Painter 上强制套用新的 Painter，
        # 极易触发 `QPainter::begin: Paint device returned engine == 0` 的底层 C++ 崩溃。
        # 这里通过申请一张全新的独立空白大图 (high_res_img)，完全在内存中进行重绘拼合。
        # -------------------------------------------------------------------------
        w = target_widget.width()
        h = target_widget.height()
        high_res_img = QImage(int(w * scale_factor), int(h * scale_factor), QImage.Format_RGB32)
        high_res_img.fill(Qt.white)  # 填入纯白底色
        
        painter = QPainter(high_res_img)
        painter.setRenderHint(QPainter.Antialiasing)
        painter.setRenderHint(QPainter.TextAntialiasing)
        painter.setRenderHint(QPainter.SmoothPixmapTransform)
        
        def is_descendant(child: QWidget, parent: QWidget) -> bool:
            """判断组件包含关系，用于准确定位各图层的相对坐标偏移量。"""
            p = child
            while p is not None:
                if p == parent: return True
                p = p.parentWidget()
            return False
            
        # A) 将来自 JS 的超高清 Plotly 底图贴入背景
        if is_descendant(web_view, target_widget):
            # 考虑 Web 容器距离顶部大容器可能有边界 Padding 偏移
            pos_web = web_view.mapTo(target_widget, QPoint(0, 0))
            target_rect = QRect(int(pos_web.x() * scale_factor), int(pos_web.y() * scale_factor), plotly_img.width(), plotly_img.height())
            painter.drawImage(target_rect, plotly_img)
        else:
            painter.drawImage(0, 0, plotly_img)
            
        # B) 内部闭包：安全的离屏原生控件抓取
        def safe_render_overlay(widget_obj):
            if not widget_obj or not widget_obj.isVisible() or not is_descendant(widget_obj, target_widget):
                return
            
            # 为当前小控件建立一个专属的高倍率透明缓冲层
            temp_img = QImage(int(widget_obj.width() * scale_factor), int(widget_obj.height() * scale_factor), QImage.Format_ARGB32_Premultiplied)
            temp_img.fill(Qt.transparent)
            
            temp_painter = QPainter(temp_img)
            temp_painter.scale(scale_factor, scale_factor) # 放大画布坐标系
            widget_obj.render(temp_painter, QPoint(0, 0))  # 此时绘制出来的文字等矢量元素将极其锐利
            temp_painter.end()
            
            # 测算控件在大容器中的实际位置并映射叠合
            pos = widget_obj.mapTo(target_widget, QPoint(0, 0))
            painter.drawImage(int(pos.x() * scale_factor), int(pos.y() * scale_factor), temp_img)

        # 遍历需要复合的叠加控件（如滑块、指示线等）并在高分图层上重绘
        for overlay_widget in context.overlay_widgets:
            safe_render_overlay(overlay_widget)
            
        painter.end()
        return DriskExportManager._save_image_to_disk(high_res_img, file_path, target_dpi)

    @staticmethod
    def _fallback_export(base_image: QImage, file_path: str, target_dpi: int) -> Optional[str]:
        """降级导出模式：直接对获取到的低分辨率屏幕截图进行强制插值放大（边缘会模糊）。"""
        scale_factor = target_dpi / 96.0
        img_to_save = base_image

        if scale_factor != 1.0:
            img_to_save = base_image.scaled(int(base_image.width() * scale_factor), int(base_image.height() * scale_factor), Qt.KeepAspectRatio, Qt.SmoothTransformation)
            
        # 移除 Alpha 通道以防止导出 JPG 时出现黑色背景
        if img_to_save.hasAlphaChannel():
            img_to_save = img_to_save.convertToFormat(QImage.Format_RGB32)

        return DriskExportManager._save_image_to_disk(img_to_save, file_path, target_dpi)

    @staticmethod
    def _save_image_to_disk(img_to_save: QImage, file_path: str, target_dpi: int) -> str:
        """物理写入磁盘，并向元数据中注入物理分辨率 (Dots Per Meter)。"""
        # DPI 转 DPM (1 英寸 = 0.0254 米)
        dpm = int(target_dpi / 0.0254)
        img_to_save.setDotsPerMeterX(dpm)
        img_to_save.setDotsPerMeterY(dpm)

        directory = os.path.dirname(file_path)
        if directory and not os.path.exists(directory): os.makedirs(directory, exist_ok=True)

        ext = os.path.splitext(file_path)[1].lower()
        if ext in ['.jpg', '.jpeg']:
            ok = img_to_save.save(file_path, quality=95) # 对 JPG 进行轻微压缩以缩小体积
        else:
            ok = img_to_save.save(file_path, quality=-1) # PNG 无损保存
            
        if not ok: raise RuntimeError(f"保存图片文件失败，请检查路径权限：{file_path}")
        return file_path
