# ui_style_manager.py
"""
本模块提供 UI 图表样式的管理与配置组件。
作为样式控制的核心模块，负责收集并下发用户的视觉配置。

主要功能模块：
1. 颜色选择组件 (Color Pickers)：提供类似 Excel 的预设调色板下拉菜单，以及调用并深度定制了系统高级取色器的功能。
2. 样式管理面板 (Style Manager Dialog)：提供一个高复用性的统一样式配置弹窗，支持对柱状图填充、折线、散点形状、均值线等多种视觉元素进行独立调节。
"""

import re

from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QPushButton, QWidget, QGroupBox,
    QGridLayout, QMenu, QSpinBox, QDoubleSpinBox, QComboBox, QTabWidget, 
    QWidgetAction, QColorDialog, QDialogButtonBox, QSlider, QAbstractSpinBox
)
from PySide6.QtCore import Qt, Signal
from PySide6.QtGui import QColor, QPainter, QIcon, QPixmap

# 引入底层的 10 色循环表，用于重置功能，以及标准对话框按钮样式
from ui_shared import DRISK_COLOR_CYCLE, DRISK_DIALOG_BTN_QSS


# =======================================================
# 1. 颜色选择组件 (Color Picker Components)
# =======================================================

class ExcelColorMenu(QMenu):
    """
    [UI 组件] 类 Excel 的颜色下拉面板。
    提供预设的主题色、标准色矩阵，并支持“无填充”与“自定义颜色”选项。
    用于在有限空间内快速选择常用颜色。
    """
    # 当用户选中某种颜色时发射，附带十六进制颜色字符串 (如 "#4472c4" 或 "transparent")
    colorSelected = Signal(str) 

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setStyleSheet("QMenu { background-color: white; border: 1px solid #ccc; font-family: 'Microsoft YaHei'; }")
        
        container = QWidget()
        layout = QVBoxLayout(container)
        layout.setContentsMargins(8, 8, 8, 8)
        layout.setSpacing(6)

        # --- 1. 定义预设色板 ---
        # 预设的 20 种办公主题色彩，与常见办公软件保持一致
        theme_colors = [
            "#ffffff", "#000000", "#e7e6e6", "#44546a", "#4472c4", "#ed7d31", "#a5a5a5", "#ffc000", "#5b9bd5", "#70ad47",
            "#f2f2f2", "#7f7f7f", "#d0cece", "#d6dce4", "#d9e1f2", "#fce4d6", "#ededed", "#fff2cc", "#deebf7", "#e2efda"
        ]
        # 预设的 10 种高饱和度标准色彩
        standard_colors = ["#c00000", "#ff0000", "#ffc000", "#ffff00", "#92d050", "#00b050", "#00b0f0", "#0070c0", "#002060", "#7030a0"]

        def create_color_btn(hex_c):
            """
            [内部闭包] 批量生成无文字的颜色块按钮。
            绑定点击事件以发射对应的颜色字符串。
            """
            btn = QPushButton()
            btn.setFixedSize(20, 20)
            btn.setStyleSheet(f"QPushButton {{ background-color: {hex_c}; border: 1px solid #d9d9d9; }} "
                              f"QPushButton:hover {{ border: 1px solid #000000; }}")
            btn.clicked.connect(lambda checked, c=hex_c: self._emit_and_close(c))
            return btn

        # --- 2. 构建主题颜色矩阵 ---
        layout.addWidget(QLabel("主题颜色"))
        theme_grid = QGridLayout()
        theme_grid.setSpacing(2)
        for i, hex_c in enumerate(theme_colors):
            # 以 10 列为基准进行网格排布
            theme_grid.addWidget(create_color_btn(hex_c), i // 10, i % 10)
        layout.addLayout(theme_grid)

        # --- 3. 构建标准色单行 ---
        layout.addWidget(QLabel("标准色"))
        std_grid = QGridLayout()
        std_grid.setSpacing(2)
        for i, hex_c in enumerate(standard_colors):
            std_grid.addWidget(create_color_btn(hex_c), 0, i)
        layout.addLayout(std_grid)

        # --- 4. 底部扩展选项 ---
        # 4.1 无填充选项
        btn_no_fill = QPushButton("  无填充(N)")
        btn_no_fill.setStyleSheet("QPushButton { text-align: left; border: none; padding: 4px; background: transparent; color: #333; } "
                                  "QPushButton:hover { background-color: #f0f0f0; }")
        btn_no_fill.clicked.connect(lambda: self._emit_and_close("transparent"))
        layout.addWidget(btn_no_fill)

        # 4.2 自定义更多颜色选项
        btn_more = QPushButton("🎨 其他颜色(M)...")
        btn_more.setStyleSheet("QPushButton { text-align: left; border: none; padding: 4px; background: transparent; color: #333; } "
                               "QPushButton:hover { background-color: #f0f0f0; }")
        btn_more.clicked.connect(self._open_more_colors)
        layout.addWidget(btn_more)

        # 将生成的布局容器挂载到 QMenu 的 Action 上
        action = QWidgetAction(self)
        action.setDefaultWidget(container)
        self.addAction(action)

    def _emit_and_close(self, color):
        """
        [事件槽] 发送颜色信号并关闭当前下拉面板。
        
        参数:
            color (str): 十六进制颜色代码或 "transparent"。
        """
        self.colorSelected.emit(color)
        self.close()

    def _open_more_colors(self):
        """
        [事件槽] 调用系统/Qt原生的增强调色板，并进行深度的 UI 汉化与美化定制。
        替换原生英文标签，修复不可见的预设色块，并统一按钮样式。
        """
        self.close()
        parent_window = self.window()
        current_hex = self.parent().current_color
        initial_color = QColor(current_hex) if current_hex != "transparent" else Qt.white
        
        # 实例化原生调色板并禁用系统原生 UI，强制使用 Qt 自绘版本以便进行定制开发
        dialog = QColorDialog(initial_color, parent_window)
        dialog.setOption(QColorDialog.DontUseNativeDialog, True)
        dialog.setWindowTitle("Drisk - 选择颜色")

        # [定制修复] 强制将所有“纯白(隐形)”的自定义网格预填充为可见的浅灰色，避免 UI 空白导致的交互困惑
        for i in range(QColorDialog.customCount()):
            if QColorDialog.customColor(i) == QColor(Qt.white):
                QColorDialog.setCustomColor(i, QColor("#e8e8e8"))

        # [界面汉化] Qt 原生对话框英文字段字典映射表
        translate_dict = {
            "Basic colors": "基本颜色",
            "Custom colors": "自定义颜色",
            "Add to Custom Colors": "添加", 
            "Pick Screen Color": "屏幕取色",
            "Hue:": "色相(H):",
            "Sat:": "饱和度(S):",
            "Val:": "亮度(L):",
            "Red:": "红色(R):",
            "Green:": "绿色(G):",
            "Blue:": "蓝色(B):",
            "Alpha channel:": "透明度(A):",
            "HTML:": "十六进制(H):"
        }

        # 遍历原生控件并替换标签文字，实现深度汉化
        for widget in dialog.findChildren(QLabel) + dialog.findChildren(QPushButton):
            clean_text = widget.text().replace('&', '')
            if clean_text in translate_dict:
                widget.setText(translate_dict[clean_text])

        # [样式统一] 拦截数值输入框，去除原生难看的上下调节箭头，统一边框样式
        for widget in dialog.findChildren(QSpinBox):
            widget.setButtonSymbols(QAbstractSpinBox.NoButtons)
            widget.setStyleSheet("""
                QSpinBox { border: 1px solid #ccc; border-radius: 3px; padding: 2px 4px; background: white; }
                QSpinBox:focus { border-color: #40a9ff; }
            """)

        # [样式统一] 拦截底部的确定和取消按钮，赋予与主界面相同的现代化配色
        bbox = dialog.findChild(QDialogButtonBox)
        if bbox:
            btn_ok = bbox.button(QDialogButtonBox.Ok)
            if btn_ok: 
                btn_ok.setText("确定")
                btn_ok.setFixedSize(60, 24)
                btn_ok.setStyleSheet("""
                    QPushButton { background-color: #0050b3; color: white; border: none; border-radius: 4px; font-size: 12px; padding: 0px; }
                    QPushButton:hover { background-color: #40a9ff; }
                """)
                
            btn_cancel = bbox.button(QDialogButtonBox.Cancel)
            if btn_cancel: 
                btn_cancel.setText("取消")
                btn_cancel.setFixedSize(60, 24)
                btn_cancel.setStyleSheet("""
                    QPushButton { background-color: white; color: #333333; border: 1px solid #cccccc; border-radius: 4px; font-size: 12px; padding: 0px; }
                    QPushButton:hover { background-color: #f5f5f5; border-color: #40a9ff; }
                """)

        # 阻塞执行并捕获取色结果
        if dialog.exec():
            color = dialog.currentColor()
            if color.isValid():
                self.colorSelected.emit(color.name(QColor.HexRgb))


class ColorPickerButton(QPushButton):
    """
    [UI 组件] 带有颜色预览矩形框的交互按钮。
    点击后弹出 ExcelColorMenu 选择颜色，并自动更新自身的预览图标与文本说明。
    """
    colorChanged = Signal(str)

    def __init__(self, default_color="#4472c4", parent=None):
        super().__init__(parent)
        self.current_color = default_color
        self.setFixedSize(95, 26) 
        self.clicked.connect(self._show_menu)
        
        self.setStyleSheet("""
            QPushButton { 
                background-color: #f5f5f5; border: 1px solid #cccccc; border-radius: 3px; 
                color: #333333; font-weight: normal; text-align: center; padding-left: 6px; font-family: Arial;
            }
            QPushButton:hover { border-color: #40a9ff; background-color: #e6f7ff; }
        """)
        self.update_color(default_color)

    @staticmethod
    def _parse_color_for_display(raw_color):
        """
        [静态工具方法] 解析颜色字符串用于图标显示。
        处理标准十六进制、Qt 命名颜色，以及 CSS 风格的 RGB/RGBA 字符串。
        
        参数:
            raw_color (str): 原始颜色字符串。
        返回:
            QColor: 解析成功后的 Qt 颜色对象，失败则返回 None。
        """
        text = str(raw_color or "").strip()
        if not text:
            return None

        qt_color = QColor(text)
        if qt_color.isValid():
            return qt_color

        # 接受类似 CSS 的 rgb()/rgba() 格式字符串，并将其标准化处理为 Qt 原生的 QColor 对象。
        match = re.fullmatch(
            r"rgba?\(\s*([-+]?\d*\.?\d+)\s*,\s*([-+]?\d*\.?\d+)\s*,\s*([-+]?\d*\.?\d+)(?:\s*,\s*([-+]?\d*\.?\d+))?\s*\)",
            text,
            flags=re.IGNORECASE,
        )
        if not match:
            return None

        def _clamp_channel(value):
            """内部闭包：将通道值限制在 0-255 范围内"""
            return int(max(0, min(255, round(float(value)))))

        r = _clamp_channel(match.group(1))
        g = _clamp_channel(match.group(2))
        b = _clamp_channel(match.group(3))
        a_raw = match.group(4)
        if a_raw is None:
            a = 255
        else:
            a_val = float(a_raw)
            a = _clamp_channel(a_val * 255.0 if 0.0 <= a_val <= 1.0 else a_val)

        return QColor(r, g, b, a)

    def update_color(self, hex_c):
        """
        [核心方法] 动态绘制包含当前颜色的正方形预览图标，并更新按钮文字描述。
        
        参数:
            hex_c (str): 目标颜色字符串。
        """
        self.current_color = hex_c
        color_text = str(hex_c or "").strip()
        
        pixmap = QPixmap(14, 14)
        if color_text.lower() == "transparent":
            # 无填充模式下：绘制带有红色对角线的透明图块
            pixmap.fill(Qt.transparent)
            painter = QPainter(pixmap)
            painter.setPen(Qt.red)
            painter.drawLine(0, 14, 14, 0)
            painter.end()
            display_text = " 无填充"
        else:
            # 常规颜色填充模式下：填充色块并附加灰色边框增强质感
            parsed_color = self._parse_color_for_display(color_text)
            if parsed_color is None:
                parsed_color = QColor("#000000")
            pixmap.fill(parsed_color)
            painter = QPainter(pixmap)
            painter.setPen(QColor("#aaaaaa"))
            painter.drawRect(0, 0, 13, 13) 
            painter.end()
            display_hex = parsed_color.name(QColor.HexRgb).upper()
            display_text = f" {display_hex}"

        self.setIcon(QIcon(pixmap))
        self.setText(display_text)

    def _show_menu(self):
        """[事件槽] 计算按钮位置并弹出下拉菜单"""
        menu = ExcelColorMenu(self)
        menu.colorSelected.connect(self._on_color_picked)
        menu.exec(self.mapToGlobal(self.rect().bottomLeft()))

    def _on_color_picked(self, hex_c):
        """[事件槽] 响应菜单的颜色选择事件，更新视图并向外转发信号"""
        self.update_color(hex_c)
        self.colorChanged.emit(hex_c)


# =======================================================
# 2. 全局样式管理面板 (Global Style Manager Dialog)
# =======================================================

class StyleManagerDialog(QDialog):
    """
    [核心控制器] 统一样式管理器弹窗。
    负责聚合展示和收集所有图表视图（直方图、散点图、箱线图、CDF等）的视觉属性（颜色、线宽、点型、透明度等）。
    支持通过参数控制各个功能面板的显隐，以适应不同图表的配置需求。
    """
    def __init__(
        self,
        display_keys,
        current_styles,
        show_bar=True,
        show_curve=True,
        show_mean=False,
        is_unified=False,
        parent=None,
        *,
        show_marker=False,
        marker_shapes=None,
        marker_size_range=(1.0, 20.0),
        style_profile=None,
    ):
        super().__init__(parent)
        self.setWindowTitle("Drisk - 样式设置")
        self.is_unified = is_unified
        self.resize(380, 520) 
        
        # --- 全局 QSS 统一定义 ---
        self.setStyleSheet("""
            QDialog { background-color: #f5f5f5; font-family: 'Microsoft YaHei'; } 
            QLabel { font-size: 12px; }
            QGroupBox { font-weight: bold; border: 1px solid #ccc; border-radius: 4px; margin-top: 10px; padding-top: 12px; }
            QGroupBox::title { subcontrol-origin: margin; left: 8px; padding: 0 3px; }
            QComboBox { border: 1px solid #cccccc; border-radius: 3px; padding: 4px 6px; background: white; }
            QComboBox:focus { border-color: #40a9ff; }
            QSpinBox, QDoubleSpinBox { border: 1px solid #cccccc; border-radius: 3px; padding: 4px 6px; background: white; }
            QSpinBox:focus, QDoubleSpinBox:focus { border-color: #40a9ff; }
        """)

        # --- 状态数据初始化 ---
        self.display_keys = display_keys
        self.styles = current_styles 
        self.current_series = display_keys[0] if display_keys else None
        
        # --- 面板显隐开关配置 ---
        self.show_bar = show_bar
        self.show_curve = show_curve
        self.show_mean = show_mean       # 是否展示“均值线”高级配置（常用于箱线图）
        self.style_profile = dict(style_profile or {})
        self.show_outline_controls = bool(self.style_profile.get("show_outline_controls", True))
        
        # 散点形状控件是可选的，主要为散点图样式的复用提供支持
        self.show_marker = bool(show_marker)
        self.marker_shapes = [str(s) for s in (marker_shapes or ["circle"])]
        try:
            min_size = float(marker_size_range[0])
            max_size = float(marker_size_range[1])
        except Exception:
            min_size, max_size = 1.0, 20.0
        if max_size <= min_size:
            max_size = min_size + 1.0
        self.marker_size_range = (min_size, max_size)

        # 渲染 UI 并加载初始数据
        self._init_ui()
        self._load_series_style(self.current_series)

    def _profile_default(self, key, fallback):
        """
        [工具方法] 从配置档案中安全提取默认值。
        
        参数:
            key (str): 配置键名。
            fallback (Any): 当键不存在或数据格式不正确时的后备值。
        """
        defaults = self.style_profile.get("defaults", {})
        if not isinstance(defaults, dict):
            return fallback
        return defaults.get(key, fallback)

    def _create_slider(self):
        """
        [工厂方法] 创建具有统一样式的透明度滑动条。
        覆写滑槽、已划过区域(sub-page)及滑块手柄的样式，保持界面质感统一。
        """
        slider = QSlider(Qt.Horizontal)
        slider.setRange(0, 100)
        
        slider.setStyleSheet("""
            QSlider::groove:horizontal { 
                border: 1px solid #ccc; height: 4px; background: #e8e8e8; border-radius: 2px; 
            }
            QSlider::sub-page:horizontal { 
                background: #8c8c8c; border-radius: 2px; 
            }
            QSlider::handle:horizontal { 
                background: #fff; border: 1px solid #888; width: 12px; 
                margin-top: -5px; margin-bottom: -5px; border-radius: 6px; 
            }
            QSlider::handle:horizontal:hover { border: 1px solid #555555; }
        """)
        return slider

    def _init_ui(self):
        """
        [核心构建] 界面构造与布局组装。
        包含三个主要区域：顶部的数据组选择器、中部的选项卡配置区、底部的确认/取消操作区。
        """
        layout = QVBoxLayout(self)

        # ==================== 顶部区域：数据组切换器 ====================
        top_layout = QHBoxLayout()
        self.label_select = QLabel("选择数据组:")
        top_layout.addWidget(self.label_select)
        
        self.combo_series = QComboBox()
        self.combo_series.setStyleSheet("QComboBox { background: white; border: 1px solid #ccc; padding: 4px 6px; }")
        self.combo_series.addItems(self.display_keys)
        self.combo_series.currentTextChanged.connect(self._on_series_changed)
        
        top_layout.addWidget(self.combo_series, 1)
        layout.addLayout(top_layout)

        # 统一模式下隐藏下拉框（此时所有系列强制应用同一套样式，无需切换）
        if getattr(self, "is_unified", False):
            self.label_select.hide()
            self.combo_series.hide()

        # ==================== 中部区域：选项卡容器 ====================
        self.tabs = QTabWidget()
        self.tabs.setStyleSheet("QTabBar::tab { padding: 6px 16px; } QTabWidget::pane { border: 1px solid #ccc; background: white; }")
        
        # -------------------- Tab 1: 柱体/填充/散点设置 --------------------
        self.tab_bar = QWidget()
        self.tab_bar.setStyleSheet("background: white;")
        bar_layout = QGridLayout(self.tab_bar)
        bar_layout.setSpacing(10)
        
        row = 0
        if self.show_marker:
            # 复用当前的对话框面板，并在“填充”选项卡中优先挂载散点配置项
            bar_layout.addWidget(QLabel("点样式"), row, 0, 1, 2)
            row += 1

            self.combo_marker_shape = QComboBox()
            self.combo_marker_shape.setStyleSheet("QComboBox { background: white; border: 1px solid #ccc; padding: 4px 6px; }")
            self.combo_marker_shape.addItems(self.marker_shapes)
            bar_layout.addWidget(QLabel("形状:"), row, 0)
            bar_layout.addWidget(self.combo_marker_shape, row, 1)
            row += 1

            self.spin_marker_size = QDoubleSpinBox()
            self.spin_marker_size.setRange(self.marker_size_range[0], self.marker_size_range[1])
            self.spin_marker_size.setSingleStep(0.5)
            self.spin_marker_size.setButtonSymbols(QAbstractSpinBox.NoButtons)
            bar_layout.addWidget(QLabel("大小:"), row, 0)
            bar_layout.addWidget(self.spin_marker_size, row, 1)
            row += 1

            # 整体不透明度联动组件 (数字输入与滑块绑定)
            self.spin_marker_opacity = QSpinBox()
            self.spin_marker_opacity.setRange(0, 100)
            self.spin_marker_opacity.setSuffix(" %")
            self.spin_marker_opacity.setButtonSymbols(QAbstractSpinBox.NoButtons)
            self.slider_marker_opacity = self._create_slider()
            # 绑定 SpinBox 和 Slider 的双向联动事件
            self.spin_marker_opacity.valueChanged.connect(self.slider_marker_opacity.setValue)
            self.slider_marker_opacity.valueChanged.connect(self.spin_marker_opacity.setValue)
            
            marker_op_layout = QVBoxLayout()
            marker_op_layout.setSpacing(4)
            marker_op_layout.addWidget(self.spin_marker_opacity)
            marker_op_layout.addWidget(self.slider_marker_opacity)
            bar_layout.addWidget(QLabel("整体不透明度:"), row, 0, Qt.AlignTop)
            bar_layout.addLayout(marker_op_layout, row, 1)
            row += 1
        else:
            self.combo_marker_shape = None
            self.spin_marker_size = None
            self.spin_marker_opacity = None
            self.slider_marker_opacity = None

        # 动态获取配置标签文本
        fill_section_title = str(self.style_profile.get("fill_section_title", "填充"))
        fill_color_label = str(self.style_profile.get("fill_color_label", "填充颜色:"))
        fill_opacity_label = str(self.style_profile.get("fill_opacity_label", "填充不透明度:"))
        outline_section_title = str(self.style_profile.get("outline_section_title", "轮廓"))
        outline_color_label = str(self.style_profile.get("outline_color_label", "轮廓颜色:"))
        outline_width_label = str(self.style_profile.get("outline_width_label", "轮廓宽度:"))

        # 填充颜色配置区块
        bar_layout.addWidget(QLabel(fill_section_title), row, 0, 1, 2)
        row += 1
        self.btn_fill_color = ColorPickerButton()
        bar_layout.addWidget(QLabel(fill_color_label), row, 0)
        bar_layout.addWidget(self.btn_fill_color, row, 1)
        row += 1

        # 填充透明度联动组件
        self.spin_fill_opacity = QSpinBox()
        self.spin_fill_opacity.setRange(0, 100)
        self.spin_fill_opacity.setSuffix(" %")
        self.spin_fill_opacity.setButtonSymbols(QAbstractSpinBox.NoButtons)
        self.slider_fill_opacity = self._create_slider()
        self.spin_fill_opacity.valueChanged.connect(self.slider_fill_opacity.setValue)
        self.slider_fill_opacity.valueChanged.connect(self.spin_fill_opacity.setValue)

        fill_op_layout = QVBoxLayout()
        fill_op_layout.setSpacing(4)
        fill_op_layout.addWidget(self.spin_fill_opacity)
        fill_op_layout.addWidget(self.slider_fill_opacity)
        bar_layout.addWidget(QLabel(fill_opacity_label), row, 0, Qt.AlignTop)
        bar_layout.addLayout(fill_op_layout, row, 1)
        row += 1

        if self.show_outline_controls:
            # 轮廓/边框颜色配置区块
            bar_layout.addWidget(QLabel(outline_section_title), row, 0, 1, 2)
            row += 1
            self.btn_line_color = ColorPickerButton()
            bar_layout.addWidget(QLabel(outline_color_label), row, 0)
            bar_layout.addWidget(self.btn_line_color, row, 1)
            row += 1

            self.spin_line_width = QDoubleSpinBox()
            self.spin_line_width.setRange(0, 10)
            self.spin_line_width.setSingleStep(0.5)
            self.spin_line_width.setButtonSymbols(QAbstractSpinBox.NoButtons)
            bar_layout.addWidget(QLabel(outline_width_label), row, 0)
            bar_layout.addWidget(self.spin_line_width, row, 1)
            row += 1
        else:
            self.btn_line_color = None
            self.spin_line_width = None
            
        bar_layout.setRowStretch(row, 1) # 压实上方的空间，保证控件紧凑
        
        # -------------------- Tab 2: 曲线/CDF/双线条设置 --------------------
        self.tab_curve = QWidget()
        self.tab_curve.setStyleSheet("background: white;")
        curve_main_layout = QVBoxLayout(self.tab_curve)
        
        # 🟢 第一组：主线条配置 (如 KDE曲线、CDF曲线)
        gb_main = QGroupBox("主线条 (曲线 / 中位数线)")
        curve_layout = QGridLayout(gb_main)
        curve_layout.setSpacing(10)
        
        self.btn_curve_color = ColorPickerButton()
        curve_layout.addWidget(QLabel("线条颜色:"), 0, 0)
        curve_layout.addWidget(self.btn_curve_color, 0, 1)

        self.spin_curve_width = QDoubleSpinBox()
        self.spin_curve_width.setRange(0.5, 10)
        self.spin_curve_width.setSingleStep(0.5)
        self.spin_curve_width.setButtonSymbols(QAbstractSpinBox.NoButtons)
        curve_layout.addWidget(QLabel("线条宽度:"), 1, 0)
        curve_layout.addWidget(self.spin_curve_width, 1, 1)

        self.combo_curve_dash = QComboBox()
        self.combo_curve_dash.setStyleSheet("QComboBox { background: white; border: 1px solid #ccc; padding: 4px 6px; }")
        self.combo_curve_dash.addItems(["实线", "虚线", "点线", "点划线"])
        curve_layout.addWidget(QLabel("线条类型:"), 2, 0)
        curve_layout.addWidget(self.combo_curve_dash, 2, 1)

        self.spin_curve_opacity = QSpinBox()
        self.spin_curve_opacity.setRange(0, 100)
        self.spin_curve_opacity.setSuffix(" %")
        self.spin_curve_opacity.setButtonSymbols(QAbstractSpinBox.NoButtons)
        self.slider_curve_opacity = self._create_slider()
        self.spin_curve_opacity.valueChanged.connect(self.slider_curve_opacity.setValue)
        self.slider_curve_opacity.valueChanged.connect(self.spin_curve_opacity.setValue)

        curve_op_layout = QVBoxLayout()
        curve_op_layout.setSpacing(4)
        curve_op_layout.addWidget(self.spin_curve_opacity)
        curve_op_layout.addWidget(self.slider_curve_opacity)
        curve_layout.addWidget(QLabel("不透明度:"), 3, 0, Qt.AlignTop)
        curve_layout.addLayout(curve_op_layout, 3, 1)
        
        curve_main_layout.addWidget(gb_main)

        # 🟢 第二组：均值线配置 (根据 `show_mean` 开关动态加载，常用于箱线图内独立控制均值线样式)
        if self.show_mean:
            gb_mean = QGroupBox("副线条 (均值线)")
            mean_layout = QGridLayout(gb_mean)
            mean_layout.setSpacing(10)
            
            self.btn_mean_color = ColorPickerButton()
            mean_layout.addWidget(QLabel("线条颜色:"), 0, 0)
            mean_layout.addWidget(self.btn_mean_color, 0, 1)

            self.spin_mean_width = QDoubleSpinBox()
            self.spin_mean_width.setRange(0.5, 10)
            self.spin_mean_width.setSingleStep(0.5)
            self.spin_mean_width.setButtonSymbols(QAbstractSpinBox.NoButtons)
            mean_layout.addWidget(QLabel("线条宽度:"), 1, 0)
            mean_layout.addWidget(self.spin_mean_width, 1, 1)

            self.combo_mean_dash = QComboBox()
            self.combo_mean_dash.setStyleSheet("QComboBox { background: white; border: 1px solid #ccc; padding: 4px 6px; }")
            self.combo_mean_dash.addItems(["实线", "虚线", "点线", "点划线"])
            mean_layout.addWidget(QLabel("线条类型:"), 2, 0)
            mean_layout.addWidget(self.combo_mean_dash, 2, 1)

            self.spin_mean_opacity = QSpinBox()
            self.spin_mean_opacity.setRange(0, 100)
            self.spin_mean_opacity.setSuffix(" %")
            self.spin_mean_opacity.setButtonSymbols(QAbstractSpinBox.NoButtons)
            self.slider_mean_opacity = self._create_slider()
            self.spin_mean_opacity.valueChanged.connect(self.slider_mean_opacity.setValue)
            self.slider_mean_opacity.valueChanged.connect(self.spin_mean_opacity.setValue)

            mean_op_layout = QVBoxLayout()
            mean_op_layout.setSpacing(4)
            mean_op_layout.addWidget(self.spin_mean_opacity)
            mean_op_layout.addWidget(self.slider_mean_opacity)
            mean_layout.addWidget(QLabel("不透明度:"), 3, 0, Qt.AlignTop)
            mean_layout.addLayout(mean_op_layout, 3, 1)

            curve_main_layout.addWidget(gb_mean)
            
        curve_main_layout.addStretch()
        
        # 将配置完成的选项卡挂载至主控面板
        if self.show_bar: self.tabs.addTab(self.tab_bar, "填充")
        if self.show_curve: self.tabs.addTab(self.tab_curve, "线条")

        layout.addWidget(self.tabs)

        # ==================== 底部区域：对话框按钮区 ====================
        self.setStyleSheet(self.styleSheet() + DRISK_DIALOG_BTN_QSS)

        btn_layout = QHBoxLayout()
        btn_layout.setContentsMargins(0, 10, 0, 0)

        # 重置按钮靠左放置，复用取消样式进行视觉弱化
        self.btn_reset = QPushButton("重置")
        self.btn_reset.setObjectName("btnCancel") 
        self.btn_reset.clicked.connect(self._reset_all_styles)
        btn_layout.addWidget(self.btn_reset)

        btn_layout.addStretch() 

        # 确定按钮在左，取消按钮在右，遵循常规窗口习惯
        self.btn_ok = QPushButton("确定")
        self.btn_ok.setObjectName("btnOk")
        self.btn_ok.setFixedSize(80, 28)
        self.btn_ok.clicked.connect(self.accept)
        
        self.btn_cancel = QPushButton("取消")
        self.btn_cancel.setObjectName("btnCancel")
        self.btn_cancel.setFixedSize(80, 28)
        self.btn_cancel.clicked.connect(self.reject)

        btn_layout.addWidget(self.btn_ok)
        btn_layout.addWidget(self.btn_cancel)
        layout.addLayout(btn_layout)

        # 统一样式：将所有的数值调整框文字左对齐居中
        for spin in self.findChildren(QSpinBox):
            spin.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        for spin in self.findChildren(QDoubleSpinBox):
            spin.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)

    def _on_series_changed(self, new_series_name):
        """
        [事件槽] 当用户在顶部下拉框切换数据序列时触发。
        自动保存当前旧视图的修改，并加载新选中序列的样式状态进行展示。
        """
        self._save_current_style()
        self._load_series_style(new_series_name)

    def _reset_all_styles(self):
        """
        [重置功能] 恢复系统默认设置：
        调用全局定义的颜色循环环 (DRISK_COLOR_CYCLE) 重新赋值，清除用户所有自定义修改。
        """
        n = len(DRISK_COLOR_CYCLE)
        marker_n = len(self.marker_shapes)
        for i, k in enumerate(self.display_keys):
            self.styles[k] = {
                "color": DRISK_COLOR_CYCLE[i % n],
                "cdf_color": DRISK_COLOR_CYCLE[i % n],
                "dash": "solid" if i < n else "dash"
            }
            if self.show_marker:
                # 在现有的状态字典中重置散点相关属性，保证后续链路可兼容读取
                self.styles[k].update({
                    "marker_symbol": self.marker_shapes[i % marker_n] if marker_n > 0 else "circle",
                    "marker_size": 4.5,
                    "fill_color": DRISK_COLOR_CYCLE[i % n],
                    "fill_opacity": 0.6,
                    "outline_color": "#ffffff",
                    "outline_width": 0.0,
                    "opacity": 1.0,
                })
        # 刷新当前 UI 面板的展示状态
        self._load_series_style(self.current_series)

    def accept(self):
        """
        [生命周期拦截] 覆写系统基类方法。
        确保用户在点击“确认”离开弹窗前，最后一次触发字典持久化写入，防止丢失最后编辑的数据组状态。
        """
        self._save_current_style()
        super().accept()

    def _load_series_style(self, series_name):
        """
        [数据流：模型 -> 视图] 
        根据当前选定的序列名，从内存字典中读取样式并绑定渲染到界面控件上。
        
        参数:
            series_name (str): 序列标识名
        """
        if not series_name: return
        self.current_series = series_name
        st = self.styles.get(series_name, {})
        
        base_color = st.get("color", "#4472c4")
        if self.show_marker and self.combo_marker_shape is not None:
            # 加载散点形状、大小、透明度等属性
            default_symbol = self.marker_shapes[0] if self.marker_shapes else "circle"
            marker_symbol = str(st.get("marker_symbol", default_symbol))
            if self.combo_marker_shape.findText(marker_symbol) < 0:
                marker_symbol = default_symbol
            self.combo_marker_shape.setCurrentText(marker_symbol)
            self.spin_marker_size.setValue(float(st.get("marker_size", 4.5)))
            self.spin_marker_opacity.setValue(int(float(st.get("opacity", 1.0)) * 100.0))

        self.btn_fill_color.update_color(st.get("fill_color", base_color))
        
        # 保持默认设置与当前加载的渲染配置文件（Profile）相一致，区分散点与箱线图的特定透明度需求。
        default_fill_opacity = 0.6 if self.show_marker else (0.85 if getattr(self, "show_mean", False) else 0.5)
        default_fill_opacity = float(self._profile_default("fill_opacity", default_fill_opacity))
        self.spin_fill_opacity.setValue(int(st.get("fill_opacity", default_fill_opacity) * 100))
        
        if self.show_outline_controls and self.btn_line_color is not None and self.spin_line_width is not None:
            # 轮廓线默认颜色分流控制：散点和常规视图默认使用白色轮廓，而箱线图族为了边界清晰默认使用深灰色
            default_outline_color = "#666666" if getattr(self, "show_mean", False) else "#ffffff"
            default_outline_color = str(self._profile_default("outline_color", default_outline_color))
            self.btn_line_color.update_color(st.get("outline_color", default_outline_color))

            default_outline_width = 0.0 if self.show_marker else (0.5 if getattr(self, "show_mean", False) else 0.8)
            default_outline_width = float(self._profile_default("outline_width", default_outline_width))
            self.spin_line_width.setValue(st.get("outline_width", default_outline_width))
        
        # 🔴 特殊区分：箱线图族（带有 show_mean 标记）的主线条（即中位数线）默认为纯白色反白，其他常规图表跟随主色
        default_curve_color = "#ffffff" if getattr(self, "show_mean", False) else base_color
        self.btn_curve_color.update_color(st.get("curve_color", default_curve_color))
        
        default_curve_width = 1.5 if getattr(self, "show_mean", False) else 1.8
        self.spin_curve_width.setValue(st.get("curve_width", default_curve_width))
        
        self.spin_curve_opacity.setValue(int(st.get("curve_opacity", 1.0) * 100))
        
        dash_map_rev = {"solid": "实线", "dash": "虚线", "dot": "点线", "dashdot": "点划线"}
        self.combo_curve_dash.setCurrentText(dash_map_rev.get(st.get("dash", "solid"), "实线"))
        
        # 独立加载副线（均值线）样式数据
        if getattr(self, "show_mean", False):
            # 🔴 特殊区分：均值线默认纯黑色，以增强其在图表中的对比度
            self.btn_mean_color.update_color(st.get("mean_color", "#000000"))
            self.spin_mean_width.setValue(st.get("mean_width", 1.5))
            self.combo_mean_dash.setCurrentText(dash_map_rev.get(st.get("mean_dash", "solid"), "实线"))
            self.spin_mean_opacity.setValue(int(st.get("mean_opacity", 1.0) * 100))

    def _save_current_style(self):
        """
        [数据流：视图 -> 模型] 
        将界面上的当前 UI 状态抓取，并覆盖写入至内部的 self.styles 字典中。
        处理完毕后等待最终向下游图表组件移交应用。
        """
        if not self.current_series: return
        
        dash_map = {"实线": "solid", "虚线": "dash", "点线": "dot", "点划线": "dashdot"}
        
        # 如果启用“统一配置”开关 (is_unified)，则把当前面板上的属性克隆覆写给所有的图表序列
        keys_to_update = self.display_keys if getattr(self, "is_unified", False) else [self.current_series]
        
        for k in keys_to_update:
            if k not in self.styles:
                self.styles[k] = {}
            outline_color = str(self._profile_default("outline_color", "rgba(0,0,0,0)"))
            outline_width = float(self._profile_default("outline_width", 0.0))
            if self.show_outline_controls and self.btn_line_color is not None and self.spin_line_width is not None:
                outline_color = self.btn_line_color.current_color
                outline_width = self.spin_line_width.value()
                
            # 写入通用填充与线条属性
            self.styles[k].update({
                "fill_color": self.btn_fill_color.current_color,
                "fill_opacity": self.spin_fill_opacity.value() / 100.0,
                "outline_color": outline_color,
                "outline_width": outline_width,
                "curve_color": self.btn_curve_color.current_color,
                "curve_width": self.spin_curve_width.value(),
                "dash": dash_map.get(self.combo_curve_dash.currentText(), "solid"),
                "curve_opacity": self.spin_curve_opacity.value() / 100.0
            })
            if self.show_marker and self.combo_marker_shape is not None:
                # 将散点的样式字段保存在同一个状态字典中，以便直接无缝传递给底层的图表渲染路径
                self.styles[k].update({
                    "marker_symbol": self.combo_marker_shape.currentText(),
                    "marker_size": self.spin_marker_size.value(),
                    "opacity": self.spin_marker_opacity.value() / 100.0,
                })
            
            # ✅ 独立保存箱线图特有的副线 (均值线) 样式
            if getattr(self, "show_mean", False):
                self.styles[k].update({
                    "mean_color": self.btn_mean_color.current_color,
                    "mean_width": self.spin_mean_width.value(),
                    "mean_dash": dash_map.get(self.combo_mean_dash.currentText(), "dash"),
                    "mean_opacity": self.spin_mean_opacity.value() / 100.0
                })