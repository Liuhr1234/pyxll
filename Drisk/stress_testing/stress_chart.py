# -*- coding: utf-8 -*-
"""
压力测试分析主窗口
"""
import sys
import os
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QSplitter, QWidget, QLabel,
    QComboBox, QPushButton, QTableWidget, QTableWidgetItem, QApplication,
    QHeaderView, QSizePolicy
)
from PySide6.QtCore import Qt
from ui_shared import get_shared_webview, set_drisk_icon
from drisk_charting import DriskChartFactory
import numpy as np


class StressTestDialog(QDialog):
    """
    压力测试分析主窗口。
    左侧显示因变量（Y）分布图表，右侧显示压力测试结果统计表。
    """

    def __init__(self, results: list[dict], y_addr: str, y_cache: dict, parent=None):
        super().__init__(parent)
        self.results = results
        self.y_addr = y_addr
        self.y_cache = y_cache
        self.x_addrs = list(set([r["x_addr"] for r in results]))
        self.current_x_index = 0

        # 调试：打印所有 x_addr
        print(f"[DEBUG] 因变量地址: {y_addr}")
        print(f"[DEBUG] 所有自变量地址: {self.x_addrs}")
        if results:
            print(f"[DEBUG] 第一个结果示例: {results[0]}")

        self.setWindowTitle("Drisk - 压力测试分析")
        set_drisk_icon(self, "simu_icon.svg")

        # 自适应屏幕大小
        screen_geo = QApplication.primaryScreen().availableGeometry()
        target_w = min(1200, int(screen_geo.width() * 0.85))
        target_h = min(800, int(screen_geo.height() * 0.75))
        self.resize(target_w, target_h)
        self.setWindowFlags(
            self.windowFlags()
            | Qt.WindowMinimizeButtonHint
            | Qt.WindowMaximizeButtonHint
            | Qt.WindowCloseButtonHint
        )

        # 样式
        self.setStyleSheet("""
            QDialog {
                background-color: white;
                font-family: 'Microsoft YaHei';
            }
            QLabel {
                color: #333;
            }
            QPushButton {
                background-color: #1890ff;
                color: white;
                border: none;
                border-radius: 4px;
                padding: 6px 16px;
                font-size: 13px;
            }
            QPushButton:hover {
                background-color: #40a9ff;
            }
            QPushButton:pressed {
                background-color: #096dd9;
            }
        """)

        self.init_ui()
        self.load_chart()

    def init_ui(self):
        """初始化UI布局"""
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)

        # 顶部工具栏
        toolbar = self._create_toolbar()
        main_layout.addWidget(toolbar)

        # 主内容区（左右分割）
        content_splitter = QSplitter(Qt.Horizontal)
        content_splitter.setHandleWidth(2)
        content_splitter.setStyleSheet("QSplitter::handle { background-color: #d9d9d9; }")

        # 左侧图表区
        left_panel = self._create_chart_panel()
        content_splitter.addWidget(left_panel)

        # 右侧统计面板
        right_panel = self._create_stats_panel()
        content_splitter.addWidget(right_panel)

        # 设置初始比例 (70% : 30%)
        content_splitter.setSizes([int(self.width() * 0.7), int(self.width() * 0.3)])

        main_layout.addWidget(content_splitter)

    def _create_toolbar(self):
        """创建顶部工具栏"""
        toolbar = QWidget()
        toolbar.setStyleSheet("background-color: #fafafa; border-bottom: 1px solid #d9d9d9;")
        toolbar.setFixedHeight(50)

        layout = QHBoxLayout(toolbar)
        layout.setContentsMargins(16, 8, 16, 8)
        layout.setSpacing(12)

        # 标题
        title_label = QLabel("压力测试分析")
        title_label.setStyleSheet("font-size: 16px; font-weight: bold; color: #262626;")
        layout.addWidget(title_label)

        layout.addStretch()

        # 自变量选择下拉框
        var_label = QLabel("自变量:")
        var_label.setStyleSheet("font-size: 13px; color: #595959;")
        layout.addWidget(var_label)

        self.var_combo = QComboBox()
        self.var_combo.setMinimumWidth(150)
        self.var_combo.addItems(self.x_addrs)
        self.var_combo.currentIndexChanged.connect(self.on_var_changed)
        layout.addWidget(self.var_combo)

        # 导出按钮
        export_btn = QPushButton("导出结果")
        export_btn.clicked.connect(self.on_export_clicked)
        layout.addWidget(export_btn)

        return toolbar

    def _create_chart_panel(self):
        """创建左侧图表面板"""
        panel = QWidget()
        panel.setMinimumWidth(500)
        panel.setStyleSheet("background-color: #ffffff;")

        layout = QVBoxLayout(panel)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(12)

        # 图表标题
        self.chart_title = QLabel()
        self.chart_title.setStyleSheet("font-size: 14px; font-weight: bold; color: #262626;")
        self.chart_title.setAlignment(Qt.AlignCenter)
        layout.addWidget(self.chart_title)

        # WebView 图表容器
        self.webview = get_shared_webview(parent_widget=panel)
        self.webview.setMinimumHeight(400)
        layout.addWidget(self.webview, 1)

        return panel

    def _create_stats_panel(self):
        """创建右侧统计面板"""
        panel = QWidget()
        panel.setMinimumWidth(300)
        panel.setStyleSheet("background-color: #fafafa;")

        layout = QVBoxLayout(panel)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(12)

        # 面板标题
        title = QLabel("压力测试结果")
        title.setStyleSheet("font-size: 14px; font-weight: bold; color: #262626;")
        layout.addWidget(title)

        # 结果表格
        self.results_table = QTableWidget()
        self.results_table.setColumnCount(4)
        self.results_table.setHorizontalHeaderLabels(["场景", "X压力值", "Y变化", "Y变化率"])
        self.results_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.results_table.setAlternatingRowColors(True)
        self.results_table.setStyleSheet("""
            QTableWidget {
                background-color: white;
                border: 1px solid #d9d9d9;
                border-radius: 4px;
            }
            QHeaderView::section {
                background-color: #fafafa;
                padding: 6px;
                border: none;
                border-bottom: 1px solid #d9d9d9;
                font-weight: bold;
            }
        """)
        layout.addWidget(self.results_table, 1)

        self._populate_results_table()

        return panel

    def _populate_results_table(self):
        """填充结果表格"""
        if not self.x_addrs:
            return

        current_x = self.x_addrs[self.current_x_index]
        x_results = [r for r in self.results if r["x_addr"] == current_x]

        self.results_table.setRowCount(len(x_results))

        for row, r in enumerate(x_results):
            # 场景
            self.results_table.setItem(row, 0, QTableWidgetItem(r["direction"]))
            # X压力值
            self.results_table.setItem(row, 1, QTableWidgetItem(f"{r['x_stressed']:.4f}"))
            # Y变化
            self.results_table.setItem(row, 2, QTableWidgetItem(f"{r['delta_y']:.4f}"))
            # Y变化率
            pct_text = f"{r['delta_y_pct']:.2f}%"
            pct_item = QTableWidgetItem(pct_text)
            # 根据正负设置颜色
            if r['delta_y_pct'] < 0:
                pct_item.setForeground(Qt.red)
            elif r['delta_y_pct'] > 0:
                pct_item.setForeground(Qt.darkGreen)
            self.results_table.setItem(row, 3, pct_item)

    def load_chart(self):
        """加载因变量（Y）的分布图表，叠加压力测试后的数据"""
        self.chart_title.setText(f"{self.y_addr} 分布直方图（20区间）")

        try:
            import numpy as np

            # 使用已有的 y_cache（原始蒙特卡洛模拟数据）
            y_data = np.asarray(self.y_cache['data'])
            y_data = y_data[np.isfinite(y_data)]  # 过滤 NaN 和 Inf

            if len(y_data) == 0:
                self.webview.setHtml("<h3>没有有效数据</h3>")
                return

            # 计算原始数据的直方图
            y_min = float(np.min(y_data))
            y_max = float(np.max(y_data))
            n_bins = 20
            bin_edges = np.linspace(y_min, y_max, n_bins + 1)
            counts, _ = np.histogram(y_data, bins=bin_edges)
            bin_centers = (bin_edges[:-1] + bin_edges[1:]) / 2
            total_samples = len(y_data)
            percentages = (counts / total_samples) * 100.0

            # 准备图表数据组 - 第一组：原始蒙特卡洛数据
            chart_groups = [{
                "name": f"{self.y_addr} (原始)",
                "x": bin_centers.tolist(),
                "y": percentages.tolist(),
                "color": "rgba(78, 205, 196, 0.6)",  # 青色，透明度 0.6
            }]

            # 收集当前选中自变量的压力测试后的 Y 值
            if self.x_addrs:
                current_x = self.x_addrs[self.current_x_index]
                x_results = [r for r in self.results if r["x_addr"] == current_x]

                # 提取所有压力测试后的 y_stressed 值
                y_stressed_values = [r["y_stressed"] for r in x_results if not np.isnan(r["y_stressed"])]

                if y_stressed_values:
                    # 使用相同的 bin_edges 计算压力测试数据的直方图
                    stressed_counts, _ = np.histogram(y_stressed_values, bins=bin_edges)
                    stressed_percentages = (stressed_counts / len(y_stressed_values)) * 100.0

                    # 第二组：压力测试后的数据
                    chart_groups.append({
                        "name": f"{current_x} 压力测试",
                        "x": bin_centers.tolist(),
                        "y": stressed_percentages.tolist(),
                        "color": "rgba(255, 107, 107, 0.6)",  # 红色，透明度 0.6
                    })

            # 计算 x 轴范围和刻度
            x_range = (float(y_min), float(y_max))
            x_dtick = (y_max - y_min) / 10  # 10个主刻度

            # 使用 DriskChartFactory 生成图表
            DriskChartFactory.VALUE_AXIS_TITLE = "因变量值"
            DriskChartFactory.VALUE_AXIS_UNIT = ""

            chart_result = DriskChartFactory.build_discrete_bar(
                data_groups=chart_groups,
                x_range=x_range,
                x_dtick=x_dtick,
                y_max=None,
                note_annotation="",
                margins=(60, 40, 30, 50),
                show_overlay_cdf=False,
                cdf_data=None,
                forced_mag=None,
                label_overrides={
                    "chart_title": {"text": ""},
                    "axes": {
                        "x": {"title": "因变量值"},
                        "y": {"title": "占比 (%)"}
                    }
                },
                axis_numeric_flags={"x": True, "y": True}
            )

            # 生成 HTML
            plot_json = chart_result.get("plot_json", "")
            if plot_json:
                html_content = f"""
                <!DOCTYPE html>
                <html>
                <head>
                    <script src="https://cdn.plot.ly/plotly-2.18.0.min.js"></script>
                </head>
                <body style="margin:0;padding:0;">
                    <div id="chart" style="width:100%;height:100%;"></div>
                    <script>
                        var plotData = {plot_json};
                        Plotly.newPlot('chart', plotData.data, plotData.layout, {{displayModeBar: false, responsive: true}});
                    </script>
                </body>
                </html>
                """
                self.webview.setHtml(html_content)

        except Exception as e:
            import traceback
            error_html = f"<h3>加载图表失败</h3><pre>{traceback.format_exc()}</pre>"
            self.webview.setHtml(error_html)


    def on_var_changed(self, index):
        """自变量选择变化"""
        self.current_x_index = index
        self.load_chart()
        self._populate_results_table()

    def on_export_clicked(self):
        """导出结果"""
        from pyxll import xlcAlert
        xlcAlert("导出功能开发中...")


def show_stress_test_dialog(results: list[dict], y_addr: str, y_cache: dict):
    """
    显示压力测试分析主窗口。

    参数：
        results: 压力测试结果列表
        y_addr: 因变量地址
        y_cache: 因变量缓存数据
    """
    try:
        app = QApplication.instance()
        if app is None:
            app = QApplication([])

        dialog = StressTestDialog(results, y_addr, y_cache)
        dialog.exec()

    except Exception as e:
        import traceback
        print(f"显示压力测试分析窗口失败: {e}\n{traceback.format_exc()}")
