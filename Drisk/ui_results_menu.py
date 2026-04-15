# ui_results_menu.py
"""
本模块提供结果视图（Results Views）的下拉菜单路由服务 ResultsMenuRouter。

主要功能：
1. UI 解耦：集中管理所有顶部控制栏按钮弹出的下拉菜单（QMenu）的构建与渲染。
2. 动作路由（Action Routing）：将用户的菜单点击事件安全地绑定并下发到主视图对象的内部方法（如 `_switch_tornado_mode`）。
3. 状态回退（State Fallback）：处理用户点击了弹出菜单但最终放弃选择（点击空白处取消）的边缘情况，自动重置触发按钮的按下状态（Checked State）。
4. 动态菜单构建：根据当前数据的特征（如连续/离散、单选/多选叠加），动态决定菜单中展示哪些图表视图选项。
"""

from PySide6.QtWidgets import QMenu

from ui_shared import drisk_combobox_qss

# =======================================================
# 1. 结果视图菜单路由服务
# =======================================================
class ResultsMenuRouter:
    """
    为结果对话框的顶部控制按钮构建并路由菜单动作。
    采用静态方法设计，通过传入 dialog 实例作为上下文来控制其 UI 状态。
    """

    # =======================================================
    # 2. 高级分析模式菜单 (龙卷风、箱线图、情景分析)
    # =======================================================
    @staticmethod
    def show_tornado_menu(dialog):
        """
        构建并展示【龙卷风图 (敏感性分析)】的下拉菜单。
        包含基于不同统计算法（分箱、回归、相关系数、方差等）的子模式。
        """
        # [安全守卫] 高级模式通常要求当前的主分析对象是“输出变量 (Output)”。
        # 如果拦截器返回 False，则取消本次操作，并弹起被按下的触发按钮。
        if not dialog._check_and_prepare_advanced_mode():
            dialog.btn_tornado.setChecked(False)
            return
            
        menu = QMenu(dialog)
        menu.setStyleSheet(drisk_combobox_qss())
        
        # 绑定不同统计算法的路由指令
        menu.addAction('飓风图 – 统计量变化').triggered.connect(lambda: dialog._switch_tornado_mode('bins'))
        menu.addAction('飓风图 – 回归系数').triggered.connect(lambda: dialog._switch_tornado_mode('reg'))
        menu.addAction('飓风图 – 回归映射值').triggered.connect(lambda: dialog._switch_tornado_mode('reg_mapped'))
        menu.addAction('飓风图 – 秩相关系数').triggered.connect(lambda: dialog._switch_tornado_mode('spearman'))
        menu.addAction('飓风图 – 方差贡献度').triggered.connect(lambda: dialog._switch_tornado_mode('variance'))
        menu.addAction('蜘蛛图 – 统计量变化').triggered.connect(lambda: dialog._switch_tornado_mode('bins_line'))
        
        # 将菜单弹出位置对齐到触发按钮的左下角
        menu.exec(dialog.btn_tornado.mapToGlobal(dialog.btn_tornado.rect().bottomLeft()))
        
        # [状态回退] 如果菜单执行完毕（含用户点击空白处取消），但主视图并未成功进入 'tornado' 模式，
        # 则说明没有发生实质性切换，需强制弹起按钮，保持 UI 视觉与底层状态一致。
        if getattr(dialog, '_current_analysis_mode', '') != 'tornado':
            dialog.btn_tornado.setChecked(False)

    @staticmethod
    def show_boxplot_menu(dialog):
        """构建并展示【多维箱线图】的下拉菜单。"""
        # 复用相同的安全守卫，防止在不支持的状态下进入高级 UI
        if not dialog._check_and_prepare_advanced_mode():
            dialog.btn_boxplot.setChecked(False)
            return
            
        menu = QMenu(dialog)
        menu.setStyleSheet(drisk_combobox_qss())
        
        menu.addAction('箱形图').triggered.connect(lambda: dialog._switch_boxplot_mode('boxplot'))
        menu.addAction('字母值图').triggered.connect(lambda: dialog._switch_boxplot_mode('letter_value'))
        menu.addAction('小提琴图').triggered.connect(lambda: dialog._switch_boxplot_mode('violin'))
        menu.addAction('趋势图').triggered.connect(lambda: dialog._switch_boxplot_mode('trend'))
        
        menu.exec(dialog.btn_boxplot.mapToGlobal(dialog.btn_boxplot.rect().bottomLeft()))
        
        # 验证当前模式是否属于箱线图家族，否则重置按钮状态
        if getattr(dialog, '_current_analysis_mode', '') not in ['boxplot', 'letter_value', 'violin', 'trend']:
            dialog.btn_boxplot.setChecked(False)

    @staticmethod
    def show_scenario_menu(dialog):
        """
        构建并展示【情景分析图 (尾部风险过滤)】的下拉菜单。
        提供预设的尾部概率区间，或唤起自定义设置面板。
        """
        if not dialog._check_and_prepare_advanced_mode():
            dialog.btn_scenario.setChecked(False)
            return
            
        menu = QMenu(dialog)
        menu.setStyleSheet(drisk_combobox_qss())
        
        # 绑定预设概率区间的触发指令
        menu.addAction('情景一：75% ~ 100%').triggered.connect(lambda: dialog._run_scenario(75.0, 100.0))
        menu.addAction('情景二：0% ~ 25%').triggered.connect(lambda: dialog._run_scenario(0.0, 25.0))
        menu.addAction('情景三：90% ~ 100%').triggered.connect(lambda: dialog._run_scenario(90.0, 100.0))
        menu.addSeparator()
        # 唤起独立的参数配置对话框
        menu.addAction('自定义情景').triggered.connect(lambda: dialog._open_scenario_settings(is_custom=True))
        
        menu.exec(dialog.btn_scenario.mapToGlobal(dialog.btn_scenario.rect().bottomLeft()))
        
        if getattr(dialog, '_current_analysis_mode', '') != 'scenario':
            dialog.btn_scenario.setChecked(False)

    # =======================================================
    # 3. 基础视图模式切换菜单 (直方图/CDF/离散图等)
    # =======================================================
    @staticmethod
    def show_view_mode_menu(dialog):
        """
        构建并展示【图表视图模式】下拉菜单。
        [动态路由] 这是最复杂的一个菜单，它需要根据当前数据是“连续(Continuous)”还是“离散(Discrete)”，
        以及当前是“单曲线”还是“多曲线叠加”，动态隐藏或显示某些不支持的渲染组合。
        """
        menu = QMenu(dialog)
        # 为基础视图菜单配置专属样式，覆盖默认的高亮蓝背景
        menu.setStyleSheet('QMenu { background: white; border: 1px solid #d9d9d9; } QMenu::item:selected { background: #e6f7ff; color: #0050b3; }')
        
        # 获取当前运行时的上下文标识
        is_discrete = bool(getattr(dialog, 'is_discrete_view', False))
        display_keys = list(getattr(dialog, 'display_keys', [dialog.current_key]))
        is_multiple = len(display_keys) > 1
        
        # 根据离散或连续数据类型，生成两套截然不同的菜单树 (Menu Tree)
        if is_discrete:
            # ---------------------------------------------------------
            # 离散模式菜单：屏蔽平滑的核密度(KDE)相关选项，主打柱状图
            # ---------------------------------------------------------
            menu.addAction('自动选择').triggered.connect(lambda: dialog._on_view_cmd_triggered('自动选择', 'auto'))
            
            m_disc = menu.addMenu('离散概率')
            m_disc.addAction('不叠加').triggered.connect(lambda: dialog._on_view_cmd_triggered('离散概率', 'discrete'))
            m_disc.addAction('叠加累积概率').triggered.connect(lambda: dialog._on_view_cmd_triggered('离散概率', 'discrete_cdf'))
            
            m_cdf = menu.addMenu('累积概率')
            m_cdf.addAction('不叠加').triggered.connect(lambda: dialog._on_view_cmd_triggered('累积概率', 'cdf'))
            m_cdf.addAction('叠加置信区间').triggered.connect(lambda: dialog._on_view_cmd_triggered('累积概率', 'cdf_dkw'))
            
        else:
            # ---------------------------------------------------------
            # 连续模式菜单：提供直方图与 KDE 曲线的各种叠加组合
            # ---------------------------------------------------------
            menu.addAction('自动选择').triggered.connect(lambda: dialog._on_view_cmd_triggered('自动选择', 'auto'))
            
            m_pdf = menu.addMenu('概率密度')
            m_pdf.addAction('不叠加').triggered.connect(lambda: dialog._on_view_cmd_triggered('概率密度', 'histogram'))
            
            # [冲突防范] 若存在多个情景叠加，在直方图上再叠加 KDE 曲线会导致画面极其混乱（线条穿模），因此隐藏。
            if not is_multiple:
                m_pdf.addAction('叠加核密度').triggered.connect(lambda: dialog._on_view_cmd_triggered('概率密度', 'histogram_kde'))
                
            m_pdf.addAction('叠加累积概率').triggered.connect(lambda: dialog._on_view_cmd_triggered('概率密度', 'histogram_cdf'))
            
            if not is_multiple:
                m_pdf.addAction('全部叠加').triggered.connect(lambda: dialog._on_view_cmd_triggered('概率密度', 'histogram_all'))
                
            menu.addAction('相对频率').triggered.connect(lambda: dialog._on_view_cmd_triggered('相对频率', 'relfreq'))
            
            m_cdf = menu.addMenu('累积概率')
            m_cdf.addAction('不叠加').triggered.connect(lambda: dialog._on_view_cmd_triggered('累积概率', 'cdf'))
            m_cdf.addAction('叠加置信区间').triggered.connect(lambda: dialog._on_view_cmd_triggered('累积概率', 'cdf_dkw'))
            
            menu.addAction('核密度').triggered.connect(lambda: dialog._on_view_cmd_triggered('核密度', 'pdfcurve'))
            
        # 唤起菜单，由于它是一个独立的 ToolButton，不需要像前面那样管理 CheckState
        menu.exec(dialog.btn_view_mode.mapToGlobal(dialog.btn_view_mode.rect().bottomLeft()))