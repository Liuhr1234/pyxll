# ui_results_dialogs.py
"""
本模块提供结果视图（Results Views）中独立弹出的高级对话框组件。

一、模块职责
1. 分析对象选择对话框（AnalysisObjectDialog）
   用于在结果分析界面切换当前“主分析对象”。
   用户可以在输入变量 / 输出变量之间切换，并进一步选择变量及其所属情景。

2. 高级叠加对话框（AdvancedOverlayDialog）
   用于为当前图表添加叠加对象。
   支持多选、批量加入、拖拽排序、上下移动、移除，以及对情景显示名进行重命名。

二、设计目的
本模块从 ui_results.py 中拆出，目的不是改变业务逻辑，而是降低主结果窗口的复杂度，
将“独立弹窗层”的职责单独收拢，便于维护、调试和后续交接。

三、核心依赖
1. 输入变量暴露控制：
   依赖 is_input_key_exposed(...) 判断某个输入变量是否允许在界面上暴露给用户。

2. 变量显示名构造：
   输入变量依赖 ui_shared / ui_variable_search 中的辅助函数，补齐变量名、单元格地址与搜索文本。
   输出变量优先使用 output_attributes 中记录的显式名称。

3. 情景显示名管理：
   依赖 ui_sim_display_names 提供的接口，对情景（Sim）显示名进行读取、修改和恢复默认。

四、数据约定
1. 当前主对象：
   由 (kind, sid, ck, lbl) 共同描述：
   - kind：变量类别，'input' 或 'output'
   - sid：情景 ID
   - ck：缓存键（Cache Key）
   - lbl：界面显示名称

2. 叠加项：
   内部通常保存为 (sid, ck, var_name, kind) 元组。
   其中旧版本历史数据可能只保存前三项，因此代码中保留了兼容逻辑。
"""

import os

from input_sample_exposure import is_input_key_exposed

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QAbstractItemView,
    QCheckBox,
    QComboBox,
    QDialog,
    QHBoxLayout,
    QInputDialog,
    QLabel,
    QListWidget,
    QListWidgetItem,
    QMenu,
    QPushButton,
    QVBoxLayout,
)

from ui_shared import (
    DRISK_DIALOG_BTN_QSS,
    extract_clean_cell_address,
    resolve_visible_variable_name,
)
from ui_variable_search import (
    SEARCH_TEXT_ROLE,
    build_input_reference_display_map,
    build_input_variable_label,
    refresh_combo_dropdown_filter,
    resolve_search_text,
    setup_searchable_combo,
    fuzzy_contains_case_insensitive,
)
from ui_sim_display_names import (
    get_custom_sim_display_name,
    get_sim_display_name,
    set_sim_display_name,
)


# =======================================================
# 0. 模块级辅助函数
# =======================================================
def _normalize_output_cache_key(cell_key):
    """
    规范化输出变量的缓存键。

    目的：
    不同情景下，同一个输出变量的缓存键可能在格式上存在细微差异，
    例如包含 / 不包含美元符号，大小写不一致，或前后带空格。
    为了实现跨情景匹配，需要先进行统一标准化处理。

    处理规则：
    1. 空值转为空字符串；
    2. 去除美元符号；
    3. 去除首尾空白；
    4. 统一转为大写。
    """
    return str(cell_key or "").replace("$", "").strip().upper()


def _collect_output_display_names(all_sims):
    """
    汇总所有情景中的输出变量显示名，并建立“规范化缓存键 -> 显示名”的映射表。

    设计意图：
    同一输出变量在多个情景中通常对应同一个业务含义，因此界面层需要优先使用
    一致的、用户可读的名称，而不是直接显示原始缓存键。

    命名优先级：
    1. 优先采用 output_attributes 中明确给出的 name；
    2. 若多个情景都存在 name，则优先采用 _name_from_attr=True 的名称；
    3. 若没有可用名称，则后续由调用方回退到缓存键本身。

    返回：
    dict[str, str]
    键为规范化后的缓存键，值为跨情景统一后的显示名。
    """
    preferred_name_map = {}

    for _sid, sim in (all_sims or {}).items():
        cache = getattr(sim, "output_cache", {}) or {}
        attrs_cache = getattr(sim, "output_attributes", {}) or {}

        for ck in cache.keys():
            attrs = attrs_cache.get(ck, {}) or {}
            name = str(attrs.get("name", "") or "").strip()
            if not name:
                continue

            norm_ck = _normalize_output_cache_key(ck)
            from_attr = bool(attrs.get("_name_from_attr", False))
            existing = preferred_name_map.get(norm_ck)

            if existing is None or (from_attr and not existing[1]):
                preferred_name_map[norm_ck] = (name, from_attr)

    return {k: v[0] for k, v in preferred_name_map.items()}


def _looks_like_raw_input_key_name(name_text, cell_key) -> bool:
    """
    判断输入变量名称是否实际上只是“原始单元格键名”的直接回填结果。

    背景：
    某些输入变量并没有真正的人类可读名称，界面层可能会把单元格地址、
    sheet!cell 形式的缓存键，甚至 makeinput 生成的技术性名称，误当作变量名显示。

    本函数用于识别这类“伪名称”，便于后续在显示层主动清空 name，
    再重新走一遍更稳妥的可见名称解析逻辑。

    判定规则：
    1. name 与 cell_key 完全一致；
    2. name 与 cell_key 去掉工作表前缀后的尾部一致；
    3. 提取净单元格地址后相同，且 name 中带有下划线或 MAKEINPUT 痕迹。
    """
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


def _resolve_input_visible_name(cell_key, attrs, *, excel_app=None) -> str:
    """
    为输入变量生成最终用于界面展示的可见名称。

    处理策略：
    1. 先读取 attrs 中的 name；
    2. 若该名称看起来只是技术性键名，则主动清空，避免误显示；
    3. 调用 resolve_visible_variable_name(...) 做标准解析；
    4. 若仍无结果，则回退到净单元格地址；
    5. 最后再回退到原始缓存键末尾部分。

    这样可以尽可能保证：
    - 优先显示用户真正可理解的变量名；
    - 实在没有变量名时，至少显示明确的单元格地址；
    - 避免把原始技术字段直接暴露给最终用户。
    """
    attrs_map = attrs if isinstance(attrs, dict) else {}
    safe_attrs = dict(attrs_map)

    raw_name = str(safe_attrs.get("name", "") or "").strip()
    if _looks_like_raw_input_key_name(raw_name, cell_key):
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


def _rename_sim_with_dialog(parent, sim_id) -> bool:
    """
    弹出输入框，对指定情景的界面显示名进行修改。

    说明：
    1. 这里只修改“界面显示名”，不改动底层情景 ID；
    2. 输入留空时，表示恢复默认显示名；
    3. 返回值表示显示名是否实际发生变化。

    返回：
    bool
    - True：显示名发生变化，调用方通常应刷新相关列表；
    - False：用户取消，或输入后与原显示效果一致。
    """
    current_custom = get_custom_sim_display_name(sim_id)
    current_effective = get_sim_display_name(sim_id)
    initial_text = current_custom if current_custom else current_effective

    text, ok = QInputDialog.getText(
        parent,
        "重命名情景",
        "输入用于绘图界面的情景显示名（留空可恢复默认）：",
        text=initial_text,
    )
    if not ok:
        return False

    before = get_sim_display_name(sim_id)
    after = set_sim_display_name(sim_id, text)
    return str(before) != str(after) or (
        not str(text or "").strip() and bool(current_custom)
    )


# =======================================================
# 1. 主分析对象选择对话框
# =======================================================
class AnalysisObjectDialog(QDialog):
    """
    单选对话框：用于在结果视图中切换当前聚焦的“主分析对象”。

    一、适用场景
    用户已经打开结果视图，希望把当前图中分析的对象切换为：
    1. 另一个输入变量；
    2. 另一个输出变量；
    3. 同一变量在不同情景下的结果。

    二、界面结构
    1. 顶部：变量类型切换（输入 / 输出）；
    2. 左侧：变量列表；
    3. 右侧：包含该变量的情景列表；
    4. 底部：确定 / 取消。

    三、交互特征
    1. 左右两列是级联关系；
    2. 顶部提供搜索框，支持变量名 / 地址模糊定位；
    3. 右侧情景列表支持右键重命名情景；
    4. 双击情景可直接确认。
    """

    # ---------------------------------------------------
    # 1.1 初始化与基础配置
    # ---------------------------------------------------
    def __init__(
        self,
        all_sims,
        current_kind,
        current_ck,
        current_sid,
        parent=None,
        lock_to_output=False,
    ):
        super().__init__(parent)

        self.setWindowTitle("Drisk - 选择分析对象")
        self.resize(460, 340)

        # 构造勾选状态图标的绝对路径，用于注入样式表。
        base_dir = os.path.dirname(os.path.abspath(__file__))
        icon_path = os.path.join(
            base_dir, "icons", "Selected_Overlay.svg"
        ).replace("\\", "/")

        # 对话框整体样式。
        # 这里保持与结果分析相关弹窗一致的视觉风格。
        self.setStyleSheet(
            "\n"
            "            QDialog { background-color: #f9f9f9; font-family: 'Microsoft YaHei'; }\n"
            "            QLabel { font-weight: bold; color: #444; font-size: 12px; }\n"
            "            QListWidget {\n"
            "                border: 1px solid #ccc;\n"
            "                border-radius: 3px;\n"
            "                outline: none;\n"
            "                font-size: 12px;\n"
            "                background-color: #ffffff;\n"
            "            }\n"
            "            QListWidget::item { padding: 6px; border-bottom: 1px solid #f5f5f5; }\n"
            "            QListWidget::item:hover { background-color: #f0f7ff; }\n"
            "            QPushButton {\n"
            "                background-color: white; border: 1px solid #ccc; border-radius: 3px;\n"
            "                padding: 4px 12px; font-size: 12px;\n"
            "            }\n"
            "            QPushButton:hover { color: #40a9ff; border-color: #40a9ff; background-color: #f0f7ff;}\n"
            "            QPushButton#btnOk { background-color: #0050b3; color: white; border: none; min-width: 70px; }\n"
            "            QCheckBox { font-size: 12px; color: #333; }\n"
            "\n"
            "            QListWidget::indicator, QCheckBox::indicator {\n"
            "                width: 16px;\n"
            "                height: 16px;\n"
            "                border-radius: 3px;\n"
            "                border: 1px solid #cbd0d6;\n"
            "                background-color: #ffffff;\n"
            "            }\n"
            "\n"
            "            QListWidget::indicator:hover, QCheckBox::indicator:hover {\n"
            "                border: 1px solid #8c9eb5;\n"
            "            }\n"
            "\n"
            "            QListWidget::indicator:checked, QCheckBox::indicator:checked {\n"
            "                background-color: #8c9eb5;\n"
            "                border: 1px solid #8c9eb5;\n"
            "                image: url('"
            + icon_path
            + "');\n"
            "            }\n"
            "\n"
            "            QListWidget::item:selected {\n"
            "                background-color: #f0f4f8;\n"
            "                color: #334455;\n"
            "                font-weight: bold;\n"
            "                border-left: 3px solid #8c9eb5;\n"
            "            }\n"
            "        "
        )

        # -----------------------------
        # 运行时上下文
        # -----------------------------
        self.all_sims = all_sims
        self.current_kind = current_kind
        self.current_ck = current_ck
        self.current_sid = current_sid
        self.lock_to_output = lock_to_output

        # 预构建输入 / 输出变量映射，减少界面切换时的重复计算。
        self.input_map = self._build_map("input")
        self.output_map = self._build_map("output")

        # 构建界面。
        self.init_ui()

    # ---------------------------------------------------
    # 1.2 数据准备：构建变量映射
    # ---------------------------------------------------
    def _build_map(self, kind):
        """
        构建“显示标签 -> [(sid, ck), ...]”的变量映射表。

        说明：
        该映射是整个对话框的基础数据源。
        对话框左侧显示的是变量标签，右侧显示的是该变量在哪些情景中存在。

        输入变量与输出变量的处理差异：
        1. 输出变量：
           - 直接遍历 output_cache；
           - 优先使用 output_attributes 中的显式名称；
           - 显示格式为“变量名 (缓存键)”。

        2. 输入变量：
           - 仅纳入 is_input_key_exposed(...) 允许暴露的项；
           - 会构建单元格引用显示映射；
           - 会借助 Excel 上下文补齐更友好的显示名称。

        同时会过滤以下无效项：
        1. 数据数组为空；
        2. 首个值是错误标记（如 #NUM!）；
        3. 首个值包含 NAN；
        4. 超出 max_scenarios 限制的情景。

        返回：
        dict[str, list[tuple]]
        例如：
        {
            "销售额 (Sheet1!B8)": [(1, "Sheet1!B8"), (2, "Sheet1!B8")],
            "成本率 [B12]": [(1, "Sheet1!B12")]
        }
        """
        var_map = {}
        input_ref_map = {}
        output_name_map = (
            _collect_output_display_names(self.all_sims) if kind == "output" else {}
        )
        excel_app = None

        if kind == "input":
            all_input_keys = []
            for _sid, _sim in self.all_sims.items():
                _cache = getattr(_sim, "input_cache", {}) or {}
                all_input_keys.extend(
                    [k for k in _cache.keys() if is_input_key_exposed(_sim, k)]
                )

            input_ref_map = build_input_reference_display_map(all_input_keys)

            try:
                from com_fixer import _safe_excel_app

                excel_app = _safe_excel_app()
            except Exception:
                excel_app = None

        for sid, sim in self.all_sims.items():
            cache = sim.output_cache if kind == "output" else getattr(sim, "input_cache", {})
            attrs_cache = (
                sim.output_attributes
                if kind == "output"
                else getattr(sim, "input_attributes", {})
            )

            for ck, data in cache.items():
                if kind == "input" and not is_input_key_exposed(sim, ck):
                    continue
                if len(data) == 0:
                    continue

                # 用首个值做快速有效性预检。
                v = str(data[0]).strip().upper()
                if v.startswith("#") or "NAN" in v:
                    continue

                pure_ck = ck.split("!")[-1] if "!" in ck else ck
                attrs = attrs_cache.get(ck, {}) if isinstance(attrs_cache, dict) else {}

                if kind == "output":
                    norm_ck = _normalize_output_cache_key(ck)
                    base_name = output_name_map.get(norm_ck, attrs.get("name", ""))
                    var_name = base_name if base_name else pure_ck
                    lbl = f"{var_name} ({ck})"
                else:
                    display_ref = input_ref_map.get(str(ck), pure_ck)
                    visible_name = _resolve_input_visible_name(
                        ck,
                        attrs,
                        excel_app=excel_app,
                    )
                    lbl = build_input_variable_label(visible_name, display_ref)

                if lbl not in var_map:
                    var_map[lbl] = []

                # max_scenarios 用于控制某变量实际允许参与展示的情景数。
                if sid <= attrs_cache.get(ck, {}).get("max_scenarios", 999):
                    var_map[lbl].append((sid, ck))

        return {k: v for k, v in var_map.items() if len(v) > 0}

    # ---------------------------------------------------
    # 1.3 界面构建
    # ---------------------------------------------------
    def init_ui(self):
        """
        构建分析对象选择对话框的完整界面。

        布局结构：
        1. 顶部变量类别切换；
        2. 中部左右分栏：
           - 左：变量搜索 + 变量列表
           - 右：情景列表
        3. 底部操作按钮。
        """
        layout = QVBoxLayout(self)
        layout.setSpacing(10)

        # 顶部：变量类型切换。
        top_layout = QHBoxLayout()
        self.lbl_kind = QLabel("变量类型:")
        self.combo_kind = QComboBox()
        self.combo_kind.addItems(["输入变量", "输出变量"])
        top_layout.addWidget(self.lbl_kind)
        top_layout.addWidget(self.combo_kind)
        top_layout.addStretch()
        layout.addLayout(top_layout)

        # 中部：左侧变量列表，右侧情景列表。
        list_layout = QHBoxLayout()

        left_vbox = QVBoxLayout()
        left_vbox.addWidget(QLabel("变量列表:"))
        self.combo_var_search = QComboBox()
        setup_searchable_combo(self.combo_var_search, "搜索变量名或地址（例如 Sheet1!B1）")
        left_vbox.addWidget(self.combo_var_search)
        self.list_vars = QListWidget()
        left_vbox.addWidget(self.list_vars)

        right_vbox = QVBoxLayout()
        right_vbox.addWidget(QLabel("包含的数据（情景 ID）:"))
        self.list_sims = QListWidget()
        self.list_sims.setContextMenuPolicy(Qt.CustomContextMenu)
        self.list_sims.customContextMenuRequested.connect(self._open_sim_context_menu)
        right_vbox.addWidget(self.list_sims)

        # 左右区域比例：3:2。
        list_layout.addLayout(left_vbox, stretch=3)
        list_layout.addLayout(right_vbox, stretch=2)
        layout.addLayout(list_layout)

        # 追加统一按钮样式。
        self.setStyleSheet(self.styleSheet() + DRISK_DIALOG_BTN_QSS)

        # 底部：确定 / 取消。
        btn_layout = QHBoxLayout()
        btn_layout.addStretch()

        self.btn_ok = QPushButton("确定")
        self.btn_ok.setObjectName("btnOk")
        self.btn_ok.clicked.connect(self.accept)
        self.btn_ok.setEnabled(False)

        btn_cancel = QPushButton("取消")
        btn_cancel.setObjectName("btnCancel")
        btn_cancel.clicked.connect(self.reject)

        self.btn_ok.setFixedSize(80, 28)
        btn_cancel.setFixedSize(80, 28)
        btn_layout.addWidget(self.btn_ok)
        btn_layout.addWidget(btn_cancel)
        layout.addLayout(btn_layout)

        # -----------------------------
        # 信号绑定
        # -----------------------------
        self.combo_kind.currentIndexChanged.connect(self._on_kind_changed)
        self.combo_var_search.currentIndexChanged.connect(self._on_var_search_picked)

        if self.combo_var_search.lineEdit() is not None:
            self.combo_var_search.lineEdit().textChanged.connect(
                self._on_var_search_text_changed
            )

        self.list_vars.currentItemChanged.connect(self._on_var_changed)
        self.list_sims.currentItemChanged.connect(self._on_sim_changed)
        self.list_sims.itemDoubleClicked.connect(self.accept)

        # -----------------------------
        # 初始状态同步
        # -----------------------------
        idx = 0 if self.current_kind == "input" else 1
        self.combo_kind.setCurrentIndex(idx)
        self._on_kind_changed(idx)

        # 某些图表仅允许输出变量，此时隐藏变量类别切换控件。
        if self.lock_to_output:
            self.combo_kind.setCurrentIndex(1)
            self.lbl_kind.hide()
            self.combo_kind.hide()

    # ---------------------------------------------------
    # 1.4 变量搜索与列表同步
    # ---------------------------------------------------
    def _rebuild_var_search_options(self, labels, current_map):
        """
        重建顶部搜索下拉框的数据源。

        说明：
        搜索框与左侧变量列表共用同一批变量标签。
        每一项除了显示文本外，还额外挂载 SEARCH_TEXT_ROLE，
        以便支持“变量名 + 地址”的综合模糊搜索。
        """
        self.combo_var_search.blockSignals(True)
        self.combo_var_search.clear()

        for lbl in labels:
            self.combo_var_search.addItem(lbl, lbl)
            idx = self.combo_var_search.count() - 1
            refs = " ".join(
                dict.fromkeys(str(ck) for _, ck in current_map.get(lbl, []))
            )
            self.combo_var_search.setItemData(
                idx,
                resolve_search_text(lbl, refs),
                SEARCH_TEXT_ROLE,
            )

        self.combo_var_search.setCurrentIndex(-1)
        self.combo_var_search.blockSignals(False)
        refresh_combo_dropdown_filter(self.combo_var_search, "")

    def _select_var_item_by_label(self, label):
        """
        根据变量标签，在左侧变量列表中定位并选中对应项。
        """
        if not label:
            return

        for i in range(self.list_vars.count()):
            item = self.list_vars.item(i)
            item_label = item.data(Qt.UserRole) or item.text()
            if str(item_label) == str(label) and not item.isHidden():
                self.list_vars.setCurrentItem(item)
                self.list_vars.scrollToItem(item)
                return

    def _on_var_search_picked(self, idx):
        """
        当用户从搜索下拉框中直接选择某一项时，
        同步选中左侧变量列表中的对应变量。
        """
        if idx < 0:
            return

        label = self.combo_var_search.itemData(idx, Qt.UserRole)
        if not label:
            label = self.combo_var_search.itemText(idx)

        self._select_var_item_by_label(label)

    def _on_var_search_text_changed(self, text):
        """
        当搜索框文本变化时：
        1. 刷新搜索下拉候选；
        2. 过滤左侧变量列表的显示项；
        3. 自动把当前选中项切换到第一个可见匹配项；
        4. 若没有任何匹配，则清空右侧情景列表并禁用确定按钮。
        """
        query = str(text or "")
        refresh_combo_dropdown_filter(self.combo_var_search, query)

        first_visible = None
        for i in range(self.list_vars.count()):
            item = self.list_vars.item(i)
            label = item.data(Qt.UserRole) or item.text()
            search_text = item.data(SEARCH_TEXT_ROLE) or resolve_search_text(label)
            matched = fuzzy_contains_case_insensitive(query, search_text)
            item.setHidden(not matched)

            if matched and first_visible is None:
                first_visible = item

        current = self.list_vars.currentItem()
        if first_visible is None:
            self.list_vars.setCurrentRow(-1)
            self.list_sims.clear()
            self.btn_ok.setEnabled(False)
            return

        if first_visible is not None and (current is None or current.isHidden()):
            self.list_vars.setCurrentItem(first_visible)

    # ---------------------------------------------------
    # 1.5 左右级联联动
    # ---------------------------------------------------
    def _on_kind_changed(self, idx):
        """
        当变量类别（输入 / 输出）切换时，重建左侧变量候选列表。

        处理步骤：
        1. 清空现有变量与情景列表；
        2. 根据当前类别切换到 input_map 或 output_map；
        3. 重建搜索框数据源；
        4. 重建左侧变量列表；
        5. 若当前主对象仍在列表中，则优先定位到它；
        6. 否则默认选择第一项。
        """
        self.list_vars.clear()
        self.list_sims.clear()
        self.btn_ok.setEnabled(False)

        kind = "input" if idx == 0 else "output"
        current_map = self.input_map if kind == "input" else self.output_map
        labels = sorted(current_map.keys())

        self._rebuild_var_search_options(labels, current_map)
        target_item = None

        for lbl in labels:
            item = QListWidgetItem(lbl)
            item.setData(Qt.UserRole, lbl)

            refs = " ".join(
                dict.fromkeys(str(ck) for _, ck in current_map.get(lbl, []))
            )
            item.setData(SEARCH_TEXT_ROLE, resolve_search_text(lbl, refs))
            self.list_vars.addItem(item)

            if kind == self.current_kind and any(
                (c == self.current_ck for s, c in current_map[lbl])
            ):
                target_item = item

        if target_item is not None:
            self.list_vars.setCurrentItem(target_item)
        elif self.list_vars.count() > 0:
            self.list_vars.setCurrentRow(0)

        if self.combo_var_search.lineEdit() is not None:
            self.combo_var_search.lineEdit().clear()

        self._on_var_search_text_changed("")

    def _on_var_changed(self, current, previous):
        """
        当左侧变量选择变化时，刷新右侧情景列表。

        右侧每一项都写入完整定位信息：
        (sid, ck, lbl, kind)

        其中：
        - sid：情景 ID
        - ck：缓存键
        - lbl：变量显示标签
        - kind：变量类别（input / output）

        额外处理：
        若某项正是当前主对象，则在文字后附加“(当前)”标记，并加粗显示。
        """
        self.list_sims.clear()
        if not current:
            return

        lbl = current.data(Qt.UserRole)
        kind = "input" if self.combo_kind.currentIndex() == 0 else "output"
        current_map = self.input_map if kind == "input" else self.output_map

        preferred_item = None
        first_item = None

        for sid, ck in current_map.get(lbl, []):
            sim_label = get_sim_display_name(sid)
            item = QListWidgetItem(sim_label)
            item.setData(Qt.UserRole, (sid, ck, lbl, kind))

            is_current = (
                kind == self.current_kind
                and sid == self.current_sid
                and (ck == self.current_ck)
            )

            if is_current:
                item.setText(f"{sim_label} (当前)")
                font = item.font()
                font.setBold(True)
                item.setFont(font)
                preferred_item = item

            self.list_sims.addItem(item)

            if first_item is None:
                first_item = item

        # 保证右侧始终存在一个有效选择，便于直接确认。
        target_item = preferred_item or first_item
        if target_item is not None:
            self.list_sims.setCurrentItem(target_item)

    def _on_sim_changed(self, current, previous):
        """
        只有右侧已经选定某个具体情景后，才允许点击“确定”。
        """
        self.btn_ok.setEnabled(current is not None)

    # ---------------------------------------------------
    # 1.6 右键菜单：情景重命名
    # ---------------------------------------------------
    def _open_sim_context_menu(self, pos):
        """
        在右侧情景列表上打开右键菜单。

        当前仅提供“重命名情景”操作。
        若名称修改成功，则刷新当前变量下的情景列表显示。
        """
        item = self.list_sims.itemAt(pos)
        if item is None:
            return

        data = item.data(Qt.UserRole)
        sid = data[0] if isinstance(data, (tuple, list)) and len(data) > 0 else None
        if sid is None:
            return

        menu = QMenu(self)
        act_rename = menu.addAction("重命名情景")
        chosen = menu.exec(self.list_sims.mapToGlobal(pos))
        if chosen != act_rename:
            return

        if _rename_sim_with_dialog(self, sid):
            self._on_var_changed(self.list_vars.currentItem(), None)

    # ---------------------------------------------------
    # 1.7 对外结果接口
    # ---------------------------------------------------
    def get_selection(self):
        """
        返回用户最终选择的主分析对象信息。

        返回格式：
        (kind, sid, ck, lbl)

        若用户尚未形成有效选择，则返回 None。
        """
        kind = "input" if self.combo_kind.currentIndex() == 0 else "output"
        sim_item = self.list_sims.currentItem()
        if not sim_item:
            return None

        data = sim_item.data(Qt.UserRole)
        if not data:
            return None

        sid = data[0] if len(data) > 0 else None
        ck = data[1] if len(data) > 1 else ""
        lbl = data[2] if len(data) > 2 else ""
        return (kind, sid, ck, lbl)


# =======================================================
# 2. 高级叠加变量选择对话框（支持多选与排序）
# =======================================================
class AdvancedOverlayDialog(QDialog):
    """
    多选与排序对话框：用于管理图表中的叠加对象列表。

    一、适用场景
    用户希望在当前主图之上叠加更多变量 / 更多情景，以进行对比观察。

    二、界面结构
    1. 顶部：变量类别切换（输入 / 输出）；
    2. 中部：
       - 左：可选变量列表；
       - 右：当前变量在各情景下的候选项，可勾选；
    3. 中间按钮：加入叠加列表；
    4. 底部：
       - 已添加叠加项列表；
       - 上移 / 下移 / 移除 / 清空操作按钮；
       - 确定 / 取消。

    三、核心特性
    1. 防重：当前主对象不可重复加入；
    2. 防重：已经在叠加列表中的对象不可再次加入；
    3. 候选区支持全选；
    4. 已添加列表支持拖拽排序，也支持按钮式顺位调整；
    5. 支持对候选区和已添加区中的情景执行右键重命名。
    """

    # ---------------------------------------------------
    # 2.1 初始化与基础配置
    # ---------------------------------------------------
    def __init__(
        self,
        all_sims,
        current_kind,
        current_ck,
        current_sid,
        current_overlays,
        parent=None,
        lock_to_output=False,
    ):
        super().__init__(parent)

        self.setWindowTitle("Drisk - 添加叠加变量")
        self.resize(520, 520)

        # 复用统一的勾选图标样式策略。
        base_dir = os.path.dirname(os.path.abspath(__file__))
        icon_path = os.path.join(
            base_dir, "icons", "Selected_Overlay.svg"
        ).replace("\\", "/")

        self.setStyleSheet(
            "\n"
            "            QDialog { background-color: #f9f9f9; font-family: 'Microsoft YaHei'; }\n"
            "            QLabel { font-weight: bold; color: #444; font-size: 12px; }\n"
            "            QListWidget {\n"
            "                border: 1px solid #ccc;\n"
            "                border-radius: 3px;\n"
            "                outline: none;\n"
            "                font-size: 12px;\n"
            "                background-color: #ffffff;\n"
            "            }\n"
            "            QListWidget::item { padding: 6px; border-bottom: 1px solid #f5f5f5; }\n"
            "            QListWidget::item:hover { background-color: #f0f7ff; }\n"
            "            QPushButton {\n"
            "                background-color: white; border: 1px solid #ccc; border-radius: 3px;\n"
            "                padding: 4px 12px; font-size: 12px;\n"
            "            }\n"
            "            QPushButton:hover { color: #40a9ff; border-color: #40a9ff; background-color: #f0f7ff;}\n"
            "            QPushButton#btnOk { background-color: #0050b3; color: white; border: none; min-width: 70px; }\n"
            "            QCheckBox { font-size: 12px; color: #333; }\n"
            "\n"
            "            QListWidget::indicator, QCheckBox::indicator {\n"
            "                width: 16px;\n"
            "                height: 16px;\n"
            "                border-radius: 3px;\n"
            "                border: 1px solid #cbd0d6;\n"
            "                background-color: #ffffff;\n"
            "            }\n"
            "\n"
            "            QListWidget::indicator:hover, QCheckBox::indicator:hover {\n"
            "                border: 1px solid #8c9eb5;\n"
            "            }\n"
            "\n"
            "            QListWidget::indicator:checked, QCheckBox::indicator:checked {\n"
            "                background-color: #8c9eb5;\n"
            "                border: 1px solid #8c9eb5;\n"
            "                image: url('"
            + icon_path
            + "');\n"
            "            }\n"
            "\n"
            "            QListWidget::item:selected {\n"
            "                background-color: #f0f4f8;\n"
            "                color: #334455;\n"
            "                font-weight: bold;\n"
            "                border-left: 3px solid #8c9eb5;\n"
            "            }\n"
            "        "
        )

        self.all_sims = all_sims
        self.current_kind = current_kind
        self.current_ck = current_ck
        self.current_sid = current_sid

        # 内部真实结果列表。
        # 该列表与下方“已添加叠加项”列表保持同步，是最终返回给外部调用方的数据源。
        self.result_items = list(current_overlays)
        self.lock_to_output = lock_to_output

        # 预构建输入 / 输出变量映射。
        self.input_map = self._build_map("input")
        self.output_map = self._build_map("output")

        self.init_ui()

    # ---------------------------------------------------
    # 2.2 数据准备：构建变量映射
    # ---------------------------------------------------
    def _build_map(self, kind):
        """
        构建叠加对话框使用的变量映射表。

        该函数与 AnalysisObjectDialog._build_map(...) 的思路基本一致，
        目的是让两个弹窗在“变量可见范围、变量命名、无效数据过滤”上保持一致。

        返回：
        dict[str, list[tuple]]
        结构为：
        {
            "变量显示标签": [(sid, ck), ...]
        }
        """
        var_map = {}
        input_ref_map = {}
        output_name_map = (
            _collect_output_display_names(self.all_sims) if kind == "output" else {}
        )
        excel_app = None

        if kind == "input":
            all_input_keys = []
            for _sid, _sim in self.all_sims.items():
                _cache = getattr(_sim, "input_cache", {}) or {}
                all_input_keys.extend(
                    [k for k in _cache.keys() if is_input_key_exposed(_sim, k)]
                )

            input_ref_map = build_input_reference_display_map(all_input_keys)

            try:
                from com_fixer import _safe_excel_app

                excel_app = _safe_excel_app()
            except Exception:
                excel_app = None

        for sid, sim in self.all_sims.items():
            cache = sim.output_cache if kind == "output" else getattr(sim, "input_cache", {})
            attrs_cache = (
                sim.output_attributes
                if kind == "output"
                else getattr(sim, "input_attributes", {})
            )

            for ck, data in cache.items():
                if kind == "input" and not is_input_key_exposed(sim, ck):
                    continue
                if len(data) == 0:
                    continue

                v = str(data[0]).strip().upper()
                if v.startswith("#") or "NAN" in v:
                    continue

                pure_ck = ck.split("!")[-1] if "!" in ck else ck
                attrs = attrs_cache.get(ck, {}) if isinstance(attrs_cache, dict) else {}

                if kind == "output":
                    norm_ck = _normalize_output_cache_key(ck)
                    base_name = output_name_map.get(norm_ck, attrs.get("name", ""))
                    var_name = base_name if base_name else pure_ck
                    lbl = f"{var_name} ({ck})"
                else:
                    display_ref = input_ref_map.get(str(ck), pure_ck)
                    visible_name = _resolve_input_visible_name(
                        ck,
                        attrs,
                        excel_app=excel_app,
                    )
                    lbl = build_input_variable_label(visible_name, display_ref)

                if lbl not in var_map:
                    var_map[lbl] = []

                if sid <= attrs_cache.get(ck, {}).get("max_scenarios", 999):
                    var_map[lbl].append((sid, ck))

        return {k: v for k, v in var_map.items() if len(v) > 0}

    # ---------------------------------------------------
    # 2.3 界面构建
    # ---------------------------------------------------
    def init_ui(self):
        """
        构建高级叠加对话框的完整界面。

        布局结构：
        1. 顶部：变量类型切换；
        2. 中部：变量列表 + 候选情景列表；
        3. 中间：加入叠加列表按钮；
        4. 底部：已添加叠加项列表及顺位控制按钮；
        5. 最下方：确定 / 取消。
        """
        layout = QVBoxLayout(self)
        layout.setSpacing(10)

        # 顶部：变量类型切换。
        top_layout = QHBoxLayout()
        self.lbl_kind = QLabel("变量类型:")
        self.combo_kind = QComboBox()
        self.combo_kind.addItems(["输入变量", "输出变量"])
        top_layout.addWidget(self.lbl_kind)
        top_layout.addWidget(self.combo_kind)
        top_layout.addStretch()
        layout.addLayout(top_layout)

        # 中部：左边变量，右边候选情景。
        middle_layout = QHBoxLayout()

        left_vbox = QVBoxLayout()
        left_vbox.addWidget(QLabel("1. 选择对比变量:"))
        self.combo_var_search = QComboBox()
        setup_searchable_combo(self.combo_var_search, "搜索变量名或地址（例如 Sheet1!B1）")
        left_vbox.addWidget(self.combo_var_search)
        self.list_vars = QListWidget()
        left_vbox.addWidget(self.list_vars)

        right_vbox = QVBoxLayout()
        right_vbox.addWidget(QLabel("2. 勾选需要叠加的模拟情景:"))
        self.chk_all = QCheckBox("全选所有情景")
        right_vbox.addWidget(self.chk_all)
        self.list_sims = QListWidget()
        self.list_sims.setContextMenuPolicy(Qt.CustomContextMenu)
        self.list_sims.customContextMenuRequested.connect(
            self._open_candidate_sim_context_menu
        )
        right_vbox.addWidget(self.list_sims)

        middle_layout.addLayout(left_vbox, stretch=1)
        middle_layout.addLayout(right_vbox, stretch=1)
        layout.addLayout(middle_layout)

        # 中间：加入叠加列表。
        btn_add = QPushButton("↓ 添加到叠加列表 ↓")
        btn_add.setObjectName("btnAdd")
        btn_add.clicked.connect(self._add_to_overlay)
        layout.addWidget(btn_add)

        # 底部：已添加叠加项列表。
        layout.addWidget(QLabel("3. 当前叠加列表（可拖拽或箭头调整顺位）:"))
        self.list_added = QListWidget()
        self.list_added.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.list_added.setDragDropMode(QAbstractItemView.InternalMove)
        self.list_added.setContextMenuPolicy(Qt.CustomContextMenu)
        self.list_added.customContextMenuRequested.connect(
            self._open_added_sim_context_menu
        )
        layout.addWidget(self.list_added)

        # 顺位调整与移除操作区。
        action_layout = QHBoxLayout()
        action_layout.setContentsMargins(0, 0, 0, 0)

        btn_up = QPushButton("↑ 上移")
        btn_up.clicked.connect(self._move_item_up)

        btn_down = QPushButton("↓ 下移")
        btn_down.clicked.connect(self._move_item_down)

        btn_remove = QPushButton("移除选中项")
        btn_remove.clicked.connect(self._remove_selected)

        btn_remove_all = QPushButton("移除所有项")
        btn_remove_all.setStyleSheet(
            "QPushButton:hover { color: #ff4d4f; border-color: #ff4d4f; background-color: #fff1f0; }"
        )
        btn_remove_all.clicked.connect(self._remove_all)

        action_layout.addWidget(btn_up)
        action_layout.addWidget(btn_down)
        action_layout.addWidget(btn_remove)
        action_layout.addWidget(btn_remove_all)
        layout.addLayout(action_layout)

        # 对话框底部按钮区。
        btn_box = QHBoxLayout()
        btn_box.addStretch()

        btn_ok = QPushButton("确定")
        btn_ok.setObjectName("btnOk")
        btn_ok.clicked.connect(self.accept)

        btn_cancel = QPushButton("取消")
        btn_cancel.setObjectName("btnCancel")
        btn_cancel.clicked.connect(self.reject)

        btn_ok.setFixedSize(80, 28)
        btn_cancel.setFixedSize(80, 28)
        btn_box.addWidget(btn_ok)
        btn_box.addWidget(btn_cancel)
        layout.addLayout(btn_box)

        # -----------------------------
        # 信号绑定
        # -----------------------------
        self.combo_kind.currentIndexChanged.connect(self._on_kind_changed)
        self.combo_var_search.currentIndexChanged.connect(self._on_var_search_picked)

        if self.combo_var_search.lineEdit() is not None:
            self.combo_var_search.lineEdit().textChanged.connect(
                self._on_var_search_text_changed
            )

        self.list_vars.currentItemChanged.connect(self._on_var_changed)
        self.chk_all.stateChanged.connect(self._on_check_all)

        # -----------------------------
        # 初始状态同步
        # -----------------------------
        idx = 0 if self.current_kind == "input" else 1
        self.combo_kind.setCurrentIndex(idx)
        self._on_kind_changed(idx)
        self._refresh_added_list()

        if self.lock_to_output:
            self.combo_kind.setCurrentIndex(1)
            self.lbl_kind.hide()
            self.combo_kind.hide()

    # ---------------------------------------------------
    # 2.4 变量搜索与列表同步
    # ---------------------------------------------------
    def _rebuild_var_search_options(self, labels, current_map):
        """
        重建顶部搜索下拉框的数据源。

        这里与 AnalysisObjectDialog 中的处理保持一致，
        使“分析对象选择”和“叠加对象选择”的搜索体验统一。
        """
        self.combo_var_search.blockSignals(True)
        self.combo_var_search.clear()

        for lbl in labels:
            self.combo_var_search.addItem(lbl, lbl)
            idx = self.combo_var_search.count() - 1
            refs = " ".join(
                dict.fromkeys(str(ck) for _, ck in current_map.get(lbl, []))
            )
            self.combo_var_search.setItemData(
                idx,
                resolve_search_text(lbl, refs),
                SEARCH_TEXT_ROLE,
            )

        self.combo_var_search.setCurrentIndex(-1)
        self.combo_var_search.blockSignals(False)
        refresh_combo_dropdown_filter(self.combo_var_search, "")

    def _select_var_item_by_label(self, label):
        """
        根据变量标签，选中左侧变量列表中的对应项。
        """
        if not label:
            return

        for i in range(self.list_vars.count()):
            item = self.list_vars.item(i)
            item_label = item.data(Qt.UserRole) or item.text()
            if str(item_label) == str(label) and not item.isHidden():
                self.list_vars.setCurrentItem(item)
                self.list_vars.scrollToItem(item)
                return

    def _on_var_search_picked(self, idx):
        """
        当用户从搜索下拉框中直接选择变量时，同步定位到左侧列表。
        """
        if idx < 0:
            return

        label = self.combo_var_search.itemData(idx, Qt.UserRole)
        if not label:
            label = self.combo_var_search.itemText(idx)

        self._select_var_item_by_label(label)

    def _on_var_search_text_changed(self, text):
        """
        搜索框文本变化时，过滤左侧变量列表。

        若存在匹配项，则自动选中第一个可见项；
        若没有任何匹配，则清空右侧候选情景区。
        """
        query = str(text or "")
        refresh_combo_dropdown_filter(self.combo_var_search, query)

        first_visible = None
        for i in range(self.list_vars.count()):
            item = self.list_vars.item(i)
            label = item.data(Qt.UserRole) or item.text()
            search_text = item.data(SEARCH_TEXT_ROLE) or resolve_search_text(label)
            matched = fuzzy_contains_case_insensitive(query, search_text)
            item.setHidden(not matched)

            if matched and first_visible is None:
                first_visible = item

        current = self.list_vars.currentItem()
        if first_visible is None:
            self.list_vars.setCurrentRow(-1)
            self._on_var_changed(None, None)
            return

        if first_visible is not None and (current is None or current.isHidden()):
            self.list_vars.setCurrentItem(first_visible)

    # ---------------------------------------------------
    # 2.5 左右级联联动
    # ---------------------------------------------------
    def _on_kind_changed(self, idx):
        """
        当变量类型切换时，重建当前可选变量列表，并重置右侧候选区。

        处理逻辑：
        1. 清空变量列表与候选情景列表；
        2. 重置“全选”复选框；
        3. 根据当前类型切换到 input_map / output_map；
        4. 重建左侧变量列表；
        5. 若当前主对象对应变量仍可见，则优先定位到它。
        """
        self.list_vars.clear()
        self.list_sims.clear()

        self.chk_all.blockSignals(True)
        self.chk_all.setChecked(False)
        self.chk_all.blockSignals(False)

        kind = "input" if idx == 0 else "output"
        current_map = self.input_map if kind == "input" else self.output_map
        labels = sorted(current_map.keys())

        self._rebuild_var_search_options(labels, current_map)
        target_item = None

        for lbl in labels:
            item = QListWidgetItem(lbl)
            item.setData(Qt.UserRole, lbl)

            refs = " ".join(
                dict.fromkeys(str(ck) for _, ck in current_map.get(lbl, []))
            )
            item.setData(SEARCH_TEXT_ROLE, resolve_search_text(lbl, refs))
            self.list_vars.addItem(item)

            if kind == self.current_kind and any(
                (c == self.current_ck for s, c in current_map[lbl])
            ):
                target_item = item

        if target_item is not None:
            self.list_vars.setCurrentItem(target_item)
        elif self.list_vars.count() > 0:
            self.list_vars.setCurrentRow(0)

        if self.combo_var_search.lineEdit() is not None:
            self.combo_var_search.lineEdit().clear()

        self._on_var_search_text_changed("")

    def _on_var_changed(self, current, previous):
        """
        当左侧变量变化时，重建右侧候选情景列表。

        候选项有三种状态：
        1. 当前主对象：不可选，显示“(当前主对象)”；
        2. 已经加入叠加列表：不可选，显示“(已在叠加列表中)”；
        3. 普通可选项：带复选框，可被加入叠加列表。

        额外处理：
        若存在情景 1（即 sim1）且该项可用，则默认勾选它，
        便于用户快速以基准情景作为首个叠加对象。
        """
        self.list_sims.clear()

        self.chk_all.blockSignals(True)
        self.chk_all.setChecked(False)
        self.chk_all.blockSignals(False)

        if not current:
            return

        lbl = current.data(Qt.UserRole)
        kind = "input" if self.combo_kind.currentIndex() == 0 else "output"
        current_map = self.input_map if kind == "input" else self.output_map

        sim1_item = None

        for sid, ck in current_map.get(lbl, []):
            is_current = (
                kind == self.current_kind
                and sid == self.current_sid
                and (ck == self.current_ck)
            )
            is_added = False

            for d in self.result_items:
                d_sid = d[0]
                d_ck = d[1]
                # 兼容旧版本历史数据：旧结构可能只有 (sid, ck, var_name) 三项。
                d_kind = d[3] if len(d) == 4 else self.current_kind
                if sid == d_sid and ck == d_ck and (kind == d_kind):
                    is_added = True
                    break

            sim_label = get_sim_display_name(sid)
            item = QListWidgetItem(sim_label)
            item.setData(Qt.UserRole, (sid, ck, lbl, kind))

            if is_current:
                item.setText(f"{sim_label} (当前主对象)")
                item.setFlags(item.flags() & ~Qt.ItemIsEnabled & ~Qt.ItemIsUserCheckable)
            elif is_added:
                item.setText(f"{sim_label} (已在叠加列表中)")
                item.setFlags(item.flags() & ~Qt.ItemIsEnabled & ~Qt.ItemIsUserCheckable)
            else:
                item.setFlags(item.flags() | Qt.ItemIsUserCheckable | Qt.ItemIsEnabled)

            item.setCheckState(Qt.Unchecked)

            try:
                sid_num = int(sid)
            except Exception:
                sid_num = None

            if (item.flags() & Qt.ItemIsEnabled) and sid_num == 1:
                sim1_item = item

            self.list_sims.addItem(item)

        if sim1_item is not None:
            sim1_item.setCheckState(Qt.Checked)

    # ---------------------------------------------------
    # 2.6 右键菜单：候选区 / 已添加区情景重命名
    # ---------------------------------------------------
    def _open_candidate_sim_context_menu(self, pos):
        """
        在右侧候选情景列表中打开右键菜单。

        若完成重命名：
        1. 刷新已添加叠加项列表；
        2. 刷新当前变量下的候选项列表。
        """
        item = self.list_sims.itemAt(pos)
        if item is None:
            return

        data = item.data(Qt.UserRole)
        sid = data[0] if isinstance(data, (tuple, list)) and len(data) > 0 else None
        if sid is None:
            return

        menu = QMenu(self)
        act_rename = menu.addAction("重命名情景")
        chosen = menu.exec(self.list_sims.mapToGlobal(pos))
        if chosen != act_rename:
            return

        if _rename_sim_with_dialog(self, sid):
            self._refresh_added_list()
            self._on_var_changed(self.list_vars.currentItem(), None)

    def _open_added_sim_context_menu(self, pos):
        """
        在下方已添加叠加项列表中打开右键菜单。

        若完成重命名：
        1. 刷新已添加列表本身；
        2. 刷新候选情景列表，确保文字状态同步。
        """
        item = self.list_added.itemAt(pos)
        if item is None:
            return

        data = item.data(Qt.UserRole)
        sid = data[0] if isinstance(data, (tuple, list)) and len(data) > 0 else None
        if sid is None:
            return

        menu = QMenu(self)
        act_rename = menu.addAction("重命名情景")
        chosen = menu.exec(self.list_added.mapToGlobal(pos))
        if chosen != act_rename:
            return

        if _rename_sim_with_dialog(self, sid):
            self._refresh_added_list()
            self._on_var_changed(self.list_vars.currentItem(), None)

    # ---------------------------------------------------
    # 2.7 候选区批量勾选
    # ---------------------------------------------------
    def _on_check_all(self, state):
        """
        对当前候选区中“可用”的情景执行全选 / 全不选。

        注意：
        当前主对象与已在叠加列表中的对象会被禁用，
        不会被“全选所有情景”强行勾选。
        """
        check_state = Qt.Checked if state == Qt.Checked.value else Qt.Unchecked
        for i in range(self.list_sims.count()):
            item = self.list_sims.item(i)
            if item.flags() & Qt.ItemIsEnabled:
                item.setCheckState(check_state)

    # ---------------------------------------------------
    # 2.8 内部状态同步
    # ---------------------------------------------------
    def _sync_internal_data(self):
        """
        使内部 result_items 与界面上的已添加列表顺序保持完全一致。

        说明：
        由于 list_added 支持拖拽排序与上下移动，
        仅修改界面顺序是不够的，必须同步回内部数据列表，
        才能保证最终返回给外部调用方的顺序正确。
        """
        self.result_items = [
            self.list_added.item(i).data(Qt.UserRole)
            for i in range(self.list_added.count())
        ]

    # ---------------------------------------------------
    # 2.9 叠加列表增删改
    # ---------------------------------------------------
    def _add_to_overlay(self):
        """
        将候选区中已勾选的项目加入叠加列表。

        处理要点：
        1. 先同步内部顺序，避免界面与内存状态不一致；
        2. 只处理“可用且已勾选”的候选项；
        3. 再次执行防重校验，避免快速点击造成重复加入；
        4. 加入完成后，刷新已添加列表；
        5. 重新刷新当前变量的候选项区，使已加入对象灰化不可选。
        """
        self._sync_internal_data()

        for i in range(self.list_sims.count()):
            item = self.list_sims.item(i)
            if item.flags() & Qt.ItemIsEnabled and item.checkState() == Qt.Checked:
                data = item.data(Qt.UserRole)

                if not any((d[0] == data[0] and d[1] == data[1] for d in self.result_items)):
                    self.result_items.append(data)

        self._refresh_added_list()
        self._on_var_changed(self.list_vars.currentItem(), None)

    def _remove_selected(self):
        """
        移除下方已添加列表中当前选中的项。

        移除后会同步内部数据，并重新刷新候选区，
        使这些被移除的对象重新变为可选状态。
        """
        for item in self.list_added.selectedItems():
            self.list_added.takeItem(self.list_added.row(item))

        self._sync_internal_data()
        self._on_var_changed(self.list_vars.currentItem(), None)

    def _remove_all(self):
        """
        清空全部叠加项。
        """
        if self.list_added.count() == 0:
            return

        self.list_added.clear()
        self._sync_internal_data()
        self._on_var_changed(self.list_vars.currentItem(), None)

    # ---------------------------------------------------
    # 2.10 顺位调整
    # ---------------------------------------------------
    def _move_item_up(self):
        """
        将已添加列表中的选中项上移。

        多选上移策略：
        1. 按当前行号从小到大遍历；
        2. 若某项上方不是另一已选项，则允许交换；
        3. 这样可以避免多选情况下顺序错乱。
        """
        selected = self.list_added.selectedItems()
        for item in sorted(selected, key=lambda x: self.list_added.row(x)):
            row = self.list_added.row(item)
            if row > 0 and (not self.list_added.item(row - 1).isSelected()):
                self.list_added.insertItem(row - 1, self.list_added.takeItem(row))
                item.setSelected(True)

        self._sync_internal_data()

    def _move_item_down(self):
        """
        将已添加列表中的选中项下移。

        多选下移策略：
        1. 按当前行号从大到小遍历；
        2. 若某项下方不是另一已选项，则允许交换；
        3. 这样可以避免多选移动时互相覆盖。
        """
        selected = self.list_added.selectedItems()
        for item in sorted(
            selected,
            key=lambda x: self.list_added.row(x),
            reverse=True,
        ):
            row = self.list_added.row(item)
            if row < self.list_added.count() - 1 and (
                not self.list_added.item(row + 1).isSelected()
            ):
                self.list_added.insertItem(row + 1, self.list_added.takeItem(row))
                item.setSelected(True)

        self._sync_internal_data()

    # ---------------------------------------------------
    # 2.11 已添加列表渲染
    # ---------------------------------------------------
    def _refresh_added_list(self):
        """
        根据内部 result_items 重新渲染下方“已添加叠加项”列表。

        显示格式：
        [输入] 变量名 - 情景显示名
        [输出] 变量名 - 情景显示名

        兼容逻辑：
        旧版本历史数据可能不包含 kind 字段，此时默认沿用 current_kind。
        """
        self.list_added.clear()

        for data in self.result_items:
            if len(data) == 4:
                sid, ck, var_name, kind = data
            else:
                sid, ck, var_name = data
                kind = self.current_kind

            kind_lbl = "输入" if kind == "input" else "输出"
            sim_label = get_sim_display_name(sid)
            item = QListWidgetItem(f"[{kind_lbl}] {var_name} - {sim_label}")
            item.setData(Qt.UserRole, (sid, ck, var_name, kind))
            self.list_added.addItem(item)

    # ---------------------------------------------------
    # 2.12 对外结果接口
    # ---------------------------------------------------
    def get_results(self):
        """
        返回最终叠加列表结果。

        调用前先同步一次内部状态，确保拖拽排序、上下移动等操作
        已经完整反映到 result_items 中。
        """
        self._sync_internal_data()
        return self.result_items