# -*- coding: utf-8 -*-
"""
用于可搜索变量选择器的共享辅助工具模块。
提供文本归一化、引用提取、模糊匹配以及 Qt 控件的搜索过滤配置功能。
"""

from __future__ import annotations

import re
from typing import Dict, Iterable, Optional

from PySide6.QtCore import Qt
from PySide6.QtWidgets import QComboBox, QCompleter


# =======================================================
# 1. 常量定义
# =======================================================
# 专用角色 (Role)，用于在下拉框/列表项的底层数据模型中，存储预先构建好的可搜索文本。
SEARCH_TEXT_ROLE = Qt.UserRole + 17


# =======================================================
# 2. 字符串提取与预处理
# =======================================================
def normalize_for_search(value: object) -> str:
    """
    统一字符串格式以供搜索匹配：将其转换为小写 (casefold 比 lower 更彻底) 并去除首尾空格。
    """
    return str(value or "").casefold().strip()


def extract_reference_from_label(label: object) -> str:
    """
    从诸如 "Name (Sheet1!B1)" 格式的标签中提取引用字符串。
    如果不存在尾部的 "(...)" 片段，则返回空字符串。
    """
    text = str(label or "").strip()
    if not text.endswith(")") or "(" not in text:
        return ""
    # 从右侧按 '(' 切割一次，提取括号内的内容
    return text.rsplit("(", 1)[-1][:-1].strip()


_CELL_REF_WITH_SUFFIX_RE = re.compile(r"^(?P<addr>[A-Za-z]{1,3}\d{1,7})(?:_[A-Za-z0-9]+)+$")
_PLAIN_CELL_REF_RE = re.compile(r"^[A-Za-z]{1,3}\d{1,7}$")


def normalize_input_reference_for_display(
    cell_key: object,
    *,
    include_sheet: bool = False,
) -> str:
    """
    Convert a technical input key to a user-facing cell reference.
    It removes trailing engine suffixes such as "_1" while keeping the key untouched elsewhere.
    """
    text = str(cell_key or "").strip().replace("$", "")
    if not text:
        return ""

    sheet = ""
    ref = text
    if "!" in text:
        sheet, ref = text.split("!", 1)
        sheet = sheet.strip()
        ref = ref.strip()

    ref_upper = ref.upper()
    m = _CELL_REF_WITH_SUFFIX_RE.match(ref_upper)
    if m:
        clean_ref = m.group("addr")
    else:
        left = ref_upper.split("_", 1)[0] if "_" in ref_upper else ref_upper
        clean_ref = left if _PLAIN_CELL_REF_RE.match(left) else ref_upper

    if include_sheet and sheet:
        return f"{sheet}!{clean_ref}"
    return clean_ref


def build_input_reference_display_map(cell_keys: Iterable[object]) -> Dict[str, str]:
    """
    Build a key->display-reference map.
    Always prefer sheet-qualified cleaned references when sheet information exists.
    """
    resolved: Dict[str, str] = {}
    for raw_key in cell_keys or ():
        key_text = str(raw_key or "")
        sheet_ref = normalize_input_reference_for_display(key_text, include_sheet=True)
        plain_ref = normalize_input_reference_for_display(key_text, include_sheet=False)
        resolved[key_text] = sheet_ref or plain_ref or key_text
    return resolved


def build_input_variable_label(
    display_name: object,
    cell_reference: object,
) -> str:
    """Build the normalized UI label: name(cell_reference)."""
    name_text = str(display_name or "").strip()
    ref_text = str(cell_reference or "").strip()
    if name_text and ref_text:
        return f"{name_text}({ref_text})"
    if name_text:
        return name_text
    return ref_text


def build_search_text(display_name: object, reference: object = "") -> str:
    """
    构建一个支持按名称或地址进行包含匹配的可搜索文本块 (blob)。
    将显示名称与引用地址拼接在一起，方便后续通过单一字符串进行全量模糊检索。
    """
    name_text = str(display_name or "").strip()
    ref_text = str(reference or "").strip()
    # 如果引用文本存在且未包含在名称中，则将二者拼接
    if ref_text and ref_text not in name_text:
        return f"{name_text} {ref_text}".strip()
    return name_text


def resolve_search_text(
    label: object,
    explicit_reference: Optional[object] = None,
) -> str:
    """
    解析并构建最终的可搜索文本的高级封装方法。
    如果未显式提供引用 (explicit_reference)，则尝试自动从标签 (label) 中提取。
    """
    ref = str(explicit_reference or "").strip()
    if not ref:
        ref = extract_reference_from_label(label)
    return build_search_text(label, ref)


# =======================================================
# 3. 核心匹配算法
# =======================================================
def fuzzy_contains_case_insensitive(query: object, *parts: object) -> bool:
    """
    所有变量选择器通用的、不区分大小写的包含匹配逻辑。
    只要用户的搜索词 (query) 包含在拼接好的目标文本 (parts) 中，即返回 True。
    """
    query_text = normalize_for_search(query)
    # 如果搜索词为空，默认匹配所有选项
    if not query_text:
        return True
    # 将传入的多个内容片段拼接为干草堆 (haystack) 用于检索
    haystack = " ".join(normalize_for_search(p) for p in parts if str(p or "").strip())
    return query_text in haystack


# =======================================================
# 4. Qt UI 控件配置与视图过滤
# =======================================================
def setup_searchable_combo(combo: QComboBox, placeholder: str) -> None:
    """
    将普通的下拉框 (ComboBox) 转换为可编辑、可搜索的下拉输入框。
    配置行编辑器 (LineEdit) 与自动完成器 (QCompleter) 的行为。
    """
    combo.setEditable(True)
    # 禁止用户通过回车将新文本插入到下拉列表中
    combo.setInsertPolicy(QComboBox.NoInsert)
    combo.setMaxVisibleItems(16)
    
    line_edit = combo.lineEdit()
    if line_edit is not None:
        line_edit.setPlaceholderText(placeholder)
        line_edit.setClearButtonEnabled(True) # 开启右侧的一键清除 (X) 按钮
        
    completer = combo.completer()
    if isinstance(completer, QCompleter):
        # 设置补全器：不区分大小写、包含匹配、弹出列表模式
        completer.setCaseSensitivity(Qt.CaseInsensitive)
        completer.setFilterMode(Qt.MatchContains)
        completer.setCompletionMode(QCompleter.PopupCompletion)


def refresh_combo_dropdown_filter(
    combo: QComboBox,
    query: object,
    *,
    search_role: int = SEARCH_TEXT_ROLE,
) -> None:
    """
    就地 (in-place) 过滤下拉选项行，使得只有符合包含匹配的选项保持可见。
    通过隐藏/显示列表视图 (View) 中的行来实现实时的搜索筛选效果。
    """
    view = combo.view()
    if view is None:
        return
        
    # 遍历下拉框中的所有选项
    for row in range(combo.count()):
        # 优先从专门的 search_role 中获取预构建的搜索文本
        text = combo.itemData(row, search_role)
        if text is None:
            # 如果没有，降级使用原本显示的文本
            text = combo.itemText(row)
            
        # 执行模糊匹配，如果不匹配则隐藏该行
        matched = fuzzy_contains_case_insensitive(query, text)
        view.setRowHidden(row, not matched)
