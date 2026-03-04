import re
import sys
import json
import os
import ctypes
import threading
from contextlib import suppress
from dataclasses import dataclass
from datetime import datetime
from typing import Dict, List, Optional, Tuple

import winreg
from PySide6.QtCore import QObject, QSignalBlocker, QTimer, Qt, Signal
from PySide6.QtGui import QColor, QFont
from PySide6.QtWidgets import (
    QApplication,
    QAbstractItemView,
    QCheckBox,
    QComboBox,
    QFrame,
    QGridLayout,
    QHBoxLayout,
    QFileDialog,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QPlainTextEdit,
    QSizePolicy,
    QSplitter,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)


@dataclass(frozen=True)
class Scope:
    id: str
    label: str
    hive: int
    hive_name: str
    path: str
    category: str
    kind: str  # shell | handler


@dataclass
class MenuItem:
    uid: str
    scope_id: str
    hive: int
    hive_name: str
    path: str
    category: str
    kind: str
    key_name: str
    display_name: str
    command: str
    enabled: bool
    has_submenu: bool


@dataclass
class ImportPlanItem:
    line_no: int
    scope_id: str
    key_name: str
    display_name: str
    command: str
    enabled: bool
    action: str  # create | update | unchanged | blocked | skip
    reason: str = ""


@dataclass
class TreeStateChange:
    hive: int
    path: str
    before: object
    after: object


@dataclass
class HistoryRecord:
    label: str
    changes: List[TreeStateChange]


def build_scopes() -> List[Scope]:
    hives = [
        (winreg.HKEY_CURRENT_USER, "HKCU", "当前用户"),
        (winreg.HKEY_LOCAL_MACHINE, "HKLM", "系统级"),
    ]
    locations = [
        ("file", "文件", r"Software\Classes\*"),
        ("allfs", "所有文件系统对象", r"Software\Classes\AllFilesystemObjects"),
        ("dir", "目录", r"Software\Classes\Directory"),
        ("dirbg", "目录背景", r"Software\Classes\Directory\Background"),
        ("folder", "文件夹", r"Software\Classes\Folder"),
        ("drive", "磁盘", r"Software\Classes\Drive"),
        ("desktopbg", "桌面背景", r"Software\Classes\DesktopBackground"),
    ]
    scopes: List[Scope] = []
    for hive, hive_name, hive_label in hives:
        for loc_id, loc_label, root in locations:
            scopes.append(
                Scope(
                    id=f"{hive_name.lower()}_{loc_id}_shell",
                    label=f"{hive_label} - {loc_label} (Shell)",
                    hive=hive,
                    hive_name=hive_name,
                    path=f"{root}\\shell",
                    category=loc_label,
                    kind="shell",
                )
            )
            scopes.append(
                Scope(
                    id=f"{hive_name.lower()}_{loc_id}_handler",
                    label=f"{hive_label} - {loc_label} (Handler)",
                    hive=hive,
                    hive_name=hive_name,
                    path=f"{root}\\shellex\\ContextMenuHandlers",
                    category=loc_label,
                    kind="handler",
                )
            )
    return scopes


SCOPES = build_scopes()
SUBMENU_PLACEHOLDER = "[子菜单/级联命令]"


class RefreshEmitter(QObject):
    completed = Signal(object, str, str, bool)


class RegistryManager:
    def __init__(self):
        self.scopes = SCOPES
        self.scope_map = {scope.id: scope for scope in self.scopes}

    @staticmethod
    def _slugify(text: str) -> str:
        text = re.sub(r"[\\/\[\]\*%:;|=,]+", "_", text.strip())
        text = re.sub(r"\s+", "_", text)
        text = re.sub(r"[^0-9A-Za-z_\-\u4e00-\u9fff]", "", text)
        return text or "CustomMenu"

    @staticmethod
    def _read_default(key) -> str:
        try:
            value, _ = winreg.QueryValueEx(key, "")
            return str(value)
        except OSError:
            return ""

    @staticmethod
    def _read_value(key, name: str) -> str:
        try:
            value, _ = winreg.QueryValueEx(key, name)
            return str(value)
        except OSError:
            return ""

    @staticmethod
    def _has_value(key, name: str) -> bool:
        try:
            winreg.QueryValueEx(key, name)
            return True
        except OSError:
            return False

    @staticmethod
    def _has_subkey(hive: int, path: str, subkey: str) -> bool:
        with suppress(OSError):
            with winreg.OpenKey(hive, f"{path}\\{subkey}", 0, winreg.KEY_READ):
                return True
        return False

    @staticmethod
    def _set_enabled_flag(item_key, enabled: bool):
        if enabled:
            with suppress(OSError):
                winreg.DeleteValue(item_key, "LegacyDisable")
        else:
            winreg.SetValueEx(item_key, "LegacyDisable", 0, winreg.REG_SZ, "")

    def _item_path(self, item: MenuItem) -> str:
        return f"{item.path}\\{item.key_name}"

    def _uid(self, scope: Scope, key_name: str) -> str:
        return f"{scope.id}|{key_name}"

    def uid_to_hive_path(self, uid: str) -> Tuple[int, str]:
        scope_id, key_name = uid.split("|", 1)
        scope = self.scope_map[scope_id]
        return scope.hive, f"{scope.path}\\{key_name}"

    def scope_key_to_hive_path(self, scope_id: str, key_name: str) -> Tuple[int, str]:
        scope = self.scope_map[scope_id]
        return scope.hive, f"{scope.path}\\{key_name}"

    def _extract_command(self, hive: int, item_path: str) -> Tuple[str, bool]:
        with suppress(OSError):
            with winreg.OpenKey(hive, f"{item_path}\\command", 0, winreg.KEY_READ) as cmd_key:
                cmd = self._read_default(cmd_key)
                if cmd:
                    return cmd, False
        with suppress(OSError):
            with winreg.OpenKey(hive, item_path, 0, winreg.KEY_READ) as key:
                has_sub = bool(self._read_value(key, "SubCommands")) or self._has_subkey(hive, item_path, "shell")
                if has_sub:
                    return SUBMENU_PLACEHOLDER, True
        return "", False

    def list_items(self) -> List[MenuItem]:
        rows: List[MenuItem] = []
        for scope in self.scopes:
            with suppress(OSError):
                with winreg.OpenKey(scope.hive, scope.path, 0, winreg.KEY_READ) as base:
                    index = 0
                    while True:
                        try:
                            key_name = winreg.EnumKey(base, index)
                            index += 1
                        except OSError:
                            break

                        item_path = f"{scope.path}\\{key_name}"
                        try:
                            with winreg.OpenKey(scope.hive, item_path, 0, winreg.KEY_READ) as item_key:
                                display_name = self._read_value(item_key, "MUIVerb") or self._read_default(item_key) or key_name
                                enabled = not self._has_value(item_key, "LegacyDisable")
                                if scope.kind == "handler":
                                    command = self._read_default(item_key) or "[COM Handler]"
                                    enabled = True
                                    has_submenu = False
                                else:
                                    command, has_submenu = self._extract_command(scope.hive, item_path)
                        except OSError:
                            continue

                        rows.append(
                            MenuItem(
                                uid=self._uid(scope, key_name),
                                scope_id=scope.id,
                                hive=scope.hive,
                                hive_name=scope.hive_name,
                                path=scope.path,
                                category=scope.category,
                                kind=scope.kind,
                                key_name=key_name,
                                display_name=display_name,
                                command=command,
                                enabled=enabled,
                                has_submenu=has_submenu,
                            )
                        )
        rows.sort(key=lambda x: (x.hive_name, x.category, x.kind, x.display_name.lower()))
        return rows

    def _list_scope_keys(self, scope: Scope) -> List[str]:
        keys: List[str] = []
        with suppress(OSError):
            with winreg.OpenKey(scope.hive, scope.path, 0, winreg.KEY_READ) as base:
                index = 0
                while True:
                    try:
                        keys.append(winreg.EnumKey(base, index))
                        index += 1
                    except OSError:
                        break
        return keys

    def create_shell_item(
        self, scope_id: str, display_name: str, command: str, key_name: str = "", enabled: bool = True
    ) -> str:
        scope = self.scope_map[scope_id]
        if scope.kind != "shell":
            raise ValueError("只能在 Shell 路径新增。")

        display_name = display_name.strip()
        command = command.strip()
        if not display_name:
            raise ValueError("显示名称不能为空。")
        if not command:
            raise ValueError("执行命令不能为空。")

        base_name = self._slugify(key_name or display_name)
        existing = set(self._list_scope_keys(scope))
        final_key = base_name
        index = 2
        while final_key in existing:
            final_key = f"{base_name}_{index}"
            index += 1

        item_path = f"{scope.path}\\{final_key}"
        with winreg.CreateKeyEx(scope.hive, item_path, 0, winreg.KEY_SET_VALUE) as item_key:
            winreg.SetValueEx(item_key, "MUIVerb", 0, winreg.REG_SZ, display_name)
            self._set_enabled_flag(item_key, enabled)
        with winreg.CreateKeyEx(scope.hive, f"{item_path}\\command", 0, winreg.KEY_SET_VALUE) as cmd_key:
            winreg.SetValueEx(cmd_key, "", 0, winreg.REG_SZ, command)

        return self._uid(scope, final_key)

    def upsert_shell_item(
        self, scope_id: str, display_name: str, command: str, key_name: str = "", enabled: bool = True
    ) -> Tuple[str, bool]:
        scope = self.scope_map[scope_id]
        if scope.kind != "shell":
            raise ValueError("只能在 Shell 路径导入。")

        display_name = display_name.strip()
        command = command.strip()
        if not display_name:
            raise ValueError("显示名称不能为空。")
        if not command:
            raise ValueError("执行命令不能为空。")

        base_name = self._slugify(key_name or display_name)
        item_path = f"{scope.path}\\{base_name}"
        existed = False

        with suppress(OSError):
            with winreg.OpenKey(scope.hive, item_path, 0, winreg.KEY_READ):
                existed = True

        if existed:
            with winreg.OpenKey(scope.hive, item_path, 0, winreg.KEY_SET_VALUE) as key:
                winreg.SetValueEx(key, "MUIVerb", 0, winreg.REG_SZ, display_name)
                self._set_enabled_flag(key, enabled)
            with winreg.CreateKeyEx(scope.hive, f"{item_path}\\command", 0, winreg.KEY_SET_VALUE) as cmd_key:
                winreg.SetValueEx(cmd_key, "", 0, winreg.REG_SZ, command)
            return self._uid(scope, base_name), False

        uid = self.create_shell_item(
            scope_id=scope_id,
            display_name=display_name,
            command=command,
            key_name=base_name,
            enabled=enabled,
        )
        return uid, True

    def update_shell_item(self, item: MenuItem, display_name: str, command: str, enabled: bool):
        if item.kind != "shell":
            raise ValueError("Handler 条目只读。")

        display_name = display_name.strip()
        command = command.strip()
        if not display_name:
            raise ValueError("显示名称不能为空。")
        if not command:
            raise ValueError("执行命令不能为空。")

        item_path = self._item_path(item)
        with winreg.OpenKey(item.hive, item_path, 0, winreg.KEY_SET_VALUE) as key:
            winreg.SetValueEx(key, "MUIVerb", 0, winreg.REG_SZ, display_name)
            self._set_enabled_flag(key, enabled)
        with winreg.CreateKeyEx(item.hive, f"{item_path}\\command", 0, winreg.KEY_SET_VALUE) as cmd_key:
            winreg.SetValueEx(cmd_key, "", 0, winreg.REG_SZ, command)

    def set_enabled(self, item: MenuItem, enabled: bool):
        if item.kind != "shell":
            raise ValueError("Handler 条目不支持启用/禁用。")
        with winreg.OpenKey(item.hive, self._item_path(item), 0, winreg.KEY_SET_VALUE) as key:
            self._set_enabled_flag(key, enabled)

    def delete_shell_item(self, item: MenuItem):
        if item.kind != "shell":
            raise ValueError("Handler 条目只读。")
        self._delete_tree(item.hive, self._item_path(item))

    def _delete_tree(self, hive: int, path: str):
        children: List[str] = []
        with winreg.OpenKey(hive, path, 0, winreg.KEY_READ) as key:
            index = 0
            while True:
                try:
                    children.append(winreg.EnumKey(key, index))
                    index += 1
                except OSError:
                    break
        for child in children:
            self._delete_tree(hive, f"{path}\\{child}")
        winreg.DeleteKey(hive, path)

    def _delete_tree_if_exists(self, hive: int, path: str):
        with suppress(OSError):
            self._delete_tree(hive, path)

    @staticmethod
    def _export_tree_from_open_key(key) -> Dict:
        values: List[Tuple[str, object, int]] = []
        value_index = 0
        while True:
            try:
                name, value, value_type = winreg.EnumValue(key, value_index)
                values.append((name, value, value_type))
                value_index += 1
            except OSError:
                break

        children: Dict[str, Dict] = {}
        child_index = 0
        while True:
            try:
                child_name = winreg.EnumKey(key, child_index)
                child_index += 1
            except OSError:
                break
            with winreg.OpenKey(key, child_name, 0, winreg.KEY_READ) as child_key:
                children[child_name] = RegistryManager._export_tree_from_open_key(child_key)

        return {"values": values, "children": children}

    def export_tree_if_exists(self, hive: int, path: str):
        with suppress(OSError):
            with winreg.OpenKey(hive, path, 0, winreg.KEY_READ) as key:
                return self._export_tree_from_open_key(key)
        return None

    def export_uid_tree_if_exists(self, uid: str):
        hive, path = self.uid_to_hive_path(uid)
        return self.export_tree_if_exists(hive, path)

    def _import_tree_to_open_key(self, key, tree: Dict):
        for name, value, value_type in tree.get("values", []):
            winreg.SetValueEx(key, name, 0, value_type, value)
        for child_name, child_tree in tree.get("children", {}).items():
            with winreg.CreateKeyEx(key, child_name, 0, winreg.KEY_SET_VALUE | winreg.KEY_CREATE_SUB_KEY) as child_key:
                self._import_tree_to_open_key(child_key, child_tree)

    def apply_tree_state(self, hive: int, path: str, tree_state):
        self._delete_tree_if_exists(hive, path)
        if tree_state is None:
            return
        with winreg.CreateKeyEx(hive, path, 0, winreg.KEY_SET_VALUE | winreg.KEY_CREATE_SUB_KEY) as root_key:
            self._import_tree_to_open_key(root_key, tree_state)

    def can_write(self, item: MenuItem) -> bool:
        if item.kind != "shell":
            return False
        with suppress(OSError):
            with winreg.OpenKey(item.hive, self._item_path(item), 0, winreg.KEY_SET_VALUE):
                return True
        return False

    def can_write_scope(self, scope_id: str) -> bool:
        scope = self.scope_map.get(scope_id)
        if not scope or scope.kind != "shell":
            return False
        with suppress(OSError):
            with winreg.OpenKey(scope.hive, scope.path, 0, winreg.KEY_CREATE_SUB_KEY):
                return True
        return False


APP_STYLESHEET = """
QWidget {
    background: #eef3f9;
    color: #1f2937;
    font-family: "Microsoft YaHei UI";
    font-size: 13px;
}
#Header {
    background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #f8fbff, stop:1 #e9f2ff);
    border: 1px solid #cfe0f6;
    border-radius: 16px;
}
#HeaderTitle {
    color: #0f172a;
    font-size: 30px;
    font-weight: 700;
}
#HeaderSubtitle {
    color: #334155;
    font-size: 14px;
    font-weight: 500;
}
#Card {
    background: #ffffff;
    border: 1px solid #dbe3ef;
    border-radius: 14px;
}
#StatValue {
    color: #0f172a;
    font-size: 20px;
    font-weight: 700;
}
#StatLabel {
    color: #64748b;
    font-size: 12px;
}
QLineEdit, QComboBox, QPlainTextEdit {
    background: #f8fafc;
    border: 1px solid #d4deeb;
    border-radius: 10px;
    padding: 7px 10px;
}
QLineEdit:focus, QComboBox:focus, QPlainTextEdit:focus {
    border: 1px solid #3b82f6;
    background: #ffffff;
}
QComboBox::drop-down {
    border: 0px;
    width: 20px;
}
QCheckBox {
    spacing: 8px;
}
QPushButton {
    background: #0ea5e9;
    color: #ffffff;
    border: 0;
    border-radius: 10px;
    padding: 8px 12px;
    font-weight: 600;
}
QPushButton:hover {
    background: #0284c7;
}
QPushButton:pressed {
    background: #0369a1;
}
QPushButton:disabled {
    background: #cbd5e1;
    color: #f8fafc;
}
#GhostButton {
    background: #e2e8f0;
    color: #334155;
}
#GhostButton:hover {
    background: #cbd5e1;
}
QTableWidget {
    background: #ffffff;
    border: 1px solid #dbe3ef;
    border-radius: 12px;
    gridline-color: #f1f5f9;
    selection-background-color: #dbeafe;
    selection-color: #0f172a;
}
QHeaderView::section {
    background: #f1f5f9;
    color: #1e293b;
    border: 0;
    border-bottom: 1px solid #dbe3ef;
    padding: 8px;
    font-weight: 700;
}
"""


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("鼠标右键管理器")
        self.resize(1420, 900)
        self.setMinimumSize(1180, 760)

        self.manager = RegistryManager()
        self.items: List[MenuItem] = []
        self.visible_items: List[MenuItem] = []
        self.item_map: Dict[str, MenuItem] = {}
        self.current_uid = ""

        self.shell_scopes = [scope for scope in SCOPES if scope.kind == "shell"]
        self.scope_labels = [scope.label for scope in self.shell_scopes]
        self.scope_by_label = {scope.label: scope.id for scope in self.shell_scopes}
        self.categories = ["全部"] + sorted({scope.category for scope in SCOPES})
        self.sort_col = 3
        self.sort_order = Qt.AscendingOrder
        self.filter_timer = QTimer(self)
        self.filter_timer.setSingleShot(True)
        self.filter_timer.setInterval(180)
        self.refresh_emitter = RefreshEmitter()
        self.refresh_emitter.completed.connect(self._on_refresh_completed)
        self.refresh_inflight = False
        self.pending_refresh: Optional[Tuple[str, bool]] = None
        self.third_party_cache: Dict[str, bool] = {}
        self.undo_stack: List[HistoryRecord] = []
        self.redo_stack: List[HistoryRecord] = []
        self.max_history = 30

        self._build_ui()
        self._wire_events()
        self._update_history_buttons()
        self._refresh()

    def _build_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        root = QVBoxLayout(central)
        root.setContentsMargins(14, 14, 14, 12)
        root.setSpacing(10)

        header = QFrame()
        header.setObjectName("Header")
        header_layout = QVBoxLayout(header)
        header_layout.setContentsMargins(18, 14, 18, 14)
        title = QLabel("右键菜单管理器")
        title.setObjectName("HeaderTitle")
        subtitle = QLabel("全量扫描 Shell + ContextMenuHandlers · 点击表头可排序")
        subtitle.setObjectName("HeaderSubtitle")
        header_layout.addWidget(title)
        header_layout.addWidget(subtitle)
        root.addWidget(header)

        stat_row = QHBoxLayout()
        stat_row.setSpacing(10)
        self.stat_total = self._build_stat_card("总条目")
        self.stat_visible = self._build_stat_card("当前可见")
        self.stat_disabled = self._build_stat_card("禁用 (Shell)")
        stat_row.addWidget(self.stat_total[0])
        stat_row.addWidget(self.stat_visible[0])
        stat_row.addWidget(self.stat_disabled[0])
        root.addLayout(stat_row)

        splitter = QSplitter(Qt.Horizontal)
        splitter.setChildrenCollapsible(False)
        splitter.setHandleWidth(10)
        root.addWidget(splitter, 1)

        left = self._build_left_panel()
        right = self._build_right_panel()
        splitter.addWidget(left)
        splitter.addWidget(right)
        splitter.setStretchFactor(0, 3)
        splitter.setStretchFactor(1, 7)

        foot = QHBoxLayout()
        self.status_label = QLabel("准备就绪")
        self.status_label.setStyleSheet("color:#475569;")
        self.exit_btn = QPushButton("退出")
        self.exit_btn.setObjectName("GhostButton")
        self.exit_btn.setFixedWidth(96)
        foot.addWidget(self.status_label, 1)
        foot.addWidget(self.exit_btn)
        root.addLayout(foot)

    def _build_stat_card(self, label: str) -> Tuple[QFrame, QLabel]:
        card = QFrame()
        card.setObjectName("Card")
        card.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        lay = QVBoxLayout(card)
        lay.setContentsMargins(14, 10, 14, 10)
        value = QLabel("0")
        value.setObjectName("StatValue")
        text = QLabel(label)
        text.setObjectName("StatLabel")
        lay.addWidget(value)
        lay.addWidget(text)
        return card, value

    def _build_left_panel(self) -> QFrame:
        card = QFrame()
        card.setObjectName("Card")
        lay = QVBoxLayout(card)
        lay.setContentsMargins(14, 14, 14, 14)
        lay.setSpacing(10)

        title = QLabel("筛选与编辑")
        title.setFont(QFont("Microsoft YaHei UI", 12, QFont.Bold))
        lay.addWidget(title)

        form = QGridLayout()
        form.setHorizontalSpacing(8)
        form.setVerticalSpacing(8)

        self.search_edit = QLineEdit()
        self.search_edit.setPlaceholderText("搜索显示名称 / 键名 / 命令 ...")

        self.hive_combo = QComboBox()
        self.hive_combo.addItems(["全部", "HKCU", "HKLM"])

        self.category_combo = QComboBox()
        self.category_combo.addItems(self.categories)

        self.kind_combo = QComboBox()
        self.kind_combo.addItems(["全部", "Shell", "Handler"])

        self.third_party_combo = QComboBox()
        self.third_party_combo.addItems(["全部", "仅第三方", "仅系统项"])

        self.scope_combo = QComboBox()
        self.scope_combo.addItems(self.scope_labels)

        self.name_edit = QLineEdit()
        self.name_edit.setPlaceholderText("显示名称")

        self.key_edit = QLineEdit()
        self.key_edit.setPlaceholderText("键名（可选）")

        self.command_edit = QPlainTextEdit()
        self.command_edit.setPlaceholderText('示例: notepad.exe "%1"')
        self.command_edit.setMinimumHeight(160)

        self.enabled_check = QCheckBox("启用该菜单项")
        self.enabled_check.setChecked(True)

        labels = [
            ("搜索", self.search_edit),
            ("来源", self.hive_combo),
            ("类型", self.category_combo),
            ("条目", self.kind_combo),
            ("视图", self.third_party_combo),
            ("新增位置", self.scope_combo),
            ("显示名称", self.name_edit),
            ("键名", self.key_edit),
        ]
        for row, (text, widget) in enumerate(labels):
            form.addWidget(QLabel(text), row, 0)
            form.addWidget(widget, row, 1)

        form.addWidget(QLabel("执行命令"), len(labels), 0, Qt.AlignTop)
        form.addWidget(self.command_edit, len(labels), 1)
        form.addWidget(self.enabled_check, len(labels) + 1, 1)
        lay.addLayout(form)

        btn_grid = QGridLayout()
        btn_grid.setHorizontalSpacing(8)
        btn_grid.setVerticalSpacing(8)

        self.btn_add = QPushButton("新增")
        self.btn_update = QPushButton("更新已选")
        self.btn_toggle = QPushButton("启用/禁用")
        self.btn_enable = QPushButton("批量启用")
        self.btn_disable = QPushButton("批量禁用")
        self.btn_delete = QPushButton("删除(单条)")
        self.btn_batch_delete = QPushButton("批量删除")
        self.btn_import = QPushButton("导入JSON")
        self.btn_import_preview = QPushButton("导入预览")
        self.btn_export = QPushButton("导出JSON")
        self.btn_undo = QPushButton("撤销")
        self.btn_redo = QPushButton("重做")
        self.btn_refresh = QPushButton("刷新")
        self.btn_clear = QPushButton("清空")
        self.btn_refresh.setObjectName("GhostButton")
        self.btn_clear.setObjectName("GhostButton")
        self.btn_import.setObjectName("GhostButton")
        self.btn_import_preview.setObjectName("GhostButton")
        self.btn_export.setObjectName("GhostButton")
        self.btn_undo.setObjectName("GhostButton")
        self.btn_redo.setObjectName("GhostButton")

        btn_grid.addWidget(self.btn_add, 0, 0)
        btn_grid.addWidget(self.btn_update, 0, 1)
        btn_grid.addWidget(self.btn_toggle, 0, 2)
        btn_grid.addWidget(self.btn_delete, 0, 3)
        btn_grid.addWidget(self.btn_enable, 1, 0)
        btn_grid.addWidget(self.btn_disable, 1, 1)
        btn_grid.addWidget(self.btn_batch_delete, 1, 2)
        btn_grid.addWidget(self.btn_refresh, 1, 3)
        btn_grid.addWidget(self.btn_import, 2, 0)
        btn_grid.addWidget(self.btn_export, 2, 1)
        btn_grid.addWidget(self.btn_import_preview, 2, 2)
        btn_grid.addWidget(self.btn_clear, 2, 3)
        btn_grid.addWidget(self.btn_undo, 3, 0)
        btn_grid.addWidget(self.btn_redo, 3, 1)
        lay.addLayout(btn_grid)

        return card

    def _build_right_panel(self) -> QFrame:
        card = QFrame()
        card.setObjectName("Card")
        lay = QVBoxLayout(card)
        lay.setContentsMargins(14, 14, 14, 14)
        lay.setSpacing(10)

        title = QLabel("菜单列表")
        title.setFont(QFont("Microsoft YaHei UI", 12, QFont.Bold))
        lay.addWidget(title)

        self.table = QTableWidget(0, 7)
        self.table.setHorizontalHeaderLabels(["来源", "类型", "条目", "显示名称", "键名", "命令/CLSID", "状态"])
        self.table.setAlternatingRowColors(False)
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.setSelectionMode(QTableWidget.ExtendedSelection)
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.table.setVerticalScrollMode(QAbstractItemView.ScrollPerPixel)
        self.table.setHorizontalScrollMode(QAbstractItemView.ScrollPerPixel)
        self.table.setShowGrid(False)
        self.table.setWordWrap(False)
        self.table.verticalHeader().setVisible(False)
        self.table.verticalHeader().setDefaultSectionSize(30)
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.horizontalHeader().setSortIndicatorShown(True)
        self.table.horizontalHeader().setSortIndicator(self.sort_col, self.sort_order)
        self.table.setColumnWidth(0, 70)
        self.table.setColumnWidth(1, 110)
        self.table.setColumnWidth(2, 90)
        self.table.setColumnWidth(3, 220)
        self.table.setColumnWidth(4, 220)
        self.table.setColumnWidth(5, 520)
        self.table.setColumnWidth(6, 90)
        self.table.verticalScrollBar().setSingleStep(16)
        self.table.horizontalScrollBar().setSingleStep(24)
        lay.addWidget(self.table, 1)
        return card

    def _wire_events(self):
        self.search_edit.textChanged.connect(lambda _t: self._schedule_filter())
        self.hive_combo.currentIndexChanged.connect(lambda _i: self._apply_filter())
        self.category_combo.currentIndexChanged.connect(lambda _i: self._apply_filter())
        self.kind_combo.currentIndexChanged.connect(lambda _i: self._apply_filter())
        self.third_party_combo.currentIndexChanged.connect(lambda _i: self._apply_filter())
        self.filter_timer.timeout.connect(self._on_filter_timeout)

        self.table.itemSelectionChanged.connect(self._on_select)
        self.table.horizontalHeader().sectionClicked.connect(self._on_header_clicked)

        self.btn_add.clicked.connect(self._add_item)
        self.btn_update.clicked.connect(self._update_item)
        self.btn_toggle.clicked.connect(self._toggle_item)
        self.btn_enable.clicked.connect(lambda: self._batch_set_enabled(True))
        self.btn_disable.clicked.connect(lambda: self._batch_set_enabled(False))
        self.btn_delete.clicked.connect(self._delete_single_item)
        self.btn_batch_delete.clicked.connect(self._delete_batch_items)
        self.btn_import.clicked.connect(self._import_json)
        self.btn_import_preview.clicked.connect(self._preview_import)
        self.btn_export.clicked.connect(self._export_json)
        self.btn_undo.clicked.connect(self._undo)
        self.btn_redo.clicked.connect(self._redo)
        self.btn_refresh.clicked.connect(self._refresh)
        self.btn_clear.clicked.connect(self._clear_form)
        self.exit_btn.clicked.connect(self.close)

    def _set_status(self, text: str):
        stamp = datetime.now().strftime("%H:%M:%S")
        self.status_label.setText(f"[{stamp}] {text}")

    def _schedule_filter(self):
        self.filter_timer.start()

    def _on_filter_timeout(self):
        self._apply_filter(self._selected_uid() or self.current_uid)

    def _selected_uid(self) -> str:
        uids = self._selected_uids()
        return uids[0] if uids else ""

    def _selected_uids(self) -> List[str]:
        model = self.table.selectionModel()
        if not model:
            return []
        uids: List[str] = []
        for idx in model.selectedRows(0):
            uid = idx.data(Qt.UserRole)
            if uid:
                uids.append(str(uid))
        return uids

    def _selected_items(self) -> List[MenuItem]:
        return [self.item_map[uid] for uid in self._selected_uids() if uid in self.item_map]

    def _current_item(self) -> Optional[MenuItem]:
        uid = self._selected_uid() or self.current_uid
        return self.item_map.get(uid)

    def _command_get(self) -> str:
        return self.command_edit.toPlainText().strip()

    def _command_set(self, text: str):
        self.command_edit.setPlainText(text)

    @staticmethod
    def _history_change_key(hive: int, path: str) -> str:
        return f"{hive}:{path.lower()}"

    @staticmethod
    def _item_hive_path(item: MenuItem) -> Tuple[int, str]:
        return item.hive, f"{item.path}\\{item.key_name}"

    def _capture_change_before(self, change_map: Dict[str, TreeStateChange], hive: int, path: str):
        key = self._history_change_key(hive, path)
        if key in change_map:
            return
        change_map[key] = TreeStateChange(
            hive=hive,
            path=path,
            before=self.manager.export_tree_if_exists(hive, path),
            after=None,
        )

    def _capture_change_after(self, change_map: Dict[str, TreeStateChange], hive: int, path: str):
        key = self._history_change_key(hive, path)
        change = change_map.get(key)
        if not change:
            change = TreeStateChange(hive=hive, path=path, before=None, after=None)
            change_map[key] = change
        change.after = self.manager.export_tree_if_exists(hive, path)

    def _push_history(self, label: str, change_map: Dict[str, TreeStateChange]):
        changes = [change for change in change_map.values() if change.before != change.after]
        if not changes:
            return
        self.undo_stack.append(HistoryRecord(label=label, changes=changes))
        if len(self.undo_stack) > self.max_history:
            self.undo_stack = self.undo_stack[-self.max_history :]
        self.redo_stack.clear()
        self._update_history_buttons()

    def _update_history_buttons(self):
        self.btn_undo.setEnabled(bool(self.undo_stack))
        self.btn_redo.setEnabled(bool(self.redo_stack))

    def _apply_history_record(self, record: HistoryRecord, undo: bool) -> Tuple[int, int]:
        success = 0
        failed = 0
        changes = sorted(record.changes, key=lambda change: len(change.path), reverse=True)
        for change in changes:
            try:
                target_state = change.before if undo else change.after
                self.manager.apply_tree_state(change.hive, change.path, target_state)
                success += 1
            except Exception:
                failed += 1
        self._refresh(set_status=False)
        return success, failed

    def _undo(self):
        if not self.undo_stack:
            self._warn("提示", "没有可撤销的操作。")
            return
        record = self.undo_stack.pop()
        success, failed = self._apply_history_record(record, undo=True)
        if success > 0:
            self.redo_stack.append(record)
        else:
            self.undo_stack.append(record)
        self._update_history_buttons()
        self._set_status(f"撤销 '{record.label}'：成功 {success}，失败 {failed}")

    def _redo(self):
        if not self.redo_stack:
            self._warn("提示", "没有可重做的操作。")
            return
        record = self.redo_stack.pop()
        success, failed = self._apply_history_record(record, undo=False)
        if success > 0:
            self.undo_stack.append(record)
        else:
            self.redo_stack.append(record)
        self._update_history_buttons()
        self._set_status(f"重做 '{record.label}'：成功 {success}，失败 {failed}")

    @staticmethod
    def _parse_bool(value, default: bool = True) -> bool:
        if isinstance(value, bool):
            return value
        if isinstance(value, (int, float)):
            return bool(value)
        if isinstance(value, str):
            text = value.strip().lower()
            if text in {"1", "true", "yes", "y", "on"}:
                return True
            if text in {"0", "false", "no", "n", "off"}:
                return False
        return default

    @staticmethod
    def _extract_exec_token(command: str) -> str:
        text = command.strip()
        if not text:
            return ""
        if text.startswith('"'):
            end = text.find('"', 1)
            if end > 1:
                return text[1:end].strip()
        parts = text.split(None, 1)
        return parts[0].strip() if parts else ""

    def _is_system_command(self, command: str) -> bool:
        text = command.strip()
        if not text:
            return False

        text_lower = text.lower()
        if "%windir%" in text_lower or "%systemroot%" in text_lower:
            return True

        token = self._extract_exec_token(text)
        if not token:
            return False

        token = token.strip().strip('"').replace("/", "\\")
        token_lower = token.lower()
        basename = os.path.basename(token_lower)
        builtin_names = {
            "cmd.exe",
            "powershell.exe",
            "pwsh.exe",
            "wscript.exe",
            "cscript.exe",
            "rundll32.exe",
            "regsvr32.exe",
            "msiexec.exe",
            "explorer.exe",
            "notepad.exe",
            "control.exe",
            "mshta.exe",
        }
        if basename in builtin_names:
            return True

        expanded = os.path.expandvars(token).replace("/", "\\").rstrip("\\")
        expanded_lower = expanded.lower()
        windows_dir = os.environ.get("WINDIR", r"C:\Windows").replace("/", "\\").rstrip("\\").lower()
        return expanded_lower == windows_dir or expanded_lower.startswith(f"{windows_dir}\\")

    def _is_third_party_item(self, item: MenuItem) -> bool:
        cached = self.third_party_cache.get(item.uid)
        if cached is not None:
            return cached

        if item.kind == "shell":
            command = item.command.strip()
            result = bool(command) and command != SUBMENU_PLACEHOLDER and not self._is_system_command(command)
        else:
            text = f"{item.display_name} {item.key_name} {item.command}".lower()
            system_markers = ("microsoft", "windows", "system", "explorer", "shell")
            result = not any(marker in text for marker in system_markers)

        self.third_party_cache[item.uid] = result
        return result

    def _load_import_entries(self, path: str) -> Optional[List]:
        try:
            with open(path, "r", encoding="utf-8") as fp:
                data = json.load(fp)
        except Exception as exc:
            self._error("导入失败", f"读取 JSON 失败：{exc}")
            return None

        if isinstance(data, dict):
            entries = data.get("items", [])
        elif isinstance(data, list):
            entries = data
        else:
            self._error("导入失败", "JSON 格式无效，需为对象或数组。")
            return None

        if not isinstance(entries, list) or not entries:
            self._warn("提示", "JSON 中没有可导入条目。")
            return None

        return entries

    def _build_import_plan(self, entries: List) -> Tuple[List[ImportPlanItem], Dict[str, int]]:
        current_shell_items = {item.uid: item for item in self.manager.list_items() if item.kind == "shell"}
        scope_write_cache: Dict[str, bool] = {}
        plan: List[ImportPlanItem] = []
        counts = {"total": len(entries), "create": 0, "update": 0, "unchanged": 0, "blocked": 0, "skip": 0}

        for idx, entry in enumerate(entries, 1):
            if not isinstance(entry, dict):
                plan.append(ImportPlanItem(idx, "", "", "", "", True, "skip", "条目不是对象"))
                counts["skip"] += 1
                continue

            scope_id = str(entry.get("scope_id", "")).strip()
            scope = self.manager.scope_map.get(scope_id)
            if not scope or scope.kind != "shell":
                plan.append(ImportPlanItem(idx, scope_id, "", "", "", True, "skip", "scope_id 无效或非 Shell"))
                counts["skip"] += 1
                continue

            display_name = str(entry.get("display_name", "")).strip()
            command = str(entry.get("command", "")).strip()
            enabled = self._parse_bool(entry.get("enabled", True), True)
            has_submenu = self._parse_bool(entry.get("has_submenu", False), False)
            raw_key_name = str(entry.get("key_name", "")).strip()
            key_name = self.manager._slugify(raw_key_name or display_name)

            if has_submenu or (not display_name) or (not command) or command == SUBMENU_PLACEHOLDER:
                plan.append(ImportPlanItem(idx, scope_id, key_name, display_name, command, enabled, "skip", "缺少有效字段或子菜单占位项"))
                counts["skip"] += 1
                continue

            uid = f"{scope_id}|{key_name}"
            existing = current_shell_items.get(uid)
            if existing:
                changed = (
                    existing.display_name != display_name
                    or existing.command != command
                    or existing.enabled != enabled
                )
                if not changed:
                    plan.append(ImportPlanItem(idx, scope_id, key_name, display_name, command, enabled, "unchanged", "与现有内容一致"))
                    counts["unchanged"] += 1
                    continue
                if not self.manager.can_write(existing):
                    plan.append(ImportPlanItem(idx, scope_id, key_name, display_name, command, enabled, "blocked", "目标项不可写"))
                    counts["blocked"] += 1
                    continue
                plan.append(ImportPlanItem(idx, scope_id, key_name, display_name, command, enabled, "update", "将更新现有项"))
                counts["update"] += 1
                current_shell_items[uid] = MenuItem(
                    uid=uid,
                    scope_id=existing.scope_id,
                    hive=existing.hive,
                    hive_name=existing.hive_name,
                    path=existing.path,
                    category=existing.category,
                    kind="shell",
                    key_name=existing.key_name,
                    display_name=display_name,
                    command=command,
                    enabled=enabled,
                    has_submenu=False,
                )
                continue

            can_write_scope = scope_write_cache.get(scope_id)
            if can_write_scope is None:
                can_write_scope = self.manager.can_write_scope(scope_id)
                scope_write_cache[scope_id] = can_write_scope
            if not can_write_scope:
                plan.append(ImportPlanItem(idx, scope_id, key_name, display_name, command, enabled, "blocked", "目标位置不可写"))
                counts["blocked"] += 1
                continue

            plan.append(ImportPlanItem(idx, scope_id, key_name, display_name, command, enabled, "create", "将新增"))
            counts["create"] += 1
            current_shell_items[uid] = MenuItem(
                uid=uid,
                scope_id=scope_id,
                hive=scope.hive,
                hive_name=scope.hive_name,
                path=scope.path,
                category=scope.category,
                kind="shell",
                key_name=key_name,
                display_name=display_name,
                command=command,
                enabled=enabled,
                has_submenu=False,
            )

        return plan, counts

    @staticmethod
    def _format_import_preview(path: str, plan: List[ImportPlanItem], counts: Dict[str, int]) -> str:
        header = [
            f"文件: {path}",
            f"总计: {counts['total']}",
            f"新增: {counts['create']}  更新: {counts['update']}  不变: {counts['unchanged']}",
            f"不可写: {counts['blocked']}  跳过: {counts['skip']}",
            "",
            "明细（最多展示前 120 条）:",
        ]
        action_label = {
            "create": "[新增]",
            "update": "[更新]",
            "unchanged": "[不变]",
            "blocked": "[不可写]",
            "skip": "[跳过]",
        }
        lines: List[str] = []
        for item in plan[:120]:
            label = action_label.get(item.action, "[未知]")
            scope_key = f"{item.scope_id}|{item.key_name}" if item.scope_id else "-"
            name = item.display_name or "(无显示名)"
            lines.append(f"{label} #{item.line_no} {scope_key} {name} - {item.reason}")
        if len(plan) > 120:
            lines.append(f"... 其余 {len(plan) - 120} 条未展示")
        return "\n".join(header + lines)

    def _set_action_state(self, single_editable: bool, batch_editable: bool):
        self.btn_update.setEnabled(single_editable)
        self.btn_toggle.setEnabled(single_editable)
        self.btn_delete.setEnabled(single_editable)
        self.btn_batch_delete.setEnabled(batch_editable)
        self.btn_enable.setEnabled(batch_editable)
        self.btn_disable.setEnabled(batch_editable)

    def _refresh(self, keep_uid: str = "", set_status: bool = True):
        selected_uid = keep_uid or self._selected_uid() or self.current_uid
        if self.refresh_inflight:
            self.pending_refresh = (selected_uid, set_status)
            return

        self.refresh_inflight = True
        self.pending_refresh = None
        self.btn_refresh.setEnabled(False)
        self._set_status("正在刷新菜单数据...")

        worker = threading.Thread(
            target=self._run_refresh_worker,
            args=(selected_uid, set_status),
            daemon=True,
        )
        worker.start()

    def _run_refresh_worker(self, keep_uid: str, set_status: bool):
        try:
            rows = self.manager.list_items()
            self.refresh_emitter.completed.emit(rows, "", keep_uid, set_status)
        except Exception as exc:
            self.refresh_emitter.completed.emit([], str(exc), keep_uid, set_status)

    def _on_refresh_completed(self, rows_obj, error: str, keep_uid: str, set_status: bool):
        self.refresh_inflight = False
        self.btn_refresh.setEnabled(True)

        if error:
            self._error("刷新失败", error)
        else:
            rows = rows_obj if isinstance(rows_obj, list) else []
            self.items = rows
            self.item_map = {item.uid: item for item in self.items}
            self.third_party_cache.clear()
            self._apply_filter(keep_uid)
            if set_status:
                self._set_status(f"已加载 {len(self.items)} 条菜单项")

        if self.pending_refresh:
            queued_keep_uid, queued_status = self.pending_refresh
            self.pending_refresh = None
            self._refresh(queued_keep_uid, queued_status)

    def _apply_filter(self, keep_uid: str = ""):
        if self.filter_timer.isActive():
            self.filter_timer.stop()
        keyword = self.search_edit.text().strip().lower()
        hive_filter = self.hive_combo.currentText()
        category_filter = self.category_combo.currentText()
        kind_filter = self.kind_combo.currentText().lower()
        third_party_filter = self.third_party_combo.currentText()

        rows = self.items
        if hive_filter != "全部":
            rows = [item for item in rows if item.hive_name == hive_filter]
        if category_filter != "全部":
            rows = [item for item in rows if item.category == category_filter]
        if kind_filter == "shell":
            rows = [item for item in rows if item.kind == "shell"]
        elif kind_filter == "handler":
            rows = [item for item in rows if item.kind == "handler"]
        if third_party_filter == "仅第三方":
            rows = [item for item in rows if self._is_third_party_item(item)]
        elif third_party_filter == "仅系统项":
            rows = [item for item in rows if not self._is_third_party_item(item)]
        if keyword:
            rows = [
                item
                for item in rows
                if keyword in item.display_name.lower() or keyword in item.key_name.lower() or keyword in item.command.lower()
            ]

        rows = self._sort_rows(rows)
        self.visible_items = rows
        self._render_rows(rows, keep_uid=keep_uid)
        self.stat_total[1].setText(str(len(self.items)))
        self.stat_visible[1].setText(str(len(rows)))
        self.stat_disabled[1].setText(str(sum(1 for item in rows if item.kind == "shell" and not item.enabled)))

    def _sort_rows(self, rows: List[MenuItem]) -> List[MenuItem]:
        reverse = self.sort_order == Qt.DescendingOrder

        def sort_key(item: MenuItem):
            if self.sort_col == 0:
                return item.hive_name.lower()
            if self.sort_col == 1:
                return item.category.lower()
            if self.sort_col == 2:
                return item.kind.lower()
            if self.sort_col == 3:
                return item.display_name.lower()
            if self.sort_col == 4:
                return item.key_name.lower()
            if self.sort_col == 5:
                return item.command.lower()
            if item.kind == "handler":
                return "2_handler"
            return "0_enabled" if item.enabled else "1_disabled"

        return sorted(rows, key=sort_key, reverse=reverse)

    def _on_header_clicked(self, col: int):
        if self.sort_col == col:
            self.sort_order = Qt.DescendingOrder if self.sort_order == Qt.AscendingOrder else Qt.AscendingOrder
        else:
            self.sort_col = col
            self.sort_order = Qt.AscendingOrder
        self.table.horizontalHeader().setSortIndicator(self.sort_col, self.sort_order)
        self._apply_filter(self._selected_uid() or self.current_uid)
        order_text = "升序" if self.sort_order == Qt.AscendingOrder else "降序"
        header_name = self.table.horizontalHeaderItem(self.sort_col).text()
        self._set_status(f"已按 {header_name} {order_text} 排序")

    def _render_rows(self, rows: List[MenuItem], keep_uid: str = ""):
        blocker = QSignalBlocker(self.table)
        self.table.setUpdatesEnabled(False)
        self.table.clearContents()
        self.table.setRowCount(len(rows))
        alt_bg = QColor("#f8fbff")
        hklm_bg = QColor("#fff8ef")
        disabled_fg = QColor("#94a3b8")
        handler_fg = QColor("#0284c7")

        for row_index, item in enumerate(rows):
            kind_text = "Shell" if item.kind == "shell" else "Handler"
            status_text = "只读" if item.kind == "handler" else ("启用" if item.enabled else "禁用")
            values = [item.hive_name, item.category, kind_text, item.display_name, item.key_name, item.command, status_text]

            for col, val in enumerate(values):
                table_item = QTableWidgetItem(val)
                if col == 0:
                    table_item.setData(Qt.UserRole, item.uid)
                table_item.setFlags(table_item.flags() ^ Qt.ItemIsEditable)

                if row_index % 2 == 1:
                    table_item.setBackground(alt_bg)
                if item.hive_name == "HKLM":
                    table_item.setBackground(hklm_bg)
                if item.kind == "handler":
                    table_item.setForeground(handler_fg)
                elif not item.enabled:
                    table_item.setForeground(disabled_fg)

                self.table.setItem(row_index, col, table_item)

        self.table.setUpdatesEnabled(True)
        del blocker

        if keep_uid:
            self._reselect_uid(keep_uid)
        else:
            self.current_uid = ""
            self._set_action_state(False, False)

    def _reselect_uid(self, uid: str):
        for row in range(self.table.rowCount()):
            item = self.table.item(row, 0)
            if item and item.data(Qt.UserRole) == uid:
                self.table.selectRow(row)
                self.table.scrollToItem(item)
                self.current_uid = uid
                return
        self.current_uid = ""
        self._set_action_state(False, False)

    def _on_select(self):
        selected_items = self._selected_items()
        if not selected_items:
            self._set_action_state(False, False)
            self.enabled_check.setEnabled(True)
            return

        selected_shell = [item for item in selected_items if item.kind == "shell"]
        writable_shell = [item for item in selected_shell if self.manager.can_write(item)]

        if len(selected_items) == 1:
            item = selected_items[0]
            self.current_uid = item.uid
            self.name_edit.setText(item.display_name)
            self.key_edit.setText(item.key_name)
            self.enabled_check.setChecked(item.enabled if item.kind == "shell" else True)
            self.enabled_check.setEnabled(item.kind == "shell")
            self._command_set("" if item.has_submenu else item.command)

            if item.kind == "shell":
                for label, scope_id in self.scope_by_label.items():
                    if scope_id == item.scope_id:
                        self.scope_combo.setCurrentText(label)
                        break

            editable = item.kind == "shell" and self.manager.can_write(item)
            self._set_action_state(editable, editable)
            if item.kind == "handler":
                self._set_status("Handler 条目为只读展示。")
            elif editable:
                self._set_status("已选中可编辑的 Shell 条目。")
            else:
                self._set_status("只读：可能需要管理员权限。")
            return

        self.current_uid = selected_items[0].uid
        self.enabled_check.setEnabled(False)
        self._set_action_state(False, len(writable_shell) > 0)
        self._set_status(
            f"已选择 {len(selected_items)} 项（Shell {len(selected_shell)}，可写 {len(writable_shell)}）"
        )

    def _clear_form(self):
        self.current_uid = ""
        self.name_edit.clear()
        self.key_edit.clear()
        self._command_set("")
        self.enabled_check.setChecked(True)
        self.enabled_check.setEnabled(True)
        if self.scope_labels:
            self.scope_combo.setCurrentIndex(0)
        self.table.clearSelection()
        self._set_action_state(False, False)
        self._set_status("已清空编辑区。")

    def _add_item(self):
        scope_label = self.scope_combo.currentText()
        scope_id = self.scope_by_label.get(scope_label)
        if not scope_id:
            self._error("新增失败", "新增位置无效。")
            return

        try:
            change_map: Dict[str, TreeStateChange] = {}
            uid = self.manager.create_shell_item(
                scope_id=scope_id,
                display_name=self.name_edit.text().strip(),
                command=self._command_get(),
                key_name=self.key_edit.text().strip(),
                enabled=self.enabled_check.isChecked(),
            )
            hive, path = self.manager.uid_to_hive_path(uid)
            self._capture_change_after(change_map, hive, path)
            self._push_history("新增菜单项", change_map)
            self._refresh(keep_uid=uid, set_status=False)
            self._set_status("新增成功。")
        except PermissionError:
            self._error("新增失败", "权限不足。写入 HKLM 可能需要管理员权限。")
        except Exception as exc:
            self._error("新增失败", str(exc))

    def _update_item(self):
        if len(self._selected_items()) != 1:
            self._warn("提示", "更新仅支持单选，请只选择一条 Shell 菜单。")
            return
        item = self._current_item()
        if not item:
            self._warn("提示", "请先选择一条 Shell 菜单。")
            return

        command = self._command_get()
        if item.has_submenu and not command:
            command = item.command

        try:
            change_map: Dict[str, TreeStateChange] = {}
            hive, path = self._item_hive_path(item)
            self._capture_change_before(change_map, hive, path)
            self.manager.update_shell_item(
                item=item,
                display_name=self.name_edit.text().strip(),
                command=command,
                enabled=self.enabled_check.isChecked(),
            )
            self._capture_change_after(change_map, hive, path)
            self._push_history("更新菜单项", change_map)
            self._refresh(keep_uid=item.uid, set_status=False)
            self._set_status("更新成功。")
        except PermissionError:
            self._error("更新失败", "权限不足。写入 HKLM 可能需要管理员权限。")
        except Exception as exc:
            self._error("更新失败", str(exc))

    def _toggle_item(self):
        if len(self._selected_items()) != 1:
            self._warn("提示", "启用/禁用切换仅支持单选。")
            return
        item = self._current_item()
        if not item:
            self._warn("提示", "请先选择一条 Shell 菜单。")
            return
        try:
            change_map: Dict[str, TreeStateChange] = {}
            hive, path = self._item_hive_path(item)
            self._capture_change_before(change_map, hive, path)
            self.manager.set_enabled(item, not item.enabled)
            self._capture_change_after(change_map, hive, path)
            self._push_history("切换启用状态", change_map)
            self._refresh(keep_uid=item.uid, set_status=False)
            self._set_status("状态切换成功。")
        except PermissionError:
            self._error("操作失败", "权限不足。写入 HKLM 可能需要管理员权限。")
        except Exception as exc:
            self._error("操作失败", str(exc))

    def _delete_single_item(self):
        if len(self._selected_items()) != 1:
            self._warn("提示", "单条删除仅支持单选。")
            return
        item = self._current_item()
        if not item:
            self._warn("提示", "请先选择一条 Shell 菜单。")
            return
        if item.kind != "shell" or not self.manager.can_write(item):
            self._warn("提示", "该条目不可删除（可能是 Handler 或无写权限）。")
            return

        ok = QMessageBox.question(
            self,
            "确认删除",
            f"确认删除 '{item.display_name}' ({item.key_name}) 吗？",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No,
        )
        if ok != QMessageBox.Yes:
            return
        try:
            change_map: Dict[str, TreeStateChange] = {}
            hive, path = self._item_hive_path(item)
            self._capture_change_before(change_map, hive, path)
            self.manager.delete_shell_item(item)
            self._capture_change_after(change_map, hive, path)
            self._push_history("删除菜单项", change_map)
            self._refresh(set_status=False)
            self._clear_form()
            self._set_status("单条删除成功。")
        except PermissionError:
            self._error("删除失败", "权限不足。写入 HKLM 可能需要管理员权限。")
        except Exception as exc:
            self._error("删除失败", str(exc))

    def _delete_batch_items(self):
        selected = self._selected_items()
        if not selected:
            self._warn("提示", "请先选择要删除的条目。")
            return
        candidates = [item for item in selected if item.kind == "shell" and self.manager.can_write(item)]
        if not candidates:
            self._warn("提示", "所选条目均不可删除（可能是 Handler 或无写权限）。")
            return

        count = len(candidates)
        ok = QMessageBox.question(
            self,
            "确认删除",
            f"确认删除 {count} 条可写 Shell 菜单吗？",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No,
        )
        if ok != QMessageBox.Yes:
            return
        deleted = 0
        failed = 0
        change_map: Dict[str, TreeStateChange] = {}
        for item in candidates:
            try:
                hive, path = self._item_hive_path(item)
                self._capture_change_before(change_map, hive, path)
                self.manager.delete_shell_item(item)
                self._capture_change_after(change_map, hive, path)
                deleted += 1
            except Exception:
                with suppress(Exception):
                    self._capture_change_after(change_map, hive, path)
                failed += 1
        try:
            self._push_history("批量删除菜单项", change_map)
            self._refresh(set_status=False)
            self._clear_form()
            self._set_status(f"批量删除完成：成功 {deleted}，失败 {failed}")
        except Exception as exc:
            self._error("删除失败", str(exc))

    def _selected_writable_shell(self) -> List[MenuItem]:
        return [item for item in self._selected_items() if item.kind == "shell" and self.manager.can_write(item)]

    def _batch_set_enabled(self, enabled: bool):
        targets = self._selected_writable_shell()
        if not targets:
            self._warn("提示", "请先多选可写的 Shell 条目。")
            return
        success = 0
        failed = 0
        change_map: Dict[str, TreeStateChange] = {}
        for item in targets:
            try:
                hive, path = self._item_hive_path(item)
                self._capture_change_before(change_map, hive, path)
                self.manager.set_enabled(item, enabled)
                self._capture_change_after(change_map, hive, path)
                success += 1
            except Exception:
                with suppress(Exception):
                    self._capture_change_after(change_map, hive, path)
                failed += 1
        self._push_history("批量切换启用状态", change_map)
        self._refresh(set_status=False)
        action = "启用" if enabled else "禁用"
        self._set_status(f"批量{action}完成：成功 {success}，失败 {failed}")

    def _export_json(self):
        selected_shell = [item for item in self._selected_items() if item.kind == "shell"]
        source_items = selected_shell if selected_shell else [item for item in self.visible_items if item.kind == "shell"]
        export_items = [
            item
            for item in source_items
            if (not item.has_submenu) and item.command.strip() and item.command.strip() != SUBMENU_PLACEHOLDER
        ]
        skipped = len(source_items) - len(export_items)
        if not export_items:
            self._warn("提示", "没有可导出的 Shell 条目。")
            return

        path, _ = QFileDialog.getSaveFileName(self, "导出 JSON", "context_menu_backup.json", "JSON Files (*.json)")
        if not path:
            return

        payload = {
            "version": 1,
            "exported_at": datetime.now().isoformat(timespec="seconds"),
            "count": len(export_items),
            "skipped_unexportable": skipped,
            "items": [
                {
                    "scope_id": item.scope_id,
                    "display_name": item.display_name,
                    "key_name": item.key_name,
                    "command": item.command,
                    "enabled": item.enabled,
                }
                for item in export_items
            ],
        }

        try:
            with open(path, "w", encoding="utf-8") as fp:
                json.dump(payload, fp, ensure_ascii=False, indent=2)
            self._set_status(f"导出成功：{len(export_items)} 条，跳过 {skipped} 条 -> {path}")
        except Exception as exc:
            self._error("导出失败", str(exc))

    def _import_json(self):
        path, _ = QFileDialog.getOpenFileName(self, "导入 JSON", "", "JSON Files (*.json)")
        if not path:
            return
        entries = self._load_import_entries(path)
        if entries is None:
            return

        try:
            plan, counts = self._build_import_plan(entries)
        except Exception as exc:
            self._error("导入失败", f"预扫描失败：{exc}")
            return

        actionable = [item for item in plan if item.action in {"create", "update"}]
        if not actionable:
            self._warn("提示", "没有可导入变更（可能全部相同、不可写或无效）。")
            self._set_status(
                f"导入跳过：不变 {counts['unchanged']}，不可写 {counts['blocked']}，跳过 {counts['skip']}"
            )
            return

        ok = QMessageBox.question(
            self,
            "确认导入",
            (
                f"预览结果：新增 {counts['create']}，更新 {counts['update']}，不变 {counts['unchanged']}，"
                f"不可写 {counts['blocked']}，跳过 {counts['skip']}。\n\n确认执行导入吗？"
            ),
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.Yes,
        )
        if ok != QMessageBox.Yes:
            return

        created_count = 0
        updated_count = 0
        fail_count = 0
        change_map: Dict[str, TreeStateChange] = {}
        for item in actionable:
            try:
                hive, path = self.manager.scope_key_to_hive_path(item.scope_id, item.key_name)
                self._capture_change_before(change_map, hive, path)
                _uid, created = self.manager.upsert_shell_item(
                    scope_id=item.scope_id,
                    display_name=item.display_name,
                    command=item.command,
                    key_name=item.key_name,
                    enabled=item.enabled,
                )
                self._capture_change_after(change_map, hive, path)
                if created:
                    created_count += 1
                else:
                    updated_count += 1
            except Exception:
                with suppress(Exception):
                    self._capture_change_after(change_map, hive, path)
                fail_count += 1

        self._push_history("导入 JSON", change_map)
        self._refresh(set_status=False)
        self._set_status(
            f"导入完成：新增 {created_count}，更新 {updated_count}，失败 {fail_count}，"
            f"不变 {counts['unchanged']}，不可写 {counts['blocked']}，跳过 {counts['skip']}"
        )

    def _preview_import(self):
        path, _ = QFileDialog.getOpenFileName(self, "导入预览 (dry-run)", "", "JSON Files (*.json)")
        if not path:
            return
        entries = self._load_import_entries(path)
        if entries is None:
            return
        try:
            plan, counts = self._build_import_plan(entries)
        except Exception as exc:
            self._error("预览失败", f"预扫描失败：{exc}")
            return

        text = self._format_import_preview(path, plan, counts)
        QMessageBox.information(self, "导入预览 (dry-run)", text)
        self._set_status(
            f"预览完成：新增 {counts['create']}，更新 {counts['update']}，不变 {counts['unchanged']}，"
            f"不可写 {counts['blocked']}，跳过 {counts['skip']}"
        )

    def _error(self, title: str, text: str):
        QMessageBox.critical(self, title, text)

    def _warn(self, title: str, text: str):
        QMessageBox.warning(self, title, text)


def main():
    if sys.platform != "win32":
        raise SystemExit("该程序仅支持 Windows。")
    if not _ensure_admin():
        raise SystemExit("需要管理员权限运行。")
    app = QApplication(sys.argv)
    app.setStyleSheet(APP_STYLESHEET)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())


def _is_admin() -> bool:
    with suppress(Exception):
        return bool(ctypes.windll.shell32.IsUserAnAdmin())
    return False


def _quote_arg(arg: str) -> str:
    if not arg:
        return '""'
    if re.search(r'[\s"]', arg):
        return f'"{arg.replace(chr(34), chr(92) + chr(34))}"'
    return arg


def _ensure_admin() -> bool:
    if _is_admin():
        return True

    exe_path = sys.executable
    if getattr(sys, "frozen", False):
        params = " ".join(_quote_arg(arg) for arg in sys.argv[1:])
    else:
        script_path = os.path.abspath(__file__)
        params = " ".join([_quote_arg(script_path)] + [_quote_arg(arg) for arg in sys.argv[1:]])

    ret = ctypes.windll.shell32.ShellExecuteW(None, "runas", exe_path, params, None, 1)
    return int(ret) > 32


if __name__ == "__main__":
    main()
