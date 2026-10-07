# -*- coding: utf-8 -*-
"""
DataTransformer 图形界面(PySide6 / Qt)
=====================================
输入:SQL 数据库(MySQL / PostgreSQL,支持 SSH 隧道)、Excel / JSON / CSV 文件
输出:Excel(.xlsx)/ JSON(.json)/ CSV(.csv)/ SQL 数据库

命名约定(库-表-字段三层一一对应):
    数据库名  == 输出文件名(Excel/JSON/CSV)
    表名      == Excel Sheet 名 == JSON 一级键
    字段名    == Excel 首行单元格 == JSON HeaderFields == CSV 表头

设计说明:
    * 全部业务逻辑位于 dt_core.py(无 Qt 依赖),本文件只负责界面与交互,
      因此 ./data_transformer.sh 的命令行向导与图形界面行为完全一致;
    * SSH 隧道:勾选后由 paramiko 建立本地端口转发,界面里填写【远程服务器上】
      数据库的地址(通常 127.0.0.1),即可像本地库一样导入导出远端 SQL 数据库;
    * 大数据量使用服务端流式游标 + 批量入库,Excel 使用 openpyxl 流式模式。
"""

import contextlib
import html
import os
import sys
import threading
import traceback
from pathlib import Path

from dt_core import (
    ADAPTERS, APP_TITLE, DEFAULT_BATCH, DELIM_LABELS, ENC_LABELS, INPUT_LABELS,
    MODE_KEYS, MODE_LABELS, OUTPUT_LABELS, SSH_DEFAULT_PORT, VERSION,
    FieldMapping, Settings, SourceSpec, SSHConfig, SSHTunnel, TargetSpec,
    config_for, convert, csv_output_paths, describe_db_error, describe_ssh_error,
    json_top_keys, open_reader, paramiko_available, read_csv_header, safe_filename,
    safe_sheet_title, scan_csv_sections, to_int, workbook_sheet_headers,
    config_file_path,
)
from dt_core import run_self_test as core_self_test
from dt_numeric import FidelityError

try:
    from PySide6.QtCore import Qt, QObject, QPoint, QTimer, Signal
    from PySide6.QtGui import QColor, QGuiApplication, QPainter, QPolygon
    from PySide6.QtWidgets import (
        QApplication, QMainWindow, QWidget, QLabel, QLineEdit, QComboBox,
        QCheckBox, QPushButton, QFrame, QGroupBox, QGridLayout, QHBoxLayout,
        QVBoxLayout, QPlainTextEdit, QFileDialog, QMessageBox, QDialog,
        QProgressBar, QTableWidget, QTableWidgetItem,
        QTreeWidget, QTreeWidgetItem, QSpinBox, QScrollArea,
    )
except ImportError as exc:  # pragma: no cover - 服务器上通常没有 PySide6
    print(f"未安装 PySide6({exc})。\n"
          "Linux 服务器可直接使用命令行向导: ./data_transformer.sh\n"
          "如需图形界面: pip install PySide6")
    if "--selftest" in sys.argv:
        raise SystemExit(0 if core_self_test() else 1)
    raise SystemExit(1)

# ---------------------------------------------------------------------------
# 配色 / 样式(clam 主题)
# ---------------------------------------------------------------------------
BG = "#eef2f8"               # 窗体背景
CARD = "#ffffff"             # 卡片背景
NAVY = "#16294d"             # 顶部横幅
TEXT = "#1f2733"             # 主文字
MUTED = "#5b6b7f"            # 次要文字
BORDER = "#c9d4e4"           # 卡片边框
PRIMARY = "#1f6feb"          # 主蓝色(SQL)
GREEN = "#1a9e5c"            # Excel
ORANGE = "#e8833a"           # JSON
PURPLE = "#8e5bd8"           # CSV
RED = "#d64545"              # 错误
TEAL = "#0e9aa7"             # 青色(辅助)
AMBER = "#f0a12e"            # 琥珀(辅助)
INDIGO = "#5a67d8"           # 靛蓝(辅助)
MAGENTA = "#c256b1"          # 品红(辅助)
CORAL = "#e56b4f"            # 珊瑚(辅助)
CYAN = "#22b8cf"             # 湖蓝(辅助)
STATUS_OK = "#1a9e5c"
STATUS_OFF = "#9aa7b8"

ACCENT_STRIP = [PRIMARY, CYAN, TEAL, GREEN, AMBER, ORANGE, CORAL, MAGENTA, PURPLE]

# 需要数据库连接的输入/输出类型(MongoDB 的 Database 对象没有 close(),关闭要走适配器)
DB_KINDS = ("SQL", "Mongo")

# 数据库端口下拉预设(可直接手填其他端口;留空 = 用该类型的默认端口)
DB_PORT_PRESETS = (3306, 5432, 27017)

# 窗口默认 / 最小尺寸(逻辑像素,高 DPI 下由 Qt 换算为物理像素)
DEFAULT_WINDOW_SIZE = (1080, 720)
MIN_WINDOW_SIZE = (900, 560)


def default_window_size():
    """默认 1080x720;屏幕可用区域更小时按可用区域收窄,避免窗口被系统钳制后布局被压扁。"""
    width, height = DEFAULT_WINDOW_SIZE
    screen = QGuiApplication.primaryScreen()
    if screen is not None:
        available = screen.availableGeometry()
        width = min(width, max(available.width() - 40, 640))
        height = min(height, max(available.height() - 40, 480))
    return width, height

LOG_COLORS = {
    "info": "#3a4a5c", "ok": GREEN, "warn": "#b8741a", "err": RED,
    "head": PRIMARY, "section": TEAL, "db": INDIGO,
}

QSS = f"""
* {{ font-family: "Microsoft YaHei"; font-size: 10pt; }}
QMainWindow, QWidget {{ background: {BG}; color: {TEXT}; }}
QLabel#appTitle {{ background: {NAVY}; color: white; font-size: 15pt;
                   font-weight: bold; padding: 6px 16px; }}
QFrame#bar {{ background: #e8f0fb; border-radius: 6px; }}
QLabel#stepIn {{ background: {INDIGO}; color: white; font-weight: bold;
                 padding: 6px 14px; border-radius: 4px; }}
QLabel#stepOut {{ background: {CORAL}; color: white; font-weight: bold;
                  padding: 6px 14px; border-radius: 4px; }}
QLabel#arrow {{ color: {TEAL}; font-weight: bold; font-size: 11pt; }}
QLabel#muted {{ color: {MUTED}; }}
QLabel#chip {{ background: {STATUS_OFF}; color: white; padding: 2px 12px;
               border-radius: 10px; }}
QLabel#chip[state="ok"] {{ background: {STATUS_OK}; }}
QLabel#chip[state="busy"] {{ background: {ORANGE}; }}
QGroupBox {{ background: {CARD}; border: 2px solid {BORDER}; border-radius: 8px;
             margin-top: 18px; padding: 4px 8px; font-weight: bold; }}
QGroupBox::title {{ subcontrol-origin: margin; subcontrol-position: top left;
                    left: 12px; padding: 0 6px; background: {CARD}; color: {TEXT}; }}
QGroupBox[mode="SQL"] {{ border-color: {PRIMARY}; }}
QGroupBox[mode="SQL"]::title {{ color: {PRIMARY}; }}
QGroupBox[mode="Excel"] {{ border-color: {GREEN}; }}
QGroupBox[mode="Excel"]::title {{ color: {GREEN}; }}
QGroupBox[mode="JSON"] {{ border-color: {ORANGE}; }}
QGroupBox[mode="JSON"]::title {{ color: {ORANGE}; }}
QGroupBox[mode="CSV"] {{ border-color: {PURPLE}; }}
QGroupBox[mode="CSV"]::title {{ color: {PURPLE}; }}
QGroupBox#log {{ border-color: {BORDER}; }}
QGroupBox#log::title {{ color: {MUTED}; }}
QPushButton {{ background: #e7edf6; border: 1px solid {BORDER}; border-radius: 5px;
               padding: 5px 14px; }}
QPushButton:hover {{ background: #d5e2f5; }}
QPushButton:disabled {{ background: #f0f3f8; color: {STATUS_OFF}; }}
QPushButton#primary {{ background: {GREEN}; color: white; font-weight: bold;
                       padding: 8px 18px; }}
QPushButton#primary:hover {{ background: #21b36b; }}
QPushButton#blue {{ background: {PRIMARY}; color: white; }}
QPushButton#blue:hover {{ background: #3b82f6; }}
QPushButton#teal {{ background: {TEAL}; color: white; }}
QPushButton#teal:hover {{ background: #12b6c4; }}
QPushButton#amber {{ background: {AMBER}; color: white; }}
QPushButton#amber:hover {{ background: #f7b64e; }}
QPushButton#indigo {{ background: {INDIGO}; color: white; }}
QPushButton#indigo:hover {{ background: #6f7deb; }}
QPushButton#warn {{ background: {ORANGE}; color: white; }}
QPushButton#warn:hover {{ background: #f09645; }}
QLineEdit, QComboBox, QSpinBox {{ background: white; border: 1px solid {BORDER};
                                  border-radius: 4px; padding: 3px 8px; }}
QLineEdit:focus, QComboBox:focus {{ border: 1px solid {PRIMARY}; }}
QComboBox::drop-down {{ border: none; width: 22px;
                        subcontrol-origin: padding; subcontrol-position: top right; }}
QComboBox QAbstractItemView {{ background: white;
                               selection-background-color: {PRIMARY};
                               selection-color: white; }}
QCheckBox {{ font-weight: bold; color: {PRIMARY}; }}
QPlainTextEdit {{ background: #fbfcfe; border: 1px solid {BORDER};
                  border-radius: 4px; font-family: Consolas; font-size: 9pt; }}
QTreeWidget {{ background: #fbfcfe; border: 1px solid {BORDER}; border-radius: 4px; }}
QTreeWidget::item {{ height: 24px; }}
QTreeWidget::item:selected {{ background: #dbe7ff; color: {NAVY}; }}
QScrollArea {{ background: transparent; border: none; }}
QDialog {{ background: {CARD}; }}
"""


class Bridge(QObject):
    """worker 线程 → 主线程的信号桥。"""

    log_s = Signal(str, str)
    status_s = Signal(str, bool)
    tree_s = Signal(dict)
    sheets_s = Signal(list)
    excel_tree_s = Signal(dict)
    keys_s = Signal(list)
    csv_sections_s = Signal(list)
    error_s = Signal(str)
    info_s = Signal(str)
    done_s = Signal()


class DropDownComboBox(QComboBox):
    """下拉框:右侧自绘一个小三角(自定义主题会把 Qt 默认箭头吃掉,看不出能下拉)。

    行为保持标准:点右侧箭头展开候选项,点/双击输入区可直接键入自定义值。
    """

    def paintEvent(self, event):
        super().paintEvent(event)
        if self.width() < 40 or self.height() < 14:
            return
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing, True)
        painter.setPen(Qt.NoPen)
        painter.setBrush(QColor(MUTED))
        right, center = self.width() - 10, self.height() // 2
        painter.drawPolygon(QPolygon([QPoint(right - 8, center - 3),
                                      QPoint(right, center - 3),
                                      QPoint(right - 4, center + 3)]))
        painter.end()


class HierTableSelect(QWidget):
    """层级勾选控件(表 → 字段):按钮 + 弹窗树形勾选。"""

    def __init__(self, parent=None, text="选择数据表与字段"):
        super().__init__(parent)
        self.title_text = text
        self.tables = {}
        self.checked = {}
        self.command = None
        self._tree = None
        self._updating = False
        lay = QHBoxLayout(self)
        lay.setContentsMargins(0, 0, 0, 0)
        self.button = QPushButton(f"{text}: (未加载)")
        self.button.clicked.connect(self._open)
        lay.addWidget(self.button, 1)

    # -- 对外接口 -----------------------------------------------------------
    def set_tables(self, tables, checked=None):
        """tables: {表名: [字段...]};checked: {表名: [勾选字段...]}(None=全选)。"""
        self.tables = {str(table): [str(col) for col in cols] for table, cols in tables.items()}
        if checked is None:
            self.checked = {table: set(cols) for table, cols in self.tables.items()}
        else:
            self.checked = {table: {col for col in cols if col in self.tables[table]}
                            for table, cols in checked.items() if table in self.tables}
        self.button.setEnabled(True)
        self._refresh_text()

    def set_hint(self, text):
        self.tables = {}
        self.checked = {}
        self.button.setText(str(text))
        self.button.setEnabled(False)

    def get_selection(self):
        return {table: [col for col in self.tables[table] if col in self.checked.get(table, set())]
                for table in self.tables if self.checked.get(table)}

    def get_tables(self):
        return [table for table in self.tables if self.checked.get(table)]

    def set_command(self, fn):
        self.command = fn

    # -- 内部实现 -----------------------------------------------------------
    def _refresh_text(self):
        if not self.tables:
            self.button.setText(f"{self.title_text}: (未加载)")
            return
        self.button.setText(f"{self.title_text}: {len(self.get_tables())}/{len(self.tables)} 表"
                            f" · {sum(map(len, self.checked.values()))} 字段")

    def _rebuild_tree(self):
        self._updating = True
        tree = self._tree
        tree.clear()
        for table, cols in self.tables.items():
            top = QTreeWidgetItem([table])
            # 注意:不要加 Qt.ItemIsAutoTristate。它会让 Qt 在子节点变化时
            # 自动重算父节点三态并再次触发 itemChanged,与下面的手动同步逻辑
            # 互相覆盖,导致用户点击字段后勾选状态被立刻还原。
            top.setFlags(top.flags() | Qt.ItemIsUserCheckable)
            selected = self.checked.get(table, set())
            state = (Qt.Checked if selected == set(cols)
                     else Qt.PartiallyChecked if selected else Qt.Unchecked)
            top.setCheckState(0, state)
            top.setData(0, Qt.ItemDataRole.UserRole, ("t", table))
            tree.addTopLevelItem(top)
            for col in cols:
                child = QTreeWidgetItem([col])
                child.setFlags(child.flags() | Qt.ItemIsUserCheckable)
                child.setCheckState(0, Qt.Checked if col in selected else Qt.Unchecked)
                child.setData(0, Qt.ItemDataRole.UserRole, ("c", table, col))
                top.addChild(child)
            top.setExpanded(True)
        self._updating = False

    def _on_item_changed(self, item, _column):
        if self._updating or self._tree is None:
            return
        self._updating = True
        kind = item.data(0, Qt.ItemDataRole.UserRole)
        if kind and kind[0] == "t":
            _, table = kind
            state = Qt.Checked if item.checkState(0) == Qt.PartiallyChecked else item.checkState(0)
            self.checked[table] = set(self.tables[table]) if state == Qt.Checked else set()
            for index in range(item.childCount()):
                item.child(index).setCheckState(0, state)
        elif kind and kind[0] == "c":
            _, table, col = kind
            selected = set(self.checked.get(table, set()))
            selected.add(col) if item.checkState(0) == Qt.Checked else selected.discard(col)
            self.checked[table] = selected
            parent, all_cols = item.parent(), set(self.tables[table])
            parent.setCheckState(0, Qt.Checked if selected == all_cols
                                 else Qt.PartiallyChecked if selected else Qt.Unchecked)
        self._updating = False

    def _open(self):
        if not self.tables:
            return
        dialog = QDialog(self.window())
        dialog.setWindowTitle(self.title_text)
        dialog.resize(580, 500)
        lay = QVBoxLayout(dialog)
        hint = QLabel("第一层:数据表(勾选整表)   ·   第二层:字段(每张表独立勾选)")
        hint.setObjectName("muted")
        lay.addWidget(hint)
        self._tree = QTreeWidget()
        self._tree.setHeaderHidden(True)
        lay.addWidget(self._tree, 1)
        buttons = QHBoxLayout()
        for text, name, slot in (("全部勾选", "teal", self._check_all),
                                 ("全部取消", "warn", self._check_none),
                                 ("确定", "primary", dialog.accept)):
            button = QPushButton(text)
            button.setObjectName(name)
            button.clicked.connect(slot)
            buttons.addWidget(button)
        buttons.insertStretch(2, 1)
        lay.addLayout(buttons)
        self._rebuild_tree()
        self._tree.itemChanged.connect(self._on_item_changed)
        snapshot = {table: set(cols) for table, cols in self.checked.items()}
        accepted = dialog.exec() == QDialog.Accepted
        dialog.deleteLater()
        self._tree = None
        if not accepted:                    # 取消/关闭对话框:回滚勾选,不触发回调
            self.checked = snapshot
            self._refresh_text()
            return
        self._refresh_text()
        if self.command:
            self.command(self.get_selection())

    def _check_all(self):
        self.checked = {table: set(cols) for table, cols in self.tables.items()}
        self._rebuild_tree()

    def _check_none(self):
        self.checked = {}
        self._rebuild_tree()


class App(QMainWindow):
    """Qt 主窗口:输入/输出双面板 + 数据库连接卡片 + SSH 隧道卡片 + 彩色主题。"""

    def __init__(self):
        super().__init__()
        self.setWindowTitle(APP_TITLE)
        self.resize(*default_window_size())
        self.setMinimumSize(*MIN_WINDOW_SIZE)

        self.bridge = Bridge()
        self.conn = None
        self.tunnel = None
        self.adapter = ADAPTERS["MySQL"]
        self._busy = False
        self.selected_table = None
        self.selected_tables = []
        self.selected_columns = {}
        self.excel_selection = {}   # {Sheet名: [勾选字段...]}(Excel→SQL 层级勾选)
        self.excel_sheets = []
        self._auto_sheet = ""
        self._auto_table = ""

        self._build_vars()
        self._build_ui()
        self._connect_signals()
        self._load_config()
        self.on_input_mode_changed()
        self.on_output_mode_changed()
        self._update_ui_state()

    # ------------------------------------------------------------------ 变量
    def _build_vars(self):
        # 界面参数一律以控件为准(避免出现和控件不同步的“影子状态”),这里只放非控件状态
        self.field_mapping = FieldMapping()

    # ------------------------------------------------------------------ UI
    def _build_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        root = QVBoxLayout(central)
        root.setContentsMargins(10, 0, 10, 6)
        root.setSpacing(6)

        title = QLabel(f"{APP_TITLE}  v{VERSION}")
        title.setObjectName("appTitle")
        root.addWidget(title)
        strip = QWidget()
        strip.setFixedHeight(5)
        strip_lay = QHBoxLayout(strip)
        strip_lay.setContentsMargins(0, 0, 0, 0)
        strip_lay.setSpacing(0)
        for color in ACCENT_STRIP:
            segment = QFrame()
            segment.setStyleSheet(f"background: {color}; border: none;")
            strip_lay.addWidget(segment, 1)
        root.addWidget(strip)

        # 上:左“输入配置” / 右“输出配置”;中:SSH 隧道横条;下:日志 + 开始转换
        cards = QWidget()
        cards_lay = QVBoxLayout(cards)
        cards_lay.setContentsMargins(0, 0, 0, 0)
        cards_lay.setSpacing(6)
        top_row = QHBoxLayout()
        top_row.setContentsMargins(0, 0, 0, 0)
        top_row.setSpacing(6)
        self._build_input_card(top_row)
        self._build_output_card(top_row)
        self._build_db_card()
        top_row.setStretch(0, 1)
        top_row.setStretch(1, 1)
        self._build_ssh_card(cards_lay)      # SSH 隧道横条放在最上方
        cards_lay.addLayout(top_row)

        # 窗口比内容还矮(小屏 / 高缩放)时出现纵向滚动条,而不是把卡片压扁到边框吞字
        self.card_scroll = QScrollArea()
        self.card_scroll.setObjectName("cardScroll")
        self.card_scroll.setWidget(cards)
        self.card_scroll.setWidgetResizable(True)
        self.card_scroll.setFrameShape(QFrame.NoFrame)
        self.card_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        root.addWidget(self.card_scroll, 1)

        # 下:运行日志(占大部分) + 右侧按钮列(清空日志 / 初始化 / 开始转换)
        bottom = QHBoxLayout()
        bottom.setContentsMargins(0, 0, 0, 0)
        bottom.setSpacing(6)
        log_group = QGroupBox("④ 运行日志")
        log_group.setObjectName("log")
        log_lay = QVBoxLayout(log_group)
        self.log_text = QPlainTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setMinimumHeight(30)
        log_lay.addWidget(self.log_text, 1)
        bottom.addWidget(log_group, 1)

        run_col = QVBoxLayout()
        run_col.setContentsMargins(0, 0, 0, 0)
        run_col.setSpacing(6)
        self.progress = QProgressBar()
        self.progress.setRange(0, 100)
        self.progress.setValue(0)
        self.progress.setVisible(False)
        run_col.addWidget(self.progress)

        self.clear_btn = QPushButton("清空日志")
        self.clear_btn.setObjectName("blue")
        self.clear_btn.clicked.connect(self.on_clear_log)
        run_col.addWidget(self.clear_btn)

        self.reset_btn = QPushButton("初始化")
        self.reset_btn.setObjectName("warn")
        self.reset_btn.setToolTip("清空界面上所有已填写 / 已修改的参数,并断开数据库与 SSH 隧道")
        self.reset_btn.clicked.connect(self.on_reset_all)
        run_col.addWidget(self.reset_btn)

        self.convert_btn = QPushButton("开始转换 ▶")
        self.convert_btn.setObjectName("primary")
        self.convert_btn.setMinimumHeight(46)
        self.convert_btn.clicked.connect(self.on_convert)
        run_col.addWidget(self.convert_btn)
        run_col.addStretch(1)
        bottom.addLayout(run_col)
        root.addLayout(bottom, 2)

        self.status_chip = QLabel("未连接数据库 | 就绪")
        self.status_chip.setObjectName("chip")
        self.status_chip.setProperty("state", "idle")
        root.addWidget(self.status_chip)

    def _build_input_card(self, parent_lay):
        self.input_card = QGroupBox("① 输入配置")
        self.input_card.setProperty("mode", "SQL")
        lay = QVBoxLayout(self.input_card)
        lay.setSpacing(6)

        head = QHBoxLayout()
        badge = QLabel("输入格式")
        badge.setObjectName("stepIn")
        head.addWidget(badge)
        self.input_combo = DropDownComboBox()
        self.input_combo.addItems(list(INPUT_LABELS.values()))
        head.addWidget(self.input_combo, 1)
        lay.addLayout(head)
        self._input_lay = lay          # _sync_db_card() 把数据库连接块插到这里

        # ---- 输入 = SQL:层级勾选(第一层表,第二层字段) ----
        self.src_sql_frame = QWidget()
        sql_lay = QGridLayout(self.src_sql_frame)
        sql_lay.setContentsMargins(0, 0, 0, 0)
        sql_lay.addWidget(QLabel("表/字段:"), 0, 0)
        self.tree_select = HierTableSelect(text="选择数据表与字段")
        self.tree_select.set_command(self.on_tree_confirmed)
        self.tree_select.setToolTip("第一层勾选数据表(可多选);第二层分别勾选每张表的导出字段")
        sql_lay.addWidget(self.tree_select, 0, 1)
        self.refresh_btn = QPushButton("刷新表")
        self.refresh_btn.setObjectName("teal")
        self.refresh_btn.clicked.connect(self.on_refresh_tables)
        sql_lay.addWidget(self.refresh_btn, 0, 2)
        sql_lay.setColumnStretch(1, 1)
        lay.addWidget(self.src_sql_frame)

        # ---- 输入 = 文件 ----
        self.src_file_frame = QWidget()
        file_lay = QGridLayout(self.src_file_frame)
        file_lay.setContentsMargins(0, 0, 0, 0)
        file_lay.addWidget(QLabel("输入文件:"), 0, 0)
        self.in_file_edit = QLineEdit()
        file_lay.addWidget(self.in_file_edit, 0, 1)
        self.pick_btn = QPushButton("选择文件...")
        self.pick_btn.setObjectName("teal")
        self.pick_btn.clicked.connect(self.on_pick_input_file)
        file_lay.addWidget(self.pick_btn, 0, 2)
        file_lay.setColumnStretch(1, 1)
        lay.addWidget(self.src_file_frame)

        # 文件类型专属参数行(按输入格式切换显示)
        self.file_excel_opts = QWidget()
        excel_lay = QGridLayout(self.file_excel_opts)
        excel_lay.setContentsMargins(0, 0, 0, 0)
        excel_lay.setSpacing(4)
        excel_lay.addWidget(QLabel("工作表(单表转换):"), 0, 0)
        self.in_sheet_combo = DropDownComboBox()
        excel_lay.addWidget(self.in_sheet_combo, 0, 1)
        self.excel_tree_label = QLabel("SQL 导入(层级勾选):")
        self.excel_tree_label.setObjectName("muted")
        excel_lay.addWidget(self.excel_tree_label, 1, 0)
        self.excel_tree = HierTableSelect(text="选择工作表与字段")
        self.excel_tree.set_command(self.on_excel_tree_confirmed)
        excel_lay.addWidget(self.excel_tree, 1, 1)
        excel_lay.setColumnStretch(1, 1)
        lay.addWidget(self.file_excel_opts)

        self.file_json_opts = QWidget()
        json_lay = QGridLayout(self.file_json_opts)
        json_lay.setContentsMargins(0, 0, 0, 0)
        json_lay.addWidget(QLabel("一级键(表名):"), 0, 0)
        self.in_key_combo = DropDownComboBox()
        json_lay.addWidget(self.in_key_combo, 0, 1)
        json_lay.setColumnStretch(1, 1)
        lay.addWidget(self.file_json_opts)

        self.file_csv_opts = QWidget()
        csv_lay = QGridLayout(self.file_csv_opts)
        csv_lay.setContentsMargins(0, 0, 0, 0)
        csv_lay.addWidget(QLabel("分节(Sheet):"), 0, 0)
        self.in_section_combo = DropDownComboBox()
        csv_lay.addWidget(self.in_section_combo, 0, 1)
        csv_lay.addWidget(QLabel("分隔符:"), 0, 2)
        self.in_delim_combo = DropDownComboBox()
        self.in_delim_combo.addItems(list(DELIM_LABELS))
        csv_lay.addWidget(self.in_delim_combo, 0, 3)
        csv_lay.addWidget(QLabel("编码:"), 0, 4)
        self.in_enc_combo = DropDownComboBox()
        self.in_enc_combo.addItems(list(ENC_LABELS))
        csv_lay.addWidget(self.in_enc_combo, 0, 5)
        csv_lay.setColumnStretch(1, 1)
        lay.addWidget(self.file_csv_opts)
        lay.addStretch(1)
        parent_lay.addWidget(self.input_card, 1)

    def _build_db_card(self):
        """数据库连接块:由 _sync_db_card() 决定放进左侧(输入为数据库)还是右侧(输出为数据库)。"""
        self.db_card = QGroupBox("数据库连接")
        self.db_card.setProperty("mode", "SQL")
        db_lay = QGridLayout(self.db_card)
        db_lay.setSpacing(6)

        db_lay.addWidget(QLabel("类型:"), 0, 0)
        self.db_type_combo = DropDownComboBox()
        self.db_type_combo.addItems(list(ADAPTERS))
        db_lay.addWidget(self.db_type_combo, 0, 1, 1, 3)

        db_lay.addWidget(QLabel("主机:"), 1, 0)
        self.host_edit = QLineEdit("127.0.0.1")
        db_lay.addWidget(self.host_edit, 1, 1, 1, 3)

        db_lay.addWidget(QLabel("端口:"), 2, 0)
        self.port_edit = DropDownComboBox()      # 常用端口可选,也可直接手工填写
        self.port_edit.setEditable(True)
        self.port_edit.addItems([str(port) for port in DB_PORT_PRESETS])
        self.port_edit.setCurrentText("")
        self.port_edit.lineEdit().setPlaceholderText("留空=默认端口")
        self.port_edit.setToolTip("点击下拉可选 3306 / 5432 / 27017,也可直接输入自定义端口;"
                                  "留空则用当前数据库类型的默认端口")
        db_lay.addWidget(self.port_edit, 2, 1)
        db_lay.addWidget(QLabel("用户名:"), 2, 2)
        self.user_edit = QLineEdit("")
        db_lay.addWidget(self.user_edit, 2, 3)

        db_lay.addWidget(QLabel("密码:"), 3, 0)
        self.pwd_edit = QLineEdit("")
        self.pwd_edit.setEchoMode(QLineEdit.Password)
        db_lay.addWidget(self.pwd_edit, 3, 1)
        db_lay.addWidget(QLabel("数据库:"), 3, 2)
        self.database_edit = QLineEdit("")
        db_lay.addWidget(self.database_edit, 3, 3)

        db_lay.addWidget(QLabel("超时(秒):"), 4, 0)
        self.timeout_edit = QLineEdit("10")
        db_lay.addWidget(self.timeout_edit, 4, 1)

        # MongoDB 专用:连接 URI 与认证库(其他数据库类型自动隐藏)
        self.mongo_uri_label = QLabel("连接 URI:")
        self.mongo_uri_edit = QLineEdit("")
        self.mongo_uri_edit.setPlaceholderText("(可留空)填了 URI 会忽略上面的主机/端口/账号,也不走 SSH 隧道")
        self.mongo_auth_label = QLabel("认证库:")
        self.mongo_auth_edit = QLineEdit("")
        self.mongo_auth_edit.setPlaceholderText("用户建在哪个库就填哪个(常见填 admin);留空=用【数据库】名")
        db_lay.addWidget(self.mongo_uri_label, 5, 0)
        db_lay.addWidget(self.mongo_uri_edit, 5, 1, 1, 3)
        db_lay.addWidget(self.mongo_auth_label, 6, 0)
        db_lay.addWidget(self.mongo_auth_edit, 6, 1, 1, 3)

        btn_box = QHBoxLayout()
        self.test_btn = QPushButton("测试连接")
        self.test_btn.setObjectName("amber")
        self.test_btn.clicked.connect(self.on_test_connection)
        self.conn_btn = QPushButton("连接并加载表")
        self.conn_btn.setObjectName("blue")
        self.conn_btn.clicked.connect(self.on_connect)
        self.disc_btn = QPushButton("断开")
        self.disc_btn.clicked.connect(self.disconnect)
        for button in (self.test_btn, self.conn_btn, self.disc_btn):
            btn_box.addWidget(button)
        btn_box.addStretch(1)
        db_lay.addLayout(btn_box, 7, 0, 1, 4)

        db_lay.setColumnStretch(1, 1)
        db_lay.setColumnStretch(3, 1)

    def _build_ssh_card(self, parent_lay):
        self.ssh_card = QGroupBox("SSH 隧道(连接指定 IP 服务器上的 SQL 数据库)")
        self.ssh_card.setProperty("mode", "SQL")
        lay = QGridLayout(self.ssh_card)

        self.ssh_check = QCheckBox("启用 SSH 隧道")
        self.ssh_host_edit = QLineEdit("")
        self.ssh_port_edit = QLineEdit(str(SSH_DEFAULT_PORT))
        self.ssh_port_edit.setFixedWidth(70)
        self.ssh_user_edit = QLineEdit("")
        self.ssh_pwd_edit = QLineEdit("")
        self.ssh_pwd_edit.setEchoMode(QLineEdit.Password)
        for label, widget, column in (("SSH 主机:", self.ssh_host_edit, 2),
                                      ("端口:", self.ssh_port_edit, 4),
                                      ("用户名:", self.ssh_user_edit, 6),
                                      ("密码:", self.ssh_pwd_edit, 8)):
            lay.addWidget(QLabel(label), 0, column - 1)
            lay.addWidget(widget, 0, column)
        for column in (2, 6, 8):                 # 输入框吸收多余宽度,连接按钮靠右上
            lay.setColumnStretch(column, 1)

        lay.addWidget(self.ssh_check, 0, 0)

        self.ssh_conn_btn = QPushButton("连接 SSH")
        self.ssh_conn_btn.setObjectName("indigo")
        self.ssh_conn_btn.clicked.connect(self.on_ssh_connect)
        self.ssh_disc_btn = QPushButton("断开 SSH")
        self.ssh_disc_btn.clicked.connect(self.on_ssh_disconnect)
        lay.addWidget(self.ssh_conn_btn, 0, 10)
        lay.addWidget(self.ssh_disc_btn, 0, 11)

        hint = QLabel("启用后,下方【输入配置】填远程服务器上的数据库地址"
                      "(如 127.0.0.1:5432 / 27017),本机自动走隧道端口。")
        if not paramiko_available():
            hint.setText(hint.text() + "\n当前未安装 paramiko,SSH 隧道不可用:"
                                       "pip install paramiko(文件互转不受影响)")
            hint.setStyleSheet(f"color: {RED};")
        hint.setObjectName("muted")
        hint.setWordWrap(True)
        self.ssh_state = QLabel("SSH 隧道:未连接")
        self.ssh_state.setObjectName("muted")
        lay.addWidget(hint, 1, 0, 1, 9)
        lay.addWidget(self.ssh_state, 1, 9, 1, 3)
        parent_lay.addWidget(self.ssh_card)

    def _build_output_card(self, parent_lay):
        self.output_card = QGroupBox("② 输出配置")
        self.output_card.setProperty("mode", "Excel")
        lay = QVBoxLayout(self.output_card)
        lay.setSpacing(6)

        head = QHBoxLayout()
        badge = QLabel("输出格式")
        badge.setObjectName("stepOut")
        head.addWidget(badge)
        self.output_combo = DropDownComboBox()
        self.output_combo.addItems(list(OUTPUT_LABELS.values()))
        head.addWidget(self.output_combo, 1)
        lay.addLayout(head)
        self._output_lay = lay         # _sync_db_card() 有可能把数据库连接块插到这里

        self.out_db_frame = QWidget()
        out_db_lay = QGridLayout(self.out_db_frame)
        out_db_lay.setContentsMargins(0, 0, 0, 0)
        out_db_lay.addWidget(QLabel("目标表名:"), 0, 0)
        self.out_table_edit = QLineEdit("")
        out_db_lay.addWidget(self.out_table_edit, 0, 1)
        out_db_lay.addWidget(QLabel("写入模式:"), 1, 0)
        self.out_mode_combo = DropDownComboBox()
        self.out_mode_combo.addItems(MODE_LABELS)
        out_db_lay.addWidget(self.out_mode_combo, 1, 1)
        out_db_lay.addWidget(QLabel("每批行数:"), 2, 0)
        self.batch_spin = QSpinBox()
        self.batch_spin.setRange(1, 1000000)
        self.batch_spin.setValue(DEFAULT_BATCH)
        out_db_lay.addWidget(self.batch_spin, 2, 1)
        out_db_lay.setColumnStretch(1, 1)
        lay.addWidget(self.out_db_frame)

        self.out_shared_hint = QLabel("输出与输入都是数据库:两者共用同一个连接(参数见左侧输入区)。")
        self.out_shared_hint.setObjectName("muted")
        self.out_shared_hint.setWordWrap(True)
        self.out_shared_hint.setVisible(False)
        lay.addWidget(self.out_shared_hint)

        # 字段映射:自动同名 + 手动覆盖(对应 Sheet 名 ↔ 表名、首行表头 ↔ 字段名规则)
        map_row = QHBoxLayout()
        self.map_btn = QPushButton("字段映射预览 / 调整...")
        self.map_btn.setObjectName("indigo")
        self.map_btn.clicked.connect(self.on_edit_mapping)
        self.map_label = QLabel("字段映射: 同名映射")
        self.map_label.setObjectName("muted")
        self.map_label.setWordWrap(True)
        map_row.addWidget(self.map_btn)
        map_row.addWidget(self.map_label, 1)
        lay.addLayout(map_row)

        self.out_file_frame = QWidget()
        out_file_lay = QGridLayout(self.out_file_frame)
        out_file_lay.setContentsMargins(0, 0, 0, 0)
        out_file_lay.addWidget(QLabel("Sheet名/JSON键:"), 0, 0)
        self.out_sheet_edit = QLineEdit("")
        out_file_lay.addWidget(self.out_sheet_edit, 0, 1)
        hint = QLabel("(Excel 用 Sheet 名,JSON 用一级键,CSV 忽略此项)")
        hint.setObjectName("muted")
        hint.setWordWrap(True)
        out_file_lay.addWidget(hint, 1, 0, 1, 2)
        out_file_lay.setColumnStretch(1, 1)
        lay.addWidget(self.out_file_frame)

        self.out_csv_opts = QWidget()
        out_csv_lay = QGridLayout(self.out_csv_opts)
        out_csv_lay.setContentsMargins(0, 0, 0, 0)
        out_csv_lay.addWidget(QLabel("分隔符:"), 0, 0)
        self.out_delim_combo = DropDownComboBox()
        self.out_delim_combo.addItems(list(DELIM_LABELS))
        out_csv_lay.addWidget(self.out_delim_combo, 0, 1)
        out_csv_lay.addWidget(QLabel("编码:"), 0, 2)
        self.out_enc_combo = DropDownComboBox()
        self.out_enc_combo.addItems(list(ENC_LABELS))
        out_csv_lay.addWidget(self.out_enc_combo, 0, 3)
        hint = QLabel("(SQL→CSV 时:目录名 = 数据库名,每个表一个 .csv 文件)")
        hint.setObjectName("muted")
        hint.setWordWrap(True)
        out_csv_lay.addWidget(hint, 1, 0, 1, 4)
        out_csv_lay.setColumnStretch(1, 1)
        lay.addWidget(self.out_csv_opts)
        lay.addStretch(1)
        parent_lay.addWidget(self.output_card, 1)

    def _connect_signals(self):
        bridge = self.bridge
        bridge.log_s.connect(self._on_log)
        bridge.status_s.connect(self._on_status)
        bridge.tree_s.connect(self._on_tree_data)
        bridge.sheets_s.connect(self._on_sheets)
        bridge.excel_tree_s.connect(self._on_excel_tree_data)
        bridge.keys_s.connect(self._on_keys)
        bridge.csv_sections_s.connect(self._on_csv_sections)
        bridge.error_s.connect(lambda message: QMessageBox.critical(self, "错误", message))
        bridge.info_s.connect(lambda message: QMessageBox.information(self, "提示", message))
        bridge.done_s.connect(self._on_done)
        self.input_combo.currentTextChanged.connect(self.on_input_mode_changed)
        self.output_combo.currentTextChanged.connect(self.on_output_mode_changed)
        self.db_type_combo.currentTextChanged.connect(self.on_db_type_changed)
        self.in_sheet_combo.currentTextChanged.connect(self.on_in_sheet_selected)
        self.in_key_combo.currentTextChanged.connect(self.on_in_key_selected)
        self.in_section_combo.currentTextChanged.connect(self.on_in_section_selected)
        self.ssh_check.toggled.connect(lambda *_: self._update_ui_state())

        # 连接参数 / 输入文件变化时,刷新按钮置灰/亮起状态
        for edit in (self.host_edit, self.user_edit, self.pwd_edit,
                     self.database_edit, self.timeout_edit, self.in_file_edit,
                     self.ssh_host_edit, self.ssh_port_edit, self.ssh_user_edit,
                     self.ssh_pwd_edit):
            edit.textChanged.connect(lambda *_: self._update_ui_state())
        self.port_edit.currentTextChanged.connect(lambda *_: self._update_ui_state())

    # ------------------------------------------------------- 线程 / 消息桥
    def log(self, message, level="info"):
        self.bridge.log_s.emit(str(message), level)

    def set_status(self, message, connected=False):
        self.bridge.status_s.emit(str(message), bool(connected))

    def _repolish(self, widget):
        widget.style().unpolish(widget)
        widget.style().polish(widget)

    def _set_chip_state(self, state):
        self.status_chip.setProperty("state", state)
        self._repolish(self.status_chip)

    def _on_log(self, message, level):
        color = LOG_COLORS.get(level, LOG_COLORS["info"])
        escaped = html.escape(message).replace("\n", "<br>")
        self.log_text.appendHtml(f'<span style="color:{color};">{escaped}</span>')

    def _on_status(self, message, connected):
        self.status_chip.setText(message)
        self._set_chip_state("ok" if connected else "idle")

    def on_clear_log(self):
        self.log_text.clear()
        self.log("日志已清空", "info")

    def on_reset_all(self):
        """初始化:清空界面上所有可填写 / 可修改的参数,并断开数据库与 SSH 隧道。"""
        if self._busy:
            QMessageBox.information(self, "提示", "已有任务正在执行,请等待完成后再试。")
            return
        if QMessageBox.question(
                self, "初始化",
                "将清空界面上所有已填写 / 已修改的参数,并断开数据库连接与 SSH 隧道。\n"
                "确定继续吗?",
                QMessageBox.Yes | QMessageBox.No, QMessageBox.No) != QMessageBox.Yes:
            return
        self._stop_tunnel()
        self._disconnect_conn()
        self.input_combo.setCurrentText(INPUT_LABELS["SQL"])
        self.output_combo.setCurrentText(OUTPUT_LABELS["Excel"])
        self.db_type_combo.setCurrentText("MySQL")
        self.host_edit.setText("127.0.0.1")
        for edit in (self.user_edit, self.pwd_edit, self.database_edit,
                     self.mongo_uri_edit, self.mongo_auth_edit, self.ssh_host_edit,
                     self.ssh_user_edit, self.ssh_pwd_edit,
                     self.in_file_edit, self.out_table_edit, self.out_sheet_edit):
            edit.clear()
        self.port_edit.setCurrentText("")
        self.timeout_edit.setText("10")
        self.ssh_port_edit.setText(str(SSH_DEFAULT_PORT))
        self.ssh_check.setChecked(False)
        self.in_sheet_combo.clear()
        self.in_key_combo.clear()
        self.in_section_combo.clear()
        self.in_delim_combo.setCurrentIndex(0)
        self.in_enc_combo.setCurrentIndex(0)
        self.tree_select.set_hint("选择数据表与字段")
        self.excel_tree.set_hint("选择工作表与字段")
        self.selected_table = None
        self.selected_tables = []
        self.selected_columns = {}
        self.excel_selection = {}
        self.excel_sheets = []
        self.out_mode_combo.setCurrentIndex(0)
        self.batch_spin.setValue(DEFAULT_BATCH)
        self.out_delim_combo.setCurrentIndex(0)
        self.out_enc_combo.setCurrentIndex(0)
        self.field_mapping = FieldMapping()
        self.map_label.setText("字段映射: 同名映射")
        self._auto_sheet = ""
        self._auto_table = ""
        self.log_text.clear()
        self.log("已初始化界面参数:所有输入项恢复默认,数据库与 SSH 隧道已断开", "warn")
        self._refresh_status_chip()
        self._update_ui_state()

    def run_task(self, fn, describe=None):
        """在后台线程执行 fn;describe 用于把异常翻译成中文提示。"""
        if self._busy:
            QMessageBox.information(self, "提示", "已有任务正在执行,请等待完成后再试。")
            return False
        self._busy = True
        self.status_chip.setText("运行中...")
        self._set_chip_state("busy")
        self.progress.setRange(0, 0)              # 不确定进度条(忙碌指示)
        self.progress.setVisible(True)
        self._update_ui_state()
        cfg = self.get_cfg()
        adapter = self.adapter
        describe = describe or (lambda exc: describe_db_error(exc, cfg, adapter))

        def worker():
            try:
                fn()
            except Exception as exc:  # noqa: BLE001 - 统一转成中文提示弹窗
                self.log(f"发生错误: {exc}", "err")
                self.log(traceback.format_exc(), "err")
                self.bridge.error_s.emit(describe(exc))
            finally:
                self.bridge.done_s.emit()

        threading.Thread(target=worker, daemon=True).start()
        return True

    def _on_done(self):
        self._busy = False
        self.progress.setRange(0, 100)
        self.progress.setValue(0)
        self.progress.setVisible(False)
        self._save_config()          # 只在主线程读写界面控件,避免跨线程访问 Qt 对象
        self._refresh_status_chip()
        self._update_ui_state()

    def _refresh_status_chip(self):
        parts = [f"{'已连接数据库' if self.conn is not None else '未连接数据库'}"]
        if self.tunnel is not None and self.tunnel.active:
            parts.append(f"SSH 隧道 {self.tunnel.local_port}")
        parts.append("就绪")
        self.status_chip.setText(" | ".join(parts))
        self._set_chip_state("ok" if self.conn is not None else "idle")

    def _on_tree_data(self, table_columns):
        if not table_columns:                # 连上了但库里没有表/集合:按钮上直接说明
            self.tree_select.set_hint("数据库里没有可用的表/集合")
            self.on_tree_confirmed({})
            return
        self.tree_select.set_tables(table_columns)
        self.on_tree_confirmed(self.tree_select.get_selection())

    def _on_sheets(self, names):
        self.in_sheet_combo.blockSignals(True)
        self.in_sheet_combo.clear()
        self.in_sheet_combo.addItems(names)
        self.in_sheet_combo.blockSignals(False)
        if names:
            self.in_sheet_combo.setCurrentIndex(0)
            self.on_in_sheet_selected()
        self._update_ui_state()

    def _on_excel_tree_data(self, sheet_headers):
        """Excel 文件解析完成:{Sheet名: [表头字段...]}。"""
        self.excel_tree.set_tables(sheet_headers)
        self.on_excel_tree_confirmed(self.excel_tree.get_selection())

    def _on_keys(self, keys):
        self.in_key_combo.blockSignals(True)
        self.in_key_combo.clear()
        self.in_key_combo.addItems(keys)
        self.in_key_combo.blockSignals(False)
        if keys:
            self.in_key_combo.setCurrentIndex(0)
            self.on_in_key_selected()
        self._update_ui_state()

    def _on_csv_sections(self, names):
        self.in_section_combo.blockSignals(True)
        self.in_section_combo.clear()
        self.in_section_combo.addItems(names or ["(普通CSV,无分节)"])
        self.in_section_combo.blockSignals(False)
        self.in_section_combo.setCurrentIndex(0)
        self.on_in_section_selected()
        self._update_ui_state()

    # ------------------------------------------------------------- 模式切换
    @staticmethod
    def _key_of(value, mapping):
        return next((key for key, label in mapping.items() if label == value), None)

    def input_key(self):
        return self._key_of(self.input_combo.currentText(), INPUT_LABELS)

    def output_key(self):
        return self._key_of(self.output_combo.currentText(), OUTPUT_LABELS)

    def on_input_mode_changed(self, *_):
        key = self.input_key()
        self.input_card.setProperty("mode", key)
        self._repolish(self.input_card)
        sql_mode = key == "SQL"
        self.src_sql_frame.setVisible(sql_mode)
        self.src_file_frame.setVisible(not sql_mode)
        self.file_excel_opts.setVisible(key == "Excel")
        self.file_json_opts.setVisible(key == "JSON")
        self.file_csv_opts.setVisible(key == "CSV")
        self._sync_db_card()
        self._sync_excel_tree()
        self._auto_fill_output_names()
        self._update_ui_state()
        self.log(f"输入格式切换为: {INPUT_LABELS[key]}", "info")

    def on_output_mode_changed(self, *_):
        key = self.output_key()
        self.output_card.setProperty("mode", key)
        self._repolish(self.output_card)
        sql_mode = key == "SQL"
        self.out_db_frame.setVisible(sql_mode)
        self.out_file_frame.setVisible(not sql_mode)
        self.out_csv_opts.setVisible(key == "CSV")
        self._sync_db_card()
        self._sync_excel_tree()
        self._auto_fill_output_names()
        self._update_ui_state()
        self.log(f"输出格式切换为: {OUTPUT_LABELS[key]}", "info")

    def _sync_db_card(self):
        """数据库连接块跟着需要它的那一侧走:输入=数据库放左侧,否则(输出=数据库)放右侧。"""
        in_sql = self.input_key() == "SQL"
        out_sql = self.output_key() == "SQL"
        need_db = in_sql or out_sql
        self.ssh_card.setVisible(need_db)
        self.db_card.setVisible(need_db)
        self.out_shared_hint.setVisible(in_sql and out_sql)
        if not need_db:
            return
        holder, slot = ((self.input_card, self._input_lay) if in_sql
                        else (self.output_card, self._output_lay))
        if self.db_card.parentWidget() is not holder:
            old = self.db_card.parentWidget()
            if old is not None and old.layout() is not None:
                old.layout().removeWidget(self.db_card)
            self.db_card.setParent(None)
            slot.insertWidget(1, self.db_card)
        self.db_card.setTitle("数据库连接(输入 / 输出共用同一连接)" if in_sql and out_sql
                              else "数据库连接(输入源)" if in_sql
                              else "数据库连接(输出目标)")
        self._sync_db_type_rows()

    def _convert_ready(self, in_key, out_key):
        """判断“开始转换”按钮是否可点击:输入/输出都就绪才亮起。"""
        if in_key == "SQL":
            return bool(self.selected_columns)
        path = self.in_file_edit.text()
        if not (path and os.path.isfile(path)):
            return False
        if in_key == "Excel":
            return (bool(self.excel_selection) if out_key == "SQL"
                    else bool(self.in_sheet_combo.currentText()))
        if in_key == "JSON":
            return bool(self.in_key_combo.currentText())
        if in_key == "CSV":
            return bool(self.in_section_combo.currentText())
        return False

    def _update_ui_state(self):
        """统一刷新按钮/输入项的置灰(不可点击)与亮起(可点击)状态。

        规则:
        * 任务运行中 → 所有操作按钮与输入项置灰;
        * 未填写完整 主机/用户名/数据库 → 测试连接、连接并加载表 置灰
          (MongoDB 允许用连接 URI 代替主机/库名,也不强制账号);
        * 未连接 → 断开 置灰;
        * 启用 SSH 但未填写完整 → 连接 SSH 置灰;
        * 输入/输出未就绪 → 开始转换 置灰;
        * 层级勾选控件只要还没加载过数据 → 置灰。
        """
        busy = self._busy
        in_key = self.input_key()
        out_key = self.output_key()
        connected = self.conn is not None
        need_db = in_key == "SQL" or out_key == "SQL"
        host = self.host_edit.text().strip()
        database = self.database_edit.text().strip()
        if self.adapter.NAME == "MongoDB":
            # MongoDB 可以用连接 URI 代替主机,也允许免账号认证的实例
            uri = self.mongo_uri_edit.text().strip()
            db_ready = bool(host or uri) and bool(database or uri)
        else:
            db_ready = all((host, self.user_edit.text().strip(), database))
        can_db = connected or db_ready
        ssh_on = self.ssh_check.isChecked()
        ssh_ready = all((self.ssh_host_edit.text().strip(), self.ssh_user_edit.text().strip(),
                         self.ssh_pwd_edit.text()))

        self.test_btn.setEnabled(not busy and need_db and db_ready)
        self.conn_btn.setEnabled(not busy and need_db and db_ready)
        self.disc_btn.setEnabled(not busy and connected)
        self.refresh_btn.setEnabled(not busy and in_key == "SQL" and can_db)
        self.pick_btn.setEnabled(not busy and in_key in ("Excel", "JSON", "CSV"))
        self.tree_select.button.setEnabled(not busy and bool(self.tree_select.tables))
        self.excel_tree.button.setEnabled(not busy and bool(self.excel_tree.tables))
        self.ssh_conn_btn.setEnabled(not busy and need_db and ssh_on and ssh_ready)
        self.ssh_disc_btn.setEnabled(not busy and self.tunnel is not None)
        self.reset_btn.setEnabled(not busy)
        self.convert_btn.setEnabled(
            not busy and (can_db if need_db else True)
            and self._convert_ready(in_key, out_key))

        # 任务执行期间禁止修改配置/输入,避免后台任务读取到中途变化的数据
        for widget in (self.input_combo, self.output_combo, self.db_type_combo,
                       self.host_edit, self.port_edit, self.user_edit,
                       self.pwd_edit, self.database_edit, self.timeout_edit,
                       self.in_file_edit, self.in_sheet_combo, self.in_key_combo,
                       self.in_section_combo, self.in_delim_combo, self.in_enc_combo,
                       self.out_table_edit, self.out_mode_combo, self.batch_spin,
                       self.out_sheet_edit, self.out_delim_combo, self.out_enc_combo,
                       self.ssh_check, self.ssh_host_edit, self.ssh_port_edit,
                       self.ssh_user_edit, self.ssh_pwd_edit,
                       self.mongo_uri_edit, self.mongo_auth_edit, self.map_btn):
            widget.setEnabled(not busy)
        for widget in (self.ssh_host_edit, self.ssh_port_edit, self.ssh_user_edit,
                       self.ssh_pwd_edit):
            widget.setEnabled(not busy and ssh_on)
        self.ssh_state.setText(
            f"SSH 隧道:{self.tunnel.describe()}" if self.tunnel is not None
            else ("SSH 隧道:未连接(勾选并点击“连接 SSH”)" if ssh_on else "SSH 隧道:未启用"))
        self._sync_ssh_state()
        self._sync_card_scroll()

    def _sync_card_scroll(self):
        """滚动区最小高度跟随卡片内容:空间够时完整显示,不够时才出现纵向滚动条。

        数据库连接块会在左右两栏之间搬动,Qt 要等一拍事件循环才刷新尺寸提示,
        所以立刻量一次之后再延后量一次,避免用过期的最小高度把卡片压扁。
        """
        self._apply_scroll_minimum()
        QTimer.singleShot(0, self._apply_scroll_minimum)

    def _apply_scroll_minimum(self):
        """滚动区最小高度 = 卡片内容的真实最小高度,但绝不吃掉日志/状态栏的位置。

        卡片确实放不下时(小屏 / 高缩放 / 数据库连接块搬到右侧),滚动区改为内部滚动,
        而不是把整个窗口撑到超出屏幕、把底部内容挤没。
        """
        cards = self.card_scroll.widget()
        cards.layout().activate()
        root = self.centralWidget().layout()
        others = sum(root.itemAt(index).minimumSize().height()
                     for index in range(root.count())
                     if root.itemAt(index).widget() is not self.card_scroll)
        margins = root.contentsMargins()
        others += root.spacing() * (root.count() - 1) + margins.top() + margins.bottom()
        needed = max(cards.minimumSizeHint().height(), cards.layout().minimumSize().height())
        self.card_scroll.setMinimumHeight(max(min(needed, self.height() - others), 80))

    def resizeEvent(self, event):
        super().resizeEvent(event)
        if getattr(self, "card_scroll", None) is not None:
            self._apply_scroll_minimum()

    def _sync_ssh_state(self, *_):
        enabled = self.ssh_check.isChecked() and not self._busy
        for widget in (self.ssh_host_edit, self.ssh_port_edit, self.ssh_user_edit,
                       self.ssh_pwd_edit):
            widget.setEnabled(enabled)

    def _auto_fill_output_names(self):
        in_key = self.input_key()
        base = ""
        if in_key == "SQL":
            if len(self.selected_tables) > 1:
                self._reset_auto_name(self.out_sheet_edit, "_auto_sheet")
                self._reset_auto_name(self.out_table_edit, "_auto_table")
                return
            base = (self.selected_table or "").split(".")[-1]
        elif in_key == "Excel":
            base = self.in_sheet_combo.currentText().strip() or Path(self.in_file_edit.text()).stem
        elif in_key == "JSON":
            base = self.in_key_combo.currentText().strip() or Path(self.in_file_edit.text()).stem
        elif in_key == "CSV":
            section = self.in_section_combo.currentText().strip()
            base = section if section and not section.startswith("(") \
                else Path(self.in_file_edit.text()).stem
        if not base:
            return
        for edit, attr in ((self.out_sheet_edit, "_auto_sheet"),
                           (self.out_table_edit, "_auto_table")):
            if not edit.text() or edit.text() == getattr(self, attr):
                edit.setText(base)
                setattr(self, attr, base)

    def _reset_auto_name(self, edit, attr):
        if not edit.text() or edit.text() == getattr(self, attr):
            edit.setText("")
            setattr(self, attr, "")

    # ------------------------------------------------------------- 连接管理
    def get_cfg(self):
        """界面参数 → 连接配置;MongoDB 走 MongoConfig,SSH 隧道生效时自动改走隧道端口。"""
        values = {
            "host": self.host_edit.text().strip() or "127.0.0.1",
            "port": to_int(self.port_edit.currentText(), 0),
            "user": self.user_edit.text().strip(),
            "password": self.pwd_edit.text(),
            "database": self.database_edit.text().strip(),
            "timeout": to_int(self.timeout_edit.text(), 10) or 10,
        }
        if self.adapter.NAME == "MongoDB":
            values.update({"uri": self.mongo_uri_edit.text().strip(),
                           "auth_source": self.mongo_auth_edit.text().strip()})
        cfg = config_for(self.adapter.NAME, values)
        if self.tunnel is not None and self.tunnel.active:
            return type(cfg)(**{**cfg.as_dict(), "host": SSHTunnel.BIND_HOST,
                                "port": self.tunnel.local_port})
        return cfg

    def remote_endpoint(self):
        """SSH 隧道要转发到的【远程服务器上】数据库地址。"""
        return (self.host_edit.text().strip() or "127.0.0.1",
                to_int(self.port_edit.currentText(), self.adapter.DEFAULT_PORT))

    def ssh_config(self) -> SSHConfig:
        return SSHConfig(
            enabled=self.ssh_check.isChecked(),
            host=self.ssh_host_edit.text().strip(),
            port=to_int(self.ssh_port_edit.text(), SSH_DEFAULT_PORT) or SSH_DEFAULT_PORT,
            user=self.ssh_user_edit.text().strip(),
            password=self.ssh_pwd_edit.text(),
        )

    def connect_db(self, cfg=None):
        cfg = (cfg or self.get_cfg()).validated(self.adapter)
        self._disconnect_conn()
        self.conn = self.adapter.connect(cfg)
        return self.adapter.server_version(self.conn)

    def _load_config(self):
        path = config_file_path()
        if not path.exists():
            return
        try:
            settings = Settings.load(path)
        except Exception as exc:
            self.log(f"读取记忆配置失败(使用空白配置): {exc}", "warn")
            return
        if settings.password:
            self.pwd_edit.setText(settings.password)
        if settings.database:
            self.database_edit.setText(settings.database)
        if settings.ssh.get("host"):
            ssh = SSHConfig.from_mapping(settings.ssh)
            self.ssh_check.setChecked(bool(ssh.enabled))
            self.ssh_host_edit.setText(ssh.host)
            self.ssh_port_edit.setText(str(ssh.port))
            self.ssh_user_edit.setText(ssh.user)
            self.ssh_pwd_edit.setText(ssh.password)
        self.log(f"已载入记忆的连接配置(密码/数据库名/SSH): {path}", "info")
        self._update_ui_state()

    def _save_config(self):
        try:
            settings = Settings.load()
        except Exception:
            settings = Settings()
        settings.password = self.pwd_edit.text()
        settings.database = self.database_edit.text()
        settings.ssh = self.ssh_config().as_dict()
        try:
            settings.save()
        except Exception:
            pass

    def _close_conn(self, conn):
        """按适配器类型关闭连接:SQL 驱动直接 close,MongoDB 要经适配器关 client。"""
        if conn is None:
            return
        with contextlib.suppress(Exception):
            if self.adapter.NAME == "MongoDB":
                self.adapter.close(conn)
            else:
                conn.close()

    def _disconnect_conn(self):
        self._close_conn(self.conn)
        self.conn = None

    def _stop_tunnel(self):
        if self.tunnel is not None:
            with contextlib.suppress(Exception):
                self.tunnel.stop()
            self.tunnel = None

    def disconnect(self):
        self._disconnect_conn()
        self._refresh_status_chip()
        self._update_ui_state()

    def _ping(self):
        try:
            if self.adapter.NAME == "MySQL":
                self.conn.ping(reconnect=False)
                return True
            if self.adapter.NAME == "MongoDB":
                self.conn.command("ping")          # Database 没有 .closed,要真的探一次
                return True
            return not self.conn.closed
        except Exception:
            return False

    def ensure_conn(self, cfg=None):
        with contextlib.suppress(Exception):
            if self.conn is not None and self._ping():
                return self.conn
        self._disconnect_conn()
        version = self.connect_db(cfg)
        self.log(f"已自动连接 {self.adapter.NAME} {version}", "ok")
        self.set_status(f"已连接 {self.adapter.NAME} {version}", connected=True)
        return self.conn

    # --------------------------------------------------------- SSH 隧道事件
    def on_ssh_connect(self):
        cfg = self.ssh_config()
        cfg.enabled = True
        if not cfg.host:
            QMessageBox.warning(self, "提示", "请填写 SSH 主机地址")
            return
        if not cfg.password:
            QMessageBox.warning(self, "提示", "请填写 SSH 密码(本工具统一使用 用户名 + 密码 认证)")
            return
        remote_host, remote_port = self.remote_endpoint()

        def work():
            self._stop_tunnel()
            self.tunnel = SSHTunnel(cfg, remote_host=remote_host,
                                    remote_port=remote_port).start()
            self.log(f"SSH 隧道已建立: {self.tunnel.describe()}", "ok")
        self.run_task(work, describe=lambda exc: describe_ssh_error(exc, cfg))

    def on_ssh_disconnect(self):
        self._stop_tunnel()
        self._disconnect_conn()          # 隧道没了,经隧道建立的连接也不能再用
        self.log("SSH 隧道已断开", "warn")
        self._refresh_status_chip()
        self._update_ui_state()

    # ------------------------------------------------------------- UI 事件
    def on_db_type_changed(self, *_):
        name = self.db_type_combo.currentText()
        self.adapter = ADAPTERS[name]
        self.disconnect()
        self._sync_db_type_rows()
        self._sync_port_presets()
        self.log(f"已切换到 {name}(端口留空时自动使用默认端口 "
                 f"{self.adapter.DEFAULT_PORT})", "info")
        self._update_ui_state()

    def _sync_port_presets(self):
        """端口下拉:当前数据库类型的默认端口排第一,其余常用端口跟在后面。"""
        current = self.port_edit.currentText()
        default = str(self.adapter.DEFAULT_PORT)
        ports = [default] + [str(port) for port in DB_PORT_PRESETS if str(port) != default]
        self.port_edit.blockSignals(True)
        self.port_edit.clear()
        self.port_edit.addItems(ports)
        self.port_edit.setCurrentText(current)
        self.port_edit.blockSignals(False)

    def _sync_db_type_rows(self):
        """MongoDB 专用输入行(连接 URI / 认证库)仅在选中 MongoDB 时显示。"""
        mongo = self.adapter.NAME == "MongoDB"
        for widget in (self.mongo_uri_label, self.mongo_uri_edit,
                       self.mongo_auth_label, self.mongo_auth_edit):
            widget.setVisible(mongo)

    def on_test_connection(self):
        cfg = self.get_cfg()

        def work():
            # 用临时连接试连,不动用户已经建立/已加载表结构的连接
            probe = self.adapter.connect(cfg.validated(self.adapter))
            try:
                version = self.adapter.server_version(probe)
            finally:
                self._close_conn(probe)
            self.log(f"连接成功: {self.adapter.NAME} {version}", "ok")
            self.bridge.info_s.emit(f"连接成功!\n{self.adapter.NAME} {version}")
        self.run_task(work)

    def _load_table_columns(self, conn, tables):
        table_columns = {}
        for table in tables:
            try:
                table_columns[table] = self.adapter.list_columns(conn, table)
            except Exception as exc:
                self.log(f"读取表【{table}】字段失败: {exc}", "warn")
                table_columns[table] = []
        return table_columns

    def _open_connection(self, cfg, announce=True):
        version = self.connect_db(cfg)
        tables = self.adapter.list_tables(self.conn)
        if announce:
            self.log(f"连接成功: {self.adapter.NAME} {version},"
                     f"共 {len(tables)} 个{self._entity_label()}", "ok")
            self.log("正在加载字段结构 ...", "info")
            self._log_empty_hint(tables, cfg.database)
        self.bridge.tree_s.emit(self._load_table_columns(self.conn, tables))
        self.set_status(f"已连接 {self.adapter.NAME} {version}", connected=True)

    def _entity_label(self):
        """MongoDB 里叫集合,关系库里叫数据表(提示语用对词,避免用户找错对象)。"""
        return "集合" if self.adapter.NAME == "MongoDB" else "数据表"

    def _log_empty_hint(self, tables, database):
        """一个集合/表都没找到时说明原因,而不是只报 0:多数情况是空库或库名/认证库不对。"""
        if tables:
            return
        name = str(database or "").strip() or "(未填写)"
        if self.adapter.NAME == "MongoDB":
            self.log(f"数据库【{name}】里没有任何集合:MongoDB 的空库不占空间、也不会出现在 "
                     'show dbs 里,请先建集合或插入文档(例: db.students.insertMany([{name:"a"}])),'
                     "再点【刷新表】;若库里确实有数据,请核对左侧【数据库】名与【认证库 authSource】。",
                     "warn")
        else:
            self.log(f"数据库【{name}】里没有任何可读数据表:请核对库名与账号权限后重试。", "warn")

    def on_connect(self):
        cfg = self.get_cfg()

        def work():
            try:
                self._open_connection(cfg)
            except Exception:
                self._disconnect_conn()
                raise
        self.run_task(work)

    def on_refresh_tables(self):
        cfg = self.get_cfg()

        def work():
            try:
                conn = self.ensure_conn(cfg)
                tables = self.adapter.list_tables(conn)
                self.log(f"共发现 {len(tables)} 个{self._entity_label()},正在加载字段结构 ...", "info")
                self._log_empty_hint(tables, cfg.database)
                self.bridge.tree_s.emit(self._load_table_columns(conn, tables))
            except Exception:
                self._disconnect_conn()
                raise
        self.run_task(work)

    def on_tree_confirmed(self, selection):
        self.selected_columns = {table: list(cols) for table, cols in selection.items() if cols}
        self.selected_tables = list(self.selected_columns)
        self.selected_table = self.selected_tables[0] if self.selected_tables else None
        self._auto_fill_output_names()
        if self.selected_tables:
            self.log(f"已选择 {len(self.selected_tables)} 张表、"
                     f"{sum(map(len, self.selected_columns.values()))} 个字段", "ok")
        else:
            self.log("未勾选任何表/字段", "warn")
        self._update_ui_state()

    def on_pick_input_file(self):
        key = self.input_key()
        filters = {
            "Excel": "Excel 工作簿 (*.xlsx);;所有文件 (*.*)",
            "JSON": "JSON 文件 (*.json);;所有文件 (*.*)",
            "CSV": "CSV 文件 (*.csv);;所有文件 (*.*)",
        }[key]
        path, _ = QFileDialog.getOpenFileName(
            self, f"选择输入文件 ({INPUT_LABELS[key]})", "", filters)
        if not path:
            return
        self.in_file_edit.setText(path)
        self._auto_fill_output_names()
        self._update_ui_state()

        if key == "Excel":
            def work():
                sheet_headers = workbook_sheet_headers(path)
                self.log(f"工作簿包含工作表: {', '.join(sheet_headers)}", "info")
                self.bridge.sheets_s.emit(list(sheet_headers))
                self.bridge.excel_tree_s.emit(sheet_headers)
            self.run_task(work)
        elif key == "JSON":
            def work():
                keys = json_top_keys(path)
                self.log(f"JSON 包含一级键: {', '.join(keys)}", "info")
                self.bridge.keys_s.emit(keys)
            self.run_task(work)
        elif key == "CSV":
            delimiter = DELIM_LABELS[self.in_delim_combo.currentText()]
            encoding = ENC_LABELS[self.in_enc_combo.currentText()]

            def work():
                if sections := scan_csv_sections(path, delimiter, encoding):
                    self.log(f"分节 CSV,共 {len(sections)} 个分节(Sheet): "
                             f"{', '.join(sections)}", "section")
                    self.bridge.csv_sections_s.emit(sections)
                else:
                    self.log(f"普通 CSV,表头列: "
                             f"{read_csv_header(path, delimiter, encoding)}", "info")
                    self.bridge.csv_sections_s.emit([])
            self.run_task(work)

    def on_in_sheet_selected(self, *_):
        self._auto_fill_output_names()
        self._update_ui_state()

    def on_in_key_selected(self, *_):
        self._auto_fill_output_names()
        self._update_ui_state()

    def on_in_section_selected(self, *_):
        self._auto_fill_output_names()
        self._update_ui_state()

    def on_excel_tree_confirmed(self, selection):
        """Excel→SQL 层级勾选确认:selection = {Sheet名: [勾选字段...]}。"""
        self.excel_selection = {sheet: list(cols) for sheet, cols in selection.items() if cols}
        self.excel_sheets = list(self.excel_selection)
        if self.excel_sheets:
            first = self.excel_sheets[0]
            if not self.out_table_edit.text() or self.out_table_edit.text() == self._auto_table:
                self.out_table_edit.setText(first)
                self._auto_table = first
            self.log(f"Excel 已选择 {len(self.excel_sheets)} 个 Sheet、"
                     f"{sum(map(len, self.excel_selection.values()))} 个字段", "ok")
        else:
            self.log("Excel 未勾选任何 Sheet/字段", "warn")
        self._update_ui_state()

    def _sync_excel_tree(self):
        """仅当 输入=Excel 且 输出=SQL 时显示层级勾选控件。"""
        visible = self.input_key() == "Excel" and self.output_key() == "SQL"
        self.excel_tree.setVisible(visible)
        self.excel_tree_label.setVisible(visible)

    # ------------------------------------------------------------- 冲突对话框
    def ask_conflict_mode(self, existing_paths):
        """同名文件冲突纠错对话框,返回 'replace' / 'merge' / None(取消)。"""
        return self._ask_choice(
            "检测到同名文件",
            "以下文件已存在:\n" + "\n".join(f"  • {path}" for path in existing_paths)
            + "\n\n请选择处理方式:",
            [("覆盖整个文件", "replace", "warn"),
             ("合并写入(覆盖同名Sheet/JSON键;CSV追加/覆盖分节)", "merge", "blue"),
             ("取消", None, None)],
            default=None)

    def _ask_sheet_mode(self, sheet):
        """检测到数据库同名表时,询问 覆盖写入 / 追加写入 / 跳过该表。"""
        return self._ask_choice(
            "检测到同名数据表",
            f"数据库已存在与 Sheet 同名的表【{sheet}】。\n请选择写入方式:",
            [("覆盖写入(删除原表并重建导入)", "replace", "warn"),
             ("追加写入(保留原表,追加数据行)", "append", "blue"),
             ("跳过该表", "skip", None)],
            default="skip")

    def _ask_choice(self, title, message, options, default=None):
        """通用三选一对话框:options = [(按钮文字, 返回值, 按钮样式), ...]。"""
        dialog = QDialog(self)
        dialog.setWindowTitle(title)
        lay = QVBoxLayout(dialog)
        label = QLabel(message)
        label.setWordWrap(True)
        lay.addWidget(label)
        buttons = QHBoxLayout()
        result = {"value": default}

        def pick(value):
            result["value"] = value
            dialog.accept()

        for text, value, style in options:
            button = QPushButton(text)
            if style:
                button.setObjectName(style)
            button.clicked.connect(lambda _=False, v=value: pick(v))
            buttons.addWidget(button)
        lay.addLayout(buttons)
        dialog.exec()
        dialog.deleteLater()
        return result["value"]

    def _ask_output_path(self, out_key, default_base):
        ext = {"Excel": ".xlsx", "JSON": ".json", "CSV": ".csv"}[out_key]
        label = OUTPUT_LABELS[out_key]
        initial = str(Path.home() / "Documents" / f"{safe_filename(default_base)}{ext}")
        path, _ = QFileDialog.getSaveFileName(
            self, f"保存输出文件 ({label})", initial, f"{label} (*{ext})")
        if not path:
            return None, None
        if os.path.exists(path):
            mode = self.ask_conflict_mode([path])
            if mode is None:
                self.log("已取消转换(存在同名文件)", "warn")
                return None, None
        else:
            mode = "replace"
        return path, mode

    def _default_output_base(self, in_key):
        if in_key == "SQL":
            return self.database_edit.text().strip() or (self.selected_table or "export")
        return Path(self.in_file_edit.text()).stem or "export"

    # ------------------------------------------------------------- 字段映射
    def _current_header(self) -> list:
        """当前输入源的字段列表(用于字段映射预览)。"""
        in_key = self.input_key()
        if in_key == "SQL":
            return list(self.selected_columns.get(self.selected_table or "", []))
        if in_key == "Excel" and self.output_key() == "SQL":
            # Excel→SQL 走层级勾选:预览必须用勾选后的那个 Sheet 与字段
            first = self.excel_sheets[0] if self.excel_sheets else ""
            return list(self.excel_selection.get(first, []))
        path = self.in_file_edit.text().strip()
        if not path:
            return []
        try:
            if in_key == "Excel":
                spec = SourceSpec(kind="Excel", path=path,
                                  table=self.in_sheet_combo.currentText())
            elif in_key == "JSON":
                spec = SourceSpec(kind="JSON", path=path,
                                  table=self.in_key_combo.currentText())
            else:
                section = self.in_section_combo.currentText().strip()
                spec = SourceSpec(kind="CSV", path=path,
                                  table="" if section.startswith("(") else section,
                                  delimiter=DELIM_LABELS[self.in_delim_combo.currentText()],
                                  encoding=ENC_LABELS[self.in_enc_combo.currentText()])
            header, _, closer = open_reader(spec)
            closer()
            return list(header)
        except Exception as exc:                 # noqa: BLE001 - 预览失败不应阻断操作
            self.log(f"读取字段列表失败: {exc}", "warn")
            return []

    def on_edit_mapping(self):
        """字段映射预览与手动调整(Sheet 名 ↔ 表名、首行表头 ↔ 数据库字段名)。"""
        fields = self._current_header()
        if not fields:
            QMessageBox.information(self, "提示", "请先选择输入数据源(数据库表 / 文件)")
            return
        dialog = QDialog(self)
        dialog.setWindowTitle("字段映射(源字段 → 目标字段)")
        dialog.resize(520, 460)
        lay = QVBoxLayout(dialog)
        hint = QLabel("左侧为源字段(首行表头),右侧可改为目标字段名;留空表示保持同名。")
        hint.setObjectName("muted")
        hint.setWordWrap(True)
        lay.addWidget(hint)
        table = QTableWidget(len(fields), 2, dialog)
        table.setHorizontalHeaderLabels(["源字段", "目标字段"])
        for row, name in enumerate(fields):
            item = QTableWidgetItem(name)
            item.setFlags(item.flags() & ~Qt.ItemIsEditable)
            table.setItem(row, 0, item)
            table.setItem(row, 1, QTableWidgetItem(self.field_mapping.target_of(name)))
        table.resizeColumnsToContents()
        lay.addWidget(table, 1)

        def reset_names():
            for row, name in enumerate(fields):
                table.setItem(row, 1, QTableWidgetItem(name))

        buttons = QHBoxLayout()
        reset_btn = QPushButton("恢复同名")
        reset_btn.clicked.connect(reset_names)
        ok_btn = QPushButton("确定")
        ok_btn.setObjectName("primary")
        ok_btn.clicked.connect(dialog.accept)
        buttons.addWidget(reset_btn)
        buttons.addStretch(1)
        buttons.addWidget(ok_btn)
        lay.addLayout(buttons)

        accepted = dialog.exec() == QDialog.Accepted
        dialog.deleteLater()
        if not accepted:
            return
        pairs = {}
        for row, name in enumerate(fields):
            item = table.item(row, 1)
            target = (item.text().strip() if item else "") or name
            if target != name:
                pairs[name] = target
        self.field_mapping = FieldMapping.from_pairs(pairs)
        self.map_label.setText(f"字段映射: {self.field_mapping.describe()}")
        self.log(f"字段映射已更新: {self.field_mapping.describe()}", "info")

    # ------------------------------------------------------------- 主转换流程
    def _collect_sources(self, in_key, out_key):
        """按层级勾选结果组装输入源;校验不通过时抛 ValueError(中文提示)。"""
        if in_key == "SQL":
            selection = {table: cols for table, cols in self.selected_columns.items() if cols}
            if not selection:
                raise ValueError("请连接数据库,并在“选择数据表与字段”层级下拉框中勾选表和字段")
            kind = "Mongo" if self.adapter.NAME == "MongoDB" else "SQL"
            return [SourceSpec(kind=kind, table=table, columns=list(cols))
                    for table, cols in selection.items()]

        path = self.in_file_edit.text().strip()
        if not (path and os.path.isfile(path)):
            raise ValueError(f"请先选择有效的 {in_key} 输入文件")
        if in_key == "Excel":
            if out_key == "SQL":
                if not self.excel_selection:
                    raise ValueError("请在“选择工作表与字段”层级下拉框中勾选要导入的 Sheet 与字段")
                return [SourceSpec(kind="Excel", path=path, table=sheet, columns=list(cols))
                        for sheet, cols in self.excel_selection.items()]
            return [SourceSpec(kind="Excel", path=path, table=self.in_sheet_combo.currentText())]
        if in_key == "JSON":
            key = self.in_key_combo.currentText()
            if not key:
                raise ValueError("请选择 JSON 一级键(表名)")
            return [SourceSpec(kind="JSON", path=path, table=key)]
        section = self.in_section_combo.currentText().strip()
        return [SourceSpec(kind="CSV", path=path,
                           table="" if section.startswith("(") else section,
                           delimiter=DELIM_LABELS[self.in_delim_combo.currentText()],
                           encoding=ENC_LABELS[self.in_enc_combo.currentText()],
                           parse=(out_key == "SQL"))]

    def _prepare_target(self, in_key, out_key, sources):
        """在主线程完成全部对话框交互,返回 TargetSpec(None = 用户取消)。"""
        delimiter = DELIM_LABELS[self.out_delim_combo.currentText()]
        encoding = ENC_LABELS[self.out_enc_combo.currentText()]

        if out_key == "SQL":
            if in_key == "Excel":
                # Excel→SQL:目标表 = Sheet 名;同名表在主线程逐个询问 覆盖/追加/跳过
                if not self._apply_sheet_decisions(sources):
                    return None
                return TargetSpec(kind="SQL", mode="create_if_missing",
                                  batch=self.batch_spin.value())
            multi = len(sources) > 1
            table_out = self.out_table_edit.text().strip()
            if multi and table_out:
                QMessageBox.warning(self, "提示", "多表拷贝时目标表名请留空(将按源表同名建表)")
                return None
            if not multi and not table_out:
                QMessageBox.warning(self, "提示", "请填写目标表名")
                return None
            mode = MODE_KEYS.get(self.out_mode_combo.currentText(), "append")
            if mode == "append" and not multi and table_out in {src.table for src in sources}:
                # 目标表与输入表同名 + 追加 = INSERT ... SELECT 自身,会把整表数据复制一遍
                if QMessageBox.question(
                        self, "危险操作",
                        f"目标表【{table_out}】与输入表同名:追加写入会把整表数据再复制一遍。\n"
                        "确定继续吗?(建议改用“创建新表”或先改目标表名)",
                        QMessageBox.Yes | QMessageBox.No) != QMessageBox.Yes:
                    return None
            if mode == "replace":
                message = ("多表同名拷贝:将先删除并重建每张勾选的源表再重新写入,确定继续吗?"
                           if multi else
                           f"将先删除表【{table_out}】(若存在)再重建并写入,确定继续吗?")
                if QMessageBox.question(self, "危险操作", message,
                                        QMessageBox.Yes | QMessageBox.No) != QMessageBox.Yes:
                    return None
            return TargetSpec(kind="SQL", table="" if multi else table_out, mode=mode,
                              batch=self.batch_spin.value())

        if out_key == "CSV" and in_key == "SQL":
            # SQL→CSV:目录模式(目录名 = 数据库名,每个表一个 {表名}.csv)
            default_dir = safe_filename(self._default_output_base(in_key))
            parent_dir = QFileDialog.getExistingDirectory(
                self, f"选择 CSV 输出的父目录(将创建目录 {default_dir},"
                      f"其中每个表一个 .csv 文件)")
            if not parent_dir:
                return None
            target_dir = os.path.join(parent_dir, default_dir)
            if existing := [p for p in csv_output_paths(target_dir, sources) if os.path.exists(p)]:
                mode = self.ask_conflict_mode(existing)
                if mode is None:
                    return None
            else:
                mode = "replace"
            return TargetSpec(kind="CSV", path=target_dir, directory=True, mode=mode,
                              delimiter=delimiter, encoding=encoding)

        default_base = self._default_output_base(in_key)
        path, mode = self._ask_output_path(out_key, default_base)
        if path is None:
            return None
        name = self.out_sheet_edit.text().strip() or default_base
        return TargetSpec(kind=out_key, path=path, mode=mode, sheet=safe_sheet_title(name),
                          key=name, delimiter=delimiter, encoding=encoding)

    def _apply_sheet_decisions(self, sources):
        """Excel→SQL:同名表逐个询问写入方式(返回 False 表示用户取消)。"""
        try:
            conn = self.ensure_conn()
        except Exception as exc:
            QMessageBox.critical(self, "错误",
                                 describe_db_error(exc, self.get_cfg(), self.adapter))
            return False
        for src in sources:
            try:
                quoted = self.adapter.quote_ident(src.table)
            except ValueError as exc:
                QMessageBox.warning(self, "提示",
                                    f"工作表名【{src.table}】不能直接作为数据库表名。\n{exc}\n"
                                    f"请重命名工作表,或改用其他输出格式(JSON / CSV / Excel)。")
                return False
            if self.adapter.table_exists(conn, quoted):
                choice = self._ask_sheet_mode(src.table)
                src.mode = choice or "skip"
                self.log(f"同名表【{src.table}】已选择: {src.mode}",
                         "warn" if src.mode == "skip" else "info")
            else:
                src.mode = "create_if_missing"
        return True

    @staticmethod
    def _needs_db(sources, target):
        return target.kind in DB_KINDS or any(src.kind in DB_KINDS for src in sources)

    def on_convert(self):
        if self._busy:
            QMessageBox.information(self, "提示", "已有任务正在执行,请等待完成后再试。")
            return
        in_key, out_key = self.input_key(), self.output_key()
        adapter = self.adapter
        try:
            sources = self._collect_sources(in_key, out_key)
        except ValueError as exc:
            QMessageBox.warning(self, "提示", str(exc))
            return
        target = self._prepare_target(in_key, out_key, sources)
        if target is None:
            return
        target.fields = self.field_mapping            # 应用用户设置的字段映射
        cfg = self.get_cfg()
        need_db = self._needs_db(sources, target)

        def work():
            conn = self.ensure_conn(cfg) if need_db else None
            try:
                result = convert(sources, target, conn=conn, adapter=adapter,
                                 cfg=cfg, log=self.log)
            except FidelityError as exc:
                # 数值保真校验未通过:输出差异报告(数据已写出,由用户决定后续处理)
                self.log(str(exc), "err")
                self.bridge.error_s.emit(
                    "数值保真校验未通过(数据已写出,请核对):\n"
                    + "\n".join(exc.report.summary_lines()))
                return
            self.log(f"✅ 转换完成: {result.summary()}", "ok")
            if result.fidelity.checked_values:
                for line in result.fidelity.summary_lines():
                    self.log(line, "ok" if result.fidelity.ok else "warn")
            self.bridge.info_s.emit(f"转换完成!\n{result.summary()}")
        self.run_task(work)

    # ------------------------------------------------------------- 退出
    def closeEvent(self, event):
        if self._busy and QMessageBox.question(
                self, "任务进行中",
                "正在执行转换,现在退出会中断写入,可能留下不完整的输出文件。\n确定退出吗?",
                QMessageBox.Yes | QMessageBox.No, QMessageBox.No) != QMessageBox.Yes:
            event.ignore()
            return
        self._save_config()
        self._disconnect_conn()
        self._stop_tunnel()
        event.accept()


def run_self_test():
    """--selftest: 无界面自检(打包后验证驱动/引擎是否齐全),结果写入 selftest.log。"""
    passed = core_self_test()
    log_path = Path.cwd() / "selftest.log"
    line = "[OK]   PySide6 图形界面依赖可用\n"
    with contextlib.suppress(Exception):
        log_path.write_text(log_path.read_text(encoding="utf-8") + line, encoding="utf-8")
    return passed


def main():
    if "--selftest" in sys.argv:
        sys.exit(0 if run_self_test() else 1)
    if "--version" in sys.argv:
        print(VERSION)
        sys.exit(0)
    app = QApplication(sys.argv)
    app.setStyleSheet(QSS)
    window = App()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
