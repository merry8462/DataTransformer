# -*- coding: utf-8 -*-
"""图形界面关键逻辑测试(未安装 PySide6 时自动跳过)。

覆盖:主窗口构造、MongoDB / SQL 配置切换、字段映射应用、层级勾选控件、
输出基准名与写入模式映射、自定义查询语句(后台校验预览 → 单源 → 与层级勾选互斥)。
"""

import os
import sys
import time

import pytest

pytest.importorskip("PySide6", reason="未安装 PySide6,跳过图形界面测试")

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")          # 无显示器环境可运行
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

from PySide6.QtWidgets import QApplication, QLabel, QSizePolicy   # noqa: E402

import data_transformer as gui                               # noqa: E402
import src.dt_core as core                                   # noqa: E402


@pytest.fixture(scope="module")
def window():
    app = QApplication.instance() or QApplication([])
    app.setStyleSheet(gui.QSS)
    window = gui.App()
    yield window
    window.close()


def test_window_builds_with_all_adapters(window):
    assert window.input_key() == "SQL" and window.output_key() == "Excel"
    assert set(core.ADAPTERS) == {"MySQL", "PostgreSQL", "MongoDB"}


def test_mode_switching_covers_every_combination(window):
    for text in core.INPUT_LABELS.values():
        window.input_combo.setCurrentText(text)
        for out_text in core.OUTPUT_LABELS.values():
            window.output_combo.setCurrentText(out_text)
    assert window.input_key() in core.INPUT_LABELS


def test_connection_config_follows_adapter(window):
    window.ssh_check.setChecked(False)      # 记忆配置可能带着 SSH 勾选:取配置前先关掉
    window.db_type_combo.setCurrentText("MongoDB")
    assert isinstance(window.get_cfg(), core.MongoConfig)
    window.db_type_combo.setCurrentText("PostgreSQL")
    assert isinstance(window.get_cfg(), core.DBConfig)
    window.db_type_combo.setCurrentText("MySQL")
    assert isinstance(window.get_cfg(), core.DBConfig)


def test_db_card_follows_the_side_that_needs_it(window):
    """数据库连接块跟着需要它的那一侧:输入=数据库在左,只有输出=数据库时在右。"""
    window.input_combo.setCurrentText(core.INPUT_LABELS["SQL"])
    window.output_combo.setCurrentText(core.OUTPUT_LABELS["SQL"])
    assert window.db_card.parentWidget() is window.input_card
    assert "共用" in window.db_card.title()

    window.input_combo.setCurrentText(core.INPUT_LABELS["Excel"])
    assert window.db_card.parentWidget() is window.output_card
    assert "输出" in window.db_card.title()

    window.output_combo.setCurrentText(core.OUTPUT_LABELS["Excel"])
    assert window.db_card.isHidden()


def test_ssh_buttons_sit_on_the_field_row(window):
    """SSH 卡片压成两行:字段行右侧放连接/断开按钮,不再单独占一行空带。"""
    grid = window.ssh_card.layout()
    conn_pos = grid.getItemPosition(grid.indexOf(window.ssh_conn_btn))
    disc_pos = grid.getItemPosition(grid.indexOf(window.ssh_disc_btn))
    host_pos = grid.getItemPosition(grid.indexOf(window.ssh_host_edit))
    assert conn_pos[0] == disc_pos[0] == host_pos[0] == 0     # 与输入框同一行
    assert conn_pos[1] > host_pos[1] and disc_pos[1] > conn_pos[1]
    assert grid.getItemPosition(grid.indexOf(window.ssh_state))[0] == 1


def test_port_is_editable_combo_with_presets(window):
    """端口改成可选预设 + 可手填的下拉框;留空时仍用该类型的默认端口。"""
    items = [window.port_edit.itemText(index) for index in range(window.port_edit.count())]
    assert {"3306", "5432", "27017"} <= set(items)
    assert window.port_edit.isEditable()                    # 能自己填
    assert "3306" in window.port_edit.toolTip()             # 提示里说明可选常用端口
    assert window.port_edit.lineEdit().placeholderText()    # 有"留空=默认端口"提示

    window.ssh_check.setChecked(False)      # 记忆配置可能带着 SSH 勾选:取配置前先关掉
    window.db_type_combo.setCurrentText("MySQL")
    window.host_edit.setText("127.0.0.1")
    window.user_edit.setText("u")
    window.database_edit.setText("d")
    assert window.port_edit.itemText(0) == "3306"           # 当前类型的默认端口排最前
    window.db_type_combo.setCurrentText("MongoDB")
    assert window.port_edit.itemText(0) == "27017"
    window.db_type_combo.setCurrentText("MySQL")

    window.port_edit.setCurrentText("15432")            # 手动填写
    assert window.get_cfg().port == 15432
    window.port_edit.setCurrentText("")                 # 留空 → 连接时用默认端口
    cfg = window.get_cfg()
    assert cfg.port == 0
    assert cfg.validated(window.adapter).port == core.ADAPTERS["MySQL"].DEFAULT_PORT


def test_ssh_card_uses_password_only_and_sits_on_top(window):
    """SSH 卡片:去掉私钥文件,统一 用户名+密码 认证;并位于输入/输出配置之上。"""
    assert not hasattr(window, "ssh_key_edit") and not hasattr(window, "ssh_key_btn")
    assert window.card_scroll.widget().layout().itemAt(0).widget() is window.ssh_card

    window.input_combo.setCurrentText(core.INPUT_LABELS["SQL"])     # 涉及数据库才有 SSH
    window.output_combo.setCurrentText(core.OUTPUT_LABELS["Excel"])
    window.ssh_host_edit.setText("1.2.3.4")
    window.ssh_user_edit.setText("cml")
    window.ssh_pwd_edit.clear()
    window.ssh_check.setChecked(True)
    window._update_ui_state()
    assert not window.ssh_conn_btn.isEnabled()          # 没填密码 → 不能连接

    window.ssh_pwd_edit.setText("secret")
    window._update_ui_state()
    assert window.ssh_conn_btn.isEnabled()
    assert window.ssh_config().key_file == ""           # 引擎字段保留,GUI 不再使用


def test_reset_all_restores_defaults(window, monkeypatch):
    """初始化按钮:清空所有可填写参数、恢复默认值并断开连接。"""
    monkeypatch.setattr(gui.QMessageBox, "question",
                        lambda *args, **kwargs: gui.QMessageBox.Yes)
    window.host_edit.setText("10.0.0.9")
    window.user_edit.setText("someone")
    window.database_edit.setText("db1")
    window.in_file_edit.setText("C:/tmp/in.xlsx")
    window.out_table_edit.setText("t1")
    window.ssh_host_edit.setText("1.2.3.4")
    window.ssh_check.setChecked(True)
    window.batch_spin.setValue(999)

    window.on_reset_all()

    assert window.host_edit.text() == "127.0.0.1"
    assert window.timeout_edit.text() == "10"
    assert not any((window.user_edit.text(), window.database_edit.text(),
                    window.pwd_edit.text(), window.in_file_edit.text(),
                    window.out_table_edit.text(), window.out_sheet_edit.text()))
    assert not window.ssh_check.isChecked() and window.ssh_host_edit.text() == ""
    assert window.ssh_port_edit.text() == str(core.SSH_DEFAULT_PORT)
    assert window.batch_spin.value() == core.DEFAULT_BATCH
    assert window.input_key() == "SQL" and window.output_key() == "Excel"
    assert not window.selected_columns and not window.excel_selection


def test_empty_database_is_explained_in_the_ui(window):
    """连上但库里没有集合/表时,界面要说明原因,而不是只显示“共 0 张表”。"""
    window.db_type_combo.setCurrentText("MongoDB")
    assert window._entity_label() == "集合"
    window._log_empty_hint([], "mgdb2026-10-03")
    assert "没有任何集合" in window.log_text.toPlainText()

    window._on_tree_data({})
    assert "没有可用" in window.tree_select.button.text()

    window.db_type_combo.setCurrentText("MySQL")
    assert window._entity_label() == "数据表"


def test_mongodb_input_is_collected_as_mongo_source(window):
    """选 MongoDB 时输入源必须是 Mongo 类型:否则引擎按 SQL 去找游标,直接 AttributeError。"""
    window.db_type_combo.setCurrentText("MongoDB")
    window.selected_columns = {"people": ["name", "age"]}
    sources = window._collect_sources("SQL", "Excel")
    assert sources and all(src.kind == "Mongo" for src in sources)
    assert window._needs_db(sources, core.TargetSpec(kind="Excel", path="out.xlsx"))

    window.db_type_combo.setCurrentText("MySQL")
    sources = window._collect_sources("SQL", "Excel")
    assert sources and all(src.kind == "SQL" for src in sources)


def test_tunnel_rewrites_connection_target(window):
    class FakeTunnel:
        local_port = 55555
        active = True
        remote_host = "127.0.0.1"
        remote_port = 3306

        def describe(self):
            return "127.0.0.1:55555"

        def retarget(self, host, port):
            self.remote_host, self.remote_port = host, port
            return self

        def stop(self):
            pass

    window.db_type_combo.setCurrentText("PostgreSQL")
    window.tunnel = FakeTunnel()
    tunneled = window.get_cfg()
    assert tunneled.host == core.SSHTunnel.BIND_HOST and tunneled.port == 55555
    window.tunnel = None
    assert window.get_cfg().port != 55555


def test_field_mapping_is_applied_to_target(window):
    window.field_mapping = core.FieldMapping.from_pairs({"姓名": "Name"})
    target = core.TargetSpec(kind="JSON")
    target.fields = window.field_mapping
    assert core.apply_field_mapping(["姓名", "金额"], target) == ["Name", "金额"]


def test_hierarchical_selection_widget(window):
    window.tree_select.set_tables({"t1": ["A", "B"], "t2": ["C"]},
                                  checked={"t1": ["A"]})
    assert window.tree_select.get_selection() == {"t1": ["A"]}
    assert window.tree_select.get_tables() == ["t1"]
    window.tree_select.set_hint("(未加载)")
    assert window.tree_select.get_tables() == []


def test_output_mode_key_mapping(window):
    assert set(core.MODE_LABELS) <= set(core.MODE_KEYS)     # 每个下拉标签都要有对应模式
    window.out_mode_combo.setCurrentText(core.MODE_REPLACE)
    assert core.MODE_KEYS.get(window.out_mode_combo.currentText(), "append") == "replace"
    window.out_mode_combo.setCurrentText(core.MODE_APPEND)
    assert core.MODE_KEYS.get(window.out_mode_combo.currentText(), "append") == "append"
    window.out_mode_combo.setCurrentText(core.MODE_CREATE)
    assert core.MODE_KEYS.get(window.out_mode_combo.currentText(), "append") == "create_if_missing"


def test_mongodb_can_connect_without_user_or_with_uri(window):
    """MongoDB 允许免账号实例,也允许用连接 URI 代替主机:按钮不能一直置灰。"""
    window.input_combo.setCurrentText(core.INPUT_LABELS["SQL"])
    window.output_combo.setCurrentText(core.OUTPUT_LABELS["Excel"])
    window.db_type_combo.setCurrentText("MongoDB")
    window.host_edit.setText("127.0.0.1")
    window.user_edit.clear()
    window.database_edit.setText("MyData")
    window.mongo_uri_edit.clear()
    window._update_ui_state()
    assert window.conn_btn.isEnabled() and window.test_btn.isEnabled()

    window.database_edit.clear()                            # 只填 URI 也应当可以连接
    window.mongo_uri_edit.setText("mongodb://127.0.0.1:27017/MyData")
    window._update_ui_state()
    assert window.conn_btn.isEnabled()

    window.mongo_uri_edit.clear()                           # 既无主机也无 URI → 置灰
    window.host_edit.clear()
    window._update_ui_state()
    assert not window.conn_btn.isEnabled()

    window.db_type_combo.setCurrentText("MySQL")
    window.host_edit.setText("127.0.0.1")


def test_progress_bar_lifecycle(window):
    window.progress.setRange(0, 0)
    window.progress.setVisible(True)
    window._on_done()
    assert not window.progress.isVisible() and window.progress.value() == 0

# --------------------------------------------------- 自定义查询语句(第二条输入路径)
def _stub_preview(monkeypatch, preview):
    """把 preview_query 换成固定结果,记录被调用的 SourceSpec(不连真实数据库)。"""
    calls = []

    def fake(src, conn=None, adapter=None, limit=5):
        calls.append((src, conn, adapter, limit))
        return preview

    monkeypatch.setattr(gui, "preview_query", fake)
    return calls


def _wait_until(predicate, timeout=5.0):
    """后台预览是线程 + 信号:离屏测试里手动转事件循环等它回填。"""
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        QApplication.processEvents()
        if predicate():
            return True
        time.sleep(0.01)
    return predicate()


def test_query_dialog_previews_before_unlocking_ok(window, monkeypatch):
    """对话框:只读语句预览后填表头 + 锁定 3 行样本;一改语句确定就重新置灰。"""
    calls = _stub_preview(monkeypatch, core.QueryPreview(
        ["Id", "Name"], [[1, "张"], [2, "李"]], True, ""))
    window.db_type_combo.setCurrentText("MySQL")
    dialog = gui.QueryDialog(window)
    window.conn = object()                       # 不需要真的连库
    try:
        statement = "SELECT Id, Name FROM VirtualProfile LIMIT 2"
        dialog.editor.setPlainText(statement)
        assert not dialog.ok_btn.isEnabled()     # 没预览过 → 不能确定
        dialog.on_preview()
        assert _wait_until(lambda: dialog.preview_table.rowCount() == 2)
        assert dialog.ok_btn.isEnabled()
        spec, _conn, _adapter, limit = calls[-1]
        assert (spec.kind, spec.query, limit) == ("SQL", statement, 3)   # 只取 3 行样本
        assert dialog.preview_table.columnCount() == 2
        assert dialog.preview_table.rowCount() == 2
        assert "预览通过" in dialog.status.text()
        assert dialog.result_pair() == (statement, "VirtualProfile")
        dialog.editor.setPlainText("SELECT 1")
        assert not dialog.ok_btn.isEnabled()     # 语句变了 → 预览作废
        assert dialog.preview_table.rowCount() == 0
    finally:
        window.conn = None
        dialog.deleteLater()


def test_query_dialog_asks_again_for_a_writing_statement(window, monkeypatch):
    """非只读语句:状态标签显示引擎提示,弹窗选 No 不解锁确定,选 Yes 才解锁。"""
    _stub_preview(monkeypatch, core.QueryPreview(
        [], [], False, "该语句不是只读查询,预览已跳过;执行会修改数据库"))
    asked = []
    monkeypatch.setattr(gui.QMessageBox, "question",
                        lambda *args, **kwargs: asked.append(args[2]) or gui.QMessageBox.No)
    window.db_type_combo.setCurrentText("MySQL")
    dialog = gui.QueryDialog(window)
    window.conn = object()
    try:
        dialog.editor.setPlainText("UPDATE t SET a = 1")
        dialog.on_preview()
        assert _wait_until(lambda: bool(asked))
        # 弹窗口径与引擎一致:非只读语句有结果集时仍会导出
        assert "不是只读" in asked[0] and "会修改数据库" in asked[0]
        assert "返回结果集时仍会导出" in asked[0]
        assert "不是只读" in dialog.status.text()
        assert not dialog.ok_btn.isEnabled()                  # 用户选了 No

        monkeypatch.setattr(gui.QMessageBox, "question",
                            lambda *args, **kwargs: gui.QMessageBox.Yes)
        dialog.on_preview()
        assert _wait_until(lambda: dialog.ok_btn.isEnabled())  # 二次确认后可用
        # 语句里没有 from,标签退化成默认值(有 from 时取表名)
        assert dialog.result_pair() == ("UPDATE t SET a = 1", "自定义查询")
    finally:
        window.conn = None
        dialog.deleteLater()


def test_custom_query_drives_sources_ready_and_reset(window, monkeypatch):
    """确认查询后:就绪由 False 变 True、单源且语句原样、表头来自预览、初始化后复原。"""
    window.input_combo.setCurrentText(core.INPUT_LABELS["SQL"])
    window.output_combo.setCurrentText(core.OUTPUT_LABELS["Excel"])
    window.db_type_combo.setCurrentText("MySQL")
    window.selected_columns = {}
    window.selected_table = None
    window.selected_tables = []
    window._clear_query_source()
    window.tree_select.set_tables({"t1": ["A"]})
    window._update_ui_state()
    assert window.tree_select.button.isEnabled()
    assert not window._convert_ready("SQL", "Excel")

    statement = "SELECT Id, Name FROM VirtualProfile WHERE Name LIKE '%张%'"
    # 模拟对话框“确定”:主窗口拿到 (语句, 标签) 之后的接线(查询源不占用层级勾选状态)
    window.query_source = core.SourceSpec(kind="SQL", table="VirtualProfile", query=statement)
    window.query_header = ["Id", "Name"]
    window.query_label = "VirtualProfile"
    window.selected_table = None
    window.selected_tables = []
    window.selected_columns = {}
    window.query_btn.setText("自定义查询:VirtualProfile(点击修改)")
    window._update_ui_state()

    assert window._convert_ready("SQL", "Excel")
    sources = window._collect_sources("SQL", "Excel")
    assert len(sources) == 1 and sources[0].query == statement
    assert sources[0] is not window.query_source        # 返回副本,convert 写不进界面状态
    sources[0].mode = "replace"
    assert window.query_source.mode != "replace"
    assert window._current_header() == ["Id", "Name"]
    assert not window.tree_select.button.isEnabled()    # 查询源生效时层级勾选置灰
    assert window.query_btn.text().startswith("自定义查询")

    monkeypatch.setattr(gui.QMessageBox, "question",
                        lambda *args, **kwargs: gui.QMessageBox.Yes)
    window.on_reset_all()
    assert window.query_source is None and window.query_header == []
    assert window.query_label == ""
    assert window.query_btn.text() == "自定义查询…"
    window.tree_select.set_tables({"t1": ["A"]})
    window._update_ui_state()
    assert window.tree_select.button.isEnabled()        # 层级勾选入口恢复可用


def test_custom_query_is_cleared_by_the_other_input_paths(window):
    """互斥:切换输入格式 / 显式确认层级勾选 / 断开连接 都会停用自定义查询。"""
    def enable():
        window.query_source = core.SourceSpec(kind="SQL", table="t1", query="SELECT Id FROM t1")
        window.query_header = ["Id"]
        window.query_btn.setText("自定义查询:t1(点击修改)")

    window.input_combo.setCurrentText(core.INPUT_LABELS["SQL"])
    for clear in (lambda: window.on_input_mode_changed(),
                  lambda: window.on_tree_confirmed({"t1": ["A"]}),
                  window.disconnect):
        enable()
        clear()
        assert window.query_source is None and window.query_header == []
        assert window.query_btn.text() == "自定义查询…"

    window.selected_columns = {}
    window.selected_tables = []
    window.selected_table = None
    window._update_ui_state()

# ------------------------------------------------- 自定义查询语句:界面加固(复审整改)
def test_refreshing_tables_keeps_the_active_query_source(window):
    """复核 S1:层级勾选 → 启用查询 → 刷新表结构,数据源不能被旧勾选悄悄换回去。"""
    window.input_combo.setCurrentText(core.INPUT_LABELS["SQL"])
    window.output_combo.setCurrentText(core.OUTPUT_LABELS["Excel"])
    window.db_type_combo.setCurrentText("MySQL")
    window.on_tree_confirmed({"t1": ["A", "B"]})               # 先走层级勾选
    assert window.selected_tables == ["t1"]

    statement = "SELECT Id, Name FROM VirtualProfile WHERE Name LIKE '%张%'"
    window.query_source = core.SourceSpec(kind="SQL", table="VirtualProfile", query=statement)
    window.query_header = ["Id", "Name"]
    window.query_label = "VirtualProfile"
    window.query_btn.setText("自定义查询:VirtualProfile(点击修改)")

    window._on_tree_data({"t1": ["A", "B"], "t2": ["C"]})       # 点“刷新表”
    assert window.query_source is not None                      # 查询源保持有效
    sources = window._collect_sources("SQL", "Excel")
    assert len(sources) == 1 and sources[0].query == statement
    assert window._convert_ready("SQL", "Excel")

    window.on_tree_confirmed({"t1": ["A", "B"]})                # 用户显式勾选并确定
    assert window.query_source is None and window.query_header == []
    assert window._collect_sources("SQL", "Excel")[0].table == "t1"


def test_query_dialog_runs_the_preview_off_the_main_thread(window, monkeypatch):
    """慢查询不能冻住窗口:on_preview 立即返回并置忙,结果回填后填表。"""
    def slow(src, conn=None, adapter=None, limit=5):
        time.sleep(0.3)
        return core.QueryPreview(["Id", "Name"], [[1, "张"]], True, "")

    monkeypatch.setattr(gui, "preview_query", slow)
    window.db_type_combo.setCurrentText("MySQL")
    window.conn = object()
    dialog = gui.QueryDialog(window)
    try:
        dialog.editor.setPlainText("SELECT Id, Name FROM VirtualProfile LIMIT 1")
        started = time.monotonic()
        dialog.on_preview()
        assert time.monotonic() - started < 0.25                # 不等慢查询,立刻返回
        assert not dialog.preview_btn.isEnabled()               # 预览中:按钮与输入框置忙
        assert not dialog.ok_btn.isEnabled()
        assert not dialog.editor.isEnabled()
        assert "正在执行语句" in dialog.status.text()
        assert _wait_until(lambda: dialog.preview_table.rowCount() == 1)
        assert dialog.preview_btn.isEnabled() and dialog.editor.isEnabled()
        assert dialog.ok_btn.isEnabled()
        assert dialog.preview_table.columnCount() == 2
        assert "预览通过" in dialog.status.text()
    finally:
        window.conn = None
        dialog.deleteLater()


def test_query_dialog_reports_a_failed_preview(window, monkeypatch):
    """预览失败:状态标签变成错误文案,「确定」保持禁用。"""
    monkeypatch.setattr(gui.QMessageBox, "critical",
                        lambda *args, **kwargs: gui.QMessageBox.Ok)

    def boom(src, conn=None, adapter=None, limit=5):
        raise ValueError("Table 'nope' doesn't exist")

    monkeypatch.setattr(gui, "preview_query", boom)
    window.db_type_combo.setCurrentText("MySQL")
    window.conn = object()
    dialog = gui.QueryDialog(window)
    try:
        dialog.editor.setPlainText("SELECT Id FROM nope")
        dialog.on_preview()
        assert not dialog.preview_btn.isEnabled()
        assert _wait_until(lambda: dialog.preview_btn.isEnabled())
        assert "校验失败" in dialog.status.text()
        assert "nope" in dialog.status.text()
        assert not dialog.ok_btn.isEnabled()
        assert dialog.preview_table.rowCount() == 0
    finally:
        window.conn = None
        dialog.deleteLater()


def test_cancelled_query_dialog_still_refreshes_the_connection_state(window, monkeypatch):
    """预览里 ensure_conn 已经连上库、用户又取消:断开按钮与状态条必须如实更新。"""
    window.input_combo.setCurrentText(core.INPUT_LABELS["SQL"])
    window.output_combo.setCurrentText(core.OUTPUT_LABELS["Excel"])
    window.db_type_combo.setCurrentText("MySQL")
    window.conn = None
    window._update_ui_state()
    assert not window.disc_btn.isEnabled()

    class CancelledDialog:
        header = []

        def __init__(self, app):
            self.app = app

        def exec(self):
            self.app.conn = object()            # 模拟预览里的 ensure_conn 真的连上了库
            return gui.QDialog.Rejected

        def result_pair(self):
            return "", ""

        def deleteLater(self):
            pass

    monkeypatch.setattr(gui, "QueryDialog", CancelledDialog)
    try:
        window.on_custom_query()
        assert window.conn is not None
        assert window.disc_btn.isEnabled()
        assert "已连接数据库" in window.status_chip.text()
        assert window.query_source is None
    finally:
        window.conn = None
        window._update_ui_state()


def test_query_dialog_prefills_the_active_statement(window):
    """“点击修改”打开时要看到当前语句(光标在末尾),而不是空白对话框。"""
    statement = "SELECT Id, Name FROM VirtualProfile WHERE Name LIKE '%张%'"
    window.db_type_combo.setCurrentText("MySQL")
    window.query_source = core.SourceSpec(kind="SQL", table="VirtualProfile", query=statement)
    window.query_header = ["Id", "Name"]
    dialog = gui.QueryDialog(window)
    try:
        assert dialog.editor.toPlainText() == statement
        assert dialog.editor.textCursor().position() == len(statement)
        assert "已载入当前语句" in dialog.status.text()
        assert not dialog.ok_btn.isEnabled()                 # 仍然要重新校验
    finally:
        dialog.deleteLater()
        window._clear_query_source()


def test_clearing_the_query_source_logs_the_statement_digest(window):
    """停用查询源时日志留下语句摘要:换行压成空格,只取前 60 字符。"""
    statement = ("SELECT Id, Name\n  FROM VirtualProfile WHERE Name LIKE '%张%' "
                 "ORDER BY Id LIMIT 100")
    window.query_source = core.SourceSpec(kind="SQL", table="VirtualProfile", query=statement)
    window.query_header = ["Id"]
    window.query_label = "VirtualProfile"
    window._clear_query_source("改用层级勾选")

    text = window.log_text.toPlainText()
    assert "已停用自定义查询(改用层级勾选)" in text
    assert " ".join(statement.split())[:60] in text
    assert window.query_label == ""


def test_query_source_does_not_take_over_the_output_table_name(window, monkeypatch):
    """查询源不占用目标表名:否则默认“追加”必弹一次“同名追加会复制整表”。"""
    window.input_combo.setCurrentText(core.INPUT_LABELS["SQL"])
    window.output_combo.setCurrentText(core.OUTPUT_LABELS["SQL"])
    window.db_type_combo.setCurrentText("MySQL")
    window.out_mode_combo.setCurrentText(core.MODE_APPEND)
    window.on_tree_confirmed({"VirtualProfile": ["Id"]})      # 层级勾选先自动填了表名
    assert window.out_table_edit.text() == "VirtualProfile"

    class AcceptedDialog:
        header = ["Id"]

        def __init__(self, app):
            pass

        def exec(self):
            return gui.QDialog.Accepted

        def result_pair(self):
            return "SELECT Id FROM VirtualProfile", "VirtualProfile"

        def deleteLater(self):
            pass

    monkeypatch.setattr(gui, "QueryDialog", AcceptedDialog)
    window.on_custom_query()
    assert window.selected_table is None and window.selected_tables == []
    assert window.out_table_edit.text() == ""                 # 自动值被清掉,由用户自己填
    assert window.out_sheet_edit.text() == "VirtualProfile"    # Sheet 名 / JSON 键照旧推导
    assert window.query_label == "VirtualProfile"
    assert window.query_btn.text() == "自定义查询:VirtualProfile(点击修改)"

    asked, warnings = [], []
    monkeypatch.setattr(gui.QMessageBox, "question",
                        lambda *args, **kwargs: asked.append(args[2]) or gui.QMessageBox.Yes)
    monkeypatch.setattr(gui.QMessageBox, "warning",
                        lambda *args, **kwargs: warnings.append(args[2]))
    sources = window._collect_sources("SQL", "SQL")
    assert window._prepare_target("SQL", "SQL", sources) is None
    assert warnings and "请填写目标表名" in warnings[0]
    assert not any("同名" in message for message in asked)     # 不再误弹“同名追加”

    window._on_tree_data({"VirtualProfile": ["Id"]})           # 刷新表也不能把表名填回来
    assert window.out_table_edit.text() == "" and window.query_source is not None

    window._clear_query_source()
    window.output_combo.setCurrentText(core.OUTPUT_LABELS["Excel"])
    window.selected_columns = {}
    window.selected_tables = []
    window.selected_table = None
    window._update_ui_state()

# ------------------------------------------- 右下按钮列 / SSH 隧道 / 错误提示
def test_run_column_buttons_fill_the_column(window):
    """按钮列末尾没有 stretch,三个按钮纵向填满,顺序为 初始化 → 清空日志 → 开始转换。"""
    run_col = window.run_col
    assert run_col.itemAt(run_col.count() - 1).spacerItem() is None
    ordered = [run_col.itemAt(index).widget() for index in range(run_col.count())]
    assert ordered == [window.progress, window.reset_btn,
                       window.clear_btn, window.convert_btn]
    for button in (window.reset_btn, window.clear_btn, window.convert_btn):
        assert button.sizePolicy().verticalPolicy() == QSizePolicy.Expanding
    # 只交换位置:样式与点击处理仍留在各自按钮上
    assert (window.reset_btn.objectName(), window.clear_btn.objectName(),
            window.convert_btn.objectName()) == ("warn", "blue", "primary")
    assert window.convert_btn.minimumHeight() == 46


class _FakeTunnel:
    """记录 retarget 调用的假隧道:不连网络,也不允许被停止。"""

    local_port = 55555
    remote_host = "127.0.0.1"
    remote_port = 3306

    def __init__(self):
        self.active = True
        self.stopped = False
        self.calls = []

    def retarget(self, host, port):
        self.calls.append((host, port))
        self.remote_host, self.remote_port = host, port
        return self

    def describe(self):
        return f"127.0.0.1:{self.local_port} → {self.remote_host}:{self.remote_port}"

    def probe(self):
        return self.active

    def stop(self):
        self.stopped = True


def test_switching_db_type_retargets_the_tunnel(window):
    """切换数据库类型不断隧道:仍是同一个隧道对象,只把目标端口改到新类型的默认端口。"""
    window.db_type_combo.setCurrentText("MySQL")
    window.host_edit.setText("127.0.0.1")
    window.port_edit.setCurrentText("")          # 端口留空 = 用该类型的默认端口
    window.log_text.clear()
    tunnel = _FakeTunnel()
    window.tunnel = tunnel
    try:
        window.db_type_combo.setCurrentText("PostgreSQL")
        assert window.tunnel is tunnel            # 同一个对象:没被 stop,也没被替换
        assert not tunnel.stopped
        assert tunnel.calls[-1] == ("127.0.0.1", 5432)
        assert "SSH 隧道已重新指向 127.0.0.1:5432" in window.log_text.toPlainText()
        assert window.get_cfg().port == tunnel.local_port    # 连接仍走隧道端口
    finally:
        window.tunnel = None
        window.db_type_combo.setCurrentText("MySQL")


def test_dead_tunnel_is_reported_once(window):
    """探活:隧道由活变死才提示一次,并断开连接、刷新状态条;再查不会重复刷日志。"""
    window.db_type_combo.setCurrentText("MySQL")
    tunnel = _FakeTunnel().retarget("127.0.0.1", 5432)   # 隧道已指向 PostgreSQL
    tunnel.calls.clear()
    window.tunnel = tunnel
    window.conn = object()                       # 假设已经通过隧道连上库
    window._update_ui_state()
    window.log_text.clear()

    window._sync_tunnel_timer()                  # 隧道活着 → 探活定时器启动
    assert window.tunnel_timer.isActive()

    window._check_tunnel_state()                 # 隧道还活着:什么都不做
    assert window.tunnel is tunnel and window.conn is not None
    assert window.log_text.toPlainText() == ""

    tunnel.active = False
    window._check_tunnel_state()
    text = window.log_text.toPlainText()
    assert "SSH 隧道已断开" in text and "127.0.0.1:5432" in text
    assert tunnel.stopped and window.tunnel is None and window.conn is None
    assert "未连接数据库" in window.status_chip.text()
    assert not window.tunnel_timer.isActive()    # 隧道没了 → 探活定时器停止

    window.log_text.clear()
    window._check_tunnel_state()                 # 隧道已清理:不再刷第二条
    assert "SSH 隧道已断开" not in window.log_text.toPlainText()


def test_error_log_hides_the_traceback_unless_debug(window, monkeypatch):
    """出错只给中文说明;默认不打印回溯与源码行号,设置 DT_DEBUG 才输出。"""
    monkeypatch.setattr(gui.QMessageBox, "critical",
                        lambda *args, **kwargs: gui.QMessageBox.Ok)
    monkeypatch.delenv("DT_DEBUG", raising=False)
    window.db_type_combo.setCurrentText("MySQL")
    window.tunnel = None

    def boom():
        raise OSError("模拟写入失败")

    window.log_text.clear()
    assert window.run_task(boom, describe=lambda exc: f"写库失败:{exc}")
    assert _wait_until(lambda: not window._busy)
    text = window.log_text.toPlainText()
    assert "发生错误:写库失败:模拟写入失败" in text
    assert "Traceback (most recent call last)" not in text
    assert '.py", line ' not in text

    monkeypatch.setenv("DT_DEBUG", "1")
    window.log_text.clear()
    assert window.run_task(boom, describe=lambda exc: "写库失败")
    assert _wait_until(lambda: not window._busy)
    assert "Traceback (most recent call last)" in window.log_text.toPlainText()


def test_query_dialog_shows_both_examples_without_the_old_note(window):
    """对话框只留两行示例(SQL / MongoDB),旧说明段删除,SQL 示例用 %ExampleText%。"""
    window.db_type_combo.setCurrentText("MySQL")
    dialog = gui.QueryDialog(window)
    try:
        text = "\n".join(label.text() for label in dialog.findChildren(QLabel))
        assert "SQL 示例:" in text and "MongoDB 示例:" in text
        assert "%ExampleText%" in text
        assert "db.TestCollections.find" in text and "$expr" in text
        assert "另一类示例" not in text                # 旧的说明段整块删掉了
        assert "%张%" not in text
    finally:
        dialog.deleteLater()

# -------------------------------------- 复审整改:隧道守卫 / 端口 / 中文异常
def test_tunnel_check_is_skipped_while_a_task_runs(window):
    """任务进行中不清理连接(worker 可能正在用),任务结束后由定时器下一拍照常处理。"""
    window.db_type_combo.setCurrentText("MySQL")
    tunnel = _FakeTunnel()
    tunnel.active = False
    window.tunnel = tunnel
    window.conn = object()
    window.log_text.clear()
    window._busy = True
    try:
        window._check_tunnel_state()
        assert window.tunnel is tunnel and window.conn is not None
        assert not tunnel.stopped
        assert window.log_text.toPlainText() == ""
    finally:
        window._busy = False

    window._check_tunnel_state()
    assert window.tunnel is None and window.conn is None
    assert "SSH 隧道已断开" in window.log_text.toPlainText()


def test_port_follows_the_adapter_type(window):
    """端口框跟着类型走:别的引擎默认端口/留空改成新默认值,手填的自定义端口保留。"""
    window.ssh_check.setChecked(False)
    window.host_edit.setText("127.0.0.1")
    window.db_type_combo.setCurrentText("MySQL")
    window.port_edit.setCurrentText("3306")            # 别的引擎的默认端口 → 跟着切换
    tunnel = _FakeTunnel()
    window.tunnel = tunnel
    try:
        window.db_type_combo.setCurrentText("PostgreSQL")
        assert window.port_edit.currentText() == "5432"
        assert tunnel.calls[-1] == ("127.0.0.1", 5432)
        assert "127.0.0.1:5432" in window.log_text.toPlainText()

        window.port_edit.setCurrentText("15432")       # 手填自定义端口 → 保持不动
        window.db_type_combo.setCurrentText("MySQL")
        assert window.port_edit.currentText() == "15432"
        assert tunnel.calls[-1] == ("127.0.0.1", 15432)

        window.port_edit.setCurrentText("")            # 留空 → 用新类型的默认端口
        window.db_type_combo.setCurrentText("PostgreSQL")
        assert window.port_edit.currentText() == "5432"
        assert tunnel.calls[-1] == ("127.0.0.1", 5432)
    finally:
        window.tunnel = None
        window.db_type_combo.setCurrentText("MySQL")


def test_get_cfg_requires_an_established_tunnel(window, monkeypatch):
    """勾了 SSH 但隧道没建立/已断开:必须报中文错,不能静默回退成直连本机同名库。"""
    asked = []
    monkeypatch.setattr(gui.QMessageBox, "warning",
                        lambda *args, **kwargs: asked.append(args[2]) or gui.QMessageBox.Ok)
    window.db_type_combo.setCurrentText("MySQL")
    window.host_edit.setText("127.0.0.1")
    window.tunnel = None
    window.ssh_check.setChecked(True)
    try:
        with pytest.raises(gui.DTError) as missing:
            window.get_cfg()
        assert "隧道" in str(missing.value)
        assert window.config_or_warn() is None         # 弹窗提示,不向外抛
        assert asked and "隧道" in asked[-1]

        dead = _FakeTunnel()
        dead.active = False                            # 隧道对象还在,但已经断了
        window.tunnel = dead
        with pytest.raises(gui.DTError):
            window.get_cfg()

        class BadTunnel(_FakeTunnel):
            def retarget(self, host, port):
                raise ValueError("远程数据库端口不合法:0")

        window.ssh_check.setChecked(False)             # retarget 的 ValueError 也包成 DTError
        window.tunnel = BadTunnel()
        with pytest.raises(gui.DTError) as broken:
            window.get_cfg()
        assert "端口不合法" in str(broken.value)
    finally:
        window.tunnel = None
        window.ssh_check.setChecked(False)


def test_entry_points_stop_when_the_tunnel_is_missing(window, monkeypatch):
    """五个取配置的入口:隧道缺失时提示并直接返回,不再启动任务。"""
    monkeypatch.setattr(gui.QMessageBox, "warning", lambda *args, **kwargs: gui.QMessageBox.Ok)
    calls = []
    monkeypatch.setattr(window, "run_task",
                        lambda *args, **kwargs: calls.append("run_task") or True)
    monkeypatch.setattr(window, "ensure_conn",
                        lambda *args, **kwargs: calls.append("ensure_conn"))
    window.db_type_combo.setCurrentText("MySQL")
    window.host_edit.setText("127.0.0.1")
    window.tunnel = None
    window.ssh_check.setChecked(True)
    try:
        window.on_test_connection()
        window.on_connect()
        window.on_refresh_tables()
        assert calls == []

        monkeypatch.setattr(window, "_collect_sources", lambda *args, **kwargs: ["src"])
        monkeypatch.setattr(window, "_prepare_target", lambda *args, **kwargs:
                            core.TargetSpec(kind="Excel", path="out.xlsx"))
        window.on_convert()
        assert calls == []

        assert window._apply_sheet_decisions([]) is False
        assert calls == []
    finally:
        window.tunnel = None
        window.ssh_check.setChecked(False)


def test_error_message_only_short_circuits_dt_error():
    """DTError 原文即中文;pymysql 自己的 RuntimeError 也必须走 describe 翻译。"""
    seen = []

    def describe(exc):
        seen.append(str(exc))
        return f"翻译后:{exc}"

    assert isinstance(gui.DTError("中文A"), RuntimeError)
    assert gui.error_message(gui.DTError("中文A"), describe) == "中文A"
    assert seen == []

    plain = RuntimeError("cryptography package is required for sha256_password")
    assert gui.error_message(plain, describe) == f"翻译后:{plain}"
    assert seen == [str(plain)]


def test_sheet_decisions_show_dt_error_text_verbatim(window, monkeypatch):
    """_apply_sheet_decisions:DTError 原文直接展示,不再套第二层“原始错误:”。"""
    shown = []
    monkeypatch.setattr(gui.QMessageBox, "critical",
                        lambda *args, **kwargs: shown.append(args[2]))
    window.db_type_combo.setCurrentText("MySQL")
    window.ssh_check.setChecked(False)
    window.tunnel = None

    def boom(cfg=None):
        raise gui.DTError("中文A")

    monkeypatch.setattr(window, "ensure_conn", boom)
    assert window._apply_sheet_decisions([]) is False
    assert shown == ["中文A"]


def test_tunnel_target_is_an_atomic_snapshot():
    """retarget 整体替换转发目标:_forward 取到成对的 host/port,读不出半新半旧。"""
    tunnel = core.SSHTunnel(core.SSHConfig(host="ssh.example", port=22,
                                           user="u", password="p"),
                            remote_host="127.0.0.1", remote_port=3306)
    assert tunnel.retarget("10.0.0.9", 5432) is tunnel
    assert (tunnel.remote_host, tunnel.remote_port) == ("10.0.0.9", 5432)
    assert "10.0.0.9:5432" in tunnel.describe()
    with pytest.raises(AttributeError):
        tunnel.remote_host = "1.1.1.1"           # 只读:只能整体替换

    seen = {}

    class _Sock:
        def recv(self, _size):
            return b""                           # 立刻结束转发,不留后台线程

        def sendall(self, _data):
            pass

        def close(self):
            pass

        def getpeername(self):
            return ("127.0.0.1", 4321)

    class _Transport:
        def open_channel(self, kind, destination, source):
            seen.update(kind=kind, destination=destination, source=source)
            return _Sock()

    tunnel._transport = _Transport()
    tunnel._forward(_Sock())
    assert seen["destination"] == ("10.0.0.9", 5432)
    assert seen["kind"] == "direct-tcpip"


def test_tunnel_break_hint_requires_a_dead_tunnel():
    """只有“链路中断 + 隧道当前不活跃”才追加提示:锁等待超时不会被误报成隧道断了。"""
    class _Tunnel:
        active = True

    alive, dead = _Tunnel(), _Tunnel()
    dead.active = False
    assert gui.tunnel_break_hint("Lock wait timeout exceeded", dead) == ""
    assert gui.tunnel_break_hint("connection reset by peer", alive) == ""
    assert gui.tunnel_break_hint("connection reset by peer", None) == ""
    assert gui.tunnel_break_hint("connection reset by peer", dead) == gui.TUNNEL_BREAK_HINT
    assert gui.tunnel_break_hint("Software caused connection abort", dead) == gui.TUNNEL_BREAK_HINT


def test_tunnel_probe_detects_a_dead_transport():
    """probe() 用一次真实往返确认通道:send_ignore 发不出去就视为断开。"""
    class _Transport:
        def __init__(self, broken):
            self.broken = broken
            self.sent = 0

        def is_active(self):
            return True

        def send_ignore(self):
            self.sent += 1
            if self.broken:
                raise OSError("socket is closed")

    tunnel = core.SSHTunnel(core.SSHConfig(host="ssh.example", port=22,
                                           user="u", password="p"),
                            remote_host="127.0.0.1", remote_port=5432)
    tunnel._listener = object()                  # active 还要求监听套接字还在
    healthy = _Transport(broken=False)
    tunnel._transport = healthy
    assert tunnel.probe() is True and healthy.sent == 1

    broken = _Transport(broken=True)
    tunnel._transport = broken
    assert tunnel.probe() is False and broken.sent == 1

    tunnel._listener = None
    assert tunnel.probe() is False               # 本地标志都不活跃:不必再发探活包

