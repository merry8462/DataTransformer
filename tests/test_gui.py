# -*- coding: utf-8 -*-
"""图形界面关键逻辑测试(未安装 PySide6 时自动跳过)。

覆盖:主窗口构造、MongoDB / SQL 配置切换、字段映射应用、层级勾选控件、
输出基准名与写入模式映射。
"""

import os
import sys

import pytest

pytest.importorskip("PySide6", reason="未安装 PySide6,跳过图形界面测试")

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")          # 无显示器环境可运行
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

from PySide6.QtWidgets import QApplication                   # noqa: E402

import data_transformer as gui                               # noqa: E402
import dt_core as core                                       # noqa: E402


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

        def describe(self):
            return "127.0.0.1:55555"

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
