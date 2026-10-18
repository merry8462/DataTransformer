# -*- coding: utf-8 -*-
"""命令行向导测试:以管道输入驱动完整交互流程(非 TTY 回退模式)。"""

import os
import subprocess
import sys
from pathlib import Path

import pytest

from src import dt_cli
from src.dt_core import DBConfig, FieldMapping, QueryPreview, SourceSpec, TargetSpec
from src.dt_numeric import loads_json

ROOT = Path(__file__).resolve().parent.parent


def run_cli(answers, *args, timeout=120):
    env = dict(os.environ, PYTHONIOENCODING="utf-8", DT_PLAIN="1", PYTHONUTF8="1")
    return subprocess.run(
        [sys.executable, "-m", "src.dt_cli", *args],
        input="\n".join(answers) + "\n", capture_output=True, text=True,
        encoding="utf-8", errors="replace", cwd=str(ROOT), env=env, timeout=timeout)


def test_cli_version_and_help():
    version = run_cli([], "--version")
    assert version.returncode == 0 and version.stdout.strip().count(".") >= 1
    help_text = run_cli([], "--help")
    assert help_text.returncode == 0 and "--demo" in help_text.stdout


def test_cli_csv_to_json_full_flow(tmp_path):
    """按向导顺序:输入 CSV → 选分隔符/编码 → 输出 JSON → 导出路径 → 执行。"""
    source = tmp_path / "people.csv"
    source.write_text("Name,Amount,Code\n张三,12345678901234567890.12345,007\n",
                      encoding="utf-8-sig")
    target = tmp_path / "people.json"
    proc = run_cli([
        "6",                        # ① 输入格式:CSV
        str(source),                # 输入文件路径
        "1",                        # CSV 分隔符:逗号
        "1",                        # CSV 编码:UTF-8 with BOM
        "5",                        # ③ 输出格式:JSON
        "n",                        # CSV 不按类型解析(保持文本原样)
        str(target),                # 导出文件路径
        "n",                        # 字段映射:保持同名
        "y",                        # 确认开始转换
        "n",                        # 不再继续
    ])
    assert proc.returncode == 0, proc.stderr
    assert target.exists(), proc.stdout
    text = target.read_text(encoding="utf-8")
    document = loads_json(text)                     # 用 Decimal 读回,证明文件本身未失真
    assert document["people"]["HeaderFields"] == ["Name", "Amount", "Code"]
    # CSV→文件默认保持字面量不变(长浮点数不做 float 推断,前导零不丢)
    assert document["people"]["Data"]["Row1"] == ["张三", "12345678901234567890.12345", "007"]
    assert "12345678901234567890.12345" in text and "e+" not in text
    assert "数值保真校验" in proc.stdout                      # 校验报告可见


def test_cli_reports_missing_file_and_exits_cleanly(tmp_path):
    proc = run_cli([str(tmp_path / "nope.csv")])
    assert proc.returncode == 0
    assert "Traceback" not in proc.stdout + proc.stderr
    assert "输入结束" in proc.stdout or "文件不存在" in proc.stdout


def test_cli_writes_nothing_on_abort(tmp_path):
    source = tmp_path / "a.csv"
    source.write_text("A\n1\n", encoding="utf-8-sig")
    target = tmp_path / "a.json"
    proc = run_cli([
        "6", str(source), "1", "1", "5", str(target),
        "n", "n", "n",              # 确认开始转换 → 否
        "n",
    ])
    assert proc.returncode == 0 and not target.exists()
    assert "已取消转换" in proc.stdout


@pytest.mark.parametrize("mode_name,index", [("CSV", "6"), ("JSON", "5")])
def test_cli_all_file_modes_present(tmp_path, mode_name, index):
    proc = run_cli([index, ""], timeout=60)                  # 只走到选择输入文件
    assert f"{mode_name}" in proc.stdout


# ---------------------------------------------------------------------------
# 自由填写查询语句分支:替换交互原语与预览函数,不连接真实数据库
# ---------------------------------------------------------------------------
SQL_TEXT = "SELECT Id, Name FROM VirtualProfile WHERE Name LIKE '%张%' ORDER BY Id LIMIT 100"
MONGO_TEXT = 'db.Coll.find({"a": 1}).limit(2)'          # 引擎只认 db.<集合>.find 两段写法


class FakeAdapter:
    NAME = "MySQL"
    KIND = "SQL"
    DEFAULT_PORT = 3306

    def list_tables(self, conn):
        return ["t1"]

    def list_columns(self, conn, name, where=None):
        return ["Id", "Name"]


def fake_db(name="MySQL"):
    return dt_cli.Database(name=name, adapter=FakeAdapter(), conn=object(),
                           cfg=DBConfig(database="demo"))


def patch_interaction(monkeypatch, statements, confirm=None, logs=None):
    """把 ask / ask_yes_no / log 换成脚本化回答;confirm(prompt, default) 决定确认结果。"""
    queue = iter(statements)
    monkeypatch.setattr(dt_cli, "ask", lambda prompt, default="", allow_empty=False,
                        param=False: next(queue))
    monkeypatch.setattr(dt_cli, "ask_yes_no",
                        confirm or (lambda prompt, default=False: default))
    monkeypatch.setattr(dt_cli, "log", lambda message, level="info":
                        logs.append((level, message)) if logs is not None else None)


def test_pick_query_source_readonly_returns_source(monkeypatch):
    seen = {}

    def fake_preview(src, conn=None, adapter=None, limit=5):
        seen["src"], seen["limit"] = src, limit
        return QueryPreview(["Id", "Name"], [["1", "张三"], ["2", "李四"]], True)

    monkeypatch.setattr(dt_cli, "preview_query", fake_preview)
    patch_interaction(monkeypatch, [SQL_TEXT])

    sources, header = dt_cli.pick_query_source(fake_db())

    assert len(sources) == 1
    assert sources[0].kind == "SQL"
    assert sources[0].query == SQL_TEXT                 # 语句逐字保留,不裁剪字段
    assert sources[0].table == "VirtualProfile"         # label 取自语句里的表名
    assert sources[0].columns == []
    assert header == ["Id", "Name"]                     # 预览表头沿调用链交给字段映射
    assert seen["limit"] == 3 and seen["src"].query == SQL_TEXT


def test_pick_query_source_retries_after_preview_error(monkeypatch):
    attempts, logs = [], []

    def fake_preview(src, conn=None, adapter=None, limit=5):
        attempts.append(src.query)
        if len(attempts) == 1:
            raise ValueError("Table 'nope' doesn't exist")
        return QueryPreview(["Id"], [["1"]], True)

    monkeypatch.setattr(dt_cli, "preview_query", fake_preview)
    patch_interaction(monkeypatch, ["SELECT * FROM nope", SQL_TEXT], logs=logs)

    sources, _header = dt_cli.pick_query_source(fake_db())

    assert attempts == ["SELECT * FROM nope", SQL_TEXT]  # 第一次失败后重新填写
    assert sources[0].query == SQL_TEXT
    assert [level for level, _ in logs].count("err") == 1


def test_pick_query_source_non_readonly_retries_on_default_answer(monkeypatch):
    logs = []

    def fake_preview(src, conn=None, adapter=None, limit=5):
        if src.query.lower().startswith("update"):
            return QueryPreview([], [], False, "该语句不是只读查询,预览已跳过")
        return QueryPreview(["A"], [["1"]], True)

    monkeypatch.setattr(dt_cli, "preview_query", fake_preview)
    patch_interaction(monkeypatch, ["update t set a=1", SQL_TEXT], logs=logs)

    sources, _header = dt_cli.pick_query_source(fake_db())

    assert sources[0].query == SQL_TEXT                  # 默认(No)时危险语句被丢弃
    assert ("warn", "该语句不是只读查询,预览已跳过") in logs


def test_pick_query_source_non_readonly_accepted_when_confirmed(monkeypatch):
    monkeypatch.setattr(dt_cli, "preview_query", lambda src, conn=None, adapter=None,
                        limit=5: QueryPreview([], [], False, "该语句不是只读查询,预览已跳过"))
    patch_interaction(monkeypatch, ["update t set a=1"], confirm=lambda prompt, default=False: True)

    sources, _header = dt_cli.pick_query_source(fake_db())

    assert len(sources) == 1 and sources[0].query == "update t set a=1"
    assert sources[0].table == "自定义查询"              # 语句里没有 from,label 走兜底


def test_pick_query_source_mongo_label(monkeypatch):
    def fake_preview(src, conn=None, adapter=None, limit=5):
        assert src.kind == "Mongo"
        return QueryPreview(["a"], [["1"]], True)

    monkeypatch.setattr(dt_cli, "preview_query", fake_preview)
    patch_interaction(monkeypatch, [MONGO_TEXT])

    sources, header = dt_cli.pick_query_source(fake_db("MongoDB"))

    assert sources[0].kind == "Mongo"
    assert sources[0].query == MONGO_TEXT
    assert sources[0].table == "Coll"                    # 集合名来自引擎的 parse_mongo_query
    assert header == ["a"]


def test_pick_query_source_aborts_on_quit(monkeypatch):
    monkeypatch.setattr(dt_cli, "preview_query",
                        lambda *a, **k: pytest.fail(":q 不应触发预览"))
    patch_interaction(monkeypatch, [":q"])

    with pytest.raises(dt_cli.Aborted):
        dt_cli.pick_query_source(fake_db())


def test_pick_database_sources_routes_to_query_branch(monkeypatch):
    monkeypatch.setattr(dt_cli, "choose_option",
                        lambda prompt, options, default=1, param=False: "query")
    monkeypatch.setattr(dt_cli, "pick_query_source",
                        lambda db: (["FROM_QUERY_BRANCH"], ["Id"]))
    assert dt_cli.pick_database_sources(fake_db()) == (["FROM_QUERY_BRANCH"], ["Id"])


def test_pick_database_sources_hierarchy_returns_no_header(monkeypatch):
    """层级勾选分支行为不变:表 + 字段照旧,表头交给后端 _preview_header 取。"""
    monkeypatch.setattr(dt_cli, "choose_option",
                        lambda prompt, options, default=1, param=False: "pick")
    monkeypatch.setattr(dt_cli, "pick_many",
                        lambda prompt, names, checked=None: list(names))

    sources, header = dt_cli.pick_database_sources(fake_db())

    assert header is None
    assert [source.table for source in sources] == ["t1"]
    assert sources[0].columns == ["Id", "Name"]


def test_preview_header_skips_query_source(monkeypatch):
    def boom(*args, **kwargs):
        raise AssertionError("查询源不应再执行一次语句")

    monkeypatch.setattr(dt_cli, "open_reader", boom)
    source = SourceSpec(kind="SQL", table="t", query="update t set a=1")
    assert dt_cli._preview_header([source], fake_db()) == []

def test_pick_query_source_gives_up_after_five_failed_attempts(monkeypatch):
    """预览一直失败时不再无限重填:第 5 次之后抛 Aborted。"""
    attempts = []

    def fake_preview(src, conn=None, adapter=None, limit=5):
        attempts.append(src.query)
        raise ValueError("syntax error at or near")

    monkeypatch.setattr(dt_cli, "preview_query", fake_preview)
    patch_interaction(monkeypatch, [f"SELECT {index} FROM nope" for index in range(6)])

    with pytest.raises(dt_cli.Aborted, match="重试次数过多"):
        dt_cli.pick_query_source(fake_db())
    assert len(attempts) == dt_cli.QUERY_MAX_ATTEMPTS == 5


def test_run_once_passes_the_query_header_to_field_mapping(monkeypatch):
    """查询源的表头要真的交给 ask_field_mapping:否则字段映射被静默跳过。"""
    seen = {}

    def fake_mapping(header):
        seen["header"] = list(header)
        return FieldMapping()

    target = TargetSpec(kind="CSV", path="out.csv")
    monkeypatch.setattr(dt_cli, "load_settings", lambda: {})
    monkeypatch.setattr(dt_cli, "choose_mode",
                        lambda prompt: "MySQL" if "输入" in prompt else "CSV")
    monkeypatch.setattr(dt_cli, "open_database", lambda settings, mode, use_demo=False: fake_db())
    monkeypatch.setattr(dt_cli, "pick_database_sources",
                        lambda db: ([SourceSpec(kind="SQL", table="VirtualProfile",
                                                query=SQL_TEXT)], ["Id", "Name"]))
    monkeypatch.setattr(dt_cli, "ask_name_mapping", lambda names, label: {})
    monkeypatch.setattr(dt_cli, "build_file_target",
                        lambda mode, in_mode, sources, base, names=None: target)
    monkeypatch.setattr(dt_cli, "ask_field_mapping", fake_mapping)
    monkeypatch.setattr(dt_cli, "ask_yes_no", lambda prompt, default=False: default)
    monkeypatch.setattr(dt_cli, "_save_profile", lambda settings, in_db, out_db: None)
    monkeypatch.setattr(dt_cli.Database, "close", lambda self: None)
    monkeypatch.setattr(dt_cli, "convert", lambda *args, **kwargs: FakeResult())

    dt_cli.run_once()

    assert seen["header"] == ["Id", "Name"]


class FakeResult:
    """convert 的最小返回值:只提供 run_once 收尾要用的两个成员。"""

    class Fidelity:
        checked_values = []

    fidelity = Fidelity()

    def summary(self):
        return "1 行"

