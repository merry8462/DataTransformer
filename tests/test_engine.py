# -*- coding: utf-8 -*-
"""转换引擎与映射规则测试(含长浮点数跨格式往返的验收用例)。"""

from decimal import Decimal

import pytest

import dt_core as core
from dt_numeric import FidelityError, format_number

LONG_VALUES = [
    "12345678901234567890.12345",          # 20 位整数 + 5 位小数
    "1.500",                               # 尾零
    "-0.000000000000000000001",            # 极小负数
    "9007199254740993",                    # 超出双精度可精确表示的整数
    "1E+20",                               # 科学计数法输入 → 定点输出
    "0.1",
    "100.00",
]
EXPECTED = [format_number(value) for value in LONG_VALUES]


def write_source_csv(path) -> None:
    core.write_csv_file(str(path), ["Id", "Amount"],
                        iter([[index + 1, value] for index, value in enumerate(LONG_VALUES)]))


def read_csv_column(path, index=1):
    lines = path.read_text(encoding="utf-8-sig").strip().splitlines()[1:]
    return [line.split(",")[index] for line in lines]


def test_append_readback_ignores_pre_existing_rows(tmp_path):
    """追加写入已有目标时,回读要能找到本次写进去的行,而不是只看开头。

    否则目标里原有的旧行会占满回读窗口,本次的数据被误报成"(缺失)"
    (xlsx → 已有 MySQL 表 append 就是这么报错的)。
    """
    header = ["Id", "IdCard"]
    first = tmp_path / "first.csv"
    second = tmp_path / "second.csv"
    combined = tmp_path / "combined.csv"
    core.write_csv_file(str(first), header,
                        iter([[index, f"1{index:017d}"] for index in range(1, 301)]))
    core.write_csv_file(str(second), header,
                        iter([[70000 + index, f"{70000 + index:018d}"] for index in range(1, 301)]))

    core.convert(
        [core.SourceSpec(kind="CSV", path=str(first), parse=True)],
        core.TargetSpec(kind="CSV", path=str(combined))
    )
    merged = core.convert(
        [core.SourceSpec(kind="CSV", path=str(second), parse=True)],
        core.TargetSpec(kind="CSV", path=str(combined), mode="merge")
    )

    assert merged.fidelity.ok, merged.fidelity.summary_lines()
    lines = combined.read_text(encoding="utf-8-sig").strip().splitlines()
    assert len(lines) == 601                       # 表头 + 两次各 300 行
    assert lines[-1].split(",")[0] == "70300"


@pytest.mark.parametrize("kind,ext,extra", [
    ("CSV", "csv", {}), ("JSON", "json", {"key": "info"}), ("Excel", "xlsx", {"sheet": "info"}),
])
def test_long_float_roundtrip_between_files(tmp_path, kind, ext, extra):
    """CSV → 目标格式 → CSV:数值与文本都必须与输入一致,且输出不得有科学计数法。"""
    source = tmp_path / "in.csv"
    write_source_csv(source)
    produced = tmp_path / f"out.{ext}"

    forward = core.convert([core.SourceSpec(kind="CSV", path=str(source), parse=True)],
                           core.TargetSpec(kind=kind, path=str(produced), **extra))
    assert forward.fidelity.ok and forward.fidelity.checked_values >= len(LONG_VALUES)

    if ext != "xlsx":                       # xlsx 是二进制包,只能解析后校验数值
        text = produced.read_text(encoding="utf-8")
        assert "e+" not in text and "E+" not in text

    back = tmp_path / "back.csv"
    core.convert(
        [core.SourceSpec(kind=kind, path=str(produced), table="" if kind == "CSV" else "info")],
        core.TargetSpec(kind="CSV", path=str(back))
    )
    assert read_csv_column(back) == EXPECTED


def test_scientific_notation_input_is_normalised_to_fixed_point(tmp_path):
    """科学计数法输入按定点输出(数值等价,文本规范化)。"""
    source = tmp_path / "sci.csv"
    core.write_csv_file(str(source), ["A", "B"], iter([["1E+20", "-2.5e-7"]]))
    produced = tmp_path / "sci.json"
    core.convert(
        [core.SourceSpec(kind="CSV", path=str(source), parse=True)],
        core.TargetSpec(kind="JSON", path=str(produced), key="t")
    )
    text = produced.read_text(encoding="utf-8")
    assert "100000000000000000000" in text and "-0.00000025" in text
    assert "e+" not in text and "E+" not in text


def lossy_csv(monkeypatch):
    """把输出值截断到 8 位的“有损写入器”,用于验证保真校验能阻断流程。"""
    original = core.write_csv_file

    def write_csv_file(path, columns, row_iterator, **kwargs):
        rows = [[format_number(cell)[:8] for cell in row] for row in row_iterator]
        return original(path, columns, iter(rows), **kwargs)

    monkeypatch.setattr(core, "write_csv_file", write_csv_file)


def test_fidelity_failure_blocks_conversion(tmp_path, monkeypatch):
    """输出侧丢失精度时必须阻断流程并给出差异报告。"""
    source = tmp_path / "in.csv"
    write_source_csv(source)
    lossy_csv(monkeypatch)
    with pytest.raises(FidelityError) as excinfo:
        core.convert(
            [core.SourceSpec(kind="CSV", path=str(source), parse=True)],
            core.TargetSpec(kind="CSV", path=str(tmp_path / "out.csv"))
        )
    assert "数值保真校验未通过" in str(excinfo.value)
    assert excinfo.value.report.mismatches


def test_allow_mismatch_lets_conversion_finish(tmp_path, monkeypatch):
    source = tmp_path / "in.csv"
    write_source_csv(source)
    lossy_csv(monkeypatch)
    result = core.convert(
        [core.SourceSpec(kind="CSV", path=str(source), parse=True)],
        core.TargetSpec(kind="CSV", path=str(tmp_path / "ok.csv"), allow_mismatch=True)
    )
    assert not result.fidelity.ok and result.rows == len(LONG_VALUES)


def test_field_mapping_renames_output_fields(tmp_path):
    source = tmp_path / "in.csv"
    core.write_csv_file(str(source), ["姓名", "金额"], iter([["张三", "1.500"]]))
    produced = tmp_path / "mapped.json"
    target = core.TargetSpec(
        kind="JSON", path=str(produced), key="info",
        fields=core.FieldMapping.from_pairs({"姓名": "Name", "金额": "Amount"})
    )
    core.convert([core.SourceSpec(kind="CSV", path=str(source), parse=True)], target)
    header, rows = core.load_json_stream(str(produced), "info")
    assert header == ["Name", "Amount"] and list(rows)[0] == ["张三", Decimal("1.500")]


def test_field_mapping_rejects_duplicate_targets(tmp_path):
    source = tmp_path / "in.csv"
    core.write_csv_file(str(source), ["A", "B"], iter([[1, 2]]))
    target = core.TargetSpec(
        kind="JSON", path=str(tmp_path / "dup.json"), key="k",
        fields=core.FieldMapping.from_pairs({"A": "X", "B": "X"})
    )
    with pytest.raises(ValueError, match="重名字段"):
        core.convert([core.SourceSpec(kind="CSV", path=str(source))], target)


def test_name_mapping_renames_sheets(tmp_path):
    """Sheet 名 ↔ 表名映射规则。"""
    source = tmp_path / "in.csv"
    core.write_csv_file(str(source), ["A"], iter([[1]]))
    produced = tmp_path / "out.xlsx"
    target = core.TargetSpec(kind="Excel", path=str(produced), names={"raw": "renamed"})
    core.convert([core.SourceSpec(kind="CSV", path=str(source), table="raw")], target)
    assert "renamed" in core.workbook_sheet_headers(str(produced))


def test_multi_source_to_excel_and_json(tmp_path):
    source = tmp_path / "t.csv"
    core.write_csv_file(str(source), ["A"], iter([[1]]))
    sources = [
        core.SourceSpec(kind="CSV", path=str(source), table="t1"),
        core.SourceSpec(kind="CSV", path=str(source), table="t2")
    ]
    excel = tmp_path / "m.xlsx"
    result = core.convert(sources, core.TargetSpec(kind="Excel", path=str(excel)))
    assert set(core.workbook_sheet_headers(str(excel))) == {"t1", "t2"} and result.rows == 2

    js = tmp_path / "m.json"
    core.convert(
        [core.SourceSpec(kind="CSV", path=str(source), table="t1"), core.SourceSpec(kind="CSV", path=str(source), table="t2")],
        core.TargetSpec(kind="JSON", path=str(js))
    )
    assert set(core.json_top_keys(str(js))) == {"t1", "t2"}


def test_skip_and_directory_modes(tmp_path):
    source = tmp_path / "t.csv"
    core.write_csv_file(str(source), ["A"], iter([[1]]))
    target_dir = tmp_path / "csvdir"
    result = core.convert(
        [core.SourceSpec(kind="CSV", path=str(source), table="t1", mode="skip"), core.SourceSpec(kind="CSV", path=str(source), table="t2")],
        core.TargetSpec(kind="CSV", path=str(target_dir), directory=True)
    )
    assert result.skipped == 1 and result.sources == 1
    assert (target_dir / "t2.csv").exists() and not (target_dir / "t1.csv").exists()


def test_guard_rails(tmp_path):
    source = tmp_path / "t.csv"
    core.write_csv_file(str(source), ["A"], iter([[1]]))
    with pytest.raises(ValueError, match="没有可转换的输入源"):
        core.convert([], core.TargetSpec(kind="JSON", path=str(tmp_path / "x.json")))
    with pytest.raises(ValueError, match="需要先建立数据库连接"):
        core.convert([core.SourceSpec(kind="CSV", path=str(source))], core.TargetSpec(kind="SQL", table="t"))
    with pytest.raises(ValueError, match="目录模式"):
        core.convert(
            [core.SourceSpec(kind="CSV", path=str(source), table="a"), core.SourceSpec(kind="CSV", path=str(source), table="b")],
            core.TargetSpec(kind="CSV", path=str(tmp_path / "one.csv"))
        )


def test_section_csv_compatibility(tmp_path):
    """旧版分节 CSV 仍可读取(向后兼容)。"""
    legacy = tmp_path / "legacy.csv"
    legacy.write_text("[Sheet:info]\nId,Name\n1,张三\n", encoding="utf-8-sig")
    assert core.scan_csv_sections(str(legacy), ",", "utf-8-sig") == ["info"]
    produced = tmp_path / "legacy.json"
    result = core.convert(
        [core.SourceSpec(kind="CSV", path=str(legacy), table="info")],
        core.TargetSpec(kind="JSON", path=str(produced), key="info")
    )
    assert result.rows == 1
    assert list(core.load_json_stream(str(produced), "info")[1])[0] == ["1", "张三"]


def test_readback_source_covers_every_target_kind(tmp_path):
    src = core.SourceSpec(kind="CSV", path=str(tmp_path / "t.csv"), table="t1")
    assert core.readback_source(src, core.TargetSpec(kind="SQL", table="x"), ["A"]).table == "x"
    assert core.readback_source(src, core.TargetSpec(kind="Mongo", table="c"), ["A"]).kind == "Mongo"
    excel = core.readback_source(src, core.TargetSpec(kind="Excel", sheet="s"), ["A"])
    assert excel.table == "s" and excel.kind == "Excel"
    js = core.readback_source(src, core.TargetSpec(kind="JSON", key="k"), ["A"])
    assert js.table == "k" and js.kind == "JSON"


def test_apply_field_mapping_keeps_order(tmp_path):
    header = ["A", "B", "C"]
    target = core.TargetSpec(kind="JSON", fields=core.FieldMapping.from_pairs({"B": "Bee"}))
    assert core.apply_field_mapping(header, target) == ["A", "Bee", "C"]


def test_multi_source_excel_readback_verifies_each_sheet(tmp_path):
    """多表写同一个 xlsx 时,回读必须按各源的 Sheet 名校验,而不是用户填写的 Sheet 名。

    否则回读会去找一个不存在的 Sheet,数值保真校验被整段静默跳过。
    """
    source = tmp_path / "in.csv"
    core.write_csv_file(str(source), ["Amount"],
                        iter([["1.500"], ["12345678901234567890.12345"]]))
    excel = tmp_path / "multi.xlsx"
    result = core.convert(
        [core.SourceSpec(kind="CSV", path=str(source), table="t1"), core.SourceSpec(kind="CSV", path=str(source), table="t2")],
        core.TargetSpec(kind="Excel", path=str(excel), sheet="用户填写的Sheet名"))

    assert set(core.workbook_sheet_headers(str(excel))) == {"t1", "t2"}
    assert result.fidelity.checked_values == 4           # 2 张表 × 2 个数值
    assert result.fidelity.ok, result.fidelity.summary_lines()
    assert not [note for note in result.fidelity.notes if "回读输出失败" in note]


def test_empty_source_is_reported_clearly(tmp_path):
    """空数组(没有任何字段)要给出中文提示,而不是让下游拼出非法 SQL。"""
    empty = tmp_path / "empty.json"
    empty.write_text('{"nothing": []}', encoding="utf-8")
    with pytest.raises(ValueError, match="没有任何字段"):
        core.convert(
            [core.SourceSpec(kind="JSON", path=str(empty), table="nothing")],
            core.TargetSpec(kind="JSON", path=str(tmp_path / "out.json"), key="k")
        )
