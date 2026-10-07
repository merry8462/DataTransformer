# -*- coding: utf-8 -*-
"""长浮点数 / 高精度数值保真专项测试(对应需求的数值一致性要求)。"""

from decimal import Decimal

import pytest

from dt_numeric import (
    FidelityReport, NumericRecorder, dumps_json, excel_cell_value, format_number,
    has_scientific_notation, is_number_literal, loads_json, parse_number,
    significant_digits, to_decimal,
)

LONG_VALUES = [
    "12345678901234567890.12345",         # 20 位整数 + 5 位小数
    "0.1000000000000000055511151231257827",
    "-1234567890.000000000000000001",
    "1.500",                              # 尾零必须保留
    "-0.000000000000000000001",
    "9007199254740993",                   # 超出双精度可表示范围的整数
    "1e-30",                              # 科学计数法输入 → 定点输出
    "1E+20",
]


@pytest.mark.parametrize("literal", LONG_VALUES)
def test_parse_and_format_roundtrip_keeps_text(literal):
    """科学计数法输入按定点输出,其余字面量文本逐字保持。"""
    parsed = parse_number(literal)
    assert parsed is not None
    text = format_number(parsed)
    assert not has_scientific_notation(text)
    assert to_decimal(literal) == to_decimal(text)      # 数值等值
    if "e" not in literal.lower():
        assert text == literal                          # 文本等值(尾零不丢)


def test_decimal_from_float_uses_decimal_literal_not_binary():
    assert to_decimal(0.1) == Decimal("0.1")
    assert format_number(0.1) == "0.1"
    assert format_number(1e21) == "1000000000000000000000"


def test_no_scientific_notation_in_any_output():
    for value in (Decimal("1E+20"), 1e20, Decimal("1e-30"), -1e-7, "2.5e3"):
        text = format_number(value)
        assert not has_scientific_notation(text), text
        assert "e" not in text.lower()


def test_trailing_zeros_are_preserved():
    assert format_number(Decimal("1.500")) == "1.500"
    assert format_number(Decimal("100.00")) == "100.00"
    assert format_number(Decimal("-0.0")) == "-0.0"


def test_sign_rules():
    assert format_number(Decimal("-1.5")) == "-1.5"
    assert format_number(Decimal("+1.5")) == "1.5"        # 冗余正号按规范输出
    assert format_number("-0.000000000000000000001") == "-0.000000000000000000001"


def test_padded_numbers_stay_text():
    for padded in ("007", "0012", "+008"):
        assert not is_number_literal(padded)
        assert parse_number(padded) is None
        assert format_number(padded) == padded
    assert parse_number("007", keep_padded=False) == 7


def test_non_numeric_text_untouched():
    for text in ("", "abc", "12a", "1,000", "1.2.3", "1e", "--1"):
        assert parse_number(text) is None
        assert format_number(text) == text


def test_plain_integers_become_int_not_float():
    assert parse_number("123456789012345678901234567890") == 123456789012345678901234567890
    assert isinstance(parse_number("42"), int)
    assert not isinstance(parse_number("42"), float)


def test_json_dump_keeps_decimal_precision_and_no_exponent():
    document = {"a": Decimal("123.4500"), "b": Decimal("1E+20"), "c": 1e22,
                "d": 9007199254740993, "e": "1.500", "f": None, "g": True}
    text = dumps_json(document)
    assert "e+" not in text and "E+" not in text
    assert "123.4500" in text
    assert "100000000000000000000" in text
    assert "10000000000000000000000" in text
    assert '"1.500"' in text and "9007199254740993" in text
    assert loads_json(text)["a"] == Decimal("123.4500")


def test_json_load_uses_decimal():
    data = loads_json('{"x": 0.1000000000000000055511151231257827, "y": 1.500}')
    assert data["x"] == Decimal("0.1000000000000000055511151231257827")
    assert isinstance(data["y"], Decimal)


def test_json_escapes_strings_safely():
    text = dumps_json({"k": '引号"与\n换行'})
    assert loads_json(text)["k"] == '引号"与\n换行'


def test_significant_digits():
    assert significant_digits(Decimal("1.500")) == 2
    assert significant_digits(Decimal("123456789012345678")) == 18
    assert significant_digits(Decimal("0")) == 1


def test_excel_cell_keeps_long_values_as_text():
    assert excel_cell_value(Decimal("1.5")) == 1.5
    assert excel_cell_value(Decimal("12345678901234567890.12345")) == "12345678901234567890.12345"
    assert excel_cell_value(Decimal("NaN")) == "NaN"


def test_excel_cell_keeps_trailing_zeros_as_text():
    """Excel 会把 1.500 显示成 1.5,因此必须按文本写出才能保真。"""
    assert excel_cell_value(Decimal("1.500")) == "1.500"
    assert excel_cell_value(Decimal("100.00")) == "100.00"
    assert excel_cell_value(Decimal("0.10")) == "0.10"


def test_fidelity_report_detects_truncation_and_text_change():
    report = FidelityReport()
    assert report.ok
    assert report.compare("amount", Decimal("1.500"), Decimal("1.500"), "1")
    assert report.ok and report.checked_values == 1

    assert not report.compare("amount", Decimal("1.500"), Decimal("1.5"), "2")   # 尾零丢失
    assert not report.ok
    assert not report.compare("amount", Decimal("123.456"), Decimal("123.46"), "3")  # 舍入
    assert any("文本表示变化" in line or "数值不等" in line
               for line in report.summary_lines())
    assert len(report.mismatches) == 2


def test_fidelity_report_flags_scientific_notation_output():
    report = FidelityReport()
    assert not report.compare("big", Decimal("1E+20"), "1E+20", "1")
    assert not report.ok


def test_numeric_recorder_multiset_compare():
    recorder = NumericRecorder(limit=10)
    for value in ("1.500", "2.25", "1.500", 7):
        recorder.record("amount", value)
    assert recorder.active

    good = recorder.compare_with({"amount": ["1.500", "1.500", "2.25", "7"]})
    assert good.ok and good.checked_values == 4

    bad = recorder.compare_with({"amount": ["1.5", "2.25", "7"]})
    assert not bad.ok
    assert bad.mismatches[0].value_equal and not bad.mismatches[0].text_equal

    lost = recorder.compare_with({"amount": ["1.500"]})
    assert not lost.ok and not lost.mismatches[0].value_equal


def test_numeric_recorder_ignores_non_numeric():
    recorder = NumericRecorder()
    recorder.record("name", "张三")
    recorder.record("code", "007")
    assert not recorder.active


def test_recorder_limit_bounds_memory():
    recorder = NumericRecorder(limit=3)
    for index in range(100):
        recorder.record("n", index)
    assert len(recorder.samples["n"]) == 3


def test_recorder_samples_the_same_row_window_as_readback():
    """采样窗口按“行”计,不能因为行内有补零编号就整体后移。

    否则回读(只读前 limit 行)会把窗口之后的值报成“(缺失)”,出现假报错。
    """
    recorder = NumericRecorder(limit=4)
    header = ["IdCard"]
    for value in ("018956", "526018", "603082", "865797", "123456"):
        recorder.record_row(header, [value])
    assert recorder.samples["IdCard"] == ["526018", "603082", "865797"]

    report = recorder.compare_with({"IdCard": ["018956", "526018", "603082", "865797"]})
    assert report.ok, report.summary_lines()
