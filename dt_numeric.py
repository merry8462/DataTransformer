# -*- coding: utf-8 -*-
"""
高精度数值保真层(dt_numeric)
============================
设计目标(对应需求中的「长浮点数保真 + 禁止科学计数法」):

* **禁止二进制 float 作为中转**:外部字面量一律解析为 ``decimal.Decimal`` 或
  Python 任意精度 ``int``;确实只能拿到 float 时(Excel 单元格、FLOAT/DOUBLE 列),
  通过 ``Decimal(str(value))`` 取十进制字面量,不做任何额外舍入。
* **定点输出**:所有输出(xlsx / csv / json / 数据库)统一走 :func:`format_number`,
  以定点十进制呈现,永不出现 ``e`` / ``E`` 科学计数法。
* **不丢尾零、不截断有效数字**:``Decimal('1.500')`` 输出仍是 ``1.500``。
* **前导零规则**:``007`` / ``0012`` 这类补零数字视为编号而非数值,保持文本原样
  (可通过 :func:`parse_number` 的 ``keep_padded`` 参数改变)。
* **符号规则**:负号始终保留;冗余正号 ``+1.5`` 在数值等值前提下按十进制规范输出
  (即 ``1.5``),不做无意义保留。
* **数值等值 + 文本等值双重校验**:转换写入后回读输出,按
  :class:`FidelityReport` 输出差异报告,发现精度丢失 / 截断 / 科学计数法时阻断流程。
"""

from __future__ import annotations

import json
import re
from collections import Counter
from dataclasses import dataclass, field
from decimal import Decimal, InvalidOperation, localcontext

# 带符号定点/科学计数法字面量(用于判定“这是不是一个数字”)
NUMBER_RE = re.compile(r"^[+-]?(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?$")
INTEGER_RE = re.compile(r"^[+-]?\d+$")
PADDED_RE = re.compile(r"^[+-]?0\d+$")           # 007 / 0012:补零编号
SCIENTIFIC_RE = re.compile(r"[eE][+-]?\d+$")

DEFAULT_SAMPLE_LIMIT = 200       # 每个字段最多保留多少个样本用于回读比对
EXCEL_SIGNIFICANT_LIMIT = 15     # Excel 双精度可精确表示的有效数字位数


class FidelityError(RuntimeError):
    """数值保真校验失败(精度丢失 / 文本表示变化 / 出现科学计数法)。"""

    def __init__(self, report: "FidelityReport"):
        self.report = report
        super().__init__("数值保真校验未通过:\n" + "\n".join(report.summary_lines()))


# ---------------------------------------------------------------------------
# 解析:文本 → int / Decimal(绝不产生 float)
# ---------------------------------------------------------------------------
def is_number_literal(text) -> bool:
    """判断文本是否是可以安全数值化的十进制字面量(补零编号不算)。"""
    if not isinstance(text, str):
        return False
    value = text.strip()
    if not value or not NUMBER_RE.match(value):
        return False
    return not PADDED_RE.match(value)


def parse_number(text, keep_padded: bool = True):
    """把十进制文本解析成 ``int``(纯整数)或 ``Decimal``(小数 / 科学计数法)。

    ``keep_padded=True``(默认)时,``007`` 这类补零字面量保持字符串,避免丢失前导零。
    无法解析或按规则应保留文本时返回 ``None``。
    """
    if not isinstance(text, str):
        return None
    value = text.strip()
    if not value or not NUMBER_RE.match(value):
        return None
    if keep_padded and PADDED_RE.match(value):
        return None
    if INTEGER_RE.match(value):
        return int(value)
    try:
        return Decimal(value)
    except InvalidOperation:                     # pragma: no cover - 正则已挡掉
        return None


def to_decimal(value, keep_padded: bool = True):
    """任意数值 → ``Decimal``(float 经 ``str`` 中转,不引入二进制误差)。"""
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, Decimal):
        return value
    if isinstance(value, int):
        return Decimal(value)
    if isinstance(value, float):
        return Decimal(str(value))
    if isinstance(value, str):
        parsed = parse_number(value, keep_padded=keep_padded)
        if parsed is None:
            return None
        return Decimal(parsed) if isinstance(parsed, int) else parsed
    return None


# ---------------------------------------------------------------------------
# 格式化:任意数值 → 定点十进制文本(禁止科学计数法)
# ---------------------------------------------------------------------------
def format_decimal(value: Decimal) -> str:
    """``Decimal`` → 定点文本;保留尾零,永不输出科学计数法。"""
    if not value.is_finite():                    # NaN / Infinity 原样输出
        return str(value)
    with localcontext() as ctx:
        ctx.prec = max(len(value.as_tuple().digits), 28) + 2   # 禁止上下文精度截断
        return format(value, "f")


def format_number(value, keep_padded: bool = True) -> str:
    """任意 Python 值 → 定点十进制文本;不是数值时返回 ``str(value)``。"""
    if isinstance(value, bool) or value is None:
        return "" if value is None else str(value)
    if isinstance(value, int):
        return str(value)
    if isinstance(value, Decimal):
        return format_decimal(value)
    if isinstance(value, float):
        return format_decimal(Decimal(str(value)))
    if isinstance(value, str):
        parsed = parse_number(value, keep_padded=keep_padded)
        if parsed is None:
            return value
        return format_decimal(Decimal(parsed) if isinstance(parsed, int) else parsed)
    return str(value)


def has_scientific_notation(text) -> bool:
    """文本末尾是否为科学计数法指数部分(用于校验输出)。"""
    return bool(isinstance(text, str) and SCIENTIFIC_RE.search(text.strip()))


def significant_digits(value: Decimal) -> int:
    """有效数字位数(忽略首尾零)。"""
    digits = value.as_tuple().digits
    if not digits:
        return 1
    stripped = 0
    for digit in reversed(digits):               # 去掉尾部零
        if digit:
            break
        stripped += 1
    return max(1, len(digits) - stripped)


def excel_cell_value(value: Decimal):
    """``Decimal`` → Excel 单元格值。

    Excel 内部是 IEEE-754 双精度:超过 15 位有效数字会被截断/舍入,尾零也会被
    Excel 直接吞掉(``1.500`` 变 ``1.5``),因此这类值按**文本**写出,保证与输入
    字面量完全一致;只有文本形式与数值形式完全等价时才写数值,便于 Excel 计算。
    """
    if not value.is_finite():
        return format_decimal(value)
    text = format_decimal(value)
    if significant_digits(value) > EXCEL_SIGNIFICANT_LIMIT:
        return text
    as_float = float(value)
    return as_float if format_decimal(Decimal(str(as_float))) == text else text


# ---------------------------------------------------------------------------
# JSON:定点数字序列化(标准库无法直接输出 Decimal 数字)
# ---------------------------------------------------------------------------
def dumps_json(document, indent: int = 4) -> str:
    """把 dict/list 结构序列化为 JSON 文本。

    与 ``json.dumps`` 的差异:``Decimal`` / ``float`` 一律以**定点数字**输出
    (``123.4500`` 保留尾零、``1E+20`` 展开为 ``100000000000000000000``),
    从而保证长浮点数在 JSON 中数值与文本都不失真。
    """
    return _dump(document, indent, 0)


def _dump(value, indent: int, level: int) -> str:
    pad = " " * (indent * (level + 1))
    end_pad = " " * (indent * level)

    if value is None:
        return "null"
    if value is True:
        return "true"
    if value is False:
        return "false"
    if isinstance(value, Decimal):
        return format_decimal(value)
    if isinstance(value, int):
        return str(value)
    if isinstance(value, float):
        return format_decimal(Decimal(str(value)))
    if isinstance(value, str):
        return json.dumps(value, ensure_ascii=False)
    if isinstance(value, (list, tuple)):
        if not value:
            return "[]"
        items = [f"{pad}{_dump(item, indent, level + 1)}" for item in value]
        return "[\n" + ",\n".join(items) + f"\n{end_pad}]"
    if isinstance(value, dict):
        if not value:
            return "{}"
        items = [f"{pad}{json.dumps(str(key), ensure_ascii=False)}: "
                 f"{_dump(item, indent, level + 1)}" for key, item in value.items()]
        return "{\n" + ",\n".join(items) + f"\n{end_pad}}}"
    return json.dumps(str(value), ensure_ascii=False)


def loads_json(text):
    """读取 JSON:小数一律解析为 ``Decimal``,整数保持任意精度 ``int``。"""
    return json.loads(text, parse_float=Decimal, parse_int=int)


def read_json_file(path):
    """读取 JSON 文件(``parse_float=Decimal``,长浮点数不失真)。"""
    with open(path, "r", encoding="utf-8") as fp:
        return json.load(fp, parse_float=Decimal, parse_int=int)


# ---------------------------------------------------------------------------
# 保真校验:记录原始字面量 → 回读输出 → 双重比对
# ---------------------------------------------------------------------------
@dataclass
class NumericSample:
    """一条数值比对记录。"""

    field: str
    original: str
    output: str
    value_equal: bool = True
    text_equal: bool = True
    reference: str = ""

    @property
    def ok(self) -> bool:
        return self.value_equal and self.text_equal

    def describe(self) -> str:
        if self.ok:
            return f"{self.field}={self.original}(一致)"
        if not self.value_equal:
            reason = "数值不等"
        else:
            reason = "文本表示变化"
        return (f"{self.field} 行{self.reference}: 输入 {self.original!r} → "
                f"输出 {self.output!r}({reason})")


@dataclass
class FidelityReport:
    """数值保真校验报告(数值等值 + 文本等值双维度)。"""

    checked_fields: set = field(default_factory=set)
    checked_values: int = 0
    mismatches: list = field(default_factory=list)
    notes: list = field(default_factory=list)

    @property
    def ok(self) -> bool:
        return not self.mismatches

    def add_note(self, note: str) -> None:
        self.notes.append(str(note))

    def compare(self, field: str, original, output, reference: str = "") -> bool:
        """比对单个字段的输入 / 输出表示,返回是否完全一致。

        检查三个维度:输出原始文本不得出现科学计数法、定点文本必须一致、数值必须等值。
        """
        self.checked_fields.add(field)
        self.checked_values += 1
        raw_output = output if isinstance(output, str) else format_number(output)
        scientific = has_scientific_notation(raw_output)
        left, right = format_number(original), format_number(output)
        if left == right and not scientific:
            return True
        left_value, right_value = to_decimal(left), to_decimal(right)
        sample = NumericSample(
            field=field, original=left, output=raw_output, reference=str(reference),
            value_equal=bool(left_value is not None and left_value == right_value),
            text_equal=left == right and not scientific)
        if len(self.mismatches) < 50:            # 报告只保留前 50 条,避免刷屏
            self.mismatches.append(sample)
        return False

    def merge(self, other: "FidelityReport") -> None:
        self.checked_fields |= other.checked_fields
        self.checked_values += other.checked_values
        self.mismatches.extend(other.mismatches[:max(0, 50 - len(self.mismatches))])
        self.notes.extend(other.notes)

    def summary_lines(self) -> list:
        head = (f"数值保真校验:{'通过' if self.ok else '未通过'}  "
                f"(字段 {len(self.checked_fields)} 个 / 比对 {self.checked_values} 个值)")
        lines = [head]
        if self.mismatches:
            lines.append(f"差异条目 {len(self.mismatches)} 条,样例:")
            lines.extend(f"  - {sample.describe()}" for sample in self.mismatches[:10])
        lines.extend(f"  · {note}" for note in self.notes)
        return lines


class NumericRecorder:
    """按字段记录输入侧的数值字面量,供写入后回读比对。

    采样窗口固定为**前 limit 行**,与回读侧的 ``islice(rows, limit)`` 对齐:
    若改成"前 limit 个数值",行内被跳过的值(如 ``007`` 这类补零编号)会让采样窗口
    整体后移,回读时就会把窗口之后的值误报成"(缺失)"。
    """

    def __init__(self, limit: int = DEFAULT_SAMPLE_LIMIT):
        self.limit = limit
        self.samples = {}                        # {字段: [定点文本, ...]}
        self.fields = []                         # 记录顺序
        self._row = 0                            # 已采样的行数

    def watch(self, field: str) -> None:
        if field not in self.samples:
            self.samples[field] = []
            self.fields.append(field)

    def record_row(self, header, row) -> None:
        """按行采样:只记录前 ``limit`` 行,行内不符合数值字面量的值自动跳过。"""
        if self._row >= self.limit:
            return
        self._row += 1
        for name, value in zip(header, row):
            self.record(str(name), value)

    def record(self, field: str, value) -> None:
        """记录一个数值样本;非数值或样本已满则忽略(不建立空桶,避免无谓开销)。"""
        bucket = self.samples.get(field)
        if bucket is not None and len(bucket) >= self.limit:
            return
        if isinstance(value, str):
            if not is_number_literal(value):
                return
        elif to_decimal(value) is None:
            return
        if bucket is None:
            bucket = self.samples[field] = []
            self.fields.append(field)
        bucket.append(format_number(value))

    @property
    def active(self) -> bool:
        return any(self.samples.values())

    def compare_with(self, rows, report: FidelityReport | None = None) -> FidelityReport:
        """用回读到的行数据与记录值做多重集比对。"""
        report = report or FidelityReport()
        for name, originals in self.samples.items():
            if not originals:
                continue
            report.checked_fields.add(name)
            report.checked_values += len(originals)
            left = Counter(originals)
            right = Counter(format_number(value) for value in rows.get(name, [])
                            if value is not None)
            for text, count in left.items():
                missing = count - right.get(text, 0)
                if missing <= 0:
                    continue
                # 该文本在输出中缺失:区分「整条样本丢失」与「数值相同但文本不同」
                decimal_value = to_decimal(text)
                candidates = [other for other in right
                              if other != text and decimal_value is not None
                              and to_decimal(other) == decimal_value]
                report.mismatches.append(NumericSample(
                    field=name, original=text,
                    output=candidates[0] if candidates else "(缺失)",
                    value_equal=bool(candidates),
                    text_equal=False))
        return report


def collect_numeric_fields(header, sample_rows) -> dict:
    """从样本行中挑出“包含数值”的字段,返回 {字段序号: 是否长精度}。"""
    fields = {}
    for index, name in enumerate(header):
        for row in sample_rows:
            if index >= len(row):
                continue
            value = row[index]
            if isinstance(value, (int, Decimal)) and not isinstance(value, bool):
                long_precision = (isinstance(value, Decimal)
                                  and significant_digits(value) > EXCEL_SIGNIFICANT_LIMIT)
                fields[index] = fields.get(index, False) or long_precision
                break
            if isinstance(value, float) or (isinstance(value, str) and is_number_literal(value)):
                fields[index] = fields.get(index, False)
                break
    return fields
