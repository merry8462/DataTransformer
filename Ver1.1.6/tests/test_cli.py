# -*- coding: utf-8 -*-
"""命令行向导测试:以管道输入驱动完整交互流程(非 TTY 回退模式)。"""

import os
import subprocess
import sys
from pathlib import Path

import pytest

from dt_numeric import loads_json

ROOT = Path(__file__).resolve().parent.parent
CLI = ROOT / "dt_cli.py"


def run_cli(answers, *args, timeout=120):
    env = dict(os.environ, PYTHONIOENCODING="utf-8", DT_PLAIN="1", PYTHONUTF8="1")
    return subprocess.run(
        [sys.executable, str(CLI), *args],
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
