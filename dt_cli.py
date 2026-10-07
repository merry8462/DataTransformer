# -*- coding: utf-8 -*-
"""
DataTransformer 终端交互式向导(questionary 动态交互版)
=====================================================
* 使用 **questionary** 动态选择 Input/Output 格式、数据库类型、连接参数、
  表/集合/工作表、字段映射与导入模式;无需图形化界面;
* 兼容 Ubuntu / Debian / Fedora / Arch / CentOS 等主流发行版;
* 非交互环境(管道、CI、无 TTY)自动回退到纯文本问答,便于脚本化调用与测试。

启动方式(推荐):
    ./data_transformer.sh                 # 交互式向导
    ./data_transformer.sh --demo          # 预填 Demo/Example.txt 的 SSH + PostgreSQL 示例
    python3 dt_cli.py --selftest          # 无界面自检

交互流程:
    ① 选择输入格式(MySQL / PostgreSQL / MongoDB / Xlsx / Json / Csv)
    ② 连接参数(主机 / SSH 隧道 / 用户名 / 密码 / 库 / 端口 / URI)
    ③ 层级勾选:表 / 集合 → 字段
    ④ 选择输出格式、写入模式、字段映射、导出路径
    ⑤ 执行转换 → 数值保真校验报告 → 继续 / 中止
"""

from __future__ import annotations

import contextlib
import getpass
import os
import sys
import unicodedata
from dataclasses import dataclass
from pathlib import Path

try:
    from dt_core import (
        ADAPTERS, APP_TITLE, DB_KINDS, DEFAULT_BATCH, DELIM_LABELS, ENC_LABELS,
        MODE_APPEND, MODE_CREATE, MODE_REPLACE, VERSION,
        DBConfig, FieldMapping, Settings, SourceSpec, SSHConfig, SSHTunnel, TargetSpec,
        connect_with_ssh, config_file_path, convert, csv_output_paths,
        describe_db_error, describe_ssh_error, json_top_keys, load_csv_section,
        load_json_stream, open_reader, read_csv_header, run_self_test, safe_filename,
        safe_sheet_title, scan_csv_sections, to_int, workbook_sheet_headers,
    )
    from dt_mongo import MongoConfig, pymongo_available
    from dt_numeric import FidelityError
except ImportError as exc:  # pragma: no cover - 依赖缺失时给出明确指引
    print(f"[错误] 缺少核心模块({exc})。\n"
          f"请确认 dt_cli.py 与 dt_core.py / dt_numeric.py / dt_mongo.py 位于同一目录,"
          f"或改用 ./data_transformer.sh 启动。")
    raise SystemExit(1)

try:
    import questionary
except Exception:                                     # pragma: no cover
    questionary = None

try:                                                  # 与 prompt_toolkit 渲染口径一致
    from prompt_toolkit.layout.processors import Processor, Transformation
    from prompt_toolkit.utils import get_cwidth as display_width
except Exception:                                     # pragma: no cover
    Processor = Transformation = None

    def display_width(text: str) -> int:
        """终端显示宽度:全角字符按 2 列计。"""
        return sum(2 if unicodedata.east_asian_width(ch) in "WF" else 1
                   for ch in text)


TAB_STOP = 8                                          # 制表位宽度(与终端一致)

# 勾选列表外观:已勾选 = 实心圆 ● + 高亮色,未勾选 = 空心圆 ○ + 终端默认色
CHECKBOX_STYLE = None
if questionary is not None:                            # pragma: no cover - 终端交互
    with contextlib.suppress(Exception):
        # 圆点符号在 questionary 2.0 / 2.1 之间不同,这里固定为 ● / ○
        from questionary.prompts import common as _questionary_common
        _questionary_common.INDICATOR_SELECTED = "●"
        _questionary_common.INDICATOR_UNSELECTED = "○"
        CHECKBOX_STYLE = questionary.Style([
            ("selected", "fg:#00d7af bold"),           # 已勾选:高亮显示
            ("highlighted", "bold"),                   # 光标所在行(未勾选时仍用默认配色)
            ("pointer", "fg:#ff9d00 bold"),
        ])


def q_message(message: str) -> str:
    """把问题文本补齐到下一个制表位:回显为「参数名 <制表位> 参数值」。

    prompt_toolkit 会把真正的制表符转义成 ``^I`` 显示(实测),因此这里用空格补齐到
    制表位,终端里的对齐效果与制表符分隔一致。
    """
    text = message.rstrip()
    width = display_width(text)
    stop = ((2 + width) // TAB_STOP + 1) * TAB_STOP     # "? " + 问题 之后的下一个制表位
    return text + " " * (stop - 3 - width)              # questionary 在问题后固定补一个空格


class ParamBrackets(Processor if Processor is not None else object):
    """给参数值套上 ``[ ]`` 边框:只影响显示,不进入取值,也不改变光标位置。"""

    def __init__(self, left="[", right="]", style="class:answer"):
        self.left, self.right, self.style = left, right, style

    def apply_transformation(self, ti):                 # pragma: no cover - 终端交互
        if ti.lineno != 0:                              # 参数行都是单行输入
            return Transformation(ti.fragments)
        fragments = ([(self.style, self.left)] + list(ti.fragments)
                     + [(self.style, self.right)])
        shift = len(self.left)                          # 方括号为半角,占 1 列
        return Transformation(
            fragments,
            source_to_display=lambda position: position + shift,
            display_to_source=lambda position: position - shift)


PARAM_INPUT = {"input_processors": [ParamBrackets()]}   # 参数行的输入处理器

PARAM_LABEL_WIDTH = 18                                  # 参数名(含方括号)的显示宽度


def param_message(label: str) -> str:
    """参数行提示:``[参数名]`` 左对齐补齐到固定宽度,使各行的值左括号竖向对齐。"""
    text = f"[{label}]"
    return text + " " * max(1, PARAM_LABEL_WIDTH - display_width(text))

MODE_CHOICES = (
    ("MySQL", "MySQL(SQL 数据库)"),
    ("PostgreSQL", "PostgreSQL(SQL 数据库)"),
    ("MongoDB", "MongoDB(文档数据库)"),
    ("Excel", "Xlsx(Excel 工作簿)"),
    ("JSON", "Json"),
    ("CSV", "Csv"),
)
MODE_TITLES = dict(MODE_CHOICES)
SQL_MODES = ("MySQL", "PostgreSQL")
DB_MODES = ("MySQL", "PostgreSQL", "MongoDB")

# Demo/Example.txt 中的示例环境(Ubuntu 26.04 + PostgreSQL + SSH)
DEMO_VALUES = {
    "ssh": {"host": "172.24.208.28", "port": 22, "user": "cml", "password": "123456"},
    "PostgreSQL": {"host": "127.0.0.1", "port": 5432, "user": "Ubuntu2604",
                   "password": "123456", "database": "virtual_data",
                   "table": "VirtualProfile"},
}

COLOR = (sys.stdout.isatty() and not os.environ.get("NO_COLOR")
         and os.environ.get("TERM") != "dumb")


def paint(text, code) -> str:
    return f"\033[{code}m{text}\033[0m" if COLOR else str(text)


def title(text) -> str:
    return paint(text, "1;36")


def ok(text) -> str:
    return paint(text, "1;32")


def warn(text) -> str:
    return paint(text, "1;33")


def section(text) -> str:
    return f"\n{paint('─' * 4 + ' ' + text + ' ' + '─' * max(4, 56 - len(text)), '1;34')}"


class Aborted(Exception):
    """用户主动放弃当前步骤。"""


# ---------------------------------------------------------------------------
# 交互层:questionary(动态选择)+ 纯文本回退
# ---------------------------------------------------------------------------
class UI:
    """统一问答接口:TTY 下走 questionary,非 TTY 下走纯文本(便于管道与测试)。"""

    def __init__(self, dynamic=None):
        self.dynamic = (questionary is not None and sys.stdin.isatty()
                        and not os.environ.get("DT_PLAIN")
                        if dynamic is None else bool(dynamic))

    def select(self, message, choices, default=None, param=False):
        options = list(choices)
        if self.dynamic:
            answer = questionary.select(message if param else q_message(message),
                                        choices=options, default=default).ask()
            if answer is None:
                raise KeyboardInterrupt
            return answer
        print(f"\n  {message}:")
        for index, label in enumerate(options, start=1):
            print(f"  {index}) {label}")
        preset = options.index(default) + 1 if default in options else 1
        while True:
            raw = input(f"  请选择 [1-{len(options)}](默认 {preset}): ").strip() or str(preset)
            if raw.isdigit() and 1 <= int(raw) <= len(options):
                return options[int(raw) - 1]
            print(warn(f"  ⚠ 请输入 1-{len(options)} 之间的序号"))

    def checkbox(self, message, choices, checked=None):
        options = list(choices)
        if self.dynamic:
            answer = questionary.checkbox(
                q_message(message),
                choices=[questionary.Choice(title=item,
                                            checked=bool(checked and item in checked))
                         for item in options],
                style=CHECKBOX_STYLE).ask()
            if answer is None:
                raise KeyboardInterrupt
            return answer
        print(f"\n  {message}:")
        width = len(str(len(options)))
        for index, item in enumerate(options, start=1):
            print(f"  {index:>{width}}) {item}")
        while True:
            picked = parse_selection(ask("  输入序号(如 1,3-5;回车=全部)", default="all"),
                                     len(options))
            if picked:
                return [options[index] for index in picked]
            print(warn("  ⚠ 没有匹配的序号,请重新输入"))

    def text(self, message, default="", validate=None, allow_empty=False, param=False):
        if self.dynamic:
            checker = validate or (None if allow_empty else
                                   (lambda value: True if value.strip() else "不能为空"))
            answer = questionary.text(message if param else q_message(message),
                                      default=str(default or ""), validate=checker,
                                      **(PARAM_INPUT if param else {})).ask()
            if answer is None:
                raise KeyboardInterrupt
            return answer.strip() or str(default or "")
        hint = f" [{default}]" if default not in ("", None) else ""
        while True:
            raw = input(f"{message}{hint}: ").strip()
            if raw:
                if validate is not None and validate(raw) is not True:
                    print(warn(f"  ⚠ {validate(raw)}"))
                    continue
                return raw
            if default not in ("", None):
                return str(default)
            if allow_empty:
                return ""
            print(warn("  ⚠ 不能为空,请重新输入"))

    def password(self, message, default="", param=False):
        if self.dynamic:
            answer = questionary.password(message if param else q_message(message),
                                          **(PARAM_INPUT if param else {})).ask()
            if answer is None:
                raise KeyboardInterrupt
            return answer or default
        hint = "(回车沿用已保存的密码)" if default else ""
        try:
            value = (getpass.getpass(f"{message}{hint}: ") if sys.stdin.isatty()
                     else input(f"{message}{hint}: "))
        except Exception:
            value = input(f"{message}{hint}: ")
        return value or default or ""

    def confirm(self, message, default=False):
        if self.dynamic:
            answer = questionary.confirm(q_message(message), default=default).ask()
            if answer is None:
                raise KeyboardInterrupt
            return bool(answer)
        hint = "Y/n" if default else "y/N"
        while True:
            raw = input(f"{message} [{hint}]: ").strip().lower()
            if not raw:
                return default
            if raw in ("y", "yes", "是"):
                return True
            if raw in ("n", "no", "否"):
                return False
            print(warn("  ⚠ 请输入 y 或 n"))

    def path(self, message, default="", validate=None):
        if self.dynamic:
            answer = questionary.path(q_message(message), default=str(default or ""),
                                      validate=validate).ask()
            if answer is None:
                raise KeyboardInterrupt
            return os.path.expanduser(os.path.expandvars(answer.strip()))
        return os.path.expanduser(os.path.expandvars(
            self.text(message, default=default,
                      validate=None if validate is None else
                      (lambda value: validate(value)))))


UI_INSTANCE = UI()


def parse_selection(text, count) -> list:
    """解析序号选择:all / 回车(全部) / 1,3-5 / 2-4 等,返回去重后的 0 基序号。"""
    text = (text or "").strip().lower()
    if not text or text in ("all", "*", "a"):
        return list(range(count))
    picked = []
    for chunk in text.replace(",", " ").replace("，", " ").split():
        start, _, end = chunk.partition("-")
        if start.isdigit() and end.isdigit():
            picked.extend(range(int(start), int(end) + 1))
        elif start.isdigit():
            picked.append(int(start))
    return sorted({index - 1 for index in picked if 1 <= index <= count})


def ask(prompt, default="", allow_empty=False, param=False) -> str:
    """param=True 时按参数行排版:``[参数名]`` 定宽 + 带方括号的值。"""
    return UI_INSTANCE.text(param_message(prompt) if param else prompt,
                            default=default, allow_empty=allow_empty, param=param)


def ask_int(prompt, default, minimum=None, maximum=None, param=False) -> int:
    def validate(value):
        try:
            number = int(str(value).strip())
        except ValueError:
            return "Please enter an integer" if param else "请输入整数"
        if (minimum is not None and number < minimum) or (maximum is not None and number > maximum):
            return (f"Please enter an integer between {minimum} and {maximum}" if param
                    else f"请输入 {minimum} ~ {maximum} 之间的整数")
        return True
    return int(UI_INSTANCE.text(param_message(prompt) if param else prompt,
                                default=default, validate=validate, param=param))


def ask_yes_no(prompt, default=False) -> bool:
    return UI_INSTANCE.confirm(prompt, default=default)


def ask_secret(prompt, default="", param=False) -> str:
    return UI_INSTANCE.password(param_message(prompt) if param else prompt,
                                default=default, param=param)


def ask_param_switch(label, default=False) -> bool:
    """参数行中的开关项(如 SSH 隧道):选项与答案同样带方括号,与其它参数行对齐。"""
    return choose_option(param_message(label), [("[No]", False), ("[Yes]", True)],
                         2 if default else 1, param=True)


def choose_option(prompt, options, default=1, param=False):
    """options = [(显示文本, 返回值), ...],返回被选项的返回值。"""
    labels = [label for label, _ in options]
    chosen = UI_INSTANCE.select(prompt, labels, default=labels[default - 1], param=param)
    return dict(options)[chosen]


def choose_mapping(prompt, mapping, default=1):
    """从 {显示名: 取值} 映射中选一项,返回对应的取值。"""
    names = list(mapping)
    chosen = UI_INSTANCE.select(prompt, names, default=names[default - 1])
    return mapping[chosen]


def choose_mode(prompt: str, default: int = 1) -> str:
    """选择输入 / 输出格式,返回内部标识(MySQL / PostgreSQL / MongoDB / Excel / JSON / CSV)。"""
    return choose_option(prompt, [(label, key) for key, label in MODE_CHOICES], default)


def pick_many(prompt, items, checked=None) -> list:
    """层级勾选:多选表 / 集合 / 字段(questionary.checkbox 或序号选择)。"""
    if not items:
        return []
    return UI_INSTANCE.checkbox(prompt, items, checked=checked)


def ask_existing_path(prompt, default="") -> str:
    return UI_INSTANCE.path(
        prompt, default=default,
        validate=lambda value: True if os.path.isfile(value) else "文件不存在")


def ask_output_dir(prompt, default) -> str:
    path = UI_INSTANCE.path(prompt, default=default)
    if os.path.isdir(path):
        return path
    if ask_yes_no(f"  目录不存在,是否创建 {path}", default=True):
        os.makedirs(path, exist_ok=True)
        return path
    raise Aborted("已取消转换(未选择输出目录)")


def ask_save_path(prompt, default, ext="") -> tuple:
    """询问输出文件路径,返回 (路径, 写入模式)。"""
    while True:
        path = UI_INSTANCE.path(prompt, default=default)
        if ext and not path.lower().endswith(ext):
            path += ext
        if not os.path.exists(path):
            return path, "replace"
        print(warn(f"  ⚠ 同名文件已存在: {path}"))
        mode = choose_option("请选择处理方式", [
            ("覆盖整个文件", "replace"),
            ("合并写入(覆盖同名 Sheet / JSON 键;CSV 追加)", "merge"),
            ("取消本次转换", "cancel")])
        if mode == "cancel":
            raise Aborted("已取消转换(存在同名文件)")
        return path, mode


def log(message, level="info") -> None:
    palette = {"ok": "1;32", "warn": "1;33", "err": "1;31", "head": "1;36", "db": "1;35"}
    print(f"  {paint(message, palette[level])}" if level in palette else f"  {message}")


def load_settings() -> Settings:
    try:
        return Settings.load()
    except Exception as exc:
        log(f"读取记忆配置失败(使用空白配置): {exc}", "warn")
        return Settings()


# ---------------------------------------------------------------------------
# 数据库会话(MySQL / PostgreSQL / MongoDB,支持 SSH 隧道)
# ---------------------------------------------------------------------------
@dataclass
class Database:
    """一个可用的数据库会话(含按需建立的 SSH 隧道)。"""

    name: str                                    # MySQL / PostgreSQL / MongoDB
    adapter: object
    conn: object
    cfg: object                                  # DBConfig 或 MongoConfig(用户填写的地址)
    connect_cfg: object | None = None            # 实际连接用的地址(SSH 时指向隧道端口)
    ssh: SSHConfig | None = None
    tunnel: object = None

    @property
    def kind(self) -> str:
        return "Mongo" if self.name == "MongoDB" else "SQL"

    def close(self) -> None:
        if self.conn is not None:
            if self.kind == "Mongo":
                if hasattr(self.adapter, "close"):
                    with contextlib.suppress(Exception):
                        self.adapter.close(self.conn)
            else:
                with contextlib.suppress(Exception):
                    self.conn.close()
            self.conn = None
        if self.tunnel is not None:
            with contextlib.suppress(Exception):
                self.tunnel.stop()
            self.tunnel = None

    def effective_cfg(self):
        return self.connect_cfg or self.cfg

    def label(self) -> str:
        suffix = f" (经 SSH {self.ssh.host})" if (self.tunnel is not None) else ""
        return (f"{self.name} {getattr(self.cfg, 'database', '')}"
                f"@{getattr(self.cfg, 'host', '')}:{getattr(self.cfg, 'port', '')}{suffix}")


def ask_ssh_config(settings: Settings, use_demo: bool) -> SSHConfig:
    saved = SSHConfig.from_mapping(DEMO_VALUES["ssh"] if use_demo else settings.ssh)
    if not ask_param_switch("SSH Tunnel", default=bool(saved.host)):
        return SSHConfig(enabled=False)
    return SSHConfig(
        enabled=True,
        host=ask("SSH Host", default=saved.host, param=True),
        port=ask_int("SSH Port", default=saved.port, minimum=1, maximum=65535, param=True),
        user=ask("SSH User", default=saved.user, param=True),
        password=ask_secret("SSH Password", default=saved.password, param=True),
        key_file=ask("SSH Key File", default=saved.key_file, allow_empty=True, param=True),
        timeout=saved.timeout,
    )


def open_database(settings: Settings, mode: str, use_demo: bool = False) -> Database:
    """询问连接参数并连接(必要时先建立 SSH 隧道);mode ∈ MySQL/PostgreSQL/MongoDB。"""
    adapter = ADAPTERS[mode]
    profile = dict(DEMO_VALUES.get(mode) or {}) if use_demo else settings.profile(mode)
    print(section(f"连接 {mode}"))
    ssh_cfg = ask_ssh_config(settings, use_demo)
    if ssh_cfg.enabled:
        print("  提示:下面的「数据库主机/端口」填写【远程服务器上】数据库的地址(通常为 127.0.0.1)")

    if mode == "MongoDB":
        db_cfg = MongoConfig(
            host=ask("MongoDB Host", default=profile.get("host") or "127.0.0.1", param=True),
            port=ask_int("MongoDB Port", default=to_int(profile.get("port"), 27017),
                         minimum=1, maximum=65535, param=True),
            uri=ask("Connection URI", default=profile.get("uri", ""),
                    allow_empty=True, param=True),
            user=ask("Username", default=profile.get("user", ""), allow_empty=True, param=True),
            password=ask_secret("Password", default=profile.get("password", ""), param=True),
            database=ask("Database", default=profile.get("database"), param=True),
            auth_source=ask("authSource", default=profile.get("auth_source", ""),
                            allow_empty=True, param=True),
            timeout=ask_int("Timeout (s)", default=to_int(profile.get("timeout"), 10),
                            minimum=1, maximum=3600, param=True),
        )
    else:
        db_cfg = DBConfig(
            host=ask("Database Host", default=profile.get("host") or "127.0.0.1", param=True),
            port=ask_int("Database Port", default=to_int(profile.get("port"), adapter.DEFAULT_PORT),
                         minimum=1, maximum=65535, param=True),
            user=ask("Username", default=profile.get("user"), param=True),
            password=ask_secret("Password", default=profile.get("password"), param=True),
            database=ask("Database", default=profile.get("database"), param=True),
            timeout=ask_int("Timeout (s)", default=to_int(profile.get("timeout"), 10),
                            minimum=1, maximum=3600, param=True),
        )
    try:
        db_cfg = db_cfg.validated(adapter)
    except ValueError as exc:
        log(str(exc), "err")
        raise Aborted("连接参数不完整") from exc

    print("  正在连接 ...")
    try:
        conn, tunnel = connect_with_ssh(db_cfg, ssh_cfg, adapter, log=log)
    except Exception as exc:
        log(describe_db_error(exc, db_cfg, adapter), "err")
        if ssh_cfg.enabled:
            log(describe_ssh_error(exc, ssh_cfg), "err")
        raise Aborted("数据库连接失败") from exc
    log(f"已连接 {adapter.server_version(conn)}"
        + (f",SSH 隧道 {tunnel.describe()}" if tunnel is not None else ""), "ok")
    connect_cfg = (type(db_cfg)(**{**db_cfg.as_dict(), "host": SSHTunnel.BIND_HOST,
                                   "port": tunnel.local_port})
                   if tunnel is not None else db_cfg)
    return Database(name=mode, adapter=adapter, conn=conn, cfg=db_cfg,
                    connect_cfg=connect_cfg, ssh=ssh_cfg, tunnel=tunnel)


# ---------------------------------------------------------------------------
# ①② 选择输入源与字段
# ---------------------------------------------------------------------------
def pick_database_sources(db: Database) -> list:
    """第 1 层选表 / 集合,第 2 层选字段。"""
    names = db.adapter.list_tables(db.conn)
    label = "集合" if db.kind == "Mongo" else "数据表"
    if not names:
        raise Aborted(f"{db.name} 的 {getattr(db.cfg, 'database', '')} 中没有可用的{label}")
    sources = []
    for name in pick_many(f"选择{label}(层级第 1 层:空格选择 / a 全选 / 回车确认)", names):
        columns = db.adapter.list_columns(db.conn, name)
        picked = pick_many(f"{label}【{name}】字段(层级第 2 层)", columns, checked=columns)
        sources.append(SourceSpec(kind=db.kind, table=name, columns=picked))
    if not sources:
        raise Aborted("未选择任何数据源")
    return sources


def pick_file_sources(mode: str) -> list:
    """文件输入的层级勾选(Excel Sheet / JSON 一级键 / CSV 分节 → 字段)。"""
    print(section(f"选择 {mode} 输入文件"))
    path = ask_existing_path("  输入文件路径")
    delimiter, encoding = ",", "utf-8-sig"
    if mode == "CSV":
        delimiter = choose_mapping("CSV 分隔符", DELIM_LABELS)
        encoding = choose_mapping("CSV 编码", ENC_LABELS)

    try:
        levels = _input_levels(mode, path, delimiter, encoding)
    except Exception as exc:
        log(f"读取输入文件失败: {exc}", "err")
        raise Aborted("无法解析输入文件") from exc

    sources = []
    for name, columns in pick_levels(levels):
        sources.append(SourceSpec(kind=mode, path=path, table=name, columns=columns,
                                  delimiter=delimiter, encoding=encoding))
    return sources


def _input_levels(mode: str, path: str, delimiter: str, encoding: str) -> dict:
    """返回 {第一层名称: [字段...]}(CSV 无分节时用空名表示整份文件)。"""
    if mode == "Excel":
        return workbook_sheet_headers(path)
    if mode == "JSON":
        return {key: load_json_stream(path, key)[0] for key in json_top_keys(path)}
    sections = scan_csv_sections(path, delimiter, encoding)
    if sections:
        return {name: load_csv_section(path, name, delimiter, encoding)[0] for name in sections}
    return {"": read_csv_header(path, delimiter, encoding)}


def pick_levels(levels: dict) -> list:
    """两层勾选:先选第一层名称,再为每个名称选字段。"""
    if not levels:
        raise Aborted("输入文件中没有可用的数据表")
    names = list(levels)
    if len(names) == 1 and not names[0]:
        return [("", list(levels[""]))]
    chosen = pick_many("选择数据表(层级第 1 层)", names)
    return [(name, pick_many(f"表【{name}】字段(层级第 2 层)", levels[name],
                             checked=levels[name])) for name in chosen]


# ---------------------------------------------------------------------------
# ③④ 字段映射 / 输出目标 / 导出路径
# ---------------------------------------------------------------------------
def ask_field_mapping(header) -> FieldMapping:
    """字段自动映射(同名)+ 可选的手动覆盖,对应需求的字段映射规则。"""
    if not header:
        return FieldMapping()
    print("  字段映射预览(源字段 → 目标字段):")
    print("    " + "、".join(f"{name} → {name}" for name in header))
    if not ask_yes_no("  是否手动调整字段映射(默认保持同名)", default=False):
        return FieldMapping()
    pairs = {}
    for name in header:
        target = ask(f"    {name} →", default=name)
        if target and target != name:
            pairs[name] = target
    return FieldMapping.from_pairs(pairs)


def ask_name_mapping(names, label: str) -> dict:
    """可选地重命名表 / 集合 / Sheet(对应「Sheet 名 ↔ 表名」映射规则)。"""
    names = [name for name in names if name]
    if not names or not ask_yes_no(f"  是否重命名{label}名称(默认与源同名)", default=False):
        return {}
    mapping = {}
    for name in names:
        target = ask(f"    {name} →", default=name)
        if target and target != name:
            mapping[name] = target
    return mapping


def build_file_target(mode: str, in_mode: str, sources: list, base_name: str,
                      names=None) -> TargetSpec:
    """文件输出:询问导出路径(及同名冲突处理)。"""
    print(section(f"选择 {mode} 导出路径"))
    default_dir = Path.home() / "Documents"
    ext = {"Excel": ".xlsx", "JSON": ".json", "CSV": ".csv"}[mode]
    base = safe_filename(base_name or "export")
    if mode == "CSV" and (in_mode in DB_MODES or len(sources) > 1):
        # 多表导出 CSV:与图形界面一致,使用目录模式(每个表一个 .csv),避免互相覆盖
        target_dir = ask_output_dir("  输出目录(每个表一个 .csv 文件)",
                                    str(Path.cwd() / base))
        target = TargetSpec(kind="CSV", path=target_dir, directory=True, mode="replace",
                            names=dict(names or {}))
        # 同名检测必须按【重命名之后】的文件名,否则改名后的旧文件会被直接覆盖
        paths = csv_output_paths(target_dir, sources, target)
        if existing := [path for path in paths if os.path.exists(path)]:
            print(warn(f"  ⚠ 目录中已存在 {len(existing)} 个同名文件(如 {existing[0]})"))
            target.mode = choose_option("请选择处理方式", [
                ("覆盖这些文件", "replace"),
                ("合并写入(追加数据行)", "merge"),
                ("取消本次转换", "cancel")])
            if target.mode == "cancel":
                raise Aborted("已取消转换(存在同名文件)")
        target.delimiter = choose_mapping("CSV 分隔符", DELIM_LABELS)
        target.encoding = choose_mapping("CSV 编码", ENC_LABELS)
        return target

    path, mode_key = ask_save_path(f"  导出文件路径({ext})",
                                   str(default_dir / f"{base}{ext}"), ext)
    target = TargetSpec(kind=mode, path=path, sheet=safe_sheet_title(base), key=base,
                        mode=mode_key)
    if mode == "CSV":
        target.delimiter = choose_mapping("CSV 分隔符", DELIM_LABELS)
        target.encoding = choose_mapping("CSV 编码", ENC_LABELS)
    return target


def build_database_target(db: Database, sources: list, base_name: str, names=None) -> TargetSpec:
    """数据库输出:询问目标表 / 集合名、写入模式,并按需逐表确认覆盖 / 追加 / 跳过。"""
    print(section(f"选择 {db.name} 写入方式"))
    names = dict(names or {})
    label = "集合" if db.kind == "Mongo" else "表"
    multi = len(sources) > 1
    table = ""
    if multi:
        print(f"  已勾选多个数据源:将按源{label}同名创建,目标{label}名留空。")
    else:
        source_name = sources[0].table
        table = ask(f"  目标{label}名", default=names.get(source_name, source_name) or base_name)
    mode = choose_option("写入模式", [
        (MODE_CREATE, "create_if_missing"),
        (MODE_APPEND, "append"),
        (MODE_REPLACE, "replace")], 1)
    batch = ask_int("  每批行数", default=DEFAULT_BATCH, minimum=1, maximum=1000000)

    if mode == "replace" and not ask_yes_no(
            f"  将先删除同名{label}(若存在)再重建并写入,确定继续吗", default=False):
        raise Aborted("已取消转换")
    for src in sources:
        name = table or src.table
        if not name:
            raise Aborted(f"无法推断目标{label}名,请指定")
        try:
            adapter_name = db.adapter.quote_ident(name)
        except ValueError as exc:
            log(f"目标{label}名【{name}】不合法:{exc}", "err")
            raise Aborted(f"目标{label}名不合法") from exc
        if src.kind not in DB_KINDS and db.adapter.table_exists(db.conn, adapter_name):
            print(warn(f"  ⚠ 数据库已存在同名{label}【{name}】"))
            src.mode = choose_option(f"{label}【{name}】写入方式", [
                ("覆盖写入(删除原数据后重建)", "replace"),
                ("追加写入(保留原有数据)", "append"),
                ("跳过该数据源", "skip")], 2)
        else:
            src.mode = mode
    return TargetSpec(kind=db.kind, table="" if multi else table, mode=mode, batch=batch)


# ---------------------------------------------------------------------------
# 主流程
# ---------------------------------------------------------------------------
def output_base_name(sources: list, db: Database | None) -> str:
    """默认输出名:数据库输入用库名(目录名 = 库名),文件输入用表名 / 文件名。"""
    if db is not None:
        return (getattr(db.cfg, "database", "")
                or (sources[0].table.split(".")[-1] if sources else "") or "export")
    if sources:
        first = sources[0]
        return first.table or Path(first.path).stem or "export"
    return "export"


def spec_kind(mode: str) -> str:
    """输入 / 输出模式 → SourceSpec.kind(MongoDB 映射为 Mongo)。"""
    return "Mongo" if mode == "MongoDB" else ("SQL" if mode in SQL_MODES else mode)


def run_once(use_demo: bool = False) -> None:
    settings = load_settings()
    in_mode = choose_mode("① 选择输入格式:MySQL / PostgreSQL / MongoDB / Xlsx / Json / Csv")
    print(f"\n{title('输入格式:')} {MODE_TITLES[in_mode]}")

    in_db = None
    out_db = None
    try:
        if in_mode in DB_MODES:
            in_db = open_database(settings, in_mode, use_demo)
            print(section("选择数据表 / 集合与字段"))
            sources = pick_database_sources(in_db)
        else:
            sources = pick_file_sources(in_mode)

        out_mode = choose_mode("③ 选择输出格式:MySQL / PostgreSQL / MongoDB / Xlsx / Json / Csv")
        if spec_kind(out_mode) in DB_KINDS:
            # CSV 文本导入数据库时按类型解析(数字 / 高精度小数还原为原生类型)
            for src in sources:
                src.parse = src.kind == "CSV"
        elif in_mode == "CSV" and ask_yes_no(
                "  CSV 是否按类型解析(数值 → int/Decimal、日期 → 日期类型)", default=False):
            # 默认保持文本原样:CSV→文件不做类型推断,保证字面量逐字一致
            for src in sources:
                src.parse = src.kind == "CSV"
        base_name = output_base_name(sources, in_db)
        # 表 / 集合 / Sheet 重命名放在目标之前:CSV 目录模式的同名检测要用改名后的文件名
        names = ask_name_mapping([src.table for src in sources],
                                 "集合" if spec_kind(out_mode) == "Mongo" else "数据表")

        if out_mode in DB_MODES:
            if in_db is None:
                out_db = open_database(settings, out_mode, use_demo)
            elif out_mode == in_mode:
                out_db = in_db                       # 同类型数据库:同一服务器内做表拷贝
            else:
                log(f"输入为 {in_mode}、输出为 {out_mode},跨数据库类型请先用文件中转"
                    f"(例如 {in_mode} → CSV → {out_mode})。", "err")
                raise Aborted("暂不支持跨数据库类型直接拷贝")
            target = build_database_target(out_db, sources, base_name, names)
        else:
            target = build_file_target(out_mode, in_mode, sources, base_name, names)

        target.fields = ask_field_mapping(_preview_header(sources, in_db))
        target.names = names

        print(section("转换摘要"))
        print(f"  输入: {MODE_TITLES[in_mode]} × {len(sources)} 个数据源")
        print(f"  输出: {MODE_TITLES[out_mode]}")
        print(f"  目标: {target.path or target.table or '(按源名称同名创建)'}")
        if target.kind in DB_KINDS:
            print(f"  模式: {target.mode},每批 {target.batch} 行")
        if target.fields.pairs:
            print(f"  字段映射: {target.fields.describe()}")
        if not ask_yes_no("\n确认开始转换", default=True):
            raise Aborted("已取消转换")

        print(section("执行转换"))
        db = out_db or in_db
        try:
            result = convert(sources, target, conn=db.conn if db else None,
                             adapter=db.adapter if db else None,
                             cfg=db.effective_cfg() if db else None, log=log)
        except FidelityError as exc:
            # 数值保真校验未通过:输出差异报告,由用户决定继续还是中止(需求要求)
            log(str(exc), "err")
            if not ask_yes_no("数值保真校验未通过,是否仍然保留已写出的结果", default=False):
                raise Aborted("已按用户选择中止(输出已写出,请核对差异报告)") from exc
            log("已保留输出,但数值校验未通过,请核对上面的差异报告", "warn")
        else:
            print()
            log(f"✅ 转换完成: {result.summary()}", "ok")
            _report_fidelity(result)
        _save_profile(settings, in_db, out_db)
    finally:
        for db in (out_db, in_db):
            if db is not None:
                db.close()


def _preview_header(sources: list, db: "Database | None" = None) -> list:
    """取第一个数据源的表头用于字段映射预览(读取失败则跳过映射步骤)。

    数据库输入优先用已勾选的字段名:既不用为预览再开一次游标,也不会因为
    缺少连接而把整个字段映射步骤静默跳过。
    """
    if not sources:
        return []
    first = sources[0]
    if first.columns:
        return [str(name) for name in first.columns]
    try:
        header, _, closer = open_reader(first,
                                        conn=db.conn if db else None,
                                        adapter=db.adapter if db else None)
        closer()
        return list(header)
    except Exception:
        return []


def _report_fidelity(result) -> None:
    """打印数值保真校验摘要(需求要求可见的校验机制与差异报告)。"""
    if not result.fidelity.checked_values:
        return
    for line in result.fidelity.summary_lines():
        log(line, "ok" if result.fidelity.ok else "warn")


def _save_profile(settings: Settings, in_db, out_db) -> None:
    """把本次连接参数写入记忆配置(密码为明文,文件权限已收紧为 600)。"""
    databases = [db for db in (in_db, out_db) if db is not None]
    if not databases:
        return
    try:
        for db in databases:
            settings.save_profile(db.name, db.cfg, db.ssh)
        settings.save()
        log(f"已记住本次连接参数: {config_file_path()}", "info")
    except Exception as exc:
        log(f"保存记忆配置失败(不影响本次转换): {exc}", "warn")


def banner() -> None:
    print(title("=" * 64))
    print(title(f" {APP_TITLE}  v{VERSION}"))
    print(title("=" * 64))
    if not pymongo_available():
        print(warn(" 提示:未安装 pymongo,本次无法使用 MongoDB(pip install pymongo 可启用)"))


def print_help() -> None:
    print(__doc__.strip())
    print("\n参数:")
    print("  --demo        预填 Demo/Example.txt 的 SSH + PostgreSQL 示例参数")
    print("  --selftest    无界面自检(依赖 / 核心算法 / 文件互转 / 数值保真)")
    print("  --plain       强制使用纯文本问答(不调用 questionary)")
    print("  --version     显示版本号")
    print("  -h, --help    显示本帮助")


def main(argv=None) -> int:
    global UI_INSTANCE
    args = list(sys.argv[1:] if argv is None else argv)
    if any(arg in ("-h", "--help") for arg in args):
        print_help()
        return 0
    if "--version" in args:
        print(VERSION)
        return 0
    if "--selftest" in args:
        return 0 if run_self_test() else 1
    if "--plain" in args:
        UI_INSTANCE = UI(dynamic=False)

    use_demo = "--demo" in args
    banner()
    if use_demo:
        print(ok(" 已启用 --demo:选择 PostgreSQL 时连接参数将默认使用 Demo/Example.txt 中的示例"))
    try:
        while True:
            run_once(use_demo)
            if not ask_yes_no("\n是否继续执行下一次转换", default=False):
                break
    except Aborted as exc:
        log(str(exc), "warn")
    except KeyboardInterrupt:
        print("\n已中断。")
    except EOFError:
        print("\n输入结束,退出。")
    except Exception as exc:                       # noqa: BLE001 - 兜底提示,避免栈回溯吓到用户
        log(f"发生未预期的错误: {exc}", "err")
        if os.environ.get("DT_DEBUG"):
            raise
    print(ok("\n感谢使用 DataTransformer,再见!"))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
