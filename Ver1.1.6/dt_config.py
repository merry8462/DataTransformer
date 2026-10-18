# -*- coding: utf-8 -*-
"""
常量 / 连接配置 / 记忆配置(dt_config)
====================================
集中放置跨模块共用的常量、数据库连接配置(DBConfig / SSHConfig)、
SSH 隧道参数、记忆配置文件读写(Settings)以及少量通用工具函数。
"""

from __future__ import annotations

import contextlib
import json
import os
import re
from dataclasses import dataclass, field, replace
from pathlib import Path

# ---------------------------------------------------------------------------
# 常量
# ---------------------------------------------------------------------------
APP_TITLE = "SQL ↔ Excel / JSON / CSV 数据转换工具"
VERSION = "1.1.6"
SAMPLE_SIZE = 200                 # 建表时用于推断字段类型的采样行数
DEFAULT_BATCH = 1000              # 默认每批读写行数
EXCEL_MAX_CELL_CHARS = 32767      # Excel 单元格字符上限
SSH_DEFAULT_PORT = 22

INPUT_LABELS = {
    "SQL": "SQL 数据库",
    "Excel": "Excel 文件 (.xlsx)",
    "JSON": "JSON 文件 (.json)",
    "CSV": "CSV 文件 (.csv)",
}
OUTPUT_LABELS = {
    "Excel": "Excel 文件 (.xlsx)",
    "JSON": "JSON 文件 (.json)",
    "CSV": "CSV 文件 (.csv)",
    "SQL": "SQL 数据库",
}

MODE_APPEND = "追加到已有表 (append)"
MODE_CREATE = "不存在则创建,存在则追加 (create if missing)"
MODE_REPLACE = "删除重建 (replace)"
MODE_LABELS = [MODE_APPEND, MODE_CREATE, MODE_REPLACE]
MODE_KEYS = {MODE_APPEND: "append", MODE_CREATE: "create_if_missing",
             MODE_REPLACE: "replace"}      # 界面下拉标签 → 引擎写入模式(必须覆盖 MODE_LABELS)
MODES = ("replace", "create_if_missing", "append")
DB_KINDS = ("SQL", "Mongo")          # 需要数据库适配器的输入 / 输出类型
FILE_KINDS = ("Excel", "JSON", "CSV")

DELIM_LABELS = {"逗号 (,)": ",", "分号 (;)": ";", "制表符 (TAB)": "\t", "竖线 (|)": "|"}
ENC_LABELS = {
    "UTF-8 with BOM (Excel 友好)": "utf-8-sig",
    "UTF-8": "utf-8",
    "GBK (中文 Excel)": "gbk",
}

IDENT_RE = re.compile(r"^[^\W\d]\w*$")            # 标识符:字母/下划线开头,允许中文、数字、下划线
CSV_SECTION_RE = re.compile(r"^\[Sheet:(.+?)\]$")  # 多表 CSV 分节标记,如 [Sheet:departments]

log_noop = lambda *args, **kwargs: None  # noqa: E731 - 默认日志回调


# ---------------------------------------------------------------------------
# 通用工具
# ---------------------------------------------------------------------------
def config_file_path() -> Path:
    """连接配置记忆文件路径。

    Windows: %APPDATA%\\DataTransformer\\config.json
    Linux/macOS: $XDG_CONFIG_HOME/DataTransformer/config.json(默认 ~/.config)
    """
    if os.name == "nt":
        base = Path(os.environ.get("APPDATA") or Path.home())
    else:
        base = Path(os.environ.get("XDG_CONFIG_HOME") or (Path.home() / ".config"))
    return base / "DataTransformer" / "config.json"


def to_int(value, default=0) -> int:
    """宽容地把界面文本转成整数,失败返回 default。"""
    try:
        return int(str(value).strip())
    except (TypeError, ValueError):
        return default


def _close_quietly(obj) -> None:
    """尽力关闭对象(游标 / 连接 / 工作簿),忽略任何异常。"""
    with contextlib.suppress(Exception):
        obj.close()


# ---------------------------------------------------------------------------
# 数据库连接参数 / SSH 参数
# ---------------------------------------------------------------------------
@dataclass
class DBConfig:
    """数据库连接参数(MySQL / PostgreSQL)。port=0 表示使用适配器默认端口。"""

    host: str = "127.0.0.1"
    port: int = 0
    user: str = ""
    password: str = ""
    database: str = ""
    charset: str = "utf8mb4"
    timeout: int = 10

    @classmethod
    def from_mapping(cls, data=None) -> "DBConfig":
        values = data or {}
        return cls(
            host=str(values.get("host") or "").strip(),
            port=to_int(values.get("port"), 0),
            user=str(values.get("user") or "").strip(),
            password=str(values.get("password") or ""),
            database=str(values.get("database") or "").strip(),
            charset=str(values.get("charset") or "utf8mb4"),
            timeout=to_int(values.get("timeout"), 10) or 10,
        )

    def resolved_port(self, adapter) -> int:
        return self.port or int(getattr(adapter, "DEFAULT_PORT", 0) or 0)

    def validated(self, adapter) -> "DBConfig":
        """校验必填项与取值范围,返回补全端口 / 超时后的新配置。"""
        if missing := [label for label, value in
                       (("主机", self.host), ("用户名", self.user), ("数据库", self.database))
                       if not value]:
            raise ValueError("请填写" + "、".join(missing) + "后再连接")
        if not 1 <= self.resolved_port(adapter) <= 65535:
            raise ValueError(f"端口号无效:{self.port!r},请输入 1-65535 的整数")
        if not 1 <= self.timeout <= 3600:
            raise ValueError(f"超时秒数无效:{self.timeout!r},请输入 1-3600 的整数")
        return replace(self, port=self.resolved_port(adapter))

    def as_dict(self) -> dict:
        return {"host": self.host, "port": self.port, "user": self.user,
                "password": self.password, "database": self.database,
                "charset": self.charset, "timeout": self.timeout}


@dataclass
class SSHConfig:
    """SSH 跳板机 / 远程服务器参数(本地端口转发)。"""

    enabled: bool = False
    host: str = ""
    port: int = SSH_DEFAULT_PORT
    user: str = ""
    password: str = ""
    key_file: str = ""
    passphrase: str = ""
    timeout: int = 15

    @classmethod
    def from_mapping(cls, data=None) -> "SSHConfig":
        values = data or {}
        return cls(
            enabled=bool(values.get("enabled")),
            host=str(values.get("host") or "").strip(),
            port=to_int(values.get("port"), SSH_DEFAULT_PORT) or SSH_DEFAULT_PORT,
            user=str(values.get("user") or "").strip(),
            password=str(values.get("password") or ""),
            key_file=str(values.get("key_file") or "").strip(),
            passphrase=str(values.get("passphrase") or ""),
            timeout=to_int(values.get("timeout"), 15) or 15,
        )

    def as_dict(self) -> dict:
        return {"enabled": self.enabled, "host": self.host, "port": self.port,
                "user": self.user, "password": self.password,
                "key_file": self.key_file, "passphrase": self.passphrase,
                "timeout": self.timeout}

    def validated(self) -> "SSHConfig":
        if missing := [label for label, value in (("SSH 主机", self.host), ("SSH 用户名", self.user))
                       if not value]:
            raise ValueError("请填写" + "、".join(missing))
        if not 1 <= self.port <= 65535:
            raise ValueError(f"SSH 端口无效:{self.port!r},请输入 1-65535 的整数")
        if not any((self.password, self.key_file, os.environ.get("SSH_AUTH_SOCK"))):
            raise ValueError("请填写 SSH 密码或指定私钥文件(使用 ssh-agent 免密登录时请先启动 ssh-agent)")
        return self


# ---------------------------------------------------------------------------
# 记忆配置(config.json)
# ---------------------------------------------------------------------------
@dataclass
class Settings:
    """记忆的连接配置:密码 / 数据库名(兼容旧版)+ 各适配器档案 + SSH 参数。"""

    password: str = ""
    database: str = ""
    profiles: dict = field(default_factory=dict)   # {"MySQL": {...}, "PostgreSQL": {...}}
    ssh: dict = field(default_factory=dict)

    @classmethod
    def load(cls, path=None) -> "Settings":
        """读取配置;文件不存在返回默认值,内容损坏时抛异常由调用方提示。"""
        path = path or config_file_path()
        if not path.exists():
            return cls()
        data = json.loads(path.read_text(encoding="utf-8"))
        if not isinstance(data, dict):
            return cls()
        return cls(
            password=str(data.get("password") or ""),
            database=str(data.get("database") or ""),
            profiles={name: dict(values) for name, values in (data.get("profiles") or {}).items()
                      if isinstance(values, dict)},
            ssh=dict(data.get("ssh") or {}),
        )

    def save(self, path=None) -> None:
        path = path or config_file_path()
        path.parent.mkdir(parents=True, exist_ok=True)
        payload = json.dumps(
            {"password": self.password, "database": self.database,
             "profiles": self.profiles, "ssh": self.ssh},
            ensure_ascii=False, indent=2)
        temp = path.with_name(path.name + ".tmp")
        temp.write_text(payload, encoding="utf-8")
        os.replace(temp, path)                     # 原子替换,避免写一半损坏配置
        if os.name != "nt":
            with contextlib.suppress(OSError):
                os.chmod(path, 0o600)              # 配置含明文密码,收紧为仅本人可读写

    def profile(self, name: str) -> dict:
        return dict(self.profiles.get(name) or {})

    def save_profile(self, name: str, db_cfg, ssh_cfg=None) -> None:
        """记住某个适配器的连接参数(密码一并记忆,文件权限已收紧)。"""
        self.profiles[name] = dict(db_cfg.as_dict())
        if ssh_cfg is not None:
            self.ssh = ssh_cfg.as_dict()
        self.password = str(getattr(db_cfg, "password", "") or "")
        self.database = str(getattr(db_cfg, "database", "") or "")
