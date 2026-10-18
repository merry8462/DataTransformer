# -*- coding: utf-8 -*-
"""
MongoDB 适配器(dt_mongo)
=======================
* 集合(collection)↔ Excel Sheet 名 / CSV 文件 / JSON 一级键;
* 文档字段 ↔ 首行表头;
* 数值保真:``Decimal`` 原样写入(bson Decimal128),超过 Decimal128 精度或
  超出 int64 范围的数值改以**文本**写入,保证数值与文本都不被截断;
* 嵌套文档 / 数组:导出为 JSON 文本单元格,导入时自动还原为嵌套结构
  (键名含 ``.`` 或以 ``$`` 开头时保持文本,避免 MongoDB 键名限制导致写入失败)。

本模块不依赖 dt_core,由 dt_core 单向引用,避免循环依赖。
"""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime, time
from decimal import Decimal

from dt_numeric import format_number, loads_json

try:
    import pymongo
    from bson import Binary, ObjectId
    from bson.decimal128 import Decimal128
    from pymongo import MongoClient
    from pymongo.errors import OperationFailure, PyMongoError, ServerSelectionTimeoutError
except Exception:                                     # pragma: no cover - 未安装驱动
    pymongo = MongoClient = None
    Binary = ObjectId = Decimal128 = None
    PyMongoError = ServerSelectionTimeoutError = OperationFailure = Exception

NAME = "MongoDB"
KIND = "Mongo"
DEFAULT_PORT = 27017
SCHEMA_SAMPLE = 200          # 推断集合字段时采样的文档数
INT64_MIN, INT64_MAX = -2 ** 63, 2 ** 63 - 1
OBJECT_ID_RE = r"^[0-9a-fA-F]{24}$"


def pymongo_available() -> bool:
    return pymongo is not None


def _to_int(value, default: int) -> int:
    """宽容地把界面 / 环境变量文本转成整数(与 DBConfig.from_mapping 口径一致)。"""
    try:
        return int(str(value).strip())
    except (TypeError, ValueError):
        return default


@dataclass
class MongoConfig:
    """MongoDB 连接参数;填了 uri 时优先使用 URI。"""

    host: str = "127.0.0.1"
    port: int = DEFAULT_PORT
    user: str = ""
    password: str = ""
    database: str = ""
    auth_source: str = ""
    uri: str = ""
    timeout: int = 10

    @classmethod
    def from_mapping(cls, data=None) -> "MongoConfig":
        values = data or {}
        return cls(
            host=str(values.get("host") or "127.0.0.1").strip(),
            port=_to_int(values.get("port"), DEFAULT_PORT) or DEFAULT_PORT,
            user=str(values.get("user") or "").strip(),
            password=str(values.get("password") or ""),
            database=str(values.get("database") or "").strip(),
            auth_source=str(values.get("auth_source") or "").strip(),
            uri=str(values.get("uri") or "").strip(),
            timeout=_to_int(values.get("timeout"), 10) or 10,
        )

    def as_dict(self) -> dict:
        return {"host": self.host, "port": self.port, "user": self.user,
                "password": self.password, "database": self.database,
                "auth_source": self.auth_source, "uri": self.uri,
                "timeout": self.timeout}

    def validated(self, adapter=None) -> "MongoConfig":
        """校验必填项(签名与 DBConfig.validated 保持一致,便于上层统一调用)。"""
        if not self.uri:
            if not self.host:
                raise ValueError("请填写 MongoDB 主机或连接 URI")
            if not self.database:
                raise ValueError("请填写 MongoDB 数据库名")
        if not 1 <= self.port <= 65535:
            raise ValueError(f"MongoDB 端口无效:{self.port!r},请输入 1-65535 的整数")
        if not 1 <= self.timeout <= 3600:
            raise ValueError(f"超时秒数无效:{self.timeout!r},请输入 1-3600 的整数")
        return self

    def resolved_port(self, adapter=None) -> int:
        return int(self.port or DEFAULT_PORT)

    def target_uri(self, scheme: str = "mongodb") -> str:
        """拼出连接串(密码中的特殊字符按 URI 规范转义)。"""
        if self.uri:
            return self.uri
        auth = ""
        if self.user:
            user = quote(self.user)
            password = quote(self.password) if self.password else ""
            auth = f"{user}:{password}@" if password else f"{user}@"
        source = self.auth_source or ("admin" if auth else "")
        query = f"/?authSource={source}" if source else "/"
        return f"{scheme}://{auth}{self.host}:{self.port}{query}"


def quote(text: str) -> str:
    from urllib.parse import quote as _quote
    return _quote(str(text), safe="")


def describe_mongo_error(exc, cfg: MongoConfig | None = None) -> str:
    """把 MongoDB 异常翻译成便于定位的中文提示(保留原始错误)。"""
    raw = str(exc)
    low = raw.lower()
    target = (cfg.uri or f"{getattr(cfg, 'host', '?')}:{getattr(cfg, 'port', '?')}") if cfg else "?"
    if isinstance(exc, (ValueError, TypeError)):
        return raw
    if "default database" in low:
        return (f"没有指定 MongoDB 数据库。\n请在【数据库】填写库名,"
                f"或在【连接 URI】里带上 /库名(例如 mongodb://127.0.0.1:27017/MyData)。\n"
                f"原始错误: {raw}")
    if isinstance(exc, ServerSelectionTimeoutError) or "serverselectiontimeout" in low \
            or "no servers found" in low:
        return (f"无法连接 MongoDB {target}。\n"
                f"请检查【主机】【端口】、mongod 是否启动、防火墙与副本集配置。\n原始错误: {raw}")
    if isinstance(exc, OperationFailure) or "authentication failed" in low:
        if "auth" in low or "credential" in low:
            return (f"MongoDB 认证失败。\n请检查【用户名】【密码】【认证库 authSource】。\n"
                    f"原始错误: {raw}")
        return f"MongoDB 操作被拒绝:{raw}"
    if "timed out" in low or "timeout" in low:
        return f"连接 MongoDB {target} 超时。\n请检查主机/端口/网络与防火墙。\n原始错误: {raw}"
    if "name or service not known" in low or "getaddrinfo" in low:
        return f"MongoDB 主机名解析失败:{target}。\n请检查【主机】。\n原始错误: {raw}"
    return raw


# ---------------------------------------------------------------------------
# 值转换(MongoDB ↔ 单元格值)
# ---------------------------------------------------------------------------
def to_object_id(value):
    """24 位十六进制文本 → ObjectId(导入时还原 _id);失败返回 None。"""
    if ObjectId is None or not isinstance(value, str):
        return None
    text = value.strip()
    if len(text) == 24:
        try:
            return ObjectId(text)
        except Exception:
            return None
    return None


def mongo_cell_value(value):
    """MongoDB 文档字段值 → 单元格值(Decimal128 / ObjectId / Binary 归一化)。"""
    if Decimal128 is not None and isinstance(value, Decimal128):
        return value.to_decimal()
    if ObjectId is not None and isinstance(value, ObjectId):
        return str(value)
    if Binary is not None and isinstance(value, Binary):
        try:
            return bytes(value).decode("utf-8")
        except UnicodeDecodeError:
            return bytes(value).hex()
    return value


def _nested_keys_are_safe(value) -> bool:
    """MongoDB 要求键名不含 '.' 且不以 '$' 开头(嵌套对象 / 数组内部同样如此)。"""
    if isinstance(value, dict):
        return all("." not in str(key) and not str(key).startswith("$")
                   and _nested_keys_are_safe(item) for key, item in value.items())
    if isinstance(value, (list, tuple)):
        return all(_nested_keys_are_safe(item) for item in value)
    return True


def _store_decimal(value: Decimal):
    """``Decimal`` → ``Decimal128``;超出位数或指数范围时降级为定点文本保真。"""
    if Decimal128 is not None and value.is_finite():
        try:
            return Decimal128(value)               # 位数 / 指数超出 Decimal128 会抛异常
        except (ArithmeticError, ValueError, TypeError):
            pass
    return format_number(value)


def _store_scalar(value):
    """单个(非容器)单元格值 → 可写入 MongoDB 的值。"""
    if value is None or isinstance(value, (bool, datetime, date, time)):
        return value
    if isinstance(value, Decimal):
        return _store_decimal(value)
    if isinstance(value, int):
        return value if INT64_MIN <= value <= INT64_MAX else str(value)
    if isinstance(value, float):
        return value
    if isinstance(value, (bytes, bytearray)):
        return Binary(bytes(value)) if Binary is not None else bytes(value)
    if isinstance(value, (dict, list, tuple)):
        return _store_nested(value)
    if isinstance(value, str):
        text = value.strip()
        if text[:1] in ("{", "["):                  # 导出的 JSON 文本单元格 → 还原嵌套结构
            try:
                parsed = loads_json(text)
            except (ValueError, TypeError):
                return value
            if not _nested_keys_are_safe(parsed):
                return value
            return _store_nested(parsed)
        return value
    return str(value)


def _store_nested(value):
    """容器值 → 可写入 MongoDB 的结构(逐层换算 Decimal / 超范围整数,否则 bson 无法编码)。"""
    if isinstance(value, dict):
        return {str(key): _store_nested(item) for key, item in value.items()}
    if isinstance(value, (list, tuple)):
        return [_store_nested(item) for item in value]
    return _store_scalar(value)


def mongo_store_value(value):
    """单元格值 → 可写入 MongoDB 的值(保证数值不截断、嵌套结构可还原)。"""
    return _store_nested(value)


def row_to_document(header, row) -> dict:
    """一行单元格 → 一个 MongoDB 文档(按表头映射字段名)。"""
    document = {}
    for name, value in zip(header, row):
        column = str(name)
        if column == "_id":
            # 24 位十六进制还原为 ObjectId;其他类型的 _id(整数/字符串)原样保留;
            # 空值交给 MongoDB 自动生成,避免多条 null _id 触发唯一键冲突。
            if (object_id := to_object_id(value)) is not None:
                document["_id"] = object_id
            elif value is not None and str(value).strip():
                document["_id"] = mongo_store_value(value)
            continue
        document[column] = mongo_store_value(value)
    return document


def document_to_row(document, header) -> list:
    """一个 MongoDB 文档 → 一行单元格(按表头取值)。"""
    return [mongo_cell_value(document.get(str(name))) for name in header]


# ---------------------------------------------------------------------------
# 适配器
# ---------------------------------------------------------------------------
class MongoAdapter:
    """MongoDB 适配器:接口与 SQL 适配器保持一致(connect/list_tables/... )。"""

    NAME = NAME
    KIND = KIND
    DEFAULT_PORT = DEFAULT_PORT
    QUOTE = ""

    # -- 连接 --------------------------------------------------------------
    def connect(self, cfg: MongoConfig):
        if pymongo is None:
            raise RuntimeError("未安装 pymongo,请先执行: pip install pymongo")
        cfg = cfg.validated()
        options = {
            "serverSelectionTimeoutMS": int(cfg.timeout) * 1000,
            "connectTimeoutMS": int(cfg.timeout) * 1000,
            "socketTimeoutMS": int(cfg.timeout) * 1000,
        }
        if cfg.uri:
            client = MongoClient(cfg.uri, **options)
        else:
            client = MongoClient(host=cfg.host, port=int(cfg.port),
                                 username=cfg.user or None, password=cfg.password or None,
                                 authSource=cfg.auth_source or cfg.database or "admin",
                                 **options)
        database = client[cfg.database] if cfg.database else client.get_default_database()
        client.admin.command("ping")                # 立即暴露连接/认证问题
        return database

    @staticmethod
    def close(database) -> None:
        try:
            database.client.close()
        except Exception:
            pass

    def server_version(self, database) -> str:
        info = database.client.server_info()
        return f"MongoDB {info.get('version', '?')}"

    def quote_ident(self, name) -> str:
        """集合名无需引用符,但需校验,避免空名或非法字符导致命令异常。"""
        text = str(name).strip()
        if not text:
            raise ValueError("集合名不能为空")
        if "\x00" in text or text.startswith("system."):
            raise ValueError(f"非法集合名: {name!r}")
        return text

    # -- 结构枚举 ----------------------------------------------------------
    def list_tables(self, database):
        return sorted(name for name in database.list_collection_names()
                      if not name.startswith("system."))

    def table_exists(self, database, name) -> bool:
        return str(name) in database.list_collection_names()

    def list_columns(self, database, collection, limit: int = SCHEMA_SAMPLE):
        """采样若干文档,按出现顺序汇总字段名(异构文档取并集)。"""
        names = []
        for document in database[str(collection)].find({}, limit=limit):
            for key in document:
                if key not in names:
                    names.append(str(key))
        return names

    def count(self, database, collection) -> int:
        return int(database[str(collection)].estimated_document_count())


ADAPTER = MongoAdapter()


# ---------------------------------------------------------------------------
# 读取 / 写入(由 dt_core 的转换管道调用)
# ---------------------------------------------------------------------------
def load_mongo_stream(database, adapter, collection, columns=None, batch: int = 1000,
                      limit: int = 0):
    """读取集合,返回 (表头, 行生成器, 关闭函数)。"""
    header = list(columns) if columns else adapter.list_columns(database, collection)
    projection = {str(name): 1 for name in header}
    cursor = database[str(collection)].find({}, projection).batch_size(batch)
    if limit:
        cursor = cursor.limit(limit)

    def gen():
        for document in cursor:
            yield document_to_row(document, header)

    return header, gen(), cursor.close


def write_mongo_collection(database, adapter, collection, header, rows, mode, batch,
                           log, sample=None) -> int:
    """把行流写入集合(mode: replace 删除重建 / append 追加 / create_if_missing)。"""
    if mode not in ("replace", "create_if_missing", "append"):
        raise ValueError(f"未知写入模式: {mode!r}")
    name = adapter.quote_ident(collection)
    exists = adapter.table_exists(database, name)
    if mode == "append" and not exists:
        raise ValueError(f"集合【{name}】不存在。请改用“不存在则创建”或“删除重建”模式")
    if mode == "replace" and exists:
        database[name].drop()
        log(f"已删除集合: {name}")
    target = database[name]
    total, chunk = 0, []
    try:
        for row in rows:
            chunk.append(row_to_document(header, row))
            if len(chunk) >= batch:
                target.insert_many(chunk, ordered=False)
                total += len(chunk)
                chunk = []
                log(f"  已写入 {total:,} 行 ...")
        if chunk:
            target.insert_many(chunk, ordered=False)
            total += len(chunk)
    except Exception as exc:
        low = str(exc).lower()
        if "duplicate key" in low:
            raise RuntimeError(
                f"写入失败:{exc}\n提示:集合中已存在相同的 _id,"
                f"请改用“覆盖(删除重建)”模式,或在导出时去掉 _id 字段。") from exc
        if "document too large" in low:
            raise RuntimeError(
                f"写入失败:{exc}\n提示:单个文档超过 16MB 上限,"
                f"请拆分嵌套字段或改为文本存储。") from exc
        raise
    if total == 0 and not exists:
        database.create_collection(name)            # 空结果也建集合,保证表结构存在
    log(f"已写入集合: {name}({total:,} 行,模式: {mode})")
    return total
