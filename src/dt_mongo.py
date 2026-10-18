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

import re
from dataclasses import dataclass, field
from datetime import date, datetime, time
from decimal import Decimal

from .dt_numeric import format_number, loads_json

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

    def list_columns(self, database, collection, limit: int = SCHEMA_SAMPLE, where=None):
        """采样若干文档,按出现顺序汇总字段名(异构文档取并集)。

        ``where`` 把采样限制在同一过滤条件内:自由查询语句回退表头时若不传它,
        异构集合里只出现在被过滤掉文档中的字段就会被静默丢掉。
        """
        documents = database[str(collection)].find(where or {}, limit=limit)
        return list(dict.fromkeys(str(key) for document in documents for key in document))

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


# ---------------------------------------------------------------------------
# 自由查询语句:db.<集合>.find(<filter>)[.sort / .limit / .projection]
# ---------------------------------------------------------------------------
QUERY_USAGE = ('查询语句格式: db.<集合>.find(<filter>)[.sort(<json>)]'
               '[.limit(<整数>)][.projection(<json>)],例如: '
               'db.VirtualProfile.find({"Name": {"$regex": "张"}}).sort({"Id": 1}).limit(100)')
QUERY_METHODS = ("sort", "limit", "projection")
QUERY_HEAD_RE = re.compile(r"db\s*\.\s*([A-Za-z0-9_]+)\s*\.\s*find\s*\(")
QUERY_CALL_RE = re.compile(r"\.\s*([A-Za-z_][A-Za-z0-9_]*)\s*\(")
QUERY_LIMIT_RE = re.compile(r"^\d+$")
# 裸键字符集(字母/数字/下划线/$):mongosh 对象字面量里可以不加引号的键
BARE_KEY_CHARS = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789_$"
# 实参合法性的口径说明(与"兼容 mongosh 写法"的实际能力保持一致)
QUERY_JSON_HINT = ("查询条件必须是合法的 JSON 对象:本项目兼容 mongosh 的单引号与裸键写法,"
                   "但仍需是 JSON 字面量,不支持函数调用 / ObjectId() / ISODate() 这类构造函数")


@dataclass
class MongoQuery:
    """自由查询语句的解析结果:集合名 + find 的 filter / sort / limit / projection。"""

    collection: str
    filter: dict = field(default_factory=dict)
    sort: dict = field(default_factory=dict)
    limit: int | None = None
    projection: dict = field(default_factory=dict)

    def as_dict(self) -> dict:
        return {"collection": self.collection, "filter": self.filter, "sort": self.sort,
                "limit": self.limit, "projection": self.projection}


def _query_value(value):
    """查询参数 → pymongo 可编码形式(键名原样保留:$ 操作符与含点的键合法)。

    写库路径会把超范围数值降级成文本以保真,查询路径**不能**降级:条件里的数值一旦
    变成字符串,查询就静默匹配不到,用户只会看到"结果为 0 行"。
    """
    if isinstance(value, dict):
        return {str(key): _query_value(item) for key, item in value.items()}
    if isinstance(value, (list, tuple)):
        return [_query_value(item) for item in value]
    if isinstance(value, Decimal):
        stored = _store_decimal(value)
        if isinstance(stored, str):
            raise ValueError("查询条件里的数值超出 Decimal128 表示范围,请改用更小的数值")
        return stored
    if isinstance(value, int) and not isinstance(value, bool):
        if not INT64_MIN <= value <= INT64_MAX:
            raise ValueError("查询条件里的数值超出 int64 表示范围,请改用更小的数值")
        return value
    return value


def _skip_blanks(source: str, pos: int) -> int:
    """跳过空白,返回下一个非空白位置(识别裸键后的冒号与末尾逗号用)。"""
    while pos < len(source) and source[pos].isspace():
        pos += 1
    return pos


def _normalize_query_json(text: str) -> str:
    """mongosh 风格对象字面量 → 严格 JSON:裸键加引号、单引号串改双引号、去掉末尾多余逗号。

    按字符状态机扫描:字符串内的冒号 / 逗号 / $ 值原样保留,只改写字符串之外的裸键。
    """
    source = str(text)
    out, pos, size = [], 0, len(source)
    while pos < size:
        char = source[pos]
        if char in "\"'":                           # 字符串:统一改写为双引号形式
            delimiter, chars, pos = char, [], pos + 1
            while pos < size:
                cur = source[pos]
                if cur == "\\" and pos + 1 < size:
                    nxt, pos = source[pos + 1], pos + 2
                    if nxt == "'":
                        chars.append("'")           # JSON 里 \' 不是合法转义,退化为普通字符
                    elif nxt == '"':
                        chars.append('\\"')
                    else:
                        chars.append("\\" + nxt)    # 保留 \n / \\ / \uXXXX 等既有转义
                    continue
                if cur == delimiter:
                    pos += 1
                    break
                chars.append('\\"' if cur == '"' else cur)
                pos += 1
            out.append('"' + "".join(chars) + '"')
            continue
        end = pos
        while end < size and source[end] in BARE_KEY_CHARS:
            end += 1
        follow = _skip_blanks(source, end)          # 裸键:后面紧跟冒号的单词补双引号
        if end > pos and follow < size and source[follow] == ":":
            out.append('"' + source[pos:end] + '"')
            pos = end
            continue
        if char == ",":                             # 对象 / 数组末尾的多余逗号直接丢弃
            follow = _skip_blanks(source, pos + 1)
            if follow < size and source[follow] in "}]":
                pos += 1
                continue
        out.append(char)
        pos += 1
    return "".join(out)


def _query_object(text: str, name: str) -> dict:
    """解析 .find / .sort / .projection 的对象实参(兼容 mongosh 的裸键 / 单引号写法)。"""
    raw = str(text).strip()
    excerpt = raw if len(raw) <= 200 else raw[:200] + "..."
    try:
        parsed = loads_json(_normalize_query_json(raw))
    except (ValueError, TypeError) as exc:
        raise ValueError(
            f"查询语句里 {name} 不是合法 JSON: {exc}\n"
            f"规整前的原文: {excerpt}\n"
            f"{QUERY_JSON_HINT}\n{QUERY_USAGE}") from exc
    if not isinstance(parsed, dict):
        raise ValueError(f'查询语句里 {name} 必须是 JSON 对象,例如 {{"Id": 1}}\n'
                         f"{QUERY_JSON_HINT}\n{QUERY_USAGE}")
    return _query_value(parsed)


def _query_argument(statement: str, pos: int):
    """取 pos 起括号内的一个实参(跳过字符串里的括号与转义),返回 (实参文本, 右括号位置)。"""
    start, depth, quote_char = pos, 0, ""
    while pos < len(statement):
        char = statement[pos]
        if quote_char:
            if char == "\\":
                pos += 2
                continue
            if char == quote_char:
                quote_char = ""
        elif char in "\"'":
            quote_char = char
        elif char in "{[(":
            depth += 1
        elif char in "}])":
            if depth == 0:
                if char != ")":
                    raise ValueError(f"查询语句括号不匹配:{statement!r}\n{QUERY_USAGE}")
                return statement[start:pos], pos
            depth -= 1
        pos += 1
    raise ValueError(f"查询语句缺少右括号:{statement!r}\n{QUERY_USAGE}")


def parse_mongo_query(text) -> MongoQuery:
    """解析 db.<集合>.find(<filter>) 及可选的 .sort / .limit / .projection(任意顺序)。"""
    statement = str(text or "").strip()
    head = QUERY_HEAD_RE.match(statement)
    if head is None:
        raise ValueError(f"查询语句语法不正确,应以 db.<集合>.find(<filter>) 开头\n{QUERY_USAGE}")
    query = MongoQuery(collection=head.group(1))
    argument, pos = _query_argument(statement, head.end())
    query.filter = _query_object(argument, "filter")
    pos += 1
    seen = set()
    while True:
        pos = _skip_blanks(statement, pos)
        if pos >= len(statement) or statement[pos] != ".":
            break
        call = QUERY_CALL_RE.match(statement, pos)
        if call is None:
            raise ValueError(f"查询语句语法不正确:{statement[pos:]!r}\n{QUERY_USAGE}")
        name = call.group(1).lower()
        if name not in QUERY_METHODS:
            raise ValueError(f"查询语句不支持 .{name}()(只支持 {', '.join(QUERY_METHODS)})\n"
                             f"{QUERY_USAGE}")
        if name in seen:
            raise ValueError(f"查询语句里 .{name}() 重复出现,请只写一次\n{QUERY_USAGE}")
        seen.add(name)
        argument, pos = _query_argument(statement, call.end())
        pos += 1
        if name == "limit":
            value = argument.strip()
            if not QUERY_LIMIT_RE.match(value):
                raise ValueError(f"查询语句里 limit 的参数必须是非负整数,当前为 {value!r}\n"
                                 f"{QUERY_USAGE}")
            query.limit = int(value)
        else:
            setattr(query, name, _query_object(argument, name))
    tail = statement[pos:].strip().rstrip(";").strip()
    if tail:
        raise ValueError(f"查询语句在 {tail!r} 处有多余内容\n{QUERY_USAGE}")
    return query


def load_mongo_query(database, adapter, statement, batch: int = 1000):
    """执行自由查询语句对应的 find,返回 (表头, 行生成器, 关闭函数)。"""
    query = parse_mongo_query(statement)
    collection = adapter.quote_ident(query.collection)
    header = [str(key) for key, value in query.projection.items() if value]
    if not header:
        # 纯排除式 projection 与未写 projection 都按集合字段决定导出列
        header = adapter.list_columns(database, collection, where=query.filter)
        header = [name for name in header if name not in query.projection]
    if not header:
        raise ValueError(f"集合【{collection}】里没有可用字段,无法作为导出数据源"
                         f"(请检查集合是否为空)")
    cursor = database[collection].find(query.filter, query.projection).batch_size(batch)
    if query.sort:
        cursor = cursor.sort(list(query.sort.items()))
    if query.limit is not None:
        cursor = cursor.limit(query.limit)

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
