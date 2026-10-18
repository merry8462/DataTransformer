# -*- coding: utf-8 -*-
"""
数据库适配层(dt_sql)
====================
MySQL(PyMySQL)/ PostgreSQL(psycopg2)适配器:连接、库表枚举、字段结构、
服务端流式游标读取,以及批量写入、类型推断与自动建表。

MongoDB 适配器见 dt_mongo.py,两者接口保持一致(connect / list_tables /
list_columns / table_exists / server_version),因此转换引擎无需区分数据库种类。
"""

from __future__ import annotations

import contextlib
import os
import re
from datetime import date, datetime
from decimal import Decimal
from itertools import islice

try:
    import pymysql
except Exception:  # pragma: no cover - 未安装驱动时仍可使用文件互转
    pymysql = None

try:
    import psycopg2
except Exception:  # pragma: no cover
    psycopg2 = None

from .dt_config import (DEFAULT_BATCH, IDENT_RE, MODE_CREATE, MODE_REPLACE, MODES,
                        SAMPLE_SIZE, DBConfig)
from .dt_files import normalize_row, to_db_value
from .dt_mongo import MongoAdapter, MongoConfig, describe_mongo_error


class DBAdapter:
    """数据库适配器基类,MySQL 与 PostgreSQL 各自实现差异点。"""

    NAME = ""
    KIND = "SQL"
    DEFAULT_PORT = 3306
    QUOTE = '"'

    def connect(self, cfg: DBConfig):
        raise NotImplementedError

    def open_query_cursor(self, conn):
        """打开一个流式查询游标(避免一次性把全部结果缓冲到客户端内存)。"""
        raise NotImplementedError

    def list_tables(self, conn):
        raise NotImplementedError

    def list_columns(self, conn, table):
        raise NotImplementedError

    def is_missing_table_error(self, exc):
        raise NotImplementedError

    # -- 通用能力 ----------------------------------------------------------
    def quote_ident(self, name) -> str:
        """校验并引用标识符,支持 schema.table 形式,防止 SQL 注入。

        采用白名单校验:字母 / 下划线 / 中文开头,其余字符仅允许字母、数字、下划线,
        因此表名与字段名中不能出现空格、括号、点号(点号用于分隔 schema)等字符。
        """
        parts = [part.strip() for part in str(name).strip().split(".") if part.strip()]
        if not parts:
            raise ValueError("表名/字段名不能为空")
        if illegal := [part for part in parts if not IDENT_RE.match(part)]:
            raise ValueError(
                f"非法标识符 {illegal[0]!r}:只允许字母 / 下划线 / 中文开头,"
                f"且只能包含字母、数字、下划线(不能有空格、括号、横线等字符)")
        quoted = self.QUOTE
        return ".".join(f"{quoted}{part}{quoted}" for part in parts)

    def table_exists(self, conn, quoted_table) -> bool:
        """通过 SELECT 探测表是否存在(跨库、跨 schema 均可靠)。"""
        try:
            with conn.cursor() as cur:
                cur.execute(f"SELECT 1 FROM {quoted_table} WHERE 1=0")
            return True
        except Exception as exc:
            # PostgreSQL 中任何出错语句都会使当前事务进入 aborted 状态,
            # 必须先 rollback 才能继续后续 SQL。
            with contextlib.suppress(Exception):
                conn.rollback()
            if self.is_missing_table_error(exc):
                return False
            raise

    def server_version(self, conn) -> str:
        with conn.cursor() as cur:
            cur.execute("SELECT VERSION()")
            return str(cur.fetchone()[0])


class MySQLAdapter(DBAdapter):
    NAME = "MySQL"
    DEFAULT_PORT = 3306
    QUOTE = "`"

    def connect(self, cfg: DBConfig):
        if pymysql is None:
            raise RuntimeError("未安装 PyMySQL,请先执行: pip install pymysql")
        return pymysql.connect(
            host=cfg.host,
            port=int(cfg.port),
            user=cfg.user,
            password=cfg.password,
            database=cfg.database,
            charset=cfg.charset or "utf8mb4",
            connect_timeout=int(cfg.timeout),
        )

    def open_query_cursor(self, conn):
        # SSCursor = 服务端无缓冲游标,大数据量导出不占客户端内存
        return conn.cursor(pymysql.cursors.SSCursor)

    def list_tables(self, conn):
        with conn.cursor() as cur:
            cur.execute("SHOW TABLES")
            return [str(row[0]) for row in cur.fetchall()]

    def list_columns(self, conn, table):
        with conn.cursor() as cur:
            cur.execute(f"SHOW COLUMNS FROM {self.quote_ident(table)}")
            return [str(row[0]) for row in cur.fetchall()]

    def is_missing_table_error(self, exc) -> bool:
        errno = exc.args[0] if getattr(exc, "args", None) else None
        return isinstance(exc, pymysql.err.ProgrammingError) and errno == 1146


class PostgresAdapter(DBAdapter):
    NAME = "PostgreSQL"
    DEFAULT_PORT = 5432
    QUOTE = '"'

    def connect(self, cfg: DBConfig):
        if psycopg2 is None:
            raise RuntimeError("未安装 psycopg2,请先执行: pip install psycopg2-binary")
        return psycopg2.connect(
            host=cfg.host,
            port=int(cfg.port),
            user=cfg.user,
            password=cfg.password,
            dbname=cfg.database,
            connect_timeout=int(cfg.timeout),
        )

    def open_query_cursor(self, conn):
        # 服务端命名游标:结果按 itersize 分批从服务器拉取
        cur = conn.cursor(name=f"dt_cur_{os.getpid()}_{id(conn)}")
        cur.itersize = 5000
        return cur

    def list_tables(self, conn):
        with conn.cursor() as cur:
            cur.execute(
                "SELECT table_schema, table_name FROM information_schema.tables "
                "WHERE table_schema NOT IN ('pg_catalog', 'information_schema') "
                "ORDER BY table_schema, table_name"
            )
            return [f"{schema}.{name}" if schema != "public" else str(name)
                    for schema, name in cur.fetchall()]

    def list_columns(self, conn, table):
        parts = str(table).split(".")
        schema, tname = (parts[0], parts[-1]) if len(parts) == 2 else ("public", parts[-1])
        with conn.cursor() as cur:
            cur.execute(
                "SELECT column_name FROM information_schema.columns "
                "WHERE table_schema = %s AND table_name = %s "
                "ORDER BY ordinal_position",
                (schema, tname),
            )
            return [str(row[0]) for row in cur.fetchall()]

    def is_missing_table_error(self, exc) -> bool:
        return getattr(exc, "pgcode", None) == "42P01"


ADAPTERS = {"MySQL": MySQLAdapter(), "PostgreSQL": PostgresAdapter(),
            "MongoDB": MongoAdapter()}


def describe_db_error(exc, cfg=None, adapter=None) -> str:
    """把数据库/网络异常翻译成便于用户定位的中文提示(不吞掉原始错误)。"""
    if adapter is not None and getattr(adapter, "KIND", "SQL") == "Mongo":
        mongo_cfg = cfg if isinstance(cfg, MongoConfig) else MongoConfig.from_mapping(cfg)
        return describe_mongo_error(exc, mongo_cfg)
    raw = str(exc)
    low = raw.lower()
    view = cfg.as_dict() if isinstance(cfg, DBConfig) else dict(cfg or {})
    host = view.get("host") or "?"
    port = view.get("port") or "?"
    name = adapter.NAME if adapter is not None else ""

    if name == "MySQL":
        first = exc.args[0] if getattr(exc, "args", None) else None
        errno = first if isinstance(first, int) else None
        hints = {
            1130: (f"MySQL 服务器拒绝了客户端主机 {host!r} 的连接。\n"
                   "请确认该主机是否在 MySQL 用户授权范围内(例如用户需要以 "
                   "'用户名'@'%' 或 '用户名'@'该IP' 形式授权)。\n"),
            1045: "MySQL 用户名或密码错误。\n请检查【用户名】【密码】。\n",
            1044: "MySQL 当前用户没有访问【数据库】的权限。\n请检查数据库名与用户授权。\n",
            1049: "MySQL 数据库不存在。\n请检查【数据库】名称。\n",
            2002: (f"无法连接 MySQL 服务器 {host}:{port}。\n"
                   "请检查【主机】【端口】是否正确、MySQL 服务是否启动、防火墙是否放行。\n"),
            2003: (f"无法连接 MySQL 服务器 {host}:{port}。\n"
                   "请检查【主机】【端口】是否正确、MySQL 服务是否启动、防火墙是否放行。\n"),
        }
        if errno in hints:
            return hints[errno] + f"原始错误: {raw}"
        if any(key in low for key in ("access denied", "password")):
            return f"MySQL 用户名或密码错误。\n请检查【用户名】【密码】。\n原始错误: {raw}"
        if "unknown database" in low:
            return f"MySQL 数据库不存在。\n请检查【数据库】名称。\n原始错误: {raw}"

    elif name == "PostgreSQL":
        pgcode = getattr(exc, "pgcode", None)
        if pgcode == "28P01" or "password authentication failed" in low:
            return f"PostgreSQL 用户名或密码错误。\n请检查【用户名】【密码】。\n原始错误: {raw}"
        if pgcode == "3D000" or "does not exist" in low:
            return f"PostgreSQL 数据库不存在。\n请检查【数据库】名称。\n原始错误: {raw}"
        if pgcode == "28000" or "no pg_hba.conf" in low:
            return f"PostgreSQL 登录被拒绝(认证方式不允许)。\n请检查用户与 pg_hba.conf。\n原始错误: {raw}"
        if any(key in low for key in ("could not connect", "connection refused")):
            return (f"无法连接 PostgreSQL 服务器 {host}:{port}。\n"
                    f"请检查【主机】【端口】、服务状态与防火墙。\n原始错误: {raw}")

    if isinstance(exc, (ValueError, TypeError)):
        return raw                      # 校验类错误已是中文提示,直接展示
    if any(key in low for key in ("timed out", "timeout", "connect timeout")):
        return f"连接 {host}:{port} 超时。\n请检查主机/端口/网络/防火墙。\n原始错误: {raw}"
    if any(key in low for key in ("name or service not known", "getaddrinfo", "gaierror")):
        return f"主机名解析失败:{host!r}。\n请检查【主机】名称。\n原始错误: {raw}"
    if "refused" in low:
        return f"连接被拒绝:{host}:{port}。\n请检查【端口】与服务是否启动。\n原始错误: {raw}"
    return raw


# ---------------------------------------------------------------------------
# 字段类型推断与建表
# ---------------------------------------------------------------------------
class ColumnStats:
    """单列的采样类型统计,用于 CREATE TABLE 的类型推断。

    ``Decimal`` 与 ``float`` 分开统计:前者对应精确数值(建 NUMERIC / DECIMAL),
    后者才是双精度浮点(建 DOUBLE)。否则高精度小数会被数据库按 float 存储而丢精度。
    """

    __slots__ = ("text", "max_len", "integer", "big", "real", "numeric", "boolean",
                 "stamp", "day")

    def __init__(self):
        self.text = False
        self.max_len = 0
        self.integer = False
        self.big = False
        self.real = False
        self.numeric = False
        self.boolean = False
        self.stamp = False
        self.day = False

    def observe(self, value) -> None:
        if value is None:
            return
        if isinstance(value, bool):
            self.boolean = True
        elif isinstance(value, int):
            self.integer = True
            self.big = self.big or not -(2 ** 31) <= value <= 2 ** 31 - 1
        elif isinstance(value, Decimal):
            self.numeric = True                      # 精确小数:绝不经 float
        elif isinstance(value, float):
            self.real = True
        elif isinstance(value, datetime):
            self.stamp = True
        elif isinstance(value, date):
            self.day = True
        else:
            self.text = True
            self.max_len = max(self.max_len, len(str(value)))


_TYPE_RULES = (
    (lambda st: st.text, "TEXT", "TEXT"),
    (lambda st: st.stamp, "TIMESTAMP", "TIMESTAMP"),
    (lambda st: st.day, "DATE", "DATE"),
    (lambda st: st.numeric, "DECIMAL(65,30)", "NUMERIC"),          # 高精度小数
    (lambda st: st.real, "DOUBLE", "DOUBLE PRECISION"),
    (lambda st: st.boolean and not st.integer, "TINYINT(1)", "BOOLEAN"),
    (lambda st: st.integer and st.big, "BIGINT", "BIGINT"),
    (lambda st: st.integer, "INT", "INT"),
)


def decide_type(adapter, stats: ColumnStats) -> str:
    """按采样统计决定字段类型。

    字符串列统一建为 TEXT:采样只覆盖前 SAMPLE_SIZE 行,若按采样长度建
    VARCHAR(n),后续出现更长的值会触发 "value too long" 写入失败。
    TEXT(MySQL 上限 64KB,PostgreSQL 无上限)可彻底规避该问题。
    """
    mysql = adapter.NAME == "MySQL"
    return next((mysql_type if mysql else pg_type
                 for rule, mysql_type, pg_type in _TYPE_RULES if rule(stats)), "TEXT")


def collect_stats(header, sample_rows) -> list:
    """统计采样行中各列出现的数据类型。"""
    stats = [ColumnStats() for _ in header]
    for row in sample_rows:
        for column, value in zip(stats, normalize_row(row, len(stats))):
            column.observe(value)
    return stats


def build_create_table(adapter, table, header, stats) -> str:
    columns = ", ".join(f"{adapter.quote_ident(name)} {decide_type(adapter, stat)}"
                        for name, stat in zip(header, stats))
    return f"CREATE TABLE {adapter.quote_ident(table)} ({columns})"


# ---------------------------------------------------------------------------
# 数据库读写
# ---------------------------------------------------------------------------
def close_cursor(conn, cur) -> None:
    """关闭流式游标并结束事务(PostgreSQL 命名游标必须 rollback 才能收尾)。"""
    if cur is not None:
        with contextlib.suppress(Exception):
            cur.close()
    if conn is not None:
        with contextlib.suppress(Exception):
            conn.rollback()


def load_sql_stream(conn, adapter, table, columns=None):
    """读取数据库表,返回 (表头, 行生成器, 游标)。调用方负责关闭游标。"""
    column_sql = ", ".join(adapter.quote_ident(name) for name in columns) if columns else "*"
    cur = adapter.open_query_cursor(conn)
    cur.execute(f"SELECT {column_sql} FROM {adapter.quote_ident(table)}")
    # PostgreSQL 命名游标的 description 要等首次 fetch 后才就绪
    first = cur.fetchmany(DEFAULT_BATCH)
    header = [str(desc[0]) for desc in cur.description]

    def gen():
        yield from (list(row) for row in first)
        while rows := cur.fetchmany(DEFAULT_BATCH):
            yield from (list(row) for row in rows)

    return header, gen(), cur


# ---------------------------------------------------------------------------
# 自由查询语句(语句由用户自行书写,不做标识符引用)
# ---------------------------------------------------------------------------
READONLY_KEYWORDS = frozenset({"select", "with", "show", "describe", "desc",
                               "explain", "values", "table"})
# WITH 可能包裹改数据的 CTE(WITH x AS (UPDATE ...) SELECT ...),需再扫一遍改数据关键字
WRITE_KEYWORDS_RE = re.compile(r"\b(insert|update|delete|merge|drop|alter|create|truncate|"
                               r"grant|revoke|call|copy)\b", re.IGNORECASE)


class QueryNoResultError(RuntimeError):
    """查询语句执行成功但没有返回结果集(不能作为导出数据源)。"""


def _strip_leading_comments(statement: str) -> str:
    """去掉语句开头的空白与注释(-- 行注释、/* */ 块注释)。"""
    while True:
        statement = statement.lstrip()
        if statement.startswith("--"):
            end = statement.find("\n")
            statement = "" if end < 0 else statement[end + 1:]
        elif statement.startswith("/*"):
            end = statement.find("*/", 2)
            statement = "" if end < 0 else statement[end + 2:]
        else:
            return statement


def ensure_single_statement(statement) -> str:
    """校验一次只执行一条语句,返回原语句。

    引号与注释之外若出现分号且其后仍有内容,说明填写了多条语句:这条边界同时保护
    预览(只读判断)与执行(psycopg2 简单查询协议会真的执行全部语句)。
    """
    text = str(statement or "")
    pos, quote_char = 0, ""
    while pos < len(text):
        char = text[pos]
        if quote_char:
            if char == "\\":
                pos += 2
                continue
            if char == quote_char:
                quote_char = ""
        elif char in ("'", '"', "`"):
            quote_char = char
        elif text.startswith("--", pos):
            end = text.find("\n", pos)
            pos = len(text) if end < 0 else end + 1
            continue
        elif text.startswith("/*", pos):
            end = text.find("*/", pos + 2)
            pos = len(text) if end < 0 else end + 2
            continue
        elif char == ";" and _strip_leading_comments(text[pos + 1:]).strip():
            raise ValueError("一次只允许一条语句:语句中出现了多个分号分隔的语句")
        pos += 1
    return text


def classify_sql(statement) -> bool:
    """判定语句是否为只读查询(只有只读语句允许被预览执行)。

    首关键字在只读白名单内才算只读;WITH 另需确认语句里没有改数据的关键字,
    避免改数据的 CTE 被当成只读。多语句(分号后仍有语句)一律拒绝:预览会真的执行语句,
    只读判定不能盖过“这条语句还能改库”的风险。
    """
    text = _strip_leading_comments(ensure_single_statement(statement))
    if not text:
        raise ValueError("查询语句不能为空")
    keyword = re.match(r"[A-Za-z]+", text)
    if keyword is None:
        return False                # 首个词不是关键字(如以括号开头):按非只读处理
    first = keyword.group(0).lower()
    if first not in READONLY_KEYWORDS:
        return False
    return first != "with" or WRITE_KEYWORDS_RE.search(text) is None


def load_sql_query(conn, adapter, statement):
    """执行用户自由书写的 SQL 语句,返回 (表头, 行生成器, 关闭函数)。

    语句原样执行(不做 quote_ident):允许任意 SQL,包含会修改数据库的 DML / DDL,
    风险由使用者自负;语句不返回结果集时抛 `QueryNoResultError`。
    只读语句走服务端流式游标;非只读语句改走普通客户端游标,因为 PostgreSQL 的命名游标
    只能 DECLARE ... SELECT,无法执行 UPDATE / DDL。非只读语句执行成功即提交,否则
    INSERT / UPDATE / DELETE ... RETURNING 这类有结果集的写语句会被 closer 的 rollback 撤掉。
    """
    readonly = classify_sql(statement)
    cur = adapter.open_query_cursor(conn) if readonly else conn.cursor()
    try:
        cur.execute(statement)
    except Exception:
        with contextlib.suppress(Exception):
            cur.close()                 # 执行失败也要收尾,避免命名游标泄漏
        raise
    if not readonly:
        try:
            conn.commit()               # 提交失败要让使用者看到,不 suppress
        except Exception:
            with contextlib.suppress(Exception):
                cur.close()
            raise
    first = None
    if cur.description is None:
        # PostgreSQL 命名游标的 description 要等首次 fetch 后才就绪,所以先取一批:
        # 连一批都取不出来才是真正的“没有结果集”(驱动会直接报错)。
        with contextlib.suppress(Exception):
            first = cur.fetchmany(DEFAULT_BATCH)
    if cur.description is None:
        affected = getattr(cur, "rowcount", None)
        affected = affected if isinstance(affected, int) and affected >= 0 else -1
        with contextlib.suppress(Exception):
            cur.close()
        raise QueryNoResultError(
            f"语句已执行(影响 {affected} 行),但没有返回结果集,无法作为导出数据源")
    header = [str(desc[0]) for desc in cur.description]
    if first is None:
        first = cur.fetchmany(DEFAULT_BATCH)

    def gen():
        yield from (list(row) for row in first)
        while rows := cur.fetchmany(DEFAULT_BATCH):
            yield from (list(row) for row in rows)

    return header, gen(), lambda: close_cursor(conn, cur)


def take_sample(rows, size=SAMPLE_SIZE, width=None):
    """预读前 size 行用于类型推断,返回 (采样行, 采样行+剩余行 的生成器)。"""
    normalize = (lambda row: normalize_row(row, width)) if width else list
    sample = [normalize(row) for row in islice(rows, size)]

    def chained():
        yield from sample
        yield from map(normalize, rows)

    return sample, chained()


def write_sql_table(conn, adapter, table, header, rows_iter, mode, batch, log,
                    sample=None) -> int:
    """把行流写入数据库表。

    sample=None 时单遍采样前 SAMPLE_SIZE 行推断类型(适合文件输入);
    调用方也可自行提供 sample 并传入未消费的行流(适合 SQL 输入,避免游标与 DDL 冲突)。
    """
    if mode not in MODES:
        raise ValueError(f"未知写入模式: {mode!r}(可选: {', '.join(MODES)})")
    if sample is None:
        sample, rows_iter = take_sample(rows_iter, SAMPLE_SIZE, len(header))

    quoted_table = adapter.quote_ident(table)
    exists = adapter.table_exists(conn, quoted_table)
    if mode == "append" and not exists:
        raise ValueError(f"目标表【{table}】不存在。请改用“{MODE_CREATE}”或“{MODE_REPLACE}”模式")
    if mode == "replace" or (mode == "create_if_missing" and not exists):
        if mode == "replace" and exists:
            with conn.cursor() as cur:
                cur.execute(f"DROP TABLE {quoted_table}")
            log(f"已删除旧表: {table}")
        with conn.cursor() as cur:
            cur.execute(build_create_table(adapter, table, header,
                                           collect_stats(header, sample)))
        log(f"已创建表: {table}")

    insert_sql = (f"INSERT INTO {quoted_table} "
                  f"({', '.join(adapter.quote_ident(name) for name in header)}) "
                  f"VALUES ({', '.join(['%s'] * len(header))})")
    total, chunk = 0, []
    try:
        with conn.cursor() as cur:
            for row in rows_iter:
                chunk.append(tuple(to_db_value(value) for value in row))
                if len(chunk) >= batch:
                    cur.executemany(insert_sql, chunk)
                    total += len(chunk)
                    chunk = []
                    log(f"  已写入 {total:,} 行 ...")
            if chunk:
                cur.executemany(insert_sql, chunk)
                total += len(chunk)
        conn.commit()
    except Exception as exc:
        with contextlib.suppress(Exception):
            conn.rollback()
        if any(key in str(exc).lower()
               for key in ("too long", "truncat", "太长了", "data too long")):
            raise RuntimeError(
                f"写入失败:{exc}\n"
                "提示:目标表字符串列长度不足。若目标是数据库中已存在的表,"
                "请将对应列改宽或改为 TEXT;若由本工具自动建表,"
                "请使用最新版本(自动建表已统一使用 TEXT 类型)。"
            ) from exc
        raise
    return total
