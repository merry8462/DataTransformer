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

from dt_config import (DEFAULT_BATCH, IDENT_RE, MODE_CREATE, MODE_REPLACE, MODES,
                       SAMPLE_SIZE, DBConfig)
from dt_files import normalize_row, to_db_value
from dt_mongo import MongoAdapter, MongoConfig, describe_mongo_error


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


def take_sample(rows, size=SAMPLE_SIZE, width=None):
    """预读前 size 行用于类型推断,返回 (采样行, 采样行+剩余行 的生成器)。"""
    normalize = (lambda row: normalize_row(row, width)) if width else list
    sample = [normalize(row) for row in islice(rows, size)]

    def chained():
        yield from sample
        for row in rows:
            yield normalize(row)

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
