# -*- coding: utf-8 -*-
"""自由查询语句测试:SQL 语句分类 / SQL 自由查询读取 / MongoDB find 语句解析与读取 / 查询预览。"""

from decimal import Decimal

import pytest

import src.dt_core as core
import src.dt_mongo as mongo
import src.dt_sql as sql


# ---------------------------------------------------------------------------
# 假的连接 / 游标 / 适配器(无需真实数据库)
# ---------------------------------------------------------------------------
class FakeConnection:
    def __init__(self, plain_cursor=None):
        self.commits = 0
        self.rollbacks = 0
        self.plain_cursor = plain_cursor       # conn.cursor() 返回的普通客户端游标
        self.plain_cursor_calls = 0

    def cursor(self):
        self.plain_cursor_calls += 1
        return self.plain_cursor

    def commit(self):
        self.commits += 1

    def rollback(self):
        self.rollbacks += 1


class FailingCommitConnection(FakeConnection):
    """commit 直接失败的假连接(验证非只读分支提交失败也要收尾)。"""

    def commit(self):
        raise RuntimeError("commit 失败")


class FakeCursor:
    """有结果集的假游标:按批次给出 fetchmany 结果。"""

    def __init__(self, header=None, batches=(), rowcount=0):
        self.description = ([(name, None, None, None, None, None, None) for name in header]
                            if header else None)
        self._batches = [list(batch) for batch in batches]
        self.rowcount = rowcount
        self.closed = False
        self.statements = []

    def execute(self, statement):
        self.statements.append(statement)

    def fetchmany(self, size):
        return self._batches.pop(0) if self._batches else []

    def close(self):
        self.closed = True


class NoResultCursor(FakeCursor):
    """没有结果集的假游标:fetchmany 直接报错(与真实驱动行为一致)。"""

    def __init__(self, rowcount=0):
        super().__init__(None, (), rowcount)

    def fetchmany(self, size):
        raise RuntimeError("no results to fetch")


class NoRowcountCursor(NoResultCursor):
    """驱动没有可用行数时(取不到 rowcount)。"""

    def __init__(self):
        super().__init__()
        del self.rowcount


class BrokenCursor(FakeCursor):
    """execute 直接抛错的假游标(验证执行失败时仍会收尾关闭)。"""

    def __init__(self):
        super().__init__(header=["a"])

    def execute(self, statement):
        raise RuntimeError("execute 失败")


class FakeSQLAdapter:
    """流式游标适配器:未给定游标即视为“不该走流式游标”,一用就失败。"""

    def __init__(self, cursor=None):
        self.cursor = cursor
        self.open_calls = 0

    def open_query_cursor(self, conn):
        self.open_calls += 1
        if self.cursor is None:
            raise AssertionError("非只读语句不应使用流式游标")
        return self.cursor


class FakeKindAdapter:
    """只带 KIND 的伪适配器(转换入口的引擎一致性校验用,不连库)。"""

    def __init__(self, kind="SQL"):
        self.KIND = kind


class FakeMongoCursor:
    """假 Mongo 游标:记录链式调用,迭代时给出文档。"""

    def __init__(self, documents):
        self.documents = list(documents)
        self.batch = None
        self.sort_spec = None
        self.limit_value = None
        self.closed = False

    def batch_size(self, size):
        self.batch = size
        return self

    def sort(self, spec):
        self.sort_spec = spec
        return self

    def limit(self, count):
        self.limit_value = count
        return self

    def close(self):
        self.closed = True

    def __iter__(self):
        return iter(self.documents)


class FakeMongoCollection:
    def __init__(self, documents):
        self.cursor = FakeMongoCursor(documents)
        self.finds = []

    def find(self, filter=None, projection=None, limit=None):
        self.finds.append((filter, projection))
        return self.cursor


class FakeMongoDatabase:
    def __init__(self, collection):
        self.collection = collection
        self.names = []

    def __getitem__(self, name):
        self.names.append(name)
        return self.collection


# ---------------------------------------------------------------------------
# SQL 语句分类(只读判定)
# ---------------------------------------------------------------------------
def test_classify_sql_keywords():
    for statement in ("select 1", "SELECT * FROM t", "  with x as (select 1) select * from x",
                      "show tables", "describe t", "desc t", "explain select 1",
                      "values (1)", "table t"):
        assert sql.classify_sql(statement) is True, statement
    for statement in ("update t set a = 1", "insert into t values (1)", "delete from t",
                      "create table t (a int)", "drop table t", "alter table t add b int",
                      "truncate table t", "(select 1)"):
        assert sql.classify_sql(statement) is False, statement


def test_classify_sql_skips_leading_comments():
    assert sql.classify_sql("-- 导出用的查询\nselect 1") is True
    assert sql.classify_sql("/* 只读 */  select 1") is True
    assert sql.classify_sql("\n\t-- 注释\n/* 再注释 */ update t set a = 1") is False


def test_classify_sql_rejects_data_modifying_cte():
    for statement in ("WITH x AS (DELETE FROM t RETURNING *) SELECT 1",
                      "with x as (update t set a = 1 returning a) select * from x",
                      "WITH x AS (INSERT INTO t VALUES (1) RETURNING *) SELECT 1",
                      "WITH x AS (SELECT 1) DELETE FROM t"):
        assert sql.classify_sql(statement) is False, statement
    for statement in ("WITH x AS (SELECT update_time FROM t) SELECT * FROM x",
                      "WITH x AS (SELECT created_at, update_time FROM t) SELECT * FROM x",
                      "WITH x AS (SELECT 1) SELECT * FROM x",
                      "WITH x AS (SELECT 1) SELECT * FROM copy_log"):
        assert sql.classify_sql(statement) is True, statement


def test_classify_sql_requires_statement():
    for statement in ("", "   ", "-- 只有注释", "/* 只有注释 */"):
        with pytest.raises(ValueError, match="查询语句不能为空"):
            sql.classify_sql(statement)


def test_classify_sql_rejects_multiple_statements():
    """多语句必须拒绝:只读判定不能盖过"这条语句还能改库"的风险。"""
    for statement in ("SELECT 1; DROP TABLE users",
                      "select * from t; update t set a=1",
                      "SELECT 1;;"):
        with pytest.raises(ValueError, match="只允许一条语句"):
            sql.classify_sql(statement)


def test_classify_sql_ignores_semicolons_in_quotes_and_comments():
    for statement in ("SELECT 'a;b' FROM t", 'SELECT "a;b" FROM t', "SELECT `a;b` FROM t",
                      "SELECT 1;", "-- 注释里的分号;不算\nSELECT 1", "/* ; */ SELECT 1",
                      "select * from t; -- 结尾注释"):
        assert sql.classify_sql(statement) is True, statement


def test_ensure_single_statement_returns_original_text():
    statement = "SELECT 'a;b' FROM t"
    assert sql.ensure_single_statement(statement) == statement


def test_load_sql_query_rejects_multiple_statements_before_opening_cursor():
    conn = FakeConnection()
    adapter = FakeSQLAdapter()                    # 一旦打开流式游标就会失败

    with pytest.raises(ValueError, match="只允许一条语句"):
        sql.load_sql_query(conn, adapter, "SELECT 1; DROP TABLE users")

    assert adapter.open_calls == 0 and conn.plain_cursor_calls == 0


# ---------------------------------------------------------------------------
# SQL 自由查询读取
# ---------------------------------------------------------------------------
def test_load_sql_query_streams_batches_and_closes():
    statement = 'select a, b from "t" where a > 1'          # 语句原样执行,不做 quote_ident
    cur = FakeCursor(header=["a", "b"], batches=[[(1, "x")], [(2, "y")]])
    conn = FakeConnection()
    header, rows, closer = sql.load_sql_query(conn, FakeSQLAdapter(cur), statement)

    assert header == ["a", "b"]
    assert list(rows) == [[1, "x"], [2, "y"]]
    assert cur.statements == [statement]
    closer()
    assert cur.closed is True and conn.rollbacks == 1        # 复用 close_cursor 语义


def test_load_sql_query_without_result_set_commits_and_raises():
    cur = NoResultCursor(rowcount=3)
    conn = FakeConnection(plain_cursor=cur)
    with pytest.raises(sql.QueryNoResultError) as excinfo:
        sql.load_sql_query(conn, FakeSQLAdapter(cur), "update t set a = 1")
    message = str(excinfo.value)
    assert "语句已执行" in message and "影响 3 行" in message and "没有返回结果集" in message
    assert conn.commits == 1 and cur.closed is True           # DML / DDL 已提交生效


def test_load_sql_query_reports_unknown_rowcount_as_minus_one():
    conn = FakeConnection(plain_cursor=NoRowcountCursor())
    with pytest.raises(sql.QueryNoResultError, match="影响 -1 行"):
        sql.load_sql_query(conn, FakeSQLAdapter(), "create table t (a int)")
    assert conn.commits == 1                                # 提交恰好一次,不在分支里重复提交


def test_non_readonly_statement_uses_plain_cursor():
    cur = NoResultCursor(rowcount=2)
    conn = FakeConnection(plain_cursor=cur)
    adapter = FakeSQLAdapter()                              # 一用流式游标就失败

    with pytest.raises(sql.QueryNoResultError, match="影响 2 行"):
        sql.load_sql_query(conn, adapter, "update t set a = 1")

    assert adapter.open_calls == 0                          # 没有使用流式游标
    assert conn.plain_cursor_calls == 1                     # 走普通客户端游标
    assert conn.commits == 1 and cur.closed is True


def test_readonly_statement_keeps_streaming_cursor():
    cur = FakeCursor(header=["a"], batches=[[(1,)]])
    conn = FakeConnection(plain_cursor=NoResultCursor())
    adapter = FakeSQLAdapter(cur)

    header, rows, closer = sql.load_sql_query(conn, adapter, "select a from t")

    assert header == ["a"] and list(rows) == [[1]]
    assert adapter.open_calls == 1 and conn.plain_cursor_calls == 0
    closer()
    assert cur.closed is True


def test_non_readonly_result_set_commits_before_reading_rows():
    """INSERT / UPDATE ... RETURNING:有结果集也要提交,否则 closer 的 rollback 会撤掉写入。"""
    cur = FakeCursor(header=["a"], batches=[[(1,)], [(2,)]], rowcount=2)
    conn = FakeConnection(plain_cursor=cur)

    header, rows, closer = sql.load_sql_query(
        conn, FakeSQLAdapter(), "insert into t values (1), (2) returning a")

    assert header == ["a"] and list(rows) == [[1], [2]]
    assert conn.commits == 1                                # 读取前已提交,写语句真正生效
    closer()
    assert conn.rollbacks == 1                              # closer 的 rollback 只是空操作


def test_readonly_result_set_does_not_commit():
    cur = FakeCursor(header=["a"], batches=[[(1,)]])
    conn = FakeConnection(plain_cursor=NoResultCursor())

    header, rows, closer = sql.load_sql_query(conn, FakeSQLAdapter(cur), "select a from t")

    assert header == ["a"] and list(rows) == [[1]]
    assert conn.commits == 0                                # 只读语句绝不引入提交
    closer()


@pytest.mark.parametrize("statement", ["select a from t", "update t set a = 1"])
def test_execute_failure_closes_cursor_and_reraises(statement):
    cur = BrokenCursor()
    conn = FakeConnection(plain_cursor=cur)

    with pytest.raises(RuntimeError) as excinfo:
        sql.load_sql_query(conn, FakeSQLAdapter(cur), statement)

    assert type(excinfo.value) is RuntimeError              # 异常原样抛出,不改写
    assert str(excinfo.value) == "execute 失败"
    assert cur.closed is True                               # 执行失败也要收尾关闭
    assert conn.commits == 0                                # 失败的事务不提交


def test_commit_failure_closes_cursor_and_reraises():
    cur = NoResultCursor(rowcount=1)
    conn = FailingCommitConnection(plain_cursor=cur)

    with pytest.raises(RuntimeError) as excinfo:
        sql.load_sql_query(conn, FakeSQLAdapter(), "update t set a = 1")

    assert type(excinfo.value) is RuntimeError              # 提交失败原样抛出
    assert str(excinfo.value) == "commit 失败"
    assert cur.closed is True                               # 提交失败也要关游标


# ---------------------------------------------------------------------------
# MongoDB 自由查询语句解析
# ---------------------------------------------------------------------------
def test_parse_mongo_query_full_chain():
    query = mongo.parse_mongo_query(
        'db.VirtualProfile.find({"Name": {"$regex": "张"}, "Id": {"$gt": 10}})'
        '.sort({"Id": -1}).limit(100).projection({"Name": 1, "Id": 1})')

    assert query.collection == "VirtualProfile"
    assert query.filter == {"Name": {"$regex": "张"}, "Id": {"$gt": 10}}
    assert query.sort == {"Id": -1}
    assert query.limit == 100
    assert query.projection == {"Name": 1, "Id": 1}
    assert query.as_dict() == {"collection": "VirtualProfile", "filter": query.filter,
                               "sort": query.sort, "limit": 100, "projection": query.projection}


def test_parse_mongo_query_optional_parts_and_whitespace():
    plain = mongo.parse_mongo_query('db.c.find({})')
    assert (plain.collection, plain.filter, plain.sort, plain.limit, plain.projection) == (
        "c", {}, {}, None, {})

    chained = mongo.parse_mongo_query("db.c.find({}).limit(5)\n  .projection({\"a\": 1})"
                                      "  .sort({ \"a\": 1 })   ;")
    assert chained.limit == 5 and list(chained.projection) == ["a"]
    assert chained.sort == {"a": 1}

    # JSON 字符串里的括号 / 转义不参与括号配对
    text = mongo.parse_mongo_query('db.c.find({"Name": "a)b}c"})')
    assert text.filter == {"Name": "a)b}c"}


BAD_QUERIES = [
    'db.Person.aggregate([{"$match": {}}])',                    # 不支持 aggregate
    'db.Person.distinct("Name")',                               # 不支持其它方法
    'Person.find({})',                                          # 缺少 db. 前缀
    'db.Person.find([1, 2])',                                   # filter 必须是 JSON 对象
    'db.Person.find({"a": })',                                  # JSON 非法
    'db.Person.find({}).limit(-1)',                             # limit 必须非负整数
    'db.Person.find({}).limit(1.5)',                            # limit 必须整数
    'db.Person.find({}).sort({"a": 1}).sort({"b": 1})',         # 方法重复
    'db.Person.find({"a": 1}',                                  # 缺少右括号
    'db.Person.find({}) extra',                                 # 有多余内容
]


@pytest.mark.parametrize("statement", BAD_QUERIES)
def test_parse_mongo_query_rejects_bad_syntax(statement):
    with pytest.raises(ValueError) as excinfo:
        mongo.parse_mongo_query(statement)
    message = str(excinfo.value)
    assert "db." in message and ".find(" in message             # 提示里带可用示例


def test_parse_mongo_query_numbers_are_pymongo_encodable():
    query = mongo.parse_mongo_query(
        'db.Person.find({"Amount": 1.5, '
        '"$or": [{"Name": {"$regex": "张"}}], "Nested.Field": 3})')

    amount = query.filter["Amount"]
    if hasattr(amount, "to_decimal"):                            # 装了 pymongo:Decimal128
        assert amount.to_decimal() == Decimal("1.5")
    else:                                                        # 未装驱动:降级为定点文本
        assert amount == "1.5"
    with pytest.raises(ValueError, match="int64"):                # 超范围不再降级为文本
        mongo.parse_mongo_query('db.Person.find({"Big": 18446744073709551616})')
    assert query.filter["Nested.Field"] == 3                     # 含点的键原样保留
    assert query.filter["$or"] == [{"Name": {"$regex": "张"}}]   # $ 操作符原样保留


def test_query_value_rejects_out_of_range_numbers():
    """查询条件里的数值不做降级:降级成文本会静默匹配不到,只剩"0 行"。"""
    with pytest.raises(ValueError, match="int64"):
        mongo.parse_mongo_query('db.Person.find({"Big": 9223372036854775808})')
    with pytest.raises(ValueError, match="Decimal128"):
        mongo.parse_mongo_query('db.Person.find({"Huge": 1e7000})')


@pytest.mark.skipif(not mongo.pymongo_available(), reason="未安装 pymongo")
def test_query_value_keeps_in_range_numbers_encodable():
    query = mongo.parse_mongo_query(
        'db.Person.find({"Big": 4611686018427387904, "Amount": 1.5})')

    assert query.filter["Big"] == 2 ** 62                    # 范围内整数原样保留
    assert query.filter["Amount"].to_decimal() == Decimal("1.5")


def test_write_path_still_degrades_oversized_numbers():
    """写库路径继续按文本保真降级(与查询路径的严格校验互不影响)。"""
    assert mongo.mongo_store_value(2 ** 63) == str(2 ** 63)
    assert isinstance(mongo.mongo_store_value(Decimal("1E+7000")), str)


# GUI 里给用户的示例:mongosh 写法(裸键 + $ 操作符 + 嵌套数组)必须能原样解析
MONGOSH_EXAMPLE = ('db.TestCollections.find({ Age: { $gte: 25, $lte: 45 }, '
                   'Example: "Text", $expr: { $and: [{ $isArray: "$Hobbies" }, '
                   '{ $gt: [ { $size: "$Hobbies" }, 3 ] }]}})')


def test_parse_mongo_query_accepts_mongosh_single_quotes():
    """原 test_parse_mongo_query_single_quotes_points_to_double_quotes:单引号现已兼容。"""
    query = mongo.parse_mongo_query("db.c.find({'Name': 'x'})")
    assert query.collection == "c"
    assert query.filter == {"Name": "x"}
    assert mongo.parse_mongo_query("db.c.find({ 'Name': 'x' })").filter == {"Name": "x"}


def test_parse_mongo_query_accepts_mongosh_bare_keys_example():
    query = mongo.parse_mongo_query(MONGOSH_EXAMPLE)

    assert query.collection == "TestCollections"
    assert query.filter["Age"] == {"$gte": 25, "$lte": 45}
    assert query.filter["Example"] == "Text"
    expr = query.filter["$expr"]
    assert expr["$and"][0] == {"$isArray": "$Hobbies"}
    assert expr["$and"][1] == {"$gt": [{"$size": "$Hobbies"}, 3]}
    assert expr["$and"][1]["$gt"][0]["$size"] == "$Hobbies"    # 字符串 "$Hobbies" 未被规整破坏
    assert expr["$and"][1]["$gt"][1] == 3                      # 数字字面量保持数字


def test_parse_mongo_query_bare_keys_with_single_quotes():
    query = mongo.parse_mongo_query("db.c.find({ Name: '张三', Age: { $gt: 20 } })")
    assert query.collection == "c"
    assert query.filter == {"Name": "张三", "Age": {"$gt": 20}}

    # 链式调用里的 sort / projection 同样兼容裸键与单引号
    chained = mongo.parse_mongo_query(
        "db.c.find({ Name: '张三' }).sort({ Age: -1 }).projection({ Name: 1 })")
    assert chained.sort == {"Age": -1}
    assert chained.projection == {"Name": 1}

    # 字面量 true / false / null 与数字不被当成裸键
    literals = mongo.parse_mongo_query('db.c.find({ a: true, b: null, c: 1.5, d: -3 })')
    amount = literals.filter.pop("c")
    assert (amount.to_decimal() if hasattr(amount, "to_decimal") else amount) == Decimal("1.5")
    assert literals.filter == {"a": True, "b": None, "d": -3}


@pytest.mark.parametrize("text", ["{ a: 1, }", "{ a: 1,}", "[1, 2, ]", "[1, 2,]", "{ a: [1,], }"])
def test_normalize_query_json_drops_trailing_commas(text):
    normalized = mongo._normalize_query_json(text)
    assert ",}" not in normalized and ",]" not in normalized
    mongo.loads_json(normalized)                               # 规整后必须是合法 JSON


def test_parse_mongo_query_accepts_trailing_commas():
    assert mongo.parse_mongo_query("db.c.find({ a: 1, })").filter == {"a": 1}
    assert mongo.parse_mongo_query('db.c.find({"a": 1, }).limit(2)').limit == 2


def test_normalize_query_json_keeps_string_content_intact():
    text = '{ "url": "http://a:8080/x,y", "note": "Age: 1, $gte:" }'
    assert mongo.loads_json(mongo._normalize_query_json(text)) == {
        "url": "http://a:8080/x,y", "note": "Age: 1, $gte:"}
    assert mongo.parse_mongo_query(text.join(("db.c.find(", ")"))).filter == {
        "url": "http://a:8080/x,y", "note": "Age: 1, $gte:"}


def test_normalize_query_json_handles_quote_escapes():
    assert mongo.loads_json(mongo._normalize_query_json(r"""{ q: 'it\'s "ok"' }""")) == {
        "q": 'it\'s "ok"'}
    assert mongo.loads_json(mongo._normalize_query_json(r'{"q": "a\"b"}')) == {"q": 'a"b'}
    assert mongo.loads_json(mongo._normalize_query_json(r"{ q: 'a\\b' }")) == {"q": "a\\b"}


@pytest.mark.parametrize("statement", ["db.c.find(notjson)", "db.c.find([1, 2])",
                                       "db.c.find({ a: })", "db.c.find({ 'a': })"])
def test_parse_mongo_query_rejects_invalid_json_with_hint(statement):
    with pytest.raises(ValueError) as excinfo:
        mongo.parse_mongo_query(statement)
    message = str(excinfo.value)
    assert "ObjectId()" in message and "ISODate()" in message   # 说明不支持构造函数
    assert "db." in message and ".find(" in message             # 仍带可用示例


def test_parse_mongo_query_bad_json_quotes_original_text():
    """真正非法的实参:报错要带上规整前的原文,便于定位。"""
    with pytest.raises(ValueError) as excinfo:
        mongo.parse_mongo_query('db.c.find({ a: })')
    message = str(excinfo.value)
    assert "规整前的原文" in message and "{ a: }" in message
    assert "ObjectId()" in message and "db." in message and ".find(" in message


# ---------------------------------------------------------------------------
# MongoDB 自由查询读取
# ---------------------------------------------------------------------------
def test_load_mongo_query_inclusion_projection_header():
    documents = [{"_id": 1, "Name": "张三", "Age": 30, "City": "北京"}]
    collection = FakeMongoCollection(documents)
    database = FakeMongoDatabase(collection)

    header, rows, closer = mongo.load_mongo_query(
        database, mongo.ADAPTER,
        'db.Person.find({"Age": {"$gte": 18}}).projection({"Name": 1, "Id": 1})'
        '.sort({"Id": -1}).limit(50)', batch=7)

    assert header == ["Name", "Id"]                              # 声明顺序,只取真值键
    assert list(rows) == [["张三", None]]
    assert database.names == ["Person"]
    assert collection.finds[0] == ({"Age": {"$gte": 18}},
                                   {"Name": 1, "Id": 1})         # 原样传给 find
    cursor = collection.cursor
    assert (cursor.batch, cursor.sort_spec, cursor.limit_value, cursor.closed) == (
        7, [("Id", -1)], 50, False)
    closer()
    assert cursor.closed is True


def test_load_mongo_query_field_only_projection_uses_collection_fields():
    documents = [{"_id": 1, "Name": "张三", "Secret": "x"}]
    database = FakeMongoDatabase(FakeMongoCollection(documents))

    header, rows, closer = mongo.load_mongo_query(
        database, mongo.ADAPTER, 'db.Person.find({}).projection({"Secret": 0})')
    assert header == ["_id", "Name"]                             # 排除式:集合字段去掉被排除键
    assert list(rows) == [[1, "张三"]]
    closer()

    header, rows, closer = mongo.load_mongo_query(database, mongo.ADAPTER, 'db.Person.find({})')
    assert header == ["_id", "Name", "Secret"]                   # 未写 projection:全部字段
    closer()


def test_load_mongo_query_rejects_collection_without_fields():
    database = FakeMongoDatabase(FakeMongoCollection([]))
    with pytest.raises(ValueError, match="没有可用字段"):
        mongo.load_mongo_query(database, mongo.ADAPTER, 'db.Empty.find({})')
    with pytest.raises(ValueError, match="没有可用字段"):
        mongo.load_mongo_query(database, mongo.ADAPTER, 'db.Empty.find({}).projection({"a": 0})')


def test_load_mongo_query_sample_honours_filter():
    """回退表头的采样必须带上用户 filter,否则异构集合会静默丢列。"""
    collection = FakeMongoCollection([{"_id": 1, "Age": 30}])
    database = FakeMongoDatabase(collection)

    header, rows, closer = mongo.load_mongo_query(
        database, mongo.ADAPTER, 'db.Person.find({"Age": {"$gte": 18}})')

    assert header == ["_id", "Age"]
    assert collection.finds[0] == ({"Age": {"$gte": 18}}, None)   # 采样:带上 filter
    assert collection.finds[1][0] == {"Age": {"$gte": 18}}        # 查询本身
    closer()


def test_list_columns_without_where_keeps_old_behaviour():
    collection = FakeMongoCollection([{"_id": 1, "A": 1}])
    database = FakeMongoDatabase(collection)

    assert mongo.ADAPTER.list_columns(database, "c") == ["_id", "A"]
    assert collection.finds[0] == ({}, None)                      # 不传 where 仍是空过滤


# ---------------------------------------------------------------------------
# 引擎侧:查询源与预览
# ---------------------------------------------------------------------------
def test_query_source_requires_connection():
    with pytest.raises(ValueError, match="SQL 输入需要先连接数据库"):
        core.open_reader(core.SourceSpec(kind="SQL", table="标签", query="select 1"))
    with pytest.raises(ValueError, match="MongoDB 输入需要先连接数据库"):
        core.open_reader(core.SourceSpec(kind="Mongo", table="标签", query="db.c.find({})"))


def test_query_source_rejects_columns():
    with pytest.raises(ValueError, match="不能再指定 columns"):
        core.open_reader(core.SourceSpec(kind="SQL", table="标签", query="select 1",
                                        columns=["a"]), conn=object(), adapter=object())
    with pytest.raises(ValueError, match="不能再指定 columns"):
        core.open_reader(core.SourceSpec(kind="Mongo", table="标签", query="db.c.find({})",
                                        columns=["a"]), conn=object(), adapter=object())


def test_open_reader_query_source_delegates_to_query_loader(monkeypatch):
    sql_calls, mongo_calls = [], []

    def fake_load_sql_query(conn, adapter, statement):
        sql_calls.append(statement)
        return ["H"], iter([[1]]), lambda: None

    def fake_load_mongo_query(conn, adapter, statement):
        mongo_calls.append(statement)
        return ["M"], iter([["x"]]), lambda: None

    monkeypatch.setattr(core, "load_sql_query", fake_load_sql_query)
    monkeypatch.setattr(core, "load_mongo_query", fake_load_mongo_query)

    header, rows, _ = core.open_reader(
        core.SourceSpec(kind="SQL", table="标签", query="  select 1  "),
        conn=object(), adapter=object())
    assert sql_calls == ["select 1"] and header == ["H"] and list(rows) == [[1]]

    header, rows, _ = core.open_reader(
        core.SourceSpec(kind="Mongo", table="标签", query="db.c.find({})"),
        conn=object(), adapter=object())
    assert mongo_calls == ["db.c.find({})"] and header == ["M"] and list(rows) == [["x"]]


def test_open_reader_without_query_keeps_old_behaviour(tmp_path):
    source = tmp_path / "plain.csv"
    core.write_csv_file(str(source), ["A", "B"], iter([[1, 2]]))
    src = core.SourceSpec(kind="CSV", path=str(source))
    assert src.query == ""

    header, rows, closer = core.open_reader(src)
    assert header == ["A", "B"] and list(rows) == [["1", "2"]]
    closer()


def test_preview_query_requires_statement():
    with pytest.raises(ValueError, match="没有填写查询语句"):
        core.preview_query(core.SourceSpec(kind="SQL"))
    with pytest.raises(ValueError, match="没有填写查询语句"):
        core.preview_query(core.SourceSpec(kind="SQL", query="   "))


def test_query_label_is_derived_from_the_statement():
    """标签:SQL 取第一个 from 后的表名(允许 schema.table / 引号),取不到用“自定义查询”。"""
    assert core.query_label("SELECT Id FROM VirtualProfile WHERE Name LIKE '%张%'",
                            "SQL") == "VirtualProfile"
    assert core.query_label('select * from "s"."t" join u on 1', "SQL") == "s.t"
    assert core.query_label("SELECT 1 FROM `db`.`tbl`", "SQL") == "db.tbl"
    assert core.query_label("select * from [dbo].[tbl]", "SQL") == "dbo.tbl"
    assert core.query_label("UPDATE t SET a = 1", "SQL") == "自定义查询"
    assert core.query_label("", "SQL") == "自定义查询"
    assert core.query_label('db.people.find({"a": 1}).sort({"Id": 1})', "Mongo") == "people"
    assert core.query_label("这不是查询语句", "Mongo") == "自定义查询"
    assert "query_label" in core.__all__


def test_preview_query_skips_non_readonly(monkeypatch):
    def forbidden(*args, **kwargs):
        raise AssertionError("非只读语句不允许执行")

    monkeypatch.setattr(core, "open_reader", forbidden)
    preview = core.preview_query(core.SourceSpec(kind="SQL", query="delete from t"))

    assert preview.readonly is False
    assert preview.header == [] and preview.rows == []
    # 非只读语句并非一定不产出数据:引擎支持 INSERT / UPDATE ... RETURNING
    assert "不是只读查询" in preview.message and "会修改数据库" in preview.message
    assert "返回结果集时仍会导出" in preview.message


def test_preview_query_reads_limited_sample_and_closes(monkeypatch):
    closed = []

    def fake_open_reader(src, conn=None, adapter=None):
        return (["Id", "Name"], iter([[index, "n%d" % index] for index in range(1, 11)]),
                lambda: closed.append(True))

    monkeypatch.setattr(core, "open_reader", fake_open_reader)

    preview = core.preview_query(core.SourceSpec(kind="SQL", query="select Id, Name from t"),
                                 limit=3)
    assert preview.readonly is True and preview.message == ""
    assert preview.header == ["Id", "Name"]
    assert preview.rows == [[1, "n1"], [2, "n2"], [3, "n3"]]
    assert closed == [True]                                  # 取完立刻关闭游标

    preview = core.preview_query(core.SourceSpec(kind="Mongo", query="db.c.find({})"), limit=1)
    assert preview.readonly is True and preview.rows == [[1, "n1"]]
    assert closed == [True, True]


def test_query_source_converts_to_csv(tmp_path, monkeypatch):
    monkeypatch.setattr(
        core, "open_reader",
        lambda src, conn=None, adapter=None: (["Id", "Name"],
                                              iter([[1, "张三"], [2, "李四"]]), lambda: None))
    produced = tmp_path / "query.csv"
    result = core.convert(
        [core.SourceSpec(kind="SQL", table="标签", query="select Id, Name from t")],
        core.TargetSpec(kind="CSV", path=str(produced), verify=False),
        conn=object(), adapter=FakeKindAdapter("SQL"))

    assert result.rows == 2 and produced.exists()
    header = core.read_csv_header(str(produced), ",", "utf-8-sig")
    assert header == ["Id", "Name"]
    assert list(core.iter_csv_rows(str(produced), ",", "utf-8-sig", len(header))) == [
        ["1", "张三"], ["2", "李四"]]


def test_facade_exports_query_api():
    names = ["classify_sql", "load_sql_query", "QueryNoResultError", "MongoQuery",
             "parse_mongo_query", "load_mongo_query", "QueryPreview", "preview_query"]
    assert [name for name in names if name not in core.__all__] == []
    assert all(getattr(core, name) is not None for name in names)
    assert core.QueryPreview() == core.QueryPreview([], [], False, "")
