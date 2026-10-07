# -*- coding: utf-8 -*-
"""MongoDB 适配层测试:值转换 / 配置校验 / 真实服务端往返(无服务时自动跳过)。"""

from decimal import Decimal
from datetime import datetime
import json

import pytest

import dt_core as core
import dt_mongo as mongo
from dt_numeric import format_number

TEST_DB = "dtx_pytest"
LONG_VALUES = ["12345678901234567890.12345", "1.500", "9007199254740993", "0.1"]


# ---------------------------------------------------------------------------
# 纯函数:不依赖服务端
# ---------------------------------------------------------------------------
def test_mongo_config_validation():
    assert mongo.MongoConfig(host="h", database="d").validated().port == 27017
    assert mongo.MongoConfig(uri="mongodb://a/b").validated().uri == "mongodb://a/b"
    with pytest.raises(ValueError, match="数据库名"):
        mongo.MongoConfig(host="h").validated()
    with pytest.raises(ValueError, match="端口"):
        mongo.MongoConfig(host="h", database="d", port=70000).validated()
    with pytest.raises(ValueError, match="超时"):
        mongo.MongoConfig(host="h", database="d", timeout=0).validated()


def test_mongo_uri_building_escapes_credentials():
    cfg = mongo.MongoConfig(host="1.2.3.4", port=27018, user="u@x", password="p:w", database="d")
    uri = cfg.target_uri()
    assert uri.startswith("mongodb://u%40x:p%3Aw@1.2.3.4:27018/")
    assert "authSource=admin" in uri
    assert mongo.MongoConfig(uri="mongodb://keep/me").target_uri() == "mongodb://keep/me"


@pytest.mark.skipif(not mongo.pymongo_available(), reason="未安装 pymongo")
def test_object_id_and_cell_conversion():
    from bson import ObjectId
    from bson.decimal128 import Decimal128
    oid = ObjectId()
    assert mongo.to_object_id(str(oid)) == oid
    assert mongo.to_object_id("not-an-object-id") is None
    assert mongo.mongo_cell_value(Decimal128("1.500")) == Decimal("1.500")
    assert mongo.mongo_cell_value(oid) == str(oid)


def test_store_value_keeps_precision():
    """MongoDB 侧:Decimal128 可容纳的按 Decimal128 存,超出的按文本保真。"""
    stored = mongo.mongo_store_value(Decimal("1.500"))
    assert format_number(mongo.mongo_cell_value(stored)) == "1.500"      # 尾零保留
    long_value = mongo.mongo_store_value(Decimal("12345678901234567890.12345"))
    assert format_number(mongo.mongo_cell_value(long_value)) == "12345678901234567890.12345"
    huge = Decimal("1" * 40)
    assert mongo.mongo_store_value(huge) == "1" * 40
    assert mongo.mongo_store_value(2 ** 63) == str(2 ** 63)          # 超出 int64
    assert mongo.mongo_store_value(42) == 42
    assert mongo.mongo_store_value(None) is None


def test_store_value_restores_nested_json_but_not_unsafe_keys():
    assert mongo.mongo_store_value('{"ip": "1.2.3.4"}') == {"ip": "1.2.3.4"}
    assert mongo.mongo_store_value('[1, 2]') == [1, 2]
    assert mongo.mongo_store_value('{"a.b": 1}') == '{"a.b": 1}'      # 键名含点 → 保持文本
    assert mongo.mongo_store_value('{"$set": 1}') == '{"$set": 1}'    # 以 $ 开头 → 保持文本
    assert mongo.mongo_store_value('{"a": {"$set": 1}}') == '{"a": {"$set": 1}}'  # 嵌套同样判定
    assert mongo.mongo_store_value("普通文本") == "普通文本"


@pytest.mark.skipif(not mongo.pymongo_available(), reason="未安装 pymongo")
def test_nested_values_can_be_encoded_by_bson():
    """还原后的嵌套结构里若仍是 decimal.Decimal,bson 会编码失败(导入直接报错)。

    所以嵌套文档 / 数组里的数值必须逐层换算成 Decimal128(超范围才降级为文本)。
    """
    from bson import encode
    from bson.decimal128 import Decimal128

    stored = mongo.mongo_store_value('{"score": 1.5, "tags": [1, 2.5]}')
    assert isinstance(stored["score"], Decimal128)
    assert all(isinstance(item, Decimal128) for item in stored["tags"][1:])

    document = mongo.row_to_document(["_id", "device"],
                                     ["0123456789abcdef01234567",
                                      mongo.loads_json('{"score": 1.5, "tags": [1, 2.5]}')])
    encode(document)                                     # 不抛异常即说明可写入


def test_mongo_config_tolerates_invalid_numeric_text():
    """端口 / 超时填了非数字时回落默认值(与 DBConfig 一致),不抛英文转换错误。"""
    cfg = mongo.MongoConfig.from_mapping({"host": "h", "port": "abc", "timeout": "xyz"})
    assert cfg.port == mongo.DEFAULT_PORT and cfg.timeout == 10


def test_store_value_handles_decimal128_range_limits():
    """位数不超但指数越界的 Decimal 不能抛异常,应降级为定点文本。"""
    assert mongo.mongo_store_value(Decimal("1E+7000")) == "1" + "0" * 7000
    assert mongo.mongo_store_value(Decimal("1E-7000")) == "0." + "0" * 6999 + "1"


def test_row_and_document_roundtrip():
    header = ["_id", "Name", "Amount"]
    oid_text = "0123456789abcdef01234567"
    document = mongo.row_to_document(header, [oid_text, "张三", Decimal("1.500")])
    assert str(document["_id"]) == oid_text
    assert format_number(mongo.mongo_cell_value(document["Amount"])) == "1.500"
    assert mongo.document_to_row(document, header) == [oid_text, "张三", Decimal("1.500")]
    assert mongo.document_to_row(mongo.row_to_document(["_id", "A"], [None, 1]), ["_id", "A"]) \
        == [None, 1]                                                     # 空 _id 交给服务端生成


def test_describe_mongo_error_is_chinese():
    text = mongo.describe_mongo_error(Exception("connection refused"), mongo.MongoConfig(host="h"))
    assert isinstance(text, str) and text


# ---------------------------------------------------------------------------
# 真实服务端(本机无 mongod 时自动跳过)
# ---------------------------------------------------------------------------
@pytest.fixture(scope="module")
def database():
    if not mongo.pymongo_available():
        pytest.skip("未安装 pymongo")
    cfg = mongo.MongoConfig(host="127.0.0.1", port=27017, database=TEST_DB, timeout=5)
    try:
        db = mongo.ADAPTER.connect(cfg)
    except Exception as exc:                                    # pragma: no cover
        pytest.skip(f"本机没有可用的 MongoDB:{exc}")
    yield db
    db.client.drop_database(TEST_DB)
    mongo.ADAPTER.close(db)


def test_server_metadata(database):
    assert mongo.ADAPTER.server_version(database).startswith("MongoDB")


def test_engine_csv_to_mongo_and_back(tmp_path, database):
    """CSV → MongoDB 集合 → CSV:长浮点数与文本逐字一致。"""
    source = tmp_path / "in.csv"
    core.write_csv_file(str(source), ["Name", "Amount"],
                        iter([["张三", LONG_VALUES[0]], ["李四", LONG_VALUES[1]]]))

    forward = core.convert(
        [core.SourceSpec(kind="CSV", path=str(source), parse=True)],
        core.TargetSpec(kind="Mongo", table="people", mode="replace"),
        conn=database, adapter=mongo.ADAPTER)
    assert forward.fidelity.ok and forward.rows == 2
    assert database["people"].count_documents({}) == 2

    back = tmp_path / "back.csv"
    core.convert([core.SourceSpec(kind="Mongo", table="people")],
                 core.TargetSpec(kind="CSV", path=str(back)),
                 conn=database, adapter=mongo.ADAPTER)
    lines = back.read_text(encoding="utf-8-sig").strip().splitlines()
    header = lines[0].split(",")
    assert header == ["_id", "Name", "Amount"]               # 集合默认带 _id 字段
    index = header.index("Amount")
    assert [line.split(",")[index] for line in lines[1:]] == [LONG_VALUES[0], LONG_VALUES[1]]


def test_engine_mongo_to_json_keeps_precision(tmp_path, database):
    database["nums"].drop()
    database["nums"].insert_many([{"_id": index, "Value": mongo.mongo_store_value(Decimal(value))}
                                  for index, value in enumerate(LONG_VALUES, start=1)])
    produced = tmp_path / "nums.json"
    core.convert([core.SourceSpec(kind="Mongo", table="nums")],
                 core.TargetSpec(kind="JSON", path=str(produced), key="nums"),
                 conn=database, adapter=mongo.ADAPTER)
    text = produced.read_text(encoding="utf-8")
    assert "e+" not in text and "E+" not in text
    header, rows = core.load_json_stream(str(produced), "nums")
    assert header == ["_id", "Value"]
    assert [row[1] for row in rows] == [Decimal(value) for value in LONG_VALUES]


def test_write_modes_and_schema_sampling(tmp_path, database):
    csv_path = tmp_path / "t.csv"
    core.write_csv_file(str(csv_path), ["A", "B"], iter([[1, "x"]]))

    core.convert([core.SourceSpec(kind="CSV", path=str(csv_path))],
                 core.TargetSpec(kind="Mongo", table="modes", mode="create_if_missing"),
                 conn=database, adapter=mongo.ADAPTER)
    core.convert([core.SourceSpec(kind="CSV", path=str(csv_path))],
                 core.TargetSpec(kind="Mongo", table="modes", mode="append"),
                 conn=database, adapter=mongo.ADAPTER)
    assert database["modes"].count_documents({}) == 2

    core.convert([core.SourceSpec(kind="CSV", path=str(csv_path))],
                 core.TargetSpec(kind="Mongo", table="modes", mode="replace"),
                 conn=database, adapter=mongo.ADAPTER)
    assert database["modes"].count_documents({}) == 1
    assert mongo.ADAPTER.list_columns(database, "modes") == ["_id", "A", "B"]
    assert "modes" in mongo.ADAPTER.list_tables(database)


def test_append_to_missing_collection_raises(tmp_path, database):
    csv_path = tmp_path / "t.csv"
    core.write_csv_file(str(csv_path), ["A"], iter([[1]]))
    with pytest.raises(ValueError, match="不存在"):
        core.convert([core.SourceSpec(kind="CSV", path=str(csv_path))],
                     core.TargetSpec(kind="Mongo", table="missing_one", mode="append"),
                     conn=database, adapter=mongo.ADAPTER)


def test_nested_document_roundtrip_through_files(tmp_path, database):
    """嵌套文档 ↔ JSON 文本单元格的双向转换。"""
    database["nested"].drop()
    database["nested"].insert_one({"_id": 1, "Name": "张三",
                                   "Device": {"ip": "192.168.1.59", "os": "Android"},
                                   "Tags": ["a", "b"]})
    produced = tmp_path / "nested.csv"
    core.convert([core.SourceSpec(kind="Mongo", table="nested")],
                 core.TargetSpec(kind="CSV", path=str(produced)),
                 conn=database, adapter=mongo.ADAPTER)
    text = produced.read_text(encoding="utf-8-sig")
    assert '""ip"": ""192.168.1.59""' in text or '"ip"' in text

    core.convert([core.SourceSpec(kind="CSV", path=str(produced), parse=True)],
                 core.TargetSpec(kind="Mongo", table="nested_back", mode="replace"),
                 conn=database, adapter=mongo.ADAPTER)
    restored = database["nested_back"].find_one({"_id": 1})
    assert restored["Device"] == {"ip": "192.168.1.59", "os": "Android"}
    assert restored["Tags"] == ["a", "b"]


def test_mongo_to_json_is_direct_document_mapping(tmp_path, database):
    """MongoDB → JSON 用"文档数组"直接映射:嵌套子文档/数组保持结构,不套 HeaderFields/Data。"""
    database["members"].drop()
    database["members"].insert_many([
        {"_id": 1, "name": "merry", "age": 22, "height": 158,
         "hobbies": [{"outDoor": "Bybike", "inDoor": "watchShortVideos"},
                     {"Eat": "cookingTomatoes"}]},
        {"_id": 2, "name": "bob", "age": 30, "height": 180, "hobbies": []},
    ])
    produced = tmp_path / "members.json"
    core.convert([core.SourceSpec(kind="Mongo", table="members")],
                 core.TargetSpec(kind="JSON", path=str(produced), key="members"),
                 conn=database, adapter=mongo.ADAPTER)

    payload = json.loads(produced.read_text(encoding="utf-8"))
    assert isinstance(payload["members"], list)
    first = payload["members"][0]
    assert first["name"] == "merry" and first["age"] == 22
    assert first["hobbies"][0]["outDoor"] == "Bybike"          # 嵌套结构原样保留
    assert payload["members"][1]["hobbies"] == []

    header, rows = core.load_json_stream(str(produced), "members")
    assert "hobbies" in header                                  # 回读仍可用
    assert list(rows)[0][header.index("name")] == "merry"


def test_datetime_roundtrip(tmp_path, database):
    moment = datetime(2026, 9, 28, 12, 30, 45)
    database["times"].drop()
    database["times"].insert_one({"_id": 1, "At": moment})
    produced = tmp_path / "times.json"
    core.convert([core.SourceSpec(kind="Mongo", table="times")],
                 core.TargetSpec(kind="JSON", path=str(produced), key="times"),
                 conn=database, adapter=mongo.ADAPTER)
    header, rows = core.load_json_stream(str(produced), "times")
    assert list(rows)[0][header.index("At")] == "2026-09-28 12:30:45"
