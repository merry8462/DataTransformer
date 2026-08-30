# SQL / Excel / JSON / CSV 数据转换工具

基于 Tkinter 的图形化工具,支持 **MySQL / PostgreSQL** 与 **Excel(.xlsx)/ JSON(.json)/ CSV(.csv)** 任意互转,并可用 Nuitka 打包为独立 exe。

## 文件说明

| 文件 | 说明 |
| --- | --- |
| `data_transformer.py` | 主程序(Tkinter GUI,彩色主题) |
| `requirements.txt` | 运行依赖 |
| `run.bat` | 直接用 Python 运行 |
| `build_exe.bat` | Nuitka 一键打包脚本(输出 `build\DataTransformer.exe`) |

## 命名约定(库-表-字段三层一一对应)

| 层级 | 对应关系 | 示例 |
| --- | --- | --- |
| 数据库名 | == 输出文件名(Excel / JSON / CSV) | `db0829.xlsx` / `db0829.json` / `db0829.csv` |
| 表名 | == Excel Sheet 名 == JSON 一级参数 | `"info"` |
| 字段名 | == Excel 首行单元格 == JSON 二级参数 == CSV 表头 | `"Id"`、`"Name"`、`"CreateTime"` … |

JSON 采用**列数组**结构:

```json
{
    "info": {
        "Id": [1, 2, 3],
        "Name": ["迟柳", "张浩", "陈柳"],
        "CreateTime": ["2026-08-29 21:32:35"]
    }
}
```

## 功能:输入 × 输出全组合

顶部模式栏选择 **① 输入格式** 与 **② 输出格式**,点击 **开始转换 ▶**:

| 输入 \ 输出 | Excel | JSON | CSV | SQL |
| --- | :-: | :-: | :-: | :-: |
| SQL(勾选表 + 勾选字段) | ✅ | ✅ | ✅ | ✅ 表拷贝 |
| Excel | ✅ | ✅ | ✅ | ✅ 导入建表 |
| JSON | ✅ | ✅ | ✅ | ✅ 导入建表 |
| CSV | ✅ | ✅ | ✅ | ✅ 导入建表 |

- 数据库连接面板在**输入或输出为 SQL 时自动展开**,否则自动隐藏;
- 输入为文件时点击 **选择文件...** 打开选择窗口;输出为文件时点转换后弹出保存窗口;
- 输入/输出面板分别以蓝(SQL)、绿(Excel)、橙(JSON)、紫(CSV)着色区分。

### ① SQL 输入:自由勾选 TableName + ColumnName(支持多表)

1. 输入格式选 **SQL 数据库**,填写连接信息后点 **连接并加载表**;
2. 点 **选择数据表** 下拉框,**可多选勾选**要导出的表:
   - 勾选 **一张表**:自动加载字段,点 **选择字段** 下拉框多选勾选要导出的字段(默认全选,支持全选/全不选);
   - 勾选 **多张表**:字段选择自动禁用,每张表导出全部字段;
3. 输出与多表的对应关系:
   - **Excel**:一个工作簿,每张表一个 **Sheet(Sheet 名 = 表名)**;
   - **JSON**:一个文件,每张表一个 **一级键(键名 = 表名)**;
   - **CSV**:选择一个输出目录,每张表生成一个 `表名.csv` 文件;
   - **SQL**:多表同名表拷贝,目标表名请留空(按源表同名建表);
4. 开始转换。

- 导出使用服务端流式游标(MySQL `SSCursor` / PostgreSQL 命名游标),大数据量不占客户端内存;
- 表拷贝(SQL→SQL)使用独立输出连接,避免流式游标与建表 DDL 冲突;
- 首次启动时仅主机默认为 `127.0.0.1`,端口/用户名/密码/数据库名均为空(端口留空自动用 3306/5432),
  密码与数据库名在连接成功后会被记忆并自动填入。

### ② 文件输入

- **Excel**:选择 `.xlsx` 后自动列出工作表供选择;
- **JSON**:选择 `.json` 后自动列出一级键(表名)供选择;
- **CSV**:支持分隔符(逗号 / 分号 / TAB / 竖线)与编码(UTF-8 BOM / UTF-8 / GBK)选择。

### ③ 输出为文件:同名文件纠错

保存时若检测到**同名文件已存在**,弹出纠错对话框:

| 处理方式 | Excel | JSON | CSV |
| --- | --- | --- | --- |
| **覆盖整个文件** | 重建工作簿 | 重建文件 | 重建文件(含表头) |
| **合并写入** | 保留其他 Sheet,覆盖同名 Sheet | 保留其他一级键,覆盖同名键 | 追加数据行(不重复表头) |
| **取消** | 不写入 | 不写入 | 不写入 |

### ④ 输出为 SQL

- **追加到已有表 (append)** / **不存在则创建,存在则追加 (create if missing)** / **删除重建 (replace)**;
- 建表时按前 200 行自动推断字段类型;CSV 导入会自动把短整数、小数、日期字符串解析为对应类型;
- 空行跳过、空单元格写 `NULL`;表名/字段名白名单校验 + 引用符防注入。

### ⑤ 其他小功能

- **一键清空日志**:运行日志子窗口右上角提供"一键清空"按钮;
- **连接配置记忆**:连接成功后自动记住**密码**与**数据库名**(同时保留在退出时),下次启动自动填入;
  配置文件位于 `%APPDATA%\DataTransformer\config.json`,注意其中密码为明文存储,请勿在多用户共用电脑上使用敏感密码。

## 环境与运行

```bat
pip install -r requirements.txt
python data_transformer.py
```

## Nuitka 打包 exe

前置条件:

1. Python 3.8+ 及 pip;
2. 已安装 C 编译器:**Visual Studio Build Tools(C++ 桌面开发工作负载)** 或 MinGW64;
3. 网络可访问 pip 源(脚本会自动安装 nuitka 等依赖)。

执行:

```bat
build_exe.bat
```

成功后生成 `build\DataTransformer.exe`(单文件,双击即用,无需安装 Python)。

> 若 onefile 打包失败(个别杀软误报或环境问题),可删除 `build_exe.bat` 中 `--onefile` 一行,
> 改为目录模式,产物在 `build\DataTransformer.dist\DataTransformer.exe`,整个文件夹一起分发。

### 打包后自检(可选)

exe 内置 `--selftest` 无界面自检模式:

```bat
set SELFTEST_MYSQL_USER=root
set SELFTEST_MYSQL_PASSWORD=你的密码
set SELFTEST_MYSQL_DATABASE=你的库名
set SELFTEST_PG_USER=postgres
set SELFTEST_PG_PASSWORD=你的密码
set SELFTEST_PG_DATABASE=postgres

DataTransformer.exe --selftest
```

结果写入运行目录下的 `selftest.log`(不设置环境变量时只检查依赖导入与 Excel/JSON/CSV 读写)。

## 技术选型说明

- **数据库**:MySQL 用 `PyMySQL`,PostgreSQL 用 `psycopg2`;两者占位符均为 `%s`,统一 `executemany` 批量写入;
- **Excel**:openpyxl `read_only / write_only` 流式模式,比 pandas(底层同样调 openpyxl 且整表物化)更省内存;
- **JSON**:标准库 `json`,`ensure_ascii=False` 保留中文,无额外依赖;
- **CSV**:标准库 `csv`,可选分隔符与编码;
- **UI**:ttk `clam` 主题 + 自定义彩色样式(标题横幅、分区卡片、状态胶囊、彩色按钮)。

## 注意事项

- JSON 输出为列数组结构,需要把整表暂存内存;千万行级数据建议输出 Excel / CSV 以保持流式低内存;
- 类型推断按前 200 行采样,若后续行类型差异较大,请先手工建表再用 **append** 模式;
- 大批量导入建议适当调大 **每批行数**(默认 1000),一般 1000~5000 效率最佳;
- CSV 长数字(身份证号、手机号等)会按字符串处理,不会丢失前导零;
- PostgreSQL `json/jsonb` 字段导出 **Excel / CSV** 时自动转为 JSON 文本(如
  `{"ip": "192.168.1.59", "os": "XiaomiNote8"}`);导出 **JSON** 时保留原始嵌套结构;
  SQL→SQL 表拷贝时以 JSON 文本入库。
