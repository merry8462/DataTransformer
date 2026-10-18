<div align="center">

<img src="images/img03.png" alt="DataTransformer main window" width="100%">

<p>MySQL / PostgreSQL / MongoDB ↔ xlsx / csv / json bidirectional data conversion tool (Windows GUI + Linux terminal wizard).</p>
<p>Long floats and high-precision values are kept exact end-to-end with <code>Decimal</code>; the output never contains scientific notation, and every write is verified by reading it back automatically.</p>

</div>

<p align="center">
<a href="README.zh.md">简体中文</a> | English
</p>


<div align="center">
<img src="images/img03.png" alt="DataTransformer main window, full view" width="100%">
</div>

From top to bottom:
Title and colour strip → **SSH 通道** → **① 输入配置** → **② 输出配置** → **④ 运行日志**→ status bar at the bottom.
The shot shows a **MySQL → CSV directory-mode** export that has just finished.


# Screenshots


## Database connection and SSH tunnel

<div align="center">
<img src="images/img02.png" alt="Database connection card and SSH tunnel card" width="100%">
</div>

The **数据库连接** (database connection) card supports MySQL / PostgreSQL / MongoDB; the「端口」("port") box is an
editable dropdown, **leave it empty and the default port of that engine is used automatically**
(MySQL 3306 / PostgreSQL 5432 / MongoDB 27017).

The **SSH 通道** (SSH tunnel) card: once it is ticked, the「主机 / 端口」("host / port") you fill in is the address
of the database **on the remote server** (usually `127.0.0.1:5432`), and local connections automatically go through
the tunnel port; while disconnected the status reads `SSH 通道:未连接`, once connected it reads
`127.0.0.1:<本地端口> → SSH 用户@主机:22 → 127.0.0.1:5432`.


## Connection success and tunnel log

<div align="center">
<img src="images/img01.png" alt="Connection success dialog and SSH tunnel log lines" width="100%">
</div>

After clicking「测试连接」("test connection") a dialog reports the server version (PostgreSQL 18.6), and the log area
prints
`SSH 通道已建立: 127.0.0.1:55960 → SSH …@172.24.208.28:22 → 127.0.0.1:5432` and
`连接成功: PostgreSQL …`


## Table import progress hints

<div align="center">
<img src="images/img04.png" alt="PostgreSQL over SSH to Excel export in progress" width="100%">
</div>

PostgreSQL (over an SSH tunnel) → Excel: the log prints once every 20,000 rows, and all input fields and action
buttons are greyed out automatically during the run, which both prevents a background task from reading a
"half-edited" configuration and prevents double clicks.


## One-million-row table export

<div align="center">
<img src="images/img05.png" alt="1,000,001 rows exported with fidelity check passed" width="100%">
</div>

`Excel 导出完成: D:\Downloads\virtual_data.xlsx（Sheet [VirtualProfile]，共 1,000,001 行）`;
after writing, the output is read back automatically for a double comparison of「数值等值 + 文本等值」("numeric
equality + text equality")
`数值保真校验：通过（字段 3 个 / 比对 575 个值）`
Million-row volumes stay stable: the server-side streaming cursor never reads the whole table into memory, and Excel
uses openpyxl's streaming mode.


## MongoDB → JSON

<div align="center">
<img src="images/img06.png" alt="Exporting a MongoDB collection to JSON" width="100%">
</div>

When MongoDB is selected the database card grows two more rows, **连接 URI** ("connection URI") and
**认证库 authSource** ("auth database"); the other types hide them automatically.

Here the connection goes over an SSH tunnel to `127.0.0.1:27017` with the auth database, and the collection is
exported to `MyData.json` (top-level key `[VirtualProfile]`, 4 rows in total); collection name ↔ JSON top-level key
and document field ↔ first header row map one to one.


## Nested documents / arrays round-trip

<div align="center">
<img src="images/img07.png" alt="mongosh documents next to the exported JSON" width="100%">
</div>

Left: the original documents in mongosh (`_id` as `ObjectId`, `Hobbies` as an array, `Birthday` / `CreateTime` as
`ISODate`). Right: the same collection exported to JSON — `_id` is written as 24-character hexadecimal text, nested
arrays are expanded as they are, and dates are written uniformly as `YYYY-MM-DD HH:MM:SS`. Importing back turns
`_id` into an `ObjectId` again, and JSON text cells are restored to nested structures automatically (keys containing
`.` or starting with `$` stay text, so write failures caused by MongoDB key restrictions are avoided).


## The Linux GUI

<div align="center">
<img src="images/img10.png" alt="The same GUI running on Linux" width="100%">
</div>

The same window runs a **MySQL → CSV directory-mode** export on Ubuntu: the directory name equals the database name,
each table becomes one `.csv` under it (here exported to `/home/.../Cty/VirtualData.csv`, 20,000 rows in total), and
the bottom status bar reads「已连接数据库 | 就绪」("database connected | ready").


## First start (automatic dependency install)

<div align="center">
<img src="images/img08.png" alt="data_transformer.sh creating a virtualenv on first run" width="70%">
</div>

When `./data_transformer.sh` detects that the system interpreter lacks the required dependency `openpyxl`, it asks
「是否在项目目录创建虚拟环境 .venv 并自动安装依赖」("create a virtualenv .venv in the project directory and install
the dependencies?"); pressing Enter creates `.venv` and installs
`openpyxl / pymysql / psycopg2-binary / pymongo / questionary / paramiko`; later runs reuse that virtualenv directly.


## GUI: launch the GUI once dependencies are ready

<div align="center">
<img src="images/img09.png" alt="data_transformer.sh --gui installing dependencies" width="70%">
</div>

`./data_transformer.sh --gui` goes through the same dependency self-check and starts the GUI once the dependencies are
ready.
Note: **the command-line wizard itself does not need PySide6**, so on a server without a desktop environment the
default `./data_transformer.sh` is enough.


## Terminal wizard: pick a format → connect to MySQL or PostgreSQL or MongoDB → hierarchical picker of database / table / collection / field

<div align="center">
<img src="images/img11.png" alt="Terminal wizard selecting the input format and MongoDB parameters" width="100%">
</div>

Left: the terminal wizard of `./data_transformer.sh` — ① choose the input format (MySQL / PostgreSQL / MongoDB /
Xlsx / Json / Csv) → `[SSH Tunnel]` switch → one prompt per parameter (Host / Port / Connection URI / Username /
Password / Database / authSource / Timeout) → `已连接 MongoDB 9.0.2` → ② **hierarchical selection**: level 1 picks
collections, level 2 ticks the fields of each collection individually.
Right: the real documents of `MongoData.MongoCollections` in mongosh at that very moment, for comparison with the
exported result.


## Terminal wizard: JSON export finished

<div align="center">
<img src="images/img12.png" alt="Terminal wizard finishing a MongoDB to JSON export" width="100%">
</div>

③ Choose the output format `Json` → export file path → whether to rename tables (default: same name as the source) →
**conversion summary** → confirm
→ `JSON 导出完成: …/MongoData.json（键 [MongoData]，共 20,000 行）`
→ `数值保真校验：通过（字段 1 个 / 比对 200 个值）` → `已记住本次连接参数`.
Right: the tail of the exported JSON — `_id` / dates / the nested `Hobbies` array are all presented by the rules of
the previous figure.


# Support matrix

| in \ out | xlsx | json | csv | MySQL | PostgreSQL | MongoDB |
| :---: | :---: | :---: | :---: | :---: | :---: | :---: |
| **xlsx** | — | ✅ | ✅ | ✅ hierarchical import | ✅ hierarchical import | ✅ hierarchical import |
| **json** | ✅ | — | ✅ | ✅ create table on import | ✅ create table on import | ✅ import into collection |
| **csv** | ✅ | ✅ | — | ✅ create table on import | ✅ create table on import | ✅ import into collection |
| **MySQL** | ✅ | ✅ | ✅ directory mode | ✅ table copy | ✅ table copy | ✅ import into collection |
| **PostgreSQL** | ✅ | ✅ | ✅ directory mode | ✅ table copy | ✅ table copy | ✅ import into collection |
| **MongoDB** | ✅ | ✅ | ✅ directory mode | — | — | ✅ collection copy |

> Database-to-database conversion happens inside one server (same engine → direct table / collection copy); to cross
> database types, go through a file (xlsx / json / csv) first.


# Project structure

```text
DataTransformer/
├── data_transformer.py    Windows GUI entry point (PySide6 + Nuitka packaging target)
├── src/                   Engine package (modules import each other relatively)
│   ├── __init__.py        Package marker
│   ├── dt_cli.py          Linux / macOS interactive terminal entry (questionary dynamic choices)
│   ├── dt_core.py         Core conversion engine (Reader/Writer pipeline, specs, SSH, self-check) + unified facade
│   ├── dt_config.py       Constants / connection config (DBConfig, SSHConfig) / remembered settings (Settings)
│   ├── dt_files.py        File format layer (xlsx / csv / json read-write and value conversion)
│   ├── dt_sql.py          Database layer (MySQL / PostgreSQL: streaming read, batch write, type inference, free queries)
│   ├── dt_mongo.py        MongoDB layer (collection listing, schema sampling, Decimal128 read-write, find parsing)
│   └── dt_numeric.py      High-precision number fidelity layer (Decimal parsing / fixed-point formatting / reports)
├── tests/                 pytest unit and integration tests (including the long-float suite)
│   ├── test_numeric.py    Numeric fidelity: trailing zeros, huge integers, scientific notation, JSON/Excel fixed-point output
│   ├── test_engine.py     Conversion engine: mapping rules, multi-table output, skip/directory modes, check blocking
│   ├── test_mongo.py      MongoDB value conversion and real-server round-trips (skipped automatically without a server)
│   ├── test_cli.py        End-to-end command-line wizard interaction (driven through a pipe)
│   ├── test_gui.py        Key GUI logic (offscreen)
│   └── test_query.py      Custom query statements: read-only classification, free-statement reading, Mongo find parsing, preview
├── data_transformer.sh    Linux / macOS launcher (detects Python, creates .venv on demand, then runs python -m src.dt_cli)
├── build_exe.bat          Windows packaging script (Nuitka onefile)
├── run.bat                Run from source on Windows
├── conftest.py            pytest root configuration (puts the project root on sys.path)
├── requirements.txt       Dependency list
├── README.zh.md           Chinese documentation
└── README.en.md           English documentation
```


# Dependencies

| Purpose | Dependency | Notes |
| :--- | :--- | :--- |
| Runtime | Python **3.10+** | uses `decimal`, dataclasses, the walrus operator, … |
| GUI | PySide6 ≥ 6.6 | only needed for the Windows GUI |
| Terminal interaction | questionary ≥ 2.0 | only needed for the Linux / macOS command-line wizard |
| MySQL | PyMySQL ≥ 1.1 | server-side streaming cursor (SSCursor) |
| PostgreSQL | psycopg2-binary ≥ 2.9 | server-side named cursor (named cursor) |
| MongoDB | pymongo ≥ 4.6 | exact numeric read-write with Decimal128 |
| Spreadsheet files | openpyxl ≥ 3.1 | read_only / write_only streaming modes |
| SSH tunnel | paramiko ≥ 3.4 (+ cryptography) | local port forwarding |
| Packaging | Nuitka | only needed for packaging |
| Tests | pytest ≥ 8.0 | optional |


# Download

Three ways are provided for now: the Windows portable build, running the sources with a Windows batch file, and the
Linux / macOS command-line wizard.


## Windows

**System requirements:**

- Windows 10 or newer
- 64-bit (x86_64) systems only

- **Portable**: `DataTransformer.exe` is a Nuitka-packaged single-file program, no Python installation needed — just
  double-click it after downloading.
- **Batch run**: `run.bat` runs the sources with the local Python interpreter, which suits development and debugging.


## Linux / macOS

**System requirements:**

- Any mainstream distribution will do: Debian / Ubuntu / CentOS / Rocky / Fedora, plus macOS
- Python 3.10+ (the launcher probes `python3` / `python3.x` / `python` automatically)

**Usage:**

```bash
chmod +x data_transformer.sh   # first use: grant execute permission
./data_transformer.sh          # start the interactive command-line wizard
```

It opens an interactive command window; choose in order following the prompts:

```text
① choose the input format: MySQL / PostgreSQL / Xlsx / Json / Csv
② choose the data table
③ choose the output format and write mode
④ choose the export path (file or directory)
⑤ run the conversion and print the log
```

**The launcher does this automatically:**

- detects an available Python3 interpreter;
- checks the dependencies and, when something is missing, asks whether to create `.venv` in the project directory and
  install them automatically;
- reports the install status of `pymysql` / `psycopg2` / `paramiko` (they only affect the matching database / SSH
  feature; Excel / JSON / CSV conversion always works).

**Common arguments:**

```bash
./data_transformer.sh --check      # check the runtime environment only, do not enter the wizard
./data_transformer.sh --demo       # prefill the SSH + PostgreSQL sample parameters
./data_transformer.sh --selftest   # headless self-test, result written to selftest.log
./data_transformer.sh --gui        # start the GUI (needs PySide6)
./data_transformer.sh --help       # all arguments
```

> The command-line wizard on Linux does **not** need PySide6, so it also works on a server without a desktop
> environment.


## Installing dependencies per distribution

The script probes by itself and can create `.venv` in the project directory; if you install manually, use the table
below:

| Distribution | Install Python and venv | Install dependencies |
| :--- | :--- | :--- |
| Ubuntu / Debian | `sudo apt update && sudo apt install -y python3 python3-venv python3-pip` | `python3 -m venv .venv && .venv/bin/pip install -r requirements.txt` |
| CentOS / Rocky | `sudo dnf install -y python3 python3-pip` | as above (on CentOS 7 run `sudo yum install -y centos-release-scl` first) |
| Fedora | `sudo dnf install -y python3 python3-pip` | as above |
| macOS | `brew install python` | as above |

> psycopg2 needs build dependencies on a few distributions: on Debian / Ubuntu use
> `sudo apt install -y libpq-dev`,
>
> quick install of the core dependencies without a virtualenv (no PySide6):
>
> ```bash
> python3 -m pip install --user openpyxl pymysql psycopg2-binary pymongo questionary paramiko
> ```


# Running from source

Clone the project first:

```bash
git clone https://github.com/merry8462/DataTransformer.git
cd DataTransformer
```

Install the dependencies:

```bash
pip install -r requirements.txt
```

> ```bash
> pip install PySide6 -i 镜像源
> ```

**Start the main program:**

```bash
python data_transformer.py
```

or double-click `run.bat`.

On Linux / macOS you can also skip the GUI and run the wizard directly (no PySide6 needed):

```bash
pip install openpyxl pymysql psycopg2-binary pymongo questionary paramiko   # core dependencies only
python -m src.dt_cli                                       # same as ./data_transformer.sh
python -m src.dt_cli --demo                                # with the Demo sample parameters
python -m src.dt_cli --plain                               # force plain-text prompts
python -m src.dt_cli --selftest                            # headless self-test
```


Double-click `build_exe.bat`

```powershell
build_exe.bat
```

Packaged artifact:
```text
build\DataTransformer.exe
```

**Full packaging command**:

```powershell
python -m nuitka --standalone --onefile --enable-plugin=pyside6 ^
  --windows-console-mode=attach --nofollow-import-to=pandas ^
  --include-package=pymysql --include-package=psycopg2 --include-package-data=psycopg2 ^
  --include-package=pymongo --include-package=bson ^
  --include-package=paramiko --include-package=cryptography ^
  --include-package=et_xmlfile --include-package=src ^
  --output-dir=build --output-filename=DataTransformer data_transformer.py
```

# Features

MySQL / PostgreSQL / MongoDB and Excel / JSON / CSV convert into each other in any direction:

| in \ out | Excel | JSON | CSV | SQL database | MongoDB |
| :---: | :---: | :---: | :---: | :---: | :---: |
| SQL | ✅ | ✅ | ✅ directory mode | ✅ table copy | ✅ import into collection |
| MongoDB | ✅ | ✅ | ✅ directory mode | — | ✅ collection copy |
| Excel | ✅ | ✅ | ✅ | ✅ hierarchical import | ✅ hierarchical import |
| JSON | ✅ | ✅ | ✅ | ✅ create table on import | ✅ import into collection |
| CSV | ✅ | ✅ | ✅ | ✅ create table on import | ✅ import into collection |

Main features:

- **Hierarchical selection**: a two-level tree dropdown — level 1 multi-selects data tables / collections (or Excel
  sheets), level 2 ticks the fields of each table independently;
- **Custom queries**: an option that sits next to the hierarchical picker — write a SQL statement directly for a SQL
  database, or a `db.<集合>.find(<filter>)` chain for MongoDB; the header and sample rows are validated and
  previewed first, then it becomes a data source;
- **Long-float fidelity**: everything travels through `Decimal`, trailing zeros are not lost, significant digits are
  not truncated, **the output never contains scientific notation**, and after writing the output is read back for a
  double check of 「数值等值 + 文本等值」("numeric equality + text equality") with a difference report;
- **Field mapping**: automatic mapping by identical name (sheet name ↔ table name, first header row ↔ database column
  names), with manual overrides and duplicate-name validation;
- **SSH tunnel**: reach MySQL / PostgreSQL / MongoDB on a server with a specific IP without opening the database port
  on that server;
- **Command-line wizard**: `./data_transformer.sh` picks the Input/Output format, connection parameters, tables /
  collections and fields dynamically through **questionary**;
- **Excel → SQL hierarchical import**: the target table name mirrors the sheet name one to one; a missing table is
  created automatically from the sheet name, and for existing tables you choose per sheet between **overwrite /
  append / skip that table**;
- **CSV directory-mode export**: the directory name equals the database name and each table becomes one
  `表名.csv` ("<table name>.csv") under it;
- **Duplicate-file handling**: when a file with the same name is detected, choose overwrite the whole file / merge into
  it / cancel;
- **Large volumes with low memory**: server-side streaming cursors plus batched inserts; million-row imports and
  exports are stable in practice;
- **Automatic table creation**: field types are inferred from the first 200 sampled rows (`int → BIGINT/INT`,
  `Decimal → NUMERIC/DECIMAL`, `float → DOUBLE`, date/time → `TIMESTAMP/DATE`, strings are always `TEXT`);
- **MongoDB nested structures**: nested documents / arrays are exported as JSON text cells and restored to nested
  structures automatically on import;
- **Remembered connection settings**: passwords, database names and SSH parameters are remembered and filled in
  automatically next time;
- **Friendly connection errors**: a wrong host / port / user name / password / database name is recognised
  automatically with a localised Chinese hint (the raw error is kept as well);
- **Button state coupling**: while input / output are not ready, the database information is incomplete, no connection
  exists or a task is running, the related buttons are greyed out automatically against repeated operations;
- **Bottom button area**: in the bottom-right corner the three buttons read **初始化 → 清空日志 → 开始转换 ▶**
  ("reset → clear log → start conversion ▶") from top to bottom, all three use
  `QSizePolicy.Preferred, Expanding` and fill that column vertically (there is no stretch at the end of the column),
  so maximising the window stretches the buttons with the column and leaves no blank space underneath;
- **One-click self-test**: built-in `--selftest` validates the drivers, the engine, file conversion and numeric
  fidelity;
- **Injection safety**: table / field names are whitelisted and quoted (MySQL backticks / PostgreSQL double quotes).


## Naming conventions

Database – table – field correspond one to one across the three levels:

| Level | Correspondence | Example |
| :---: | :--- | :--- |
| Database name | Excel/JSON file name == CSV directory name == MongoDB database | `xxx.xlsx` / `xxx.json` / `xxx/` |
| Table name | database table name == MongoDB collection name == Excel sheet name == JSON top-level key == CSV file name | `"info"` → `xxx/info.csv` |
| Field name | database field name == MongoDB document field == first Excel row cell == JSON HeaderFields == CSV header | `"Id"`, `"Name"`, `"CreateTime"` |

> Mapping is **automatic by identical name** by default, and can be overridden manually in the GUI's
> 「字段映射预览 / 调整」("field mapping preview / adjust") or in the command-line wizard
> (for example map `姓名` to `Name`); duplicate field names after mapping are intercepted with a hint before writing.

JSON uses a **header + row records** structure.

```json
{
    "info": {
        "HeaderFields": ["Id", "Name", "CreateTime"],
        "Data": {
            "Row1": [1, "张三", "2026-08-29 21:32:35"],
            "Row2": [2, "李四", "2026-08-29 21:32:36"]
        }
    }
}
```

> `HeaderFields` lists the field names and `RowN` holds row N (rows are numbered from 1); the older **column-array** layout (field names as keys, whole columns as arrays) is still readable, so old files keep working.


## GUI operation steps

1. The card at the very top of the window is 【SSH 通道】(see figure 2); tick **启用 SSH 通道** ("enable SSH
   tunnel");
2. Fill in the SSH host / IP, port (default 22), user name and password (this tool always authenticates with
   「用户名 + 密码」("user name + password"); the command-line wizard / engine still keep the private-key parameter,
   the GUI does not offer private-key selection);
3. Click **连接 SSH**; the status area shows `127.0.0.1:<本地端口> → SSH 用户@主机:22 → 127.0.0.1:5432`
   (see figure 3);
4. Go back to **数据库连接** and fill in the address / port / user name / password / database name of the remote
   database, then click **连接并加载表**;
5. Afterwards tick「表 → 字段」as in the normal flow and export or import (for large-table progress see figure 4 and
   figure 5);
6. To change the database type simply switch **数据库类型**: **the tunnel is not dropped**, the program only points
   the forward target at the new「主机:端口」and the log says
   `SSH 通道已重新指向 127.0.0.1:5432` (the port is whatever you actually filled in); then click
   **连接并加载表** again.

> The tunnel only listens on the local `127.0.0.1` and never exposes a port to the outside network; clicking
> **断开 SSH** or closing the program releases the tunnel automatically.

> **Switching the database type does not drop the tunnel**: the old database connection becomes invalid on switching
> (you need to "connect and load tables" again), but the SSH tunnel is kept as it is —
> `SSHTunnel.retarget` on the engine side only resets the forward target, so the established local listening port is
> unaffected; and every time the connection configuration is fetched (connect, read, preview query) the tunnel target
> is first synced to the currently entered「主机:端口」, falling back to that engine's default port when the port is
> empty or out of range.

> **Tunnel liveness check**: the GUI probes once every 3 seconds; as soon as the tunnel is found broken the log says
> so **only once** — `SSH 通道已断开:<通道信息>,请重新连接` — and then the database connection is cleaned up
> automatically and the status bar and button states are refreshed; the probe only runs while a tunnel exists, so it
> never floods the log repeatedly. Click **连接 SSH** to rebuild the tunnel and continue.


## Command-line wizard

When asking for connection parameters `./data_transformer.sh` first asks
「是否通过 SSH 通道连接远程服务器」("connect to the remote server through an SSH tunnel?"), the interaction matches
the GUI; adding `--demo` prefills the SSH + PostgreSQL sample parameters from `Demo/Example.txt` directly:

```bash
./data_transformer.sh --demo
```

> SSH support needs `paramiko` (`pip install paramiko`). Without it the button is greyed out with a hint, and file
> conversion is unaffected.


# Command-line arguments

DataTransformer supports the following command-line arguments, which help with debugging and troubleshooting.

The arguments are used like this:

**From source:**

```bash
python data_transformer.py 参数
```

**Compiled build:**

```powershell
DataTransformer.exe 参数
```

---

## `--selftest`

Headless self-test mode, used to verify that the drivers and the engine are complete inside a packaged artifact; the
result is written to `selftest.log` in the working directory.

Example:

```powershell
DataTransformer.exe --selftest
```

Together with the database environment variables the MySQL / PostgreSQL connectivity can be self-tested as well:

```powershell
set SELFTEST_MYSQL_USER=root
set SELFTEST_MYSQL_PASSWORD=你的密码
set SELFTEST_MYSQL_DATABASE=你的库名
set SELFTEST_PG_USER=你的用户名
set SELFTEST_PG_PASSWORD=你的密码
set SELFTEST_PG_DATABASE=你的数据库名

DataTransformer.exe --selftest
```

> Without those environment variables only the dependency imports and the Excel / JSON / CSV write-read round-trips
> are checked.

---

## `--version`

Prints the version number and exits, which is convenient for scripted calls:

```powershell
DataTransformer.exe --version
```

The equivalent on Linux / macOS:

```bash
./data_transformer.sh --version
```

---


# Logging and debugging

The program run log is shown in the "④ 运行日志" area of the main window; it is colour-coded by log level and
supports **one-click clearing**.

```powershell
set DT_DEBUG=1
DataTransformer.exe --selftest
```

```bash
DT_DEBUG=1 ./data_transformer.sh
```

> When `DT_DEBUG` is present the GUI appends `traceback.format_exc()` to the log; the command-line wizard re-raises
> the exception directly in its last-resort handler.
> `DT_DEBUG` is off by default, so normal use only shows that Chinese hint.

If you hit an import or export error, please keep the log content and provide the relevant log when reporting the
problem or opening an issue.


# Configuration and remembered settings

After a successful connection the program remembers the connection parameters automatically (the Windows GUI
remembers the **password** and the **database name**; the command-line wizard additionally remembers the **host /
port / user name / MongoDB URI / authSource / SSH parameters**), and fills them in automatically next time.

The configuration file lives at:

```text
Windows      : %APPDATA%\DataTransformer\config.json
Linux / macOS: $XDG_CONFIG_HOME/DataTransformer/config.json (default ~/.config/DataTransformer/config.json)
```

Delete that file and restart the program to get the first-start state back (only the host defaults to `127.0.0.1`;
port, user name, password and database name are all empty).

> Passwords are stored **in plain text** in the configuration file, so do not use sensitive passwords on a computer
> shared by several users.
> On Linux / macOS the permissions are tightened to `600` (owner read-write only) when writing the configuration, and
> a "write a temporary file first, then replace atomically" strategy avoids a half-written configuration.


# Security notes

- **Identifier validation**: every table / field name is whitelisted first (must start with `字母 / 下划线 / 中文`,
  i.e. a letter, an underscore or a Chinese character, and may only contain letters, digits and underscores), then
  quoted per database dialect (MySQL backticks / PostgreSQL double quotes), which prevents SQL injection through
  string concatenation.
  As a consequence, Excel sheet names and header column names containing spaces, brackets, dashes and the like must be
  renamed before importing into a database (exporting to JSON / CSV / Excel has no such restriction);
- **Custom query statements**: the statement is written by the user and executed verbatim, with no identifier
  whitelist validation and no quoting — this is the only entry point that bypasses the validation above; a
  non-read-only statement modifies the database and cannot be rolled back, so only use it on statements you trust;
- **SSH tunnel**: the local listening address is fixed to `127.0.0.1` and no port is exposed to the LAN;
- **Configuration permissions**: on Linux / macOS the configuration file permissions are tightened to `600`;
- **CSV formula injection**: exported CSV / Excel files **keep the data content as it is**. If the source data contains
  text starting with `=`, `+`, `-` or `@`, opening the file directly in Excel may execute it as a formula — that is
  Excel's existing behaviour. When handling untrusted data, inspect it in a text editor first, or export to JSON
  instead.


# FAQ

| Symptom | Fix |
| :--- | :--- |
| `ModuleNotFoundError: No module named 'questionary'` | The Linux wizard needs questionary: `pip install questionary`, or simply use `./data_transformer.sh` (it installs it automatically) |
| `ModuleNotFoundError: No module named 'pymongo'` | Install it when you need MongoDB features: `pip install pymongo`; ignore it if you do not use MongoDB |
| The hint「非法标识符」("illegal identifier") appears | Table / field names (including Excel sheet names and headers) must start with a letter, underscore or Chinese character and may only contain letters, digits and underscores; rename names containing spaces / brackets / dashes first, or switch to JSON / CSV / Excel output |
| Opening the exported file in Excel distorts long numbers or shows scientific notation | The long values this tool writes are **text cells** (Excel cannot represent more than 15 significant digits exactly); use JSON / CSV to keep numeric semantics, or format that column as text in Excel before importing |
| `数值保真校验未通过` ("numeric fidelity check failed") | It means precision loss / trailing-zero loss / scientific notation happened on the output side; the log and the dialog list sample differences — check the target column type (it should be NUMERIC / DECIMAL rather than DOUBLE) or switch to a text column |
| MySQL/PostgreSQL import reports `value too long` | The string column of the target table is too short; automatic table creation already uses TEXT uniformly, so widen that column for hand-made tables |
| Chinese text is garbled | Run `chcp 65001` in the Windows terminal first; choose `UTF-8 with BOM` or `GBK` as the CSV encoding |
| `无法连接 MongoDB ... ServerSelectionTimeoutError` | Check whether mongod is running, the port and `authSource`; tick the SSH tunnel for remote scenarios |
| Million-row imports are slow | Increase 「每批行数」("rows per batch") (default 1000, 5000~20000 is possible); on the database side mind `max_allowed_packet` (MySQL) and disk I/O |
| After an error the log only holds one Chinese hint — how do I locate the exact code? | Neither the traceback nor source line numbers are printed by default; set the environment variable `DT_DEBUG=1` before launching (`set DT_DEBUG=1` on Windows, `DT_DEBUG=1 ./data_transformer.sh` on Linux / macOS) and run again — the log then appends the full `traceback` |
| The log shows `SSH 通道已断开:…,请重新连接` | The 3-second liveness probe found the tunnel gone; the program has already cleaned up the database connection and refreshed the buttons — click **连接 SSH** to rebuild the tunnel, then "connect and load tables" again (switching the database type itself never drops the tunnel) |


# Acceptance checklist

- [x] Windows 10/11 GUI (PySide6): Input parameter area, Output parameter area, run log + start conversion button,
      progress bar;
- [x] Windows single-file executable packaged with Nuitka (`build_exe.bat`, validates size and SHA-256 automatically);
- [x] Linux / macOS interactive terminal tool (`data_transformer.sh` → `python -m src.dt_cli`), **questionary**
      dynamic selection of format and parameters, no GUI required;
- [x] Dependency installation and run instructions covering Ubuntu, Debian, Fedora and CentOS;
- [x] MySQL, PostgreSQL and **MongoDB** support, plus xlsx, csv and json files;
- [x] Bidirectional database ↔ file conversion (including SQL→SQL table copy, MongoDB collection copy and file-to-file
      conversion);
- [x] Mapping rules: xlsx `sheetname` ↔ database table name / MongoDB collection name; first header row ↔ database
      field names; CSV / JSON offer the equivalent mapping with manual overrides;
- [x] Long-float and high-precision fidelity: numeric equality + text identity, no lost trailing zeros, no truncated
      significant digits, no `e`/`E` scientific notation;
- [x] A visible numeric verification mechanism with a difference report (read back after writing, block or confirm to
      let through);
- [x] Data type conversion, NULLs, date/time, encoding, batch writes, error handling, logging and progress hints;
- [x] Unit tests and integration tests (including the long-float suite and real database round-trips);
- [x] Free-form SQL / MongoDB query statements with validation and preview first (the GUI "自定义查询…" dialog and the
      terminal wizard "② 自由填写查询语句");
- [x] The SQL / MongoDB examples built into the custom query dialog can be used directly (the MongoDB example uses
      mongosh syntax: bare keys, single quotes and trailing commas all work);
- [x] Switching the database type never drops the SSH tunnel (only the forward target is reset), and when the tunnel
      breaks unexpectedly the log says so only once while the connection and button states are cleaned up
      automatically;


# AI-assisted development

This project was developed with DeepSeek Harness assisting the programming (Vibe Coding) during development,
including writing code, refactoring, debugging and problem analysis.

The overall design, feature planning, code review and final maintenance of the project are the author's
responsibility.


# Credits

DataTransformer mainly uses the following projects:

- [Python](https://www.python.org/)
- [PySide6](https://doc.qt.io/qtforpython-6/)
- [questionary](https://questionary.readthedocs.io/) (dynamic terminal interaction)
- [PyMySQL](https://github.com/PyMySQL/PyMySQL)
- [psycopg2](https://www.psycopg.org/)
- [pymongo](https://pymongo.readthedocs.io/) (MongoDB)
- [openpyxl](https://openpyxl.readthedocs.io/)
- [paramiko](https://www.paramiko.org/) (SSH tunnel)
- [Nuitka](https://nuitka.net/)



Copyright © 2026 merry8462.
