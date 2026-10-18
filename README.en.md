<div align="center">

<img src="images/img03.png" alt="DataTransformer main window" width="100%">

<p>MySQL / PostgreSQL / MongoDB ↔ xlsx / csv / json bidirectional data converter (Windows GUI + Linux terminal wizard).</p>
<p>Long floats and high-precision numbers are kept exact end-to-end with <code>Decimal</code>; output never uses
scientific notation, and every write is verified by reading the result back.</p>

</div>

<p align="center">
<a href="README.zh.md">简体中文</a> | English
</p>


# Screenshots

> All screenshots below come from real runs of v1.1.6 (historical captures): the Windows GUI, the Linux GUI and the
> Linux terminal wizard. The current version is v1.1.7; each caption names the feature it demonstrates, so you can
> read the docs by picture.


## Figure 1 · The main window

<div align="center">
<img src="images/img03.png" alt="DataTransformer main window" width="100%">
</div>

From top to bottom: title and colour strip → **SSH tunnel** → **① Input** → **② Output** → **④ Run log** → status bar.
The shot was taken right after a successful **MySQL → CSV directory-mode** export: the log shows
`CSV 导出完成: …\Cty\VirtualData.csv（共 20,000 行）` ("CSV export finished, 20,000 rows") and
`数值保真校验：通过（字段 3 个 / 比对 586 个值）` ("fidelity check passed, 3 fields / 586 values compared"),
with the result summary in the dialog. The run buttons (and every input field) are greyed out while a task runs.


## Figure 2 · Database connection and SSH tunnel

<div align="center">
<img src="images/img02.png" alt="Database connection card and SSH tunnel card" width="100%">
</div>

The **Database connection** card supports MySQL / PostgreSQL / MongoDB. The **Port** box is an editable dropdown:
leave it empty and the adapter's default port is used (MySQL 3306 / PostgreSQL 5432 / MongoDB 27017) — the red arrow
points at that hint. The **SSH tunnel** card sits on top: the host/port you type there belong to the database
**on the remote server** (usually `127.0.0.1:5432`), and local connections are transparently rerouted through the
tunnel. While disconnected it reads `SSH 隧道:未连接`; once connected it shows
`127.0.0.1:<local port> → SSH user@host:22 → 127.0.0.1:5432`.


## Figure 3 · Connection result and tunnel log

<div align="center">
<img src="images/img01.png" alt="Connection success dialog and SSH tunnel log lines" width="100%">
</div>

**Test connection** reports the server version (PostgreSQL 18.6) in a dialog, while the log prints
`SSH 隧道已建立: 127.0.0.1:55960 → SSH …@172.24.208.28:22 → 127.0.0.1:5432` and `连接成功: PostgreSQL …`.
User names and passwords are masked in the picture; fill in your own values.


## Figure 4 · Progress while exporting a large table

<div align="center">
<img src="images/img04.png" alt="PostgreSQL over SSH to Excel export in progress" width="100%">
</div>

PostgreSQL (over an SSH tunnel) → Excel: the log prints `已写出 100,000 行 ...` every 20,000 rows, the status bar
shows "运行中..." ("running") and the progress bar switches to the indeterminate busy state. Every input field and
action button is disabled during the run, so a background task can never read a half-edited configuration and
double clicks are impossible.


## Figure 5 · One million rows and the fidelity check

<div align="center">
<img src="images/img05.png" alt="1,000,001 rows exported with fidelity check passed" width="100%">
</div>

`Excel 导出完成: D:\Downloads\virtual_data.xlsx（Sheet [VirtualProfile]，共 1,000,001 行）`; the freshly written
output is read back and compared value-wise **and** text-wise:
`数值保真校验：通过（字段 3 个 / 比对 575 个值）`. Million-row exports stay stable — the server-side streaming
cursor never buffers the whole table, and Excel uses openpyxl's streaming writer.


## Figure 6 · MongoDB → JSON

<div align="center">
<img src="images/img06.png" alt="Exporting a MongoDB collection to JSON" width="100%">
</div>

When MongoDB is selected the connection card grows a **Connection URI** row and an **authSource** row (hidden for the
other engines). Here a collection is exported to `MyData.json` over an SSH tunnel (`127.0.0.1:27017`, authSource
`admin`): the JSON top-level key `[VirtualProfile]` holds 4 documents — collection ↔ top-level key and document
field ↔ header row map one to one.


## Figure 7 · Nested documents and arrays round-trip

<div align="center">
<img src="images/img07.png" alt="mongosh documents next to the exported JSON" width="100%">
</div>

Left: the original documents in mongosh (`_id` as `ObjectId`, `Hobbies` as an array, `Birthday` / `CreateTime` as
`ISODate`). Right: the same collection exported to JSON — `_id` becomes a 24-character hex string, nested arrays are
kept as arrays, dates are written as `YYYY-MM-DD HH:MM:SS`. Importing back turns `_id` into an `ObjectId` again and
restores JSON text cells into nested structures (keys containing `.` or starting with `$` stay text, because MongoDB
would reject them).


## Figure 8 · The Linux GUI

<div align="center">
<img src="images/img10.png" alt="The same GUI running on Linux" width="100%">
</div>

Exactly the same window on Ubuntu 24.04, running a **MySQL → CSV directory-mode** export: the directory name equals
the database name and each table becomes one `.csv` file (here `/home/.../Cty/VirtualData.csv`, 20,000 rows). The
status bar shows "已连接数据库 | 就绪" ("database connected | ready").


## Figure 9 · Automatic dependency install on first start

<div align="center">
<img src="images/img08.png" alt="data_transformer.sh creating a virtualenv on first run" width="70%">
</div>

When `./data_transformer.sh` notices that the system interpreter lacks the required `openpyxl`, it asks whether to
create a project-local `.venv` and install the dependencies; pressing Enter creates the virtualenv and installs
`openpyxl / pymysql / psycopg2-binary / pymongo / questionary / paramiko`. Later runs reuse that virtualenv.


## Figure 10 · `--gui`: start the GUI once dependencies are ready

<div align="center">
<img src="images/img09.png" alt="data_transformer.sh --gui installing dependencies" width="70%">
</div>

`./data_transformer.sh --gui` follows the same dependency self-check and then launches the GUI. Note that the
**terminal wizard itself does not need PySide6** — on a headless server just run `./data_transformer.sh`.


## Figure 11 · Terminal wizard: pick a format → connect to MongoDB → hierarchical selection

<div align="center">
<img src="images/img11.png" alt="Terminal wizard selecting the input format and MongoDB parameters" width="100%">
</div>

Left: the `./data_transformer.sh` wizard — ① choose the input format (MySQL / PostgreSQL / MongoDB / Xlsx / Json / Csv)
→ `[SSH Tunnel]` switch → one prompt per parameter (Host / Port / Connection URI / Username / Password / Database /
authSource / Timeout) → `已连接 MongoDB 9.0.2` → ② **hierarchical selection**: level 1 picks collections, level 2
picks the fields of each collection individually. Right: the real documents of `MongoData.MongoCollections` in
mongosh, for comparison with the exported result.


## Figure 12 · Terminal wizard: JSON export finished

<div align="center">
<img src="images/img12.png" alt="Terminal wizard finishing a MongoDB to JSON export" width="100%">
</div>

③ Choose the output format `Json` → output file path → whether to rename tables (default: same as the source) →
**conversion summary** → confirm → `JSON 导出完成: …/MongoData.json（键 [MongoData]，共 20,000 行）` →
`数值保真校验：通过（字段 1 个 / 比对 200 个值）` → `已记住本次连接参数` ("connection parameters remembered").
Right: the tail of the exported JSON — `_id`, dates and the nested `Hobbies` array follow exactly the rules shown in
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

> Database-to-database conversion happens inside one server (same engine → direct table / collection copy).
> To cross database engines, go through a file (xlsx / json / csv) first.


# Project structure

```text
DataTransformer/
├── data_transformer.py    Windows GUI entry point (PySide6 + Nuitka target)
├── src/                   Engine package (the modules import each other relatively)
│   ├── __init__.py        Package marker
│   ├── dt_cli.py          Linux / macOS interactive terminal entry (questionary)
│   ├── dt_core.py         Core conversion engine (reader/writer pipeline, specs, SSH, self-test) + facade
│   ├── dt_config.py       Constants / connection config (DBConfig, SSHConfig) / remembered settings
│   ├── dt_files.py        File format layer (xlsx / csv / json read-write and value conversion)
│   ├── dt_sql.py          Database layer (MySQL / PostgreSQL: streaming read, batch write, type inference, free queries)
│   ├── dt_mongo.py        MongoDB layer (collection listing, schema sampling, Decimal128 read-write, find parsing)
│   └── dt_numeric.py      High-precision number layer (Decimal parsing / fixed-point formatting / reports)
├── tests/                 pytest unit and integration tests (incl. long-float suite)
│   ├── test_numeric.py    Numeric fidelity: trailing zeros, huge ints, scientific notation, JSON/Excel output
│   ├── test_engine.py     Engine: mapping rules, multi-table output, skip/directory modes, blocked conversion
│   ├── test_mongo.py      MongoDB value conversion and real-server round-trips (skipped without a server)
│   ├── test_cli.py        End-to-end terminal wizard interaction (driven through a pipe)
│   ├── test_gui.py        Key GUI logic (offscreen)
│   └── test_query.py      Custom queries: read-only classification, free-statement loading, find parsing, preview
├── data_transformer.sh    Linux / macOS launcher (detects Python, creates .venv on demand, runs python -m src.dt_cli)
├── build_exe.bat          Windows packaging script (Nuitka onefile)
├── run.bat                Run from source on Windows
├── conftest.py            pytest root configuration (puts the project root on sys.path)
├── requirements.txt       Dependency list
├── README.zh.md           Chinese documentation (this file's sibling)
├── README.en.md           English documentation (this file)
└── Demo/                  Sample data and notes (Example.txt / Example.xlsx)
```


# Dependencies

| Purpose | Dependency | Notes |
| :--- | :--- | :--- |
| Runtime | Python **3.10+** | uses `decimal`, dataclasses, the walrus operator, … |
| GUI | PySide6 ≥ 6.6 | only needed for the Windows GUI |
| Terminal wizard | questionary ≥ 2.0 | only needed for the Linux / macOS wizard |
| MySQL | PyMySQL ≥ 1.1 | server-side streaming cursor (SSCursor) |
| PostgreSQL | psycopg2-binary ≥ 2.9 | server-side named cursor |
| MongoDB | pymongo ≥ 4.6 | exact numeric read-write with Decimal128 |
| Spreadsheets | openpyxl ≥ 3.1 | read_only / write_only streaming modes |
| SSH tunnel | paramiko ≥ 3.4 (+ cryptography) | local port forwarding |
| Packaging | Nuitka | only needed to build the exe |
| Tests | pytest ≥ 8.0 | optional |


# Download

Three ways to use it: the Windows portable build, running the sources with a Windows batch file, and the
Linux / macOS terminal wizard.


## Windows

**Requirements:**

- Windows 10 or newer
- 64-bit (x86_64) only

**Builds:**

- **Portable**: `DataTransformer.exe` is a Nuitka onefile build — no Python needed, just double-click it.
- **Batch run**: `run.bat` runs the sources with the local Python interpreter (handy while developing).

The current version is **Ver1.1.7** (`VERSION` in `src/dt_config.py`; `DataTransformer.exe --version` prints it);
historical versions are archived in `Ver1.0.0` … `Ver1.1.6` (the names `Ver1.2.0` … `Ver1.6.0` and 1.7/1.8 are
retired).


## Linux / macOS

**Requirements:**

- Any mainstream distribution: Debian / Ubuntu / CentOS / / Fedora, and macOS
- Python 3.8+ (the launcher probes `python3` / `python3.x` / `python`)

**Usage:**

```bash
chmod +x data_transformer.sh   # first time only: make it executable
./data_transformer.sh          # start the interactive terminal wizard
```

The wizard then walks you through:

```text
① pick the input format: MySQL / PostgreSQL / Xlsx / Json / Csv
② pick the table (level 1) → pick the fields (level 2)
③ pick the output format and write mode
④ pick the export path (file or directory)
⑤ run the conversion and print the log
```

**The launcher automatically:**

- detects an available Python 3.8+ interpreter;
- checks the dependencies and, when something is missing, offers to create a project-local `.venv` and install
  `requirements.txt` (recent Debian / Ubuntu pip is protected by PEP 668, so a virtualenv is the safest route —
  see figure 9);
- reports the presence of `pymysql` / `psycopg2` / `paramiko` (they only affect the matching database or SSH
  feature; Excel / JSON / CSV conversion always works);
- reuses an existing `.venv` on later runs.

**Common arguments:**

```bash
./data_transformer.sh --check      # check the environment only, do not enter the wizard
./data_transformer.sh --demo       # prefill the SSH + PostgreSQL sample from Demo/Example.txt
./data_transformer.sh --selftest   # headless self-test, written to selftest.log
./data_transformer.sh --gui        # start the GUI (needs PySide6)
./data_transformer.sh --help       # all arguments
```

> The Linux wizard does **not** need PySide6, so it also works on a server without a desktop.


## Installing dependencies per distribution

The script probes and can create a project-local `.venv` by itself; for manual installs use the table below:

| Distribution | Install Python and venv | Install dependencies |
| :--- | :--- | :--- |
| Ubuntu / Debian | `sudo apt update && sudo apt install -y python3 python3-venv python3-pip` | `python3 -m venv .venv && .venv/bin/pip install -r requirements.txt` |
| CentOS  | `sudo dnf install -y python3 python3-pip` | as above (on CentOS 7 run `sudo yum install -y centos-release-scl` first) |
| Fedora | `sudo dnf install -y python3 python3-pip` | as above |
| macOS | `brew install python` | as above |

> On a few distributions psycopg2 needs build dependencies: `sudo apt install -y libpq-dev` on Debian/Ubuntu,
> `sudo dnf install -y postgresql-devel` on RHEL-family systems. Using `psycopg2-binary` (the default dependency)
> avoids that.
>
> Quick install of the core dependencies without a virtualenv (no PySide6):
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

> PySide6 is large; if the official index is slow, use a mirror:
>
> ```bash
> pip install PySide6 -i <mirror-url>
> ```

**Start the main program:**

```bash
python data_transformer.py
```

or double-click `run.bat`.

On Linux / macOS you can skip the GUI and run the wizard directly (no PySide6 required):

```bash
pip install openpyxl pymysql psycopg2-binary pymongo questionary paramiko   # core dependencies only
python -m src.dt_cli                                       # same as ./data_transformer.sh
python -m src.dt_cli --demo                                # with the Demo sample parameters
python -m src.dt_cli --plain                               # force plain-text prompts
python -m src.dt_cli --selftest                            # headless self-test
```


# Building the exe

Double-click `build_exe.bat`: it installs the dependencies and calls Nuitka (the first build takes about 5–8 minutes).

```powershell
build_exe.bat
```

Artifact:

```text
build\DataTransformer.exe
```

> If the onefile build fails, drop `--onefile` from the script and ship the folder build instead
> (`build\data_transformer.dist\DataTransformer.exe`).

> **Always verify the artifact before shipping**: `build\DataTransformer.exe` must be a **single file of about
> 44–48 MB** (it contains PySide6 + pymongo/bson + paramiko/cryptography + openpyxl). The **v1.1.6 published
> artifact** measured 46,789,632 bytes (44.6 MB),
> SHA-256 `9ed118bead348a2230b722d3d4596a8734534242f57e1139c55ff27ad9eff8d5` — a historical record of a shipped
> build, **not** a value to check v1.1.7 against: v1.1.7 has to be repackaged, and only then are its own size and
> SHA-256 recorded.
> If the target machine reports a `DataTransformer.exe` of about **219 KB**, the file transfer was truncated or only
> the launcher was copied:
>
> - **onefile mode**: copy the **complete** `build\DataTransformer.exe` (zipping it first is recommended) and compare
>   size and SHA-256 on the target machine;
> - **folder mode** (built without `--onefile`): copy the **entire `build\data_transformer.dist` folder**; the `.exe`
>   inside it cannot run on its own;
> - `build_exe.bat` now validates the size and prints the SHA-256 after packaging.
>
> Verify on the target machine:
>
> ```powershell
> Get-FileHash .\DataTransformer.exe -Algorithm SHA256
> .\DataTransformer.exe --selftest
> ```
>
> The self-test writes `selftest.log` next to the executable: `import paramiko` covers the SSH tunnel,
> `import pymongo` covers MongoDB and the "long float fidelity" line covers the numeric pipeline — everything must be
> `[OK]` for the build to be complete.
>
> **Full packaging command** (exactly what `build_exe.bat` runs, useful for troubleshooting):
>
> ```powershell
> python -m nuitka --standalone --onefile --enable-plugin=pyside6 ^
>   --windows-console-mode=attach --nofollow-import-to=pandas ^
>   --include-package=pymysql --include-package=psycopg2 --include-package-data=psycopg2 ^
>   --include-package=pymongo --include-package=bson ^
>   --include-package=paramiko --include-package=cryptography ^
>   --include-package=et_xmlfile ^
>   --include-package=src ^
>   --output-dir=build --output-filename=DataTransformer data_transformer.py
> ```
>
> `--windows-console-mode=attach`: double-clicking still opens no console window, but running
> `DataTransformer.exe --version` or `--selftest` from a terminal shows the output (`disable` makes both
> completely invisible).
>
> Notes: ① Python 3.10+ and MSVC (Visual Studio Build Tools, "Desktop development with C++") or MinGW64 are required;
> ② **close any running `DataTransformer.exe` before packaging**, otherwise Nuitka fails with `WinError 5 Access denied`
> when replacing the artifact; ③ the first build takes 5–15 minutes, later builds are much faster thanks to the C
> compilation cache; ④ changing `VERSION` in `src/dt_config.py` only takes effect after rebuilding.
>
> **Release package**: `release\DataTransformer_v1.1.6_win64.zip` is the **published 1.1.6 artifact (historical
> record)** — it contains `DataTransformer.exe`, an end-user `README.txt` and `SHA256SUMS.txt`; that build measured
> 46,789,632 bytes with SHA-256 `9ed118be…` and is kept for traceability only, it is **not** the current version.
> **1.1.7 needs a fresh package**: run `build_exe.bat` again, then record the real size and SHA-256 of the new
> artifact (printed by the script, or ask with `Get-FileHash .\DataTransformer.exe -Algorithm SHA256`) and update
> `release\SHA256SUMS.txt` together with the maintainer-facing Markdown release notes `release\README.md`.
> Unzip the package, verify the hash, run `--selftest` and the artifact is confirmed intact.

## Packaged exe does not start / crashes

### Symptoms and triage

When double-clicking `DataTransformer.exe` does nothing, flashes and closes, or shows an error, check two things first:

1. **the exe size**: a healthy onefile build is about 44–54 MB. A few hundred KB means the transfer was truncated or
   only the launcher was copied;
2. **Event Viewer → Windows Logs → Application** for an `Application Error` entry, and note its **faulting module** and
   **exception code**.

### Three typical outcomes

① **`Exception code 0xc0000005` with faulting module `data_transformer.dll`** (inside the onefile unpack directory, e.g.
`%TEMP%\onefile_*\`) = the compiled artifact itself is broken, not the Python code (the crash happens before any Python
code runs).

Fix: clean rebuild — delete `build\` and the `clcache`, `module-cache` and `scons-msvc-config` entries under
`%LOCALAPPDATA%\Nuitka\Nuitka\Cache`, and **prefer the Nuitka stable release** (this machine used 4.3, which Nuitka
itself reports as newer than the stable 4.2.2: `python -m pip install "nuitka==4.2.2"`).

② **Faulting module `Qt6Core.dll` / `qwindows.dll`, or a dialog saying the platform plugin is missing** (`could not find
or load the Qt platform plugin "windows"`) = the Qt plugins did not ship with the artifact.

Fix: make sure the build command carries `--enable-plugin=pyside6`, and that you distribute the **complete artifact**
(the whole exe in onefile mode; the whole `build\data_transformer.dist` folder in folder mode).

③ **No event log entry at all, the process vanishes instantly** = most likely quarantined by antivirus, or the onefile
unpack step is blocked.

Fix: whitelist the exe (and the `%TEMP%` unpack directory), or switch to folder mode (`build_exe.bat folder`), which
skips the unpack step.

### Bisecting

1. **Folder mode**: `build_exe.bat folder` → `build\data_transformer.dist\DataTransformer.exe` (no argument keeps onefile).
   If folder mode starts and onefile does not, the problem is in the onefile unpacker/launcher layer;
2. **Console subsystem**: change `CONSOLE_MODE` at the top of `build_exe.bat` from `attach` to `disable` and rebuild;
3. **Narrow the dependencies**: temporarily drop `--include-package=src` from the build command to see whether the crash
   goes away (tells you whether a module inside the `src` package is involved).

The build log lives in `build\nuitka_build.log`; the script echoes its last 30 lines before it ends.

### Run it from a terminal to see the output

```powershell
cd /d <build directory>
DataTransformer.exe --selftest
```

Keep `selftest.log` and the full terminal output when it fails. A Python-level startup exception is appended to
`startup-error.log` next to the executable (in the current working directory when running from source). Note that a
**native crash (such as `0xc0000005`) never reaches Python**, so `startup-error.log` stays empty for it — only the Event
Viewer's faulting module and fault offset can locate that.


# Features

MySQL / PostgreSQL / MongoDB and Excel / JSON / CSV convert into each other in any direction:

| in \ out | Excel | JSON | CSV | SQL database | MongoDB |
| :---: | :---: | :---: | :---: | :---: | :---: |
| SQL | ✅ | ✅ | ✅ directory mode | ✅ table copy | ✅ import into collection |
| MongoDB | ✅ | ✅ | ✅ directory mode | — | ✅ collection copy |
| Excel | ✅ | ✅ | ✅ | ✅ hierarchical import | ✅ hierarchical import |
| JSON | ✅ | ✅ | ✅ | ✅ create table on import | ✅ import into collection |
| CSV | ✅ | ✅ | ✅ | ✅ create table on import | ✅ import into collection |

Highlights:

- **Hierarchical selection**: a two-level tree — level 1 multi-selects tables / collections (or Excel sheets),
  level 2 ticks the fields of each table independently;
- **Custom queries**: an alternative to the hierarchical picker — write a SQL statement for MySQL /
  PostgreSQL, or a `db.<collection>.find(<filter>)` chain for MongoDB; the statement is validated and
  previewed before it becomes a data source (see the dedicated section below);
- **Long-float fidelity**: everything travels through `Decimal`, trailing zeros and significant digits are preserved,
  **scientific notation is never emitted**, and after writing the output is read back for a "numeric equality + text
  equality" check with a difference report (see the dedicated section below);
- **Field mapping**: automatic by name (sheet name ↔ table name, first header row ↔ database column names), with manual
  overrides and duplicate-name validation;
- **SSH tunnel**: reach MySQL / PostgreSQL / MongoDB on a specific IP without exposing the database port;
- **Terminal wizard**: `./data_transformer.sh` uses **questionary** for dynamic selection of input/output format,
  connection parameters, tables / collections and fields;
- **Excel → SQL hierarchical import**: the target table name mirrors the sheet name; a missing table is created
  automatically, an existing one is handled per sheet with **overwrite / append / skip**;
- **CSV directory mode**: the directory is named after the database and holds one `table.csv` per table;
- **Duplicate-file handling**: overwrite the whole file / merge / cancel;
- **Low memory on large data sets**: server-side streaming cursors plus batched inserts; million-row imports and
  exports are stable in practice;
- **Automatic table creation**: types are inferred from the first 200 sampled rows (`int → BIGINT/INT`,
  `Decimal → NUMERIC/DECIMAL`, `float → DOUBLE`, date/time → `TIMESTAMP/DATE`, strings → `TEXT`);
- **MongoDB nested structures**: nested documents / arrays are exported as JSON text cells and restored to nested
  structures on import (see figures 6 and 7);
- **Remembered connection settings**: passwords, database names and SSH parameters are filled in automatically next
  time;
- **Friendly connection errors**: wrong host / port / user / password / database is diagnosed with a localised hint
  (the raw error is always kept);
- **Button state coupling**: controls are greyed out until input/output are ready, the database parameters are
  complete, a connection exists and no task is running — no double submissions;
- **Bottom button column**: to the right of the run log the three buttons read **Reset → Clear log → Start conversion ▶** from
  top to bottom; all three use `QSizePolicy.Preferred, Expanding` and fill the column vertically (there is no trailing
  stretch), so maximising the window stretches them instead of leaving blank space underneath;
- **One-click self-test**: built-in `--selftest` validates the drivers, the engine, file conversions and numeric
  fidelity;
- **Injection safety**: table and column names are whitelisted and quoted (MySQL backticks / PostgreSQL double quotes).


## Naming conventions

Database – table – field map one to one:

| Level | Correspondence | Example |
| :---: | :--- | :--- |
| Database | Excel/JSON file name == CSV directory name == MongoDB database | `xxx.xlsx` / `xxx.json` / `xxx/` |
| Table | database table == MongoDB collection == Excel sheet == JSON top-level key == CSV file name | `"info"` → `xxx/info.csv` |
| Field | database column == MongoDB document field == first Excel cell == JSON HeaderFields == CSV header | `"Id"`, `"Name"`, `"CreateTime"` |

> Mapping is **automatic by identical name** and can be overridden in the GUI's "field mapping preview / adjust"
> dialog or in the terminal wizard (for example map `姓名` to `Name`). Duplicate target names are rejected before
> writing.

JSON uses a **header + row records** structure:

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

> `HeaderFields` lists the field names and `RowN` holds row N (rows are numbered from 1).
> The older **column-array** layout (`{"field": [values...]}`) is still readable, so old files keep working.


# Custom queries

Besides the "table → field" hierarchical picker, a data source can be a statement you write yourself: SQL for
MySQL / PostgreSQL, and `db.<collection>.find(<filter>)` chains for MongoDB. The result goes through the same
reader pipeline as a picked table, so field mapping, output naming and the numeric read-back check after writing
all keep working.

> Implementation: `SourceSpec.query` / `preview_query` in `src/dt_core.py`, `load_sql_query` in
> `src/dt_sql.py` and `parse_mongo_query` in `src/dt_mongo.py`; the GUI entry is `QueryDialog` in
> `data_transformer.py`, the wizard entry is `pick_query_source` in `src/dt_cli.py`.


## How to get there

| Front end | Entry point | Interaction |
| :--- | :--- | :--- |
| GUI | Click **自定义查询…** ("custom query") to the right of the **表/字段:** ("tables / fields") row in **① Input** | Type the statement in the dialog → **校验并预览** ("validate and preview") → the header and 3 sample rows appear → **确定** ("OK") |
| Terminal wizard | Choose **② 自由填写查询语句** ("write a query yourself") under 选择数据来源方式 ("data source method") | Type the statement → a read-only statement shows the header and 3 sample rows and asks whether to use it → `:q` cancels (five failed previews in a row abort the step) |

> In the GUI the dialog and the hierarchical picker are mutually exclusive: once enabled the button reads
> "自定义查询:<label>(点击修改)", and opening it loads the current statement.
> The query source is cleared only when you **pick tables / fields in the tree and confirm**, switch the input
> format, drop the database / SSH connection, or press reset; **refreshing the table list (refresh / connect and
> load) keeps the current data source**, so a statement you already confirmed is never silently replaced by the
> previous hierarchical selection.


## SQL examples

The SQL example built into the GUI's "custom query" dialog can be used as-is (rename table / columns):

```sql
SELECT Id, Name FROM VirtualProfile WHERE Name LIKE '%ExampleText%' ORDER BY Id LIMIT 100
```

Read-only statements starting with `WITH` work too:

```sql
WITH big AS (SELECT Id, Amount FROM VirtualProfile WHERE Amount > 1000)
SELECT Id, Amount FROM big ORDER BY Amount DESC
```

Any statement is allowed (including `UPDATE` / `DELETE` and DDL); it runs verbatim without identifier quoting,
**at your own risk**. Read-only classification looks only at the first keyword (leading `--` line comments and
`/* */` block comments are skipped first), against the whitelist
`select` / `with` / `show` / `describe` / `desc` / `explain` / `values` / `table`. A statement starting with `with`
is scanned again, and any of `insert` / `update` / `delete` / `merge` / `drop` / `alter` / `create` / `truncate` /
`grant` / `revoke` / `call` / `copy` makes it non-read-only.


## MongoDB examples

The MongoDB example built into the dialog is mongosh syntax itself (unquoted keys, an `$expr` compound condition) and
can be copied verbatim:

```javascript
db.TestCollections.find({ Age: { $gte: 25, $lte: 45 }, Example: "Text", $expr: { $and: [{ $isArray: "$Hobbies" }, { $gt: [ { $size: "$Hobbies" }, 3 ] }]}})
```

Strict JSON keeps working too:

```javascript
db.VirtualProfile.find({"Amount": {"$gt": 1000}})
  .projection({"Id": 1, "Name": 1, "Amount": 1})
  .sort({"Amount": -1})
  .limit(50)
```

Only `db.<collection>.find(<filter>)` is supported, optionally followed by `.sort({...})` / `.limit(n)` /
`.projection({...})`: any order, each written once, newlines and a trailing semicolon allowed, and `.limit()` takes
a non-negative integer.

**The arguments accept mongosh syntax**: keys may be unquoted (`{ Age: { $gte: 25 } }`), strings may use single
quotes, and trailing commas inside objects / arrays are dropped. The parser first rewrites the argument into strict
JSON with a character-level state machine (bare keys get quoted, single-quoted strings are normalised) and parses it
only afterwards, so a statement copied straight out of mongosh can be pasted in. The result must still be a JSON
literal: **no** constructors such as `ObjectId()` / `ISODate()` / `NumberDecimal()` (use a plain literal, e.g. the
date `"2026-08-29T00:00:00"`), and **not supported**: `aggregate`, `findOne`, `count`, `distinct` and other
methods, or `db.<db>.<collection>.find(...)`; when the statement does not match, the error message carries a usable
example plus the "raw text before normalisation", which pinpoints what to fix.


## Read-only vs non-read-only

| Statement | Preview | Execution | Usable as an export source |
| :--- | :--- | :--- | :--- |
| Read-only (first keyword on the whitelist; every MongoDB query counts as read-only) | **Runs once** for the header and 3 sample rows, then closes the cursor immediately | Normal execution; read-only SQL uses the server-side streaming cursor | Yes |
| Non-read-only (`UPDATE` / `DELETE` / `INSERT` / DDL …) | **Not executed**; only the note "该语句不是只读查询,预览已跳过" ("not read-only, preview skipped") | Runs after a second confirmation, on a plain client cursor (a named cursor can only `DECLARE ... SELECT`) | Yes with a result set, no without one |

- **Execute means commit**: a successful non-read-only statement is committed at once, and writes **cannot be rolled
  back**;
- **No result set**: after `UPDATE` / `DELETE` / DDL the tool reports "语句已执行(影响 N 行),但没有返回结果集,无法作为导出
  数据源" ("executed, N rows affected, but no result set — it cannot be an export source"); an unknown row count shows
  as `-1`;
- **Second confirmation**: in the GUI it is the question "该语句不是只读查询,执行会修改数据库(语句返回结果集时仍会导出),确定使用吗?"
  (default answer "No"); in the terminal wizard it is a prompt whose default answer is "no" (five failed previews
  in a row abort the step).


## How this interacts with existing features

- **No column pruning**: the result is exactly what the statement returns (for MongoDB, what `projection` selects);
  there is no field-subset picker;
- **Field mapping still applies**: the "field mapping preview / adjust" dialog can map query columns onto target
  fields, with duplicate-name validation;
- **Output naming**: the query label is used by default — for SQL the table name after the first `from`
  (`schema.table` keeps both parts), falling back to "自定义查询" ("custom query"); for MongoDB the collection name.
  It can still be renamed in the output configuration;
- **Everything else is unchanged**: directory mode, duplicate-file handling, batched writes and the numeric
  read-back check (with blocking on failure) behave exactly as before.


## Common errors

| Message | Cause | Fix |
| :--- | :--- | :--- |
| `查询语句语法不正确,应以 db.<集合>.find(<filter>) 开头` ("syntax error, should start with db.<collection>.find(<filter>)") | The statement is not a `db.<collection>.find(...)` chain (for example `aggregate`) | Rewrite it along the example in the message |
| `查询语句里 filter 不是合法 JSON: ...` ("filter is not valid JSON") | A value is not a valid JSON literal — typically an `ObjectId()` / `ISODate()` / `NumberDecimal()` constructor, or an unbalanced quote. Unquoted keys, single-quoted strings and trailing commas are **accepted** | Replace the constructor with a plain value (write the hex id as `"507f1f77bcf86cd799439011"`), use the "raw text before normalisation" in the message to locate it, then validate again |
| `查询语句不支持 .aggregate()` / `查询语句在 ... 处有多余内容` ("method not supported" / "trailing content") | A method other than `find`, or a second statement after the semicolon | Use only `sort` / `limit` / `projection`, one statement at a time |
| The driver's SQL syntax / missing table / permission error | The statement itself, or the connection lacks privileges | The preview reports it verbatim (the GUI adds a localised hint); fix it and validate again |
| `该语句不是只读查询,预览已跳过;执行会修改数据库(语句返回结果集时仍会导出)` ("not read-only, preview skipped") | The first keyword is not whitelisted, or `with` contains a data-modifying keyword | Refused by default; answer "yes" to the second confirmation only when you really mean to write |
| `语句已执行(影响 N 行),但没有返回结果集,无法作为导出数据源` ("executed, N rows affected, no result set") | A write statement or DDL returns no result set | The statement is already committed; switch to `... RETURNING ...` or to a read-only query |
| `集合【X】里没有可用字段,无法作为导出数据源(请检查集合是否为空)` ("collection has no usable fields — is it empty?") | An empty collection, or an exclusion-only `projection` that drops every field | Make sure the collection is not empty and use an inclusion `projection` |
| `没有填写查询语句` ("no query statement") | Previewing with an empty statement | Type a statement first |


# Long-float fidelity and avoiding scientific notation

High-precision values (many decimals, very large integers, values written in scientific notation) must stay
**numerically equal and textually identical** in every direction. Implementation notes (see `src/dt_numeric.py`):

| Stage | Rule |
| :--- | :--- |
| Intermediate representation | Always `decimal.Decimal` or Python's arbitrary-precision `int`; binary `float` is **never** used as a carrier. When only a float is available (Excel cells, FLOAT columns) the decimal literal is taken through `Decimal(str(value))` |
| Parsing | CSV text is classified by `parse_number`: integers → `int`, decimals / scientific notation → `Decimal`; JSON is read with `parse_float=Decimal`; zero-padded codes such as `007` are treated as **text** so the leading zeros survive |
| Formatting | Every output goes through `format_number`, which renders **fixed-point** decimals: `1.500` keeps its trailing zeros, `1E+20` expands to `100000000000000000000`, and `e` / `E` never appears |
| JSON output | Uses the built-in `dumps_json` serializer (the standard library cannot emit Decimal numbers), so numbers in JSON are fixed-point too |
| Excel output | `Decimal` maps to `NUMERIC/DECIMAL` (never DOUBLE) when creating tables; cells holding more than 15 significant digits or trailing zeros are written as **text** (Excel stores doubles internally, so writing a number would truncate it or eat the trailing zeros) |
| Database output | `Decimal` is handed to the driver as a parameter (psycopg2 / PyMySQL / pymongo `Decimal128`), never through a float; MongoDB values beyond `Decimal128` (34 digits) or the `int64` range degrade to text so nothing is truncated |
| Read-back check | After writing, the output is read back and every field is compared **both numerically and textually**; any precision loss, truncation or scientific notation **blocks the run** and prints a difference report (`allow_mismatch` lets it through, and the wizard asks whether to keep the result) |

Example output (input `1.500` and `12345678901234567890.12345` after a CSV → JSON → Excel → CSV round-trip, unchanged):

```json
{
    "info": {
        "HeaderFields": ["Id", "Amount"],
        "Data": {
            "Row1": [1, 12345678901234567890.12345],
            "Row2": [2, 1.500]
        }
    }
}
```

> A passing report looks like `数值保真校验:通过 (字段 2 个 / 比对 4 个值)`; on failure the differences are listed, e.g.
> `Amount 行1: 输入 '1.500' → 输出 '1.5'(文本表示变化)`.


# SSH tunnel to remote databases

When the database only listens on `127.0.0.1` and you must log in to the server over SSH first, the tool can set up
local port forwarding: tick **Enable SSH tunnel** and enter the address of the database **on the remote server**
(usually `127.0.0.1:5432`) in the connection card — the local connection is rerouted through the tunnel port and the
import/export flow stays identical to a local database.


## GUI walkthrough

1. The **SSH tunnel** card is at the top of the window (figure 2) — tick **Enable SSH tunnel**;
2. fill in the SSH host/IP, port (default 22), user name and password (this tool uses user name + password
   authentication; the wizard and engine still accept a key file, the GUI does not offer one);
3. click **Connect SSH**; the status area shows `127.0.0.1:<local port> → SSH user@host:22 → 127.0.0.1:5432`
   (figure 3);
4. back in the **Database connection** card enter the remote database address / port / user / password / database and
   click **Connect and load tables**;
5. then tick "table → fields" as usual and export or import (see figures 4 and 5 for large-table progress);
6. to switch database type just change **Database type**: the **tunnel stays up** — the tool only re-points the
   forward target at the new `host:port` and logs `SSH 隧道已重新指向 127.0.0.1:5432` (the port is whatever you
   typed); then click **Connect and load tables** again.

> The tunnel only listens on local `127.0.0.1` and never exposes a port to the network; clicking **Disconnect SSH**
> or closing the program releases it.

> **Switching database type never drops the tunnel**: the old database connection is discarded (reconnect with
> **Connect and load tables**), but the SSH channel is reused as-is — `SSHTunnel.retarget` in the engine only resets
> the forward target, so the local listening port already in use keeps working. On top of that, every call that reads
> the connection configuration (connect, read, query preview) first syncs the tunnel target to the current
> `host:port`, falling back to the default port of the selected database type when the field is empty or out of range.

> **Tunnel liveness check**: the GUI probes the tunnel every 3 seconds; as soon as it notices the tunnel is gone it
> logs `SSH 隧道已断开:<tunnel info>,请重新连接` ("tunnel lost — reconnect") **exactly once**, then automatically
> clears the database connection and refreshes the status bar and button states. The probe only runs while a tunnel
> exists, so the warning never floods the log. Click **Connect SSH** to rebuild the channel and carry on.


## One-click Demo/Example.txt sample

**Load the Example.txt sample** fills the form from `Demo/Example.txt`:

| Item | Value |
| :--- | :--- |
| SSH | `cml@172.24.208.28:22`, password `123456` |
| Database | PostgreSQL `Ubuntu2604@127.0.0.1:5432`, password `123456` |
| Database / table | `virtual_data` / `VirtualProfile` |

That is the equivalent of `ssh cml@172.24.208.28` followed by
`psql -h 127.0.0.1 -p 5432 -U Ubuntu2604 -d virtual_data`. Replace the table with your own values in another
environment.


## Terminal wizard

`./data_transformer.sh` first asks whether to connect through an SSH tunnel, exactly like the GUI; `--demo` prefills
the values above (the full interaction is shown in figures 11 and 12):

```bash
./data_transformer.sh --demo
```

> SSH support needs `paramiko` (`pip install paramiko`). Without it the buttons are greyed out with a hint and file
> conversion is unaffected.


# Command-line arguments

DataTransformer accepts the following arguments for debugging and troubleshooting.

**From source:**

```bash
python data_transformer.py <argument>
```

**Compiled build:**

```powershell
DataTransformer.exe <argument>
```

---

## `--selftest`

Headless self-test: verifies that the drivers and the engine are complete inside a packaged build and writes the
result to `selftest.log` in the working directory.

```powershell
DataTransformer.exe --selftest
```

Set the environment variables below to also test MySQL / PostgreSQL connectivity:

```powershell
set SELFTEST_MYSQL_USER=root
set SELFTEST_MYSQL_PASSWORD=your-password
set SELFTEST_MYSQL_DATABASE=your-database
set SELFTEST_PG_USER=your-user
set SELFTEST_PG_PASSWORD=your-password
set SELFTEST_PG_DATABASE=your-database

DataTransformer.exe --selftest
```

> Without those variables only the imports and the Excel / JSON / CSV round-trips are checked.

---

## `--version`

Prints the version and exits, which is convenient for scripts:

```powershell
DataTransformer.exe --version
```

The equivalent on Linux / macOS:

```bash
./data_transformer.sh --version
```

---


# Logging and debugging

The run log is shown in the "④ 运行日志" area of the main window, colour-coded by level, with a one-click
**Clear log** button.

**An error yields one localised hint, nothing more**: the log and the dialog show the Chinese text produced by
`describe_db_error` / `describe_ssh_error` (usually one or two lines that name the failing part — host / port /
user / password / database name / privileges — and keep the raw error on the last line), and by default **no Python
traceback and no source line numbers** are printed, so end users are not handed a stack dump.
When an SSH tunnel is active and the error looks like a broken connection (`abort` / `timeout` / `timed out` /
`refused`), one extra hint is appended: "当前启用了 SSH 隧道,隧道可能已断开,请检查【连接 SSH】状态后重试"
("an SSH tunnel is enabled — it may have dropped, check the Connect SSH state and retry").

For the full traceback, set the environment variable `DT_DEBUG=1` and run again:

```powershell
set DT_DEBUG=1
DataTransformer.exe --selftest
```

```bash
DT_DEBUG=1 ./data_transformer.sh
```

> With `DT_DEBUG` set the GUI appends `traceback.format_exc()` to the log, and the terminal wizard re-raises the
> exception in its last-resort handler. `DT_DEBUG` is off by default, which is why normal runs only show that
> localised hint.

If an import or export fails, keep the log content and attach it when you report the issue or open a ticket.


# Configuration and remembered settings

After a successful connection the tool remembers the connection parameters (the Windows GUI remembers the
**password** and **database name**; the terminal wizard also remembers the **host / port / user / MongoDB URI /
authSource / SSH parameters**) and fills them in next time.

The configuration file lives at:

```text
Windows      : %APPDATA%\DataTransformer\config.json
Linux / macOS: $XDG_CONFIG_HOME/DataTransformer/config.json (default ~/.config/DataTransformer/config.json)
```

Delete that file and restart to get a fresh start (only the host defaults to `127.0.0.1`; port, user, password and
database name are empty).

> Passwords are stored **in plain text** in that file, so do not use sensitive passwords on a shared machine.
> On Linux / macOS the file is written with permissions `600` (owner only) using a "write a temporary file, then
> replace atomically" strategy so a half-written configuration cannot happen.


# Security notes

- **Identifier validation**: every table and column name is whitelisted (must start with a letter, underscore or CJK
  character and contain only letters, digits and underscores) and then quoted per dialect (MySQL backticks /
  PostgreSQL double quotes), which prevents SQL injection through name concatenation. As a consequence, Excel sheet
  names and header cells containing spaces, brackets or dashes must be renamed before importing into a database
  (exporting to JSON / CSV / Excel has no such restriction);
- **Custom queries**: the statement is written by the user and executed verbatim, so it bypasses the identifier
  whitelist and quoting above — it is the only entry point that does; a non-read-only statement changes the
  database and cannot be rolled back, so only run statements you trust;
- **SSH tunnel**: the local listener always binds `127.0.0.1`, so no port is exposed to the local network;
- **Configuration permissions**: tightened to `600` on Linux / macOS;
- **CSV formula injection**: exported CSV / Excel keeps the data as-is. Text starting with `=`, `+`, `-` or `@` may be
  interpreted as a formula when opened directly in Excel — that is Excel's own behaviour. When handling untrusted
  data, inspect the file in a text editor first or export to JSON instead.


# Tests

```bash
python -m pytest tests/ -q                  # all unit and integration tests (189 in total)
python -m pytest tests/test_numeric.py -q   # long-float fidelity suite
python -m src.dt_core                       # headless self-test (same as --selftest), writes selftest.log
```

| Test file | Coverage |
| :--- | :--- |
| `tests/test_numeric.py` | Decimal parsing / fixed-point formatting / trailing-zero and leading-zero rules / JSON fixed-point serialisation / Excel cell policy / fidelity reports and multiset comparison |
| `tests/test_engine.py` | Long floats across xlsx / json / csv round-trips, scientific-notation normalisation, blocked and allowed fidelity failures, field and name mapping, multi-table output (including per-sheet read-back), skip and directory modes, guard rails |
| `tests/test_mongo.py` | MongoDB value conversion (Decimal128 / int64 overflow / nested structures / unsafe key names), `_id` handling and **real server** round-trips (skipped when no mongod is running) |
| `tests/test_cli.py` | End-to-end terminal wizard interaction (piped): full flow, output paths, nothing written on abort, no stack traces |
| `tests/test_gui.py` | Key GUI logic (offscreen): window construction, switching between the three database engines, connection target rewriting for SSH tunnels, field mapping, hierarchical selection, progress-bar lifecycle, connection-button readiness |
| `tests/test_query.py` | Custom queries: SQL read-only classification (including the `WITH` data-modifying CTE hardening), streaming vs plain cursor routing and commit semantics for free statements, the no-result-set error, MongoDB `find` chain parsing and error hints, **mongosh compatibility (bare keys / single quotes / trailing commas)**, `preview_query` leaving non-read-only statements unexecuted, and an end-to-end query source written to CSV |

> Cases per file: `test_cli.py` 17 / `test_engine.py` 21 / `test_gui.py` 41 / `test_mongo.py` 18 /
> `test_numeric.py` 27 / `test_query.py` 65 — **189** in total.

> The integration tests need a real environment: PostgreSQL / MySQL / MongoDB cases are skipped automatically when
> the service or the environment variables are missing, so they never fail the suite.


# FAQ

| Symptom | Fix |
| :--- | :--- |
| `ModuleNotFoundError: No module named 'questionary'` | The Linux wizard needs questionary: `pip install questionary`, or simply run `./data_transformer.sh` (it installs it for you) |
| `ModuleNotFoundError: No module named 'pymongo'` | Install it if you need MongoDB: `pip install pymongo`; ignore it otherwise |
| "非法标识符" (illegal identifier) | Table / column names (including Excel sheet names and headers) must start with a letter, underscore or CJK character and may only contain letters, digits and underscores. Rename names containing spaces / brackets / dashes, or export to JSON / CSV / Excel instead |
| Long numbers change or turn into scientific notation in Excel | Long values are written as **text cells** on purpose (Excel cannot represent more than 15 significant digits exactly). Use JSON / CSV to preserve numeric semantics, or format the column as text in Excel before importing |
| `数值保真校验未通过` (fidelity check failed) | The output lost precision, trailing zeros or produced scientific notation. The log and the dialog list sample differences: check the target column type (it should be NUMERIC / DECIMAL, not DOUBLE) or switch it to text |
| MySQL/PostgreSQL import fails with `value too long` | The target string column is too short; auto-created tables already use TEXT, so widen the column of a hand-made table |
| Mojibake with Chinese text | Run `chcp 65001` in the Windows console first, and pick `UTF-8 with BOM` or `GBK` as the CSV encoding |
| `无法连接 MongoDB ... ServerSelectionTimeoutError` | Check that mongod is running, the port and `authSource`; for remote servers tick the SSH tunnel |
| Million-row imports are slow | Increase "rows per batch" (default 1000, try 5000–20000) and check the server side (`max_allowed_packet` for MySQL, disk I/O) |
| The log shows only a localised hint after a failure — how do I see the code path? | Tracebacks and source line numbers are off by default; set `DT_DEBUG=1` before launching (`set DT_DEBUG=1` on Windows, `DT_DEBUG=1 ./data_transformer.sh` on Linux / macOS) and the full `traceback` is appended to the log |
| The log says `SSH 隧道已断开:…,请重新连接` ("tunnel lost, reconnect") | The 3-second liveness probe found the channel gone and cleaned up the database connection and buttons automatically; click **Connect SSH** to rebuild the channel, then **Connect and load tables** again (switching the database type does not drop the tunnel) |


# Acceptance checklist

- [x] Windows 10/11 GUI (PySide6): input parameters, output parameters, run log + start button, progress bar;
- [x] Windows single-file executable built with Nuitka (`build_exe.bat`, validating size and SHA-256);
- [x] Linux / macOS interactive terminal tool (`data_transformer.sh` → `python -m src.dt_cli`) using **questionary** for dynamic
      format and parameter selection, no GUI required;
- [x] Dependency installation and usage instructions for Ubuntu, Debian, Fedora and CentOS;
- [x] MySQL, PostgreSQL and **MongoDB** support plus xlsx, csv and json files;
- [x] Bidirectional database ↔ file conversion (including SQL→SQL table copy, MongoDB collection copy and file-to-file
      conversion);
- [x] Mapping rules: xlsx `sheetname` ↔ database table / MongoDB collection; first header row ↔ database column names;
      CSV / JSON offer the equivalent mapping with manual overrides;
- [x] Long-float and high-precision fidelity: numerically equal and textually identical, trailing zeros kept, no
      truncated significant digits, no `e`/`E` scientific notation;
- [x] A visible numeric verification mechanism with a difference report (read back after writing, block or confirm);
- [x] Data type conversion, NULLs, date/time, encoding, batch writes, error handling, logging and progress reporting;
- [x] Unit and integration tests (including the long-float suite and real database round-trips);
- [x] Free-form SQL / MongoDB query statements, validated and previewed before use (the GUI "自定义查询…" dialog
      and the wizard's "② 自由填写查询语句" branch);
- [x] The SQL and MongoDB examples built into the custom-query dialog can be used as-is (the MongoDB one is mongosh
      syntax: bare keys, single quotes and trailing commas all parse);
- [x] Switching the database type never drops the SSH tunnel (only the forward target is reset), and a dropped tunnel
      is reported once in the log with the connection and button state cleaned up automatically.


# AI-assisted development

This project was developed with the help of DeepSeek Harness (vibe coding) for writing code, refactoring, debugging
and problem analysis.

The overall design, feature planning, code review and final maintenance are the author's responsibility.


# Credits

DataTransformer builds on these projects:

- [Python](https://www.python.org/)
- [PySide6](https://doc.qt.io/qtforpython-6/)
- [questionary](https://questionary.readthedocs.io/) (dynamic terminal prompts)
- [PyMySQL](https://github.com/PyMySQL/PyMySQL)
- [psycopg2](https://www.psycopg.org/)
- [pymongo](https://pymongo.readthedocs.io/) (MongoDB)
- [openpyxl](https://openpyxl.readthedocs.io/)
- [paramiko](https://www.paramiko.org/) (SSH tunnel)
- [Nuitka](https://nuitka.net/)



Copyright © 2026 merry8462.
