<div align="center">

<img src="images/img03.png" alt="DataTransformer 主界面" width="100%">

<p>MySQL / PostgreSQL / MongoDB ↔ xlsx / csv / json 双向数据转换工具（Windows 图形界面 + Linux 终端向导）。</p>
<p>长浮点数与高精度数值全程以 <code>Decimal</code> 保真，输出不出现科学计数法，并在写入后自动回读校验。</p>

</div>

<p align="center">
简体中文 | <a href="README.en.md">English</a>
</p>


<div align="center">
<img src="images/img03.png" alt="DataTransformer 主界面全貌" width="100%">
</div>

自上而下依次是：
标题与彩色分隔条 → **SSH 通道** → **① 输入配置** → **② 输出配置** → **④ 运行日志**→ 底部状态栏。
图中是一次 **MySQL → CSV 目录模式**导出刚完成的状态。


# 截图


## 数据库连接与 SSH 通道

<div align="center">
<img src="images/img02.png" alt="数据库连接卡片与 SSH 通道卡片" width="100%">
</div>

**数据库连接**卡片支持 MySQL / PostgreSQL / MongoDB；「端口」是可下拉、可手填的输入框，
**留空即自动使用该类型的默认端口**（MySQL 3306 / PostgreSQL 5432 / MongoDB 27017）。
最上方的 **SSH 通道**卡片：勾选后填写的「主机 / 端口」是**远程服务器上**数据库的地址（通常 `127.0.0.1:5432`），
本机连接自动改走通道端口；未连接时状态显示 `SSH 通道:未连接`，连接后显示
`127.0.0.1:<本地端口> → SSH 用户@主机:22 → 127.0.0.1:5432`。


## 连接成功与通道日志

<div align="center">
<img src="images/img01.png" alt="连接成功提示与 SSH 通道日志" width="100%">
</div>

点「测试连接」后弹窗给出服务端版本（PostgreSQL 18.6），日志区同时打印
`SSH 通道已建立: 127.0.0.1:55960 → SSH …@172.24.208.28:22 → 127.0.0.1:5432` 与
`连接成功: PostgreSQL …`。图中的用户名、密码等敏感字段已打码，实际使用时按自己的环境填写。


## 表格导与进度提示

<div align="center">
<img src="images/img04.png" alt="PostgreSQL 经 SSH 通道导出 Excel 的进度" width="100%">
</div>

PostgreSQL（经 SSH 通道）→ Excel：日志每 2 万行打印一次，
执行期间所有输入框与操作按钮自动置灰，既防止中途改参数让后台任务读到“半截”配置，也防止重复点击。


## 百万行表格导出

<div align="center">
<img src="images/img05.png" alt="100 万行导出完成与数值保真校验" width="100%">
</div>

`Excel 导出完成: D:\Downloads\virtual_data.xlsx（Sheet [VirtualProfile]，共 1,000,001 行）`；
写入后自动回读输出做「数值等值 + 文本等值」双重比对，`数值保真校验：通过（字段 3 个 / 比对 575 个值）`
—— 百万行级别依然稳定：服务端流式游标不会把整表读进内存，Excel 使用 openpyxl 流式模式。


## MongoDB → JSON

<div align="center">
<img src="images/img06.png" alt="MongoDB 集合导出为 JSON" width="100%">
</div>

选中 MongoDB 时，数据库卡片会多出 **连接 URI** 与 **认证库 authSource** 两行（其他类型自动隐藏）。
图中经 SSH 通道连接 `127.0.0.1:27017`、认证库，把集合导出为 `MyData.json`
（一级键 `[VirtualProfile]`，共 4 行）；集合名 ↔ JSON 一级键、文档字段 ↔ 首行表头一一对应。


## 嵌套文档 / 数组的往返对照

<div align="center">
<img src="images/img07.png" alt="mongosh 原始文档与导出 JSON 的对照" width="100%">
</div>

左侧是 mongosh 中的原始文档（`_id` 为 `ObjectId`、`Hobbies` 为数组、`Birthday` / `CreateTime` 为 `ISODate`），
右侧是同一集合导出的 JSON：`_id` 输出为 24 位十六进制文本，嵌套数组原样展开，
日期统一写成 `YYYY-MM-DD HH:MM:SS`。反向导入时 `_id` 会还原成 `ObjectId`，
JSON 文本单元格自动还原为嵌套结构（键名含 `.` 或以 `$` 开头时按文本保留，避免 MongoDB 键名限制导致写入失败）。


## Linux 图形界面

<div align="center">
<img src="images/img10.png" alt="Linux 上的图形界面" width="100%">
</div>

同一套界面在 Ubuntu 24.04 上运行 **MySQL → CSV 目录模式**：目录名 = 数据库名，目录下每个表一个 `.csv`
（图中导出到 `/home/.../Cty/VirtualData.csv`，共 20,000 行），底部状态栏显示「已连接数据库 | 就绪」。


## 首次启动(自动安装依赖项)

<div align="center">
<img src="images/img08.png" alt="data_transformer.sh 首次运行自动创建虚拟环境" width="70%">
</div>

`./data_transformer.sh` 探测到系统解释器缺少必需依赖 `openpyxl` 时，会询问
「是否在项目目录创建虚拟环境 .venv 并自动安装依赖」，回车即自动创建 `.venv` 并安装
`openpyxl / pymysql / psycopg2-binary / pymongo / questionary / paramiko`；再次运行会直接复用该虚拟环境。


## GUI界面 依赖就绪后启动图形界面

<div align="center">
<img src="images/img09.png" alt="data_transformer.sh --gui 自动安装依赖" width="70%">
</div>

`./data_transformer.sh --gui` 走的是同一条依赖自检流程，依赖就绪后启动图形界面。
注意：**命令行向导本身不需要 PySide6**，无桌面环境的服务器用默认的 `./data_transformer.sh` 即可。


## 终端向导：选择格式 → 连接 MySQL or PostgreSQL or MongoDB → 数据库/表/集合/字段 的层级勾选

<div align="center">
<img src="images/img11.png" alt="终端向导:选择输入格式与 MongoDB 连接参数" width="100%">
</div>

左侧是 `./data_transformer.sh` 的终端向导：① 选择输入格式（MySQL / PostgreSQL / MongoDB / Xlsx / Json / Csv）
→ `[SSH Tunnel]` 开关 → 逐个参数行（Host / Port / Connection URI / Username / Password / Database / authSource / Timeout）
→ `已连接 MongoDB 9.0.2` → ② **层级勾选**：第 1 层选集合，第 2 层给每个集合单独勾选字段。
右侧是同一时刻 mongosh 中 `MongoData.MongoCollections` 的真实文档，可与导出结果对照。


## 终端向导：导出 JSON 完成

<div align="center">
<img src="images/img12.png" alt="终端向导:MongoDB 导出 JSON 完成" width="100%">
</div>

③ 选择输出格式 `Json` → 导出文件路径 → 是否重命名数据表（默认与源同名）→ **转换摘要** → 确认
→ `JSON 导出完成: …/MongoData.json（键 [MongoData]，共 20,000 行）`
→ `数值保真校验：通过（字段 1 个 / 比对 200 个值）` → `已记住本次连接参数`。
右侧是导出的 JSON 尾部内容：`_id` / 日期 / 嵌套 `Hobbies` 数组都按上一张图的规则呈现。


# 支持矩阵

| 输入 \ 输出 | xlsx | json | csv | MySQL | PostgreSQL | MongoDB |
| :---: | :---: | :---: | :---: | :---: | :---: | :---: |
| **xlsx** | — | ✅ | ✅ | ✅ 层级导入 | ✅ 层级导入 | ✅ 层级导入 |
| **json** | ✅ | — | ✅ | ✅ 导入建表 | ✅ 导入建表 | ✅ 导入集合 |
| **csv** | ✅ | ✅ | — | ✅ 导入建表 | ✅ 导入建表 | ✅ 导入集合 |
| **MySQL** | ✅ | ✅ | ✅ 目录模式 | ✅ 表拷贝 | ✅ 表拷贝 | ✅ 导入集合 |
| **PostgreSQL** | ✅ | ✅ | ✅ 目录模式 | ✅ 表拷贝 | ✅ 表拷贝 | ✅ 导入集合 |
| **MongoDB** | ✅ | ✅ | ✅ 目录模式 | — | — | ✅ 集合拷贝 |

> 数据库之间的转换在同一服务器内进行（同类型数据库直接做表 / 集合拷贝）；跨数据库类型请先用文件（xlsx / json / csv）中转。


# 项目结构

```text
DataTransformer/
├── data_transformer.py    Windows GUI 入口（PySide6 + Nuitka 打包目标）
├── src/                   引擎包（模块之间使用相对导入）
│   ├── __init__.py        包标记
│   ├── dt_cli.py          Linux / macOS 终端交互式入口（questionary 动态选择）
│   ├── dt_core.py         核心转换引擎（Reader/Writer 管道、规格、SSH、自检）+ 统一门面
│   ├── dt_config.py       常量 / 连接配置（DBConfig、SSHConfig）/ 记忆配置（Settings）
│   ├── dt_files.py        文件格式适配层（xlsx / csv / json 读写与值转换）
│   ├── dt_sql.py          数据库适配层（MySQL / PostgreSQL：流式读、批量写、类型推断、自由查询）
│   ├── dt_mongo.py        MongoDB 适配层（集合枚举、Schema 采样、Decimal128 读写、find 语句解析）
│   └── dt_numeric.py      高精度数值保真层（Decimal 解析 / 定点格式化 / 保真校验报告）
├── tests/                 pytest 单元测试与集成测试（含长浮点数专项）
│   ├── test_numeric.py    数值保真：尾零、超大整数、科学计数法、JSON/Excel 定点输出
│   ├── test_engine.py     转换引擎：映射规则、多表输出、跳过/目录模式、校验阻断
│   ├── test_mongo.py      MongoDB 值转换与真实服务端往返（无服务时自动跳过）
│   ├── test_cli.py        命令行向导端到端交互（管道驱动）
│   ├── test_gui.py        图形界面关键逻辑（离屏运行）
│   └── test_query.py      自定义查询语句：只读判定、自由语句读取、Mongo find 解析、预览
├── data_transformer.sh    Linux / macOS 启动脚本（探测 Python、按需创建 .venv 后运行 python -m src.dt_cli）
├── build_exe.bat          Windows 打包脚本（Nuitka onefile）
├── run.bat                Windows 源码运行
├── conftest.py            pytest 根目录配置（把项目根加入 sys.path）
├── requirements.txt       依赖清单
├── README.zh.md           中文文档
└── README.en.md           英文文档
```


# 依赖清单

| 用途 | 依赖 | 说明 |
| :--- | :--- | :--- |
| 运行环境 | Python **3.10+** | 使用 `decimal`、dataclass、海象运算符等 |
| 图形界面 | PySide6 ≥ 6.6 | 仅 Windows GUI 需要 |
| 终端交互 | questionary ≥ 2.0 | 仅 Linux / macOS 命令行向导需要 |
| MySQL | PyMySQL ≥ 1.1 | 服务端流式游标（SSCursor） |
| PostgreSQL | psycopg2-binary ≥ 2.9 | 服务端命名游标（named cursor） |
| MongoDB | pymongo ≥ 4.6 | Decimal128 精确数值读写 |
| 表格文件 | openpyxl ≥ 3.1 | read_only / write_only 流式模式 |
| SSH 通道 | paramiko ≥ 3.4（+ cryptography） | 本地端口转发 |
| 打包 | Nuitka | 仅打包时需要 |
| 测试 | pytest ≥ 8.0 | 可选 |


# 下载

目前提供 Windows 便携版、Windows bat 源码运行,以及 Linux / macOS 命令行向导三种方式。


## Windows

**系统要求：**

- Windows 10 及以上版本
- 仅支持 64 位（x86_64）系统

**版本说明：**

- **当前版本**：**Ver1.1.7**（见 `src/dt_config.py` 的 `VERSION`，`DataTransformer.exe --version` 打印的就是它）；
  历史版本归档在 `Ver1.0.0` … `Ver1.1.6`。
- **便携版**：`DataTransformer.exe` 为 Nuitka 打包的单文件程序，无需安装 Python，下载后双击即可使用。
- **bat 运行**：`run.bat` 使用本机 Python 直接运行源码，适合开发调试。


## Linux / macOS

**系统要求：**

- 主流发行版均可：Debian / Ubuntu / CentOS / Rocky / Fedora，以及 macOS
- Python 3.8+（启动脚本会自动探测 `python3` / `python3.x` / `python`）

**使用方式：**

```bash
chmod +x data_transformer.sh   # 首次使用：赋予执行权限
./data_transformer.sh          # 启动交互式命令行向导
```

执行后会展开交互式命令窗口，按提示依次选择：

```text
① 选择输入格式：MySQL / PostgreSQL / Xlsx / Json / Csv
② 选择数据表
③ 选择输出格式与写入方式
④ 选择导出路径（文件或目录）
⑤ 执行转换并输出日志
```

**启动脚本会自动完成：**

- 探测可用的 Python 3.8+ 解释器；
- 检查依赖，缺少时询问是否在项目目录创建 `.venv` 并自动安装
  （Debian / Ubuntu 新版 pip 受 PEP 668 保护，虚拟环境是最稳妥的方式，实际界面见 图 9）；
- 提示 `pymysql` / `psycopg2` / `paramiko` 的安装情况（仅影响对应的数据库 / SSH 功能，
  Excel / JSON / CSV 互转始终可用）；
- 再次运行会直接复用已建好的 `.venv`。

**常用参数：**

```bash
./data_transformer.sh --check      # 只检查运行环境，不进入向导
./data_transformer.sh --demo       # 预填 Demo/Example.txt 的 SSH + PostgreSQL 示例参数
./data_transformer.sh --selftest   # 无界面自检，结果写入 selftest.log
./data_transformer.sh --gui        # 启动图形界面（需 PySide6）
./data_transformer.sh --help       # 全部参数
```

> Linux 上的命令行向导**不需要 PySide6**，在无桌面环境的服务器上也能直接使用。


## 各发行版依赖安装

脚本会自行探测并可在项目目录创建 `.venv`；如需手动安装，可参考下表：

| 发行版 | 安装 Python 与 venv | 安装依赖 |
| :--- | :--- | :--- |
| Ubuntu / Debian | `sudo apt update && sudo apt install -y python3 python3-venv python3-pip` | `python3 -m venv .venv && .venv/bin/pip install -r requirements.txt` |
| CentOS / Rocky | `sudo dnf install -y python3 python3-pip` | 同上（CentOS 7 建议先 `sudo yum install -y centos-release-scl`） |
| Fedora | `sudo dnf install -y python3 python3-pip` | 同上 |
| macOS | `brew install python` | 同上 |

> psycopg2 在少数发行版需要编译依赖：Debian/Ubuntu 用 `sudo apt install -y libpq-dev`，
>
> 免虚拟环境快速安装核心依赖（不含 PySide6）：
>
> ```bash
> python3 -m pip install --user openpyxl pymysql psycopg2-binary pymongo questionary paramiko
> ```


# 从源码运行

首先克隆项目：

```bash
git clone https://github.com/merry8462/DataTransformer.git
cd DataTransformer
```

安装依赖：

```bash
pip install -r requirements.txt
```

> ```bash
> pip install PySide6 -i 镜像源
> ```

**启动主程序：**

```bash
python data_transformer.py
```

或双击 `run.bat`。

Linux / macOS 上也可以跳过图形界面直接运行向导（无需 PySide6）：

```bash
pip install openpyxl pymysql psycopg2-binary pymongo questionary paramiko   # 只装核心依赖
python -m src.dt_cli                                       # 等价于 ./data_transformer.sh
python -m src.dt_cli --demo                                # 带 Demo 示例参数
python -m src.dt_cli --plain                               # 强制纯文本问答
python -m src.dt_cli --selftest                            # 无界面自检
```


# 打包为 exe

双击 `build_exe.bat`，脚本会自动安装依赖并调用 Nuitka 打包（首次约 5~8 分钟）：

```powershell
build_exe.bat
```

打包产物：

```text
build\DataTransformer.exe
```

**完整打包命令**（与 `build_exe.bat` 一致，`--include-package=src` 用于把 7 个引擎模块所在的 `src` 包一并收进产物）：

```powershell
python -m nuitka --standalone --onefile --enable-plugin=pyside6 ^
  --windows-console-mode=attach --nofollow-import-to=pandas ^
  --include-package=pymysql --include-package=psycopg2 --include-package-data=psycopg2 ^
  --include-package=pymongo --include-package=bson ^
  --include-package=paramiko --include-package=cryptography ^
  --include-package=et_xmlfile --include-package=src ^
  --output-dir=build --output-filename=DataTransformer data_transformer.py
```

> `--windows-console-mode=attach`：双击启动时依然没有控制台窗口，但从命令行运行 `DataTransformer.exe --version`
> 或 `--selftest` 时能看到输出（`disable` 会让这两条命令的输出完全不可见）。

> **发布包**：`release\DataTransformer_v1.1.6_win64.zip` 是 **1.1.6 已发布产物（历史记录）**——内含
> `DataTransformer.exe` + 面向最终用户的 `README.txt` + `SHA256SUMS.txt`，当时的产物为 46,789,632 字节、
> SHA-256 `9ed118be…`，仅供回溯核对，**不代表当前版本**。
> **1.1.7 需重新打包后记录实际大小与 SHA-256**：重新执行 `build_exe.bat` 生成新的 exe，打包结束后按脚本打印的
> 大小 / 哈希（或 `Get-FileHash .\DataTransformer.exe -Algorithm SHA256`）记录实际值，并同步更新
> `release\SHA256SUMS.txt` 与面向维护者的 Markdown 发布说明 `release\README.md`。
> 解压后核对哈希并执行 `--selftest` 即可确认产物完整。

## 打包后的 exe 无法启动 / 崩溃排查

### 症状与判据

双击 `DataTransformer.exe` 没反应、一闪而过或直接报错时，先看两件事：

1. **exe 大小**：onefile 正常约 44–54 MB。只有几百 KB 说明传输被截断或只复制了启动器；
2. **事件查看器 → Windows 日志 → 应用程序**里有没有 `Application Error` 记录，记下其中的**出错模块**与**异常代码**。

### 三种典型结果与处理

① **`异常代码 0xc0000005`，出错模块是 `data_transformer.dll`**（位于 onefile 解包目录内，如 `%TEMP%\onefile_*\`）=
编译产物本身有问题，与 Python 代码无关（该崩溃发生在 Python 代码开始执行之前）。

处理：干净重建——删掉 `build\` 与 `%LOCALAPPDATA%\Nuitka\Nuitka\Cache` 下的 `clcache`、`module-cache`、
`scons-msvc-config`，并**优先使用 Nuitka 稳定版**（本机曾用 4.3，Nuitka 自己提示它比稳定版 4.2.2 更新：
`python -m pip install "nuitka==4.2.2"`）。

② **出错模块是 `Qt6Core.dll` / `qwindows.dll`，或弹窗提示找不到平台插件**（`could not find or load the Qt platform
plugin "windows"`）= Qt 插件没有随包。

处理：确认打包命令带 `--enable-plugin=pyside6`，并且分发的是**完整产物**（onefile 复制整个 exe；folder 模式复制整个
`build\data_transformer.dist` 文件夹）。

③ **完全没有事件日志、进程瞬间消失** = 大概率被杀软隔离，或 onefile 解包被拦。

处理：给 exe 与 `%TEMP%` 解包目录加白名单，或改用 folder 模式（`build_exe.bat folder`）绕开解包环节。

### 二分手段

1. **目录模式对比**：`build_exe.bat folder` → 产物 `build\data_transformer.dist\DataTransformer.exe`（不带参数时仍是 onefile）。
   目录模式能启动而 onefile 不能，说明问题出在 onefile 解包/启动器这一层；
2. **控制台子系统对比**：把 `build_exe.bat` 顶部的 `CONSOLE_MODE` 由 `attach` 改成 `disable` 后重新打包；
3. **收窄依赖**：必要时临时去掉打包命令里的 `--include-package=src`，看崩溃是否消失（用于判断是否与 `src` 包内模块相关）。

构建日志在 `build\nuitka_build.log`，脚本会在构建结束前回显最后 30 行。

### 从命令行运行以便看到输出

```powershell
cd /d <build 目录>
DataTransformer.exe --selftest
```

失败时把 `selftest.log` 与命令行完整输出一并留存。启动期若抛出 Python 层异常，会追加写入 exe 同级目录的
`startup-error.log`（源码运行时写到当前工作目录）。注意**原生崩溃（如 `0xc0000005`）不经过 Python 层**，
`startup-error.log` 里不会有它的记录，只能靠事件查看器的出错模块与错误偏移定位。


# 功能说明

支持 MySQL / PostgreSQL / MongoDB 与 Excel / JSON / CSV 任意互转：

| 输入 \ 输出 | Excel | JSON | CSV | SQL 数据库 | MongoDB |
| :---: | :---: | :---: | :---: | :---: | :---: |
| SQL | ✅ | ✅ | ✅ 目录模式 | ✅ 表拷贝 | ✅ 导入集合 |
| MongoDB | ✅ | ✅ | ✅ 目录模式 | — | ✅ 集合拷贝 |
| Excel | ✅ | ✅ | ✅ | ✅ 层级导入 | ✅ 层级导入 |
| JSON | ✅ | ✅ | ✅ | ✅ 导入建表 | ✅ 导入集合 |
| CSV | ✅ | ✅ | ✅ | ✅ 导入建表 | ✅ 导入集合 |

主要特性：

- **层级勾选**：树形下拉框两层结构，第一层多选数据表 / 集合（或 Excel Sheet），第二层每张表独立勾选字段；
- **自定义查询**：与层级勾选并列的可选项——SQL 数据库直接写 SQL 语句，MongoDB 写 `db.<集合>.find(<filter>)` 链式语句，先校验预览表头与样本行再作为数据源（详见下文专项说明）；
- **长浮点数保真**：全程 `Decimal` 中转，不丢尾零、不截断有效数字、**输出不出现科学计数法**，写入后自动回读做「数值等值 + 文本等值」双重校验并输出差异报告（详见下文专项说明）；
- **字段映射**：自动按同名映射（Sheet 名 ↔ 表名、首行表头 ↔ 数据库字段名），支持手动覆盖并做重名校验；
- **SSH 通道**：连接指定 IP 服务器上的 MySQL / PostgreSQL / MongoDB，无需在服务器上开放数据库端口；
- **命令行向导**：`./data_transformer.sh` 通过 **questionary** 动态选择 Input/Output 格式、连接参数、表 / 集合与字段；
- **Excel → SQL 层级导入**：目标表名与 Sheet 名一一对应；无同名表时自动按 Sheet 名新建，存在同名表时逐个选择 **覆盖写入 / 追加写入 / 跳过该表**；
- **目录模式 CSV 导出**：目录名 = 数据库名，目录下每个表一个 `表名.csv`；
- **同名文件纠错**：检测到同名文件时选择 覆盖整个文件 / 合并写入 / 取消；
- **大批量低内存**：服务端流式游标 + 批量入库，实测百万行级数据稳定导入导出；
- **自动建表**：按前 200 行采样推断字段类型（`int → BIGINT/INT`、`Decimal → NUMERIC/DECIMAL`、`float → DOUBLE`、日期时间 → `TIMESTAMP/DATE`，字符串统一 `TEXT`）；
- **MongoDB 嵌套结构**：嵌套文档 / 数组导出为 JSON 文本单元格，导入时自动还原为嵌套结构；
- **连接配置记忆**：密码、数据库名与 SSH 参数自动记忆，下次启动自动填入；
- **连接异常友好提示**：主机 / 端口 / 用户名 / 密码 / 数据库名错误自动识别并弹出中文定位提示（同时保留原始错误）；
- **按钮状态联动**：输入 / 输出未就绪、数据库信息未填全、未连接或任务执行中时，相关按钮自动置灰，防止重复操作；
- **底部按钮区**：右下角三个按钮自上而下依次是 **初始化 → 清空日志 → 开始转换 ▶**，三者都设为
  `QSizePolicy.Preferred, Expanding` 并纵向填满该列（列尾不再有 stretch），窗口最大化后按钮随列拉伸、下方不留空白；
- **一键自检**：内置 `--selftest`，验证驱动、引擎、文件互转与数值保真；
- **安全防注入**：表名 / 字段名白名单校验并加引用符（MySQL 反引号 / PostgreSQL 双引号）。


## 命名约定

库-表-字段三层一一对应：

| 层级 | 对应关系 | 示例 |
| :---: | :--- | :--- |
| 数据库名 | Excel/JSON 文件名 == CSV 目录名 == MongoDB 数据库 | `xxx.xlsx` / `xxx.json` / `xxx/` |
| 表名 | 数据库表名 == MongoDB 集合名 == Excel Sheet 名 == JSON 一级键 == CSV 文件名 | `"info"` → `xxx/info.csv` |
| 字段名 | 数据库字段名 == MongoDB 文档字段 == Excel 首行单元格 == JSON HeaderFields == CSV 表头 | `"Id"`、`"Name"`、`"CreateTime"` |

> 映射默认**按同名自动匹配**，并可在图形界面「字段映射预览 / 调整」或命令行向导中选择手动覆盖
> （例如把 `姓名` 映射为 `Name`）；映射后的重名字段会在写入前被拦截并提示。

JSON 采用**表头 + 行记录**结构：

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

> `HeaderFields` 为字段名列表，`Data` 中的 `RowN` 为第 N 行数据数组（行号从 1 开始）。
> 旧版**列数组**结构（`{"字段名": [值...]}`）仍可正常读取，向后兼容。


# 自定义查询语句

数据源除了「表 → 字段」层级勾选，也可以直接写一条查询语句：SQL 数据库（MySQL / PostgreSQL）写 SQL，
MongoDB 写 `db.<集合>.find(<filter>)` 及其链式调用。查询结果与层级勾选共用同一条 Reader 管道，
字段映射、输出命名与写入后的数值保真回读校验都照常生效。

> 实现位置：`src/dt_core.py` 的 `SourceSpec.query` / `preview_query`、`src/dt_sql.py` 的 `load_sql_query`、
> `src/dt_mongo.py` 的 `parse_mongo_query`；图形界面入口是 `data_transformer.py` 的 `QueryDialog`，
> 终端向导入口是 `src/dt_cli.py` 的 `pick_query_source`。


## 入口

| 前端 | 进入方式 | 交互 |
| :--- | :--- | :--- |
| 图形界面 | **① 输入配置** 的「表/字段:」行右侧点击 **自定义查询…** | 弹窗内填写语句 → 点 **校验并预览** → 显示表头与前 3 行样本 → 点 **确定** |
| 终端向导 | **选择数据来源方式** 选 **② 自由填写查询语句** | 输入语句 → 只读语句显示表头与 3 行样本并询问是否采用 → 输入 `:q` 取消（连续 5 次预览失败会自动中止） |

> 图形界面里「自定义查询…」与层级勾选互斥：启用后按钮变为「自定义查询:<标签>(点击修改)」，点开会载入当前语句。
> 只有**在树里重新勾选表 / 字段并确认**、切换输入格式、断开数据库或 SSH 通道、点「初始化」时才停用该查询数据源；
> **刷新表结构（刷新表 / 连接并加载表）不会改变当前数据源**，避免已确认的语句被悄悄换回层级勾选。


## SQL 示例

图形界面「自定义查询语句」对话框里内置的 SQL 示例可以直接套用（改表名 / 字段名即可）：

```sql
SELECT Id, Name FROM VirtualProfile WHERE Name LIKE '%ExampleText%' ORDER BY Id LIMIT 100
```

以 `WITH` 开头的只读语句同样可用：

```sql
WITH big AS (SELECT Id, Amount FROM VirtualProfile WHERE Amount > 1000)
SELECT Id, Amount FROM big ORDER BY Amount DESC
```

SQL 允许任意语句（含 `UPDATE` / `DELETE` / DDL），语句原样执行、不做标识符引用，**风险由使用者自负**。
只读判定只看首个关键字（开头的 `--` 行注释与 `/* */` 块注释会先被跳过），白名单为
`select` / `with` / `show` / `describe` / `desc` / `explain` / `values` / `table`；
以 `with` 开头时还会再扫一遍整条语句，出现 `insert` / `update` / `delete` / `merge` / `drop` / `alter` /
`create` / `truncate` / `grant` / `revoke` / `call` / `copy` 等关键字即判为非只读。


## MongoDB 示例

对话框里内置的 MongoDB 示例本身就是 mongosh 写法（键不带引号、含 `$expr` 组合条件），可原样使用：

```javascript
db.TestCollections.find({ Age: { $gte: 25, $lte: 45 }, Example: "Text", $expr: { $and: [{ $isArray: "$Hobbies" }, { $gt: [ { $size: "$Hobbies" }, 3 ] }]}})
```

严格 JSON 写法也照旧支持：

```javascript
db.VirtualProfile.find({"Amount": {"$gt": 1000}})
  .projection({"Id": 1, "Name": 1, "Amount": 1})
  .sort({"Amount": -1})
  .limit(50)
```

只支持 `db.<集合>.find(<filter>)`，其后可链式追加 `.sort({...})` / `.limit(n)` / `.projection({...})`：
三者任意顺序、各写一次，允许换行与结尾分号，`.limit()` 的参数必须是非负整数。

**实参兼容 mongosh 写法**：键可以不带引号（`{ Age: { $gte: 25 } }`）、字符串可以用单引号、对象 / 数组末尾多余的
逗号会被忽略——解析器先按字符状态机把实参规整成严格 JSON（裸键补双引号、单引号串统一改写），再交给 JSON 解析，
因此从 mongosh 复制出来的语句可以直接粘贴。规整之后仍然必须是 JSON 字面量：
**不支持** `ObjectId()` / `ISODate()` / `NumberDecimal()` 这类构造函数（请换成普通字面量，例如日期写
`"2026-08-29T00:00:00"`），也**不支持** `aggregate`、`findOne`、`count`、`distinct` 等其它方法以及
`db.<库>.<集合>.find(...)`；不符合格式时报错信息会附上可用示例与「规整前的原文」，便于定位。


## 只读与非只读的行为差异

| 语句类型 | 预览 | 执行 | 能否作为导出数据源 |
| :--- | :--- | :--- | :--- |
| 只读（首关键字在白名单内；MongoDB 查询全部按只读处理） | **执行一次**取表头与前 3 行样本，取完立即关闭游标 | 正常执行，SQL 只读语句走服务端流式游标 | 能 |
| 非只读（`UPDATE` / `DELETE` / `INSERT` / DDL 等） | **不执行**，只提示「该语句不是只读查询,预览已跳过」 | 二次确认后才执行，改用普通客户端游标（命名游标只能 `DECLARE ... SELECT`） | 有结果集时能，没有结果集时不能 |

- **执行即提交**：非只读语句一执行成功就提交，写操作**不可回滚**；
- **没有结果集**：`UPDATE` / `DELETE` / DDL 执行后提示
  「语句已执行（影响 N 行），但没有返回结果集，无法作为导出数据源」，影响行数未知时显示 `-1`；
- **二次确认**：图形界面是「该语句不是只读查询,执行会修改数据库(语句返回结果集时仍会导出),确定使用吗?」，
  终端向导是一个默认回答为「否」的询问（连续 5 次预览失败会自动中止）。


## 与既有能力的关系

- **不裁剪字段**：查询结果就是语句返回的列（MongoDB 由 `projection` 决定），不提供字段子集勾选；
- **字段映射照常可用**：仍可在「字段映射预览 / 调整」里把查询列映射到目标字段，并做重名校验；
- **输出命名**：默认取查询标签——SQL 取第一个 `from` 后的表名（`schema.table` 保留为两段），
  取不到时用「自定义查询」；MongoDB 取集合名。仍可在输出配置里继续改名；
- **其余能力不变**：目录模式、同名文件纠错、批量写入、数值保真回读校验与失败阻断都照常生效。


## 常见错误提示与处理

| 提示 | 原因 | 处理 |
| :--- | :--- | :--- |
| `查询语句语法不正确,应以 db.<集合>.find(<filter>) 开头` | MongoDB 语句不是 `db.<集合>.find(...)`（例如写了 `aggregate`） | 按报错附带的示例改写 |
| `查询语句里 filter 不是合法 JSON: ...` | 值不是合法 JSON 字面量（例如用了 `ObjectId()` / `ISODate()` / `NumberDecimal()` 构造器，或字符串里引号没配对）；键不加引号、单引号字符串、末尾多余逗号本身是**允许**的 | 把构造器换成普通值（如十六进制串直接写 `"507f1f77bcf86cd799439011"`），按报错附带的「规整前的原文」定位改写后重新「校验并预览」 |
| `查询语句不支持 .aggregate()` / `查询语句在 ... 处有多余内容` | 用了 `find` 之外的方法，或分号后面还有第二条语句 | 只用 `sort` / `limit` / `projection`，一次只写一条语句 |
| 驱动返回的 SQL 语法错误 / 表或字段不存在 / 权限不足 | 语句本身或连接权限有问题 | 预览阶段就原样报出（图形界面给中文定位提示），改写后重新「校验并预览」 |
| `该语句不是只读查询,预览已跳过;执行会修改数据库(语句返回结果集时仍会导出)` | 首关键字不在白名单，或 `with` 里含改数据关键字 | 默认拒绝；确需写库时再在二次确认里选「是」 |
| `语句已执行(影响 N 行),但没有返回结果集,无法作为导出数据源` | 写语句 / DDL 不返回结果集 | 语句已提交；改用 `... RETURNING ...` 或只读查询后再导出 |
| `集合【X】里没有可用字段,无法作为导出数据源(请检查集合是否为空)` | 空集合，或纯排除式 `projection` 把字段全排掉了 | 确认集合非空，`projection` 改用包含写法 |
| `没有填写查询语句` | 语句为空就点了「校验并预览」 | 先填一条语句 |


# 长浮点数保真与科学计数法规避

高精度数值（多位小数、超大整数、以科学计数法书写的字符串）在任意方向转换后都必须**数值等值、文本一致**。
实现要点（代码见 `src/dt_numeric.py`）：

| 环节 | 规则 |
| :--- | :--- |
| 中间表示 | 一律使用 `decimal.Decimal` 或 Python 任意精度 `int`，**禁止二进制 `float` 作为中转容器**；只能拿到 float 时（Excel 单元格、FLOAT 列）经 `Decimal(str(value))` 取十进制字面量 |
| 解析 | CSV 文本按 `parse_number` 判定：整数 → `int`，小数 / 科学计数法 → `Decimal`；JSON 读取使用 `parse_float=Decimal`；`007` 这类补零编号识别为**文本**，不丢前导零 |
| 格式化 | 所有输出统一经 `format_number` 以**定点十进制**呈现：`1.500` 保留尾零、`1E+20` 展开为 `100000000000000000000`，永不出现 `e` / `E` |
| JSON 写出 | 使用自带序列化器 `dumps_json`（标准库无法直接输出 Decimal 数字），保证 JSON 里的数字同样是定点形式 |
| Excel 写出 | 建表时 `Decimal` → `NUMERIC/DECIMAL`（不是 DOUBLE）；单元格中超过 15 位有效数字或带尾零的数值按**文本**写出（Excel 内部是双精度，写数值必然被截断/吞尾零） |
| 数据库写出 | `Decimal` 交给驱动参数化写入（psycopg2 / PyMySQL / pymongo 的 `Decimal128`），不经 float；超出 `Decimal128`(34 位) 或 `int64` 范围的 Mongo 数值自动降级为文本保真 |
| 回读校验 | 写入后自动回读输出，按字段做「数值等值 + 文本等值」双重比对；发现精度丢失 / 截断 / 科学计数法即**阻断流程并输出差异报告**（可用 `allow_mismatch` 放行，命令行会询问是否保留结果） |

示例输出（输入 `1.500` 与 `12345678901234567890.12345`，经 CSV → JSON → Excel → CSV 往返后完全一致）：

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

> 校验报告形如：`数值保真校验:通过 (字段 2 个 / 比对 4 个值)`；未通过时列出差异样例，例如
> `Amount 行1: 输入 '1.500' → 输出 '1.5'(文本表示变化)`。


# SSH 连接远程数据库

针对「数据库只监听 `127.0.0.1`、必须先 SSH 登录服务器才能访问」的场景,程序内置 SSH 本地端口转发：
勾选 **启用 SSH 通道** 后,【数据库连接】里的主机/端口填写**远程服务器上**数据库的地址（通常 `127.0.0.1:5432`）,
本机连接会自动改走通道端口,导入导出流程与本地数据库完全一致。


## 图形界面操作步骤

1. 窗口最上方是【SSH 通道】卡片（见 图 2），勾选 **启用 SSH 通道**；
2. 填写 SSH 主机/IP、端口（默认 22）、用户名、密码（本工具统一使用「用户名 + 密码」认证，
   命令行向导 / 引擎仍保留私钥参数，图形界面不提供私钥选择）；
3. 点击 **连接 SSH**，状态区显示 `127.0.0.1:<本地端口> → SSH 用户@主机:22 → 127.0.0.1:5432`（见 图 3）；
4. 回到【数据库连接】填写远程库的地址 / 端口 / 用户名 / 密码 / 数据库名，点击 **连接并加载表**；
5. 之后按普通流程勾选「表 → 字段」，导出或导入即可（大表导出进度见 图 4、图 5）；
6. 需要换数据库类型时直接切【数据库类型】：**隧道不会断开**，程序只把转发目标改到新的「主机:端口」，
   日志提示 `SSH 隧道已重新指向 127.0.0.1:5432`（端口以实际填写为准），再点 **连接并加载表** 即可。

> 通道只监听本机 `127.0.0.1`，不会对外网暴露端口；点击 **断开 SSH** 或关闭程序会自动释放通道。

> **切换数据库类型不断开隧道**：切换时旧的数据库连接会作废（需要重新「连接并加载表」），但 SSH 通道原样沿用——
> 引擎侧 `SSHTunnel.retarget` 只重设转发目标，已建立的本地监听端口不受影响；而且每次取连接配置（连接、读取、
> 预览查询）都会先把隧道目标同步到当前填写的「主机:端口」，端口留空或越界时回退到该类型的默认端口。

> **隧道存活检查**：图形界面每 3 秒探活一次，一旦发现隧道断开，日志**只提示一次**
> `SSH 隧道已断开:<通道信息>,请重新连接`，随后自动清理数据库连接并刷新状态栏与按钮状态；探活只在隧道存在时运行，
> 因此不会反复刷屏。点 **连接 SSH** 重建通道即可继续。


## 命令行向导

`./data_transformer.sh` 在询问连接参数时会先问「是否通过 SSH 通道连接远程服务器」,交互与图形界面一致;
加 `--demo` 可直接预填上表参数（完整交互见 图 11、图 12）：

```bash
./data_transformer.sh --demo
```

> SSH 功能依赖 `paramiko`（`pip install paramiko`）。未安装时按钮置灰并给出提示,文件互转不受影响。


# 启动参数

DataTransformer 支持以下启动参数，可用于调试与故障排查。

启动参数的使用方式如下：

**源码运行：**

```bash
python data_transformer.py 参数
```

**已编译版本：**

```powershell
DataTransformer.exe 参数
```

---

## `--selftest`

无界面自检模式，用于验证打包产物中的驱动与引擎是否完整，结果写入运行目录下的 `selftest.log`。

示例：

```powershell
DataTransformer.exe --selftest
```

配合数据库环境变量可同时自检 MySQL / PostgreSQL 连通性：

```powershell
set SELFTEST_MYSQL_USER=root
set SELFTEST_MYSQL_PASSWORD=你的密码
set SELFTEST_MYSQL_DATABASE=你的库名
set SELFTEST_PG_USER=你的用户名
set SELFTEST_PG_PASSWORD=你的密码
set SELFTEST_PG_DATABASE=你的数据库名

DataTransformer.exe --selftest
```

> 不设置环境变量时，仅检查依赖导入与 Excel / JSON / CSV 写读往返。

---

## `--version`

打印版本号后退出，便于脚本化调用：

```powershell
DataTransformer.exe --version
```

Linux / macOS 下等价的写法：

```bash
./data_transformer.sh --version
```

---


# 日志与调试

程序运行日志显示在主界面"④ 运行日志"区域，可按日志级别着色显示，并支持 **一键清空**。

**出错时只输出一段中文定位提示**：日志与弹窗给的是 `describe_db_error` / `describe_ssh_error` 生成的
中文说明（通常一到两行，指明是主机 / 端口 / 用户名 / 密码 / 数据库名 / 权限哪一环出错，末行保留「原始错误: …」），
默认**不再打印 Python 回溯与源码行号**，避免把栈信息直接摊给最终用户。
若当前启用了 SSH 隧道且错误特征像连接中断（`abort` / `timeout` / `timed out` / `refused`），
会再追加一句「当前启用了 SSH 隧道,隧道可能已断开,请检查【连接 SSH】状态后重试」。

需要完整回溯排查时设置环境变量 `DT_DEBUG=1` 后重跑：

```powershell
set DT_DEBUG=1
DataTransformer.exe --selftest
```

```bash
DT_DEBUG=1 ./data_transformer.sh
```

> 图形界面在 `DT_DEBUG` 存在时把 `traceback.format_exc()` 追加到日志；命令行向导则在兜底捕获处直接重新抛出。
> `DT_DEBUG` 默认关闭，所以正常使用只会看到那段中文提示。

如果遇到导入导出异常，请保留日志内容，并在反馈问题或提交 Issue 时一并提供相关日志。


# 配置与记忆

连接成功后，程序会自动记住连接参数（Windows 图形界面记忆 **密码** 与 **数据库名**；
命令行向导还会记忆 **主机 / 端口 / 用户名 / MongoDB URI / authSource / SSH 参数**），下次启动自动填入。

配置文件位于：

```text
Windows      : %APPDATA%\DataTransformer\config.json
Linux / macOS: $XDG_CONFIG_HOME/DataTransformer/config.json（默认 ~/.config/DataTransformer/config.json）
```

删除该文件并重启程序，即可恢复首次打开状态（仅主机默认为 `127.0.0.1`，端口、用户名、密码、数据库名均为空）。

> 配置文件中密码为**明文**存储，请勿在多用户共用电脑上使用敏感密码。
> Linux / macOS 下写入配置时会自动把权限收紧为 `600`（仅本人可读写），并采用“先写临时文件再原子替换”的方式避免写坏配置。


# 安全说明

- **标识符校验**：所有表名 / 字段名先做白名单校验（`字母 / 下划线 / 中文` 开头，只允许字母、数字、下划线），
  再按数据库方言加引用符（MySQL 反引号 / PostgreSQL 双引号），避免拼接 SQL 造成注入。
  因此 Excel Sheet 名与表头列名若含空格、括号、横线等字符，需要先重命名再导入数据库
  （导出为 JSON / CSV / Excel 不受此限制）；
- **自定义查询语句**：语句由使用者自行书写并原样执行，不做标识符白名单校验，也不加引用符——这是唯一绕过上面这套校验的入口；非只读语句会改动数据库且不可回滚，请只在可信语句上使用；
- **SSH 通道**：本地监听地址固定为 `127.0.0.1`，不对局域网暴露端口；
- **配置权限**：Linux / macOS 下配置文件权限收紧为 `600`；
- **CSV 公式注入**：导出的 CSV / Excel 会**原样保留**数据内容。若源数据中带有以 `=`、`+`、`-`、`@`
  开头的文本，直接用 Excel 打开时可能被当作公式执行——这是 Excel 的既有行为。处理不受信任的数据时，
  请先用文本编辑器查看，或改用 JSON 格式导出。


# 测试

```bash
python -m pytest tests/ -q          # 全部单元测试与集成测试(当前 189 项)
python -m pytest tests/test_numeric.py -q   # 长浮点数保真专项
python -m src.dt_core               # 无界面自检(等价 --selftest),结果写入 selftest.log
```

| 测试文件 | 覆盖内容 |
| :--- | :--- |
| `tests/test_numeric.py` | Decimal 解析 / 定点格式化 / 尾零与前导零规则 / JSON 定点序列化 / Excel 单元格策略 / 保真报告与多重集比对 |
| `tests/test_engine.py` | 长浮点数跨 xlsx / json / csv 往返、科学计数法归一化、保真校验阻断与放行、字段与名称映射、多表输出、跳过与目录模式、错误护栏 |
| `tests/test_mongo.py` | MongoDB 值转换（Decimal128 / int64 溢出 / 嵌套结构 / 不安全键名）、`_id` 处理，以及**真实服务端**往返（本机无 mongod 时自动跳过） |
| `tests/test_cli.py` | 命令行向导端到端交互（管道驱动）：完整流程、输出路径、取消不落盘、无栈回溯 |
| `tests/test_gui.py` | 图形界面关键逻辑（离屏运行）：窗口构造、三种数据库配置切换、SSH 通道改写连接目标、字段映射、层级勾选、进度条生命周期 |
| `tests/test_query.py` | 自定义查询语句：SQL 只读判定（含 `WITH` 里改数据 CTE 的加固）、自由语句的流式 / 普通游标分流与提交语义、无结果集异常、MongoDB `find` 链解析与错误提示、**mongosh 兼容（裸键 / 单引号 / 尾逗号）**、`preview_query` 非只读不执行、查询源端到端转 CSV |

> 逐文件用例数：`test_cli.py` 17 / `test_engine.py` 21 / `test_gui.py` 41 / `test_mongo.py` 18 /
> `test_numeric.py` 27 / `test_query.py` 65，合计 **189**。

> 集成测试需要真实环境：PostgreSQL / MySQL / MongoDB 用例在缺少服务或环境变量时会自动跳过，不会导致失败。


# 常见问题

| 现象 | 处理办法 |
| :--- | :--- |
| `ModuleNotFoundError: No module named 'questionary'` | Linux 向导依赖 questionary：`pip install questionary`，或直接用 `./data_transformer.sh`（会自动安装） |
| `ModuleNotFoundError: No module named 'pymongo'` | 需要 MongoDB 功能时安装：`pip install pymongo`；不使用 MongoDB 可忽略 |
| 提示「非法标识符」 | 表名 / 字段名（含 Excel Sheet 名与表头）必须由字母、下划线或中文开头，且只含字母、数字、下划线；含空格 / 括号 / 横线时请先重命名，或改用 JSON / CSV / Excel 输出 |
| Excel 打开导出的文件，长数字变形或以科学计数法显示 | 本工具写出的长数值是**文本单元格**（Excel 无法精确表示超过 15 位有效数字）；如需保持数值语义请用 JSON / CSV，或在 Excel 中把该列设为文本格式再导入 |
| `数值保真校验未通过` | 说明输出侧发生了精度丢失 / 尾零丢失 / 科学计数法；日志与弹窗会列出差异样例，请检查目标列类型（应使用 NUMERIC / DECIMAL 而非 DOUBLE）或改用文本列 |
| MySQL/PostgreSQL 导入报 `value too long` | 目标表字符串列长度不足；自动建表已统一使用 TEXT，手工建表请把该列改宽 |
| 中文乱码 | Windows 终端先执行 `chcp 65001`；CSV 编码选择 `UTF-8 with BOM` 或 `GBK` |
| `无法连接 MongoDB ... ServerSelectionTimeoutError` | 检查 mongod 是否启动、端口与 `authSource`；远程场景可勾选 SSH 通道 |
| 百万行数据导入很慢 | 调大「每批行数」（默认 1000，可设 5000~20000）；数据库侧注意 `max_allowed_packet`（MySQL）与磁盘 IO |
| 出错时日志只有一段中文提示，想定位到具体代码 | 默认不打印回溯与源码行号；启动前设置环境变量 `DT_DEBUG=1`（Windows 下 `set DT_DEBUG=1`，Linux / macOS 下 `DT_DEBUG=1 ./data_transformer.sh`）后重跑，日志会追加完整 `traceback` |
| 日志出现 `SSH 隧道已断开:…,请重新连接` | 3 秒探活发现通道已断，程序已自动清理数据库连接并刷新按钮；点 **连接 SSH** 重建通道后重新「连接并加载表」（切换数据库类型本身不会断开隧道） |


# 验收清单

- [x] Windows 10/11 图形界面（PySide6）：Input 参数区、Output 参数区、运行日志 + 开始转换按钮、进度条；
- [x] Windows 使用 Nuitka 打包为单文件可执行程序（`build_exe.bat`，自动校验大小与 SHA-256）；
- [x] Linux / macOS 终端交互式工具（`data_transformer.sh` → `python -m src.dt_cli`），**questionary** 动态选择格式与参数，无需图形界面；
- [x] 覆盖 Ubuntu、Debian、Fedora、CentOS 的依赖安装与运行说明；
- [x] 适配 MySQL、PostgreSQL、**MongoDB**，以及 xlsx、csv、json 文件；
- [x] 数据库 ↔ 文件双向转换（含 SQL→SQL 表拷贝、MongoDB 集合拷贝、文件间互转）；
- [x] 映射规则：xlsx `sheetname` ↔ 数据库表名 / MongoDB 集合名；首行表头 ↔ 数据库字段名；CSV / JSON 提供等价映射并提供手动覆盖；
- [x] 长浮点数与高精度数值保真：数值等值 + 文本一致、不丢尾零、不截断有效数字、不出现 `e`/`E` 科学计数法；
- [x] 提供可见的数值校验机制与差异报告（写入后回读比对，失败可阻断或确认放行）；
- [x] 数据类型转换、空值、日期时间、编码、批量写入、错误处理、日志与进度提示；
- [x] 单元测试与集成测试（含长浮点数专项与真实数据库往返）；
- [x] 支持自由填写 SQL / MongoDB 查询语句并先校验预览（图形界面「自定义查询…」与终端向导「② 自由填写查询语句」）；
- [x] 自定义查询对话框里内置的 SQL / MongoDB 示例可直接使用（MongoDB 示例为 mongosh 写法：裸键、单引号、尾逗号均兼容）；
- [x] 切换数据库类型不断开 SSH 隧道（只重设转发目标），隧道意外断开时日志只提示一次并自动清理连接与按钮状态；


# AI 辅助开发

本项目开发过程中使用了 DeepSeek Harness 辅助编程（Vibe Coding），包括代码编写、重构、调试、问题分析等。

项目的整体设计、功能规划、代码审查与最终维护由作者负责。


# 依赖

DataTransformer 主要使用以下项目：

- [Python](https://www.python.org/)
- [PySide6](https://doc.qt.io/qtforpython-6/)
- [questionary](https://questionary.readthedocs.io/)（终端动态交互）
- [PyMySQL](https://github.com/PyMySQL/PyMySQL)
- [psycopg2](https://www.psycopg.org/)
- [pymongo](https://pymongo.readthedocs.io/)（MongoDB）
- [openpyxl](https://openpyxl.readthedocs.io/)
- [paramiko](https://www.paramiko.org/)（SSH 通道）
- [Nuitka](https://nuitka.net/)



Copyright © 2026 merry8462.
