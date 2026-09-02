# data_generator.py 使用说明

> 虚拟数据生成脚本:批量生成中文虚拟个人信息,同时导出 Excel 并写入 MySQL,纯终端运行,无图形界面。

## 功能

- 使用 [Faker](https://faker.readthedocs.io/)(`zh_CN`)随机生成虚拟个人信息
- 导出 Excel:自动创建 `Excel/` 目录,按时间戳命名 `.xlsx` 文件(Sheet 名 `VirtualProfile`,自动列宽)
- 写入 MySQL:批量 `executemany` 插入 `VirtualProfile` 表,失败自动回滚
- 终端输出带颜色提示(绿色 = Excel 导出成功,蓝色 = 数据库插入成功)

## 环境依赖

```bash
pip install faker pymysql openpyxl colorama
```

> 建议 Python 3.8+;Windows 终端如需正常显示中文,可先执行 `chcp 65001`。

## 数据库准备

脚本写入的 MySQL 表结构(脚本文件头部注释中附有完整建表语句):

```sql
CREATE TABLE VirtualProfile (
    Id INT NOT NULL AUTO_INCREMENT COMMENT '主键ID，自增',
    Name VARCHAR(50) COMMENT '姓名',
    Province VARCHAR(50) COMMENT '省份',
    PhoneNumber VARCHAR(20) COMMENT '电话号码',
    IdCard VARCHAR(30) COMMENT '身份证号',
    Email VARCHAR(100) COMMENT '电子邮箱',
    Birthday DATE COMMENT '生日',
    Company VARCHAR(100) COMMENT '公司名称',
    Bank VARCHAR(100) COMMENT '银行',
    BankAccount VARCHAR(100) COMMENT '银行账户',
    LicensePlate VARCHAR(100) COMMENT '车牌号',
    Address VARCHAR(200) COMMENT '地址',
    CreateTime TIMESTAMP DEFAULT CURRENT_TIMESTAMP COMMENT '创建时间',
    PRIMARY KEY (Id)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci COMMENT='虚拟信息表';
```

**数据库连接参数硬编码在脚本中**,请按需修改 `data_generator.py` 第 85~91 行:

```python
connection = connect(
    host="127.0.0.1",   # 数据库主机
    user="root",        # 用户名
    password="123",     # 密码
    database="db",      # 数据库名
    charset="utf8mb4"
)
```

## 使用方法

在终端中运行:

```bash
python data_generator.py [生成条数]
```

- 不传参数时默认生成 **15** 条:

```bash
python data_generator.py
```

- 指定生成条数(正整数):

```bash
python data_generator.py 1000
```

> 参数不是数字时同样回退为默认 15 条。

## 输出说明

### ① Excel 文件

- 输出目录:`脚本所在目录/Excel/`
- 文件名:`YYYY-MM-DD_HHMMSS.xlsx`(按生成时间命名)
- Sheet 名:`VirtualProfile`
- 表头 11 列:`Name, Province, PhoneNumber, IdCard, Email, Birthday, Company, Bank, BankAccount, LicensePlate, Address`
- 列宽按内容自动调整

### ② MySQL 表

- 目标表:`VirtualProfile`
- 写入 11 列(与 Excel 表头一致,`Id` 自增、`CreateTime` 由数据库默认值填充)
- 全部成功才提交,出错自动回滚

## 字段与数据来源

| 字段 | 含义 | Faker 方法 |
| --- | --- | --- |
| Name | 姓名 | `fk.name()` |
| Province | 省份 | `fk.province()` |
| PhoneNumber | 手机号 | `fk.phone_number()` |
| IdCard | 身份证号 | `fk.ssn()` |
| Email | 邮箱 | `fk.ascii_free_email()` |
| Birthday | 生日(18~65 岁) | `fk.date_of_birth(minimum_age=18, maximum_age=65)` |
| Company | 公司 | `fk.company()` |
| Bank | 银行名称 | `fk.bank()` |
| BankAccount | 银行账号 | `fk.bban()` |
| LicensePlate | 车牌号 | `fk.license_plate()` |
| Address | 地址 | `fk.address()` |

## 运行示例

```text
数据已成功导出到Excel文件:<D:\...\Demo\Excel\xxx.xlsx>。
成功插入了 <1000> 条数据。
```

## 常见问题

| 现象 | 处理办法 |
| --- | --- |
| `插入数据时发生错误: ... Connection refused` | MySQL 未启动,或 host/port 配置有误 |
| `Access denied for user ...` | 修改脚本中的 `user` / `password` |
| `Unknown database 'db'` | 修改脚本中的 `database`,或先创建对应数据库 |
| `Table 'db.VirtualProfile' doesn't exist` | 先执行上文建表语句,或检查表名大小写 |
| 中文输出乱码 | 先执行 `chcp 65001`,或使用 Windows Terminal |
| Excel 打开文件报"格式损坏" | 生成过程中被杀毒软件锁定文件,重新运行生成即可 |
