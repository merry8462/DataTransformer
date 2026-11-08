'''建表（集合）语句 —— mongosh

use MongoData

db.createCollection("MongoCollections", {
  validator: {
    $jsonSchema: {
      bsonType: "object",
      title: "虚拟信息表",
      required: ["Name", "CreateTime"],
      properties: {
        _id:          { bsonType: "objectId",                description: "主键ID" },
        Name:         { bsonType: "string", maxLength: 20,   description: "姓名" },
        Gender:       { bsonType: "string", enum: ["男", "女"], description: "性别" },
        Age:          { bsonType: "int", minimum: 16, maximum: 99, description: "年龄" },
        Province:     { bsonType: "string", maxLength: 15,   description: "省份" },
        PhoneNumber:  { bsonType: "string", maxLength: 12,   description: "电话号码" },
        IdCard:       { bsonType: "string", maxLength: 20,   description: "身份证号" },
        Email:        { bsonType: "string", maxLength: 50,   description: "电子邮箱" },
        Occupation:   { bsonType: "string", maxLength: 30,   description: "职业" },
        UserName:     { bsonType: "string", maxLength: 30,   description: "用户名" },
        PrivateIP:    { bsonType: "string", maxLength: 15,   description: "IPv4地址" },
        MacAddress:   { bsonType: "string", maxLength: 17,   description: "MAC地址" },
        Birthday:     { bsonType: "date",                    description: "生日" },
        Company:      { bsonType: "string", maxLength: 30,   description: "公司名称" },
        Bank:         { bsonType: "string", maxLength: 20,   description: "银行" },
        BankAccount:  { bsonType: "string", maxLength: 30,   description: "银行账户" },
        LicensePlate: { bsonType: "string", maxLength: 10,   description: "车牌号" },
        Address:      { bsonType: "string", maxLength: 100,  description: "地址" },
        Hobbies:      { bsonType: "array", items: { bsonType: "string" }, minItems: 1, maxItems: 5, description: "爱好(1~5个)" },
        Remark:       { bsonType: "string", maxLength: 200,  description: "备注" },
        CreateTime:   { bsonType: "date",                    description: "创建时间" }
      }
    }
  }
})

db.MongoCollections.createIndex({ "CreateTime": -1 })
'''


from faker import Faker
from pymongo import MongoClient
from typing import Any, Dict, List
from sys import argv
from pathlib import Path
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from datetime import datetime, time
from colorama import init, Fore, Style

init(autoreset=True)

# ==================== MongoDB 连接参数 ====================
MONGO_URI = "mongodb://127.0.0.1:27017"
DB_NAME = "MongoData"
COLLECTION_NAME = "MongoCollections"

# 性别与爱好取值池，可按需增删；爱好明细统一用英文
GENDERS = ("男", "女")
HOBBIES = (
    "Reading", "Swimming", "Running", "Basketball", "Football", "Badminton",
    "Painting", "Calligraphy", "Singing", "Guitar", "Piano", "Cooking",
    "Baking", "Travel", "Fishing", "Chess", "Movies", "Anime", "Gaming",
    "Dancing", "Yoga", "Gardening", "Pets", "Handcraft", "Programming",
    "Skiing", "Skating", "Surfing", "Diving", "Tennis", "Archery",
    "Rock Climbing", "Karaoke", "Magic", "Origami", "Scrapbooking",
    "Woodworking", "Pottery", "Calligraphy", "Quilting", "Knitting",
    "Sewing", "Jewelry Making", "Model Building", "Puzzle Solving",
    "Card Games", "Board Games", "Cosplay", "Traveling", "Language Learning",
    "Meditation", "Bird Watching", "Stargazing", "Wine Tasting", "Beer Brewing",
    "Cheese Making", "Candle Making", "Soap Making", "Perfume Making",
    "Leatherworking", "Metalworking", "Glassblowing", "Sculpting",
    "Photography", "Videography", "Blogging", "Podcasting", "Volunteering",
    "Charity Work", "Community Service", "Public Speaking", "Debating",
    "Teaching", "Mentoring", "Coaching", "Fitness Training",
    "Martial Arts", "Self-Defense", "First Aid", "Survival Skills", "Camping",
    "Hiking", "Backpacking", "Rock Climbing", "Caving", "Sailing", "Kayaking",
    "Canoeing", "Rafting", "Paragliding", "Skydiving", "Bungee Jumping", "Scuba Diving"
)
HOBBIES_MIN, HOBBIES_MAX = 1, 5
AGE_MIN, AGE_MAX = 16, 99

num = int(argv[1]) if len(argv) > 1 and argv[1].isdigit() else 15

fk = Faker(locale="zh_CN")
data: List[Dict[str, Any]] = [
    {
        "Name": fk.name(),
        "Gender": fk.random_element(elements=GENDERS),
        "Age": fk.random_int(min=AGE_MIN, max=AGE_MAX),
        "Province": fk.province(),
        "PhoneNumber": fk.phone_number(),
        "IdCard": fk.ssn(),
        "Email": fk.ascii_free_email(),
        "Occupation": fk.job(),
        "UserName": fk.user_name(),
        "PrivateIP": fk.ipv4_private(),
        "MacAddress": fk.mac_address(),
        "Birthday": fk.date_of_birth(minimum_age=18, maximum_age=65),
        "Company": fk.company(),
        "Bank": fk.bank(),
        "BankAccount": fk.bban(),
        "LicensePlate": fk.license_plate(),
        "Address": fk.address(),
        "Hobbies": fk.random_elements(
            elements=HOBBIES,
            length=fk.random_int(min=HOBBIES_MIN, max=HOBBIES_MAX),
            unique=True
        ),
        "Remark": fk.text(max_nb_chars=50)
    } for _ in range(num)
]

# ==================== Excel 导出（不变） ====================
headers = ["Name", "Gender", "Age", "Province", "PhoneNumber", "IdCard", "Email", "Occupation", "UserName", "PrivateIP", "MacAddress", "Birthday", "Company", "Bank", "BankAccount", "LicensePlate", "Address", "Hobbies", "Remark"]
excel_folder = Path(__file__).parent / "Excel"
excel_folder.mkdir(exist_ok=True)

excel_filename = f"{datetime.now().strftime('%Y-%m-%d_%H%M%S')}.xlsx"
excel_path = excel_folder / excel_filename

try:
    wb = Workbook()
    ws = wb.active
    ws.title = "MongoCollections"
    ws.append(headers)
    for row in data:
        values = [
            ", ".join(row[header]) if isinstance(row[header], list) else row[header]
            for header in headers
        ]
        ws.append(values)
    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            if cell.value:
                cell_len = len(str(cell.value))
                if cell_len > max_len:
                    max_len = cell_len
        ws.column_dimensions[col_letter].width = max_len + 2
    wb.save(excel_path)
    print(f"{Fore.GREEN}数据已成功导出到Excel文件：<{excel_path}>。{Style.RESET_ALL}")
except Exception as e:
    print(f"导出 Excel 文件时出错：{e}")

# ==================== 数据库连接 ====================
client = MongoClient(
    host="127.0.0.1", port=27017,
    username="merry8462", password="8462",
    authSource="admin",
)
# ==================== 插入 MongoCollections ====================
try:
    client.admin.command("ping")
    documents = [
        {
            **row,
            "Birthday": datetime.combine(row["Birthday"], time.min),
            "CreateTime": datetime.now()
        } for row in data
    ]
    result = client[DB_NAME][COLLECTION_NAME].insert_many(documents, ordered=True)
    print(f"{Fore.BLUE}{Style.BRIGHT}成功插入了 <{len(result.inserted_ids)}> 条数据。{Style.RESET_ALL}")
except Exception as err:
    print(f"插入数据时发生错误：<{err}>")

# ==================== 关闭连接 ====================
client.close()
