from __future__ import annotations

import argparse
import json
import re
import sys
from collections import Counter
from datetime import date
from pathlib import Path

import requests
from bs4 import BeautifulSoup

COLUMNS = ["平台", "店铺", "品牌", "型号", "商品名称", "价格", "销量原值", "销量", "口径", "商品链接", "抓取日期"]

PLATFORM_ALIAS = {"JD": "京东", "JD.COM": "京东", "京东": "京东",
                  "TB": "淘宝", "TM": "淘宝", "TMALL": "淘宝", "TAOBAO": "淘宝", "淘宝": "淘宝", "天猫": "淘宝"}
BRAND_ALIAS = {"ASUS": "华硕", "华硕": "华硕", "MSI": "微星", "微星": "微星",
               "THINKPAD": "联想", "LENOVO": "联想", "联想": "联想",
               "GIGABYTE": "技嘉", "技嘉": "技嘉", "AORUS": "技嘉"}

# 页面导航模块会被误当成商品, 标题命中这些词就丢弃
NAV_NOISE_RE = re.compile(r"全部分类|^首页\s|所有宝贝|返回顶部|客服|购物车")

NUM = r"([\d.]+\s*万|[\d.]+)"
# (正则, 口径) —— 口径决定这一列数字的含义, 不跨平台求和
SALES_RULES = (
    (rf"已售\s*{NUM}", "已售"),
    (rf"月销\s*{NUM}", "月销"),
    (rf"{NUM}\s*\+?\s*人付款", "人付款"),
    (rf"{NUM}\s*\+?\s*人收货", "人收货"),
    (rf"销量\s*{NUM}", "销量"),
    (rf"累计评价\s*{NUM}", "评价数"),
    (rf"{NUM}\s*\+?\s*条评价", "评价数"),
)
# 口径由命中的规则决定(已售/月销/人付款/评价数), 不再按平台硬套默认值

JD_ITEM = re.compile(r"item\.jd\.com/(\d+)\.html")
TB_ITEM = re.compile(r"(?:item\.taobao\.com|detail\.tmall\.com)/item\.htm\?[^\"']*\bid=(\d+)")
# 型号尾巴只吃字母数字, 不能用 \w(会把"台式主机""电竞显示器"一起吞掉)
MODEL_RE = re.compile(
    r"(天选\s*[A-Za-z0-9-]{0,12}|幻\s*\d+[A-Za-z0-9-]{0,12}|魔霸\s*[A-Za-z0-9-]{0,12}|"
    r"枪神\s*[A-Za-z0-9-]{0,12}|无畏\s*[A-Za-z0-9-]{0,12}|灵耀\s*[A-Za-z0-9-]{0,12}|"
    r"无双\s*[A-Za-z0-9-]{0,12}|破晓\s*[A-Za-z0-9-]{0,12}|adolbook\s*[A-Za-z0-9-]{0,12}|a豆\s*[A-Za-z0-9-]{0,12}|"
    r"泰坦\s*[A-Za-z0-9-]{0,12}|绝影\s*[A-Za-z0-9-]{0,12}|雷影\s*[A-Za-z0-9-]{0,12}|"
    r"神影\s*[A-Za-z0-9-]{0,12}|星影\s*[A-Za-z0-9-]{0,12}|尊爵\s*[A-Za-z0-9-]{0,12}|"
    r"侠客\s*[A-Za-z0-9-]{0,12}|创造者\s*[A-Za-z0-9-]{0,12}|"
    r"拯救者\s*[A-Za-z0-9-]{0,12}|小新\s*[A-Za-z0-9-]{0,12}|YOGA\s*[A-Za-z0-9-]{0,12}|"
    r"ThinkBook\s*[A-Za-z0-9-]{0,12}|ThinkPad\s*[A-Za-z0-9-]{0,12}|Legion\s*[A-Za-z0-9-]{0,12}|"
    r"AORUS\s*[A-Za-z0-9-]{0,12}|猎鹰\s*[A-Za-z0-9-]{0,12}|傲世\s*[A-Za-z0-9-]{0,12}|"
    r"小金刚\s*[A-Za-z0-9-]{0,12}|大金刚\s*[A-Za-z0-9-]{0,12}|ProArt\s*[A-Za-z0-9-]{0,12}|"
    r"TUF\s*[A-Za-z0-9-]{0,12}|ROG\s*[A-Za-z0-9-]{0,12}|"
    r"MateBook\s*[A-Za-z0-9-]{0,12}|MacBook\s*[A-Za-z0-9-]{0,12}|"
    r"RedmiBook\s*[A-Za-z0-9-]{0,12}|MagicBook\s*[A-Za-z0-9-]{0,12})", re.I)
MODEL_KEYWORD_RE = re.compile(
    r"^(天选|幻|魔霸|枪神|无畏|灵耀|无双|破晓|adolbook|a豆|泰坦|绝影|雷影|神影|星影|尊爵|侠客|创造者|"
    r"拯救者|小新|YOGA|ThinkBook|ThinkPad|Legion|AORUS|猎鹰|傲世|小金刚|大金刚|ProArt|TUF|ROG|"
    r"MateBook|MacBook|RedmiBook|MagicBook)", re.I)
BRAND_STRIP_RE = re.compile(
    r"^\s*(?:【[^】]*】|\[[^\]]*\]|\s)*\s*(?:Asus|ASUS|华硕|微星|MSI|联想|Lenovo|ThinkPad|戴尔|DELL|惠普|HP|"
    r"华为|HUAWEI|苹果|Apple|宏碁|Acer|机械革命|雷神|神舟)(?:\s*（[^）]*）|\s*\([^)]*\))?\s*", re.I)
CODE_RE = re.compile(r"([A-Za-z]{1,6}-?\d{2,5}[A-Za-z0-9-]{0,8})")
CODE_BLACKLIST = re.compile(r"^(?:RTX|GTX|MX|DDR|USB|HDMI|WIFI|WIN|OLED|IPS|SSD|HDD|IPS|HZ|GB|TB|AI)",
                            re.I)
LAPTOP_RE = re.compile(r"笔记本|游戏本|轻薄本|全能本|商务本|笔记本电脑|本本|laptop|ThinkBook|ThinkPad|小新|拯救者|YOGA|无畏|灵耀|天选|泰坦|绝影|雷影|尊爵|\d{2}(?:\.\d)?\s*英寸", re.I)
NOT_LAPTOP_RE = re.compile(r"台式机|主机|主板|显卡|显示器|机械键盘|无线鼠标|鼠标垫|散热器|内存条|固态硬盘|U盘|路由器|插排|电源适配器|机箱|耳机|摄像头|电脑包|手写板|扩展坞|贴膜|保护套|支架", re.I)
# 标题里出现这些词就是笔记本本身, 优先于排除词("满血RTX5070Ti显卡"里的"显卡"不该判成显卡商品)
STRONG_LAPTOP_RE = re.compile(r"笔记本电脑|游戏本|轻薄本|全能本|商务本|电竞本|笔记本|本本|laptop", re.I)
ACCESSORY_TAIL_RE = re.compile(r"(支架|贴膜|保护套|电脑包|扩展坞|鼠标垫|手写板|散热器|电源适配器)\s*$")
PRICE_RE = re.compile(r"[¥￥]\s*([\d,]+(?:\.\d{1,2})?)")
# 京东列表页里裸 ¥ 常来自券/分期等噪声, 只认带价签的数字
PRICE_LABELED_RE = re.compile(r"(?:京东价|到手价|售价|价格|促销价|活动价)\s*[：:]?\s*[¥￥]\s*([\d,]+(?:\.\d{1,2})?)")
TITLE_NOISE_RE = re.compile(r"^(?:\s*(?:关注|对比|收藏|加购|包邮|预售)\s*)+")
TITLE_CUT_RE = re.compile(r"(?:京东价|好评率|加入购物车|立即购买|券\s*[¥￥]|赠\s*$).*$", re.S)


def load_html(path: Path) -> str:
    """按文件声明的 charset 解码, 淘宝系页面多为 GBK。"""
    raw = path.read_bytes()
    m = re.search(rb'charset=["\']?([\w-]+)', raw[:4000], re.I)
    declared = m.group(1).decode("ascii", "ignore") if m else "utf-8"
    for encoding in (declared, "utf-8", "gb18030"):
        try:
            return raw.decode(encoding)
        except (UnicodeDecodeError, LookupError):
            continue
    return raw.decode("gb18030", errors="ignore")


def parse_count(text: str) -> int:
    if not text:
        return 0
    cleaned = text.replace(",", "").replace("+", "").replace(" ", "")
    multiplier = 1
    if cleaned.endswith("万"):
        cleaned, multiplier = cleaned[:-1], 10000
    try:
        return int(float(cleaned) * multiplier)
    except ValueError:
        return 0


def find_sales(text: str) -> tuple:
    """返回 (原值, 数值, 口径)。找不到就 ("", "", "")。"""
    for pattern, caliber in SALES_RULES:
        match = re.search(pattern, text)
        if match:
            raw = match.group(1).replace(" ", "")
            value = parse_count(raw)
            if value:
                return raw, value, caliber
    return "", "", ""


def find_price(text: str, labeled_only: bool = False) -> str:
    match = (PRICE_LABELED_RE if labeled_only else PRICE_RE).search(text)
    return match.group(1).replace(",", "") if match else ""


def clean_title(text: str) -> str:
    text = TITLE_CUT_RE.sub("", text or "")
    # 卡片首行常是翻页箭头, 价格是字体反爬乱码: 到第一个货币符就截断
    text = re.split(r"[¥￥]", text, maxsplit=1)[0]
    text = TITLE_NOISE_RE.sub("", text.strip()).strip()
    return re.sub(r"^[<>‹›«»\s\-—*·|]+", "", text).strip()


def is_laptop_title(name: str) -> bool:
    if NAV_NOISE_RE.search(name) or ACCESSORY_TAIL_RE.search(name):
        return False
    if STRONG_LAPTOP_RE.search(name):
        return True
    if NOT_LAPTOP_RE.search(name):
        return False
    return bool(LAPTOP_RE.search(name) or find_model(name))


def find_model(title: str) -> str:
    """型号: 先按系列关键词匹配, 匹配不到就从标题里找"字母+数字"的型号码。"""
    if not title:
        return ""
    found = MODEL_RE.findall(title)
    if found:
        # 取最长的一条: "ROG枪神10" 里应选"枪神10"而不是"ROG"
        raw = max(found, key=len)
        spaced = bool(re.search(r"\s", raw))
        model = re.sub(r"\s+", "", raw)
        # "TUF GAMING" 是品牌+品类, 不是型号: 只留关键词那截
        if spaced and not re.search(r"\d", model):
            keyword = MODEL_KEYWORD_RE.match(model)
            if keyword:
                model = keyword.group(1)
        return model[:20]

    cleaned = BRAND_STRIP_RE.sub("", title)
    for code in CODE_RE.findall(cleaned):
        if len(code) < 4 or CODE_BLACKLIST.match(code):
            continue
        return code[:20]
    return ""


def shop_name_from_file(path: Path) -> str:
    name = path.stem
    name = re.sub(r"\s*-\s*(京东|天猫\s*Tmall\.com|淘宝|Tmall\.com|淘宝网).*$", "", name, flags=re.I)
    name = re.sub(r"^(默认宝贝分类页|首页|笔记本|商品详情)-", "", name)
    return name.strip()


def extract_tmall_new(soup) -> list:
    """天猫新版卡片: 类名带哈希前缀(container--/price--/sales--), 用前缀匹配。"""
    cards = soup.select("a[class*=itemCard--], a[class*=shop-item-card]")
    if not cards:
        cards = soup.select("div[class^=container--]")

    rows = []
    for card in cards:
        link = card if card.name == "a" and card.get("href") else card.select_one("a[href*='item.htm']")
        if not link:
            continue
        match = TB_ITEM.search(link["href"])
        if not match:
            continue
        title = ""
        for el in card.select("[title]"):
            candidate = (el.get("title") or "").strip()
            if len(candidate) >= 8:
                title = candidate
                break
        if not title:
            title = card.get_text(" ", strip=True)[:80]
        price_el = card.select_one("span[class^=price--]") or card.select_one("span[class*=price--]")
        sales_el = card.select_one("span[class^=sales--]") or card.select_one("[class*=sales--]")
        sales_raw, sales, caliber = find_sales(sales_el.get_text(" ", strip=True)) if sales_el else ("", "", "")
        rows.append({
            "商品名称": clean_title(title)[:100],
            "价格": find_price(price_el.get_text(" ", strip=True)) if price_el else "",
            "销量原值": sales_raw, "销量": sales, "口径": caliber,
            "商品链接": link["href"], "_item_id": match.group(1),
        })
    return rows


def extract_tmall_legacy(soup) -> list:
    """天猫/淘宝旧版类目页: dl.item 列表, 有商品名, 价格多为字体反爬, 没有销量。"""
    rows = []
    for dl in soup.select("dl.item"):
        link = dl.select_one("a[href*='item.htm?id=']") or dl.select_one("a[href*='item.htm']")
        if not link:
            continue
        match = TB_ITEM.search(link["href"])
        if not match:
            continue
        name_el = dl.select_one("dd.detail a") or dl.select_one("img[alt]")
        title = (name_el.get_text(" ", strip=True) if name_el and name_el.name == "a"
                 else (name_el.get("alt", "") if name_el else ""))
        if not title:
            detail = dl.select_one("dd.detail")
            title = detail.get_text(" ", strip=True)[:100] if detail else ""
        text = dl.get_text(" ", strip=True)
        sales_raw, sales, caliber = find_sales(text)
        rows.append({
            "商品名称": clean_title(title)[:100],
            "价格": find_price(text),
            "销量原值": sales_raw, "销量": sales, "口径": caliber,
            "商品链接": link["href"], "_item_id": match.group(1),
        })
    return rows


def extract_jd_list(soup) -> list:
    """京东店铺/列表页: 商品可能是图片热区(无独立信息), 有正经商品模块时才提取。"""
    rows = []
    for anchor in soup.find_all("a", href=True):
        match = JD_ITEM.search(anchor["href"])
        if not match:
            continue
        node, card = anchor, None
        for _ in range(7):
            node = node.parent
            if node is None or node.name in ("body", "html", "[document]"):
                break
            text = node.get_text(" ", strip=True)
            if len(text) < 10:
                continue
            if len(text) > 600:
                break
            card = node
            break
        if card is None:
            continue
        text = card.get_text(" ", strip=True)
        title = clean_title(anchor.get("title") or anchor.get_text(" ", strip=True) or text)
        if len(title) < 4:
            img = card.select_one("img[alt]")
            title = clean_title(img.get("alt", "") if img else text)
        sales_raw, sales, caliber = find_sales(text)
        rows.append({
            "商品名称": title[:100],
            "价格": find_price(text, labeled_only=True),
            "销量原值": sales_raw, "销量": sales, "口径": caliber,
            "商品链接": anchor["href"], "_item_id": match.group(1),
        })
    return rows


# 详情页的商品 ID 可能藏在链接/canonical/JS 变量里, 都要试
DETAIL_ID_PATTERNS = (
    JD_ITEM,
    TB_ITEM,
    re.compile(r'"skuId"\s*:\s*"?(\d{6,})"?'),
    re.compile(r'"productId"\s*:\s*"?(\d{6,})"?'),
    re.compile(r'"itemId"\s*:\s*"?(\d{6,})"?'),
    re.compile(r'rel=["\']canonical["\'][^>]*[?&/](\d{6,})'),
    re.compile(r'[?&]id=(\d{6,})'),
)


# 推荐位/猜你喜欢里的价格和评价是别的商品的, 必须先剔掉再取价
RECOMMEND_RE = re.compile(r"recommend|guess-you|similar-goods|ad-banner", re.I)


def strip_recommendations(soup) -> None:
    for element in soup.find_all(True):
        if element.attrs is None:
            continue
        marker = " ".join(element.get("class") or []) + " " + (element.get("id") or "")
        if marker.strip() and RECOMMEND_RE.search(marker):
            element.decompose()


TITLE_TAIL_RE = re.compile(r"\s*[-_|]\s*(京东|淘宝|天猫|Tmall\.com|tmall\.com|淘宝网).*$|【行情[^】]*】", re.I)
TITLE_BAD_RE = re.compile(r"计算器|购物车|优惠前|登录|注册|搜索|客服|进店逛逛")


def _detail_title_candidates(soup) -> list:
    candidates = []
    for selector in ("div.sku-name", "[class*=sku-name]", "[class*=mainTitle]", "[class*=itemTitle]",
                     "[class*=goodsName]", "[class*=title--]", "h1", "title"):
        for element in soup.select(selector)[:3]:
            candidates.append(element.get_text(" ", strip=True))
    for meta in soup.find_all("meta"):
        if (meta.get("property") or meta.get("name") or "") in ("og:title", "title"):
            candidates.append(meta.get("content") or "")
    return candidates


def pick_detail_title(candidates: list) -> str:
    best = ""
    for raw in candidates:
        text = clean_title(TITLE_TAIL_RE.sub("", raw or ""))
        if len(text) < 8 or TITLE_BAD_RE.search(text):
            continue
        score = 0
        if LAPTOP_RE.search(text):
            score += 2
        if 15 <= len(text) <= 110:
            score += 1
        if score >= 2 and len(text) > len(best):
            best = text
    return best


def extract_detail(soup, html: str, path: Path = None) -> list:
    """商品详情页: 一个页面一个商品, 标题/价格/销量都在页面上。"""
    strip_recommendations(soup)
    title = pick_detail_title(_detail_title_candidates(soup))

    price = ""
    for selector in ("span.price", "div.price", "span[class*=price]", "em[class*=price]"):
        element = soup.select_one(selector)
        if element is None:
            continue
        price = find_price(element.get_text(" ", strip=True))
        if price:
            break
    text = soup.get_text(" ", strip=True)
    if not price:
        price = find_price(text)

    sales_raw, sales, caliber = find_sales(text)
    platform_is_jd = bool(re.search(r"jd\.com", html[:300000], re.I))
    item_id = ""
    for pattern in DETAIL_ID_PATTERNS:
        match = pattern.search(html)
        if match:
            item_id = match.group(1)
            break
    if not item_id and path is not None:
        digits = re.findall(r"\d{6,}", path.stem)
        item_id = digits[0] if digits else ""

    return [{
        "商品名称": title[:100],
        "价格": price,
        "销量原值": sales_raw, "销量": sales, "口径": caliber,
        "商品链接": (f"https://item.jd.com/{item_id}.html" if platform_is_jd else
                     f"https://item.taobao.com/item.htm?id={item_id}") if item_id else "",
        "_item_id": item_id,
    }]


def parse_page(path: Path, platform: str, brand: str, kind: str) -> list:
    html = load_html(path)
    soup = BeautifulSoup(html, features="html.parser")

    if kind == "detail":
        rows = extract_detail(soup, html, path)
    elif platform == "京东":
        rows = extract_jd_list(soup)
    else:
        rows = extract_tmall_new(soup) + extract_tmall_legacy(soup)

    deduped, seen = [], set()
    for row in rows:
        key = row.get("_item_id") or row["商品链接"] or row["商品名称"]
        if key in seen:
            continue
        seen.add(key)
        deduped.append({
            "平台": platform,
            "店铺": shop_name_from_file(path),
            "品牌": brand,
            "型号": find_model(row["商品名称"]),
            "商品名称": row["商品名称"],
            "价格": row["价格"],
            "销量原值": row["销量原值"],
            "销量": row["销量"],
            "口径": row["口径"],   # 口径只跟真实销量一起存在, 无销量就留空
            "商品链接": row["商品链接"],
            "抓取日期": date.today().isoformat(),
            "_item_id": row.get("_item_id", ""),
            "_kind": kind,
        })

    # 图片热区页: 多个商品链接共享同一段模块文字, 不是独立商品, 判为无数据
    if kind == "list" and deduped:
        titles = {r["商品名称"] for r in deduped}
        if len(titles) < max(2, len(deduped) * 0.2):
            return []
    return deduped


def merge_rows(rows: list) -> list:
    """同一商品可能既有店铺页行(有名字)又有详情页行(有销量), 合并成一行。"""
    merged = {}
    for row in rows:
        key = (row["平台"], row.get("_item_id") or row["商品链接"] or row["商品名称"])
        current = merged.get(key)
        if current is None:
            merged[key] = dict(row)
            continue
        for field in ("商品名称", "价格", "销量", "销量原值", "口径", "商品链接"):
            if not current.get(field) and row.get(field):
                current[field] = row[field]
        # 标题取更长更完整的那条
        if len(row.get("商品名称", "")) > len(current.get("商品名称", "")):
            current["商品名称"] = row["商品名称"]
        if row.get("_kind") == "detail":
            current["_kind"] = "detail"
    return list(merged.values())


# ---------------- 浏览器抓取 (DrissionPage) ----------------

# ---------------- 浏览器抓取 (DrissionPage) ----------------
# 全部走元素 API(page.eles / ele.attr / ele.parent / scroll), 不使用 run_js:
# 实测在某些环境下 run_js 恒返回 None, 而元素定位正常。

PRODUCT_CSS = ('css:a[href*="item.jd.com/"], a[href*="item.taobao.com/item.htm"], '
               'a[href*="detail.tmall.com/item.htm"]')
# 店铺页上"分类"链接的形态: 京东 view_search / index-xxx?categoryId=, 淘宝 category-/catName
CATEGORY_HREF_RE = re.compile(
    r"view_search|categoryId=|/category-|catName=|cat_id=|search\.htm|category\.htm", re.I)
BRAND_WORDS = (("华硕", "华硕"), ("ASUS", "华硕"), ("微星", "微星"), ("MSI", "微星"),
               ("联想", "联想"), ("Lenovo", "联想"), ("ThinkPad", "联想"),
               ("小米", "小米"), ("Xiaomi", "小米"), ("红米", "小米"), ("Redmi", "小米"),
               ("iQOO", "iQOO"), ("华为", "华为"), ("HUAWEI", "华为"), ("荣耀", "荣耀"),
               ("OPPO", "OPPO"), ("vivo", "vivo"), ("戴尔", "戴尔"), ("DELL", "戴尔"),
               ("惠普", "惠普"), ("HP", "惠普"), ("机械革命", "机械革命"), ("雷神", "雷神"))


def normalize_url(url: str) -> str:
    """同一 URL 出现重复查询参数时(如两个 catName), 只保留最后一个, 否则服务器会取错分类。"""
    if not url:
        return ""
    url = url.split("#")[0]
    if "?" not in url:
        return url
    base, query = url.split("?", 1)
    latest = {}
    for pair in [p for p in query.split("&") if p]:
        latest[pair.split("=", 1)[0]] = pair
    return base + "?" + "&".join(latest.values())


def store_list_url(store_url: str) -> str:
    """淘宝/天猫店铺的"本店所有商品"页, 商品都在 dl.item 容器里。"""
    match = re.match(r"(https?://[^/]+)", store_url or "")
    if not match:
        return ""
    host = match.group(1)
    if "tmall.com" in host or "taobao.com" in host:
        return host + "/search.htm?search=y"
    return ""


def store_category_url(store_url: str) -> str:
    """淘宝/天猫店铺的"所有分类"页, 分类侧栏在这里。"""
    match = re.match(r"(https?://[^/]+)", store_url or "")
    if not match:
        return ""
    host = match.group(1)
    if "tmall.com" in host or "taobao.com" in host:
        return host + "/category.htm"
    return ""


def guess_brand(*texts) -> str:
    for text in texts:
        if not text:
            continue
        for word, brand in BRAND_WORDS:
            if word in text:
                return brand
    return ""


def launch_browser(attach: str = "", headless: bool = False, profile: str = ".dp_profile"):
    from DrissionPage import ChromiumPage, ChromiumOptions
    if attach:
        return ChromiumPage(attach)
    options = ChromiumOptions()
    options.set_user_data_path(str(Path(profile).resolve()))
    options.auto_port()
    options.set_argument("--no-first-run")
    if headless:
        options.headless()
    return ChromiumPage(options)


def absolute(href: str, page_url: str = "") -> str:
    if not href:
        return ""
    if href.startswith("//"):
        return "https:" + href
    if href.startswith("/"):
        match = re.match(r"(https?://[^/]+)", page_url or "")
        return (match.group(1) if match else "") + href
    return href


def finding_links(scope) -> list:
    try:
        return scope.eles(PRODUCT_CSS) or []
    except Exception:
        return []


def safe_text(element) -> str:
    try:
        return (element.text or "").strip()
    except Exception:
        return ""


# 推荐位/猜你喜欢/热卖榜里的商品不属于当前分类, 必须排除
RECOMMEND_MARK = re.compile(r"recommend|guess|similar|you-?like|hot-?sale|rank|ad-?box|promo|推荐|猜你喜欢", re.I)


def in_recommend_block(element) -> bool:
    node = element
    for _ in range(10):
        try:
            node = node.parent()
        except Exception:
            return False
        if node is None:
            return False
        marker = f"{node.attr('class') or ''} {node.attr('id') or ''}".strip()
        if marker and RECOMMEND_MARK.search(marker):
            return True
    return False


def card_element(anchor):
    """从商品链接向上找卡片容器。"""
    for level in range(1, 8):
        try:
            parent = anchor.parent(level)
        except Exception:
            break
        if parent is None:
            break
        text = safe_text(parent)
        if len(text) > 600:
            break
        if len(text) >= 12:
            return parent
    return None


def card_text(anchor) -> str:
    card = card_element(anchor)
    return safe_text(card) if card is not None else safe_text(anchor)


# 卡片里混着的导航/按钮/库存字样, 不能当商品名
TITLE_NOISE = re.compile(
    r"^(?:首页|全部|所有|上一页|下一页|<<|>>|<|>|无货|有货|上海无货|北京无货|加入购物车|立即购买|"
    r"关注|对比|收藏|客服|进店|店铺|查看|更多|销量|好评|¥[\d.]+|\d{1,3}|[-—\s]*)$")
# 店铺页里的装修模块标签, 绝不能当商品名
TITLE_BAD = re.compile(r"设计师自定义模块|自定义模块|装修模块|自由模块|图文模块|轮播模块|本店所有商品|"
                       r"按综合|按销量|按新品|按价格|所有宝贝|店铺首页|^首页[-–—\s]|淘！我喜欢", re.I)


def looks_like_title(text: str) -> bool:
    if not text or not (8 <= len(text) <= 130):
        return False
    if TITLE_NOISE.match(text.strip()) or TITLE_BAD.search(text):
        return False
    if not re.search(r"[\u4e00-\u9fa5A-Za-z]", text):
        return False
    return True


def title_from_card(card, anchor) -> str:
    """商品名: 在卡片里挑最长的一条像样的文本(优先带 title 属性的链接)。"""
    candidates = []
    if anchor is not None:
        candidates.append(anchor.attr("title") or "")
        candidates.append(safe_text(anchor))
    if card is not None:
        try:
            for link in (card.eles("tag:a") or [])[:12]:
                candidates.append(link.attr("title") or "")
                candidates.append(safe_text(link))
        except Exception:
            pass
        candidates.extend(safe_text(card).split("\n"))

    best = ""
    for candidate in candidates:
        text = clean_title(candidate or "")
        if looks_like_title(text) and len(text) > len(best):
            best = text
    return best


def wait_products(page, timeout: int = 15) -> int:
    """商品是异步渲染的: 轮询等出现, 再滚到底触发懒加载。"""
    import time

    deadline = time.time() + timeout
    links = finding_links(page)
    while not links and time.time() < deadline:
        time.sleep(0.5)
        links = finding_links(page)

    for _ in range(4):
        try:
            page.scroll.to_bottom()
        except Exception:
            try:
                page.scroll.down(600)
            except Exception:
                break
        time.sleep(0.8)

    return len(finding_links(page))


def products_from_cards(cards: list, page_url: str = "") -> list:
    """只从"商品列表容器"(如淘宝旧版 dl.item)里取, 天然排除推荐位。"""
    products, seen = [], set()
    for card in cards:
        try:
            links = card.eles("css:a[href]") or []
        except Exception:
            continue
        for anchor in links:
            href = absolute(anchor.attr("href") or "", page_url)
            match = JD_ITEM.search(href) or TB_ITEM.search(href)
            if not match:
                continue
            item_id = match.group(1)
            if item_id in seen:
                continue
            seen.add(item_id)
            text = safe_text(card)
            title = title_from_card(card, anchor)   # 取不到先留空, 详情页会补
            sales_raw, sales, caliber = find_sales(text)
            products.append({
                "商品名称": title, "价格": find_price(text),
                "销量原值": sales_raw, "销量": sales, "口径": caliber,
                "商品链接": (f"https://item.jd.com/{item_id}.html" if JD_ITEM.search(href)
                             else f"https://item.taobao.com/item.htm?id={item_id}"),
                "_item_id": item_id,
            })
            break
    return products


def products_from_scope(scope, page_url: str = "") -> list:
    products, seen = [], set()
    for anchor in finding_links(scope):
        if in_recommend_block(anchor):
            continue
        href = absolute(anchor.attr("href") or "", page_url)
        match = JD_ITEM.search(href) or TB_ITEM.search(href)
        if not match:
            continue
        item_id = match.group(1)
        if item_id in seen:
            continue
        seen.add(item_id)
        card = card_element(anchor)
        text = safe_text(card) if card is not None else safe_text(anchor)
        title = title_from_card(card, anchor)   # 取不到先留空, 详情页会补
        sales_raw, sales, caliber = find_sales(text)
        products.append({
            "商品名称": title, "价格": find_price(text),
            "销量原值": sales_raw, "销量": sales, "口径": caliber,
            "商品链接": (f"https://item.jd.com/{item_id}.html" if JD_ITEM.search(href)
                         else f"https://item.taobao.com/item.htm?id={item_id}"),
            "_item_id": item_id,
        })
    return products


def collect_products(page) -> tuple:
    """先按商品列表容器取(排除推荐位), 取不到再全页扫描, 最后遍历子 frame。"""
    page_url = getattr(page, "url", "") or ""
    try:
        cards = page.eles("css:dl.item") or []
    except Exception:
        cards = []
    items = products_from_cards(cards, page_url) if cards else []
    if not items:
        items = products_from_scope(page, page_url)
    if items:
        return items, "main"
    try:
        frames = page.get_frames()
    except Exception:
        frames = []
    for index, frame in enumerate(frames):
        try:
            sub = products_from_scope(frame, page_url)
        except Exception:
            continue
        if sub:
            return sub, f"frame{index}"
    return [], "none"


def page_diagnostics(page) -> str:
    """抓不到东西时打印现场。"""
    try:
        all_links = page.eles("css:a[href]") or []
        item_links = finding_links(page)
        body = ""
        try:
            body = page.ele("tag:body", timeout=1).text or ""
        except Exception:
            pass
        frame_count = 0
        try:
            frame_count = len(page.get_frames())
        except Exception:
            pass
        lines = [f"    URL: {getattr(page, 'url', None)}",
                 f"    标题: {(getattr(page, 'title', '') or '')[:50]}",
                 f"    a[href] 总数={len(all_links)} 其中商品链接={len(item_links)} "
                 f"正文长度={len(body)} 子frame={frame_count}"]
        if item_links:
            sample = [absolute(e.attr("href") or "", page.url) for e in item_links[:3]]
            lines.append(f"    链接样例: {sample}")
        else:
            sample = [absolute(e.attr("href") or "", getattr(page, "url", "")) for e in all_links[:5]]
            lines.append(f"    前几个链接: {sample}")
        if "登录" in body:
            lines.append("    页面出现『登录』字样: 可能需要登录态(用 --attach 接自己已登录的浏览器)")
        return "\n".join(lines)
    except Exception as exc:
        return f"    诊断失败: {type(exc).__name__} {str(exc)[:120]}"


# 详情页里的价格/评价区块, 命中就用, 避免把"推荐位"里别的商品价格当成当前商品
DETAIL_PRICE_SELECTORS = ("css:.summary-price", "css:.p-price", "css:#jd-price", "css:.tm-price",
                          "css:[class*=priceText]", "css:[class*=price--]", "css:[class*=Price]")
DETAIL_SALES_SELECTORS = ("css:#comment-count", "css:.comment-count", "css:[class*=comment-count]",
                          "css:[class*=sellCount]", "css:[class*=sales--]", "css:[class*=SellCount]")


def read_body_text(page, limit: int = 0) -> str:
    try:
        body = page.ele("tag:body", timeout=3)
        text = (body.text or "") if body else ""
    except Exception:
        text = ""
    return text[:limit] if limit else text


DETAIL_TITLE_TAIL = re.compile(
    r"\s*[-–|]\s*(京东|淘宝|天猫|Tmall\.com|tmall\.com|淘宝网|京东商城).*$|【[^】]*行情[^】]*】.*$", re.I)


def read_detail_title(page) -> str:
    """详情页标题最可靠: 页面标题里就是完整商品名。"""
    raw = getattr(page, "title", "") or ""
    if not raw:
        return ""
    text = clean_title(DETAIL_TITLE_TAIL.sub("", raw)).strip()
    return text if looks_like_title(text) else ""


def enrich_from_detail(page, product: dict) -> bool:
    """进商品详情页补价格、销量, 并用详情页标题校正商品名。返回是否有更新。"""
    import time

    try:
        page.get(product["商品链接"], retry=1, timeout=30)
    except Exception:
        return False
    time.sleep(1.2)

    updated = False
    real_title = read_detail_title(page)
    if real_title:
        product["商品名称"] = real_title
        updated = True

    price = ""
    for selector in DETAIL_PRICE_SELECTORS:
        try:
            element = page.ele(selector, timeout=0.5)
        except Exception:
            element = None
        if element is not None:
            price = find_price(safe_text(element))
            if price:
                break

    head = read_body_text(page, 3000)
    if not price:
        price = find_price(head)

    sales_raw, sales, caliber = find_sales(head)
    if not sales:
        for selector in DETAIL_SALES_SELECTORS:
            try:
                element = page.ele(selector, timeout=0.5)
            except Exception:
                element = None
            if element is not None:
                sales_raw, sales, caliber = find_sales(safe_text(element))
                if sales:
                    break
    if not sales:
        sales_raw, sales, caliber = find_sales(read_body_text(page))

    if price:
        product["价格"] = price
        updated = True
    if sales:
        product["销量原值"], product["销量"], product["口径"] = sales_raw, sales, caliber
        updated = True
    return updated


def read_categories(page) -> dict:
    """店铺页上的分类名 -> 分类页地址"""
    categories = {}
    try:
        links = page.eles("css:a[href]") or []
    except Exception:
        links = []
    page_url = getattr(page, "url", "") or ""
    for link in links:
        href = link.attr("href") or ""
        if not CATEGORY_HREF_RE.search(href):
            continue
        name = safe_text(link).replace("\n", "")
        if 2 <= len(name) <= 20 and name not in categories:
            url = normalize_url(absolute(href, page_url))
            # 排除指回本页的链接(店铺 logo 之类); 京东分类是 index-xxx.html?categoryId= 同 base 不同参数, 不能只比 base
            if url and url != normalize_url(page_url):
                categories[name] = url
    return categories


def visit_products(page, url: str, max_pages: int, log, debug: bool = False,
                   expect_name: str = "") -> list:
    collected = []
    page.get(url, retry=1, timeout=30)
    # 页面标题里没有分类名, 说明没真正进到该分类(分类页会回落到全店), 抓出来的东西是错的
    if expect_name:
        title = (getattr(page, "title", "") or "")
        if expect_name not in title:
            log(f"    [!] 页面标题里没有『{expect_name}』, 可能没进到该分类: {title[:36]}")

    for index in range(max_pages):
        wait_products(page, timeout=15)
        batch, where = collect_products(page)
        if not batch and index == 0:
            log("    [!] 未取到商品卡片, 现场诊断:")
            log(page_diagnostics(page))
        elif where != "main" and batch:
            log(f"      商品在 {where} 里")
        collected.extend(batch)

        if index + 1 >= max_pages:
            break
        try:
            nxt = page.ele("text:下一页", timeout=2)
            if not nxt or not nxt.attr("href"):
                break
            nxt.click()
            page.wait.load_start()
            log(f"      翻到第 {index + 2} 页")
        except Exception:
            break
    return collected


ASCII_TOKEN = re.compile(r"[A-Za-z]{2,}|\d+")
CJK_CHAR = re.compile(r"[\u4e00-\u9fa5]")


def _norm(text: str) -> str:
    return re.sub(r"\s+", "", (text or "")).lower()


def _strip_brand(text: str) -> str:
    """去掉开头"第一个"英文品牌词并保留空格边界, 否则 'iQOO Neo系列' 会被整个吃掉。
    'iQOO Neo系列' -> 'Neo系列';  'iQOO旗舰系列' -> '旗舰系列'
    """
    return re.sub(r"^\s*[A-Za-z]{2,}\s*", "", text or "", count=1)


def match_category(target: str, categories: dict):
    """把配置里的目标名对到页面上的分类名。四档, 命中即返回:
    1 完全相等
    2 任一方向子串包含
    3 去掉开头品牌英文后的剩余部分子串匹配
      ("iQOO Neo系列" -> "neo系列" 命中 "Neo系列 新生代性能旗舰")
    4 中文字符重叠率 >= 0.5, 且英文标记(如 Neo/Z)必须对得上
    """
    if target in categories:
        return target

    for name in categories:
        if target in name or name in target:
            return name

    rest = _norm(_strip_brand(target))
    t = _norm(target)
    for name in categories:
        n = _norm(name)
        stripped = _norm(_strip_brand(name))
        # 注意 stripped 可能为空串(纯英文分类名), 空串"在任何字符串里"会误命中
        if rest and (rest in n or (stripped and stripped in t)):
            return name

    ascii_tokens = set(ASCII_TOKEN.findall(t))
    cjk_target = set(CJK_CHAR.findall(target))
    best, best_score = None, 0.0
    for name in categories:
        n = _norm(name)
        if ascii_tokens and not (ascii_tokens & set(ASCII_TOKEN.findall(n))):
            continue
        cjk_name = set(CJK_CHAR.findall(name))
        if not cjk_target or not cjk_name:
            continue
        score = len(cjk_target & cjk_name) / len(cjk_target | cjk_name)
        if score > best_score:
            best, best_score = name, score
    # 纯中文目标不能靠"神影/绝影"这种一字之差蒙混过关; 有英文标记时放宽
    threshold = 0.5 if ascii_tokens else 0.8
    return best if best_score >= threshold else None


def crawl_store_browser(page, store_name: str, store_url: str, targets: list,
                        max_pages: int = 2, log=print, debug: bool = False,
                        detail_limit: int = 0) -> list:
    """按配置里的分类/商品名抓取一家店铺。返回行列表。"""
    platform = "京东" if "jd.com" in store_url else (
        "淘宝" if ("tmall.com" in store_url or "taobao.com" in store_url) else "")
    if not platform:
        log(f"{store_name}: 无法从链接判断平台, 跳过")
        return []

    brand = guess_brand(store_name)
    page.get(store_url, retry=1, timeout=30)
    log(f"{store_name}: 已打开 [{(getattr(page, 'title', '') or '')[:34]}]")

    categories = read_categories(page)
    # 淘宝/天猫: 店铺首页的导航多是 banner 图片热区, 分类侧栏在 category.htm / search.htm 上
    if not categories:
        for candidate in (store_category_url(store_url), store_list_url(store_url)):
            if not candidate:
                continue
            try:
                page.get(candidate, retry=1, timeout=30)
                categories = read_categories(page)
                log(f"    首页没有分类, 转到 {candidate} 重找: 发现 {len(categories)} 个")
                if categories:
                    break
            except Exception as exc:
                log(f"    打开 {candidate} 失败: {type(exc).__name__}")
    log(f"    发现分类 {len(categories)} 个: {'/'.join(list(categories)[:6])}")

    wanted = [str(t) for t in targets if str(t).strip()]
    category_targets, name_keywords = [], []
    for target in wanted:
        hit = match_category(target, categories)
        if hit:
            category_targets.append(hit)
        else:
            name_keywords.append(target)

    if category_targets:
        urls = [categories[n] for n in category_targets if n in categories]
        log(f"    按分类抓取: {'/'.join(category_targets)}")
        if len(set(urls)) < len(urls):
            log("    [!] 有多个分类指向同一个 URL, 说明分类链接解析有问题, 结果可能混在一起")
    elif categories:
        urls = list(categories.values())
        log(f"    未指定分类, 抓取全部分类 {len(urls)} 个")
    else:
        fallback = store_list_url(store_url) or store_url
        urls = [fallback]
        if name_keywords:
            log(f"    未发现分类链接, 改为抓商品列表页 {fallback}; 分类名不再作商品名过滤: {name_keywords}")
            name_keywords = []

    rows, seen = [], set()
    for index, url in enumerate(urls):
        expect = category_targets[index] if index < len(category_targets) else ""
        batch = visit_products(page, url, max_pages, log, debug, expect)
        kept = 0
        for product in batch:
            # 只有"没匹配到任何分类"时, 配置项才当商品名关键词用
            if name_keywords and not category_targets:
                if not any(k in product["商品名称"] for k in name_keywords):
                    continue
            if product["_item_id"] in seen:
                continue
            seen.add(product["_item_id"])
            rows.append({
                "平台": platform, "店铺": store_name,
                "品牌": brand or guess_brand(product["商品名称"]),
                "型号": find_model(product["商品名称"]),
                "商品名称": product["商品名称"], "价格": product["价格"],
                "销量原值": product["销量原值"], "销量": product["销量"],
                "口径": product["口径"], "商品链接": product["商品链接"],
                "抓取日期": date.today().isoformat(),
                "_item_id": product["_item_id"], "_kind": "browser",
            })
            kept += 1
        log(f"    {url[-38:]}: 抓到 {len(batch)} 个, 采用 {kept} 个")

    if name_keywords and not category_targets:
        log(f"    商品名关键词过滤: {name_keywords}")

    # 分类页没有价格和销量, 抓完再逐个进详情页补。放在翻页之后, 避免打断分页。
    if detail_limit and rows:
        target = rows[:detail_limit]
        log(f"    进入详情页补价格/销量: {len(target)} 个")
        for index, row in enumerate(target, start=1):
            if enrich_from_detail(page, row):
                # 商品名被详情页校正过, 型号和品牌要跟着重算
                row["型号"] = find_model(row["商品名称"])
                if not row["品牌"]:
                    row["品牌"] = guess_brand(row["商品名称"])
                log(f"      [{index}/{len(target)}] ¥{row['价格'] or '-':<9} "
                    f"{row['销量原值'] or '无销量':>6} {row['商品名称'][:26]}")

    # 详情页也没能补出名字的, 才是真的废数据
    unnamed = sum(1 for row in rows if not row["商品名称"])
    if unnamed:
        log(f"    [!] 有 {unnamed} 个商品拿不到名称(卡片和详情页都没有), 已丢弃")
        rows = [row for row in rows if row["商品名称"]]

    if not rows:
        log("    [!] 这家店一个商品都没抓到, 现场诊断:")
        log(page_diagnostics(page))
    return rows


def load_config(path: Path) -> tuple:
    """返回 (文件任务, 在线抓取任务)"""
    config = json.loads(path.read_text(encoding="utf-8"))
    tasks, live = [], []
    for platform_key, platform_conf in (config.get("Platform") or {}).items():
        platform = PLATFORM_ALIAS.get(str(platform_key).upper(), str(platform_key))
        base = Path(platform_conf.get("HtmlPath") or path.parent)
        for section, files_by_brand in platform_conf.items():
            if not isinstance(files_by_brand, dict) or section == "HtmlPath":
                continue
            kind = "detail" if "detail" in section.lower() or "详情" in section else "list"
            for brand_key, files in files_by_brand.items():
                brand = BRAND_ALIAS.get(str(brand_key).upper(), str(brand_key))
                if isinstance(files, str):
                    files = [files]
                for name in files:
                    if str(name).lower().endswith((".html", ".htm")):
                        tasks.append((platform, brand, base / name, kind))

    # 顶层 Details: {"店铺名": {"url": "店铺链接", "商品类别": [...], "具体商品": [...]}}
    # 兼容旧写法 {"店铺名": {"店铺链接": [分类...]}}
    for key in ("Details", "Detail", "详情"):
        section = config.get(key)
        if not isinstance(section, dict):
            continue
        for store_name, store_conf in section.items():
            if not isinstance(store_conf, dict):
                store_conf = { "url": store_conf } if isinstance(store_conf, str) else {}
            store_url, targets = "", []
            for field, value in store_conf.items():
                if isinstance(field, str) and field.startswith("http"):
                    store_url = store_url or field
                if isinstance(value, str) and value.startswith("http"):
                    store_url = store_url or value
                elif isinstance(value, list):
                    targets.extend(str(v).strip() for v in value if str(v).strip())
            if store_url:
                live.append((store_name, store_url, targets))
    return tasks, live
    return tasks, live


def write_xlsx(rows: list, out: Path) -> Path:
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter

    wb = Workbook()
    ws = wb.active
    ws.title = "月销统计"
    ws.append(COLUMNS)
    head_font = Font(bold=True, color="FFFFFF")
    head_fill = PatternFill("solid", fgColor="4B3F72")
    for cell in ws[1]:
        cell.font = head_font
        cell.fill = head_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")

    for row in rows:
        values = []
        for col in COLUMNS:
            value = row.get(col, "")
            if col in ("价格", "销量") and value not in ("", None):
                try:
                    value = float(value) if col == "价格" else int(value)
                except (ValueError, TypeError):
                    pass
            # 没有销量/口径就写 Null, 不要留空让人误以为是 0
            if col in ("销量", "口径") and value in ("", None):
                value = "Null"
            values.append(value)
        ws.append(values)

    widths = {"平台": 10, "店铺": 30, "品牌": 10, "型号": 16, "商品名称": 48, "价格": 10,
              "销量原值": 12, "销量": 10, "口径": 10, "商品链接": 50, "抓取日期": 12}
    for idx, col in enumerate(COLUMNS, start=1):
        ws.column_dimensions[get_column_letter(idx)].width = widths.get(col, 14)
    for cell in ws[get_column_letter(COLUMNS.index("价格") + 1)][1:]:
        cell.number_format = "0.00"
    for cell in ws[get_column_letter(COLUMNS.index("销量") + 1)][1:]:
        cell.number_format = "#,##0"

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    wb.save(out)
    return out


def selftest(live_tasks: list) -> int:
    print("1) 依赖检查")
    try:
        import DrissionPage
        print("   DrissionPage:", getattr(DrissionPage, "__version__", "ok"))
    except Exception as exc:
        print("   导入失败:", type(exc).__name__, exc)
        print("   先执行: pip install DrissionPage -i https://pypi.mirrors.ustc.edu.cn/simple")
        return 1

    print("2) 配置检查")
    print("   Details 店铺数:", len(live_tasks))
    for name, url, targets in live_tasks:
        print(f"     {name} -> {url}  目标={targets or '全部'}")

    if not live_tasks:
        print("   Info.json 里没有 Details 段, 无法自检")
        return 1

    print("3) 浏览器启动与页面读取")
    page = None
    try:
        page = launch_browser()
        page.get(live_tasks[0][1], retry=1, timeout=30)
        print("   页面标题:", page.title[:50])
        print("   发现分类:", list(read_categories(page))[:8])
        print("\n自检通过, 可以执行: python product_sales.py")
        return 0
    except Exception as exc:
        print("   失败:", type(exc).__name__, str(exc)[:300])
        print("   常见原因: 浏览器已在用同一用户目录 / 无界面环境需加 --headless")
        print("   也可以自己用 --remote-debugging-port=9222 启动浏览器后加 --attach 127.0.0.1:9222")
        return 1
    finally:
        if page is not None:
            try:
                page.quit()
            except Exception:
                pass


def main() -> int:
    parser = argparse.ArgumentParser(description="按 Info.json 抓取京东/淘宝店铺商品销售数据")
    parser.add_argument("--config", default="crawler_config.json", help="配置文件路径")
    parser.add_argument("--out", default="商品销售统计.xlsx", help="输出 xlsx 文件名")
    parser.add_argument("--mode", choices=("browser", "html"), default="browser",
                        help="browser=浏览器实时抓取(默认), html=只解析本地已保存的 HTML")
    parser.add_argument("--attach", default="", help="接入已运行的浏览器, 例如 127.0.0.1:9222")
    parser.add_argument("--headless", action="store_true", help="无头模式(看不到窗口)")
    parser.add_argument("--max-pages", type=int, default=2, help="每个分类最多翻几页")
    parser.add_argument("--selftest", action="store_true", help="检查依赖/配置/浏览器能否工作")
    parser.add_argument("--probe", action="store_true", help="只打印结果, 不写 xlsx")
    parser.add_argument("--debug", action="store_true", help="抓不到商品时打印页面现场诊断")
    parser.add_argument("--detail-limit", type=int, default=30,
                        help="最多进多少个商品详情页补价格/销量, 0=不进详情页")
    parser.add_argument("--no-detail", action="store_true", help="完全不进详情页(快, 但只有商品名和链接)")
    args = parser.parse_args()

    config_path = Path(args.config)
    if not config_path.exists():
        raise SystemExit(f"配置文件不存在: {config_path.resolve()}")
    tasks, live_tasks = load_config(config_path)

    if args.selftest:
        return selftest(live_tasks)

    collected = []
    if args.mode == "html":
        if not tasks:
            raise SystemExit("Platform 段里没有可解析的 HTML 文件")
        report = []
        for platform, brand, html_path, kind in tasks:
            if not html_path.exists():
                report.append((platform, brand, kind, html_path.name, -1, -1, "文件不存在"))
                continue
            page_rows = parse_page(html_path, platform, brand, kind)
            with_sales = sum(1 for row in page_rows if row["销量"])
            note = "" if page_rows else (
                "详情页无可解析内容" if kind == "detail" else "无独立商品数据(图片热区或未渲染)")
            report.append((platform, brand, kind, html_path.name, len(page_rows), with_sales, note))
            collected.extend(page_rows)

        print(f"{'平台':<5}{'品牌':<6}{'类型':<8}{'商品':<6}{'带销量':<7}文件")
        for platform, brand, kind, name, total, with_sales, note in report:
            flag = f"  ← {note}" if note else ""
            print(f"{platform:<5}{brand:<6}{kind:<8}{total:<6}{with_sales:<7}{name[:36]}{flag}")
    else:
        if not live_tasks:
            raise SystemExit("配置里没有 Details 段, 没有可抓取的店铺")
        print("--- 浏览器抓取 ---")
        page = None
        try:
            page = launch_browser(args.attach, args.headless)
        except Exception as exc:
            raise SystemExit(
                f"浏览器启动失败: {type(exc).__name__} {str(exc)[:200]}\n"
                "可先跑 --selftest 定位, 或 --attach 127.0.0.1:9222 接入已开的浏览器")
        try:
            for store_name, store_url, targets in live_tasks:
                limit = 0 if args.no_detail else args.detail_limit
                collected.extend(crawl_store_browser(page, store_name, store_url,
                                                     targets, args.max_pages, print, args.debug, limit))
        finally:
            try:
                page.quit()
            except Exception:
                pass

    rows = merge_rows(collected)
    print(f"\n合并后共 {len(rows)} 条")
    if not rows:
        print("没有抓到任何商品: 检查配置里的分类名是否和店铺页一致")
        return 1

    rows.sort(key=lambda r: (0 if r["销量"] else 1, -(r["销量"] or 0), 0 if r["价格"] else 1))
    missing = sum(1 for r in rows if not r["销量"])
    print(f"共 {len(rows)} 条, 带销量 {len(rows) - missing} 条, 带价格 {sum(1 for r in rows if r['价格'])} 条")

    if args.probe:
        for row in rows[:20]:
            print(f"  [{row['平台']}/{row['品牌'] or '-'}] {row['型号'] or '-':<10} "
                  f"{row['销量原值'] or '无销量':>8} ¥{row['价格'] or '-':>8} "
                  f"{row['商品名称'][:34]}")
        return 0

    out = write_xlsx(rows, Path(args.out))
    summary = Counter((r["平台"], r["店铺"]) for r in rows)
    print("\n按平台/店铺:")
    for (platform, store), count in sorted(summary.items()):
        print(f"  {platform} {store}: {count} 条")
    print(f"\n已导出: {out.resolve()}")
    return 0


if __name__ == "__main__":
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    raise SystemExit(main())
