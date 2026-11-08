"""解析 .xlsx 表格, 按表头自动识别字段类型, 生成图表可视化看板。

用法:
    python xlsx_dashboard.py --xlsx 笔记本月销统计.xlsx
    python xlsx_dashboard.py --xlsx 任意表.xlsx --out 看板.html --top 15
"""

from __future__ import annotations

import argparse
import base64
import io
import math
import re
import sys
from datetime import date, datetime
from pathlib import Path

from openpyxl import load_workbook

PALETTE = ["#4B3F72", "#FF5000", "#00A6A6", "#F0A202", "#C41E3A", "#2E86AB", "#8A4FFF", "#2FBF71"]
NUMERIC_HINTS = re.compile(r"数|量|额|价|销量|人数|占比|比例|评分|size|count|amount|price|sales|qty", re.I)
# 表格里的占位符按"缺失"处理: 不参与类型推断, 也不参与求和
NULL_TOKENS = {"null", "none", "na", "n/a", "-", "--", "—", "无", "未知"}
# 这些列是主键/标识, 长得像数字也不能当度量参与求和与散点
ID_RE = re.compile(r"订单号|单号|编号|编码|条码|流水号|^id$|_id$|sku|手机|电话", re.I)
DATE_RE = re.compile(r"^\d{4}[-/]\d{1,2}[-/]\d{1,2}")


def read_sheet(path: Path) -> tuple:
    wb = load_workbook(path, data_only=True, read_only=True)
    ws = wb[wb.sheetnames[0]]
    rows = [list(r) for r in ws.iter_rows(values_only=True)]
    wb.close()
    rows = [r for r in rows if any(v not in (None, "") for v in r)]
    if not rows:
        raise SystemExit(f"{path.name} 是空表")
    header = [str(c).strip() if c not in (None, "") else f"列{i + 1}" for i, c in enumerate(rows[0])]
    data = [r + [None] * (len(header) - len(r)) for r in rows[1:]]
    return header, data


def to_number(value):
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    if value is None:
        return None
    text = str(value).strip().replace(",", "").replace("¥", "").replace("￥", "")
    if not text or text.lower() in NULL_TOKENS:
        return None
    multiplier = 1
    if text.endswith("万"):
        text, multiplier = text[:-1], 10000
    try:
        return float(text) * multiplier
    except ValueError:
        return None


def infer_types(header: list, data: list) -> list:
    types = []
    for i, name in enumerate(header):
        values = [row[i] for row in data
                  if row[i] not in (None, "") and str(row[i]).strip().lower() not in NULL_TOKENS]
        if not values:
            types.append("empty")
            continue
        numeric = sum(1 for v in values if to_number(v) is not None) / len(values)
        dated = sum(1 for v in values if DATE_RE.match(str(v)) or isinstance(v, (date, datetime))) / len(values)
        numbers = [to_number(v) for v in values]
        looks_like_id = (
            all(n is not None and n == int(n) and n > 1e8 for n in numbers)
            and len({n for n in numbers}) / len(numbers) > 0.9
        )
        if dated >= 0.7:
            types.append("date")
        elif ID_RE.search(name) or looks_like_id:
            types.append("category")
        elif numeric >= 0.7 or (numeric >= 0.5 and NUMERIC_HINTS.search(name)):
            types.append("number")
        else:
            types.append("category")
    return types


def pick_fields(header: list, types: list, data: list) -> dict:
    cats = [h for h, t in zip(header, types) if t == "category"]
    nums = [h for h, t in zip(header, types) if t == "number"]
    dates = [h for h, t in zip(header, types) if t == "date"]

    # 分类字段优先选不同取值少、看起来像维度的(平台/店铺/品牌)
    def cat_score(name):
        idx = header.index(name)
        distinct = len({str(r[idx]) for r in data})
        weight = 0
        if re.search(r"平台|店铺|品牌|类目|渠道|型号|分类", name):
            weight -= 100
        if re.search(r"名称|标题|链接|url", name, re.I):
            weight += 1000
        return weight + distinct

    # 只有一个取值的分类列没有信息量, 不参与分组
    informative = [c for c in cats if len({str(r[header.index(c)]) for r in data}) >= 2]
    cats = informative or cats
    cats.sort(key=cat_score)

    def num_score(name):
        derived = 1 if re.search(r"原值|原始|链接|url|日期|时间", name, re.I) else 0
        core = 0 if re.search(r"销量|数量|count|sales|qty", name, re.I) else 1
        return (derived, core)

    nums.sort(key=num_score)
    return {"cats": cats, "nums": nums, "dates": dates}


def pick_label_field(cats: list, header: list, data: list) -> str:
    """Top N 图的标签要能区分每一行, 所以取区分度最高的分类字段, 排除链接类。"""
    candidates = [c for c in cats if not re.search(r"链接|url|地址|图片", c, re.I)]
    if not candidates:
        return cats[0] if cats else header[0]

    def distinct(name):
        i = header.index(name)
        return len({str(row[i]) for row in data})

    return max(candidates, key=distinct)


def aggregate(cat_name: str, num_name: str, header: list, data: list, top: int = 12) -> list:
    ci, ni = header.index(cat_name), header.index(num_name)
    buckets = {}
    for row in data:
        key = str(row[ci]) if row[ci] not in (None, "") else "(空)"
        value = to_number(row[ni]) or 0.0
        buckets[key] = buckets.get(key, 0.0) + value
    pairs = sorted(buckets.items(), key=lambda kv: kv[1], reverse=True)[:top]
    return pairs


def ranking(num_name: str, label_name: str, header: list, data: list, top: int = 10) -> list:
    ni, li = header.index(num_name), header.index(label_name)
    items = []
    for row in data:
        value = to_number(row[ni])
        if value is None:
            continue
        label = str(row[li]) if row[li] not in (None, "") else "(空)"
        items.append((label[:26], value))
    items.sort(key=lambda kv: kv[1], reverse=True)
    return items[:top]


def fmt(value: float) -> str:
    if value >= 100000000:
        return f"{value / 100000000:.2f}亿"
    if value >= 10000:
        return f"{value / 10000:.1f}万"
    if value == int(value):
        return f"{int(value):,}"
    return f"{value:,.2f}"


def esc(text) -> str:
    return (str(text).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
            .replace('"', "&quot;"))


def svg_barh(pairs: list, width: int = 640, row_h: int = 30) -> str:
    if not pairs:
        return "<p>无可绘制的数据</p>"
    top = max(v for _, v in pairs) or 1
    height = len(pairs) * row_h + 24
    label_w = 150
    bar_w = width - label_w - 90
    parts = [f'<svg viewBox="0 0 {width} {height}" width="100%" height="{height}" font-family="Microsoft YaHei, sans-serif" font-size="13">']
    for i, (name, value) in enumerate(pairs):
        y = i * row_h + 12
        w = max(2, bar_w * value / top)
        color = PALETTE[i % len(PALETTE)]
        parts.append(f'<text x="0" y="{y + 14}" fill="#333">{esc(name[:18])}</text>')
        parts.append(f'<rect x="{label_w}" y="{y + 2}" width="{w:.1f}" height="{row_h - 10}" rx="3" fill="{color}"/>')
        parts.append(f'<text x="{label_w + w + 8:.1f}" y="{y + 14}" fill="#666">{fmt(value)}</text>')
    parts.append("</svg>")
    return "".join(parts)


def svg_pie(pairs: list, size: int = 380) -> str:
    total = sum(v for _, v in pairs)
    if not pairs or total <= 0:
        return "<p>无可绘制的数据</p>"
    cx = cy = size / 2
    r = size / 2 - 12
    angle = -math.pi / 2
    parts = [f'<svg viewBox="0 0 {size} {size}" width="{size}" height="{size}" font-family="Microsoft YaHei, sans-serif" font-size="12">']
    for i, (name, value) in enumerate(pairs):
        sweep = 2 * math.pi * value / total
        x1, y1 = cx + r * math.cos(angle), cy + r * math.sin(angle)
        angle += sweep
        x2, y2 = cx + r * math.cos(angle), cy + r * math.sin(angle)
        large = 1 if sweep > math.pi else 0
        color = PALETTE[i % len(PALETTE)]
        if len(pairs) == 1:
            parts.append(f'<circle cx="{cx}" cy="{cy}" r="{r}" fill="{color}"/>')
        else:
            parts.append(f'<path d="M{cx},{cy} L{x1:.2f},{y1:.2f} A{r},{r} 0 {large} 1 {x2:.2f},{y2:.2f} Z" fill="{color}" stroke="#fff" stroke-width="1.5"/>')
    legend = "".join(
        f'<div class="lg"><span style="background:{PALETTE[i % len(PALETTE)]}"></span>{esc(n[:16])} {fmt(v)} ({v / total:.0%})</div>'
        for i, (n, v) in enumerate(pairs))
    parts.append("</svg>")
    return "".join(parts) + f'<div class="legend">{legend}</div>'


def svg_scatter(points: list, x_name: str, y_name: str, width: int = 640, height: int = 340) -> str:
    if not points:
        return "<p>无可绘制的数据</p>"
    xs = [p[1] for p in points]
    ys = [p[2] for p in points]
    xmax = max(xs) or 1
    ymax = max(ys) or 1
    pad = 46
    parts = [f'<svg viewBox="0 0 {width} {height}" width="100%" height="{height}" font-family="Microsoft YaHei, sans-serif" font-size="12">']
    parts.append(f'<line x1="{pad}" y1="{height - pad}" x2="{width - 16}" y2="{height - pad}" stroke="#ccc"/>')
    parts.append(f'<line x1="{pad}" y1="16" x2="{pad}" y2="{height - pad}" stroke="#ccc"/>')
    for i, (name, x, y) in enumerate(points):
        px = pad + (width - pad - 24) * x / xmax
        py = (height - pad) - (height - pad - 24) * y / ymax
        parts.append(f'<circle cx="{px:.1f}" cy="{py:.1f}" r="5" fill="{PALETTE[i % len(PALETTE)]}" opacity="0.8"><title>{esc(name[:20])}: {fmt(x)} / {fmt(y)}</title></circle>')
    parts.append(f'<text x="{pad}" y="{height - pad + 22}" fill="#666">{esc(x_name)} →</text>')
    parts.append(f'<text x="{pad + 4}" y="12" fill="#666">↑ {esc(y_name)}</text>')
    parts.append("</svg>")
    return "".join(parts)


def try_matplotlib(charts: list, out_dir: Path) -> dict:
    try:
        import os
        os.environ.setdefault("MPLCONFIGDIR", str(Path.cwd() / ".mplconfig"))
        import matplotlib
        matplotlib.use("Agg")
        import matplotlib.pyplot as plt
        matplotlib.rcParams["font.sans-serif"] = ["Microsoft YaHei", "SimHei", "SimSun", "DengXian"]
        matplotlib.rcParams["axes.unicode_minus"] = False
    except Exception:
        return {}
    made = {}
    for kind, spec, title in charts:
        try:
            fig, ax = plt.subplots(figsize=(7, 3.6), dpi=110)
            if kind == "bar":
                names = [n for n, _ in spec][::-1]
                values = [v for _, v in spec][::-1]
                ax.barh(names, values, color=PALETTE[0])
                ax.bar_label(ax.containers[0], labels=[fmt(v) for v in values], fontsize=8, padding=4)
                ax.set_xlim(0, (max(values) or 1) * 1.18)
            elif kind == "pie":
                ax.pie([v for _, v in spec], labels=[n[:10] for n, _ in spec], autopct="%1.0f%%",
                       colors=PALETTE[:len(spec)], textprops={"fontsize": 8})
            elif kind == "scatter":
                ax.scatter([p[1] for p in spec], [p[2] for p in spec], color=PALETTE[1], alpha=0.8)
            buf = io.BytesIO()
            fig.tight_layout()
            fig.savefig(buf, format="png")
            plt.close(fig)
            made[title] = base64.b64encode(buf.getvalue()).decode()
        except Exception as exc:
            print(f"  matplotlib 出图失败({title}): {exc}")
    return made


def build_html(header, types, data, fields, charts, kpis, pngs, top) -> str:
    blocks = []
    for title, svg in charts:
        png = pngs.get(title)
        body = f'<img src="data:image/png;base64,{png}" style="max-width:100%"/>' if png else svg
        tag = '<span class="tag">matplotlib PNG</span>' if png else '<span class="tag alt">内联 SVG</span>'
        blocks.append(f'<section class="card"><h2>{esc(title)} {tag}</h2>{body}</section>')

    head = "".join(f"<th>{esc(h)}<i>{t}</i></th>" for h, t in zip(header, types))
    body_rows = []
    for row in data[:top * 20]:
        cells = "".join(f"<td>{esc('' if c is None else c)}</td>" for c in row)
        body_rows.append(f"<tr>{cells}</tr>")

    kpi_html = "".join(f'<div class="kpi"><b>{esc(v)}</b><span>{esc(k)}</span></div>' for k, v in kpis)
    return f"""<!doctype html><html lang="zh-CN"><head><meta charset="utf-8">
<title>数据看板</title>
<style>
:root{{--bg:#f5f6fa;--card:#fff;--ink:#242436;--mute:#8a8aa0;--p:#4B3F72}}
*{{box-sizing:border-box}}
body{{margin:0;padding:24px;background:var(--bg);color:var(--ink);font-family:"Microsoft YaHei",system-ui,sans-serif}}
h1{{font-size:20px;margin:0 0 4px}}
.sub{{color:var(--mute);font-size:13px;margin-bottom:18px}}
.kpis{{display:flex;gap:14px;flex-wrap:wrap;margin-bottom:18px}}
.kpi{{background:var(--card);border-radius:10px;padding:14px 18px;box-shadow:0 1px 3px rgba(0,0,0,.06);min-width:130px}}
.kpi b{{display:block;font-size:22px;color:var(--p)}}
.kpi span{{font-size:12px;color:var(--mute)}}
.grid{{display:grid;grid-template-columns:repeat(auto-fit,minmax(430px,1fr));gap:16px}}
.card{{background:var(--card);border-radius:12px;padding:16px 18px;box-shadow:0 1px 3px rgba(0,0,0,.06)}}
.card h2{{font-size:14px;margin:0 0 12px;font-weight:600}}
.tag{{font-size:10px;background:#eef;color:#55c;padding:2px 6px;border-radius:6px;font-weight:400;margin-left:6px}}
.tag.alt{{background:#ffe9e0;color:#c40}}
.legend{{display:flex;flex-wrap:wrap;gap:8px;margin-top:8px;font-size:12px;color:#555}}
.legend span{{display:inline-block;width:10px;height:10px;border-radius:2px;margin-right:5px}}
table{{width:100%;border-collapse:collapse;font-size:12px}}
th{{position:sticky;top:0;background:#f0f0f7;text-align:left;padding:8px;cursor:pointer;white-space:nowrap}}
th i{{display:block;font-style:normal;font-size:9px;color:var(--mute);font-weight:400}}
td{{padding:7px 8px;border-top:1px solid #eee;max-width:340px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}}
tbody tr:hover{{background:#fafafd}}
.tablewrap{{max-height:520px;overflow:auto}}
</style></head><body>
<h1>数据看板</h1>
<div class="sub">数据源 <b>{esc(Path.cwd().name)}</b> · 生成于 {datetime.now():%Y-%m-%d %H:%M} · {len(data)} 行 / {len(header)} 字段</div>
<div class="kpis">{kpi_html}</div>
<div class="grid">{"".join(blocks)}</div>
<section class="card" style="margin-top:16px">
<h2>明细数据（点表头可排序）</h2>
<div class="tablewrap"><table id="t"><thead><tr>{head}</tr></thead><tbody>{"".join(body_rows)}</tbody></table></div>
</section>
<script>
document.querySelectorAll('#t th').forEach((th,i)=>{{
  let asc=true;
  th.onclick=()=>{{
    const tb=th.closest('table').tBodies[0];
    const rows=[...tb.rows].sort((a,b)=>{{
      const x=a.cells[i].innerText,y=b.cells[i].innerText;
      const nx=parseFloat(x.replace(/[^\\d.-]/g,'')),ny=parseFloat(y.replace(/[^\\d.-]/g,''));
      const v=(!isNaN(nx)&&!isNaN(ny))?(nx-ny):x.localeCompare(y,'zh');
      return asc?v:-v;
    }});
    asc=!asc; rows.forEach(r=>tb.appendChild(r));
  }};
}});
</script>
</body></html>"""


def main() -> int:
    parser = argparse.ArgumentParser(description="解析 xlsx 并生成图表看板")
    parser.add_argument("--xlsx", required=True, help="输入 xlsx 文件")
    parser.add_argument("--out", default="dashboard.html", help="输出 HTML 路径")
    parser.add_argument("--top", type=int, default=12, help="每张图最多展示多少分类/TopN")
    args = parser.parse_args()

    path = Path(args.xlsx)
    if not path.exists():
        raise SystemExit(f"文件不存在: {path}")
    header, data = read_sheet(path)
    types = infer_types(header, data)
    fields = pick_fields(header, types, data)
    print(f"字段识别: " + ", ".join(f"{h}={t}" for h, t in zip(header, types)))

    cats, nums, dates = fields["cats"], fields["nums"], fields["dates"]
    charts, specs, chart_specs = [], [], []

    if cats and nums:
        pairs = aggregate(cats[0], nums[0], header, data, args.top)
        specs.append((f"{cats[0]} 的 {nums[0]} 汇总", svg_barh(pairs)))
        chart_specs.append(("bar", pairs, f"{cats[0]} 的 {nums[0]} 汇总"))

    if nums:
        label = pick_label_field(cats, header, data) if cats else header[0]
        items = ranking(nums[0], label, header, data, min(args.top, 10))
        specs.append((f"{label} 的 {nums[0]} Top {len(items)}", svg_barh(items)))
        chart_specs.append(("bar", items, f"{label} 的 {nums[0]} Top {len(items)}"))

    if cats and nums and len(cats) > 1:
        pairs = aggregate(cats[1], nums[0], header, data, min(args.top, 8))
        specs.append((f"{nums[0]} 按 {cats[1]} 占比", svg_pie(pairs)))
        chart_specs.append(("pie", pairs, f"{nums[0]} 按 {cats[1]} 占比"))

    if len(nums) >= 2:
        x_name, y_name = nums[0], nums[1]
        xi, yi = header.index(x_name), header.index(y_name)
        li = header.index(cats[0]) if cats else 0
        points = []
        for row in data:
            x, y = to_number(row[xi]), to_number(row[yi])
            if x is None or y is None:
                continue
            points.append((str(row[li] or ""), x, y))
        specs.append((f"{x_name} vs {y_name} 分布", svg_scatter(points, x_name, y_name)))
        chart_specs.append(("scatter", points, f"{x_name} vs {y_name} 分布"))

    kpis = [("数据行数", len(data)), ("字段数", len(header)),
            ("分类字段", len(cats)), ("数值字段", len(nums))]
    for name in nums[:2]:
        idx = header.index(name)
        total = sum(to_number(r[idx]) or 0 for r in data)
        kpis.append((f"{name}合计", fmt(total)))

    print("生成图表: " + ", ".join(t for t, _ in specs))
    pngs = try_matplotlib(chart_specs, path.parent)
    print("matplotlib:" + ("已出图 " + ", ".join(pngs) if pngs else "未安装或不可用, 使用内联 SVG"))

    html = build_html(header, types, data, fields, specs, kpis, pngs, args.top)
    out = Path(args.out)
    out.write_text(html, encoding="utf-8")
    print(f"看板已生成: {out.resolve()}")
    return 0


if __name__ == "__main__":
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    raise SystemExit(main())
