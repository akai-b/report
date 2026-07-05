from __future__ import annotations

import json
import math
import posixpath
import statistics
import xml.etree.ElementTree as ET
from collections import defaultdict
from datetime import datetime, timedelta
from pathlib import Path
from zipfile import ZipFile

BASE_DIR = Path("D:/Downloads/test")
XLSX_PATH = BASE_DIR / "26年销售数据.xlsx"
HTML_PATH = BASE_DIR / "销售数据透视看板.html"

NS = {
    "a": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}


def column_index(cell_ref: str) -> int:
    letters = ""
    for ch in cell_ref:
        if ch.isalpha():
            letters += ch
        else:
            break
    result = 0
    for ch in letters:
        result = result * 26 + (ord(ch.upper()) - ord("A") + 1)
    return result - 1


def read_shared_strings(z: ZipFile) -> list[str]:
    if "xl/sharedStrings.xml" not in z.namelist():
        return []
    root = ET.fromstring(z.read("xl/sharedStrings.xml"))
    values: list[str] = []
    for si in root.findall("a:si", NS):
        texts = []
        for t in si.iter("{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t"):
            texts.append(t.text or "")
        values.append("".join(texts))
    return values


def read_first_sheet_rows(path: Path) -> tuple[str, list[list[object]]]:
    with ZipFile(path) as z:
        shared = read_shared_strings(z)
        workbook = ET.fromstring(z.read("xl/workbook.xml"))
        rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
        relmap = {rel.attrib["Id"]: rel.attrib["Target"] for rel in rels}
        sheet = workbook.find("a:sheets/a:sheet", NS)
        if sheet is None:
            raise RuntimeError("Excel 中未找到工作表")
        title = sheet.attrib.get("name", "Sheet1")
        rid = sheet.attrib["{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"]
        target = relmap[rid]
        if target.startswith("/"):
            sheet_path = target.lstrip("/")
        else:
            sheet_path = posixpath.normpath(posixpath.join("xl", target))
        root = ET.fromstring(z.read(sheet_path))

        rows: list[list[object]] = []
        for row in root.findall("a:sheetData/a:row", NS):
            values: list[object] = []
            for cell in row.findall("a:c", NS):
                idx = column_index(cell.attrib.get("r", "A1"))
                while len(values) <= idx:
                    values.append("")
                cell_type = cell.attrib.get("t")
                value_node = cell.find("a:v", NS)
                value: object = ""
                if cell_type == "inlineStr":
                    text_node = cell.find("a:is/a:t", NS)
                    value = text_node.text if text_node is not None else ""
                elif value_node is not None:
                    raw = value_node.text or ""
                    if cell_type == "s":
                        value = shared[int(raw)] if raw else ""
                    elif cell_type == "b":
                        value = raw == "1"
                    else:
                        try:
                            num = float(raw)
                            value = int(num) if num.is_integer() else num
                        except ValueError:
                            value = raw
                values[idx] = value
            if any(v != "" for v in values):
                rows.append(values)
        return title, rows


def to_float(value: object) -> float:
    if value is None or value == "":
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).replace(",", "").replace("¥", "").replace("￥", "").strip()
    if not text:
        return 0.0
    try:
        return float(text)
    except ValueError:
        return 0.0


def excel_date_to_datetime(value: object) -> datetime | None:
    if value is None or value == "":
        return None
    if isinstance(value, datetime):
        return value
    if isinstance(value, (int, float)):
        return datetime(1899, 12, 30) + timedelta(days=float(value))
    text = str(value).strip()
    for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%Y.%m.%d", "%Y-%m-%d %H:%M:%S", "%Y/%m/%d %H:%M:%S"):
        try:
            return datetime.strptime(text, fmt)
        except ValueError:
            continue
    return None


def round2(value: float) -> float:
    return round(float(value) + 1e-9, 2)


def build_table_from_group(group: dict[str, dict[str, float]], key_name: str) -> list[dict[str, object]]:
    rows = []
    for key, item in group.items():
        qty = item["quantity"]
        amount = item["amount"]
        orders = int(item["orders"])
        rows.append({
            key_name: key,
            "销售额": round2(amount),
            "销售数量": round2(qty),
            "订单数": orders,
            "平均单价": round2(amount / qty) if qty else 0,
            "客单价": round2(amount / orders) if orders else 0,
        })
    rows.sort(key=lambda x: x["销售额"], reverse=True)
    return rows


def aggregate(records: list[dict[str, object]]) -> dict[str, object]:
    total_amount = sum(float(r["amount"]) for r in records)
    total_qty = sum(float(r["quantity"]) for r in records)
    valid_price_records = [float(r["price"]) for r in records if float(r["price"]) > 0]

    product_group: dict[str, dict[str, float]] = defaultdict(lambda: {"amount": 0.0, "quantity": 0.0, "orders": 0.0})
    seller_group: dict[str, dict[str, float]] = defaultdict(lambda: {"amount": 0.0, "quantity": 0.0, "orders": 0.0})
    month_group: dict[str, dict[str, float]] = defaultdict(lambda: {"amount": 0.0, "quantity": 0.0, "orders": 0.0})
    cross_group: dict[tuple[str, str], dict[str, float]] = defaultdict(lambda: {"amount": 0.0, "quantity": 0.0, "orders": 0.0})
    price_by_product: dict[str, list[float]] = defaultdict(list)

    for r in records:
        product = str(r["product"])
        seller = str(r["seller"])
        month = str(r["month"])
        qty = float(r["quantity"])
        price = float(r["price"])
        amount = float(r["amount"])
        for group, key in ((product_group, product), (seller_group, seller), (month_group, month)):
            group[key]["amount"] += amount
            group[key]["quantity"] += qty
            group[key]["orders"] += 1
        cross_group[(seller, product)]["amount"] += amount
        cross_group[(seller, product)]["quantity"] += qty
        cross_group[(seller, product)]["orders"] += 1
        if price > 0:
            price_by_product[product].append(price)

    product_table = build_table_from_group(product_group, "产品")
    seller_table = build_table_from_group(seller_group, "销售员")
    month_table = build_table_from_group(month_group, "月份")
    month_table.sort(key=lambda x: str(x["月份"]))

    cross_table = []
    for (seller, product), item in cross_group.items():
        qty = item["quantity"]
        amount = item["amount"]
        orders = int(item["orders"])
        cross_table.append({
            "销售员": seller,
            "产品": product,
            "销售额": round2(amount),
            "销售数量": round2(qty),
            "订单数": orders,
            "平均单价": round2(amount / qty) if qty else 0,
        })
    cross_table.sort(key=lambda x: x["销售额"], reverse=True)

    anomalies = []
    product_price_stats = {}
    for product, prices in price_by_product.items():
        if prices:
            mean = statistics.mean(prices)
            stdev = statistics.pstdev(prices) if len(prices) > 1 else 0
            product_price_stats[product] = (mean, stdev)

    for r in records:
        reasons = []
        qty = float(r["quantity"])
        price = float(r["price"])
        amount = float(r["amount"])
        if price <= 0:
            reasons.append("单价为 0 或负数")
        if qty <= 0:
            reasons.append("数量为 0 或负数")
        if amount <= 0:
            reasons.append("销售额为 0 或负数")
        mean_stdev = product_price_stats.get(str(r["product"]))
        if mean_stdev and mean_stdev[1] > 0 and price > 0:
            mean, stdev = mean_stdev
            if abs(price - mean) > 2 * stdev:
                reasons.append("单价偏离该产品均值较大")
        if reasons:
            anomalies.append({
                "日期": r["date"],
                "月份": r["month"],
                "产品": r["product"],
                "销售员": r["seller"],
                "数量": round2(qty),
                "销售单价": round2(price),
                "销售额": round2(amount),
                "异常原因": "；".join(reasons),
            })

    return {
        "kpi": {
            "总销售额": round2(total_amount),
            "总销售数量": round2(total_qty),
            "订单数": len(records),
            "平均销售单价": round2(total_amount / total_qty) if total_qty else 0,
            "平均订单金额": round2(total_amount / len(records)) if records else 0,
            "产品数量": len(product_group),
            "销售人员数量": len(seller_group),
            "正价单平均单价": round2(sum(valid_price_records) / len(valid_price_records)) if valid_price_records else 0,
        },
        "monthTable": month_table,
        "productTable": product_table,
        "sellerTable": seller_table,
        "crossTable": cross_table,
        "anomalies": anomalies,
        "topProductsByAmount": product_table[:15],
        "topProductsByQuantity": sorted(product_table, key=lambda x: x["销售数量"], reverse=True)[:15],
        "topSellersByAmount": seller_table[:15],
    }


def normalize_rows(rows: list[list[object]]) -> tuple[list[str], list[dict[str, object]]]:
    if not rows:
        raise RuntimeError("Excel 工作表为空")
    headers = [str(v).strip() for v in rows[0]]
    base_required = ["产品", "数量", "销售单价", "成交时间"]
    dimension_field = "销售员" if "销售员" in headers else "产品线" if "产品线" in headers else ""
    required = base_required + ([dimension_field] if dimension_field else [])
    missing = [name for name in base_required if name not in headers]
    if not dimension_field:
        missing.append("销售员或产品线")
    if missing:
        raise RuntimeError(f"缺少必要字段：{', '.join(missing)}；当前字段：{headers}")
    index = {name: headers.index(name) for name in required}

    records = []
    for i, row in enumerate(rows[1:], start=2):
        def get(name: str) -> object:
            col = index[name]
            return row[col] if col < len(row) else ""

        product = str(get("产品")).strip() or "未填写产品"
        seller = str(get(dimension_field)).strip() or f"未填写{dimension_field}"
        quantity = to_float(get("数量"))
        price = to_float(get("销售单价"))
        amount = quantity * price
        dt = excel_date_to_datetime(get("成交时间"))
        date_text = dt.strftime("%Y-%m-%d") if dt else str(get("成交时间"))
        month_text = dt.strftime("%Y-%m") if dt else "未知月份"
        records.append({
            "row": i,
            "product": product,
            "seller": seller,
            "quantity": round2(quantity),
            "price": round2(price),
            "amount": round2(amount),
            "date": date_text,
            "month": month_text,
        })
    return headers, records


def js_json(data: object) -> str:
    return json.dumps(data, ensure_ascii=False, separators=(",", ":"))


def build_html(data: dict[str, object]) -> str:
    return f"""<!doctype html>
<html lang=\"zh-CN\">
<head>
  <meta charset=\"utf-8\" />
  <meta name=\"viewport\" content=\"width=device-width, initial-scale=1\" />
  <title>2026 年销售数据透视看板</title>
  <script src=\"https://cdn.jsdelivr.net/npm/echarts@5/dist/echarts.min.js\"></script>
  <style>
    :root {{
      --bg:#f5f7fb; --card:#ffffff; --text:#172033; --muted:#697386; --line:#e6eaf2;
      --blue:#326bff; --cyan:#00a6c8; --green:#16a34a; --orange:#f59e0b; --red:#ef4444;
      --shadow:0 10px 24px rgba(15,23,42,.08); --radius:18px;
    }}
    * {{ box-sizing:border-box; }}
    body {{ margin:0; font-family:-apple-system,BlinkMacSystemFont,"Segoe UI","Microsoft YaHei",Arial,sans-serif; background:var(--bg); color:var(--text); }}
    .page {{ width:min(1440px,100%); margin:0 auto; padding:24px; }}
    .hero {{ display:flex; justify-content:space-between; gap:16px; align-items:flex-end; margin-bottom:18px; }}
    .hero h1 {{ margin:0 0 8px; font-size:30px; letter-spacing:.2px; }}
    .hero p {{ margin:0; color:var(--muted); line-height:1.7; }}
    .badge {{ white-space:nowrap; color:#fff; background:linear-gradient(135deg,var(--blue),var(--cyan)); padding:10px 14px; border-radius:999px; font-size:13px; box-shadow:var(--shadow); }}
    .kpi-grid {{ display:grid; grid-template-columns:repeat(4,1fr); gap:14px; margin-bottom:16px; }}
    .kpi {{ background:var(--card); border:1px solid var(--line); border-radius:var(--radius); padding:18px; box-shadow:var(--shadow); position:relative; overflow:hidden; }}
    .kpi:after {{ content:""; position:absolute; right:-30px; top:-35px; width:95px; height:95px; border-radius:50%; background:rgba(50,107,255,.08); }}
    .kpi .label {{ color:var(--muted); font-size:13px; margin-bottom:10px; }}
    .kpi .value {{ font-size:25px; font-weight:800; word-break:break-all; }}
    .kpi .sub {{ margin-top:8px; color:var(--muted); font-size:12px; }}
    .grid {{ display:grid; grid-template-columns:1fr 1fr; gap:16px; }}
    .card {{ background:var(--card); border:1px solid var(--line); border-radius:var(--radius); padding:18px; box-shadow:var(--shadow); min-width:0; }}
    .card.full {{ grid-column:1 / -1; }}
    .card-title {{ display:flex; justify-content:space-between; align-items:center; gap:12px; margin-bottom:12px; }}
    .card-title h2 {{ margin:0; font-size:18px; }}
    .card-title span {{ color:var(--muted); font-size:12px; }}
    .chart {{ width:100%; height:360px; }}
    .chart.tall {{ height:460px; }}
    .table-wrap {{ width:100%; overflow:auto; border:1px solid var(--line); border-radius:14px; }}
    table {{ width:100%; min-width:780px; border-collapse:collapse; font-size:13px; background:#fff; }}
    th,td {{ padding:11px 12px; border-bottom:1px solid var(--line); text-align:right; white-space:nowrap; }}
    th:first-child,td:first-child {{ text-align:left; }}
    th {{ position:sticky; top:0; z-index:1; background:#f8fafc; color:#475569; cursor:pointer; user-select:none; font-weight:700; }}
    tr:hover td {{ background:#fafcff; }}
    .danger {{ color:var(--red); font-weight:700; }}
    .tabs {{ display:flex; gap:8px; flex-wrap:wrap; margin-bottom:12px; }}
    .tab {{ border:1px solid var(--line); background:#fff; color:#475569; padding:8px 12px; border-radius:999px; cursor:pointer; }}
    .tab.active {{ background:#172033; color:#fff; border-color:#172033; }}
    .note {{ color:var(--muted); font-size:12px; line-height:1.7; margin-top:10px; }}
    @media (max-width: 960px) {{
      .page {{ padding:16px; }} .hero {{ align-items:flex-start; flex-direction:column; }} .hero h1 {{ font-size:24px; }}
      .kpi-grid {{ grid-template-columns:repeat(2,1fr); }} .grid {{ grid-template-columns:1fr; }}
      .chart {{ height:320px; }} .chart.tall {{ height:380px; }}
    }}
    @media (max-width: 560px) {{
      .page {{ padding:12px; }} .kpi-grid {{ grid-template-columns:1fr; gap:10px; }} .kpi,.card {{ border-radius:14px; padding:14px; }}
      .kpi .value {{ font-size:22px; }} .chart {{ height:280px; }} .chart.tall {{ height:340px; }}
      table {{ font-size:12px; min-width:700px; }} th,td {{ padding:9px 10px; }}
    }}
  </style>
</head>
<body>
  <main class=\"page\">
    <section class=\"hero\">
      <div>
        <h1>2026 年销售数据透视看板</h1>
        <p>基于 Excel 原始销售记录生成，覆盖月份、产品、销售员、交叉透视与异常数据复核。</p>
      </div>
      <div class=\"badge\">数据更新：{data['generatedAt']}</div>
    </section>

    <section class=\"kpi-grid\" id=\"kpiGrid\"></section>

    <section class=\"grid\">
      <div class=\"card full\">
        <div class=\"card-title\"><h2>月度销售趋势</h2><span>销售额 / 数量 / 订单数</span></div>
        <div id=\"monthTrend\" class=\"chart\"></div>
      </div>
      <div class=\"card\">
        <div class=\"card-title\"><h2>产品销售额 TOP</h2><span>按销售额降序</span></div>
        <div id=\"productAmount\" class=\"chart tall\"></div>
      </div>
      <div class=\"card\">
        <div class=\"card-title\"><h2>产品销售数量 TOP</h2><span>按数量降序</span></div>
        <div id=\"productQty\" class=\"chart tall\"></div>
      </div>
      <div class=\"card\">
        <div class=\"card-title\"><h2>销售员业绩排行</h2><span>销售额排行</span></div>
        <div id=\"sellerRank\" class=\"chart\"></div>
      </div>
      <div class=\"card\">
        <div class=\"card-title\"><h2>销售员销售额占比</h2><span>结构占比</span></div>
        <div id=\"sellerPie\" class=\"chart\"></div>
      </div>
      <div class=\"card full\">
        <div class=\"card-title\"><h2>数据透视表</h2><span>点击表头可排序，移动端可横向滑动</span></div>
        <div class=\"tabs\">
          <button class=\"tab active\" data-table=\"productTable\">产品透视</button>
          <button class=\"tab\" data-table=\"sellerTable\">人员/产品线透视</button>
          <button class=\"tab\" data-table=\"monthTable\">月份透视</button>
          <button class=\"tab\" data-table=\"crossTable\">维度 × 产品</button>
          <button class=\"tab\" data-table=\"anomalies\">异常数据</button>
        </div>
        <div class=\"table-wrap\" id=\"tableWrap\"></div>
        <div class=\"note\">异常数据规则：单价 ≤ 0、数量 ≤ 0、销售额 ≤ 0，或单价相对同产品均值偏离较大。</div>
      </div>
    </section>
  </main>

<script>
const DATA = {js_json(data)};
const fmtNumber = (n, digits=2) => Number(n || 0).toLocaleString('zh-CN', {{maximumFractionDigits:digits, minimumFractionDigits:digits}});
const fmtInt = n => Number(n || 0).toLocaleString('zh-CN');
const isMoneyKey = key => /销售额|单价|客单价|金额/.test(key);

function renderKpis() {{
  const k = DATA.kpi;
  const items = [
    ['总销售额', '¥' + fmtNumber(k['总销售额']), '所有订单数量 × 单价汇总'],
    ['总销售数量', fmtNumber(k['总销售数量']), '累计销售数量'],
    ['订单数', fmtInt(k['订单数']), '原始有效销售记录数'],
    ['平均销售单价', '¥' + fmtNumber(k['平均销售单价']), '总销售额 / 总数量'],
    ['平均订单金额', '¥' + fmtNumber(k['平均订单金额']), '总销售额 / 订单数'],
    ['产品数量', fmtInt(k['产品数量']), '去重产品型号数'],
    [DATA.dimensionName + '数量', fmtInt(k['销售人员数量']), '去重' + DATA.dimensionName + '数量'],
    ['异常记录数', fmtInt(DATA.anomalies.length), '需业务复核的记录'],
  ];
  document.getElementById('kpiGrid').innerHTML = items.map(([label,value,sub]) => `
    <article class="kpi"><div class="label">${{label}}</div><div class="value">${{value}}</div><div class="sub">${{sub}}</div></article>
  `).join('');
}}

function initChart(id, option) {{
  const chart = echarts.init(document.getElementById(id));
  chart.setOption(option);
  window.addEventListener('resize', () => chart.resize());
  return chart;
}}

function renderCharts() {{
  const months = DATA.monthTable.map(x => x['月份']);
  initChart('monthTrend', {{
    tooltip: {{ trigger:'axis' }}, legend: {{ top:0 }}, grid: {{ left:50, right:42, top:54, bottom:36 }},
    xAxis: {{ type:'category', data:months }},
    yAxis: [{{ type:'value', name:'销售额' }}, {{ type:'value', name:'数量/订单' }}],
    series: [
      {{ name:'销售额', type:'line', smooth:true, data:DATA.monthTable.map(x=>x['销售额']), areaStyle:{{opacity:.12}}, itemStyle:{{color:'#326bff'}} }},
      {{ name:'销售数量', type:'bar', yAxisIndex:1, data:DATA.monthTable.map(x=>x['销售数量']), itemStyle:{{color:'#00a6c8'}} }},
      {{ name:'订单数', type:'line', yAxisIndex:1, data:DATA.monthTable.map(x=>x['订单数']), itemStyle:{{color:'#f59e0b'}} }}
    ]
  }});

  const productAmount = [...DATA.topProductsByAmount].reverse();
  initChart('productAmount', {{
    tooltip: {{ trigger:'axis', axisPointer:{{type:'shadow'}} }}, grid: {{ left:90, right:28, top:16, bottom:24 }},
    xAxis: {{ type:'value' }}, yAxis: {{ type:'category', data:productAmount.map(x=>x['产品']) }},
    series: [{{ name:'销售额', type:'bar', data:productAmount.map(x=>x['销售额']), itemStyle:{{color:'#326bff', borderRadius:[0,8,8,0]}} }}]
  }});

  const productQty = [...DATA.topProductsByQuantity].reverse();
  initChart('productQty', {{
    tooltip: {{ trigger:'axis', axisPointer:{{type:'shadow'}} }}, grid: {{ left:90, right:28, top:16, bottom:24 }},
    xAxis: {{ type:'value' }}, yAxis: {{ type:'category', data:productQty.map(x=>x['产品']) }},
    series: [{{ name:'销售数量', type:'bar', data:productQty.map(x=>x['销售数量']), itemStyle:{{color:'#16a34a', borderRadius:[0,8,8,0]}} }}]
  }});

  initChart('sellerRank', {{
    tooltip: {{ trigger:'axis', axisPointer:{{type:'shadow'}} }}, grid: {{ left:52, right:20, top:26, bottom:60 }},
    xAxis: {{ type:'category', data:DATA.topSellersByAmount.map(x=>x['销售员']), axisLabel:{{rotate:35}} }},
    yAxis: {{ type:'value' }},
    series: [{{ name:'销售额', type:'bar', data:DATA.topSellersByAmount.map(x=>x['销售额']), itemStyle:{{color:'#00a6c8', borderRadius:[8,8,0,0]}} }}]
  }});

  initChart('sellerPie', {{
    tooltip: {{ trigger:'item' }}, legend: {{ type:'scroll', bottom:0 }},
    series: [{{ name:'销售额占比', type:'pie', radius:['42%','68%'], center:['50%','45%'], data:DATA.topSellersByAmount.map(x=>({{name:x['销售员'], value:x['销售额']}})) }}]
  }});
}}

let currentTable = 'productTable';
let sortState = {{ key: '销售额', dir: -1 }};
function renderTable(name=currentTable) {{
  currentTable = name;
  const rows = [...(DATA[name] || [])];
  if (!rows.length) {{
    document.getElementById('tableWrap').innerHTML = '<table><tbody><tr><td>暂无数据</td></tr></tbody></table>';
    return;
  }}
  const keys = Object.keys(rows[0]);
  rows.sort((a,b) => {{
    const av = a[sortState.key], bv = b[sortState.key];
    if (typeof av === 'number' && typeof bv === 'number') return (av-bv)*sortState.dir;
    return String(av ?? '').localeCompare(String(bv ?? ''), 'zh-CN') * sortState.dir;
  }});
  const html = `<table><thead><tr>${{keys.map(k=>`<th data-key="${{k}}">${{k}} ${{sortState.key===k ? (sortState.dir>0?'▲':'▼') : ''}}</th>`).join('')}}</tr></thead><tbody>${{rows.map(r=>`<tr>${{keys.map(k=>{{
    const v = r[k];
    const cls = k === '异常原因' ? ' class="danger"' : '';
    const text = typeof v === 'number' ? (isMoneyKey(k) ? '¥' + fmtNumber(v) : fmtNumber(v)) : (v ?? '');
    return `<td${{cls}}>${{text}}</td>`;
  }}).join('')}}</tr>`).join('')}}</tbody></table>`;
  document.getElementById('tableWrap').innerHTML = html;
  document.querySelectorAll('th[data-key]').forEach(th => {{
    th.addEventListener('click', () => {{
      const key = th.dataset.key;
      sortState = sortState.key === key ? {{key, dir: -sortState.dir}} : {{key, dir: -1}};
      renderTable(currentTable);
    }});
  }});
}}

function bindTabs() {{
  document.querySelectorAll('.tab').forEach(btn => {{
    btn.addEventListener('click', () => {{
      document.querySelectorAll('.tab').forEach(x=>x.classList.remove('active'));
      btn.classList.add('active');
      sortState = {{ key: btn.dataset.table === 'anomalies' ? '销售额' : '销售额', dir: -1 }};
      renderTable(btn.dataset.table);
    }});
  }});
}}

renderKpis();
renderCharts();
bindTabs();
renderTable('productTable');
</script>
</body>
</html>
"""


def main() -> None:
    sheet_name, rows = read_first_sheet_rows(XLSX_PATH)
    headers, records = normalize_rows(rows)
    summary = aggregate(records)
    data = {
        **summary,
        "sheetName": sheet_name,
        "headers": headers,
        "recordCount": len(records),
        "generatedAt": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "dimensionName": "销售员" if "销售员" in headers else "产品线",
    }
    HTML_PATH.write_text(build_html(data), encoding="utf-8")
    print(json.dumps({
        "output": str(HTML_PATH),
        "sheet": sheet_name,
        "headers": headers,
        "records": len(records),
        "kpi": data["kpi"],
        "anomalies": len(data["anomalies"]),
    }, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
