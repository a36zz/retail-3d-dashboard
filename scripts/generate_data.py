"""
Generates a fully SYNTHETIC retail sales workbook (data/sales_data.xlsx).

Nothing here is real: names, branches, prices and quantities are random numbers
drawn with a fixed seed so the result is reproducible.

    python scripts/generate_data.py            # default 4,200 orders
    python scripts/generate_data.py 8000       # custom size
"""
import sys, random, math, datetime, pathlib
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

ROOT = pathlib.Path(__file__).resolve().parent.parent
OUT = ROOT / "data" / "sales_data.xlsx"

N = int(sys.argv[1]) if len(sys.argv) > 1 else 4200
random.seed(2026)

# region -> (weight, branches)
REGIONS = {
    "North America": (1.25, ["New York", "Toronto", "Chicago", "Los Angeles"]),
    "Europe":        (1.10, ["London", "Berlin", "Madrid", "Paris"]),
    "Middle East":   (0.95, ["Dubai", "Riyadh", "Doha", "Istanbul"]),
    "Asia Pacific":  (1.20, ["Singapore", "Tokyo", "Sydney", "Mumbai"]),
    "Latin America": (0.70, ["Sao Paulo", "Mexico City", "Bogota", "Santiago"]),
    "Africa":        (0.55, ["Lagos", "Nairobi", "Cairo", "Cape Town"]),
}

# category -> (weight, margin, [(product, base price)])
CATEGORIES = {
    "Electronics":      (1.30, 0.22, [("Wireless Headphones", 129), ("Smart Watch", 249), ("4K Monitor", 389), ("Portable Speaker", 89)]),
    "Home & Living":    (1.00, 0.34, [("Ergonomic Chair", 279), ("Standing Desk", 459), ("Air Purifier", 199), ("LED Lamp Set", 59)]),
    "Fashion":          (1.10, 0.48, [("Running Shoes", 119), ("Leather Jacket", 229), ("Travel Backpack", 79), ("Sunglasses", 99)]),
    "Sports & Outdoor": (0.80, 0.38, [("Yoga Mat Pro", 45), ("Trail Tent", 189), ("Cycling Helmet", 69), ("Smart Bottle", 39)]),
    "Beauty & Care":    (0.75, 0.55, [("Skincare Set", 85), ("Hair Dryer", 119), ("Fragrance 100ml", 99), ("Electric Toothbrush", 79)]),
    "Grocery & Gourmet": (0.65, 0.27, [("Specialty Coffee 1kg", 34), ("Olive Oil Box", 52), ("Organic Tea Set", 28), ("Gourmet Gift Basket", 74)]),
}

CHANNELS = [("Store", 0.40), ("Online", 0.35), ("Mobile App", 0.17), ("Partner", 0.08)]
SEGMENTS = [("Consumer", 0.60), ("Corporate", 0.25), ("Small Business", 0.15)]

START = datetime.date(2025, 1, 1)
END = datetime.date(2026, 10, 31)
DAYS = (END - START).days + 1

# month seasonality (1.0 = average) and steady growth over time
SEASON = {1: .86, 2: .84, 3: .95, 4: .98, 5: 1.02, 6: .97, 7: .93, 8: .96, 9: 1.08, 10: 1.12, 11: 1.35, 12: 1.45}


def pick(pairs):
    r, acc = random.random() * sum(w for _, w in pairs), 0
    for v, w in pairs:
        acc += w
        if r <= acc:
            return v
    return pairs[-1][0]


# branch weights (some stronger than others)
branches = []
for reg, (rw, bl) in REGIONS.items():
    for i, b in enumerate(bl):
        branches.append((reg, b, rw * (1.35 - i * 0.22) * random.uniform(0.85, 1.2)))

# day weights: seasonality * growth, weekend bump
day_w = []
for d in range(DAYS):
    dt = START + datetime.timedelta(days=d)
    growth = 1 + 0.55 * d / DAYS
    wk = 1.18 if dt.weekday() in (4, 5) else 1.0
    day_w.append(SEASON[dt.month] * growth * wk)
cum = []
s = 0
for w in day_w:
    s += w
    cum.append(s)


def rand_day():
    r = random.random() * s
    lo, hi = 0, len(cum) - 1
    while lo < hi:
        mid = (lo + hi) // 2
        if cum[mid] < r:
            lo = mid + 1
        else:
            hi = mid
    return START + datetime.timedelta(days=lo)


cat_pairs = [(c, v[0]) for c, v in CATEGORIES.items()]
br_pairs = [((r, b), w) for r, b, w in branches]

rows = []
for _ in range(N):
    dt = rand_day()
    (reg, br) = pick(br_pairs)
    cat = pick(cat_pairs)
    cw, margin, prods = CATEGORIES[cat]
    prod, base = random.choice(prods)
    ch = pick(CHANNELS)
    sg = pick(SEGMENTS)
    units = max(1, int(round(random.lognormvariate(0.35, 0.55))))
    if sg == "Corporate":
        units = int(units * random.uniform(1.6, 3.2)) + 1
    elif sg == "Small Business":
        units = int(units * random.uniform(1.2, 2.0)) + 1
    units = min(units, 40)
    disc = 1 - (0.06 if ch == "Online" else 0.0) - (0.1 if sg == "Corporate" else 0.0) - random.choice([0, 0, 0, .05, .1])
    price = round(base * random.lognormvariate(0, 0.07) * disc, 2)
    m = min(0.7, max(0.05, margin + random.gauss(0, 0.05) - (0.08 if sg == "Corporate" else 0)))
    unit_cost = round(price * (1 - m), 2)
    rows.append([dt, reg, br, cat, prod, ch, sg, units, round(price * units, 2), round(unit_cost * units, 2)])

rows.sort(key=lambda r: r[0])

wb = Workbook()
ws = wb.active
ws.title = "Sales"
HEAD = ["Order ID", "Order Date", "Region", "Branch", "Category", "Product", "Channel", "Segment", "Units", "Revenue", "Cost"]
ws.append(HEAD)
for i, r in enumerate(rows):
    ws.append([f"ORD-{100001 + i}"] + r)

thin = Side(style="thin", color="D9DEEA")
for c in ws[1]:
    c.font = Font(bold=True, color="FFFFFF", name="Calibri")
    c.fill = PatternFill("solid", fgColor="312E81")
    c.alignment = Alignment(horizontal="center", vertical="center")
widths = [13, 13, 16, 15, 18, 22, 12, 15, 8, 12, 12]
for i, w in enumerate(widths, 1):
    ws.column_dimensions[get_column_letter(i)].width = w
for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
    row[1].number_format = "yyyy-mm-dd"
    row[9].number_format = "#,##0.00"
    row[10].number_format = "#,##0.00"
ws.freeze_panes = "A2"
tab = Table(displayName="SalesTable", ref=f"A1:K{ws.max_row}")
tab.tableStyleInfo = TableStyleInfo(name="TableStyleLight1", showRowStripes=True)
ws.add_table(tab)

about = wb.create_sheet("About")
about.column_dimensions["A"].width = 26
about.column_dimensions["B"].width = 90
info = [
    ("About this workbook", ""),
    ("Nature of the data", "100% synthetic. Every name, number and date was generated randomly for demonstration."),
    ("Used by", "index.html (3D retail analytics dashboard). The dashboard reads sheet 1 (\"Sales\"), columns A:K."),
    ("How to update", "Edit or append rows in the Sales sheet, save, then run Update-Data.bat (Windows) or push to GitHub."),
    ("Columns", "Order ID, Order Date, Region, Branch, Category, Product, Channel, Segment, Units, Revenue, Cost"),
    ("Profit", "Calculated by the dashboard as Revenue - Cost."),
    ("Rules", "Keep headers in row 1, dates as real Excel dates, Units / Revenue / Cost as numbers."),
]
for r in info:
    about.append(list(r))
about["A1"].font = Font(bold=True, size=14, color="312E81")
for r in range(2, len(info) + 1):
    about.cell(r, 1).font = Font(bold=True)
    about.cell(r, 2).alignment = Alignment(wrap_text=True, vertical="top")

OUT.parent.mkdir(exist_ok=True)
wb.save(OUT)
tot = sum(r[8] for r in rows)
print(f"{len(rows)} orders | revenue {tot:,.0f} | {rows[0][0]} -> {rows[-1][0]} | saved {OUT}")
