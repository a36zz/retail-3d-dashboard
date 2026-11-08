"""
Builds index.html (a single, self-contained file) from:
    src/template.html  +  data/sales_data.xlsx  +  config.json  +  vendor/*.js

    python scripts/build.py                     # normal build
    python scripts/build.py --xlsx other.xlsx   # use another workbook

Needs only:  pip install openpyxl
"""
import argparse, datetime, html, json, pathlib, sys
from openpyxl import load_workbook

ROOT = pathlib.Path(__file__).resolve().parent.parent
FIELDS = ["Order ID", "Order Date", "Region", "Branch", "Category", "Product", "Channel", "Segment", "Units", "Revenue", "Cost"]


def read_rows(xlsx: pathlib.Path):
    wb = load_workbook(xlsx, read_only=True, data_only=True)
    ws = wb.worksheets[0]  # first sheet, whatever its name is
    it = ws.iter_rows(values_only=True)
    head = [str(h).strip() if h is not None else "" for h in next(it)]
    idx = {}
    for f in FIELDS:
        if f not in head:
            sys.exit(f'Column "{f}" was not found in sheet "{ws.title}". Found: {head}')
        idx[f] = head.index(f)
    rows = []
    for r in it:
        if r is None or r[idx["Order ID"]] in (None, ""):
            continue
        d = r[idx["Order Date"]]
        if isinstance(d, datetime.datetime):
            d = d.date()
        if isinstance(d, datetime.date):
            d = d.isoformat()
        else:
            d = str(d or "")[:10]
        num = lambda v: round(float(v or 0), 2)
        rows.append([
            str(r[idx["Order ID"]]).strip(), d,
            str(r[idx["Region"]] or "").strip(), str(r[idx["Branch"]] or "").strip(),
            str(r[idx["Category"]] or "").strip(), str(r[idx["Product"]] or "").strip(),
            str(r[idx["Channel"]] or "").strip(), str(r[idx["Segment"]] or "").strip(),
            int(float(r[idx["Units"]] or 0)), num(r[idx["Revenue"]]), num(r[idx["Cost"]]),
        ])
    return rows


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--xlsx", default=str(ROOT / "data" / "sales_data.xlsx"))
    ap.add_argument("--out", default=str(ROOT / "index.html"))
    a = ap.parse_args()

    cfg = json.loads((ROOT / "config.json").read_text(encoding="utf-8"))
    rows = read_rows(pathlib.Path(a.xlsx))
    data = json.dumps({"built": datetime.datetime.now().strftime("%Y-%m-%d %H:%M"), "rows": rows},
                      ensure_ascii=False, separators=(",", ":")).replace("</", "<\\/")

    three = (ROOT / "vendor" / "three.min.js").read_text(encoding="utf-8")
    orbit = (ROOT / "vendor" / "OrbitControls.js").read_text(encoding="utf-8")
    assert "</script" not in three and "</script" not in orbit

    links = []
    if cfg.get("github"):
        links.append(f'<a href="{html.escape(cfg["github"])}" target="_blank" rel="noopener">GitHub</a>')
    if cfg.get("linkedin"):
        links.append(f'<a href="{html.escape(cfg["linkedin"])}" target="_blank" rel="noopener">LinkedIn</a>')
    who = html.escape(cfg.get("author", "").strip())
    if who:
        author = f"Built by <strong style=\"color:#c7d0ea\">{who}</strong>" + ("  ·  " + "  ·  ".join(links) if links else "")
    else:
        author = "Built with Three.js and vanilla JavaScript" + ("  ·  " + "  ·  ".join(links) if links else "")

    t = (ROOT / "src" / "template.html").read_text(encoding="utf-8")
    for k, v in [("{{TITLE}}", html.escape(cfg["title"])), ("{{SUBTITLE}}", html.escape(cfg["subtitle"])),
                 ("{{AUTHOR_HTML}}", author), ("__DATA__", data),
                 ("/*__THREE__*/", three), ("/*__ORBIT__*/", orbit)]:
        assert k in t, k
        t = t.replace(k, v)
    out = pathlib.Path(a.out)
    out.write_text(t, encoding="utf-8")
    print(f"{len(rows)} orders -> {out.name}  ({out.stat().st_size/1048576:.2f} MB)")


if __name__ == "__main__":
    main()
