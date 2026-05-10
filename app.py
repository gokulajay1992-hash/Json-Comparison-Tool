from flask import Flask, render_template, request, jsonify, send_file
import oracledb
import json
import io
from datetime import datetime
import openpyxl
from openpyxl.styles import PatternFill, Font

app = Flask(__name__)

# ---------------------------------------------------------------------------
# Config loader
# ---------------------------------------------------------------------------

def load_config():
    with open("config.json", "r") as f:
        return json.load(f)

# ---------------------------------------------------------------------------
# Oracle helpers
# ---------------------------------------------------------------------------

def get_connection(host, port, service, username, password):
    dsn = oracledb.makedsn(host, int(port), service_name=service)
    return oracledb.connect(user=username, password=password, dsn=dsn)

# ---------------------------------------------------------------------------
# JSON diff engine (recursive, handles nested objects and arrays)
# ---------------------------------------------------------------------------

def deep_diff(a, b, path=""):
    diffs = []

    if type(a) != type(b):
        diffs.append({"path": path or "root", "old": _fmt(a), "new": _fmt(b), "type": "changed"})
        return diffs

    if isinstance(a, dict):
        for key in set(a) | set(b):
            p = f"{path}.{key}" if path else key
            if key not in a:
                diffs.append({"path": p, "old": "", "new": _fmt(b[key]), "type": "added"})
            elif key not in b:
                diffs.append({"path": p, "old": _fmt(a[key]), "new": "", "type": "removed"})
            else:
                diffs.extend(deep_diff(a[key], b[key], p))
    elif isinstance(a, list):
        for i in range(max(len(a), len(b))):
            p = f"{path}[{i}]"
            if i >= len(a):
                diffs.append({"path": p, "old": "", "new": _fmt(b[i]), "type": "added"})
            elif i >= len(b):
                diffs.append({"path": p, "old": _fmt(a[i]), "new": "", "type": "removed"})
            else:
                diffs.extend(deep_diff(a[i], b[i], p))
    else:
        if a != b:
            diffs.append({"path": path or "root", "old": _fmt(a), "new": _fmt(b), "type": "changed"})

    return diffs


def _fmt(val):
    if val is None:
        return ""
    if isinstance(val, (dict, list)):
        return json.dumps(val)
    return str(val)


def parse_json(val):
    if val is None:
        return None
    if isinstance(val, str):
        try:
            return json.loads(val)
        except Exception:
            return val
    return val

# ---------------------------------------------------------------------------
# Routes
# ---------------------------------------------------------------------------

@app.route("/")
def index():
    cfg = load_config()
    return render_template("index.html", config=cfg)


@app.route("/config")
def get_config():
    return jsonify(load_config())


@app.route("/query", methods=["POST"])
def query():
    body = request.json
    cfg = load_config()
    col = cfg["columns"]
    table = cfg["table_name"]

    # Build SELECT column list from config only — never hardcode
    select_cols = [col["primary_id"], col["datetime"]]
    select_cols += col.get("extra_ids", [])
    for s in col["json_sets"]:
        for side in (s["left"], s["right"]):
            if side not in select_cols:
                select_cols.append(side)

    col_list = ", ".join(select_cols)

    dt_col = col["datetime"]
    sql = (
        f"SELECT {col_list} FROM {table} "
        f"WHERE {dt_col} >= TO_TIMESTAMP(:date_from, 'YYYY-MM-DD HH24:MI:SS') "
        f"AND   {dt_col} <= TO_TIMESTAMP(:date_to,   'YYYY-MM-DD HH24:MI:SS') "
        f"ORDER BY {dt_col}"
    )

    try:
        conn = get_connection(
            body["host"], body["port"], body["service"],
            body["username"], body["password"]
        )
        cursor = conn.cursor()
        cursor.execute(sql, date_from=body["date_from"], date_to=body["date_to"])
        rows = cursor.fetchall()
        cursor.close()
        conn.close()
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500

    results = []
    for row in rows:
        row_map = {c: row[i] for i, c in enumerate(select_cols)}

        record = {
            "id":       str(row_map.get(col["primary_id"], "")),
            "dat_time": str(row_map.get(col["datetime"], "")),
            "extra":    {k: str(row_map.get(k, "") or "") for k in col.get("extra_ids", [])},
            "sets":     []
        }

        for s in col["json_sets"]:
            left  = parse_json(row_map.get(s["left"]))
            right = parse_json(row_map.get(s["right"]))
            if left is not None and right is not None:
                diffs = deep_diff(left, right)
            else:
                diffs = []
            record["sets"].append({
                "label":    s["label"],
                "left_col": s["left"],
                "right_col": s["right"],
                "diffs":    diffs,
                "has_diff": len(diffs) > 0
            })

        results.append(record)

    return jsonify({"status": "ok", "data": results, "count": len(results)})


@app.route("/export", methods=["POST"])
def export():
    body = request.json
    results = body["results"]
    cfg = load_config()
    extra_id_cols = cfg["columns"].get("extra_ids", [])

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "JSON Diff"

    hdr_fill    = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    hdr_font    = Font(color="FFFFFF", bold=True)
    add_fill    = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    rem_fill    = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    chg_fill    = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")

    base_headers = [cfg["columns"]["primary_id"]] + extra_id_cols + [cfg["columns"]["datetime"]]
    headers = base_headers + ["Comparison Set", "Left Column", "Right Column",
                               "Path", "Difference Type", "Old Value", "New Value"]
    ws.append(headers)
    for cell in ws[1]:
        cell.fill = hdr_fill
        cell.font = hdr_font

    type_fill = {"added": add_fill, "removed": rem_fill, "changed": chg_fill}

    for rec in results:
        base_vals = [rec["id"]] + [rec["extra"].get(k, "") for k in extra_id_cols] + [rec["dat_time"]]
        for s in rec["sets"]:
            if not s["diffs"]:
                ws.append(base_vals + [s["label"], s["left_col"], s["right_col"],
                                        "(no differences)", "", "", ""])
            for diff in s["diffs"]:
                row_vals = base_vals + [s["label"], s["left_col"], s["right_col"],
                                         diff["path"], diff["type"],
                                         diff["old"], diff["new"]]
                ws.append(row_vals)
                fill = type_fill.get(diff["type"])
                if fill:
                    for cell in ws[ws.max_row]:
                        cell.fill = fill

    for col_cells in ws.columns:
        width = max(len(str(c.value or "")) for c in col_cells)
        ws.column_dimensions[col_cells[0].column_letter].width = min(width + 4, 60)

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)

    fname = f"json_diff_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    return send_file(out,
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                     as_attachment=True,
                     download_name=fname)


@app.route("/demo", methods=["GET"])
def demo():
    cfg = load_config()
    col = cfg["columns"]
    extra_ids = col.get("extra_ids", [])

    sample_records = [
        {
            "id": "TXN-1001",
            "dat_time": "2026-05-10 09:15:00",
            "extra": {k: v for k, v in zip(extra_ids, ["C001", "ORD-55", "REG-AU"])},
            "sets": [
                {
                    "label": "Set 1",
                    "left_col": col["json_sets"][0]["left"],
                    "right_col": col["json_sets"][0]["right"],
                    "diffs": [
                        {"path": "customer.name",    "old": "John Smith",  "new": "John A. Smith", "type": "changed"},
                        {"path": "customer.email",   "old": "john@old.com","new": "john@new.com",  "type": "changed"},
                        {"path": "order.total",      "old": "250.00",      "new": "275.50",        "type": "changed"},
                        {"path": "order.discount",   "old": "",            "new": "10.00",         "type": "added"},
                    ],
                    "has_diff": True,
                },
                {
                    "label": "Set 2",
                    "left_col": col["json_sets"][1]["left"] if len(col["json_sets"]) > 1 else "Json3",
                    "right_col": col["json_sets"][1]["right"] if len(col["json_sets"]) > 1 else "Json4",
                    "diffs": [],
                    "has_diff": False,
                },
            ],
        },
        {
            "id": "TXN-1002",
            "dat_time": "2026-05-10 10:30:22",
            "extra": {k: v for k, v in zip(extra_ids, ["C002", "ORD-56", "REG-US"])},
            "sets": [
                {
                    "label": "Set 1",
                    "left_col": col["json_sets"][0]["left"],
                    "right_col": col["json_sets"][0]["right"],
                    "diffs": [],
                    "has_diff": False,
                },
                {
                    "label": "Set 2",
                    "left_col": col["json_sets"][1]["left"] if len(col["json_sets"]) > 1 else "Json3",
                    "right_col": col["json_sets"][1]["right"] if len(col["json_sets"]) > 1 else "Json4",
                    "diffs": [
                        {"path": "shipping.address.street", "old": "12 Main St",   "new": "99 Oak Ave",    "type": "changed"},
                        {"path": "shipping.address.city",   "old": "Springfield",  "new": "Shelbyville",   "type": "changed"},
                        {"path": "shipping.method",         "old": "standard",     "new": "express",       "type": "changed"},
                        {"path": "shipping.tracking",       "old": "TRK-OLD-999",  "new": "",              "type": "removed"},
                        {"path": "shipping.insurance",      "old": "",             "new": "true",          "type": "added"},
                    ],
                    "has_diff": True,
                },
            ],
        },
        {
            "id": "TXN-1003",
            "dat_time": "2026-05-10 11:05:47",
            "extra": {k: v for k, v in zip(extra_ids, ["C003", "ORD-57", "REG-UK"])},
            "sets": [
                {
                    "label": "Set 1",
                    "left_col": col["json_sets"][0]["left"],
                    "right_col": col["json_sets"][0]["right"],
                    "diffs": [
                        {"path": "items[0].qty",      "old": "2",       "new": "5",       "type": "changed"},
                        {"path": "items[0].price",    "old": "19.99",   "new": "17.99",   "type": "changed"},
                        {"path": "items[1].sku",      "old": "SKU-AAA", "new": "",        "type": "removed"},
                        {"path": "items[2]",          "old": "",        "new": "{\"sku\":\"SKU-ZZZ\",\"qty\":1}", "type": "added"},
                        {"path": "payment.method",    "old": "card",    "new": "wallet",  "type": "changed"},
                    ],
                    "has_diff": True,
                },
                {
                    "label": "Set 2",
                    "left_col": col["json_sets"][1]["left"] if len(col["json_sets"]) > 1 else "Json3",
                    "right_col": col["json_sets"][1]["right"] if len(col["json_sets"]) > 1 else "Json4",
                    "diffs": [
                        {"path": "status",            "old": "pending", "new": "confirmed", "type": "changed"},
                        {"path": "approvedBy",        "old": "",        "new": "mgr_jane",  "type": "added"},
                    ],
                    "has_diff": True,
                },
            ],
        },
        {
            "id": "TXN-1004",
            "dat_time": "2026-05-10 14:22:10",
            "extra": {k: v for k, v in zip(extra_ids, ["C004", "ORD-58", "REG-SG"])},
            "sets": [
                {
                    "label": "Set 1",
                    "left_col": col["json_sets"][0]["left"],
                    "right_col": col["json_sets"][0]["right"],
                    "diffs": [],
                    "has_diff": False,
                },
                {
                    "label": "Set 2",
                    "left_col": col["json_sets"][1]["left"] if len(col["json_sets"]) > 1 else "Json3",
                    "right_col": col["json_sets"][1]["right"] if len(col["json_sets"]) > 1 else "Json4",
                    "diffs": [],
                    "has_diff": False,
                },
            ],
        },
    ]

    return jsonify({"status": "ok", "data": sample_records, "count": len(sample_records)})


if __name__ == "__main__":
    app.run(debug=True, port=5000)
