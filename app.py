"""
Canteen Expenses Dashboard – Flask backend.

Period format in Excel: "DD.MM.YYYY - DD.MM.YYYY"  (fortnightly blocks).
All period lists returned by the API are sorted chronologically by start date.
"""

import os
import re
from collections import defaultdict
from datetime import datetime

import pandas as pd
from flask import Flask, jsonify, render_template, request

app = Flask(__name__)

EXCEL_PATH = os.environ.get("EXCEL_PATH", "Canteen.xlsx")

# ── nationality config ────────────────────────────────────────────────────────
NAT_SHEETS = ["Bangladeshi", "Indian", "Malagasy", "Srilankan"]
NAT_LABELS = {
    "Bangladeshi": "Bangladeshi",
    "Indian": "Indian",
    "Malagasy": "Malagasy",
    "Srilankan": "Sri Lankan",
}
NAT_COLORS = {
    "Bangladeshi": "#4472ca",
    "Indian":      "#d4a017",
    "Malagasy":    "#70ad47",
    "Srilankan":   "#c55a11",
}
EMP_COLS = {
    "Bangladeshi": "BANGLADESHI",
    "Indian":      "INDIAN",
    "Malagasy":    "MALAGASY",
    "Srilankan":   "SRILANKAN",
}

# ── period helpers ─────────────────────────────────────────────────────────────

def _parse_start(period: str):
    """Return datetime of the start date in 'DD.MM.YYYY ...' or None."""
    if not period:
        return None
    m = re.match(r"(\d{2})\.(\d{2})\.(\d{4})", str(period).strip())
    if not m:
        return None
    try:
        return datetime(int(m.group(3)), int(m.group(2)), int(m.group(1)))
    except ValueError:
        return None


def _normalize(period: str) -> str:
    """Normalise dash variants to ' - '."""
    s = str(period).strip()
    return re.sub(r"\s*[-–—]\s*", " - ", s)


def _period_key(p: str):
    dt = _parse_start(p)
    return dt if dt else datetime.min


def _month_label(p: str) -> str:
    dt = _parse_start(p)
    return dt.strftime("%b %Y") if dt else "Unknown"


def _month_dt(p: str):
    dt = _parse_start(p)
    return datetime(dt.year, dt.month, 1) if dt else None


# ── data loading ───────────────────────────────────────────────────────────────

_cache: dict | None = None
_cache_mtime: float | None = None


def _load() -> dict:
    global _cache, _cache_mtime

    if not os.path.exists(EXCEL_PATH):
        return _sample_data()

    mtime = os.path.getmtime(EXCEL_PATH)
    if _cache is not None and _cache_mtime == mtime:
        return _cache

    try:
        xl = pd.ExcelFile(EXCEL_PATH)
        sheet_map = {s.lower(): s for s in xl.sheet_names}
        data: dict = {}

        for nat in NAT_SHEETS:
            sheet = sheet_map.get(nat.lower())
            if sheet:
                df = xl.parse(sheet)
                df.columns = [str(c).strip().upper() for c in df.columns]
                if "PERIOD" in df.columns:
                    df["PERIOD"] = df["PERIOD"].apply(
                        lambda v: _normalize(v) if pd.notna(v) else v
                    )
                data[nat] = df
            else:
                data[nat] = pd.DataFrame(
                    columns=["PERIOD", "ITEMS", "QTY", "UNIT", "UNIT PRICE", "TOTAL"]
                )

        emp_sheet = sheet_map.get("employees")
        if emp_sheet:
            df = xl.parse(emp_sheet)
            df.columns = [str(c).strip().upper() for c in df.columns]
            if "PERIOD" in df.columns:
                df["PERIOD"] = df["PERIOD"].apply(
                    lambda v: _normalize(v) if pd.notna(v) else v
                )
            data["EMPLOYEES"] = df
        else:
            data["EMPLOYEES"] = pd.DataFrame(
                columns=["PERIOD", "BANGLADESHI", "INDIAN", "MALAGASY", "SRILANKAN"]
            )

        _cache = data
        _cache_mtime = mtime
        return data

    except Exception as exc:  # noqa: BLE001
        print(f"[ERROR] Could not load Excel file: {exc}")
        return _sample_data()


def _sample_data() -> dict:
    """Return realistic demo data when no Excel file is present."""
    import random

    random.seed(42)

    periods = [
        "01.04.2026 - 15.04.2026",
        "16.04.2026 - 30.04.2026",
        "01.05.2026 - 15.05.2026",
        "16.05.2026 - 31.05.2026",
        "01.06.2026 - 15.06.2026",
    ]
    emp_counts = {
        "BANGLADESHI": [700, 720, 750, 730, 760],
        "INDIAN":      [450, 470, 500, 480, 510],
        "MALAGASY":    [130, 140, 130, 135, 130],
        "SRILANKAN":   [165, 310, 100, 290, 170],
    }

    items_cfg = {
        "Bangladeshi": [
            ("Rice",       "kg", 80,  120),
            ("Chicken",    "kg", 250, 400),
            ("Vegetables", "kg", 40,  80),
            ("Cooking Oil","L",  180, 250),
            ("Dal",        "kg", 90,  130),
            ("Bread",      "pc", 5,   15),
            ("Eggs",       "pc", 8,   12),
        ],
        "Indian": [
            ("Basmati Rice", "kg", 120, 180),
            ("Mutton",       "kg", 600, 900),
            ("Spices",       "kg", 200, 400),
            ("Cooking Oil",  "L",  180, 250),
            ("Flour",        "kg", 60,  100),
            ("Sugar",        "kg", 70,  110),
            ("Tea",          "kg", 300, 500),
        ],
        "Malagasy": [
            ("Rice",      "kg", 80,  120),
            ("Fish",      "kg", 200, 350),
            ("Cassava",   "kg", 40,  80),
            ("Palm Oil",  "L",  160, 230),
            ("Beans",     "kg", 70,  110),
        ],
        "Srilankan": [
            ("Rice",         "kg", 80,  120),
            ("Coconut",      "pc", 30,  50),
            ("Curry Leaves", "kg", 80,  150),
            ("Coconut Oil",  "L",  220, 320),
            ("Dhal",         "kg", 90,  140),
            ("Chilli",       "kg", 200, 400),
        ],
    }

    data: dict = {}
    for nat, item_list in items_cfg.items():
        rows = []
        for i, period in enumerate(periods):
            for item, unit, lo, hi in item_list:
                qty = round(random.uniform(30, 200), 1)
                up  = round(random.uniform(lo, hi), 2)
                rows.append(
                    {
                        "PERIOD":     period,
                        "ITEMS":      item,
                        "QTY":        qty,
                        "UNIT":       unit,
                        "UNIT PRICE": up,
                        "TOTAL":      round(qty * up, 2),
                    }
                )
        data[nat] = pd.DataFrame(rows)

    data["EMPLOYEES"] = pd.DataFrame(
        {
            "PERIOD":       periods,
            "BANGLADESHI":  emp_counts["BANGLADESHI"],
            "INDIAN":       emp_counts["INDIAN"],
            "MALAGASY":     emp_counts["MALAGASY"],
            "SRILANKAN":    emp_counts["SRILANKAN"],
        }
    )
    return data


# ── period utilities ───────────────────────────────────────────────────────────

def _all_periods(data: dict) -> list[str]:
    """Return every unique period from expense + employee sheets, sorted chronologically."""
    seen: set[str] = set()
    for nat in NAT_SHEETS:
        df = data.get(nat, pd.DataFrame())
        if "PERIOD" in df.columns:
            seen.update(df["PERIOD"].dropna().astype(str).tolist())
    emp = data.get("EMPLOYEES", pd.DataFrame())
    if "PERIOD" in emp.columns:
        seen.update(emp["PERIOD"].dropna().astype(str).tolist())
    return sorted([p for p in seen if p.strip()], key=_period_key)


def _filter(data: dict, from_p: str, to_p: str) -> dict:
    """Slice data to only the requested period range."""
    all_p = _all_periods(data)
    if not all_p:
        return data
    fp = all_p.index(from_p) if from_p in all_p else 0
    tp = all_p.index(to_p)   if to_p   in all_p else len(all_p) - 1
    keep = set(all_p[fp : tp + 1])
    out: dict = {}
    for k, df in data.items():
        if "PERIOD" in df.columns:
            out[k] = df[df["PERIOD"].isin(keep)].copy()
        else:
            out[k] = df.copy()
    return out


# ── helpers ────────────────────────────────────────────────────────────────────

def _expense_by_period(data: dict, periods: list[str]) -> dict[str, list[float]]:
    result: dict[str, list[float]] = {nat: [] for nat in NAT_SHEETS}
    for p in periods:
        for nat in NAT_SHEETS:
            df = data.get(nat, pd.DataFrame())
            val = 0.0
            if not df.empty and "TOTAL" in df.columns and "PERIOD" in df.columns:
                val = float(df.loc[df["PERIOD"] == p, "TOTAL"].sum())
            result[nat].append(val)
    return result


def _emp_by_period(data: dict, periods: list[str]) -> dict[str, list[int]]:
    emp_df = data.get("EMPLOYEES", pd.DataFrame())
    result: dict[str, list[int]] = {nat: [] for nat in NAT_SHEETS}
    for p in periods:
        for nat in NAT_SHEETS:
            col = EMP_COLS[nat]
            val = 0
            if not emp_df.empty and "PERIOD" in emp_df.columns and col in emp_df.columns:
                rows = emp_df.loc[emp_df["PERIOD"] == p, col]
                val = int(rows.sum()) if not rows.empty else 0
            result[nat].append(val)
    return result


def _labeled(d: dict[str, list]) -> dict[str, list]:
    """Re-key from internal nat names to display labels."""
    return {NAT_LABELS[n]: v for n, v in d.items()}


def _colors_labeled() -> dict[str, str]:
    return {NAT_LABELS[n]: NAT_COLORS[n] for n in NAT_SHEETS}


# ── routes – pages ─────────────────────────────────────────────────────────────

@app.route("/")
def dashboard():
    return render_template("dashboard.html", periods=_all_periods(_load()))


@app.route("/employees")
def employees():
    return render_template("employees.html", periods=_all_periods(_load()))


@app.route("/trend")
def trend():
    return render_template("trend.html", periods=_all_periods(_load()))


@app.route("/comparison")
def comparison():
    return render_template("comparison.html", periods=_all_periods(_load()))


@app.route("/reports")
def reports():
    return render_template("reports.html", periods=_all_periods(_load()))


# ── routes – API ───────────────────────────────────────────────────────────────

@app.route("/api/dashboard")
def api_dashboard():
    data = _load()
    fp = request.args.get("from_period", "")
    tp = request.args.get("to_period",   "")
    if fp or tp:
        data = _filter(data, fp, tp)

    periods  = _all_periods(data)
    emp_df   = data.get("EMPLOYEES", pd.DataFrame())
    exp_by_p = _expense_by_period(data, periods)
    emp_by_p = _emp_by_period(data, periods)

    # Totals
    totals = {
        nat: float(data[nat]["TOTAL"].sum())
        if not data.get(nat, pd.DataFrame()).empty and "TOTAL" in data[nat].columns
        else 0.0
        for nat in NAT_SHEETS
    }
    grand = sum(totals.values())

    # Latest employee count (last period chronologically)
    latest_emp: dict[str, int] = {nat: 0 for nat in NAT_SHEETS}
    if not emp_df.empty and "PERIOD" in emp_df.columns:
        last_p = sorted(emp_df["PERIOD"].unique(), key=_period_key)[-1]
        row = emp_df.loc[emp_df["PERIOD"] == last_p].iloc[0]
        for nat in NAT_SHEETS:
            col = EMP_COLS[nat]
            latest_emp[nat] = int(row[col]) if col in row.index else 0

    # Average CPH across all displayed periods
    cph: dict[str, float] = {}
    for nat in NAT_SHEETS:
        col = EMP_COLS[nat]
        if not emp_df.empty and col in emp_df.columns:
            avg_emp = emp_df[col].mean()
            cph[nat] = totals[nat] / avg_emp if avg_emp > 0 else 0.0
        else:
            cph[nat] = 0.0

    return jsonify(
        {
            "periods":          periods,
            "grand_total":      round(grand, 2),
            "totals":           {NAT_LABELS[n]: round(totals[n], 2)   for n in NAT_SHEETS},
            "cph":              {NAT_LABELS[n]: round(cph[n],    2)   for n in NAT_SHEETS},
            "latest_employees": {NAT_LABELS[n]: latest_emp[n]         for n in NAT_SHEETS},
            "total_employees":  sum(latest_emp.values()),
            "expense_trend":    _labeled(exp_by_p),
            "emp_trend":        _labeled(emp_by_p),
            "colors":           _colors_labeled(),
        }
    )


@app.route("/api/employees")
def api_employees():
    data = _load()
    fp = request.args.get("from_period", "")
    tp = request.args.get("to_period",   "")
    if fp or tp:
        data = _filter(data, fp, tp)

    periods  = _all_periods(data)
    emp_by_p = _emp_by_period(data, periods)

    table_rows = []
    for i, p in enumerate(periods):
        row = {"period": p}
        total = 0
        for nat in NAT_SHEETS:
            v = emp_by_p[nat][i]
            row[NAT_LABELS[nat]] = v
            total += v
        row["Total"] = total
        table_rows.append(row)

    return jsonify(
        {
            "periods": periods,
            "table":   table_rows,
            "chart":   _labeled(emp_by_p),
            "colors":  _colors_labeled(),
        }
    )


@app.route("/api/trend")
def api_trend():
    data = _load()
    fp = request.args.get("from_period", "")
    tp = request.args.get("to_period",   "")
    if fp or tp:
        data = _filter(data, fp, tp)

    periods  = _all_periods(data)
    exp_by_p = _expense_by_period(data, periods)
    emp_by_p = _emp_by_period(data, periods)

    # CPH per period
    cph_ft: dict[str, list[float]] = {nat: [] for nat in NAT_SHEETS}
    for i, _ in enumerate(periods):
        for nat in NAT_SHEETS:
            exp = exp_by_p[nat][i]
            emp = emp_by_p[nat][i]
            cph_ft[nat].append(round(exp / emp, 2) if emp > 0 else 0.0)

    # Monthly aggregation – order by first occurrence
    month_order: list[str] = []
    month_exp: dict[str, dict[str, float]] = defaultdict(lambda: {n: 0.0 for n in NAT_SHEETS})
    month_emp: dict[str, dict[str, int]]   = defaultdict(lambda: {n: 0   for n in NAT_SHEETS})

    for i, p in enumerate(periods):
        ml = _month_label(p)
        if ml not in month_order:
            month_order.append(ml)
        for nat in NAT_SHEETS:
            month_exp[ml][nat] += exp_by_p[nat][i]
            month_emp[ml][nat] += emp_by_p[nat][i]

    monthly_exp: dict[str, list[float]] = {nat: [month_exp[m][nat] for m in month_order] for nat in NAT_SHEETS}
    monthly_cph: dict[str, list[float]] = {
        nat: [
            round(month_exp[m][nat] / month_emp[m][nat], 2) if month_emp[m][nat] > 0 else 0.0
            for m in month_order
        ]
        for nat in NAT_SHEETS
    }

    return jsonify(
        {
            "fortnightly": {
                "periods": periods,
                "expense": _labeled(exp_by_p),
                "cph":     _labeled(cph_ft),
            },
            "monthly": {
                "periods": month_order,
                "expense": _labeled(monthly_exp),
                "cph":     _labeled(monthly_cph),
            },
            "colors": _colors_labeled(),
        }
    )


@app.route("/api/comparison")
def api_comparison():
    data = _load()
    fp  = request.args.get("from_period", "")
    tp  = request.args.get("to_period",   "")
    nat_filter = request.args.get("nationality", "")  # internal sheet name or ""

    if fp or tp:
        data = _filter(data, fp, tp)

    emp_df  = data.get("EMPLOYEES", pd.DataFrame())
    nat_list = [nat_filter] if nat_filter in NAT_SHEETS else NAT_SHEETS

    result: dict[str, list] = {}
    for nat in nat_list:
        df = data.get(nat, pd.DataFrame())
        if df.empty or "ITEMS" not in df.columns:
            result[NAT_LABELS[nat]] = []
            continue

        col = EMP_COLS[nat]
        total_emp = int(emp_df[col].sum()) if not emp_df.empty and col in emp_df.columns else 0

        agg_dict: dict = {"QTY": "sum", "TOTAL": "sum"}
        if "UNIT" in df.columns:
            agg_dict["UNIT"] = "first"
        if "UNIT PRICE" in df.columns:
            agg_dict["UNIT PRICE"] = "mean"

        grouped = df.groupby("ITEMS").agg(agg_dict).reset_index()

        items_data = []
        for _, row in grouped.iterrows():
            qty        = float(row.get("QTY",   0) or 0)
            total      = float(row.get("TOTAL", 0) or 0)
            unit       = str(row.get("UNIT", ""))
            unit_price = float(row.get("UNIT PRICE", 0) or 0)
            items_data.append(
                {
                    "item":          str(row["ITEMS"]),
                    "qty":           round(qty,        2),
                    "unit":          unit,
                    "unit_price":    round(unit_price, 2),
                    "total":         round(total,      2),
                    "employees":     total_emp,
                    "qty_per_head":  round(qty   / total_emp, 4) if total_emp else 0,
                    "cost_per_head": round(total / total_emp, 2) if total_emp else 0,
                }
            )
        items_data.sort(key=lambda x: x["total"], reverse=True)
        result[NAT_LABELS[nat]] = items_data

    return jsonify(
        {
            "data":          result,
            "nationalities": [NAT_LABELS[n] for n in NAT_SHEETS],
            "nat_keys":      NAT_SHEETS,
            "colors":        _colors_labeled(),
        }
    )


@app.route("/api/reports")
def api_reports():
    data = _load()
    fp = request.args.get("from_period", "")
    tp = request.args.get("to_period",   "")
    if fp or tp:
        data = _filter(data, fp, tp)

    periods  = _all_periods(data)
    exp_by_p = _expense_by_period(data, periods)
    emp_by_p = _emp_by_period(data, periods)
    emp_df   = data.get("EMPLOYEES", pd.DataFrame())

    # Expense summary table
    totals_by_nat = {nat: 0.0 for nat in NAT_SHEETS}
    exp_rows = []
    for i, p in enumerate(periods):
        row = {"period": p}
        pt  = 0.0
        for nat in NAT_SHEETS:
            v = exp_by_p[nat][i]
            row[NAT_LABELS[nat]] = round(v, 2)
            totals_by_nat[nat] += v
            pt += v
        row["Total"] = round(pt, 2)
        exp_rows.append(row)

    grand = sum(totals_by_nat.values())

    # Employee count table
    emp_rows = []
    for i, p in enumerate(periods):
        row = {"period": p}
        tot = 0
        for nat in NAT_SHEETS:
            v = emp_by_p[nat][i]
            row[NAT_LABELS[nat]] = v
            tot += v
        row["Total"] = tot
        emp_rows.append(row)

    # CPH table
    cph_rows = []
    cph_chart: dict[str, list[float]] = {nat: [] for nat in NAT_SHEETS}
    for i, p in enumerate(periods):
        row = {"period": p}
        for nat in NAT_SHEETS:
            exp = exp_by_p[nat][i]
            emp = emp_by_p[nat][i]
            v   = round(exp / emp, 2) if emp > 0 else 0.0
            row[NAT_LABELS[nat]] = v
            cph_chart[nat].append(v)
        cph_rows.append(row)

    return jsonify(
        {
            "periods": periods,
            "expense": {
                "rows":        exp_rows,
                "grand_total": {NAT_LABELS[n]: round(totals_by_nat[n], 2) for n in NAT_SHEETS},
                "grand_all":   round(grand, 2),
                "chart":       _labeled(exp_by_p),
            },
            "employees": {
                "rows":  emp_rows,
                "chart": _labeled(emp_by_p),
            },
            "cph": {
                "rows":  cph_rows,
                "chart": _labeled(cph_chart),
            },
            "colors":        _colors_labeled(),
            "nationalities": [NAT_LABELS[n] for n in NAT_SHEETS],
        }
    )


if __name__ == "__main__":
    debug = os.environ.get("FLASK_DEBUG", "0") == "1"
    app.run(debug=debug, host="0.0.0.0", port=5000)
