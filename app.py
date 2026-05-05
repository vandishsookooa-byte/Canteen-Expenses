import os
import io
import random

import pandas as pd
from flask import Flask, jsonify, render_template, request, send_file

app = Flask(__name__)

EXCEL_PATH = os.environ.get("EXCEL_PATH", "Canteen.xlsx")
NATIONALITIES = ["Bangladeshi", "Indian", "Malagasy", "Srilankan"]


# ---------------------------------------------------------------------------
# Sample / demo data (used when Canteen.xlsx is absent)
# ---------------------------------------------------------------------------

def _sample_data() -> dict:
    """Return realistic demo data matching the Excel schema."""
    rng = random.Random(42)
    periods = ["Jan-24", "Feb-24", "Mar-24", "Apr-24", "May-24", "Jun-24"]

    # EMPLOYEES sheet -------------------------------------------------------
    base_counts = {"BANGLADESHI": 45, "INDIAN": 30, "MALAGASY": 22, "SRILANKAN": 18}
    emp_rows = []
    for i, period in enumerate(periods):
        row = {"PERIOD": period}
        for col, base in base_counts.items():
            row[col] = base + (i % 3)
        emp_rows.append(row)
    employees_df = pd.DataFrame(emp_rows)

    # Nationality expense sheets -------------------------------------------
    items_catalog = [
        ("Rice",        "Kg",     200,  85),
        ("Bread",       "Piece",  500,  12),
        ("Vegetables",  "Kg",     150,  60),
        ("Chicken",     "Kg",     100, 380),
        ("Fish",        "Kg",      80, 280),
        ("Cooking Oil", "Ltr",     50, 220),
        ("Sugar",       "Kg",      40, 110),
        ("Tea",         "Box",     20, 350),
        ("Dal",         "Kg",      60, 180),
        ("Eggs",        "Dozen",   30, 200),
    ]
    multipliers = {"Bangladeshi": 1.0, "Indian": 0.9, "Malagasy": 0.7, "Srilankan": 0.6}

    expense_frames: dict[str, pd.DataFrame] = {}
    for nat in NATIONALITIES:
        mult = multipliers[nat]
        rows = []
        for period in periods:
            for item, unit, base_qty, base_price in items_catalog:
                qty = int(base_qty * mult * (0.9 + rng.random() * 0.2))
                price = int(base_price * (0.95 + rng.random() * 0.1))
                rows.append({
                    "PERIOD":     period,
                    "ITEMS":      item,
                    "QTY":        qty,
                    "UNIT":       unit,
                    "UNIT PRICE": price,
                    "TOTAL":      qty * price,
                })
        expense_frames[nat] = pd.DataFrame(rows)

    return {"employees": employees_df, "expenses": expense_frames}


def _load_data() -> dict:
    """Load from Excel if available, else fall back to sample data."""
    if os.path.exists(EXCEL_PATH):
        try:
            xl = pd.read_excel(EXCEL_PATH, sheet_name=None)
            employees = xl.get("EMPLOYEES", pd.DataFrame())
            expenses = {nat: xl.get(nat, pd.DataFrame()) for nat in NATIONALITIES}
            return {"employees": employees, "expenses": expenses}
        except Exception as exc:
            print(f"[WARN] Could not read {EXCEL_PATH}: {exc}")
    return _sample_data()


# ---------------------------------------------------------------------------
# Helper: serialise a DataFrame row to a plain dict with safe types
# ---------------------------------------------------------------------------

def _safe_row(row, mapping: dict) -> dict:
    """Extract fields from a row using (result_key, source_key, default) mapping."""
    out = {}
    for result_key, source_key, default in mapping:
        val = row.get(source_key, default)
        if isinstance(val, float) and val != val:      # NaN guard
            val = default
        if isinstance(default, int):
            val = int(val)
        elif isinstance(default, float):
            val = float(val)
        else:
            val = str(val) if val is not None else ""
        out[result_key] = val
    return out


# ---------------------------------------------------------------------------
# Routes – pages
# ---------------------------------------------------------------------------

@app.route("/")
def index():
    return render_template("index.html")


# ---------------------------------------------------------------------------
# Routes – API
# ---------------------------------------------------------------------------

@app.route("/api/employees")
def api_employees():
    data = _load_data()
    df = data["employees"].fillna(0)
    records = []
    for _, row in df.iterrows():
        bd  = int(row.get("BANGLADESHI", 0))
        ind = int(row.get("INDIAN",      0))
        mal = int(row.get("MALAGASY",    0))
        sri = int(row.get("SRILANKAN",   0))
        records.append({
            "period":      str(row.get("PERIOD", "")),
            "bangladeshi": bd,
            "indian":      ind,
            "malagasy":    mal,
            "srilankan":   sri,
            "total":       bd + ind + mal + sri,
        })
    return jsonify({"data": records})


@app.route("/api/expenses/<nationality>")
def api_expenses(nationality):
    if nationality not in NATIONALITIES:
        return jsonify({"error": "Invalid nationality"}), 400
    data = _load_data()
    df = data["expenses"].get(nationality, pd.DataFrame()).fillna(0)
    records = []
    for _, row in df.iterrows():
        records.append({
            "period":     str(row.get("PERIOD",     "")),
            "items":      str(row.get("ITEMS",      "")),
            "qty":        float(row.get("QTY",       0)),
            "unit":       str(row.get("UNIT",       "")),
            "unit_price": float(row.get("UNIT PRICE", 0)),
            "total":      float(row.get("TOTAL",     0)),
        })
    return jsonify({"data": records})


@app.route("/api/comparison")
def api_comparison():
    data = _load_data()
    period_map: dict[str, dict] = {}
    for nat in NATIONALITIES:
        df = data["expenses"].get(nat, pd.DataFrame()).fillna(0)
        for _, row in df.iterrows():
            period = str(row.get("PERIOD", ""))
            total  = float(row.get("TOTAL", 0))
            if period not in period_map:
                period_map[period] = {
                    "period":      period,
                    "bangladeshi": 0.0,
                    "indian":      0.0,
                    "malagasy":    0.0,
                    "srilankan":   0.0,
                }
            period_map[period][nat.lower()] += total

    records = list(period_map.values())
    for r in records:
        r["grand_total"] = r["bangladeshi"] + r["indian"] + r["malagasy"] + r["srilankan"]
    return jsonify({"data": records})


@app.route("/api/reports")
def api_reports():
    period_filter = request.args.get("period", "").strip()
    nat_filter    = request.args.get("nationality", "").strip()

    data = _load_data()

    # Employees -----------------------------------------------------------
    emp_df = data["employees"].fillna(0)
    employees = []
    for _, row in emp_df.iterrows():
        period = str(row.get("PERIOD", ""))
        if period_filter and period_filter.lower() not in period.lower():
            continue
        bd  = int(row.get("BANGLADESHI", 0))
        ind = int(row.get("INDIAN",      0))
        mal = int(row.get("MALAGASY",    0))
        sri = int(row.get("SRILANKAN",   0))
        employees.append({
            "period":      period,
            "bangladeshi": bd,
            "indian":      ind,
            "malagasy":    mal,
            "srilankan":   sri,
            "total":       bd + ind + mal + sri,
        })

    # Expenses ------------------------------------------------------------
    nats = ([nat_filter] if nat_filter and nat_filter in NATIONALITIES
            else NATIONALITIES)
    expenses: dict[str, list] = {}
    for nat in nats:
        df = data["expenses"].get(nat, pd.DataFrame()).fillna(0)
        rows = []
        for _, row in df.iterrows():
            period = str(row.get("PERIOD", ""))
            if period_filter and period_filter.lower() not in period.lower():
                continue
            rows.append({
                "period":     period,
                "items":      str(row.get("ITEMS",      "")),
                "qty":        float(row.get("QTY",       0)),
                "unit":       str(row.get("UNIT",       "")),
                "unit_price": float(row.get("UNIT PRICE", 0)),
                "total":      float(row.get("TOTAL",     0)),
            })
        expenses[nat] = rows

    grand_total      = sum(r["total"] for rows in expenses.values() for r in rows)
    total_employees  = sum(r["total"] for r in employees)
    all_periods      = sorted({r["period"] for r in employees})

    return jsonify({
        "employees": employees,
        "expenses":  expenses,
        "summary": {
            "grand_total":     grand_total,
            "total_employees": total_employees,
            "periods":         all_periods,
            "nationalities":   nats,
        },
    })


@app.route("/api/overview")
def api_overview():
    data = _load_data()

    # Total spend across all nationalities
    total_spend = 0.0
    for nat in NATIONALITIES:
        df = data["expenses"].get(nat, pd.DataFrame()).fillna(0)
        if not df.empty:
            total_spend += float(df["TOTAL"].sum())

    # Total employee-periods (sum of all employee counts across rows)
    emp_df = data["employees"].fillna(0)
    nat_cols = ["BANGLADESHI", "INDIAN", "MALAGASY", "SRILANKAN"]
    total_employees = int(emp_df[[c for c in nat_cols if c in emp_df.columns]].sum().sum())

    # Per-nationality totals
    nat_totals = {}
    for nat in NATIONALITIES:
        df = data["expenses"].get(nat, pd.DataFrame()).fillna(0)
        nat_totals[nat.lower()] = float(df["TOTAL"].sum()) if not df.empty else 0.0

    # Periods
    periods = list(emp_df["PERIOD"].astype(str).unique()) if not emp_df.empty else []

    return jsonify({
        "total_spend":     total_spend,
        "total_employees": total_employees,
        "nat_totals":      nat_totals,
        "periods":         periods,
        "data_source":     "Excel" if os.path.exists(EXCEL_PATH) else "Demo",
    })


# ---------------------------------------------------------------------------
# Routes – downloads
# ---------------------------------------------------------------------------

def _build_df_filter(df: pd.DataFrame, period_filter: str) -> pd.DataFrame:
    if period_filter and not df.empty:
        mask = df["PERIOD"].astype(str).str.lower().str.contains(
            period_filter.lower(), na=False
        )
        df = df[mask]
    return df


@app.route("/download/csv")
def download_csv():
    period_filter = request.args.get("period", "").strip()
    nat_filter    = request.args.get("nationality", "").strip()

    data   = _load_data()
    output = io.StringIO()

    # Employees
    emp_df = _build_df_filter(data["employees"].fillna(0), period_filter)
    output.write("EMPLOYEE COUNT BY PERIOD\n")
    emp_df.to_csv(output, index=False)
    output.write("\n")

    # Expenses
    nats = ([nat_filter] if nat_filter and nat_filter in NATIONALITIES
            else NATIONALITIES)
    for nat in nats:
        df = _build_df_filter(
            data["expenses"].get(nat, pd.DataFrame()).fillna(0), period_filter
        )
        output.write(f"{nat.upper()} EXPENSES\n")
        df.to_csv(output, index=False)
        output.write("\n")

    output.seek(0)
    return send_file(
        io.BytesIO(output.read().encode("utf-8-sig")),
        mimetype="text/csv",
        as_attachment=True,
        download_name="canteen_report.csv",
    )


@app.route("/download/excel")
def download_excel():
    period_filter = request.args.get("period", "").strip()
    nat_filter    = request.args.get("nationality", "").strip()

    data   = _load_data()
    output = io.BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        emp_df = _build_df_filter(data["employees"].fillna(0), period_filter)
        emp_df.to_excel(writer, sheet_name="Employees", index=False)

        nats = ([nat_filter] if nat_filter and nat_filter in NATIONALITIES
                else NATIONALITIES)
        for nat in nats:
            df = _build_df_filter(
                data["expenses"].get(nat, pd.DataFrame()).fillna(0), period_filter
            )
            df.to_excel(writer, sheet_name=nat[:31], index=False)

    output.seek(0)
    return send_file(
        output,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name="canteen_report.xlsx",
    )


if __name__ == "__main__":
    app.run(debug=True, host="0.0.0.0", port=5000)
