import os
import json
from flask import Flask, render_template, jsonify, request
import pandas as pd

app = Flask(__name__)

EXCEL_PATH = os.environ.get("EXCEL_PATH", "Canteen.xlsx")

NATIONALITIES = ["Bangladeshi", "Malagasy", "Indian", "Srilankan"]


def load_excel():
    """Load all sheets from the Excel workbook."""
    if not os.path.exists(EXCEL_PATH):
        return None, None
    try:
        xl = pd.ExcelFile(EXCEL_PATH)
        sheets = {}
        for nat in NATIONALITIES:
            if nat in xl.sheet_names:
                df = xl.parse(nat)
                df.columns = [str(c).strip().upper() for c in df.columns]
                df = df.dropna(how="all")
                if "TOTAL" in df.columns:
                    df["TOTAL"] = pd.to_numeric(df["TOTAL"], errors="coerce").fillna(0)
                if "QTY" in df.columns:
                    df["QTY"] = pd.to_numeric(df["QTY"], errors="coerce").fillna(0)
                if "UNIT PRICE" in df.columns:
                    df["UNIT PRICE"] = pd.to_numeric(df["UNIT PRICE"], errors="coerce").fillna(0)
                if "PERIOD" in df.columns:
                    df["PERIOD"] = df["PERIOD"].astype(str).str.strip()
                if "ITEMS" in df.columns:
                    df["ITEMS"] = df["ITEMS"].astype(str).str.strip()
                df["NATIONALITY"] = nat
                sheets[nat] = df

        emp_df = None
        if "EMPLOYEES" in xl.sheet_names:
            emp_df = xl.parse("EMPLOYEES")
            emp_df.columns = [str(c).strip().upper() for c in emp_df.columns]
            emp_df = emp_df.dropna(how="all")
            for col in ["BANGLADESHI", "INDIAN", "MALAGASY", "SRILANKAN"]:
                if col in emp_df.columns:
                    emp_df[col] = pd.to_numeric(emp_df[col], errors="coerce").fillna(0)

        return sheets, emp_df
    except Exception as e:
        print(f"Error loading Excel: {e}")
        return None, None


def get_combined(sheets):
    if not sheets:
        return pd.DataFrame()
    return pd.concat(sheets.values(), ignore_index=True)


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/api/summary")
def api_summary():
    sheets, emp_df = load_excel()
    if not sheets:
        return jsonify({"error": "Excel file not found or unreadable"}), 404

    combined = get_combined(sheets)
    summary = {}

    # Overall totals
    summary["grand_total"] = float(combined["TOTAL"].sum()) if "TOTAL" in combined.columns else 0
    summary["total_items"] = int(len(combined))

    # Per-nationality totals
    nat_totals = {}
    for nat in NATIONALITIES:
        if nat in sheets and "TOTAL" in sheets[nat].columns:
            nat_totals[nat] = float(sheets[nat]["TOTAL"].sum())
        else:
            nat_totals[nat] = 0
    summary["nat_totals"] = nat_totals

    # Per-period totals (all nationalities combined)
    if "PERIOD" in combined.columns and "TOTAL" in combined.columns:
        period_totals = (
            combined.groupby("PERIOD")["TOTAL"].sum().reset_index()
            .sort_values("PERIOD")
        )
        summary["period_totals"] = period_totals.to_dict(orient="records")
    else:
        summary["period_totals"] = []

    # Per-nationality per-period
    nat_period = {}
    for nat in NATIONALITIES:
        if nat in sheets and "PERIOD" in sheets[nat].columns and "TOTAL" in sheets[nat].columns:
            pt = (
                sheets[nat].groupby("PERIOD")["TOTAL"].sum().reset_index()
                .sort_values("PERIOD")
            )
            nat_period[nat] = pt.to_dict(orient="records")
        else:
            nat_period[nat] = []
    summary["nat_period"] = nat_period

    # Employee counts
    if emp_df is not None and "PERIOD" in emp_df.columns:
        summary["employees"] = emp_df.to_dict(orient="records")
    else:
        summary["employees"] = []

    # Top items by total spend
    if "ITEMS" in combined.columns and "TOTAL" in combined.columns:
        top_items = (
            combined.groupby("ITEMS")["TOTAL"].sum()
            .sort_values(ascending=False)
            .head(10)
            .reset_index()
        )
        summary["top_items"] = top_items.to_dict(orient="records")
    else:
        summary["top_items"] = []

    return jsonify(summary)


@app.route("/api/report")
def api_report():
    nationality = request.args.get("nationality", "All")
    period = request.args.get("period", "All")
    item_search = request.args.get("item", "").strip().lower()
    sort_by = request.args.get("sort", "TOTAL")
    sort_dir = request.args.get("dir", "desc")

    sheets, _ = load_excel()
    if not sheets:
        return jsonify({"error": "Excel file not found"}), 404

    combined = get_combined(sheets)

    if nationality != "All" and "NATIONALITY" in combined.columns:
        combined = combined[combined["NATIONALITY"] == nationality]

    if period != "All" and "PERIOD" in combined.columns:
        combined = combined[combined["PERIOD"] == period]

    if item_search and "ITEMS" in combined.columns:
        combined = combined[combined["ITEMS"].str.lower().str.contains(item_search, na=False)]

    # Sort
    valid_sorts = ["ITEMS", "TOTAL", "QTY", "UNIT PRICE", "PERIOD", "NATIONALITY", "UNIT"]
    if sort_by not in valid_sorts:
        sort_by = "TOTAL"
    ascending = sort_dir == "asc"
    if sort_by in combined.columns:
        combined = combined.sort_values(sort_by, ascending=ascending)

    # Select display columns
    display_cols = [c for c in ["PERIOD", "NATIONALITY", "ITEMS", "QTY", "UNIT", "UNIT PRICE", "TOTAL"] if c in combined.columns]
    result = combined[display_cols].fillna("").to_dict(orient="records")

    # Periods and nationalities for filter dropdowns
    sheets2, _ = load_excel()
    full = get_combined(sheets2)
    periods = sorted(full["PERIOD"].dropna().unique().tolist()) if "PERIOD" in full.columns else []
    nats = NATIONALITIES

    return jsonify({
        "rows": result,
        "total": float(combined["TOTAL"].sum()) if "TOTAL" in combined.columns else 0,
        "count": len(result),
        "periods": periods,
        "nationalities": nats,
    })


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)
