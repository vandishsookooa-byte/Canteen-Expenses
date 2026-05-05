import os
from datetime import datetime

import pandas as pd
from flask import Flask, render_template, request

app = Flask(__name__)


# ---------------------------------------------------------------------------
# Jinja2 custom filters
# ---------------------------------------------------------------------------

def _format_num(value) -> str:
    """Format a numeric value as a 2-decimal-place string with thousands separator."""
    try:
        return f"{float(value):,.2f}"
    except (TypeError, ValueError):
        return str(value)


def _format_qty(value) -> str:
    """Format a quantity: show 2 decimals only when fractional, else integer."""
    try:
        f = float(value)
        return f"{f:,.2f}" if f != int(f) else f"{int(f):,}"
    except (TypeError, ValueError):
        return str(value)


app.jinja_env.filters["format_num"] = _format_num
app.jinja_env.filters["format_qty"] = _format_qty


# ---------------------------------------------------------------------------
# Context processor – injects 'periods' and 'now' into every template
# ---------------------------------------------------------------------------

@app.context_processor
def inject_globals():
    sheets = _load_excel()
    return {
        "periods": _all_periods(sheets),
        "now": datetime.utcnow(),
    }

EXCEL_PATH = os.environ.get("EXCEL_PATH", "Canteen.xlsx")

NAT_SHEETS = ["Bangladeshi", "Indian", "Malagasy", "Srilankan"]

NAT_COLOURS = {
    "Bangladeshi": "success",
    "Indian":      "warning",
    "Malagasy":    "info",
    "Srilankan":   "danger",
}

EMP_COLS = {
    "Bangladeshi": "BANGLADESHI",
    "Indian":      "INDIAN",
    "Malagasy":    "MALAGASY",
    "Srilankan":   "SRILANKAN",
}

# ---------------------------------------------------------------------------
# Data loading helpers
# ---------------------------------------------------------------------------

def _load_excel() -> dict[str, pd.DataFrame]:
    """Return a dict of {sheet_name: DataFrame} for all sheets."""
    try:
        xf = pd.ExcelFile(EXCEL_PATH)
        return {s: xf.parse(s) for s in xf.sheet_names}
    except FileNotFoundError:
        return {}


def _nat_df(sheets: dict, nat: str) -> pd.DataFrame:
    """Return cleaned nationality DataFrame or empty frame."""
    df = sheets.get(nat, pd.DataFrame())
    if df.empty:
        return df
    df.columns = [str(c).strip().upper() for c in df.columns]
    df = df.dropna(subset=["PERIOD"])
    df["PERIOD"] = df["PERIOD"].astype(str).str.strip()
    for col in ["QTY", "UNIT PRICE", "TOTAL"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
    if "ITEMS" in df.columns:
        df["ITEMS"] = df["ITEMS"].astype(str).str.strip()
    return df


def _emp_df(sheets: dict) -> pd.DataFrame:
    """Return cleaned EMPLOYEES DataFrame."""
    df = sheets.get("EMPLOYEES", pd.DataFrame())
    if df.empty:
        return df
    df.columns = [str(c).strip().upper() for c in df.columns]
    df = df.dropna(subset=["PERIOD"])
    df["PERIOD"] = df["PERIOD"].astype(str).str.strip()
    for col in EMP_COLS.values():
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
    return df


def _parse_start(period: str) -> datetime:
    """Parse 'DD.MM.YYYY - DD.MM.YYYY' and return start date for sorting."""
    try:
        return datetime.strptime(period.split(" - ")[0].strip(), "%d.%m.%Y")
    except Exception:
        return datetime.min


def _all_periods(sheets: dict) -> list[str]:
    """Return all distinct periods across all nationality sheets, sorted chronologically."""
    periods: set[str] = set()
    for nat in NAT_SHEETS:
        df = _nat_df(sheets, nat)
        if not df.empty and "PERIOD" in df.columns:
            periods.update(df["PERIOD"].unique())
    return sorted(periods, key=_parse_start)


# ---------------------------------------------------------------------------
# Routes
# ---------------------------------------------------------------------------

@app.route("/")
def index():
    sheets = _load_excel()
    periods = _all_periods(sheets)

    # Summary cards: total spend per nationality (all periods)
    nat_totals: dict[str, float] = {}
    for nat in NAT_SHEETS:
        df = _nat_df(sheets, nat)
        nat_totals[nat] = float(df["TOTAL"].sum()) if not df.empty and "TOTAL" in df.columns else 0.0

    grand_total = sum(nat_totals.values())

    return render_template(
        "index.html",
        periods=periods,
        nat_totals=nat_totals,
        grand_total=grand_total,
        nat_colours=NAT_COLOURS,
        excel_missing=(not sheets),
    )


@app.route("/period/<path:period>")
def period_detail(period: str):
    sheets = _load_excel()
    all_periods = _all_periods(sheets)

    emp_df = _emp_df(sheets)
    emp_row: dict[str, int] = {}
    if not emp_df.empty:
        row = emp_df[emp_df["PERIOD"] == period]
        if not row.empty:
            emp_row = {nat: int(row.iloc[0].get(col, 0)) for nat, col in EMP_COLS.items()}

    nat_data: dict[str, list[dict]] = {}
    nat_subtotals: dict[str, float] = {}

    for nat in NAT_SHEETS:
        df = _nat_df(sheets, nat)
        if df.empty:
            nat_data[nat] = []
            nat_subtotals[nat] = 0.0
            continue
        sub = df[df["PERIOD"] == period].copy()
        nat_subtotals[nat] = float(sub["TOTAL"].sum()) if "TOTAL" in sub.columns else 0.0
        nat_data[nat] = sub.to_dict("records")

    grand_total = sum(nat_subtotals.values())

    return render_template(
        "period.html",
        period=period,
        all_periods=all_periods,
        nat_data=nat_data,
        nat_subtotals=nat_subtotals,
        grand_total=grand_total,
        emp_row=emp_row,
        nat_colours=NAT_COLOURS,
    )


@app.route("/report/item")
def report_item():
    sheets = _load_excel()
    all_periods = _all_periods(sheets)

    # Optional period filter
    selected_period = request.args.get("period", "")

    # Aggregate across all nationalities
    frames = []
    for nat in NAT_SHEETS:
        df = _nat_df(sheets, nat)
        if df.empty:
            continue
        if selected_period:
            df = df[df["PERIOD"] == selected_period]
        df = df.copy()
        df["NATIONALITY"] = nat
        frames.append(df)

    if not frames:
        items_agg = pd.DataFrame()
        item_rows = []
        grand_total = 0.0
    else:
        combined = pd.concat(frames, ignore_index=True)
        # Per-item totals across all nationalities
        items_agg = (
            combined.groupby("ITEMS", as_index=False)
            .agg(
                TOTAL_QTY=("QTY", "sum"),
                TOTAL_SPEND=("TOTAL", "sum"),
            )
            .sort_values("TOTAL_SPEND", ascending=False)
        )
        item_rows = items_agg.to_dict("records")
        grand_total = float(items_agg["TOTAL_SPEND"].sum())

    return render_template(
        "report_item.html",
        item_rows=item_rows,
        grand_total=grand_total,
        all_periods=all_periods,
        selected_period=selected_period,
    )


@app.route("/report/nationality")
def report_nationality():
    sheets = _load_excel()
    all_periods = _all_periods(sheets)

    selected_period = request.args.get("period", "")

    emp_df = _emp_df(sheets)

    nat_rows = []
    for nat in NAT_SHEETS:
        df = _nat_df(sheets, nat)
        if not df.empty and selected_period:
            df = df[df["PERIOD"] == selected_period]

        total_spend = float(df["TOTAL"].sum()) if not df.empty and "TOTAL" in df.columns else 0.0
        total_qty = float(df["QTY"].sum()) if not df.empty and "QTY" in df.columns else 0.0
        items_count = int(df["ITEMS"].nunique()) if not df.empty and "ITEMS" in df.columns else 0

        # Employee count: sum across filtered periods or all
        emp_count = 0
        if not emp_df.empty:
            col = EMP_COLS[nat]
            if col in emp_df.columns:
                if selected_period:
                    emp_sub = emp_df[emp_df["PERIOD"] == selected_period]
                else:
                    emp_sub = emp_df
                emp_count = int(emp_sub[col].sum())

        cost_per_head = (total_spend / emp_count) if emp_count > 0 else 0.0

        # Top 5 items for this nationality
        top_items: list[dict] = []
        if not df.empty and "ITEMS" in df.columns and "TOTAL" in df.columns:
            top = (
                df.groupby("ITEMS", as_index=False)["TOTAL"]
                .sum()
                .sort_values("TOTAL", ascending=False)
                .head(5)
            )
            top_items = top.to_dict("records")

        nat_rows.append(
            {
                "nationality": nat,
                "total_spend": total_spend,
                "total_qty": total_qty,
                "items_count": items_count,
                "emp_count": emp_count,
                "cost_per_head": cost_per_head,
                "top_items": top_items,
                "colour": NAT_COLOURS[nat],
            }
        )

    grand_total = sum(r["total_spend"] for r in nat_rows)

    return render_template(
        "report_nationality.html",
        nat_rows=nat_rows,
        grand_total=grand_total,
        all_periods=all_periods,
        selected_period=selected_period,
    )


# ---------------------------------------------------------------------------

if __name__ == "__main__":
    debug = os.environ.get("FLASK_DEBUG", "0") == "1"
    app.run(host="0.0.0.0", port=5000, debug=debug)
