"""Flask application for Canteen Expenses Dashboard."""
import io
import logging

from flask import Flask, render_template, jsonify, request, send_file, Response

from helpers.data_loader import (
    get_periods,
    get_expense_periods,
    get_kpi_data,
    get_trend_data,
    get_comparison_data,
    get_employees,
    get_all_expenses,
    get_report_perhead,
    get_report_detailed,
    get_report_summary,
    get_report_monthly,
    get_nationality_sheets,
    get_item_comparison,
    NATIONALITY_SHEETS,
    format_rs,
)

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

app = Flask(__name__)
import os as _os
app.secret_key = _os.environ.get("SECRET_KEY", _os.urandom(32))


# ─── Helper ────────────────────────────────────────────────────────────────────

def _get_filter_args():
    period_from = request.args.get("from") or None
    period_to = request.args.get("to") or None
    nationality = request.args.get("nationality") or None
    return period_from, period_to, nationality


def _download_response(df, filename_base: str, fmt: str) -> Response:
    """Return a file download response for CSV or Excel."""
    if fmt == "excel":
        buf = io.BytesIO()
        with __import__("xlsxwriter").Workbook(buf) as wb:
            ws = wb.add_worksheet("Report")
            if not df.empty:
                for col_idx, col_name in enumerate(df.columns):
                    ws.write(0, col_idx, col_name)
                for row_idx, row in enumerate(df.itertuples(index=False), start=1):
                    for col_idx, val in enumerate(row):
                        ws.write(row_idx, col_idx, val if val is not None else "")
        buf.seek(0)
        return send_file(
            buf,
            as_attachment=True,
            download_name=f"{filename_base}.xlsx",
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    else:
        csv_data = df.to_csv(index=False)
        return Response(
            csv_data,
            mimetype="text/csv",
            headers={"Content-Disposition": f"attachment; filename={filename_base}.csv"},
        )


# ─── Page Routes ───────────────────────────────────────────────────────────────

@app.route("/")
def dashboard():
    periods = get_periods()
    # latest_period comes from expense sheets only — avoids defaulting to a future
    # period that has employee data but no expense data yet.
    expense_periods = get_expense_periods()
    latest = expense_periods[-1] if expense_periods else ""
    return render_template("dashboard.html", periods=periods, latest_period=latest)


@app.route("/employees")
def employees():
    periods = get_periods()
    return render_template("employees.html", periods=periods)


@app.route("/trends")
def trends():
    periods = get_periods()
    nationalities = get_nationality_sheets() or NATIONALITY_SHEETS
    return render_template("trends.html", periods=periods, nationalities=nationalities)


@app.route("/comparison")
def comparison():
    periods = get_periods()
    nationalities = get_nationality_sheets() or NATIONALITY_SHEETS
    return render_template("comparison.html", periods=periods, nationalities=nationalities)


@app.route("/reports")
def reports():
    periods = get_periods()
    nationalities = get_nationality_sheets() or NATIONALITY_SHEETS
    return render_template("reports.html", periods=periods, nationalities=nationalities)


# ─── JSON API ──────────────────────────────────────────────────────────────────

@app.route("/api/periods")
def api_periods():
    try:
        return jsonify(get_periods())
    except Exception as exc:
        logger.error("api_periods error: %s", exc)
        return jsonify([])


@app.route("/api/kpi")
def api_kpi():
    try:
        period_from, period_to, _ = _get_filter_args()
        return jsonify(get_kpi_data(period_from=period_from, period_to=period_to))
    except Exception as exc:
        logger.error("api_kpi error: %s", exc)
        return jsonify({"error": "An internal error occurred"})


@app.route("/api/charts/nationality")
def api_chart_nationality():
    """Employee by nationality donut data."""
    try:
        period_from, period_to, _ = _get_filter_args()
        kpi = get_kpi_data(period_from=period_from, period_to=period_to)
        nat_emp = kpi.get("nat_employees", {})
        labels = list(nat_emp.keys())
        data = list(nat_emp.values())
        return jsonify({
            "labels": labels,
            "datasets": [{
                "data": data,
                "backgroundColor": ["#3a86ff", "#06d6a0", "#ffd166", "#ef233c"],
                "borderColor": "#0d1b2a",
                "borderWidth": 2,
            }]
        })
    except Exception as exc:
        logger.error("api_chart_nationality error: %s", exc)
        return jsonify({"error": "An internal error occurred"})


@app.route("/api/charts/expenditure")
def api_chart_expenditure():
    """Expenditure by nationality bar chart data."""
    try:
        period_from, period_to, _ = _get_filter_args()
        kpi = get_kpi_data(period_from=period_from, period_to=period_to)
        nat_exp = kpi.get("nat_expenditure", {})
        labels = list(nat_exp.keys())
        data = list(nat_exp.values())
        return jsonify({
            "labels": labels,
            "datasets": [{
                "label": "Expenditure (Rs)",
                "data": data,
                "backgroundColor": ["#3a86ff", "#06d6a0", "#ffd166", "#ef233c"],
                "borderRadius": 6,
            }]
        })
    except Exception as exc:
        logger.error("api_chart_expenditure error: %s", exc)
        return jsonify({"error": "An internal error occurred"})


@app.route("/api/charts/trend")
def api_chart_trend():
    try:
        period_from, period_to, nationality = _get_filter_args()
        return jsonify(get_trend_data(nationality=nationality, period_from=period_from, period_to=period_to))
    except Exception as exc:
        logger.error("api_chart_trend error: %s", exc)
        return jsonify({"error": "An internal error occurred"})


@app.route("/api/employees")
def api_employees():
    try:
        period_from, period_to, _ = _get_filter_args()
        df = get_employees(period_from=period_from, period_to=period_to)
        if df.empty:
            return jsonify([])
        nat_cols = [c for c in ["BANGLADESHI", "INDIAN", "MALAGASY", "SRILANKAN"] if c in df.columns]
        df = df.copy()
        df["TOTAL"] = df[nat_cols].sum(axis=1)
        rows = []
        for _, row in df.iterrows():
            rows.append({
                "period": str(row.get("PERIOD", "")),
                "bangladeshi": int(row.get("BANGLADESHI", 0)),
                "indian": int(row.get("INDIAN", 0)),
                "malagasy": int(row.get("MALAGASY", 0)),
                "srilankan": int(row.get("SRILANKAN", 0)),
                "total": int(row.get("TOTAL", 0)),
            })
        return jsonify(rows)
    except Exception as exc:
        logger.error("api_employees error: %s", exc)
        return jsonify({"error": "An internal error occurred"})


@app.route("/api/comparison")
def api_comparison():
    try:
        period_from, period_to, _ = _get_filter_args()
        return jsonify(get_comparison_data(period_from=period_from, period_to=period_to))
    except Exception as exc:
        logger.error("api_comparison error: %s", exc)
        return jsonify({"error": "An internal error occurred"})


@app.route("/api/items/comparison")
def api_items_comparison():
    """Item-by-item consumption comparison across nationalities."""
    try:
        period_from, period_to, nationality = _get_filter_args()
        return jsonify(get_item_comparison(nationality=nationality, period_from=period_from, period_to=period_to))
    except Exception as exc:
        logger.error("api_items_comparison error: %s", exc)
        return jsonify({"error": "An internal error occurred"})


@app.route("/api/report/data/<report_type>")
def api_report_data(report_type: str):
    """Return report data as JSON for inline preview."""
    try:
        period_from, period_to, nationality = _get_filter_args()
        if report_type == "perhead":
            df = get_report_perhead(period_from=period_from, period_to=period_to, nationality=nationality)
        elif report_type == "detailed":
            df = get_report_detailed(period_from=period_from, period_to=period_to, nationality=nationality)
        elif report_type == "summary":
            df = get_report_summary(period_from=period_from, period_to=period_to)
        elif report_type == "monthly":
            df = get_report_monthly(period_from=period_from, period_to=period_to)
        else:
            return jsonify({"error": "Unknown report type"}), 404

        if df.empty:
            return jsonify({"columns": [], "rows": []})

        return jsonify({
            "columns": list(df.columns),
            "rows": df.fillna("").values.tolist(),
        })
    except Exception as exc:
        logger.error("api_report_data error: %s", exc)
        return jsonify({"error": "An internal error occurred"})


# ─── Report Downloads ──────────────────────────────────────────────────────────

@app.route("/reports/download/perhead")
def download_perhead():
    try:
        period_from, period_to, nationality = _get_filter_args()
        fmt = request.args.get("format", "csv")
        df = get_report_perhead(period_from=period_from, period_to=period_to, nationality=nationality)
        return _download_response(df, "perhead_consumption", fmt)
    except Exception as exc:
        logger.error("download_perhead error: %s", exc)
        return jsonify({"error": "An internal error occurred"}), 500


@app.route("/reports/download/detailed")
def download_detailed():
    try:
        period_from, period_to, nationality = _get_filter_args()
        fmt = request.args.get("format", "csv")
        df = get_report_detailed(period_from=period_from, period_to=period_to, nationality=nationality)
        return _download_response(df, "detailed_expenditure", fmt)
    except Exception as exc:
        logger.error("download_detailed error: %s", exc)
        return jsonify({"error": "An internal error occurred"}), 500


@app.route("/reports/download/summary")
def download_summary():
    try:
        period_from, period_to, _ = _get_filter_args()
        fmt = request.args.get("format", "csv")
        df = get_report_summary(period_from=period_from, period_to=period_to)
        return _download_response(df, "summary_report", fmt)
    except Exception as exc:
        logger.error("download_summary error: %s", exc)
        return jsonify({"error": "An internal error occurred"}), 500


@app.route("/reports/download/monthly")
def download_monthly():
    try:
        period_from, period_to, _ = _get_filter_args()
        fmt = request.args.get("format", "csv")
        df = get_report_monthly(period_from=period_from, period_to=period_to)
        return _download_response(df, "monthly_fortnight_report", fmt)
    except Exception as exc:
        logger.error("download_monthly error: %s", exc)
        return jsonify({"error": "An internal error occurred"}), 500


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=False)
