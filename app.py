import os
import io
import csv
import json
import re
from datetime import datetime

import pandas as pd
from flask import Flask, render_template, request, jsonify, Response

# ---------------------------------------------------------------------------
# App setup
# ---------------------------------------------------------------------------
app = Flask(__name__)

EXCEL_PATH = os.environ.get('EXCEL_PATH', 'Canteen.xlsx')
NAT_SHEETS = ['Bangladeshi', 'Indian', 'Malagasy', 'Srilankan']
EMP_COLS = {
    'Bangladeshi': 'BANGLADESHI',
    'Indian':      'INDIAN',
    'Malagasy':    'MALAGASY',
    'Srilankan':   'SRILANKAN',
}
NAT_COLORS = {
    'Bangladeshi': '#4fc3f7',
    'Indian':      '#ff8a65',
    'Malagasy':    '#81c784',
    'Srilankan':   '#ce93d8',
}

# ---------------------------------------------------------------------------
# Jinja2 filter – Rs currency formatting
# ---------------------------------------------------------------------------
@app.template_filter('rs')
def rs_format(value):
    """Format a number as Rs with comma grouping (e.g. Rs 1,000)."""
    try:
        v = float(value)
        if v == int(v):
            return 'Rs {:,.0f}'.format(v)
        return 'Rs {:,.2f}'.format(v)
    except (TypeError, ValueError):
        return 'Rs 0'


# ---------------------------------------------------------------------------
# Excel helpers
# ---------------------------------------------------------------------------

# Alternative column names used for the period column across sheets
_PERIOD_COL_CANDIDATES = ['PERIOD', 'SN', 'DATE', 'MONTH', 'WEEK', 'FORTNIGHTLY']

# Regex pattern that matches "DD.MM.YYYY" at the start of a cell value
_DATE_RANGE_RE = re.compile(r'^\d{2}\.\d{2}\.\d{4}')


def _normalise_period(value: str) -> str:
    """Ensure a period string has spaces around the dash separator.

    Handles both "DD.MM.YYYY - DD.MM.YYYY" and "DD.MM.YYYY-DD.MM.YYYY".
    """
    # Replace dash(es) surrounded by optional whitespace with ' - '
    return re.sub(r'\s*-\s*', ' - ', value.strip())


def _detect_period_col(df: pd.DataFrame) -> str | None:
    """Return the name of the column that holds period/date-range strings.

    Tries known candidate names first, then falls back to scanning every
    column's values for a DD.MM.YYYY date-range pattern.
    """
    cols = list(df.columns)
    # 1. Try well-known names
    for candidate in _PERIOD_COL_CANDIDATES:
        if candidate in cols:
            return candidate
    # 2. Scan columns for date-range-like content
    for col in cols:
        sample = df[col].dropna().astype(str).head(10)
        if sample.str.match(r'^\d{2}\.\d{2}\.\d{4}').any():
            return col
    return None


def _parse_start(period: str) -> datetime:
    """Return datetime for the start of a period string 'DD.MM.YYYY - DD.MM.YYYY'."""
    try:
        return datetime.strptime(period.split(' - ')[0].strip(), '%d.%m.%Y')
    except Exception:
        return datetime.min


def _load_data() -> dict:
    """Load all relevant sheets from the Excel file."""
    data: dict = {}
    if not os.path.exists(EXCEL_PATH):
        return data
    try:
        xl = pd.ExcelFile(EXCEL_PATH)
        for sheet in NAT_SHEETS:
            if sheet in xl.sheet_names:
                df = xl.parse(sheet)
                df.columns = [str(c).strip().upper() for c in df.columns]
                # Find whichever column holds the period, rename it to PERIOD
                period_col = _detect_period_col(df)
                if period_col and period_col != 'PERIOD':
                    df = df.rename(columns={period_col: 'PERIOD'})
                # Normalise PERIOD values (strip whitespace, uniform dash spacing)
                if 'PERIOD' in df.columns:
                    df['PERIOD'] = (
                        df['PERIOD']
                        .astype(str)
                        .str.strip()
                        .apply(lambda v: _normalise_period(v) if _DATE_RANGE_RE.match(v) else v)
                    )
                data[sheet] = df
        if 'EMPLOYEES' in xl.sheet_names:
            emp = xl.parse('EMPLOYEES')
            emp.columns = [str(c).strip().upper() for c in emp.columns]
            period_col = _detect_period_col(emp)
            if period_col and period_col != 'PERIOD':
                emp = emp.rename(columns={period_col: 'PERIOD'})
            if 'PERIOD' in emp.columns:
                emp['PERIOD'] = (
                    emp['PERIOD']
                    .astype(str)
                    .str.strip()
                    .apply(lambda v: _normalise_period(v) if _DATE_RANGE_RE.match(v) else v)
                )
            data['EMPLOYEES'] = emp
    except Exception as exc:
        print(f'[ERROR] loading Excel: {exc}')
    return data


def _all_periods(data: dict) -> list:
    """Return all unique periods sorted chronologically."""
    periods: set = set()
    for sheet in NAT_SHEETS:
        if sheet in data:
            df = data[sheet]
            if 'PERIOD' in df.columns:
                periods.update(df['PERIOD'].dropna().unique())
    return sorted(periods, key=_parse_start)


def _employee_counts(data: dict, period: str) -> dict:
    """Return {nationality: employee_count} for a given period."""
    counts = {n: 0 for n in NAT_SHEETS}
    if 'EMPLOYEES' not in data:
        return counts
    emp = data['EMPLOYEES']
    row = emp[emp['PERIOD'] == period]
    if row.empty:
        return counts
    row = row.iloc[0]
    for nat in NAT_SHEETS:
        col = EMP_COLS[nat]
        if col in emp.columns:
            try:
                counts[nat] = int(row[col]) if not pd.isna(row[col]) else 0
            except (ValueError, TypeError):
                counts[nat] = 0
    return counts


def _period_summary(data: dict, period: str) -> dict:
    """Return per-nationality expense totals and items for a given period."""
    summary = {}
    for nat in NAT_SHEETS:
        if nat not in data:
            summary[nat] = {'total': 0, 'item_list': []}
            continue
        df = data[nat]
        pf = df[df['PERIOD'] == period] if 'PERIOD' in df.columns else pd.DataFrame()
        total = float(pf['TOTAL'].sum()) if 'TOTAL' in pf.columns else 0.0
        item_list = []
        if not pf.empty:
            cols_needed = ['ITEMS', 'QTY', 'UNIT', 'UNIT PRICE', 'TOTAL']
            available = [c for c in cols_needed if c in pf.columns]
            for _, row in pf[available].iterrows():
                item = {}
                for c in available:
                    v = row[c]
                    if pd.isna(v):
                        item[c.lower().replace(' ', '_')] = None
                    elif hasattr(v, 'item'):
                        item[c.lower().replace(' ', '_')] = v.item()
                    else:
                        item[c.lower().replace(' ', '_')] = v
                item_list.append(item)
        summary[nat] = {'total': total, 'item_list': item_list}
    return summary


def _build_dashboard(data: dict, period: str) -> dict:
    """Assemble all dashboard data for a given period."""
    summary = _period_summary(data, period)
    emp_counts = _employee_counts(data, period)

    grand_total = sum(v['total'] for v in summary.values())
    total_employees = sum(emp_counts.values())

    nat_data = []
    for nat in NAT_SHEETS:
        total = summary[nat]['total']
        emp = emp_counts[nat]
        per_head = round(total / emp, 2) if emp > 0 else 0
        nat_data.append({
            'name': nat,
            'total': total,
            'employees': emp,
            'per_head': per_head,
            'color': NAT_COLORS[nat],
            'item_list': summary[nat]['item_list'],
        })

    overall_per_head = round(grand_total / total_employees, 2) if total_employees > 0 else 0

    return {
        'period': period,
        'grand_total': grand_total,
        'total_employees': total_employees,
        'overall_per_head': overall_per_head,
        'nat_data': nat_data,
    }


def _trend_data(data: dict, periods: list) -> list:
    """Return trend data across all periods for each nationality."""
    trend = []
    for period in periods:
        summary = _period_summary(data, period)
        emp_counts = _employee_counts(data, period)
        entry = {'period': period}
        for nat in NAT_SHEETS:
            entry[nat] = summary[nat]['total']
            emp = emp_counts[nat]
            total = summary[nat]['total']
            entry[f'{nat}_per_head'] = round(total / emp, 2) if emp > 0 else 0
        trend.append(entry)
    return trend


# ---------------------------------------------------------------------------
# Routes
# ---------------------------------------------------------------------------
@app.route('/')
def index():
    data = _load_data()
    periods = _all_periods(data)
    selected = request.args.get('period', periods[-1] if periods else None)
    dashboard = _build_dashboard(data, selected) if selected else None
    trend = _trend_data(data, periods)
    has_excel = os.path.exists(EXCEL_PATH)
    return render_template(
        'index.html',
        periods=periods,
        selected_period=selected,
        dashboard=dashboard,
        trend_json=json.dumps(trend),
        nat_sheets=NAT_SHEETS,
        nat_colors=NAT_COLORS,
        has_excel=has_excel,
    )


@app.route('/api/report')
def api_report():
    """Return JSON report data filtered by period and optional nationality."""
    data = _load_data()
    periods = _all_periods(data)

    period_filter = request.args.get('period', 'all')
    nat_filter = request.args.get('nationality', 'all')

    target_periods = periods if period_filter == 'all' else [period_filter]
    target_nats = NAT_SHEETS if nat_filter == 'all' else [nat_filter]

    rows = []
    for period in target_periods:
        summary = _period_summary(data, period)
        emp_counts = _employee_counts(data, period)
        for nat in target_nats:
            if nat not in summary:
                continue
            for item in summary[nat]['item_list']:
                emp = emp_counts.get(nat, 0)
                total = float(item.get('total', 0) or 0)
                qty = float(item.get('qty', 0) or 0)
                per_head_qty = round(qty / emp, 4) if emp > 0 else 0
                per_head_cost = round(total / emp, 2) if emp > 0 else 0
                rows.append({
                    'period': period,
                    'nationality': nat,
                    'item': item.get('items', ''),
                    'qty': qty,
                    'unit': item.get('unit', ''),
                    'unit_price': float(item.get('unit_price', 0) or 0),
                    'total': total,
                    'employees': emp,
                    'qty_per_head': per_head_qty,
                    'cost_per_head': per_head_cost,
                })

    return jsonify({'rows': rows, 'periods': periods, 'nationalities': NAT_SHEETS})


@app.route('/api/download')
def api_download():
    """Download filtered report as CSV."""
    data = _load_data()
    periods = _all_periods(data)

    period_filter = request.args.get('period', 'all')
    nat_filter = request.args.get('nationality', 'all')

    target_periods = periods if period_filter == 'all' else [period_filter]
    target_nats = NAT_SHEETS if nat_filter == 'all' else [nat_filter]

    output = io.StringIO()
    writer = csv.writer(output)
    writer.writerow([
        'Period', 'Nationality', 'Item', 'Qty', 'Unit',
        'Unit Price (Rs)', 'Total (Rs)', 'Employees',
        'Qty/Head', 'Cost/Head (Rs)',
    ])

    for period in target_periods:
        summary = _period_summary(data, period)
        emp_counts = _employee_counts(data, period)
        for nat in target_nats:
            if nat not in summary:
                continue
            for item in summary[nat]['item_list']:
                emp = emp_counts.get(nat, 0)
                qty = float(item.get('qty', 0) or 0)
                total = float(item.get('total', 0) or 0)
                writer.writerow([
                    period,
                    nat,
                    item.get('items', ''),
                    qty,
                    item.get('unit', ''),
                    float(item.get('unit_price', 0) or 0),
                    total,
                    emp,
                    round(qty / emp, 4) if emp > 0 else 0,
                    round(total / emp, 2) if emp > 0 else 0,
                ])

    output.seek(0)
    filename = f'canteen_report_{period_filter}_{nat_filter}.csv'
    return Response(
        output.getvalue(),
        mimetype='text/csv',
        headers={'Content-Disposition': f'attachment; filename={filename}'},
    )


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------
if __name__ == '__main__':
    debug = os.environ.get('FLASK_DEBUG', '0') == '1'
    app.run(host='0.0.0.0', port=5000, debug=debug)
