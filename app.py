import os
import io
import csv
import json

import pandas as pd
from flask import Flask, render_template, jsonify, request, make_response, Response

app = Flask(__name__)

EXCEL_PATH = os.environ.get('EXCEL_PATH', 'Canteen.xlsx')

NATIONALITIES = ['Bangladeshi', 'Indian', 'Malagasy', 'Srilankan']

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def fmt_currency(value):
    """Format a numeric value as  Rs X,XXX  (no decimal places for whole numbers)."""
    try:
        v = float(value)
        if v == int(v):
            return f"Rs {int(v):,}"
        return f"Rs {v:,.2f}"
    except (TypeError, ValueError):
        return "Rs 0"


def fmt_number(value):
    """Format a plain integer with comma separators."""
    try:
        return f"{int(value):,}"
    except (TypeError, ValueError):
        return "0"


# ---------------------------------------------------------------------------
# Excel loading
# ---------------------------------------------------------------------------

def load_excel():
    if not os.path.exists(EXCEL_PATH):
        return None
    try:
        return pd.ExcelFile(EXCEL_PATH)
    except Exception as exc:
        print(f"Error loading Excel: {exc}")
        return None


def _normalise_cols(df):
    df.columns = [str(c).strip().upper() for c in df.columns]
    return df


def get_expense_df(xl, sheet_name):
    try:
        df = _normalise_cols(xl.parse(sheet_name))
        if 'PERIOD' not in df.columns:
            return pd.DataFrame()
        df = df.dropna(subset=['PERIOD'])
        if 'TOTAL' in df.columns:
            df['TOTAL'] = pd.to_numeric(df['TOTAL'], errors='coerce').fillna(0)
        if 'UNIT PRICE' in df.columns:
            df['UNIT PRICE'] = pd.to_numeric(df['UNIT PRICE'], errors='coerce').fillna(0)
        if 'QTY' in df.columns:
            df['QTY'] = pd.to_numeric(df['QTY'], errors='coerce').fillna(0)
        return df
    except Exception as exc:
        print(f"Error reading sheet {sheet_name}: {exc}")
        return pd.DataFrame()


def get_employee_df(xl):
    try:
        df = _normalise_cols(xl.parse('EMPLOYEES'))
        if 'PERIOD' not in df.columns:
            return pd.DataFrame()
        df = df.dropna(subset=['PERIOD'])
        for col in ['BANGLADESHI', 'INDIAN', 'MALAGASY', 'SRILANKAN']:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0).astype(int)
        return df
    except Exception as exc:
        print(f"Error reading EMPLOYEES: {exc}")
        return pd.DataFrame()


def all_periods(xl):
    periods = set()
    for nat in NATIONALITIES:
        df = get_expense_df(xl, nat)
        if not df.empty and 'PERIOD' in df.columns:
            periods.update(df['PERIOD'].dropna().astype(str).unique())
    emp_df = get_employee_df(xl)
    if not emp_df.empty and 'PERIOD' in emp_df.columns:
        periods.update(emp_df['PERIOD'].dropna().astype(str).unique())
    return sorted(periods)


def filter_by_period(df, period_from, period_to, available):
    if not period_from and not period_to:
        return df
    indices = []
    if period_from and period_from in available:
        indices.append(available.index(period_from))
    else:
        indices.append(0)
    if period_to and period_to in available:
        indices.append(available.index(period_to))
    else:
        indices.append(len(available) - 1)
    subset = available[indices[0]: indices[1] + 1]
    return df[df['PERIOD'].astype(str).isin(subset)]


# ---------------------------------------------------------------------------
# Sample / fallback data
# ---------------------------------------------------------------------------

SAMPLE_PERIODS = [
    '01.04.2026 - 15.04.2026',
    '16.04.2026 - 30.04.2026',
    '01.05.2026 - 15.05.2026',
    '16.05.2026 - 31.05.2026',
    '01.06.2026 - 15.06.2026',
]

SAMPLE_EXPENSES = {
    'Bangladeshi': [125000, 132500, 130000, 131250, 140000],
    'Indian':      [87500,  87500,  78750,  74025,  87500],
    'Malagasy':    [24150,  24150,  24150,  25475,  70400],
    'Srilankan':   [35000,  70000,  43750,  61250,  87500],
}

SAMPLE_EMPLOYEES = {
    'Bangladeshi': [700, 750, 750, 750, 800],
    'Indian':      [500, 500, 500, 450, 500],
    'Malagasy':    [138, 138, 138, 145, 400],
    'Srilankan':   [200, 400, 350, 350, 500],
}

# Detailed items for the expense detail view (sample only)
SAMPLE_ITEMS = {
    'Bangladeshi': [
        ('Rice', 2000, 'Kg', 25),
        ('Chicken', 500, 'Kg', 120),
        ('Vegetables', 800, 'Kg', 30),
        ('Oil', 200, 'L', 90),
        ('Spices', 50, 'Kg', 300),
    ],
    'Indian': [
        ('Rice', 1500, 'Kg', 25),
        ('Dal', 400, 'Kg', 80),
        ('Vegetables', 600, 'Kg', 30),
        ('Oil', 150, 'L', 90),
        ('Spices', 40, 'Kg', 300),
    ],
    'Malagasy': [
        ('Rice', 800, 'Kg', 25),
        ('Fish', 200, 'Kg', 80),
        ('Vegetables', 300, 'Kg', 30),
    ],
    'Srilankan': [
        ('Rice', 1000, 'Kg', 25),
        ('Fish', 300, 'Kg', 100),
        ('Vegetables', 400, 'Kg', 30),
        ('Coconut', 500, 'Pc', 10),
    ],
}


def sample_expense_rows(nationality, period):
    rows = []
    items = SAMPLE_ITEMS.get(nationality, [])
    for name, qty, unit, price in items:
        rows.append({
            'period': period,
            'item': name,
            'qty': qty,
            'unit': unit,
            'unit_price': fmt_currency(price),
            'total': fmt_currency(qty * price),
        })
    return rows


# ---------------------------------------------------------------------------
# Routes – pages
# ---------------------------------------------------------------------------

@app.route('/')
def dashboard():
    return render_template('dashboard.html', active='dashboard')


@app.route('/employees')
def employees():
    return render_template('employees.html', active='employees')


@app.route('/trends')
def trends():
    return render_template('trends.html', active='trends')


@app.route('/comparison')
def comparison():
    return render_template('comparison.html', active='comparison')


@app.route('/reports')
def reports():
    return render_template('reports.html', active='reports')


# ---------------------------------------------------------------------------
# API – periods list
# ---------------------------------------------------------------------------

@app.route('/api/periods')
def api_periods():
    xl = load_excel()
    if xl is None:
        return jsonify(SAMPLE_PERIODS)
    return jsonify(all_periods(xl))


# ---------------------------------------------------------------------------
# API – dashboard summary
# ---------------------------------------------------------------------------

@app.route('/api/dashboard')
def api_dashboard():
    xl = load_excel()
    period_from = request.args.get('from', '')
    period_to   = request.args.get('to', '')

    if xl is None:
        periods = SAMPLE_PERIODS
        expenses = SAMPLE_EXPENSES
        emps     = SAMPLE_EMPLOYEES
    else:
        periods = all_periods(xl)
        expenses = {}
        emps     = {}
        for nat in NATIONALITIES:
            df = get_expense_df(xl, nat)
            if df.empty:
                expenses[nat] = [0] * len(periods)
            else:
                g = df.groupby('PERIOD')['TOTAL'].sum()
                expenses[nat] = [float(g.get(p, 0)) for p in periods]
        emp_df = get_employee_df(xl)
        for nat in NATIONALITIES:
            col = nat.upper()
            if not emp_df.empty and col in emp_df.columns:
                col_map = emp_df.set_index('PERIOD')[col]
                emps[nat] = [int(col_map.get(p, 0)) for p in periods]
            else:
                emps[nat] = [0] * len(periods)

    # Apply period filter
    idx_from = 0
    idx_to   = len(periods) - 1
    if period_from and period_from in periods:
        idx_from = periods.index(period_from)
    if period_to and period_to in periods:
        idx_to = periods.index(period_to)
    sel_periods  = periods[idx_from: idx_to + 1]
    sel_expenses = {nat: expenses[nat][idx_from: idx_to + 1] for nat in NATIONALITIES}
    sel_emps     = {nat: emps[nat][idx_from: idx_to + 1]     for nat in NATIONALITIES}

    totals = {nat: sum(sel_expenses[nat]) for nat in NATIONALITIES}
    grand_total = sum(totals.values())

    # Cost per head: total cost / total employee-periods
    cph = {}
    for nat in NATIONALITIES:
        total_cost = sum(sel_expenses[nat])
        total_emp  = sum(sel_emps[nat])
        cph[nat] = total_cost / total_emp if total_emp > 0 else 0

    latest_emps = {nat: sel_emps[nat][-1] if sel_emps[nat] else 0 for nat in NATIONALITIES}

    return jsonify({
        'periods':          sel_periods,
        'expense_by_period': {nat: sel_expenses[nat] for nat in NATIONALITIES},
        'emp_by_period':    {nat: sel_emps[nat]      for nat in NATIONALITIES},
        'totals':           {nat: fmt_currency(v)    for nat, v in totals.items()},
        'grand_total':      fmt_currency(grand_total),
        'cost_per_head':    {nat: fmt_currency(v)    for nat, v in cph.items()},
        'latest_employees': latest_emps,
        'total_employees':  sum(latest_emps.values()),
    })


# ---------------------------------------------------------------------------
# API – employees
# ---------------------------------------------------------------------------

@app.route('/api/employees')
def api_employees():
    xl = load_excel()
    period_from = request.args.get('from', '')
    period_to   = request.args.get('to', '')

    if xl is None:
        periods = SAMPLE_PERIODS
        emps    = SAMPLE_EMPLOYEES
    else:
        periods = all_periods(xl)
        emp_df  = get_employee_df(xl)
        emps    = {}
        for nat in NATIONALITIES:
            col = nat.upper()
            if not emp_df.empty and col in emp_df.columns:
                col_map = emp_df.set_index('PERIOD')[col]
                emps[nat] = [int(col_map.get(p, 0)) for p in periods]
            else:
                emps[nat] = [0] * len(periods)

    idx_from = 0
    idx_to   = len(periods) - 1
    if period_from and period_from in periods:
        idx_from = periods.index(period_from)
    if period_to and period_to in periods:
        idx_to = periods.index(period_to)

    rows = []
    for i in range(idx_from, idx_to + 1):
        bd  = emps['Bangladeshi'][i]
        ind = emps['Indian'][i]
        mal = emps['Malagasy'][i]
        sri = emps['Srilankan'][i]
        rows.append({
            'period':      periods[i],
            'bangladeshi': bd,
            'indian':      ind,
            'malagasy':    mal,
            'srilankan':   sri,
            'total':       bd + ind + mal + sri,
        })
    return jsonify({'data': rows})


# ---------------------------------------------------------------------------
# API – trends
# ---------------------------------------------------------------------------

@app.route('/api/trends')
def api_trends():
    xl = load_excel()
    period_from = request.args.get('from', '')
    period_to   = request.args.get('to', '')

    if xl is None:
        periods  = SAMPLE_PERIODS
        expenses = SAMPLE_EXPENSES
        emps     = SAMPLE_EMPLOYEES
    else:
        periods  = all_periods(xl)
        expenses = {}
        emps     = {}
        for nat in NATIONALITIES:
            df = get_expense_df(xl, nat)
            if df.empty:
                expenses[nat] = [0] * len(periods)
            else:
                g = df.groupby('PERIOD')['TOTAL'].sum()
                expenses[nat] = [float(g.get(p, 0)) for p in periods]
        emp_df = get_employee_df(xl)
        for nat in NATIONALITIES:
            col = nat.upper()
            if not emp_df.empty and col in emp_df.columns:
                cm = emp_df.set_index('PERIOD')[col]
                emps[nat] = [int(cm.get(p, 0)) for p in periods]
            else:
                emps[nat] = [0] * len(periods)

    idx_from = 0
    idx_to   = len(periods) - 1
    if period_from and period_from in periods:
        idx_from = periods.index(period_from)
    if period_to and period_to in periods:
        idx_to = periods.index(period_to)

    sel_periods  = periods[idx_from: idx_to + 1]
    sel_expenses = {nat: expenses[nat][idx_from: idx_to + 1] for nat in NATIONALITIES}
    sel_emps     = {nat: emps[nat][idx_from: idx_to + 1]     for nat in NATIONALITIES}

    # Cost per head per period
    cph = {}
    for nat in NATIONALITIES:
        cph[nat] = []
        for cost, emp in zip(sel_expenses[nat], sel_emps[nat]):
            cph[nat].append(round(cost / emp, 2) if emp > 0 else 0)

    # Trend table rows
    rows = []
    for i, period in enumerate(sel_periods):
        row = {'period': period}
        for nat in NATIONALITIES:
            row[nat.lower() + '_expense'] = fmt_currency(sel_expenses[nat][i])
            row[nat.lower() + '_emp']     = sel_emps[nat][i]
            row[nat.lower() + '_cph']     = fmt_currency(cph[nat][i])
        rows.append(row)

    return jsonify({
        'periods':  sel_periods,
        'expenses': sel_expenses,
        'cph':      cph,
        'rows':     rows,
    })


# ---------------------------------------------------------------------------
# API – comparison
# ---------------------------------------------------------------------------

@app.route('/api/comparison')
def api_comparison():
    xl = load_excel()
    period_from = request.args.get('from', '')
    period_to   = request.args.get('to', '')

    if xl is None:
        periods  = SAMPLE_PERIODS
        expenses = SAMPLE_EXPENSES
        emps     = SAMPLE_EMPLOYEES
    else:
        periods  = all_periods(xl)
        expenses = {}
        emps     = {}
        for nat in NATIONALITIES:
            df = get_expense_df(xl, nat)
            if df.empty:
                expenses[nat] = [0] * len(periods)
            else:
                g = df.groupby('PERIOD')['TOTAL'].sum()
                expenses[nat] = [float(g.get(p, 0)) for p in periods]
        emp_df = get_employee_df(xl)
        for nat in NATIONALITIES:
            col = nat.upper()
            if not emp_df.empty and col in emp_df.columns:
                cm = emp_df.set_index('PERIOD')[col]
                emps[nat] = [int(cm.get(p, 0)) for p in periods]
            else:
                emps[nat] = [0] * len(periods)

    idx_from = 0
    idx_to   = len(periods) - 1
    if period_from and period_from in periods:
        idx_from = periods.index(period_from)
    if period_to and period_to in periods:
        idx_to = periods.index(period_to)

    sel_periods  = periods[idx_from: idx_to + 1]
    sel_expenses = {nat: expenses[nat][idx_from: idx_to + 1] for nat in NATIONALITIES}
    sel_emps     = {nat: emps[nat][idx_from: idx_to + 1]     for nat in NATIONALITIES}

    rows = []
    for i, period in enumerate(sel_periods):
        row = {'period': period}
        total_cost = 0
        total_emp  = 0
        for nat in NATIONALITIES:
            cost = sel_expenses[nat][i]
            emp  = sel_emps[nat][i]
            cph  = round(cost / emp, 2) if emp > 0 else 0
            row[nat.lower() + '_cost'] = fmt_currency(cost)
            row[nat.lower() + '_emp']  = emp
            row[nat.lower() + '_cph']  = fmt_currency(cph)
            total_cost += cost
            total_emp  += emp
        row['total_cost'] = fmt_currency(total_cost)
        row['total_emp']  = total_emp
        row['total_cph']  = fmt_currency(round(total_cost / total_emp, 2) if total_emp > 0 else 0)
        rows.append(row)

    # Aggregated totals across selected periods
    agg = {}
    for nat in NATIONALITIES:
        total_cost = sum(sel_expenses[nat])
        total_emp  = sum(sel_emps[nat])
        agg[nat] = {
            'total_cost': fmt_currency(total_cost),
            'total_emp':  total_emp,
            'avg_cph':    fmt_currency(round(total_cost / total_emp, 2) if total_emp > 0 else 0),
        }

    return jsonify({
        'periods':  sel_periods,
        'rows':     rows,
        'agg':      agg,
        'expenses': sel_expenses,
        'cph': {
            nat: [
                round(sel_expenses[nat][i] / sel_emps[nat][i], 2) if sel_emps[nat][i] > 0 else 0
                for i in range(len(sel_periods))
            ]
            for nat in NATIONALITIES
        },
    })


# ---------------------------------------------------------------------------
# API – reports  (consolidated report data)
# ---------------------------------------------------------------------------

@app.route('/api/reports')
def api_reports():
    """Return all report data for the Reports page."""
    xl = load_excel()
    period_from = request.args.get('from', '')
    period_to   = request.args.get('to', '')

    if xl is None:
        periods  = SAMPLE_PERIODS
        expenses = SAMPLE_EXPENSES
        emps     = SAMPLE_EMPLOYEES
    else:
        periods  = all_periods(xl)
        expenses = {}
        emps     = {}
        for nat in NATIONALITIES:
            df = get_expense_df(xl, nat)
            if df.empty:
                expenses[nat] = [0] * len(periods)
            else:
                g = df.groupby('PERIOD')['TOTAL'].sum()
                expenses[nat] = [float(g.get(p, 0)) for p in periods]
        emp_df = get_employee_df(xl)
        for nat in NATIONALITIES:
            col = nat.upper()
            if not emp_df.empty and col in emp_df.columns:
                cm = emp_df.set_index('PERIOD')[col]
                emps[nat] = [int(cm.get(p, 0)) for p in periods]
            else:
                emps[nat] = [0] * len(periods)

    idx_from = 0
    idx_to   = len(periods) - 1
    if period_from and period_from in periods:
        idx_from = periods.index(period_from)
    if period_to and period_to in periods:
        idx_to = periods.index(period_to)

    sel_periods  = periods[idx_from: idx_to + 1]
    sel_expenses = {nat: expenses[nat][idx_from: idx_to + 1] for nat in NATIONALITIES}
    sel_emps     = {nat: emps[nat][idx_from: idx_to + 1]     for nat in NATIONALITIES}

    # --- Report 1: Expense Summary ---
    expense_rows = []
    for i, period in enumerate(sel_periods):
        row = {'period': period}
        total = 0
        for nat in NATIONALITIES:
            v = sel_expenses[nat][i]
            row[nat.lower()] = fmt_currency(v)
            total += v
        row['total'] = fmt_currency(total)
        expense_rows.append(row)

    # --- Report 2: Employee Count ---
    emp_rows = []
    for i, period in enumerate(sel_periods):
        row = {'period': period}
        total = 0
        for nat in NATIONALITIES:
            v = sel_emps[nat][i]
            row[nat.lower()] = v
            total += v
        row['total'] = total
        emp_rows.append(row)

    # --- Report 3: Cost Per Head ---
    cph_rows = []
    for i, period in enumerate(sel_periods):
        row = {'period': period}
        total_cost = 0
        total_emp  = 0
        for nat in NATIONALITIES:
            cost = sel_expenses[nat][i]
            emp  = sel_emps[nat][i]
            cph  = round(cost / emp, 2) if emp > 0 else 0
            row[nat.lower()] = fmt_currency(cph)
            total_cost += cost
            total_emp  += emp
        row['total'] = fmt_currency(round(total_cost / total_emp, 2) if total_emp > 0 else 0)
        cph_rows.append(row)

    # --- Report 4: Grand totals ---
    totals_row = {}
    for nat in NATIONALITIES:
        totals_row[nat.lower()] = fmt_currency(sum(sel_expenses[nat]))
    totals_row['total'] = fmt_currency(sum(sum(sel_expenses[nat]) for nat in NATIONALITIES))

    return jsonify({
        'periods':       sel_periods,
        'expense_rows':  expense_rows,
        'emp_rows':      emp_rows,
        'cph_rows':      cph_rows,
        'totals':        totals_row,
        'expenses_raw':  sel_expenses,
        'cph_raw': {
            nat: [
                round(sel_expenses[nat][i] / sel_emps[nat][i], 2) if sel_emps[nat][i] > 0 else 0
                for i in range(len(sel_periods))
            ]
            for nat in NATIONALITIES
        },
    })


# ---------------------------------------------------------------------------
# CSV export endpoints
# ---------------------------------------------------------------------------

def _csv_response(rows, headers, filename):
    buf = io.StringIO()
    writer = csv.DictWriter(buf, fieldnames=headers)
    writer.writeheader()
    writer.writerows(rows)
    resp = make_response(buf.getvalue())
    resp.headers['Content-Disposition'] = f'attachment; filename="{filename}"'
    resp.headers['Content-Type'] = 'text/csv; charset=utf-8'
    return resp


@app.route('/export/employees/csv')
def export_employees_csv():
    data = api_employees().get_json()
    rows = [
        {
            'Period':      r['period'],
            'Bangladeshi': r['bangladeshi'],
            'Indian':      r['indian'],
            'Malagasy':    r['malagasy'],
            'Sri Lankan':  r['srilankan'],
            'Total':       r['total'],
        }
        for r in data['data']
    ]
    return _csv_response(rows, ['Period', 'Bangladeshi', 'Indian', 'Malagasy', 'Sri Lankan', 'Total'],
                         'employees.csv')


@app.route('/export/expenses/csv')
def export_expenses_csv():
    data = api_reports().get_json()
    rows = [
        {
            'Period':      r['period'],
            'Bangladeshi': r['bangladeshi'],
            'Indian':      r['indian'],
            'Malagasy':    r['malagasy'],
            'Sri Lankan':  r['srilankan'],
            'Total':       r['total'],
        }
        for r in data['expense_rows']
    ]
    return _csv_response(rows, ['Period', 'Bangladeshi', 'Indian', 'Malagasy', 'Sri Lankan', 'Total'],
                         'expense_summary.csv')


@app.route('/export/cph/csv')
def export_cph_csv():
    data = api_reports().get_json()
    rows = [
        {
            'Period':      r['period'],
            'Bangladeshi': r['bangladeshi'],
            'Indian':      r['indian'],
            'Malagasy':    r['malagasy'],
            'Sri Lankan':  r['srilankan'],
            'Total':       r['total'],
        }
        for r in data['cph_rows']
    ]
    return _csv_response(rows, ['Period', 'Bangladeshi', 'Indian', 'Malagasy', 'Sri Lankan', 'Total'],
                         'cost_per_head.csv')


@app.route('/export/comparison/csv')
def export_comparison_csv():
    data = api_comparison().get_json()
    headers = ['Period',
               'BD Cost', 'BD Emp', 'BD CPH',
               'IN Cost', 'IN Emp', 'IN CPH',
               'ML Cost', 'ML Emp', 'ML CPH',
               'SL Cost', 'SL Emp', 'SL CPH',
               'Total Cost', 'Total Emp', 'Total CPH']
    rows = []
    for r in data['rows']:
        rows.append({
            'Period':       r['period'],
            'BD Cost':      r.get('bangladeshi_cost', ''),
            'BD Emp':       r.get('bangladeshi_emp', ''),
            'BD CPH':       r.get('bangladeshi_cph', ''),
            'IN Cost':      r.get('indian_cost', ''),
            'IN Emp':       r.get('indian_emp', ''),
            'IN CPH':       r.get('indian_cph', ''),
            'ML Cost':      r.get('malagasy_cost', ''),
            'ML Emp':       r.get('malagasy_emp', ''),
            'ML CPH':       r.get('malagasy_cph', ''),
            'SL Cost':      r.get('srilankan_cost', ''),
            'SL Emp':       r.get('srilankan_emp', ''),
            'SL CPH':       r.get('srilankan_cph', ''),
            'Total Cost':   r.get('total_cost', ''),
            'Total Emp':    r.get('total_emp', ''),
            'Total CPH':    r.get('total_cph', ''),
        })
    return _csv_response(rows, headers, 'comparison.csv')


if __name__ == '__main__':
    debug = os.environ.get('FLASK_DEBUG', '0') == '1'
    app.run(debug=debug, host='0.0.0.0', port=5000)
