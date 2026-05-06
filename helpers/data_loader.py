"""Data loading and parsing module for Canteen Expenses."""
import os
import logging
import math
import threading
from datetime import datetime, date

import pandas as pd

logger = logging.getLogger(__name__)

NATIONALITY_SHEETS = ["Bangladeshi", "Malagasy", "Indian", "Srilankan"]
EXPENSE_COLUMNS = ["PERIOD", "ITEMS", "QTY", "UNIT", "UNIT PRICE", "TOTAL"]
EMPLOYEE_COLUMNS = ["PERIOD", "BANGLADESHI", "INDIAN", "MALAGASY", "SRILANKAN"]

_cache: dict = {}
_cache_path: str = ""
_cache_lock = threading.Lock()


def _get_excel_path() -> str:
    return os.environ.get("EXCEL_PATH", os.path.join(os.path.dirname(os.path.dirname(__file__)), "Canteen.xlsx"))


def load_workbook_data() -> dict:
    """Load and cache all sheets from the Excel workbook."""
    global _cache, _cache_path
    path = _get_excel_path()
    with _cache_lock:
        if _cache and _cache_path == path:
            return _cache

        _cache_path = path
        if not os.path.exists(path):
            logger.warning("Excel file not found")
            _cache = {"nationality": {}, "employees": pd.DataFrame()}
            return _cache

        try:
            xl = pd.ExcelFile(path, engine="openpyxl")
            sheets = xl.sheet_names

            # Build case-insensitive, whitespace-stripped sheet name lookup
            sheet_lookup = {s.strip().lower(): s for s in sheets}

            def _find_sheet(target: str):
                """Return actual sheet name matching target (case-insensitive)."""
                exact = target if target in sheets else None
                if exact:
                    return exact
                return sheet_lookup.get(target.strip().lower())

            _BAD_PERIODS = {"", "nan", "none", "nat", "period"}

            def _valid_period(p: str) -> bool:
                return p.strip().lower() not in _BAD_PERIODS

            nationality_data: dict = {}
            for sheet in NATIONALITY_SHEETS:
                actual_sheet = _find_sheet(sheet)
                if actual_sheet:
                    try:
                        df = xl.parse(actual_sheet)
                        df.columns = [str(c).strip().upper() for c in df.columns]
                        for col in ["PERIOD", "TOTAL"]:
                            if col not in df.columns:
                                df[col] = None
                        if "ITEMS" not in df.columns:
                            df["ITEMS"] = ""
                        if "QTY" not in df.columns:
                            df["QTY"] = 0
                        if "UNIT" not in df.columns:
                            df["UNIT"] = ""
                        if "UNIT PRICE" not in df.columns:
                            df["UNIT PRICE"] = 0
                        df["PERIOD"] = df["PERIOD"].astype(str).str.strip()
                        df = df[df["PERIOD"].apply(_valid_period)]
                        df["TOTAL"] = pd.to_numeric(df["TOTAL"], errors="coerce").fillna(0)
                        df["QTY"] = pd.to_numeric(df["QTY"], errors="coerce").fillna(0)
                        df["UNIT PRICE"] = pd.to_numeric(df["UNIT PRICE"], errors="coerce").fillna(0)
                        df["nationality"] = sheet  # always use canonical name
                        nationality_data[sheet] = df
                    except Exception:
                        logger.error("Error parsing sheet %r", sheet, exc_info=True)
                        nationality_data[sheet] = pd.DataFrame()

            employees_df = pd.DataFrame()
            for name in sheets:
                if name.strip().upper() == "EMPLOYEES":
                    try:
                        employees_df = xl.parse(name)
                        employees_df.columns = [str(c).strip().upper() for c in employees_df.columns]
                        employees_df["PERIOD"] = employees_df["PERIOD"].astype(str).str.strip()
                        employees_df = employees_df[employees_df["PERIOD"].apply(_valid_period)]
                        for col in ["BANGLADESHI", "INDIAN", "MALAGASY", "SRILANKAN"]:
                            if col in employees_df.columns:
                                employees_df[col] = pd.to_numeric(employees_df[col], errors="coerce").fillna(0)
                            else:
                                employees_df[col] = 0
                    except Exception:
                        logger.error("Error parsing EMPLOYEES sheet", exc_info=True)

            _cache = {"nationality": nationality_data, "employees": employees_df}
        except Exception:
            logger.error("Failed to load workbook")
            _cache = {"nationality": {}, "employees": pd.DataFrame()}

        return _cache


def get_nationality_sheets() -> list:
    """Return list of nationality sheet names found in workbook."""
    data = load_workbook_data()
    return [s for s in NATIONALITY_SHEETS if s in data["nationality"] and not data["nationality"][s].empty]


def parse_period_date(period_str: str):
    """Parse '01.04.2026 - 15.04.2026' into (start_date, end_date). Returns (None, None) on failure."""
    if not period_str or period_str == "nan":
        return None, None
    try:
        # Normalise separators: remove spaces around hyphens for splitting
        s = str(period_str).strip()
        # Split on ' - ' first, then on plain '-' but watch for day.month.year-day.month.year
        if " - " in s:
            parts = s.split(" - ", 1)
        elif s.count("-") == 1:
            parts = s.split("-", 1)
        elif s.count("-") >= 2:
            # Could be '01.04.2026-15.04.2026' (no spaces)
            # Find the hyphen that is NOT inside a date (dates use dots)
            idx = s.index("-")
            # Check if both sides look like dates
            parts = [s[:idx].strip(), s[idx + 1:].strip()]
        else:
            return None, None

        if len(parts) != 2:
            return None, None

        start_str = parts[0].strip()
        end_str = parts[1].strip()

        for fmt in ("%d.%m.%Y", "%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y"):
            try:
                start = datetime.strptime(start_str, fmt).date()
                end = datetime.strptime(end_str, fmt).date()
                return start, end
            except ValueError:
                continue
        return None, None
    except Exception:
        return None, None


def _period_in_range(period_str: str, period_from: str | None, period_to: str | None) -> bool:
    """Return True if the period overlaps with the [period_from, period_to] range."""
    if not period_from and not period_to:
        return True
    start, end = parse_period_date(period_str)
    if start is None:
        return True  # can't determine, include it

    if period_from:
        from_start, _ = parse_period_date(period_from)
        if from_start and end and end < from_start:
            return False
    if period_to:
        _, to_end = parse_period_date(period_to)
        if to_end and start and start > to_end:
            return False
    return True


def get_all_expenses(nationality: str | None = None, period_from: str | None = None, period_to: str | None = None) -> pd.DataFrame:
    """Return filtered expense rows from nationality sheets."""
    data = load_workbook_data()
    sheets = data["nationality"]
    if not sheets:
        return pd.DataFrame()

    targets = [nationality] if nationality and nationality in sheets else list(sheets.keys())
    frames = []
    for name in targets:
        df = sheets.get(name, pd.DataFrame())
        if df.empty:
            continue
        frames.append(df)

    if not frames:
        return pd.DataFrame()

    combined = pd.concat(frames, ignore_index=True)
    if period_from or period_to:
        mask = combined["PERIOD"].apply(lambda p: _period_in_range(p, period_from, period_to))
        combined = combined[mask]
    return combined


def get_employees(period_from: str | None = None, period_to: str | None = None) -> pd.DataFrame:
    """Return employee rows optionally filtered by period range."""
    data = load_workbook_data()
    df = data["employees"]
    if df.empty:
        return df
    if period_from or period_to:
        mask = df["PERIOD"].apply(lambda p: _period_in_range(p, period_from, period_to))
        df = df[mask]
    return df.copy()


def get_periods() -> list:
    """Return sorted list of all unique parseable periods found across all sheets (including EMPLOYEES)."""
    data = load_workbook_data()
    periods = set()
    for df in data["nationality"].values():
        if not df.empty and "PERIOD" in df.columns:
            periods.update(df["PERIOD"].dropna().unique().tolist())
    emp = data["employees"]
    if not emp.empty and "PERIOD" in emp.columns:
        periods.update(emp["PERIOD"].dropna().unique().tolist())

    def sort_key(p):
        start, _ = parse_period_date(str(p))
        return start or date.min

    # Only include periods that successfully parse to a real date
    valid = []
    for p in periods:
        s = str(p)
        if s.lower() in ("nan", "none", "", "period"):
            continue
        start, _ = parse_period_date(s)
        if start is not None:
            valid.append(s)

    return sorted(valid, key=sort_key)


def get_expense_periods() -> list:
    """Return sorted list of parseable periods from expense sheets only (excludes EMPLOYEES sheet).

    Use this for the dashboard default period so that the latest selected period
    always corresponds to a period with actual expense data.
    """
    data = load_workbook_data()
    periods = set()
    for df in data["nationality"].values():
        if not df.empty and "PERIOD" in df.columns:
            periods.update(df["PERIOD"].dropna().unique().tolist())

    def sort_key(p):
        start, _ = parse_period_date(str(p))
        return start or date.min

    valid = []
    for p in periods:
        s = str(p)
        if s.lower() in ("nan", "none", "", "period"):
            continue
        start, _ = parse_period_date(s)
        if start is not None:
            valid.append(s)

    return sorted(valid, key=sort_key)


def format_rs(amount) -> str:
    """Format as Mauritian Rupee with standard Western grouping: Rs 1,000,000"""
    try:
        if amount is None or (isinstance(amount, float) and math.isnan(amount)):
            return "Rs 0"
        amount = float(amount)
    except (TypeError, ValueError):
        return "Rs 0"

    is_negative = amount < 0
    amount = abs(amount)

    int_part = int(amount)
    dec_part = round((amount - int_part) * 100)

    # Standard Western grouping: groups of 3
    result = f"{int_part:,}"

    if dec_part > 0:
        result = f"{result}.{dec_part:02d}"

    prefix = "Rs -" if is_negative else "Rs "
    return f"{prefix}{result}"


def _days_in_period(period_str: str) -> int:
    """Return number of days in a fortnight period (default 15)."""
    start, end = parse_period_date(period_str)
    if start and end:
        return max(1, (end - start).days + 1)
    return 15


def get_kpi_data(period_from: str | None = None, period_to: str | None = None) -> dict:
    """Return KPI dict for dashboard."""
    expenses = get_all_expenses(period_from=period_from, period_to=period_to)
    employees = get_employees(period_from=period_from, period_to=period_to)

    total_expenditure = float(expenses["TOTAL"].sum()) if not expenses.empty else 0.0

    # Total employees: peak count in any single period
    nat_col_keys = ["BANGLADESHI", "INDIAN", "MALAGASY", "SRILANKAN"]
    nat_cols = [c for c in nat_col_keys if not employees.empty and c in employees.columns]
    total_employees = 0
    if not employees.empty and nat_cols:
        _emp = employees.copy()
        _emp["_total"] = _emp[nat_cols].sum(axis=1)
        total_employees = int(_emp["_total"].max())

    avg_per_head = 0.0
    if total_employees > 0:
        avg_per_head = total_expenditure / total_employees

    # Top nationality by expenditure
    top_nat = ""
    if not expenses.empty and "nationality" in expenses.columns:
        nat_totals = expenses.groupby("nationality")["TOTAL"].sum()
        if not nat_totals.empty:
            top_nat = str(nat_totals.idxmax())

    # Nationality expenditure — always include all 4 in NATIONALITY_SHEETS order so bar chart
    # colours are always consistent and every nationality shows (with 0 if no data).
    nat_exp: dict = {nat: 0.0 for nat in NATIONALITY_SHEETS}
    if not expenses.empty and "nationality" in expenses.columns:
        for nat, grp in expenses.groupby("nationality"):
            if str(nat) in nat_exp:
                nat_exp[str(nat)] = float(grp["TOTAL"].sum())

    # Employee counts — always include all 4 in NATIONALITY_SHEETS order (consistent colours)
    nat_emp: dict = {}
    nat_col_map = {n: n.upper() for n in NATIONALITY_SHEETS}
    for nat in NATIONALITY_SHEETS:
        col = nat_col_map[nat]
        if not employees.empty and col in employees.columns:
            nat_emp[nat] = int(employees[col].sum())
        else:
            nat_emp[nat] = 0

    return {
        "total_employees": total_employees,
        "total_expenditure": total_expenditure,
        "total_expenditure_fmt": format_rs(total_expenditure),
        "avg_per_head": avg_per_head,
        "avg_per_head_fmt": format_rs(avg_per_head),
        "top_nationality": top_nat,
        "nat_expenditure": nat_exp,
        "nat_employees": nat_emp,
    }


def get_trend_data(nationality: str | None = None, period_from: str | None = None, period_to: str | None = None) -> dict:
    """Return trend chart data grouped by period."""
    expenses = get_all_expenses(nationality=nationality, period_from=period_from, period_to=period_to)
    if expenses.empty:
        return {"labels": [], "datasets": []}

    # Per-period totals
    period_totals = expenses.groupby("PERIOD")["TOTAL"].sum().reset_index()

    def sort_key(p):
        start, _ = parse_period_date(str(p))
        return start or date.min

    period_totals["_sort"] = period_totals["PERIOD"].apply(sort_key)
    period_totals = period_totals.sort_values("_sort")

    labels = period_totals["PERIOD"].tolist()
    values = [float(v) for v in period_totals["TOTAL"].tolist()]

    # Monthly aggregation
    monthly: dict = {}
    for _, row in period_totals.iterrows():
        start, _ = parse_period_date(str(row["PERIOD"]))
        if start:
            key = start.strftime("%b %Y")
        else:
            key = str(row["PERIOD"])
        monthly[key] = monthly.get(key, 0.0) + float(row["TOTAL"])

    # Per-nationality per-period for stacked chart
    nat_datasets = []
    if not nationality and "nationality" in expenses.columns:
        colors = {
            "Bangladeshi": "#3a86ff",
            "Indian": "#06d6a0",
            "Malagasy": "#ffd166",
            "Srilankan": "#ef233c",
        }
        for nat in NATIONALITY_SHEETS:
            nat_df = expenses[expenses["nationality"] == nat]
            if nat_df.empty:
                continue
            nat_period = nat_df.groupby("PERIOD")["TOTAL"].sum()
            nat_values = [float(nat_period.get(p, 0)) for p in labels]
            nat_datasets.append({
                "label": nat,
                "data": nat_values,
                "borderColor": colors.get(nat, "#ffffff"),
                "backgroundColor": colors.get(nat, "#ffffff") + "33",
                "fill": False,
                "tension": 0.3,
            })

    return {
        "labels": labels,
        "values": values,
        "monthly_labels": list(monthly.keys()),
        "monthly_values": list(monthly.values()),
        "nat_datasets": nat_datasets,
    }


def get_comparison_data(period_from: str | None = None, period_to: str | None = None) -> list:
    """Return comparison table data per nationality."""
    employees = get_employees(period_from=period_from, period_to=period_to)
    rows = []

    nat_col_map = {
        "Bangladeshi": "BANGLADESHI",
        "Indian": "INDIAN",
        "Malagasy": "MALAGASY",
        "Srilankan": "SRILANKAN",
    }

    for nat in NATIONALITY_SHEETS:
        expenses = get_all_expenses(nationality=nat, period_from=period_from, period_to=period_to)
        total_exp = float(expenses["TOTAL"].sum()) if not expenses.empty else 0.0

        emp_col = nat_col_map.get(nat, nat.upper())
        total_emp = 0
        if not employees.empty and emp_col in employees.columns:
            total_emp = int(employees[emp_col].sum())

        per_head = total_exp / total_emp if total_emp > 0 else 0.0

        # Per-day: sum days across periods
        total_days = 0
        if not expenses.empty:
            for p in expenses["PERIOD"].unique():
                total_days += _days_in_period(str(p))
        per_day = total_exp / total_days if total_days > 0 else 0.0

        rows.append({
            "nationality": nat,
            "total_employees": total_emp,
            "total_expenditure": total_exp,
            "total_expenditure_fmt": format_rs(total_exp),
            "per_head_avg": per_head,
            "per_head_avg_fmt": format_rs(per_head),
            "per_day_avg": per_day,
            "per_day_avg_fmt": format_rs(per_day),
        })

    return rows


def get_report_perhead(period_from=None, period_to=None, nationality=None) -> pd.DataFrame:
    """Per-head consumption report."""
    nat_col_map = {
        "Bangladeshi": "BANGLADESHI",
        "Indian": "INDIAN",
        "Malagasy": "MALAGASY",
        "Srilankan": "SRILANKAN",
    }
    targets = [nationality] if nationality else NATIONALITY_SHEETS
    employees = get_employees(period_from=period_from, period_to=period_to)
    rows = []
    for nat in targets:
        expenses = get_all_expenses(nationality=nat, period_from=period_from, period_to=period_to)
        if expenses.empty:
            continue
        for period, grp in expenses.groupby("PERIOD"):
            total_exp = float(grp["TOTAL"].sum())
            emp_col = nat_col_map.get(nat, nat.upper())
            emp_count = 0
            if not employees.empty and emp_col in employees.columns:
                emp_row = employees[employees["PERIOD"] == period]
                if not emp_row.empty:
                    emp_count = int(emp_row[emp_col].iloc[0])
            per_head = total_exp / emp_count if emp_count > 0 else 0.0
            rows.append({
                "Period": period,
                "Nationality": nat,
                "Employees": emp_count,
                "Total Expenditure (Rs)": total_exp,
                "Per Head Avg (Rs)": round(per_head, 2),
            })
    return pd.DataFrame(rows)


def get_report_detailed(period_from=None, period_to=None, nationality=None) -> pd.DataFrame:
    """Detailed expenditure report."""
    expenses = get_all_expenses(nationality=nationality, period_from=period_from, period_to=period_to)
    if expenses.empty:
        return pd.DataFrame()
    cols = [c for c in ["PERIOD", "nationality", "ITEMS", "QTY", "UNIT", "UNIT PRICE", "TOTAL"] if c in expenses.columns]
    df = expenses[cols].copy()
    df = df.rename(columns={"nationality": "Nationality", "PERIOD": "Period", "ITEMS": "Item",
                             "QTY": "Qty", "UNIT": "Unit", "UNIT PRICE": "Unit Price", "TOTAL": "Total"})
    return df


def get_report_summary(period_from=None, period_to=None) -> pd.DataFrame:
    """Summary report by period and nationality."""
    rows = []
    for nat in NATIONALITY_SHEETS:
        expenses = get_all_expenses(nationality=nat, period_from=period_from, period_to=period_to)
        if expenses.empty:
            continue
        for period, grp in expenses.groupby("PERIOD"):
            rows.append({
                "Period": period,
                "Nationality": nat,
                "Items Count": len(grp),
                "Total Expenditure (Rs)": float(grp["TOTAL"].sum()),
            })
    return pd.DataFrame(rows)


def get_report_monthly(period_from=None, period_to=None) -> pd.DataFrame:
    """Monthly/Fortnight aggregation report."""
    rows = []
    for nat in NATIONALITY_SHEETS:
        expenses = get_all_expenses(nationality=nat, period_from=period_from, period_to=period_to)
        if expenses.empty:
            continue
        for period, grp in expenses.groupby("PERIOD"):
            start, _ = parse_period_date(str(period))
            month = start.strftime("%B %Y") if start else str(period)
            rows.append({
                "Month": month,
                "Fortnight Period": period,
                "Nationality": nat,
                "Total Expenditure (Rs)": float(grp["TOTAL"].sum()),
            })
    df = pd.DataFrame(rows)
    if df.empty:
        return df
    df = df.sort_values(["Month", "Fortnight Period", "Nationality"])
    return df


def get_item_comparison(nationality: str | None = None, period_from: str | None = None, period_to: str | None = None) -> dict:
    """Return item-by-item comparison across nationalities.

    Returns a dict:
      {
        "nationalities": [...],
        "items": [
          {
            "item": "Rice",
            "by_nationality": {
              "Bangladeshi": {"qty": 500, "total": 25000, "total_fmt": "Rs 25,000", "per_head": 35.71, "per_head_fmt": "Rs 35.71"},
              ...
            },
            "grand_total": 75000,
            "grand_total_fmt": "Rs 75,000",
          },
          ...
        ]
      }
    """
    targets = [nationality] if nationality else NATIONALITY_SHEETS
    employees = get_employees(period_from=period_from, period_to=period_to)

    nat_col_map = {
        "Bangladeshi": "BANGLADESHI",
        "Indian": "INDIAN",
        "Malagasy": "MALAGASY",
        "Srilankan": "SRILANKAN",
    }

    # Collect per-nationality employee totals
    nat_emp: dict = {}
    for nat in targets:
        emp_col = nat_col_map.get(nat, nat.upper())
        if not employees.empty and emp_col in employees.columns:
            nat_emp[nat] = int(employees[emp_col].sum())
        else:
            nat_emp[nat] = 0

    # Collect all expense rows
    all_items: dict = {}  # item_name -> {nat -> {qty, total}}
    for nat in targets:
        exp = get_all_expenses(nationality=nat, period_from=period_from, period_to=period_to)
        if exp.empty:
            continue
        if "ITEMS" not in exp.columns:
            continue
        for item_name, grp in exp.groupby("ITEMS"):
            item_key = str(item_name).strip()
            if not item_key or item_key.lower() in ("nan", "none", ""):
                continue
            if item_key not in all_items:
                all_items[item_key] = {}
            qty = float(grp["QTY"].sum()) if "QTY" in grp.columns else 0.0
            total = float(grp["TOTAL"].sum())
            emp_count = nat_emp.get(nat, 0)
            per_head = total / emp_count if emp_count > 0 else 0.0
            all_items[item_key][nat] = {
                "qty": round(qty, 2),
                "total": round(total, 2),
                "total_fmt": format_rs(total),
                "per_head": round(per_head, 2),
                "per_head_fmt": format_rs(per_head),
            }

    # Sort items by grand total descending
    def item_grand_total(item_data):
        return sum(v["total"] for v in item_data.values())

    sorted_items = sorted(all_items.items(), key=lambda kv: item_grand_total(kv[1]), reverse=True)

    items_list = []
    for item_name, by_nat in sorted_items:
        grand_total = item_grand_total(by_nat)
        items_list.append({
            "item": item_name,
            "by_nationality": by_nat,
            "grand_total": round(grand_total, 2),
            "grand_total_fmt": format_rs(grand_total),
        })

    return {
        "nationalities": targets,
        "nat_employees": nat_emp,
        "items": items_list,
    }
