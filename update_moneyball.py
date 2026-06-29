"""
Moneyball monthly update script.
Usage: python update_moneyball.py "path/to/Monthly Excel.xlsx"

Reads the Excel file, computes P&L and cost increments, and updates data.json.
"""

import json
import sys
import os
import re
from pathlib import Path
import pandas as pd

# ── Config ───────────────────────────────────────────────────────────────────
DATA_JSON = Path(__file__).parent / "data.json"

# Reps currently on unpaid leave.
# Set return_month to None if unknown. Format: "Mon-YY"
LEAVE_CONFIG = {
    "BDA": {
        "Ivan Garville": {"leave_start": "Apr-26", "return_month": "Sep-26"},
    },
    "SBC": {}
}

TAX_RATE = 0.12  # Payroll tax: (Salary + Comm) * 12%

# ── Helpers ──────────────────────────────────────────────────────────────────
def r2(n):
    return round(n, 2)

def r4(n):
    return round(n, 4)

def month_label(dt):
    """Convert a datetime or string to 'Mon-YY' format."""
    if hasattr(dt, 'strftime'):
        return dt.strftime("%b-%y")
    return str(dt)

def get_last_val(arr, rep):
    """Get the last non-null value for a rep in a date/month array."""
    val = None
    for row in arr:
        v = row.get(rep)
        if v is not None:
            val = v
    return val

def get_last_month(arr, rep):
    """Get the last month number where a rep has data in a month array."""
    last = -1
    for row in arr:
        if row.get(rep) is not None:
            last = row["month"]
    return last

def ensure_rep_columns(arr, reps):
    """Add null columns for new reps to all existing rows."""
    for row in arr:
        for rep in reps:
            if rep not in row:
                row[rep] = None

def add_date_row(arr, date_label, rep_values):
    """Add or update a date row. rep_values: {rep: cumulative_value}"""
    row = next((r for r in arr if r["date"] == date_label), None)
    if row is None:
        # Build from last row's keys
        template = {k: None for k in arr[-1].keys()} if arr else {}
        row = dict(template)
        row["date"] = date_label
        arr.append(row)
    for rep, val in rep_values.items():
        row[rep] = val
    return arr

def add_month_row(arr, month_num, rep_values):
    """Add or update a specific tenure month row."""
    row = next((r for r in arr if r["month"] == month_num), None)
    if row is None:
        template = {k: None for k in arr[-1].keys()} if arr else {}
        row = dict(template)
        row["month"] = month_num
        arr.append(arr)  # placeholder — fill below
        arr[-1] = row
    for rep, val in rep_values.items():
        row[rep] = val
    return arr

# ── Read Excel ───────────────────────────────────────────────────────────────
def read_excel(path):
    """
    Returns {'BDA': [{rep, salary, comm, elc_rev, assoc, lead_bonus}], 'SBC': [...]}
    """
    df = pd.read_excel(path, sheet_name="Moneyball", header=None)

    result = {"BDA": [], "SBC": []}
    current_section = None
    header_row = None

    for i, row in df.iterrows():
        cell0 = str(row[0]).strip() if pd.notna(row[0]) else ""
        cell1 = str(row[1]).strip() if pd.notna(row[1]) else ""

        # Detect section headers
        if "BDA Moneyball" in cell0 or "BDA Moneyball" in cell1:
            current_section = "BDA"
            header_row = None
            continue
        if "SBC Moneyball" in cell0 or "SBC Moneyball" in cell1:
            current_section = "SBC"
            header_row = None
            continue

        # Detect column header row (Month, BDA/SBC, Salary, ...)
        if cell1 in ("BDA", "SBC") and str(row[2]).strip() == "Salary":
            header_row = i
            continue

        if current_section and header_row is not None:
            # Data row: col0=date, col1=rep name, col2=salary, col5=elcrev, col6=comm, col7=assoc, col8=lead_bonus
            rep = str(row[1]).strip() if pd.notna(row[1]) else ""
            if not rep or rep in ("nan", "Month", "BDA", "SBC"):
                continue

            def safe(val, default=0.0):
                try:
                    return float(val) if pd.notna(val) and str(val).strip() not in ("", "nan") else default
                except (ValueError, TypeError):
                    return default

            month_dt = row[0]
            result[current_section].append({
                "rep": rep,
                "month_label": month_label(month_dt),
                "salary": safe(row[2]),
                "comm":   safe(row[6]),
                "elc_rev": safe(row[5]),
                "assoc":  safe(row[7]),
                "lead_bonus": safe(row[8]),
            })

    return result

# ── Main update logic ─────────────────────────────────────────────────────────
def update(excel_path):
    data = json.loads(DATA_JSON.read_text())

    rows = read_excel(excel_path)

    for universe in ("BDA", "SBC"):
        rpr_key_date  = f"{universe}_DATE_RPR"
        rpr_key_month = f"{universe}_MONTH_RPR"
        cost_key_date  = f"{universe}_DATE_CCOST"
        cost_key_month = f"{universe}_MONTH_CCOST"
        pc_key_date   = f"{universe}_DATE_PC"
        pc_key_month  = f"{universe}_MONTH_PC"
        active_key    = f"{universe}_ACTIVE"
        reps_key      = f"{universe}_REPS"
        known_key     = f"{universe}_KNOWN_REPS"

        excel_rows = rows[universe]
        if not excel_rows:
            print(f"  No {universe} data in Excel, skipping.")
            continue

        date_label = excel_rows[0]["month_label"]
        print(f"\n{universe} — processing {date_label} ({len(excel_rows)} reps)")

        rpr_date  = data[rpr_key_date]
        rpr_month = data[rpr_key_month]
        cost_date  = data[cost_key_date]
        cost_month = data[cost_key_month]
        pc_date   = data[pc_key_date]
        pc_month  = data[pc_key_month]
        active_list = data[active_key]
        reps_list   = data[reps_key]
        known_reps  = set(data.get(known_key, []))
        leave_conf  = LEAVE_CONFIG.get(universe, {})

        new_rpr_date  = {}
        new_rpr_month = {}
        new_cost_date  = {}
        new_cost_month = {}
        new_pc_date   = {}
        new_pc_month  = {}

        processed_reps = set()

        for row in excel_rows:
            rep     = row["rep"]
            salary  = row["salary"]
            comm    = row["comm"]
            elc_rev = row["elc_rev"]
            assoc   = row["assoc"]
            lb      = row["lead_bonus"]

            # P&L formula
            payroll_tax = (salary + comm) * TAX_RATE
            monthly_rpr = elc_rev - comm - assoc - payroll_tax - salary - lb
            monthly_cost = (salary + comm) * (1 + TAX_RATE)  # direct cost (positive)

            prev_rpr  = get_last_val(rpr_date, rep) or 0.0
            prev_cost = get_last_val(cost_date, rep) or 0.0  # negative cumulative

            cum_rpr  = r2(prev_rpr + monthly_rpr)
            cum_cost = r2(prev_cost - monthly_cost)  # grows more negative

            pc_ratio = r4(cum_rpr / abs(cum_cost)) if cum_cost != 0 else None

            new_rpr_date[rep]  = cum_rpr
            new_cost_date[rep] = cum_cost
            new_pc_date[rep]   = pc_ratio

            # Month arrays: find last tenure month and advance
            last_m = get_last_month(rpr_month, rep)
            next_m = last_m + 1
            prev_rpr_m  = get_last_val(rpr_month, rep) or 0.0
            prev_cost_m = get_last_val(cost_month, rep) or 0.0
            cum_rpr_m  = r2(prev_rpr_m + monthly_rpr)
            cum_cost_m = r2(prev_cost_m - monthly_cost)
            pc_ratio_m = r4(cum_rpr_m / abs(cum_cost_m)) if cum_cost_m != 0 else None
            new_rpr_month[rep]  = (next_m, cum_rpr_m)
            new_cost_month[rep] = (next_m, cum_cost_m)
            new_pc_month[rep]   = (next_m, pc_ratio_m)

            # Auto-add new reps
            if rep not in reps_list:
                reps_list.append(rep)
                reps_list.sort()
                print(f"  + Added {rep} to {reps_key}")
            if rep not in active_list and rep not in known_reps:
                active_list.append(rep)
                active_list.sort()
                print(f"  + Added {rep} to {active_key}")

            processed_reps.add(rep)

        # Keep leave reps flat (carry forward their last RPR, no new cost)
        for rep, conf in leave_conf.items():
            if rep in processed_reps:
                continue
            prev_rpr = get_last_val(rpr_date, rep)
            prev_cost = get_last_val(cost_date, rep)
            if prev_rpr is not None:
                new_rpr_date[rep]  = prev_rpr   # flat
                new_cost_date[rep] = prev_cost  # flat (no cost during leave)
                new_pc_date[rep]   = r4(prev_rpr / abs(prev_cost)) if prev_cost else None
                # Month arrays: leave period doesn't advance tenure month
                print(f"  ~ {rep} on leave — held flat at {prev_rpr}")

        # Ensure new rep columns exist in all rows
        all_new_reps = list(processed_reps)
        ensure_rep_columns(rpr_date, all_new_reps)
        ensure_rep_columns(cost_date, all_new_reps)
        ensure_rep_columns(pc_date, all_new_reps)
        ensure_rep_columns(rpr_month, all_new_reps)
        ensure_rep_columns(cost_month, all_new_reps)
        ensure_rep_columns(pc_month, all_new_reps)

        # Apply date arrays
        add_date_row(rpr_date,  date_label, new_rpr_date)
        add_date_row(cost_date, date_label, new_cost_date)
        # Compute PC from updated RPR and CCOST
        pc_row_vals = {}
        for rep, rpr_v in new_rpr_date.items():
            cost_v = new_cost_date.get(rep)
            if cost_v and cost_v != 0:
                pc_row_vals[rep] = r4(rpr_v / abs(cost_v))
        add_date_row(pc_date, date_label, pc_row_vals)

        # Apply month arrays
        month_rpr_by_m  = {}
        month_cost_by_m = {}
        month_pc_by_m   = {}
        for rep, (m, val) in new_rpr_month.items():
            month_rpr_by_m.setdefault(m, {})[rep] = val
        for rep, (m, val) in new_cost_month.items():
            month_cost_by_m.setdefault(m, {})[rep] = val
        for rep, (m, val) in new_pc_month.items():
            month_pc_by_m.setdefault(m, {})[rep] = val

        for month_num, rep_vals in month_rpr_by_m.items():
            target = next((r for r in rpr_month if r["month"] == month_num), None)
            if target is None:
                new_row = {k: None for k in rpr_month[-1].keys()}
                new_row["month"] = month_num
                rpr_month.append(new_row)
                target = rpr_month[-1]
            for rep, val in rep_vals.items():
                target[rep] = val

        for month_num, rep_vals in month_cost_by_m.items():
            target = next((r for r in cost_month if r["month"] == month_num), None)
            if target is None:
                new_row = {k: None for k in cost_month[-1].keys()}
                new_row["month"] = month_num
                cost_month.append(new_row)
                target = cost_month[-1]
            for rep, val in rep_vals.items():
                target[rep] = val

        for month_num, rep_vals in month_pc_by_m.items():
            target = next((r for r in pc_month if r["month"] == month_num), None)
            if target is None:
                new_row = {k: None for k in pc_month[-1].keys()}
                new_row["month"] = month_num
                pc_month.append(new_row)
                target = pc_month[-1]
            for rep, val in rep_vals.items():
                target[rep] = val

        print(f"  OK {universe} {date_label} written to all arrays")

    DATA_JSON.write_text(json.dumps(data, indent=2), encoding="utf-8")
    print(f"\nSaved {DATA_JSON}")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python update_moneyball.py \"path/to/Monthly Excel.xlsx\"")
        sys.exit(1)
    update(sys.argv[1])
