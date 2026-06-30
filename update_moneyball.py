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

def get_last_val(arr, rep, exclude_date=None):
    """Get the last non-null value for a rep in a date/month array.

    When exclude_date is given, rows whose "date" equals it are skipped. This
    makes re-running an already-written month idempotent: the "previous"
    cumulative is read from the prior month, not from the row being overwritten
    in place (which would double-count the increment).
    """
    val = None
    for row in arr:
        if exclude_date is not None and row.get("date") == exclude_date:
            continue
        v = row.get(rep)
        if v is not None:
            val = v
    return val

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

def rebuild_month_arrays(data, universe):
    """Regenerate the MONTH (tenure) arrays from the authoritative DATE (calendar)
    arrays. For each rep, their contiguous calendar months (first..last non-null
    RPR) are reindexed as tenure months 0..N. All three metrics (RPR/PC/CCOST)
    share the same per-rep spine, so they stay perfectly consistent and never
    develop the index-skip / shifted-start / extra-point artifacts that the old
    incremental approach accumulated. Real gaps (leave / inactive months) are
    preserved as null, exactly as they appear in the calendar arrays.
    """
    d_rpr = data[f"{universe}_DATE_RPR"]
    d_pc  = data[f"{universe}_DATE_PC"]
    d_cc  = data[f"{universe}_DATE_CCOST"]

    # preserve first-seen rep order across the date rows
    reps, seen = [], set()
    for row in d_rpr:
        for k in row:
            if k != "date" and k not in seen:
                seen.add(k); reps.append(k)

    # per-rep spine = [first non-null RPR row .. last non-null RPR row]
    spine, max_t = {}, -1
    for rep in reps:
        first = last = -1
        for i, row in enumerate(d_rpr):
            if row.get(rep) is not None:
                if first < 0:
                    first = i
                last = i
        if first < 0:
            continue
        spine[rep] = (first, last)
        max_t = max(max_t, last - first)

    def build(date_arr):
        rows = []
        for k in range(max_t + 1):
            row = {"month": k}
            for rep in reps:
                sp = spine.get(rep)
                row[rep] = date_arr[sp[0] + k].get(rep) if (sp and k <= sp[1] - sp[0]) else None
            rows.append(row)
        return rows

    data[f"{universe}_MONTH_RPR"]   = build(d_rpr)
    data[f"{universe}_MONTH_PC"]    = build(d_pc)
    data[f"{universe}_MONTH_CCOST"] = build(d_cc)

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
        # All rows in one sheet must belong to the same month, and the label must
        # be a valid Mon-YY — otherwise add_date_row would append a malformed/orphan row.
        labels = {r["month_label"] for r in excel_rows}
        if len(labels) != 1:
            raise ValueError(f"{universe}: Excel spans multiple months {sorted(labels)}; expected one.")
        if not re.match(r"^[A-Z][a-z]{2}-\d{2}$", date_label):
            raise ValueError(f"{universe}: bad month label {date_label!r}; expected 'Mon-YY' (e.g. Apr-26).")
        print(f"\n{universe} — processing {date_label} ({len(excel_rows)} reps)")

        rpr_date  = data[rpr_key_date]
        cost_date  = data[cost_key_date]
        pc_date   = data[pc_key_date]
        active_list = data[active_key]
        reps_list   = data[reps_key]
        known_reps  = set(data.get(known_key, []))
        leave_conf  = LEAVE_CONFIG.get(universe, {})

        new_rpr_date  = {}
        new_cost_date  = {}

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

            prev_rpr  = get_last_val(rpr_date, rep, exclude_date=date_label) or 0.0
            prev_cost = get_last_val(cost_date, rep, exclude_date=date_label) or 0.0  # negative cumulative

            cum_rpr  = r2(prev_rpr + monthly_rpr)
            cum_cost = r2(prev_cost - monthly_cost)  # grows more negative

            new_rpr_date[rep]  = cum_rpr
            new_cost_date[rep] = cum_cost
            # PC is recomputed authoritatively from RPR/CCOST after the date rows
            # are written (see pc_row_vals below); no need to stash it here.

            # (Month/tenure arrays are regenerated from the date arrays below,
            #  after all date rows are written — see rebuild_month_arrays.)

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
            prev_rpr = get_last_val(rpr_date, rep, exclude_date=date_label)
            prev_cost = get_last_val(cost_date, rep, exclude_date=date_label)
            if prev_rpr is not None:
                new_rpr_date[rep]  = prev_rpr   # flat
                new_cost_date[rep] = prev_cost  # flat (no cost during leave)
                # PC recomputed from RPR/CCOST below (pc_row_vals)
                # Month arrays: leave period doesn't advance tenure month
                print(f"  ~ {rep} on leave — held flat at {prev_rpr}")

        # Ensure new rep columns exist in all date rows
        # (month arrays are rebuilt fresh below, so they don't need this)
        all_new_reps = list(processed_reps)
        ensure_rep_columns(rpr_date, all_new_reps)
        ensure_rep_columns(cost_date, all_new_reps)
        ensure_rep_columns(pc_date, all_new_reps)

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

        # Regenerate the tenure-month arrays from the freshly-updated date arrays.
        # This keeps MONTH_* perfectly in sync with DATE_* every run and prevents
        # the index-skip / shifted / extra-point drift the old approach caused.
        rebuild_month_arrays(data, universe)

        print(f"  OK {universe} {date_label} written to all arrays")

    DATA_JSON.write_text(json.dumps(data, indent=2), encoding="utf-8")
    print(f"\nSaved {DATA_JSON}")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python update_moneyball.py \"path/to/Monthly Excel.xlsx\"")
        sys.exit(1)
    update(sys.argv[1])
