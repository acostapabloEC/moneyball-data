# moneyball-data

Single source of truth for all three Moneyball dashboards:
- [moneyball-gray.vercel.app](https://moneyball-gray.vercel.app) — Cumulative P&L
- [bda-payback.vercel.app](https://bda-payback.vercel.app) — Gross Payback Multiple
- [bda-projection.vercel.app](https://bda-projection.vercel.app) — Projection

## Monthly update workflow

1. Receive the monthly Excel file from Gene/John
2. Run the update script:
   ```
   python update_moneyball.py "June 2026 Moneyball.xlsx"
   ```
3. Review the diff in `data.json`
4. Commit and push — all three dashboards update automatically

## Leave configuration

Edit `LEAVE_CONFIG` at the top of `update_moneyball.py` to manage reps on unpaid leave.
During leave: RPR held flat, no cost accrual.

## Files

| File | Purpose |
|------|---------|
| `data.json` | All dashboard data — **the only file you edit via the Python script** |
| `update_moneyball.py` | Monthly updater — reads Excel, writes data.json |
| `extract_data.cjs` | One-time extractor — used to seed data.json from the old JS files |
| `strip_jun26.cjs` | Dev utility — strips a month for testing |
