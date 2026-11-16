# lte-3g-kpi-degradation-gui

Tkinter desktop GUI for spotting degraded cells and sites across 3G and LTE by comparing a chosen KPI's current value against the same weekday over the past 2 or 4 weeks.

## Problem

When a KPI dips or spikes on a given day, the question is always "is this normal for this day of the week, or is something actually wrong?" Pulling that comparison by hand — same weekday, several weeks back, for every WBTS/WCEL (3G) or LNBTS/LNCEL (LTE) in an Excel export — is slow and easy to get wrong when repeated for every KPI and every investigation. This tool turns it into a few clicks.

## Approach

- Load a Nokia KPI export (Excel, with a `Period start time` column and one row per object per period) through a file picker; the KPI dropdown auto-populates from the file's own column headers, so there's no need to hardcode which KPIs exist.
- Pick a start date, a lookback window (14 or 30 days — i.e. the same weekday for the last 2 or 4 weeks), a technology (3G or LTE), and a granularity (Site or Cell).
- Pick a comparison operator depending on which direction means degradation for that KPI: "current > historical" for a KPI where lower is better (e.g. a drop rate), or "current < historical" for a KPI where higher is better (e.g. a success rate).
- For every object, the tool compares the current value against each of the historical same-weekday values and flags it as degraded only when the comparison holds against all of them — a consistent, not one-off, move.
- Results are pivoted into an Excel sheet with one column per historical date (degradation flag + numeric difference) plus the min/max difference across all compared dates, so the worst-case movement for each object is visible at a glance.

## Tech stack

Python · pandas · NumPy · tkinter · openpyxl

## How to run

```bash
pip install -r requirements.txt
python main.py
```

This opens the GUI. Steps:

1. Browse to the input Excel file (KPI dropdown populates automatically from its columns).
2. Browse to where the output Excel file should be saved.
3. Enter the time period: `14` or `30`.
4. Enter the start date (`YYYY-MM-DD`) to compare against.
5. Pick the KPI from the dropdown.
6. Pick the comparison operator matching the KPI's "good" direction.
7. Pick the technology (3G or LTE).
8. Check Site or Cell for the granularity.
9. Click "Generate Excel".

## Project history

Originally written as a set of near-identical per-technology, per-granularity functions (3G site, 3G cell, LTE site, LTE cell) wrapped behind one GUI, since each combination needed a slightly different grouping column (`WBTS name`, `WCEL name`, `LNBTS name`, `LNCEL name`) but otherwise followed the same comparison logic. Kept as-is here since the duplication is minor and the GUI dispatch (`generate_excel`) keeps the user-facing flow identical regardless of technology.

## Note on data

No input data or company-specific file paths are included. All file locations are chosen at runtime through the GUI's file pickers.
