"""
ytd_dashboard.py
-----------------
STANDALONE, INDEPENDENT SCRIPT -- imports NOTHING from the rest of this
project (no excel_utils.py, file_mgmt.py, dashboard.py, order_sheet.py,
etc.). Deliberately decoupled per Harish's explicit request: he is
actively iterating on multiple versions of the trading pipeline
(01_Master_Code.py, 02_Master_Code_3Indicator.py, and whatever comes
next) and does not want this reporting tool to break or need rework
every time that code changes underneath it.

WHAT THIS DOES
    Scans F:\05_Claude_Automation (top level only, no subfolders) for
    every "DD-Mon-YY FNO.xlsx"-style backtest output file, pulls each
    day's 'Orders' sheet (every closed trade -- BACKTEST mode never
    leaves a position open overnight, so every row already has an Exit
    Time/Reason), and builds/updates a single consolidated
    "YTD Dashboard.xlsx" in the same folder with:

        Trade Log  -- one row per individual trade, across every file,
                      raw Orders columns preserved as-is, with a
                      prepended Date column. The single source of truth
                      everything else below is computed from.
        Monthly    -- one row per TRADING DATE (Harish's own naming for
                      this sheet -- the granularity is daily; "Monthly"
                      refers to it being the source the Overall sheet's
                      monthly rollup is built FROM).
        Overall    -- a YTD overview block, followed by one row per
                      calendar month, both aggregated from the Monthly
                      (daily) sheet.
        _RunLog    -- hidden. {Date: source file's mtime last processed}.
                      Never shown to Harish, just internal state.

INCREMENTAL REFRESH
    Comparing "have I loaded this date before" is not enough -- Harish
    frequently re-runs the SAME date against different/updated pipeline
    code while iterating, and a naive "skip once ever seen" rule would
    silently freeze the first result forever. Instead, each date's
    _RunLog entry stores the source file's last-modified time: unchanged
    mtime -> skip (fast path, satisfies "don't reload everything every
    run"); new date OR changed mtime (a re-run with fresh numbers) ->
    that date's old Trade Log rows are deleted and replaced with
    freshly-read ones. Dates already logged are never forgotten even if
    the source .xlsx is later deleted from disk (this is meant to be a
    persistent YTD ledger, not a live mirror of the folder's contents).

HOW TO RUN
    python ytd_dashboard.py
    (Or the equivalent `py ytd_dashboard.py` on Harish's Windows setup.)
    Prints a summary of which dates were reprocessed vs. skipped.

Disclaimer: this is purely a reporting/aggregation tool over historical
backtest data already produced elsewhere. It contains no trading/signal
logic of its own, and none of the underlying numbers are a guarantee of
future performance.
"""

import os
import re
import sys
import traceback
from datetime import datetime

import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter

# ---------------------------------------------------------------------------
# Config -- same ALGO_BASE_DIR portability convention the rest of the
# project uses (see file_mgmt.py), reproduced here inline rather than
# imported, per the "no dependency on other codes" requirement.
# ---------------------------------------------------------------------------
BASE_DIR = os.environ.get("ALGO_BASE_DIR", r"F:\05_Claude_Automation")
YTD_FILENAME = "YTD Dashboard.xlsx"
YTD_PATH = os.path.join(BASE_DIR, YTD_FILENAME)

# Matches "08-Jul-26 FNO.xlsx" and also tolerates a trailing suffix before
# the extension (e.g. upload-dedup artifacts like "08-Jul-26 FNO-450ae647.xlsx")
# so the same script behaves the same way against either naming.
FNO_FILE_PATTERN = re.compile(r'^(\d{2}-[A-Za-z]{3}-\d{2})\s+FNO.*\.xlsx$', re.IGNORECASE)

ORDERS_SHEET = "Orders"
TRADE_LOG_SHEET = "Trade Log"
MONTHLY_SHEET = "Monthly"
OVERALL_SHEET = "Overall"
RUNLOG_SHEET = "_RunLog"

PL_COL = "Net P/L (Rs)"

# ---------------------------------------------------------------------------
# Styling (mirrors the green/red/gray + dark-header convention already
# used across this project's Dashboard/indicator sheets, reproduced here
# rather than imported)
# ---------------------------------------------------------------------------
FILL_TITLE = PatternFill(start_color="111827", end_color="111827", fill_type="solid")
FILL_SECTION = PatternFill(start_color="374151", end_color="374151", fill_type="solid")
FILL_HEADER = PatternFill(start_color="1F2937", end_color="1F2937", fill_type="solid")
FILL_GREEN = PatternFill(start_color="26A69A", end_color="26A69A", fill_type="solid")
FILL_RED = PatternFill(start_color="EF5350", end_color="EF5350", fill_type="solid")
FILL_GRAY = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
FONT_TITLE = Font(color="FFFFFF", bold=True, size=14)
FONT_SECTION = Font(color="FFFFFF", bold=True, size=11)
FONT_HEADER = Font(color="FFFFFF", bold=True)
FONT_BOLD = Font(bold=True)
FONT_WHITE_BOLD = Font(color="FFFFFF", bold=True)


def autofit_columns(ws, min_width=8, max_width=42, padding=2):
    widths = {}
    for row in ws.iter_rows():
        for cell in row:
            if cell.value is None:
                continue
            col = cell.column_letter
            length = len(str(cell.value))
            widths[col] = max(widths.get(col, min_width), min(length + padding, max_width))
    for col, w in widths.items():
        ws.column_dimensions[col].width = w


# ---------------------------------------------------------------------------
# Step 1: Discover candidate files on disk
# ---------------------------------------------------------------------------
def discover_fno_files(base_dir):
    """Returns {date_str ('DD-Mon-YY'): (full_path, mtime)} for every
    top-level file matching the FNO naming pattern. If more than one file
    matches the same date (shouldn't happen in normal use -- Harish
    deletes before re-running -- but guarded anyway), the most recently
    modified one wins."""
    found = {}
    if not os.path.isdir(base_dir):
        raise RuntimeError(f"Base directory not found: {base_dir}")
    for name in os.listdir(base_dir):
        full_path = os.path.join(base_dir, name)
        if not os.path.isfile(full_path):
            continue
        m = FNO_FILE_PATTERN.match(name)
        if not m:
            continue
        date_str = m.group(1)
        mtime = os.path.getmtime(full_path)
        if date_str not in found or mtime > found[date_str][1]:
            found[date_str] = (full_path, mtime)
    return found


def _parse_date(date_str):
    """'08-Jul-26' -> datetime.date(2026, 7, 8)."""
    return datetime.strptime(date_str, "%d-%b-%y").date()


def _date_str(d):
    return d.strftime("%d-%b-%y")


# ---------------------------------------------------------------------------
# Step 2: Read one day's Orders sheet
# ---------------------------------------------------------------------------
def read_orders_for_date(file_path, date_str):
    """Returns a DataFrame of that date's trades (raw Orders columns, plus
    a prepended 'Date' column), or an empty DataFrame if the sheet is
    missing/empty (zero-trade day, e.g. no signals fired) -- never raises
    for that reason alone."""
    try:
        df = pd.read_excel(file_path, sheet_name=ORDERS_SHEET)
    except Exception as e:
        print(f"    [INFO] {date_str}: no '{ORDERS_SHEET}' sheet found ({e}) -- zero-trade day.")
        return pd.DataFrame()

    if df.empty:
        return pd.DataFrame()

    df.insert(0, 'Date', _parse_date(date_str))
    return df


# ---------------------------------------------------------------------------
# Step 3: Load existing YTD workbook state (Trade Log + RunLog), if any
# ---------------------------------------------------------------------------
def load_existing_state(ytd_path):
    """Returns (trade_log_df, runlog_dict). Both empty/blank if the YTD
    file doesn't exist yet (first-ever run)."""
    if not os.path.exists(ytd_path):
        return pd.DataFrame(), {}

    try:
        trade_log_df = pd.read_excel(ytd_path, sheet_name=TRADE_LOG_SHEET)
        if 'Date' in trade_log_df.columns:
            trade_log_df['Date'] = pd.to_datetime(trade_log_df['Date']).dt.date
    except Exception:
        trade_log_df = pd.DataFrame()

    runlog = {}
    try:
        runlog_df = pd.read_excel(ytd_path, sheet_name=RUNLOG_SHEET)
        for _, row in runlog_df.iterrows():
            runlog[str(row['Date'])] = float(row['Last Processed Mtime'])
    except Exception:
        runlog = {}

    return trade_log_df, runlog


# ---------------------------------------------------------------------------
# Step 4: Incrementally refresh the Trade Log
# ---------------------------------------------------------------------------
def refresh_trade_log(base_dir, ytd_path):
    """Core incremental logic. Returns (trade_log_df, runlog, processed,
    skipped) -- runlog accumulates every date ever processed (old entries
    are carried forward untouched), so history persists even if a source
    .xlsx is later deleted from disk."""
    found_files = discover_fno_files(base_dir)
    trade_log_df, runlog = load_existing_state(ytd_path)

    processed, skipped = [], []
    new_frames = []

    for date_str, (file_path, mtime) in sorted(found_files.items(), key=lambda kv: _parse_date(kv[0])):
        last_mtime = runlog.get(date_str)
        if last_mtime is not None and abs(last_mtime - mtime) < 1e-6:
            skipped.append(date_str)
            continue

        print(f"  [PROCESS] {date_str}: reading '{ORDERS_SHEET}' from {os.path.basename(file_path)}...")
        day_df = read_orders_for_date(file_path, date_str)

        # Drop any existing rows for this date before re-inserting -- handles
        # both first-time load and a re-run with different results.
        if not trade_log_df.empty and 'Date' in trade_log_df.columns:
            trade_log_df = trade_log_df[trade_log_df['Date'] != _parse_date(date_str)]

        if not day_df.empty:
            new_frames.append(day_df)

        runlog[date_str] = mtime
        processed.append(date_str)

    if new_frames:
        trade_log_df = pd.concat([trade_log_df] + new_frames, ignore_index=True, sort=False)

    return trade_log_df, runlog, processed, skipped


# ---------------------------------------------------------------------------
# Step 5: Aggregation
# ---------------------------------------------------------------------------
def _pf(gross_win, gross_loss):
    if gross_loss > 0:
        return round(gross_win / gross_loss, 2)
    return 'Inf' if gross_win > 0 else ''


def build_daily_rollup(trade_log_df, all_dates):
    """One row per trading date (written to the 'Monthly' sheet, per
    Harish's naming -- see module docstring)."""
    rows = []
    for d in sorted(all_dates):
        day_trades = trade_log_df[trade_log_df['Date'] == d] if not trade_log_df.empty else pd.DataFrame()
        n = len(day_trades)
        if n == 0:
            rows.append({
                'Date': d, 'Trades': 0, 'Wins': 0, 'Losses': 0,
                'Win Rate (%)': '', 'Total Net P/L (Rs)': 0.0,
                'Profit Factor': '', 'Best Symbol': '', 'Worst Symbol': '',
            })
            continue
        pl = pd.to_numeric(day_trades[PL_COL], errors='coerce').fillna(0)
        wins = int((pl > 0).sum())
        losses = int((pl < 0).sum())
        gross_win = pl[pl > 0].sum()
        gross_loss = abs(pl[pl < 0].sum())
        best_idx = pl.idxmax()
        worst_idx = pl.idxmin()
        rows.append({
            'Date': d, 'Trades': n, 'Wins': wins, 'Losses': losses,
            'Win Rate (%)': round(wins / n * 100, 1),
            'Total Net P/L (Rs)': round(pl.sum(), 2),
            'Profit Factor': _pf(gross_win, gross_loss),
            'Best Symbol': day_trades.loc[best_idx, 'Symbol'],
            'Worst Symbol': day_trades.loc[worst_idx, 'Symbol'],
        })
    return pd.DataFrame(rows)


def build_monthly_rollup(trade_log_df, all_dates):
    """One row per calendar month -- recomputed from the underlying trades
    in that month (never by averaging daily ratios), embedded in the
    'Overall' sheet."""
    if not all_dates:
        return pd.DataFrame()
    month_of = {d: d.strftime('%b-%Y') for d in all_dates}
    months_sorted = sorted(set(month_of.values()), key=lambda m: datetime.strptime(m, '%b-%Y'))

    rows = []
    for month_label in months_sorted:
        month_dates = [d for d, m in month_of.items() if m == month_label]
        month_trades = trade_log_df[trade_log_df['Date'].isin(month_dates)] if not trade_log_df.empty else pd.DataFrame()
        n = len(month_trades)
        trading_days = len(month_dates)
        if n == 0:
            rows.append({
                'Month': month_label, 'Trading Days': trading_days, 'Trades': 0,
                'Wins': 0, 'Losses': 0, 'Win Rate (%)': '',
                'Net P/L (Rs)': 0.0, 'Profit Factor': '',
            })
            continue
        pl = pd.to_numeric(month_trades[PL_COL], errors='coerce').fillna(0)
        wins = int((pl > 0).sum())
        losses = int((pl < 0).sum())
        gross_win = pl[pl > 0].sum()
        gross_loss = abs(pl[pl < 0].sum())
        rows.append({
            'Month': month_label, 'Trading Days': trading_days, 'Trades': n,
            'Wins': wins, 'Losses': losses,
            'Win Rate (%)': round(wins / n * 100, 1),
            'Net P/L (Rs)': round(pl.sum(), 2),
            'Profit Factor': _pf(gross_win, gross_loss),
        })
    return pd.DataFrame(rows)


def compute_max_drawdown(trade_log_df):
    """Peak-to-trough decline on the chronological (Date + Exit Time)
    cumulative Net P/L curve across the full Trade Log."""
    if trade_log_df.empty or PL_COL not in trade_log_df.columns:
        return 0.0
    df = trade_log_df.copy()
    exit_time = df['Exit Time'].astype(str) if 'Exit Time' in df.columns else ''
    df['_sort_key'] = pd.to_datetime(df['Date'].astype(str) + ' ' + exit_time, errors='coerce')
    df = df.sort_values('_sort_key')
    cum = pd.to_numeric(df[PL_COL], errors='coerce').fillna(0).cumsum()
    running_max = cum.cummax()
    drawdown = running_max - cum
    return round(drawdown.max(), 2) if not drawdown.empty else 0.0


# ---------------------------------------------------------------------------
# Step 6: Write sheets
# ---------------------------------------------------------------------------
def write_trade_log_sheet(wb, trade_log_df):
    if TRADE_LOG_SHEET in wb.sheetnames:
        del wb[TRADE_LOG_SHEET]
    ws = wb.create_sheet(TRADE_LOG_SHEET)

    if trade_log_df.empty:
        ws['A1'] = "No trades recorded yet."
        ws['A1'].font = FONT_BOLD
        return

    df = trade_log_df.sort_values(['Date'], kind='stable') if 'Date' in trade_log_df.columns else trade_log_df
    cols = list(df.columns)
    ws.append(cols)
    for cell in ws[1]:
        cell.fill = FILL_HEADER
        cell.font = FONT_HEADER

    for _, row in df.iterrows():
        ws.append([row[c] for c in cols])

    if PL_COL in cols:
        pl_idx = cols.index(PL_COL) + 1
        for r in range(2, ws.max_row + 1):
            cell = ws.cell(row=r, column=pl_idx)
            if isinstance(cell.value, (int, float)):
                if cell.value > 0:
                    cell.fill = FILL_GREEN
                elif cell.value < 0:
                    cell.fill = FILL_RED

    ws.freeze_panes = "A2"
    autofit_columns(ws)


def write_daily_sheet(wb, daily_df):
    """Written to the sheet named 'Monthly' -- see module docstring for
    why this daily-granularity table has that name."""
    if MONTHLY_SHEET in wb.sheetnames:
        del wb[MONTHLY_SHEET]
    ws = wb.create_sheet(MONTHLY_SHEET)

    if daily_df.empty:
        ws['A1'] = "No trading days recorded yet."
        ws['A1'].font = FONT_BOLD
        return

    cols = list(daily_df.columns)
    ws.append(cols)
    for cell in ws[1]:
        cell.fill = FILL_HEADER
        cell.font = FONT_HEADER

    pl_idx = cols.index('Total Net P/L (Rs)') + 1
    for _, row in daily_df.iterrows():
        ws.append([row[c] for c in cols])
        r = ws.max_row
        val = ws.cell(row=r, column=pl_idx).value
        if isinstance(val, (int, float)) and val != 0:
            ws.cell(row=r, column=pl_idx).fill = FILL_GREEN if val > 0 else FILL_RED

    ws.freeze_panes = "A2"
    autofit_columns(ws)


def write_overall_sheet(wb, trade_log_df, monthly_df, all_dates):
    if OVERALL_SHEET in wb.sheetnames:
        del wb[OVERALL_SHEET]
    ws = wb.create_sheet(OVERALL_SHEET, 0)  # first tab

    ws.merge_cells('B1:F1')
    ws['B1'] = f"YTD F&O TRADING DASHBOARD -- generated {datetime.now().strftime('%d %b %Y, %H:%M')} IST"
    ws['B1'].font = FONT_TITLE
    ws['B1'].fill = FILL_TITLE
    for col in range(2, 7):
        ws.cell(row=1, column=col).fill = FILL_TITLE

    total_trades = len(trade_log_df)
    total_days = len(all_dates)
    days_with_trades = trade_log_df['Date'].nunique() if not trade_log_df.empty and 'Date' in trade_log_df.columns else 0

    if total_trades:
        pl = pd.to_numeric(trade_log_df[PL_COL], errors='coerce').fillna(0)
        wins = int((pl > 0).sum())
        losses = int((pl < 0).sum())
        win_rate = round(wins / total_trades * 100, 1)
        net_pl = round(pl.sum(), 2)
        gross_win = pl[pl > 0].sum()
        gross_loss = abs(pl[pl < 0].sum())
        pf = _pf(gross_win, gross_loss)
        avg_win = round(pl[pl > 0].mean(), 2) if wins else 0
        avg_loss = round(pl[pl < 0].mean(), 2) if losses else 0
        max_dd = compute_max_drawdown(trade_log_df)
    else:
        wins = losses = 0
        win_rate = net_pl = pf = avg_win = avg_loss = max_dd = 0

    best_month = worst_month = ('', '')
    if monthly_df is not None and not monthly_df.empty:
        traded_months = monthly_df[monthly_df['Trades'] > 0]
        if not traded_months.empty:
            best_row = traded_months.loc[traded_months['Net P/L (Rs)'].idxmax()]
            worst_row = traded_months.loc[traded_months['Net P/L (Rs)'].idxmin()]
            best_month = (best_row['Month'], best_row['Net P/L (Rs)'])
            worst_month = (worst_row['Month'], worst_row['Net P/L (Rs)'])

    row = 3
    ws.cell(row=row, column=2, value="YTD OVERVIEW")
    ws.cell(row=row, column=2).font = FONT_SECTION
    ws.cell(row=row, column=2).fill = FILL_SECTION
    for col in range(2, 7):
        ws.cell(row=row, column=col).fill = FILL_SECTION
    row += 1

    stats = [
        ("Total Trading Days Logged", total_days),
        ("Days With At Least 1 Trade", days_with_trades),
        ("Total Trades", total_trades),
        ("Wins / Losses", f"{wins} / {losses}"),
        ("Overall Win Rate", f"{win_rate}%" if total_trades else "N/A"),
        ("Overall Profit Factor", pf if total_trades else "N/A"),
        ("Total Net P/L (Rs)", net_pl),
        ("Avg Profit per Win (Rs)", avg_win),
        ("Avg Loss per Loss (Rs)", avg_loss),
        ("Max Drawdown (Rs)", max_dd),
        ("Best Month", f"{best_month[0]} ({best_month[1]})" if best_month[0] else "N/A"),
        ("Worst Month", f"{worst_month[0]} ({worst_month[1]})" if worst_month[0] else "N/A"),
    ]
    for label, value in stats:
        ws.cell(row=row, column=2, value=label).font = FONT_BOLD
        cell = ws.cell(row=row, column=3, value=value)
        if label == "Total Net P/L (Rs)" and isinstance(value, (int, float)):
            cell.fill = FILL_GREEN if value > 0 else (FILL_RED if value < 0 else FILL_GRAY)
        row += 1

    row += 1
    ws.cell(row=row, column=2, value="MONTHLY P&L")
    ws.cell(row=row, column=2).font = FONT_SECTION
    ws.cell(row=row, column=2).fill = FILL_SECTION
    for col in range(2, 7):
        ws.cell(row=row, column=col).fill = FILL_SECTION
    row += 1

    if monthly_df is None or monthly_df.empty:
        ws.cell(row=row, column=2, value="No trading months recorded yet.")
    else:
        header_row = row
        for c_idx, col_name in enumerate(monthly_df.columns, start=2):
            cell = ws.cell(row=header_row, column=c_idx, value=col_name)
            cell.fill = FILL_HEADER
            cell.font = FONT_HEADER
        row += 1
        pl_col_pos = list(monthly_df.columns).index('Net P/L (Rs)') + 2
        for _, m_row in monthly_df.iterrows():
            for c_idx, col_name in enumerate(monthly_df.columns, start=2):
                ws.cell(row=row, column=c_idx, value=m_row[col_name])
            val = ws.cell(row=row, column=pl_col_pos).value
            if isinstance(val, (int, float)) and val != 0:
                ws.cell(row=row, column=pl_col_pos).fill = FILL_GREEN if val > 0 else FILL_RED
            row += 1

    ws.column_dimensions['A'].width = 3
    autofit_columns(ws)


def write_runlog_sheet(wb, runlog):
    if RUNLOG_SHEET in wb.sheetnames:
        del wb[RUNLOG_SHEET]
    ws = wb.create_sheet(RUNLOG_SHEET)
    ws.sheet_state = 'hidden'
    ws.append(['Date', 'Last Processed Mtime'])
    for date_str, mtime in sorted(runlog.items(), key=lambda kv: _parse_date(kv[0])):
        ws.append([date_str, mtime])


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
def main():
    print("=" * 60)
    print("   YTD DASHBOARD -- standalone consolidator")
    print("=" * 60)
    print(f"[SYSTEM] Base directory: {BASE_DIR}")
    print(f"[SYSTEM] Output file:    {YTD_PATH}")

    try:
        trade_log_df, runlog, processed, skipped = refresh_trade_log(BASE_DIR, YTD_PATH)
    except Exception as e:
        print(f"\n[FATAL] Failed to refresh Trade Log: {e}")
        print(traceback.format_exc())
        sys.exit(1)

    print("-" * 60)
    print(f"[SUMMARY] {len(processed)} date(s) (re)processed, {len(skipped)} date(s) unchanged/skipped.")
    if processed:
        print(f"          Processed: {', '.join(processed)}")
    if skipped:
        print(f"          Skipped:   {', '.join(skipped)}")

    all_dates = sorted({_parse_date(ds) for ds in runlog.keys()})
    if not all_dates:
        print("\n[WARNING] No FNO.xlsx files found in the base directory -- nothing to consolidate yet.")

    daily_df = build_daily_rollup(trade_log_df, all_dates)
    monthly_df = build_monthly_rollup(trade_log_df, all_dates)

    try:
        if os.path.exists(YTD_PATH):
            wb = load_workbook(YTD_PATH)
        else:
            wb = Workbook()
            wb.remove(wb.active)

        write_trade_log_sheet(wb, trade_log_df)
        write_daily_sheet(wb, daily_df)
        write_overall_sheet(wb, trade_log_df, monthly_df, all_dates)
        write_runlog_sheet(wb, runlog)

        wb.save(YTD_PATH)
        print(f"\n[SUCCESS] {YTD_PATH} updated.")
    except Exception as e:
        print(f"\n[FATAL] Failed to write {YTD_PATH}: {e}")
        print(traceback.format_exc())
        sys.exit(1)

    print("=" * 60)


if __name__ == "__main__":
    main()


# ---------------------------------------------------------------------------
# Disclaimer: this is purely a reporting/aggregation tool over historical
# backtest data already produced elsewhere. It contains no trading/signal
# logic of its own, and none of the underlying numbers are a guarantee of
# future performance.
# ---------------------------------------------------------------------------
