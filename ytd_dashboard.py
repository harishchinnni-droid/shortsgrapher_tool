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

VISUAL STYLE
    [CHANGED -- 18-Jul-26, Harish's feedback: "very boring"] Reuses the
    same navy/section-blue/table-head-blue palette, zebra striping,
    borders, and threshold-based green/red KPI coloring as dashboard.py's
    per-day 'Dashboard' sheet (see that module's style_dashboard_sheet()),
    reproduced here as local color constants/helpers rather than
    imported, so both files LOOK like one product without this one
    depending on that one's code.

CAPITAL DEPLOYED
    [ADDED -- 18-Jul-26, Harish's request] Same formula dashboard.py uses
    for its own 'Total Capital Deployed (Rs)' KPI: sum(Entry LTP x
    Quantity (Units)) across a day's trades. The 'Monthly' (daily) sheet
    now shows this per day; the 'Overall' sheet's Monthly P&L table shows
    the MIN and MAX of those daily figures within each calendar month, so
    Harish can see how much his capital usage swung day-to-day within a
    month, not just the P&L outcome.

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
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
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
# Styling -- same palette as dashboard.py's per-day 'Dashboard' sheet
# (NAVY/SECTION_BLUE/TABLE_HEAD_BLUE/GREEN/RED/zebra banding), reproduced
# locally so this file stays import-free from the rest of the project.
# ---------------------------------------------------------------------------
NAVY = "1F3864"
SECTION_BLUE = "2F5596"
TABLE_HEAD_BLUE = "4472C4"
GREEN = "1E9E4C"
RED = "D33B2C"
BAND_LIGHT = "FFFFFF"
BAND_DARK = "F2F2F2"
BORDER_COLOR = "D9D9D9"

FILL_NAVY = PatternFill(start_color=NAVY, end_color=NAVY, fill_type="solid")
FILL_SECTION = PatternFill(start_color=SECTION_BLUE, end_color=SECTION_BLUE, fill_type="solid")
FILL_TABLE_HEAD = PatternFill(start_color=TABLE_HEAD_BLUE, end_color=TABLE_HEAD_BLUE, fill_type="solid")
FILL_GREEN = PatternFill(start_color=GREEN, end_color=GREEN, fill_type="solid")
FILL_RED = PatternFill(start_color=RED, end_color=RED, fill_type="solid")

FONT_TITLE = Font(color="FFFFFF", bold=True, size=14)
FONT_SECTION = Font(color="FFFFFF", bold=True, size=11)
FONT_TABLE_HEAD = Font(color="FFFFFF", bold=True)
FONT_VALUE_ON_FILL = Font(color="FFFFFF", bold=True)
FONT_GREEN_TEXT = Font(color=GREEN, bold=True)
FONT_RED_TEXT = Font(color=RED, bold=True)
FONT_BOLD = Font(bold=True)
FONT_LABEL = Font(bold=True)

THIN = Side(style="thin", color=BORDER_COLOR)
CELL_BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
LEFT = Alignment(horizontal="left", vertical="center")
RIGHT = Alignment(horizontal="right", vertical="center")
CENTER = Alignment(horizontal="center", vertical="center")

# Keyword-based coloring for KPI label/value pairs (Overall sheet's
# YTD OVERVIEW panel) -- same idea as dashboard.py's _classify_kpi_label(),
# tuned for this sheet's own label set.
THRESHOLD_RULES = (
    ("win rate", lambda v: v >= 50),
    ("profit factor", lambda v: v >= 1),
    ("net p/l", lambda v: v >= 0),
)
GOOD_LABEL_KEYWORDS = ("avg profit", "best month")
BAD_LABEL_KEYWORDS = ("avg loss", "max drawdown", "worst month")


def _classify_kpi_label(label_text):
    label = str(label_text).strip().lower()
    for keyword, is_good_fn in THRESHOLD_RULES:
        if keyword in label:
            return "__threshold__", is_good_fn
    if any(k in label for k in GOOD_LABEL_KEYWORDS):
        return "good", None
    if any(k in label for k in BAD_LABEL_KEYWORDS):
        return "bad", None
    return None, None


def _parse_numeric(value):
    if value is None or value == "":
        return None
    if isinstance(value, (int, float)):
        return float(value)
    s = str(value).strip().replace("Rs.", "").replace("Rs", "").replace(",", "").replace("%", "").strip()
    try:
        return float(s)
    except ValueError:
        return None


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


def style_title_bar(ws, title_text, last_col_letter, row=1):
    ws.merge_cells(f"B{row}:{last_col_letter}{row}")
    cell = ws.cell(row=row, column=2, value=title_text)
    cell.font = FONT_TITLE
    cell.alignment = CENTER
    for c in range(2, get_column_letter_to_index(last_col_letter) + 1):
        ws.cell(row=row, column=c).fill = FILL_NAVY
    ws.row_dimensions[row].height = 28


def get_column_letter_to_index(letter):
    from openpyxl.utils import column_index_from_string
    return column_index_from_string(letter)


def style_section_header(ws, row, start_col, end_col, text):
    cell = ws.cell(row=row, column=start_col, value=text)
    if end_col > start_col:
        ws.merge_cells(start_row=row, start_column=start_col, end_row=row, end_column=end_col)
    for c in range(start_col, end_col + 1):
        ws.cell(row=row, column=c).fill = FILL_SECTION
    cell.font = FONT_SECTION
    cell.alignment = LEFT
    ws.row_dimensions[row].height = 20


def style_kpi_rows(ws, start_row, end_row, label_col, value_col):
    """Zebra-striped label/value panel with threshold-based green/red
    value coloring -- mirrors dashboard.py's _style_kpi_panel()."""
    for i, r in enumerate(range(start_row, end_row + 1)):
        label_cell = ws.cell(row=r, column=label_col)
        value_cell = ws.cell(row=r, column=value_col)
        if label_cell.value in (None, ""):
            continue
        band = BAND_LIGHT if i % 2 == 0 else BAND_DARK
        band_fill = PatternFill(start_color=band, end_color=band, fill_type="solid")
        for c in range(label_col, value_col + 1):
            ws.cell(row=r, column=c).fill = band_fill
            ws.cell(row=r, column=c).border = CELL_BORDER
        label_cell.font = FONT_LABEL
        label_cell.alignment = LEFT

        classification, threshold_fn = _classify_kpi_label(label_cell.value)
        numeric_val = _parse_numeric(value_cell.value)
        fill = None
        if classification == "__threshold__" and numeric_val is not None:
            fill = FILL_GREEN if threshold_fn(numeric_val) else FILL_RED
        elif classification == "good":
            fill = FILL_GREEN
        elif classification == "bad":
            fill = FILL_RED

        if fill:
            value_cell.fill = fill
            value_cell.font = FONT_VALUE_ON_FILL
        else:
            value_cell.font = Font(color="000000", bold=True)
        value_cell.alignment = RIGHT


def style_wide_table(ws, header_row, start_col, end_col, pl_col_names=()):
    """Blue header row + zebra-striped body, with green/red text on any
    column named in pl_col_names -- mirrors dashboard.py's
    _style_wide_table()."""
    pl_cols = set()
    for c in range(start_col, end_col + 1):
        cell = ws.cell(row=header_row, column=c)
        cell.fill = FILL_TABLE_HEAD
        cell.font = FONT_TABLE_HEAD
        cell.alignment = CENTER
        cell.border = CELL_BORDER
        if str(cell.value).strip() in pl_col_names:
            pl_cols.add(c)
    ws.row_dimensions[header_row].height = 18

    band_idx = 0
    for r in range(header_row + 1, ws.max_row + 1):
        row_vals = [ws.cell(row=r, column=c).value for c in range(start_col, end_col + 1)]
        if all(v in (None, "") for v in row_vals):
            continue
        band = BAND_LIGHT if band_idx % 2 == 0 else BAND_DARK
        band_fill = PatternFill(start_color=band, end_color=band, fill_type="solid")
        band_idx += 1
        for c in range(start_col, end_col + 1):
            cell = ws.cell(row=r, column=c)
            cell.fill = band_fill
            cell.border = CELL_BORDER
            if c in pl_cols:
                numeric_val = _parse_numeric(cell.value)
                cell.font = FONT_GREEN_TEXT if (numeric_val is None or numeric_val >= 0) else FONT_RED_TEXT
                cell.alignment = RIGHT
            elif isinstance(cell.value, (int, float)):
                cell.font = Font(color="000000")
                cell.alignment = RIGHT
            else:
                cell.font = Font(color="000000")
                cell.alignment = LEFT if c == start_col else CENTER


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


def _capital_deployed(trades_df):
    """Same formula as dashboard.py's 'Total Capital Deployed (Rs)':
    sum(Entry LTP x Quantity (Units)) across the given trades."""
    if trades_df.empty:
        return 0.0
    entry_ltp = pd.to_numeric(trades_df.get('Entry LTP', pd.Series(dtype=float)), errors='coerce').fillna(0)
    qty = pd.to_numeric(trades_df.get('Quantity (Units)', pd.Series(dtype=float)), errors='coerce').fillna(0)
    return round((entry_ltp * qty).sum(), 2)


def build_daily_rollup(trade_log_df, all_dates):
    """One row per trading date (written to the 'Monthly' sheet, per
    Harish's naming -- see module docstring)."""
    rows = []
    for d in sorted(all_dates):
        day_trades = trade_log_df[trade_log_df['Date'] == d] if not trade_log_df.empty else pd.DataFrame()
        n = len(day_trades)
        capital_deployed = _capital_deployed(day_trades)
        if n == 0:
            rows.append({
                'Date': d, 'Trades': 0, 'Wins': 0, 'Losses': 0,
                'Win Rate (%)': '', 'Total Net P/L (Rs)': 0.0,
                'Profit Factor': '', 'Capital Deployed (Rs)': capital_deployed,
                'Best Symbol': '', 'Worst Symbol': '',
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
            'Capital Deployed (Rs)': capital_deployed,
            'Best Symbol': day_trades.loc[best_idx, 'Symbol'],
            'Worst Symbol': day_trades.loc[worst_idx, 'Symbol'],
        })
    return pd.DataFrame(rows)


def build_monthly_rollup(trade_log_df, daily_df, all_dates):
    """One row per calendar month -- P&L/win-rate figures recomputed from
    the underlying trades in that month (never by averaging daily
    ratios); Min/Max Capital Deployed pulled from the daily_df's already-
    computed per-day figures. Embedded in the 'Overall' sheet."""
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

        month_daily = daily_df[daily_df['Date'].isin(month_dates)] if daily_df is not None and not daily_df.empty else pd.DataFrame()
        if not month_daily.empty and 'Capital Deployed (Rs)' in month_daily.columns:
            min_cap = round(month_daily['Capital Deployed (Rs)'].min(), 2)
            max_cap = round(month_daily['Capital Deployed (Rs)'].max(), 2)
        else:
            min_cap = max_cap = 0.0

        if n == 0:
            rows.append({
                'Month': month_label, 'Trading Days': trading_days, 'Trades': 0,
                'Wins': 0, 'Losses': 0, 'Win Rate (%)': '',
                'Net P/L (Rs)': 0.0, 'Profit Factor': '',
                'Min Capital Deployed (Rs)': min_cap, 'Max Capital Deployed (Rs)': max_cap,
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
            'Min Capital Deployed (Rs)': min_cap, 'Max Capital Deployed (Rs)': max_cap,
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
    for _, row in df.iterrows():
        ws.append([row[c] for c in cols])

    style_wide_table(ws, header_row=1, start_col=1, end_col=len(cols), pl_col_names={PL_COL})
    ws.sheet_view.showGridLines = False
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
    for _, row in daily_df.iterrows():
        ws.append([row[c] for c in cols])

    style_wide_table(ws, header_row=1, start_col=1, end_col=len(cols), pl_col_names={'Total Net P/L (Rs)'})
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "A2"
    autofit_columns(ws)


def write_overall_sheet(wb, trade_log_df, monthly_df, all_dates):
    if OVERALL_SHEET in wb.sheetnames:
        del wb[OVERALL_SHEET]
    ws = wb.create_sheet(OVERALL_SHEET, 0)  # first tab

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

    last_col = 4  # B..D used by the KPI panel; monthly table may extend further
    if monthly_df is not None and not monthly_df.empty:
        last_col = max(last_col, 1 + len(monthly_df.columns))

    style_title_bar(ws, f"YTD F&O TRADING DASHBOARD  --  generated {datetime.now().strftime('%d %b %Y, %H:%M')} IST",
                     get_column_letter(last_col))

    row = 3
    style_section_header(ws, row, 2, last_col, "YTD OVERVIEW")
    row += 1
    kpi_start = row

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
        ws.cell(row=row, column=2, value=label)
        ws.cell(row=row, column=3, value=value)
        row += 1
    kpi_end = row - 1
    style_kpi_rows(ws, kpi_start, kpi_end, label_col=2, value_col=3)

    row += 1
    style_section_header(ws, row, 2, last_col, "MONTHLY P&L")
    row += 1

    if monthly_df is None or monthly_df.empty:
        ws.cell(row=row, column=2, value="No trading months recorded yet.")
    else:
        header_row = row
        for c_idx, col_name in enumerate(monthly_df.columns, start=2):
            ws.cell(row=header_row, column=c_idx, value=col_name)
        row += 1
        for _, m_row in monthly_df.iterrows():
            for c_idx, col_name in enumerate(monthly_df.columns, start=2):
                ws.cell(row=row, column=c_idx, value=m_row[col_name])
            row += 1
        style_wide_table(ws, header_row=header_row, start_col=2, end_col=1 + len(monthly_df.columns),
                          pl_col_names={'Net P/L (Rs)'})

    ws.column_dimensions['A'].width = 3
    ws.sheet_view.showGridLines = False
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
    monthly_df = build_monthly_rollup(trade_log_df, daily_df, all_dates)

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
