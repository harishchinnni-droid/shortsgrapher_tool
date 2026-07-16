"""
excel_utils.py
--------------
Small shared helpers used by every indicator module (SQZMOM, RSI, BRKPRO,
ADX, TW ALL, HTF Bias, EMA 20, VWAP, OBV CMF) so the autofit/save/
date-scoping logic lives in one place instead of being copy-pasted across
every module.

[ADDED] restrict_to_target_date() -- see its own docstring. This is the
fix for the "workbook shows a full day of candle data before the market
has even opened" bug: every indicator's Symbol x Time matrix pivot used to
dedupe purely on time-of-day ('09:15', '09:20', ...) with no date
component, keeping the LAST occurrence of each slot across up to 90 days
of cached history. Before target_date's own candle for a given slot
exists, "last occurrence" silently fell back to the most recent PRIOR
trading day's value for that same clock time -- which is how the sheet
could show 73/73 time columns fully populated even pre-market. Every
indicator module now calls this right before returning from
process_symbol(), so only target_date's own rows ever reach the matrix.
"""

import os
import tempfile
import shutil

import pandas as pd


def autofit_columns(ws, min_width=6, max_width=40, padding=2):
    """Approximate Excel's 'AutoFit Column Width'.

    openpyxl has no native autofit -- real autofit only happens inside
    Excel's own rendering engine, which knows the actual font metrics. This
    sizes each column to its longest cell's string length instead, which is
    what every openpyxl-based workaround for this does. Widths are clamped
    to [min_width, max_width] so one long outlier symbol/value can't blow a
    column out to an unusable size.
    """
    widths = {}
    for row in ws.iter_rows():
        for cell in row:
            if cell.value is None:
                continue
            col = cell.column_letter
            length = len(str(cell.value))
            if length > widths.get(col, 0):
                widths[col] = length
    for col, length in widths.items():
        ws.column_dimensions[col].width = max(min_width, min(length + padding, max_width))


def autofit_all_sheets(wb, **kwargs):
    """Autofit every sheet in an already-open workbook."""
    for ws in wb.worksheets:
        autofit_columns(ws, **kwargs)


def replace_sheet_with_matrix(wb, sheet_name, matrix_rows):
    """Clears out `sheet_name` in an already-open workbook and rewrites it
    with `matrix_rows` (a list of row-lists; row 1 is the header row).

    If a sheet by that name already exists, its position in the tab order
    is preserved (delete + recreate at the same index) instead of getting
    shoved to the end. If it doesn't exist yet, it's appended.

    This is what lets 01_Master_Code.py's _compute_and_write_matrices()
    batch several indicators into a single load -> write-all -> save pass
    instead of every write_matrix() doing its own load_workbook/save.

    Returns the new worksheet so callers can apply fills/fonts and run
    excel_utils.autofit_columns() on it.
    """
    if sheet_name in wb.sheetnames:
        idx = wb.sheetnames.index(sheet_name)
        del wb[sheet_name]
        ws = wb.create_sheet(sheet_name, idx)
    else:
        ws = wb.create_sheet(sheet_name)

    for row in matrix_rows:
        ws.append(row)

    return ws


def atomic_save(wb, output_path):
    """Write to a temp file then swap it in, so a crash mid-write never
    corrupts the workbook other pipeline steps have already written to."""
    tmp_fd, tmp_path = tempfile.mkstemp(suffix='.xlsx')
    os.close(tmp_fd)
    try:
        wb.save(tmp_path)
        shutil.move(tmp_path, output_path)
    except Exception:
        if os.path.exists(tmp_path):
            os.remove(tmp_path)
        raise


def restrict_to_target_date(df, target_date, dt_col='_sort_dt'):
    """Filters an ALREADY-COMPUTED multi-day indicator dataframe down to
    just target_date's own rows, for matrix export only.

    Call this LAST, after the indicator itself has been computed on the
    full multi-day history (EMA/RSI/ADX warmup, VWAP/OBV session resets,
    HTF Bias's 15M/Daily lookback all need the prior days on disk) -- this
    does not affect indicator correctness, only which rows are eligible to
    be written into the Symbol x Time matrix.

    Returns None (not an empty DataFrame) if target_date has no rows yet
    -- e.g. called before the market has opened, or before today's first
    candle has closed -- so callers can treat it identically to any other
    "nothing to process yet for this symbol" case instead of writing a
    stale prior-day value into today's column.
    """
    target_ts = pd.Timestamp(target_date).normalize()
    out = df[df[dt_col].dt.normalize() == target_ts]
    return out if not out.empty else None
