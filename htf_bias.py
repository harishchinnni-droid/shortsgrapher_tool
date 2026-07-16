"""
htf_bias.py
-----------
Rebuilt from scratch -- the original source was lost (only a stale
__pycache__/htf_bias.cpython-312.pyc leftover; no Pine Script equivalent
exists for this one, it's a custom module). Two other files hard-depend
on it and were BOTH silently degraded while it was missing:

    1. ema20.py imports htf_bias directly and calls compute_htf_bias(df)
       inline (see that module's docstring) to gate its own 'EMA 20
       Recomm' column -- without this file, ema20.py can't even be
       imported.
    2. order_sheet.py reads a literal 'HTF Bias' SHEET from the workbook
       -- `_load_metric_lookup(output_excel_path, 'HTF Bias', 'ADX Value')`
       and `'ATR Value'` -- for its low-momentum gate (#5) and its
       ATR-based position sizing. With the sheet missing, both of those
       silently no-op (ADX gate never rejects anything; position sizing
       falls back to its floor-percentage stop instead of a real
       ATR-adaptive one) with only a console [WARNING], not a crash. That
       is a real, currently-live degradation of the risk system, not a
       cosmetic gap -- this file is worth rebuilding for that reason
       alone, independent of ema20.py's HTF gate.

Design note -- this is a reconstruction, not a recovered file:
    Nothing in this codebase's comments states exactly how the original
    computed 'BIAS_15M' / 'BIAS_DAILY'. What IS pinned down by ema20.py's
    own docstring is the CONTRACT: two columns whose values are the
    strings 'UP' / 'DOWN', used as a veto (a BUY CE drops to WAIT if
    EITHER reads 'DOWN'; a BUY PE drops to WAIT if EITHER reads 'UP').
    This implementation fills that contract with a standard, explainable
    "higher-timeframe trend via EMA structure" rule: resample the 5-min
    series up to 15-minute and daily bars, and read UP/DOWN off whether
    the LAST FULLY CLOSED higher-timeframe bar's close is above/below its
    own 20-period EMA on that higher timeframe. Flag this to Harish if the
    lost original used a different HTF bias definition (e.g. higher-high/
    higher-low structure, or a different EMA length) -- the contract
    (UP/DOWN veto into ema20.py) is honored either way, but the specific
    rule producing UP/DOWN is this module's own reconstruction.

No-lookahead rule: a 5-min bar can only ever see the higher-timeframe
bias from the LAST HTF bar that had ALREADY CLOSED strictly before that
5-min bar's own timestamp -- never the HTF bar it's currently inside of
(which is still forming and would leak future information back onto
earlier 5-min bars once that HTF bar eventually closes).

ADX Value / ATR Value:
    ADX_VAL reuses adx_di.py's own Wilder ADX computation directly
    (imported, not re-derived) so the two sheets can never quietly drift
    onto two different ADX formulas -- order_sheet.py's momentum gate and
    the 'ADX' sheet's own Recomm are reading the literal same number.
    ATR_VAL is a standard Wilder ATR(14) on the underlying's true range,
    using the same ewm(alpha=1/period, adjust=False) approximation this
    codebase's own RSI.py already uses for its Wilder smoothing -- kept
    consistent with that established convention rather than introducing a
    third smoothing style.
"""

import os
import sys
from datetime import time as dtime

import numpy as np
import pandas as pd
from openpyxl import load_workbook

CODES_DIR = os.path.dirname(os.path.abspath(__file__))
if CODES_DIR not in sys.path:
    sys.path.append(CODES_DIR)
import data_ingestion
import excel_utils
import adx_di

CUTOFF_TIME = dtime(15, 15)
INTERVAL = "5minute"
ADX_ATR_PERIOD = 14
HTF_EMA_LEN = 20


def _wilder_atr(df, period=ADX_ATR_PERIOD):
    """Standard Wilder ATR via ewm(alpha=1/period, adjust=False) on true
    range -- same smoothing convention RSI.py already uses in this
    codebase. Deliberately NOT adx_di._wilder_running_sum(): that helper
    is an unnormalized running SUM (correct for ADX/DI, where DI+/DI-
    ratios cancel the scale) and would be off by roughly a factor of
    `period` if used as a standalone price-scale ATR."""
    high, low, close = df['high'], df['low'], df['close']
    prev_close = close.shift(1)
    true_range = pd.concat([
        high - low, (high - prev_close).abs(), (low - prev_close).abs(),
    ], axis=1).max(axis=1).fillna(0.0)
    return true_range.ewm(alpha=1.0 / period, adjust=False).mean()


def _resample_bias(df, rule, ema_len=HTF_EMA_LEN):
    """Resamples df (indexed by '_sort_dt') up to `rule` ('15min' or
    '1D'), computes an EMA(ema_len) trend on the resampled closes, and
    returns a Series aligned back to df's own index giving each 5-min bar
    the bias of the LAST HTF bar that had already fully closed before
    that bar's own timestamp -- never the still-forming current one.
    'UP' if that HTF close > its own EMA, 'DOWN' otherwise. Returns
    'DOWN' (safe/conservative default -- never silently 'UP') for any bar
    before the first HTF bar has closed."""
    htf = df.set_index('_sort_dt')['close'].resample(rule, label='right', closed='right').last().dropna()
    if htf.empty:
        return pd.Series('DOWN', index=df.index)

    htf_ema = htf.ewm(span=ema_len, adjust=False).mean()
    htf_bias = pd.Series(np.where(htf > htf_ema, 'UP', 'DOWN'), index=htf.index)

    # `label='right', closed='right'` timestamps each HTF bar at its OWN
    # close time -- so merge_asof with direction='forward' on the 5-min
    # bar's own timestamp correctly finds "the next HTF close at or after
    # this exact 5-min timestamp" only when the 5-min bar IS that close;
    # for every other (earlier) 5-min bar we want the PREVIOUS HTF close,
    # i.e. direction='backward'. Using 'backward' means a 5-min bar at
    # exactly an HTF close boundary sees that bar's OWN close (correct --
    # it has just finished), and every bar before the first HTF close
    # falls through to the 'DOWN' default via the initial fillna below.
    merged = pd.merge_asof(
        pd.DataFrame({'_sort_dt': df['_sort_dt'].values}).sort_values('_sort_dt'),
        pd.DataFrame({'_sort_dt': htf_bias.index, 'bias': htf_bias.values}).sort_values('_sort_dt'),
        on='_sort_dt', direction='backward',
    )
    return merged['bias'].fillna('DOWN').to_numpy()


def compute_htf_bias(df):
    """Requires df sorted ascending by '_sort_dt', lowercase OHLC columns
    (same precondition ema20.py's docstring already documents). Adds
    ADX_VAL, ATR_VAL, BIAS_15M, BIAS_DAILY columns and returns df.
    Computed on the FULL multi-day history handed in -- callers restrict
    to target_date only AFTERWARDS, same convention as every other
    indicator's warmup (RSI's EMA warmup, this module's own HTF resample
    lookback, etc.)."""
    df = df.copy()

    adx_df = adx_di.calculate_adx_di(df)
    df['ADX_VAL'] = adx_df['ADX']
    df['ATR_VAL'] = _wilder_atr(df)
    df['BIAS_15M'] = _resample_bias(df, '15min')
    df['BIAS_DAILY'] = _resample_bias(df, '1D')
    return df


# ---------------------------------------------------------------------------
# Per-symbol processing (same brute-force time/timezone protocol as RSI.py)
# ---------------------------------------------------------------------------
def process_symbol(symbol_data):
    symbol, df, target_date = symbol_data
    df = df.copy()
    df.columns = df.columns.astype(str).str.strip().str.lower()

    if not all(col in df.columns for col in ['open', 'high', 'low', 'close']):
        print(f"[WARNING] {symbol}: Missing required price columns. Discarding.")
        return symbol, None

    time_col = next((col for col in ['date', 'time', 'datetime', 'timestamp'] if col in df.columns), None)
    if time_col:
        parsed = pd.to_datetime(df[time_col], errors='coerce')
    elif pd.api.types.is_datetime64_any_dtype(df.index):
        parsed = pd.Series(df.index, index=df.index)
    else:
        unnamed_cols = [c for c in df.columns if 'unnamed' in c]
        if unnamed_cols:
            parsed = pd.to_datetime(df[unnamed_cols[0]], errors='coerce')
        else:
            print(f"[WARNING] {symbol}: Failed to locate any valid timestamp data. Discarding.")
            return symbol, None

    try:
        if parsed.dt.tz is not None:
            parsed = parsed.dt.tz_localize(None)
    except (AttributeError, TypeError):
        pass

    df['_sort_dt'] = parsed
    df['time_str'] = parsed.dt.strftime('%H:%M')
    df.dropna(subset=['time_str', '_sort_dt'], inplace=True)
    if df.empty:
        print(f"[WARNING] {symbol}: No rows with a valid timestamp. Discarding.")
        return symbol, None

    df.sort_values('_sort_dt', inplace=True)
    df = df[df['_sort_dt'].dt.time <= CUTOFF_TIME]
    if df.empty:
        print(f"[WARNING] {symbol}: No candles at/before {CUTOFF_TIME.strftime('%H:%M')}. Discarding.")
        return symbol, None

    df = compute_htf_bias(df)

    df = excel_utils.restrict_to_target_date(df, target_date)
    if df is None:
        return symbol, None

    return symbol, df


# ---------------------------------------------------------------------------
# Excel export -- pivoted Symbol x Time matrix. No 'Recomm' row here on
# purpose: HTF Bias is a gate consumed by ema20.py/order_sheet.py, not an
# independent vote in final_sheet.py -- see that module's docstring.
# ---------------------------------------------------------------------------
def build_matrix(data_dict, target_date, max_workers=None):
    import concurrent.futures
    results = {}
    with concurrent.futures.ProcessPoolExecutor(max_workers=max_workers) as executor:
        futures = {executor.submit(process_symbol, (sym, df, target_date)): sym for sym, df in data_dict.items()}
        for future in concurrent.futures.as_completed(futures):
            sym = futures[future]
            try:
                processed_sym, processed_df = future.result()
                if processed_df is not None and 'ADX_VAL' in processed_df.columns:
                    if not processed_df.empty and sym.lower() != 'summary':
                        results[sym] = processed_df
            except Exception as e:
                print(f"[ERROR] HTF Bias Engine Failure on symbol {sym}: {str(e)}")

    if not results:
        raise RuntimeError("HTF Bias: zero symbols processed successfully for target_date -- matrix build aborted.")

    all_times = set()
    for df in results.values():
        if 'time_str' in df.columns:
            all_times.update(df['time_str'].dropna().tolist())

    if not all_times:
        raise RuntimeError("HTF Bias: no valid timestamps extracted across all symbols -- matrix build aborted.")

    sorted_times = sorted(all_times)
    matrix_rows = [['Symbol', 'Metrics'] + sorted_times]

    for sym in sorted(results.keys()):
        df = results[sym].drop_duplicates(subset=['time_str'], keep='last')
        df_indexed = df.set_index('time_str')

        def get_metric_row(metric_name, col_name):
            row = [sym, metric_name]
            for t in sorted_times:
                if t in df_indexed.index:
                    val = df_indexed.loc[t, col_name]
                    if isinstance(val, pd.Series):
                        val = val.iloc[-1]
                    row.append(val)
                else:
                    row.append("")
            return row

        # Exact labels order_sheet.py looks up: 'ADX Value' / 'ATR Value'.
        matrix_rows.append(get_metric_row('ADX Value', 'ADX_VAL'))
        matrix_rows.append(get_metric_row('ATR Value', 'ATR_VAL'))
        matrix_rows.append(get_metric_row('Bias 15M', 'BIAS_15M'))
        matrix_rows.append(get_metric_row('Bias Daily', 'BIAS_DAILY'))

    return matrix_rows


def write_matrix_to_workbook(matrix_rows, wb):
    sheet_name = "HTF Bias"
    ws = excel_utils.replace_sheet_with_matrix(wb, sheet_name, matrix_rows)
    excel_utils.autofit_columns(ws)


def write_matrix(matrix_rows, output_excel_path):
    wb = load_workbook(output_excel_path)
    write_matrix_to_workbook(matrix_rows, wb)
    excel_utils.atomic_save(wb, output_excel_path)


def parallel_compute_and_export(data_dict, target_date, output_excel_path, max_workers=None):
    matrix_rows = build_matrix(data_dict, target_date, max_workers=max_workers)
    write_matrix(matrix_rows, output_excel_path)


def run_htf_bias_step(df_ref, target_date, output_excel_path):
    print("[SYSTEM] Loading 5-minute historical data for HTF Bias...")
    data_dict = data_ingestion.load_interval_data(df_ref, target_date, interval=INTERVAL)
    if not data_dict:
        raise RuntimeError("HTF Bias: no 5-minute data files found for any symbol.")
    print(f"[PROCESS] Computing HTF Bias matrix for {len(data_dict)} symbol(s)...")
    parallel_compute_and_export(data_dict, target_date, output_excel_path)
    print("[SUCCESS] HTF Bias matrix written to sheet 'HTF Bias'.")


# ---------------------------------------------------------------------------
# Disclaimer: this module derives a signal from historical/live price data
# only. It is not financial advice, and no result here is a guarantee of
# future performance -- paper-trade it and run it through
# scripts/backtester.py before risking real capital.
# ---------------------------------------------------------------------------
