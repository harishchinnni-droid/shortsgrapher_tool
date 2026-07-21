"""
run_pipeline_lite.py
---------------------
NEW MODULE -- part of the 3-indicator standalone pipeline (Task 52,
18-Jul-26). Near-identical to run_pipeline.py's run_pipeline_for_date()
(same file provisioning / token mapping / historical data ingestion via
file_mgmt, token_mgmt, data_ingestion -- all reused unmodified), but the
INDICATORS list is swapped down to exactly 3: TW ALL (lite_tw_all.py),
RSI (lite_rsi.py), and ADX (adx_di.py, reused completely as-is -- its
existing DI+/DI- direction + ADX>20 threshold recomm rule already matches
Harish's spec for this pipeline).

Does NOT modify or replace run_pipeline.py -- that file (and its own
13-indicator INDICATORS list) stays exactly as-is for the existing full
pipeline (01_Master_Code.py). This is the lean equivalent used by
02_Master_Code_3Indicator.py only.

Output workbook: same filename convention as the full pipeline
("DD-Mon-YY FNO.xlsx", via file_mgmt.provision_daily_trade_file()) --
Harish confirmed he will delete the old files himself before running
this, so there is no separate naming scheme here.
"""

import os
import sys
import traceback
import concurrent.futures
from ist_clock import now_ist

CODES_DIR = os.path.dirname(os.path.abspath(__file__))
if CODES_DIR not in sys.path:
    sys.path.append(CODES_DIR)

try:
    import broker_auth
    import calendar_mgmt
    import file_mgmt
    import token_mgmt
    import data_ingestion
    import excel_utils
    import adx_di
    import lite_tw_all
    import lite_rsi
except ImportError as e:
    print(f"[CRITICAL ERROR] Failed to import lite pipeline modules: {e}")
    print("Ensure all scripts are saved in the '05 Codes' directory with correct filenames.")
    sys.exit(1)

INDICATOR_INTERVAL = '5minute'

# (display label, module, sheet name written to the workbook) -- same
# interface every indicator module exposes:
#   build_matrix(data_dict, target_date, max_workers=None) -> matrix_rows
#   write_matrix(matrix_rows, output_excel_path)
# Exactly 3 indicators, per Harish's explicit spec -- no other indicator
# is computed or referenced by this pipeline.
INDICATORS = [
    ("TW All In One (lite)",   lite_tw_all,   "TW ALL"),
    ("RSI (lite)",             lite_rsi,      "RSI"),
    ("ADX & DI",               adx_di,        "ADX"),
]


def run_pipeline_for_date_lite(smart_api, kite_api, target_date, mode):
    """Steps 3-8 only, mirrors run_pipeline.run_pipeline_for_date() exactly
    except for the INDICATORS list. Returns (final_excel_path, df_ref)."""
    # STEP 3: File Provisioning
    new_filename = file_mgmt.provision_daily_trade_file(target_date)
    # STEP 4: Token Synchronization & Mapping
    df_ref = token_mgmt.update_instrument_tokens(new_filename, kite_api, target_date=target_date)
    # STEP 5: Historical Data Ingestion
    try:
        is_live = (mode == calendar_mgmt.LIVE)

        already_exists = data_ingestion.historical_data_exists(df_ref, target_date)
        stale = False
        if is_live and already_exists:
            first_sym = str(df_ref.dropna(subset=['Zerodha_Token']).iloc[0]['Symbol / StrikePrice']).strip().upper()
            date_str = target_date.strftime('%d-%b-%y')
            check_path = os.path.join(data_ingestion.HIST_DIR, f"{first_sym}_5minute_{date_str}.csv")
            stale = data_ingestion.file_has_future_candles(check_path, now_ist())

        if already_exists and not stale:
            print("[SYSTEM] Historical data already available for today -- skipping full backfill.")
            # [FIXED -- Task 69, 21-Jul-26] Same fix as run_pipeline.py --
            # only catch up incrementally in LIVE mode. See that file's
            # comment for the full reasoning: this unconditional call was
            # the actual trigger for the tz-naive/tz-aware crash, since it
            # ran a live-"now"-based fetch/merge against a BACKTEST target
            # date's already-closed, supposedly-immutable CSV.
            if is_live:
                data_ingestion.update_incremental_data(df_ref, target_date, kite_api)
            else:
                print("[SYSTEM] BACKTEST target date already fully downloaded (closed trading "
                      "day) -- no incremental catch-up needed.")
        else:
            if stale:
                print("[WARNING] Existing data on disk extends beyond the current time -- "
                      "treating as stale/leftover, not live. Purging and rebuilding.")
                data_ingestion.purge_interval_data(df_ref, target_date, interval='5minute')
            data_ingestion.download_historical_data(df_ref, target_date, kite_api, cap_to_now=is_live)
        print("[SUCCESS] Historical Data Download Complete. Pipeline fully synchronized.")
    except Exception as e:
        raise RuntimeError(f"Step 5 (Historical Ingestion) failed for {target_date.strftime('%d-%b-%y')}: {e}") from e
    print("-" * 60)

    final_excel_path = os.path.join(file_mgmt.BASE_DIR, new_filename)

    # STEP 6: Load the shared 5-minute dataset once, then compute all 3
    # indicator matrices IN PARALLEL.
    try:
        print(f"[PROCESS] Loading {INDICATOR_INTERVAL} historical data (shared across all indicators)...")
        raw_data_dict = data_ingestion.load_interval_data(df_ref, target_date, interval=INDICATOR_INTERVAL)
        if not raw_data_dict:
            raise RuntimeError(f"No {INDICATOR_INTERVAL} historical data files found for any symbol.")
    except Exception as e:
        raise RuntimeError(f"Step 6 (Historical Data Load) failed for {target_date.strftime('%d-%b-%y')}: {e}") from e

    num_indicators = len(INDICATORS)
    cpu_count = os.cpu_count() or 4
    workers_per_indicator = max(1, cpu_count // num_indicators)
    max_concurrent_indicators = min(num_indicators, max(1, cpu_count))

    print(f"[PROCESS] Launching {num_indicators} indicator engines "
          f"({max_concurrent_indicators} running at a time, "
          f"{workers_per_indicator} worker process(es) each, {cpu_count} logical CPUs detected)...")

    matrix_results = {}
    failures = {}

    with concurrent.futures.ThreadPoolExecutor(max_workers=max_concurrent_indicators) as executor:
        future_to_label = {
            executor.submit(module.build_matrix, raw_data_dict, target_date, max_workers=workers_per_indicator): label
            for label, module, sheet_name in INDICATORS
        }
        for future in concurrent.futures.as_completed(future_to_label):
            label = future_to_label[future]
            try:
                matrix_results[label] = future.result()
                print(f"  -> DONE: {label} matrix computed ({len(matrix_results[label]) - 1} rows).")
            except Exception as exc:
                failures[label] = exc
                print(f"  [ERROR] {label} failed to compute: {exc}")

    print("-" * 60)

    # STEP 7: Write every successfully-computed matrix into the shared
    # workbook ONE AT A TIME, in the main thread.
    print("[PROCESS] Writing indicator matrices to workbook (sequential to avoid file races)...")
    for label, module, sheet_name in INDICATORS:
        if label not in matrix_results:
            continue
        try:
            module.write_matrix(matrix_results[label], final_excel_path)
            print(f"  -> WRITTEN: '{sheet_name}' sheet updated.")
        except Exception as exc:
            failures[label] = exc
            print(f"  [ERROR] Failed writing '{sheet_name}' sheet: {exc}")

    print("-" * 60)

    # STEP 8: Auto-fit column widths across every sheet in the workbook.
    try:
        print("[PROCESS] Auto-fitting column widths across all sheets...")
        from openpyxl import load_workbook
        wb = load_workbook(final_excel_path)
        excel_utils.autofit_all_sheets(wb)
        excel_utils.atomic_save(wb, final_excel_path)
        print("[SUCCESS] Column widths auto-fitted.")
    except Exception as e:
        print(f"[WARNING] Auto-fit pass failed (non-fatal): {e}")

    print("=" * 60)
    if failures:
        print(f"[WARNING] {len(failures)} of {num_indicators} indicator(s) failed:")
        for label, exc in failures.items():
            print(f"   - {label}: {exc}")
        print("   Successful indicators were still written to the workbook above.")
        print("=" * 60)
        raise RuntimeError(
            f"{len(failures)} of {num_indicators} indicator(s) failed for "
            f"{target_date.strftime('%d-%b-%y')}: {list(failures.keys())}"
        )

    print("   STATUS: READY (3-indicator lite pipeline).   ")
    print("=" * 60)
    return final_excel_path, df_ref
