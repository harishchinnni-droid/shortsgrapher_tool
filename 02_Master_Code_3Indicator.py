"""
02_Master_Code_3Indicator.py
------------------------------
NEW STANDALONE ENTRY POINT (Task 52, 18-Jul-26) -- Harish's brand-new,
isolated 3-indicator pipeline: TW All-in-One (lite_tw_all.py), RSI
(lite_rsi.py), and ADX/DI (adx_di.py, reused unmodified). Signals combine
via strict unanimous confluence (lite_final_sheet.py) into the SAME
'Final Recomm' interface order_sheet.py already reads -- so order_sheet.py
itself (audit/rejection checks, PCR, TSL/SL, position management) runs
COMPLETELY UNMODIFIED, exactly as it does for the full 13-indicator
pipeline.

Mirrors 01_Master_Code.py's orchestration structure exactly (broker
login once, mode/date selection, per-date cycle), swapping only:
    run_pipeline.run_pipeline_for_date()   -> run_pipeline_lite.run_pipeline_for_date_lite()
    final_sheet.run_final_sheet_step()     -> lite_final_sheet.run_final_sheet_lite_step()
order_sheet's own step is UNCHANGED.

Does NOT modify or replace 01_Master_Code.py -- both entry points coexist.
Run THIS file (not 01_Master_Code.py) to use the 3-indicator strategy;
run 01_Master_Code.py for the existing full-indicator pipeline. They
write to the same "DD-Mon-YY FNO.xlsx" filename convention, so only run
one of them for a given date at a time -- Harish has confirmed he will
delete any old files for the dates he re-runs here.

Nothing in this file (or anything it calls) places a real order with a
broker -- BACKTEST simulates trades against historical data, and LIVE only
TRACKS simulated positions against live quotes. See order_sheet.py's and
position_manager.py's own disclaimers. This is a new, unbacktested rule
set (3-indicator unanimous confluence) -- paper-trade thoroughly before
ever wiring real order placement on top of this.
"""

import sys
import time
import traceback

import broker_auth
import calendar_mgmt
import run_pipeline_lite
import lite_final_sheet
import order_sheet
import process_log
from ist_clock import now_ist, is_before_market_open, seconds_until_market_open, MARKET_CLOSE_TIME

CYCLE_INTERVAL_SECONDS = 300  # 5 minutes -- matches every indicator's own 5-minute candle interval


def _run_full_cycle(smart_api, kite_api, target_date, mode, live_first_run_of_day=False, cycle_num=0):
    """One full pass of the lite pipeline for one date: data -> 3
    indicators -> unanimous confluence -> orders/positions (order_sheet.py,
    unmodified) -> same TSL/SL/PCR/audit logic as the full pipeline."""
    # [ADDED -- Task 73, 22-Jul-26] Process Log timer -- covers file
    # provisioning + token sync + historical data ingestion + indicator
    # computation + Final sheet write ONLY. Deliberately stops before
    # order_sheet.py is even called -- see process_log.py's own docstring
    # for why (Harish's explicit boundary: "not placing orders from order
    # sheet and TSL").
    _cycle_start = time.time()
    final_excel_path, df_ref = run_pipeline_lite.run_pipeline_for_date_lite(smart_api, kite_api, target_date, mode)

    lite_final_sheet.run_final_sheet_lite_step(final_excel_path)
    _cycle_duration = time.time() - _cycle_start
    process_log.log_cycle_timing(
        final_excel_path, mode, target_date, cycle_num, _cycle_duration, now_ist().strftime('%H:%M:%S')
    )

    if mode == calendar_mgmt.LIVE and not live_first_run_of_day:
        order_sheet.update_open_positions_live(kite_api, final_excel_path)

    order_sheet.run_order_sheet_step(final_excel_path, kite_api, df_ref, mode=mode, target_date=target_date)

    return final_excel_path


def run_backtest_range(smart_api, kite_api, trading_dates):
    """One file per trading date -- same convention as 01_Master_Code.py.
    A failure on one date does NOT stop the rest of the range."""
    print(f"\n[SYSTEM] BACKTEST range (3-indicator lite pipeline): {len(trading_dates)} trading date(s) queued.")
    succeeded, failed = [], []

    for i, target_date in enumerate(trading_dates, start=1):
        date_label = target_date.strftime('%d-%b-%y')
        print("\n" + "=" * 60)
        print(f"  BACKTEST {i}/{len(trading_dates)} -- {date_label} (3-indicator lite)")
        print("=" * 60)
        try:
            final_excel_path = _run_full_cycle(smart_api, kite_api, target_date, calendar_mgmt.BACKTEST)
            print(f"[SUCCESS] {date_label} complete -> {final_excel_path}")
            succeeded.append(date_label)
        except Exception as e:
            print(f"[ERROR] {date_label} failed -- skipping to next date in range: {e}")
            print(traceback.format_exc())
            failed.append((date_label, str(e)))

    print("\n" + "=" * 60)
    print(f"  BACKTEST RANGE COMPLETE: {len(succeeded)} succeeded, {len(failed)} failed")
    print("=" * 60)
    for date_label, err in failed:
        print(f"  [FAILED] {date_label}: {err}")


def run_live_session(smart_api, kite_api, target_date):
    """Runs today once immediately, then loops every CYCLE_INTERVAL_SECONDS
    until NSE market close."""
    if is_before_market_open():
        wait_s = seconds_until_market_open()
        print(f"[SYSTEM] Market not open yet -- waiting {wait_s / 60:.1f} minute(s) for 09:15 IST...")
        time.sleep(wait_s)

    print(f"\n[SYSTEM] LIVE session (3-indicator lite) starting for {target_date.strftime('%d-%b-%y')}.")
    cycle_num = 0
    while True:
        now = now_ist()
        if now.time() >= MARKET_CLOSE_TIME:
            print("[SYSTEM] Market closed (15:30 IST). Ending LIVE session for today.")
            break

        cycle_num += 1
        print("\n" + "=" * 60)
        print(f"  LIVE CYCLE {cycle_num} -- {now.strftime('%H:%M:%S')} IST")
        print("=" * 60)
        try:
            _run_full_cycle(smart_api, kite_api, target_date, calendar_mgmt.LIVE,
                             live_first_run_of_day=(cycle_num == 1), cycle_num=cycle_num)
        except Exception as e:
            print(f"[ERROR] LIVE cycle {cycle_num} failed: {e}")
            print(traceback.format_exc())

        now = now_ist()
        seconds_into_interval = (now.minute % 5) * 60 + now.second
        sleep_s = max(5.0, CYCLE_INTERVAL_SECONDS - seconds_into_interval)
        print(f"[SYSTEM] Sleeping {sleep_s:.0f}s until next cycle...")
        time.sleep(sleep_s)


def main():
    print("=" * 60)
    print("   02_MASTER_CODE_3INDICATOR -- TW ALL + RSI + ADX PIPELINE")
    print("=" * 60)

    # STEP 1: Broker Authentication -- once per run.
    try:
        smart_api = broker_auth.initialize_angel_one()
        kite_api = broker_auth.initialize_zerodha()
    except Exception as e:
        print(f"\n[FATAL] Broker authentication failed: {e}")
        print(traceback.format_exc())
        sys.exit(1)

    # STEP 2: Mode + date(s).
    try:
        mode, dates = calendar_mgmt.get_run_config()
    except Exception as e:
        print(f"\n[FATAL] Mode/date selection failed: {e}")
        sys.exit(1)

    if mode == calendar_mgmt.BACKTEST:
        run_backtest_range(smart_api, kite_api, dates)
    else:
        run_live_session(smart_api, kite_api, dates[0])


if __name__ == "__main__":
    main()
