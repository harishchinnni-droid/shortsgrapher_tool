"""
01_Master_Code.py
------------------
THE FILE TO RUN. Every other script in this folder (broker_auth,
calendar_mgmt, file_mgmt, token_mgmt, data_ingestion, every indicator
module, final_sheet, order_sheet, dashboard) is a building block this
file wires together in the right order. Run this one, not the individual
pieces -- they only expect to be called FROM here (or from run_pipeline.py
directly for a quick single-date manual run).

Referenced by name throughout this codebase's comments ("see
01_Master_Code.py", "01_Master_Code.run_cycle()") as the intended live-loop
entry point -- that file did not actually exist anywhere in this project
before now (only the comments pointing to it survived). Rebuilt here from
those comments' own description of what it's supposed to do, plus every
module's own "single entry point for 01_Master_Code.py" docstring, which
between them fully specify the call order below.

What this does, end to end:
    1. Logs into Angel One + Zerodha ONCE (broker_auth) -- same-day token
       caching means this is fast on a same-day re-run.
    2. Asks LIVE vs BACKTEST (calendar_mgmt.get_run_config()).
         LIVE     -> runs today, then loops every 5 minutes until market
                     close, incrementally syncing data and re-running the
                     signal/order pipeline each cycle.
         BACKTEST -> asks for a START and END date, builds one dated file
                     PER TRADING DAY in that range (weekends/NSE holidays
                     skipped automatically), and runs the full pipeline
                     once per date -- one bad date does not abort the rest
                     of the range.
    3. Per date: run_pipeline.run_pipeline_for_date() (file provisioning,
       token mapping, historical data, every indicator matrix) ->
       final_sheet.run_final_sheet_step() (confluence) ->
       order_sheet.run_order_sheet_step() (signal -> simulated order,
       real Net P/L in BACKTEST, tracked position in LIVE) -> dashboard
       (already invoked from inside run_order_sheet_step()).

Nothing in this file (or anything it calls) places a real order with a
broker -- BACKTEST simulates trades against historical data, and LIVE only
TRACKS simulated positions against live quotes. See order_sheet.py's and
position_manager.py's own disclaimers. Paper-trade thoroughly, and
understand every gate/exit rule in order_sheet.py's module docstring,
before ever wiring real order placement on top of this.
"""

import sys
import time
import traceback
from datetime import timedelta

import broker_auth
import calendar_mgmt
import run_pipeline
import final_sheet
import order_sheet
import process_log
from ist_clock import now_ist, is_before_market_open, seconds_until_market_open, MARKET_CLOSE_TIME

CYCLE_INTERVAL_SECONDS = 300  # 5 minutes -- matches every indicator's own 5-minute candle interval


def _run_full_cycle(smart_api, kite_api, target_date, mode, live_first_run_of_day=False, cycle_num=0):
    """One full pass of the pipeline for one date: data -> indicators ->
    confluence -> orders/positions -> dashboard. Used identically by both
    the BACKTEST date loop and each LIVE 5-minute cycle -- the only
    difference between them is WHICH data step Step 5 inside
    run_pipeline_for_date() takes (full backfill vs incremental), which
    that function already decides for itself based on `mode` and what's
    already on disk.
    """
    # [ADDED -- Task 73, 22-Jul-26] Process Log timer -- see
    # process_log.py's own docstring / 02_Master_Code_3Indicator.py's
    # identical comment for the exact boundary (data + indicators + Final
    # sheet only, stops before order_sheet.py is ever called).
    _cycle_start = time.time()
    final_excel_path, df_ref = run_pipeline.run_pipeline_for_date(smart_api, kite_api, target_date, mode)

    final_sheet.run_final_sheet_step(final_excel_path)
    _cycle_duration = time.time() - _cycle_start
    process_log.log_cycle_timing(
        final_excel_path, mode, target_date, cycle_num, _cycle_duration, now_ist().strftime('%H:%M:%S')
    )

    if mode == calendar_mgmt.LIVE and not live_first_run_of_day:
        # [IMPORTANT] Must run BEFORE run_order_sheet_step() -- see
        # order_sheet.update_open_positions_live()'s own docstring. This
        # resolves/updates any ALREADY-OPEN position's stop/target/
        # trailing/max-hold/EOD state against the latest live quote
        # before the same cycle's order_sheet pass looks for brand-new
        # entries. On the very first LIVE run of the day there are no
        # open positions yet (the 'Orders' sheet doesn't even exist),
        # so this step is skipped -- see that function's own
        # sheet-existence check for why that's a silent, expected no-op.
        order_sheet.update_open_positions_live(kite_api, final_excel_path)

    order_sheet.run_order_sheet_step(final_excel_path, kite_api, df_ref, mode=mode, target_date=target_date)

    return final_excel_path


def run_backtest_range(smart_api, kite_api, trading_dates):
    """One file per trading date, per Harish's spec: type a start/end
    date, get every trading day in that range as its own dated workbook,
    weekends/holidays already excluded by calendar_mgmt.get_run_config().
    A failure on one date is logged and does NOT stop the rest of the
    range -- see run_pipeline.run_pipeline_for_date()'s docstring for why
    it raises instead of sys.exit()ing."""
    print(f"\n[SYSTEM] BACKTEST range: {len(trading_dates)} trading date(s) queued.")
    succeeded, failed = [], []

    for i, target_date in enumerate(trading_dates, start=1):
        date_label = target_date.strftime('%d-%b-%y')
        print("\n" + "=" * 60)
        print(f"  BACKTEST {i}/{len(trading_dates)} -- {date_label}")
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
    until NSE market close, incrementally syncing data and re-running the
    signal/order pipeline each cycle. Waits for market open first if
    started early; exits the loop (does not crash) once the session ends
    for the day."""
    if is_before_market_open():
        wait_s = seconds_until_market_open()
        print(f"[SYSTEM] Market not open yet -- waiting {wait_s / 60:.1f} minute(s) for 09:15 IST...")
        time.sleep(wait_s)

    print(f"\n[SYSTEM] LIVE session starting for {target_date.strftime('%d-%b-%y')}.")
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
            # [IMPORTANT] A single cycle failing (a rate-limit blip, a
            # transient network error, one bad symbol) must not kill the
            # whole day's LIVE session -- log it and try again next cycle.
            print(f"[ERROR] LIVE cycle {cycle_num} failed: {e}")
            print(traceback.format_exc())

        # Sleep until the next 5-minute boundary rather than a flat
        # CYCLE_INTERVAL_SECONDS from "whenever this cycle happened to
        # finish" -- keeps cycles aligned to the same candle-close times
        # every indicator module already assumes (see ist_clock.py).
        now = now_ist()
        seconds_into_interval = (now.minute % 5) * 60 + now.second
        sleep_s = max(5.0, CYCLE_INTERVAL_SECONDS - seconds_into_interval)
        print(f"[SYSTEM] Sleeping {sleep_s:.0f}s until next cycle...")
        time.sleep(sleep_s)


def main():
    print("=" * 60)
    print("   01_MASTER_CODE -- ALGO TRADING PIPELINE ENTRY POINT")
    print("=" * 60)

    # STEP 1: Broker Authentication -- once per run, not once per date.
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
