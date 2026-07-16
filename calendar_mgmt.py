from datetime import datetime, timedelta
from ist_clock import now_ist

# Mode constants -- both run_pipeline.py and 01_Master_Code.py (live_loop.py)
# branch on these rather than the raw '1'/'2' input strings.
LIVE = "LIVE"
BACKTEST = "BACKTEST"

def is_trading_holiday(target_date):
    if target_date.weekday() >= 5:
        return True
    nse_holidays_2026 = [
        "2026-01-26", "2026-03-03", "2026-03-20", "2026-04-03",
        "2026-04-14", "2026-05-01", "2026-08-15", "2026-09-19",
        "2026-10-02", "2026-10-18", "2026-11-08", "2026-12-25"
    ]
    return target_date.strftime("%Y-%m-%d") in nse_holidays_2026

def get_execution_date():
    """Prompts for LIVE vs BACKTEST and returns (target_date, mode).

    Returning mode alongside the date (not just the date) matters: a caller
    that only gets target_date back has no reliable way to tell "today,
    picked as LIVE" apart from "today, picked as BACKTEST" -- and without
    that distinction it's easy to end up hardcoding datetime.now() and
    silently losing BACKTEST support, which is what happened to
    01_Master_Code.py (live_loop.py) before this function returned mode too.
    """
    while True:
        mode = input("Select Mode (1 for LIVE, 2 for BACKTEST): ").strip()

        if mode == '1':
            # [FIX] was datetime.now() -- the host machine's own local
            # time, which is only correct if the host happens to be set to
            # Asia/Kolkata. now_ist() is correct regardless of host config.
            target_date = now_ist()
            if is_trading_holiday(target_date):
                print("[WARNING] Today is a weekend or NSE holiday. Market is closed.")
                # You can either strictly reject or allow it to proceed for offline testing.
                proceed = input("Proceed anyway? (Y/N): ").strip().upper()
                if proceed != 'Y':
                    continue
            return target_date, LIVE

        elif mode == '2':
            date_input = input("Enter Backtest Date (DD-MMM-YY, e.g., 07-Jul-26): ").strip()
            try:
                target_date = datetime.strptime(date_input, '%d-%b-%y')
                if is_trading_holiday(target_date):
                    print(f"[WARNING] {date_input} is a recognized trading holiday or weekend.")
                return target_date, BACKTEST
            except ValueError:
                print("[ERROR] Invalid date format. Use DD-MMM-YY.")
        else:
            print("Invalid selection.")


# ---------------------------------------------------------------------------
# [ADDED] Backtest DATE RANGE support -- Harish's spec: key in a start and
# end date once (e.g. 01-Jul-26 to 15-Jul-26) and get one trading file per
# trading day in that range, weekends/NSE holidays skipped automatically,
# instead of running the whole pipeline by hand once per date.
# ---------------------------------------------------------------------------
def get_trading_dates_in_range(start_date, end_date):
    """Every calendar date from start_date to end_date INCLUSIVE, with
    weekends and NSE holidays (is_trading_holiday()) removed. Returns a
    sorted list of datetimes. Prints which dates were skipped and why, so
    a 01-Jul-26..15-Jul-26 request visibly shows its 4 weekend days (and
    any holiday) being dropped rather than silently vanishing from the
    output file count."""
    if end_date < start_date:
        start_date, end_date = end_date, start_date
        print("[WARNING] End date was before start date -- swapped so the range still makes sense.")

    trading_dates = []
    skipped = []
    current = start_date
    while current <= end_date:
        if is_trading_holiday(current):
            reason = "weekend" if current.weekday() >= 5 else "NSE holiday"
            skipped.append((current, reason))
        else:
            trading_dates.append(current)
        current += timedelta(days=1)

    total_days = (end_date - start_date).days + 1
    print(f"[SYSTEM] Date range {start_date.strftime('%d-%b-%y')} to {end_date.strftime('%d-%b-%y')}: "
          f"{total_days} calendar day(s), {len(trading_dates)} trading day(s), {len(skipped)} skipped.")
    for d, reason in skipped:
        print(f"  [SKIP] {d.strftime('%d-%b-%y')} ({reason})")

    return trading_dates


def _prompt_date(label):
    while True:
        raw = input(f"Enter {label} (DD-MMM-YY, e.g., 01-Jul-26): ").strip()
        try:
            return datetime.strptime(raw, '%d-%b-%y')
        except ValueError:
            print("[ERROR] Invalid date format. Use DD-MMM-YY.")


def get_run_config():
    """Single entry point for 01_Master_Code.py. Prompts LIVE vs BACKTEST:

    LIVE     -> returns (LIVE, [today's date]) -- today_ist() via now_ist(),
                a warning (not a hard block) if today is a weekend/holiday,
                same as get_execution_date()'s existing LIVE behavior.
    BACKTEST -> prompts for a START and END date and returns
                (BACKTEST, [trading_date, trading_date, ...]) -- the
                resolved list from get_trading_dates_in_range(), i.e. one
                entry per trading day the caller should run the full
                pipeline for and produce one dated file each.

    Unlike get_execution_date(), this always returns a LIST of dates (even
    for LIVE, a single-item list) so 01_Master_Code.py has one uniform
    "loop over these dates" code path regardless of mode.
    """
    while True:
        mode = input("Select Mode (1 for LIVE, 2 for BACKTEST): ").strip()

        if mode == '1':
            target_date = now_ist()
            if is_trading_holiday(target_date):
                print("[WARNING] Today is a weekend or NSE holiday. Market is closed.")
                proceed = input("Proceed anyway? (Y/N): ").strip().upper()
                if proceed != 'Y':
                    continue
            return LIVE, [target_date]

        elif mode == '2':
            start_date = _prompt_date("Start Date")
            end_date = _prompt_date("End Date")
            trading_dates = get_trading_dates_in_range(start_date, end_date)
            if not trading_dates:
                print("[ERROR] No trading days in that range -- every date was a weekend/holiday. Try again.")
                continue
            return BACKTEST, trading_dates

        else:
            print("Invalid selection.")
