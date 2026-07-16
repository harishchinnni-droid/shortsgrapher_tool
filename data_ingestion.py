import os
import time
import threading
import pandas as pd
from datetime import datetime, timedelta
from ist_clock import now_ist
from concurrent.futures import ThreadPoolExecutor, as_completed
import file_mgmt

# [CHANGED -- cloud/Colab portability] derives from file_mgmt.BASE_DIR --
# see file_mgmt.py's BASE_DIR docstring. historical_lookup.OPTION_HIST_DIR
# is itself derived from this constant (os.path.dirname(HIST_DIR) + join),
# so fixing it here automatically carries through without a second edit.
HIST_DIR = os.path.join(file_mgmt.BASE_DIR, "02_Historical_Data")
os.makedirs(HIST_DIR, exist_ok=True)


class RateLimiter:
    """Thread-safe pacing so however many worker threads are running, the
    WHOLE pool collectively stays under Kite's ~3 requests/second historical
    data limit. Replaces the old per-worker `time.sleep(1)`, which either
    wasted headroom (too conservative with few workers) or risked 429s (too
    aggressive with many) -- this scales cleanly with worker count instead.
    """
    def __init__(self, max_per_second=3):
        self._interval = 1.0 / max_per_second
        self._lock = threading.Lock()
        self._next_slot = time.monotonic()

    def acquire(self):
        with self._lock:
            now = time.monotonic()
            wait = self._next_slot - now
            self._next_slot = max(now, self._next_slot) + self._interval
        if wait > 0:
            time.sleep(wait)


def _strip_tz(ts):
    """Kite returns tz-aware timestamps; our from_date/to_date args elsewhere
    in this codebase are naive local (IST) datetimes. Normalize so the two
    don't silently disagree."""
    if hasattr(ts, 'tzinfo') and ts.tzinfo is not None:
        return ts.tz_localize(None)
    return ts


# --------------------------------------------------------------------------
# [FIX -- 13-Jul-26] CANDLE-CLOSE (not candle-OPEN) AWARENESS
# --------------------------------------------------------------------------
# Kite's `date` field on every candle is the candle's OPEN timestamp, not
# its close. The old guard in this file did:
#
#       df_hist[_strip_tz(df_hist.index) <= now_cutoff]
#
# which only checks "has this candle STARTED", not "has this candle
# CLOSED". At 09:20:05 (the moment the live loop wakes up to fetch the
# just-closed 09:15-09:20 candle), Kite's intraday historical_data() can
# also hand back the 09:20-09:25 candle that started 5 seconds earlier --
# its open timestamp (09:20:00) is trivially <= now_cutoff (09:20:05), so
# the old filter waved it straight through. That candle is a live, still-
# forming bar: it has O=H=L=C (no movement yet) and will keep changing for
# the next 5 minutes. Writing it to disk as if it were closed data is
# EXACTLY the bug reported -- "9:15 candle can show BUY CE, then by actual
# close at 9:20 the real numbers say BUY PE" -- because the indicator and
# confluence layers were scoring a candle that hadn't finished happening
# yet. This is a documented Kite behavior, not a hypothetical: see e.g.
# https://kite.trade/forum/discussion/11613 ("if you fetch data until
# current time, you will get the current running candle as well... you
# have to look at the 2nd last row for the last completed candle") and
# https://kite.trade/forum/discussion/10984 (same root cause on 15-minute
# candles). A second, separate Kite quirk -- the newly-closed candle's own
# OHLC can be silently REVISED for anywhere from a few seconds up to
# ~30 minutes after it closes (see
# https://kite.trade/forum/discussion/14530) -- is handled separately by
# the OVERLAP_MINUTES re-fetch-and-merge window below; it is a different
# problem from "is this candle closed at all", which is what this section
# fixes.
#
# The fix: a candle is only eligible to be persisted once its OPEN time
# PLUS its own duration is <= the cutoff, i.e. it has actually closed.
# _interval_to_timedelta() gives every fetch call the candle duration it
# needs to compute that, so the same guard works for every interval this
# pipeline might ever request, not just '5minute'.
# --------------------------------------------------------------------------

_INTERVAL_MINUTES = {
    'minute': 1, '1minute': 1,
    '3minute': 3,
    '5minute': 5,
    '10minute': 10,
    '15minute': 15,
    '30minute': 30,
    '60minute': 60, 'hour': 60,
}


def _interval_to_timedelta(interval):
    """Maps a Kite interval string to the candle's own duration, so callers
    can distinguish "this candle has STARTED" (its open timestamp <= now)
    from "this candle has CLOSED" (open timestamp + duration <= now) --
    only the latter is safe to treat as final data. Raises rather than
    silently guessing if an interval this pipeline doesn't already know
    about is passed in -- a wrong guess here would silently let partial
    candles back through the exact guard this section exists to enforce.
    """
    key = interval.strip().lower()
    if key in _INTERVAL_MINUTES:
        return timedelta(minutes=_INTERVAL_MINUTES[key])
    if key in ('day', '1day'):
        return timedelta(days=1)
    digits = ''.join(ch for ch in key if ch.isdigit())
    if digits and 'minute' in key:
        return timedelta(minutes=int(digits))
    raise ValueError(
        f"_interval_to_timedelta: unrecognized interval '{interval}' -- "
        f"add its duration to _INTERVAL_MINUTES before using it here. "
        f"Refusing to guess, since a wrong guess would silently let a "
        f"still-forming candle pass the closed-candle guard."
    )


def _drop_unclosed_candles(df, interval, cutoff, label, context):
    """Shared close-time filter used by both the full backfill and the
    incremental fetch below. Keeps only candles whose OPEN time + candle
    duration is <= cutoff (i.e. candles that have actually finished
    forming) and drops anything still-forming or, defensively, dated
    after cutoff outright. `cutoff` may be None, in which case this is a
    no-op (BACKTEST path, where the whole historical day is legitimately
    wanted and there is no "now" to compare against).

    [CHANGED -- log volume] This filter itself is NOT optional and is NOT
    being removed: dropping the still-forming tail candle is what stops
    the documented repainting bug (a 9:15 candle showing BUY CE, then
    flipping to BUY PE once it actually closes at 9:20) -- see the
    module-level note above. What WAS noisy is that every LIVE cycle
    fetches ~1 fresh candle per symbol, of which the newest one is *by
    definition* still-forming almost every time -- so this used to print
    one [GUARD] line per symbol, per cycle, forever (see the 48-line
    block in a single incremental sync). That's expected, routine
    behavior, not something worth a line per symbol. So this function no
    longer prints for the routine case (dropped <= 1); the caller
    aggregates the return value into a single one-line per-cycle summary
    instead (see update_incremental_data / download_historical_data).
    Dropping MORE than 1 candle for a symbol in one fetch is NOT routine
    (it means a real gap -- e.g. this symbol's poll was skipped for over
    a candle's length) and is still flagged immediately, per-symbol,
    right here, since that's a signal worth seeing in real time.
    Returns (filtered_df, dropped_count).
    """
    if cutoff is None or df.empty:
        return df, 0
    candle_duration = _interval_to_timedelta(interval)
    before = len(df)
    closed_mask = (_strip_tz(df.index) + candle_duration) <= cutoff
    df = df[closed_mask]
    dropped = before - len(df)
    if dropped > 1:
        print(f"  [GUARD] {label}: dropped {dropped} still-forming/unclosed "
              f"candle(s) (interval={interval}) as of {cutoff.strftime('%H:%M:%S')} "
              f"IST ({context}) -- more than the routine 1-candle tail trim, "
              f"worth checking for a missed cycle/gap for this symbol.")
    return df, dropped


# --------------------------------------------------------------------------
# AVAILABILITY CHECK -- lets callers skip the full 90-day backfill when
# today's data has already been downloaded (e.g. pipeline re-run/restart).
# --------------------------------------------------------------------------

def historical_data_exists(df_ref, target_date, interval='5minute', hist_dir=HIST_DIR, min_ratio=1.0):
    """Checks whether download_historical_data() has already run for
    target_date, by looking for each valid-token symbol's {interval} CSV
    on disk instead of re-downloading 90 days of history unconditionally.

    Returns True when at least `min_ratio` (default 100%) of valid-token
    symbols already have their file present. Lower min_ratio if a few
    missing/late-added symbols shouldn't force a full re-backfill.
    """
    date_str = target_date.strftime('%d-%b-%y')
    valid_tokens = df_ref.dropna(subset=['Zerodha_Token'])
    if valid_tokens.empty:
        return False

    total = len(valid_tokens)
    present = 0
    for _, row in valid_tokens.iterrows():
        sym = str(row['Symbol / StrikePrice']).strip().upper()
        path = os.path.join(hist_dir, f"{sym}_{interval}_{date_str}.csv")
        if os.path.exists(path):
            present += 1

    return (present / total) >= min_ratio


# --------------------------------------------------------------------------
# FULL BACKFILL -- run ONCE per day (pre-market), not on every live cycle.
# --------------------------------------------------------------------------

def fetch_historical_worker(sym, token, from_date, to_date, kite_api, date_str, intervals, limiter, now_cutoff=None):
    """Returns the total number of still-forming candles dropped across
    all requested intervals for this symbol, so the caller can fold it
    into one per-run summary line instead of a per-symbol print."""
    dropped_total = 0
    for interval in intervals:
        limiter.acquire()
        try:
            hist_data = kite_api.historical_data(int(token), from_date, to_date, interval)
            df_hist = pd.DataFrame(hist_data)
            if df_hist.empty:
                continue

            df_hist['date'] = pd.to_datetime(df_hist['date'])
            df_hist.set_index('date', inplace=True)

            # [FIX -- 13-Jul-26] CLOSE-time guard, not open-time guard --
            # see the module-level "CANDLE-CLOSE AWARENESS" note above for
            # why the old `index <= now_cutoff` check let a still-forming
            # candle through. Under normal LIVE conditions Kite typically
            # won't hand back a genuinely future-dated row, but it WILL
            # (per Kite's own forum) hand back the currently-forming
            # candle when `to_date` is close to "now" -- which is exactly
            # every LIVE call this function makes. This is what stops a
            # partial candle from ever reaching disk, regardless of a
            # stale/cached response, a wrong to_date from a caller, a
            # clock issue, or simply calling this mid-candle.
            df_hist, dropped = _drop_unclosed_candles(
                df_hist, interval, now_cutoff, sym, "full backfill"
            )
            dropped_total += dropped
            if df_hist.empty:
                continue

            save_path = os.path.join(HIST_DIR, f"{sym}_{interval}_{date_str}.csv")
            df_hist.to_csv(save_path)
        except Exception as e:
            print(f"[ERROR] Data fetch failed for {sym} at {interval}: {e}")
    return dropped_total


def download_historical_data(df_ref, target_date, kite_api, intervals=('5minute',), max_workers=4, cap_to_now=False):
    """One-time (or once-per-trading-day) full backfill.

    Only fetches the interval(s) actually consumed downstream. Every
    indicator module (SQZMOM, RSI, BRKPRO, ADX, TW ALL) reads
    ONLY '5minute' data -- '15minute' / '60minute' / 'day' were being
    downloaded on every run and never read by anything. Default changed to
    ('5minute',) accordingly; pass the old 4-tuple explicitly if a future
    indicator needs a higher timeframe.

    Call this ONCE per session (e.g. from a pre-market setup step). For the
    recurring per-candle updates during live trading, use
    update_incremental_data() below instead -- it fetches only the new
    candles rather than re-pulling 90 days every time.

    cap_to_now: pass True for a LIVE session so the backfill can never pull
    (or accept) candles past the current wall-clock time. LEAVE FALSE for
    BACKTEST, where target_date is a past, fully-closed session and the
    whole day's candles are legitimately wanted in one pass. Without this,
    a LIVE call using the old unconditional 23:59:59 end-of-day cutoff is
    what let a stale/leftover file (e.g. from an earlier BACKTEST pass
    against the same calendar date) get silently adopted as "today's live
    data" -- showing candles hours ahead of the actual market.
    """
    print("[SYSTEM] Initiating Full Historical Backfill (one-time / once-per-day)...")

    date_str = target_date.strftime('%d-%b-%y')
    valid_tokens = df_ref.dropna(subset=['Zerodha_Token'])
    from_date = target_date - timedelta(days=90)
    fetch_to_date = target_date.replace(hour=23, minute=59, second=59)

    now_cutoff = None
    if cap_to_now:
        # [FIX] was datetime.now() -- host local time, not necessarily
        # IST. NSE runs on IST regardless of what timezone the machine
        # executing this script is set to.
        now = now_ist()
        fetch_to_date = min(fetch_to_date, now)
        now_cutoff = now
        print(f"[SYSTEM] cap_to_now=True -- backfill capped at {now_cutoff.strftime('%H:%M:%S')} IST "
              f"(no candle beyond this will be requested or saved, and any candle that hasn't "
              f"FULLY CLOSED as of this timestamp will be dropped too).")

    limiter = RateLimiter(max_per_second=3)

    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        future_to_sym = {
            executor.submit(
                fetch_historical_worker,
                row['Symbol / StrikePrice'].strip().upper(),
                int(row['Zerodha_Token']),
                from_date,
                fetch_to_date,
                kite_api,
                date_str,
                list(intervals),
                limiter,
                now_cutoff,
            ): row['Symbol / StrikePrice'].strip().upper() for _, row in valid_tokens.iterrows()
        }

        dropped_total = 0
        for future in as_completed(future_to_sym):
            sym = future_to_sym[future]
            try:
                dropped_total += future.result()
                print(f"  -> SUCCESS: Extracted data for {sym}")
            except Exception as exc:
                print(f"  [CRITICAL] {sym} worker exception: {exc}")

    if dropped_total:
        print(f"[SYSTEM] Full backfill: {dropped_total} still-forming candle(s) "
              f"trimmed across all symbols (routine tail-trim under cap_to_now, "
              f"not an error).")


# --------------------------------------------------------------------------
# INCREMENTAL UPDATE -- run on EVERY live candle close.
# --------------------------------------------------------------------------

# [CHANGED -- 13-Jul-26] 15 -> 20 minutes. Kite's own forum reports the
# newly-closed candle's OHLC can be silently revised for anywhere from a
# few seconds up to ~30 minutes after it closes (see
# https://kite.trade/forum/discussion/14530). The re-fetch window here is
# relative to the LAST SAVED candle's own timestamp, and that timestamp
# advances every cycle -- so once a correction lands more than
# OVERLAP_MINUTES after its own candle closed, it stops being covered by
# any future cycle's window. 20 minutes (4 candles at 5-min) is a safer
# margin than 15 against that reported worst case without meaningfully
# increasing API load; raise it further if your own logs show late
# corrections still being missed.
DEFAULT_OVERLAP_MINUTES = 20


def fetch_incremental_worker(sym, token, kite_api, interval, csv_path, limiter, overlap_minutes=DEFAULT_OVERLAP_MINUTES):
    """Fetches only the delta since the last candle already on disk for this
    symbol, instead of re-downloading history. A small overlap window
    (default 20 min = 4 candles at 5min) is re-pulled and merged on top of
    the existing rows so a late correction to a recently-closed candle
    (Kite sometimes finalizes OHLC slightly after the raw close -- see the
    module-level note above) gets picked up too, without duplicating rows.

    Returns the number of still-forming candles this call dropped (almost
    always 0 or 1 -- see _drop_unclosed_candles), so the caller can report
    one aggregate number per cycle instead of a line per symbol.
    """
    limiter.acquire()
    # [FIX] was datetime.now() -- host local time. This value is used both
    # as the Kite API's to_date (must be IST, since that's what NSE/Kite
    # operate in) and as the future-candle guard below, so it has to be
    # the real IST "now" regardless of the host machine's own clock.
    now = now_ist()

    existing = None
    if os.path.exists(csv_path):
        try:
            existing = pd.read_csv(csv_path, parse_dates=['date'], index_col='date')
        except Exception:
            existing = None  # unreadable/corrupt -- fall through to a fresh catch-up fetch

    if existing is not None and not existing.empty:
        last_ts = _strip_tz(existing.index.max())
        from_date = last_ts - timedelta(minutes=overlap_minutes)
    else:
        # No local file yet for today (new symbol, or first call before the
        # daily backfill has run) -- self-heal with a one-off full pull for
        # just this symbol instead of silently skipping it.
        from_date = now - timedelta(days=90)

    hist_data = kite_api.historical_data(int(token), from_date, now, interval)
    df_new = pd.DataFrame(hist_data)
    if df_new.empty:
        return 0  # nothing new yet (e.g. called before the candle actually closed)

    df_new['date'] = pd.to_datetime(df_new['date'])
    df_new.set_index('date', inplace=True)

    # [FIX -- 13-Jul-26] CLOSE-time guard, not open-time guard -- see the
    # module-level "CANDLE-CLOSE AWARENESS" note above. This call's own
    # `now` is passed as `to_date` above, so Kite is explicitly being
    # asked "everything up to right now" -- which, per Kite's own forum,
    # is precisely the request shape that returns the still-forming
    # candle as the last row. The old guard (`index <= now`) only checked
    # whether that candle had STARTED, which is always true for the
    # currently-forming bar the instant it begins. Never let a candle
    # that hasn't actually finished forming reach the CSV that the
    # indicator engines and the workbook read from. This filter itself is
    # NOT being removed -- only the per-symbol print of its routine case
    # was; see _drop_unclosed_candles' docstring.
    df_new, dropped = _drop_unclosed_candles(df_new, interval, now, sym, "incremental sync")
    if df_new.empty:
        return dropped

    combined = pd.concat([existing, df_new]) if existing is not None else df_new
    combined = combined[~combined.index.duplicated(keep='last')].sort_index()
    combined.to_csv(csv_path)
    return dropped


def update_incremental_data(df_ref, target_date, kite_api, interval='5minute', max_workers=4,
                             overlap_minutes=DEFAULT_OVERLAP_MINUTES):
    """The recurring live-loop step: pulls only new candles per symbol and
    appends/merges them into the existing CSV rather than re-downloading
    everything. Call this once per candle close.
    """
    print(f"[SYSTEM] Incremental sync ({interval}) starting...")
    date_str = target_date.strftime('%d-%b-%y')
    valid_tokens = df_ref.dropna(subset=['Zerodha_Token'])
    limiter = RateLimiter(max_per_second=3)
    t0 = time.time()
    ok, failed, dropped_total = 0, 0, 0

    def _job(row):
        sym = row['Symbol / StrikePrice'].strip().upper()
        token = int(row['Zerodha_Token'])
        csv_path = os.path.join(HIST_DIR, f"{sym}_{interval}_{date_str}.csv")
        return fetch_incremental_worker(sym, token, kite_api, interval, csv_path, limiter, overlap_minutes)

    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        future_to_sym = {
            executor.submit(_job, row): row['Symbol / StrikePrice'].strip().upper()
            for _, row in valid_tokens.iterrows()
        }
        for future in as_completed(future_to_sym):
            sym = future_to_sym[future]
            try:
                dropped_total += future.result()
                ok += 1
            except Exception as exc:
                failed += 1
                print(f"  [CRITICAL] {sym} incremental worker exception: {exc}")

    elapsed = time.time() - t0
    # [CHANGED -- 13-Jul-26] One line instead of one [GUARD] line per
    # symbol. Every cycle routinely drops ~1 still-forming candle per
    # symbol (see _drop_unclosed_candles) -- that's the filter doing its
    # job, not a problem, so it no longer needs a dedicated line per
    # symbol every ~5 minutes. Anything ABOVE the routine 1-per-symbol
    # case still gets its own immediate [GUARD] line from
    # _drop_unclosed_candles itself, since that does indicate a real gap.
    print(f"[SYSTEM] Incremental sync complete: {ok} ok, {failed} failed, "
          f"{dropped_total} still-forming candle(s) trimmed (routine), {elapsed:.1f}s elapsed.")


def file_has_future_candles(csv_path, now):
    """True if csv_path's last saved candle is timestamped after `now`.

    Used by a LIVE session before it trusts an already-on-disk file (the
    "skip full backfill, data already exists" shortcut). A file can only
    legitimately reach 'now' or earlier under LIVE conditions -- anything
    later means it's stale/leftover (e.g. a BACKTEST pass run earlier
    against the same calendar date, which intentionally pulls the whole
    day), not real candle-by-candle LIVE data, and should not be adopted
    as-is.
    """
    try:
        df = pd.read_csv(csv_path, parse_dates=['date'], index_col='date')
        if df.empty:
            return False
        return _strip_tz(df.index.max()) > now
    except Exception:
        return False  # unreadable -- let the normal self-heal path handle it


def purge_interval_data(df_ref, target_date, interval='5minute', hist_dir=HIST_DIR):
    """Deletes every symbol's {interval} CSV for target_date. Used to clear
    a stale/future-dated cache before forcing a clean, now-capped rebuild,
    rather than merging new data on top of data that shouldn't be there.
    """
    date_str = target_date.strftime('%d-%b-%y')
    valid_tokens = df_ref.dropna(subset=['Zerodha_Token'])
    removed = 0
    for _, row in valid_tokens.iterrows():
        sym = str(row['Symbol / StrikePrice']).strip().upper()
        path = os.path.join(hist_dir, f"{sym}_{interval}_{date_str}.csv")
        if os.path.exists(path):
            os.remove(path)
            removed += 1
    print(f"[SYSTEM] Purged {removed} stale {interval} file(s) for {date_str}.")


def load_interval_data(df_ref, target_date, interval='5minute', hist_dir=HIST_DIR):
    """Reads back the CSVs that download_historical_data() / update_incremental_data()
    wrote to hist_dir for every symbol in df_ref, for one interval.

    This is the shared data source for the indicator steps (SQZMOM, RSI, etc).
    Note it reads from disk directly rather than from the master tracker
    Excel file -- nothing in this pipeline writes OHLC data INTO that Excel
    workbook, so pd.read_excel(final_excel_path, sheet_name=None) will only
    ever return the 'Reference' sheet. Returns {symbol: DataFrame}; symbols
    with no matching file are reported and skipped rather than raising.
    """
    date_str = target_date.strftime('%d-%b-%y')
    valid_tokens = df_ref.dropna(subset=['Zerodha_Token'])
    data_dict = {}
    for _, row in valid_tokens.iterrows():
        sym = str(row['Symbol / StrikePrice']).strip().upper()
        path = os.path.join(hist_dir, f"{sym}_{interval}_{date_str}.csv")
        if os.path.exists(path):
            data_dict[sym] = pd.read_csv(path)
        else:
            print(f"[WARNING] load_interval_data: no {interval} file for {sym} on {date_str} at '{path}'.")
    return data_dict


# ---------------------------------------------------------------------------
# Disclaimer: this module only decides WHICH candles are safe to persist as
# closed, historical data. It makes no trading decision itself. Every
# downstream indicator/confluence/order module built on top of it should
# still be paper-traded and run through scripts/backtester.py before any
# of this is trusted with real capital.
# ---------------------------------------------------------------------------