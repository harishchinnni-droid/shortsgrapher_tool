"""
historical_lookup.py
---------------------
Originally BACKTEST-only replacements for every live kite_api.quote()
call that order_sheet.py's gates otherwise depend on. It was wrong for
BACKTEST: a backtest for 06-Jul-26 run on the evening of 10-Jul-26 would
have every one of those gates evaluated against 10-Jul-26's post-close
quote, not against the actual market conditions at each historical
signal's own timestamp. This was confirmed in an actual multi-day
backtest output that motivated this fix: the same PCR value appeared
unchanged across four different intraday timestamps for the same symbol,
and a large share of rejections were "Entry LTP Rs.0.00" -- a live
quote() call returning nothing outside its own live session, not a
genuinely dead option.

[CHANGED -- Task 65, 20-Jul-26] get_historical_pcr() and get_historical_
oi_buildup() (plus the get_option_snapshot() they both share) now also
run in LIVE mode, via a new is_live=True path -- Harish's explicit
request once he saw that BACKTEST's PCR/OI-buildup gates compare actual
option candles (current bar vs. one N-minutes-earlier) while LIVE's own
versions (calculate_local_pcr() / get_oi_buildup_signal(), still defined
below but no longer called by order_sheet.py) instead compared a live
quote() snapshot against whatever the LAST POLL of that exact contract
happened to be -- a materially different method, not just a different
data source, since "last poll" isn't a fixed time window and depends on
how often that specific contract came up. LIVE now uses the SAME
candle-comparison method as BACKTEST for these two gates specifically;
see fetch_option_live_candles() below for how that's sourced safely for
an ongoing (not yet closed) trading day. Entry LTP, spot price, and VIX
are UNCHANGED -- still live kite_api.quote() calls in order_sheet.py --
this only applies to the PCR and OI-buildup gates.

Every function below returns a value AS OF a specific historical
(target_date, time_str) pair, sourced from Kite's historical_data()
endpoint -- which Kite serves for option contracts and index instruments
(India VIX) exactly the same way it does for equities, given the right
instrument_token -- or, for the underlying's own spot price, straight off
the 5-minute CSVs data_ingestion.py already downloaded to disk in Step 5
of the pipeline (no network call needed at all for that one).

Caveats, stated plainly:
  - Kite's historical F&O intraday retention window is materially shorter
    than its equity window (rolling weeks, not years). A backtest for a
    date outside that window will get None back from the option-history
    fetches here and should be expected to fail gracefully (see
    order_sheet.py's handling of a None return), not silently produce
    fabricated numbers.
  - The instrument MASTER (which strikes/expiries exist) is still fetched
    live via kite_api.instruments() in order_sheet.py, unchanged. That
    listing is materially static day-to-day and fixing its historicity
    too was judged out of scope for this pass.
  - 5-minute candle closes are used as the LTP proxy at each timestamp
    (Kite has no historical tick-by-tick LTP feed) -- the same resolution
    the rest of this pipeline's indicators already operate at.

Caching: [CHANGED] two-tier now. In-memory per (token, date) for the life
of one HistoricalCache instance, same as before -- the same option
contract's full trading day is typically looked up several times as a
symbol's signal history is scanned, and PCR alone can touch ~10 strikes.
NEW: a disk-level cache under OPTION_HIST_DIR, keyed the same way. This
was the actual bottleneck behind slow backtest re-runs -- the in-memory
cache alone only helps WITHIN one run; every fresh run of the SAME
backtest date (e.g. re-testing after a code change, which this project
has now done 3x for the same 5 dates) started from an empty cache and
re-fetched every option contract's candles, the VIX, live from Kite
again, one request every ~0.33s (the 3 req/sec limiter). A BACKTEST date
is a closed, immutable trading day -- once fetched, that data can never
change -- so disk-caching it is exactly as safe as data_ingestion.py
already does for the underlying's own 5-min CSVs, just extended to cover
option contracts and VIX too. First run of a given date still pays the
full live-fetch cost; every run after that reads straight off disk.
"""

import os

import pandas as pd

import data_ingestion  # HIST_DIR + the already-downloaded underlying 5-min CSVs
from ist_clock import now_ist

MARKET_OPEN = (9, 0)
MARKET_DATA_END = (15, 35)

# [ADDED] Disk cache for option-contract / VIX historical candles -- see
# module docstring above. Sibling directory to data_ingestion.HIST_DIR,
# same drive, so this survives just as long and needs no separate backup
# story. A BACKTEST date's data here never goes stale (the day is over),
# so unlike the underlying's own cache there is no "already exists but
# might be incomplete/live" ambiguity to check for.
OPTION_HIST_DIR = os.path.join(os.path.dirname(data_ingestion.HIST_DIR), "03_Option_Historical_Data")
os.makedirs(OPTION_HIST_DIR, exist_ok=True)


class HistoricalCache:
    """One instance per backtest run. Avoids re-fetching the same
    instrument's full-day candle history once per gate check that touches
    it (a symbol can be checked at 20+ different timestamps in one day).
    [CHANGED] The in-memory dicts here are now backed by OPTION_HIST_DIR
    on disk (see fetch_option_day_candles() / get_vix_snapshot()), so a
    cache MISS in a fresh instance can still be a fast disk read instead
    of a live Kite call."""

    def __init__(self, kite_api, max_requests_per_second=3):
        self.option_candles = {}   # (token, date_str) -> DataFrame
        self.vix_candles = {}      # date_str -> DataFrame
        self.limiter = data_ingestion.RateLimiter(max_per_second=max_requests_per_second)
        self._vix_token = None


# ---------------------------------------------------------------------------
# Small shared helpers
# ---------------------------------------------------------------------------
def time_str_to_dt(target_date, time_str):
    hh, mm = map(int, time_str.split(':'))
    return target_date.replace(hour=hh, minute=mm, second=0, microsecond=0)


def _candle_at_or_before(df, target_dt):
    """Last candle at or before target_dt, or None if the day's data
    doesn't reach that far back yet (e.g. the very first candle of the
    session, or the fetch itself failed / returned nothing)."""
    if df is None or df.empty:
        return None
    idx = df.index[df.index <= target_dt]
    if len(idx) == 0:
        return None
    return df.loc[idx.max()]


def _day_bounds(target_date):
    from_dt = target_date.replace(hour=MARKET_OPEN[0], minute=MARKET_OPEN[1], second=0, microsecond=0)
    to_dt = target_date.replace(hour=MARKET_DATA_END[0], minute=MARKET_DATA_END[1], second=0, microsecond=0)
    return from_dt, to_dt


# ---------------------------------------------------------------------------
# Option contract history (drives entry LTP, exit-simulation candles,
# volume/OI liquidity gates, and the OI-buildup gate)
# ---------------------------------------------------------------------------
def fetch_option_day_candles(kite_api, opt_token, target_date, cache, interval='5minute'):
    """Full day of 5-min candles (OHLC + OI) for one option contract,
    cached so a symbol checked repeatedly through a day only costs ONE
    Kite historical_data() call for that contract, not one per check --
    and, once that call has happened on ANY prior run for this exact
    (token, date), never again after that (see disk-cache block below)."""
    if opt_token is None:
        return pd.DataFrame()

    date_str = target_date.strftime('%Y-%m-%d')
    key = (int(opt_token), date_str)
    if key in cache.option_candles:
        return cache.option_candles[key]

    # [ADDED] Disk cache -- checked before touching the network at all.
    # target_date is always a closed BACKTEST day here, so a hit is always
    # valid; there's no "might be stale" case to guard against.
    disk_path = os.path.join(OPTION_HIST_DIR, f"{int(opt_token)}_{interval}_{date_str}.csv")
    empty_marker = disk_path + ".empty"
    if os.path.exists(disk_path):
        try:
            df = pd.read_csv(disk_path, parse_dates=['date'], index_col='date')
            cache.option_candles[key] = df
            return df
        except Exception as e:
            print(f"[WARNING] historical_lookup: disk cache read failed for token {opt_token} ({e}) -- refetching live.")
    elif os.path.exists(empty_marker):
        # A previous run already confirmed Kite has nothing for this
        # contract/date (e.g. outside the F&O historical retention window)
        # -- don't spend another live call re-discovering that.
        df = pd.DataFrame()
        cache.option_candles[key] = df
        return df

    from_dt, to_dt = _day_bounds(target_date)
    cache.limiter.acquire()
    try:
        raw = kite_api.historical_data(int(opt_token), from_dt, to_dt, interval, oi=True)
        df = pd.DataFrame(raw)
        if not df.empty:
            df['date'] = pd.to_datetime(df['date'])
            if df['date'].dt.tz is not None:
                df['date'] = df['date'].dt.tz_localize(None)
            df.set_index('date', inplace=True)
            df.sort_index(inplace=True)
            df.to_csv(disk_path)  # [ADDED] persist -- this date's data will never change
        else:
            open(empty_marker, 'w').close()  # [ADDED] remember "confirmed empty" too
    except Exception as e:
        print(f"[WARNING] historical_lookup: option history fetch failed for token {opt_token} ({e}). "
              f"Likely outside Kite's F&O historical retention window, or the contract didn't exist yet.")
        df = pd.DataFrame()

    cache.option_candles[key] = df
    return df


def fetch_option_live_candles(kite_api, opt_token, cache, interval='5minute'):
    """[ADDED -- Task 65, 20-Jul-26] LIVE equivalent of
    fetch_option_day_candles() above -- same shape of result (a candle
    DataFrame, OI included), same in-memory cache dict, but sourced
    safely for an ONGOING trading day instead of a closed BACKTEST one.
    Two differences from the BACKTEST version, both required because
    today is not immutable the way every backtest date is:

      1. Fetches from market open through NOW (not a fixed end-of-day
         to_date), and drops the still-forming tail candle the same way
         data_ingestion.py's own full backfill does for the underlying
         (see that module's "CANDLE-CLOSE AWARENESS" note) -- without
         this, a signal could get graded against a candle that hasn't
         actually finished forming yet, the exact repainting bug that
         guard exists to prevent.
      2. In-memory cached ONLY, scoped to the one HistoricalCache
         instance a single build_order_sheet() LIVE call creates for
         itself (a fresh instance every 5-minute cycle, same lifetime
         as everything else that call computes) -- NEVER written to
         OPTION_HIST_DIR's disk cache, since a partial trading day has
         to be re-fetched fresh every cycle, not frozen after its first
         partial fetch the way a genuinely closed backtest day safely
         can be.

    If opt_token is None, or the fetch fails/returns nothing, returns an
    empty DataFrame -- callers already handle that the same way they do
    for fetch_option_day_candles()."""
    if opt_token is None:
        return pd.DataFrame()

    key = (int(opt_token), 'LIVE')
    if key in cache.option_candles:
        return cache.option_candles[key]

    now = now_ist()
    from_dt = now.replace(hour=MARKET_OPEN[0], minute=MARKET_OPEN[1], second=0, microsecond=0)
    cache.limiter.acquire()
    try:
        raw = kite_api.historical_data(int(opt_token), from_dt, now, interval, oi=True)
        df = pd.DataFrame(raw)
        if not df.empty:
            df['date'] = pd.to_datetime(df['date'])
            if df['date'].dt.tz is not None:
                df['date'] = df['date'].dt.tz_localize(None)
            df.set_index('date', inplace=True)
            df.sort_index(inplace=True)
            df, _dropped = data_ingestion._drop_unclosed_candles(
                df, interval, now, f"option token {opt_token}", "LIVE PCR/OI-buildup gate"
            )
    except Exception as e:
        print(f"[WARNING] historical_lookup: LIVE option candle fetch failed for token {opt_token} ({e}).")
        df = pd.DataFrame()

    cache.option_candles[key] = df
    return df


def get_option_snapshot(kite_api, opt_token, target_date, time_str, cache, interval='5minute', is_live=False):
    """{'close','high','low','volume','oi'} for opt_token AT OR BEFORE
    time_str on target_date, or None if unavailable. [CHANGED -- Task 65]
    is_live=True sources candles via fetch_option_live_candles() instead
    of fetch_option_day_candles() -- see that function's docstring."""
    df = (fetch_option_live_candles(kite_api, opt_token, cache, interval) if is_live
          else fetch_option_day_candles(kite_api, opt_token, target_date, cache, interval))
    row = _candle_at_or_before(df, time_str_to_dt(target_date, time_str))
    if row is None:
        return None
    return {
        'close': float(row.get('close', 0) or 0),
        'high': float(row.get('high', 0) or 0),
        'low': float(row.get('low', 0) or 0),
        'volume': float(row.get('volume', 0) or 0),
        'oi': float(row.get('oi', 0) or 0),
    }


# [ADDED] Feature flag for the OI-confirmation fix below -- default False
# matches the 06-13 Jul 26 backtest audit evidence (SHORT_COVERING lost
# money on both CE and PE). Set True to restore the original (direction-
# corrected but SHORT_COVERING-inclusive) behavior for an A/B re-run.
ALLOW_SHORT_COVERING_CONFIRM = False


def get_historical_oi_buildup(kite_api, opt_token, target_date, time_str, cache, signal,
                               lookback_minutes=5, is_live=False):
    """Historical replacement for order_sheet.get_oi_buildup_signal().

    The original LIVE version (order_sheet.get_oi_buildup_signal(), still
    defined there but no longer called as of Task 65) compared this
    poll's OI+price against a JSON snapshot cache keyed only by
    opt_symbol with NO DATE in the key -- fine-ish for a single live
    session, but replaying several backtest dates for the same recurring
    contract would silently compare one day's OI against a DIFFERENT
    day's OI, and even within one live session "prior" meant "whenever
    this contract was last polled" rather than a fixed window. This
    version instead compares two points on the SAME day's own candle
    sequence (current bar vs. `lookback_minutes` earlier), which is
    unambiguous regardless of how many backtest dates get run, needs no
    persisted cache file, and can't leak across runs or drift with
    however often a contract happens to get checked.

    [CHANGED -- Task 65, 20-Jul-26] is_live=True runs this exact same
    comparison for LIVE too (Harish's request), sourcing both the
    current and prior candle from fetch_option_live_candles() instead of
    fetch_option_day_candles() -- see that function's docstring for why
    an ongoing trading day needs a different (in-memory-only, closed-
    candle-filtered) fetch than a closed BACKTEST date.
    """
    current = get_option_snapshot(kite_api, opt_token, target_date, time_str, cache, is_live=is_live)
    if current is None:
        return "INSUFFICIENT_DATA", True

    prior_dt = time_str_to_dt(target_date, time_str) - pd.Timedelta(minutes=lookback_minutes)
    df = (fetch_option_live_candles(kite_api, opt_token, cache) if is_live
          else fetch_option_day_candles(kite_api, opt_token, target_date, cache))
    prior_row = _candle_at_or_before(df, prior_dt)
    if prior_row is None:
        return "INSUFFICIENT_DATA", True

    oi_delta = current['oi'] - float(prior_row.get('oi', 0) or 0)
    price_delta = current['close'] - float(prior_row.get('close', 0) or 0)

    if price_delta >= 0 and oi_delta > 0:
        quadrant = "LONG_BUILDUP"
    elif price_delta >= 0 and oi_delta <= 0:
        quadrant = "SHORT_COVERING"
    elif price_delta < 0 and oi_delta > 0:
        quadrant = "SHORT_BUILDUP"
    else:
        quadrant = "LONG_UNWINDING"

    # [FIX -- risk_and_signal_patches audit] Was direction-blind: both
    # LONG_BUILDUP and SHORT_COVERING represent RISING underlying price,
    # so the old check let either one confirm BOTH BUY CE and BUY PE --
    # a bullish OI state was waving through bearish bets. Now CE needs a
    # bullish quadrant, PE needs a bearish one.
    #
    # SHORT_COVERING is also dropped from the auto-confirm set entirely
    # (not just direction-corrected): the 06-13 Jul 26 backtest audit
    # showed SHORT_COVERING-tagged entries lost money on BOTH sides of
    # the trade (BUY CE: -Rs 7,610/14 trades, 28.6% win; BUY PE: -Rs
    # 2,233/6 trades, 16.7% win), while LONG_BUILDUP-tagged entries were
    # the one OI bucket that was net profitable (+Rs 1,382/17 trades,
    # 58.8% win). A short-covering bounce appears too low-conviction to
    # treat as confirmation regardless of direction. Sample is small
    # (n=17-20) -- re-validate on a larger set. Toggle back on via
    # ALLOW_SHORT_COVERING_CONFIRM for an A/B re-run without touching
    # this logic.
    bullish_quadrants = {"LONG_BUILDUP"}
    if ALLOW_SHORT_COVERING_CONFIRM:
        bullish_quadrants.add("SHORT_COVERING")
    bearish_quadrants = {"SHORT_BUILDUP", "LONG_UNWINDING"}

    if signal == "BUY CE":
        confirms = quadrant in bullish_quadrants
    elif signal == "BUY PE":
        confirms = quadrant in bearish_quadrants
    else:
        confirms = True
    return quadrant, confirms


# ---------------------------------------------------------------------------
# Underlying spot price -- read straight off the already-downloaded 5-min
# CSV (Step 5 of the pipeline), no network call needed.
# ---------------------------------------------------------------------------
def get_spot_snapshot(base_symbol, target_date, time_str, hist_dir=None):
    hist_dir = hist_dir or data_ingestion.HIST_DIR
    date_str = target_date.strftime('%d-%b-%y')
    path = os.path.join(hist_dir, f"{base_symbol}_5minute_{date_str}.csv")
    if not os.path.exists(path):
        return None
    try:
        df = pd.read_csv(path, parse_dates=['date'])
        if df['date'].dt.tz is not None:
            df['date'] = df['date'].dt.tz_localize(None)
        df.set_index('date', inplace=True)
        df.sort_index(inplace=True)
    except Exception as e:
        print(f"[WARNING] historical_lookup: failed reading spot CSV for {base_symbol} ({e}).")
        return None
    row = _candle_at_or_before(df, time_str_to_dt(target_date, time_str))
    return float(row['close']) if row is not None else None


# ---------------------------------------------------------------------------
# India VIX
# ---------------------------------------------------------------------------
def _resolve_vix_token(kite_master, cache):
    if cache._vix_token is not None:
        return cache._vix_token
    match = kite_master[
        (kite_master['segment'] == 'INDICES')
        & (kite_master['name'].astype(str).str.upper().str.contains('INDIA VIX'))
    ]
    token = int(match.iloc[0]['instrument_token']) if not match.empty else None
    cache._vix_token = token
    return token


def get_vix_snapshot(kite_api, target_date, time_str, kite_master, cache, interval='5minute'):
    """India VIX AT time_str on target_date, or None if the VIX
    instrument token can't be resolved or the history fetch fails."""
    token = _resolve_vix_token(kite_master, cache)
    if token is None:
        print("[WARNING] historical_lookup: could not resolve India VIX instrument token -- VIX gate skipped.")
        return None

    date_str = target_date.strftime('%Y-%m-%d')
    if date_str not in cache.vix_candles:
        # [ADDED] Disk cache, same pattern/reasoning as fetch_option_day_candles().
        disk_path = os.path.join(OPTION_HIST_DIR, f"VIX_{token}_{interval}_{date_str}.csv")
        if os.path.exists(disk_path):
            try:
                cache.vix_candles[date_str] = pd.read_csv(disk_path, parse_dates=['date'], index_col='date')
            except Exception as e:
                print(f"[WARNING] historical_lookup: VIX disk cache read failed ({e}) -- refetching live.")
                cache.vix_candles[date_str] = None  # fall through to live fetch below
        if date_str not in cache.vix_candles or cache.vix_candles.get(date_str) is None:
            from_dt, to_dt = _day_bounds(target_date)
            cache.limiter.acquire()
            try:
                raw = kite_api.historical_data(token, from_dt, to_dt, interval)
                df = pd.DataFrame(raw)
                if not df.empty:
                    df['date'] = pd.to_datetime(df['date'])
                    if df['date'].dt.tz is not None:
                        df['date'] = df['date'].dt.tz_localize(None)
                    df.set_index('date', inplace=True)
                    df.sort_index(inplace=True)
                    df.to_csv(disk_path)  # [ADDED] persist -- this date's VIX data will never change
            except Exception as e:
                print(f"[WARNING] historical_lookup: VIX history fetch failed ({e}).")
                df = pd.DataFrame()
            cache.vix_candles[date_str] = df

    row = _candle_at_or_before(cache.vix_candles[date_str], time_str_to_dt(target_date, time_str))
    return float(row['close']) if row is not None else None


# ---------------------------------------------------------------------------
# Historical PCR -- ATM +/- N strikes, OI-based, AT a specific timestamp.
# Mirrors order_sheet.calculate_local_pcr's LIVE logic exactly, just
# sourced from each strike's own historical OI candle instead of a live
# quote() snapshot.
#
# Reliability note (carried into order_sheet.py's PCRTrendTracker): PCR
# on a SINGLE stock is materially noisier than index PCR -- one large
# trade can skew it -- and is best read as a trend across several
# readings rather than a single absolute-value cutoff. This function
# returns one point-in-time reading; see order_sheet.py for how a
# sequence of these gets turned into a trend-based gate.
# ---------------------------------------------------------------------------
def resolve_pcr_chain_tokens(base_symbol, spot_price, df_ref, kite_master,
                              strikes_each_side=5, strike_step_default=50):
    """[ADDED -- Task 60, concurrent prefetch] Returns [(instrument_token,
    tradingsymbol), ...] for the ATM +/- N strike chain get_historical_pcr
    uses -- WITHOUT fetching any candle data. Pulled out of
    get_historical_pcr() itself so both that function and the concurrent
    prefetch pass in order_sheet.py resolve the exact same chain from the
    exact same logic; they can never drift out of sync since one calls
    the other."""
    diff_val = strike_step_default
    match = df_ref[df_ref['Symbol / StrikePrice'].astype(str).str.strip().str.upper() == base_symbol]
    if not match.empty and 'Option Price Difference' in df_ref.columns:
        val = match['Option Price Difference'].values[0]
        if pd.notna(val) and val != 0:
            diff_val = val

    if not spot_price:
        return []

    atm_strike = round(spot_price / diff_val) * diff_val
    strikes = [atm_strike + i * diff_val for i in range(-strikes_each_side, strikes_each_side + 1)]
    chain = kite_master[
        (kite_master['name'] == base_symbol)
        & (kite_master['segment'] == 'NFO-OPT')
        & (kite_master['strike'].isin(strikes))
    ]
    if chain.empty:
        return []

    current_expiry = chain.sort_values('expiry').iloc[0]['expiry']
    chain = chain[chain['expiry'] == current_expiry]
    return list(zip(chain['instrument_token'].tolist(), chain['tradingsymbol'].tolist()))


def get_historical_pcr(base_symbol, spot_price, target_date, time_str, df_ref, kite_master,
                        kite_api, cache, strikes_each_side=5, strike_step_default=50, is_live=False):
    """[CHANGED -- Task 65, 20-Jul-26] is_live=True sources every strike's
    OI from fetch_option_live_candles() instead of fetch_option_day_
    candles() (via get_option_snapshot's own is_live passthrough) -- same
    put/call OI summation either way, just a LIVE-safe candle source.
    Replaces order_sheet.calculate_local_pcr() (still defined there but
    no longer called as of this change) for LIVE too, per Harish's
    request that LIVE's PCR gate use the same method BACKTEST's already
    did."""
    if not spot_price:
        return None

    chain_tokens = resolve_pcr_chain_tokens(base_symbol, spot_price, df_ref, kite_master,
                                             strikes_each_side, strike_step_default)
    if not chain_tokens:
        return None

    put_oi, call_oi = 0.0, 0.0
    for token, tsym in chain_tokens:
        snap = get_option_snapshot(kite_api, token, target_date, time_str, cache, is_live=is_live)
        if snap is None:
            continue
        tsym = str(tsym)
        if tsym.endswith('PE'):
            put_oi += snap['oi']
        elif tsym.endswith('CE'):
            call_oi += snap['oi']

    if call_oi == 0:
        return 2.0 if put_oi > 0 else None
    return put_oi / call_oi
