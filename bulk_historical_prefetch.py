"""
bulk_historical_prefetch.py
-----------------------------
ONE-TIME (or occasional) bulk historical data downloader for BACKTESTING
(Task 61, 18-Jul-26). Harish's request: "we can use 1 set of data for all
our working instead of downloading again and again."

WHY THIS EXISTS
----------------
data_ingestion.download_historical_data() already pulls a 90-day trailing
window in ONE Kite call per symbol -- but it saves that result under a
filename keyed to only the SINGLE target_date being processed
(`{sym}_5minute_{date_str}.csv`). historical_data_exists() checks for that
exact date's own file, so every NEW date in a backtest range triggers its
own fresh 90-day backfill, ~99% overlapping the one before it. That is
the concrete, traced cause of "historical data is loading for everytime."

WHAT THIS SCRIPT DOES
----------------------
1. UNDERLYING (spot) data: fetches each symbol's 5-min history for the
   WHOLE requested date range ONCE (chunked into <=85-day windows -- a
   safety margin under the 90-day single-call window this codebase's own
   download_historical_data() already relies on), slices out each
   individual trading day, and saves each slice under the EXACT SAME
   per-date filename data_ingestion.py already reads. Zero changes needed
   to data_ingestion.py / order_sheet.py / run_pipeline.py -- they will
   simply find every date pre-cached and skip their own backfill.
   Already-cached days are skipped on a re-run, so this is safe to run
   again later to extend the range forward.

2. OPTIONS: reuses historical_lookup.fetch_option_day_candles() and
   get_vix_snapshot() VERBATIM (same functions order_sheet.py already
   calls, same disk cache format) -- nothing about how option/VIX data is
   fetched or stored is reimplemented here.

IMPORTANT LIMITATION -- READ BEFORE TRUSTING OLD-MONTH OPTION DATA
---------------------------------------------------------------------
Investigated 18-Jul-26: order_sheet.resolve_option_chain() and
historical_lookup.resolve_pcr_chain_tokens() both pick whichever expiry is
SOONEST among contracts Kite's LIVE instruments() list currently has --
with no check that it's the contract that actually existed on the
backtest date. Kite's instruments() endpoint only lists currently-
tradeable contracts; an already-expired monthly option is gone from that
list entirely, no historical-snapshot API exists to recover it. So for an
old backtest date, those two functions would silently hand back TODAY's
nearest contract (e.g. July's) as if it were the option that traded back
in April -- wrong data, not a clean failure.

Per Harish's explicit instruction (18-Jul-26), this script does NOT touch
resolve_option_chain() or resolve_pcr_chain_tokens() in the existing
pipeline files -- that known gap is left as-is for now. This script only
guards ITS OWN prefetching: _resolve_safe_chain_tokens() below refuses to
prefetch option data for a symbol/day whose nearest currently-listed
expiry is implausibly far (> MAX_EXPIRY_GAP_DAYS) from that day, and logs
exactly why it skipped rather than silently caching a wrong contract.
Practically, this means option prefetching will only succeed for dates
within roughly the current expiry cycle -- the run's own console output
tells you exactly where that boundary actually falls, rather than this
comment guessing at it. UNDERLYING/spot data has no such limitation and
is fetched for the full requested range regardless.

Respects the same 3 req/sec RateLimiter the rest of this pipeline uses --
concurrency here overlaps idle/wasted time, it does not exceed Kite's
real rate limit. A multi-month, multi-symbol first run is genuinely slow
(expect tens of minutes); re-runs are fast since cached days are skipped.

HOW TO RUN
-----------
    python bulk_historical_prefetch.py
Prompts for a start and end date (DD-MMM-YY), same format as the rest of
this pipeline.
"""

import os
import sys
import shutil
import tempfile
import traceback
import concurrent.futures
from datetime import datetime, timedelta

import pandas as pd

CODES_DIR = os.path.dirname(os.path.abspath(__file__))
if CODES_DIR not in sys.path:
    sys.path.append(CODES_DIR)

try:
    import broker_auth
    import calendar_mgmt
    import file_mgmt
    import token_mgmt
    import data_ingestion
    import historical_lookup
    from order_sheet import _build_kite_master
    from ist_clock import now_ist
except ImportError as e:
    print(f"[CRITICAL ERROR] Failed to import pipeline modules: {e}")
    print("Ensure this script is saved in the '05 Codes' directory with the rest of the pipeline.")
    sys.exit(1)

# --------------------------------------------------------------------------
# Config
# --------------------------------------------------------------------------
UNDERLYING_CHUNK_DAYS = 85       # safety margin under the proven 90-day single-call window
UNDERLYING_FETCH_WORKERS = 4     # symbols fetched concurrently (shared rate limiter keeps Kite calls paced)
OPTION_PREFETCH_WORKERS = 8      # option contracts fetched concurrently (same shared-limiter pattern as Task 60)
MAX_EXPIRY_GAP_DAYS = 40         # see module docstring -- generous upper bound for a monthly contract's own lifespan


# --------------------------------------------------------------------------
# Reference sheet (symbol universe + broker tokens) -- reuses existing
# provisioning logic but writes to a THROWAWAY temp copy, never to
# Harish's canonical 01_SourceFile.xlsx.
# --------------------------------------------------------------------------
def _load_df_ref(kite_api, as_of_date):
    tmp_path = os.path.join(tempfile.gettempdir(), "bulk_prefetch_reference.xlsx")
    if os.path.exists(file_mgmt.SOURCE_FILE):
        shutil.copy2(file_mgmt.SOURCE_FILE, tmp_path)
    else:
        print("[SYSTEM] No local source template found -- pulling 'Reference' from the Google Sheet template instead.")
        file_mgmt._build_reference_from_google_sheet(tmp_path)
    df_ref = token_mgmt.update_instrument_tokens(tmp_path, kite_api, target_date=as_of_date)
    return df_ref


# --------------------------------------------------------------------------
# STEP 1 -- Underlying (spot) 5-minute data, whole range, one call per symbol.
# --------------------------------------------------------------------------
def _chunk_ranges(start_date, end_date, chunk_days=UNDERLYING_CHUNK_DAYS):
    chunks = []
    cur = start_date
    while cur <= end_date:
        chunk_end = min(cur + timedelta(days=chunk_days - 1), end_date)
        chunks.append((cur, chunk_end))
        cur = chunk_end + timedelta(days=1)
    return chunks


def _fetch_symbol_underlying(sym, token, start_date, end_date, kite_api, limiter, trading_dates):
    date_strs = [d.strftime('%d-%b-%y') for d in trading_dates]
    missing_dates = [
        (d, ds) for d, ds in zip(trading_dates, date_strs)
        if not os.path.exists(os.path.join(data_ingestion.HIST_DIR, f"{sym}_5minute_{ds}.csv"))
    ]
    already_cached = len(trading_dates) - len(missing_dates)
    if not missing_dates:
        return sym, already_cached, len(trading_dates), "already fully cached"

    frames = []
    for c_start, c_end in _chunk_ranges(start_date, end_date):
        limiter.acquire()
        try:
            raw = kite_api.historical_data(
                int(token), c_start, c_end.replace(hour=23, minute=59, second=59), '5minute'
            )
            if raw:
                frames.append(pd.DataFrame(raw))
        except Exception as e:
            print(f"[WARNING] Bulk underlying fetch failed for {sym} "
                  f"{c_start:%d-%b-%y}-{c_end:%d-%b-%y}: {e}")

    if not frames:
        return sym, already_cached, len(trading_dates), "Kite returned no data for the whole range"

    full = pd.concat(frames, ignore_index=True)
    full['date'] = pd.to_datetime(full['date'])
    if full['date'].dt.tz is not None:
        full['date'] = full['date'].dt.tz_localize(None)
    full = full.drop_duplicates(subset='date').sort_values('date')
    full.set_index('date', inplace=True)

    saved = already_cached
    for d, ds in missing_dates:
        day_df = full[full.index.date == d.date()]
        if day_df.empty:
            continue
        save_path = os.path.join(data_ingestion.HIST_DIR, f"{sym}_5minute_{ds}.csv")
        day_df.to_csv(save_path)
        saved += 1
    return sym, saved, len(trading_dates), None


def bulk_fetch_underlying(df_ref, start_date, end_date, kite_api, trading_dates):
    limiter = data_ingestion.RateLimiter(max_per_second=3)
    valid = df_ref.dropna(subset=['Zerodha_Token'])
    print(f"\n[SYSTEM] STEP 1/2 -- Bulk underlying fetch: {len(valid)} symbol(s), "
          f"{len(trading_dates)} trading day(s) each, {UNDERLYING_FETCH_WORKERS} concurrent workers "
          f"(shared 3 req/sec rate limiter)...")

    results = {}
    with concurrent.futures.ThreadPoolExecutor(max_workers=UNDERLYING_FETCH_WORKERS) as executor:
        futures = {
            executor.submit(
                _fetch_symbol_underlying,
                str(row['Symbol / StrikePrice']).strip().upper(),
                row['Zerodha_Token'],
                start_date, end_date, kite_api, limiter, trading_dates,
            ): str(row['Symbol / StrikePrice']).strip().upper()
            for _, row in valid.iterrows()
        }
        for f in concurrent.futures.as_completed(futures):
            try:
                sym, saved, total, note = f.result()
                results[sym] = (saved, total)
                tag = f" ({note})" if note else ""
                print(f"  -> {sym}: {saved}/{total} trading day(s) cached{tag}.")
            except Exception as e:
                print(f"  [ERROR] {futures[f]} worker exception: {e}")
                print(traceback.format_exc())

    fully_covered = sum(1 for saved, total in results.values() if saved == total)
    print(f"[SUCCESS] STEP 1 complete -- {fully_covered}/{len(results)} symbol(s) fully covered "
          f"for the whole {start_date:%d-%b-%y} to {end_date:%d-%b-%y} range.")
    return results


# --------------------------------------------------------------------------
# STEP 2 -- Options + VIX, reusing historical_lookup.py's own fetch/cache
# functions verbatim. See module docstring for the expiry-safety caveat.
# --------------------------------------------------------------------------
def _resolve_safe_chain_tokens(base_symbol, spot_price, df_ref, kite_master, target_date,
                                strikes_each_side=8, strike_step_default=50):
    """Own copy of historical_lookup.resolve_pcr_chain_tokens()'s strike-
    band logic, PLUS a guard resolve_pcr_chain_tokens() itself does NOT
    have: refuses to trust a resolved expiry that's implausibly far from
    target_date (see module docstring -- Kite's live instrument list can
    only offer the SOONEST currently-listed expiry, which for an old
    backtest date is not necessarily the contract that was actually alive
    back then). Returns (tokens, expiry_or_None, skip_reason_or_None, gap_days_or_None) --
    gap_days is only populated when a nearest expiry was found at all (for
    the caller's own reporting/aggregation, see bulk_fetch_options())."""
    diff_val = strike_step_default
    match = df_ref[df_ref['Symbol / StrikePrice'].astype(str).str.strip().str.upper() == base_symbol]
    if not match.empty and 'Option Price Difference' in df_ref.columns:
        val = match['Option Price Difference'].values[0]
        if pd.notna(val) and val != 0:
            diff_val = val

    if not spot_price:
        return [], None, "no underlying spot price on disk for this day", None

    atm_strike = round(spot_price / diff_val) * diff_val
    strikes = [atm_strike + i * diff_val for i in range(-strikes_each_side, strikes_each_side + 1)]
    chain = kite_master[
        (kite_master['name'] == base_symbol)
        & (kite_master['segment'] == 'NFO-OPT')
        & (kite_master['strike'].isin(strikes))
    ]
    if chain.empty:
        return [], None, "no F&O contracts currently listed for this symbol/strike band at all", None

    nearest_expiry = pd.to_datetime(chain['expiry']).min()
    gap_days = (nearest_expiry.date() - target_date.date()).days
    if gap_days > MAX_EXPIRY_GAP_DAYS:
        # [FIXED -- stable category string, not one unique string per gap
        # value] The gap itself (41d, 42d, 43d, ...) used to be baked into
        # the reason string, so a real multi-symbol/multi-month range
        # produced one distinct dict key PER DAY instead of one aggregated
        # count -- unreadable at real scale (thousands of near-duplicate
        # lines). The gap is returned separately now so the caller can
        # summarize min/max once instead.
        return [], nearest_expiry, "contract's real expiry has already rolled off Kite's live instrument list", gap_days

    chain = chain[pd.to_datetime(chain['expiry']) == nearest_expiry]
    return list(zip(chain['instrument_token'].tolist(), chain['tradingsymbol'].tolist())), nearest_expiry, None, gap_days


def bulk_fetch_options(df_ref, kite_master, kite_api, trading_dates, symbols):
    hist_cache = historical_lookup.HistoricalCache(kite_api)

    print(f"\n[SYSTEM] STEP 2/2 -- Option chain resolution: {len(symbols)} symbol(s) x "
          f"{len(trading_dates)} trading day(s) (local, no network yet)...")

    token_days = {}     # token -> set of trading dates to fetch it for
    token_symbol = {}   # token -> tradingsymbol, for the summary line
    resolved_pairs = 0
    skip_reasons = {}          # reason -> count
    skip_gap_range = {}        # reason -> [min_gap, max_gap] (only meaningful for the expiry-gap reason)

    for d in trading_dates:
        for sym in symbols:
            spot = historical_lookup.get_spot_snapshot(sym, d, "15:25") \
                or historical_lookup.get_spot_snapshot(sym, d, "09:20")
            tokens, expiry, reason, gap_days = _resolve_safe_chain_tokens(sym, spot, df_ref, kite_master, d)
            if reason:
                skip_reasons[reason] = skip_reasons.get(reason, 0) + 1
                if gap_days is not None:
                    lo, hi = skip_gap_range.get(reason, (gap_days, gap_days))
                    skip_gap_range[reason] = (min(lo, gap_days), max(hi, gap_days))
                continue
            resolved_pairs += 1
            for token, tsym in tokens:
                token_days.setdefault(int(token), set()).add(d)
                token_symbol[int(token)] = tsym

    total_pairs = len(symbols) * len(trading_dates)
    print(f"[SYSTEM] Resolved {resolved_pairs}/{total_pairs} symbol-day pair(s) to a trustworthy option chain "
          f"-- {len(token_days)} unique contract(s) to fetch.")
    for reason, count in sorted(skip_reasons.items(), key=lambda x: -x[1]):
        gap_note = ""
        if reason in skip_gap_range:
            lo, hi = skip_gap_range[reason]
            gap_note = f" (gap ranged {lo}-{hi} day(s) beyond the {MAX_EXPIRY_GAP_DAYS}d cutoff)"
        print(f"  [SKIP] {count}x -- {reason}{gap_note}")

    if not token_days:
        print("[WARNING] Nothing resolvable to prefetch -- the whole requested range is likely outside "
              "the currently-listed expiry window (see skip reasons above). Underlying data from STEP 1 "
              "is still fully cached and usable regardless.")
        return

    total_contract_days = sum(len(days) for days in token_days.values())
    print(f"[SYSTEM] Fetching {len(token_days)} option contract(s), {total_contract_days} contract-day "
          f"combination(s) total, {OPTION_PREFETCH_WORKERS} concurrent workers (shared rate limiter, "
          f"same disk cache format order_sheet.py already reads)...")

    def _fetch_one(token):
        for d in token_days[token]:
            try:
                historical_lookup.fetch_option_day_candles(kite_api, token, d, hist_cache)
            except Exception as e:
                print(f"[WARNING] Option prefetch failed for {token_symbol.get(token, token)} "
                      f"on {d:%d-%b-%y}: {e}")

    with concurrent.futures.ThreadPoolExecutor(max_workers=OPTION_PREFETCH_WORKERS) as executor:
        futures = [executor.submit(_fetch_one, t) for t in token_days]
        for f in concurrent.futures.as_completed(futures):
            f.result()

    print(f"[SUCCESS] Option prefetch complete -- {len(hist_cache.option_candles)} contract-day(s) "
          f"now disk-cached under historical_lookup.OPTION_HIST_DIR.")

    print(f"[SYSTEM] Prefetching India VIX for {len(trading_dates)} trading day(s)...")
    vix_ok = 0
    for d in trading_dates:
        try:
            if historical_lookup.get_vix_snapshot(kite_api, d, "15:25", kite_master, hist_cache) is not None:
                vix_ok += 1
        except Exception as e:
            print(f"[WARNING] VIX prefetch failed for {d:%d-%b-%y}: {e}")
    print(f"[SUCCESS] VIX cached for {vix_ok}/{len(trading_dates)} trading day(s).")


# --------------------------------------------------------------------------
# Entry point
# --------------------------------------------------------------------------
def _prompt_date(label):
    while True:
        raw = input(f"Enter {label} (DD-MMM-YY, e.g., 01-Apr-26): ").strip()
        try:
            return datetime.strptime(raw, '%d-%b-%y')
        except ValueError:
            print("[ERROR] Invalid date format. Use DD-MMM-YY.")


def main():
    print("=" * 60)
    print("   BULK HISTORICAL PREFETCH -- one-time range download")
    print("=" * 60)

    try:
        kite_api = broker_auth.initialize_zerodha()
    except Exception as e:
        print(f"\n[FATAL] Zerodha authentication failed: {e}")
        sys.exit(1)

    start_date = _prompt_date("Start Date")
    end_date = _prompt_date("End Date")
    trading_dates = calendar_mgmt.get_trading_dates_in_range(start_date, end_date)
    if not trading_dates:
        print("[ERROR] No trading days in that range -- every date was a weekend/holiday.")
        sys.exit(1)

    as_of = now_ist()
    df_ref = _load_df_ref(kite_api, as_of)
    symbols = [str(s).strip().upper() for s in df_ref.dropna(subset=['Zerodha_Token'])['Symbol / StrikePrice']]

    bulk_fetch_underlying(df_ref, start_date, end_date, kite_api, trading_dates)

    print("\n[SYSTEM] Building Kite instrument master for option resolution (fetched fresh, TODAY's live listing)...")
    kite_master = _build_kite_master(kite_api)

    bulk_fetch_options(df_ref, kite_master, kite_api, trading_dates, symbols)

    print("\n" + "=" * 60)
    print("  BULK PREFETCH COMPLETE")
    print("=" * 60)
    print("Re-run your normal backtest (01_Master_Code.py / 02_Master_Code_3Indicator.py) for any date "
          "in this range -- it will find everything pre-cached and skip its own downloads. "
          "Re-running THIS script later with a wider date range is safe: already-cached days are skipped.")


if __name__ == "__main__":
    main()
