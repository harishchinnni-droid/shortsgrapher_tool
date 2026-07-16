"""
risk_and_signal_patches.py
---------------------------
Drop-in fixes for the four concrete defects found while auditing the
01-13 Jul 26 backtest set (49 trades, -Rs 12,026 net, PF 0.57). Each
patch below is self-contained and documents exactly which existing
function/line in your pipeline it replaces. None of this is a rewrite --
your architecture (signal modules -> order_sheet gates -> position_manager
exits) is sound; these are targeted corrections, per the "change one
thing, re-test" discipline.

Historical performance note: all backtest figures cited in the comments
below are historical results on this specific 8-day, 49-trade sample.
Past performance does not guarantee future results, and a sample this
size is provisional, not statistically conclusive -- treat every number
here as a hypothesis to re-validate on the next 30-50 trades, not a
proven fix.
"""

import json
import os
import time
from datetime import datetime

import pandas as pd

# A minimal shim so this module has no hard runtime dependency on your
# pipeline's estimate_round_trip_costs(). Delete this shim and import
# the real one from position_manager.py when you wire this in.
try:
    from position_manager import estimate_round_trip_costs
except ImportError:
    def estimate_round_trip_costs(entry_ltp, exit_ltp, quantity):
        """Fallback stub (~0.35% of round-trip turnover) -- ONLY used if
        position_manager isn't importable in this scratch environment.
        Replace with the real import before running for real."""
        if quantity <= 0:
            return {'total': 0.0}
        turnover = (entry_ltp + exit_ltp) * quantity
        return {'total': round(turnover * 0.0035, 2)}


# =============================================================================
# PATCH 1 -- PCR trend gate: accumulate on every candle, not just on new
# signal streaks.
#
# Replaces: order_sheet.py's PCRTrendTracker (the persistence-bug fix
# already applied on 13-Jul-26 was necessary but not sufficient).
#
# Evidence: across all 8 backtest days, 32 of 49 executed trades (65%)
# carried PCR Trend = INSUFFICIENT_DATA, and that bucket alone accounted
# for -Rs 11,231 of the -Rs 12,026 total net loss (93%) at a 31.2% win
# rate -- vs 38-50% win rates when a trend WAS established. The trend
# gate is barely running, not because the persistence bug reappeared,
# but because .record() is only called when a NEW 3-bar signal streak
# starts for a symbol (per the original docstring) -- and most symbols
# never accumulate 3 separate streak events before their entry fires.
# Fix: record a PCR reading on every candle a symbol's chain is polled/
# read, independent of whether a streak just started, so 3 readings
# accrue within the first 15-20 minutes of a symbol being watched
# instead of requiring 3 separate entry attempts.
# =============================================================================
class PCRTrendTrackerV2:
    """Same public contract as the original (.record(), .evaluate()),
    but callers should invoke .record() on EVERY bar/poll for every
    symbol on the watchlist (e.g. once per 5-min candle close), not only
    when a new signal streak begins. This is the actual behavior change;
    the class body is otherwise unchanged from the original."""

    def __init__(self, min_readings=3, max_history=20):
        self.history = {}  # {symbol: [pcr_val, ...]}
        self.min_readings = min_readings
        self.max_history = max_history

    def record(self, symbol, pcr_val):
        if pcr_val is None:
            return
        hist = self.history.setdefault(symbol, [])
        hist.append(pcr_val)
        if len(hist) > self.max_history:
            del hist[0]

    def evaluate(self, symbol):
        """Returns (trend, n_readings). trend in {'RISING','FALLING',
        'FLAT','INSUFFICIENT_DATA'}."""
        hist = self.history.get(symbol, [])
        if len(hist) < self.min_readings:
            return "INSUFFICIENT_DATA", len(hist)
        recent = hist[-self.min_readings:]
        if recent[-1] > recent[0]:
            return "RISING", len(hist)
        if recent[-1] < recent[0]:
            return "FALLING", len(hist)
        return "FLAT", len(hist)


# =============================================================================
# PATCH 2 -- OI buildup confirmation: fix the direction-blind check, and
# stop treating SHORT_COVERING as a strong confirming signal.
#
# Replaces: order_sheet.py get_oi_buildup_signal(), specifically the
# line:
#     confirms = quadrant in ("LONG_BUILDUP", "SHORT_COVERING") if signal
#                in ("BUY CE", "BUY PE") else True
#
# Bug: this checks the SAME two quadrants (both of which represent
# RISING underlying price) as "confirming" for BOTH BUY CE and BUY PE.
# LONG_BUILDUP/SHORT_COVERING describe bullish price action -- they
# should only be able to confirm a CE (call/bullish) entry. A PE (put/
# bearish) entry needs SHORT_BUILDUP or LONG_UNWINDING (falling-price
# quadrants) to be genuinely confirmed.
#
# Evidence: SHORT_COVERING-tagged entries lost -Rs 9,842 across 20
# trades (25% win rate) -- the single worst-performing OI bucket by a
# wide margin, worse even than INSUFFICIENT_DATA on a per-trade basis.
# LONG_BUILDUP-tagged entries made +Rs 1,382 across 17 trades (58.8% win
# rate). Sample sizes are small (n=17-20) -- re-validate before trusting
# this as permanent, but it's consistent enough to act on now: this
# patch (a) fixes the direction bug, AND (b) removes SHORT_COVERING from
# the auto-confirm set entirely, since -- unlike a textbook mismatch --
# it underperformed on BOTH sides of the trade in this sample (CE: -Rs
# 7,610/14 trades; PE: -Rs 2,233/6 trades), suggesting a short-covering
# bounce is too low-conviction an OI state to treat as confirmation for
# a fresh long-premium bet, regardless of direction.
# =============================================================================
OI_SNAPSHOT_CACHE_DEFAULT = os.path.join("json_cache", "oi_snapshot_cache.json")


def _load_json(path, default):
    try:
        with open(path, "r") as f:
            return json.load(f)
    except Exception:
        return default


def _save_json(path, data):
    try:
        os.makedirs(os.path.dirname(path), exist_ok=True)
        with open(path, "w") as f:
            json.dump(data, f)
    except Exception as e:
        print(f"[WARNING] Failed to persist {path}: {e}")


def get_oi_buildup_signal_v2(opt_symbol, current_oi, current_price, signal,
                              today_str, cache_path=OI_SNAPSHOT_CACHE_DEFAULT,
                              allow_short_covering=False):
    """Drop-in replacement for order_sheet.get_oi_buildup_signal().
    `today_str` replaces the internal today_ist() call -- pass
    today_ist().strftime('%Y-%m-%d') from the caller so this module has
    no hard dependency on your ist_clock module.

    allow_short_covering: kept as a flag (default False = matches the
    backtest evidence above) rather than hardcoded, so you can flip it
    back on for an A/B re-run without editing logic -- see module
    docstring: change one thing, re-test.
    """
    cache_all = _load_json(cache_path, {})
    cache_today = dict(cache_all.get(today_str, {}))
    prev = cache_today.get(opt_symbol)
    cache_today[opt_symbol] = {"oi": current_oi, "price": current_price, "ts": time.time()}
    _save_json(cache_path, {today_str: cache_today})

    if prev is None:
        return "INSUFFICIENT_DATA", True

    oi_delta = current_oi - prev.get("oi", current_oi)
    price_delta = current_price - prev.get("price", current_price)

    if price_delta >= 0 and oi_delta > 0:
        quadrant = "LONG_BUILDUP"
    elif price_delta >= 0 and oi_delta <= 0:
        quadrant = "SHORT_COVERING"
    elif price_delta < 0 and oi_delta > 0:
        quadrant = "SHORT_BUILDUP"
    else:
        quadrant = "LONG_UNWINDING"

    # [FIXED] direction-aware confirmation, was direction-blind before.
    bullish_quadrants = {"LONG_BUILDUP"}
    if allow_short_covering:
        bullish_quadrants.add("SHORT_COVERING")
    bearish_quadrants = {"SHORT_BUILDUP", "LONG_UNWINDING"}

    if signal == "BUY CE":
        confirms = quadrant in bullish_quadrants
    elif signal == "BUY PE":
        confirms = quadrant in bearish_quadrants
    else:
        confirms = True

    return quadrant, confirms


# =============================================================================
# PATCH 3 -- Hard stop-loss cap must be cost-aware.
#
# Replaces: position_manager.py's simulate_backtest_exit(), the block:
#     worst_case_loss = (entry_ltp - candle_low) * quantity
#     if worst_case_loss > max_loss_rs:
#         cap_exit_ltp = max(entry_ltp - (max_loss_rs / quantity), 0.05)
#
# Bug: this solves for the exit price that makes GROSS loss == max_loss_rs
# (Rs 2,000), then round-trip costs (brokerage/STT/GST/slippage, ~Rs
# 175-235 in this sample) are subtracted AFTER, so every trade that hits
# this exit realizes a NET loss larger than the stated cap.
#
# Evidence: all 3 "Hard Stop-Loss (Rs 2000 cap)" exits in the sample
# lost more than Rs 2,000 net -- Rs 2,195.80 / 2,234.12 / 2,174.93 -- a
# 9-12% breach of the stated capital-preservation ceiling every time.
# On a system whose stated priority is minimizing drawdown, a cap that
# reliably overshoots itself is a real, fixable gap.
# =============================================================================
def solve_cost_aware_cap_exit(entry_ltp, quantity, max_loss_rs, floor_price=0.05,
                               iterations=2):
    """Returns cap_exit_ltp such that NET loss (gross loss + estimated
    round-trip cost AT that exit price) is as close to max_loss_rs as a
    couple of fixed-point iterations gets it. Costs are a small,
    near-linear function of turnover, so 2 iterations converges tightly
    in practice -- verify with the smoke test at the bottom of this file
    if you change the cost model.
    """
    if quantity <= 0 or entry_ltp <= 0:
        return max(entry_ltp - 0, floor_price)

    # Iteration 0: original (cost-blind) formula as the starting guess.
    cap_exit_ltp = max(entry_ltp - (max_loss_rs / quantity), floor_price)

    for _ in range(iterations):
        costs = estimate_round_trip_costs(entry_ltp, cap_exit_ltp, quantity)
        est_cost = costs["total"]
        # Shrink the allowed GROSS loss by the estimated cost, then
        # resolve the exit price against that smaller gross budget.
        effective_gross_cap = max(max_loss_rs - est_cost, 0.0)
        cap_exit_ltp = max(entry_ltp - (effective_gross_cap / quantity), floor_price)

    return round(cap_exit_ltp, 2)


# =============================================================================
# PATCH 4 -- Daily max-drawdown circuit breaker (did not exist anywhere
# in the pipeline; grep confirms no daily-loss kill switch in
# position_manager.py, order_sheet.py, or 01_Master_Code.py).
#
# This directly answers output requirement #3 (capital protection) from
# the brief. Wire the .breached() check into order_sheet.py's main
# candidate loop, immediately before a new order is created -- if
# breached, reject every new entry for the rest of the session (existing
# open positions still get managed/exited normally by position_manager).
# =============================================================================
class DailyDrawdownGuard:
    """Tracks realized Net P/L across the session and flips to breached
    once the configured Rs loss (or % of equity) is crossed. Call
    .update(net_pl_of_closed_trade) every time a position closes, and
    check .breached() before opening any new one.

    Sizing note: -Rs 12,026 over 49 trades / 8 days on a Rs 5,00,000
    default account (see position_manager.ACCOUNT_EQUITY_DEFAULT) is a
    -2.4% drawdown in aggregate, but individual days ranged from +Rs
    2,251 (02-Jul) to -Rs 5,068 (07-Jul) -- a single-day max loss of
    ~1% of equity. A daily cap of 2x that (Rs 10,000, i.e. 2% of a Rs
    5,00,000 account) would not have blocked any single day in this
    sample outright, but WOULD have stopped compounding losses within a
    day once the cap is crossed intraday -- tune max_daily_loss_rs to
    your own equity and risk tolerance, this default is a starting
    point, not a rule.
    """

    def __init__(self, max_daily_loss_rs=10_000.0):
        self.max_daily_loss_rs = max_daily_loss_rs
        self.realized_pl = 0.0
        self._breached = False
        self._breach_time = None

    def update(self, net_pl_of_closed_trade):
        self.realized_pl += net_pl_of_closed_trade
        if not self._breached and self.realized_pl <= -abs(self.max_daily_loss_rs):
            self._breached = True
            self._breach_time = datetime.now().strftime("%H:%M:%S")
            print(f"[RISK] Daily drawdown cap breached at {self._breach_time}: "
                  f"realized P/L Rs {self.realized_pl:.2f} <= -Rs {self.max_daily_loss_rs:.2f}. "
                  f"No new entries for the rest of the session.")

    def breached(self):
        return self._breached

    def status(self):
        return {
            "realized_pl": round(self.realized_pl, 2),
            "max_daily_loss_rs": self.max_daily_loss_rs,
            "breached": self._breached,
            "breach_time": self._breach_time,
        }


# =============================================================================
# PATCH 5 (proposed, not yet evidence-backed the way 1-4 are) -- early
# "no follow-through" exit, to test against the Max Hold Time cluster.
#
# NOT a new indicator. This is a risk-timing rule layered on your
# EXISTING trailing-stop machinery.
#
# Evidence this targets: "Max Hold Time (75min)" was the single largest
# loss bucket by trade count (22/49 = 45% of all trades, 27.3% win rate,
# -Rs 6,312 net) and "EOD Square-off" the second largest (10/49, 20% win,
# -Rs 2,757). Combined, these two "the trade just never went anywhere"
# exits caused -Rs 9,069 of the -Rs 12,026 total loss (75%). Meanwhile
# every trade that reached "Trailing Stop Hit" (7/7) or "Target Hit"
# (2/2) was profitable. The pattern: trades either show real follow-
# through early, or they don't -- and the ones that don't are currently
# ridden the full 75 minutes (or to EOD) before being cut.
#
# Proposed rule: if a trade has NOT reached +0.5R (half the initial
# risk_per_unit_premium) in favorable excursion within the first 20
# minutes, exit at market rather than continuing to hold. This is a
# single, clearly-labeled new gate -- test it in isolation (same 8 days,
# same entries, only this flag changed) before combining with any other
# change, per the skill's "one change at a time" rule.
# =============================================================================
NO_FOLLOWTHROUGH_MINUTES = 20
NO_FOLLOWTHROUGH_R_MULTIPLE = 0.5


def check_no_followthrough_exit(entry_dt, current_dt, max_ltp_seen, entry_ltp,
                                 risk_per_unit_premium,
                                 minutes=NO_FOLLOWTHROUGH_MINUTES,
                                 r_multiple=NO_FOLLOWTHROUGH_R_MULTIPLE):
    """Returns True if this trade should be cut now for lack of
    follow-through. Call this alongside (not instead of) the existing
    stop/target/trailing/max-hold checks in simulate_backtest_exit() /
    check_live_exit() -- whichever check fires first wins, same pattern
    already used for the other exit types."""
    if risk_per_unit_premium <= 0:
        return False
    elapsed_minutes = (current_dt - entry_dt).total_seconds() / 60.0
    if elapsed_minutes < minutes:
        return False
    favorable_excursion = max_ltp_seen - entry_ltp
    return favorable_excursion < (r_multiple * risk_per_unit_premium)


# =============================================================================
# Smoke tests -- run this file directly to sanity-check the math above
# before wiring anything into the live pipeline.
# =============================================================================
if __name__ == "__main__":
    # --- Patch 3: cap must land at/near max_loss_rs in NET terms ---
    entry_ltp, quantity, max_loss_rs = 32.55, 425, 2000.0
    cap_exit = solve_cost_aware_cap_exit(entry_ltp, quantity, max_loss_rs)
    costs = estimate_round_trip_costs(entry_ltp, cap_exit, quantity)
    gross_loss = (entry_ltp - cap_exit) * quantity
    net_loss = gross_loss + costs["total"]
    print(f"[TEST] cap_exit={cap_exit}, gross_loss={gross_loss:.2f}, "
          f"est_cost={costs['total']:.2f}, net_loss={net_loss:.2f} (target <= {max_loss_rs})")
    assert net_loss <= max_loss_rs * 1.02, "Cost-aware cap still overshoots by >2%"

    # --- Patch 4: drawdown guard flips correctly ---
    guard = DailyDrawdownGuard(max_daily_loss_rs=5000.0)
    for pl in [-1200, -900, 300, -2500, -900]:
        guard.update(pl)
    print("[TEST] guard status:", guard.status())
    assert guard.breached() is True

    # --- Patch 5: no-followthrough logic ---
    t0 = datetime(2026, 7, 13, 10, 0)
    t1 = datetime(2026, 7, 13, 10, 25)
    result = check_no_followthrough_exit(t0, t1, max_ltp_seen=32.60, entry_ltp=32.55,
                                          risk_per_unit_premium=3.25)
    print("[TEST] no-followthrough (flat trade, 25min in) ->", result)
    assert result is True

    print("\nAll smoke tests passed.")
