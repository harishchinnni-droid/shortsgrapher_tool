"""
position_manager.py
--------------------
Real position management for the order-sheet pipeline: ATR-based
stop-loss/target sizing, equity-risk-based position sizing, a cost model
for realistic net P/L, and the exit simulation order_sheet.py's original
design left out entirely. Its own old docstring said so outright: "Live
LTP polling / trailing-stop-loss management and the P&L dashboard ...
were NOT requested and are not included here" -- which is why every
trade in earlier backtests showed Entry LTP = Current LTP = Max LTP and
P/L = 0 for every single position. [NEW MODULE]

Design choices, stated plainly
-------------------------------
This system only ever BUYS premium (CE or PE) -- it never writes
options. For a premium buyer, loss is always "the premium erodes below
entry" and profit is always "the premium rises above entry", regardless
of whether the position is a CE or a PE. So stop-loss and target are
computed once, symmetrically, on the option's own premium -- there's no
separate long-above/short-below branch to get wrong.

ATR-to-premium translation: this pipeline has ATR on the UNDERLYING
(Wilder ATR(14), read from the 'HTF Bias' sheet's 'ATR Value' row -- see
adx_di.py / htf_bias.py) but no options-Greeks feed, so there's no real
delta to size a premium stop off of. resolve_option_chain() always picks
the strike nearest current spot (ATM), and ATM options sit close to
delta ~0.5 -- this module uses that as a STATED APPROXIMATION, not a
precise Greek. That's disclosed, not hidden, because it's the single
biggest source of imprecision below; swapping in a real IV/Black-Scholes
delta from the option chain is the natural next upgrade.

    premium_risk_per_unit = clamp(
        ATR_MULTIPLIER * underlying_atr * ATM_DELTA_APPROX,
        MIN_STOP_PCT * entry_ltp, MAX_STOP_PCT * entry_ltp
    )
    stop_ltp   = entry_ltp - premium_risk_per_unit
    target_ltp = entry_ltp + REWARD_MULTIPLE * premium_risk_per_unit

The clamp exists because a raw ATR-derived stop can come out absurdly
tight (near-zero ATR in a dead session) or absurdly wide (an ATR spike).
The clamp keeps every stop inside the 10-35% of premium band that
retail options-buying risk guides converge on, while letting volatility
(via ATR) decide WHERE in that band a given trade's stop sits, instead
of hardcoding one fixed percentage for every trade regardless of regime.

Position sizing follows the algo-trading-python skill's standard
formula (risk_amount / stop_distance), floored to whole lots since
options can't be bought fractionally -- see compute_position_size().

Cost model: Zerodha discount-broker defaults for Indian F&O options, as
of the Apr-2026 Budget STT revision. This is a config block, not
hardcoded math, because these rates change -- VERIFY against your own
broker's current published charges before trusting the net P/L this
produces (see the algo-trading-python skill's risk_and_backtest.md:
"check current rates rather than guessing, they change").
"""

import math
import os
from datetime import datetime, timedelta

import pandas as pd
import pytz

import historical_lookup

IST = pytz.timezone('Asia/Kolkata')

# ---------------------------------------------------------------------------
# Risk / sizing config
# ---------------------------------------------------------------------------
ACCOUNT_EQUITY_DEFAULT = 500_000.0   # override via run_order_sheet_step(account_equity=...)
RISK_PCT_PER_TRADE = 0.01            # 1% of equity risked per trade

# [ADDED -- Task 75, 23-Jul-26, Harish's request] The equity-risk formula
# below (risk_amount / (risk_per_unit x lot_size)) can size a position at
# more than 1 lot whenever a trade's ATR-based stop distance is tight
# relative to the Rs 2,000 risk budget -- e.g. HEROMOTOCO 23-Jul-26 sized
# to 3 lots (risk/unit Rs 3.8 x lot size 150 = Rs 570/lot, comfortably
# under Rs 2,000 three times over). Harish wants every trade capped at 1
# lot regardless of what the risk formula alone would allow -- simpler
# position management while this is still being tested/paper-traded.
# compute_position_size() already had a max_lots parameter (previously
# unused, defaulting to None/unlimited); this constant is now its
# default, so every caller gets the cap automatically without needing to
# pass it explicitly. Does NOT weaken the risk-based rejection itself --
# a trade whose risk/unit is so wide that even 1 lot would blow past the
# Rs 2,000 budget is still rejected (num_lots floors to 0 before this cap
# is ever applied), this only ever prevents sizing UP past 1 lot.
MAX_LOTS_PER_TRADE = 1

# [CHANGED] Rs 5,000 -> Rs 2,000 per user request. At the Rs 5,00,000
# default equity, RISK_PCT_PER_TRADE alone sized positions for a Rs 5,000
# loss at full stop -- confirmed too high by the 5-day retest (06-10 Jul,
# reversal-exit disabled): "Stop Loss Hit" was the single largest loss
# category, Rs -48,754 across 13 fires, averaging Rs -3,750/trade (already
# above the old Rs 5,000 budget once round-trip costs/slippage stack on
# top of the gross stop distance -- see estimate_round_trip_costs()).
# Used TWICE below, deliberately:
#   1. compute_position_size() caps the equity-based risk budget at this
#      figure, so a full ATR stop costs close to Rs 2,000 BY CONSTRUCTION,
#      regardless of how large account_equity is set (kept equity-scaled,
#      not a flat lot size, per the skill's sizing guidance -- this is a
#      ceiling on that scaled number, not a replacement for it).
#   2. simulate_backtest_exit() / check_live_exit() ALSO check this
#      directly against realized/unrealized rupee loss, priced off the
#      candle's actual low (backtest) or live LTP -- not the stop_ltp
#      price level -- so a single 5-min candle that gaps straight through
#      the ATR stop (the existing stop check assumes a fill exactly AT
#      stop_ltp, which is optimistic) still gets caught and closed. This
#      is the literal "hard currency cap, supersedes the ATR calculation"
#      backstop -- (1) alone doesn't cover a gap-through, this does.
MAX_LOSS_PER_TRADE_RS = 2000.0

ATR_MULTIPLIER = 1.5                 # skill reference default: stop = entry - 1.5x ATR
ATM_DELTA_APPROX = 0.5               # ATM option delta approximation -- see module docstring
MIN_STOP_PCT = 0.10                  # floor: stop no tighter than 10% of premium
MAX_STOP_PCT = 0.35                  # ceiling: stop no wider than 35% of premium
REWARD_MULTIPLE = 2.0                # target = entry + 2R (research: 1:2 to 1:3 is typical)
TRAIL_TRIGGER_R = 1.0                # once +1R in profit, start trailing
TRAIL_LOCK_FRACTION = 0.5            # lock in 50% of the running gain above entry

# [ADDED -- Task 72, 22-Jul-26, ENABLE_MULTI_TARGET_TRAIL, Harish's request]
# Today's single fixed target (REWARD_MULTIPLE = 2R = "Target 1" below) exits
# the WHOLE position the instant price touches it, even on a trade that's
# clearly still trending -- capping upside on exactly the trades that
# deserved to run further, which was Harish's specific complaint watching
# 22-Jul-26 (target hit, trailing stop never got a chance to prove itself
# because the position was already closed).
#
# Design (flag OFF by default, same as every other experimental exit-logic
# change in this file -- A/B backtest before trusting): instead of exiting
# at Target 1, TIGHTEN the trailing stop's lock fraction each time price
# reaches a further checkpoint (T1 -> T2 -> T3), and -- per Harish's own
# choice when asked -- do NOT hard-exit at T3 either. From T3 onward the
# position is governed ENTIRELY by the (now very tight) trailing stop, so
# a genuine trend day can run indefinitely while a fake breakout still
# gives back only TRAIL_LOCK_STAGE3 (not all) of its gain before closing.
# Target 1/2/3 LTP are still written to the Orders sheet as fixed
# reference price levels regardless of whether the flag is on, so Harish
# can see them either way.
ENABLE_MULTI_TARGET_TRAIL = False
TARGET_1_R = REWARD_MULTIPLE          # 2R -- unchanged from today's single target
TARGET_2_R = 3.0
TARGET_3_R = 4.0
TRAIL_LOCK_STAGE1 = 0.65              # lock fraction once Target 1 (2R) is reached
TRAIL_LOCK_STAGE2 = 0.80              # lock fraction once Target 2 (3R) is reached
TRAIL_LOCK_STAGE3 = 0.90              # lock fraction once Target 3 (4R) is reached -- no hard cap beyond this

# [ADDED -- 17-Jul-26, Harish's request] TRAIL_LOCK_FRACTION above locks
# in 50% of the running gain the INSTANT price touches that level -- a
# single wick/pullback candle is enough to close a trade that's still
# winning (HINDALCO, 17-Jul-26: exited 'Trailing Stop Hit' at 13:40 with
# the underlying then continuing favorably before recovering -- real
# money left on the table, see the 17-Jul-26 audit). This is the
# "time-delayed / N-bar confirmation" technique from trailing-stop
# research (require the adverse price to persist for a few bars before
# honoring the stop, instead of reacting to the very first touch) --
# sources: TrendSpider's ATR trailing-stop guide and the Volatility Box
# ATR/Chandelier/Keltner comparison both describe this as standard
# practice for filtering normal pullback noise from a genuine reversal.
#
# Applied ONLY to the TRAILING portion (trailing_stop > stop_ltp -- the
# trade already ran +1R and is giving some back) -- the original hard ATR
# stop_ltp and the MAX_LOSS_PER_TRADE_RS cap stay INSTANT either way.
# Those two exist to protect capital on a trade that's LOSING; delaying
# them to "hold with reason" would be a real risk-management regression,
# not an improvement -- only a trade already sitting on unrealized profit
# gets the extra patience.
#
# Off by default -- needs its own A/B backtest (same pattern as every
# other experimental flag in this codebase) before being trusted. There's
# evidence it would have helped THIS one trade, not proof it helps on
# average -- more patience also means more given-back profit on trades
# that really were reversing for good.
ENABLE_TSL_CONFIRMATION_HOLD = False
TSL_CONFIRMATION_BARS = 2            # consecutive candles/cycles closed at/below the trail before honoring it
MAX_HOLD_MINUTES = 75                # theta-decay cap -- exit on time even absent SL/target
EOD_SQUAREOFF_TIME = "15:15"         # hard square-off

# [ADDED] A 5-day / 106-trade backtest audit (06-10 Jul 26) showed 91% of
# exits (97/106) fired via "5M Reversal (Signal WAIT)" for a combined
# -Rs 35,008, while every single "Trailing Stop Hit" exit (6/106) was
# profitable. The confirmed 3-bar entry signal on this pipeline's 5-min
# timeframe was breaking down again almost immediately in most cases --
# before price had moved far enough to reach either the ATR stop or the
# 2R target -- so the reversal exit was cutting trades (winners and
# losers alike) before the ATR stop/target/trailing mechanism below ever
# got a chance to run. Set True to restore the original behavior for an
# A/B re-run; consumed in simulate_backtest_exit() and check_live_exit()
# below, and in order_sheet.py's streak-break block.
# TRADE-OFF, read before flipping this permanently: with this off, a
# trade that would previously have exited early and cheap (avg ~Rs 361
# per reversal exit) instead rides all the way to its full ATR stop
# (10-35% of premium, materially larger -- the one 'Stop Loss Hit' in
# this sample lost Rs 3,781) if it never recovers. Re-run the same 5
# backtest dates with this flag off and compare Net P/L / Profit Factor /
# Max Drawdown against the baseline before trusting it live.
ENABLE_SIGNAL_REVERSAL_EXIT = False

# [ADDED -- risk_and_signal_patches audit, 13-Jul-26] The Rs 2,000 hard
# cap above solves for the exit price at which GROSS loss == max_loss_rs,
# then round-trip costs are subtracted AFTER -- so every trade that hits
# it realizes a NET loss bigger than the stated cap. All 3 occurrences in
# the 06-13 Jul 26 sample breached it: -Rs 2,195.80 / 2,234.12 / 2,174.93
# against a Rs 2,000 ceiling (9-12% overshoot every time). Set False to
# restore the original cost-blind cap for an A/B re-run.
ENABLE_COST_AWARE_STOP_CAP = True

# [ADDED -- same audit] "Max Hold Time (75min)" was the single largest
# loss bucket by trade count in the sample (22/49 = 45% of trades, 27.3%
# win rate, -Rs 6,312 net) and "EOD Square-off" the second largest
# (10/49, 20% win, -Rs 2,757) -- combined, 75% of total losses came from
# trades that were simply held until time ran out. Meanwhile every trade
# that reached "Trailing Stop Hit" or "Target Hit" was profitable (9/9,
# 100%). This is a PROPOSED rule, not yet evidence-backed the way the
# cost-aware cap above is -- test it in isolation (same dates, only this
# flag changed) before trusting it. Set False to disable entirely.
ENABLE_NO_FOLLOWTHROUGH_EXIT = True
NO_FOLLOWTHROUGH_MINUTES = 20        # give the trade this long to show real movement
NO_FOLLOWTHROUGH_R_MULTIPLE = 0.5    # "real movement" = reached +0.5R favorable excursion

# [ADDED -- same audit] No daily loss circuit breaker existed anywhere in
# the pipeline (confirmed by grep across this file, order_sheet.py, and
# 01_Master_Code.py) despite capital preservation being the stated
# priority. -Rs 5,068.36 was the single worst day in the 8-day sample
# (07-Jul-26) on the Rs 5,00,000 default account below -- this default
# caps daily loss at roughly 2x that, tune to your own risk tolerance.
# See DailyDrawdownGuard below -- BACKTEST-mode limitation documented on
# the class itself, read it before assuming this blocks entries the same
# way in both modes.
DAILY_MAX_LOSS_RS = 10_000.0

# ---------------------------------------------------------------------------
# Cost model
# ---------------------------------------------------------------------------
BROKERAGE_PER_ORDER = 20.0           # flat fee, or a % of turnover if that's lower
BROKERAGE_PCT_CAP = 0.0003
STT_SELL_PCT = 0.0015                # 0.15% on sell-side premium (post Apr-2026 STT hike)
EXCHANGE_TXN_PCT = 0.00035           # NSE F&O options, approx, both legs
SEBI_FEE_PCT = 0.000001              # Rs 10/crore, both legs -- negligible but included
GST_PCT = 0.18                       # on (brokerage + exchange + SEBI charges)
STAMP_DUTY_BUY_PCT = 0.00003         # 0.003%, buy side only
SLIPPAGE_PCT_PER_LEG = 0.005         # 0.5% of premium per leg -- market-order slippage assumption


def _leg_brokerage(turnover):
    return min(BROKERAGE_PER_ORDER, turnover * BROKERAGE_PCT_CAP)


def estimate_round_trip_costs(entry_ltp, exit_ltp, quantity):
    """Full buy+sell cost breakdown for `quantity` UNITS (lots x lot_size,
    already multiplied) of a single option leg. Returns a dict with a
    'total' key so callers can show a full breakdown or just net it out."""
    if quantity <= 0:
        return {'brokerage': 0.0, 'exchange_txn': 0.0, 'sebi_fee': 0.0, 'stt': 0.0,
                'stamp_duty': 0.0, 'gst': 0.0, 'slippage': 0.0, 'total': 0.0}

    buy_turnover = entry_ltp * quantity
    sell_turnover = exit_ltp * quantity

    buy_brokerage = _leg_brokerage(buy_turnover)
    sell_brokerage = _leg_brokerage(sell_turnover)
    buy_exchange = buy_turnover * EXCHANGE_TXN_PCT
    sell_exchange = sell_turnover * EXCHANGE_TXN_PCT
    buy_sebi = buy_turnover * SEBI_FEE_PCT
    sell_sebi = sell_turnover * SEBI_FEE_PCT

    stt = sell_turnover * STT_SELL_PCT          # options: STT on sell side only
    stamp_duty = buy_turnover * STAMP_DUTY_BUY_PCT
    gst = GST_PCT * (buy_brokerage + sell_brokerage + buy_exchange + sell_exchange + buy_sebi + sell_sebi)
    slippage = (buy_turnover + sell_turnover) * SLIPPAGE_PCT_PER_LEG

    total = (buy_brokerage + sell_brokerage + buy_exchange + sell_exchange
             + buy_sebi + sell_sebi + stt + stamp_duty + gst + slippage)

    return {
        'brokerage': round(buy_brokerage + sell_brokerage, 2),
        'exchange_txn': round(buy_exchange + sell_exchange, 2),
        'sebi_fee': round(buy_sebi + sell_sebi, 2),
        'stt': round(stt, 2),
        'stamp_duty': round(stamp_duty, 2),
        'gst': round(gst, 2),
        'slippage': round(slippage, 2),
        'total': round(total, 2),
    }


def solve_cost_aware_cap_exit(entry_ltp, quantity, max_loss_rs, floor_price=0.05, iterations=2):
    """[ADDED] Returns cap_exit_ltp such that NET loss (gross + estimated
    round-trip cost AT that exit price) lands at/near max_loss_rs, instead
    of the old formula which only capped GROSS loss and let costs push the
    realized loss past the stated ceiling. Costs are a small, near-linear
    function of turnover, so 2 fixed-point iterations converges tightly --
    see the smoke test at the bottom of this file."""
    if quantity <= 0 or entry_ltp <= 0:
        return max(entry_ltp, floor_price)

    cap_exit_ltp = max(entry_ltp - (max_loss_rs / quantity), floor_price)
    for _ in range(iterations):
        est_cost = estimate_round_trip_costs(entry_ltp, cap_exit_ltp, quantity)['total']
        effective_gross_cap = max(max_loss_rs - est_cost, 0.0)
        cap_exit_ltp = max(entry_ltp - (effective_gross_cap / quantity), floor_price)
    return round(cap_exit_ltp, 2)


def _no_followthrough_triggered(entry_dt, current_dt, max_ltp_seen, entry_ltp,
                                 risk_per_unit_premium,
                                 minutes=NO_FOLLOWTHROUGH_MINUTES,
                                 r_multiple=NO_FOLLOWTHROUGH_R_MULTIPLE):
    """[ADDED] True if the trade has neither hit stop/target/trailing NOR
    shown at least r_multiple x risk_per_unit_premium of favorable
    excursion within `minutes` of entry -- i.e. it's dead weight, cut it
    rather than ride it to Max Hold / EOD. See ENABLE_NO_FOLLOWTHROUGH_EXIT
    docstring above for the evidence this targets."""
    if risk_per_unit_premium <= 0:
        return False
    elapsed_minutes = (current_dt - entry_dt).total_seconds() / 60.0
    if elapsed_minutes < minutes:
        return False
    favorable_excursion = max_ltp_seen - entry_ltp
    return favorable_excursion < (r_multiple * risk_per_unit_premium)


# ---------------------------------------------------------------------------
# ATR-based stop/target + equity-risk position sizing
# ---------------------------------------------------------------------------
def compute_stop_and_target(entry_ltp, underlying_atr, atr_multiplier=ATR_MULTIPLIER,
                             delta_approx=ATM_DELTA_APPROX, reward_multiple=REWARD_MULTIPLE,
                             min_stop_pct=MIN_STOP_PCT, max_stop_pct=MAX_STOP_PCT):
    """Returns (stop_ltp, target_ltp, risk_per_unit_premium). See module
    docstring for the ATR->premium translation and why it's clamped."""
    if not entry_ltp or entry_ltp <= 0:
        return 0.0, 0.0, 0.0

    raw_risk = atr_multiplier * (underlying_atr or 0) * delta_approx
    floor_risk = min_stop_pct * entry_ltp
    ceil_risk = max_stop_pct * entry_ltp
    risk_per_unit = min(max(raw_risk, floor_risk), ceil_risk)

    stop_ltp = round(max(entry_ltp - risk_per_unit, 0.05), 2)
    target_ltp = round(entry_ltp + reward_multiple * risk_per_unit, 2)
    return stop_ltp, target_ltp, round(risk_per_unit, 2)


def compute_multi_targets(entry_ltp, risk_per_unit_premium):
    """[ADDED -- Task 72, 22-Jul-26] Returns (target1_ltp, target2_ltp,
    target3_ltp) -- target1_ltp is identical to compute_stop_and_target()'s
    target_ltp (same TARGET_1_R = REWARD_MULTIPLE), included here too so
    all three checkpoints can be computed and written to the Orders sheet
    together in one call. See ENABLE_MULTI_TARGET_TRAIL's docstring above."""
    if not entry_ltp or entry_ltp <= 0 or risk_per_unit_premium <= 0:
        return 0.0, 0.0, 0.0
    t1 = round(entry_ltp + TARGET_1_R * risk_per_unit_premium, 2)
    t2 = round(entry_ltp + TARGET_2_R * risk_per_unit_premium, 2)
    t3 = round(entry_ltp + TARGET_3_R * risk_per_unit_premium, 2)
    return t1, t2, t3


def _target_stage_reached(gain, risk_per_unit_premium):
    """[ADDED -- Task 72, 22-Jul-26] 0/1/2/3 -- how many of the Target
    1/2/3 checkpoints the running gain (off max_ltp_seen, i.e. the best
    price seen so far, not the current price) has reached. Used to decide
    how tightly to trail -- see _trailing_lock_fraction() below."""
    if risk_per_unit_premium <= 0:
        return 0
    r_multiple = gain / risk_per_unit_premium
    if r_multiple >= TARGET_3_R:
        return 3
    if r_multiple >= TARGET_2_R:
        return 2
    if r_multiple >= TARGET_1_R:
        return 1
    return 0


def _trailing_lock_fraction(stage, multi_target_enabled):
    """[ADDED -- Task 72, 22-Jul-26] Lock fraction to apply to the running
    gain once TRAIL_TRIGGER_R has been crossed. Stage 0 (below Target 1)
    always uses the original TRAIL_LOCK_FRACTION regardless of the flag --
    ENABLE_MULTI_TARGET_TRAIL only changes what happens ONCE a trade has
    already reached Target 1 or further; it never makes the EARLY trail
    looser or tighter than it always was."""
    if not multi_target_enabled or stage <= 0:
        return TRAIL_LOCK_FRACTION
    return {1: TRAIL_LOCK_STAGE1, 2: TRAIL_LOCK_STAGE2, 3: TRAIL_LOCK_STAGE3}[stage]


def compute_position_size(account_equity, risk_per_unit_premium, lot_size,
                           risk_pct_per_trade=RISK_PCT_PER_TRADE, max_lots=MAX_LOTS_PER_TRADE,
                           max_loss_rs=MAX_LOSS_PER_TRADE_RS):
    """Risk-based sizing: quantity = floor(risk_amount / (risk_per_unit x
    lot_size)). Returns (num_lots, quantity, risk_amount_rupees).
    num_lots CAN be 0 if even one lot's risk exceeds the per-trade risk
    budget -- callers must treat 0 as 'skip this trade', never silently
    round up to 1 (that would blow past the stated risk_pct_per_trade).

    [CHANGED] risk_amount is now the equity-based figure CAPPED at
    max_loss_rs -- see MAX_LOSS_PER_TRADE_RS docstring above for why.

    [CHANGED -- Task 75, 23-Jul-26] max_lots now defaults to
    MAX_LOTS_PER_TRADE (1) instead of None/unlimited -- see that
    constant's docstring. Pass max_lots=None explicitly to restore
    unlimited equity-risk-based sizing if this default is ever loosened
    again."""
    risk_amount = min(account_equity * risk_pct_per_trade, max_loss_rs)
    if risk_per_unit_premium <= 0 or lot_size <= 0:
        return 0, 0, round(risk_amount, 2)

    risk_per_lot = risk_per_unit_premium * lot_size
    num_lots = math.floor(risk_amount / risk_per_lot) if risk_per_lot > 0 else 0
    if max_lots is not None:
        num_lots = min(num_lots, max_lots)
    quantity = num_lots * lot_size
    return int(num_lots), int(quantity), round(risk_amount, 2)


# ---------------------------------------------------------------------------
# BACKTEST exit simulation -- walks forward through the option's own
# historical 5-min candles from entry to whichever comes first: stop,
# target, trailing-stop, the original 5M-reversal signal exit, the
# max-hold-time cap, or end-of-day square-off. This is what turns
# 'Current LTP' / 'Max LTP' / 'P/L' from hardcoded placeholders into a
# real, historically-grounded simulated trade result.
# ---------------------------------------------------------------------------
def simulate_backtest_exit(kite_api, opt_token, target_date, entry_time_str, entry_ltp,
                            stop_ltp, target_ltp, risk_per_unit_premium, cache,
                            signal_reversal_time_str=None, max_hold_minutes=MAX_HOLD_MINUTES,
                            eod_squareoff_time=EOD_SQUAREOFF_TIME, interval='5minute',
                            quantity=0, max_loss_rs=MAX_LOSS_PER_TRADE_RS):
    """Returns a dict: exit_time, exit_ltp, exit_reason, max_ltp_seen,
    min_ltp_seen. Falls back to exit-at-entry with exit_reason
    'NO_HISTORICAL_DATA' (clearly flagged, never silently guessed) if the
    option's own candle history can't be fetched for this date -- e.g.
    Kite's F&O historical retention window doesn't reach that far back.

    [ADDED] quantity: pass the order's actual Quantity (Units) so the
    hard rupee cap below can be checked in real rupees, not premium
    points -- see MAX_LOSS_PER_TRADE_RS. Pass 0 (default) to skip this
    check entirely (e.g. a caller that only cares about premium-level
    exit simulation)."""
    df = historical_lookup.fetch_option_day_candles(kite_api, opt_token, target_date, cache, interval)
    entry_dt = historical_lookup.time_str_to_dt(target_date, entry_time_str)

    if df is None or df.empty:
        return {'exit_time': entry_time_str, 'exit_ltp': entry_ltp, 'exit_reason': 'NO_HISTORICAL_DATA',
                'max_ltp_seen': entry_ltp, 'min_ltp_seen': entry_ltp}

    # Only candles STRICTLY AFTER the entry bar -- the entry bar itself is
    # where the fill happened, not where the exit can first be checked.
    forward = df[df.index > entry_dt]

    max_ltp_seen = entry_ltp
    min_ltp_seen = entry_ltp
    trailing_stop = stop_ltp
    tsl_breach_streak = 0  # [ADDED -- ENABLE_TSL_CONFIRMATION_HOLD] consecutive confirmed-breach candles

    reversal_dt = (historical_lookup.time_str_to_dt(target_date, signal_reversal_time_str)
                   if signal_reversal_time_str else None)
    max_hold_dt = entry_dt + timedelta(minutes=max_hold_minutes)
    eod_dt = datetime.combine(target_date.date(), datetime.strptime(eod_squareoff_time, "%H:%M").time())

    for ts, candle_row in forward.iterrows():
        candle_close = float(candle_row.get('close', entry_ltp) or entry_ltp)
        candle_high = float(candle_row.get('high', candle_close) or candle_close)
        candle_low = float(candle_row.get('low', candle_close) or candle_close)

        max_ltp_seen = max(max_ltp_seen, candle_high)
        min_ltp_seen = min(min_ltp_seen, candle_low)

        # [ADDED] Hard rupee cap, checked FIRST and priced off the candle's
        # actual LOW -- the worst price this candle realistically touched
        # -- not the stop_ltp level the check below assumes a perfect fill
        # at. This is what actually "supersedes the ATR calculation": if a
        # single 5-min candle gaps straight through stop_ltp, the check
        # below would still report an exit at stop_ltp (optimistic); this
        # one reports the true rupee loss and exits at whatever price keeps
        # it at max_loss_rs, even if that's worse than stop_ltp.
        if quantity > 0:
            worst_case_loss = (entry_ltp - candle_low) * quantity
            if worst_case_loss > max_loss_rs:
                # [CHANGED] cost-aware cap -- see ENABLE_COST_AWARE_STOP_CAP
                # docstring above for why the old formula overshot its own
                # stated ceiling by 9-12% every time it fired.
                if ENABLE_COST_AWARE_STOP_CAP:
                    cap_exit_ltp = solve_cost_aware_cap_exit(entry_ltp, quantity, max_loss_rs)
                else:
                    cap_exit_ltp = max(entry_ltp - (max_loss_rs / quantity), 0.05)
                return {'exit_time': ts.strftime('%H:%M:%S'), 'exit_ltp': round(cap_exit_ltp, 2),
                        'exit_reason': f'Hard Stop-Loss (Rs {max_loss_rs:.0f} cap)',
                        'max_ltp_seen': round(max_ltp_seen, 2), 'min_ltp_seen': round(min_ltp_seen, 2)}

        # Ratchet the trailing stop up once price has run risk_per_unit_premium
        # x TRAIL_TRIGGER_R in profit, locking in a fraction of the running
        # gain above entry. Never moves the stop down.
        # [CHANGED -- Task 72, 22-Jul-26, ENABLE_MULTI_TARGET_TRAIL] The
        # lock fraction now TIGHTENS in stages as price clears Target 1/2/3
        # instead of staying fixed at TRAIL_LOCK_FRACTION forever -- see
        # _trailing_lock_fraction()'s docstring. Flag off -> byte-for-byte
        # the original behavior (stage is computed but ignored).
        gain = max_ltp_seen - entry_ltp
        stage = _target_stage_reached(gain, risk_per_unit_premium)
        if risk_per_unit_premium > 0 and gain >= TRAIL_TRIGGER_R * risk_per_unit_premium:
            lock_fraction = _trailing_lock_fraction(stage, ENABLE_MULTI_TARGET_TRAIL)
            trailing_stop = max(trailing_stop, entry_ltp + lock_fraction * gain)

        # Conservative same-candle ordering: if a single 5-min candle's
        # range spans BOTH the stop and the target, assume the stop hit
        # first -- OHLC alone can't tell true intra-candle sequencing,
        # and assuming the worse outcome is the standard conservative
        # backtest convention. [CHANGED -- ENABLE_TSL_CONFIRMATION_HOLD]
        # The original hard ATR stop (trailing_stop == stop_ltp, no profit
        # cushion yet) still exits INSTANTLY off candle_low exactly as
        # before -- only a TRAILING breach (already in profit) can be held
        # for confirmation, and only using candle CLOSE (a wick touching
        # the level doesn't count -- see flag docstring above).
        # [CHANGED -- Task 72] reason string notes the target stage already
        # cleared when the flag is on, so a trailing exit past Target 1/2
        # is distinguishable from an ordinary +1R trail in the sheet.
        stage_note = f" (past Target {stage})" if (ENABLE_MULTI_TARGET_TRAIL and stage > 0) else ""
        is_trailing = trailing_stop > stop_ltp
        if candle_low <= trailing_stop:
            if not (ENABLE_TSL_CONFIRMATION_HOLD and is_trailing):
                reason = f'Trailing Stop Hit{stage_note}' if is_trailing else 'Stop Loss Hit'
                return {'exit_time': ts.strftime('%H:%M:%S'), 'exit_ltp': round(trailing_stop, 2),
                        'exit_reason': reason, 'max_ltp_seen': round(max_ltp_seen, 2),
                        'min_ltp_seen': round(min_ltp_seen, 2)}
            if candle_close <= trailing_stop:
                tsl_breach_streak += 1
                if tsl_breach_streak >= TSL_CONFIRMATION_BARS:
                    return {'exit_time': ts.strftime('%H:%M:%S'), 'exit_ltp': round(candle_close, 2),
                            'exit_reason': f'Trailing Stop Hit (confirmed, {TSL_CONFIRMATION_BARS} bars){stage_note}',
                            'max_ltp_seen': round(max_ltp_seen, 2), 'min_ltp_seen': round(min_ltp_seen, 2)}
                # Held with reason: low wicked through but didn't close
                # beyond the trail confirmed yet -- fall through to the
                # other exit checks below instead of returning.
            else:
                tsl_breach_streak = 0
        else:
            tsl_breach_streak = 0
        # [CHANGED -- Task 72, 22-Jul-26, ENABLE_MULTI_TARGET_TRAIL] With
        # the flag ON, Target 1 (target_ltp) is no longer a hard exit --
        # the trailing stop above (now locking a bigger fraction once
        # Target 1/2/3 clear) is the ONLY thing that can close a winning
        # trade, per Harish's explicit choice: keep trailing indefinitely,
        # even past Target 3, rather than capping the position anywhere.
        # Flag off -> byte-for-byte the original hard-exit-at-target
        # behavior.
        if not ENABLE_MULTI_TARGET_TRAIL and candle_high >= target_ltp:
            return {'exit_time': ts.strftime('%H:%M:%S'), 'exit_ltp': round(target_ltp, 2),
                    'exit_reason': 'Target Hit', 'max_ltp_seen': round(max_ltp_seen, 2),
                    'min_ltp_seen': round(min_ltp_seen, 2)}
        if ENABLE_SIGNAL_REVERSAL_EXIT and reversal_dt is not None and ts >= reversal_dt:
            return {'exit_time': ts.strftime('%H:%M:%S'), 'exit_ltp': round(candle_close, 2),
                    'exit_reason': '5M Reversal (Signal WAIT)', 'max_ltp_seen': round(max_ltp_seen, 2),
                    'min_ltp_seen': round(min_ltp_seen, 2)}
        # [ADDED] see ENABLE_NO_FOLLOWTHROUGH_EXIT docstring above --
        # checked after every real price-based exit, before the time-cap
        # exits it's specifically designed to preempt.
        if ENABLE_NO_FOLLOWTHROUGH_EXIT and _no_followthrough_triggered(
                entry_dt, ts, max_ltp_seen, entry_ltp, risk_per_unit_premium):
            return {'exit_time': ts.strftime('%H:%M:%S'), 'exit_ltp': round(candle_close, 2),
                    'exit_reason': f'No Follow-Through ({NO_FOLLOWTHROUGH_MINUTES}min, <{NO_FOLLOWTHROUGH_R_MULTIPLE}R)',
                    'max_ltp_seen': round(max_ltp_seen, 2), 'min_ltp_seen': round(min_ltp_seen, 2)}
        if ts >= max_hold_dt:
            return {'exit_time': ts.strftime('%H:%M:%S'), 'exit_ltp': round(candle_close, 2),
                    'exit_reason': f'Max Hold Time ({max_hold_minutes}min)', 'max_ltp_seen': round(max_ltp_seen, 2),
                    'min_ltp_seen': round(min_ltp_seen, 2)}
        if ts >= eod_dt:
            return {'exit_time': ts.strftime('%H:%M:%S'), 'exit_ltp': round(candle_close, 2),
                    'exit_reason': 'EOD Square-off', 'max_ltp_seen': round(max_ltp_seen, 2),
                    'min_ltp_seen': round(min_ltp_seen, 2)}

    # Ran out of candles for the day without any exit condition firing --
    # square off at the last available close.
    if forward.empty:
        return {'exit_time': entry_time_str, 'exit_ltp': entry_ltp, 'exit_reason': 'EOD Square-off (no candles after entry)',
                'max_ltp_seen': entry_ltp, 'min_ltp_seen': entry_ltp}
    last_close = float(forward.iloc[-1]['close'])
    return {'exit_time': forward.index[-1].strftime('%H:%M:%S'), 'exit_ltp': round(last_close, 2),
            'exit_reason': 'EOD Square-off (data ended)', 'max_ltp_seen': round(max_ltp_seen, 2),
            'min_ltp_seen': round(min_ltp_seen, 2)}


# ---------------------------------------------------------------------------
# LIVE polling -- one call per candle-close cycle, per open position.
# Wire into 01_Master_Code.py's run_cycle() (see that file's patch
# notes). Lighter-touch than the BACKTEST path: one live quote() per
# still-open position per cycle, checked against the same stop/target/
# trailing/max-hold/EOD rules, using the real wall clock instead of a
# historical candle walk. Paper-trade thoroughly before trusting this
# with real order placement.
# ---------------------------------------------------------------------------
def check_live_exit(entry_ltp, stop_ltp, target_ltp, risk_per_unit_premium, current_ltp,
                     max_ltp_seen, entry_dt, now_dt, signal_reversal_now=False,
                     max_hold_minutes=MAX_HOLD_MINUTES, eod_squareoff_time=EOD_SQUAREOFF_TIME,
                     quantity=0, max_loss_rs=MAX_LOSS_PER_TRADE_RS, tsl_breach_streak=0):
    """Pure decision function (no I/O) so it can be unit-tested the same
    way simulate_backtest_exit() is. Returns (should_exit, exit_reason,
    exit_ltp, new_max_ltp_seen, new_trailing_stop, new_tsl_breach_streak)
    for ONE live poll.

    [ADDED] quantity: pass the order's Quantity (Units) to enable the hard
    rupee cap below -- pass 0 (default) to skip it. See
    MAX_LOSS_PER_TRADE_RS docstring.

    [ADDED] tsl_breach_streak: pass back in whatever new_tsl_breach_streak
    this function returned on the PREVIOUS poll for this same open
    position (order_sheet.py persists it in the 'TSL Breach Streak'
    column) -- LIVE has no candle history to look back over the way
    simulate_backtest_exit() does, so the confirmation count has to be
    carried across cycles by the caller instead. See
    ENABLE_TSL_CONFIRMATION_HOLD's docstring above."""
    new_max = max(max_ltp_seen, current_ltp)
    trailing_stop = stop_ltp
    # [CHANGED -- Task 72, 22-Jul-26, ENABLE_MULTI_TARGET_TRAIL] Same
    # staged lock-fraction tightening as simulate_backtest_exit() -- see
    # that function's comment and _trailing_lock_fraction()'s docstring.
    # `stage` is re-derived fresh every poll from new_max (which is
    # already persisted via the 'Max LTP' column), so no new column is
    # needed just to track it across LIVE cycles.
    gain = new_max - entry_ltp
    stage = _target_stage_reached(gain, risk_per_unit_premium)
    if risk_per_unit_premium > 0 and gain >= TRAIL_TRIGGER_R * risk_per_unit_premium:
        lock_fraction = _trailing_lock_fraction(stage, ENABLE_MULTI_TARGET_TRAIL)
        trailing_stop = max(trailing_stop, entry_ltp + lock_fraction * gain)

    eod_dt = now_dt.replace(
        hour=int(eod_squareoff_time.split(':')[0]), minute=int(eod_squareoff_time.split(':')[1]),
        second=0, microsecond=0,
    )

    # [ADDED] Checked first, same as simulate_backtest_exit() -- a live LTP
    # print can gap between polls just like a backtest candle can.
    if quantity > 0:
        # [CHANGED] cost-aware: in LIVE mode we exit at whatever current_ltp
        # is (no price to "solve for" like backtest), so instead the
        # TRIGGER fires a bit earlier -- once unrealized loss crosses
        # (max_loss_rs - estimated round-trip cost at this price) -- so
        # that by the time costs are actually deducted downstream, net
        # loss still lands at/under max_loss_rs. See ENABLE_COST_AWARE_STOP_CAP.
        if ENABLE_COST_AWARE_STOP_CAP:
            est_cost = estimate_round_trip_costs(entry_ltp, current_ltp, quantity)['total']
            effective_cap = max(max_loss_rs - est_cost, 0.0)
        else:
            effective_cap = max_loss_rs
        unrealized_loss = (entry_ltp - current_ltp) * quantity
        if unrealized_loss > effective_cap:
            return True, f'Hard Stop-Loss (Rs {max_loss_rs:.0f} cap)', current_ltp, new_max, trailing_stop, tsl_breach_streak

    # [CHANGED -- ENABLE_TSL_CONFIRMATION_HOLD, mirrors
    # simulate_backtest_exit()'s same change] Only a TRAILING breach
    # (already in profit) can be held for confirmation across polls; the
    # original hard stop_ltp still exits instantly.
    stage_note = f" (past Target {stage})" if (ENABLE_MULTI_TARGET_TRAIL and stage > 0) else ""
    is_trailing = trailing_stop > stop_ltp
    new_streak = tsl_breach_streak
    if current_ltp <= trailing_stop:
        if not (ENABLE_TSL_CONFIRMATION_HOLD and is_trailing):
            reason = f'Trailing Stop Hit{stage_note}' if is_trailing else 'Stop Loss Hit'
            return True, reason, trailing_stop, new_max, trailing_stop, new_streak
        new_streak = tsl_breach_streak + 1
        if new_streak >= TSL_CONFIRMATION_BARS:
            return (True, f'Trailing Stop Hit (confirmed, {TSL_CONFIRMATION_BARS} bars){stage_note}',
                    current_ltp, new_max, trailing_stop, new_streak)
        # Held with reason -- fall through to the other checks below
        # instead of exiting on this poll.
    else:
        new_streak = 0

    # [CHANGED -- Task 72] Same "no hard exit at Target 1 once the flag is
    # on" change as simulate_backtest_exit() -- see that function's comment.
    if not ENABLE_MULTI_TARGET_TRAIL and current_ltp >= target_ltp:
        return True, 'Target Hit', target_ltp, new_max, trailing_stop, new_streak
    if ENABLE_SIGNAL_REVERSAL_EXIT and signal_reversal_now:
        return True, '5M Reversal (Signal WAIT)', current_ltp, new_max, trailing_stop, new_streak
    if ENABLE_NO_FOLLOWTHROUGH_EXIT and _no_followthrough_triggered(
            entry_dt, now_dt, new_max, entry_ltp, risk_per_unit_premium):
        return (True, f'No Follow-Through ({NO_FOLLOWTHROUGH_MINUTES}min, <{NO_FOLLOWTHROUGH_R_MULTIPLE}R)',
                current_ltp, new_max, trailing_stop, new_streak)
    if now_dt >= entry_dt + timedelta(minutes=max_hold_minutes):
        return True, f'Max Hold Time ({max_hold_minutes}min)', current_ltp, new_max, trailing_stop, new_streak
    if now_dt >= eod_dt:
        return True, 'EOD Square-off', current_ltp, new_max, trailing_stop, new_streak
    return False, None, current_ltp, new_max, trailing_stop, new_streak


# ---------------------------------------------------------------------------
# Daily max-drawdown circuit breaker. [ADDED -- no equivalent existed
# anywhere in the pipeline before this patch; grep confirmed it across
# this file, order_sheet.py, and 01_Master_Code.py.]
# ---------------------------------------------------------------------------
class DailyDrawdownGuard:
    """Tracks realized Net P/L across a session and flips to breached once
    max_daily_loss_rs is crossed. Call .update(net_pl) every time a
    position closes; check .breached() before opening any new one.

    IMPORTANT -- mode-dependent behavior, read before wiring in:

    LIVE mode: order_sheet.build_order_sheet() is called once per
    candle-close cycle (see check_live_exit()'s docstring). Entries and
    exits both happen in real time across cycles, so a guard that's
    PERSISTED across cycles (cache_path set, mirroring PCRTrendTracker's
    fix) genuinely blocks new entries mid-session once the daily cap is
    breached. This is the mode this class is designed for.

    BACKTEST mode: [CHANGED -- Task 64, 20-Jul-26] Each order's real exit
    is now resolved IMMEDIATELY when it's created (order_sheet.py's
    _resolve_backtest_exit_now(), called from inside the per-symbol
    streak-detection loop), not in a separate pass after the whole day
    is scanned -- so Net P/L exists right away. This instance
    (build_order_sheet()'s own `dd_guard`) is still NOT updated live
    during that per-symbol scan, though, because the scan processes one
    symbol's whole day before moving to the next symbol -- checking
    .breached() there would see a LATER trade from an earlier-processed
    symbol before an EARLIER trade from a later-processed one, which
    isn't true wall-clock order. Instead, order_sheet.py's Phase 2.5
    runs a SEPARATE DailyDrawdownGuard instance through every entered
    order sorted by real entry time (true chronological order across
    every symbol) and PRUNES -- moves to Rejected -- any order entered
    after the true breach point. That's what actually enforces the cap
    in backtest now, equivalent to what LIVE's persisted instance does
    by construction (LIVE's entries genuinely happen in real time, one
    cycle at a time, so this class updating live there was already
    correct). See order_sheet.py's Phase 2.5 block for the enforcement
    itself.
    """

    def __init__(self, max_daily_loss_rs=DAILY_MAX_LOSS_RS, cache_path=None, date_key=None):
        self.max_daily_loss_rs = max_daily_loss_rs
        self.cache_path = cache_path
        self.date_key = date_key
        self.realized_pl = 0.0
        self._breached = False
        self._breach_time = None
        if cache_path and date_key:
            state = _load_json(cache_path, {}).get(date_key)
            if state:
                self.realized_pl = state.get('realized_pl', 0.0)
                self._breached = state.get('breached', False)
                self._breach_time = state.get('breach_time')

    def _persist(self):
        if self.cache_path and self.date_key:
            _save_json(self.cache_path, {self.date_key: self.status()})

    def update(self, net_pl_of_closed_trade, at_time_str=None):
        self.realized_pl += net_pl_of_closed_trade
        if not self._breached and self.realized_pl <= -abs(self.max_daily_loss_rs):
            self._breached = True
            self._breach_time = at_time_str or datetime.now().strftime('%H:%M:%S')
            print(f"[RISK] Daily drawdown cap breached at {self._breach_time}: "
                  f"realized P/L Rs {self.realized_pl:.2f} <= -Rs {self.max_daily_loss_rs:.2f}. "
                  f"No new entries for the rest of the session.")
        self._persist()

    def breached(self):
        return self._breached

    def status(self):
        return {'realized_pl': round(self.realized_pl, 2), 'max_daily_loss_rs': self.max_daily_loss_rs,
                'breached': self._breached, 'breach_time': self._breach_time}


def _load_json(path, default):
    import json
    try:
        with open(path, 'r') as f:
            return json.load(f)
    except Exception:
        return default


def _save_json(path, data):
    import json
    try:
        os.makedirs(os.path.dirname(path), exist_ok=True)
        with open(path, 'w') as f:
            json.dump(data, f)
    except Exception as e:
        print(f"[WARNING] Failed to persist {path}: {e}")


# ---------------------------------------------------------------------------
# Disclaimer: nothing in this module places real orders -- it only
# computes sizing/exit numbers for order_sheet.py to record. It is not
# financial advice, no result here is a guarantee of future performance,
# and every number should be paper-traded before risking real capital.
# ---------------------------------------------------------------------------
