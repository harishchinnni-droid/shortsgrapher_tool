"""
ist_clock.py
------------
Single source of truth for "what time is it right now" across the whole
pipeline.

Every scheduling / staleness decision this codebase makes -- when the 9:15
candle closes, whether a cached file is stale, what "today" is, whether a
candle is legitimately in the past -- has to agree with the NSE's clock,
which runs on IST (UTC+5:30, no DST), full stop. The old code called the
bare `datetime.now()` everywhere, which returns the HOST MACHINE's local
time. That's only correct if the machine running this script happens to be
configured for Asia/Kolkata. On a cloud VM defaulting to UTC, a laptop set
to a different timezone, or a system clock that's simply wrong, every
"now" in the pipeline silently disagrees with the real Indian market clock
-- which is how candles that hadn't actually happened yet in IST were able
to end up in the workbook.

now_ist() fixes this by anchoring to the IST timezone explicitly via the
stdlib `zoneinfo`, independent of the host's own clock/timezone setting.
It returns a NAIVE datetime (tzinfo stripped) because the rest of this
codebase -- and the Kite Connect API itself, whose historical_data()
expects local-IST-naive datetimes -- work in naive datetimes throughout;
mixing naive and aware datetimes raises TypeError the first time one is
compared against the other.

Every module in this pipeline that previously called `datetime.now()` for
scheduling/business-time purposes (calendar_mgmt, token_mgmt, broker_auth,
data_ingestion, 01_Master_Code / run_pipeline) should import and use
now_ist() from here instead. (pyotp's internal `.now()` for TOTP codes is
unrelated -- that's UTC-epoch-based by design and must NOT be touched.)

Windows note: zoneinfo needs the IANA tz database. Most Linux/macOS
systems already have it; on Windows, install the `tzdata` package once:
    pip install tzdata

[ADDED -- market-hours constants] NSE cash/F&O session: 09:15-15:30 IST.
The first 5-minute candle is 09:15-09:20, so it's the first one that can
possibly exist; nothing should be fetched, cycled, or displayed as
"today's data" before it closes. Centralizing these here (rather than as
magic numbers scattered across 01_Master_Code.py / run_pipeline.py) is
what let the scheduling fix in 01_Master_Code.py's _next_candle_close()
and _wait_for_market_open() be a two-line change instead of a guess at
what "9:15" meant in three different places.
"""
from datetime import datetime, time as dtime
from zoneinfo import ZoneInfo

IST = ZoneInfo("Asia/Kolkata")

# NSE equity/F&O session times, IST. Update here (only) if the exchange
# ever changes session hours -- every scheduling decision in the pipeline
# should derive from these, not a hardcoded "9:15" typed somewhere else.
MARKET_OPEN_TIME = dtime(9, 15, 0)
MARKET_CLOSE_TIME = dtime(15, 30, 0)
# First candle is 09:15-09:20; it cannot close, and therefore cannot be
# fetched or displayed as real data, before 09:20:00.
FIRST_CANDLE_CLOSE_TIME = dtime(9, 20, 0)


def now_ist():
    """Current wall-clock time in India (IST, UTC+5:30) as a naive
    datetime -- correct regardless of the host machine's own timezone."""
    return datetime.now(IST).replace(tzinfo=None)


def today_ist():
    """Today's IST calendar date, midnight-naive. Use for date-only
    comparisons (holiday checks, cache-file dating) where only the
    calendar day matters, not the time of day."""
    return now_ist().replace(hour=0, minute=0, second=0, microsecond=0)


def is_before_market_open(dt=None):
    """True if `dt` (defaults to now_ist()) is earlier than 09:15 IST on
    its own calendar day. Used to gate the live loop so no cycle -- and no
    same-day data fetch -- is attempted before the exchange has actually
    opened."""
    dt = dt or now_ist()
    return dt.time() < MARKET_OPEN_TIME


def seconds_until_market_open(dt=None):
    """Seconds remaining until 09:15 IST today, from `dt` (defaults to
    now_ist()). Returns 0 if the market is already open. Does not account
    for weekends/holidays -- pair with calendar_mgmt.is_trading_holiday()
    for that, same as get_execution_date() already does."""
    dt = dt or now_ist()
    market_open_dt = dt.replace(hour=9, minute=15, second=0, microsecond=0)
    return max(0.0, (market_open_dt - dt).total_seconds())
