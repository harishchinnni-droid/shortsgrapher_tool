import os
import json
import pandas as pd
import requests
from datetime import datetime
from ist_clock import now_ist
import file_mgmt

# Same folder broker_auth.py already uses for the daily Zerodha access-token
# cache -- keeping the token-mapping cache alongside it.
# [CHANGED -- cloud/Colab portability] derives from file_mgmt.BASE_DIR --
# see file_mgmt.py's BASE_DIR docstring.
JSON_DIR = os.path.join(file_mgmt.BASE_DIR, "01_JSON_Files")
os.makedirs(JSON_DIR, exist_ok=True)


def _cache_path(target_date):
    return os.path.join(JSON_DIR, f"instrument_token_cache_{target_date.strftime('%Y-%m-%d')}.json")


def _fetch_broker_masters(kite_api):
    """The expensive step: full Zerodha (NSE+NFO) + Angel One instrument
    master download. Token<->symbol mappings don't change intraday, so this
    should only ever run once per trading day -- update_instrument_tokens()
    below skips it entirely on a cache hit.
    """
    print("[SYSTEM] Fetching Broker Master Registries (full download)...")
    kite_nse = pd.DataFrame(kite_api.instruments(exchange=kite_api.EXCHANGE_NSE))
    kite_nfo = pd.DataFrame(kite_api.instruments(exchange=kite_api.EXCHANGE_NFO))
    kite_master = pd.concat([kite_nse, kite_nfo], ignore_index=True)

    url = "https://margincalculator.angelbroking.com/OpenAPI_File/files/OpenAPIScripMaster.json"
    angel_master = pd.DataFrame(requests.get(url, timeout=45).json())
    angel_master['symbol'] = angel_master['symbol'].astype(str).str.strip().str.upper()

    # Build O(1) lookup dicts ONCE instead of the old per-symbol boolean-mask
    # filter (kite_master[kite_master['tradingsymbol'] == symbol]), which was
    # an O(rows_ref x rows_master) full scan repeated for every single symbol.
    zerodha_lookup = (
        kite_master.drop_duplicates('tradingsymbol')
        .set_index('tradingsymbol')['instrument_token']
        .to_dict()
    )
    angel_lookup = {}
    for _, r in angel_master.iterrows():
        angel_lookup.setdefault(r['symbol'], str(r['token']))

    return zerodha_lookup, angel_lookup


def update_instrument_tokens(filepath, kite_api, target_date=None, force_refresh=False):
    """Maps Zerodha/Angel tokens onto the Reference sheet's symbol list.

    Reuses a same-day cache (JSON_DIR/instrument_token_cache_<date>.json)
    instead of re-downloading the full NSE+NFO+Angel master lists and
    re-matching every symbol on every run -- tokens are static for the whole
    trading day, so that work only needs to happen once. The cache only
    stores the (small) matched result for symbols actually in your Reference
    sheet, not the full master lists, so it stays fast to read/write.

    Self-healing: if any symbol in Reference isn't in today's cache yet
    (e.g. you added a new symbol mid-day), this automatically re-runs the
    full match for that day rather than silently leaving it blank. Pass
    force_refresh=True to force a full re-fetch regardless.
    """
    # [FIX] was datetime.now() -- host local time, not necessarily IST.
    target_date = target_date or now_ist()
    cache_file = _cache_path(target_date)

    df_ref = pd.read_excel(filepath, sheet_name='Reference')

    if 'Zerodha_Token' not in df_ref.columns:
        df_ref['Zerodha_Token'] = pd.Series(dtype='object')
    else:
        df_ref['Zerodha_Token'] = df_ref['Zerodha_Token'].astype('object')

    if 'Angel_Token' not in df_ref.columns:
        df_ref['Angel_Token'] = pd.Series(dtype='object')
    else:
        df_ref['Angel_Token'] = df_ref['Angel_Token'].astype('object')

    symbols = [str(s).strip().upper() for s in df_ref['Symbol / StrikePrice']]

    cached = {}
    if not force_refresh and os.path.exists(cache_file):
        with open(cache_file, 'r') as f:
            cached = json.load(f)

    missing = [s for s in symbols if s not in cached]

    if force_refresh or missing:
        if missing and not force_refresh:
            print(f"[SYSTEM] {len(missing)} symbol(s) not in today's token cache -- refreshing.")
        zerodha_lookup, angel_lookup = _fetch_broker_masters(kite_api)

        print("[SYSTEM] Re-mapping instrument tokens (dict lookup)...")
        for sym in symbols:
            entry = {}
            if sym in zerodha_lookup:
                entry['zerodha_token'] = str(zerodha_lookup[sym])
            angel_key = f"{sym}-EQ" if f"{sym}-EQ" in angel_lookup else sym
            if angel_key in angel_lookup:
                entry['angel_token'] = angel_lookup[angel_key]
            if entry:
                cached[sym] = entry

        with open(cache_file, 'w') as f:
            json.dump(cached, f)
    else:
        print(f"[SYSTEM] Reusing cached instrument tokens for {target_date.strftime('%d-%b-%y')} "
              f"({len(symbols)} symbols -- master download skipped).")

    for idx, sym in enumerate(symbols):
        entry = cached.get(sym, {})
        if 'zerodha_token' in entry:
            df_ref.at[idx, 'Zerodha_Token'] = entry['zerodha_token']
        if 'angel_token' in entry:
            df_ref.at[idx, 'Angel_Token'] = entry['angel_token']

    with pd.ExcelWriter(filepath, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        df_ref.to_excel(writer, sheet_name='Reference', index=False)

    print("[SYSTEM] Token mapping updated.")
    return df_ref