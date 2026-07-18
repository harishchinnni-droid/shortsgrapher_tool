"""
smc_lib.py
----------
Vendored from joshyattridge/smart-money-concepts (PyPI: smartmoneyconcepts,
version 0.0.27), https://github.com/joshyattridge/smart-money-concepts,
MIT License. Copied in unmodified (only this header/attribution block
added) rather than added as a pip dependency, matching this codebase's
existing pattern of self-contained indicator modules (adx_di.py, RSI.py,
etc. are all local translations, not external package installs) -- one
less thing that can break across Harish's Colab/local environments.

Only smart_money_concepts.py (this project's own wrapper) calls into
this file -- swing_highs_lows() and ob() (Order Blocks) are the two
functions actually used for Task 51's first version. The rest
(fvg, bos_choch, liquidity, previous_high_low, sessions, retracements)
are left in place, unused for now, in case a future pass wants them --
see smart_money_concepts.py's own docstring for what's actually wired
into the gate today vs. what's available but dormant.

MIT License (from the upstream repository):
    Copyright (c) joshyattridge and contributors.
    Permission is hereby granted, free of charge, to any person obtaining
    a copy of this software and associated documentation files, to deal
    in the Software without restriction, including without limitation
    the rights to use, copy, modify, merge, publish, distribute,
    sublicense, and/or sell copies of the Software, subject to including
    the above copyright notice in all copies or substantial portions.
"""

from functools import wraps
import pandas as pd
import numpy as np
from pandas import DataFrame, Series
from datetime import datetime


def inputvalidator(input_="ohlc"):
    def dfcheck(func):
        @wraps(func)
        def wrap(*args, **kwargs):
            args = list(args)
            i = 0 if isinstance(args[0], pd.DataFrame) else 1

            args[i] = args[i].rename(columns={c: c.lower() for c in args[i].columns})

            inputs = {
                "o": "open",
                "h": "high",
                "l": "low",
                "c": kwargs.get("column", "close").lower(),
                "v": "volume",
            }

            if inputs["c"] != "close":
                kwargs["column"] = inputs["c"]

            for l in input_:
                if inputs[l] not in args[i].columns:
                    raise LookupError(
                        'Must have a dataframe column named "{0}"'.format(inputs[l])
                    )

            return func(*args, **kwargs)

        return wrap

    return dfcheck


def apply(decorator):
    def decorate(cls):
        for attr in cls.__dict__:
            if callable(getattr(cls, attr)):
                setattr(cls, attr, decorator(getattr(cls, attr)))

        return cls

    return decorate


@apply(inputvalidator(input_="ohlc"))
class smc:
    __version__ = "0.0.27"

    @classmethod
    def fvg(cls, ohlc: DataFrame, join_consecutive=False) -> Series:
        fvg = np.where(
            (
                (ohlc["high"].shift(1) < ohlc["low"].shift(-1))
                & (ohlc["close"] > ohlc["open"])
            )
            | (
                (ohlc["low"].shift(1) > ohlc["high"].shift(-1))
                & (ohlc["close"] < ohlc["open"])
            ),
            np.where(ohlc["close"] > ohlc["open"], 1, -1),
            np.nan,
        )

        top = np.where(
            ~np.isnan(fvg),
            np.where(
                ohlc["close"] > ohlc["open"],
                ohlc["low"].shift(-1),
                ohlc["low"].shift(1),
            ),
            np.nan,
        )

        bottom = np.where(
            ~np.isnan(fvg),
            np.where(
                ohlc["close"] > ohlc["open"],
                ohlc["high"].shift(1),
                ohlc["high"].shift(-1),
            ),
            np.nan,
        )

        if join_consecutive:
            for i in range(len(fvg) - 1):
                if fvg[i] == fvg[i + 1]:
                    top[i + 1] = max(top[i], top[i + 1])
                    bottom[i + 1] = min(bottom[i], bottom[i + 1])
                    fvg[i] = top[i] = bottom[i] = np.nan

        mitigated_index = np.zeros(len(ohlc), dtype=np.int32)
        for i in np.where(~np.isnan(fvg))[0]:
            mask = np.zeros(len(ohlc), dtype=np.bool_)
            if fvg[i] == 1:
                mask = ohlc["low"][i + 2:] <= top[i]
            elif fvg[i] == -1:
                mask = ohlc["high"][i + 2:] >= bottom[i]
            if np.any(mask):
                j = np.argmax(mask) + i + 2
                mitigated_index[i] = j

        mitigated_index = np.where(np.isnan(fvg), np.nan, mitigated_index)

        return pd.concat(
            [
                pd.Series(fvg, name="FVG"),
                pd.Series(top, name="Top"),
                pd.Series(bottom, name="Bottom"),
                pd.Series(mitigated_index, name="MitigatedIndex"),
            ],
            axis=1,
        )

    @classmethod
    def swing_highs_lows(cls, ohlc: DataFrame, swing_length: int = 50) -> Series:
        swing_length *= 2
        swing_highs_lows = np.where(
            ohlc["high"]
            == ohlc["high"].shift(-(swing_length // 2)).rolling(swing_length).max(),
            1,
            np.where(
                ohlc["low"]
                == ohlc["low"].shift(-(swing_length // 2)).rolling(swing_length).min(),
                -1,
                np.nan,
            ),
        )

        while True:
            positions = np.where(~np.isnan(swing_highs_lows))[0]

            if len(positions) < 2:
                break

            current = swing_highs_lows[positions[:-1]]
            next = swing_highs_lows[positions[1:]]

            highs = ohlc["high"].iloc[positions[:-1]].values
            lows = ohlc["low"].iloc[positions[:-1]].values

            next_highs = ohlc["high"].iloc[positions[1:]].values
            next_lows = ohlc["low"].iloc[positions[1:]].values

            index_to_remove = np.zeros(len(positions), dtype=bool)

            consecutive_highs = (current == 1) & (next == 1)
            index_to_remove[:-1] |= consecutive_highs & (highs < next_highs)
            index_to_remove[1:] |= consecutive_highs & (highs >= next_highs)

            consecutive_lows = (current == -1) & (next == -1)
            index_to_remove[:-1] |= consecutive_lows & (lows > next_lows)
            index_to_remove[1:] |= consecutive_lows & (lows <= next_lows)

            if not index_to_remove.any():
                break

            swing_highs_lows[positions[index_to_remove]] = np.nan

        positions = np.where(~np.isnan(swing_highs_lows))[0]

        if len(positions) > 0:
            if swing_highs_lows[positions[0]] == 1:
                swing_highs_lows[0] = -1
            if swing_highs_lows[positions[0]] == -1:
                swing_highs_lows[0] = 1
            if swing_highs_lows[positions[-1]] == -1:
                swing_highs_lows[-1] = 1
            if swing_highs_lows[positions[-1]] == 1:
                swing_highs_lows[-1] = -1

        level = np.where(
            ~np.isnan(swing_highs_lows),
            np.where(swing_highs_lows == 1, ohlc["high"], ohlc["low"]),
            np.nan,
        )

        return pd.concat(
            [
                pd.Series(swing_highs_lows, name="HighLow"),
                pd.Series(level, name="Level"),
            ],
            axis=1,
        )

    @classmethod
    def ob(
        cls,
        ohlc: DataFrame,
        swing_highs_lows: DataFrame,
        close_mitigation: bool = False,
    ) -> Series:
        ohlc_len = len(ohlc)
        _open = ohlc["open"].values
        _high = ohlc["high"].values
        _low = ohlc["low"].values
        _close = ohlc["close"].values
        _volume = ohlc["volume"].values
        swing_hl = swing_highs_lows["HighLow"].values

        crossed = np.full(ohlc_len, False, dtype=bool)
        ob = np.zeros(ohlc_len, dtype=np.int32)
        top_arr = np.zeros(ohlc_len, dtype=np.float32)
        bottom_arr = np.zeros(ohlc_len, dtype=np.float32)
        obVolume = np.zeros(ohlc_len, dtype=np.float32)
        lowVolume = np.zeros(ohlc_len, dtype=np.float32)
        highVolume = np.zeros(ohlc_len, dtype=np.float32)
        percentage = np.zeros(ohlc_len, dtype=np.float32)
        mitigated_index = np.zeros(ohlc_len, dtype=np.int32)
        breaker = np.full(ohlc_len, False, dtype=bool)

        swing_high_indices = np.flatnonzero(swing_hl == 1)
        swing_low_indices = np.flatnonzero(swing_hl == -1)

        active_bullish = []
        for i in range(ohlc_len):
            close_index = i
            for idx in active_bullish.copy():
                if breaker[idx]:
                    if _high[close_index] > top_arr[idx]:
                        ob[idx] = 0
                        top_arr[idx] = 0.0
                        bottom_arr[idx] = 0.0
                        obVolume[idx] = 0.0
                        lowVolume[idx] = 0.0
                        highVolume[idx] = 0.0
                        mitigated_index[idx] = 0
                        percentage[idx] = 0.0
                        active_bullish.remove(idx)
                else:
                    if ((not close_mitigation and _low[close_index] < bottom_arr[idx])
                            or (close_mitigation and min(_open[close_index], _close[close_index]) < bottom_arr[idx])):
                        breaker[idx] = True
                        mitigated_index[idx] = close_index - 1

            pos = np.searchsorted(swing_high_indices, close_index)
            last_top_index = swing_high_indices[pos - 1] if pos > 0 else None

            if last_top_index is not None:
                if _close[close_index] > _high[last_top_index] and not crossed[last_top_index]:
                    crossed[last_top_index] = True
                    default_index = close_index - 1
                    obBtm = _high[default_index]
                    obTop = _low[default_index]
                    obIndex = default_index
                    if close_index - last_top_index > 1:
                        start = last_top_index + 1
                        end = close_index
                        if end > start:
                            segment = _low[start:end]
                            min_val = segment.min()
                            candidates = np.nonzero(segment == min_val)[0]
                            if candidates.size:
                                candidate_index = start + candidates[-1]
                                obBtm = _low[candidate_index]
                                obTop = _high[candidate_index]
                                obIndex = candidate_index
                    ob[obIndex] = 1
                    top_arr[obIndex] = obTop
                    bottom_arr[obIndex] = obBtm
                    vol_cur = _volume[close_index]
                    vol_prev1 = _volume[close_index - 1] if close_index >= 1 else 0.0
                    vol_prev2 = _volume[close_index - 2] if close_index >= 2 else 0.0
                    obVolume[obIndex] = vol_cur + vol_prev1 + vol_prev2
                    lowVolume[obIndex] = vol_prev2
                    highVolume[obIndex] = vol_cur + vol_prev1
                    max_vol = max(highVolume[obIndex], lowVolume[obIndex])
                    percentage[obIndex] = (min(highVolume[obIndex], lowVolume[obIndex]) / max_vol * 100.0) if max_vol != 0 else 100.0
                    active_bullish.append(obIndex)

        active_bearish = []
        for i in range(ohlc_len):
            close_index = i
            for idx in active_bearish.copy():
                if breaker[idx]:
                    if _low[close_index] < bottom_arr[idx]:
                        ob[idx] = 0
                        top_arr[idx] = 0.0
                        bottom_arr[idx] = 0.0
                        obVolume[idx] = 0.0
                        lowVolume[idx] = 0.0
                        highVolume[idx] = 0.0
                        mitigated_index[idx] = 0
                        percentage[idx] = 0.0
                        active_bearish.remove(idx)
                else:
                    if ((not close_mitigation and _high[close_index] > top_arr[idx])
                            or (close_mitigation and max(_open[close_index], _close[close_index]) > top_arr[idx])):
                        breaker[idx] = True
                        mitigated_index[idx] = close_index

            pos = np.searchsorted(swing_low_indices, close_index)
            last_btm_index = swing_low_indices[pos - 1] if pos > 0 else None

            if last_btm_index is not None:
                if _close[close_index] < _low[last_btm_index] and not crossed[last_btm_index]:
                    crossed[last_btm_index] = True
                    default_index = close_index - 1
                    obTop = _high[default_index]
                    obBtm = _low[default_index]
                    obIndex = default_index
                    if close_index - last_btm_index > 1:
                        start = last_btm_index + 1
                        end = close_index
                        if end > start:
                            segment = _high[start:end]
                            max_val = segment.max()
                            candidates = np.nonzero(segment == max_val)[0]
                            if candidates.size:
                                candidate_index = start + candidates[-1]
                                obTop = _high[candidate_index]
                                obBtm = _low[candidate_index]
                                obIndex = candidate_index
                    ob[obIndex] = -1
                    top_arr[obIndex] = obTop
                    bottom_arr[obIndex] = obBtm
                    vol_cur = _volume[close_index]
                    vol_prev1 = _volume[close_index - 1] if close_index >= 1 else 0.0
                    vol_prev2 = _volume[close_index - 2] if close_index >= 2 else 0.0
                    obVolume[obIndex] = vol_cur + vol_prev1 + vol_prev2
                    lowVolume[obIndex] = vol_cur + vol_prev1
                    highVolume[obIndex] = vol_prev2
                    max_vol = max(highVolume[obIndex], lowVolume[obIndex])
                    percentage[obIndex] = (min(highVolume[obIndex], lowVolume[obIndex]) / max_vol * 100.0) if max_vol != 0 else 100.0
                    active_bearish.append(obIndex)

        ob = np.where(ob != 0, ob, np.nan)
        top_arr = np.where(~np.isnan(ob), top_arr, np.nan)
        bottom_arr = np.where(~np.isnan(ob), bottom_arr, np.nan)
        obVolume = np.where(~np.isnan(ob), obVolume, np.nan)
        mitigated_index = np.where(~np.isnan(ob), mitigated_index, np.nan)
        percentage = np.where(~np.isnan(ob), percentage, np.nan)

        ob_series = pd.Series(ob, name="OB")
        top_series = pd.Series(top_arr, name="Top")
        bottom_series = pd.Series(bottom_arr, name="Bottom")
        obVolume_series = pd.Series(obVolume, name="OBVolume")
        mitigated_index_series = pd.Series(mitigated_index, name="MitigatedIndex")
        percentage_series = pd.Series(percentage, name="Percentage")

        return pd.concat(
            [
                ob_series,
                top_series,
                bottom_series,
                obVolume_series,
                mitigated_index_series,
                percentage_series,
            ],
            axis=1,
        )
