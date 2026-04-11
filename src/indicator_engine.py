"""
Technical Indicator Engine Module
Calculates complex technical indicators for stock analysis.
Includes ADX, RSI, Squeeze Momentum, VWAP, and OI Dynamics.
"""

import pandas as pd
import numpy as np
import urllib
from sqlalchemy import create_engine, text
from datetime import datetime, timedelta, date
import warnings

from src.logging_config import LoggingManager

logger = LoggingManager.get_logger()
warnings.simplefilter(action='ignore', category=pd.errors.PerformanceWarning)

# Indicator Constants
ADX_TREND_THRESHOLD = 20
CANDLE_BODY_RATIO_MIN = 0.40
MIN_5M_VOLUME = 25000
MAX_DEAD_CANDLES = 3
COMPOSITE_THRESHOLD = 50


class IndicatorEngine:
    """Calculates technical indicators for market analysis."""
    
    def __init__(self, server_name, db_operational, db_historical):
        """
        Initialize indicator engine.
        
        Args:
            server_name: SQL Server name
            db_operational: Operational database name
            db_historical: Historical database name
        """
        self.server_name = server_name
        self.db_operational = db_operational
        self.db_historical = db_historical
        self.logger = logger
        
        # Initialize engines
        self.engine_op = self._create_engine(db_operational)
        self.engine_hist = self._create_engine(db_historical)
    
    def _create_engine(self, db_name):
        """Create SQLAlchemy engine for database."""
        conn_str = f'DRIVER={{ODBC Driver 17 for SQL Server}};SERVER={self.server_name};DATABASE={db_name};Trusted_Connection=yes;'
        return create_engine(f"mssql+pyodbc:///?odbc_connect={urllib.parse.quote_plus(conn_str)}", fast_executemany=True)
    
    @staticmethod
    def wilder_rma(series, period):
        """
        Calculate Wilder's Modified Moving Average (Relative Moving Average).
        
        Args:
            series: Pandas series
            period: Period for calculation
            
        Returns:
            Series with RMA values
        """
        vals = series.values.astype(float)
        result = np.full(len(vals), np.nan)
        valid_mask = ~np.isnan(vals)
        
        if valid_mask.sum() < period:
            return pd.Series(result, index=series.index)
        
        first_valid = int(np.argmax(valid_mask))
        seed_end = first_valid + period
        
        if seed_end > len(vals):
            return pd.Series(result, index=series.index)
        
        result[seed_end - 1] = np.nanmean(vals[first_valid:seed_end])
        
        for i in range(seed_end, len(vals)):
            if not np.isnan(vals[i]) and not np.isnan(result[i-1]):
                result[i] = (result[i-1] * (period - 1) + vals[i]) / period
        
        return pd.Series(result, index=series.index)
    
    @staticmethod
    def calc_adx(df, period=14):
        """
        Calculate Average Directional Index (ADX).
        
        Args:
            df: DataFrame with OHLC data
            period: Period for ADX calculation
            
        Returns:
            DataFrame with ADX indicators
        """
        high, low, close = df['High'], df['Low'], df['Close']
        tr = pd.concat([high - low, (high - close.shift(1)).abs(), (low - close.shift(1)).abs()], axis=1).max(axis=1)
        
        up_move = high.diff()
        dn_move = -(low.diff())
        
        plus_dm = pd.Series(np.where((up_move > dn_move) & (up_move > 0), up_move, 0.0), index=df.index)
        minus_dm = pd.Series(np.where((dn_move > up_move) & (dn_move > 0), dn_move, 0.0), index=df.index)
        
        atr_s = IndicatorEngine.wilder_rma(tr, period)
        plus_s = IndicatorEngine.wilder_rma(plus_dm, period)
        minus_s = IndicatorEngine.wilder_rma(minus_dm, period)
        
        df['plus_di'] = (plus_s / atr_s.replace(0, np.nan)) * 100
        df['minus_di'] = (minus_s / atr_s.replace(0, np.nan)) * 100
        
        dx = (abs(df['plus_di'] - df['minus_di']) / (df['plus_di'] + df['minus_di']).replace(0, np.nan)) * 100
        df['adx'] = IndicatorEngine.wilder_rma(dx.fillna(0), period)
        
        return df
    
    @staticmethod
    def calc_rsi(df, period=9):
        """
        Calculate Relative Strength Index (RSI).
        
        Args:
            df: DataFrame with OHLC data
            period: RSI period
            
        Returns:
            DataFrame with RSI values
        """
        delta = df['Close'].diff()
        gain = delta.where(delta > 0, 0.0)
        loss = (-delta.where(delta < 0, 0.0))
        
        df['rsi'] = 100 - (100 / (1 + IndicatorEngine.wilder_rma(gain, period) / IndicatorEngine.wilder_rma(loss, period).replace(0, np.nan)))
        
        return df
    
    @staticmethod
    def get_rolling_linreg(series, window=20):
        """
        Calculate rolling linear regression.
        
        Args:
            series: Series for regression
            window: Rolling window size
            
        Returns:
            Series with regression values
        """
        vals = series.values
        out = np.full(len(vals), np.nan)
        x = np.arange(window)
        x_mean = x.mean()
        x_var = ((x - x_mean) ** 2).sum()
        
        for i in range(window - 1, len(vals)):
            y = vals[i - window + 1: i + 1]
            y_mean = y.mean()
            cov = ((x - x_mean) * (y - y_mean)).sum()
            slope = cov / x_var
            intercept = y_mean - slope * x_mean
            out[i] = intercept + slope * (window - 1)
        
        return pd.Series(out, index=series.index)
    
    @staticmethod
    def calc_sqzmom(df):
        """
        Calculate Squeeze Momentum indicator.
        
        Args:
            df: DataFrame with OHLC data
            
        Returns:
            DataFrame with squeeze momentum values
        """
        df['basis'] = df['Close'].rolling(20).mean()
        df['dev'] = 2.0 * df['Close'].rolling(20).std(ddof=0)
        df['upperBB'] = df['basis'] + df['dev']
        df['lowerBB'] = df['basis'] - df['dev']
        
        tr0 = abs(df['High'] - df['Low'])
        tr1 = abs(df['High'] - df['Close'].shift(1))
        tr2 = abs(df['Low'] - df['Close'].shift(1))
        df['tr'] = pd.concat([tr0, tr1, tr2], axis=1).max(axis=1)
        
        df['ma'] = df['Close'].rolling(20).mean()
        df['rangema'] = df['tr'].rolling(20).mean()
        df['upperKC'] = df['ma'] + df['rangema'] * 1.5
        df['lowerKC'] = df['ma'] - df['rangema'] * 1.5
        
        df['sqzOn'] = (df['lowerBB'] > df['lowerKC']) & (df['upperBB'] < df['upperKC'])
        df['sqzOff'] = (df['lowerBB'] < df['lowerKC']) & (df['upperBB'] > df['upperKC'])
        
        highest_high = df['High'].rolling(20).max()
        lowest_low = df['Low'].rolling(20).min()
        avg_hl = (highest_high + lowest_low) / 2
        
        df['mom_source'] = df['Close'] - (avg_hl + df['ma']) / 2
        df['sqz_val'] = IndicatorEngine.get_rolling_linreg(df['mom_source'], 20)
        
        return df
    
    @staticmethod
    def calc_vwap(df):
        """
        Calculate Volume Weighted Average Price (VWAP).
        
        Args:
            df: DataFrame with OHLCV data
            
        Returns:
            DataFrame with VWAP values
        """
        df = df.copy()
        df['tp'] = (df['High'] + df['Low'] + df['Close']) / 3
        df['tp_vol'] = df['tp'] * df['Volume']
        df['date_only'] = pd.to_datetime(df['Date']).dt.date
        df['cum_tp_vol'] = df.groupby('date_only')['tp_vol'].cumsum()
        df['cum_vol'] = df.groupby('date_only')['Volume'].cumsum()
        df['vwap'] = df['cum_tp_vol'] / df['cum_vol'].replace(0, np.nan)
        df['_vwap_dev'] = df['Close'] - df['vwap']
        df['vwap_std'] = df.groupby('date_only')['_vwap_dev'].transform(lambda x: x.expanding().std())
        
        df.drop(columns=['_vwap_dev', 'tp', 'tp_vol', 'cum_tp_vol', 'cum_vol', 'date_only'], inplace=True)
        
        return df
    
    @staticmethod
    def calc_oi_dynamics(df):
        """
        Calculate Open Interest dynamics.
        
        Args:
            df: DataFrame with OHLCV and OI data
            
        Returns:
            DataFrame with OI indicators
        """
        has_oi = bool(('OI' in df.columns) and (df['OI'] > 0).any())
        
        if not has_oi:
            df['OI_Recomm'] = "N/A"
            df['cum_oi_bull'] = 0
            df['cum_oi_bear'] = 0
            df['oi_acceleration'] = 0
            return df
        
        df['oi_change'] = df['OI'].diff()
        df['price_change'] = df['Close'].diff()
        df['oi_change_smooth'] = df['oi_change'].rolling(3).mean()
        df['price_change_smooth'] = df['price_change'].rolling(3).mean()
        df['oi_acceleration'] = df['oi_change_smooth'].diff()
        
        oi_rising = df['oi_change_smooth'] > 0
        oi_falling = df['oi_change_smooth'] < 0
        price_up = df['price_change_smooth'] > 0
        price_down = df['price_change_smooth'] < 0
        oi_accel_pos = df['oi_acceleration'] >= 0
        
        cond_lb = price_up & oi_rising & oi_accel_pos
        cond_sb = price_down & oi_rising & oi_accel_pos
        cond_luw = price_up & oi_falling
        cond_suw = price_down & oi_falling
        
        df['OI_Recomm'] = np.select([cond_lb, cond_sb, cond_luw, cond_suw], ["Long Buildup", "Short Buildup", "Long Unwinding", "Short Unwinding"], default="WAIT")
        df['cum_oi_bull'] = df['oi_change_smooth'].where(price_up, 0).rolling(5).sum()
        df['cum_oi_bear'] = df['oi_change_smooth'].where(price_down, 0).rolling(5).sum()
        
        return df
    
    @staticmethod
    def calc_breakout_probability(df, lookback=10):
        """
        Calculate breakout probability score.
        
        Args:
            df: DataFrame with OHLCV data
            lookback: Lookback period
            
        Returns:
            DataFrame with breakout scores
        """
        df = IndicatorEngine.calc_adx(df, period=14)
        df = IndicatorEngine.calc_rsi(df, period=9)
        df = IndicatorEngine.calc_sqzmom(df)
        
        df['vol_ma'] = df['Volume'].rolling(20).mean()
        df['vol_std_r'] = df['Volume'].rolling(20).std().replace(0, np.nan)
        df['vol_zscore'] = ((df['Volume'] - df['vol_ma']) / df['vol_std_r']).clip(-3, 3)
        vol_norm = ((df['vol_zscore'] + 3) / 6.0).clip(0, 1)
        
        adx = df['adx']
        adx_slope = adx - adx.shift(3)
        adx_above = adx > ADX_TREND_THRESHOLD
        adx_rising = adx_slope > 0
        
        df['adx_score'] = np.where(adx_above & adx_rising, np.minimum((adx - ADX_TREND_THRESHOLD) / 25.0, 1.0), 0.0)
        
        df['candle_range'] = df['High'] - df['Low']
        df['candle_body'] = abs(df['Close'] - df['Open'])
        body_ratio = (df['candle_body'] / df['candle_range'].replace(0, np.nan)).fillna(0).clip(0, 1)
        df['strong_body'] = body_ratio >= CANDLE_BODY_RATIO_MIN
        
        top_of_body = df[['Close', 'Open']].max(axis=1)
        bot_of_body = df[['Close', 'Open']].min(axis=1)
        upper_wick_r = ((df['High'] - top_of_body) / df['candle_range'].replace(0, np.nan)).fillna(0).clip(0, 1)
        lower_wick_r = ((bot_of_body - df['Low']) / df['candle_range'].replace(0, np.nan)).fillna(0).clip(0, 1)
        
        bull_cq = (body_ratio * (1 - upper_wick_r)).clip(0, 1)
        bear_cq = (body_ratio * (1 - lower_wick_r)).clip(0, 1)
        
        df['roc_3'] = df['Close'].pct_change(3) * 100
        roc_bull_score = pd.Series(np.where(df['roc_3'] > 0, np.minimum(df['roc_3'] / 1.5, 1.0), 0.0), index=df.index)
        roc_bear_score = pd.Series(np.where(df['roc_3'] < 0, np.minimum(abs(df['roc_3']) / 1.5, 1.0), 0.0), index=df.index)
        
        rsi = df['rsi'].fillna(50)
        rsi_bull_score = pd.Series(np.where((rsi >= 52) & (rsi <= 75), (rsi - 52) / 23.0, np.where(rsi > 75, 0.2, 0.0)), index=df.index)
        rsi_bear_score = pd.Series(np.where((rsi >= 25) & (rsi <= 48), (48 - rsi) / 23.0, np.where(rsi < 25, 0.2, 0.0)), index=df.index)
        
        sqz_multiplier = np.where(df['sqzOn'], 0.5, 1.0)
        adx_s_series = pd.Series(df['adx_score'].values, index=df.index)
        
        df['composite_bull'] = (adx_s_series * 30 + vol_norm * 20 + bull_cq * 25 + roc_bull_score * 15 + rsi_bull_score * 10) * sqz_multiplier
        df['composite_bear'] = (adx_s_series * 30 + vol_norm * 20 + bear_cq * 25 + roc_bear_score * 15 + rsi_bear_score * 10) * sqz_multiplier
        
        df['prob_bull'] = np.where(df['Close'] > df['Open'], 1, 0)
        df['prob_bear'] = np.where(df['Close'] < df['Open'], 1, 0)
        df['prob_bull'] = df['prob_bull'].rolling(lookback).mean() * 100
        df['prob_bear'] = df['prob_bear'].rolling(lookback).mean() * 100
        
        df['is_dead'] = df['candle_range'] <= (df['Close'] * 0.0005)
        df['dead_count'] = df['is_dead'].rolling(lookback).sum()
        
        bull_filter = ((df['composite_bull'] >= COMPOSITE_THRESHOLD) & (df['Close'] > df['Open']) & df['strong_body'] & (df['Volume'] >= MIN_5M_VOLUME) & (df['dead_count'] <= MAX_DEAD_CANDLES) & (df['adx_score'] > 0) & ~df['sqzOn'])
        bear_filter = ((df['composite_bear'] >= COMPOSITE_THRESHOLD) & (df['Close'] < df['Open']) & df['strong_body'] & (df['Volume'] >= MIN_5M_VOLUME) & (df['dead_count'] <= MAX_DEAD_CANDLES) & (df['adx_score'] > 0) & ~df['sqzOn'])
        
        df['Breakout_Score'] = np.select([bull_filter, bear_filter], ["High Prob Bull", "High Prob Bear"], default="Low Confidence")
        
        return df
    
    def run_indicators_for_date(self, target_date_str):
        """
        Run all indicators for a specific date.
        
        Args:
            target_date_str: Target date in YYYY-MM-DD format
        """
        self.logger.info(f"\n[INDICATOR ENGINE] Initializing for Target Date: {target_date_str}")
        
        try:
            target_date = datetime.strptime(target_date_str, "%Y-%m-%d").date()
            
            # Get all unique symbols
            symbols_query = text("SELECT DISTINCT Symbol FROM dbo.Historical_Candles WHERE Timeframe = '5m'")
            with self.engine_hist.connect() as conn:
                symbols_df = pd.read_sql(symbols_query, conn)
            
            if symbols_df.empty:
                self.logger.error("[FATAL] No symbols found in Data Warehouse.")
                return 0
            
            lookback_start = (target_date - timedelta(days=45)).strftime("%Y-%m-%d")
            target_end_str = (target_date + timedelta(days=1)).strftime("%Y-%m-%d")
            
            processed_count = 0
            
            for symbol in symbols_df['Symbol']:
                query = text("""
                    SELECT Date, Open, High, Low, Close, Volume, OI
                    FROM dbo.Historical_Candles 
                    WHERE Symbol = :sym AND Timeframe = '5m' 
                    AND Date >= :start AND Date < :end
                    ORDER BY Date ASC
                """)
                
                with self.engine_hist.connect() as conn:
                    df = pd.read_sql(query, conn, params={"sym": symbol, "start": lookback_start, "end": target_end_str})
                
                if df.empty or len(df) < 50:
                    continue
                
                # Calculate all indicators
                df = self.calc_oi_dynamics(df)
                df = self.calc_breakout_probability(df)
                df = self.calc_vwap(df)
                
                # Filter data for target date
                df['date_only'] = pd.to_datetime(df['Date']).dt.date
                output_df = df[df['date_only'] == target_date].copy()
                output_df.drop(columns=['date_only'], inplace=True)
                
                if output_df.empty:
                    continue
                
                output_df['Symbol'] = symbol
                output_df['Timeframe'] = '5m'
                output_df.replace([np.inf, -np.inf, np.nan], None, inplace=True)
                
                # Store to database
                try:
                    output_df.to_sql('Technical_Indicators', self.engine_hist, schema='dbo', if_exists='append', index=False)
                    self.logger.info(f"✓ [{symbol}] Wrote {len(output_df)} indicator rows")
                    processed_count += 1
                except Exception as e:
                    self.logger.error(f"✗ [SQL] Write failed for {symbol}: {e}")
            
            self.logger.info(f"Indicator calculation complete - Processed: {processed_count} symbols")
            return processed_count
            
        except Exception as e:
            self.logger.error(f"Error in indicator engine: {e}", exc_info=True)
            return 0
