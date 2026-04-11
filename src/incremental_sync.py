"""
Incremental Data Synchronization Module
Handles delta loads and market status checks.
"""

import pandas as pd
import urllib
import time
from sqlalchemy import create_engine, text
from datetime import datetime, timedelta, date

from src.logging_config import LoggingManager

logger = LoggingManager.get_logger()

# NSE Trading Holidays for 2026
NSE_HOLIDAYS = [
    date(2026, 1, 15), date(2026, 1, 26), date(2026, 3, 3), date(2026, 3, 26),
    date(2026, 3, 31), date(2026, 4, 3), date(2026, 4, 14), date(2026, 5, 1),
    date(2026, 5, 28), date(2026, 6, 26), date(2026, 9, 14), date(2026, 10, 2),
    date(2026, 10, 20), date(2026, 11, 10), date(2026, 11, 24), date(2026, 12, 25),
]


class IncrementalSync:
    """Handles incremental data synchronization."""
    
    def __init__(self, server_name, db_operational, db_historical):
        """
        Initialize incremental sync.
        
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
    
    def get_last_synced_date(self, symbol, timeframe):
        """
        Get the last synced date for a symbol.
        
        Args:
            symbol: Trading symbol
            timeframe: Candle timeframe
            
        Returns:
            datetime of last sync or None
        """
        try:
            query = text("SELECT MAX(Date) FROM dbo.Historical_Candles WHERE Symbol = :s AND Timeframe = :t")
            with self.engine_hist.connect() as conn:
                result = conn.execute(query, {"s": symbol, "t": timeframe}).scalar()
                return result
        except Exception as e:
            self.logger.error(f"Error getting last sync date for {symbol}: {e}")
            return None
    
    def is_market_open(self):
        """
        Check if NSE market is open today.
        
        Returns:
            Tuple (bool, str) - (is_open, reason)
        """
        today = date.today()
        
        if today.weekday() >= 5:
            return False, "Weekend"
        
        if today in NSE_HOLIDAYS:
            return False, "Holiday"
        
        return True, "Live"
    
    def fetch_incremental_data(self, kite, ref_df):
        """
        Execute incremental data sync with delta loading.
        
        Args:
            kite: KiteConnect API instance
            ref_df: DataFrame with symbol references
        """
        now = datetime.now()
        
        # Enforce strictly closed 5-minute candles
        target_end = now - timedelta(minutes=now.minute % 5, seconds=now.second, microseconds=now.microsecond)
        
        self.logger.info(f"Starting Incremental Sync. Target End Time: {target_end}")
        
        successful_symbols = 0
        skipped_symbols = 0
        failed_symbols = 0
        
        for _, row in ref_df.iterrows():
            symbol = row['Symbol']
            token = int(row['KiteToken'])
            
            try:
                last_date = self.get_last_synced_date(symbol, '5m')
                
                if last_date:
                    from_date = last_date
                    self.logger.info(f"[{symbol}] Base timestamp: {last_date}. Executing delta sync.")
                else:
                    from_date = now - timedelta(days=90)
                    self.logger.info(f"[{symbol}] No baseline found. Initiating 90-day bulk download.")
                
                if from_date >= target_end:
                    self.logger.info(f"[{symbol}] Sync redundant. Data is already up to date.")
                    skipped_symbols += 1
                    continue
                
                # Fetch data from Kite
                records = kite.historical_data(token, from_date, target_end, "5minute")
                if not records:
                    self.logger.warning(f"[{symbol}] No data received from Kite")
                    continue
                
                df = pd.DataFrame(records)
                df['date'] = pd.to_datetime(df['date']).dt.tz_localize(None)
                df['Symbol'] = symbol
                df['Timeframe'] = '5m'
                
                df = df[['Symbol', 'Timeframe', 'date', 'open', 'high', 'low', 'close', 'volume']]
                df.columns = ['Symbol', 'Timeframe', 'Date', 'Open', 'High', 'Low', 'Close', 'Volume']
                
                # Remove overlapping candles
                if last_date:
                    df = df[df['Date'] > last_date]
                
                if df.empty:
                    self.logger.info(f"[{symbol}] No new closed candles available.")
                    continue
                
                # Store to database
                df.to_sql('Historical_Candles', self.engine_hist, schema='dbo', if_exists='append', index=False)
                self.logger.info(f"[{symbol}] ✓ Inserted {len(df)} records to {self.db_historical}")
                
                successful_symbols += 1
                time.sleep(0.4)  # Rate limiting
                
            except Exception as e:
                self.logger.error(f"[{symbol}] ✗ Sync failed: {e}", exc_info=True)
                failed_symbols += 1
        
        self.logger.info(f"Incremental sync summary - Success: {successful_symbols}, Skipped: {skipped_symbols}, Failed: {failed_symbols}")
        return successful_symbols, skipped_symbols, failed_symbols
