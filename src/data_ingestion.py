"""
Data Ingestion Module
Fetches historical data from Kite API and stores in SQL Server Database.
"""

import pandas as pd
import urllib
import os
import time
import json
from sqlalchemy import create_engine, text
from datetime import datetime, timedelta

from src.logging_config import LoggingManager

logger = LoggingManager.get_logger()


class DataIngestion:
    """Handles data fetching and storage operations."""
    
    def __init__(self, json_dir, cache_file, server_name, db_operational, db_historical):
        """
        Initialize data ingestion.
        
        Args:
            json_dir: Directory containing credentials
            cache_file: Path to session cache
            server_name: SQL Server name
            db_operational: Operational database name
            db_historical: Historical database name
        """
        self.json_dir = json_dir
        self.cache_file = cache_file
        self.server_name = server_name
        self.db_operational = db_operational
        self.db_historical = db_historical
        self.logger = logger
        
        # Initialize engines
        self.engine_operational = self._create_engine(db_operational)
        self.engine_historical = self._create_engine(db_historical)
    
    def _create_engine(self, db_name):
        """Create SQLAlchemy engine for database."""
        conn_str = f'DRIVER={{ODBC Driver 17 for SQL Server}};SERVER={self.server_name};DATABASE={db_name};Trusted_Connection=yes;'
        return create_engine(f"mssql+pyodbc:///?odbc_connect={urllib.parse.quote_plus(conn_str)}", fast_executemany=True)
    
    def get_active_symbols(self):
        """
        Fetch active symbols from operational database.
        
        Returns:
            DataFrame with Symbol and KiteToken columns
        """
        self.logger.info(f"Fetching reference tokens from {self.db_operational}...")
        
        try:
            query = "SELECT Symbol, KiteToken FROM dbo.Reference WHERE KiteToken IS NOT NULL"
            with self.engine_operational.connect() as conn:
                df = pd.read_sql(text(query), conn)
                self.logger.info(f"Retrieved {len(df)} active symbols")
                return df
        except Exception as e:
            self.logger.error(f"Error fetching symbols: {e}", exc_info=True)
            return pd.DataFrame()
    
    def process_and_store_data(self, kite, ref_df):
        """
        Fetch data from Kite API and store in database.
        
        Args:
            kite: KiteConnect API instance
            ref_df: DataFrame with symbol references
        """
        self.logger.info(f"Starting data fetch and storage to {self.db_historical}")
        
        now = datetime.now()
        minute_remainder = now.minute % 5
        last_closed_candle_time = now - timedelta(minutes=minute_remainder, seconds=now.second, microseconds=now.microsecond)
        
        from_date = (now - timedelta(days=5)).strftime("%Y-%m-%d 09:15:00")
        to_date = last_closed_candle_time.strftime("%Y-%m-%d %H:%M:%S")
        
        self.logger.info(f"Data range: {from_date} to {to_date}")
        
        successful_symbols = 0
        failed_symbols = 0
        
        for index, row in ref_df.iterrows():
            symbol = row['Symbol']
            token = int(row['KiteToken'])
            
            try:
                self.logger.info(f"Processing: {symbol}...")
                
                records = kite.historical_data(token, from_date, to_date, "5minute")
                if not records:
                    self.logger.warning(f"No data received from Kite for {symbol}")
                    continue
                
                df_5m = pd.DataFrame(records)
                
                # Strip timezone from timestamp
                df_5m['date'] = pd.to_datetime(df_5m['date']).dt.tz_localize(None)
                
                df_5m['Symbol'] = symbol
                df_5m['Timeframe'] = '5m'
                df_5m = df_5m[['Symbol', 'Timeframe', 'date', 'open', 'high', 'low', 'close', 'volume']]
                df_5m.columns = ['Symbol', 'Timeframe', 'Date', 'Open', 'High', 'Low', 'Close', 'Volume']
                
                # Store 5m data
                try:
                    df_5m.to_sql('Historical_Candles', self.engine_historical, schema='dbo', if_exists='append', index=False)
                    self.logger.info(f"✓ Inserted {len(df_5m)} 5m candles for {symbol}")
                except Exception as sql_err:
                    self.logger.error(f"✗ SQL INSERT ERROR on 5m {symbol}: {sql_err}")
                    continue
                
                # Resample to higher timeframes
                df_5m.set_index('Date', inplace=True)
                agg = {'Open': 'first', 'High': 'max', 'Low': 'min', 'Close': 'last', 'Volume': 'sum'}
                
                for tf in ['15min', '1h', '1D']:
                    try:
                        resampled = df_5m.resample(tf).agg(agg).dropna().reset_index()
                        resampled['Symbol'] = symbol
                        resampled['Timeframe'] = tf
                        resampled.to_sql('Historical_Candles', self.engine_historical, schema='dbo', if_exists='append', index=False)
                        self.logger.info(f"✓ Inserted {len(resampled)} {tf} candles for {symbol}")
                    except Exception as resample_err:
                        self.logger.error(f"✗ SQL INSERT ERROR on {tf} {symbol}: {resample_err}")
                
                successful_symbols += 1
                time.sleep(0.4)  # Rate limiting
                
            except Exception as e:
                self.logger.error(f"✗ Error processing {symbol}: {e}", exc_info=True)
                failed_symbols += 1
        
        self.logger.info(f"Data ingestion summary - Success: {successful_symbols}, Failed: {failed_symbols}")
        return successful_symbols, failed_symbols
