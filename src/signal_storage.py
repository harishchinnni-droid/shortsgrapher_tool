"""
Signal Storage Module
Handles storing trading signals in SQLite database for audit trail and reporting.
"""

import sqlite3
import os
from datetime import datetime
from typing import List, Dict, Any
from src.logging_config import LoggingManager

logger = LoggingManager.get_logger()


class SignalStorage:
    """Manages signal storage in SQLite database."""
    
    def __init__(self, storage_dir: str = "./local_trading_data"):
        """
        Initialize signal storage.
        
        Args:
            storage_dir: Directory to store signals database
        """
        self.storage_dir = storage_dir
        self.db_path = os.path.join(storage_dir, "signals.db")
        self._initialize_database()
        logger.info(f"Signal Storage initialized at: {self.db_path}")
    
    def _initialize_database(self):
        """Create signals table if it doesn't exist."""
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            # Create signals table
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS signals (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    signal_date TEXT NOT NULL,
                    symbol TEXT NOT NULL,
                    signal_type TEXT NOT NULL,
                    signal_strength REAL,
                    rsi_value REAL,
                    vwap_value REAL,
                    adx_value REAL,
                    squeeze_momentum REAL,
                    ema_value REAL,
                    oi_dynamics REAL,
                    breakout_score REAL,
                    entry_price REAL,
                    stop_loss REAL,
                    take_profit REAL,
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    UNIQUE(signal_date, symbol, signal_type)
                )
            """)
            
            # Create index for faster queries
            cursor.execute("""
                CREATE INDEX IF NOT EXISTS idx_signal_date_symbol 
                ON signals(signal_date, symbol)
            """)
            
            conn.commit()
            conn.close()
            logger.info("✓ Signals database table created/verified")
        except Exception as e:
            logger.error(f"Failed to initialize database: {e}")
            raise
    
    def insert_signal(self, signal_data: Dict[str, Any]) -> bool:
        """
        Insert a trading signal into the database.
        
        Args:
            signal_data: Dictionary containing signal information
            
        Returns:
            bool: Success status
        """
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            cursor.execute("""
                INSERT OR REPLACE INTO signals (
                    signal_date, symbol, signal_type, signal_strength,
                    rsi_value, vwap_value, adx_value, squeeze_momentum,
                    ema_value, oi_dynamics, breakout_score,
                    entry_price, stop_loss, take_profit
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                signal_data.get("signal_date"),
                signal_data.get("symbol"),
                signal_data.get("signal_type"),
                signal_data.get("signal_strength"),
                signal_data.get("rsi_value"),
                signal_data.get("vwap_value"),
                signal_data.get("adx_value"),
                signal_data.get("squeeze_momentum"),
                signal_data.get("ema_value"),
                signal_data.get("oi_dynamics"),
                signal_data.get("breakout_score"),
                signal_data.get("entry_price"),
                signal_data.get("stop_loss"),
                signal_data.get("take_profit")
            ))
            
            conn.commit()
            conn.close()
            logger.info(f"✓ Signal stored: {signal_data.get('symbol')} - {signal_data.get('signal_type')}")
            return True
        except Exception as e:
            logger.error(f"Failed to insert signal: {e}")
            return False
    
    def get_signals_by_date(self, date_str: str) -> List[Dict[str, Any]]:
        """
        Retrieve all signals for a specific date.
        
        Args:
            date_str: Date in YYYY-MM-DD format
            
        Returns:
            List of signal dictionaries
        """
        try:
            conn = sqlite3.connect(self.db_path)
            conn.row_factory = sqlite3.Row
            cursor = conn.cursor()
            
            cursor.execute("""
                SELECT * FROM signals 
                WHERE signal_date = ? 
                ORDER BY symbol ASC
            """, (date_str,))
            
            rows = cursor.fetchall()
            conn.close()
            
            signals = [dict(row) for row in rows]
            logger.info(f"✓ Retrieved {len(signals)} signals for date: {date_str}")
            return signals
        except Exception as e:
            logger.error(f"Failed to retrieve signals: {e}")
            return []
    
    def get_all_signals(self) -> List[Dict[str, Any]]:
        """
        Retrieve all signals from database.
        
        Returns:
            List of all signal dictionaries
        """
        try:
            conn = sqlite3.connect(self.db_path)
            conn.row_factory = sqlite3.Row
            cursor = conn.cursor()
            
            cursor.execute("SELECT * FROM signals ORDER BY signal_date DESC, symbol ASC")
            
            rows = cursor.fetchall()
            conn.close()
            
            signals = [dict(row) for row in rows]
            logger.info(f"✓ Retrieved {len(signals)} total signals from database")
            return signals
        except Exception as e:
            logger.error(f"Failed to retrieve all signals: {e}")
            return []
    
    def get_database_path(self) -> str:
        """Get the full path to the signals database."""
        return self.db_path
    
    def database_exists(self) -> bool:
        """Check if database file exists."""
        return os.path.exists(self.db_path)
