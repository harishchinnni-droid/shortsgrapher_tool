"""
Main Orchestrator for Stock Automation Tool
Coordinates all modules in sequential order.

Execution Flow:
1. Broker Authentication (Angel One & Zerodha)
2. Data Ingestion (Initial data fetch)
3. Incremental Data Sync (Delta loads)
4. Technical Indicator Calculation (Market analysis)
5. Market Status Check (Live vs Backtest routing)
"""

import os
from datetime import date

from src.logging_config import LoggingManager
from src.broker_login import BrokerAuthentication
from src.data_ingestion import DataIngestion
from src.incremental_sync import IncrementalSync
from src.indicator_engine import IndicatorEngine

# Initialize logger
logger = LoggingManager.get_logger()

# Configuration
JSON_DIR = r"F:\Stock_Automation_11-Apr-26\01 JSON Files"
ZERODHA_FILE = os.path.join(JSON_DIR, "harish_zerodha.JSON")
ANGEL_FILE = os.path.join(JSON_DIR, "harish_angel_one.JSON")
CACHE_FILE = os.path.join(JSON_DIR, "session_cache.json")

SERVER_NAME = 'DESKTOP-57PQCS1'
DB_OPERATIONAL = 'FnO_Apr26'
DB_HISTORICAL = 'Historical_Database'


class StockAutomationOrchestrator:
    """Main orchestrator for stock automation workflow."""
    
    def __init__(self):
        """Initialize orchestrator with all required modules."""
        self.logger = logger
        self.zerodha_client = None
        self.angel_client = None
    
    def step_1_authenticate_brokers(self):
        """Step 1: Authenticate with brokers."""
        self.logger.info("\n" + "="*80)
        self.logger.info("STEP 1: BROKER AUTHENTICATION")
        self.logger.info("="*80)
        
        try:
            auth = BrokerAuthentication(JSON_DIR, CACHE_FILE)
            
            # Load credentials
            logger.info("Loading broker credentials...")
            zerodha_creds = auth.load_credentials(ZERODHA_FILE)
            angel_creds = auth.load_credentials(ANGEL_FILE)
            
            # Authenticate Angel One
            self.angel_client = auth.login_angel_one(angel_creds)
            
            # Authenticate Zerodha
            self.zerodha_client = auth.login_zerodha(zerodha_creds)
            
            if self.angel_client and self.zerodha_client:
                logger.info("✓ STEP COMPLETED: Both brokers authenticated successfully")
                return True
            else:
                logger.error("✗ STEP FAILED: One or more brokers failed authentication")
                return False
        
        except Exception as e:
            logger.error(f"✗ STEP FAILED: {e}", exc_info=True)
            return False
    
    def step_2_initial_data_ingestion(self):
        """Step 2: Fetch and store initial data."""
        self.logger.info("\n" + "="*80)
        self.logger.info("STEP 2: INITIAL DATA INGESTION")
        self.logger.info("="*80)
        
        if not self.zerodha_client:
            logger.error("✗ STEP SKIPPED: Zerodha client not available")
            return False
        
        try:
            ingestion = DataIngestion(JSON_DIR, CACHE_FILE, SERVER_NAME, DB_OPERATIONAL, DB_HISTORICAL)
            
            # Get active symbols
            active_symbols = ingestion.get_active_symbols()
            if active_symbols.empty:
                logger.warning("⚠ No active symbols found")
                return False
            
            # Process and store data
            logger.info(f"Processing {len(active_symbols)} symbols...")
            success_count, fail_count = ingestion.process_and_store_data(self.zerodha_client, active_symbols)
            
            logger.info(f"✓ STEP COMPLETED: Ingestion done (Success: {success_count}, Failed: {fail_count})")
            return True
        
        except Exception as e:
            logger.error(f"✗ STEP FAILED: {e}", exc_info=True)
            return False
    
    def step_3_incremental_data_sync(self):
        """Step 3: Execute incremental data synchronization."""
        self.logger.info("\n" + "="*80)
        self.logger.info("STEP 3: INCREMENTAL DATA SYNC")
        self.logger.info("="*80)
        
        if not self.zerodha_client:
            logger.error("✗ STEP SKIPPED: Zerodha client not available")
            return False
        
        try:
            sync = IncrementalSync(SERVER_NAME, DB_OPERATIONAL, DB_HISTORICAL)
            
            # Get active symbols
            from sqlalchemy import text
            query = "SELECT Symbol, KiteToken FROM dbo.Reference WHERE KiteToken IS NOT NULL"
            from sqlalchemy import create_engine
            import urllib
            
            conn_str = f'DRIVER={{ODBC Driver 17 for SQL Server}};SERVER={SERVER_NAME};DATABASE={DB_OPERATIONAL};Trusted_Connection=yes;'
            engine_op = create_engine(f"mssql+pyodbc:///?odbc_connect={urllib.parse.quote_plus(conn_str)}")
            
            import pandas as pd
            with engine_op.connect() as conn:
                active_symbols = pd.read_sql(text(query), conn)
            
            if active_symbols.empty:
                logger.warning("⚠ No active symbols found")
                return False
            
            # Execute incremental sync
            logger.info(f"Running incremental sync for {len(active_symbols)} symbols...")
            success, skipped, failed = sync.fetch_incremental_data(self.zerodha_client, active_symbols)
            
            logger.info(f"✓ STEP COMPLETED: Sync done (Success: {success}, Skipped: {skipped}, Failed: {failed})")
            return True
        
        except Exception as e:
            logger.error(f"✗ STEP FAILED: {e}", exc_info=True)
            return False
    
    def step_4_calculate_indicators(self):
        """Step 4: Calculate technical indicators."""
        self.logger.info("\n" + "="*80)
        self.logger.info("STEP 4: TECHNICAL INDICATOR CALCULATION")
        self.logger.info("="*80)
        
        try:
            indicators = IndicatorEngine(SERVER_NAME, DB_OPERATIONAL, DB_HISTORICAL)
            
            # Check market status
            sync = IncrementalSync(SERVER_NAME, DB_OPERATIONAL, DB_HISTORICAL)
            market_open, reason = sync.is_market_open()
            
            if market_open:
                target_date = date.today().strftime("%Y-%m-%d")
                logger.info(f"Market open. Running indicators for today: {target_date}")
            else:
                target_date = date.today().strftime("%Y-%m-%d")
                logger.info(f"Market closed ({reason}). Running indicators for today: {target_date}")
            
            # Run indicators
            processed = indicators.run_indicators_for_date(target_date)
            logger.info(f"✓ STEP COMPLETED: Indicators calculated for {processed} symbols")
            return True
        
        except Exception as e:
            logger.error(f"✗ STEP FAILED: {e}", exc_info=True)
            return False
    
    def step_5_system_status_check(self):
        """Step 5: Final system status check."""
        self.logger.info("\n" + "="*80)
        self.logger.info("STEP 5: SYSTEM STATUS CHECK")
        self.logger.info("="*80)
        
        try:
            sync = IncrementalSync(SERVER_NAME, DB_OPERATIONAL, DB_HISTORICAL)
            market_open, reason = sync.is_market_open()
            
            if market_open:
                logger.info("✓ Market is OPEN - System in LIVE mode")
            else:
                logger.info(f"⚠ Market is CLOSED ({reason}) - System in BACKTEST mode")
            
            logger.info("✓ STEP COMPLETED: System status verified")
            return True
        
        except Exception as e:
            logger.error(f"✗ STEP FAILED: {e}", exc_info=True)
            return False
    
    def run(self):
        """Execute the complete orchestration workflow."""
        logger.info("\n")
        logger.info("╔" + "="*78 + "╗")
        logger.info("║" + " "*20 + "STOCK AUTOMATION TOOL - MASTER SEQUENCE" + " "*20 + "║")
        logger.info("╚" + "="*78 + "╝")
        
        results = {
            "Step 1 - Authentication": self.step_1_authenticate_brokers(),
            "Step 2 - Data Ingestion": self.step_2_initial_data_ingestion(),
            "Step 3 - Incremental Sync": self.step_3_incremental_data_sync(),
            "Step 4 - Indicator Calculation": self.step_4_calculate_indicators(),
            "Step 5 - System Status": self.step_5_system_status_check(),
        }
        
        # Summary
        logger.info("\n" + "="*80)
        logger.info("EXECUTION SUMMARY")
        logger.info("="*80)
        
        for step, result in results.items():
            status = "✓ PASSED" if result else "✗ FAILED"
            logger.info(f"{step}: {status}")
        
        all_passed = all(results.values())
        
        if all_passed:
            logger.info("\n✓ ALL STEPS COMPLETED SUCCESSFULLY")
            logger.info("System is online and operational")
        else:
            logger.warning("\n⚠ SOME STEPS FAILED - Review logs above")
        
        logger.info("="*80)
        logger.info("EXECUTION FINISHED")
        logger.info("="*80 + "\n")
        
        return all_passed


def main():
    """Entry point for the application."""
    orchestrator = StockAutomationOrchestrator()
    success = orchestrator.run()
    return success


if __name__ == "__main__":
    success = main()
    exit(0 if success else 1)
