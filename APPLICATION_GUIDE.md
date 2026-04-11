"""
STOCK AUTOMATION TOOL - APPLICATION GUIDE

This document outlines the architecture, structure, and modules of the Stock Automation Tool
built from the Jupyter notebook F&O_Login.ipynb
"""

# ============================================================================
# PROJECT STRUCTURE
# ============================================================================

src/
├── __init__.py                 # Package initialization
├── logging_config.py           # Logging configuration (Timestamp-based logs)
├── broker_login.py             # Broker authentication module
├── data_ingestion.py           # Initial data fetching and storage
├── incremental_sync.py         # Incremental data syncing with NSE holidays
├── indicator_engine.py         # Technical indicator calculations
└── main.py                     # Main orchestrator (5 sequential steps)

main.py                         # Application entry point
requirements.txt                # Python dependencies
config/
└── config.yml                  # Application configuration

logs/
└── <YYYY-MM-DD_HH-MM-SS>.log  # Timestamped log files (auto-created)


# ============================================================================
# MODULE BREAKDOWN (Mapped from Jupyter Notebook Cells)
# ============================================================================

NOTEBOOK CELL 1: pip install dependencies
→ MAPPED TO: requirements.txt
   Contains: kiteconnect, smartapi-python, pyotp, selenium, sqlalchemy, pandas, etc.

NOTEBOOK CELL 2: Broker Authentication (Angel One & Zerodha)
→ MAPPED TO: src/broker_login.py
   Class: BrokerAuthentication
   Methods:
   - load_credentials(): Load JSON credential files
   - login_angel_one(): Authenticate with Angel One broker
   - login_zerodha(): Authenticate with Zerodha using Selenium automation
   - load_cache() / update_cache(): Session caching

NOTEBOOK CELL 3: Data Ingestion (Initial 5-day fetch)
→ MAPPED TO: src/data_ingestion.py
   Class: DataIngestion
   Methods:
   - get_active_symbols(): Query Reference table for active symbols
   - process_and_store_data(): Fetch from Kite API and store in SQL Server
   Features: Automatic timeframe resampling (5m → 15m, 1h, 1D)

NOTEBOOK CELL 4: Incremental Data Sync  
→ MAPPED TO: src/incremental_sync.py
   Class: IncrementalSync
   Methods:
   - fetch_incremental_data(): Delta load with overlap detection
   - get_last_synced_date(): Query last sync timestamp
   - is_market_open(): Check NSE trading calendar
   Features: Smart delta loading, NSE holiday calendar built-in

NOTEBOOK CELL 5: Technical Indicator Engine
→ MAPPED TO: src/indicator_engine.py
   Class: IndicatorEngine
   Methods:
   - calc_adx(): Average Directional Index
   - calc_rsi(): Relative Strength Index
   - calc_sqzmom(): Squeeze Momentum
   - calc_vwap(): Volume Weighted Average Price
   - calc_oi_dynamics(): Open Interest analysis
   - calc_breakout_probability(): Composite breakout scoring
   - run_indicators_for_date(): Execute all indicators for a date
   Features: Multi-timeframe analysis, OI-based signal generation


# ============================================================================
# EXECUTION FLOW (Sequential, as per Notebook)
# ============================================================================

MASTER ORCHESTRATOR: src/main.py → StockAutomationOrchestrator

Step 1: BROKER AUTHENTICATION
   - Load credentials from JSON files
   - Authenticate with Angel One (if available)
   - Authenticate with Zerodha (with Selenium automation)
   - Cache tokens for session reuse

Step 2: DATA INGESTION  
   - Query operational database for active symbols
   - Fetch 5-day historical data from Kite API
   - Store base 5m candles to database
   - Resample to 15m, 1h, 1D automatically

Step 3: INCREMENTAL DATA SYNC
   - Compare last synced timestamp with current time
   - Execute delta load (only new/closed candles)
   - Remove duplicate overlapping candles
   - Update database with latest data

Step 4: INDICATOR CALCULATION
   - Check if market is open (NSE trading calendar)
   - Retrieve 45-day lookback data for each symbol
   - Calculate all technical indicators
   - Score breakout probability for each symbol
   - Write results to Technical_Indicators table

Step 5: SYSTEM STATUS CHECK
   - Verify market status (Live vs Closed)
   - Display mode information (Live/Backtest)
   - Final system readiness check

FINAL SUMMARY:
   - Display pass/fail status for each step
   - Show overall system status


# ============================================================================
# LOGGING SYSTEM
# ============================================================================

LOGGING CONFIGURATION: src/logging_config.py
LoggingManager Class:
   - setup_logger(): Initialize logger with timestamp
   - Log file format: <YYYY-MM-DD_HH-MM-SS>.log
   - Logs stored in: logs/ directory
   - Both console and file output
   - Format: "YYYY-MM-DD HH:MM:SS - [LEVEL] - function:line - message"

Methods:
   - log_step_start(step_name): Log step initiation
   - log_step_completion(step_name, status): Log step completion
   - log_error(step_name, error_message, exception): Log errors with traceback

USAGE IN MODULES:
   from src.logging_config import LoggingManager
   logger = LoggingManager.get_logger()
   logger.info("Message here")


# ============================================================================
# CONFIGURATION & CREDENTIALS
# ============================================================================

CREDENTIALS STRUCTURE (JSON Files):
F:\Stock_Automation_11-Apr-26\01 JSON Files\

harish_zerodha.JSON:
{
  "api_key": "xxx",
  "api_secret": "xxx",
  "user_id": "xxx",
  "password": "xxx"
}

harish_angel_one.JSON:
{
  "api_key": "xxx",
  "client_id": "xxx",
  "user_id": "xxx",
  "password": "xxx",
  "totp_secret": "xxx"
}

DATABASES:
   Server: DESKTOP-57PQCS1 (Windows Authentication)
   Operational DB: FnO_Apr26 (Read-only for reference data)
   Historical DB: Historical_Database (Write-heavy for candle data)

TABLES:
   Reference: Symbol, KiteToken mapping
   Historical_Candles: Date, Open, High, Low, Close, Volume, OI
   Technical_Indicators: Date, RSI, ADX, VWAP, Breakout_Score, etc.


# ============================================================================
# HOW TO RUN
# ============================================================================

PREREQUISITE:
1. Ensure all credentials JSON files exist in correct location
2. Verify SQL Server databases are accessible
3. Install Python 3.8+ and pip

INSTALLATION:
1. cd F:\01_Stokc_Automation_V2\Stock_Automation
2. python -m venv venv
3. venv\Scripts\activate  (Windows)
4. pip install -r requirements.txt

EXECUTION:
1. python main.py

OUTPUT:
- Console logs with step-by-step progress
- Timestamped log file created in logs/ directory
- Database tables populated with results


# ============================================================================
# FEATURES
# ============================================================================

✓ Broker Agnostic: Supports both Angel One and Zerodha simultaneously
✓ Automated Selenium: Handles Zerodha TOTP input automatically
✓ Session Caching: Reuses tokens within same day (avoid re-authentication)
✓ NSE Calendar: Built-in trading holidays for India (2026)
✓ Delta Loading: Smart incremental sync (no duplicate candles)
✓ Multi-timeframe: Auto-resample 5m to 15m, 1h, 1D
✓ Comprehensive Indicators: ADX, RSI, Squeeze Momentum, VWAP, OI Dynamics
✓ Breakout Scoring: Multi-factor composite scoring system
✓ Logging: Timestamp-based file logging with detailed traceability
✓ Error Handling: Try-catch blocks throughout with detailed error messages
✓ No CLI Input: All parameters hardcoded (no user prompts)


# ============================================================================
# TROUBLESHOOTING
# ============================================================================

ISSUE: "CRITICAL ERROR: Credentials file not found"
   Solution: Verify JSON files exist in F:\Stock_Automation_11-Apr-26\01 JSON Files\

ISSUE: "TokenException" on Zerodha login
   Solution: Cached token expired. Delete session_cache.json and re-run.

ISSUE: "No mapped tokens found in Operational Reference table"
   Solution: Verify Reference table in FnO_Apr26 database has KiteToken values.

ISSUE: Browser automation hangs on TOTP input
   Solution: Verify Zerodha login credentials are correct. Check browser visibility.

ISSUE: SQL Server connection fails
   Solution: Verify Windows Authentication is enabled. Check server name and database names.

ISSUE: No indicators written to database
   Solution: Verify Historical_Candles table has sufficient data (>50 rows per symbol).


# ============================================================================
#PERFORMANCE NOTES
# ============================================================================

Data Ingestion: ~0.4s per symbol (rate-limited)
Incremental Sync: ~0.4s per symbol (depends on new candles)
Indicator Calculation: Depends on lookback period and data volume
   - 45-day lookback with ADX, RSI, VWAP, OI: ~1-2s per symbol
Total Runtime: Typically 5-15 minutes for complete run with ~50-100 symbols


# ============================================================================
# FUTURE ENHANCEMENTS
# ============================================================================

[ ] Live Market Scanning (Real-time signal generation)
[ ] Backtest Engine (Historical performance simulation)
[ ] Signal Alerting (Email/SMS notifications)
[ ] Dashboard (Web UI for monitoring)
[ ] Multi-Strategy Support
[ ] Risk Management Module
[ ] Trade Execution Module
[ ] Performance Analytics
