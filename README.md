# Stock Automation Tool

A bulletproof Python-based automation tool for algorithmic stock trading with F&O derivatives. Built with security-first architecture, strict timing controls, and complete audit trails.

## Critical Security Notice

⚠️ **NEVER commit sensitive files to GitHub:**
- `.env` (contains API keys) - **Added to .gitignore**
- `session_cache.json` - **Added to .gitignore**
- Any credential files - **Added to .gitignore**

All API keys must be loaded from environment variables via `.env` file.

## Project Structure

```
Stock_Automation/
├── src/
│   ├── logging_config.py        # Logging with timestamp-based files
│   ├── main.py                  # Main orchestrator (BULLETPROOF ARCHITECTURE)
│   ├── broker_login.py          # Broker authentication
│   ├── data_ingestion.py        # Data fetching and storage
│   ├── incremental_sync.py      # Delta loading with NSE calendar
│   └── indicator_engine.py      # Technical indicators
│
├── local_trading_data/          # Local-only storage (NEVER cloud)
│   └── indicators/
│       ├── rsi/                 # Individual RSI audit trail
│       ├── vwap/                # Individual VWAP audit trail
│       ├── adx/                 # Individual ADX audit trail
│       ├── sqzmom/              # Individual Squeeze Momentum audit trail
│       ├── ema/                 # Individual EMA audit trail
│       ├── oi_dynamics/         # Individual OI Dynamics audit trail
│       └── breakout_score/      # Individual Breakout Score audit trail
│
├── logs/                        # Timestamped log files (auto-created)
│   └── YYYY-MM-DD_HH-MM-SS.log
│
├── main.py                      # Application entry point
├── .env                         # Environment variables (SENSITIVE - in .gitignore)
├── .env.example                 # Template for .env (safe to commit)
├── .gitignore                   # Excludes .env and sensitive files
├── requirements.txt             # Python dependencies
├── LICENSE                      # MIT License
└── README.md                    # This file
```

## Installation

### Prerequisites
- Python 3.8+
- pip (Python package manager)
- Windows Authentication for SQL Server
- API credentials for Zerodha and/or Angel One

### Setup Steps

1. **Clone the repository**
   ```bash
   git clone <repository-url>
   cd Stock_Automation
   ```

2. **Create and activate virtual environment**
   ```bash
   python -m venv venv
   venv\Scripts\activate
   ```

3. **Install dependencies**
   ```bash
   pip install -r requirements.txt
   ```

4. **Configure environment variables**
   ```bash
   # Copy the template
   copy .env.example .env
   
   # Edit .env with your actual credentials
   # NOTE: .env is automatically excluded from git (in .gitignore)
   ```

5. **Verify .env is in .gitignore**
   ```bash
   # Check that .env and sensitive files are protected
   git check-ignore .env
   ```

## Configuration

### Environment Variables (.env)

Create `.env` file from `.env.example`:

```env
# Zerodha API
ZERODHA_API_KEY=your_key_here
ZERODHA_USER_ID=your_user_id
ZERODHA_PASSWORD=your_password
ZERODHA_API_SECRET=your_secret

# Angel One API
ANGEL_API_KEY=your_key_here
ANGEL_CLIENT_ID=your_client_id
ANGEL_USER_ID=your_user_id
ANGEL_PASSWORD=your_password
ANGEL_TOTP_SECRET=your_totp_secret

# Database
DB_SERVER=DESKTOP-57PQCS1
DB_OPERATIONAL=FnO_Apr26
DB_HISTORICAL=Historical_Database

# Application
LOCAL_STORAGE_DIR=./local_trading_data
LOG_LEVEL=INFO
DEBUG_MODE=False
```

**CRITICAL**: Never commit `.env` to version control. It's automatically excluded by `.gitignore`.

## Usage

### Running Live Mode

```bash
python main.py
```

The application will:
1. ✓ Load API keys from `.env`
2. ✓ Wait for 5-minute candle to close
3. ✓ Fetch fresh market data
4. ✓ Calculate indicators in individual audit matrices
5. ✓ Evaluate trading signals
6. ✓ Execute trades (if triggered)
7. ✓ Loop back to step 2

### Running Backtest Mode

Uncomment in `src/main.py`:

```python
if __name__ == "__main__":
    tool = StockAutomationTool(local_storage_dir="./local_trading_data")
    
    # Run backtest for specific date
    success = tool.run(mode="backtest", backtest_date="2026-04-10")
    
    # OR run live
    # success = tool.run(mode="live")
```

Then run:
```bash
python main.py
```

## Architecture Highlights

### 1. Security First
- ✅ API keys stored in `.env` (never hardcoded)
- ✅ `.env` excluded from git via `.gitignore`
- ✅ Environment variable loading via `python-dotenv`
- ✅ Credentials validated at startup

### 2. Strict Timing Control
- ✅ **Enforced 5-minute candle close** before data fetch
- ✅ Prevents mid-candle garbage data
- ✅ Calculates exact sleep time to candle close
- ✅ Logs all timing decisions

### 3. Local-Only Storage
- ✅ All data stored locally (`./local_trading_data/`)
- ✅ No cloud dependencies or external storage
- ✅ Individual indicator audit trails (separated matrices)
- ✅ CSV format for easy auditing

### 4. Manual Date Entry for Backtesting
- ✅ Requires explicit date input (YYYY-MM-DD)
- ✅ Validates date is in the past
- ✅ Prevents accidental data corruption
- ✅ Supports point-in-time testing

### 5. Individual Indicator Auditing
Each indicator maintains its own audit trail:
- `local_trading_data/indicators/rsi/` → RSI calculations
- `local_trading_data/indicators/vwap/` → VWAP calculations
- `local_trading_data/indicators/adx/` → ADX calculations
- ... (one folder per indicator)

Benefits:
- Isolated accountability
- Easy debugging of specific indicators
- No single point of failure
- Auditable signal generation

### 6. Comprehensive Logging
- ✅ Timestamped log files: `logs/YYYY-MM-DD_HH-MM-SS.log`
- ✅ Both console and file output
- ✅ Detailed error tracebacks
- ✅ Step-by-step execution tracking

## API Reference

### StockAutomationTool

```python
from src.main import StockAutomationTool

# Initialize
tool = StockAutomationTool(local_storage_dir="./local_trading_data")

# Run live trading
tool.run(mode="live")

# Run backtest
tool.run(mode="backtest", backtest_date="2026-04-10")

# Check next candle close
candle_info = tool._get_next_candle_close_time(timeframe_minutes=5)
# Returns: {next_close_time, seconds_to_wait, current_candle, elapsed_seconds}

# Wait for candle close
tool.fetch_data_on_candle_close(timeframe_minutes=5)
```

## Logging

Logs are automatically timestamped:
```
logs/
├── 2026-04-11_14-30-22.log
├── 2026-04-11_15-45-10.log
└── 2026-04-12_09-15-33.log
```

Format: `YYYY-MM-DD HH:MM:SS - [LEVEL] - function:line - message`

Access logs from code:
```python
from src.logging_config import LoggingManager

logger = LoggingManager.get_logger()
logger.info("Message")
logger.error("Error")
```

## Database Schema

### Operational DB (FnO_Apr26)
- **Reference**: Symbol, KiteToken mapping

### Historical DB (Historical_Database)
- **Historical_Candles**: Date, Open, High, Low, Close, Volume, OI
- **Technical_Indicators**: Full indicator suite with scores

## Trading Hours & Market Calendar

Built-in NSE trading calendar for 2026:
- Market hours: 09:15 - 15:30 IST
- Excludes weekends and holidays
- 7 trading holidays configured

## Error Handling

The tool implements defensive programming:
- Try-catch on all critical operations
- Detailed error logging with tracebacks
- Graceful degradation (skip bad symbols, continue)
- Keyboard interrupt (Ctrl+C) for safe shutdown

## Performance Metrics

- **Candle timing**: <100ms precision
- **Data fetch**: ~0.4s per symbol (rate-limited)
- **Indicator calculation**: ~1-2s per symbol (45-day lookback)
- **Complete cycle**: 5-15 minutes (~100 symbols)

## Deployment on VPS

When deploying to production VPS:

1. **Use environment variables exclusively**
   ```bash
   export ZERODHA_API_KEY=your_key
   export ZERODHA_USER_ID=your_id
   # ... etc
   ```

2. **Never use .env file on production** (if possible)
   - Use VPS environment configuration instead
   - AWS Secrets Manager, Azure Key Vault, etc.

3. **Enable IP whitelisting** on broker API
   - Restrict API access to VPS IP only
   - Reduces attack surface

4. **Use systemd or supervisor** for persistence
   ```bash
   # systemd example
   [Unit]
   Description=Stock Automation Tool
   
   [Service]
   ExecStart=/path/to/venv/bin/python /path/to/main.py
   Restart=always
   ```

5. **Rotate credentials regularly**
   - Change API keys quarterly
   - Update database passwords monthly

## Troubleshooting

### Issue: "API Key missing. Set BROKER_API_KEY in environment"
**Solution**: Environment variables not loaded. Check:
- `.env` file exists and is in correct directory
- `.env` has required API keys filled in
- Virtual environment is activated

### Issue: "Candle timing enforcement loop"
**Solution**: System waiting for candle to close. This is normal.
- Tool sleeps until 5-minute candle closes
- Check logs to see exact wait time

### Issue: "SQL Server connection failed"
**Solution**: Database configuration issue.
- Verify Windows Authentication enabled
- Check server name and database names in `.env`
- Ensure SQL Server is running

### Issue: "No symbols found in Reference table"
**Solution**: Database empty or credentials wrong.
- Verify Reference table exists in FnO_Apr26
- Check KiteToken values are populated

## Contributing

1. Fork the repository
2. Create feature branch: `git checkout -b feature/your-feature`
3. Commit changes: `git commit -am 'Add feature'`
4. **NEVER commit .env or credentials**
5. Push to branch: `git push origin feature/your-feature`
6. Submit pull request

## Security Checklist

- [ ] `.env` file created from `.env.example`
- [ ] API credentials filled in `.env`
- [ ] `.env` is in `.gitignore`
- [ ] Never committed `.env` to git
- [ ] Running `python main.py` starts tool successfully
- [ ] Logs generated in `logs/` directory
- [ ] Indicator audit files created in `local_trading_data/indicators/`

## License

MIT License - See LICENSE file for details

## Support

For issues and questions:
1. Check logs in `logs/` directory
2. Enable DEBUG_MODE=True in `.env`
3. Create issue with log excerpt
4. Include backtesting date if applicable

## Version History

- **2.0.0** - Bulletproof Architecture (2026-04-11)
  - ✅ Environment variable security
  - ✅ Strict 5-minute candle enforcement
  - ✅ Local-only storage with individual audit trails
  - ✅ Manual date entry for backtesting
  - ✅ Comprehensive error handling

- **1.0.0** - Initial Release
  - Basic project structure
  - Configuration management
  - Logging setup
  - Unit test framework

a simple tool to automate trading
