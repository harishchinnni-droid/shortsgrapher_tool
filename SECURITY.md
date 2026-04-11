# Security Architecture & Best Practices

## Executive Summary

This document outlines the bulletproof security architecture implemented in Stock Automation Tool. It covers API key management, data protection, timing controls, and deployment guidelines.

---

## 1. API Key Management

### Environment Variables (CRITICAL)

All sensitive credentials are loaded from `.env` file using `python-dotenv`:

```python
from dotenv import load_dotenv
import os

load_dotenv()
api_key = os.getenv("ZERODHA_API_KEY")
```

**Why this matters:**
- ✅ Credentials are NOT hardcoded in source code
- ✅ Different credentials per environment (dev/staging/prod)
- ✅ No accidental credential leaks in git history
- ✅ Secrets never appear in logs

### .env File Protection

The `.env` file is protected by `.gitignore`:

```gitignore
# CRITICAL: Sensitive Files - NEVER commit these to version control
.env
.env.local
.env.*.local
credentials.json
*.key
*.pem
session_cache.json
```

**Verification:**
```bash
# Check that .env is properly excluded
git check-ignore .env      # Should return .env
git status                 # .env should NOT appear
```

### Credential Validation

On startup, the tool validates that API keys are present:

```python
def _validate_api_keys(self):
    """Validate that required API keys are configured in environment."""
    zerodha_key = os.getenv("ZERODHA_API_KEY")
    angel_key = os.getenv("ANGEL_API_KEY")
    
    if not zerodha_key and not angel_key:
        raise ValueError("No broker API keys found in environment")
```

**Behavior:**
- Fails immediately if credentials missing
- Provides clear error message
- Prevents tool from running with incomplete setup

---

## 2. Data Storage Security

### Local-Only Storage

All data is stored locally in `./local_trading_data/`:

```
./local_trading_data/
├── indicators/
│   ├── rsi/
│   ├── vwap/
│   ├── adx/
│   └── ... (one folder per indicator)
└── audit_logs/
```

**Why local-only:**
- ✅ No external API calls to store data
- ✅ Data never leaves your machine/VPS
- ✅ No cloud account credentials needed
- ✅ Complete data sovereignty
- ✅ Faster performance (local disk I/O)

### Individual Indicator Audit Trails

Each indicator maintains separate CSV files:

```
local_trading_data/indicators/rsi/
├── BANKNIFTY_2026-04-11.csv
├── NIFTY_2026-04-11.csv
└── INDIAVIX_2026-04-11.csv

local_trading_data/indicators/vwap/
├── BANKNIFTY_2026-04-11.csv
├── NIFTY_2026-04-11.csv
└── INDIAVIX_2026-04-11.csv
```

**Benefits:**
- ✅ Isolated accountability (each indicator auditable)
- ✅ Easy debugging (find issues in specific indicator)
- ✅ No single point of failure (if one indicator corrupts, others safe)
- ✅ Regulatory compliance (audit trail per signal component)

### File Permissions

For production deployments:

```bash
# Restrict .env to owner only
chmod 600 .env

# Restrict local_trading_data to owner only
chmod 700 local_trading_data/
chmod 600 local_trading_data/indicators/**/*.csv
```

---

## 3. Timing Control Security

### Strict 5-Minute Candle Close Enforcement

The tool enforces exact timing before data fetch:

```python
def fetch_data_on_candle_close(self, timeframe_minutes: int = 5):
    """
    Halt execution and wait until candle officially closes.
    Prevents garbage data from mid-candle fetches.
    """
    candle_info = self._get_next_candle_close_time(timeframe_minutes)
    seconds_to_wait = candle_info["seconds_to_wait"]
    
    if seconds_to_wait > 0:
        time.sleep(seconds_to_wait)  # Sleep until exact close time
    
    return True
```

**Why this matters:**
- ✅ Prevents partial candle data (incomplete OHLCV)
- ✅ Ensures all market participants see same data
- ✅ Avoids race conditions on data fetch
- ✅ Guarantees signal consistency

**Timing Precision:**
- ✓ Sleeps until exact second of candle close
- ✓ Accounts for system time skew
- ✓ Logs all timing decisions

---

## 4. Input Validation

### Backtest Date Validation

Manual date entry prevents accidental data corruption:

```python
def run_backtest(self, manual_date: str):
    """
    Requires manual date entry in YYYY-MM-DD format.
    Prevents garbage data runs.
    """
    try:
        target_date = datetime.strptime(manual_date, "%Y-%m-%d").date()
    except ValueError:
        raise ValueError(f"Invalid date format: '{manual_date}'. Expected YYYY-MM-DD.")
    
    if target_date >= date.today():
        raise ValueError(f"Backtest date must be in the past. Provided: {target_date}")
```

**Validation checks:**
- ✅ Format validation (YYYY-MM-DD)
- ✅ Date must be in the past (no future data)
- ✅ No default dates (explicit entry required)
- ✅ Clear error messages

---

## 5. Logging & Audit Trails

### Timestamped Logging

All operations logged with timestamps:

```
logs/2026-04-11_14-30-22.log
Format: YYYY-MM-DD HH:MM:SS - [LEVEL] - function:line - message
```

**Log contents:**
- ✅ API key validation checks
- ✅ Candle timing enforcement
- ✅ Data fetch operations
- ✅ Indicator calculations
- ✅ Signal generation
- ✅ Error tracebacks
- ✅ Performance metrics

### Sensitive Data Masking

Logs do NOT contain:
- ✗ API keys or tokens
- ✗ User credentials
- ✗ Account numbers
- ✗ Trade amounts (only symbols)

---

## 6. VPS Deployment Security

### Environment Setup for Production

1. **Never use .env file on production**
   - Use VPS environment variable system instead
   - Or use container secrets (Docker/K8s)

   ```bash
   # Export credentials as environment variables
   export ZERODHA_API_KEY="xxx"
   export ZERODHA_USER_ID="xxx"
   # ... etc
   ```

2. **IP Whitelisting**
   ```bash
   # Configure broker API to accept only VPS IP
   # Restrict at firewall level:
   telnet api.zerodha.com 443  # Should only work from VPS
   ```

3. **Process Isolation with systemd**
   ```ini
   [Unit]
   Description=Stock Automation Tool
   After=network.target
   
   [Service]
   Type=simple
   User=trading
   WorkingDirectory=/opt/stock-automation
   ExecStart=/opt/stock-automation/venv/bin/python main.py
   Restart=always
   RestartSec=10
   
   # Security hardening
   PrivateTmp=yes
   NoNewPrivileges=true
   ProtectHome=yes
   ProtectSystem=strict
   ReadWritePaths=/opt/stock-automation/local_trading_data
   
   [Install]
   WantedBy=multi-user.target
   ```

4. **Database Connection Security**
   ```python
   # Window Authentication (recommended on Windows)
   conn_str = 'DRIVER={{ODBC Driver 17 for SQL Server}};SERVER=server;DATABASE=db;Trusted_Connection=yes;'
   
   # With username/password (less secure, use Vault)
   conn_str = 'DRIVER=...;SERVER=server;DATABASE=db;UID=user;PWD=password;'
   ```

5. **SSL/TLS for Broker API**
   - Zerodha uses HTTPS automatically
   - Verify certificate in production
   - Disable insecure protocols

---

## 7. Credential Rotation

### Quarterly Rotation Schedule

```bash
# March 31, June 30, September 30, December 31

# Step 1: Generate new API keys in broker console
# Step 2: Update .env or environment variables
# Step 3: Test in staging first
# Step 4: Deploy to production
# Step 5: Revoke old API keys
# Step 6: Document rotation in audit log
```

### Database Password Rotation

```bash
# Monthly rotation
# 1. Change password in SQL Server
# 2. Update environment variables
# 3. Restart process
# 4. Monitor connection logs
```

---

## 8. Incident Response

### If API Keys Compromised

1. **Immediate actions:**
   ```bash
   # 1. Revoke all API keys in broker console
   # 2. Stop the trading bot
   sudo systemctl stop stock-automation
   
   # 3. Generate new API keys
   # 4. Update .env or environment
   # 5. Restart
   sudo systemctl start stock-automation
   ```

2. **Investigation:**
   - Review logs for unauthorized access attempts
   - Check broker account for suspicious orders
   - Review all IP connections in logs
   - Check git history for accidental commits

3. **Prevention:**
   ```bash
   # Add git precommit hook to prevent .env commits
   cat > .git/hooks/pre-commit << 'EOF'
   #!/bin/bash
   if git diff --cached | grep -q "^+.*ZERODHA_API_KEY\|^+.*ANGEL_API_KEY"
   then
       echo "ERROR: Refusing to commit API keys to git!"
       exit 1
   fi
   EOF
   chmod +x .git/hooks/pre-commit
   ```

---

## 9. Code Security Best Practices

### Never Hardcode Secrets

```python
# ❌ WRONG - NEVER DO THIS
api_key = "abc123xyz789"

# ✅ CORRECT - Use environment variables
api_key = os.getenv("ZERODHA_API_KEY")
```

### Validate All Inputs

```python
# ✅ Input validation required
def set_symbols(self, symbols: list):
    if not isinstance(symbols, list):
        raise TypeError("symbols must be a list")
    if len(symbols) == 0:
        raise ValueError("symbols cannot be empty")
    if len(symbols) > 100:
        raise ValueError("Maximum 100 symbols allowed")
```

### Error Handling Without Leaking Data

```python
# ❌ WRONG - Exposes sensitive info
except Exception as e:
    logger.error(f"API call failed: {e}")  # Error might contain API key

# ✅ CORRECT - Generic error message
except Exception as e:
    logger.error(f"API call failed: Connection error")
    logger.debug(f"Details: {str(e)[:50]}")  # Log details at debug level only
```

### Regular Security Audits

```bash
# Check for secrets in git history
git log -p -S "api_key\|password\|token" --all

# Scan for hardcoded secrets
pip install detect-secrets
detect-secrets scan --all-files

# Check dependencies for vulnerabilities
pip install safety
safety check
```

---

## 10. Compliance & Auditing

### Audit Trail Components

1. **Timestamped logs** → When actions occurred
2. **Individual indicator files** → Which signals were calculated
3. **Git commits** → What code changes were made
4. **Database records** → Trade history
5. **System logs** → Process restarts, errors

### Example Audit Query

```bash
# Find all RSI calculations on 2026-04-11
ls -la local_trading_data/indicators/rsi/*2026-04-11*

# Check logs for specific error
grep ERROR logs/*2026-04-11* | head -20

# Verify git history
git log --oneline --all | head -20
```

---

## 11. Dependency Security

### Regular Updates

```bash
# Check for outdated packages
pip list --outdated

# Update all packages
pip install --upgrade pip
pip install -U -r requirements.txt

# Verify integrity after update
python -c "import kiteconnect; print(kiteconnect.__version__)"
```

### Pinned Versions (requirements.txt)

```
# Predictable versions for production
requests==2.31.0
pandas==1.5.3
numpy==1.24.3
python-dotenv==1.0.0
```

### Vulnerability Scanning

```bash
# Install and run safety check
pip install safety
safety check
# Output: No security vulnerabilities found!
```

---

## Quick Checklist for Developers

- [ ] Never commit `.env` to git
- [ ] Never hardcode API keys in source code
- [ ] Always load credentials from environment variables
- [ ] Validate all user inputs (dates, symbols, etc.)
- [ ] Don't log sensitive information
- [ ] Test in staging before deploying to production
- [ ] Rotate credentials quarterly
- [ ] Review logs regularly for anomalies
- [ ] Keep dependencies updated
- [ ] Use strong passwords for database access

---

## Support

For security issues or vulnerabilities:
1. **DO NOT** create public GitHub issues
2. Email: security@yourdomain.com
3. Provide detailed description of vulnerability
4. Allow 48-72 hours for response

---

## References

- [OWASP Secure Coding Practices](https://cheatsheetseries.owasp.org/)
- [Python-dotenv Documentation](https://python-dotenv.readthedocs.io/)
- [SEC Rule 17a-3 (Trading Records)](https://www.sec.gov/rules/final/34-41991.txt)
- [SEBI Algorithmic Trading Guidelines](https://www.sebi.gov.in/)
