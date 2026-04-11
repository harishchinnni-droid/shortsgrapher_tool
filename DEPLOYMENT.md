# VPS Deployment Guide

This guide covers deploying Stock Automation Tool to a production VPS with high availability and security.

---

## System Requirements

- **OS**: Ubuntu 20.04+ or Windows Server 2019+
- **Python**: 3.8+
- **RAM**: 2GB minimum
- **Disk**: 20GB SSD (for historical data)
- **Network**: Static IP, port 443 access to broker APIs

---

## Pre-Deployment Checklist

- [ ] VPS provisioned and running
- [ ] Python 3.8+ installed
- [ ] Git installed for repository cloning
- [ ] SQL Server access configured
- [ ] Broker API keys obtained
- [ ] Firewall configured (allow outbound HTTPS)
- [ ] SSL certificates installed (if using HTTPS)
- [ ] Backup strategy in place

---

## Linux Deployment (Ubuntu 20.04+)

### Step 1: System Setup

```bash
# Update system
sudo apt-get update
sudo apt-get upgrade -y

# Install dependencies
sudo apt-get install -y python3.9 python3.9-venv python3.9-dev
sudo apt-get install -y git curl wget unzip
sudo apt-get install -y build-essential libssl-dev libffi-dev
```

### Step 2: Clone Repository

```bash
# Create app directory
sudo mkdir -p /opt/stock-automation
sudo chown $USER:$USER /opt/stock-automation
cd /opt/stock-automation

# Clone repository
git clone <repository-url> .
```

### Step 3: Create Virtual Environment

```bash
cd /opt/stock-automation

# Create venv
python3.9 -m venv venv

# Activate venv
source venv/bin/activate

# Install dependencies
pip install --upgrade pip
pip install -r requirements.txt
```

### Step 4: Configure Environment

```bash
# Create .env file with actual credentials
cat > .env << 'EOF'
# Zerodha API
ZERODHA_API_KEY=your_actual_key_here
ZERODHA_USER_ID=your_user_id_here
ZERODHA_PASSWORD=your_password_here
ZERODHA_API_SECRET=your_secret_here

# Angel One API
ANGEL_API_KEY=your_actual_key_here
ANGEL_CLIENT_ID=your_client_id_here
ANGEL_USER_ID=your_user_id_here
ANGEL_PASSWORD=your_password_here
ANGEL_TOTP_SECRET=your_totp_secret_here

# Database
DB_SERVER=your_database_server
DB_OPERATIONAL=FnO_Apr26
DB_HISTORICAL=Historical_Database

# Application
LOCAL_STORAGE_DIR=/opt/stock-automation/local_trading_data
LOG_LEVEL=INFO
DEBUG_MODE=False
EOF

# Restrict .env permissions
chmod 600 .env

# Verify it can't be read by others
ls -la .env  # Should show: -rw------- 1 user user
```

### Step 5: Setup systemd Service

```bash
# Create service file
sudo tee /etc/systemd/system/stock-automation.service > /dev/null << 'EOF'
[Unit]
Description=Stock Automation Tool
Documentation=https://github.com/yourusername/stock-automation
After=network.target
Wants=network-online.target

[Service]
Type=simple
User=trading_user
Group=trading_group
WorkingDirectory=/opt/stock-automation

# Activate virtual environment and run
ExecStart=/opt/stock-automation/venv/bin/python main.py

# Auto-restart on failure
Restart=always
RestartSec=10
StandardOutput=journal
StandardError=journal

# Security hardening
PrivateTmp=yes
NoNewPrivileges=true
ProtectHome=yes
ProtectSystem=strict
ReadWritePaths=/opt/stock-automation/local_trading_data /opt/stock-automation/logs

# Resource limits
MemoryMax=512M
FileDescriptorMax=65536

[Install]
WantedBy=multi-user.target
EOF

# Reload systemd and enable service
sudo systemctl daemon-reload
sudo systemctl enable stock-automation
```

### Step 6: Create Dedicated User

```bash
# Create non-privileged user
sudo useradd -s /bin/bash -d /opt/stock-automation trading_user

# Set ownership
sudo chown -R trading_user:trading_user /opt/stock-automation

# Set proper permissions
sudo chmod 750 /opt/stock-automation
sudo chmod 640 /opt/stock-automation/.env
sudo chmod 755 /opt/stock-automation/local_trading_data
```

### Step 7: Start Service

```bash
# Start the service
sudo systemctl start stock-automation

# Check status
sudo systemctl status stock-automation

# View logs
sudo journalctl -u stock-automation -f  # Follow logs in real-time
```

### Step 8: Monitoring & Logs

```bash
# Check if process is running
ps aux | grep main.py

# View recent logs (last 50 lines)
sudo journalctl -u stock-automation -n 50

# View errors only
sudo journalctl -u stock-automation -p err

# Archive logs daily
sudo tee /etc/logrotate.d/stock-automation > /dev/null << 'EOF'
/opt/stock-automation/logs/*.log {
    daily
    rotate 30
    compress
    delaycompress
    notifempty
    create 640 trading_user trading_group
    sharedscripts
    postrotate
        systemctl reload stock-automation > /dev/null 2>&1 || true
    endscript
}
EOF
```

---

## Windows Server Deployment

### Step 1: Install Python

1. Download Python 3.9+ from python.org
2. Run installer, check "Add Python to PATH"
3. Verify installation:
   ```powershell
   python --version
   ```

### Step 2: Clone Repository

```powershell
# Create directory
New-Item -ItemType Directory -Path "C:\Apps\stock-automation"
cd "C:\Apps\stock-automation"

# Clone repository
git clone <repository-url> .
```

### Step 3: Create Virtual Environment

```powershell
# Create venv
python -m venv venv

# Activate venv
.\venv\Scripts\Activate.ps1

# Install dependencies
pip install --upgrade pip
pip install -r requirements.txt
```

### Step 4: Configure Environment

```powershell
# Create .env file
@"
ZERODHA_API_KEY=your_actual_key_here
ZERODHA_USER_ID=your_user_id_here
ZERODHA_PASSWORD=your_password_here
ZERODHA_API_SECRET=your_secret_here

ANGEL_API_KEY=your_actual_key_here
ANGEL_CLIENT_ID=your_client_id_here
ANGEL_USER_ID=your_user_id_here
ANGEL_PASSWORD=your_password_here
ANGEL_TOTP_SECRET=your_totp_secret_here

DB_SERVER=your_database_server
DB_OPERATIONAL=FnO_Apr26
DB_HISTORICAL=Historical_Database

LOCAL_STORAGE_DIR=C:\Apps\stock-automation\local_trading_data
LOG_LEVEL=INFO
DEBUG_MODE=False
"@ | Out-File -FilePath ".env" -Encoding UTF8

# Restrict .env access to current user
icacls ".env" /inheritance:r /grant:r "$($env:USERNAME):(F)"
```

### Step 5: Create Windows Service

```powershell
# Install NSSM (Non-Sucking Service Manager)
choco install nssm

# Or download from: https://nssm.cc/download

# Create service
nssm install StockAutomation "C:\Apps\stock-automation\venv\Scripts\python.exe" "C:\Apps\stock-automation\main.py"

# Set working directory
nssm set StockAutomation AppDirectory "C:\Apps\stock-automation"

# Set startup type to automatic
nssm set StockAutomation Start SERVICE_AUTO_START

# Start service
nssm start StockAutomation

# Check service status
nssm status StockAutomation
```

---

## Docker Deployment (Recommended)

### Dockerfile

```dockerfile
FROM python:3.9-slim

WORKDIR /app

# Install system dependencies
RUN apt-get update && apt-get install -y \
    git \
    curl \
    && rm -rf /var/lib/apt/lists/*

# Copy requirements
COPY requirements.txt .

# Install Python dependencies
RUN pip install --no-cache-dir -r requirements.txt

# Copy application
COPY . .

# Create non-root user
RUN useradd -m -u 1000 trader && \
    chown -R trader:trader /app
USER trader

# Health check
HEALTHCHECK --interval=30s --timeout=10s --start-period=5s --retries=3 \
    CMD python -c "import sys; sys.exit(0)"

# Run application
CMD ["python", "main.py"]
```

### Docker Compose

```yaml
version: '3.9'

services:
  stock-automation:
    build: .
    container_name: stock-automation
    restart: always
    
    # Environment variables
    env_file: .env
    environment:
      LOG_LEVEL: INFO
    
    # Volumes for persistence
    volumes:
      - ./local_trading_data:/app/local_trading_data
      - ./logs:/app/logs
    
    # Resource limits
    deploy:
      resources:
        limits:
          cpus: '1'
          memory: 512M
        reservations:
          cpus: '0.5'
          memory: 256M
    
    # Logging
    logging:
      driver: "json-file"
      options:
        max-size: "10m"
        max-file: "3"
    
    # Network
    networks:
      - stock-net

volumes:
  local_trading_data:
  logs:

networks:
  stock-net:
    driver: bridge
```

### Deploy with Docker

```bash
# Build image
docker-compose build

# Start service
docker-compose up -d

# Check logs
docker-compose logs -f stock-automation

# Stop service
docker-compose down
```

---

## Health Checks & Monitoring

### Health Check Script

```bash
#!/bin/bash
# health_check.sh

PIDFILE="/var/run/stock-automation.pid"
LOG_FILE="/opt/stock-automation/logs/health.log"

echo "[$(date)] Running health check..." >> $LOG_FILE

# Check if process is running
if ! pgrep -f "python main.py" > /dev/null; then
    echo "[$(date)] ERROR: Process not running!" >> $LOG_FILE
    sudo systemctl start stock-automation
    exit 1
fi

# Check if recent logs exist
if ! find /opt/stock-automation/logs -name "*.log" -mmin -60 | grep -q .; then
    echo "[$(date)] ERROR: No recent logs!" >> $LOG_FILE
    exit 1
fi

echo "[$(date)] Health check passed" >> $LOG_FILE
exit 0
```

Schedule with cron:

```bash
# Run every 5 minutes
*/5 * * * * /opt/stock-automation/health_check.sh
```

### Monitoring with Prometheus

```yaml
# prometheus.yml
global:
  scrape_interval: 15s

scrape_configs:
  - job_name: 'stock-automation'
    static_configs:
      - targets: ['localhost:9090']
```

---

## Backup Strategy

### Daily Backup Script

```bash
#!/bin/bash
# backup.sh

BACKUP_DIR="/backups/stock-automation"
TIMESTAMP=$(date +%Y%m%d_%H%M%S)

# Create backup directory
mkdir -p $BACKUP_DIR

# Backup local trading data
tar -czf $BACKUP_DIR/local_trading_data_$TIMESTAMP.tar.gz \
    /opt/stock-automation/local_trading_data

# Backup logs
tar -czf $BACKUP_DIR/logs_$TIMESTAMP.tar.gz \
    /opt/stock-automation/logs

# Keep only last 30 days
find $BACKUP_DIR -type f -mtime +30 -delete

echo "Backup completed: $TIMESTAMP"
```

Schedule daily:

```bash
# Add to crontab
0 2 * * * /opt/stock-automation/backup.sh
```

---

## Security Hardening

### Firewall Rules

```bash
# Allow only HTTPS outbound to broker APIs
sudo ufw allow out 443/tcp

# Block all other outbound
sudo ufw default deny outgoing

# Allow SSH for management
sudo ufw allow 22/tcp
```

### File Permissions

```bash
# Ensure .env is readable only by owner
chmod 600 /opt/stock-automation/.env

# Ensure logs are readable
chmod 640 /opt/stock-automation/logs/*.log

# Restrict directory access
chmod 750 /opt/stock-automation
chmod 750 /opt/stock-automation/local_trading_data
```

### Credential Rotation

```bash
# Create rotation script
cat > /opt/stock-automation/rotate_credentials.sh << 'EOF'
#!/bin/bash
# Update .env with new credentials
nano /opt/stock-automation/.env

# Restart service
sudo systemctl restart stock-automation

# Verify running
sudo systemctl status stock-automation
EOF

chmod +x /opt/stock-automation/rotate_credentials.sh

# Schedule quarterly rotation (first day of quarter)
0 0 1 */3 * /opt/stock-automation/rotate_credentials.sh
```

---

## Troubleshooting

### Service Won't Start

```bash
# Check service status
sudo systemctl status stock-automation

# View error logs
sudo journalctl -u stock-automation -p err

# Test manually
cd /opt/stock-automation
source venv/bin/activate
python main.py

# Check .env file
cat .env | head -5
```

### High CPU Usage

```bash
# Monitor processes
top -p $(pgrep -f main.py)

# Check for hanging processes
ps aux | grep main.py

# Kill and restart if needed
sudo systemctl restart stock-automation
```

### Disk Space Issues

```bash
# Check disk usage
df -h /opt/stock-automation

# Check large files
du -sh /opt/stock-automation/*

# Clean old logs
find /opt/stock-automation/logs -mtime +30 -delete
```

### Database Connection Errors

```bash
# Test SQL Server connectivity
/opt/stock-automation/venv/bin/python << 'EOF'
import os
from dotenv import load_dotenv
import pyodbc

load_dotenv()
server = os.getenv('DB_SERVER')
database = os.getenv('DB_OPERATIONAL')

try:
    conn = pyodbc.connect(f'Driver={{ODBC Driver 17 for SQL Server}};Server={server};Database={database};Trusted_Connection=yes;')
    print("✓ Connection successful")
    conn.close()
except Exception as e:
    print(f"✗ Connection failed: {e}")
EOF
```

---

## References

- [systemd Service Documentation](https://www.freedesktop.org/software/systemd/man/systemd.service.html)
- [Docker Best Practices](https://docs.docker.com/develop/dev-best-practices/)
- [NSSM Windows Service Manager](https://nssm.cc/)
- [Ubuntu Server Guide](https://ubuntu.com/server/docs)

---

## Support

For deployment issues or questions, contact your DevOps team or create an issue on GitHub with:
- [ ] OS and Python version
- [ ] Error logs (sanitized)
- [ ] Deployment method (Linux/Windows/Docker)
- [ ] Steps to reproduce
