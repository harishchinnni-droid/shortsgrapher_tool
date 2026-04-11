"""
Entry point for Stock Automation Tool

BEFORE RUNNING:
1. Copy .env.example to .env
2. Fill in your API credentials in .env
3. CRITICAL: .env is in .gitignore and must NEVER be committed

Run this script to start the tool:
    python main.py

The tool will start in LIVE mode by default.
To run backtest, uncomment the backtest line in src/main.py
"""

import sys
from src.main import main


if __name__ == "__main__":
    try:
        success = main()
        sys.exit(0 if success else 1)
    except KeyboardInterrupt:
        print("\n\nApplication terminated by user (Ctrl+C)")
        sys.exit(0)
    except Exception as e:
        print(f"\nFATAL ERROR: {e}")
        sys.exit(1)
