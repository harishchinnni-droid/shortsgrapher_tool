"""
Entry point for Stock Automation Tool
Run this file to start the complete automation workflow: python main.py
"""

from src.main import main


if __name__ == "__main__":
    success = main()
    exit(0 if success else 1)
