"""
Logging configuration module for Stock Automation Tool.
Handles logging in the format: <current_date_time>.log
"""

import logging
import os
from datetime import datetime
from pathlib import Path


class LoggingManager:
    """Manages logging configuration for the application."""
    
    _logger = None
    _log_file = None
    
    @staticmethod
    def setup_logger(log_dir="logs"):
        """
        Setup logger with timestamp-based log file.
        
        Args:
            log_dir: Directory to store log files (default: "logs")
            
        Returns:
            Configured logger instance
        """
        if LoggingManager._logger is not None:
            return LoggingManager._logger
        
        # Create log directory if it doesn't exist
        Path(log_dir).mkdir(parents=True, exist_ok=True)
        
        # Create log file with timestamp
        timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        log_file = os.path.join(log_dir, f"{timestamp}.log")
        LoggingManager._log_file = log_file
        
        # Configure logger
        logger = logging.getLogger("StockAutomation")
        logger.setLevel(logging.DEBUG)
        
        # File handler
        file_handler = logging.FileHandler(log_file, mode='a', encoding='utf-8')
        file_handler.setLevel(logging.DEBUG)
        
        # Console handler
        console_handler = logging.StreamHandler()
        console_handler.setLevel(logging.INFO)
        
        # Formatter
        formatter = logging.Formatter(
            '%(asctime)s - %(name)s - [%(levelname)s] - %(funcName)s:%(lineno)d - %(message)s',
            datefmt='%Y-%m-%d %H:%M:%S'
        )
        
        file_handler.setFormatter(formatter)
        console_handler.setFormatter(formatter)
        
        # Add handlers to logger
        logger.addHandler(file_handler)
        logger.addHandler(console_handler)
        
        LoggingManager._logger = logger
        
        logger.info("="*80)
        logger.info("STOCK AUTOMATION TOOL - INITIALIZATION")
        logger.info(f"Log file: {log_file}")
        logger.info("="*80)
        
        return logger
    
    @staticmethod
    def get_logger():
        """
        Get the configured logger instance.
        
        Returns:
            Logger instance
        """
        if LoggingManager._logger is None:
            return LoggingManager.setup_logger()
        return LoggingManager._logger
    
    @staticmethod
    def log_step_start(step_name):
        """
        Log the start of a process step.
        
        Args:
            step_name: Name of the step
        """
        logger = LoggingManager.get_logger()
        logger.info(f"\n{'='*60}")
        logger.info(f"STEP START: {step_name}")
        logger.info(f"{'='*60}")
    
    @staticmethod
    def log_step_completion(step_name, status="SUCCESS"):
        """
        Log the completion of a process step.
        
        Args:
            step_name: Name of the step
            status: Completion status (SUCCESS, FAILED, etc.)
        """
        logger = LoggingManager.get_logger()
        logger.info(f"STEP COMPLETION: {step_name} - {status}")
        logger.info(f"{'='*60}\n")
    
    @staticmethod
    def log_error(step_name, error_message, exception=None):
        """
        Log an error during a process step.
        
        Args:
            step_name: Name of the step where error occurred
            error_message: Description of the error
            exception: Optional exception object
        """
        logger = LoggingManager.get_logger()
        logger.error(f"ERROR in {step_name}: {error_message}")
        if exception:
            logger.exception(exception)
