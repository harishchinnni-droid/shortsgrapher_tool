"""
Utility functions for Stock Automation Tool
"""

import logging
import yaml
from typing import Dict, Any


def setup_logger(name: str) -> logging.Logger:
    """
    Setup logger for the application.
    
    Args:
        name: Logger name
        
    Returns:
        Configured logger instance
    """
    logger = logging.getLogger(name)
    
    if not logger.handlers:
        handler = logging.StreamHandler()
        formatter = logging.Formatter(
            '%(asctime)s - %(name)s - %(levelname)s - %(message)s'
        )
        handler.setFormatter(formatter)
        logger.addHandler(handler)
        logger.setLevel(logging.INFO)
    
    return logger


def load_config(config_path: str) -> Dict[str, Any]:
    """
    Load configuration from YAML file.
    
    Args:
        config_path: Path to configuration file
        
    Returns:
        Configuration dictionary
    """
    try:
        with open(config_path, 'r') as f:
            config = yaml.safe_load(f)
        return config if config else {}
    except FileNotFoundError:
        logging.warning(f"Config file not found: {config_path}")
        return {}
    except yaml.YAMLError as e:
        logging.error(f"Error parsing YAML config: {e}")
        return {}
