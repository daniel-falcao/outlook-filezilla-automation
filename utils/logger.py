"""utils/logger.py
Configures and returns a logger that writes to both console and a log file."""

import logging
import os
from datetime import datetime


def setup_logger(log_dir: str = 'logs') -> logging.Logger:
    '''
    Sets up and returns a configured logger instance.

    The logger writes INFO-level messages to stdout and DEBUG-level messages
    to a rotating daily log file inside `log_dir`.

    Args:
        log_dir: Directory where log files will be stored.

    Returns:
        A configured logging.Logger instance.
    '''
    os.makedirs(log_dir, exist_ok=True)

    log_filename = os.path.join(
        log_dir,
        f'automation_{datetime.now().strftime("%Y-%m-%d")}.log'
    )

    logger = logging.getLogger('automation')
    logger.setLevel(logging.DEBUG)

    # Avoid adding duplicate handlers if logger is called more than once
    if logger.handlers:
        return logger

    # Console handler — INFO and above
    console_handler = logging.StreamHandler()
    console_handler.setLevel(logging.INFO)
    console_format = logging.Formatter(
        '[%(asctime)s] %(levelname)s - %(message)s',
        datefmt='%H:%M:%S'
    )
    console_handler.setFormatter(console_format)

    # File handler — DEBUG and above
    file_handler = logging.FileHandler(log_filename, encoding='utf-8')
    file_handler.setLevel(logging.DEBUG)
    file_format = logging.Formatter(
        '[%(asctime)s] %(levelname)s [%(module)s:%(lineno)d] - %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S'
    )
    file_handler.setFormatter(file_format)

    logger.addHandler(console_handler)
    logger.addHandler(file_handler)

    return logger
