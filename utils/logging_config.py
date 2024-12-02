import logging
from pathlib import Path
from datetime import datetime

def setup_logging(name=None):
    """Set up logging with a unique name for each run"""
    # Remove any existing handlers from root logger
    root = logging.getLogger()
    for handler in root.handlers[:]:
        root.removeHandler(handler)
    
    # Create logs directory if it doesn't exist
    log_dir = Path("logs")
    log_dir.mkdir(exist_ok=True)
    
    # Create a unique logger name if not provided
    if name is None:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        name = f"xliff_converter_{timestamp}"
    
    # Get or create logger
    logger = logging.getLogger(name)
    logger.setLevel(logging.DEBUG)
    logger.handlers = []  # Clear any existing handlers
    
    # Console handler - only warnings and errors
    console_handler = logging.StreamHandler()
    console_handler.setLevel(logging.WARNING)
    console_format = logging.Formatter('%(levelname)s: %(message)s')
    console_handler.setFormatter(console_format)
    
    # File handler - detailed debug info
    log_file = log_dir / f"{name}.log"
    file_handler = logging.FileHandler(log_file, mode='w')  # 'w' mode to overwrite
    file_handler.setLevel(logging.DEBUG)
    file_format = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
    file_handler.setFormatter(file_format)
    
    # Add handlers
    logger.addHandler(console_handler)
    logger.addHandler(file_handler)
    
    return logger
    