import os
from dotenv import load_dotenv

# Load environment variables
load_dotenv()

# Azure Storage Configuration
#STORAGE_ACCOUNT_NAME = os.getenv('STORAGE_ACCOUNT_NAME', '')
#CONTAINER_NAME = os.getenv('CONTAINER_NAME', '')
#CONNECTION_STRING = os.getenv('STORAGE_CONNECTION_STRING')

# Azure Vision Configuration
VISION_ENDPOINT = os.getenv('VISION_ENDPOINT')
VISION_API_KEY = os.getenv('VISION_API_KEY')

# Processing Configuration
MAX_THREADS = int(os.getenv('MAX_THREADS', '3'))
MAX_RETRIES = int(os.getenv('MAX_RETRIES', '3'))
POLLING_INTERVAL = float(os.getenv('POLLING_INTERVAL', '1.5'))

# Output Configuration
OUTPUT_DIR = os.getenv('OUTPUT_DIR', './output')
LOG_LEVEL = os.getenv('LOG_LEVEL', 'INFO')

# Validate required environment variables
def validate_config():
    """Validate that all required environment variables are set"""
    missing_vars = []
    
    if not VISION_ENDPOINT:
        missing_vars.append('VISION_ENDPOINT')
    if not VISION_API_KEY:
        missing_vars.append('VISION_API_KEY')
        
    if missing_vars:
        raise ValueError(f"Missing required environment variables: {', '.join(missing_vars)}")
        
    return True 