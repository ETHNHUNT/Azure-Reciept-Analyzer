import re
import logging
from typing import Dict, List, Optional
from tenacity import retry, stop_after_attempt, wait_exponential
import os

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

def normalize_currency(amount: str) -> Optional[float]:
    """Normalize currency values to float."""
    try:
        # Remove currency symbols and commas
        cleaned = re.sub(r'[^\d.]', '', amount)
        return float(cleaned)
    except (ValueError, TypeError):
        logger.warning(f"Could not normalize amount: {amount}")
        return None

def validate_receipt_data(data: Dict) -> bool:
    """Validate the structure and content of receipt data."""
    required_fields = ['image_id', 'extracted_text', 'items']
    
    # Check required fields
    if not all(field in data for field in required_fields):
        logger.error(f"Missing required fields in receipt data: {data.get('image_id', 'unknown')}")
        return False
    
    # Validate items
    if not isinstance(data['items'], list):
        logger.error(f"Items must be a list: {data.get('image_id', 'unknown')}")
        return False
    
    return True

def clean_text(text: str) -> str:
    """Clean and normalize extracted text."""
    # Remove multiple spaces
    text = re.sub(r'\s+', ' ', text)
    # Remove special characters except basic punctuation
    text = re.sub(r'[^\w\s.,$]', '', text)
    return text.strip()

@retry(stop=stop_after_attempt(3), wait=wait_exponential(multiplier=1, min=4, max=10))
def process_items(items: List[str]) -> List[Dict]:
    """Process and structure receipt items."""
    processed_items = []
    for item in items:
        # Split item into components
        parts = item.split()
        if len(parts) >= 3:
            processed_items.append({
                'name': ' '.join(parts[:-2]),
                'quantity': parts[-2],
                'price': normalize_currency(parts[-1])
            })
    return processed_items

def detect_receipt_type(text: str) -> str:
    """Detect the type of receipt based on text patterns."""
    text = text.lower()
    
    if 'restaurant' in text or 'cafe' in text or 'food' in text:
        return 'restaurant'
    elif 'retail' in text or 'store' in text or 'shop' in text:
        return 'retail'
    elif 'service' in text or 'bill' in text:
        return 'service'
    else:
        return 'unknown'

