"""
MPS 工單資訊自動同步腳本
每 4 小時執行一次，從外部 MPS API 同步工單出貨日期與客戶資訊
"""
import sys
import os
import json
from datetime import datetime

# Add the project directory to the Python path
project_path = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.append(project_path)

os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'material_requisition_system.settings.production')

import django
django.setup()

# Timestamp file path
TIMESTAMP_FILE = os.path.join(project_path, 'scratch', 'mps_sync_timestamp.json')


def save_sync_timestamp(success, message):
    """Save the last sync result to a JSON file."""
    os.makedirs(os.path.dirname(TIMESTAMP_FILE), exist_ok=True)
    data = {
        'last_sync_time': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
        'success': success,
        'message': message,
    }
    with open(TIMESTAMP_FILE, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def run_sync_mps():
    """Execute the MPS sync."""
    print("Starting MPS Order Info Sync...")
    
    from requisitions.services.sap_sync import sync_external_order_info
    
    success, message = sync_external_order_info()
    
    if success:
        print(f"✅ {message}")
    else:
        print(f"❌ {message}")
    
    save_sync_timestamp(success, message)
    print(f"Timestamp saved to {TIMESTAMP_FILE}")
    return success


if __name__ == "__main__":
    run_sync_mps()
