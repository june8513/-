import os
import time
import shutil
import json
import traceback
from django.core.management import call_command
from django.conf import settings

# Configure Django settings
os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'material_requisition_system.settings')
import django
django.setup()

BASE_DIR = settings.BASE_DIR 
MONITOR_DIR = os.path.join(BASE_DIR, 'auto_upload', 'material_details') 
TIMESTAMP_FILE = os.path.join(MONITOR_DIR, 'last_processed_timestamps.json')
REQUIRED_QTY_COL = '需求數量 (EINHEIT)'

os.makedirs(MONITOR_DIR, exist_ok=True)

def load_timestamps():
    if os.path.exists(TIMESTAMP_FILE):
        with open(TIMESTAMP_FILE, 'r') as f:
            return json.load(f)
    return {}

def save_timestamps(timestamps):
    with open(TIMESTAMP_FILE, 'w') as f:
        json.dump(timestamps, f, indent=4)

def run_monitor_material_details(): # Renamed function
    print(f"--- 2. Running Material Details Monitor ---")
    last_processed_timestamps = load_timestamps()
    current_files = set()

    for filename in os.listdir(MONITOR_DIR):
        if filename.lower().endswith('.xlsx') or filename.lower().endswith('.xls'):
            file_path = os.path.join(MONITOR_DIR, filename)
            current_files.add(filename)

            try:
                current_mtime = os.path.getmtime(file_path)
                
                if filename not in last_processed_timestamps or current_mtime > last_processed_timestamps[filename]:
                    print(f"Detected change in {filename}. Attempting to upload...")
                    call_command(
                        'auto_upload_material_details', 
                        path=file_path, 
                        qty_col=REQUIRED_QTY_COL
                    )
                    print(f"Successfully processed {filename}.")
                    last_processed_timestamps[filename] = current_mtime
                else:
                    print(f"No change detected for {filename}. Skipping.")

            except Exception as e:
                print(f"An error occurred while processing {filename}:")
                traceback.print_exc()
        else:
            if not filename.endswith('.json'):
                print(f"Skipping non-Excel file: {filename}")

    files_to_remove = [f for f in last_processed_timestamps if f not in current_files]
    for f in files_to_remove:
        print(f"Removing timestamp for deleted file: {f}")
        del last_processed_timestamps[f]

    save_timestamps(last_processed_timestamps)

if __name__ == "__main__":
    run_monitor_material_details()