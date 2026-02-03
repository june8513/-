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

from django.utils import timezone
from requisitions.models import AutoUploadConfig

BASE_DIR = settings.BASE_DIR 
DEFAULT_MONITOR_DIR = os.path.join(BASE_DIR, 'auto_upload', 'material_details') 
TIMESTAMP_FILE = os.path.join(DEFAULT_MONITOR_DIR, 'last_processed_timestamps.json')
REQUIRED_QTY_COL = '需求數量 (EINHEIT)'

os.makedirs(DEFAULT_MONITOR_DIR, exist_ok=True)

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
    
    # Check AutoUploadConfig first
    config = AutoUploadConfig.objects.filter(upload_type='material_details', is_active=True).first()
    
    monitor_paths = []
    if config and config.file_path and os.path.exists(config.file_path):
        if os.path.isdir(config.file_path):
             monitor_paths.append(config.file_path)
             print(f"Using configured directory: {config.file_path}")
        else:
             # It's a file
             monitor_paths.append(config.file_path)
             print(f"Using configured file: {config.file_path}")
    else:
        monitor_paths.append(DEFAULT_MONITOR_DIR)
        print(f"Using default directory: {DEFAULT_MONITOR_DIR}")
        
    last_processed_timestamps = load_timestamps()
    current_files_processed = set()

    for path in monitor_paths:
        if os.path.isfile(path):
            files_to_check = [path]
            is_single_file = True
        else:
            files_to_check = [os.path.join(path, f) for f in os.listdir(path) if f.lower().endswith(('.xlsx', '.xls'))]
            is_single_file = False
            
        for file_path in files_to_check:
            filename = os.path.basename(file_path)
            current_files_processed.add(filename)
            
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
                    
                    if config:
                        config.last_run = timezone.now()
                        config.last_status = "Success"
                        config.save()
                else:
                    print(f"No change detected for {filename}. Skipping.")

            except Exception as e:
                print(f"An error occurred while processing {filename}:")
                traceback.print_exc()
                if config:
                    config.last_status = f"Error: {str(e)}"
                    config.save()

    save_timestamps(last_processed_timestamps)

if __name__ == "__main__":
    run_monitor_material_details()