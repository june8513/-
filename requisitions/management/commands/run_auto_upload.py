from django.core.management.base import BaseCommand
from requisitions.models import AutoUploadConfig
from requisitions.utils import process_inventory_excel, process_order_model_excel, process_material_details_excel, process_shipping_customer_excel
import os
import shutil
from django.utils import timezone
import time

class Command(BaseCommand):
    help = 'Automatically monitors and uploads SAP files based on configuration.'

    def add_arguments(self, parser):
        parser.add_argument('--loop', action='store_true', help='Run in an infinite loop')
        parser.add_argument('--interval', type=int, default=60, help='Interval in seconds for loop mode')

    def handle(self, *args, **options):
        loop_mode = options['loop']
        interval = options['interval']

        if loop_mode:
            self.stdout.write(self.style.SUCCESS(f'Starting Auto-Upload Monitor (Interval: {interval}s)...'))
            while True:
                self.process_configs()
                time.sleep(interval)
        else:
            self.process_configs()

    def process_configs(self):
        configs = AutoUploadConfig.objects.filter(is_active=True).order_by('priority')
        
        for config in configs:
            file_path = config.file_path
            
            # Check if file exists
            if not os.path.exists(file_path):
                continue

            # Check modification time
            try:
                current_mtime = os.path.getmtime(file_path)
            except OSError:
                self.stdout.write(self.style.ERROR(f"Could not read mtime for {file_path}"))
                continue

            # Compare with last processed mtime
            if config.last_processed_mtime and current_mtime <= config.last_processed_mtime:
                # File hasn't changed since last run
                continue

            self.stdout.write(f'Detected update for {config.get_upload_type_display()}: {file_path}')
            
            try:
                # Normalize type
                upload_type = str(config.upload_type).strip()
                self.stdout.write(f"DEBUG: Processing type '{upload_type}' (len={len(upload_type)})")

                # Process based on type
                if upload_type == 'inventory':
                    created, updated = process_inventory_excel(file_path)
                    msg = f"Success: Created {created}, Updated {updated}"
                    
                elif upload_type == 'order_model':
                    created, updated = process_order_model_excel(file_path)
                    msg = f"Success: Created {created}, Updated {updated}"
                    
                elif upload_type == 'material_details':
                    # Defaulting required quantity column to '需求數量'
                    created, updated, deactivated = process_material_details_excel(file_path, '需求數量') 
                    msg = f"Success: Created {created}, Updated {updated}, Deactivated {deactivated}"

                elif upload_type == 'semi_finished_model_db':
                     created, updated = process_order_model_excel(file_path)
                     msg = f"Success: Created {created}, Updated {updated}"

                elif upload_type == 'shipping_customer':
                     updated = process_shipping_customer_excel(file_path)
                     msg = f"Success: Updated {updated} work orders"

                else:
                     raise ValueError(f"未知的上傳類型：'{upload_type}' (len={len(upload_type)}) - 請檢查 models.py 或資料庫設定")

                # Update Config Status
                config.last_run = timezone.now()
                config.last_status = msg
                config.last_processed_mtime = current_mtime # Update timestamp
                config.save()
                
                self.stdout.write(self.style.SUCCESS(msg))

            except Exception as e:
                error_msg = f"Error: {str(e)}"
                self.stdout.write(self.style.ERROR(error_msg))
                config.last_status = error_msg
                config.save()
