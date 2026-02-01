from django.core.management.base import BaseCommand
from requisitions.models import AutoUploadConfig
from requisitions.utils import process_order_model_excel
import os
import datetime
from django.utils import timezone

class Command(BaseCommand):
    help = 'Auto-upload semi-finished machine model database from configured path'

    def add_arguments(self, parser):
        parser.add_argument('--path', type=str, help='Override the file path to upload.')

    def handle(self, *args, **options):
        # 取得要處理的檔案路徑
        passed_path = options.get('path')
        
        config_name = 'semi_finished_model_db'
        config = AutoUploadConfig.objects.filter(upload_type=config_name, is_active=True).first()

        if not config and not passed_path:
            self.stdout.write(self.style.WARNING(f'No active configuration found for {config_name}'))
            return

        file_path = passed_path or config.file_path
        if not file_path or not os.path.exists(file_path):
            error_msg = f'File not found: {file_path}'
            self.stdout.write(self.style.ERROR(error_msg))
            if config:
                config.last_status = error_msg
                config.last_run = timezone.now()
                config.save()
            return

        # 如果是自動執行（沒有手動給 path），檢查檔案修改時間
        if not passed_path and config:
            current_mtime = os.path.getmtime(file_path)
            if config.last_processed_mtime and current_mtime <= config.last_processed_mtime:
                self.stdout.write(self.style.SUCCESS(f'File has not changed since last run. Skipping.'))
                config.last_run = timezone.now()
                config.last_status = 'Skipped (No changes)'
                config.save()
                return

        try:
            self.stdout.write(f'Processing file: {file_path}')
            created_count, updated_count = process_order_model_excel(file_path)
            
            success_msg = f'Success! Created: {created_count}, Updated: {updated_count}'
            self.stdout.write(self.style.SUCCESS(success_msg))
            
            config.last_status = success_msg
            config.last_processed_mtime = current_mtime
            config.last_run = timezone.now()
            config.save()

        except Exception as e:
            error_msg = f'Error: {str(e)}'
            self.stdout.write(self.style.ERROR(error_msg))
            config.last_status = error_msg
            config.last_run = timezone.now()
            config.save()
