import os
from django.core.management.base import BaseCommand, CommandError
from django.conf import settings
from requisitions.utils import process_shipping_customer_excel

class Command(BaseCommand):
    help = '自動上傳出貨客戶資料 (出貨日期、客戶名稱) 從指定的 Excel 檔案。'

    def add_arguments(self, parser):
        parser.add_argument('--path', type=str, help='The path to the Excel file to upload.')

    def handle(self, *args, **options):
        excel_file_path = options['path']

        if not os.path.exists(excel_file_path):
            raise CommandError(f'File "{excel_file_path}" does not exist.')

        self.stdout.write(self.style.SUCCESS(f'Attempting to upload shipping/customer data from {excel_file_path}...'))

        try:
            updated_count = process_shipping_customer_excel(excel_file_path)
            self.stdout.write(self.style.SUCCESS(f"出貨客戶資料同步成功！更新 {updated_count} 筆工單。"))

        except Exception as e:
            raise CommandError(f"上傳檔案時發生錯誤: {e}")
