import os
from django.core.management.base import BaseCommand, CommandError
from requisitions.utils import process_inventory_excel

class Command(BaseCommand):
    help = 'Automatically uploads inventory data from a specified Excel file.'

    def add_arguments(self, parser):
        parser.add_argument('--path', type=str, required=True, help='The path to the Excel file to upload.')

    def handle(self, *args, **options):
        excel_file_path = options['path']

        if not os.path.exists(excel_file_path):
            raise CommandError(f'File "{excel_file_path}" does not exist.')

        self.stdout.write(self.style.SUCCESS(f'Attempting to upload inventory from {excel_file_path}...'))

        try:
            created_count, updated_count = process_inventory_excel(excel_file_path)
            self.stdout.write(self.style.SUCCESS(
                f"庫存資料同步成功！新增 {created_count} 筆，更新 {updated_count} 筆。"
            ))

        except Exception as e:
            import traceback
            tb_str = traceback.format_exc()
            raise CommandError(f"上傳檔案時發生錯誤: {e}\n{tb_str}")
