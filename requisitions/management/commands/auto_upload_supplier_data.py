import os
from django.core.management.base import BaseCommand, CommandError
from requisitions.utils import process_supplier_data_excel

class Command(BaseCommand):
    help = 'Automatically updates supplier data from a specified Excel file based on latest document date.'

    def add_arguments(self, parser):
        parser.add_argument('--path', type=str, required=True, help='The path to the Excel file to upload.')

    def handle(self, *args, **options):
        excel_file_path = options['path']

        if not os.path.exists(excel_file_path):
            raise CommandError(f'File "{excel_file_path}" does not exist.')

        self.stdout.write(self.style.SUCCESS(f'Attempting to update supplier data from {excel_file_path}...'))

        try:
            updated_count = process_supplier_data_excel(excel_file_path)
            self.stdout.write(self.style.SUCCESS(
                f"供應商資料更新成功！已根據最新日期更新 {updated_count} 筆物料。"
            ))

        except Exception as e:
            import traceback
            tb_str = traceback.format_exc()
            raise CommandError(f"更新供應商檔案時發生錯誤: {e}\n{tb_str}")
