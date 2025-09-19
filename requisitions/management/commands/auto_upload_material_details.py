import os
from django.core.management.base import BaseCommand, CommandError
from requisitions.utils import process_material_details_excel

class Command(BaseCommand):
    help = 'Automatically uploads material details data from a specified Excel file.'

    def add_arguments(self, parser):
        parser.add_argument('--path', type=str, required=True, help='The path to the Excel file to upload.')
        parser.add_argument('--qty-col', type=str, required=True, help='The name of the required quantity column.')

    def handle(self, *args, **options):
        excel_file_path = options['path']
        required_qty_col = options['qty_col']

        if not os.path.exists(excel_file_path):
            raise CommandError(f'File "{excel_file_path}" does not exist.')

        self.stdout.write(self.style.SUCCESS(f'Attempting to upload material details from {excel_file_path}...'))

        try:
            created_count, updated_count, deleted_count = process_material_details_excel(excel_file_path, required_qty_col)
            self.stdout.write(self.style.SUCCESS(
                f"物料明細同步成功！新增 {created_count} 筆，更新 {updated_count} 筆，刪除 {deleted_count} 筆。"
            ))

        except Exception as e:
            import traceback
            tb_str = traceback.format_exc()
            raise CommandError(f"上傳檔案時發生錯誤: {e}\n{tb_str}")