import os
from django.core.management.base import BaseCommand, CommandError
from requisitions.utils import process_material_details_excel

class Command(BaseCommand):
    help = 'Automatically uploads material details data from a specified Excel file.'

    def add_arguments(self, parser):
        parser.add_argument('--path', type=str, required=True, help='The path to the Excel file to upload.')
        parser.add_argument('--qty-col', type=str, required=False, default=None, help='The name of the required quantity column. If not provided, auto-detection will be attempted.')

    def handle(self, *args, **options):
        excel_file_path = options['path']
        required_qty_col = options['qty_col']

        if not os.path.exists(excel_file_path):
            raise CommandError(f'File "{excel_file_path}" does not exist.')

        self.stdout.write(self.style.SUCCESS(f'Attempting to upload material details from {excel_file_path}...'))

        try:
            result = process_material_details_excel(excel_file_path, required_qty_col)
            
            created_count = result.get('created_count', 0)
            updated_count = result.get('updated_count', 0)
            deleted_count = result.get('deactivated_count', 0)
            
            self.stdout.write(self.style.SUCCESS(
                f"物料明細同步成功！新增 {created_count} 筆，更新 {updated_count} 筆，移除/停用 {deleted_count} 筆。"
            ))

        except Exception as e:
            import traceback
            tb_str = traceback.format_exc()
            raise CommandError(f"上傳檔案時發生錯誤: {e}\n{tb_str}")