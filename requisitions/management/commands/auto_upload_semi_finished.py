import os
from django.core.management.base import BaseCommand, CommandError
from django.conf import settings
from requisitions.utils import process_semi_finished_excel

class Command(BaseCommand):
    help = 'Automatically uploads semi-finished goods data from a specified Excel file.'

    def add_arguments(self, parser):
        parser.add_argument('--path', type=str, help='The path to the Excel file to upload.')

    def handle(self, *args, **options):
        excel_file_path = options['path']

        if not os.path.exists(excel_file_path):
            raise CommandError(f'File "{excel_file_path}" does not exist.')

        self.stdout.write(self.style.SUCCESS(f'Attempting to upload semi-finished data from {excel_file_path}...'))

        try:
            result = process_semi_finished_excel(excel_file_path)
            self.stdout.write(self.style.SUCCESS(
                f"半成品資料同步成功！\n"
                f"新增: {result['created_count']}\n"
                f"更新: {result['updated_count']}\n"
                f"停用: {result['deactivated_count']}"
            ))

        except Exception as e:
            raise CommandError(f"上傳檔案時發生錯誤: {e}")
