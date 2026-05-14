
import os
import sys
import django

from django.utils import timezone

# Setup Django environment
sys.path.append('/home/june/material-requisition')
os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'material_requisition_system.settings.production')
django.setup()

from requisitions.models import AutoUploadConfig

def update_configs():
    base_path = '/home/june/material-requisition/Newdata'
    
    # Define mapping: upload_type -> filename
    config_mapping = {
        'inventory': '零件庫存.XLSX',
        'order_model': '成品入庫TECO狀態.XLSX', # Screenshot says this
        'material_details': '成品撥料.XLSX',
        'semi_finished': '撥料.XLSX',
        'semi_finished_model_db': '半品未撥.XLSX', 
    }

    for upload_type, filename in config_mapping.items():
        file_path = os.path.join(base_path, filename)
        
        # Check if file exists
        if not os.path.exists(file_path):
            print(f"WARNING: File not found: {file_path}")
        else:
            print(f"File found: {file_path}")

        # specific for order_model, screenshot showed two TECO files, user might want '入庫TECO狀態.XLSX' instead?
        # Screenshot says: `/home/june/material-requisition/Newdata/成品入庫TECO狀態.XLSX` for Order Models. So I'll stick with that.
        
        config, created = AutoUploadConfig.objects.get_or_create(
            upload_type=upload_type,
            defaults={'file_path': file_path, 'is_active': True}
        )
        
        if not created:
            config.file_path = file_path
            config.is_active = True
            config.save()
            print(f"Updated config for {upload_type} to {file_path}")
        else:
            print(f"Created config for {upload_type} to {file_path}")

if __name__ == '__main__':
    update_configs()
