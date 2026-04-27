import os
import django
from decimal import Decimal

# Configure Django
os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'material_requisition_system.settings')
django.setup()

from requisitions.models import RequisitionItem

def cleanup_orphaned_requisition_items():
    print("Starting cleanup of RequisitionItems with inactive source materials...")
    
    orphaned_items = RequisitionItem.objects.filter(source_material__is_active=False)
    total_count = orphaned_items.count()
    print(f"Found {total_count} orphaned RequisitionItems.")
    
    items_to_save = []
    ids_to_delete = []
    
    for item in orphaned_items:
        confirmed = item.confirmed_quantity or Decimal('0')
        if confirmed > 0:
            # 有撥料：設為待退料
            item.required_quantity = Decimal('0')
            item.alert_dismissed = False
            if "(已刪除)" not in item.item_name:
                item.item_name += " (已刪除)"
            item.dispatch_status = 'dispatched'
            items_to_save.append(item)
        else:
            # 無撥料：直接刪除
            ids_to_delete.append(item.id)
            
    if items_to_save:
        RequisitionItem.objects.bulk_update(items_to_save, ['required_quantity', 'alert_dismissed', 'item_name', 'dispatch_status'], batch_size=2000)
        print(f"Updated {len(items_to_save)} items to '(已刪除)' state.")
        
    if ids_to_delete:
        deleted_count, _ = RequisitionItem.objects.filter(id__in=ids_to_delete).delete()
        print(f"Deleted {deleted_count} items with 0 confirmed quantity.")

if __name__ == "__main__":
    cleanup_orphaned_requisition_items()
