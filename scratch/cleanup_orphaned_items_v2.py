import os
import django
from decimal import Decimal

# Configure Django
os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'material_requisition_system.settings')
django.setup()

from requisitions.models import RequisitionItem, WorkOrderMaterial

def cleanup_orphaned_data():
    print("Starting cleanup of orphaned materials and requisition items...")
    
    # 1. Fix WorkOrderMaterial required_quantity for inactive materials
    inactive_woms = WorkOrderMaterial.objects.filter(is_active=False).exclude(required_quantity=0)
    wom_count = inactive_woms.count()
    print(f"Found {wom_count} inactive materials with non-zero requirement.")
    
    if wom_count > 0:
        for wom in inactive_woms:
            wom.required_quantity = Decimal('0')
        WorkOrderMaterial.objects.bulk_update(inactive_woms, ['required_quantity'], batch_size=2000)
        print(f"Updated {wom_count} materials.")

    # 2. Fix RequisitionItems
    orphaned_items = RequisitionItem.objects.filter(source_material__is_active=False)
    # We should also handle items that might have been modified but not fully synced
    # and items whose source_material might be None but should be deleted? (unlikely)
    
    total_count = orphaned_items.count()
    print(f"Found {total_count} orphaned RequisitionItems.")
    
    items_to_save = []
    ids_to_delete = []
    
    for item in orphaned_items:
        confirmed = item.confirmed_quantity or Decimal('0')
        required = item.required_quantity or Decimal('0')
        
        # If it's orphaned and still has requirement > 0 or doesn't have "(已刪除)"
        if confirmed > 0:
            needs_update = False
            if required != 0:
                item.required_quantity = Decimal('0')
                needs_update = True
            if "(已刪除)" not in (item.item_name or ""):
                item.item_name = (item.item_name or "") + " (已刪除)"
                needs_update = True
            
            if needs_update:
                item.alert_dismissed = False
                item.dispatch_status = 'dispatched'
                items_to_save.append(item)
        else:
            ids_to_delete.append(item.id)
            
    if items_to_save:
        RequisitionItem.objects.bulk_update(items_to_save, ['required_quantity', 'alert_dismissed', 'item_name', 'dispatch_status'], batch_size=2000)
        print(f"Updated {len(items_to_save)} items to '(已刪除)' state.")
        
    if ids_to_delete:
        deleted_count, _ = RequisitionItem.objects.filter(id__in=ids_to_delete).delete()
        print(f"Deleted {deleted_count} items with 0 confirmed quantity.")

if __name__ == "__main__":
    cleanup_orphaned_data()
