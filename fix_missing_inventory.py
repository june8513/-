import os
import sys
import django
from decimal import Decimal

# Set up Django environment
os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'material_requisition_system.settings')
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
django.setup()

from requisitions.models import RequisitionItem, Inventory

def fix_missing_inventory():
    print("Checking for missing inventory records...")
    items = RequisitionItem.objects.all()
    created_count = 0
    
    for item in items:
        material_number = item.material_number
        if not material_number:
            continue
            
        inventory_item, created = Inventory.objects.get_or_create(
            material_number=material_number,
            defaults={
                'stock_quantity': Decimal('1000.00'),
                'storage_bin': 'TEST-BIN'
            }
        )
        
        if created:
            print(f"Created missing inventory for material: {material_number}")
            created_count += 1
            
    print(f"Finished. Created {created_count} missing inventory records.")

if __name__ == "__main__":
    fix_missing_inventory()
