import os
import django
from decimal import Decimal
from django.db.models import Sum

# Configure Django
os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'material_requisition_system.settings')
django.setup()

from requisitions.models import WorkOrderMaterial, RequisitionItem, Inventory

def debug_order(order_number):
    print(f"--- Debugging Order {order_number} ---")
    
    req_items_with_qty = RequisitionItem.objects.filter(order_number=order_number, confirmed_quantity__gt=0)
    print(f"RequisitionItems with confirmed_quantity > 0: {req_items_with_qty.count()}")
    
    woms = WorkOrderMaterial.objects.filter(order_number=order_number, is_active=True)
    print(f"Active WorkOrderMaterials: {woms.count()}")
    
    shortages_found = 0
    for wom in woms:
        items = RequisitionItem.objects.filter(source_material=wom)
        total_item_confirmed = items.aggregate(Sum('confirmed_quantity'))['confirmed_quantity__sum'] or Decimal('0')
        
        # Check shortage logic (similar to analysis.py)
        remaining = wom.required_quantity - wom.confirmed_quantity
        
        if remaining > 0:
            shortages_found += 1
            if shortages_found <= 20:
                print(f"Shortage: {wom.material_number} - Req: {wom.required_quantity}, Conf: {wom.confirmed_quantity}, ItemsSum: {total_item_confirmed}")

if __name__ == "__main__":
    import sys
    order = sys.argv[1] if len(sys.argv) > 1 else '100003316'
    debug_order(order)
