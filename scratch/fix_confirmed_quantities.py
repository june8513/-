import os
import django
from decimal import Decimal

# Configure Django
os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'material_requisition_system.settings')
django.setup()

from requisitions.models import WorkOrderMaterial, RequisitionItem
from django.db.models import Sum

def sync_all_confirmed_quantities():
    print("Starting synchronization of WorkOrderMaterial.confirmed_quantity...")
    
    # Get all WorkOrderMaterials
    woms = WorkOrderMaterial.objects.all()
    total_count = woms.count()
    updated_count = 0
    
    for i, wom in enumerate(woms):
        if i % 1000 == 0:
            print(f"Processing {i}/{total_count}...")
            
        # Calculate sum of items
        total_item_confirmed = RequisitionItem.objects.filter(
            source_material=wom
        ).aggregate(total=Sum('confirmed_quantity'))['total'] or Decimal('0')
        
        # We must also consider sap_withdrawn_quantity!
        # If there are no RequisitionItems, but SAP says it's withdrawn, we must honor SAP.
        sap_qty = wom.sap_withdrawn_quantity or Decimal('0')
        
        # The correct confirmed quantity is the maximum of the two.
        # Usually they are equal, but if there's no RequisitionItem, sap_qty is higher.
        correct_confirmed = max(total_item_confirmed, sap_qty)
        
        if wom.confirmed_quantity != correct_confirmed:
            wom.confirmed_quantity = correct_confirmed
            wom.save(update_fields=['confirmed_quantity'])
            updated_count += 1

    print(f"Synchronization complete. Updated {updated_count} records.")

if __name__ == "__main__":
    sync_all_confirmed_quantities()
