import os
import sys
import django

os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'material_requisition_system.settings')
sys.path.append(os.getcwd())
django.setup()

from requisitions.models import Requisition, RequisitionItem, WorkOrderMaterial
from inventory.models import Material as InvMaterial
from decimal import Decimal

def fix_order_6722():
    order_number = '200006722'
    req = Requisition.objects.filter(order_number=order_number, requisition_type='semi_finished').first()
    if not req:
        print(f"No semi-finished requisition found for order {order_number}")
        return

    print(f"Processing Requisition ID: {req.id}")
    
    # 1. Handle 3209000336C0 (The new one)
    target_mat_number = '3209000336C0'
    wom_c0 = WorkOrderMaterial.objects.filter(order_number=order_number, material_number=target_mat_number, is_active=True).first()
    
    if wom_c0:
        item_exists = RequisitionItem.objects.filter(requisition=req, material_number=target_mat_number).exists()
        if not item_exists:
            print(f"Adding {target_mat_number} to requisition...")
            # Get inventory info
            inv = InvMaterial.objects.filter(material_code=target_mat_number).first()
            RequisitionItem.objects.create(
                requisition=req,
                source_material=wom_c0,
                order_number=order_number,
                material_number=target_mat_number,
                item_name=wom_c0.item_name,
                required_quantity=wom_c0.required_quantity,
                stock_quantity=inv.system_quantity if inv else Decimal('0'),
                storage_bin=inv.bin if inv else ''
            )
            print(f"Successfully added {target_mat_number}")
        else:
            print(f"{target_mat_number} already exists in requisition.")
    else:
        print(f"Active WorkOrderMaterial for {target_mat_number} not found!")

    # 2. Handle 3209000336B0 (The old one)
    old_mat_number = '3209000336B0'
    old_item = RequisitionItem.objects.filter(requisition=req, material_number=old_mat_number).first()
    if old_item:
        if not old_item.dispatch_status:
            print(f"Removing inactive item {old_mat_number} (undispatched)...")
            old_item.delete()
            print(f"Successfully removed {old_mat_number}")
        else:
            print(f"Inactive item {old_mat_number} has status {old_item.dispatch_status}, keeping it for history.")

if __name__ == "__main__":
    fix_order_6722()
