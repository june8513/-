import os
import django
from django.db import transaction

# Set up Django environment
os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'material_requisition_system.settings')
django.setup()

from requisitions.models import Requisition, WorkOrderMaterial, WorkOrderMaterialTransaction, RequisitionItem, Inventory
from inventory.models import Material # Assuming Material model is in inventory app

print("Starting full DATABASE data cleanup script...")

print("********************************************************************************")
print("*** WARNING: This script will PERMANENTLY DELETE ALL DATA from the following tables: ***")
print("***          - Requisition (and all related RequisitionItem)                 ***")
print("***          - WorkOrderMaterial (and all related WorkOrderMaterialTransaction) ***")
print("***          - Inventory                                                    ***")
print("***          - Material                                                     ***")
print("********************************************************************************")
confirm = input("Are you absolutely sure you want to proceed? Type 'yes' to confirm: ")

if confirm.lower() != 'yes':
    print("Data cleanup cancelled by user.")
    exit()

try:
    with transaction.atomic():
        # Clear Requisition table (this will cascade delete RequisitionItems)
        num_requisitions_deleted, _ = Requisition.objects.all().delete()
        print(f"Step 1: Deleted {num_requisitions_deleted} Requisitions and associated RequisitionItems.")

        # Clear WorkOrderMaterial table
        num_work_order_materials_deleted, _ = WorkOrderMaterial.objects.all().delete()
        print(f"Step 2: Deleted {num_work_order_materials_deleted} WorkOrderMaterials.")

        # Clear WorkOrderMaterialTransaction table
        num_transactions_deleted, _ = WorkOrderMaterialTransaction.objects.all().delete()
        print(f"Step 3: Deleted {num_transactions_deleted} WorkOrderMaterialTransactions.")
        
        # Clear Inventory table
        num_inventory_deleted, _ = Inventory.objects.all().delete()
        print(f"Step 4: Deleted {num_inventory_deleted} Inventory records.")

        # Clear Material table (assuming this is part of the data you want to clear)
        num_materials_deleted, _ = Material.objects.all().delete()
        print(f"Step 5: Deleted {num_materials_deleted} Material records.")

    print("Full DATABASE data cleanup script completed successfully.")

except Exception as e:
    print(f"An error occurred during data cleanup: {e}")
    import traceback
    traceback.print_exc()
