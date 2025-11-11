import os
import django
from decimal import Decimal
from django.db import transaction

# Set up Django environment
os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'material_requisition_system.settings')
django.setup()

from requisitions.models import Requisition, WorkOrderMaterial, WorkOrderMaterialTransaction, RequisitionItem

print("Starting full data cleanup script...")

try:
    with transaction.atomic():
        # 1. Reset confirmed_quantity for WorkOrderMaterial objects that have confirmed_quantity > 0
        # but are not associated with any RequisitionItem.
        materials_to_reset = WorkOrderMaterial.objects.filter(
            confirmed_quantity__gt=Decimal('0')
        ).exclude(
            requisition_items__isnull=False
        )

        print(f"Found {materials_to_reset.count()} WorkOrderMaterial objects with confirmed_quantity > 0 but no associated RequisitionItems.")

        for material in materials_to_reset:
            print(f"  Resetting confirmed_quantity for WorkOrderMaterial PK: {material.pk}, Material Number: {material.material_number}, Old Confirmed Qty: {material.confirmed_quantity}")
            material.confirmed_quantity = Decimal('0')
            material.save()
        
        print("Step 1: Resetting confirmed_quantity for WorkOrderMaterial objects without associated RequisitionItems complete.")

        # 2. Clear Requisition table (this will cascade delete RequisitionItems)
        num_requisitions_deleted, _ = Requisition.objects.all().delete()
        print(f"Step 2: Deleted {num_requisitions_deleted} Requisitions and associated RequisitionItems.")

        # 3. Clear WorkOrderMaterialTransaction table
        num_transactions_deleted, _ = WorkOrderMaterialTransaction.objects.all().delete()
        print(f"Step 3: Deleted {num_transactions_deleted} WorkOrderMaterialTransactions.")

    print("Full data cleanup script completed successfully.")

except Exception as e:
    print(f"An error occurred during data cleanup: {e}")
    import traceback
    traceback.print_exc()
