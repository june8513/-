import os
import django
from decimal import Decimal

os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'material_requisition_system.settings')
django.setup()

from requisitions.models import WorkOrderMaterial

def debug_order_sap(order_number):
    woms = WorkOrderMaterial.objects.filter(order_number=order_number, is_active=True, sap_withdrawn_quantity__gt=0)
    print(f"Materials with SAP withdrawn quantity > 0: {woms.count()}")
    for wom in woms[:10]:
        print(f"Mat: {wom.material_number}, Req: {wom.required_quantity}, SAP Withdrawn: {wom.sap_withdrawn_quantity}, System Conf: {wom.confirmed_quantity}")

if __name__ == "__main__":
    import sys
    order = sys.argv[1] if len(sys.argv) > 1 else '100003316'
    debug_order_sap(order)
