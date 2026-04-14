import os
import django

os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'material_requisition_system.settings')
django.setup()

from requisitions.models import Requisition, RequisitionItem
from django.db.models import Q

# 嘗試找出使用者截圖中的那一張單 (訂單號 100003346, 投料點 電裝)
order_number = '100003346'
process_type = '電裝'

req = Requisition.objects.filter(order_number=order_number, process_type=process_type).first()

if req:
    print(f"Found Requisition: PK={req.pk}, Order={req.order_number}, Process={req.process_type}")
    items = RequisitionItem.objects.filter(requisition=req)
    print(f"Items count: {items.count()}")
    for item in items:
        print(f"  - Item: {item.material_number}, {item.item_name}")
else:
    print("Requisition not found with exact match, searching all...")
    all_reqs = Requisition.objects.filter(order_number__icontains='100003346')
    for r in all_reqs:
        print(f"Found: PK={r.pk}, Order={r.order_number}, Process={r.process_type}")
