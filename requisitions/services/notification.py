"""
缺料通知服務
"""
import requests
from django.utils import timezone
from django.conf import settings
from decimal import Decimal

def notify_requisition_shortages(requisition):
    """
    Sends a consolidated notification about ALL shortages for a given requisition.
    Triggered after batch allocation updates.
    """
    try:
        # Get all items for this requisition that have shortages
        # We fetch all items to calculate the full picture if needed, or just filter shortages
        # Let's send a list of items that are currently short
        items = requisition.items.all()
        
        shortage_list = []
        has_shortage = False
        
        for item in items:
            confirmed = item.confirmed_quantity or Decimal('0')
            shortage = item.required_quantity - confirmed
            
            if shortage > 0:
                has_shortage = True
                shortage_list.append({
                    "material_number": item.material_number,
                    "item_name": item.item_name,
                    "required_quantity": float(item.required_quantity),
                    "confirmed_quantity": float(confirmed),
                    "shortage_quantity": float(shortage),
                    "status": "OPEN"
                })

        # Overall status for the requisition
        overall_status = "OPEN" if has_shortage else "RESOLVED"

        payload = {
            "order_number": requisition.order_number,
            "process_type": requisition.process_type,
            "requisition_id": requisition.pk,
            "status": overall_status,
            "timestamp": timezone.now().isoformat(),
            "shortage_items": shortage_list
        }
        
        # Send Request
        external_url = getattr(settings, 'SHORTAGE_NOTIFICATION_URL', None)
        
        if external_url:
            try:
                response = requests.post(external_url, json=payload, timeout=5)
                response.raise_for_status()
                print(f"[Requisition Notification] Success! Sent {len(shortage_list)} items for Order {requisition.order_number}")
            except requests.RequestException as req_err:
                print(f"[Requisition Notification] Request failed: {req_err}")
        else:
            print(f"[Requisition Notification] URL not configured. Payload: {payload}")
            
    except Exception as e:
        print(f"[Requisition Notification] Error: {e}")
