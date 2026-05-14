"""
申請單變更提醒服務
"""
from django.utils import timezone
from requisitions.models import Requisition

def _update_requisition_alert(order_number, process_type_name, message, is_demand_increase=False):
    """
    Helper to update requisition alert status and potentially revert status if demand increases.
    """
    try:
        # Filter for the exact requisition
        # Note: There should be only one per order/process_type due to UniqueConstraint
        reqs = Requisition.objects.filter(order_number=order_number, process_type=process_type_name)
        
        for req in reqs:
            # Append message
            timestamp = timezone.localtime(timezone.now()).strftime('%Y-%m-%d %H:%M')
            new_msg = f"[{timestamp}] {message}"
            if req.alert_message:
                req.alert_message += f"\n{new_msg}"
            else:
                req.alert_message = new_msg
            
            req.has_alert = True
            
            # Revert status if demand increases (and it was completed/signed_off/archived)
            if is_demand_increase:
                if req.status in ['dispatch_completed', 'signed_off', 'archived'] or req.is_archived:
                    req.status = 'dispatch_in_progress'
                    req.is_archived = False # Un-archive
                    # We might want to keep dispatch_performed as True if partial dispatch is done?
                    # But strictly, if demand increased, dispatch is NOT fully performed.
                    req.dispatch_performed = False 
            
            req.save()
    except Exception as e:
        print(f"Error updating alert for {order_number}: {e}")

