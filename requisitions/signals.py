from django.db.models.signals import post_save, post_delete
from django.dispatch import receiver
from .models import Requisition, RequisitionItem, WorkOrderMaterial
from decimal import Decimal
from django.utils import timezone
from datetime import timedelta
from django.db import models

@receiver(post_save, sender=Requisition)
def create_announcement_on_requisition_creation(sender, instance, created, **kwargs):
    """
    當有新申請單建立時，自動發佈公告。
    """
    if created:
        from .models import Announcement
        process_type_str = f'"{instance.process_type}"' if instance.process_type else ""
        content = f"系統有新申請單一筆 ({process_type_str}工單: {instance.order_number})"
        expires_at = timezone.now() + timedelta(days=1)
        Announcement.objects.create(
            content=content,
            is_system_generated=True,
            expires_at=expires_at
        )

@receiver(post_save, sender=RequisitionItem)
def update_work_order_material_on_requisition_item_save(sender, instance, **kwargs):
    """
    When a RequisitionItem is saved, recalculate the total confirmed_quantity
    for its associated WorkOrderMaterial.
    """
    if instance.source_material:
        work_order_material = instance.source_material
        # Sum all confirmed_quantity for this WorkOrderMaterial
        total_confirmed = RequisitionItem.objects.filter(
            source_material=work_order_material
        ).aggregate(total=models.Sum('confirmed_quantity'))['total'] or Decimal('0')
        
        sap_qty = work_order_material.sap_withdrawn_quantity or Decimal('0')
        correct_confirmed = max(total_confirmed, sap_qty)
        
        if work_order_material.confirmed_quantity != correct_confirmed:
            work_order_material.confirmed_quantity = correct_confirmed
            work_order_material.save(update_fields=['confirmed_quantity'])
            print(f"DEBUG: RequisitionItem (PK: {instance.pk}) saved. Updated WorkOrderMaterial (PK: {work_order_material.pk}) confirmed_quantity to {correct_confirmed}")

@receiver(post_delete, sender=RequisitionItem)
def update_work_order_material_on_requisition_item_delete(sender, instance, **kwargs):
    """
    When a RequisitionItem is deleted, adjust the confirmed_quantity of its
    associated WorkOrderMaterial.
    """
    if instance.source_material:
        work_order_material = instance.source_material
        # Recalculate total
        total_confirmed = RequisitionItem.objects.filter(
            source_material=work_order_material
        ).aggregate(total=models.Sum('confirmed_quantity'))['total'] or Decimal('0')
        
        sap_qty = work_order_material.sap_withdrawn_quantity or Decimal('0')
        correct_confirmed = max(total_confirmed, sap_qty)
        
        if work_order_material.confirmed_quantity != correct_confirmed:
            work_order_material.confirmed_quantity = correct_confirmed
            work_order_material.save(update_fields=['confirmed_quantity'])
            print(f"DEBUG: RequisitionItem (PK: {instance.pk}) deleted. Updated WorkOrderMaterial (PK: {work_order_material.pk}) confirmed_quantity to {correct_confirmed}")
