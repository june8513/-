from django.db.models.signals import post_save, post_delete
from django.dispatch import receiver
from .models import Requisition, RequisitionItem, WorkOrderMaterial
from decimal import Decimal
from django.utils import timezone
from datetime import timedelta

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

@receiver(post_delete, sender=RequisitionItem)
def update_work_order_material_on_requisition_item_delete(sender, instance, **kwargs):
    """
    When a RequisitionItem is deleted, adjust the confirmed_quantity of its
    associated WorkOrderMaterial.
    """
    if instance.source_material and instance.confirmed_quantity is not None:
        work_order_material = instance.source_material
        # Ensure confirmed_quantity doesn't go below zero
        current_confirmed = work_order_material.confirmed_quantity if work_order_material.confirmed_quantity is not None else Decimal('0')
        work_order_material.confirmed_quantity = max(Decimal('0'), current_confirmed - instance.confirmed_quantity)
        work_order_material.save()
        print(f"DEBUG: RequisitionItem (PK: {instance.pk}) deleted. Adjusted WorkOrderMaterial (PK: {work_order_material.pk}) confirmed_quantity to {work_order_material.confirmed_quantity}")
