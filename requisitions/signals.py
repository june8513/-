from django.db.models.signals import post_delete
from django.dispatch import receiver
from .models import RequisitionItem, WorkOrderMaterial
from decimal import Decimal

@receiver(post_delete, sender=RequisitionItem)
def update_work_order_material_on_requisition_item_delete(sender, instance, **kwargs):
    """
    When a RequisitionItem is deleted, adjust the confirmed_quantity of its
    associated WorkOrderMaterial.
    """
    if instance.source_material and instance.confirmed_quantity is not None:
        work_order_material = instance.source_material
        # Ensure confirmed_quantity doesn't go below zero
        work_order_material.confirmed_quantity = max(Decimal('0'), work_order_material.confirmed_quantity - instance.confirmed_quantity)
        work_order_material.save()
        print(f"DEBUG: RequisitionItem (PK: {instance.pk}) deleted. Adjusted WorkOrderMaterial (PK: {work_order_material.pk}) confirmed_quantity to {work_order_material.confirmed_quantity}")
