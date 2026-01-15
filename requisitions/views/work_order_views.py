from django.shortcuts import render, redirect, get_object_or_404
from django.contrib.auth.decorators import login_required
from django.contrib import messages
from ..models import WorkOrder, WorkOrderMaterial, ProcessType, Requisition
from django.db.models import Sum, Value, DecimalField, OuterRef, Subquery, F
from django.db.models.functions import Coalesce
from django.urls import reverse
from django.db import transaction

@login_required
def work_order_list(request):
    sort_by = request.GET.get('sort', '-updated_at')
    direction = request.GET.get('direction', 'desc')

    # Validate sort_by parameter
    valid_sort_fields = ['order_number', 'shipping_date', 'status_message', 'machine_model_name', '-updated_at']
    if sort_by.lstrip('-') not in valid_sort_fields:
        sort_by = '-updated_at'

    order = f"{'-' if direction == 'desc' else ''}{sort_by.lstrip('-')}"

    # Annotate machine model name for sorting
    machine_model_subquery = WorkOrderMaterial.objects.filter(
        order_number=OuterRef('order_number')
    ).exclude(material_number="PARENT_SCOPE").values('machine_model__name')[:1]

    active_work_orders = WorkOrder.objects.filter(is_archived=False).annotate(
        machine_model_name=Subquery(machine_model_subquery)
    ).order_by(order)

    archived_work_orders = WorkOrder.objects.filter(is_archived=True).order_by('-updated_at')
    
    work_orders_data = []
    for wo in active_work_orders:
        materials = WorkOrderMaterial.objects.filter(order_number=wo.order_number).exclude(material_number="PARENT_SCOPE")
        
        # We already have the machine model name from the annotation
        machine_model_name = wo.machine_model_name if hasattr(wo, 'machine_model_name') else 'N/A'
        
        process_type_progress = []
        process_type_ids = materials.values_list('process_type_id', flat=True).distinct()
        process_types = ProcessType.objects.filter(id__in=process_type_ids)

        for pt in process_types:
            pt_materials = materials.filter(process_type=pt)
            total_required = pt_materials.aggregate(total=Coalesce(Sum('required_quantity'), Value(0), output_field=DecimalField()))['total']
            total_confirmed = pt_materials.aggregate(total=Coalesce(Sum('confirmed_quantity'), Value(0), output_field=DecimalField()))['total']
            
            progress = 0
            if total_required > 0:
                progress = (total_confirmed / total_required) * 100

            # Find the corresponding requisition for this process type
            requisition = Requisition.objects.filter(order_number=wo.order_number, process_type=pt.name).first()
            
            process_type_progress.append({
                'id': pt.id,
                'name': pt.name,
                'progress': round(progress, 2),
                'requisition_pk': requisition.pk if requisition else None
            })
            
        work_orders_data.append({
            'work_order': wo,
            'machine_model': machine_model_name,
            'process_type_progress': process_type_progress,
        })

    context = {
        'work_orders_data': work_orders_data,
        'archived_work_orders': archived_work_orders,
        'current_sort': sort_by.lstrip('-'),
        'current_direction': direction,
    }
    return render(request, 'requisitions/work_order_list.html', context)

@login_required
def work_order_requisitions_list(request, order_number):
    requisitions = Requisition.objects.filter(order_number=order_number).order_by('process_type')
    context = {
        'requisitions': requisitions,
        'order_number': order_number,
    }
    return render(request, 'requisitions/work_order_requisitions_list.html', context)

@login_required
def toggle_work_order_archive(request, order_number):
    if request.method == 'POST':
        with transaction.atomic():
            work_order = get_object_or_404(WorkOrder, order_number=order_number)
            new_archived_status = not work_order.is_archived
            
            # Update the WorkOrder
            work_order.is_archived = new_archived_status
            work_order.save()
            
            # Update all associated Requisitions
            Requisition.objects.filter(order_number=order_number).update(is_archived=new_archived_status)
            
            if new_archived_status:
                messages.success(request, f"工單 {order_number} 及其所有相關申請單已歸檔。")
            else:
                messages.success(request, f"工單 {order_number} 及其所有相關申請單已取消歸檔。")
            
    return redirect(reverse('requisitions:work_order_list') + '#work-order-list-table')