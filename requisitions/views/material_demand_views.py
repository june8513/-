from django.shortcuts import render, redirect, get_object_or_404
from django.contrib.auth.decorators import login_required
from django.contrib import messages
from django.utils import timezone
from ..forms import MaterialDetailsUploadForm
from ..models import Requisition, RequisitionItem, MaterialListVersion, WorkOrderMaterial, Inventory, MachineModel, ProcessType
from inventory.models import Material
from django.db import transaction
from django.conf import settings
from django.contrib.auth.models import User, Group
from django.db.models import Q, F, Value, DecimalField, OuterRef, Subquery, Exists, ExpressionWrapper, Sum, Count
from django.db import models
from django.db.models.functions import Coalesce
from django.core.paginator import Paginator, EmptyPage, PageNotAnInteger
from django.urls import reverse
from django.db import IntegrityError
import pandas as pd
from django.http import JsonResponse, HttpResponse
import json
from decimal import Decimal, InvalidOperation
import datetime

@login_required
def update_estimated_arrival_date(request):
    if request.method == 'POST':
        try:
            data = json.loads(request.body)
            print(data)
            pk = data.get('pk')
            estimated_arrival_date = data.get('estimated_arrival_date')
            if estimated_arrival_date == '':
                estimated_arrival_date = None

            if not pk:
                return JsonResponse({'success': False, 'message': '未提供物料 ID'}, status=400)

            updated_rows = WorkOrderMaterial.objects.filter(pk=pk).update(estimated_arrival_date=estimated_arrival_date)
            print(f"Updated {updated_rows} rows.")
            
            return JsonResponse({'success': True, 'message': '預計入料日期已更新'})

        except Exception as e:
            return JsonResponse({'success': False, 'message': str(e)}, status=500)

    return JsonResponse({'success': False, 'message': '無效的請求'}, status=400)

@login_required
def estimated_material_demand(request):
    if not request.user.is_superuser and not request.user.groups.filter(name='撥料人員').exists():
        messages.error(request, "您沒有權限訪問此頁面。")
        return redirect('homepage')

    material_number_filter = request.GET.get('material_number')
    shortage_date_filter = request.GET.get('shortage_date')
    purchaser_filter = request.GET.get('purchaser')
    sort_by = request.GET.get('sort_by', 'first_demand_date') # Default sort by first_demand_date
    order = request.GET.get('order', 'asc') # Default order ascending

    # Subquery to identify WorkOrderMaterial objects linked to dispatched or archived Requisitions
    dispatched_or_archived_wom_pks = WorkOrderMaterial.objects.filter(
        requisition_items__material_list_version__requisition__dispatch_performed=True
    ).values_list('pk', flat=True).distinct()
    
    archived_wom_pks = WorkOrderMaterial.objects.filter(
        requisition_items__material_list_version__requisition__is_archived=True
    ).values_list('pk', flat=True).distinct()

    # Combine the PKs to exclude
    pks_to_exclude = list(dispatched_or_archived_wom_pks) + list(archived_wom_pks)

    # Base queryset for WorkOrderMaterial
    demand_qs = WorkOrderMaterial.objects.filter(
        is_active=True
    ).exclude(
        material_number='PARENT_SCOPE'
    ).exclude(
        pk__in=pks_to_exclude # Exclude materials linked to dispatched/archived requisitions
    ).annotate(
        remaining_required_quantity=ExpressionWrapper(
            F('required_quantity') - Coalesce(F('confirmed_quantity'), Decimal('0.00')),
            output_field=DecimalField()
        ),
        current_stock=Subquery(
            Inventory.objects.filter(
                material_number=OuterRef('material_number')
            ).annotate(
                coalesced_stock=Coalesce(F('stock_quantity'), Decimal('0.00'))
            ).values('coalesced_stock')[:1],
            output_field=DecimalField()
        ),
        purchaser_username=Subquery(
            Material.objects.filter(
                material_code=OuterRef('material_number')
            ).values('purchaser__username')[:1],
            output_field=models.CharField()
        )
            ).filter(
                remaining_required_quantity__gt=0
            )

    # Apply purchaser filter
    if purchaser_filter:
        demand_qs = demand_qs.filter(purchaser_username=purchaser_filter)

    # Apply material number filter
    if material_number_filter:
        demand_qs = demand_qs.filter(material_number__icontains=material_number_filter)

    # Aggregation
    final_aggregated_data = {}
    for item in demand_qs.order_by('demand_date', 'material_number', 'machine_model__name').values(
        'pk', 'demand_date', 'material_number', 'item_name', 'machine_model__name',
        'process_type__name', 'remaining_required_quantity', 'order_number', 'current_stock', 'estimated_arrival_date', 'purchaser_username'
    ):
        material_key = (item['material_number'], item['item_name'])
        if material_key not in final_aggregated_data:
            final_aggregated_data[material_key] = {
                'pk': item['pk'],
                'material_number': item['material_number'],
                'item_name': item['item_name'],
                'current_stock': item['current_stock'] or Decimal('0.00'),
                'first_demand_date': item['demand_date'],
                'demanding_orders': set(),
                'total_required_quantity': Decimal('0.00'),
                'detail_orders': [],
                'estimated_arrival_date': item['estimated_arrival_date'],
                'purchaser': {'username': item['purchaser_username']},
            }
        
        if item['demand_date'] and (final_aggregated_data[material_key]['first_demand_date'] is None or item['demand_date'] < final_aggregated_data[material_key]['first_demand_date']):
            final_aggregated_data[material_key]['first_demand_date'] = item['demand_date']
        
        final_aggregated_data[material_key]['demanding_orders'].add(item['order_number'])
        final_aggregated_data[material_key]['total_required_quantity'] += item['remaining_required_quantity']

        final_aggregated_data[material_key]['detail_orders'].append({
            'order_number': item['order_number'],
            'demand_date': str(item['demand_date']),
            'required_quantity': str(item['remaining_required_quantity']),
            'machine_model_name': item['machine_model__name'],
            'process_type_name': item['process_type__name'],
        })

    # Calculate final_shortage after aggregation
    for material_key, data in final_aggregated_data.items():
        data['final_shortage'] = data['total_required_quantity'] - data['current_stock']
        if data['final_shortage'] < 0:
            data['final_shortage'] = Decimal('0.00')

    # Convert to list, sort, and add demanding_orders_count
    if sort_by == 'estimated_arrival_date':
        demand_list_sorted = sorted(
            final_aggregated_data.values(),
            key=lambda x: (x['estimated_arrival_date'] if x['estimated_arrival_date'] else datetime.date.max),
            reverse=(order == 'desc')
        )
    elif sort_by == 'material_number':
        demand_list_sorted = sorted(
            final_aggregated_data.values(),
            key=lambda x: x['material_number'],
            reverse=(order == 'desc')
        )
    # Add other sorting options here if needed
    else: # Default sort by first_demand_date
        demand_list_sorted = sorted(
            final_aggregated_data.values(),
            key=lambda x: (x['first_demand_date'] if x['first_demand_date'] else datetime.date.max, x['material_number']),
            reverse=(order == 'desc')
        )
    for item in demand_list_sorted:
        item['demanding_orders_count'] = len(item['demanding_orders'])
        del item['demanding_orders']

        # Calculate running stock and shortage for detail_orders
        running_stock = item['current_stock']
        item['shortage_date'] = None
        # Sort detail_orders by demand_date before calculating running stock
        item['detail_orders'].sort(key=lambda x: (datetime.datetime.strptime(x['demand_date'], '%Y-%m-%d').date() if x['demand_date'] != 'None' else datetime.date.max))
        for detail_order in item['detail_orders']:
            required_qty = Decimal(detail_order['required_quantity'])
            running_stock -= required_qty
            detail_order['running_stock'] = str(running_stock)
            detail_order['is_running_shortage'] = running_stock < 0

            if item['shortage_date'] is None and running_stock < 0:
                item['shortage_date'] = detail_order['demand_date']

        # Explicitly convert non-JSON-serializable types to strings for detail_orders
        for detail_order in item['detail_orders']:
            if isinstance(detail_order.get('demand_date'), datetime.date):
                detail_order['demand_date'] = str(detail_order['demand_date'])
        
        if isinstance(item.get('current_stock'), Decimal):
            item['current_stock'] = str(item['current_stock'])
        if isinstance(item.get('first_demand_date'), datetime.date):
            item['first_demand_date'] = str(item['first_demand_date'])
        if isinstance(item.get('estimated_arrival_date'), datetime.date):
            item['estimated_arrival_date'] = item['estimated_arrival_date'].strftime('%Y-%m-%d')

        item['detail_orders_json'] = json.dumps(item['detail_orders'])

    # Apply shortage date filter
    if shortage_date_filter:
        try:
            shortage_date = datetime.datetime.strptime(shortage_date_filter, '%Y-%m-%d').date()
            demand_list_sorted = [
                item for item in demand_list_sorted
                if item.get('shortage_date') and datetime.datetime.strptime(item['shortage_date'], '%Y-%m-%d').date() <= shortage_date
            ]
        except (ValueError, TypeError):
            messages.error(request, "無效的日期格式，請使用 YYYY-MM-DD。")


    # Pagination
    paginator = Paginator(demand_list_sorted, 20)
    page_number = request.GET.get('page')
    page_obj = paginator.get_page(page_number)

    # Get all unique purchasers for the filter dropdown
    all_purchasers = User.objects.filter(
        id__in=Material.objects.values_list('purchaser__id', flat=True).distinct()
    ).order_by('username')
    purchaser_choices = [(p.username, p.username) for p in all_purchasers if p.username]

    context = {
        'demand_list': page_obj,
        'material_number_filter': material_number_filter,
        'shortage_date_filter': shortage_date_filter,
        'purchaser_filter': purchaser_filter,
        'purchaser_choices': purchaser_choices,
        'sort_by': sort_by,
        'order': order,
    }
    return render(request, 'requisitions/estimated_material_demand.html', context)


@login_required
def shortage_materials_list(request):
    if not request.user.is_superuser and not request.user.groups.filter(name='撥料人員').exists():
        messages.error(request, "您沒有權限查看此頁面。")
        return redirect('homepage')

    # Get order_number and process_type pairs from dispatched requisitions
    dispatched_requisition_pairs = Requisition.objects.filter(
        dispatch_performed=True
    ).values_list('order_number', 'process_type')

    # Build a Q object to filter WorkOrderMaterial based on these pairs
    q_objects = Q()
    if not dispatched_requisition_pairs:
        shortage_materials_qs = WorkOrderMaterial.objects.none()
    else:
        for order_num, proc_type in dispatched_requisition_pairs:
            q_objects |= Q(order_number=order_num, process_type__name=proc_type)

        # Filter for active, backordered materials associated with dispatched requisitions
        shortage_materials_qs = WorkOrderMaterial.objects.filter(
            q_objects,
            is_active=True,
            required_quantity__gt=F('confirmed_quantity')
        ).exclude(material_number='PARENT_SCOPE').order_by('material_number', 'pk')

    # Aggregate in Python
    aggregated_shortages = {}
    for material in shortage_materials_qs:
        key = material.material_number
        if key not in aggregated_shortages:
            aggregated_shortages[key] = {
                'material_number': material.material_number,
                'item_name': material.item_name,
                'total_shortage': Decimal('0.00'),
                'orders': set(),
                'estimated_arrival_date': material.estimated_arrival_date
            }
        shortage = material.required_quantity - (material.confirmed_quantity or 0)
        if shortage > 0:
            aggregated_shortages[key]['total_shortage'] += shortage
            aggregated_shortages[key]['orders'].add(material.order_number)

    # Convert to list and format orders_str
    summarized_shortages = list(aggregated_shortages.values())
    for summary in summarized_shortages:
        summary['orders_str'] = ", ".join(sorted(list(summary['orders'])))

    context = {
        'shortage_materials': summarized_shortages,
    }
    return render(request, 'requisitions/shortage_materials_list.html', context)

@login_required
def update_shortage_arrival_dates(request):
    if request.method != 'POST':
        return redirect('shortage_materials_list')

    if not request.user.is_superuser and not request.user.groups.filter(name='撥料人員').exists():
        messages.error(request, "您沒有權限執行此操作。")
        return redirect('shortage_materials_list')

    # Get the same context for filtering as in the list view
    dispatched_requisition_pairs = Requisition.objects.filter(
        dispatch_performed=True
    ).values_list('order_number', 'process_type')

    q_objects = Q()
    if dispatched_requisition_pairs:
        for order_num, proc_type in dispatched_requisition_pairs:
            q_objects |= Q(order_number=order_num, process_type__name=proc_type)

    total_updated_count = 0
    updated_materials_count = 0

    for key, value in request.POST.items():
        is_desktop = key.startswith('arrival_date_desktop_')
        is_mobile = key.startswith('arrival_date_mobile_')

        if is_desktop or is_mobile:
            if is_desktop:
                material_number = key.replace('arrival_date_desktop_', '')
            else:
                material_number = key.replace('arrival_date_mobile_', '')
            
            arrival_date = value if value else None

            try:
                # Apply the full filter to get the exact set of materials that are on the shortage list
                shortage_materials_to_update = WorkOrderMaterial.objects.filter(
                    q_objects,
                    material_number=material_number,
                    is_active=True,
                    required_quantity__gt=F('confirmed_quantity')
                ).exclude(material_number='PARENT_SCOPE')

                    # We now allow clearing the date, so the check for `if value:` is removed.
                if not shortage_materials_to_update.exists():
                    continue

                updated_count = shortage_materials_to_update.update(estimated_arrival_date=arrival_date)
                if updated_count > 0:
                    total_updated_count += updated_count
                    updated_materials_count += 1

            except Exception as e:
                messages.error(request, f"更新物料 {material_number} 時發生錯誤: {e}")

    if updated_materials_count > 0:
        messages.success(request, f"成功更新 {updated_materials_count} 種物料，共影響 {total_updated_count} 筆欠料記錄。")

    return redirect('shortage_materials_list')
