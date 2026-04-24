from django.shortcuts import render, redirect, get_object_or_404
from django.contrib.auth.decorators import login_required
from django.contrib import messages
from ..models import Requisition, WorkOrderMaterial, Inventory, RequisitionItem
from inventory.models import Material
from django.contrib.auth.models import User
from django.core.paginator import Paginator, EmptyPage
from django.http import JsonResponse
import json
from decimal import Decimal
import datetime
from django.db.models import Q, F, Sum, Max, Value, DecimalField, OuterRef, Subquery, ExpressionWrapper
from django.db.models.functions import Coalesce
from ..analysis import get_material_demand_analysis
from ..constants import GROUP_NAMES

@login_required
def update_estimated_arrival_date(request):
    if request.method == 'POST':
        try:
            data = json.loads(request.body)
            material_number = data.get('material_number')
            estimated_arrival_date_str = data.get('estimated_arrival_date')
            
            if not material_number:
                return JsonResponse({'success': False, 'message': '未提供物料編號'}, status=400)

            estimated_arrival_date = None
            if estimated_arrival_date_str:
                try:
                    estimated_arrival_date = datetime.datetime.strptime(estimated_arrival_date_str, '%Y-%m-%d').date()
                except ValueError:
                    return JsonResponse({'success': False, 'message': '無效的預計入料日期格式'}, status=400)

            print(f"DEBUG: material_number: {material_number}")
            print(f"DEBUG: estimated_arrival_date_str: {estimated_arrival_date_str}")
            print(f"DEBUG: Parsed estimated_arrival_date: {estimated_arrival_date}")

            # Retrieve shortage_date for the material
            analysis_data = get_material_demand_analysis()
            material_analysis = analysis_data.get(material_number)
            
            shortage_date = None
            if material_analysis and material_analysis.get('shortage_date'):
                # Ensure shortage_date is a date object for comparison
                if isinstance(material_analysis['shortage_date'], str):
                    shortage_date = datetime.datetime.strptime(material_analysis['shortage_date'], '%Y-%m-%d').date()
                else:
                    shortage_date = material_analysis['shortage_date']
            
            print(f"DEBUG: Retrieved shortage_date: {shortage_date}")

            # Apply business rule: if estimated_arrival_date is earlier than shortage_date, clear it
            if estimated_arrival_date and shortage_date and estimated_arrival_date < shortage_date:
                estimated_arrival_date = None
                message = '預計入料日期早於缺料日期，已自動清除。'
                print(f"DEBUG: Condition met: estimated_arrival_date < shortage_date. estimated_arrival_date set to None.")
            else:
                message = '預計入料日期已更新'
                print(f"DEBUG: Condition not met or dates valid. estimated_arrival_date: {estimated_arrival_date}")

            # Update all WorkOrderMaterial instances with this material_number
            updated_count = WorkOrderMaterial.objects.filter(material_number=material_number, is_active=True).update(estimated_arrival_date=estimated_arrival_date)
            print(f"DEBUG: Updated {updated_count} WorkOrderMaterial instances.")
            
            return JsonResponse({'success': True, 'message': message})

        except Exception as e:
            return JsonResponse({'success': False, 'message': str(e)}, status=500)

    return JsonResponse({'success': False, 'message': '無效的請求'}, status=400)

@login_required
def estimated_material_demand(request):
    is_dispatcher_supervisor = request.user.groups.filter(name=GROUP_NAMES['DISPATCHER_SUPERVISOR']).exists()
    if not request.user.is_superuser and not is_dispatcher_supervisor:
        messages.error(request, "您沒有權限訪問此頁面。")
        return redirect('core:homepage')

    # Get all aggregated data from the shared analysis function
    all_materials_analysis = get_material_demand_analysis()

    # Convert to list for filtering and sorting
    # Filter to only include materials with a shortage (Requirement #4)
    demand_list = [item for item in all_materials_analysis.values() if item['is_shortage']]

    # Add calculated fields and format dates for template
    for item in demand_list:
        # Calculate demanding_orders_count
        item['demanding_orders_count'] = len(item['detail_orders'])

        # Format dates for display in template
        if item.get('first_shortage_shipping_date'):
            item['first_shortage_shipping_date_str'] = item['first_shortage_shipping_date'].strftime('%Y-%m-%d')
        else:
            item['first_shortage_shipping_date_str'] = ''
        
        if item.get('estimated_arrival_date'):
            item['estimated_arrival_date_str'] = item['estimated_arrival_date'].strftime('%Y-%m-%d')
        else:
            item['estimated_arrival_date_str'] = ''

    material_number_filter = request.GET.get('material_number')
    purchaser_filter = request.GET.get('purchaser')
    shortage_date_filter = request.GET.get('shortage_date') # Get the shortage date filter
    sort_by = request.GET.get('sort_by', 'first_demand_date')
    order = request.GET.get('order', 'asc')

    # Apply filters
    if material_number_filter:
        demand_list = [m for m in demand_list if material_number_filter.lower() in m['material_number'].lower()]

    if purchaser_filter:
        demand_list = [m for m in demand_list if m.get('purchaser') and purchaser_filter.lower() in m.get('purchaser', '').lower()]

    if shortage_date_filter:
        try:
            filter_date = datetime.datetime.strptime(shortage_date_filter, '%Y-%m-%d').date()
            demand_list = [m for m in demand_list if m.get('shortage_date') and m.get('shortage_date') <= filter_date]
        except ValueError:
            messages.error(request, "無效的日期格式，請使用 YYYY-MM-DD 格式。")
    
    # Sorting
    reverse_order = order == 'desc'
    if sort_by == 'material_number':
        demand_list.sort(key=lambda x: x['material_number'], reverse=reverse_order)
    elif sort_by == 'first_shortage_shipping_date': # New sort option
        demand_list.sort(key=lambda x: x['first_shortage_shipping_date'] if x['first_shortage_shipping_date'] else datetime.date.max, reverse=reverse_order)
    elif sort_by == 'estimated_arrival_date': # New sort option
        demand_list.sort(key=lambda x: x['estimated_arrival_date'] if x['estimated_arrival_date'] else datetime.date.max, reverse=reverse_order)
    # Add other sort options as needed

    # Add detail_orders_json for the frontend
    for item in demand_list:
        # Convert Decimal and date to string for JSON serialization
        for detail in item['detail_orders']:
            if isinstance(detail['required_quantity'], Decimal):
                detail['required_quantity'] = str(detail['required_quantity'])
            if isinstance(detail['demand_date'], datetime.date):
                detail['demand_date'] = str(detail['demand_date'])
            if isinstance(detail['shipping_date'], datetime.date): # Format shipping_date
                detail['shipping_date'] = str(detail['shipping_date'])
            else:
                detail['shipping_date'] = '' # Ensure it's a string even if None
        item['detail_orders_json'] = json.dumps(item['detail_orders'])

    # Pagination
    paginator = Paginator(demand_list, 20)
    page_number = request.GET.get('page')
    page_obj = paginator.get_page(page_number)

    # Get all unique purchasers for the filter dropdown
    all_purchasers = User.objects.filter(
        username__in=Material.objects.values_list('purchaser', flat=True).distinct()
    ).order_by('username')
    purchaser_choices = [(p.username, p.username) for p in all_purchasers if p.username]

    context = {
        'demand_list': page_obj,
        'material_number_filter': material_number_filter,
        'purchaser_filter': purchaser_filter,
        'shortage_date_filter': shortage_date_filter, # Add to context
        'purchaser_choices': purchaser_choices,
        'sort_by': sort_by,
        'order': order,
    }
    return render(request, 'requisitions/estimated_material_demand.html', context)


@login_required
def shortage_materials_list(request):
    is_dispatcher_supervisor = request.user.groups.filter(name=GROUP_NAMES['DISPATCHER_SUPERVISOR']).exists()
    if not request.user.is_superuser and not is_dispatcher_supervisor:
        messages.error(request, "您沒有權限查看此頁面。")
        return redirect('core:homepage')

    # 只顯示被標記為「缺料」的物料（透過簡易撥料頁面的缺料按鈕標記）
    backordered_items = RequisitionItem.objects.filter(
        dispatch_status='backordered',
        requisition__is_archived=False
    )
    
    # 聚合相同物料
    aggregated_shortages = {}
    for item in backordered_items:
        key = item.material_number
        shortage = item.required_quantity - (item.confirmed_quantity or 0)
        if shortage <= 0:
            continue
            
        if key not in aggregated_shortages:
            # 嘗試從 WorkOrderMaterial 取得預計入料日期
            latest_date = WorkOrderMaterial.objects.filter(
                material_number=key
            ).aggregate(latest_date=Max('estimated_arrival_date'))['latest_date']
            
            aggregated_shortages[key] = {
                'material_number': item.material_number,
                'item_name': item.item_name,
                'total_shortage': Decimal('0.00'),
                'orders': set(),
                'estimated_arrival_date': latest_date
            }
        aggregated_shortages[key]['total_shortage'] += shortage
        aggregated_shortages[key]['orders'].add(item.order_number)

    # Convert to list and format orders_str
    summarized_shortages = list(aggregated_shortages.values())
    for summary in summarized_shortages:
        summary['orders_str'] = ", ".join(sorted(list(summary['orders'])))

    # Get the global final arrival date from all shortage materials
    global_final_arrival_date = None
    all_shortage_arrival_dates = [
        s['estimated_arrival_date'] for s in summarized_shortages if s['estimated_arrival_date']
    ]
    if all_shortage_arrival_dates:
        global_final_arrival_date = max(all_shortage_arrival_dates)

    context = {
        'shortage_materials': summarized_shortages,
        'global_final_arrival_date': global_final_arrival_date,
    }
    return render(request, 'requisitions/shortage_materials_list.html', context)

@login_required
def update_shortage_arrival_dates(request):
    if request.method != 'POST':
        return redirect('requisitions:shortage_materials_list')

    if not request.user.is_superuser and not request.user.groups.filter(name=GROUP_NAMES['DISPATCHER_SUPERVISOR']).exists():
        messages.error(request, "您沒有權限執行此操作。")
        return redirect('requisitions:shortage_materials_list')

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
        print(f"Processing key: {key}, value: {value}") # DEBUG
        is_desktop = key.startswith('arrival_date_desktop_')
        is_mobile = key.startswith('arrival_date_mobile_')

        if is_desktop or is_mobile:
            if is_desktop:
                material_number = key.replace('arrival_date_desktop_', '')
            else:
                material_number = key.replace('arrival_date_mobile_', '')
            
            print(f"Extracted material_number: {material_number}") # DEBUG
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

    return redirect('requisitions:shortage_materials_list')