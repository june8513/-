from django.shortcuts import render, redirect, get_object_or_404
from django.contrib.auth.decorators import login_required
from django.contrib import messages
from requisitions.models import Requisition, RequisitionItem, WorkOrderMaterial, ProcessType
from inventory.models import Material
from django.db import transaction, IntegrityError
import json
from decimal import Decimal, InvalidOperation
from django.http import JsonResponse, HttpResponse
from django.db.models import F, ExpressionWrapper, DecimalField, Subquery, OuterRef, Sum, Q
from django.db.models.functions import Coalesce

@login_required
def finished_goods_dispatch(request):
    is_admin = request.user.is_superuser
    is_applicant_supervisor = request.user.groups.filter(name='申請人員主管').exists()
    is_dispatcher_supervisor = request.user.groups.filter(name='撥料人員主管').exists()
    is_supervisor = is_applicant_supervisor or is_dispatcher_supervisor
    
    context = {}
    
    # 如果是主管或管理員，撈取額外看板數據
    if is_admin or is_supervisor:
        from requisitions.models import RequisitionItem
        from django.db.models import Prefetch
        
        # 定義預加載邏輯：僅抓取狀態為 backordered 的項目
        short_items_prefetch = Prefetch(
            'items', 
            queryset=RequisitionItem.objects.filter(dispatch_status='backordered'),
            to_attr='short_items'
        )

        # 成品看板數據：所有申請單
        context['all_requisitions'] = Requisition.objects.filter(requisition_type='finished').order_by('-created_at')[:20]
        
        # 成品看板數據：有缺料的申請單 (單號分組)
        context['shortage_requisitions'] = Requisition.objects.filter(
            requisition_type='finished',
            is_archived=False,
            items__dispatch_status='backordered'
        ).distinct().prefetch_related(short_items_prefetch).order_by('-created_at')
        
        # 半成品看板數據：所有申請單
        context['semi_all_requisitions'] = Requisition.objects.filter(requisition_type='semi_finished').order_by('-created_at')[:20]
        
        # 半成品看板數據：有缺料的申請單 (單號分組)
        context['semi_shortage_requisitions'] = Requisition.objects.filter(
            requisition_type='semi_finished',
            is_archived=False,
            items__dispatch_status='backordered'
        ).distinct().prefetch_related(short_items_prefetch).order_by('-created_at')
        
    return render(request, 'requisitions/finished_goods_dispatch.html', context)

@login_required
def update_dispatch_note(request, pk):
    requisition = get_object_or_404(Requisition, pk=pk)

    is_admin = request.user.is_superuser
    is_applicant = request.user == requisition.applicant
    is_material_handler = request.user.groups.filter(name='撥料人員').exists()

    if not (is_admin or is_applicant or is_material_handler):
        messages.error(request, "您沒有權限訪問此頁面。")
        return redirect('core:homepage') # Changed from 'homepage'
    if request.method == 'POST':
        confirm_value = request.POST.get('confirm')
        if confirm_value:
            material_id, action = confirm_value.split('_')
            try:
                material = WorkOrderMaterial.objects.get(pk=material_id)
                if action == 'yes':
                    material.confirmed_quantity = material.required_quantity
                    material.is_signed_off = True
                    messages.success(request, f"物料 {material.material_number} 已確認撥料。")
                else:
                    material.confirmed_quantity = 0
                    material.is_signed_off = False
                    messages.info(request, f"物料 {material.material_number} 已標記為未撥料。")
                material.save()
            except WorkOrderMaterial.DoesNotExist:
                messages.error(request, f"物料 ID {material_id} 不存在。")
            except Exception as e:
                messages.error(request, f"更新物料時發生錯誤: {e}")
    return redirect('requisitions:generate_dispatch_note', pk=requisition.pk)

@login_required
@transaction.atomic
def update_material_dispatch_status(request, pk):
    requisition = get_object_or_404(Requisition, pk=pk)

    is_admin = request.user.is_superuser
    is_applicant = request.user == requisition.applicant
    is_material_handler = request.user.groups.filter(name='撥料人員').exists()

    if not (is_admin or is_applicant or is_material_handler):
        return JsonResponse({'success': False, 'message': '權限不足'}, status=403)

    if request.method == 'POST':
        try:
            data = json.loads(request.body)
            material_id = data.get('material_id')
            action = data.get('action') # 'yes' or 'no'

            material = get_object_or_404(WorkOrderMaterial, pk=material_id)
            requisition = get_object_or_404(Requisition, pk=pk)

            if action == 'yes':
                material.confirmed_quantity = material.required_quantity
                material.is_signed_off = True
                message = f"物料 {material.material_number} 已確認撥料。"
            elif action == 'no':
                material.confirmed_quantity = Decimal('0.00') # Set to 0 for backorder
                material.is_signed_off = False
                message = f"物料 {material.material_number} 已取消撥料並移至欠料。"
            else:
                return JsonResponse({'success': False, 'message': '無效的操作。'}, status=400)

            material.save()

            # Update Requisition status if all materials are dispatched/undispatched
            # This logic might need refinement based on exact business rules
            # For now, let's just return success and let the user refresh or handle UI updates
            return JsonResponse({'success': True, 'message': message, 'new_confirmed_quantity': str(material.confirmed_quantity), 'new_is_signed_off': material.is_signed_off})

        except WorkOrderMaterial.DoesNotExist:
            return JsonResponse({'success': False, 'message': '找不到物料。'}, status=404)
        except Requisition.DoesNotExist:
            return JsonResponse({'success': False, 'message': '找不到撥料單。'}, status=404)
        except json.JSONDecodeError:
            return JsonResponse({'success': False, 'message': '無效的 JSON 請求。'}, status=400)
        except Exception as e:
            import traceback
            return JsonResponse({'success': False, 'message': f'處理請求時發生錯誤: {e}\n{traceback.format_exc()}'}, status=500)
    return JsonResponse({'success': False, 'message': '無效的請求方法。'}, status=405)

@login_required
def generate_backorder_note(request, pk):
    requisition = get_object_or_404(Requisition, pk=pk)

    is_admin = request.user.is_superuser
    is_applicant = request.user == requisition.applicant
    is_material_handler = request.user.groups.filter(name='撥料人員').exists()

    if not (is_admin or is_applicant or is_material_handler):
        messages.error(request, "您沒有權限訪問此頁面。")
        return redirect('core:homepage') # Changed from 'requisitions:homepage'

    # Subquery to get bin and system_quantity from inventory.models.Material
    inventory_subquery_storage_bin = Subquery(
        Material.objects.filter(material_code=OuterRef('material_number')).values('bin')[:1]
    )
    inventory_subquery_stock_quantity = Subquery(
        Material.objects.filter(material_code=OuterRef('material_number')).values('system_quantity')[:1]
    )
    # Filter for RequisitionItems that are not fully dispatched (shortage)
    # We use RequisitionItem as the source of truth because it tracks confirmed_quantity
    # and includes all items (including those from Kit child process types)
    shortage_materials = RequisitionItem.objects.filter(
        requisition=requisition,
        required_quantity__gt=Coalesce(F('confirmed_quantity'), Decimal('0'))
    ).annotate(
        shortage_quantity=ExpressionWrapper(
            F('required_quantity') - Coalesce(F('confirmed_quantity'), Decimal('0')),
            output_field=DecimalField()
        ),
        storage_bin=inventory_subquery_storage_bin,
        stock_quantity=inventory_subquery_stock_quantity,
        estimated_arrival_date=F('source_material__estimated_arrival_date')
    ).order_by('order_number', 'material_number')

    context = {
        'requisition': requisition,
        'shortage_materials': shortage_materials,
    }
    return render(request, 'requisitions/backorder_note.html', context)

@login_required
def supplement_material(request):
    # Placeholder function
    return HttpResponse("This is a placeholder for supplement_material.")

@login_required
def batch_dispatch_view(request):
    """
    Displays an aggregated list of materials for batch dispatch,
    only showing materials that still require dispatching.
    """
    req_ids = request.GET.getlist('req_id')
    
    # Handle case where req_id is passed as a comma-separated string (e.g. from sorting links)
    if len(req_ids) == 1 and ',' in req_ids[0]:
        req_ids = req_ids[0].split(',')

    if not req_ids:
        messages.error(request, "沒有選擇任何申請單。")
        return redirect('requisitions:requisition_list')

    requisitions = Requisition.objects.filter(
        pk__in=req_ids,
        dispatch_performed=False,
        is_archived=False
    )
    
    if len(req_ids) != requisitions.count():
        messages.error(request, "包含無效或已完成撥料的申請單。")
        return redirect('requisitions:requisition_list')

    # --- 1. Get all items that still need dispatching, and pre-calculate their need ---
    items_to_process = RequisitionItem.objects.filter(
        requisition__in=requisitions
    ).annotate(
        current_confirmed=Coalesce('confirmed_quantity', Decimal('0'))
    ).filter(
        required_quantity__gt=F('current_confirmed')
    ).select_related('requisition', 'source_material__process_type').order_by('material_number', 'requisition__created_at')

    # --- 2. Group them by Kit or Material for the template ---
    grouped_items = {}
    
    # Pre-fetch ProcessTypes for all items to avoid N+1 queries
    # We need to know if the item's source material's process type is a kit or has a kit parent
    # RequisitionItem -> WorkOrderMaterial (source_material) -> ProcessType
    
    # Helper to get kit info
    def get_kit_info(item):
        if not item.source_material or not item.source_material.process_type:
            return None
        
        pt = item.source_material.process_type
        if pt.is_kit:
            return pt
        if pt.parent and pt.parent.is_kit:
            return pt.parent
        return None

    for item in items_to_process:
        kit = get_kit_info(item)
        
        if kit:
            # Group by Kit
            group_key = f"KIT_{kit.id}"
            group_display_name = f"台份: {kit.name}"
            is_kit_group = True
            kit_id = kit.id
        else:
            # Group by Material Number (Legacy behavior)
            group_key = item.material_number
            group_display_name = f"物料: {item.material_number}"
            is_kit_group = False
            kit_id = None

        if group_key not in grouped_items:
            # First time seeing this group
            if not is_kit_group:
                # Fetch stock info for single material
                main_material = Material.objects.filter(material_code=item.material_number).first()
                stock_quantity = main_material.system_quantity if main_material else Decimal('0')
                storage_bin = main_material.bin if main_material else ''
                item_name_display = item.item_name
            else:
                # For kits, stock info is per item, so we don't show it at group level
                stock_quantity = Decimal('0') 
                storage_bin = 'Multiple'
                item_name_display = '包含多個物料'

            grouped_items[group_key] = {
                'display_name': group_display_name,
                'is_kit': is_kit_group,
                'kit_id': kit_id,
                'item_name': item_name_display,
                'storage_bin': storage_bin,
                'stock_quantity': stock_quantity,
                'items': [],
                'material_number': item.material_number # Keep for sorting if not kit
            }
        
        # Add the item itself to the list for this group
        item.quantity_needed = item.required_quantity - item.current_confirmed
        
        # If it's a kit, we might want to fetch individual stock info for display in the table
        if is_kit_group:
             main_material = Material.objects.filter(material_code=item.material_number).first()
             item.stock_quantity_display = main_material.system_quantity if main_material else Decimal('0')
             item.storage_bin_display = main_material.bin if main_material else ''

        grouped_items[group_key]['items'].append(item)

    # Calculate total remaining need for each group
    for key, details in grouped_items.items():
        total_need = sum(item.quantity_needed for item in details['items'])
        details['total_remaining_need'] = total_need

    # --- 3. Sorting Logic ---
    sort_option = request.GET.get('sort', 'material') # Default to material sort
    
    if sort_option == 'storage_bin':
        # Sort by storage_bin (empty bins last), then by material_number
        sorted_items = dict(sorted(grouped_items.items(), key=lambda item: (item[1]['storage_bin'] == '', item[1]['storage_bin'], item[0])))
    else:
        # Default: Sort by material_number
        sorted_items = dict(sorted(grouped_items.items(), key=lambda item: item[0]))

    context = {
        'grouped_items': sorted_items,
        'requisition_ids': ",".join(req_ids),
        'current_sort': sort_option
    }
    return render(request, 'requisitions/batch_dispatch_aggregated.html', context)

@login_required
@transaction.atomic
def batch_dispatch_action(request):
    """
    Handles the submission of the batch dispatch form from the manual allocation UI.
    Processes individual quantities for each RequisitionItem.
    Updates RequisitionItem, Inventory, and parent Requisition status.
    """
    if request.method != 'POST':
        messages.error(request, "無效的請求。")
        return redirect('requisitions:requisition_list')

    requisition_ids_str = request.POST.get('requisition_ids', '')
    if not requisition_ids_str:
        messages.error(request, "沒有提供申請單 ID。")
        return redirect('requisitions:requisition_list')

    # --- 1. Parse dispatched quantities for each item from the form ---
    item_quantities = {}
    for key, value in request.POST.items():
        if key.startswith('dispatch_item_'):
            if not value or value.isspace():
                continue  # Skip if the input is empty

            try:
                item_pk = int(key.replace('dispatch_item_', ''))
                quantity = Decimal(value)
                if quantity > 0:  # Only process positive quantities
                    item_quantities[item_pk] = quantity
            except (ValueError, InvalidOperation):
                messages.error(request, f"提供的數量 '{value}' 無效。")
                return redirect('requisitions:requisition_list')

    if not item_quantities:
        messages.warning(request, "沒有輸入任何撥料數量。")
        return redirect('requisitions:requisition_list')

    # --- 2. Process each item and update inventory ---
    # Keep track of total dispatch per material for inventory update
    inventory_updates = {}
    updated_item_pks = set(item_quantities.keys())

    items_to_update = RequisitionItem.objects.filter(pk__in=updated_item_pks).select_related('requisition')
    
    for item in items_to_update:
        quantity_to_add = item_quantities.get(item.pk, Decimal('0'))
        if quantity_to_add <= 0:
            continue

        # Security/safety check
        current_confirmed = item.confirmed_quantity or Decimal('0')
        needed = item.required_quantity - current_confirmed
        if quantity_to_add > needed:
            messages.error(request, f"物料 {item.material_number} 的撥料數量 ({quantity_to_add}) 超過尚需數量 ({needed})。")
            # Rolling back transaction manually
            transaction.set_rollback(True)
            return redirect('requisitions:requisition_list')

        item.confirmed_quantity += quantity_to_add
        item.save()

        # Aggregate inventory changes
        if item.material_number not in inventory_updates:
            inventory_updates[item.material_number] = Decimal('0')
        inventory_updates[item.material_number] += quantity_to_add

    # --- 3. Update inventory in a separate loop ---
    for material_number, total_to_dispatch in inventory_updates.items():
        try:
            # Use lock for update to prevent race conditions
            inventory_item = Material.objects.select_for_update().get(material_code=material_number)
            if inventory_item.system_quantity < total_to_dispatch:
                messages.error(request, f"物料 {material_number} 的庫存 ({inventory_item.system_quantity}) 不足，無法撥料 {total_to_dispatch}。")
                transaction.set_rollback(True)
                return redirect('requisitions:requisition_list')
            
            inventory_item.system_quantity -= total_to_dispatch
            inventory_item.save()
        except Material.DoesNotExist:
            messages.error(request, f"庫存中找不到物料 {material_number}。")
            transaction.set_rollback(True)
            return redirect('requisitions:requisition_list')

    # --- 4. Update dispatch_status for all affected items ---
    for item in items_to_update:
        if item.pk in updated_item_pks:  # Redundant check, but safe
            if item.confirmed_quantity >= item.required_quantity:
                item.dispatch_status = 'dispatched'
            elif item.confirmed_quantity > 0:
                item.dispatch_status = 'dispatched' # Partially dispatched is still 'dispatched'
            else:
                # This case shouldn't be hit if we only process quantity > 0, but for completeness:
                item.dispatch_status = 'backordered'
            item.save()

    # --- 5. Update status for all affected requisitions ---
    requisition_ids = [int(id) for id in requisition_ids_str.split(',') if id.isdigit()]
    requisitions_to_update = Requisition.objects.filter(pk__in=requisition_ids)
    for req in requisitions_to_update:
        # Re-fetch all items for the requisition to get the most up-to-date state
        all_req_items = RequisitionItem.objects.filter(requisition=req)
        
        total_required = sum(i.required_quantity for i in all_req_items)
        total_confirmed = sum(i.confirmed_quantity or Decimal('0') for i in all_req_items)

        if total_confirmed >= total_required:
            req.status = 'dispatch_completed'
            req.dispatch_performed = True
        elif total_confirmed > 0:
            req.status = 'dispatch_in_progress'
            # dispatch_performed remains False
        else:
            # Status remains as it was (e.g., demand_submitted)
            pass
        req.save()

    messages.success(request, f"批量撥料操作成功！")
    return redirect('requisitions:requisition_list')
