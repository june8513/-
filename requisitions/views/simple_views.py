"""
簡易介面視圖 - 給一般申請人員和撥料人員使用
"""
from django.shortcuts import render, redirect, get_object_or_404
from django.contrib.auth.decorators import login_required
from django.contrib import messages
from django.utils import timezone
from django.db import transaction
from django.db.models import Q, Count, Sum, Case, When, Value, IntegerField
from django.http import JsonResponse
from decimal import Decimal
from datetime import date

from ..models import (
    Requisition, RequisitionItem, WorkOrderMaterial, ProcessType, 
    MachineModel, WorkOrder
)
from ..constants import GROUP_NAMES, PROCESS_CATEGORY_NAMES, PROCESS_CATEGORY_COLORS
from ..forms import RequisitionForm
from inventory.models import Material


def is_simple_applicant(user):
    """檢查是否為簡易申請人員（非主管）"""
    return user.groups.filter(name=GROUP_NAMES['APPLICANT']).exists() and \
           not user.groups.filter(name=GROUP_NAMES['APPLICANT_SUPERVISOR']).exists() and \
           not user.is_superuser


def is_simple_dispatcher(user):
    """檢查是否為簡易撥料人員（非主管）"""
    return user.groups.filter(name=GROUP_NAMES['DISPATCHER']).exists() and \
           not user.groups.filter(name=GROUP_NAMES['DISPATCHER_SUPERVISOR']).exists() and \
           not user.is_superuser


# =====================
# 簡易申請人員視圖
# =====================

@login_required
def simple_applicant_home(request):
    """簡易申請人員首頁 - 顯示自己的申請單列表"""
    if not is_simple_applicant(request.user) and not request.user.is_superuser:
        return redirect('requisitions:requisition_list')
    
    # 取得類型參數（成品/半成品）
    current_type = request.GET.get('type', 'finished')
    if current_type not in ['finished', 'semi_finished']:
        current_type = 'finished'
    
    requisitions = Requisition.objects.filter(
        applicant=request.user
    ).order_by('-created_at')
    
    # TODO: 未來可根據 current_type 過濾不同類型的申請單
    # 目前先顯示所有申請單，待申請單有 material_type 欄位後再過濾
    
    # 計算每個申請單的撥料進度
    today = timezone.now().date()
    for req in requisitions:
        items = req.items.all()
        total = items.count()
        dispatched = items.filter(dispatch_status='dispatched').count()
        req.progress = int((dispatched / total * 100) if total > 0 else 0)
        req.dispatched_count = dispatched
        req.total_count = total
        # 檢查是否逾期（需求日期已過且未完成撥料）
        req.is_overdue = (
            req.request_date < today and 
            req.status in ['demand_submitted', 'dispatch_in_progress']
        )
    
    context = {
        'requisitions': requisitions,
        'user': request.user,
        'current_type': current_type,
    }
    return render(request, 'requisitions/simple/simple_applicant_home.html', context)


@login_required
def simple_applicant_create(request):
    """簡易申請人員建立申請單"""
    if not is_simple_applicant(request.user) and not request.user.is_superuser:
        return redirect('requisitions:requisition_list')
    
    # 讀取類型參數
    current_type = request.GET.get('type', 'finished')
    if request.method == 'POST':
        current_type = request.POST.get('type', 'finished')
    
    if request.method == 'POST':
        order_number = request.POST.get('order_number')
        process_type_id = request.POST.get('process_type')
        request_date = request.POST.get('request_date')
        
        # Check if the work order is archived
        try:
            work_order = WorkOrder.objects.get(order_number=order_number)
            if work_order.is_archived:
                messages.error(request, f"工單 {order_number} 已被歸檔，無法為其建立新的撥料申請單。")
                return render(request, 'requisitions/simple/simple_applicant_create.html', {
                    'current_type': current_type,
                    'today': date.today().isoformat()
                })
        except WorkOrder.DoesNotExist:
            pass
        
        try:
            if current_type == 'semi_finished':
                # 半成品申請單
                from requisitions.models import SemiFinishedProcessType
                
                process_type_obj = get_object_or_404(SemiFinishedProcessType, id=process_type_id)
                
                # 檢查是否已存在
                existing_requisition = Requisition.objects.filter(
                    order_number=order_number,
                    process_type=process_type_obj.name,
                    requisition_type='semi_finished'
                ).first()
                
                if existing_requisition:
                    messages.error(request, "此訂單單號在該投料點已存在半成品申請單。")
                    return render(request, 'requisitions/simple/simple_applicant_create.html', {
                        'current_type': current_type,
                        'today': date.today().isoformat()
                    })
                
                # 建立申請單
                requisition = Requisition.objects.create(
                    applicant=request.user,
                    order_number=order_number,
                    process_type=process_type_obj.name,
                    status='demand_submitted',
                    requisition_type='semi_finished',
                    request_date=request_date or date.today()
                )
                
                # 新增物料項目 - 從半成品 WorkOrderMaterial
                materials_to_add = WorkOrderMaterial.objects.filter(
                    order_number=order_number,
                    material_type='semi_finished',
                    is_active=True
                )
                
                items_to_create = []
                for material in materials_to_add:
                    main_material = Material.objects.filter(material_code=material.material_number).first()
                    stock_quantity = main_material.system_quantity if main_material else Decimal('0')
                    storage_bin = main_material.bin if main_material else ''

                    items_to_create.append(
                        RequisitionItem(
                            requisition=requisition,
                            source_material=material,
                            order_number=material.order_number,
                            material_number=material.material_number,
                            item_name=material.item_name,
                            required_quantity=material.required_quantity,
                            stock_quantity=stock_quantity,
                            storage_bin=storage_bin,
                        )
                    )
                
                if items_to_create:
                    RequisitionItem.objects.bulk_create(items_to_create)
                
                messages.success(request, "半成品撥料申請單建立成功！")
                return redirect('requisitions:simple_applicant_home')
            else:
                # 成品申請單 - 原有邏輯
                # Re-generate choices for process_type based on the submitted order_number
                material_process_type_ids = WorkOrderMaterial.objects.filter(
                    order_number=order_number,
                    is_active=True
                ).values_list('process_type__id', flat=True).distinct()

                used_requisition_process_type_names = Requisition.objects.filter(
                    order_number=order_number
                ).values_list('process_type', flat=True)

                available_process_types_query = ProcessType.objects.filter(
                    id__in=material_process_type_ids
                ).exclude(
                    name__in=used_requisition_process_type_names
                ).order_by('name')
                
                form_process_type_choices = [(pt.id, pt.name) for pt in available_process_types_query]
                form = RequisitionForm(request.POST, process_type_choices=form_process_type_choices)
                
                if form.is_valid():
                    existing_requisition = Requisition.objects.filter(
                        order_number=order_number,
                        process_type=form.cleaned_data['process_type']
                    ).first()

                    if existing_requisition:
                        messages.error(request, "此訂單單號在該需求流程中已存在。")
                        return render(request, 'requisitions/simple/simple_applicant_create.html', {
                            'form': form,
                            'current_type': current_type,
                            'today': date.today().isoformat()
                        })

                    selected_process_type_id = form.cleaned_data['process_type']
                    selected_process_type_obj = get_object_or_404(ProcessType, id=selected_process_type_id)

                    requisition = form.save(commit=False)
                    requisition.applicant = request.user
                    requisition.order_number = order_number
                    requisition.process_type = selected_process_type_obj.name
                    requisition.status = 'demand_submitted'
                    requisition.save()

                    # Find related process types (including children for Kits)
                    related_process_types = [requisition.process_type]
                    try:
                        parent_pt = ProcessType.objects.filter(
                            name=requisition.process_type, 
                            machine_model__work_order_materials__order_number=requisition.order_number
                        ).distinct().get()
                        
                        children_pts = ProcessType.objects.filter(parent=parent_pt).values_list('name', flat=True)
                        related_process_types.extend(children_pts)
                    except Exception:
                        pass

                    materials_to_add = WorkOrderMaterial.objects.filter(
                        order_number=requisition.order_number,
                        process_type__name__in=related_process_types,
                        is_active=True
                    )

                    items_to_create = []
                    for material in materials_to_add:
                        main_material = Material.objects.filter(material_code=material.material_number).first()
                        stock_quantity = main_material.system_quantity if main_material else Decimal('0')
                        storage_bin = main_material.bin if main_material else ''

                        items_to_create.append(
                            RequisitionItem(
                                requisition=requisition,
                                source_material=material,
                                order_number=material.order_number,
                                material_number=material.material_number,
                                item_name=material.item_name,
                                required_quantity=material.required_quantity,
                                stock_quantity=stock_quantity,
                                storage_bin=storage_bin,
                            )
                        )
                    
                    if items_to_create:
                        RequisitionItem.objects.bulk_create(items_to_create)

                    messages.success(request, "撥料申請單建立成功！")
                    return redirect('requisitions:simple_applicant_home')
        except Exception as e:
            messages.error(request, f"建立申請單時發生錯誤：{str(e)}")
    else:
        form = RequisitionForm()
    
    context = {
        'form': form if 'form' in dir() else RequisitionForm(),
        'current_type': current_type,
        'today': date.today().isoformat()
    }
    return render(request, 'requisitions/simple/simple_applicant_create.html', context)


@login_required
def simple_applicant_detail(request, pk):
    """簡易申請人員查看申請單詳情"""
    requisition = get_object_or_404(Requisition, pk=pk)
    
    # 檢查是否為申請人本人
    if requisition.applicant != request.user and not request.user.is_superuser:
        messages.error(request, "您沒有權限查看此申請單。")
        return redirect('requisitions:simple_applicant_home')
    
    items = requisition.items.all().order_by('material_number')
    
    # 計算進度
    total = items.count()
    dispatched = items.filter(dispatch_status='dispatched').count()
    progress = int((dispatched / total * 100) if total > 0 else 0)
    
    # 標記缺料物料
    for item in items:
        if item.dispatch_status == 'backordered':
            item.is_shortage = True
            # 嘗試取得預計入料日期
            if item.source_material:
                item.expected_date = item.source_material.demand_date
            else:
                item.expected_date = None
        else:
            item.is_shortage = False
    
    context = {
        'requisition': requisition,
        'items': items,
        'progress': progress,
        'dispatched_count': dispatched,
        'total_count': total,
    }
    return render(request, 'requisitions/simple/simple_applicant_detail.html', context)


@login_required
def simple_applicant_update_process_type(request, pk):
    """申請人員修改投料點"""
    requisition = get_object_or_404(Requisition, pk=pk)
    is_ajax = request.headers.get('X-Requested-With') == 'XMLHttpRequest'
    
    # 檢查是否為申請人本人或管理員
    if requisition.applicant != request.user and not request.user.is_superuser:
        if is_ajax:
            return JsonResponse({'success': False, 'message': '您沒有權限修改此申請單。'})
        messages.error(request, "您沒有權限修改此申請單。")
        return redirect('requisitions:simple_applicant_home')
    
    # 只允許尚未開始撥料的申請單修改投料點
    if requisition.status not in ['demand_submitted']:
        if is_ajax:
            return JsonResponse({'success': False, 'message': '已開始撥料的申請單無法修改投料點。'})
        messages.error(request, "已開始撥料的申請單無法修改投料點。")
        return redirect('requisitions:simple_applicant_detail', pk=pk)
    
    if request.method == 'POST':
        new_process_type_id = request.POST.get('process_type')
        
        if not new_process_type_id:
            if is_ajax:
                return JsonResponse({'success': False, 'message': '請選擇投料點。'})
            messages.error(request, "請選擇投料點。")
            return redirect('requisitions:simple_applicant_detail', pk=pk)
        
        try:
            # 根據申請單類型取得投料點
            if requisition.requisition_type == 'semi_finished':
                from requisitions.models import SemiFinishedProcessType
                process_type_obj = get_object_or_404(SemiFinishedProcessType, id=new_process_type_id)
                new_process_type_name = process_type_obj.name
                
                # 檢查是否與其他申請單重複
                existing = Requisition.objects.filter(
                    order_number=requisition.order_number,
                    process_type=new_process_type_name,
                    requisition_type='semi_finished'
                ).exclude(pk=pk).exists()
            else:
                process_type_obj = get_object_or_404(ProcessType, id=new_process_type_id)
                new_process_type_name = process_type_obj.name
                
                # 檢查是否與其他申請單重複
                existing = Requisition.objects.filter(
                    order_number=requisition.order_number,
                    process_type=new_process_type_name
                ).exclude(pk=pk).exists()
            
            if existing:
                if is_ajax:
                    return JsonResponse({'success': False, 'message': f'該工單已有「{new_process_type_name}」的申請單。'})
                messages.error(request, f"該工單已有「{new_process_type_name}」的申請單。")
                return redirect('requisitions:simple_applicant_detail', pk=pk)
            
            # 更新投料點
            old_process_type = requisition.process_type
            requisition.process_type = new_process_type_name
            requisition.save()
            
            if is_ajax:
                return JsonResponse({
                    'success': True, 
                    'message': f'投料點已從「{old_process_type}」更新為「{new_process_type_name}」。',
                    'new_process_type': new_process_type_name
                })
            messages.success(request, f"投料點已從「{old_process_type}」更新為「{new_process_type_name}」。")
            
        except Exception as e:
            if is_ajax:
                return JsonResponse({'success': False, 'message': f'更新失敗：{str(e)}'})
            messages.error(request, f"更新失敗：{str(e)}")
    
    return redirect('requisitions:simple_applicant_detail', pk=pk)


@login_required
def simple_applicant_sign_off(request, pk):
    """簡易申請人員簽收功能"""
    requisition = get_object_or_404(Requisition, pk=pk)
    is_ajax = request.headers.get('X-Requested-With') == 'XMLHttpRequest'
    
    # 檢查是否為申請人本人
    if requisition.applicant != request.user and not request.user.is_superuser:
        if is_ajax:
            return JsonResponse({'success': False, 'message': '您沒有權限執行簽收操作。'})
        messages.error(request, "您沒有權限執行簽收操作。")
        return redirect('requisitions:simple_applicant_home')
    
    result = {'success': False, 'message': ''}
    
    if request.method == 'POST':
        with transaction.atomic():
            signed_off_count = 0
            item_pk_signed = None
            
            if 'confirm_all_sign_off' in request.POST:
                # Sign off all dispatched items
                dispatched_items = RequisitionItem.objects.filter(
                    requisition=requisition,
                    dispatch_status='dispatched',
                    is_signed_off=False
                )
                for item in dispatched_items:
                    item.is_signed_off = True
                    item.sign_off_by = request.user
                    item.sign_off_date = timezone.now()
                    item.save()
                    signed_off_count += 1
                
                if signed_off_count > 0:
                    result = {'success': True, 'message': f'成功簽收 {signed_off_count} 筆已撥料項目。', 'all_signed': True}
                else:
                    result = {'success': True, 'message': '沒有新的已撥料項目需要簽收。', 'all_signed': True}
            else:
                # Process individual sign-off
                for key in request.POST.keys():
                    if key.startswith('sign_off_item_'):
                        item_pk = key.split('_')[-1]
                        try:
                            item = RequisitionItem.objects.get(pk=item_pk, requisition=requisition)
                            if not item.is_signed_off:
                                item.is_signed_off = True
                                item.sign_off_by = request.user
                                item.sign_off_date = timezone.now()
                                item.save()
                                signed_off_count += 1
                                item_pk_signed = item_pk
                        except RequisitionItem.DoesNotExist:
                            continue
                
                if signed_off_count > 0:
                    result = {'success': True, 'message': f'成功簽收 {signed_off_count} 筆物料項目。', 'item_pk': item_pk_signed}
            
            # Check if all items are signed off
            all_relevant_items = RequisitionItem.objects.filter(
                requisition=requisition,
                dispatch_status__in=['dispatched', 'backordered']
            )
            all_signed = all_relevant_items.exists() and all(item.is_signed_off for item in all_relevant_items)
            if all_signed:
                requisition.status = 'signed_off'
                requisition.sign_off_by = request.user
                requisition.sign_off_date = timezone.now()
                requisition.save()
                result['requisition_completed'] = True
                if not result.get('all_signed'):
                    result['message'] += ' 撥料單已全部簽收完成！'
    
    if is_ajax:
        return JsonResponse(result)
    
    if result.get('success') and result.get('message'):
        messages.success(request, result['message'])
    
    return redirect('requisitions:simple_applicant_detail', pk=pk)


# =====================
# 簡易撥料人員視圖
# =====================

@login_required
def simple_dispatcher_home(request):
    """簡易撥料人員首頁 - 顯示投料點分類按鈕"""
    if not is_simple_dispatcher(request.user) and not request.user.is_superuser:
        return redirect('requisitions:requisition_list')
    
    # 取得類型參數（成品/半成品）
    current_type = request.GET.get('type', 'finished')
    if current_type not in ['finished', 'semi_finished']:
        current_type = 'finished'
    
    categories = []
    
    if current_type == 'semi_finished':
        # 載入半成品投料點
        from ..models import SemiFinishedProcessType
        semi_process_types = SemiFinishedProcessType.objects.filter(is_active=True).order_by('order', 'name')
        
        for pt in semi_process_types:
            # 計算該投料點的待撥申請單數量
            pending_count = Requisition.objects.filter(
                process_type=pt.name,
                requisition_type='semi_finished',
                status__in=['demand_submitted', 'dispatch_in_progress']
            ).count()
            
            categories.append({
                'name': pt.name,
                'color': pt.color,
                'pending_count': pending_count,
            })
    else:
        # 載入成品投料點（原有邏輯）
        for category_name in PROCESS_CATEGORY_NAMES:
            pending_count = Requisition.objects.filter(
                process_type__icontains=category_name,
                status__in=['demand_submitted', 'dispatch_in_progress']
            ).count()
            
            categories.append({
                'name': category_name,
                'color': PROCESS_CATEGORY_COLORS.get(category_name, '#6B7280'),
                'pending_count': pending_count,
            })
    
    context = {
        'categories': categories,
        'user': request.user,
        'current_type': current_type,
    }
    return render(request, 'requisitions/simple/simple_dispatcher_home.html', context)


@login_required
def simple_dispatcher_category(request, category):
    """簡易撥料人員查看特定投料點的申請單"""
    if not is_simple_dispatcher(request.user) and not request.user.is_superuser:
        return redirect('requisitions:requisition_list')
    
    # 取得類型參數（成品/半成品）
    current_type = request.GET.get('type', 'finished')
    if current_type not in ['finished', 'semi_finished']:
        current_type = 'finished'
    
    today = timezone.now().date()
    
    if current_type == 'semi_finished':
        # 半成品投料點
        from ..models import SemiFinishedProcessType
        
        # 驗證投料點存在
        if not SemiFinishedProcessType.objects.filter(name=category, is_active=True).exists():
            messages.error(request, "無效的半成品投料點。")
            return redirect('requisitions:simple_dispatcher_home')
        
        # 取得投料點顏色
        pt = SemiFinishedProcessType.objects.filter(name=category).first()
        category_color = pt.color if pt else '#6B7280'
        
        # 待撥料申請單
        pending_requisitions = Requisition.objects.filter(
            process_type=category,
            requisition_type='semi_finished',
            status__in=['demand_submitted', 'dispatch_in_progress']
        ).order_by('-created_at')
        
        # 已撥料申請單
        completed_requisitions = Requisition.objects.filter(
            process_type=category,
            requisition_type='semi_finished',
            status__in=['dispatch_completed', 'signed_off']
        ).order_by('-updated_at')[:20]
    else:
        # 成品投料點（原有邏輯）
        if category not in PROCESS_CATEGORY_NAMES:
            messages.error(request, "無效的投料點分類。")
            return redirect('requisitions:simple_dispatcher_home')
        
        category_color = PROCESS_CATEGORY_COLORS.get(category, '#6B7280')
        
        # 待撥料申請單
        pending_requisitions = Requisition.objects.filter(
            process_type__icontains=category,
            status__in=['demand_submitted', 'dispatch_in_progress']
        ).order_by('-created_at')
        
        # 已撥料申請單
        completed_requisitions = Requisition.objects.filter(
            process_type__icontains=category,
            status__in=['dispatch_completed', 'signed_off']
        ).order_by('-updated_at')[:20]
    
    # 計算每個申請單的逾期狀態
    for req in pending_requisitions:
        req.is_overdue = req.request_date < today
    
    context = {
        'category': category,
        'category_color': category_color,
        'pending_requisitions': pending_requisitions,
        'completed_requisitions': completed_requisitions,
        'current_type': current_type,
    }
    return render(request, 'requisitions/simple/simple_dispatcher_category.html', context)


@login_required
def simple_dispatcher_detail(request, category, pk):
    """簡易撥料人員撥退料操作頁面"""
    if not is_simple_dispatcher(request.user) and not request.user.is_superuser:
        return redirect('requisitions:requisition_list')
    
    requisition = get_object_or_404(Requisition, pk=pk)
    items = requisition.items.all().order_by('material_number')
    
    if request.method == 'POST':
        action = request.POST.get('action')
        item_pk = request.POST.get('item_pk')
        is_ajax = request.headers.get('X-Requested-With') == 'XMLHttpRequest'
        
        result = {'success': False, 'message': ''}
        
        if item_pk:
            try:
                item = RequisitionItem.objects.get(pk=item_pk, requisition=requisition)
                
                if action == 'dispatch':
                    # 撥料
                    dispatched_qty = request.POST.get('dispatched_qty')
                    try:
                        dispatched_qty = Decimal(dispatched_qty)
                        item.confirmed_quantity = dispatched_qty
                        item.dispatch_status = 'dispatched'
                        item.save()
                        result = {'success': True, 'message': f'物料 {item.material_number} 撥料 {dispatched_qty} 成功。', 'new_status': 'dispatched', 'dispatched_qty': str(dispatched_qty)}
                        if not is_ajax:
                            messages.success(request, result['message'])
                    except Exception as e:
                        result = {'success': False, 'message': f'撥料失敗：{str(e)}'}
                        if not is_ajax:
                            messages.error(request, result['message'])
                
                elif action == 'backorder':
                    # 標記缺料
                    item.dispatch_status = 'backordered'
                    item.save()
                    result = {'success': True, 'message': f'物料 {item.material_number} 已標記為缺料。', 'new_status': 'backordered'}
                    if not is_ajax:
                        messages.success(request, result['message'])
                
                elif action == 'undo':
                    # 退料/取消撥料 - 只有未簽收的才能撤銷
                    if item.is_signed_off:
                        result = {'success': False, 'message': f'物料 {item.material_number} 已簽收，無法撤銷。請使用補撥功能。'}
                        if not is_ajax:
                            messages.error(request, result['message'])
                    else:
                        item.confirmed_quantity = Decimal('0')
                        item.dispatch_status = None
                        item.save()
                        result = {'success': True, 'message': f'物料 {item.material_number} 已取消撥料。', 'new_status': 'pending'}
                        if not is_ajax:
                            messages.success(request, result['message'])
                
                elif action == 'supplementary':
                    # 補撥 - 為已簽收但不足量的項目建立補撥記錄
                    supplementary_qty = request.POST.get('supplementary_qty')
                    try:
                        supplementary_qty = Decimal(supplementary_qty)
                        if supplementary_qty <= 0:
                            raise ValueError("補撥數量必須大於 0")
                        
                        # 計算剩餘需求量
                        confirmed = item.confirmed_quantity if item.confirmed_quantity else item.required_quantity
                        remaining = item.required_quantity - confirmed
                        if supplementary_qty > remaining:
                            raise ValueError(f"補撥數量不能超過剩餘需求量 ({remaining})")
                        
                        # 取得最新庫存
                        from inventory.models import Material
                        main_material = Material.objects.filter(material_code=item.material_number).first()
                        stock_quantity = main_material.system_quantity if main_material else Decimal('0')
                        
                        # 建立補撥項目
                        supplementary_item = RequisitionItem.objects.create(
                            requisition=requisition,
                            source_material=item.source_material,
                            order_number=item.order_number,
                            material_number=item.material_number,
                            item_name=item.item_name,
                            required_quantity=supplementary_qty,
                            stock_quantity=stock_quantity,
                            storage_bin=item.storage_bin,
                            is_supplementary=True,
                            parent_item=item
                        )
                        
                        result = {'success': True, 'message': f'已為物料 {item.material_number} 建立 {supplementary_qty} 單位的補撥項目。'}
                        if not is_ajax:
                            messages.success(request, result['message'])
                            return redirect('requisitions:simple_dispatcher_detail', category=category, pk=pk)
                            
                    except Exception as e:
                        result = {'success': False, 'message': f'補撥失敗：{str(e)}'}
                        if not is_ajax:
                            messages.error(request, result['message'])
                
            except RequisitionItem.DoesNotExist:
                result = {'success': False, 'message': '找不到指定的物料項目。'}
                if not is_ajax:
                    messages.error(request, result['message'])
        
        # 更新申請單狀態
        all_items = requisition.items.all()
        dispatched_items = all_items.filter(dispatch_status='dispatched')
        
        if dispatched_items.count() == all_items.count():
            requisition.status = 'dispatch_completed'
        elif dispatched_items.count() > 0:
            requisition.status = 'dispatch_in_progress'
        else:
            requisition.status = 'demand_submitted'
        requisition.save()
        
        if is_ajax:
            # 計算新的進度
            total = all_items.count()
            dispatched = dispatched_items.count()
            progress = int((dispatched / total * 100) if total > 0 else 0)
            result['progress'] = progress
            result['dispatched_count'] = dispatched
            result['total_count'] = total
            return JsonResponse(result)
        
        return redirect('requisitions:simple_dispatcher_detail', category=category, pk=pk)
    
    # 計算進度
    total = items.count()
    dispatched = items.filter(dispatch_status='dispatched').count()
    progress = int((dispatched / total * 100) if total > 0 else 0)
    
    # --- 檢查更早的未撥需求 (優先工單警示) ---
    from django.db.models.functions import Coalesce
    from django.db.models import F
    
    backlog_map = {}
    target_material_numbers = list(items.values_list('material_number', flat=True))
    
    if target_material_numbers:
        # 1. 找出相同物料在其他工單的未撥需求（欠料大於 0）
        from ..models import WorkOrderMaterial
        from django.db.models import DecimalField
        from django.db.models.functions import Coalesce
        from django.db.models import F, Value
        
        other_shortages = WorkOrderMaterial.objects.filter(
            material_number__in=target_material_numbers,
            is_active=True,
        ).annotate(
            shortage=F('required_quantity') - Coalesce(F('confirmed_quantity'), Value(0), output_field=DecimalField())
        ).filter(
            shortage__gt=0  # 只選擇確實有欠料的物料
        ).exclude(
            order_number=requisition.order_number,
            process_type__name=requisition.process_type
        ).select_related('process_type')
        
        # 2. 按工單和投料點分組
        shortage_groups = {}
        for s in other_shortages:
            p_name = s.process_type.name if s.process_type else None
            key = (s.order_number, p_name)
            if key not in shortage_groups:
                shortage_groups[key] = []
            shortage_groups[key].append(s)
        
        if shortage_groups:
            # 3. 找出哪些有更早的需求日期且尚未完成撥料
            date_q = Q()
            for (o_num, p_name) in shortage_groups.keys():
                if p_name:
                    date_q |= Q(order_number=o_num, process_type=p_name)
                else:
                    date_q |= Q(order_number=o_num, process_type__isnull=True)
            
            if date_q:
                earlier_reqs = Requisition.objects.filter(
                    date_q,
                    request_date__lt=requisition.request_date,
                    status__in=['demand_submitted', 'dispatch_in_progress'],  # 只看尚未完成的申請單
                    is_archived=False
                ).values('order_number', 'process_type', 'request_date')
                
                # 4. 建立 backlog_map（只加入確實有欠料的物料）
                for req in earlier_reqs:
                    key = (req['order_number'], req['process_type'])
                    if key in shortage_groups:
                        for s in shortage_groups[key]:
                            # s.shortage 是上面 annotate 計算的欠料數量
                            if s.shortage > 0:
                                if s.material_number not in backlog_map:
                                    backlog_map[s.material_number] = []
                                
                                backlog_map[s.material_number].append({
                                    'order': s.order_number,
                                    'date': req['request_date'],
                                    'shortage': s.shortage
                                })
    
    # 過濾只顯示主項目（非補撥項目），補撥項目會透過 parent_item 關聯顯示
    main_items = items.filter(is_supplementary=False)
    
    # 將 backlog_info 附加到每個物料項目，並預處理顯示文字
    items_list = list(main_items)
    for item in items_list:
        item.backlog_info = backlog_map.get(item.material_number, [])
        item.has_backlog = bool(item.backlog_info)
        # 預處理 CSS 類別
        if item.dispatch_status == 'dispatched':
            item.card_class = 'dispatched'
            qty_display = item.confirmed_quantity if item.confirmed_quantity else item.required_quantity
            item.status_text = f'已撥 {qty_display}'
        elif item.dispatch_status == 'backordered':
            item.card_class = 'backordered'
            item.status_text = '缺料'
        else:
            item.card_class = ''
            item.status_text = ''
        # 預處理是否可撥料（未簽收且未處理）
        item.is_actionable = item.dispatch_status not in ['dispatched', 'backordered']
        
        # 補撥相關計算
        # 如果 confirmed_quantity 為 None 但已撥料，視為已撥完整需求數量
        if item.dispatch_status == 'dispatched' and item.confirmed_quantity is None:
            confirmed = item.required_quantity
        else:
            confirmed = item.confirmed_quantity or Decimal('0')
        item.remaining_quantity = item.required_quantity - confirmed
        # 只有當 確認數量 < 需求數量 時才需要補撥
        item.needs_supplementary = (
            item.is_signed_off and 
            item.dispatch_status == 'dispatched' and 
            confirmed < item.required_quantity
        )
        
        # 載入補撥子項目
        item.supplementary_list = list(item.supplementary_items.all().order_by('pk'))
        for supp in item.supplementary_list:
            # 預處理補撥項目的顯示
            if supp.dispatch_status == 'dispatched':
                supp.card_class = 'dispatched'
                supp_qty = supp.confirmed_quantity if supp.confirmed_quantity else supp.required_quantity
                supp.status_text = f'已撥 {supp_qty}'
            elif supp.dispatch_status == 'backordered':
                supp.card_class = 'backordered'
                supp.status_text = '缺料'
            else:
                supp.card_class = ''
                supp.status_text = '待撥'
            supp.is_actionable = supp.dispatch_status not in ['dispatched', 'backordered']
    
    # 取得類型參數
    current_type = request.GET.get('type', 'finished')
    if current_type not in ['finished', 'semi_finished']:
        current_type = 'finished'
    
    context = {
        'requisition': requisition,
        'items': items_list,
        'category': category,
        'category_color': PROCESS_CATEGORY_COLORS.get(category, '#6B7280'),
        'progress': progress,
        'dispatched_count': dispatched,
        'total_count': total,
        'is_completed': requisition.status in ['dispatch_completed', 'signed_off'],
        'current_type': current_type,
    }
    return render(request, 'requisitions/simple/simple_dispatcher_detail.html', context)
