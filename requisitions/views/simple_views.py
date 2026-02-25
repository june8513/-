"""
簡易介面視圖 - 給一般申請人員和撥料人員使用
"""
from django.shortcuts import render, redirect, get_object_or_404
from django.contrib.auth.decorators import login_required
from django.contrib import messages
from django.contrib.auth.models import User
from django.utils import timezone
from django.db import transaction
from django.db.models import Q, Count, Sum, Case, When, Value, IntegerField, F, DecimalField
from django.db.models import Case, When, Value, IntegerField, Sum
from django.http import JsonResponse
from decimal import Decimal
from datetime import date

from ..models import Requisition, RequisitionItem, Inventory, WorkOrderMaterial
from ..constants import GROUP_NAMES, PROCESS_CATEGORY_NAMES, PROCESS_CATEGORY_COLORS
from ..forms import RequisitionForm
from inventory.models import Material


def is_simple_applicant(user):
    """檢查是否為簡易申請人員（包含主管）"""
    return user.groups.filter(name__in=[GROUP_NAMES['APPLICANT'], GROUP_NAMES['APPLICANT_SUPERVISOR']]).exists() and \
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
    
    base_qs = Requisition.objects.filter(applicant=request.user)
    
    # 如果是主管，顯示所有申請單
    is_supervisor = request.user.groups.filter(name=GROUP_NAMES['APPLICANT_SUPERVISOR']).exists()
    if is_supervisor:
        base_qs = Requisition.objects.all()
    
    # 依狀態分類
    pending_reqs = list(base_qs.filter(status='demand_submitted').order_by('-created_at'))
    in_progress_reqs = list(base_qs.filter(status='dispatch_in_progress').order_by('-created_at'))
    completed_reqs = list(base_qs.filter(status='dispatch_completed').order_by('-updated_at'))
    signed_off_reqs = list(base_qs.filter(status='signed_off').order_by('-updated_at')[:20])
    
    # 計算進度和逾期
    today = timezone.now().date()
    for req_list in [pending_reqs, in_progress_reqs, completed_reqs, signed_off_reqs]:
        for req in req_list:
            items = req.items.all()
            total = items.count()
            dispatched = items.filter(dispatch_status='dispatched').count()
            req.progress = int((dispatched / total * 100) if total > 0 else 0)
            req.dispatched_count = dispatched
            req.total_count = total
            req.is_overdue = (
                req.request_date < today and 
                req.status in ['demand_submitted', 'dispatch_in_progress']
            )
    
    context = {
        'pending_reqs': pending_reqs,
        'in_progress_reqs': in_progress_reqs,
        'completed_reqs': completed_reqs,
        'signed_off_reqs': signed_off_reqs,
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
    
    # 取得申請人列表 (供主管選擇)
    # 找出所有屬於 '申請人員' 群組的用戶
    applicants = User.objects.filter(groups__name=GROUP_NAMES['APPLICANT']).order_by('username')

    if request.method == 'POST':
        action = request.POST.get('action', 'create')
        
        if current_type == 'semi_finished':
            # --- 半成品：批量建立 & 預覽 ---
            if action == 'preview':
                # 步驟 1: 預覽
                order_numbers_raw = request.POST.get('order_numbers', '')
                default_applicant_id = request.POST.get('applicant_id')
                default_request_date = request.POST.get('request_date')
                
                # 解析單號 (一行一個)
                order_numbers = [line.strip() for line in order_numbers_raw.splitlines() if line.strip()]
                
                preview_items = []
                for order_number in order_numbers:
                    # 查詢對應機型 (從 WorkOrderMaterial 找)
                    # 邏輯：找該單號關聯的第一個有效的 WorkOrderMaterial 的機型
                    material_obj = WorkOrderMaterial.objects.filter(
                        order_number=order_number, 
                        machine_model__isnull=False
                    ).first()
                    
                    machine_model_name = material_obj.machine_model.name if material_obj and material_obj.machine_model else "未知機型"
                    
                    preview_items.append({
                        'order_number': order_number,
                        'machine_model': machine_model_name,
                        'applicant_id': default_applicant_id,
                        'request_date': default_request_date
                    })
                
                return render(request, 'requisitions/simple/simple_applicant_batch_preview.html', {
                    'preview_items': preview_items,
                    'applicants': applicants,
                    'today': date.today().strftime('%Y-%m-%d')
                })
                
            elif action == 'confirm_create':
                order_numbers = request.POST.getlist('order_number')
                applicant_ids = request.POST.getlist('applicant_id')
                request_dates = request.POST.getlist('request_date')
                machine_models = request.POST.getlist('machine_model')

                created_count = 0
                error_count = 0
                errors = []

                if len(order_numbers) > 0:
                    for i in range(len(order_numbers)):
                        if i >= len(applicant_ids) or i >= len(request_dates) or i >= len(machine_models):
                            continue
                            
                        order_number = order_numbers[i]
                        applicant_id = applicant_ids[i]
                        request_date_str = request_dates[i]
                        model_name = machine_models[i]

                        if not (order_number and applicant_id and request_date_str):
                             continue

                        try:
                            if applicant_id:
                                target_applicant = User.objects.get(pk=applicant_id)
                            else:
                                target_applicant = request.user
                            
                            # --- Logic to Update/Save Machine Model ---
                            if model_name:
                                 machine_model_obj, _ = MachineModel.objects.get_or_create(name=model_name.strip())
                                 
                                 # Update or Create WorkOrderMaterial 'PARENT_SCOPE'
                                 WorkOrderMaterial.objects.update_or_create(
                                     order_number=order_number,
                                     material_number='PARENT_SCOPE',
                                     defaults={
                                         'machine_model': machine_model_obj,
                                         'item_name': '訂單機型範圍',
                                         'required_quantity': 0,
                                         'is_active': True 
                                     }
                                 )
                            # ------------------------------------------

                            with transaction.atomic():
                                # Duplicate check
                                existing = Requisition.objects.filter(
                                    order_number=order_number,
                                    requisition_type='semi_finished',
                                    process_type='SEMI'
                                ).first()
                                if existing:
                                    error_count += 1
                                    errors.append(f"工單 {order_number} 已存在")
                                    continue

                                requisition = Requisition.objects.create(
                                    applicant=target_applicant,
                                    order_number=order_number,
                                    process_type='SEMI',
                                    status='demand_submitted',
                                    requisition_type='semi_finished',
                                    request_date=request_date_str or date.today()
                                )
                                
                                # Auto-add semi-finished materials
                                materials_to_add = WorkOrderMaterial.objects.filter(
                                    order_number=order_number,
                                    material_type='semi_finished',
                                    is_active=True
                                )
                                
                                items_to_create = []
                                for material in materials_to_add:
                                    items_to_create.append(RequisitionItem(
                                        requisition=requisition,
                                        source_material=material,
                                        order_number=material.order_number,
                                        material_number=material.material_number,
                                        item_name=material.item_name,
                                        required_quantity=material.required_quantity,
                                        stock_quantity=0
                                    ))
                                if items_to_create:
                                    RequisitionItem.objects.bulk_create(items_to_create)
                                
                                created_count += 1
                                
                        except Exception as e:
                            print(f"Error creating requisition for {order_number}: {e}")
                            error_count += 1
                            if hasattr(e, 'message'):
                                errors.append(str(e.message))
                            else:
                                errors.append(str(e))

                    if created_count > 0:
                        messages.success(request, f"成功建立 {created_count} 張半成品申請單！")
                    if error_count > 0:
                         if created_count == 0:
                              messages.error(request, f"建立失敗：<br>{'<br>'.join(errors)}")
                         else:
                              messages.warning(request, f"有 {error_count} 筆未建立：<br>{'<br>'.join(errors)}")
                    
                    return redirect('requisitions:simple_applicant_home')
                
                else:
                    messages.error(request, "未提交任何資料")
                    return redirect('requisitions:simple_applicant_create')

            else:
                 return redirect('requisitions:simple_applicant_create')

        else:
            # --- 成品：單筆建立 (原有邏輯) ---
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
                        'today': date.today().isoformat(),
                        'applicants': applicants
                    })
            except WorkOrder.DoesNotExist:
                # 若工單不存在，是否阻擋？視需求。目前邏輯似乎是依賴 WorkOrderMaterial
                pass

            process_type_obj = get_object_or_404(ProcessType, id=process_type_id)

            # 檢查是否已存在
            existing_requisition = Requisition.objects.filter(
                order_number=order_number,
                process_type=process_type_obj.name,
                requisition_type='finished'
            ).first()

            if existing_requisition:
                messages.error(request, "此訂單單號在該投料點已存在申請單。")
                return render(request, 'requisitions/simple/simple_applicant_create.html', {
                    'current_type': current_type,
                    'today': date.today().isoformat(),
                    'applicants': applicants
                })

            # 建立申請單
            with transaction.atomic():
                requisition = Requisition.objects.create(
                    applicant=request.user,
                    order_number=order_number,
                    process_type=process_type_obj.name,
                    status='demand_submitted',
                    requisition_type='finished',
                    request_date=request_date or date.today()
                )

                # 新增物料項目
                materials_to_add = WorkOrderMaterial.objects.filter(
                    order_number=order_number,
                    process_type=process_type_obj,
                    material_type='finished',
                    is_active=True
                )
                
                items_to_create = []
                for material in materials_to_add:
                    # 嘗試抓取即時庫存
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
                            confirmed_quantity=Decimal('0'),
                        )
                    )
                
                if items_to_create:
                    RequisitionItem.objects.bulk_create(items_to_create)
                    messages.info(request, f"已自動為此申請單加入 {len(items_to_create)} 筆物料項目。")
                else:
                    messages.warning(request, "警告：此申請單在對應的需求流程中沒有找到任何有效的物料項目。")

            messages.success(request, "撥料申請單建立成功！")
            return redirect('requisitions:simple_applicant_home')

    return render(request, 'requisitions/simple/simple_applicant_create.html', {
        'current_type': current_type,
        'today': date.today().isoformat(),
        'applicants': applicants
    })



@login_required
def simple_applicant_delete(request, pk):
    """簡易申請人員刪除申請單"""
    if not is_simple_applicant(request.user) and not request.user.is_superuser:
        return redirect('requisitions:requisition_list')

    requisition = get_object_or_404(Requisition, pk=pk)

    # 權限檢查：只能刪除自己的申請單（除非是超級管理員 或 申請人員主管）
    is_supervisor = request.user.groups.filter(name=GROUP_NAMES['APPLICANT_SUPERVISOR']).exists()
    if requisition.applicant != request.user and not is_supervisor and not request.user.is_superuser:
        messages.error(request, '您無權限刪除此申請單。')
        return redirect('requisitions:simple_applicant_detail', pk=pk)

    # 狀態檢查：只能刪除 'demand_submitted' 狀態的申請單
    if requisition.status != 'demand_submitted':
        messages.error(request, '只能刪除尚未撥料的申請單。')
        return redirect('requisitions:simple_applicant_detail', pk=pk)

    if request.method == 'POST':
        requisition.delete()
        messages.success(request, '申請單已刪除。')
        return redirect('requisitions:simple_applicant_home')


@login_required
def simple_applicant_detail(request, pk):
    """簡易申請人員查看申請單詳情"""
    requisition = get_object_or_404(Requisition, pk=pk)
    
    # 檢查是否為申請人本人 (或主管)
    is_supervisor = request.user.groups.filter(name=GROUP_NAMES['APPLICANT_SUPERVISOR']).exists()
    if requisition.applicant != request.user and not is_supervisor and not request.user.is_superuser:
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
            
        # 檢查庫存是否不足
        item.is_insufficient_stock = False
        if item.stock_quantity is not None and item.required_quantity is not None:
             # 如果尚未撥料且庫存 < 需求，標記為庫存不足
             if item.dispatch_status != 'dispatched' and item.stock_quantity < item.required_quantity:
                 item.is_insufficient_stock = True
                 # 嘗試取得預計入料日期
                 if item.source_material:
                     item.expected_date = item.source_material.estimated_arrival_date
                 else:
                     item.expected_date = None
    
    # 查詢機型
    machine_model_name = ''
    wom = WorkOrderMaterial.objects.filter(
        order_number=requisition.order_number,
        machine_model__isnull=False
    ).select_related('machine_model').first()
    if wom and wom.machine_model:
        machine_model_name = wom.machine_model.name

    context = {
        'requisition': requisition,
        'items': items,
        'progress': progress,
        'dispatched_count': dispatched,
        'total_count': total,
        'machine_model_name': machine_model_name,
    }
    return render(request, 'requisitions/simple/simple_applicant_detail.html', context)


@login_required
def simple_applicant_update_process_type(request, pk):
    """申請人員修改投料點"""
    requisition = get_object_or_404(Requisition, pk=pk)
    is_ajax = request.headers.get('X-Requested-With') == 'XMLHttpRequest'
    
    # 檢查是否為申請人本人或管理員 (或主管)
    is_supervisor = request.user.groups.filter(name=GROUP_NAMES['APPLICANT_SUPERVISOR']).exists()
    if requisition.applicant != request.user and not is_supervisor and not request.user.is_superuser:
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
def simple_applicant_update_request_date(request, pk):
    """申請人員修改需求日期"""
    requisition = get_object_or_404(Requisition, pk=pk)
    is_ajax = request.headers.get('X-Requested-With') == 'XMLHttpRequest'
    
    # Check permission
    is_supervisor = request.user.groups.filter(name=GROUP_NAMES['APPLICANT_SUPERVISOR']).exists()
    if requisition.applicant != request.user and not is_supervisor and not request.user.is_superuser:
        message = "您沒有權限修改此申請單。"
        if is_ajax:
            return JsonResponse({'success': False, 'message': message})
        messages.error(request, message)
        return redirect('requisitions:simple_applicant_home')
    
    # Check status
    if requisition.status not in ['demand_submitted']:
        message = "已開始撥料的申請單無法修改需求日期。"
        if is_ajax:
            return JsonResponse({'success': False, 'message': message})
        messages.error(request, message)
        return redirect('requisitions:simple_applicant_detail', pk=pk)
    
    if request.method == 'POST':
        new_date_str = request.POST.get('request_date')
        
        if not new_date_str:
            message = "請選擇有效的日期。"
            if is_ajax:
                return JsonResponse({'success': False, 'message': message})
            messages.error(request, message)
        else:
            try:
                # Validate and convert date
                from datetime import datetime
                new_date = datetime.strptime(new_date_str, '%Y-%m-%d').date()
                
                requisition.request_date = new_date
                requisition.save()
                
                message = f"需求日期已更新為 {new_date_str}。"
                if is_ajax:
                    return JsonResponse({'success': True, 'message': message})
                messages.success(request, message)
                
            except ValueError:
                message = "日期格式錯誤。"
                if is_ajax:
                    return JsonResponse({'success': False, 'message': message})
                messages.error(request, message)
            
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
        # 半成品：改依「領料人 (Applicant)」分組
        # 找出所有狀態為「待撥」或「撥料中」的半成品申請單
        pending_requisitions = Requisition.objects.filter(
            requisition_type='semi_finished',
            status__in=['demand_submitted', 'dispatch_in_progress']
        ).values('applicant__username').annotate(
            pending_count=Count('id')
        ).order_by('applicant__username')
        
        for item in pending_requisitions:
            username = item['applicant__username']
            count = item['pending_count']
            
            # 為了介面一致性，這裡的 name 放 username
            # color 可以隨機或固定，這裡暫時給一個預設色
            categories.append({
                'name': username, 
                'color': '#8B5CF6', # Purple for applicants
                'pending_count': count,
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
    """簡易撥料人員查看特定投料點（或申請人）的申請單"""
    if not is_simple_dispatcher(request.user) and not request.user.is_superuser:
        return redirect('requisitions:requisition_list')
    
    # 取得類型參數（成品/半成品）
    current_type = request.GET.get('type', 'finished')
    if current_type not in ['finished', 'semi_finished']:
        current_type = 'finished'
    
    today = timezone.now().date()
    
    if current_type == 'semi_finished':
        # 半成品：category 代表 applicant.username
        target_username = category
        category_color = '#8B5CF6' # Purple
        
        # 待撥料申請單 (依領料人過濾)
        pending_requisitions = Requisition.objects.filter(
            applicant__username=target_username,
            requisition_type='semi_finished',
            status__in=['demand_submitted', 'dispatch_in_progress']
        ).order_by('-created_at')
        
        # 已撥料申請單 (依領料人過濾)
        completed_requisitions = Requisition.objects.filter(
            applicant__username=target_username,
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
    
    # 計算每個申請單的逾期狀態和撥料進度
    for req in pending_requisitions:
        req.is_overdue = req.request_date < today
        items = req.items.all()
        total = items.count()
        dispatched = items.filter(dispatch_status='dispatched').count()
        req.progress = int((dispatched / total * 100) if total > 0 else 0)
        req.dispatched_count = dispatched
        req.total_count = total
    
    for req in completed_requisitions:
        items = req.items.all()
        total = items.count()
        dispatched = items.filter(dispatch_status='dispatched').count()
        req.progress = int((dispatched / total * 100) if total > 0 else 0)
        req.dispatched_count = dispatched
        req.total_count = total
    
    # 彙整所有待撥申請單中的缺料項目
    shortage_items = RequisitionItem.objects.filter(
        requisition__in=pending_requisitions,
        dispatch_status='backordered'
    ).values(
        'material_number', 'item_name', 'storage_bin'
    ).annotate(
        total_required=Sum('required_quantity')
    ).order_by('storage_bin', 'material_number')
    
    context = {
        'category': category,
        'category_color': category_color,
        'pending_requisitions': pending_requisitions,
        'completed_requisitions': completed_requisitions,
        'current_type': current_type,
        'shortage_items': shortage_items,
    }
    return render(request, 'requisitions/simple/simple_dispatcher_category.html', context)


@login_required
def simple_dispatcher_detail(request, category, pk):
    """簡易撥料人員撥退料操作頁面"""
    if not is_simple_dispatcher(request.user) and not request.user.is_superuser:
        return redirect('requisitions:requisition_list')
    
    requisition = get_object_or_404(Requisition, pk=pk)
    
    # Get sort parameter
    sort_param = request.GET.get('sort', 'material')
    if request.method == 'POST':
        sort_param = request.POST.get('sort', sort_param)
        
    # Apply sorting
    if sort_param == 'bin':
        items = requisition.items.all().order_by('storage_bin', 'material_number')
    elif sort_param == 'name':
        items = requisition.items.all().order_by('item_name', 'material_number')
    elif sort_param == 'status':
        # Status sort: Pending (None/Empty) -> Backordered -> Dispatched
        items = requisition.items.all().annotate(
            status_order=Case(
                When(dispatch_status__isnull=True, then=Value(1)),
                When(dispatch_status='', then=Value(1)),
                When(dispatch_status='backordered', then=Value(2)),
                When(dispatch_status='dispatched', then=Value(3)),
                default=Value(1),
                output_field=IntegerField(),
            )
        ).order_by('status_order', 'material_number')
    else:
        # Default: material number
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
        

        
        return redirect(f"{reverse('requisitions:simple_dispatcher_detail', kwargs={'category': category, 'pk': pk})}?sort={sort_param}")
    
    # 計算進度
    total = items.count()
    dispatched = items.filter(dispatch_status='dispatched').count()
    progress = int((dispatched / total * 100) if total > 0 else 0)
    
    # --- 檢查更早的未撥需求 (優先工單警示) ---
    backlog_map = {}
    target_material_numbers = list(items.values_list('material_number', flat=True))
    
    if target_material_numbers:
        # 1. 找出相同物料在其他工單的未撥需求（欠料大於 0）
        from ..models import WorkOrderMaterial
        
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
    
    # 即時更新庫存（從 Material 表取得最新庫存數量）
    from inventory.models import Material as InvMaterial
    material_codes = [item.material_number for item in items_list]
    live_stock = dict(
        InvMaterial.objects.filter(material_code__in=material_codes)
        .values_list('material_code', 'system_quantity')
    )
    for item in items_list:
        live_qty = live_stock.get(item.material_number)
        if live_qty is not None:
            item.stock_quantity = live_qty
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
            
        # 檢查庫存是否不足 (針對主項目)
        item.is_insufficient_stock = False
        if item.stock_quantity is not None and item.required_quantity is not None:
             # 如果尚未撥料且庫存 < 需求，標記為庫存不足
             if item.dispatch_status != 'dispatched' and item.stock_quantity < item.required_quantity:
                 item.is_insufficient_stock = True
                 # 嘗試取得預計入料日期
                 if item.source_material:
                     item.expected_date = item.source_material.estimated_arrival_date
                 else:
                     item.expected_date = None
    
    # 取得類型參數
    current_type = request.GET.get('type', 'finished')
    if current_type not in ['finished', 'semi_finished']:
        current_type = 'finished'
    
    # 查詢機型
    machine_model_name = ''
    wom = WorkOrderMaterial.objects.filter(
        order_number=requisition.order_number,
        machine_model__isnull=False
    ).select_related('machine_model').first()
    if wom and wom.machine_model:
        machine_model_name = wom.machine_model.name

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
        'current_sort': sort_param,
        'machine_model_name': machine_model_name,
    }
    return render(request, 'requisitions/simple/simple_dispatcher_detail.html', context)
