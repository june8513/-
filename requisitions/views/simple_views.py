"""
簡易介面視圖 - 給一般申請人員和撥料人員使用
"""
from django.shortcuts import render, redirect, get_object_or_404
from django.urls import reverse
from django.contrib.auth.decorators import login_required
from django.contrib import messages
from django.db.models import Q
from django.contrib.auth.models import User
from django.utils import timezone
from django.db import transaction
from django.db.models import Prefetch, F, Exists, OuterRef, Case, When, Value, BooleanField, Q, Sum, DecimalField, Count, IntegerField, CharField
from django.db.models.functions import Coalesce
from django.http import JsonResponse, HttpResponse
from django.views.decorators.http import require_POST
from decimal import Decimal
from datetime import date
import pandas as pd
import io
import traceback

from ..models import Requisition, RequisitionItem, WorkOrderMaterial, WorkOrder, ProcessType, RequisitionShareGroup, Announcement, MachineModel, WorkOrderMaterialTransaction
from ..constants import GROUP_NAMES, PROCESS_CATEGORY_NAMES, PROCESS_CATEGORY_COLORS
from ..forms import RequisitionForm
from inventory.models import Material
from ..utils import get_sap_user, _update_requisition_alert


def is_simple_applicant(user):
    """檢查是否為簡易申請人員（包含主管與管理員）"""
    return user.groups.filter(name__in=[GROUP_NAMES['APPLICANT'], GROUP_NAMES['APPLICANT_SUPERVISOR']]).exists() or \
           user.is_superuser


def is_simple_dispatcher(user):
    """檢查是否為簡易撥料人員（包含主管與管理員）"""
    return user.groups.filter(name__in=[GROUP_NAMES['DISPATCHER'], GROUP_NAMES['DISPATCHER_SUPERVISOR']]).exists() or \
           user.is_superuser


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
    
    # 取得被授權查看的使用者列表 (同一共享群組的成員)
    share_groups = request.user.requisition_share_groups.all()
    viewable_owners = User.objects.filter(requisition_share_groups__in=share_groups).distinct()
    viewable_owner_names = list(viewable_owners.exclude(pk=request.user.pk).values_list('username', flat=True))
    
    base_qs = Requisition.objects.filter(
        Q(applicant=request.user) | Q(applicant__in=viewable_owners),
        requisition_type=current_type
    )
    

    
    # 依狀態分類
    pending_reqs = list(base_qs.filter(status='demand_submitted', is_archived=False).order_by('-created_at'))
    in_progress_reqs = list(base_qs.filter(status='dispatch_in_progress', is_archived=False).order_by('-created_at'))
    completed_reqs = list(base_qs.filter(status='dispatch_completed', is_archived=False).order_by('-updated_at'))
    signed_off_reqs = list(base_qs.filter(status='signed_off', is_archived=False).order_by('-updated_at')[:20])
    archived_reqs = list(base_qs.filter(is_archived=True).order_by('-updated_at')[:20])
    
    # 計算進度和逾期
    today = timezone.now().date()
    for req_list in [pending_reqs, in_progress_reqs, completed_reqs, signed_off_reqs, archived_reqs]:
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
            # 標記是否為他人的申請單
            req.is_from_other = (req.applicant != request.user)
    
    context = {
        'pending_reqs': pending_reqs,
        'in_progress_reqs': in_progress_reqs,
        'completed_reqs': completed_reqs,
        'signed_off_reqs': signed_off_reqs,
        'archived_reqs': archived_reqs,
        'user': request.user,
        'current_type': current_type,
        'viewable_owner_names': viewable_owner_names,
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
                        'request_date': default_request_date,
                        'remarks': request.POST.get('remarks', '')
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
                remarks_list = request.POST.getlist('remarks')

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
                        remark = remarks_list[i] if i < len(remarks_list) else ''

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
                                    request_date=request_date_str or date.today(),
                                    remarks=remark
                                )
                                
                                # Auto-add semi-finished materials
                                materials_to_add = WorkOrderMaterial.objects.filter(
                                    order_number=order_number,
                                    material_type='semi_finished',
                                    is_active=True
                                )
                                
                                items_to_create = []
                                for material in materials_to_add:
                                    # Check if pre-dispatched via fast-dispatch OR SAP
                                    pre_confirmed = material.confirmed_quantity or Decimal('0')
                                    sap_withdrawn = material.sap_withdrawn_quantity or Decimal('0')
                                    
                                    # 依據使用者需求：如果系統 < SAP，則更新為 SAP 數量
                                    final_confirmed = pre_confirmed
                                    dispatched_by = None
                                    dispatched_at = None
                                    
                                    if sap_withdrawn > pre_confirmed:
                                        final_confirmed = sap_withdrawn
                                        dispatched_by = get_sap_user()
                                        dispatched_at = timezone.now()
                                    
                                    is_dispatched = final_confirmed > 0

                                    # 嘗試抓取即時庫存與儲格 (半成品)
                                    main_material = Material.objects.filter(material_code=material.material_number).first()
                                    stock_quantity = main_material.system_quantity if main_material else Decimal('0')
                                    storage_bin = main_material.bin if main_material else ''

                                    items_to_create.append(RequisitionItem(
                                        requisition=requisition,
                                        source_material=material,
                                        order_number=material.order_number,
                                        material_number=material.material_number,
                                        item_name=material.item_name,
                                        required_quantity=material.required_quantity,
                                        stock_quantity=stock_quantity,
                                        storage_bin=storage_bin,
                                        confirmed_quantity=final_confirmed,
                                        dispatch_status='dispatched' if is_dispatched else None,
                                        dispatched_by=dispatched_by,
                                        dispatched_at=dispatched_at
                                    ))
                                if items_to_create:
                                    RequisitionItem.objects.bulk_create(items_to_create)
                                    
                                    # Update requisition status if there are pre-dispatched items
                                    dispatched_count = sum(1 for item in items_to_create if item.dispatch_status == 'dispatched')
                                    if dispatched_count > 0:
                                        if dispatched_count == len(items_to_create):
                                            requisition.status = 'dispatch_completed'
                                        else:
                                            requisition.status = 'dispatch_in_progress'
                                        requisition.save()
                                
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
            remarks = request.POST.get('remarks')
            
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
                    request_date=request_date or date.today(),
                    remarks=remarks
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

                    # SAP 繼承邏輯
                    pre_confirmed = material.confirmed_quantity or Decimal('0')
                    sap_withdrawn = material.sap_withdrawn_quantity or Decimal('0')
                    
                    final_confirmed = pre_confirmed
                    dispatched_by = None
                    dispatched_at = None
                    
                    if sap_withdrawn > pre_confirmed:
                        final_confirmed = sap_withdrawn
                        dispatched_by = get_sap_user()
                        dispatched_at = timezone.now()

                    is_dispatched = final_confirmed > 0 and final_confirmed >= material.required_quantity

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
                            confirmed_quantity=final_confirmed,
                            dispatch_status='dispatched' if is_dispatched else None,
                            dispatched_by=dispatched_by,
                            dispatched_at=dispatched_at
                        )
                    )
                
                if items_to_create:
                    RequisitionItem.objects.bulk_create(items_to_create)
                    
                    # Update requisition status if there are pre-dispatched items
                    dispatched_count = sum(1 for item in items_to_create if item.dispatch_status == 'dispatched')
                    if dispatched_count > 0:
                        if dispatched_count == len(items_to_create):
                            requisition.status = 'dispatch_completed'
                        else:
                            requisition.status = 'dispatch_in_progress'
                        requisition.save()
                    
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

    # 權限檢查：只能刪除自己的申請單（或是同一群組成員、主管、超級管理員）
    is_supervisor = request.user.groups.filter(name=GROUP_NAMES['APPLICANT_SUPERVISOR']).exists()
    is_group_member = RequisitionShareGroup.objects.filter(
        members=request.user
    ).filter(members=requisition.applicant).exists()
    
    if requisition.applicant != request.user and not is_supervisor and not request.user.is_superuser and not is_group_member:
        messages.error(request, '您無權限刪除此申請單。')
        return redirect('requisitions:simple_applicant_detail', pk=pk)

    # 狀態檢查：只能刪除 'demand_submitted' 狀態的申請單
    if requisition.status != 'demand_submitted':
        messages.error(request, '只能刪除尚未撥料的申請單。')
        return redirect('requisitions:simple_applicant_detail', pk=pk)

    # 檢查是否已歸檔
    if requisition.is_archived:
        messages.error(request, '此申請單所屬工單已歸檔，無法刪除。')
        return redirect('requisitions:simple_applicant_detail', pk=pk)

    if request.method == 'POST':
        requisition.delete()
        messages.success(request, '申請單已刪除。')
        return redirect('requisitions:simple_applicant_home')


@login_required
def simple_applicant_detail(request, pk):
    """簡易申請人員查看申請單詳情"""
    requisition = get_object_or_404(Requisition, pk=pk)
    
    # 檢查是否為申請人本人 (或主管 或同一群組成員)
    is_supervisor = request.user.groups.filter(name=GROUP_NAMES['APPLICANT_SUPERVISOR']).exists()
    is_group_member = RequisitionShareGroup.objects.filter(
        members=request.user
    ).filter(members=requisition.applicant).exists()
    
    if requisition.applicant != request.user and not is_supervisor and not request.user.is_superuser and not is_group_member:
        messages.error(request, "您沒有權限查看此申請單。")
        return redirect('requisitions:simple_applicant_home')
    
    items = requisition.items.all().order_by('material_number')
    
    # 計算進度
    total = items.count()
    dispatched = items.filter(dispatch_status='dispatched').count()
    progress = int((dispatched / total * 100) if total > 0 else 0)
    
    # 取得即時庫存資料 (避免顯示建立時的過期數據)
    material_numbers = [item.material_number for item in items]
    realtime_stocks = {m.material_code: m for m in Material.objects.filter(material_code__in=material_numbers)}
    
    # 標記缺料物料與即時庫存檢查
    for item in items:
        # 更新即時庫存與儲格資訊
        m_info = realtime_stocks.get(item.material_number)
        if m_info:
            item.stock_quantity = m_info.system_quantity
            item.storage_bin = m_info.bin
        
        if item.dispatch_status == 'backordered':
            item.is_shortage = True
            # 嘗試取得預計入料日期
            if item.source_material:
                item.expected_date = item.source_material.demand_date
            else:
                item.expected_date = None
        else:
            item.is_shortage = False
            
        # 檢查庫存是否不足 (使用即時庫存)
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

    # 檢查是否包含「已撥料需退料」的項目
    has_over_dispatched_items = False
    if requisition.has_alert:
        has_over_dispatched_items = items.filter(
            confirmed_quantity__gt=F('required_quantity'),
            alert_dismissed=False
        ).exists()

    context = {
        'requisition': requisition,
        'items': items,
        'progress': progress,
        'dispatched_count': dispatched,
        'total_count': total,
        'machine_model_name': machine_model_name,
        'has_over_dispatched_items': has_over_dispatched_items,
    }
    return render(request, 'requisitions/simple/simple_applicant_detail.html', context)


@login_required
def simple_applicant_update_process_type(request, pk):
    """申請人員修改投料點"""
    requisition = get_object_or_404(Requisition, pk=pk)
    is_ajax = request.headers.get('X-Requested-With') == 'XMLHttpRequest'
    
    # 檢查是否為申請人本人或管理員 (或主管 或同一群組成員)
    is_supervisor = request.user.groups.filter(name=GROUP_NAMES['APPLICANT_SUPERVISOR']).exists()
    is_group_member = RequisitionShareGroup.objects.filter(
        members=request.user
    ).filter(members=requisition.applicant).exists()
    
    if requisition.applicant != request.user and not is_supervisor and not request.user.is_superuser and not is_group_member:
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

    # 檢查是否已歸檔
    if requisition.is_archived:
        if is_ajax:
            return JsonResponse({'success': False, 'message': '此申請單所屬工單已歸檔，無法修改投料點。'})
        messages.error(request, "此申請單所屬工單已歸檔，無法修改投料點。")
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
    is_group_member = RequisitionShareGroup.objects.filter(
        members=request.user
    ).filter(members=requisition.applicant).exists()
    
    if requisition.applicant != request.user and not is_supervisor and not request.user.is_superuser and not is_group_member:
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

    # 檢查是否已歸檔
    if requisition.is_archived:
        message = "此申請單所屬工單已歸檔，無法修改需求日期。"
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
    
    # 檢查是否為申請人本人 (或同一群組成員)
    is_group_member = RequisitionShareGroup.objects.filter(
        members=request.user
    ).filter(members=requisition.applicant).exists()
    
    if requisition.applicant != request.user and not request.user.is_superuser and not is_group_member:
        if is_ajax:
            return JsonResponse({'success': False, 'message': '您沒有權限執行簽收操作。'})
        messages.error(request, "您沒有權限執行簽收操作。")
        return redirect('requisitions:simple_applicant_home')

    # 檢查是否已歸檔
    if requisition.is_archived:
        if is_ajax:
            return JsonResponse({'success': False, 'message': '此申請單所屬工單已歸檔，無法進行簽收操作。'})
        messages.error(request, "此申請單所屬工單已歸檔，無法進行簽收操作。")
        return redirect('requisitions:simple_applicant_detail', pk=pk)
    
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
                    is_signed_off=False,
                    has_issue=False
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
                            if not item.is_signed_off and not item.has_issue:
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
            all_items = requisition.items.all()
            all_signed = all_items.exists() and all(item.is_signed_off for item in all_items)
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

@login_required
@require_POST
def report_item_issue(request, item_id):
    """申請人回報物料異況"""
    item = get_object_or_404(RequisitionItem, pk=item_id)
    description = request.POST.get('description', '').strip()
    
    if not description:
        return JsonResponse({'success': False, 'message': '請填寫異況說明。'})
        
    if item.is_signed_off:
        return JsonResponse({'success': False, 'message': '此項目已簽收，無法回報異況。'})
        
    if item.dispatch_status != 'dispatched':
        return JsonResponse({'success': False, 'message': '此項目尚未完成撥料，無法回報異況。'})
        
    with transaction.atomic():
        from requisitions.models import RequisitionItemIssue
        RequisitionItemIssue.objects.create(
            requisition_item=item,
            reported_by=request.user,
            description=description
        )
        item.has_issue = True
        item.save()
        
    return JsonResponse({'success': True, 'message': '異況已回報。'})

@login_required
@require_POST
def resolve_item_issue(request, item_id):
    """撥料員解除物料異況"""
    item = get_object_or_404(RequisitionItem, pk=item_id)
    resolution_notes = request.POST.get('resolution_notes', '').strip()
    
    # 權限檢查：只有撥料員或管理員可以解除
    if not (is_simple_dispatcher(request.user) or request.user.is_superuser):
        return JsonResponse({'success': False, 'message': '您沒有權限解除異況。'})
        
    with transaction.atomic():
        from requisitions.models import RequisitionItemIssue
        # 標記所有未解決的異況為已解決
        unresolved_issues = item.issues.filter(is_resolved=False)
        for issue in unresolved_issues:
            issue.is_resolved = True
            issue.resolved_by = request.user
            issue.resolved_at = timezone.now()
            issue.resolution_notes = resolution_notes
            issue.save()
            
        item.has_issue = False
        item.save()
        
    return JsonResponse({'success': True, 'message': '異況已解除。'})

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
        from datetime import timedelta
        alert_threshold = timezone.now() - timedelta(hours=24)
        
        for category_name in PROCESS_CATEGORY_NAMES:
            pending_count = Requisition.objects.filter(
                process_type__icontains=category_name,
                status__in=['demand_submitted', 'dispatch_in_progress']
            ).count()
            
            # 檢查 SAP 扣帳異常 (> 24 小時)
            has_sap_issue = WorkOrderMaterial.objects.filter(
                process_type__name__icontains=category_name,
                sap_sync_issue=True,
                sap_sync_issue_since__lte=alert_threshold,
                is_active=True
            ).exists()
            
            categories.append({
                'name': category_name,
                'color': PROCESS_CATEGORY_COLORS.get(category_name, '#6B7280'),
                'pending_count': pending_count,
                'has_sap_issue': has_sap_issue,
            })
    
    # 取得所有最新且未過期的公告
    now = timezone.now()
    announcements = Announcement.objects.filter(
        Q(is_active=True) & (Q(expires_at__gt=now) | Q(expires_at__isnull=True))
    ).order_by('-created_at')
    
    can_publish = request.user.is_superuser or (hasattr(request.user, 'profile') and request.user.profile.can_publish_announcements)
    
    context = {
        'categories': categories,
        'user': request.user,
        'current_type': current_type,
        'announcements': announcements,
        'can_publish': can_publish,
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
    
    # 處理快速撥料（補料）請求
    if request.method == 'POST' and request.POST.get('action') == 'quick_dispatch':
        item_pk = request.POST.get('item_pk')
        try:
            item = get_object_or_404(RequisitionItem, pk=item_pk)
            # 將數量設為需求量，狀態改為已撥料
            item.confirmed_quantity = item.required_quantity
            item.dispatch_status = 'dispatched'
            item.save()
            # 更新申請單狀態
            req = item.requisition
            all_items = req.items.all()
            dispatched = all_items.filter(dispatch_status='dispatched').count()
            if dispatched == all_items.count():
                req.status = 'dispatch_completed'
            elif dispatched > 0:
                req.status = 'dispatch_in_progress'
            req.save()
            return JsonResponse({'success': True})
        except Exception as e:
            return JsonResponse({'success': False, 'message': str(e)})
    
    if current_type == 'semi_finished':
        # 半成品：category 代表 applicant.username
        target_username = category
        category_color = '#8B5CF6' # Purple
        
        # 待撥料申請單 (依領料人過濾)
        pending_requisitions = Requisition.objects.filter(
            applicant__username=target_username,
            requisition_type='semi_finished',
            status__in=['demand_submitted', 'dispatch_in_progress'],
            is_archived=False
        ).order_by('request_date', '-created_at')
        
        # 已撥料申請單 (依領料人過濾)
        completed_requisitions = Requisition.objects.filter(
            applicant__username=target_username,
            requisition_type='semi_finished',
            status__in=['dispatch_completed', 'signed_off'],
            is_archived=False
        ).order_by('-updated_at')[:20]
        
        # 已歸檔申請單 (依領料人過濾)
        archived_requisitions = Requisition.objects.filter(
            applicant__username=target_username,
            requisition_type='semi_finished',
            is_archived=True
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
            status__in=['demand_submitted', 'dispatch_in_progress'],
            is_archived=False
        ).order_by('request_date', '-created_at')
        
        # 已撥料申請單
        completed_requisitions = Requisition.objects.filter(
            process_type__icontains=category,
            status__in=['dispatch_completed', 'signed_off'],
            is_archived=False
        ).order_by('-updated_at')[:20]

        # 已歸檔申請單
        archived_requisitions = Requisition.objects.filter(
            process_type__icontains=category,
            is_archived=True
        ).order_by('-updated_at')[:20]
    
    # 計算每個申請單的逾期狀態和撥料進度
    from datetime import timedelta
    alert_threshold = timezone.now() - timedelta(hours=24)
    
    for req in pending_requisitions:
        req.is_overdue = req.request_date < today
        items = req.items.all()
        total = items.count()
        dispatched = items.filter(dispatch_status='dispatched').count()
        req.progress = int((dispatched / total * 100) if total > 0 else 0)
        req.dispatched_count = dispatched
        req.total_count = total
        req.undispatched_count = total - dispatched
        
        # 檢查 SAP 扣帳異常
        req.has_sap_issue = items.filter(
            source_material__sap_sync_issue=True,
            source_material__sap_sync_issue_since__lte=alert_threshold
        ).exists()
    
    for req in completed_requisitions:
        items = req.items.all()
        total = items.count()
        dispatched = items.filter(dispatch_status='dispatched').count()
        req.progress = int((dispatched / total * 100) if total > 0 else 0)
        req.dispatched_count = dispatched
        req.total_count = total

    for req in archived_requisitions:
        items = req.items.all()
        total = items.count()
        dispatched = items.filter(dispatch_status='dispatched').count()
        req.progress = int((dispatched / total * 100) if total > 0 else 0)
        req.dispatched_count = dispatched
        req.total_count = total
    
    # 取得所有待撥申請單中的缺料項目 (不彙整，以便單獨補料)
    shortage_items = RequisitionItem.objects.filter(
        requisition__in=pending_requisitions,
        dispatch_status='backordered'
    ).order_by('storage_bin', 'order_number', 'material_number')
    
    context = {
        'category': category,
        'category_color': category_color,
        'pending_requisitions': pending_requisitions,
        'completed_requisitions': completed_requisitions,
        'archived_requisitions': archived_requisitions,
        'current_type': current_type,
        'shortage_items': shortage_items,
    }
    return render(request, 'requisitions/simple/simple_dispatcher_category.html', context)


@login_required
def simple_dispatcher_shortage(request, category):
    """待撥欠料彙整 - 獨立網頁頁面呈現該分類下的所有欠料物料"""
    if not is_simple_dispatcher(request.user) and not request.user.is_superuser:
        return redirect('requisitions:requisition_list')

    current_type = request.GET.get('type', 'finished')
    category_color = PROCESS_CATEGORY_COLORS.get(category, '#6B7280')

    # 取得該分類下的待處理申請單
    pending_requisitions = Requisition.objects.filter(
        process_type__icontains=category,
        status__in=['demand_submitted', 'dispatch_in_progress'],
        is_archived=False
    )
    
    # 取得所有待撥申請單中的缺料項目
    shortage_items = RequisitionItem.objects.filter(
        requisition__in=pending_requisitions,
        dispatch_status='backordered'
    ).select_related('requisition', 'source_material').order_by('storage_bin', 'order_number', 'material_number')
    
    context = {
        'category': category,
        'category_color': category_color,
        'shortage_items': shortage_items,
        'current_type': current_type,
    }
    return render(request, 'requisitions/simple/simple_dispatcher_shortage.html', context)


@login_required
@require_POST
def simple_dispatch_item_ajax(request, item_id):
    """通用單項撥料 AJAX 處理"""
    from django.http import JsonResponse
    try:
        item = RequisitionItem.objects.get(pk=item_id)
        
        # 標記為已撥料
        item.confirmed_quantity = item.required_quantity
        item.dispatch_status = 'dispatched'
        item.dispatched_by = request.user
        item.dispatched_at = timezone.now()
        item.save()
        
        # 更新申請單狀態
        requisition = item.requisition
        items = requisition.items.all()
        dispatched_count = items.filter(dispatch_status='dispatched').count()
        if dispatched_count == items.count():
            requisition.status = 'dispatch_completed'
        else:
            requisition.status = 'dispatch_in_progress'
        requisition.save()
        
        return JsonResponse({'success': True, 'message': '撥料成功'})
    except RequisitionItem.DoesNotExist:
        return JsonResponse({'success': False, 'message': '找不到指定的物料項目'})
    except Exception as e:
        error_msg = traceback.format_exc()
        print(f"Replenishment error: {error_msg}")
        return JsonResponse({'success': False, 'message': f'系統錯誤: {str(e)}'})


@login_required
def simple_dispatcher_merge(request, category):
    """合併撥料 - 將多張工單的物料合併呈現，集中撥料"""
    if not is_simple_dispatcher(request.user) and not request.user.is_superuser:
        return redirect('requisitions:requisition_list')

    current_type = request.GET.get('type', 'finished')
    order_numbers = request.GET.getlist('orders')

    if not order_numbers:
        messages.error(request, '請選擇至少一張工單。')
        return redirect('requisitions:simple_dispatcher_category', category=category)

    # 處理 AJAX 撥料/缺料請求
    if request.method == 'POST':
        is_ajax = request.headers.get('X-Requested-With') == 'XMLHttpRequest'
        action = request.POST.get('action')

        if action == 'dispatch_item':
            item_pk = request.POST.get('item_pk')
            try:
                item = get_object_or_404(RequisitionItem, pk=item_pk)
                item.confirmed_quantity = item.required_quantity
                item.dispatch_status = 'dispatched'
                item.save()
                # 更新申請單狀態
                req = item.requisition
                all_items = req.items.all()
                dispatched = all_items.filter(dispatch_status='dispatched').count()
                if dispatched == all_items.count():
                    req.status = 'dispatch_completed'
                elif dispatched > 0:
                    req.status = 'dispatch_in_progress'
                req.save()
                return JsonResponse({'success': True, 'message': '已撥料'})
            except Exception as e:
                return JsonResponse({'success': False, 'message': str(e)})

        elif action == 'backorder_item':
            item_pk = request.POST.get('item_pk')
            try:
                item = get_object_or_404(RequisitionItem, pk=item_pk)
                item.dispatch_status = 'backordered'
                item.save()
                return JsonResponse({'success': True, 'message': '已標記缺料'})
            except Exception as e:
                return JsonResponse({'success': False, 'message': str(e)})

        elif action == 'dispatch_material':
            material_number = request.POST.get('material_number')
            try:
                items = RequisitionItem.objects.filter(
                    (Q(requisition__applicant__username=category) if current_type == 'semi_finished' else Q(requisition__process_type__icontains=category)),
                    requisition__order_number__in=order_numbers,
                    material_number=material_number,
                    requisition__is_archived=False
                ).exclude(dispatch_status='dispatched')
                
                count = 0
                affected_reqs = set()
                for item in items:
                    item.confirmed_quantity = item.required_quantity
                    item.dispatch_status = 'dispatched'
                    item.save()
                    affected_reqs.add(item.requisition_id)
                    count += 1

                # 更新所有受影響的申請單狀態
                for req in Requisition.objects.filter(pk__in=affected_reqs):
                    all_items = req.items.all()
                    dispatched = all_items.filter(dispatch_status='dispatched').count()
                    if dispatched == all_items.count():
                        req.status = 'dispatch_completed'
                    elif dispatched > 0:
                        req.status = 'dispatch_in_progress'
                    req.save()

                return JsonResponse({'success': True, 'message': f'已將物料 {material_number} 的 {count} 筆需求完成撥料'})
            except Exception as e:
                return JsonResponse({'success': False, 'message': str(e)})

        elif action == 'backorder_material':
            material_number = request.POST.get('material_number')
            try:
                items = RequisitionItem.objects.filter(
                    (Q(requisition__applicant__username=category) if current_type == 'semi_finished' else Q(requisition__process_type__icontains=category)),
                    requisition__order_number__in=order_numbers,
                    material_number=material_number,
                    requisition__is_archived=False
                ).exclude(dispatch_status='dispatched')
                
                count = 0
                for item in items:
                    item.dispatch_status = 'backordered'
                    item.save()
                    count += 1

                return JsonResponse({'success': True, 'message': f'已將物料 {material_number} 的 {count} 筆需求標記為缺料'})
            except Exception as e:
                return JsonResponse({'success': False, 'message': str(e)})

        return JsonResponse({'success': False, 'message': '未知操作'})

    # GET: 查詢所選工單中的未撥料物料
    # 根據類型決定過濾條件
    if current_type == 'semi_finished':
        filter_q = Q(requisition__applicant__username=category)
    else:
        filter_q = Q(requisition__process_type__icontains=category)

    undispatched_items = RequisitionItem.objects.filter(
        filter_q,
        requisition__order_number__in=order_numbers,
        requisition__is_archived=False
    ).exclude(dispatch_status='dispatched').select_related('requisition')

    sort_param = request.GET.get('sort', 'bin')
    if sort_param == 'material':
        sort_args = ['material_number']
    elif sort_param == 'name':
        sort_args = ['item_name', 'material_number']
    else:
        sort_args = ['storage_bin', 'material_number']
        sort_param = 'bin'

    # 取得即時儲位資訊 (避免顯示過期或空白數據)
    from inventory.models import Material as InvMaterial
    material_codes = list(undispatched_items.values_list('material_number', flat=True).distinct())
    real_bins = dict(InvMaterial.objects.filter(material_code__in=material_codes).values_list('material_code', 'bin'))

    # 按物料編號分組合併
    from collections import OrderedDict
    merged = OrderedDict()
    for item in undispatched_items:
        key = item.material_number
        if key not in merged:
            # 優先使用即時儲位
            live_bin = real_bins.get(key)
            merged[key] = {
                'material_number': item.material_number,
                'item_name': item.item_name,
                'storage_bin': live_bin if live_bin else (item.storage_bin or '-'),
                'orders': [],
                'total_qty': 0,
            }
        merged[key]['orders'].append({
            'pk': item.pk,
            'order_number': item.requisition.order_number,
            'required_quantity': item.required_quantity,
            'status': item.dispatch_status,
            'request_date': item.requisition.request_date,
        })
        merged[key]['total_qty'] += item.required_quantity

    from datetime import date
    # Ensure orders within each material are sorted by request_date
    for mat_data in merged.values():
        mat_data['orders'] = sorted(mat_data['orders'], key=lambda x: x['request_date'] or date.today())

    category_color = PROCESS_CATEGORY_COLORS.get(category, '#6B7280')

    merged_items_list = list(merged.values())
    if sort_param == 'material':
        merged_items_list.sort(key=lambda x: x['material_number'] or '')
    elif sort_param == 'name':
        merged_items_list.sort(key=lambda x: (x['item_name'] or '', x['material_number'] or ''))
    else:
        merged_items_list.sort(key=lambda x: (x['storage_bin'] or '', x['material_number'] or ''))

    context = {
        'category': category,
        'category_color': category_color,
        'current_type': current_type,
        'order_numbers': order_numbers,
        'merged_items': merged_items_list,
        'total_items': undispatched_items.count(),
        'sort_param': sort_param,
    }
    return render(request, 'requisitions/simple/simple_dispatcher_merge.html', context)

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
        items = requisition.items.all().select_related('dispatched_by').annotate(
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
        items = requisition.items.all().select_related('dispatched_by').order_by('material_number')
    
    if request.method == 'POST':
        action = request.POST.get('action')
        item_pk = request.POST.get('item_pk')
        is_ajax = request.headers.get('X-Requested-With') == 'XMLHttpRequest'
        
        result = {'success': False, 'message': ''}

        # 檢查是否已歸檔
        if requisition.is_archived:
            result = {'success': False, 'message': '此申請單所屬工單已歸檔，無法進行任何操作。'}
            if is_ajax:
                return JsonResponse(result)
            messages.error(request, result['message'])
            return redirect(f"{reverse('requisitions:simple_dispatcher_detail', kwargs={'category': category, 'pk': pk})}?sort={sort_param}")
        
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
                        item.dispatched_by = request.user
                        item.dispatched_at = timezone.now()
                        item.save()
                        
                        # 自定義顯示名稱
                        dispatcher_display_name = f"{request.user.first_name}{request.user.last_name}"
                        if not (request.user.first_name or request.user.last_name):
                            dispatcher_display_name = request.user.username

                        result = {
                            'success': True, 
                            'message': f'物料 {item.material_number} 撥料 {dispatched_qty} 成功。', 
                            'new_status': 'dispatched', 
                            'dispatched_qty': str(dispatched_qty),
                            'dispatched_by_name': dispatcher_display_name,
                            'dispatched_by_id': request.user.id
                        }
                        if not is_ajax:
                            messages.success(request, result['message'])
                    except Exception as e:
                        result = {'success': False, 'message': f'撥料失敗：{str(e)}'}
                        if not is_ajax:
                            messages.error(request, result['message'])
                
                elif action == 'undo' or action == 'return':
                    # 退料/取消撥料
                    # 正常情況下已簽收不能撤銷，但如果是「物料已刪除」或「需求變更導致多撥」，應允許退料
                    is_deactivated = item.source_material and not item.source_material.is_active
                    has_surplus = item.confirmed_quantity > item.required_quantity
                    
                    if item.is_signed_off and not (is_deactivated or has_surplus) and action == 'undo':
                        result = {'success': False, 'message': f'物料 {item.material_number} 已簽收，無法撤銷。'}
                        if not is_ajax:
                            messages.error(request, result['message'])
                    else:
                        old_qty = item.confirmed_quantity or Decimal('0')
                        item.confirmed_quantity = Decimal('0')
                        item.dispatch_status = None
                        item.dispatched_by = None
                        item.dispatched_at = None
                        item.save()
                        
                        # 記錄交易 (如果是退料)
                        if old_qty > 0:
                            # 新增：記錄到物料變更歷程
                            user_display_name = f"{request.user.first_name}{request.user.last_name}"
                            if not user_display_name:
                                user_display_name = request.user.username
                            
                            log_msg = f"執行退料 {item.material_number} (退料者：{user_display_name}，退料數量：{old_qty})"
                            _update_requisition_alert(requisition.order_number, requisition.process_type, log_msg)
                            requisition.refresh_from_db()

                            if item.source_material:
                                from requisitions.models import WorkOrderMaterialTransaction
                                WorkOrderMaterialTransaction.objects.create(
                                    work_order_material=item.source_material,
                                    user=request.user,
                                    transaction_type='RETURN',
                                    quantity_change=-old_qty,
                                    new_confirmed_quantity=Decimal('0'),
                                    notes=f"簡易畫面執行退料 (物料狀態: {'已刪除' if is_deactivated else '正常'})"
                                )

                        # 如果是已刪除的物料，歸零後直接移除 RequisitionItem
                        if is_deactivated:
                            item.delete()
                            result = {'success': True, 'message': f'物料 {item.material_number} 已成功退料並從清單移除。', 'new_status': 'deleted'}
                        else:
                            result = {'success': True, 'message': f'物料 {item.material_number} 已成功退料。', 'new_status': 'pending'}
                        
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
                elif action == 'backorder':
                    # 標記為缺料
                    item.dispatch_status = 'backordered'
                    item.confirmed_quantity = Decimal('0')
                    item.save()
                    result = {'success': True, 'message': f'物料 {item.material_number} 已標記為缺料。', 'new_status': 'backordered'}
                    if not is_ajax:
                        messages.warning(request, result['message'])
                
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
    
    # 即時更新庫存與預計入料日期（從 Material 與 WorkOrderMaterial 取得最新數據）
    from inventory.models import Material as InvMaterial
    material_codes = [item.material_number for item in items_list]
    
    # 庫存對照表 (包含儲位)
    live_inventory = {
        m[0]: {'qty': m[1], 'bin': m[2]}
        for m in InvMaterial.objects.filter(material_code__in=material_codes).values_list('material_code', 'system_quantity', 'bin')
    }
    
    # 預計入料日期對照表 (從 WorkOrderMaterial 取得最新的日期)
    from django.db.models import Max
    arrival_dates = dict(
        WorkOrderMaterial.objects.filter(material_number__in=material_codes, is_active=True)
        .values('material_number')
        .annotate(latest_date=Max('estimated_arrival_date'))
        .values_list('material_number', 'latest_date')
    )

    for item in items_list:
        # 庫存與儲位
        inv_info = live_inventory.get(item.material_number)
        if inv_info:
            item.stock_quantity = inv_info['qty']
            item.storage_bin = inv_info['bin']
            
        # 預計入料日期
        item.estimated_arrival_date = arrival_dates.get(item.material_number)
    for item in items_list:
        item.backlog_info = backlog_map.get(item.material_number, [])
        item.has_backlog = bool(item.backlog_info)
        # 預處理 CSS 類別
        if item.dispatch_status == 'dispatched':
            item.card_class = 'dispatched'
            qty_display = item.confirmed_quantity if item.confirmed_quantity else item.required_quantity
            # 取得撥料人員名稱
            dispatcher_name = ""
            if item.dispatched_by:
                dispatcher_name = f" ({item.dispatched_by.first_name}{item.dispatched_by.last_name}"
                if not (item.dispatched_by.first_name or item.dispatched_by.last_name):
                    dispatcher_name = f" ({item.dispatched_by.username}"
                dispatcher_name += ")"
            
            item.status_text = f'已撥 {qty_display}{dispatcher_name}'
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
                # 取得撥料人員名稱
                supp_dispatcher = ""
                if supp.dispatched_by:
                    supp_dispatcher = f" ({supp.dispatched_by.first_name}{supp.dispatched_by.last_name}"
                    if not (supp.dispatched_by.first_name or supp.dispatched_by.last_name):
                        supp_dispatcher = f" ({supp.dispatched_by.username}"
                    supp_dispatcher += ")"
                
                supp.status_text = f'已撥 {supp_qty}{supp_dispatcher}'
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

    # 如果是以儲位排序，需要在更新即時儲位後重新排序
    if sort_param == 'bin':
        items_list.sort(key=lambda x: (x.storage_bin or '', x.material_number or ''))
    
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

    # 檢查是否包含「已撥料需退料」的項目
    has_over_dispatched_items = False
    if requisition.has_alert:
        has_over_dispatched_items = items.filter(
            confirmed_quantity__gt=F('required_quantity'),
            alert_dismissed=False
        ).exists()

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
        'has_over_dispatched_items': has_over_dispatched_items,
    }
    return render(request, 'requisitions/simple/simple_dispatcher_detail.html', context)


@login_required
def update_announcement(request):
    """更新系統公告 (管理員或授權人員)"""
    can_publish = request.user.is_superuser or (hasattr(request.user, 'profile') and request.user.profile.can_publish_announcements)
    if not can_publish:
        return JsonResponse({'success': False, 'message': '權限不足'})
    
    if request.method == 'POST':
        action = request.POST.get('action', 'create')
        
        if action == 'delete':
            announcement_id = request.POST.get('announcement_id')
            if announcement_id:
                announcement = get_object_or_404(Announcement, id=announcement_id)
                announcement.is_active = False
                announcement.save()
                return JsonResponse({'success': True, 'message': '公告已刪除'})
            return JsonResponse({'success': False, 'message': '找不到公告'})

        content = request.POST.get('content', '').strip()
        if content:
            # 建立新的公告，而不是覆蓋舊的
            Announcement.objects.create(
                content=content,
                created_by=request.user,
                is_active=True
            )
            return JsonResponse({'success': True, 'message': '公告已發佈'})
        else:
            return JsonResponse({'success': False, 'message': '內容不能為空'})
    
    return JsonResponse({'success': False, 'message': '無效的請求'})


@login_required
def export_simple_applicant_requisitions_excel(request):
    """申請人員匯出 Excel"""
    if not is_simple_applicant(request.user) and not request.user.is_superuser:
        return redirect('requisitions:requisition_list')
    
    current_type = request.GET.get('type', 'finished')
    
    # 取得被授權查看的使用者列表
    share_groups = request.user.requisition_share_groups.all()
    viewable_owners = User.objects.filter(requisition_share_groups__in=share_groups).distinct()
    
    base_qs = Requisition.objects.filter(
        Q(applicant=request.user) | Q(applicant__in=viewable_owners)
    ).filter(requisition_type=current_type)
    
    # 如果是主管，顯示所有申請單
    is_supervisor = request.user.groups.filter(name=GROUP_NAMES['APPLICANT_SUPERVISOR']).exists()
    if is_supervisor:
        base_qs = Requisition.objects.filter(requisition_type=current_type)

    # 包含所有狀態
    requisitions = base_qs.order_by('-created_at')
    return _generate_simple_export_excel(requisitions, f"applicant_requisitions_{current_type}")


@login_required
def export_simple_dispatcher_requisitions_excel(request, category):
    """撥料人員匯出 Excel (特定分類)"""
    if not is_simple_dispatcher(request.user) and not request.user.is_superuser:
        return redirect('requisitions:requisition_list')
    
    current_type = request.GET.get('type', 'finished')
    order_numbers = request.GET.getlist('orders')
    
    if current_type == 'semi_finished':
        # 半成品：category 是 applicant.username
        qs = Requisition.objects.filter(
            applicant__username=category,
            requisition_type='semi_finished'
        )
    else:
        # 成品：category 是 process_type 的一部分
        qs = Requisition.objects.filter(
            process_type__icontains=category,
            requisition_type='finished'
        )
        
    if order_numbers:
        qs = qs.filter(order_number__in=order_numbers)
        
    requisitions = qs.order_by('-created_at')

    return _generate_simple_export_excel(requisitions, f"dispatcher_requisitions_{category}_{current_type}")

@login_required
def simple_dispatcher_fast_dispatch(request):
    """
    處理快速撥料輸入，進行解析與預覽
    """
    if not (request.user.is_superuser or request.user.groups.filter(name=GROUP_NAMES['DISPATCHER']).exists() or request.user.groups.filter(name=GROUP_NAMES['DISPATCHER_SUPERVISOR']).exists()):
        messages.error(request, "您沒有撥料權限。")
        return redirect('requisitions:simple_dispatcher_home')

    if request.method == 'POST':
        mode = request.POST.get('dispatch_mode', 'batch_mixed')
        parsed_items = []
        
        if mode == 'single_order':
            order_number = request.POST.get('single_order_number', '').strip()
            raw_materials = request.POST.get('material_numbers_data', '')
            if order_number:
                lines = [line.strip() for line in raw_materials.splitlines() if line.strip()]
                for line in lines:
                    import re
                    # 允許逗號或空白分隔同一列貼上的多個品號
                    parts = [p.strip() for p in re.split(r'[\t ,]+', line) if p.strip()]
                    for mat_no in parts:
                        parsed_items.append({
                            'order_number': order_number,
                            'material_number': mat_no
                        })
        else:
            raw_text = request.POST.get('fast_dispatch_data', '')
            lines = [line.strip() for line in raw_text.splitlines() if line.strip()]
            
            for line in lines:
                import re
                parts = [p.strip() for p in re.split(r'[\t ,]+', line) if p.strip()]
                if len(parts) >= 2:
                    order_number = parts[0]
                    material_number = parts[1]
                    parsed_items.append({
                        'order_number': order_number,
                        'material_number': material_number
                    })
        
        ready_items = []
        already_full_items = []
        conflict_items = []
        not_found_items = []
        
        from collections import defaultdict
        grouped_items = defaultdict(list)
        for item in parsed_items:
            grouped_items[(item['order_number'], item['material_number'])].append(item)
            
        for (order_no, mat_no), items in grouped_items.items():
            woms = WorkOrderMaterial.objects.filter(
                order_number=order_no, 
                material_number=mat_no, 
                is_active=True
            ).select_related('process_type')
            
            wom_count = woms.count()
            if wom_count == 0:
                not_found_items.append({'order_number': order_no, 'material_number': mat_no})
            elif wom_count == 1:
                wom = woms.first()
                # 檢查 WOM 本身的 confirmed_quantity
                wom_confirmed = wom.confirmed_quantity or Decimal('0')
                # 也檢查 RequisitionItem 的 confirmed_quantity（一般撥料流程只更新 RequisitionItem）
                req_item_confirmed = RequisitionItem.objects.filter(
                    source_material=wom
                ).order_by('-confirmed_quantity').values_list('confirmed_quantity', flat=True).first() or Decimal('0')
                # 取兩者中較大值作為實際已撥數量
                confirmed = max(wom_confirmed, req_item_confirmed)
                is_full = confirmed >= wom.required_quantity
                
                if is_full:
                    already_full_items.append({
                        'order_number': order_no,
                        'material_number': mat_no,
                        'item_name': wom.item_name,
                        'process_type': wom.process_type.name if wom.process_type else '未指定',
                        'required_quantity': wom.required_quantity,
                        'confirmed_quantity': confirmed,
                    })
                else:
                    ready_items.append({
                        'order_number': order_no,
                        'material_number': mat_no,
                        'item_name': wom.item_name,
                        'process_type': wom.process_type.name if wom.process_type else '未指定',
                        'required_quantity': wom.required_quantity,
                        'confirmed_quantity': confirmed,
                        'wom_id': wom.id
                    })
            else:
                candidates = []
                all_full = True
                for wom in woms:
                    wom_confirmed = wom.confirmed_quantity or Decimal('0')
                    req_item_confirmed = RequisitionItem.objects.filter(
                        source_material=wom
                    ).order_by('-confirmed_quantity').values_list('confirmed_quantity', flat=True).first() or Decimal('0')
                    confirmed = max(wom_confirmed, req_item_confirmed)
                    is_full = confirmed >= wom.required_quantity
                    if not is_full:
                        all_full = False
                    candidates.append({
                        'wom_id': wom.id,
                        'process_type': wom.process_type.name if wom.process_type else '未指定',
                        'required_quantity': wom.required_quantity,
                        'confirmed_quantity': confirmed,
                        'is_full': is_full,
                    })
                if all_full:
                    already_full_items.append({
                        'order_number': order_no,
                        'material_number': mat_no,
                        'item_name': woms.first().item_name,
                        'process_type': ', '.join([c['process_type'] for c in candidates]),
                        'required_quantity': sum(c['required_quantity'] for c in candidates),
                        'confirmed_quantity': sum(c['confirmed_quantity'] for c in candidates),
                    })
                else:
                    conflict_items.append({
                        'order_number': order_no,
                        'material_number': mat_no,
                        'item_name': woms.first().item_name,
                        'candidates': candidates
                    })
                
        return render(request, 'requisitions/simple/simple_fast_dispatch_preview.html', {
            'ready_items': ready_items,
            'already_full_items': already_full_items,
            'conflict_items': conflict_items,
            'not_found_items': not_found_items,
        })
        
    return render(request, 'requisitions/simple/simple_fast_dispatch.html')


@login_required
@require_POST
def simple_dispatcher_fast_dispatch_execute(request):
    """
    執行快速撥料
    """
    if not (request.user.is_superuser or request.user.groups.filter(name=GROUP_NAMES['DISPATCHER']).exists() or request.user.groups.filter(name=GROUP_NAMES['DISPATCHER_SUPERVISOR']).exists()):
        messages.error(request, "您沒有撥料權限。")
        return redirect('requisitions:simple_dispatcher_home')

    wom_ids = request.POST.getlist('wom_ids')
    if not wom_ids:
        messages.error(request, "未選擇任何有效的物料項目。")
        return redirect('requisitions:simple_dispatcher_home')
        
    success_count = 0
    try:
        with transaction.atomic():
            woms = WorkOrderMaterial.objects.filter(id__in=wom_ids, is_active=True)
            for wom in woms:
                qty_to_dispatch = wom.required_quantity
                
                if wom.confirmed_quantity is None:
                    wom.confirmed_quantity = Decimal('0')
                    
                if wom.confirmed_quantity >= wom.required_quantity:
                    continue
                    
                qty_change = qty_to_dispatch - (wom.confirmed_quantity or Decimal('0'))
                
                # 更新 WOM
                wom.confirmed_quantity = qty_to_dispatch
                wom.save()
                
                # 新增交易紀錄
                WorkOrderMaterialTransaction.objects.create(
                    work_order_material=wom,
                    user=request.user,
                    transaction_type='ALLOCATION',
                    quantity_change=qty_change,
                    new_confirmed_quantity=wom.confirmed_quantity,
                    notes="快速撥料自動寫入"
                )
                
                # 更新已存在的申請單項目
                req_items = RequisitionItem.objects.filter(source_material=wom)
                for req_item in req_items:
                    req_item.confirmed_quantity = qty_to_dispatch
                    req_item.dispatch_status = 'dispatched'
                    req_item.dispatched_by = request.user
                    req_item.dispatched_at = timezone.now()
                    req_item.save()
                    
                    req = req_item.requisition
                    if req:
                        all_dispatched = not req.items.exclude(dispatch_status='dispatched').exists()
                        has_dispatched = req.items.filter(dispatch_status='dispatched').exists()
                        if all_dispatched:
                            req.status = 'dispatch_completed'
                        elif has_dispatched:
                            req.status = 'dispatch_in_progress'
                        req.save()
                
                success_count += 1
                
        if success_count > 0:
            messages.success(request, f"成功快速撥出 {success_count} 筆物料！")
        else:
            messages.info(request, "所選物料皆已達需撥料數量，無新的撥料動作。")
            
    except Exception as e:
        messages.error(request, f"快速撥料發生錯誤：{str(e)}")
        
    return redirect('requisitions:simple_dispatcher_home')


def _generate_simple_export_excel(requisitions, filename_prefix):
    """通用產生 Excel 函數 - 調整為以物料清單為主，並對齊系統標準格式"""
    
    # 1. 準備申請單摘要資料 (將用於第二個分頁)
    req_list_data = []
    for r in requisitions:
        # 取得申請人顯示名稱 (與網頁一致)
        applicant_name = f"{r.applicant.first_name}{r.applicant.last_name}"
        if not applicant_name:
            applicant_name = r.applicant.username

        req_list_data.append({
            "訂單": r.order_number,
            "需求流程": r.process_type,
            "申請人": applicant_name,
            "需求日期": r.request_date.strftime('%Y-%m-%d') if r.request_date else '',
            "狀態": r.get_status_display(),
            "建立時間": r.created_at.strftime('%Y-%m-%d %H:%M'),
        })
    df_reqs = pd.DataFrame(req_list_data)

    # 2. 準備物料清單資料 (主表，第一個分頁)
    item_data = []
    # 使用 select_related 抓取撥料人員，避免 N+1 查詢
    items = RequisitionItem.objects.filter(requisition__in=requisitions).select_related('requisition', 'requisition__applicant', 'dispatched_by').order_by('requisition__order_number', 'material_number')
    for item in items:
        # 取得申請人顯示名稱
        applicant_name = f"{item.requisition.applicant.first_name}{item.requisition.applicant.last_name}"
        if not applicant_name:
            applicant_name = item.requisition.applicant.username

        # 取得撥料人員顯示名稱
        dispatched_name = ""
        if item.dispatched_by:
            dispatched_name = f"{item.dispatched_by.first_name}{item.dispatched_by.last_name}"
            if not dispatched_name:
                dispatched_name = item.dispatched_by.username

        item_data.append({
            "訂單單號": item.requisition.order_number,
            "需求流程": item.requisition.process_type,
            "物料": item.material_number,
            "品名": item.item_name,
            "需求數量": item.required_quantity,
            "庫存數量": item.stock_quantity if item.stock_quantity is not None else 0,
            "撥料數量 (實際撥出)": item.confirmed_quantity if item.confirmed_quantity is not None else '',
            "最終簽收已確認": "是" if item.is_signed_off else "否",
            "儲格": item.storage_bin,
            "撥料人員": dispatched_name,
            "申請人": applicant_name,
        })
    df_items = pd.DataFrame(item_data)

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # 重要：將「物料明細」放在第一個分頁索引，確保使用者一開啟就能看到
        if not df_items.empty:
            df_items.to_excel(writer, index=False, sheet_name='撥料物料明細')
        else:
            # 若無物料，建立帶有標題的空工作表
            empty_df_items = pd.DataFrame(columns=[
                "訂單單號", "需求流程", "物料", "品名", "需求數量", "庫存數量", "撥料數量 (實際撥出)", "最終簽收已確認"
            ])
            empty_df_items.to_excel(writer, index=False, sheet_name='撥料物料明細')
            
        # 將「申請單摘要」放在第二個分頁
        df_reqs.to_excel(writer, index=False, sheet_name='撥料申請單')
    
    output.seek(0)
    response = HttpResponse(
        output,
        content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )
    filename = f"{filename_prefix}_{timezone.now().strftime('%Y%m%d_%H%M')}.xlsx"
    # 確保檔名是 ASCII 或是經過編碼，這裡簡單使用英文
    response['Content-Disposition'] = f'attachment; filename="{filename}"'
    return response


@login_required
def export_single_requisition_excel(request, pk):
    """單張申請單詳情匯出 Excel"""
    requisition = get_object_or_404(Requisition, pk=pk)
    
    # 權限檢查邏輯與詳情頁面一致
    is_applicant = requisition.applicant == request.user
    is_dispatcher = is_simple_dispatcher(request.user)
    is_supervisor = request.user.groups.filter(name__in=[GROUP_NAMES['APPLICANT_SUPERVISOR'], GROUP_NAMES['DISPATCHER_SUPERVISOR']]).exists()
    
    is_group_member = RequisitionShareGroup.objects.filter(
        members=request.user
    ).filter(members=requisition.applicant).exists()
    
    if not (is_applicant or is_dispatcher or is_supervisor or request.user.is_superuser or is_group_member):
        return redirect('requisitions:requisition_list')
        
    # 重用之前的導出函數，傳入只包含此單的 queryset
    requisitions = Requisition.objects.filter(pk=pk)
    return _generate_simple_export_excel(requisitions, f"requisition_{requisition.order_number}")

@login_required
def simple_requisition_change_detail(request, pk):
    """
    需求變更管理分頁 - 簡易版
    分成上下兩個欄位：
    1. 上：已撥料需退料（confirmed > required && not dismissed）
    2. 下：新增的、未撥料的，以及上面已解除的明細
    """
    requisition = get_object_or_404(Requisition, pk=pk)
    
    # 基本權限檢查
    if not is_simple_applicant(request.user) and not is_simple_dispatcher(request.user) and not request.user.is_superuser:
        messages.error(request, "您沒有權限訪問此頁面。")
        return redirect('core:homepage')

    items = requisition.items.all().select_related('alert_dismissed_by', 'dispatched_by')
    
    # 分類
    to_return_items = []
    other_changes = []
    
    for item in items:
        # 已撥料需退料: 已撥 > 需求
        if item.confirmed_quantity > item.required_quantity:
            to_return_items.append(item)
        elif item.required_quantity > item.confirmed_quantity:
            # 需求 > 已撥 (新增或增量)
            other_changes.append(item)
        elif item.alert_dismissed:
            # 即使數量一致，但若曾經有警示紀錄（且被標記過），也顯示在變更紀錄中
            other_changes.append(item)

    context = {
        'requisition': requisition,
        'to_return_items': to_return_items,
        'other_changes': other_changes,
        'alert_messages': requisition.alert_message.split('\n') if requisition.alert_message else [],
    }
    return render(request, 'requisitions/simple/simple_requisition_change_detail.html', context)


@login_required
def dismiss_requisition_item_alert(request, item_pk):
    """
    解除單筆物料的警示
    """
    item = get_object_or_404(RequisitionItem, pk=item_pk)
    requisition = item.requisition
    
    if request.method == 'POST':
        item.alert_dismissed = True
        item.alert_dismissed_by = request.user
        item.alert_dismissed_at = timezone.now()
        item.save()
        
        # 檢查是否所有項目都已處理
        # 如果所有 confirmed > required 的項目都已解除，
        # 且沒有其他未處理的新增項目（這部分邏輯可以視需求調整）
        remaining_alerts = requisition.items.filter(
            confirmed_quantity__gt=F('required_quantity'),
            alert_dismissed=False
        ).exists()
        
        if not remaining_alerts:
             # 如果沒有待退料的了，可以考慮是否自動關閉全單警示
             # 或者留給使用者手動關閉。照計畫書先保留全單警示手動關閉，
             # 但如果使用者所有單項都按了解除，全單警示可能也該關閉。
             pass

        if request.headers.get('x-requested-with') == 'XMLHttpRequest':
            return JsonResponse({
                'success': True, 
                'dismissed_by': request.user.username,
                'dismissed_at': item.alert_dismissed_at.strftime('%Y-%m-%d %H:%M')
            })
            
        messages.success(request, f"物料 {item.material_number} 的警示已解除。")
        return redirect('requisitions:simple_requisition_change_detail', pk=requisition.pk)

    return HttpResponse("method not allowed", status=405)


@login_required
def shortage_inquiry(request):
    """
    缺料查詢 - 給撥料人員主管使用
    批量帶入工單號碼，自動查出庫存量低於未滿足需求數量的物料
    """
    is_supervisor = request.user.groups.filter(name=GROUP_NAMES['DISPATCHER_SUPERVISOR']).exists()
    if not (request.user.is_superuser or is_supervisor):
        messages.error(request, "您沒有權限使用此功能。")
        return redirect('core:homepage')

    results = None
    submitted_orders = ''
    order_count = 0
    total_items = 0
    shortage_rate = 0
    selected_process_type = ''
    # 收集系統中所有固定的投料點名稱
    from ..models import ProcessType, SemiFinishedProcessType
    pt_set = set(ProcessType.objects.values_list('name', flat=True))
    pt_set.update(SemiFinishedProcessType.objects.values_list('name', flat=True))
    process_types = sorted([name for name in pt_set if name])

    if request.method == 'POST':
        import re
        raw_text = request.POST.get('order_numbers', '')
        submitted_orders = raw_text
        # 解析工單號碼：支援換行、逗號、空白分隔
        order_numbers = [o.strip() for o in re.split(r'[\n\r,\s\t]+', raw_text) if o.strip()]
        selected_process_type = request.POST.get('process_type', '')

        if not order_numbers:
            messages.warning(request, "請輸入至少一個工單號碼。")
        else:
            from ..analysis import get_material_demand_analysis

            # 使用全域分析函數計算所有工單的物料需求與缺料狀態
            all_materials_analysis = get_material_demand_analysis()

            # 預先查詢所有相關工單的客戶名稱
            work_order_map = {}
            for wo in WorkOrder.objects.filter(order_number__in=order_numbers):
                work_order_map[wo.order_number] = wo

            # 預先查詢供應商資訊
            all_material_numbers = set()
            for material_key, data in all_materials_analysis.items():
                if data['is_shortage']:
                    for detail in data['detail_orders']:
                        if detail['order_number'] in order_numbers:
                            all_material_numbers.add(material_key)
                            break

            supplier_map = {}
            for mat in Material.objects.filter(material_code__in=all_material_numbers):
                supplier_map[mat.material_code] = mat.purchaser

            results = []
            seen_orders = set()
            total_items = 0
            
            for material_key, data in all_materials_analysis.items():
                # 篩選出屬於查詢工單的明細
                relevant_details = [d for d in data['detail_orders'] if d['order_number'] in order_numbers]
                
                # 投料點篩選 (在此處篩選會影響 total_items 的統計)
                if selected_process_type:
                    relevant_details = [d for d in relevant_details if d.get('process_type_name') == selected_process_type]

                if not relevant_details:
                    continue
                
                # 計算總筆數 (反映篩選後的總數)
                total_items += len(relevant_details)

                # 只顯示全域分析後確認為缺料的物料
                if not data['is_shortage']:
                    continue

                # 全域庫存
                current_stock = float(data['current_stock'])

                for detail in relevant_details:
                    order_num = detail['order_number']
                    wo = work_order_map.get(order_num)
                    
                    # 這一點的累計需求
                    cumulative_demand = float(detail['cumulative_demand'])
                    
                    # 計算这一點的缺料量 (累計需求 - 庫存，最小為 0，且不能大於該單需求量)
                    required_qty = float(detail['required_quantity'])
                    row_shortage = min(required_qty, max(0.0, cumulative_demand - current_stock))

                    # 如果此工單在累計需求下還沒造成缺料，則視情況跳過？
                    # 使用者要求是「缺料查詢」，所以如果 row_shortage 為 0，代表此工單的需求目前庫存還夠支應
                    if row_shortage <= 0:
                        continue

                    seen_orders.add(order_num)
                    results.append({
                        'order_number': order_num,
                        'machine_model': detail.get('machine_model_name'),
                        'material_number': data['material_number'],
                        'item_name': data['item_name'],
                        'process_type': detail.get('process_type_name'),
                        'required_quantity': float(detail['required_quantity']),
                        'stock_quantity': current_stock,
                        'shortage': row_shortage,
                        'total_demand': cumulative_demand,
                        'demand_date': detail.get('demand_date'),
                        'shipping_date': detail.get('shipping_date') or (wo.shipping_date if wo else None),
                        'estimated_arrival': data.get('estimated_arrival_date'),
                        'customer_name': wo.customer_name if wo else None,
                        'supplier': supplier_map.get(data['material_number']),
                    })

            # 排序：依工單號 -> 品號
            results.sort(key=lambda x: (x['order_number'], x['material_number']))
            order_count = len(seen_orders)
            shortage_rate = (len(results) / total_items * 100) if total_items > 0 else 0

    return render(request, 'requisitions/simple/simple_shortage_inquiry.html', {
        'results': results,
        'submitted_orders': submitted_orders,
        'order_count': order_count,
        'total_items': total_items,
        'shortage_rate': shortage_rate,
        'process_types': process_types,
        'selected_process_type': selected_process_type,
    })


@login_required
@require_POST
def shortage_inquiry_export(request):
    """
    缺料查詢結果匯出 Excel
    """
    is_supervisor = request.user.groups.filter(name=GROUP_NAMES['DISPATCHER_SUPERVISOR']).exists()
    if not (request.user.is_superuser or is_supervisor):
        messages.error(request, "您沒有權限使用此功能。")
        return redirect('core:homepage')

    import re
    raw_text = request.POST.get('order_numbers', '')
    order_numbers = [o.strip() for o in re.split(r'[\n\r,\s\t]+', raw_text) if o.strip()]
    selected_process_type = request.POST.get('process_type', '')

    if not order_numbers:
        messages.error(request, "無可匯出的資料。")
        return redirect('requisitions:shortage_inquiry')

    from ..analysis import get_material_demand_analysis

    # 使用全域分析函數
    all_materials_analysis = get_material_demand_analysis()

    # 預先查詢工單資訊
    work_order_map = {}
    for wo in WorkOrder.objects.filter(order_number__in=order_numbers):
        work_order_map[wo.order_number] = wo

    # 預先查詢供應商資訊
    all_material_numbers = set()
    for material_key, data in all_materials_analysis.items():
        if data['is_shortage']:
            for detail in data['detail_orders']:
                if detail['order_number'] in order_numbers:
                    all_material_numbers.add(material_key)
                    break

    supplier_map = {}
    for mat in Material.objects.filter(material_code__in=all_material_numbers):
        supplier_map[mat.material_code] = mat.purchaser

    rows = []
    total_items = 0
    seen_orders = set()

    for material_key, data in all_materials_analysis.items():
        relevant_details = [d for d in data['detail_orders'] if d['order_number'] in order_numbers]
        
        # 投料點篩選
        if selected_process_type:
            relevant_details = [d for d in relevant_details if d.get('process_type_name') == selected_process_type]

        if not relevant_details:
            continue

        # 計算總筆數
        total_items += len(relevant_details)

        if not data['is_shortage']:
            continue

        current_stock = float(data['current_stock'])

        for detail in relevant_details:
            order_num = detail['order_number']
            wo = work_order_map.get(order_num)
            
            cumulative_demand = float(detail['cumulative_demand'])
            # 計算这一點的缺料量 (累計需求 - 庫存，最小為 0，且不能大於該單需求量)
            required_qty = float(detail['required_quantity'])
            row_shortage = min(required_qty, max(0.0, cumulative_demand - current_stock))

            if row_shortage <= 0:
                continue

            seen_orders.add(order_num)

            rows.append({
                '工單號碼': order_num,
                '客戶': wo.customer_name if wo and wo.customer_name else '',
                '機型': detail.get('machine_model_name', ''),
                '品號': data['material_number'],
                '品名': data['item_name'],
                '供應商': supplier_map.get(data['material_number'], ''),
                '投料點': detail.get('process_type_name', ''),
                '此工單需求量': float(detail['required_quantity']),
                '累計未滿足需求': cumulative_demand,
                '庫存量': current_stock,
                '缺料': row_shortage,
                '需求日期': detail.get('demand_date') or '',
                '出貨日期': detail.get('shipping_date') or (wo.shipping_date if wo else '') or '',
                '預計入料': data.get('estimated_arrival_date') or '',
            })

    rows.sort(key=lambda x: (x['工單號碼'], x['品號']))

    # 計算統計資料
    shortage_count = len(rows)
    shortage_rate = (shortage_count / total_items * 100) if total_items > 0 else 0
    order_count = len(seen_orders)
    summary_text = f"總數/缺料數:{total_items}/{shortage_count}筆 缺料率{shortage_rate:.1f}% 工單{order_count}張"

    df = pd.DataFrame(rows)
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='缺料查詢', startrow=0)
        worksheet = writer.sheets['缺料查詢']
        
        # 寫入統計資料到最後一列
        last_row = len(rows) + 2 # 表頭佔1列，資料佔 len(rows) 列，下一列為 last_row
        cell = worksheet.cell(row=last_row, column=1, value=summary_text)
        worksheet.merge_cells(start_row=last_row, start_column=1, end_row=last_row, end_column=5)
        
        # 設定粗體與些微樣式
        from openpyxl.styles import Font
        cell.font = Font(bold=True, size=12)

    output.seek(0)

    response = HttpResponse(
        output.getvalue(),
        content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )
    response['Content-Disposition'] = f'attachment; filename="shortage_inquiry_{date.today().isoformat()}.xlsx"'
    return response


@login_required
def sync_mps_order_info(request):
    """
    從外部 API 同步工單出貨資訊 (客戶與出貨日期)
    """
    from ..utils import sync_external_order_info
    
    success, message = sync_external_order_info()
    if success:
        messages.success(request, message)
    else:
        messages.error(request, message)
        
    return redirect('requisitions:shortage_inquiry')

