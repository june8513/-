"""
簡易申請人員視圖 - 申請單建立、查看、修改、簽收
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
from common.permissions import is_simple_applicant, is_simple_dispatcher
from ..services.requisition_service import RequisitionService

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
        Q(applicant=request.user) | Q(applicant__in=viewable_owners) | Q(demand_person=request.user),
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
                                ).exclude(material_number='PARENT_SCOPE')
                                
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
    
    if requisition.applicant != request.user and requisition.demand_person != request.user and not is_supervisor and not request.user.is_superuser and not is_group_member:
        messages.error(request, "您沒有權限查看此申請單。")
        return redirect('requisitions:simple_applicant_home')
    
    items = requisition.items.exclude(material_number='PARENT_SCOPE').order_by('material_number')
    
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
        'images': requisition.images.all().order_by('-uploaded_at'),
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
    
    if requisition.applicant != request.user and requisition.demand_person != request.user and not request.user.is_superuser and not is_group_member:
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


@login_required
def simple_applicant_batch_sign_off(request):
    """批量簽收 - 申請人可以選擇多張申請單並一次性簽收其中的物料"""
    if not is_simple_applicant(request.user) and not request.user.is_superuser:
        return redirect('requisitions:requisition_list')
    
    current_type = request.GET.get('type', 'finished')
    action = request.POST.get('action')
    
    # 取得被授權查看的使用者列表
    share_groups = request.user.requisition_share_groups.all()
    viewable_owners = User.objects.filter(requisition_share_groups__in=share_groups).distinct()
    
    if request.method == 'POST' and action == 'select_requisitions':
        # 階段 2：顯示選中申請單的物料
        selected_ids = request.POST.getlist('selected_requisitions')
        if not selected_ids:
            messages.warning(request, "請至少選擇一張申請單。")
            return redirect(f"{reverse('requisitions:simple_applicant_batch_sign_off')}?type={current_type}")
        
        requisitions = Requisition.objects.filter(
            id__in=selected_ids,
            is_archived=False
        ).filter(
            Q(applicant=request.user) | Q(applicant__in=viewable_owners) | Q(demand_person=request.user)
        )
        
        # 取得這些申請單中「已撥料」且「未簽收」且「無異況」的項目
        items = RequisitionItem.objects.filter(
            requisition__in=requisitions,
            dispatch_status='dispatched',
            is_signed_off=False,
            has_issue=False
        ).select_related('requisition').order_by('order_number', 'material_number')
        
        return render(request, 'requisitions/simple/simple_applicant_batch_sign_off_items.html', {
            'requisitions': requisitions,
            'items': items,
            'current_type': current_type,
            'selected_ids': ','.join(selected_ids)
        })

    elif request.method == 'POST' and action == 'execute_sign_off':
        # 階段 3：執行簽收
        item_ids = request.POST.getlist('selected_items')
        if not item_ids:
            messages.warning(request, "請至少選擇一個物料項目。")
            return redirect(f"{reverse('requisitions:simple_applicant_home')}?type={current_type}")
        
        signed_count = RequisitionService.sign_off_items(item_ids, request.user)
        messages.success(request, f"成功簽收 {signed_count} 筆物料項目。")
        return redirect(f"{reverse('requisitions:simple_applicant_home')}?type={current_type}")

    elif request.method == 'POST' and action == 'execute_all_sign_off':
        # 階段 3：執行全部簽收
        req_ids_str = request.POST.get('requisition_ids', '')
        if not req_ids_str:
            return redirect(f"{reverse('requisitions:simple_applicant_home')}?type={current_type}")
            
        req_ids = req_ids_str.split(',')
        signed_count = RequisitionService.sign_off_all_items_in_requisitions(req_ids, request.user)
        messages.success(request, f"成功簽收全部 {signed_count} 筆物料項目。")
        return redirect(f"{reverse('requisitions:simple_applicant_home')}?type={current_type}")

    # 階段 1：顯示可選申請單列表 (GET)
    requisitions = Requisition.objects.filter(
        Q(applicant=request.user) | Q(applicant__in=viewable_owners) | Q(demand_person=request.user),
        requisition_type=current_type,
        status__in=['dispatch_in_progress', 'dispatch_completed'],
        is_archived=False
    ).annotate(
        unsigned_items_count=Count('items', filter=Q(items__dispatch_status='dispatched', items__is_signed_off=False))
    ).filter(unsigned_items_count__gt=0).order_by('-updated_at')
    
    return render(request, 'requisitions/simple/simple_applicant_batch_sign_off_select.html', {
        'requisitions': requisitions,
        'current_type': current_type
    })


@login_required
@require_POST
def simple_upload_requisition_images(request, pk):
    """簡易申請單上傳/拍照圖片"""
    requisition = get_object_or_404(Requisition, pk=pk)
    
    # 檢查權限：與詳情頁一致，新增允許撥料/派工人員上傳
    is_supervisor = request.user.groups.filter(name=GROUP_NAMES['APPLICANT_SUPERVISOR']).exists()
    is_group_member = RequisitionShareGroup.objects.filter(
        members=request.user
    ).filter(members=requisition.applicant).exists()
    is_dispatcher = is_simple_dispatcher(request.user)
    
    if (requisition.applicant != request.user and 
        requisition.demand_person != request.user and 
        not is_supervisor and 
        not request.user.is_superuser and 
        not is_group_member and 
        not is_dispatcher):
        return JsonResponse({'success': False, 'message': '您沒有權限為此申請單上傳圖片。'})
        
    if requisition.is_archived:
        return JsonResponse({'success': False, 'message': '此申請單已歸檔，無法上傳圖片。'})

    if 'image' in request.FILES or 'images' in request.FILES:
        from requisitions.models import RequisitionImage
        # 支援單張或多張上傳
        uploaded_files = request.FILES.getlist('image') or request.FILES.getlist('images')
        
        saved_images = []
        for f in uploaded_files:
            img_obj = RequisitionImage.objects.create(
                requisition=requisition,
                image=f,
                uploaded_by=request.user
            )
            
            uploader_name = img_obj.uploaded_by.first_name + img_obj.uploaded_by.last_name if img_obj.uploaded_by and (img_obj.uploaded_by.first_name or img_obj.uploaded_by.last_name) else (img_obj.uploaded_by.username if img_obj.uploaded_by else '未知')
            
            saved_images.append({
                'id': img_obj.id,
                'url': img_obj.image.url,
                'uploaded_at': img_obj.uploaded_at.strftime('%Y-%m-%d %H:%M'),
                'uploaded_by': uploader_name
            })
            
        return JsonResponse({
            'success': True,
            'message': f'成功上傳 {len(saved_images)} 張圖片！',
            'images': saved_images
        })
    else:
        return JsonResponse({'success': False, 'message': '未檢測到上傳的圖片檔案。'})


@login_required
@require_POST
def simple_delete_requisition_image(request, pk, img_id):
    """刪除上傳的圖片"""
    requisition = get_object_or_404(Requisition, pk=pk)
    
    # 檢查權限：與詳情頁一致，新增允許撥料/派工人員刪除自己上傳的圖片
    is_supervisor = request.user.groups.filter(name=GROUP_NAMES['APPLICANT_SUPERVISOR']).exists()
    is_group_member = RequisitionShareGroup.objects.filter(
        members=request.user
    ).filter(members=requisition.applicant).exists()
    is_dispatcher = is_simple_dispatcher(request.user)
    
    if (requisition.applicant != request.user and 
        requisition.demand_person != request.user and 
        not is_supervisor and 
        not request.user.is_superuser and 
        not is_group_member and 
        not is_dispatcher):
        return JsonResponse({'success': False, 'message': '您沒有權限刪除此圖片。'})
        
    if requisition.is_archived:
        return JsonResponse({'success': False, 'message': '此申請單已歸檔，無法刪除圖片。'})
        
    from requisitions.models import RequisitionImage
    img_obj = get_object_or_404(RequisitionImage, pk=img_id, requisition=requisition)
    
    # 僅限上傳者本人或主管、超級管理員可以刪除
    if img_obj.uploaded_by != request.user and not is_supervisor and not request.user.is_superuser:
        return JsonResponse({'success': False, 'message': '您只能刪除自己上傳的圖片。'})
        
    img_obj.image.delete(save=False) # 刪除實體檔案
    img_obj.delete() # 刪除資料庫紀錄
    
    return JsonResponse({'success': True, 'message': '圖片已成功刪除。'})
