"""
快速撥料視圖 - 快速批量撥料功能
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


