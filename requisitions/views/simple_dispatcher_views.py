"""
簡易撥料人員視圖 - 撥料員首頁、分類、詳情、合併撥料
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
        ).values('applicant__username', 'applicant__first_name', 'applicant__last_name').annotate(
            pending_count=Count('id')
        ).order_by('applicant__username')
        
        for item in pending_requisitions:
            username = item['applicant__username']
            first_name = item['applicant__first_name'] or ''
            last_name = item['applicant__last_name'] or ''
            display_name = f"{first_name}{last_name}".strip() or username
            count = item['pending_count']
            
            # 為了介面一致性，這裡的 name 放 username
            # color 可以隨機或固定，這裡暫時給一個預設色
            categories.append({
                'name': username, 
                'display_name': display_name,
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
    
    display_name = category
    if current_type == 'semi_finished':
        # 半成品：category 代表 applicant.username
        target_username = category
        target_user = User.objects.filter(username=target_username).first()
        if target_user:
            u_first = target_user.first_name or ''
            u_last = target_user.last_name or ''
            display_name = f"{u_first}{u_last}".strip() or target_username
        
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
        'display_name': display_name,
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
    # 根據類型決定過濾條件
    if current_type == 'semi_finished':
        filter_q = Q(applicant__username=category, requisition_type='semi_finished')
        category_color = '#8B5CF6' # Purple for applicants
    else:
        filter_q = Q(process_type__icontains=category)
        category_color = PROCESS_CATEGORY_COLORS.get(category, '#6B7280')

    # 取得該分類下的待處理申請單
    pending_requisitions = Requisition.objects.filter(
        filter_q,
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
                ).exclude(dispatch_status='dispatched').exclude(material_number='PARENT_SCOPE')
                
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
                ).exclude(dispatch_status='dispatched').exclude(material_number='PARENT_SCOPE')
                
                count = 0
                for item in items:
                    item.dispatch_status = 'backordered'
                    item.save()
                    count += 1

                return JsonResponse({'success': True, 'message': f'已將物料 {material_number} 的 {count} 筆需求標記為缺料'})
            except Exception as e:
                return JsonResponse({'success': False, 'message': str(e)})

        elif action == 'undo_item':
            item_pk = request.POST.get('item_pk')
            try:
                item = get_object_or_404(RequisitionItem, pk=item_pk)
                item.confirmed_quantity = Decimal('0')
                item.dispatch_status = None
                item.dispatched_by = None
                item.dispatched_at = None
                item.save()
                # 更新申請單狀態
                req = item.requisition
                all_items = req.items.all()
                dispatched = all_items.filter(dispatch_status='dispatched').count()
                if dispatched == 0:
                    req.status = 'demand_submitted'
                else:
                    req.status = 'dispatch_in_progress'
                req.save()
                return JsonResponse({'success': True, 'message': '已還原'})
            except Exception as e:
                return JsonResponse({'success': False, 'message': str(e)})

        elif action == 'undo_material':
            material_number = request.POST.get('material_number')
            try:
                items = RequisitionItem.objects.filter(
                    (Q(requisition__applicant__username=category) if current_type == 'semi_finished' else Q(requisition__process_type__icontains=category)),
                    requisition__order_number__in=order_numbers,
                    material_number=material_number,
                    requisition__is_archived=False
                ).filter(Q(dispatch_status='dispatched') | Q(dispatch_status='backordered'))
                
                count = 0
                affected_reqs = set()
                for item in items:
                    item.confirmed_quantity = Decimal('0')
                    item.dispatch_status = None
                    item.dispatched_by = None
                    item.dispatched_at = None
                    item.save()
                    affected_reqs.add(item.requisition_id)
                    count += 1
                
                # 更新受影響申請單
                for req in Requisition.objects.filter(pk__in=affected_reqs):
                    all_items = req.items.all()
                    dispatched = all_items.filter(dispatch_status='dispatched').count()
                    if dispatched == 0:
                        req.status = 'demand_submitted'
                    else:
                        req.status = 'dispatch_in_progress'
                    req.save()
                    
                return JsonResponse({'success': True, 'message': f'已將物料 {material_number} 的 {count} 筆紀錄還原'})
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
    ).exclude(dispatch_status='dispatched').exclude(material_number='PARENT_SCOPE').select_related('requisition')

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
        items = requisition.items.exclude(material_number='PARENT_SCOPE').order_by('storage_bin', 'material_number')
    elif sort_param == 'name':
        items = requisition.items.exclude(material_number='PARENT_SCOPE').order_by('item_name', 'material_number')
    elif sort_param == 'status':
        # Status sort: Pending (None/Empty) -> Backordered -> Dispatched
        items = requisition.items.exclude(material_number='PARENT_SCOPE').select_related('dispatched_by').annotate(
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
        items = requisition.items.exclude(material_number='PARENT_SCOPE').select_related('dispatched_by').order_by('material_number')
    
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
        # 預處理是否可撥料（未簽收且未處理，且來源物料需為啟用狀態）
        item.is_actionable = (
            item.dispatch_status not in ['dispatched', 'backordered'] and 
            (item.source_material is None or item.source_material.is_active)
        )
        
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


