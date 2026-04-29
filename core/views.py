from django.shortcuts import render, redirect, get_object_or_404
from django.utils import timezone
from datetime import timedelta
from django.db.models import Count, Q, F, DecimalField, Prefetch
from requisitions.constants import GROUP_NAMES
from requisitions.models import Requisition, RequisitionItem, ProcessType, SemiFinishedProcessType

def homepage(request):
    """
    首頁視圖 - 根據使用者角色導向不同介面
    """
    if request.user.is_authenticated:
        is_admin = request.user.is_superuser
        
        # 檢查是否為簡易人員（非主管）
        is_simple_applicant = request.user.groups.filter(
            name=GROUP_NAMES['APPLICANT']
        ).exists() and not request.user.groups.filter(
            name=GROUP_NAMES['APPLICANT_SUPERVISOR']
        ).exists() and not is_admin
        
        is_simple_dispatcher = request.user.groups.filter(
            name=GROUP_NAMES['DISPATCHER']
        ).exists() and not request.user.groups.filter(
            name=GROUP_NAMES['DISPATCHER_SUPERVISOR']
        ).exists() and not is_admin
        
        # 簡易人員自動導向簡易介面
        if is_simple_applicant:
            return redirect('requisitions:simple_applicant_home')
        if is_simple_dispatcher:
            return redirect('requisitions:simple_dispatcher_home')
        
        # 主管和管理員使用完整介面
        is_applicant_supervisor = request.user.groups.filter(
            name=GROUP_NAMES['APPLICANT_SUPERVISOR']
        ).exists()
        is_dispatcher_supervisor = request.user.groups.filter(
            name=GROUP_NAMES['DISPATCHER_SUPERVISOR']
        ).exists()
        
        # 相容舊版群組名稱
        is_applicant = is_applicant_supervisor or request.user.groups.filter(name='申請人員').exists()
        is_material_handler = is_dispatcher_supervisor or request.user.groups.filter(name='撥料人員').exists()
        
        # 為主管準備看板數據 (投料點匯總)
        process_summaries = []
        if is_admin or is_applicant_supervisor or is_dispatcher_supervisor:
            # 計算本週起始日 (週一)
            now = timezone.now()
            start_of_week = now - timedelta(days=now.weekday())
            start_of_week = start_of_week.replace(hour=0, minute=0, second=0, microsecond=0)
            
            # 1. 成品投料點 (使用固定清單)
            target_finished_names = ['機械', '系統', '電裝', '鑄件', '護蓋', '刀庫', '出貨', '組件', '軟體研發部', '其他']
            finished_stats = {item['process_type']: item for item in Requisition.objects.filter(
                requisition_type='finished',
                is_archived=False
            ).values('process_type').annotate(
                total_count=Count('id'),
                week_count=Count('id', filter=Q(created_at__gte=start_of_week))
            )}
            
            for name in target_finished_names:
                stat = finished_stats.get(name, {'total_count': 0, 'week_count': 0})
                process_summaries.append({
                    'name': name,
                    'display_name': name,
                    'total': stat['total_count'],
                    'week': stat['week_count'],
                    'type': 'finished'
                })
            
            # 2. 半成品投料點 (依領料人分組，比照撥料介面)
            semi_stats = Requisition.objects.filter(
                requisition_type='semi_finished',
                is_archived=False
            ).values(
                'applicant__username', 
                'applicant__first_name', 
                'applicant__last_name'
            ).annotate(
                total_count=Count('id'),
                week_count=Count('id', filter=Q(created_at__gte=start_of_week))
            ).order_by('applicant__username')
            
            for stat in semi_stats:
                username = stat['applicant__username']
                first_name = stat['applicant__first_name']
                last_name = stat['applicant__last_name']
                
                if username:
                    # 格式化姓名：如果有姓或名就組合，否則用帳號
                    display_name = f"{last_name}{first_name}" if (last_name or first_name) else username
                    
                    process_summaries.append({
                        'name': username,
                        'display_name': display_name,
                        'total': stat['total_count'],
                        'week': stat['week_count'],
                        'type': 'semi_finished'
                    })

        context = {
            'is_admin': is_admin,
            'is_applicant': is_applicant,
            'is_material_handler': is_material_handler,
            'is_applicant_supervisor': is_applicant_supervisor,
            'is_dispatcher_supervisor': is_dispatcher_supervisor,
            'is_supervisor': is_applicant_supervisor or is_dispatcher_supervisor,
            'process_summaries': process_summaries,
        }
        return render(request, 'core/homepage.html', context)
    else:
        return render(request, 'core/landing.html')

def supervisor_process_detail(request, process_name):
    """
    主管點進特定投料點後的詳細看板
    """
    if not request.user.is_authenticated:
        return redirect('login')
        
    is_admin = request.user.is_superuser
    is_supervisor = request.user.groups.filter(
        name__in=[GROUP_NAMES['APPLICANT_SUPERVISOR'], GROUP_NAMES['DISPATCHER_SUPERVISOR']]
    ).exists()
    
    if not (is_admin or is_supervisor):
        return redirect('core:homepage')

    # 計算本週起始日
    now = timezone.now()
    start_of_week = now - timedelta(days=now.weekday())
    start_of_week = start_of_week.replace(hour=0, minute=0, second=0, microsecond=0)
    
    # 預加載缺料項
    short_items_prefetch = Prefetch(
        'items', 
        queryset=RequisitionItem.objects.filter(dispatch_status='backordered'),
        to_attr='short_items'
    )
    
    # 獲取該投料點的所有申請單，並計算進度
    # 成品依據 process_type，半成品依據 applicant__username
    requisitions = Requisition.objects.filter(
        Q(process_type=process_name) | Q(applicant__username=process_name),
        is_archived=False
    ).annotate(
        total_items_count=Count('items'),
        dispatched_items_count=Count('items', filter=Q(items__dispatch_status='dispatched'))
    ).prefetch_related(short_items_prefetch).order_by('-created_at')
    
    # 嘗試獲取顯示名稱 (如果是領料人則顯示姓名)
    display_name = process_name
    first_req = requisitions.first()
    if first_req and first_req.requisition_type == 'semi_finished':
        u = first_req.applicant
        display_name = f"{u.last_name}{u.first_name}" if (u.last_name or u.first_name) else u.username

    # 獲取所有相關的機型名稱
    order_numbers = [r.order_number for r in requisitions]
    machine_model_map = {item['order_number']: item['machine_model__name'] for item in WorkOrderMaterial.objects.filter(
        order_number__in=order_numbers,
        machine_model__isnull=False
    ).values('order_number', 'machine_model__name')}

    # 為了計算百分比，我們在 Python 中處理
    for req in requisitions:
        if req.total_items_count > 0:
            req.progress = int((req.dispatched_items_count / req.total_items_count) * 100)
        else:
            req.progress = 0
        # 賦予機型名稱
        req.machine_model_name = machine_model_map.get(req.order_number, "未知機型")

    this_week_reqs = [r for r in requisitions if r.created_at >= start_of_week]
    other_reqs = [r for r in requisitions if r.created_at < start_of_week]
    
    context = {
        'process_name': process_name,
        'display_name': display_name,
        'this_week_reqs': this_week_reqs,
        'other_reqs': other_reqs,
        'start_of_week': start_of_week,
    }
    return render(request, 'core/supervisor_process_detail.html', context)