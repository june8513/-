from django.shortcuts import render, redirect
from requisitions.constants import GROUP_NAMES


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
        
        # 為主管準備看板數據
        all_requisitions = []
        shortage_requisitions = []
        semi_all_requisitions = []
        semi_shortage_requisitions = []
        
        if is_admin or is_applicant_supervisor or is_dispatcher_supervisor:
            from requisitions.models import Requisition, RequisitionItem
            from django.db.models import Prefetch
            
            # 預加載缺料項
            short_items_prefetch = Prefetch(
                'items', 
                queryset=RequisitionItem.objects.filter(dispatch_status='backordered'),
                to_attr='short_items'
            )

            # 成品數據
            all_requisitions = Requisition.objects.filter(requisition_type='finished').order_by('-created_at')[:20]
            shortage_requisitions = Requisition.objects.filter(
                requisition_type='finished',
                is_archived=False,
                items__dispatch_status='backordered'
            ).distinct().prefetch_related(short_items_prefetch).order_by('-created_at')
            
            # 半成品數據
            semi_all_requisitions = Requisition.objects.filter(requisition_type='semi_finished').order_by('-created_at')[:20]
            semi_shortage_requisitions = Requisition.objects.filter(
                requisition_type='semi_finished',
                is_archived=False,
                items__dispatch_status='backordered'
            ).distinct().prefetch_related(short_items_prefetch).order_by('-created_at')

        context = {
            'is_admin': is_admin,
            'is_applicant': is_applicant,
            'is_material_handler': is_material_handler,
            'is_applicant_supervisor': is_applicant_supervisor,
            'is_dispatcher_supervisor': is_dispatcher_supervisor,
            'is_supervisor': is_applicant_supervisor or is_dispatcher_supervisor,
            # 看板數據
            'all_requisitions': all_requisitions,
            'shortage_requisitions': shortage_requisitions,
            'semi_all_requisitions': semi_all_requisitions,
            'semi_shortage_requisitions': semi_shortage_requisitions,
        }
        return render(request, 'core/homepage.html', context)
    else:
        return render(request, 'core/landing.html')