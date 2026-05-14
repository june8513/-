from django.shortcuts import render, redirect, get_object_or_404
from django.contrib.auth.decorators import login_required
from django.contrib import messages
from django.http import JsonResponse
from requisitions.models import UserSelectedMaterial, WorkOrderMaterial, Requisition, RequisitionItem
from requisitions.services.special_request_service import SpecialRequestService
from django.db.models import Q

def special_request_permission_required(view_func):
    """檢查是否有特殊申請權限的裝飾器"""
    def _wrapped_view(request, *args, **kwargs):
        if not request.user.is_authenticated:
            return redirect('login')
        if not hasattr(request.user, 'profile') or not request.user.profile.can_access_special_request:
            messages.error(request, "您沒有權限訪問此功能。")
            return redirect('requisitions:simple_applicant_home')
        return view_func(request, *args, **kwargs)
    return _wrapped_view

@login_required
@special_request_permission_required
def special_request_home(request):
    """特殊申請功能入口"""
    return render(request, 'requisitions/special/home.html')

@login_required
@special_request_permission_required
def update_user_materials_view(request):
    """更新個人物料資料庫"""
    if request.method == 'POST':
        raw_text = request.POST.get('materials', '')
        count = SpecialRequestService.update_user_materials(request.user, raw_text)
        messages.success(request, f"成功更新 {count} 筆物料到您的個人資料庫。")
        return redirect('requisitions:special_request_home')
    
    # 獲取目前已有的物料
    current_materials = UserSelectedMaterial.objects.filter(user=request.user).values_list('material_number', flat=True)
    raw_text = "\n".join(current_materials)
    
    return render(request, 'requisitions/special/update_materials.html', {'raw_text': raw_text})

@login_required
@special_request_permission_required
def bulk_requisition_form_view(request):
    """批量工單申請表單"""
    if request.method == 'POST':
        step = request.POST.get('step')
        
        if step == 'match':
            # 第一步：解析工單並進行比對
            raw_orders = request.POST.get('order_numbers', '')
            order_numbers = [o.strip() for o in raw_orders.split('\n') if o.strip()]
            
            if not order_numbers:
                messages.error(request, "請輸入至少一個工單單號。")
                return render(request, 'requisitions/special/bulk_request_step1.html')
            
            matching_items = SpecialRequestService.get_matching_materials(request.user, order_numbers)
            
            return render(request, 'requisitions/special/bulk_request_step2.html', {
                'items': matching_items,
                'order_numbers': ",".join(order_numbers)
            })
            
        elif step == 'submit':
            # 第二步：提交申請
            order_numbers_raw = request.POST.get('order_numbers', '')
            order_numbers = order_numbers_raw.split(',')
            request_date = request.POST.get('request_date')
            demand_person = request.POST.get('demand_person')
            
            # 收集項目確認數據
            items_data = []
            for key, value in request.POST.items():
                if key.startswith('qty_'):
                    wom_id = key.replace('qty_', '')
                    items_data.append({
                        'wom_id': int(wom_id),
                        'request_qty': float(value)
                    })
            
            try:
                requisition = SpecialRequestService.create_special_requisition(
                    request.user, order_numbers, request_date, demand_person, items_data
                )
                messages.success(request, f"申請單 {requisition.id} 已成功建立。")
                return redirect('requisitions:simple_applicant_home')
            except Exception as e:
                messages.error(request, f"建立申請單時發生錯誤: {str(e)}")
                return redirect('requisitions:special_request_home')

    return render(request, 'requisitions/special/bulk_request_step1.html')

@login_required
def global_dispatch_search_view(request):
    """
    撥料員全域搜尋 - 搜尋功能
    顯示特定投料點所有未歸檔且待撥料的項目
    """
    if not request.user.is_superuser and not any(group.name == '簡易撥料員' for group in request.user.groups.all()):
        messages.error(request, "您沒有權限訪問此功能。")
        return redirect('requisitions:simple_dispatcher_home')

    # 從參數取得投料點，預設為「組件」
    process_type = request.GET.get('category', '組件')
    query = request.GET.get('q', '')

    # 獲取投料點清單用於模板切換
    from requisitions.constants import PROCESS_CATEGORY_NAMES, PROCESS_CATEGORY_COLORS
    
    # 獲取項目基礎查詢集 (包含已歸檔項目，以便搜尋歷史)
    items_qs = RequisitionItem.objects.filter(
        requisition__process_type__icontains=process_type
    ).select_related('requisition', 'source_material')

    if query:
        # 有搜尋關鍵字時
        items_qs = items_qs.filter(
            Q(order_number__icontains=query) |
            Q(material_number__icontains=query) |
            Q(item_name__icontains=query)
        )
    
    # 依照工單號碼與建立時間排序，並限制最近的 1000 筆以免頁面過重
    items = items_qs.order_by('-requisition__created_at', 'order_number')[:1000]

    return render(request, 'requisitions/special/dispatcher_global_search.html', {
        'items': items,
        'process_type': process_type,
        'categories': PROCESS_CATEGORY_NAMES,
        'query': query,
        'category_colors': PROCESS_CATEGORY_COLORS
    })
