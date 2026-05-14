from django.http import JsonResponse
from django.db.models import Q


def shortage_materials_api(request):
    """
    API endpoint to get shortage materials data (items marked as 'backordered').
    Returns JSON data for external programs to use.
    """
    from decimal import Decimal
    from requisitions.models import RequisitionItem, WorkOrderMaterial
    from django.db.models import Max
    
    # 只取得被標記為「缺料」的物料
    backordered_items = RequisitionItem.objects.filter(
        dispatch_status='backordered',
        requisition__is_archived=False
    )
    
    # 聚合相同物料
    aggregated_shortages = {}
    for item in backordered_items:
        key = item.material_number
        shortage = float(item.required_quantity - (item.confirmed_quantity or 0))
        if shortage <= 0:
            continue
            
        if key not in aggregated_shortages:
            # 嘗試從 WorkOrderMaterial 取得預計入料日期
            latest_date = WorkOrderMaterial.objects.filter(
                material_number=key
            ).aggregate(latest_date=Max('estimated_arrival_date'))['latest_date']
            
            aggregated_shortages[key] = {
                'material_number': item.material_number,
                'item_name': item.item_name,
                'total_shortage': 0.0,
                'orders': [],
                'estimated_arrival_date': str(latest_date) if latest_date else None
            }
        aggregated_shortages[key]['total_shortage'] += shortage
        if item.order_number not in aggregated_shortages[key]['orders']:
            aggregated_shortages[key]['orders'].append(item.order_number)
    
    # 轉換為列表
    result = list(aggregated_shortages.values())
    
    return JsonResponse({
        'success': True,
        'count': len(result),
        'shortage_materials': result
    })

def requisition_items_shortages_api(request):
    """
    回傳詳細的申請單缺料清單。
    支援 query parameter:
    - type: 'finished' 或 'semi_finished'
    - req_id: 指定單號 ID
    """
    from requisitions.models import RequisitionItem
    
    requisition_type = request.GET.get('type')
    req_id = request.GET.get('req_id')
    
    # 基礎過濾條件：未歸檔且狀態為 backordered
    filters = Q(dispatch_status='backordered', requisition__is_archived=False)
    
    if requisition_type:
        filters &= Q(requisition__requisition_type=requisition_type)
    
    if req_id:
        filters &= Q(requisition_id=req_id)
        
    items = RequisitionItem.objects.filter(filters).select_related('requisition', 'requisition__applicant')
    
    data = []
    for item in items:
        data.append({
            'requisition_id': item.requisition.id,
            'order_number': item.order_number,
            'material_number': item.material_number,
            'item_name': item.item_name,
            'required_quantity': float(item.required_quantity),
            'confirmed_quantity': float(item.confirmed_quantity or 0),
            'shortage_quantity': float(item.required_quantity - (item.confirmed_quantity or 0)),
            'storage_bin': item.storage_bin,
            'request_date': str(item.requisition.request_date),
            'applicant': item.requisition.applicant.username,
            'requisition_type': item.requisition.requisition_type,
            'status': item.requisition.get_status_display()
        })
    
    return JsonResponse({
        'success': True,
        'count': len(data),
        'items': data
    })
