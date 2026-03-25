from django.http import JsonResponse
from django.db.models import Q
import json
import threading
import uuid
from datetime import date, datetime, timedelta # Import date and timedelta
import re # Import re for regex
from .action_handlers import handle_search, handle_export
from .ollama_integration import get_ollama_response # Import the new Ollama integration
from requisitions.tasks import run_ollama_task, TASK_RESULTS # Import task runner and results store
from requisitions.models import AIUserCorrection

def natural_action_view(request):
    if request.method != 'GET':
        return JsonResponse({'error': 'Only GET requests are supported'}, status=405)

    query_text = request.GET.get('q', '')
    history_str = request.GET.get('history', '[]')
    
    if not query_text:
        return JsonResponse({'error': 'Query parameter "q" is missing'}, status=400)

    try:
        history = json.loads(history_str)
    except Exception:
        history = []

    # Generate a unique task ID
    task_id = str(uuid.uuid4())

    # Start the Ollama processing in a new thread
    thread = threading.Thread(target=run_ollama_task, args=(task_id, query_text, request.user.id, history))
    thread.start()

    # Immediately return the task ID to the client
    return JsonResponse({'task_id': task_id, 'status': 'PROCESSING', 'message': '正在思考中...'})

def check_task_status(request, task_id):
    """
    API endpoint for clients to poll the status of a background task.
    """
    task_info = TASK_RESULTS.get(task_id)

    if not task_info:
        return JsonResponse({'status': 'NOT_FOUND', 'message': '任務不存在或已過期。'}, status=404)

    status = task_info['status']
    result = task_info['result']
    error = task_info['error']

    if status == 'SUCCESS':
        # Optionally remove the task from TASK_RESULTS after successful retrieval
        # del TASK_RESULTS[task_id]
        return JsonResponse({'status': status, 'result': result})
    elif status == 'FAILURE':
        # Optionally remove the task from TASK_RESULTS after failure
        # del TASK_RESULTS[task_id]
        return JsonResponse({'status': status, 'error': error}, status=500)
    else: # PENDING or PROCESSING
        return JsonResponse({'status': status, 'message': '任務仍在處理中...'})


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

def save_ai_correction(request):
    """
    API endpoint to save user-provided corrections for AI responses.
    """
    if request.method != 'POST':
        return JsonResponse({'error': 'Only POST requests are supported'}, status=405)

    try:
        data = json.loads(request.body)
        query_text = data.get('query_text')
        incorrect_response = data.get('incorrect_response')
        correction_text = data.get('correction_text')

        if not all([query_text, incorrect_response, correction_text]):
            return JsonResponse({'error': 'Missing required fields'}, status=400)

        correction = AIUserCorrection.objects.create(
            user=request.user,
            query_text=query_text,
            incorrect_response=incorrect_response,
            correction_text=correction_text
        )

        return JsonResponse({
            'success': True,
            'message': '修正已記錄，我會記住的！',
            'id': correction.id
        })
    except Exception as e:
        return JsonResponse({'error': str(e)}, status=500)
