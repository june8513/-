from django.http import JsonResponse
from django.http import JsonResponse
from django.http import JsonResponse
import json
import threading
import uuid
from datetime import date, datetime, timedelta # Import date and timedelta
import re # Import re for regex
from .action_handlers import handle_search, handle_export
from .ollama_integration import get_ollama_response # Import the new Ollama integration
from requisitions.tasks import run_ollama_task, TASK_RESULTS # Import task runner and results store

def natural_action_view(request):
    if request.method != 'GET':
        return JsonResponse({'error': 'Only GET requests are supported'}, status=405)

    query_text = request.GET.get('q', '')
    if not query_text:
        return JsonResponse({'error': 'Query parameter "q" is missing'}, status=400)

    # Generate a unique task ID
    task_id = str(uuid.uuid4())

    # Start the Ollama processing in a new thread
    thread = threading.Thread(target=run_ollama_task, args=(task_id, query_text, request.user.id))
    thread.start()

    # Immediately return the task ID to the client
    return JsonResponse({'task_id': task_id, 'status': 'PROCESSING', 'message': '正在處理您的請求，請稍候...'})

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
