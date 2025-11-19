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
