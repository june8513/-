import threading
import uuid
import time
import json # Import json for parsing JsonResponse content
from datetime import date, datetime, timedelta
from decimal import Decimal

from django.http import JsonResponse
from django.contrib.auth import get_user_model
from django.db.models.functions import Coalesce # Needed for shortage calculation in action_handlers
from django.db.models import F, ExpressionWrapper, DecimalField # Needed for shortage calculation in action_handlers

from .views.ollama_integration import get_ollama_response
from .views.action_handlers import handle_search, handle_export

# Global dictionary to store task results
# In a production environment, this would typically be a database or a dedicated cache (e.g., Redis)
TASK_RESULTS = {}

User = get_user_model()

class DummyRequest:
    """
    A simplified mock for HttpRequest object to pass to handle_search/handle_export.
    """
    def __init__(self, user_obj):
        self.user = user_obj
        self.GET = {} # handle_search/export might access request.GET
        self.method = 'GET' # Assume GET for assistant queries

def run_ollama_task(task_id: str, query_text: str, user_id: int):
    """
    Executes the Ollama query and subsequent data processing in a background thread.
    Stores the result or error in the global TASK_RESULTS dictionary.
    """
    TASK_RESULTS[task_id] = {'status': 'PENDING', 'result': None, 'error': None}

    try:
        user = User.objects.get(id=user_id)
        dummy_request = DummyRequest(user)

        # Call Ollama model to parse the natural language query
        ollama_response = get_ollama_response(query_text, user_id)

        if "error" in ollama_response:
            TASK_RESULTS[task_id].update({'status': 'FAILURE', 'error': ollama_response['error']})
            return

        # Extract parameters from Ollama's structured response
        intent = ollama_response.get("intent", "SEARCH")
        data_source = ollama_response.get("data_source", "Requisition")
        filters = ollama_response.get("filters", {})
        aggregation = ollama_response.get("aggregation", {})

        # Prepare parameters for action handlers
        parameters = {
            "data_source": data_source,
            "filters": filters,
            "aggregation": aggregation
        }

        # Dispatch actions
        if intent == "EXPORT":
            parameters['file_type'] = "Excel"
            # handle_export returns HttpResponse, which is not directly JSON serializable.
            # We need to store the file content and return a URL or base64 encoded data.
            # For simplicity, let's assume handle_export can be modified to return file content.
            # For now, we'll just indicate success.
            # A better approach would be to save the file to media and return its URL.
            TASK_RESULTS[task_id].update({'status': 'FAILURE', 'error': 'Export via async task not fully implemented yet. Please use search for now.'})
            return
        elif intent == "SEARCH":
            # handle_search returns JsonResponse, extract its content
            search_response = handle_search(dummy_request, parameters)
            if isinstance(search_response, JsonResponse):
                result_data = json.loads(search_response.content)
                TASK_RESULTS[task_id].update({'status': 'SUCCESS', 'result': result_data})
            else:
                TASK_RESULTS[task_id].update({'status': 'FAILURE', 'error': 'Search handler returned unexpected response.'})
        else:
            TASK_RESULTS[task_id].update({'status': 'FAILURE', 'error': '無法理解您的指令，請嘗試更明確的表達。'})

    except Exception as e:
        TASK_RESULTS[task_id].update({'status': 'FAILURE', 'error': str(e)})
    finally:
        # In a real system, you might want to set a timeout for task results
        # or implement a cleanup mechanism.
        pass
