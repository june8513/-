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

def run_ollama_task(task_id: str, query_text: str, user_id: int, history: list = None):
    """
    Executes the Ollama query and subsequent data processing in a background thread.
    Stores the result or error in the global TASK_RESULTS dictionary.
    """
    TASK_RESULTS[task_id] = {'status': 'PENDING', 'result': None, 'error': None}

    try:
        user = User.objects.get(id=user_id)
        dummy_request = DummyRequest(user)

        # Call Ollama model to parse the natural language query
        ollama_response = get_ollama_response(query_text, user_id, history)

        if "error" in ollama_response and "response_to_user" not in ollama_response:
            TASK_RESULTS[task_id].update({'status': 'FAILURE', 'error': ollama_response['error']})
            return

        # Extract textual response
        response_to_user = ollama_response.get("response_to_user", "好的，我為您處理中。")
        action = ollama_response.get("action", {})

        if not action:
            # Just a conversational response, no system action needed
            TASK_RESULTS[task_id].update({
                'status': 'SUCCESS', 
                'result': {'response_to_user': response_to_user, 'results': []}
            })
            return

        # Extract parameters from Ollama's structured action
        intent = action.get("intent", "SEARCH")
        data_source = action.get("data_source")
        filters = action.get("filters", {})
        aggregation = action.get("aggregation", {})

        if not data_source:
             TASK_RESULTS[task_id].update({
                'status': 'SUCCESS', 
                'result': {'response_to_user': response_to_user, 'results': []}
            })
             return

        # Prepare parameters for action handlers
        parameters = {
            "data_source": data_source,
            "filters": filters,
            "aggregation": aggregation
        }

        # Dispatch actions
        if intent == "EXPORT":
            TASK_RESULTS[task_id].update({'status': 'FAILURE', 'error': '匯出功能目前需由手動點擊，建議先使用查詢查看結果。'})
            return
        elif intent == "SEARCH":
            # handle_search returns JsonResponse, extract its content
            search_response = handle_search(dummy_request, parameters)
            if isinstance(search_response, JsonResponse):
                result_data = json.loads(search_response.content)
                # Combine conversational response with data results
                final_result = {
                    'response_to_user': response_to_user,
                    'results': result_data.get('results', [])
                }
                TASK_RESULTS[task_id].update({'status': 'SUCCESS', 'result': final_result})
            else:
                TASK_RESULTS[task_id].update({'status': 'FAILURE', 'error': '資料庫查詢回傳異常。'})
        else:
            TASK_RESULTS[task_id].update({
                'status': 'SUCCESS', 
                'result': {'response_to_user': response_to_user, 'results': []}
            })
    except Exception as e:
        TASK_RESULTS[task_id].update({'status': 'FAILURE', 'error': str(e)})

    except Exception as e:
        TASK_RESULTS[task_id].update({'status': 'FAILURE', 'error': str(e)})
    finally:
        # In a real system, you might want to set a timeout for task results
        # or implement a cleanup mechanism.
        pass
