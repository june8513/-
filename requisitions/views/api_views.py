from django.http import JsonResponse
import spacy
import json
from datetime import date, datetime, timedelta # Import date and timedelta
import re # Import re for regex
from .action_handlers import handle_search, handle_export

# Load the spaCy Chinese model
# This should ideally be loaded once when the application starts, not on every request.
# For simplicity in this example, we'll load it here. In a real app, use a global variable or Django's AppConfig.
try:
    nlp = spacy.load("zh_core_web_sm")
except OSError:
    # Handle case where model is not found, e.g., if it wasn't downloaded correctly
    print("spaCy model 'zh_core_web_sm' not found. Please run 'python -m spacy download zh_core_web_sm'")
    nlp = None # Or raise an exception to prevent the app from starting

def natural_action_view(request):
    if request.method != 'GET':
        return JsonResponse({'error': 'Only GET requests are supported'}, status=405)

    query_text = request.GET.get('q', '')
    if not query_text:
        return JsonResponse({'error': 'Query parameter "q" is missing'}, status=400)

    if nlp is None:
        return JsonResponse({'error': 'spaCy model not loaded. Please check server logs.'}, status=500)

    doc = nlp(query_text)

    # --- Phase 1: spaCy-based Rule Engine Logic ---
    # This is where the core logic for parsing natural language into structured queries will go.
    # We will extract intent, data_source, filters, and aggregation parameters.

    intent = "SEARCH"  # Default intent
    data_source = "Requisition" # Default data source
    filters = {}
    aggregation = {} # Initialize aggregation here
    parameters = {}
    response_to_user = "我已收到您的查詢，正在處理中。"

    # Detect Data Source
    if "撥出" in query_text or "出庫" in query_text:
        data_source = "MaterialTransaction"
    elif "入料" in query_text:
        data_source = "WorkOrderMaterial"
        filters.pop('transaction_type', None) # Remove unnecessary filter
    elif "申請單" in query_text:
        data_source = "Requisition"
    # If no specific data source keyword, it defaults to Requisition
    else:
        data_source = "Requisition"

    # 2. Detect Intent
    if "excel" in query_text.lower() or "報表" in query_text or "匯出" in query_text:
        intent = "EXPORT"
        response_to_user = "好的，正在為您準備 Excel 報表。"
    elif "缺料" in query_text or "沒料了" in query_text:
        intent = "SEARCH"
        data_source = "Requisition" # Override data_source for '缺料'
        response_to_user = "正在查詢缺料資訊。"
    else:
        intent = "SEARCH"

    # --- Filter Extraction (Date Ranges) ---
    current_year = datetime.now().year
    today = datetime.now().date()
    start_date = None
    end_date = None

    # Keywords for relative dates
    if "今日" in query_text or "今天" in query_text:
        start_date = today
        end_date = today
    elif "昨日" in query_text or "昨天" in query_text:
        start_date = today - timedelta(days=1)
        end_date = today - timedelta(days=1)
    elif "本週" in query_text or "這週" in query_text:
        start_date = today - timedelta(days=today.weekday())
        end_date = start_date + timedelta(days=6)
    elif "上週" in query_text:
        start_date = today - timedelta(days=today.weekday() + 7)
        end_date = start_date + timedelta(days=6)
    elif "本月" in query_text or "這個月" in query_text:
        start_date = today.replace(day=1)
        end_date = (today.replace(day=1) + timedelta(days=32)).replace(day=1) - timedelta(days=1)
    elif "上個月" in query_text:
        end_date = today.replace(day=1) - timedelta(days=1)
        start_date = end_date.replace(day=1)
    elif "今年" in query_text:
        start_date = date(current_year, 1, 1)
        end_date = date(current_year, 12, 31)
    elif "去年" in query_text:
        start_date = date(current_year - 1, 1, 1)
        end_date = date(current_year - 1, 12, 31)

    # Specific date parsing (e.g., '2025年10月26日')
    date_match = re.search(r'(\d{4})年(\d{1,2})月(\d{1,2})日', query_text)
    if date_match:
        year, month, day = map(int, date_match.groups())
        try:
            specific_date = date(year, month, day)
            start_date = specific_date
            end_date = specific_date
        except ValueError:
            pass # Invalid date, ignore

    # Apply date filters if found
    if start_date and end_date:
        if data_source == "MaterialTransaction":
            filters['timestamp__date__gte'] = start_date
            filters['timestamp__date__lte'] = end_date
        elif data_source == "Requisition":
            filters['created_at__date__gte'] = start_date
            filters['created_at__date__lte'] = end_date
        elif data_source == "WorkOrderMaterial":
            filters['estimated_arrival_date__gte'] = start_date
            filters['estimated_arrival_date__lte'] = end_date

        # --- Filter Extraction (User) ---

        # --- Filter Extraction (User) ---
        if request.user.is_authenticated:
            if data_source == "Requisition":
                filters['applicant__id'] = request.user.id
            elif data_source == "MaterialTransaction":
                filters['user__id'] = request.user.id
            # Note: WorkOrderMaterial does not have a user field, so no filter is applied.

        # --- Filter Extraction (Status) ---
        if "待處理" in query_text:
            if data_source == "Requisition":
                filters['status'] = 'pending'
            elif data_source == "MaterialTransaction":
                pass # MaterialTransaction doesn't have a 'pending' status directly
        elif "已確認" in query_text or "已處理" in query_text:
            if data_source == "Requisition":
                filters['status'] = 'completed' # Assuming '已處理' maps to 'completed'
            elif data_source == "MaterialTransaction":
                pass # All MaterialTransactions are 'processed' once created
        elif "已撥料" in query_text:
            if data_source == "Requisition":
                filters['dispatch_performed'] = True
            elif data_source == "MaterialTransaction":
                filters['transaction_type'] = 'ALLOCATION'
        elif "缺料" in query_text:
            if data_source == "Requisition":
                filters['status'] = 'pending' # Or a more specific 'shortage' status if it exists
            pass # For WorkOrderMaterial, '缺料' implies required_quantity > confirmed_quantity

        # --- Aggregation Extraction (Example: '件數' for SUM) ---
        if "件數" in query_text or "總數" in query_text:
            aggregation = {
                "group_by": "material__material_code", # Use relation to group by material code
                "function": "SUM",
                "field": "quantity_change" # Correct field name for transactions
            }


    # --- User Context ---
    if request.user.is_authenticated:
        # Apply user-specific filters if intent is about 'my' work or dispatched items
        if "我" in query_text or "我的" in query_text:
            # if data_source == "DispatchedItem":
            #     filters['dispatched_by_id'] = request.user.id
            if data_source == "Requisition":
                # Assuming 'my' requisitions are those created by the user
                filters['created_by_id'] = request.user.id # You'll need to add created_by_id to your dummy Requisition model

    # --- Dispatch Actions ---
    parameters = {
        "data_source": data_source,
        "filters": filters,
        "aggregation": aggregation
    }

    if intent == "EXPORT":
        parameters['file_type'] = "Excel"
        return handle_export(request, parameters)
    elif intent == "SEARCH":
        return handle_search(request, parameters)
    else:
        return JsonResponse({'error': '無法理解您的指令，請嘗試更明確的表達。'}, status=400)
