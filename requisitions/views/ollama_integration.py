import ollama
import json
from datetime import date, datetime, timedelta
from decimal import Decimal

# This function will interact with the Ollama model to parse natural language queries
def get_ollama_response(query_text: str, user_id: int) -> dict:
    """
    Sends the user's natural language query to the Ollama model and
    returns a structured JSON response containing intent, data_source, filters, and aggregation.
    """

    # Define the system prompt to guide the Ollama model
    system_prompt = f"""
    你是一個智慧自然語言助理，專門用於處理製造業的物料管理系統。
    你的任務是將用戶的自然語言查詢轉換為結構化的 JSON 物件，以便系統能夠執行資料庫查詢、匯出報表或執行其他操作。

    **可用的資料來源 (data_source) 及其關鍵字：**
    - "Requisition" (撥料申請單): 關鍵字包括 "申請單", "撥料單", "我的申請單"。
        - 可過濾欄位:
            - `created_at__date__gte`, `created_at__date__lte`: 日期範圍 (例如: "今年", "上個月", "2025年10月26日")
            - `status`: 狀態 (例如: "待處理" -> "demand_submitted", "已確認" -> "signed_off", "已處理" -> "signed_off", "已撥料" -> "dispatch_completed")
            - `applicant_id`: 申請人ID (當用戶說 "我的" 或 "我" 時，使用當前用戶的ID: {user_id})
            - `dispatch_performed`: 是否已撥料 (當用戶說 "已撥料" 時，設為 True)
    - "WorkOrderMaterial" (工單物料): 關鍵字包括 "入料", "工單物料", "缺料項目"。
        - 可過濾欄位:
            - `estimated_arrival_date__gte`, `estimated_arrival_date__lte`: 預計入料日期範圍
            - `is_shortage`: 是否缺料 (當用戶說 "缺料" 時，設為 True。這會轉換為 `required_quantity > confirmed_quantity`)
    - "MaterialTransaction" (物料交易紀錄): 關鍵字包括 "撥出", "出庫", "交易紀錄"。
        - 可過濾欄位:
            - `timestamp__date__gte`, `timestamp__date__lte`: 交易日期範圍
            - `user_id`: 操作人員ID (當用戶說 "我的" 或 "我" 時，使用當前用戶的ID: {user_id})
            - `transaction_type`: 交易類型 (例如: "撥出" -> "ALLOCATION")

    **可用的意圖 (intent)：**
    - "SEARCH": 查詢資料並顯示。
    - "EXPORT": 匯出資料到 Excel。

    **可用的聚合 (aggregation)：**
    - 如果用戶要求 "總數" 或 "件數"，則需要設定聚合。
    - 聚合物件應包含:
        - `function`: "SUM" (總和) 或 "COUNT" (計數)。
        - `field`: 要聚合的欄位 (例如: "quantity_change" 用於 MaterialTransaction 的總數)。
        - `group_by`: 可選。如果用戶要求按某個維度分組 (例如: "每種物料的總數")，則指定分組欄位 (例如: "material__material_code")。如果沒有指定，則為總計。

    **日期處理：**
    - "今日", "今天": 轉換為今天的日期。
    - "昨日", "昨天": 轉換為昨天的日期。
    - "本週", "這週": 轉換為本週的開始和結束日期。
    - "上週": 轉換為上週的開始和結束日期。
    - "本月", "這個月": 轉換為本月的開始和結束日期。
    - "上個月": 轉換為上個月的開始和結束日期。
    - "今年": 轉換為今年的開始和結束日期。
    - "去年": 轉換為去年的開始和結束日期。
    - 特定日期格式: "YYYY年MM月DD日" (例如: "2025年10月26日")。

    **輸出格式：**
    請嚴格以 JSON 格式輸出，不要包含任何額外的文字或解釋。
    JSON 範例：
    ```json
    {{
      "intent": "SEARCH",
      "data_source": "Requisition",
      "filters": {{
        "created_at__date__gte": "2025-01-01",
        "created_at__date__lte": "2025-12-31",
        "status": "signed_off",
        "applicant_id": {user_id}
      }},
      "aggregation": {{
        "function": "SUM",
        "field": "quantity_change",
        "group_by": "material__material_code"
      }}
    }}
    ```
    如果沒有聚合，`aggregation` 應為空物件 {{}}。
    如果沒有過濾條件，`filters` 應為空物件 {{}}。
    如果無法理解用戶的查詢，請返回一個包含錯誤訊息的 JSON：
    ```json
    {{
      "error": "無法理解您的指令，請嘗試更明確的表達。"
    }}
    ```
    請確保所有日期都格式化為 "YYYY-MM-DD"。
    請確保所有數值類型（如 user_id）都是正確的 JSON 數值。
    """

    # Get current date information for relative date calculations
    current_year = datetime.now().year
    today = datetime.now().date()

    # Construct the user message
    user_message = f"用戶查詢: {query_text}\n當前日期: {today.strftime('%Y-%m-%d')}"

    try:
        # Call Ollama model
        response = ollama.chat(
            model='phi3', # Use the phi3 model
            messages=[
                {'role': 'system', 'content': system_prompt},
                {'role': 'user', 'content': user_message},
            ],
            format='json' # Request JSON format output
        )
        
        # Extract and parse the JSON content
        response_content = response['message']['content']
        parsed_json = json.loads(response_content)
        return parsed_json

    except json.JSONDecodeError as e:
        print(f"Ollama response was not valid JSON: {response_content}. Error: {e}")
        return {"error": "Ollama 返回了無效的 JSON 格式。"}
    except Exception as e:
        print(f"Error calling Ollama: {e}")
        return {"error": f"與 Ollama 溝通時發生錯誤: {e}"}
