import ollama
import json
import re
from datetime import date, datetime, timedelta
from django.contrib.auth.models import User
from requisitions.models import AIUserCorrection

def get_ollama_response(query_text: str, user_id: int, history: list = None) -> dict:
    today = datetime.now().date()
    yesterday = today - timedelta(days=1)
    
    # 獲取使用者的修正紀錄
    corrections = AIUserCorrection.objects.filter(is_active=True).order_by('-created_at')[:3]
    corr_p = ""
    if corrections:
        corr_p = "\n修正建議:" + "".join([f"\n- 問「{c.query_text}」時參考: {c.correction_text}" for c in corrections])

    system_prompt = f"""你是一個精簡的撥料系統管家，請用「繁體中文」回答。{corr_p}
今日日期: {today}。昨天: {yesterday}。
JSON 格式: {{ "response_to_user": "文字", "action": {{ "data_source": "Requisition"|"WorkOrderMaterial", "filters": {{...}} }} }}
規範:
1. 查「撥料單、狀態、列表」選 Requisition。
2. 查「物料、缺料、欠料、料呢」選 WorkOrderMaterial。問缺料必加 "is_shortage": true。
3. 遇到「這張/這筆/這單」，務必從 history 的 [系統內容] 提取 order_number 加入 filters。
4. 範例：{{ "response_to_user": "單號 100003482 的缺料如下：", "action": {{ "data_source": "WorkOrderMaterial", "filters": {{ "order_number": "100003482", "is_shortage": true }} }} }}"""

    messages = [{'role': 'system', 'content': system_prompt}]
    if history:
        recent = history[-6:]
        for i, msg in enumerate(recent):
            content = msg['content']
            if msg['role'] == 'assistant' and i < len(recent) - 2:
                content = re.sub(r'\s\[系統內容:.*?\]', '', content)
            messages.append({'role': msg['role'], 'content': content})
    messages.append({'role': 'user', 'content': query_text})
    
    try:
        response = ollama.chat(model='mistral:7b', messages=messages, format='json')
        return json.loads(response['message']['content'])
    except Exception as e:
        return {"response_to_user": "抱歉，解析出了點問題。", "action": {}}
