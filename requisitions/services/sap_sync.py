"""
SAP / MPS 外部資料同步服務
"""
import io
import pandas as pd
import requests
from django.db import transaction
from requisitions.models import WorkOrder

def sync_external_order_info():
    """
    從外部 API (MPS) 同步工單的出貨日期與客戶資訊
    """
    api_url = "http://192.168.1.89/MPS/GetExcel?depno=F004"
    try:
        response = requests.get(api_url, timeout=60)
        response.raise_for_status()
        
        content = response.content
        if not content:
            return False, "API 回傳內容為空"

        # 嘗試使用不同的引擎讀取 Excel (xlsx 用 openpyxl, xls 用 xlrd)
        df_upload = None
        error_msgs = []

        # 1. 嘗試新版 xlsx (openpyxl)
        try:
            df_upload = pd.read_excel(io.BytesIO(content), dtype=str, engine='openpyxl')
        except Exception as e:
            error_msgs.append(f"xlsx 讀取失敗: {str(e)}")
            
            # 2. 嘗試舊版 xls (xlrd)
            try:
                df_upload = pd.read_excel(io.BytesIO(content), dtype=str, engine='xlrd')
            except Exception as e2:
                error_msgs.append(f"xls 讀取失敗: {str(e2)}")
                
                # 3. 嘗試讀取為 HTML (有些系統回傳的是假的 xls, 其實是 html table)
                try:
                    dfs = pd.read_html(io.BytesIO(content))
                    if dfs:
                        df_upload = dfs[0]
                        # 轉換所有欄位為字串
                        df_upload = df_upload.astype(str)
                except Exception as e3:
                    error_msgs.append(f"HTML 讀取失敗: {str(e3)}")

        if df_upload is None:
            # 如果都失敗了，回傳最後的錯誤訊息，並檢查是否為純文字 (可能是 API 報錯)
            try:
                text_content = content.decode('utf-8')[:200]
                return False, f"無法解析 Excel 內容。回傳內容前段：{text_content}"
            except:
                return False, f"無法解析 Excel 內容。詳細錯誤：{'; '.join(error_msgs)}"

        # --- 強力偵測標頭列 ---
        # 有些 Excel 前面幾行是空的或是標題，我們掃描前 10 行找出真正的 header
        actual_header_index = 0
        found_header = False
        
        # 先轉成 list of list 方便掃描
        all_values = df_upload.values.tolist()
        # 把原本的 columns 也當作第一行加入
        all_values.insert(0, df_upload.columns.tolist())
        
        target_keywords = ['Order NO', 'Order', '工單', '訂單']
        for i, row in enumerate(all_values[:10]):
            row_str = [str(cell) for cell in row]
            if any(key in s for s in row_str for key in target_keywords):
                actual_header_index = i
                found_header = True
                break
        
        if found_header:
            # 重新以該行作為 header
            if actual_header_index == 0:
                # header 就是原本的 columns，不用動
                pass
            else:
                # 以該行重新讀取
                df_upload = pd.read_excel(io.BytesIO(content), skiprows=actual_header_index, dtype=str)
        
        df_upload.columns = df_upload.columns.str.strip()
        
        # 再次偵測欄位
        order_col = None
        for col_name in ['Machine Number', 'Order NO', 'Order', '工單號碼', '工單單號', '訂單單號', '訂單', '工單']:
            if col_name in df_upload.columns:
                order_col = col_name
                break
        
        shipping_date_col = None
        for col_name in ['Order Scheduled Shipping Date', '訂單預出貨日', '客戶原始預交日', '出貨日期', '預交日期', '交期']:
            if col_name in df_upload.columns:
                shipping_date_col = col_name
                break
                
        customer_col = None
        for col_name in ['Order Customer Name', '下單客戶名稱', '客戶名稱', '客戶', '客戶名']:
            if col_name in df_upload.columns:
                customer_col = col_name
                break
        
        if not order_col:
            # 如果找不到欄位，可能是因為 Excel 格式不對
            available_cols = ", ".join([str(c) for c in df_upload.columns[:10]])
            return False, f"找不到單號欄位。偵測到的欄位有：{available_cols}"

        # --- 偵錯日誌：記錄抓到了什麼 ---
        try:
            with open('scratch/sync_debug.log', 'w', encoding='utf-8') as f:
                f.write(f"Detected Columns: order={order_col}, customer={customer_col}, date={shipping_date_col}\n")
                f.write(f"All Columns: {list(df_upload.columns)}\n")
                f.write("First 5 rows data:\n")
                f.write(df_upload.head(5).to_string())
        except:
            pass

        updated_count = 0
        with transaction.atomic():
            for _, row in df_upload.iterrows():
                # 去掉前導零，確保與系統內的工單號碼格式一致
                order_number = str(row.get(order_col, '')).strip().lstrip('0')
                if not order_number or order_number == 'nan' or order_number == '': continue
                
                defaults = {}
                # 解析日期
                if shipping_date_col:
                    date_val = row.get(shipping_date_col)
                    if pd.notna(date_val) and str(date_val).strip() not in ['nan', '', 'None']:
                        try:
                            # 支援多種日期格式
                            defaults['shipping_date'] = pd.to_datetime(date_val).date()
                        except: pass
                
                # 解析客戶
                if customer_col:
                    val = row.get(customer_col)
                    if pd.notna(val) and str(val).strip() not in ['nan', '', 'None']:
                        defaults['customer_name'] = str(val).strip()
                
                if defaults:
                    WorkOrder.objects.update_or_create(
                        order_number=order_number,
                        defaults=defaults
                    )
                    updated_count += 1
                    
        return True, f"同步成功，已更新 {updated_count} 筆工單的出貨資訊。"
    except Exception as e:
        return False, f"同步失敗：{str(e)}"

