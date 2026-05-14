"""
同步外部交期 API 資料到 WorkOrderMaterial.estimated_arrival_date

從 http://192.168.6.119:5002/api/delivery/nearest 取得各物料的預計交期，
並更新到資料庫中。
"""
import sys
import os
import requests
from datetime import datetime

# Add the project directory to the Python path
project_path = os.path.dirname(os.path.abspath(__file__))
sys.path.append(project_path)

# Set up Django environment
os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'material_requisition_system.settings.production')

import django
django.setup()

from requisitions.models import WorkOrderMaterial

# API 設定
API_URL = "http://192.168.6.119:5002/api/delivery/nearest"
API_TIMEOUT = 30  # 秒


def fetch_delivery_dates():
    """
    從外部 API 取得交期資料
    
    Returns:
        dict: API 回傳的 data 物件，key 為物料編號，value 為交期資訊
        None: 如果 API 呼叫失敗
    """
    try:
        print(f"正在連線至 API: {API_URL}")
        response = requests.get(API_URL, timeout=API_TIMEOUT)
        response.raise_for_status()
        
        result = response.json()
        data = result.get('data', {})
        count = result.get('count', 0)
        timestamp = result.get('timestamp', '')
        
        print(f"成功取得 {count} 筆物料交期資料 (時間戳記: {timestamp})")
        return data
        
    except requests.exceptions.Timeout:
        print(f"錯誤: API 連線逾時 ({API_TIMEOUT}秒)")
        return None
    except requests.exceptions.ConnectionError:
        print(f"錯誤: 無法連線至 API ({API_URL})")
        return None
    except requests.exceptions.RequestException as e:
        print(f"錯誤: API 請求失敗 - {e}")
        return None
    except ValueError as e:
        print(f"錯誤: 無法解析 API 回應 - {e}")
        return None


def sync_to_database(delivery_data):
    """
    將取得的交期資料同步到資料庫
    
    Args:
        delivery_data: dict, key 為物料編號，value 為交期資訊
        
    Returns:
        tuple: (更新成功數, 跳過數, 錯誤數)
    """
    if not delivery_data:
        print("沒有資料需要同步")
        return (0, 0, 0)
    
    updated_count = 0
    skipped_count = 0
    error_count = 0
    
    for material_number, info in delivery_data.items():
        try:
            date_str = info.get('date')
            if not date_str:
                skipped_count += 1
                continue
            
            # 解析日期
            try:
                arrival_date = datetime.strptime(date_str, '%Y-%m-%d').date()
            except ValueError:
                print(f"  警告: 物料 {material_number} 的日期格式無效: {date_str}")
                error_count += 1
                continue
            
            # 更新資料庫 - 只更新 is_active=True 的記錄
            affected_rows = WorkOrderMaterial.objects.filter(
                material_number=material_number,
                is_active=True
            ).update(estimated_arrival_date=arrival_date)
            
            if affected_rows > 0:
                source = info.get('source', 'unknown')
                ref_id = info.get('ref_id', '')
                print(f"  ✓ {material_number}: {date_str} ({source}, ref: {ref_id}) - 更新 {affected_rows} 筆")
                updated_count += 1
            else:
                skipped_count += 1
                
        except Exception as e:
            print(f"  ✗ 處理物料 {material_number} 時發生錯誤: {e}")
            error_count += 1
    
    return (updated_count, skipped_count, error_count)


def run_sync_estimated_arrival_dates():
    """
    主執行函數 - 供 run_all_monitors.py 呼叫
    """
    print("=" * 50)
    print("開始同步預計入料日期...")
    print("=" * 50)
    
    # 1. 從 API 取得資料
    delivery_data = fetch_delivery_dates()
    
    if delivery_data is None:
        print("\n同步失敗: 無法取得 API 資料")
        return False
    
    if not delivery_data:
        print("\n同步完成: API 回傳空資料")
        return True
    
    # 2. 同步到資料庫
    print(f"\n開始更新資料庫...")
    updated, skipped, errors = sync_to_database(delivery_data)
    
    # 3. 輸出結果
    print("\n" + "-" * 50)
    print(f"同步結果:")
    print(f"  - 更新成功: {updated} 種物料")
    print(f"  - 跳過 (無對應記錄): {skipped} 種物料")
    print(f"  - 錯誤: {errors} 種物料")
    print("=" * 50)
    
    return errors == 0


if __name__ == "__main__":
    success = run_sync_estimated_arrival_dates()
    sys.exit(0 if success else 1)
