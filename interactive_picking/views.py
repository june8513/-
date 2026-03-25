from django.shortcuts import render, redirect, get_object_or_404
from django.contrib.auth.decorators import login_required
from django.utils import timezone
from django.http import JsonResponse
from decimal import Decimal, InvalidOperation
import uuid
import pandas as pd
from django.contrib import messages

from .models import WarehouseLocation
from requisitions.models import Requisition, RequisitionItem, WorkOrderMaterialTransaction
from .forms import UploadWarehouseLocationForm

def _get_warehouse_location(storage_bin):
    """
    智慧對應儲位座標 (前綴比對)：
    例如物料在 'A01-01-01' (A01區 01架 01層)
    若地圖只有建立 'A01-01' 的平面座標，演算法也會自動將該物料對應到 'A01-01' 返回。
    """
    if not storage_bin:
        return None
        
    storage_bin = str(storage_bin).strip().upper()
    
    # 1. 嘗試完全比對 (忽略大小寫)
    loc = WarehouseLocation.objects.filter(name__iexact=storage_bin).first()
    if loc:
        return loc
        
    # 2. 嘗試前綴比對 (找出最長吻合的前綴，忽略大小寫)
    all_locs = WarehouseLocation.objects.all()
    best_loc = None
    longest_match = 0
    for l in all_locs:
        db_name = l.name.strip().upper()
        if storage_bin.startswith(db_name) and len(db_name) > longest_match:
            best_loc = l
            longest_match = len(db_name)
            
    return best_loc


@login_required
def index(request):
    """
    任務清單首頁：列出所有真實「尚未撥料完成」的申請單 (包含已提交與撥料中)
    """
    base_query = Requisition.objects.filter(status__in=['demand_submitted', 'dispatch_in_progress'])
    
    # 取得不重複的投料點清單供篩選器使用 (排除 None/空值)
    process_types = base_query.exclude(process_type__isnull=True).exclude(process_type='').values_list('process_type', flat=True).distinct()
    
    # 處理前端傳來的篩選條件
    selected_process_type = request.GET.get('process_type')
    if selected_process_type:
        tasks = base_query.filter(process_type=selected_process_type).order_by('-created_at')
    else:
        tasks = base_query.order_by('-created_at')
    
    # 計算每個單據剩餘未處理項目
    task_data = []
    for task in tasks:
        pending_count = task.items.filter(dispatch_status__isnull=True).count()
        task_data.append({
            'id': task.id,
            'task_number': task.order_number,
            'process_type': task.process_type,
            'created_at': task.created_at,
            'pending_count': pending_count,
            'is_completed': pending_count == 0,
            'total_count': task.items.count(),
        })
        
    return render(request, 'interactive_picking/index.html', {
        'tasks': task_data,
        'process_types': process_types,
        'selected_process_type': selected_process_type
    })
def _get_next_optimized_picking_item(task, current_x=0.0, current_y=0.0):
    """
    核心演算法：給定當前座標 (預設起點 0,0)，計算並回傳距離最近的下一個待揀物料 (Nearest Neighbor)
    """
    pending_items = task.items.filter(dispatch_status__isnull=True)
    if not pending_items.exists():
        return None
        
    closest_item = None
    min_distance = float('inf')
    
    for item in pending_items:
        # 透過智慧字串比對找到最精確的對應座標節點 (支援前綴)
        loc = _get_warehouse_location(item.storage_bin)
        
        if not loc:
            dist = 999999.0 # 如果沒有建檔座標，給予極大距離排在後面
        else:
            # 計算歐幾里得距離
            dist = ((loc.x_coordinate - current_x)**2 + (loc.y_coordinate - current_y)**2) ** 0.5
            
        if dist < min_distance:
            min_distance = dist
            closest_item = item
            
    if not closest_item:
        closest_item = pending_items.first()
            
    return closest_item


def _get_sorted_picking_list(task, start_x=0.0, start_y=0.0):
    """
    使用最近鄰啟發式 (Nearest Neighbor) 建立完整的路徑優化排序清單。
    效能優化：一次性載入所有 WarehouseLocation 到記憶體快取，避免 N² 次 DB 查詢。
    """
    pending = list(task.items.filter(dispatch_status__isnull=True))
    if not pending:
        return []
    
    # ★ 效能關鍵：一次載入所有座標，建立 name(大寫) -> loc 的字典快取
    all_locs = list(WarehouseLocation.objects.all())
    loc_cache = {}
    for loc in all_locs:
        loc_cache[loc.name.strip().upper()] = loc
    
    def _find_loc_fast(storage_bin):
        """用記憶體快取做前綴比對，不再打資料庫"""
        if not storage_bin:
            return None
        key = str(storage_bin).strip().upper()
        # 完全比對
        if key in loc_cache:
            return loc_cache[key]
        # 前綴比對 (最長吻合)
        best = None
        longest = 0
        for name, loc in loc_cache.items():
            if key.startswith(name) and len(name) > longest:
                best = loc
                longest = len(name)
        return best
    
    # 預先算好每個 item 對應的座標 (只算一次)
    item_loc_map = {}
    for item in pending:
        item_loc_map[item.id] = _find_loc_fast(item.storage_bin)
    
    sorted_list = []
    current_x, current_y = start_x, start_y
    
    while pending:
        closest_item = None
        min_distance = float('inf')
        
        for item in pending:
            loc = item_loc_map.get(item.id)
            if not loc:
                dist = 999999.0
            else:
                dist = ((loc.x_coordinate - current_x)**2 + (loc.y_coordinate - current_y)**2) ** 0.5
            
            if dist < min_distance:
                min_distance = dist
                closest_item = item
        
        if closest_item:
            sorted_list.append(closest_item)
            pending.remove(closest_item)
            loc = item_loc_map.get(closest_item.id)
            if loc:
                current_x = loc.x_coordinate
                current_y = loc.y_coordinate
        else:
            sorted_list.extend(pending)
            break
    
    return sorted_list


@login_required
def picking_wizard(request, task_id):
    """
    互動式揀料精靈視圖：
    1. 接收任務 ID
    2. 建立完整的路徑優化排序清單
    3. 支援透過 item_index 參數在待揀物料間切換 (上一筆/下一筆)
    """
    task = get_object_or_404(Requisition, id=task_id)
    
    # 檢查是否已全部完成
    pending_count = task.items.filter(dispatch_status__isnull=True).count()
    if pending_count == 0:
        if task.status != 'dispatch_completed':
            task.status = 'dispatch_completed'
            task.dispatch_performed = True
            task.save()
    else:
        # 單據若在「已提交」尚未變更狀態時，進入精靈即改為「撥料中」
        if task.status == 'demand_submitted':
            task.status = 'dispatch_in_progress'
            task.save()
        
    # 取得當前人員位置 (模擬從 session，如果沒有則從 0,0 起點計算)
    current_x = request.session.get('current_picking_x', 0.0)
    current_y = request.session.get('current_picking_y', 0.0)
    
    # 建立完整的路徑優化排序清單 (使用最近鄰啟發式)
    sorted_items = _get_sorted_picking_list(task, current_x, current_y)
    
    # 透過 GET 參數取得目前的索引位置
    try:
        item_index = int(request.GET.get('item_index', 0))
    except (ValueError, TypeError):
        item_index = 0
    
    # 邊界保護
    if len(sorted_items) > 0:
        item_index = max(0, min(item_index, len(sorted_items) - 1))
        next_item = sorted_items[item_index]
    else:
        next_item = None
        item_index = 0
    
    # 防錯：如果有 pending item 但沒有 location 的極端情況
    if not next_item and pending_count > 0:
        next_item = task.items.filter(dispatch_status__isnull=True).first()

    # 取得最後一筆被處理的項目 (由 session 取得)，用於「回上一步」功能
    last_item_id = request.session.get('last_picked_item_id')
    last_processed_item = None
    if last_item_id:
        last_processed_item = task.items.filter(id=last_item_id).first()

    # 檢查該次拿取的品項是否有建檔真正的地圖座標 (或母區域座標)
    missing_loc = False
    if next_item:
        loc_check = _get_warehouse_location(next_item.storage_bin)
        if not loc_check:
            missing_loc = True

    context = {
        'task': task,
        'item': next_item,
        'pending_count': pending_count,
        'total_count': task.items.count(),
        'last_processed_item': last_processed_item,
        'missing_loc': missing_loc,
        # 上下筆導航
        'item_index': item_index,
        'item_total': len(sorted_items),
        'has_prev': item_index > 0,
        'has_next': item_index < len(sorted_items) - 1,
        'prev_index': item_index - 1,
        'next_index': item_index + 1,
    }
    
    return render(request, 'interactive_picking/wizard.html', context)


@login_required
def process_picking_action(request, item_id):
    """
    處理撥料人員按下「撥料」或「缺料」的動作
    """
    if request.method == 'POST':
        item = get_object_or_404(RequisitionItem, id=item_id)
        action = request.POST.get('action') # 'picked' or 'shortage'
        actual_quantity = request.POST.get('actual_quantity')
        
        if action in ['picked', 'shortage']:
            if action == 'picked':
                item.dispatch_status = 'dispatched'
                try:
                    actual_qty_dec = Decimal(actual_quantity)
                    item.confirmed_quantity = actual_qty_dec
                    
                    # 同步回寫至主物料清單 (WorkOrderMaterial) 並記錄交易足跡
                    if item.source_material:
                        wom = item.source_material
                        current_confirmed = wom.confirmed_quantity if wom.confirmed_quantity is not None else Decimal('0')
                        new_confirmed = current_confirmed + actual_qty_dec
                        wom.confirmed_quantity = new_confirmed
                        wom.save()
                        
                        WorkOrderMaterialTransaction.objects.create(
                            work_order_material=wom,
                            user=request.user,
                            transaction_type='ALLOCATION',
                            quantity_change=actual_qty_dec,
                            new_confirmed_quantity=new_confirmed,
                            notes="由互動式揀料系統操作"
                        )
                except (InvalidOperation, TypeError):
                    item.confirmed_quantity = Decimal('0')
            else:
                item.dispatch_status = 'backordered'
                item.confirmed_quantity = Decimal('0')
            
            item.save()
            
            # 使用 Session 紀錄上一次處理的 ID 作為 Undo 來源
            request.session['last_picked_item_id'] = item.id
            
            # 更新 Session 中的當前座標，讓下一步演算法能從該點出發
            loc = _get_warehouse_location(item.storage_bin)
            if loc:
                request.session['current_picking_x'] = loc.x_coordinate
                request.session['current_picking_y'] = loc.y_coordinate
                
    return redirect('interactive_picking:wizard', task_id=item.requisition.id)


@login_required
def undo_last_action(request, task_id):
    """
    撤銷最後一次的動作：將最後處理的物料改回待處理狀態。
    """
    if request.method == 'POST':
        task = get_object_or_404(Requisition, id=task_id)
        
        last_item_id = request.session.get('last_picked_item_id')
        if last_item_id:
            last_item = task.items.filter(id=last_item_id).first()
            if last_item:
                # 如果是 Undo 一筆原本是有已拿取數量的，也要把 WorkOrderMaterial 加回去的扣掉
                if last_item.dispatch_status == 'dispatched' and last_item.confirmed_quantity and last_item.source_material:
                    wom = last_item.source_material
                    current_confirmed = wom.confirmed_quantity if wom.confirmed_quantity is not None else Decimal('0')
                    # 確保扣除後不低於 0
                    new_confirmed = max(Decimal('0'), current_confirmed - last_item.confirmed_quantity)
                    wom.confirmed_quantity = new_confirmed
                    wom.save()
                    
                    WorkOrderMaterialTransaction.objects.create(
                        work_order_material=wom,
                        user=request.user,
                        transaction_type='RETURN',
                        quantity_change=-last_item.confirmed_quantity,
                        new_confirmed_quantity=new_confirmed,
                        notes="由互動式揀料系統退回上一步"
                    )

                # 還原 RequisitionItem 狀態
                last_item.dispatch_status = None
                last_item.confirmed_quantity = None
                last_item.save()
                
                # 因為退回了物品，如果任務本來已經標記完成，也要改回未完成
                if task.status == 'dispatch_completed':
                    task.status = 'dispatch_in_progress'
                    task.save()
                

                # 清除 session 紀錄避免重複復原
                if 'last_picked_item_id' in request.session:
                    del request.session['last_picked_item_id']
                
    return redirect('interactive_picking:wizard', task_id=task_id)


@login_required
def upload_warehouse_locations(request):
    """
    透過上傳 Excel 批次建檔倉儲儲位 (WarehouseLocation) X/Y 座標
    若您的使用者不是管理員，建議限制權限
    """
    if not request.user.is_superuser:
         messages.error(request, "只有管理員才能匯入倉儲圖資。")
         return redirect('interactive_picking:index')

    if request.method == 'POST':
        # API: 處理前端 3D 互動地圖的 AJAX 請求
        if request.content_type == 'application/json':
            import json
            try:
                data = json.loads(request.body)
                action = data.get('action')
                if action == 'save':
                    name = str(data.get('name', '')).strip().upper()
                    old_name = data.get('old_name', '').strip().upper() # 用於判斷是否編輯既有標籤
                    x = float(data.get('x', 0))
                    y = float(data.get('y', 0))
                    floor = int(data.get('floor', 1))
                    
                    if not name:
                        return JsonResponse({'status': 'error', 'msg': '儲位代號不可為空'}, status=400)
                        
                    # 檢查新名稱是否已被其他座標使用 (排除掉原本自己)
                    duplicate = WarehouseLocation.objects.filter(name=name).exclude(name=old_name).first()
                    if duplicate:
                        return JsonResponse({'status': 'error', 'msg': f'代號 [{name}] 已存在於座標 (X:{duplicate.x_coordinate}, Y:{duplicate.y_coordinate})'}, status=400)
                        
                    if old_name:
                        # 更新既有標籤 (不論舊資料是大寫還是小寫)
                        loc = WarehouseLocation.objects.filter(name__iexact=old_name).first()
                        if loc:
                            loc.name = name
                            loc.x_coordinate = x
                            loc.y_coordinate = y
                            loc.floor = floor
                            loc.save()
                        else:
                            WarehouseLocation.objects.create(name=name, x_coordinate=x, y_coordinate=y, floor=floor)
                    else:
                        # 在同一座標下新增另一個層架標籤
                        WarehouseLocation.objects.create(name=name, x_coordinate=x, y_coordinate=y, floor=floor)
                        
                    return JsonResponse({'status': 'ok', 'msg': f'已儲存儲位 {name}'})
                    
                elif action == 'delete':
                    name = str(data.get('name', '')).strip().upper()
                    if name:
                        deleted_count, _ = WarehouseLocation.objects.filter(name__iexact=name).delete()
                        if deleted_count > 0:
                            return JsonResponse({'status': 'ok', 'msg': f'已成功刪除標籤：{name}'})
                        else:
                            return JsonResponse({'status': 'error', 'msg': f'找不到要刪除的標籤：{name}'})
                    return JsonResponse({'status': 'error', 'msg': '未提供刪除目標名稱'})
                    
            except Exception as e:
                return JsonResponse({'status': 'error', 'msg': str(e)}, status=500)
        
        # 傳統表單: Excel 上傳
        form = UploadWarehouseLocationForm(request.POST, request.FILES)
        if form.is_valid():
            excel_file = request.FILES['excel_file']
            try:
                df = pd.read_excel(excel_file, engine='openpyxl')
                
                # 預期欄位名稱 mapping
                # 例如可以接受 "儲位代號", "X座標", "Y座標", "描述"
                required_cols = ['儲位代號', 'X座標', 'Y座標']
                for col in required_cols:
                    if col not in df.columns:
                        messages.error(request, f"上傳失敗：找不到必要欄位 '{col}'")
                        return redirect('interactive_picking:upload_locations')

                created_count = 0
                updated_count = 0
                
                for index, row in df.iterrows():
                    name = str(row['儲位代號']).strip()
                    if not name or name == 'nan':
                        continue
                        
                    x_coord = float(row.get('X座標', 0.0))
                    y_coord = float(row.get('Y座標', 0.0))
                    floor_val = str(row.get('樓層', '1')).strip()
                    try:
                        floor_val = int(floor_val)
                    except ValueError:
                        floor_val = 1
                    
                    desc = str(row.get('描述', '')).strip()
                    if desc == 'nan':
                        desc = ''

                    obj, created = WarehouseLocation.objects.update_or_create(
                        name=name,
                        defaults={
                            'x_coordinate': x_coord,
                            'y_coordinate': y_coord,
                            'floor': floor_val,
                            'description': desc
                        }
                    )
                    if created:
                        created_count += 1
                    else:
                        updated_count += 1
                
                messages.success(request, f"成功匯入！新增 {created_count} 筆儲位，更新 {updated_count} 筆儲位座標。")
                return redirect('interactive_picking:index')

            except Exception as e:
                messages.error(request, f"解析 Excel 失敗：{str(e)}")
                return redirect('interactive_picking:upload_locations')
    else:
        form = UploadWarehouseLocationForm()


    # 一併取得現有清單以供檢視與渲染
    locations = WarehouseLocation.objects.all().order_by('name')
    
    import json
    locations_json = json.dumps([
        {'id': loc.id, 'name': loc.name, 'x': loc.x_coordinate, 'y': loc.y_coordinate, 'floor': loc.floor}
        for loc in locations
    ])
    
    return render(request, 'interactive_picking/upload_locations.html', {
        'form': form,
        'locations': locations,
        'locations_json': locations_json
    })
