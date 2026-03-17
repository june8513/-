from django.shortcuts import render, redirect, get_object_or_404
from django.contrib.auth.decorators import login_required
from django.utils import timezone
from django.http import JsonResponse
import random
import uuid

from .models import WarehouseLocation, MockPickingTask, MockPickingItem

@login_required
def index(request):
    """
    任務清單首頁：列出所有模擬檢料任務
    """
    tasks = MockPickingTask.objects.all().order_by('-created_at')
    return render(request, 'interactive_picking/index.html', {'tasks': tasks})

@login_required
def generate_mock_task(request):
    """
    自動生成模擬揀料任務與假資料供測試使用
    """
    # 建立或確認倉儲節點存在 (假定一個簡單的 3x3 網格倉庫)
    locations = []
    zones = ['A', 'B', 'C']
    for x, zone in enumerate(zones):
        for y in range(1, 4):
            loc, created = WarehouseLocation.objects.get_or_create(
                name=f"{zone}-0{y}",
                defaults={'x_coordinate': x * 10, 'y_coordinate': y * 5}
            )
            locations.append(loc)
            
    # 建立一個新任務
    new_task = MockPickingTask.objects.create(
        task_number=f"MOCK-{uuid.uuid4().hex[:6].upper()}"
    )
    
    # 隨機挑選 3 - 6 個儲位產生揀料需求
    selected_locations = random.sample(locations, random.randint(3, 6))
    for i, loc in enumerate(selected_locations):
        MockPickingItem.objects.create(
            task=new_task,
            material_name=f"測試物料-Type{i+1}",
            quantity_required=random.randint(5, 50),
            location=loc
        )
        
    return redirect('interactive_picking:index')


def _get_next_optimized_picking_item(task, current_x=0.0, current_y=0.0):
    """
    核心演算法：給定當前座標 (預設起點 0,0)，計算並回傳距離最近的下一個待揀物料 (Nearest Neighbor)
    """
    pending_items = task.items.filter(status='pending')
    if not pending_items.exists():
        return None
        
    # 計算最近距離的項目
    closest_item = None
    min_distance = float('inf')
    
    for item in pending_items:
        if not item.location:
            continue
        # 計算歐幾里得距離
        dist = ((item.location.x_coordinate - current_x)**2 + (item.location.y_coordinate - current_y)**2) ** 0.5
        if dist < min_distance:
            min_distance = dist
            closest_item = item
            
    return closest_item


@login_required
def picking_wizard(request, task_id):
    """
    互動式揀料精靈視圖：
    1. 接收任務 ID
    2. 如果有前一步的座標 (由前端回傳或 session 紀錄)，以該座標計算下一個最近物料
    3. 渲染單步 UI，只顯示下一個應該拿的物料
    """
    task = get_object_or_404(MockPickingTask, id=task_id)
    
    # 檢查是否已全部完成
    pending_count = task.items.filter(status='pending').count()
    if pending_count == 0:
        if not task.is_completed:
            task.is_completed = True
            task.save()
        # 原本這裡會 return redirect，現在拿掉，讓它繼續渲染 wizard.html 以顯示「完成畫面」與「復原按鈕」
        
    # 取得當前人員位置 (模擬從 session，如果沒有則從 0,0 起點計算)
    current_x = request.session.get('current_picking_x', 0.0)
    current_y = request.session.get('current_picking_y', 0.0)
    
    # 呼叫演算法取得下一筆
    next_item = _get_next_optimized_picking_item(task, current_x, current_y)
    
    # 防錯：如果有 pending item 但沒有 location 的極端情況
    if not next_item and pending_count > 0:
        next_item = task.items.filter(status='pending').first()

    # 取得最後一筆被處理的項目 (依據 picked_at 排序)，用於「回上一步」功能
    last_processed_item = task.items.exclude(status='pending').order_by('-picked_at').first()

    context = {
        'task': task,
        'item': next_item,
        'pending_count': pending_count,
        'total_count': task.items.count(),
        'last_processed_item': last_processed_item
    }
    
    return render(request, 'interactive_picking/wizard.html', context)


@login_required
def process_picking_action(request, item_id):
    """
    處理撥料人員按下「撥料」或「缺料」的動作
    """
    if request.method == 'POST':
        item = get_object_or_404(MockPickingItem, id=item_id)
        action = request.POST.get('action') # 'picked' or 'shortage'
        actual_quantity = request.POST.get('actual_quantity')
        
        if action in ['picked', 'shortage']:
            item.status = action
            item.picked_at = timezone.now()
            
            # 儲存實際拿取數量 (如果是缺料或未填寫則預設 0 或不存)
            if actual_quantity and action == 'picked':
                try:
                    item.quantity_picked = int(actual_quantity)
                except ValueError:
                    item.quantity_picked = 0
            
            item.save()
            
            # 更新 Session 中的當前座標，讓下一步演算法能從該點出發
            if item.location:
                request.session['current_picking_x'] = item.location.x_coordinate
                request.session['current_picking_y'] = item.location.y_coordinate
                
    return redirect('interactive_picking:wizard', task_id=item.task.id)


@login_required
def undo_last_action(request, task_id):
    """
    撤銷最後一次的動作：將最後處理的物料改回待處理狀態。
    """
    if request.method == 'POST':
        task = get_object_or_404(MockPickingTask, id=task_id)
        
        # 找尋這個任務底下，最後一次被標記為 picked 或 shortage 的項目
        last_item = task.items.exclude(status='pending').order_by('-picked_at').first()
        
        if last_item:
            # 還原狀態
            last_item.status = 'pending'
            last_item.picked_at = None
            last_item.quantity_picked = None
            last_item.save()
            
            # 因為退回了物品，如果任務本來已經標記完成，也要改回未完成
            if task.is_completed:
                task.is_completed = False
                task.save()
                
            # 將 Session 回推到上一個儲位（為了簡單起見，這裡我們讓演算法以退回的這個儲位為中心出發）
            if last_item.location:
                request.session['current_picking_x'] = last_item.location.x_coordinate
                request.session['current_picking_y'] = last_item.location.y_coordinate
                
    return redirect('interactive_picking:wizard', task_id=task_id)
