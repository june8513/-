"""
缺料查詢與管理視圖 - 缺料彙整、查詢、匯出
"""
from django.shortcuts import render, redirect, get_object_or_404
from django.urls import reverse
from django.contrib.auth.decorators import login_required
from django.contrib import messages
from django.db.models import Q
from django.contrib.auth.models import User
from django.utils import timezone
from django.db import transaction
from django.db.models import Prefetch, F, Exists, OuterRef, Case, When, Value, BooleanField, Q, Sum, DecimalField, Count, IntegerField, CharField
from django.db.models.functions import Coalesce
from django.http import JsonResponse, HttpResponse
from django.views.decorators.http import require_POST
from decimal import Decimal
from datetime import date
import pandas as pd
import io
import traceback
import os

from ..models import Requisition, RequisitionItem, WorkOrderMaterial, WorkOrder, ProcessType, RequisitionShareGroup, Announcement, MachineModel, WorkOrderMaterialTransaction
from ..constants import GROUP_NAMES, PROCESS_CATEGORY_NAMES, PROCESS_CATEGORY_COLORS
from ..forms import RequisitionForm
from inventory.models import Material
from ..utils import get_sap_user, _update_requisition_alert
from common.permissions import is_simple_applicant, is_simple_dispatcher

@login_required
def simple_requisition_change_detail(request, pk):
    """
    需求變更管理分頁 - 簡易版
    分成上下兩個欄位：
    1. 上：已撥料需退料（confirmed > required && not dismissed）
    2. 下：新增的、未撥料的，以及上面已解除的明細
    """
    requisition = get_object_or_404(Requisition, pk=pk)
    
    # 基本權限檢查
    if not is_simple_applicant(request.user) and not is_simple_dispatcher(request.user) and not request.user.is_superuser:
        messages.error(request, "您沒有權限訪問此頁面。")
        return redirect('core:homepage')

    items = requisition.items.exclude(material_number='PARENT_SCOPE').select_related('alert_dismissed_by', 'dispatched_by')
    
    # 分類
    to_return_items = []
    other_changes = []
    
    for item in items:
        # 已撥料需退料: 已撥 > 需求
        if (item.confirmed_quantity or 0) > (item.required_quantity or 0):
            to_return_items.append(item)
        elif (item.confirmed_quantity or 0) > (item.required_quantity or 0):
            # 需求 > 已撥 (新增或增量)
            other_changes.append(item)
        elif item.alert_dismissed:
            # 即使數量一致，但若曾經有警示紀錄（且被標記過），也顯示在變更紀錄中
            other_changes.append(item)

    context = {
        'requisition': requisition,
        'to_return_items': to_return_items,
        'other_changes': other_changes,
        'alert_messages': requisition.alert_message.split('\n') if requisition.alert_message else [],
    }
    return render(request, 'requisitions/simple/simple_requisition_change_detail.html', context)


@login_required
def dismiss_requisition_item_alert(request, item_pk):
    """
    解除單筆物料的警示
    """
    item = get_object_or_404(RequisitionItem, pk=item_pk)
    requisition = item.requisition
    
    if request.method == 'POST':
        item.alert_dismissed = True
        item.alert_dismissed_by = request.user
        item.alert_dismissed_at = timezone.now()
        item.save()
        
        # 檢查是否所有項目都已處理
        # 如果所有 confirmed > required 的項目都已解除，
        # 且沒有其他未處理的新增項目（這部分邏輯可以視需求調整）
        remaining_alerts = requisition.items.filter(
            confirmed_quantity__gt=F('required_quantity'),
            alert_dismissed=False
        ).exists()
        
        if not remaining_alerts:
             # 如果沒有待退料的了，可以考慮是否自動關閉全單警示
             # 或者留給使用者手動關閉。照計畫書先保留全單警示手動關閉，
             # 但如果使用者所有單項都按了解除，全單警示可能也該關閉。
             pass

        if request.headers.get('x-requested-with') == 'XMLHttpRequest':
            return JsonResponse({
                'success': True, 
                'dismissed_by': request.user.username,
                'dismissed_at': item.alert_dismissed_at.strftime('%Y-%m-%d %H:%M')
            })
            
        messages.success(request, f"物料 {item.material_number} 的警示已解除。")
        return redirect('requisitions:simple_requisition_change_detail', pk=requisition.pk)

    return HttpResponse("method not allowed", status=405)



@login_required
def shortage_inquiry(request):
    """
    缺料查詢 - 給撥料人員主管使用
    批量帶入工單號碼，自動查出庫存量低於未滿足需求數量的物料
    """
    is_supervisor = request.user.groups.filter(name=GROUP_NAMES['DISPATCHER_SUPERVISOR']).exists()
    if not (request.user.is_superuser or is_supervisor):
        messages.error(request, "您沒有權限使用此功能。")
        return redirect('core:homepage')

    results = None
    submitted_orders = ''
    order_count = 0
    total_items = 0
    shortage_rate = 0
    selected_process_type = ''
    # 收集系統中所有固定的投料點名稱
    from ..models import ProcessType, SemiFinishedProcessType
    pt_set = set(ProcessType.objects.values_list('name', flat=True))
    pt_set.update(SemiFinishedProcessType.objects.values_list('name', flat=True))
    process_types = sorted([name for name in pt_set if name])

    if request.method == 'POST':
        import re
        raw_text = request.POST.get('order_numbers', '')
        submitted_orders = raw_text
        # 解析工單號碼：支援換行、逗號、空白分隔
        order_numbers = [o.strip() for o in re.split(r'[\n\r,\s\t]+', raw_text) if o.strip()]
        selected_process_type = request.POST.get('process_type', '')

        if not order_numbers:
            messages.warning(request, "請輸入至少一個工單號碼。")
        else:
            from ..analysis import get_material_demand_analysis

            # 使用全域分析函數計算所有工單的物料需求與缺料狀態
            all_materials_analysis = get_material_demand_analysis()

            # 預先查詢所有相關工單的客戶名稱
            work_order_map = {}
            for wo in WorkOrder.objects.filter(order_number__in=order_numbers):
                work_order_map[wo.order_number] = wo

            # 預先查詢供應商資訊
            all_material_numbers = set()
            for material_key, data in all_materials_analysis.items():
                if data['is_shortage']:
                    for detail in data['detail_orders']:
                        if detail['order_number'] in order_numbers:
                            all_material_numbers.add(material_key)
                            break

            supplier_map = {}
            for mat in Material.objects.filter(material_code__in=all_material_numbers):
                supplier_map[mat.material_code] = mat.purchaser

            results = []
            seen_orders = set()
            total_items = 0
            
            for material_key, data in all_materials_analysis.items():
                # 篩選出屬於查詢工單的明細
                relevant_details = [d for d in data['detail_orders'] if d['order_number'] in order_numbers]
                
                # 投料點篩選 (在此處篩選會影響 total_items 的統計)
                if selected_process_type:
                    relevant_details = [d for d in relevant_details if d.get('process_type_name') == selected_process_type]

                if not relevant_details:
                    continue
                
                # 計算總筆數 (反映篩選後的總數)
                total_items += len(relevant_details)

                # 只顯示全域分析後確認為缺料的物料
                if not data['is_shortage']:
                    continue

                # 全域庫存
                current_stock = float(data['current_stock'])

                for detail in relevant_details:
                    order_num = detail['order_number']
                    wo = work_order_map.get(order_num)
                    
                    # 這一點的累計需求
                    cumulative_demand = float(detail['cumulative_demand'])
                    
                    # 計算这一點的缺料量 (累計需求 - 庫存，最小為 0，且不能大於該單需求量)
                    required_qty = float(detail['required_quantity'])
                    row_shortage = min(required_qty, max(0.0, cumulative_demand - current_stock))

                    # 如果此工單在累計需求下還沒造成缺料，則視情況跳過？
                    # 使用者要求是「缺料查詢」，所以如果 row_shortage 為 0，代表此工單的需求目前庫存還夠支應
                    if row_shortage <= 0:
                        continue

                    seen_orders.add(order_num)
                    results.append({
                        'order_number': order_num,
                        'machine_model': detail.get('machine_model_name'),
                        'material_number': data['material_number'],
                        'item_name': data['item_name'],
                        'process_type': detail.get('process_type_name'),
                        'required_quantity': float(detail['required_quantity']),
                        'stock_quantity': current_stock,
                        'shortage': row_shortage,
                        'total_demand': cumulative_demand,
                        'demand_date': detail.get('demand_date'),
                        'shipping_date': detail.get('shipping_date') or (wo.shipping_date if wo else None),
                        'estimated_arrival': data.get('estimated_arrival_date'),
                        'customer_name': wo.customer_name if wo else None,
                        'supplier': supplier_map.get(data['material_number']),
                    })

            # 排序：依工單號 -> 品號
            results.sort(key=lambda x: (x['order_number'], x['material_number']))
            order_count = len(seen_orders)
            shortage_rate = (len(results) / total_items * 100) if total_items > 0 else 0

    # 讀取 MPS 同步時間
    mps_sync_time = None
    import json as _json
    _ts_file = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))), 'scratch', 'mps_sync_timestamp.json')
    try:
        if os.path.exists(_ts_file):
            with open(_ts_file, 'r', encoding='utf-8') as f:
                _ts_data = _json.load(f)
                mps_sync_time = _ts_data.get('last_sync_time', '')
    except:
        pass

    return render(request, 'requisitions/simple/simple_shortage_inquiry.html', {
        'results': results,
        'submitted_orders': submitted_orders,
        'order_count': order_count,
        'total_items': total_items,
        'shortage_rate': shortage_rate,
        'process_types': process_types,
        'selected_process_type': selected_process_type,
        'mps_sync_time': mps_sync_time,
    })


@login_required
@require_POST
def shortage_inquiry_export(request):
    """
    缺料查詢結果匯出 Excel
    """
    is_supervisor = request.user.groups.filter(name=GROUP_NAMES['DISPATCHER_SUPERVISOR']).exists()
    if not (request.user.is_superuser or is_supervisor):
        messages.error(request, "您沒有權限使用此功能。")
        return redirect('core:homepage')

    import re
    raw_text = request.POST.get('order_numbers', '')
    order_numbers = [o.strip() for o in re.split(r'[\n\r,\s\t]+', raw_text) if o.strip()]
    selected_process_type = request.POST.get('process_type', '')

    if not order_numbers:
        messages.error(request, "無可匯出的資料。")
        return redirect('requisitions:shortage_inquiry')

    from ..analysis import get_material_demand_analysis

    # 使用全域分析函數
    all_materials_analysis = get_material_demand_analysis()

    # 預先查詢工單資訊
    work_order_map = {}
    for wo in WorkOrder.objects.filter(order_number__in=order_numbers):
        work_order_map[wo.order_number] = wo

    # 預先查詢供應商資訊
    all_material_numbers = set()
    for material_key, data in all_materials_analysis.items():
        if data['is_shortage']:
            for detail in data['detail_orders']:
                if detail['order_number'] in order_numbers:
                    all_material_numbers.add(material_key)
                    break

    supplier_map = {}
    for mat in Material.objects.filter(material_code__in=all_material_numbers):
        supplier_map[mat.material_code] = mat.purchaser

    rows = []
    total_items = 0
    seen_orders = set()

    for material_key, data in all_materials_analysis.items():
        relevant_details = [d for d in data['detail_orders'] if d['order_number'] in order_numbers]
        
        # 投料點篩選
        if selected_process_type:
            relevant_details = [d for d in relevant_details if d.get('process_type_name') == selected_process_type]

        if not relevant_details:
            continue

        # 計算總筆數
        total_items += len(relevant_details)

        if not data['is_shortage']:
            continue

        current_stock = float(data['current_stock'])

        for detail in relevant_details:
            order_num = detail['order_number']
            wo = work_order_map.get(order_num)
            
            cumulative_demand = float(detail['cumulative_demand'])
            # 計算这一點的缺料量 (累計需求 - 庫存，最小為 0，且不能大於該單需求量)
            required_qty = float(detail['required_quantity'])
            row_shortage = min(required_qty, max(0.0, cumulative_demand - current_stock))

            if row_shortage <= 0:
                continue

            seen_orders.add(order_num)

            rows.append({
                '工單號碼': order_num,
                '客戶': wo.customer_name if wo and wo.customer_name else '',
                '機型': detail.get('machine_model_name', ''),
                '品號': data['material_number'],
                '品名': data['item_name'],
                '供應商': supplier_map.get(data['material_number'], ''),
                '投料點': detail.get('process_type_name', ''),
                '此工單需求量': float(detail['required_quantity']),
                '累計未滿足需求': cumulative_demand,
                '庫存量': current_stock,
                '缺料': row_shortage,
                '需求日期': detail.get('demand_date') or '',
                '出貨日期': detail.get('shipping_date') or (wo.shipping_date if wo else '') or '',
                '預計入料': data.get('estimated_arrival_date') or '',
            })

    rows.sort(key=lambda x: (x['工單號碼'], x['品號']))

    # 計算統計資料
    shortage_count = len(rows)
    shortage_rate = (shortage_count / total_items * 100) if total_items > 0 else 0
    order_count = len(seen_orders)
    summary_text = f"總數/缺料數:{total_items}/{shortage_count}筆 缺料率{shortage_rate:.1f}% 工單{order_count}張"

    df = pd.DataFrame(rows)
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='缺料查詢', startrow=0)
        worksheet = writer.sheets['缺料查詢']
        
        # 寫入統計資料到最後一列
        last_row = len(rows) + 2 # 表頭佔1列，資料佔 len(rows) 列，下一列為 last_row
        cell = worksheet.cell(row=last_row, column=1, value=summary_text)
        worksheet.merge_cells(start_row=last_row, start_column=1, end_row=last_row, end_column=5)
        
        # 設定粗體與些微樣式
        from openpyxl.styles import Font
        cell.font = Font(bold=True, size=12)

    output.seek(0)

    response = HttpResponse(
        output.getvalue(),
        content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )
    response['Content-Disposition'] = f'attachment; filename="shortage_inquiry_{date.today().isoformat()}.xlsx"'
    return response


@login_required
def sync_mps_order_info(request):
    """
    從外部 API 同步工單出貨資訊 (客戶與出貨日期)
    """
    from ..utils import sync_external_order_info
    import json as _json
    from datetime import datetime as _dt
    
    success, message = sync_external_order_info()
    
    # 儲存同步時間戳
    _ts_file = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))), 'scratch', 'mps_sync_timestamp.json')
    try:
        os.makedirs(os.path.dirname(_ts_file), exist_ok=True)
        with open(_ts_file, 'w', encoding='utf-8') as f:
            _json.dump({
                'last_sync_time': _dt.now().strftime('%Y-%m-%d %H:%M:%S'),
                'success': success,
                'message': message,
            }, f, ensure_ascii=False, indent=2)
    except:
        pass
    
    if success:
        messages.success(request, message)
    else:
        messages.error(request, message)
        
    return redirect('requisitions:shortage_inquiry')

