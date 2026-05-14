"""
簡易介面匯出視圖 - Excel 匯出功能
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

from ..models import Requisition, RequisitionItem, WorkOrderMaterial, WorkOrder, ProcessType, RequisitionShareGroup, Announcement, MachineModel, WorkOrderMaterialTransaction
from ..constants import GROUP_NAMES, PROCESS_CATEGORY_NAMES, PROCESS_CATEGORY_COLORS
from ..forms import RequisitionForm
from inventory.models import Material
from ..utils import get_sap_user, _update_requisition_alert
from common.permissions import is_simple_applicant, is_simple_dispatcher

@login_required
def export_simple_applicant_requisitions_excel(request):
    """申請人員匯出 Excel"""
    if not is_simple_applicant(request.user) and not request.user.is_superuser:
        return redirect('requisitions:requisition_list')
    
    current_type = request.GET.get('type', 'finished')
    
    # 取得被授權查看的使用者列表
    share_groups = request.user.requisition_share_groups.all()
    viewable_owners = User.objects.filter(requisition_share_groups__in=share_groups).distinct()
    
    base_qs = Requisition.objects.filter(
        Q(applicant=request.user) | Q(applicant__in=viewable_owners)
    ).filter(requisition_type=current_type)
    
    # 如果是主管，顯示所有申請單
    is_supervisor = request.user.groups.filter(name=GROUP_NAMES['APPLICANT_SUPERVISOR']).exists()
    if is_supervisor:
        base_qs = Requisition.objects.filter(requisition_type=current_type)

    # 包含所有狀態
    requisitions = base_qs.order_by('-created_at')
    return _generate_simple_export_excel(requisitions, f"applicant_requisitions_{current_type}")


@login_required
def export_simple_dispatcher_requisitions_excel(request, category):
    """撥料人員匯出 Excel (特定分類)"""
    if not is_simple_dispatcher(request.user) and not request.user.is_superuser:
        return redirect('requisitions:requisition_list')
    
    current_type = request.GET.get('type', 'finished')
    order_numbers = request.GET.getlist('orders')
    
    if current_type == 'semi_finished':
        # 半成品：category 是 applicant.username
        qs = Requisition.objects.filter(
            applicant__username=category,
            requisition_type='semi_finished'
        )
    else:
        # 成品：category 是 process_type 的一部分
        qs = Requisition.objects.filter(
            process_type__icontains=category,
            requisition_type='finished'
        )
        
    if order_numbers:
        qs = qs.filter(order_number__in=order_numbers)
        
    requisitions = qs.order_by('-created_at')

    return _generate_simple_export_excel(requisitions, f"dispatcher_requisitions_{category}_{current_type}")


def _generate_simple_export_excel(requisitions, filename_prefix):
    """通用產生 Excel 函數 - 調整為以物料清單為主，並對齊系統標準格式"""
    
    # 1. 準備申請單摘要資料 (將用於第二個分頁)
    req_list_data = []
    for r in requisitions:
        # 取得申請人顯示名稱 (與網頁一致)
        applicant_name = f"{r.applicant.first_name}{r.applicant.last_name}"
        if not applicant_name:
            applicant_name = r.applicant.username

        req_list_data.append({
            "訂單": r.order_number,
            "需求流程": r.process_type,
            "申請人": applicant_name,
            "需求日期": r.request_date.strftime('%Y-%m-%d') if r.request_date else '',
            "狀態": r.get_status_display(),
            "建立時間": r.created_at.strftime('%Y-%m-%d %H:%M'),
        })
    df_reqs = pd.DataFrame(req_list_data)

    # 2. 準備物料清單資料 (主表，第一個分頁)
    item_data = []
    # 使用 select_related 抓取撥料人員，避免 N+1 查詢
    items = RequisitionItem.objects.filter(requisition__in=requisitions).exclude(material_number='PARENT_SCOPE').select_related('requisition', 'requisition__applicant', 'dispatched_by').order_by('requisition__order_number', 'material_number')
    for item in items:
        # 取得申請人顯示名稱
        applicant_name = f"{item.requisition.applicant.first_name}{item.requisition.applicant.last_name}"
        if not applicant_name:
            applicant_name = item.requisition.applicant.username

        # 取得撥料人員顯示名稱
        dispatched_name = ""
        if item.dispatched_by:
            dispatched_name = f"{item.dispatched_by.first_name}{item.dispatched_by.last_name}"
            if not dispatched_name:
                dispatched_name = item.dispatched_by.username

        item_data.append({
            "訂單單號": item.requisition.order_number,
            "需求流程": item.requisition.process_type,
            "物料": item.material_number,
            "品名": item.item_name,
            "需求數量": item.required_quantity,
            "庫存數量": item.stock_quantity if item.stock_quantity is not None else 0,
            "撥料數量 (實際撥出)": item.confirmed_quantity if item.confirmed_quantity is not None else '',
            "最終簽收已確認": "是" if item.is_signed_off else "否",
            "儲格": item.storage_bin,
            "撥料人員": dispatched_name,
            "申請人": applicant_name,
        })
    df_items = pd.DataFrame(item_data)

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # 重要：將「物料明細」放在第一個分頁索引，確保使用者一開啟就能看到
        if not df_items.empty:
            df_items.to_excel(writer, index=False, sheet_name='撥料物料明細')
        else:
            # 若無物料，建立帶有標題的空工作表
            empty_df_items = pd.DataFrame(columns=[
                "訂單單號", "需求流程", "物料", "品名", "需求數量", "庫存數量", "撥料數量 (實際撥出)", "最終簽收已確認"
            ])
            empty_df_items.to_excel(writer, index=False, sheet_name='撥料物料明細')
            
        # 將「申請單摘要」放在第二個分頁
        df_reqs.to_excel(writer, index=False, sheet_name='撥料申請單')
    
    output.seek(0)
    response = HttpResponse(
        output,
        content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )
    filename = f"{filename_prefix}_{timezone.now().strftime('%Y%m%d_%H%M')}.xlsx"
    # 確保檔名是 ASCII 或是經過編碼，這裡簡單使用英文
    response['Content-Disposition'] = f'attachment; filename="{filename}"'
    return response


@login_required
def export_single_requisition_excel(request, pk):
    """單張申請單詳情匯出 Excel"""
    requisition = get_object_or_404(Requisition, pk=pk)
    
    # 權限檢查邏輯與詳情頁面一致
    is_applicant = requisition.applicant == request.user
    is_dispatcher = is_simple_dispatcher(request.user)
    is_supervisor = request.user.groups.filter(name__in=[GROUP_NAMES['APPLICANT_SUPERVISOR'], GROUP_NAMES['DISPATCHER_SUPERVISOR']]).exists()
    
    is_group_member = RequisitionShareGroup.objects.filter(
        members=request.user
    ).filter(members=requisition.applicant).exists()
    
    if not (is_applicant or is_dispatcher or is_supervisor or request.user.is_superuser or is_group_member):
        return redirect('requisitions:requisition_list')
        
    # 重用之前的導出函數，傳入只包含此單的 queryset
    requisitions = Requisition.objects.filter(pk=pk)
    return _generate_simple_export_excel(requisitions, f"requisition_{requisition.order_number}")
