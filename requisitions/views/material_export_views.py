from django.shortcuts import render, redirect, get_object_or_404
from django.contrib.auth.decorators import login_required
from django.contrib import messages
from django.utils import timezone
from ..forms import RequisitionForm, UploadFileForm, OrderModelUploadForm, MaterialDetailsUploadForm, RequisitionItemMaterialConfirmationFormSet, RequisitionItemSignOffFormSet, UpdateProcessTypeDBForm, UploadInventoryFileForm, ProcessTypeForm, RequisitionImageUploadForm, WorkOrderMaterialImageUploadForm
from ..models import Requisition, RequisitionItem, MaterialListVersion, WorkOrderMaterial, Inventory, MachineModel, ProcessType, RequisitionImage, WorkOrderMaterialTransaction, WorkOrderMaterialImage
from inventory.models import Material
from django.db import transaction
import openpyxl
import os
from django.conf import settings
from django.contrib.auth.models import User, Group
from django.db.models import Q, F, Value, DecimalField, OuterRef, Subquery, Exists, ExpressionWrapper, Sum, Count
from django.db import models
from django.db.models.functions import Coalesce
from django.core.paginator import Paginator, EmptyPage, PageNotAnInteger
from django.urls import reverse
from django.db import IntegrityError
import pandas as pd
from django.http import JsonResponse, HttpResponse
from django.template.loader import render_to_string
import io
import json
from decimal import Decimal, InvalidOperation
import tempfile
from requisitions.utils import process_order_model_excel, process_material_details_excel
import datetime
from .requisition_management_views import _filter_requisitions

@login_required
def export_archived_requisitions_excel(request):
    requisitions = Requisition.objects.filter(is_archived=True)
    
    # Prepare data for Requisitions sheet
    requisition_data = {
        "訂單": [r.order_number for r in requisitions],
        "需求流程": [r.get_process_type_display() for r in requisitions],
        "申請人": [r.applicant.username for r in requisitions],
        "需求日期": [r.request_date.strftime('%Y-%m-%d') if r.request_date else '' for r in requisitions],
        "狀態": [r.get_status_display() for r in requisitions],
        "建立時間": [r.created_at.strftime('%Y-%m-%d %H:%M') for r in requisitions],
        "是否已歸檔": ["是" if r.is_archived else "否" for r in requisitions],
    }
    df_requisitions = pd.DataFrame(requisition_data)

    # Prepare data for Requisition Items sheet
    all_requisition_items = []
    for req in requisitions:
        if req.current_material_list_version:
            items = req.current_material_list_version.items.all().select_related('source_material')
            for item in items:
                all_requisition_items.append({
                    "訂單單號": req.order_number,
                    "需求流程": req.get_process_type_display(),
                    "物料": item.material_number,
                    "品名": item.item_name,
                    "需求數量": item.required_quantity,
                    "庫存數量": item.stock_quantity,
                    "撥料數量 (實際撥出)": item.confirmed_quantity if item.confirmed_quantity is not None else '',
                    "最終簽收已確認": "是" if item.is_signed_off else "否",
                })
    df_items = pd.DataFrame(all_requisition_items)

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_requisitions.to_excel(writer, index=False, sheet_name='已歸檔撥料申請單')
        if not df_items.empty:
            df_items.to_excel(writer, index=False, sheet_name='已歸檔撥料物料明細')
        else:
            empty_df_items = pd.DataFrame(columns=[
                "訂單", "需求流程", "物料", "物料說明", "需求數量", "撥料數量 (實際撥出)", "最終簽收已確認"
            ])
            empty_df_items.to_excel(writer, index=False, sheet_name='已歸檔撥料物料明細')

    output.seek(0)

    response = HttpResponse(
        output,
        content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )
    response['Content-Disposition'] = f'attachment; filename="archived_requisitions.xlsx"'
    return response

@login_required
def export_requisitions_excel(request):
    requisitions = _filter_requisitions(request)
    
    # Prepare data for Requisitions sheet
    requisition_data = {
        "訂單": [r.order_number for r in requisitions],
        "需求流程": [r.get_process_type_display() for r in requisitions],
        "申請人": [r.applicant.username for r in requisitions],
        "需求日期": [r.request_date.strftime('%Y-%m-%d') if r.request_date else '' for r in requisitions],
        "狀態": [r.get_status_display() for r in requisitions],
        "建立時間": [r.created_at.strftime('%Y-%m-%d %H:%M') for r in requisitions],
    }
    df_requisitions = pd.DataFrame(requisition_data)

    # Prepare data for Requisition Items sheet
    all_requisition_items = []
    for req in requisitions:
        if req.current_material_list_version:
            items = req.current_material_list_version.items.all().select_related('source_material')
            for item in items:
                all_requisition_items.append({
                    "訂單單號": req.order_number,
                    "需求流程": req.get_process_type_display(),
                    "物料": item.material_number,
                    "品名": item.item_name,
                    "需求數量": item.required_quantity,
                    "庫存數量": item.stock_quantity,
                    "撥料數量 (實際撥出)": item.confirmed_quantity if item.confirmed_quantity is not None else '',
                    "最終簽收已確認": "是" if item.is_signed_off else "否",
                })
    df_items = pd.DataFrame(all_requisition_items)

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_requisitions.to_excel(writer, index=False, sheet_name='撥料申請單')
        if not df_items.empty:
            df_items.to_excel(writer, index=False, sheet_name='撥料物料明細')
        else:
            # If no items, create an empty sheet with headers
            empty_df_items = pd.DataFrame(columns=[
                "訂單", "需求流程", "物料", "物料說明", "需求數量", "撥料數量 (實際撥出)", "最終簽收已確認"
            ])
            empty_df_items.to_excel(writer, index=False, sheet_name='撥料物料明細')

    output.seek(0)

    response = HttpResponse(
        output,
        content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )
    response['Content-Disposition'] = 'attachment; filename="requisition_list_with_materials.xlsx"'
    return response

@login_required
def export_work_order_materials_excel(request):
    order_number = request.GET.get('order_number', None)
    materials = WorkOrderMaterial.objects.none()

    if order_number:
        materials = WorkOrderMaterial.objects.filter(order_number=order_number)
    
    data = {
        "訂單單號": [m.order_number for m in materials],
        "物料": [m.material_number for m in materials],
        "物料說明": [m.item_name for m in materials],
        "需求數量": [m.required_quantity for m in materials],
        "投料點": [m.get_process_type_display() for m in materials],
        "已撥料數量": [m.confirmed_quantity if m.confirmed_quantity is not None else '' for m in materials],
        "簽收狀態": ["已簽收" if m.is_signed_off else "未簽收" for m in materials],
    }
    df = pd.DataFrame(data)

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='WorkOrderMaterials')
    output.seek(0)

    response = HttpResponse(
        output,
        content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )
    response['Content-Disposition'] = f'attachment; filename="order_materials_{order_number}.xlsx"'
    return response

@login_required
def export_archived_work_order_materials_excel(request):
    order_number = request.GET.get('order_number', None)
    materials = WorkOrderMaterial.objects.none()

    if order_number:
        materials = WorkOrderMaterial.objects.filter(order_number=order_number, is_active=False)
    
    data = {
        "訂單單號": [m.order_number for m in materials],
        "物料": [m.material_number for m in materials],
        "物料說明": [m.item_name for m in materials],
        "需求數量": [m.required_quantity for m in materials],
        "投料點": [m.get_process_type_display() for m in materials],
        "已撥料數量": [m.confirmed_quantity if m.confirmed_quantity is not None else '' for m in materials],
        "簽收狀態": ["已簽收" if m.is_signed_off else "未簽收" for m in materials],
        "是否啟用": ["是" if m.is_active else "否" for m in materials],
    }
    df = pd.DataFrame(data)

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='ArchivedWorkOrderMaterials')
    output.seek(0)

    response = HttpResponse(
        output,
        content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )
    return response

@login_required
def generate_dispatch_note(request, pk):
    requisition = get_object_or_404(Requisition, pk=pk)

    is_admin = request.user.is_superuser
    is_applicant = request.user == requisition.applicant
    is_material_handler = request.user.groups.filter(name='撥料人員').exists()

    if not (is_admin or is_applicant or is_material_handler):
        messages.error(request, "您沒有權限訪問此頁面。")
        return redirect('requisitions:homepage')

    dispatcher_subquery = WorkOrderMaterialTransaction.objects.filter(
        work_order_material=OuterRef('pk'),
        transaction_type='ALLOCATION'
    ).order_by('-timestamp').values('user__username')[:1]

    materials = WorkOrderMaterial.objects.filter(
        order_number=requisition.order_number,
        process_type__name=requisition.process_type,
        is_active=True # Only active materials
    ).filter(confirmed_quantity__gt=0).annotate(
        dispatcher_name=Subquery(dispatcher_subquery)
    )

    # Excel Export Logic
    if 'excel' in request.path:
        data = {
            "物料": [m.material_number for m in materials],
            "品名": [m.item_name for m in materials],
            "需求數量": [m.required_quantity for m in materials],
            "已撥料數量": [m.confirmed_quantity for m in materials],
            "撥料人員": [m.dispatcher_name for m in materials],
        }
        df = pd.DataFrame(data)

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='DispatchNote')
        output.seek(0)

        response = HttpResponse(
            output,
            content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
        response['Content-Disposition'] = f'attachment; filename="dispatch_note_{requisition.order_number}_{requisition.process_type}.xlsx"'
        return response

    # Fetch images related to this requisition and its process type
    if requisition.is_archived:
        dispatch_note_images = WorkOrderMaterialImage.objects.none()
    else:
        dispatch_note_images = WorkOrderMaterialImage.objects.filter(
            requisition=requisition,
            process_type__name=requisition.process_type
        ).order_by('-uploaded_at')

    context = {
        'requisition': requisition,
        'materials': materials,
        'dispatch_note_images': dispatch_note_images,
    }
    return render(request, 'requisitions/dispatch_note.html', context)

@login_required
def export_backorder_note_excel(request, pk):
    requisition = get_object_or_404(Requisition, pk=pk)

    is_admin = request.user.is_superuser
    is_applicant = request.user == requisition.applicant
    is_material_handler = request.user.groups.filter(name='撥料人員').exists()

    if not (is_admin or is_applicant or is_material_handler):
        messages.error(request, "您沒有權限訪問此頁面。")
        return redirect('requisitions:homepage')

    # Subquery to get storage_bin and stock_quantity from Inventory
    inventory_subquery_storage_bin = Subquery(
        Inventory.objects.filter(material_number=OuterRef('material_number')).values('storage_bin')[:1]
    )
    inventory_subquery_stock_quantity = Subquery(
        Inventory.objects.filter(material_number=OuterRef('material_number')).values('stock_quantity')[:1]
    )

    # Filter for active materials where required_quantity > confirmed_quantity
    # and are associated with this specific requisition and its process type
    shortage_materials = WorkOrderMaterial.objects.filter(
        is_active=True,
        required_quantity__gt=F('confirmed_quantity'),
        order_number=requisition.order_number,
        process_type__name=requisition.process_type
    ).annotate(
        shortage_quantity=ExpressionWrapper(
            F('required_quantity') - Coalesce(F('confirmed_quantity'), 0),
            output_field=DecimalField()
        ),
        storage_bin=inventory_subquery_storage_bin,
        stock_quantity=inventory_subquery_stock_quantity
    ).order_by('order_number', 'material_number').distinct()

    data = {
        "物料": [m.material_number for m in shortage_materials],
        "品名": [m.item_name for m in shortage_materials],
        "需求數量": [m.required_quantity for m in shortage_materials],
        "已撥料數量": [m.confirmed_quantity for m in shortage_materials],
        "欠料數量": [m.shortage_quantity for m in shortage_materials],
        "儲格": [m.storage_bin for m in shortage_materials],
        "庫存": [m.stock_quantity for m in shortage_materials],
    }
    df = pd.DataFrame(data)

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Backorder')

    response = HttpResponse(
        output.getvalue(),
        content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )
    response['Content-Disposition'] = f'attachment; filename=backorder_{pk}.xlsx'
    return response

@login_required
def export_material_confirmation_excel(request, pk):
    requisition = get_object_or_404(Requisition, pk=pk)
    
    if not requisition.current_material_list_version:
        messages.error(request, "此申請單沒有當前物料清單，無法匯出。")
        return redirect('material_confirmation', pk=pk)

    items = RequisitionItem.objects.filter(material_list_version=requisition.current_material_list_version)

    data = {
        "工單單號": [item.order_number for item in items],
        "物料": [item.material_number for item in items],
        "物料說明": [item.item_name for item in items],
        "需求數量": [item.required_quantity for item in items],
        
        "撥料數量 (實際撥出)": [item.confirmed_quantity if item.confirmed_quantity is not None else '' for item in items],
    }
    df = pd.DataFrame(data)

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='MaterialConfirmation')
    output.seek(0);

    response = HttpResponse(
        output,
        content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )
    response['Content-Disposition'] = f'attachment; filename="material_confirmation_{requisition.order_number}.xlsx"'
    return response

@login_required
def export_all_pending_materials_excel(request):
    # Filter for requisitions that are in 'pending' status
    pending_requisitions = Requisition.objects.filter(status='pending').select_related('applicant')

    all_pending_requisition_items = []
    for req in pending_requisitions:
        if req.current_material_list_version:
            items = req.current_material_list_version.items.all().select_related('source_material')
            for item in items:
                all_pending_requisition_items.append({
                    "訂單單號": req.order_number,
                    "需求流程": req.get_process_type_display(),
                    "物料": item.material_number,
                    "品名": item.item_name,
                    "需求數量": item.required_quantity,
                    "庫存數量": item.stock_quantity,
                    "撥料數量 (實際撥出)": item.confirmed_quantity if item.confirmed_quantity is not None else '',
                    "最終簽收已確認": "是" if item.is_signed_off else "否",
                    "申請單狀態": req.get_status_display(),
                    "申請人": req.applicant.username,
                    "申請日期": req.request_date.strftime('%Y-%m-%d') if req.request_date else '',
                })
    
    df_pending_items = pd.DataFrame(all_pending_requisition_items)

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        if not df_pending_items.empty:
            df_pending_items.to_excel(writer, index=False, sheet_name='所有待撥料物料')
        else:
            # If no items, create an empty sheet with headers
            empty_df_pending_items = pd.DataFrame(columns=[
                "訂單", "需求流程", "物料", "物料說明", "需求數量", "撥料數量 (實際撥出)", "最終簽收已確認", "申請單狀態", "申請人", "申請日期"
            ])
            empty_df_pending_items.to_excel(writer, index=False, sheet_name='所有待撥料物料')

    output.seek(0)

    response = HttpResponse(
        output,
        content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )
    response['Content-Disposition'] = 'attachment; filename="all_pending_materials.xlsx"'
    return response