from django.shortcuts import render, redirect, get_object_or_404
from django.contrib.auth.decorators import login_required
from django.contrib import messages
from django.utils import timezone
from ..forms import UploadFileForm, OrderModelUploadForm, MaterialDetailsUploadForm, UpdateProcessTypeDBForm, UploadInventoryFileForm, BulkUploadForm
from ..models import Requisition, RequisitionItem, WorkOrderMaterial, Inventory, MachineModel, ProcessType
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
from requisitions.utils import process_order_model_excel, process_material_details_excel, process_inventory_excel
import datetime


@login_required
def bulk_upload(request):
    """一鍵更新：同時處理庫存、訂單機型、物料明細的上傳"""
    if not request.user.is_superuser:
        messages.error(request, "您沒有權限執行此操作。")
        return redirect('core:homepage')

    if request.method == 'POST':
        form = BulkUploadForm(request.POST, request.FILES)
        if form.is_valid():
            results = []
            
            # Process Inventory File
            inventory_file = request.FILES.get('inventory_file')
            if inventory_file:
                try:
                    with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as temp_file:
                        for chunk in inventory_file.chunks():
                            temp_file.write(chunk)
                        temp_file_path = temp_file.name
                    created, updated = process_inventory_excel(temp_file_path)
                    results.append(f"✅ 庫存資料：新增 {created} 筆，更新 {updated} 筆")
                    os.unlink(temp_file_path)
                except Exception as e:
                    results.append(f"❌ 庫存資料錯誤：{str(e)}")
            
            # Process Order Model File
            order_model_file = request.FILES.get('order_model_file')
            if order_model_file:
                try:
                    with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as temp_file:
                        for chunk in order_model_file.chunks():
                            temp_file.write(chunk)
                        temp_file_path = temp_file.name
                    created, updated = process_order_model_excel(temp_file_path)
                    results.append(f"✅ 訂單機型：新增 {created} 筆，更新 {updated} 筆")
                    os.unlink(temp_file_path)
                except Exception as e:
                    results.append(f"❌ 訂單機型錯誤：{str(e)}")
            
            # Process Material Details File
            material_details_file = request.FILES.get('material_details_file')
            if material_details_file:
                try:
                    with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as temp_file:
                        for chunk in material_details_file.chunks():
                            temp_file.write(chunk)
                        temp_file_path = temp_file.name
                    required_qty_col = form.cleaned_data.get('required_quantity_col') or '需求數量'
                    result = process_material_details_excel(temp_file_path, required_qty_col)
                    created = result['created_count']
                    updated = result['updated_count']
                    deactivated = result['deactivated_count']
                    unknown_ops = result.get('unknown_operations', [])
                    results.append(f"✅ 成品物料明細：新增 {created} 筆，更新 {updated} 筆，停用 {deactivated} 筆")
                    if unknown_ops:
                        request.session['unknown_operations'] = unknown_ops
                        results.append(f"⚠️ 發現 {len(unknown_ops)} 個未知的作業說明，需設定投料點")
                    os.unlink(temp_file_path)
                except Exception as e:
                    results.append(f"❌ 成品物料明細錯誤：{str(e)}")
            
            # Process Semi-Finished Material File
            semi_finished_file = request.FILES.get('semi_finished_file')
            if semi_finished_file:
                try:
                    from requisitions.utils import process_semi_finished_excel
                    with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as temp_file:
                        for chunk in semi_finished_file.chunks():
                            temp_file.write(chunk)
                        temp_file_path = temp_file.name
                    required_qty_col = form.cleaned_data.get('required_quantity_col') or '需求數量'
                    result = process_semi_finished_excel(temp_file_path, required_qty_col)
                    created = result['created_count']
                    updated = result['updated_count']
                    deactivated = result['deactivated_count']
                    results.append(f"✅ 半成品物料明細：新增 {created} 筆，更新 {updated} 筆，停用 {deactivated} 筆")
                    os.unlink(temp_file_path)
                except Exception as e:
                    results.append(f"❌ 半成品物料明細錯誤：{str(e)}")

            # Process Semi-Finished Model Database File (New)
            semi_finished_model_file = request.FILES.get('semi_finished_model_file')
            if semi_finished_model_file:
                try:
                    with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as temp_file:
                        for chunk in semi_finished_model_file.chunks():
                            temp_file.write(chunk)
                        temp_file_path = temp_file.name
                    # Reuse process_order_model_excel as logic is identical (Order -> Machine Model)
                    created, updated = process_order_model_excel(temp_file_path)
                    results.append(f"✅ 半成品機型對照表：新增 {created} 筆，更新 {updated} 筆")
                    os.unlink(temp_file_path)
                except Exception as e:
                    results.append(f"❌ 半成品機型對照表錯誤：{str(e)}")
            
            if not results:
                messages.warning(request, "請至少選擇一個檔案上傳。")
            else:
                for result in results:
                    if result.startswith("✅"):
                        messages.success(request, result)
                    else:
                        messages.error(request, result)
                return redirect('requisitions:bulk_upload')
    else:
        form = BulkUploadForm()
    
    return render(request, 'requisitions/bulk_upload.html', {'form': form})



@login_required
def upload_order_model_excel(request):
    if not request.user.is_superuser:
        messages.error(request, "您沒有權限執行此操作。")
        return redirect('core:homepage')

    if request.method == 'POST':
        form = OrderModelUploadForm(request.POST, request.FILES)
        if form.is_valid():
            excel_file = request.FILES['file']
            
            # Save the uploaded file to a temporary location
            with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as temp_file:
                for chunk in excel_file.chunks():
                    temp_file.write(chunk)
                temp_file_path = temp_file.name

            try:
                created_count, updated_count = process_order_model_excel(temp_file_path)
                messages.success(request, f"訂單與機型資料同步成功！新增 {created_count} 筆，更新 {updated_count} 筆。")
                return redirect('core:homepage')

            except Exception as e:
                messages.error(request, f"上傳檔案時發生錯誤: {e}")
                import traceback
                print(traceback.format_exc())
            finally:
                # Clean up the temporary file
                os.unlink(temp_file_path)

    # For GET request, render the upload form page
    form = OrderModelUploadForm()
    return render(request, 'requisitions/upload_order_model.html', {'form': form})

@login_required
def upload_material_details_excel(request):
      if not request.user.is_superuser:
          messages.error(request, "您沒有權限執行此操作。")
          return redirect('core:homepage')

      if request.method == 'POST':
          form = MaterialDetailsUploadForm(request.POST, request.FILES)
          if form.is_valid():
              excel_file = request.FILES['file']
              required_qty_col = form.cleaned_data['required_quantity_col']
              
              # Save to temp file because utils function expects a path
              with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as temp_file:
                  for chunk in excel_file.chunks():
                      temp_file.write(chunk)
                  temp_file_path = temp_file.name

              try:
                  # Call the optimized utility function
                  result = process_material_details_excel(temp_file_path, required_qty_col)
                  created_count = result['created_count']
                  updated_count = result['updated_count']
                  deactivated_count = result['deactivated_count']
                  unknown_ops = result.get('unknown_operations', [])
                  
                  messages.success(request, f"物料明細同步成功！新增 {created_count} 筆，更新 {updated_count} 筆，停用 {deactivated_count} 筆。")
                  
                  if unknown_ops:
                      request.session['unknown_operations'] = unknown_ops
                      return redirect('requisitions:classify_operations')
                  
                  return redirect('core:homepage')

              except Exception as e:
                  messages.error(request, f"上傳檔案時發生錯誤: {e}")
                  import traceback
                  print(traceback.format_exc())
              finally:
                  if os.path.exists(temp_file_path):
                      os.unlink(temp_file_path)
          else:
              messages.error(request, "表單驗證失敗。")

      # For GET request, render the upload form page
      form = MaterialDetailsUploadForm()
      context = {'form': form}
      return render(request, 'requisitions/upload_material_details.html', context)


@login_required
def upload_inventory_data(request): # This will now be for stock quantity only
    if not request.user.is_superuser:
        messages.error(request, "您沒有權限執行此操作。")
        return redirect('core:homepage')

    if request.method == 'POST':
        form = UploadInventoryFileForm(request.POST, request.FILES)
        if form.is_valid():
            excel_file = request.FILES['file']
            try:
                df = pd.read_excel(excel_file, dtype=str)
                df.columns = df.columns.str.strip()

                if '物料' not in df.columns:
                    raise ValueError("Excel 檔案中找不到 '物料' 欄位。")
                if '未限制' not in df.columns:
                    raise ValueError("Excel 檔案中找不到 '未限制' (庫存數量) 欄位。")
                # '儲格' is no longer expected here

                updated_count = 0
                created_count = 0

                with transaction.atomic():
                    for index, row in df.iterrows():
                        material_number = row.get('物料')
                        stock_quantity_str = row.get('未限制')

                        if not material_number or not stock_quantity_str:
                            messages.warning(request, f"跳過第 {index+2} 行: 物料或庫存數量為空。")
                            continue

                        try:
                            stock_quantity = float(stock_quantity_str)
                        except ValueError:
                            messages.warning(request, f"跳過第 {index+2} 行: 無效的庫存數量 '{stock_quantity_str}'.")
                            continue

                        # Only update stock_quantity, storage_bin should be preserved if it exists
                        obj, created = Inventory.objects.update_or_create(
                            material_number=material_number,
                            defaults={
                                'stock_quantity': stock_quantity,
                            }
                        )
                        if created:
                            created_count += 1
                        else:
                            updated_count += 1
                
                messages.success(request, f"庫存資料上傳成功！新增 {created_count} 筆，更新 {updated_count} 筆。")
                return redirect('core:homepage')

            except Exception as e:
                messages.error(request, f"上傳檔案時發生錯誤: {e}")
                import traceback
                print(traceback.format_exc())
        else: # Add this else block for invalid form
            messages.error(request, "上傳失敗：請檢查檔案格式或欄位是否正確。")
    # For GET request, render the upload form page
    form = UploadInventoryFileForm()
    return render(request, 'requisitions/upload_inventory_data.html', {'form': form})

@login_required
def update_process_type_db(request):
    if not request.user.is_superuser:
        messages.error(request, "您沒有權限執行此操作。")
        return redirect('core:homepage')

    if request.method == 'POST':
        form = UpdateProcessTypeDBForm(request.POST, request.FILES)
        if form.is_valid():
            excel_file = request.FILES['file']
            db_path = os.path.join(settings.BASE_DIR, 'output.xlsx')
            
            try:
                if not excel_file.name.endswith(('.xlsx', '.xls')):
                    raise Exception("上傳的檔案必須是 Excel 檔案 (.xlsx, .xls)。")

                with open(db_path, 'wb+') as destination:
                    for chunk in excel_file.chunks():
                        destination.write(chunk)
                
                messages.success(request, "投料點資料庫 (output.xlsx) 已成功更新！")
                return redirect('core:homepage')
            except Exception as e:
                messages.error(request, f"更新資料庫時發生錯誤: {e}")
        else: # Add this else block for invalid form
            messages.error(request, "上傳失敗：請檢查檔案格式是否正確。")
    # For GET request, render the upload form page
    form = UpdateProcessTypeDBForm()
    return render(request, 'requisitions/update_process_type_db.html', {'form': form})